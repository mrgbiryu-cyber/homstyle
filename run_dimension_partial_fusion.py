from __future__ import annotations

import itertools
import json
import sqlite3
from collections import defaultdict
from datetime import datetime
from typing import Any, Iterable

from bulk_homestyle_collect import DB_PATH
from build_dimension_context_normalization import (
    candidate_is_physically_plausible,
    candidate_matches_title_pair,
)
from dimension_context_normalizer import category_profile


RUN_NAME = "dimension_partial_fusion_v1"
SOURCE_RUN = "dimension_targeted_reocr_remaining_v1"
MAX_COMBINATIONS_PER_GROUP = 24


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def init_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS stg_dimension_targeted_ocr_candidate (
            run_name TEXT NOT NULL,
            processed_at TEXT NOT NULL,
            product_id TEXT NOT NULL,
            image_url TEXT,
            crop_no INTEGER,
            candidate_no INTEGER,
            candidate_role TEXT,
            decision_status TEXT,
            raw_notation TEXT,
            normalized_axis_mapping TEXT,
            option_label TEXT,
            shape_type TEXT,
            w_mm REAL,
            d_mm REAL,
            h_mm REAL,
            diameter_mm REAL,
            candidate_score INTEGER,
            evidence_text TEXT,
            candidate_json TEXT,
            PRIMARY KEY(run_name, product_id, image_url, crop_no, candidate_no)
        );
        """
    )


def dimension_values(
    rows: Iterable[sqlite3.Row],
) -> dict[str, list[float]]:
    values: dict[str, set[float]] = {"w_mm": set(), "d_mm": set(), "h_mm": set()}
    for row in rows:
        for axis in values:
            value = row[axis]
            if value is not None and 0 < float(value) <= 20_000:
                values[axis].add(float(value))
    return {axis: sorted(axis_values) for axis, axis_values in values.items()}


def fusion_groups(
    rows: list[sqlite3.Row],
) -> list[tuple[int, str, list[sqlite3.Row]]]:
    """Prefer same-crop fusion, then use same-image fusion for unresolved axes."""

    by_crop: dict[tuple[int, str], list[sqlite3.Row]] = defaultdict(list)
    for row in rows:
        label = str(row["option_label"] or "").upper()
        by_crop[(int(row["crop_no"]), label)].append(row)

    groups: list[tuple[int, str, list[sqlite3.Row]]] = []
    complete_labels: set[str] = set()
    for (crop_no, label), crop_rows in sorted(by_crop.items()):
        values = dimension_values(crop_rows)
        if all(values[axis] for axis in ("w_mm", "d_mm", "h_mm")):
            groups.append((crop_no, label, crop_rows))
            complete_labels.add(label)

    by_label: dict[str, list[sqlite3.Row]] = defaultdict(list)
    for row in rows:
        by_label[str(row["option_label"] or "").upper()].append(row)
    for label, image_rows in sorted(by_label.items()):
        if label in complete_labels:
            continue
        values = dimension_values(image_rows)
        if all(values[axis] for axis in ("w_mm", "d_mm", "h_mm")):
            groups.append((9000, label, image_rows))
    return groups


def combined_context(rows: Iterable[sqlite3.Row]) -> str:
    parts: list[str] = []
    for row in rows:
        text = " ".join(str(row["evidence_text"] or "").split())
        if text and text not in parts:
            parts.append(text)
    return " | ".join(parts)[:12_000]


def build_candidate(
    product: sqlite3.Row,
    rows: list[sqlite3.Row],
    *,
    option_label: str,
    w_mm: float,
    d_mm: float,
    h_mm: float,
) -> dict[str, Any]:
    evidence = combined_context(rows)
    section_role = (
        "PRODUCT_SIZE_SECTION"
        if any(
            json.loads(str(row["candidate_json"] or "{}")).get("section_role")
            == "PRODUCT_SIZE_SECTION"
            for row in rows
        )
        else "UNSCOPED"
    )
    product_name_match_score = max(
        (
            int(
                json.loads(str(row["candidate_json"] or "{}")).get(
                    "product_name_match_score"
                )
                or 0
            )
            for row in rows
        ),
        default=0,
    )
    raw_notation = f"W={w_mm:g}mm, D={d_mm:g}mm, H={h_mm:g}mm"
    return {
        "rule_id": "CTX_PARTIAL_FUSION_V1",
        "raw_notation": raw_notation,
        "context_text": evidence,
        "section_role": section_role,
        "candidate_role": "PRODUCT_DIMENSION",
        "source_axis_signature": "W,D,H",
        "normalized_axis_mapping": "PARTIAL_FUSION->W,D,H",
        "option_label": option_label,
        "shape_type": "RECTANGLE",
        "unit_status": "UNIT_PRESENT",
        "unit_text": "mm",
        "w_raw": w_mm,
        "d_raw": d_mm,
        "h_raw": h_mm,
        "l_raw": None,
        "r_raw": None,
        "value_1_raw": None,
        "value_2_raw": None,
        "value_3_raw": None,
        "w_mm": w_mm,
        "d_mm": d_mm,
        "h_mm": h_mm,
        "diameter_mm": None,
        "product_name_match_score": product_name_match_score,
        "candidate_score": min(
            79, max((int(row["candidate_score"] or 0) for row in rows), default=0) + 5
        ),
        # A fusion is evidence for the comparison sheet, never an automatic
        # product-size confirmation.
        "decision_status": "HUMAN_REVIEW",
        "rejection_reason": "PARTIAL_AXIS_VALUES_FUSED_FOR_COMPARISON",
    }


def main() -> None:
    timestamp = now_text()
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA journal_mode=WAL")
    connection.execute("PRAGMA synchronous=NORMAL")
    init_schema(connection)

    products = {
        row["product_id"]: row
        for row in connection.execute(
            """
            SELECT m.*
            FROM vw_dimension_classification_master_current m
            JOIN fact_dimension_resolution_ledger l
              ON l.product_id = m.product_id
            WHERE l.resolution_status IN ('OCR_REQUIRED', 'NO_CANDIDATE')
               OR m.product_id IN (
                    SELECT DISTINCT product_id
                    FROM stg_dimension_targeted_ocr_candidate
                    WHERE run_name = 'dimension_partial_fusion_v1'
               )
               OR m.product_id IN (
                    SELECT DISTINCT product_id
                    FROM fact_dimension_comparison_candidate
                    WHERE normalized_axis_mapping = 'PARTIAL_FUSION->W,D,H'
               )
            """
        )
    }
    source_rows: dict[tuple[str, str], list[sqlite3.Row]] = defaultdict(list)
    for row in connection.execute(
        """
        SELECT *
        FROM stg_dimension_targeted_ocr_candidate
        WHERE run_name = ?
          AND decision_status != 'REJECT'
          AND candidate_role = 'PRODUCT_DIMENSION'
          AND (
              w_mm IS NOT NULL
              OR d_mm IS NOT NULL
              OR h_mm IS NOT NULL
          )
        ORDER BY product_id, image_url, crop_no, candidate_no
        """,
        (SOURCE_RUN,),
    ):
        if row["product_id"] in products:
            source_rows[(row["product_id"], row["image_url"])].append(row)

    connection.execute(
        "DELETE FROM stg_dimension_targeted_ocr_candidate WHERE run_name = ?",
        (RUN_NAME,),
    )

    inserted = 0
    product_ids: set[str] = set()
    truncated_groups = 0
    for (product_id, image_url), rows in sorted(source_rows.items()):
        product = products[product_id]
        if category_profile(
            str(product["product_name"] or ""),
            str(product["small_category"] or ""),
        ) == "AREA_2D":
            # Flat artwork/rug dimensions use W/H with D=N/A. Combining
            # independent OCR values into a synthetic W/D/H triple would
            # incorrectly turn size options into depth/height.
            continue
        for group_no, (crop_no, option_label, group_rows) in enumerate(
            fusion_groups(rows), 1
        ):
            values = dimension_values(group_rows)
            combinations = list(
                itertools.product(
                    values["w_mm"], values["d_mm"], values["h_mm"]
                )
            )
            if len(combinations) > MAX_COMBINATIONS_PER_GROUP:
                truncated_groups += 1
                combinations = combinations[:MAX_COMBINATIONS_PER_GROUP]

            output_crop_no = crop_no if crop_no < 9000 else 9000 + group_no
            candidate_no = 0
            seen_dimensions: set[tuple[float, float, float]] = set()
            for w_mm, d_mm, h_mm in combinations:
                dimension = (w_mm, d_mm, h_mm)
                if dimension in seen_dimensions:
                    continue
                seen_dimensions.add(dimension)
                candidate = build_candidate(
                    product,
                    group_rows,
                    option_label=option_label,
                    w_mm=w_mm,
                    d_mm=d_mm,
                    h_mm=h_mm,
                )
                if not candidate_is_physically_plausible(product, candidate):
                    continue
                if not candidate_matches_title_pair(product, candidate):
                    continue
                candidate_no += 1
                connection.execute(
                    """
                    INSERT INTO stg_dimension_targeted_ocr_candidate
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """,
                    (
                        RUN_NAME,
                        timestamp,
                        product_id,
                        image_url,
                        output_crop_no,
                        candidate_no,
                        candidate["candidate_role"],
                        candidate["decision_status"],
                        candidate["raw_notation"],
                        candidate["normalized_axis_mapping"],
                        candidate["option_label"],
                        candidate["shape_type"],
                        candidate["w_mm"],
                        candidate["d_mm"],
                        candidate["h_mm"],
                        candidate["diameter_mm"],
                        candidate["candidate_score"],
                        candidate["context_text"],
                        json.dumps(candidate, ensure_ascii=False),
                    ),
                )
                inserted += 1
                product_ids.add(product_id)

    connection.executescript(
        """
        DROP VIEW IF EXISTS vw_dimension_targeted_ocr_candidates_current;
        CREATE VIEW vw_dimension_targeted_ocr_candidates_current AS
        SELECT *
        FROM stg_dimension_targeted_ocr_candidate
        WHERE run_name IN (
            'dimension_targeted_reocr_user_cases_v1',
            'dimension_targeted_reocr_remaining_v1',
            'dimension_partial_fusion_v1',
            'dimension_blocked_candidate_exposure_v1',
            'dimension_pass2_raw_exposure_v1'
        );
        """
    )
    connection.commit()
    print(
        json.dumps(
            {
                "run_name": RUN_NAME,
                "source_run": SOURCE_RUN,
                "source_product_image_groups": len(source_rows),
                "products_with_fusion_candidates": len(product_ids),
                "fusion_candidates": inserted,
                "truncated_groups": truncated_groups,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    connection.close()


if __name__ == "__main__":
    main()
