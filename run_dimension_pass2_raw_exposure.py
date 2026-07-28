from __future__ import annotations

import json
import re
import sqlite3
from collections import defaultdict
from datetime import datetime
from typing import Any

from bulk_homestyle_collect import DB_PATH
from build_dimension_context_normalization import USER_CASES
from dimension_context_normalizer import category_profile


RUN_NAME = "dimension_pass2_raw_exposure_v1"
SOURCE_RUN = "dimension_scan_pass2_layout_v1"
MAX_CANDIDATES_PER_PRODUCT = 24


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


def unit_factor(row: sqlite3.Row, profile: str) -> tuple[float, str]:
    unit = str(row["unit_text"] or "").casefold()
    notation = str(row["raw_notation"] or "").casefold()
    if unit == "cm" or ("cm" in notation and "mm" not in notation):
        return 10.0, "EXPLICIT_CM"
    if unit == "mm" or "mm" in notation:
        return 1.0, "EXPLICIT_MM"
    if unit in {"inch", "in"} or re.search(r"(?:inch|\\bin\\b|[″”\"])", notation):
        return 25.4, "EXPLICIT_INCH"

    values = [
        float(row[column])
        for column in (
            "w_raw",
            "d_raw",
            "h_raw",
            "l_raw",
            "value_1_raw",
            "value_2_raw",
            "value_3_raw",
        )
        if row[column] is not None
    ]
    maximum = max(values, default=0)
    # Unitless detail drawings generally use millimetres above 300 and
    # centimetres below that scale. This remains comparison-only evidence.
    if maximum and maximum <= 300 and profile in {
        "FURNITURE",
        "AREA_2D",
        "ROUND_FURNITURE",
        "OVAL_FURNITURE",
    }:
        return 10.0, "INFERRED_CM_FOR_COMPARISON"
    return 1.0, "INFERRED_MM_FOR_COMPARISON"


def valid(value: float | None) -> bool:
    return value is not None and 0 < float(value) <= 20_000


def to_mm(value: Any, factor: float) -> float | None:
    if value is None:
        return None
    converted = round(float(value) * factor, 3)
    return converted if valid(converted) else None


def observation_candidate(
    row: sqlite3.Row,
    product: sqlite3.Row,
) -> dict[str, Any] | None:
    profile = category_profile(
        str(product["product_name"] or ""),
        str(product["small_category"] or ""),
    )
    factor, factor_reason = unit_factor(row, profile)
    w_mm = to_mm(row["w_mm"], 1.0) or to_mm(row["w_raw"], factor)
    d_mm = to_mm(row["d_mm"], 1.0) or to_mm(row["d_raw"], factor)
    h_mm = to_mm(row["h_mm"], 1.0) or to_mm(row["h_raw"], factor)
    l_mm = to_mm(row["l_mm"], 1.0) or to_mm(row["l_raw"], factor)
    value_1 = to_mm(row["value_1_raw"], factor)
    value_2 = to_mm(row["value_2_raw"], factor)
    value_3 = to_mm(row["value_3_raw"], factor)
    axis_signature = str(row["axis_signature"] or "")
    mapping_detail = axis_signature or str(row["candidate_type"] or "RAW_VALUES")

    if l_mm is not None and w_mm is not None and h_mm is not None and d_mm is None:
        # L/W/H in furniture diagrams means overall length / depth / height.
        w_mm, d_mm = l_mm, w_mm
        mapping_detail = "L,W,H->W,D,H"
    elif w_mm is None and d_mm is None and h_mm is None:
        if value_1 is not None and value_2 is not None and value_3 is not None:
            w_mm, d_mm, h_mm = value_1, value_2, value_3
            mapping_detail = "RAW_ORDER_1,2,3->W,D,H"
        elif value_1 is not None and value_2 is not None:
            if profile == "AREA_2D":
                w_mm, h_mm = value_1, value_2
                mapping_detail = "RAW_ORDER_1,2->W,H;D=N/A"
            else:
                w_mm, d_mm = value_1, value_2
                mapping_detail = "RAW_ORDER_1,2->W,D"
        elif value_1 is not None:
            w_mm = value_1
            mapping_detail = "RAW_VALUE_1->W"
        elif l_mm is not None:
            w_mm = l_mm
            mapping_detail = "RAW_L->W"

    if not any(valid(value) for value in (w_mm, d_mm, h_mm)):
        return None

    shape_type = "AREA_2D" if profile == "AREA_2D" and d_mm is None else "RECTANGLE"
    evidence = " ".join(str(row["evidence_text"] or "").split())
    mapping_note = f"{mapping_detail};{factor_reason}"
    candidate = {
        "rule_id": "CTX_PASS2_RAW_EXPOSURE_V1",
        "raw_notation": str(row["raw_notation"] or ""),
        "context_text": (
            f"[pass2 raw comparison / {mapping_note}] {evidence}"
        )[:12_000],
        "section_role": "UNSCOPED",
        "candidate_role": "PRODUCT_DIMENSION",
        "source_axis_signature": axis_signature,
        "normalized_axis_mapping": "PASS2_RAW_VALUES->COMPARISON",
        "option_label": "",
        "shape_type": shape_type,
        "unit_status": (
            str(row["unit_status"] or "") + "|" + factor_reason
        ).strip("|"),
        "unit_text": "mm",
        "w_raw": row["w_raw"],
        "d_raw": row["d_raw"],
        "h_raw": row["h_raw"],
        "l_raw": row["l_raw"],
        "r_raw": row["r_raw"],
        "value_1_raw": row["value_1_raw"],
        "value_2_raw": row["value_2_raw"],
        "value_3_raw": row["value_3_raw"],
        "w_mm": w_mm,
        "d_mm": d_mm,
        "h_mm": h_mm,
        "diameter_mm": None,
        "product_name_match_score": 0,
        "candidate_score": min(79, int(row["candidate_score"] or 0)),
        "decision_status": "HUMAN_REVIEW",
        "rejection_reason": (
            "PASS2_RAW_VALUES_EXPOSED_WITHOUT_AUTOMATIC_CONFIRMATION"
        ),
    }
    return candidate


def forbidden_user_value(product_id: str, candidate: dict[str, Any]) -> bool:
    for case in USER_CASES:
        if case["product_id"] != product_id:
            continue
        expected = (
            case.get("forbidden_w_mm"),
            case.get("forbidden_d_mm"),
            case.get("forbidden_h_mm"),
        )
        if not any(value is not None for value in expected):
            continue
        actual = (
            candidate.get("w_mm"),
            candidate.get("d_mm"),
            candidate.get("h_mm"),
        )
        if all(
            right is None
            or (left is not None and abs(float(left) - float(right)) <= 1.0)
            for left, right in zip(actual, expected)
        ):
            return True
    return False


def candidate_key(candidate: dict[str, Any]) -> tuple[Any, ...]:
    return (
        candidate["w_mm"],
        candidate["d_mm"],
        candidate["h_mm"],
        candidate["shape_type"],
        candidate["raw_notation"],
    )


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
            WHERE l.resolution_status = 'NO_CANDIDATE'
               OR m.product_id IN (
                    SELECT DISTINCT product_id
                    FROM stg_dimension_targeted_ocr_candidate
                    WHERE run_name = ?
               )
               OR m.product_id IN (
                    SELECT DISTINCT product_id
                    FROM fact_dimension_comparison_candidate
                    WHERE normalized_axis_mapping = 'PASS2_RAW_VALUES->COMPARISON'
               )
            """,
            (RUN_NAME,),
        )
    }

    observations: dict[str, list[sqlite3.Row]] = defaultdict(list)
    for row in connection.execute(
        """
        SELECT DISTINCT
            pi.product_id,
            pi.image_order,
            o.*
        FROM stg_dimension_scan_pass2_product_image pi
        JOIN stg_dimension_scan_pass2_observation o
          ON o.run_name = pi.run_name
         AND o.url_hash = pi.url_hash
        WHERE pi.run_name = ?
        ORDER BY
            pi.product_id,
            CASE WHEN COALESCE(o.unit_text, '') != '' THEN 0 ELSE 1 END,
            COALESCE(o.candidate_score, 0) DESC,
            pi.image_order,
            o.observation_no
        """,
        (SOURCE_RUN,),
    ):
        if row["product_id"] in products:
            observations[row["product_id"]].append(row)

    connection.execute(
        "DELETE FROM stg_dimension_targeted_ocr_candidate WHERE run_name = ?",
        (RUN_NAME,),
    )

    inserted = 0
    product_count = 0
    inferred_unit_candidates = 0
    truncated_products = 0
    for product_id, rows in sorted(observations.items()):
        product = products[product_id]
        candidates: list[tuple[sqlite3.Row, dict[str, Any]]] = []
        seen: set[tuple[Any, ...]] = set()
        for row in rows:
            candidate = observation_candidate(row, product)
            if candidate is None or forbidden_user_value(product_id, candidate):
                continue
            key = candidate_key(candidate)
            if key in seen:
                continue
            seen.add(key)
            candidates.append((row, candidate))
        if len(candidates) > MAX_CANDIDATES_PER_PRODUCT:
            truncated_products += 1
            candidates = candidates[:MAX_CANDIDATES_PER_PRODUCT]
        if not candidates:
            continue
        product_count += 1
        counters_by_image: dict[str, int] = defaultdict(int)
        for row, candidate in candidates:
            image_url = str(row["image_url"] or "")
            counters_by_image[image_url] += 1
            candidate_no = counters_by_image[image_url]
            if "INFERRED_" in candidate["unit_status"]:
                inferred_unit_candidates += 1
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
                    9900,
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
                "target_products": len(products),
                "products_with_raw_comparison": product_count,
                "raw_comparison_candidates": inserted,
                "unit_inferred_candidates": inferred_unit_candidates,
                "truncated_products": truncated_products,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    connection.close()


if __name__ == "__main__":
    main()
