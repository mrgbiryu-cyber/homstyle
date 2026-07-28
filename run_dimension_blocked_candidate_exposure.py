from __future__ import annotations

import json
import sqlite3
from collections import defaultdict
from datetime import datetime
from typing import Any

from bulk_homestyle_collect import DB_PATH


RUN_NAME = "dimension_blocked_candidate_exposure_v1"
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


def is_complete(row: sqlite3.Row) -> bool:
    if row["shape_type"] == "AREA_2D":
        return row["w_mm"] is not None and row["h_mm"] is not None
    return all(row[axis] is not None for axis in ("w_mm", "d_mm", "h_mm"))


def dimension_key(row: sqlite3.Row) -> tuple[Any, ...]:
    return (
        None if row["w_mm"] is None else round(float(row["w_mm"]), 3),
        None if row["d_mm"] is None else round(float(row["d_mm"]), 3),
        None if row["h_mm"] is None else round(float(row["h_mm"]), 3),
        str(row["shape_type"] or ""),
    )


def exposed_candidate(row: sqlite3.Row) -> dict[str, Any]:
    return {
        "rule_id": "CTX_BLOCKED_COMPLETE_EXPOSURE_V1",
        "raw_notation": str(row["raw_notation"] or ""),
        "context_text": str(row["context_text"] or ""),
        "section_role": str(row["section_role"] or "UNSCOPED"),
        "candidate_role": "PRODUCT_DIMENSION",
        "source_axis_signature": str(row["source_axis_signature"] or ""),
        "normalized_axis_mapping": "BLOCKED_COMPLETE->COMPARISON",
        "option_label": str(row["option_label"] or ""),
        "shape_type": str(row["shape_type"] or "RECTANGLE"),
        "unit_status": str(row["unit_status"] or ""),
        "unit_text": str(row["unit_text"] or ""),
        "w_raw": row["w_raw"],
        "d_raw": row["d_raw"],
        "h_raw": row["h_raw"],
        "l_raw": row["l_raw"],
        "r_raw": row["r_raw"],
        "value_1_raw": row["value_1_raw"],
        "value_2_raw": row["value_2_raw"],
        "value_3_raw": row["value_3_raw"],
        "w_mm": row["w_mm"],
        "d_mm": row["d_mm"],
        "h_mm": row["h_mm"],
        "diameter_mm": row["diameter_mm"],
        "product_name_match_score": int(row["product_name_match_score"] or 0),
        "candidate_score": min(79, int(row["candidate_score"] or 0)),
        "decision_status": "HUMAN_REVIEW",
        "rejection_reason": (
            "COMPLETE_DIMENSION_EXISTS_BUT_AUTOMATIC_SELECTION_WAS_BLOCKED"
        ),
    }


def main() -> None:
    timestamp = now_text()
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA journal_mode=WAL")
    connection.execute("PRAGMA synchronous=NORMAL")
    init_schema(connection)

    target_ids = {
        row["product_id"]
        for row in connection.execute(
            """
            SELECT product_id
            FROM fact_dimension_resolution_ledger
            WHERE resolution_status IN ('OCR_REQUIRED', 'NO_CANDIDATE')
            UNION
            SELECT DISTINCT product_id
            FROM stg_dimension_targeted_ocr_candidate
            WHERE run_name = ?
            UNION
            SELECT DISTINCT product_id
            FROM fact_dimension_comparison_candidate
            WHERE normalized_axis_mapping = 'BLOCKED_COMPLETE->COMPARISON'
            """,
            (RUN_NAME,),
        )
    }

    by_product: dict[str, list[sqlite3.Row]] = defaultdict(list)
    for row in connection.execute(
        """
        SELECT *
        FROM stg_dimension_context_candidate
        WHERE is_current = 1
          AND candidate_role = 'PRODUCT_DIMENSION'
          AND decision_status != 'REJECT'
          AND source_ref NOT LIKE 'dimension_blocked_candidate_exposure_v1:%'
        ORDER BY
            product_id,
            COALESCE(product_name_match_score, 0) DESC,
            COALESCE(candidate_score, 0) DESC,
            COALESCE(source_priority, 0) DESC,
            candidate_key
        """
    ):
        if row["product_id"] in target_ids and is_complete(row):
            by_product[row["product_id"]].append(row)

    connection.execute(
        "DELETE FROM stg_dimension_targeted_ocr_candidate WHERE run_name = ?",
        (RUN_NAME,),
    )

    inserted = 0
    product_count = 0
    truncated_products = 0
    for product_id, rows in sorted(by_product.items()):
        deduplicated: list[sqlite3.Row] = []
        seen: set[tuple[Any, ...]] = set()
        for row in rows:
            key = dimension_key(row)
            if key in seen:
                continue
            seen.add(key)
            deduplicated.append(row)
        if len(deduplicated) > MAX_CANDIDATES_PER_PRODUCT:
            truncated_products += 1
            deduplicated = deduplicated[:MAX_CANDIDATES_PER_PRODUCT]
        if not deduplicated:
            continue
        product_count += 1
        for candidate_no, row in enumerate(deduplicated, 1):
            candidate = exposed_candidate(row)
            evidence = (
                f"[comparison exposure from {row['source_type']} / "
                f"{row['normalized_axis_mapping']}] "
                f"{candidate['context_text']}"
            )[:12_000]
            candidate["context_text"] = evidence
            connection.execute(
                """
                INSERT INTO stg_dimension_targeted_ocr_candidate
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """,
                (
                    RUN_NAME,
                    timestamp,
                    product_id,
                    str(row["image_url"] or ""),
                    9800,
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
                "target_products": len(target_ids),
                "products_with_exposed_candidates": product_count,
                "exposed_candidates": inserted,
                "truncated_products": truncated_products,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    connection.close()


if __name__ == "__main__":
    main()
