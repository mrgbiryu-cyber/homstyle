from __future__ import annotations

import hashlib
import json
import sqlite3
from collections import defaultdict
from typing import Any

from bulk_homestyle_collect import DB_PATH
from build_dimension_context_normalization import (
    ENGINE_VERSION,
    SOURCE_PRIORITY,
    USER_CASES,
    category_profile,
    choose_product_options,
    evaluate_regression,
    init_schema,
    insert_dict_rows,
    is_complete_candidate,
    now_text,
    upsert_testcases,
)


TARGETED_RUNS = (
    "dimension_targeted_reocr_user_cases_v1",
    "dimension_targeted_reocr_remaining_v1",
    "dimension_partial_fusion_v1",
    "dimension_blocked_candidate_exposure_v1",
    "dimension_pass2_raw_exposure_v1",
)


def targeted_candidate_key(
    run_name: str,
    product_id: str,
    image_url: str,
    crop_no: int,
    candidate_no: int,
    candidate: dict[str, Any],
) -> str:
    raw = json.dumps(
        [
            run_name,
            product_id,
            image_url,
            crop_no,
            candidate_no,
            candidate.get("raw_notation"),
            candidate.get("normalized_axis_mapping"),
            candidate.get("w_mm"),
            candidate.get("d_mm"),
            candidate.get("h_mm"),
        ],
        ensure_ascii=False,
        separators=(",", ":"),
    )
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:32]


def merge_targeted_candidates(
    connection: sqlite3.Connection,
    snapshot_id: str,
    timestamp: str,
) -> int:
    connection.execute(
        """
        DELETE FROM stg_dimension_context_candidate
        WHERE snapshot_id = ?
          AND source_type IN ('TARGETED_FULL_IMAGE_OCR', 'TARGETED_REGION_OCR')
        """,
        (snapshot_id,),
    )
    rows: list[dict[str, Any]] = []
    for source in connection.execute(
        """
        SELECT
            t.*,
            m.product_name,
            m.mid_category,
            m.small_category
        FROM stg_dimension_targeted_ocr_candidate t
        JOIN vw_dimension_classification_master_current m
          ON m.product_id = t.product_id
        WHERE t.run_name IN ({targeted_run_placeholders})
        ORDER BY t.product_id, t.image_url, t.crop_no, t.candidate_no
        """.format(
            targeted_run_placeholders=",".join("?" for _ in TARGETED_RUNS)
        ),
        TARGETED_RUNS,
    ):
        candidate = json.loads(source["candidate_json"])
        source_type = (
            "TARGETED_FULL_IMAGE_OCR"
            if int(source["crop_no"]) == 0
            else "TARGETED_REGION_OCR"
        )
        row = {
            "snapshot_id": snapshot_id,
            "normalized_at": timestamp,
            "is_current": 1,
            "engine_version": ENGINE_VERSION,
            "candidate_key": targeted_candidate_key(
                source["run_name"],
                source["product_id"],
                source["image_url"],
                int(source["crop_no"]),
                int(source["candidate_no"]),
                candidate,
            ),
            "product_id": source["product_id"],
            "product_name": source["product_name"],
            "mid_category": source["mid_category"],
            "small_category": source["small_category"],
            "category_profile": category_profile(
                source["product_name"], source["small_category"]
            ),
            "source_type": source_type,
            "source_ref": (
                f"{source['run_name']}:crop={source['crop_no']}:candidate={source['candidate_no']}"
            ),
            "source_order": int(source["crop_no"]),
            "source_priority": SOURCE_PRIORITY[source_type],
            "image_url": source["image_url"],
            "candidate_no": int(source["candidate_no"]),
        }
        row.update(candidate)
        rows.append(row)
    insert_dict_rows(connection, "stg_dimension_context_candidate", rows)
    return len(rows)


def rebuild_selection(
    connection: sqlite3.Connection,
    snapshot_id: str,
    timestamp: str,
) -> dict[str, int]:
    products = list(
        connection.execute(
            """
            SELECT *
            FROM vw_dimension_classification_master_current
            WHERE dimension_value_confirmed = 0
            ORDER BY product_id
            """
        )
    )
    candidates_by_product: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in connection.execute(
        """
        SELECT *
        FROM stg_dimension_context_candidate
        WHERE snapshot_id = ?
        """,
        (snapshot_id,),
    ):
        candidates_by_product[row["product_id"]].append(dict(row))

    for table in (
        "stg_product_dimension_option",
        "stg_dimension_context_product",
        "stg_dimension_targeted_reocr_queue",
        "stg_dimension_context_regression_result",
    ):
        connection.execute(f"DELETE FROM {table} WHERE snapshot_id = ?", (snapshot_id,))

    option_rows: list[dict[str, Any]] = []
    product_rows: list[dict[str, Any]] = []
    queue_rows: list[dict[str, Any]] = []
    user_case_ids = {case["product_id"] for case in USER_CASES}

    for product in products:
        product_id = product["product_id"]
        candidates = candidates_by_product.get(product_id, [])
        status, options, requires_reocr, requires_review, next_action = (
            choose_product_options(product, candidates)
        )
        for option_no, row in enumerate(options, 1):
            option_rows.append(
                {
                    "snapshot_id": snapshot_id,
                    "normalized_at": timestamp,
                    "is_current": 1,
                    "engine_version": ENGINE_VERSION,
                    "product_id": product_id,
                    "product_name": product["product_name"],
                    "option_no": option_no,
                    "is_primary": 1 if option_no == 1 else 0,
                    "option_label": row.get("option_label")
                    or ("대표" if option_no == 1 else f"후보{option_no - 1}"),
                    "selection_status": (
                        "AMBIGUOUS_CANDIDATE"
                        if status == "AMBIGUOUS_MULTI_CANDIDATE"
                        else "PRIMARY_CANDIDATE"
                        if option_no == 1
                        else "OPTION_CANDIDATE"
                    ),
                    "candidate_key": row["candidate_key"],
                    "shape_type": row.get("shape_type"),
                    "diameter_mm": row.get("diameter_mm"),
                    "w_mm": row.get("w_mm"),
                    "d_mm": row.get("d_mm"),
                    "h_mm": row.get("h_mm"),
                    "d_applicability": (
                        "N/A_2D_CATEGORY"
                        if row.get("shape_type") == "AREA_2D"
                        else "APPLICABLE"
                    ),
                    "source_axis_signature": row.get("source_axis_signature"),
                    "normalized_axis_mapping": row.get("normalized_axis_mapping"),
                    "unit_status": row.get("unit_status"),
                    "candidate_score": row.get("candidate_score"),
                    "source_type": row.get("source_type"),
                    "image_url": row.get("image_url"),
                    "evidence_text": row.get("context_text"),
                }
            )
        representative = options[0] if options else {}
        product_rows.append(
            {
                "snapshot_id": snapshot_id,
                "normalized_at": timestamp,
                "is_current": 1,
                "engine_version": ENGINE_VERSION,
                "product_id": product_id,
                "product_name": product["product_name"],
                "mid_category": product["mid_category"],
                "small_category": product["small_category"],
                "category_profile": category_profile(
                    product["product_name"], product["small_category"]
                ),
                "context_status": status,
                "candidate_count": len(candidates),
                "rejected_candidate_count": sum(
                    row["decision_status"] == "REJECT" for row in candidates
                ),
                "complete_candidate_count": sum(
                    is_complete_candidate(row) for row in candidates
                ),
                "automatic_candidate_count": sum(
                    row["decision_status"] in {"AUTO_ACCEPT", "CATEGORY_NORMALIZED"}
                    and is_complete_candidate(row)
                    for row in candidates
                ),
                "option_count": len(options),
                "representative_candidate_key": representative.get("candidate_key"),
                "representative_w_mm": representative.get("w_mm"),
                "representative_d_mm": representative.get("d_mm"),
                "representative_h_mm": representative.get("h_mm"),
                "representative_shape_type": representative.get("shape_type"),
                "representative_d_applicability": (
                    "N/A_2D_CATEGORY"
                    if representative.get("shape_type") == "AREA_2D"
                    else "APPLICABLE"
                    if representative
                    else ""
                ),
                "requires_reocr": requires_reocr,
                "requires_human_review": requires_review,
                "next_action": next_action,
            }
        )
        if requires_reocr or product_id in user_case_ids:
            best_image_url = str(product["best_image_url"] or "")
            if not best_image_url and candidates:
                best_image_url = next(
                    (str(row.get("image_url") or "") for row in candidates if row.get("image_url")),
                    "",
                )
            axis_status = "/".join(
                axis
                for axis, value in (
                    ("W", product["candidate_w_mm"]),
                    ("D", product["candidate_d_mm"]),
                    ("H", product["candidate_h_mm"]),
                )
                if value is not None
            )
            queue_rows.append(
                {
                    "snapshot_id": snapshot_id,
                    "queued_at": timestamp,
                    "is_current": 1,
                    "product_id": product_id,
                    "product_name": product["product_name"],
                    "mid_category": product["mid_category"],
                    "small_category": product["small_category"],
                    "queue_priority": (
                        1
                        if product_id in user_case_ids
                        else 2
                        if status == "REOCR_REQUIRED"
                        else 3
                    ),
                    "queue_reason": status,
                    "target_heading": "SIZE|DIMENSION|규격|제품 사이즈|INFO|한 눈에 보기|DETAIL",
                    "target_strategy": (
                        "1차 제목/구역 탐지 → 구역 좌표 확대 → 2차 좌표 OCR → 문맥 정규화"
                    ),
                    "best_image_url": best_image_url,
                    "existing_axis_status": axis_status or "NONE",
                    "existing_evidence_text": product["evidence_text"],
                    "user_review_case": 1 if product_id in user_case_ids else 0,
                }
            )

    insert_dict_rows(connection, "stg_product_dimension_option", option_rows)
    insert_dict_rows(connection, "stg_dimension_context_product", product_rows)
    insert_dict_rows(connection, "stg_dimension_targeted_reocr_queue", queue_rows)
    regression_rows = evaluate_regression(connection, snapshot_id, timestamp)
    insert_dict_rows(
        connection, "stg_dimension_context_regression_result", regression_rows
    )
    return {
        "products": len(product_rows),
        "options": len(option_rows),
        "queue": len(queue_rows),
        "regression": len(regression_rows),
    }


def main() -> None:
    timestamp = now_text()
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA journal_mode=WAL")
    connection.execute("PRAGMA synchronous=NORMAL")
    init_schema(connection)
    upsert_testcases(connection, timestamp)
    snapshot_row = connection.execute(
        """
        SELECT snapshot_id
        FROM stg_dimension_context_product
        WHERE is_current = 1
        GROUP BY snapshot_id
        ORDER BY MAX(normalized_at) DESC
        LIMIT 1
        """
    ).fetchone()
    if not snapshot_row:
        raise RuntimeError("current dimension context snapshot not found")
    snapshot_id = snapshot_row["snapshot_id"]

    merged = merge_targeted_candidates(connection, snapshot_id, timestamp)
    counts = rebuild_selection(connection, snapshot_id, timestamp)
    connection.commit()
    summary = {
        "snapshot_id": snapshot_id,
        "targeted_candidates_merged": merged,
        **counts,
        "context_status": {
            row["context_status"]: row["product_count"]
            for row in connection.execute("SELECT * FROM vw_dimension_context_summary")
        },
        "regression": {
            row["result_status"]: row["count"]
            for row in connection.execute(
                """
                SELECT result_status, COUNT(*) AS count
                FROM vw_dimension_context_regression_current
                GROUP BY result_status
                """
            )
        },
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
