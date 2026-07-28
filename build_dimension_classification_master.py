from __future__ import annotations

import json
import sqlite3
from collections import Counter
from datetime import datetime
from typing import Any

from bulk_homestyle_collect import DB_PATH, RUN_DIR
from low_dimension_quality_policy import assess_low_dimension


TABLE = "stg_dimension_classification_master"
SNAPSHOT_ID = "dimension-classification-master-v1-20260723"
TAXONOMY_VERSION = "dimension-classification-master-v1.0"
SUMMARY_PATH = RUN_DIR / "dimension_classification_master_latest.json"

PRE_PASS2_GROUPS = {
    "G01_COMPLETE_CANDIDATE",
    "G02_CURRENT_PARTIAL",
    "G03_OCR_PARTIAL_AXES",
    "G04_SIZE_LABEL_UNIT_MISSING",
    "G05_UNCLASSIFIED_OCR",
}


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def create_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        f"""
        CREATE TABLE IF NOT EXISTS {TABLE} (
            snapshot_id TEXT NOT NULL,
            classified_at TEXT NOT NULL,
            is_current INTEGER NOT NULL,
            taxonomy_version TEXT NOT NULL,
            product_id TEXT NOT NULL,
            product_name TEXT,
            mid_category TEXT,
            small_category TEXT,
            classification_complete INTEGER NOT NULL,
            classification_status TEXT NOT NULL,
            classification_stage TEXT NOT NULL,
            classification_code TEXT NOT NULL,
            classification_name TEXT NOT NULL,
            dimension_resolution_status TEXT NOT NULL,
            dimension_value_confirmed INTEGER NOT NULL,
            confirmed_w_mm REAL,
            confirmed_d_mm REAL,
            confirmed_h_mm REAL,
            candidate_w_mm REAL,
            candidate_d_mm REAL,
            candidate_h_mm REAL,
            pass1_work_group_code TEXT,
            pass1_pattern_code TEXT,
            pass2_status TEXT,
            quality_flag TEXT,
            requires_followup INTEGER NOT NULL,
            followup_priority INTEGER,
            next_action TEXT,
            best_image_url TEXT,
            evidence_text TEXT,
            PRIMARY KEY(snapshot_id,product_id)
        );

        CREATE INDEX IF NOT EXISTS idx_dimension_classification_master_current
            ON {TABLE}(is_current,classification_complete,classification_stage);
        CREATE INDEX IF NOT EXISTS idx_dimension_classification_master_resolution
            ON {TABLE}(is_current,dimension_resolution_status,requires_followup);
        CREATE INDEX IF NOT EXISTS idx_dimension_pattern_queue_current_product
            ON stg_dimension_pattern_work_queue(is_current,product_id);
        CREATE INDEX IF NOT EXISTS idx_dimension_pass2_reclass_current_product
            ON stg_dimension_pass2_reclassification(is_current,product_id);

        DROP VIEW IF EXISTS vw_dimension_classification_master_current;
        DROP VIEW IF EXISTS vw_dimension_classification_complete_current;
        DROP VIEW IF EXISTS vw_dimension_classification_remaining_current;
        DROP VIEW IF EXISTS vw_dimension_classification_summary;
        DROP VIEW IF EXISTS vw_dimension_resolution_summary;

        CREATE VIEW vw_dimension_classification_master_current AS
        SELECT * FROM {TABLE} WHERE is_current=1;

        CREATE VIEW vw_dimension_classification_complete_current AS
        SELECT * FROM {TABLE}
        WHERE is_current=1 AND classification_complete=1;

        CREATE VIEW vw_dimension_classification_remaining_current AS
        SELECT * FROM {TABLE}
        WHERE is_current=1 AND classification_complete=0
        ORDER BY followup_priority,small_category,product_id;

        CREATE VIEW vw_dimension_classification_summary AS
        SELECT classification_status,classification_stage,
               COUNT(*) AS product_count
        FROM {TABLE}
        WHERE is_current=1
        GROUP BY classification_status,classification_stage
        ORDER BY classification_status,classification_stage;

        CREATE VIEW vw_dimension_resolution_summary AS
        SELECT dimension_resolution_status,dimension_value_confirmed,
               requires_followup,COUNT(*) AS product_count
        FROM {TABLE}
        WHERE is_current=1
        GROUP BY dimension_resolution_status,dimension_value_confirmed,requires_followup
        ORDER BY dimension_value_confirmed DESC,dimension_resolution_status;
        """
    )


def load_products(connection: sqlite3.Connection) -> list[sqlite3.Row]:
    return connection.execute(
        """
        SELECT
            m.product_id,m.product_name,m.mid_category_value,m.small_category_value,
            m.size_wdh_ok,m.w_mm,m.d_mm,m.h_mm,
            q.work_group_code,q.work_group_name,q.pattern_code,q.pattern_name,
            q.next_action AS pass1_next_action,
            r.observation_classified,r.pattern_classified,r.resolution_state,
            r.requeue_pattern_code,r.requeue_pattern_name,r.pass2_status,
            r.quality_flag,r.resolved_w_mm,r.resolved_d_mm,r.resolved_h_mm,
            r.priority,r.next_action AS pass2_next_action,
            r.best_image_url,r.evidence_text
        FROM stg_mandatory_pass m
        LEFT JOIN vw_dimension_pattern_work_queue_current q
          ON q.product_id=m.product_id
        LEFT JOIN stg_dimension_pass2_reclassification r
          ON r.product_id=m.product_id AND r.is_current=1
        WHERE m.is_current=1
        ORDER BY m.product_id
        """
    ).fetchall()


def base_row(row: sqlite3.Row, classified_at: str) -> dict[str, Any]:
    return {
        "snapshot_id": SNAPSHOT_ID,
        "classified_at": classified_at,
        "is_current": 1,
        "taxonomy_version": TAXONOMY_VERSION,
        "product_id": row["product_id"],
        "product_name": row["product_name"] or "",
        "mid_category": row["mid_category_value"] or "",
        "small_category": row["small_category_value"] or "",
        "confirmed_w_mm": row["w_mm"],
        "confirmed_d_mm": row["d_mm"],
        "confirmed_h_mm": row["h_mm"],
        "candidate_w_mm": row["resolved_w_mm"],
        "candidate_d_mm": row["resolved_d_mm"],
        "candidate_h_mm": row["resolved_h_mm"],
        "pass1_work_group_code": row["work_group_code"] or "",
        "pass1_pattern_code": row["pattern_code"] or "",
        "pass2_status": row["pass2_status"] or "",
        "quality_flag": row["quality_flag"] or "",
        "best_image_url": row["best_image_url"] or "",
        "evidence_text": row["evidence_text"] or "",
    }


def classify(row: sqlite3.Row, classified_at: str) -> dict[str, Any]:
    result = base_row(row, classified_at)

    confirmed_values = (row["w_mm"], row["d_mm"], row["h_mm"])
    confirmed_values_physically_valid = all(
        isinstance(value, (int, float))
        and not isinstance(value, bool)
        and float(value) > 0
        for value in confirmed_values
    )
    low_value_assessment = assess_low_dimension(
        str(row["product_name"] or ""),
        str(row["mid_category_value"] or ""),
        str(row["small_category_value"] or ""),
        row["w_mm"],
        row["d_mm"],
        row["h_mm"],
    )
    if (
        int(row["size_wdh_ok"] or 0) == 1
        and confirmed_values_physically_valid
        and not low_value_assessment.requires_review
    ):
        result.update(
            {
                "classification_complete": 1,
                "classification_status": "COMPLETE",
                "classification_stage": "BASELINE_WDH_CONFIRMED",
                "classification_code": "B01_CONFIRMED_WDH",
                "classification_name": "기존 W/D/H 확정값 보유",
                "dimension_resolution_status": "CONFIRMED_WDH",
                "dimension_value_confirmed": 1,
                "requires_followup": 0,
                "followup_priority": None,
                "next_action": "완료",
            }
        )
        return result
    if (
        int(row["size_wdh_ok"] or 0) == 1
        and low_value_assessment.requires_review
    ):
        result["quality_flag"] = low_value_assessment.code

    work_group = str(row["work_group_code"] or "")
    if work_group in PRE_PASS2_GROUPS:
        resolution_by_group = {
            "G01_COMPLETE_CANDIDATE": "COMPLETE_CANDIDATE_REVIEW",
            "G02_CURRENT_PARTIAL": "CURRENT_PARTIAL_REVIEW",
            "G03_OCR_PARTIAL_AXES": "OCR_PARTIAL_AXES_REVIEW",
            "G04_SIZE_LABEL_UNIT_MISSING": "UNIT_OR_AXIS_INFERENCE_REVIEW",
            "G05_UNCLASSIFIED_OCR": "SEMANTIC_PATTERN_REVIEW",
        }
        result.update(
            {
                "classification_complete": 1,
                "classification_status": "COMPLETE",
                "classification_stage": "PRE_PASS2_PATTERN_CLASSIFIED",
                "classification_code": row["pattern_code"] or work_group,
                "classification_name": row["pattern_name"] or row["work_group_name"] or "",
                "dimension_resolution_status": resolution_by_group[work_group],
                "dimension_value_confirmed": 0,
                "requires_followup": 1,
                "followup_priority": 10,
                "next_action": row["pass1_next_action"] or "",
            }
        )
        return result

    if int(row["observation_classified"] or 0) == 1:
        pass2_resolution = str(row["resolution_state"] or "")
        resolution_by_pass2 = {
            "DIRECT_WDH_CANDIDATE": "DIRECT_WDH_CANDIDATE",
            "PATTERN_CLASSIFIED_REVIEW": "PASS2_PATTERN_REVIEW",
            "VALUE_RECLASSIFICATION_REQUIRED": "PASS2_VALUE_PATTERN_REVIEW",
        }
        result.update(
            {
                "classification_complete": 1,
                "classification_status": "COMPLETE",
                "classification_stage": "PASS2_OCR_CLASSIFIED",
                "classification_code": row["requeue_pattern_code"] or row["pass2_status"] or "",
                "classification_name": row["requeue_pattern_name"] or "",
                "dimension_resolution_status": resolution_by_pass2.get(
                    pass2_resolution,
                    "PASS2_PATTERN_REVIEW",
                ),
                "dimension_value_confirmed": 0,
                "requires_followup": 1,
                "followup_priority": row["priority"],
                "next_action": row["pass2_next_action"] or "",
            }
        )
        return result

    result.update(
        {
            "classification_complete": 0,
            "classification_status": "REMAINING",
            "classification_stage": "NO_OBSERVATION_REMAINING",
            "classification_code": row["requeue_pattern_code"] or row["pattern_code"] or "UNCLASSIFIED",
            "classification_name": row["requeue_pattern_name"] or row["work_group_name"] or "미분류",
            "dimension_resolution_status": "NO_DIMENSION_OBSERVATION",
            "dimension_value_confirmed": 0,
            "requires_followup": 1,
            "followup_priority": row["priority"] if row["priority"] is not None else 99,
            "next_action": row["pass2_next_action"] or row["pass1_next_action"] or "",
        }
    )
    return result


def build() -> dict[str, Any]:
    classified_at = now_text()
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    create_schema(connection)
    source_rows = load_products(connection)
    rows = [classify(row, classified_at) for row in source_rows]
    columns = list(rows[0])
    placeholders = ",".join(f":{column}" for column in columns)
    with connection:
        connection.execute(f"UPDATE {TABLE} SET is_current=0 WHERE is_current=1")
        connection.execute(f"DELETE FROM {TABLE} WHERE snapshot_id=?", (SNAPSHOT_ID,))
        connection.executemany(
            f"INSERT INTO {TABLE} ({','.join(columns)}) VALUES ({placeholders})",
            rows,
        )

    status_counts = Counter(row["classification_status"] for row in rows)
    stage_counts = Counter(row["classification_stage"] for row in rows)
    resolution_counts = Counter(row["dimension_resolution_status"] for row in rows)
    remaining_counts = Counter(
        row["classification_code"]
        for row in rows
        if row["classification_complete"] == 0
    )
    summary = {
        "snapshot_id": SNAPSHOT_ID,
        "classified_at": classified_at,
        "taxonomy_version": TAXONOMY_VERSION,
        "total_products": len(rows),
        "classification_status": dict(status_counts),
        "classification_stage": dict(stage_counts),
        "classification_complete_rate": status_counts["COMPLETE"] / len(rows) if rows else 0,
        "dimension_value_confirmed_products": sum(
            row["dimension_value_confirmed"] for row in rows
        ),
        "dimension_resolution_status": dict(resolution_counts),
        "remaining_reason": dict(remaining_counts),
        "table": TABLE,
        "views": [
            "vw_dimension_classification_master_current",
            "vw_dimension_classification_complete_current",
            "vw_dimension_classification_remaining_current",
            "vw_dimension_classification_summary",
            "vw_dimension_resolution_summary",
        ],
        "source_dimension_values_written": False,
        "excel_written": False,
    }
    SUMMARY_PATH.write_text(
        json.dumps(summary, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    connection.close()
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return summary


if __name__ == "__main__":
    build()
