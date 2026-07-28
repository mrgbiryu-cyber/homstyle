from __future__ import annotations

import json
import sqlite3
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any

from bulk_homestyle_collect import DB_PATH, RUN_DIR


PASS2_RUN_NAME = "dimension_scan_pass2_layout_v1"
SNAPSHOT_ID = "dimension-pass2-reclassification-v1-20260723"
TAXONOMY_VERSION = "dimension-pass2-reclassification-v1.0"
TABLE = "stg_dimension_pass2_reclassification"
SUMMARY_PATH = RUN_DIR / "dimension_pass2_reclassification_latest.json"

PASS1_SCOPE_GROUPS = {
    "G06_PASS1_IMAGE_CANDIDATE",
    "G07_PASS1_NO_SIGNAL",
    "G08_NO_DETAIL_IMAGE",
    "G09_SCAN_INCOMPLETE",
}

AREA_2D_TERMS = (
    "러그",
    "카펫",
    "매트",
    "액자",
    "포스터",
    "스프레드",
    "태피스트리",
)

CLASSIFIED_STATUS = {
    "COMPLETE_EXPLICIT_WDH_CANDIDATE": (
        "DIRECT_WDH_CANDIDATE",
        "C01_EXPLICIT_WDH_WITH_UNIT",
        "명시적 W/D/H와 단위",
        "대표 규격·옵션 충돌 검증 후 적용",
    ),
    "COMPLETE_WDH_UNIT_MISSING": (
        "PATTERN_CLASSIFIED_REVIEW",
        "C02_EXPLICIT_WDH_UNIT_REQUIRED",
        "명시적 W/D/H·단위 검증 필요",
        "단위를 확인한 뒤 W/D/H 정규화",
    ),
    "COMPLETE_NONSTANDARD_AXES_REVIEW": (
        "PATTERN_CLASSIFIED_REVIEW",
        "C03_NONSTANDARD_AXES",
        "L/W/H 등 비표준 축",
        "카테고리와 도면 기준으로 L/W/H 의미 해석",
    ),
    "TRIPLET_MAPPING_REVIEW": (
        "PATTERN_CLASSIFIED_REVIEW",
        "C04_UNLABELED_TRIPLET",
        "축 미표기 3개 값",
        "카테고리별 순서 규칙으로 W/D/H 매핑",
    ),
    "PARTIAL_AXES_REVIEW": (
        "PATTERN_CLASSIFIED_REVIEW",
        "C05_PARTIAL_LABELED_AXES",
        "일부 축 확보",
        "확보 축을 유지하고 누락 축 보강",
    ),
}


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def category_profile(category: str) -> str:
    return "AREA_2D_CANDIDATE" if any(term in category for term in AREA_2D_TERMS) else "OBJECT_3D_CANDIDATE"


def unit_present(unit_status: str) -> bool:
    return unit_status not in {"", "UNIT_MISSING"}


def normalized_values_mm(values: list[float], unit_text: str) -> list[float] | None:
    unit = unit_text.strip().lower().replace("㎜", "mm").replace("㎝", "cm")
    factor = {"mm": 1.0, "cm": 10.0, "m": 1000.0}.get(unit)
    if factor is None:
        return None
    return [value * factor for value in values]


def quality_flag(values: list[float], unit_text: str, unit_status: str) -> str:
    if not values:
        return "NO_RAW_VALUES"
    if any(value <= 0 for value in values):
        return "NONPOSITIVE_VALUE_REVIEW"
    normalized = normalized_values_mm(values, unit_text)
    if normalized is None:
        ratio = max(values) / min(values) if min(values) else float("inf")
        if max(values) >= 10000 or ratio >= 100:
            return "OCR_DIGIT_JOIN_REVIEW"
        return "UNIT_REQUIRED"
    if min(normalized) < 10 or max(normalized) > 5000:
        return "PHYSICAL_RANGE_REVIEW"
    if not unit_present(unit_status):
        return "UNIT_REQUIRED"
    return "BASIC_RANGE_PLAUSIBLE"


def numeric_requeue_pattern(
    candidate_type: str,
    profile: str,
    has_unit: bool,
) -> tuple[str, str, int, str]:
    value_form = "PAIR" if candidate_type == "UNLABELED_PAIR" else "TRIPLE"
    profile_code = "2D" if profile == "AREA_2D_CANDIDATE" else "3D"
    unit_code = "UNIT" if has_unit else "NO_UNIT"
    patterns = {
        ("PAIR", "2D", "UNIT"): (
            "V01_2D_PAIR_WITH_UNIT",
            "2D 카테고리·2개 값·단위 있음",
            1,
            "W/H 또는 가로/세로 패턴 후보로 묶고 두께 비적용 정책 검토",
        ),
        ("TRIPLE", "3D", "UNIT"): (
            "V02_3D_TRIPLE_WITH_UNIT",
            "3D 카테고리·3개 값·단위 있음",
            2,
            "OCR 근거에서 값 순서를 확인해 W/D/H 매핑 규칙 생성",
        ),
        ("TRIPLE", "2D", "UNIT"): (
            "V03_2D_TRIPLE_WITH_UNIT",
            "2D 카테고리·3개 값·단위 있음",
            3,
            "두께 포함 규격인지 옵션별 크기 나열인지 구분",
        ),
        ("PAIR", "3D", "UNIT"): (
            "V04_3D_PAIR_WITH_UNIT",
            "3D 카테고리·2개 값·단위 있음",
            4,
            "누락 축 또는 지름×높이 표기 여부 확인",
        ),
        ("TRIPLE", "3D", "NO_UNIT"): (
            "V05_3D_TRIPLE_UNIT_REQUIRED",
            "3D 카테고리·3개 값·단위 없음",
            5,
            "상세 이미지 문맥·카테고리 기준으로 단위부터 검증",
        ),
        ("TRIPLE", "2D", "NO_UNIT"): (
            "V06_2D_TRIPLE_UNIT_REQUIRED",
            "2D 카테고리·3개 값·단위 없음",
            6,
            "두께/옵션 구분과 단위를 함께 검증",
        ),
        ("PAIR", "3D", "NO_UNIT"): (
            "V07_3D_PAIR_UNIT_REQUIRED",
            "3D 카테고리·2개 값·단위 없음",
            7,
            "단위와 누락 축을 함께 보강",
        ),
        ("PAIR", "2D", "NO_UNIT"): (
            "V08_2D_PAIR_UNIT_REQUIRED",
            "2D 카테고리·2개 값·단위 없음",
            8,
            "가로/세로 단위를 검증하고 두께 비적용 정책 검토",
        ),
    }
    return patterns[(value_form, profile_code, unit_code)]


def no_observation_pattern(
    pass2_status: str,
    work_group_code: str,
) -> tuple[str, str, int, str]:
    if pass2_status == "OCR_INCOMPLETE_NO_CANDIDATE":
        return (
            "N01_PASS2_IMAGE_FETCH_RETRY",
            "2차 이미지 재수집 필요",
            20,
            "실패 URL을 다시 수집한 뒤 좌표 OCR 재실행",
        )
    if pass2_status == "NO_DIMENSION_OBSERVATION":
        return (
            "N02_PASS2_NO_PATTERN",
            "2차 OCR 성공·규격 패턴 없음",
            21,
            "OCR 전체 원문과 이미지를 사람이 확인하고 탐지 어휘 확장",
        )
    if work_group_code == "G09_SCAN_INCOMPLETE":
        return (
            "N03_PASS1_SOURCE_RETRY",
            "1차 이미지 소스 재수집 필요",
            22,
            "상세 이미지 URL을 다시 수집한 뒤 1차 스캔부터 재실행",
        )
    if work_group_code == "G08_NO_DETAIL_IMAGE":
        return (
            "N04_NO_DETAIL_IMAGE",
            "상세 이미지 없음",
            23,
            "HTML/API/외부 판매처에서 상세 이미지 또는 규격 정보 보강",
        )
    return (
        "N05_PASS1_NO_SIZE_SIGNAL",
        "전체 이미지 규격 신호 없음",
        24,
        "HTML·FAQ/Q&A·옵션·외부 상품 설명 순서로 다른 소스 조사",
    )


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
            small_category TEXT,
            pass1_work_group_code TEXT,
            pass1_pattern_code TEXT,
            pass1_product_classification TEXT,
            pass2_run_name TEXT,
            pass2_status TEXT,
            observation_classified INTEGER NOT NULL,
            pattern_classified INTEGER NOT NULL,
            resolution_state TEXT NOT NULL,
            requeue_required INTEGER NOT NULL,
            requeue_pattern_code TEXT NOT NULL,
            requeue_pattern_name TEXT NOT NULL,
            category_profile TEXT,
            value_form TEXT,
            value_count INTEGER NOT NULL,
            unit_status TEXT,
            unit_text TEXT,
            quality_flag TEXT,
            raw_notation TEXT,
            raw_value_1 REAL,
            raw_value_2 REAL,
            raw_value_3 REAL,
            resolved_w_mm REAL,
            resolved_d_mm REAL,
            resolved_h_mm REAL,
            best_image_url TEXT,
            evidence_text TEXT,
            priority INTEGER NOT NULL,
            next_action TEXT,
            PRIMARY KEY(snapshot_id,product_id)
        );

        CREATE INDEX IF NOT EXISTS idx_dimension_pass2_reclass_current
            ON {TABLE}(is_current,resolution_state,requeue_required,priority);
        CREATE INDEX IF NOT EXISTS idx_dimension_pass2_reclass_pattern
            ON {TABLE}(is_current,requeue_pattern_code,small_category);

        DROP VIEW IF EXISTS vw_dimension_pass2_observed_classified_current;
        DROP VIEW IF EXISTS vw_dimension_pass2_pattern_classified_current;
        DROP VIEW IF EXISTS vw_dimension_pass2_value_reclassification_current;
        DROP VIEW IF EXISTS vw_dimension_pass2_unclassified_current;
        DROP VIEW IF EXISTS vw_dimension_pass2_requeue_current;
        DROP VIEW IF EXISTS vw_dimension_pass2_requeue_summary;

        CREATE VIEW vw_dimension_pass2_observed_classified_current AS
        SELECT * FROM {TABLE}
        WHERE is_current=1 AND observation_classified=1;

        CREATE VIEW vw_dimension_pass2_pattern_classified_current AS
        SELECT * FROM {TABLE}
        WHERE is_current=1 AND pattern_classified=1;

        CREATE VIEW vw_dimension_pass2_value_reclassification_current AS
        SELECT * FROM {TABLE}
        WHERE is_current=1 AND resolution_state='VALUE_RECLASSIFICATION_REQUIRED';

        CREATE VIEW vw_dimension_pass2_unclassified_current AS
        SELECT * FROM {TABLE}
        WHERE is_current=1 AND resolution_state='NO_OBSERVATION_REQUEUE';

        CREATE VIEW vw_dimension_pass2_requeue_current AS
        SELECT * FROM {TABLE}
        WHERE is_current=1 AND requeue_required=1
        ORDER BY priority,quality_flag,small_category,product_id;

        CREATE VIEW vw_dimension_pass2_requeue_summary AS
        SELECT resolution_state,requeue_pattern_code,requeue_pattern_name,
               category_profile,value_form,unit_status,quality_flag,
               MIN(priority) AS priority,COUNT(*) AS product_count
        FROM {TABLE}
        WHERE is_current=1 AND requeue_required=1
        GROUP BY resolution_state,requeue_pattern_code,requeue_pattern_name,
                 category_profile,value_form,unit_status,quality_flag
        ORDER BY priority,requeue_pattern_code,product_count DESC;
        """
    )


def load_scope(connection: sqlite3.Connection) -> list[sqlite3.Row]:
    placeholders = ",".join("?" for _ in PASS1_SCOPE_GROUPS)
    return connection.execute(
        f"""
        SELECT
            q.product_id,q.product_name,q.small_category,
            q.work_group_code,q.pattern_code,q.pass1_product_classification,
            p.run_name AS pass2_run_name,p.pass2_status,
            p.best_candidate_type,p.best_raw_notation,p.best_unit_status,
            p.resolved_w_mm,p.resolved_d_mm,p.resolved_h_mm,p.best_image_url,
            o.unit_text,o.value_1_raw,o.value_2_raw,o.value_3_raw,o.evidence_text
        FROM vw_dimension_pattern_work_queue_current q
        LEFT JOIN stg_dimension_scan_pass2_product p
          ON p.product_id=q.product_id AND p.run_name=?
        LEFT JOIN stg_dimension_scan_pass2_product_image pi
          ON pi.run_name=p.run_name
         AND pi.product_id=p.product_id
         AND pi.image_order=p.best_image_order
        LEFT JOIN stg_dimension_scan_pass2_observation o
          ON o.run_name=pi.run_name
         AND o.url_hash=pi.url_hash
         AND o.observation_no=p.best_observation_no
        WHERE q.work_group_code IN ({placeholders})
        ORDER BY q.product_id
        """,
        (PASS2_RUN_NAME, *sorted(PASS1_SCOPE_GROUPS)),
    ).fetchall()


def build_row(row: sqlite3.Row, classified_at: str) -> dict[str, Any]:
    pass2_status = str(row["pass2_status"] or "")
    category = str(row["small_category"] or "")
    unit_status = str(row["best_unit_status"] or "")
    values = [
        float(value)
        for value in (row["value_1_raw"], row["value_2_raw"], row["value_3_raw"])
        if value is not None
    ]
    base: dict[str, Any] = {
        "snapshot_id": SNAPSHOT_ID,
        "classified_at": classified_at,
        "is_current": 1,
        "taxonomy_version": TAXONOMY_VERSION,
        "product_id": row["product_id"],
        "product_name": row["product_name"],
        "small_category": category,
        "pass1_work_group_code": row["work_group_code"],
        "pass1_pattern_code": row["pattern_code"],
        "pass1_product_classification": row["pass1_product_classification"],
        "pass2_run_name": row["pass2_run_name"] or "",
        "pass2_status": pass2_status,
        "category_profile": category_profile(category),
        "value_form": "",
        "value_count": len(values),
        "unit_status": unit_status,
        "unit_text": row["unit_text"] or "",
        "quality_flag": "",
        "raw_notation": row["best_raw_notation"] or "",
        "raw_value_1": row["value_1_raw"],
        "raw_value_2": row["value_2_raw"],
        "raw_value_3": row["value_3_raw"],
        "resolved_w_mm": row["resolved_w_mm"],
        "resolved_d_mm": row["resolved_d_mm"],
        "resolved_h_mm": row["resolved_h_mm"],
        "best_image_url": row["best_image_url"] or "",
        "evidence_text": row["evidence_text"] or "",
    }

    if pass2_status in CLASSIFIED_STATUS:
        state, code, name, next_action = CLASSIFIED_STATUS[pass2_status]
        base.update(
            {
                "observation_classified": 1,
                "pattern_classified": 1,
                "resolution_state": state,
                "requeue_required": 0,
                "requeue_pattern_code": code,
                "requeue_pattern_name": name,
                "quality_flag": "PATTERN_ASSIGNED",
                "priority": 0,
                "next_action": next_action,
            }
        )
        return base

    if pass2_status == "NUMERIC_CLUSTER_REVIEW":
        candidate_type = str(row["best_candidate_type"] or "")
        form = "PAIR" if candidate_type == "UNLABELED_PAIR" else "TRIPLE"
        profile = base["category_profile"]
        code, name, priority, next_action = numeric_requeue_pattern(
            candidate_type,
            profile,
            unit_present(unit_status),
        )
        flag = quality_flag(values, str(row["unit_text"] or ""), unit_status)
        if flag not in {"BASIC_RANGE_PLAUSIBLE", "UNIT_REQUIRED"}:
            priority += 10
        base.update(
            {
                "observation_classified": 1,
                "pattern_classified": 0,
                "resolution_state": "VALUE_RECLASSIFICATION_REQUIRED",
                "requeue_required": 1,
                "requeue_pattern_code": code,
                "requeue_pattern_name": name,
                "value_form": form,
                "quality_flag": flag,
                "priority": priority,
                "next_action": next_action,
            }
        )
        return base

    code, name, priority, next_action = no_observation_pattern(
        pass2_status,
        str(row["work_group_code"] or ""),
    )
    base.update(
        {
            "observation_classified": 0,
            "pattern_classified": 0,
            "resolution_state": "NO_OBSERVATION_REQUEUE",
            "requeue_required": 1,
            "requeue_pattern_code": code,
            "requeue_pattern_name": name,
            "quality_flag": "NO_OBSERVATION",
            "priority": priority,
            "next_action": next_action,
        }
    )
    return base


def build() -> dict[str, Any]:
    classified_at = now_text()
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    create_schema(connection)
    scope = load_scope(connection)
    rows = [build_row(row, classified_at) for row in scope]
    columns = list(rows[0])
    placeholders = ",".join(f":{column}" for column in columns)
    with connection:
        connection.execute(f"UPDATE {TABLE} SET is_current=0 WHERE is_current=1")
        connection.execute(f"DELETE FROM {TABLE} WHERE snapshot_id=?", (SNAPSHOT_ID,))
        connection.executemany(
            f"INSERT INTO {TABLE} ({','.join(columns)}) VALUES ({placeholders})",
            rows,
        )

    resolution_counts = Counter(row["resolution_state"] for row in rows)
    pattern_counts = Counter(
        row["requeue_pattern_code"] for row in rows if row["requeue_required"]
    )
    quality_counts = Counter(
        row["quality_flag"] for row in rows if row["requeue_required"]
    )
    summary = {
        "snapshot_id": SNAPSHOT_ID,
        "classified_at": classified_at,
        "taxonomy_version": TAXONOMY_VERSION,
        "scope_products": len(rows),
        "observation_classified_products": sum(row["observation_classified"] for row in rows),
        "pattern_classified_products": sum(row["pattern_classified"] for row in rows),
        "requeue_products": sum(row["requeue_required"] for row in rows),
        "resolution_state": dict(resolution_counts),
        "requeue_pattern": dict(pattern_counts),
        "requeue_quality_flag": dict(quality_counts),
        "table": TABLE,
        "views": [
            "vw_dimension_pass2_observed_classified_current",
            "vw_dimension_pass2_pattern_classified_current",
            "vw_dimension_pass2_value_reclassification_current",
            "vw_dimension_pass2_unclassified_current",
            "vw_dimension_pass2_requeue_current",
            "vw_dimension_pass2_requeue_summary",
        ],
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
