from __future__ import annotations

import json
import re
import sqlite3
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

from bulk_homestyle_collect import DB_PATH
from low_dimension_quality_policy import assess_low_dimension


LEDGER_VERSION = "dimension-resolution-ledger-v2-comparison"
LOCKED_STATUSES = {"SOURCE_CONFIRMED", "RULE_RESOLVED", "MANUAL_CONFIRMED"}

RULES = [
    (
        "SOURCE_DB_CONFIRMED",
        "SOURCE",
        "기존 DB 확정 규격",
        "ACTIVE_CONFIRMED",
        "기존 dimension_value_confirmed=1",
        "원본 확정값을 최우선 잠금",
    ),
    (
        "EXPLICIT_WDH",
        "NORMALIZE",
        "명시적 W/D/H",
        "ACTIVE_CONFIRMED",
        "W,D,H->W,D,H",
        "제품 규격 구역과 단위가 확인된 명시적 축",
    ),
    (
        "ORDERED_TRIPLE",
        "NORMALIZE",
        "가구 순서형 3개 값",
        "ACTIVE_CONFIRMED",
        "ORDERED_TRIPLE->W,D,H",
        "제품 규격 구역 및 가구 카테고리 문맥",
    ),
    (
        "LDH_TO_WDH",
        "NORMALIZE",
        "L/D/H를 W/D/H로 변환",
        "ACTIVE_CONFIRMED",
        "L,D,H->W,D,H",
        "가구 제품 규격 구역에서 L을 대표 가로로 해석",
    ),
    (
        "WLH_TO_WDH",
        "NORMALIZE",
        "W/L/H를 W/D/H로 변환",
        "ACTIVE_CONFIRMED",
        "W,L,H->W,D,H",
        "가구 제품 규격 구역에서 L을 깊이로 해석",
    ),
    (
        "LWH_TO_WDH",
        "NORMALIZE",
        "L/W/H를 W/D/H로 변환",
        "ACTIVE_CONFIRMED",
        "L,W,H->W,D,H",
        "Length/Width/Height 순서형 표기",
    ),
    (
        "ROUND_DIAMETER",
        "NORMALIZE",
        "원형 지름 정규화",
        "ACTIVE_CONFIRMED",
        "DIAMETER,H 또는 D,H",
        "shape_type=ROUND와 diameter_mm 보존",
    ),
    (
        "AREA_2D",
        "NORMALIZE",
        "2D 제품 규격",
        "ACTIVE_CONFIRMED",
        "2D_PAIR->W,H;D=N/A",
        "액자·아트웍 등은 D 비적용",
    ),
    (
        "TITLE_MODEL_MATCH",
        "SELECT",
        "상품명 숫자·모델코드 일치",
        "ACTIVE_CONFIRMED",
        "AUTO_ACCEPT_TITLE_MATCH",
        "라인업 복수 규격 중 현재 상품과 일치하는 값 우선",
    ),
    (
        "MULTI_OPTION_LABEL",
        "SELECT",
        "옵션명 기반 복수 규격",
        "ACTIVE_CONFIRMED",
        "MULTI_OPTION_CANDIDATE",
        "행 기반으로 모든 옵션을 잠금 저장",
    ),
    (
        "COMPARISON_CANDIDATES_EXPOSED",
        "PRESENT",
        "미확정 복수 후보 비교정보 제공",
        "ACTIVE_POLICY",
        "HUMAN_REVIEW|AMBIGUOUS_MULTI_CANDIDATE",
        "대표 후보와 비교 후보를 엑셀에 쉼표 구분으로 제공하고 사람 확인 대기열에서는 제외",
    ),
    (
        "EXCLUDE_DELIVERY",
        "EXCLUDE",
        "배송·반입 규격 제외",
        "ACTIVE_CONFIRMED",
        "DELIVERY_CLEARANCE",
        "엘리베이터·출입문·사다리차 값 제외",
    ),
    (
        "EXCLUDE_COMPONENT",
        "EXCLUDE",
        "구성품 규격 제외",
        "ACTIVE_CONFIRMED",
        "COMPONENT_DIMENSION",
        "가드·발통·헤드보드 등 대표 규격 제외",
    ),
    (
        "EXCLUDE_LINEUP",
        "EXCLUDE",
        "다른 라인업 규격 제외",
        "ACTIVE_CONFIRMED",
        "LINEUP_OTHER_MODEL",
        "후보 주변 모델명과 현재 상품명 비교",
    ),
    (
        "EXCLUDE_TOLERANCE",
        "EXCLUDE",
        "허용오차 제외",
        "ACTIVE_CONFIRMED",
        "MEASUREMENT_TOLERANCE",
        "± 값과 오차범위를 대표 규격에서 제외",
    ),
    (
        "OCR_REPAIR_AXIS_VALUE",
        "OCR_REPAIR",
        "OCR 축·숫자 결합 보정",
        "ACTIVE_CONFIRMED",
        "H41 O/WI 165/세로기054 등",
        "축과 단위 주변에 한정된 문자 치환",
    ),
]


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def init_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS ref_dimension_resolution_rule (
            rule_code TEXT PRIMARY KEY,
            rule_group TEXT NOT NULL,
            rule_name TEXT NOT NULL,
            rule_status TEXT NOT NULL,
            applies_to TEXT,
            notes TEXT,
            ledger_version TEXT NOT NULL,
            confirmed_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS fact_dimension_resolution_ledger (
            product_id TEXT PRIMARY KEY,
            product_name TEXT,
            mid_category TEXT,
            small_category TEXT,
            resolution_status TEXT NOT NULL,
            is_locked INTEGER NOT NULL,
            pass_status TEXT NOT NULL,
            needs_human_review INTEGER NOT NULL,
            needs_ocr INTEGER NOT NULL,
            locked_w_mm REAL,
            locked_d_mm REAL,
            locked_h_mm REAL,
            locked_d_applicability TEXT,
            locked_shape_type TEXT,
            locked_option_count INTEGER NOT NULL DEFAULT 0,
            candidate_w_mm REAL,
            candidate_d_mm REAL,
            candidate_h_mm REAL,
            context_status TEXT,
            resolution_rule_code TEXT,
            resolution_source TEXT,
            source_snapshot_id TEXT,
            representative_candidate_key TEXT,
            evidence_text TEXT,
            first_resolved_at TEXT,
            last_transition_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            ledger_version TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS fact_dimension_resolution_option (
            product_id TEXT NOT NULL,
            option_no INTEGER NOT NULL,
            option_label TEXT,
            is_primary INTEGER NOT NULL,
            w_mm REAL,
            d_mm REAL,
            h_mm REAL,
            d_applicability TEXT,
            shape_type TEXT,
            diameter_mm REAL,
            normalized_axis_mapping TEXT,
            resolution_rule_code TEXT,
            source_candidate_key TEXT,
            source_type TEXT,
            evidence_text TEXT,
            locked_at TEXT NOT NULL,
            ledger_version TEXT NOT NULL,
            PRIMARY KEY(product_id, option_no)
        );

        CREATE TABLE IF NOT EXISTS fact_dimension_comparison_candidate (
            product_id TEXT NOT NULL,
            candidate_no INTEGER NOT NULL,
            is_representative INTEGER NOT NULL DEFAULT 0,
            comparison_target TEXT,
            w_mm REAL,
            d_mm REAL,
            h_mm REAL,
            shape_type TEXT,
            normalized_axis_mapping TEXT,
            source_type TEXT,
            raw_notation TEXT,
            context_text TEXT,
            candidate_key TEXT,
            source_snapshot_id TEXT,
            captured_at TEXT NOT NULL,
            ledger_version TEXT NOT NULL,
            PRIMARY KEY(product_id, candidate_no)
        );

        CREATE TABLE IF NOT EXISTS hist_dimension_resolution_event (
            event_id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_at TEXT NOT NULL,
            product_id TEXT NOT NULL,
            old_status TEXT,
            new_status TEXT NOT NULL,
            transition_reason TEXT,
            source_snapshot_id TEXT,
            old_w_mm REAL,
            old_d_mm REAL,
            old_h_mm REAL,
            new_w_mm REAL,
            new_d_mm REAL,
            new_h_mm REAL,
            ledger_version TEXT NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_dimension_resolution_status
            ON fact_dimension_resolution_ledger(resolution_status);
        CREATE INDEX IF NOT EXISTS idx_dimension_resolution_work
            ON fact_dimension_resolution_ledger(is_locked, needs_human_review, needs_ocr);
        CREATE INDEX IF NOT EXISTS idx_dimension_resolution_category
            ON fact_dimension_resolution_ledger(mid_category, small_category);

        DROP VIEW IF EXISTS vw_dimension_resolution_ledger_current;
        CREATE VIEW vw_dimension_resolution_ledger_current AS
        SELECT *
        FROM fact_dimension_resolution_ledger;

        DROP VIEW IF EXISTS vw_dimension_locked_current;
        CREATE VIEW vw_dimension_locked_current AS
        SELECT *
        FROM fact_dimension_resolution_ledger
        WHERE is_locked = 1
        ORDER BY product_id;

        DROP VIEW IF EXISTS vw_dimension_remaining_human_review_current;
        CREATE VIEW vw_dimension_remaining_human_review_current AS
        SELECT *
        FROM fact_dimension_resolution_ledger
        WHERE needs_human_review = 1
          AND is_locked = 0
        ORDER BY mid_category, small_category, product_id;

        DROP VIEW IF EXISTS vw_dimension_comparison_provided_current;
        CREATE VIEW vw_dimension_comparison_provided_current AS
        SELECT *
        FROM fact_dimension_resolution_ledger
        WHERE resolution_status = 'COMPARISON_PROVIDED'
        ORDER BY mid_category, small_category, product_id;

        DROP VIEW IF EXISTS vw_dimension_context_comparison_candidates_eligible;
        CREATE VIEW vw_dimension_context_comparison_candidates_eligible AS
        WITH eligible AS (
            SELECT
                l.product_id,
                l.product_name,
                l.mid_category,
                l.small_category,
                c.candidate_key,
                c.option_label,
                c.source_axis_signature,
                c.section_role,
                c.source_type,
                c.raw_notation,
                c.context_text,
                c.w_mm,
                c.d_mm,
                c.h_mm,
                c.shape_type,
                c.normalized_axis_mapping,
                c.candidate_score,
                CASE
                    WHEN c.candidate_key = l.representative_candidate_key THEN 1
                    ELSE 0
                END AS is_representative,
                ROW_NUMBER() OVER (
                    PARTITION BY
                        c.product_id,
                        COALESCE(ROUND(c.w_mm, 3), -1),
                        COALESCE(ROUND(c.d_mm, 3), -1),
                        COALESCE(ROUND(c.h_mm, 3), -1)
                    ORDER BY
                        CASE
                            WHEN c.candidate_key = l.representative_candidate_key THEN 1
                            ELSE 0
                        END DESC,
                        COALESCE(c.candidate_score, 0) DESC,
                        COALESCE(c.source_priority, 0) DESC,
                        COALESCE(c.source_order, 999999),
                        c.candidate_key
                ) AS duplicate_rank
            FROM fact_dimension_resolution_ledger l
            JOIN stg_dimension_context_candidate c
              ON c.product_id = l.product_id
             AND c.is_current = 1
            WHERE l.resolution_status = 'COMPARISON_PROVIDED'
              AND c.candidate_role = 'PRODUCT_DIMENSION'
              AND COALESCE(c.decision_status, '') != 'REJECT'
              AND (
                    (c.w_mm IS NOT NULL)
                  + (c.d_mm IS NOT NULL)
                  + (c.h_mm IS NOT NULL)
              ) >= 1
        ),
        deduplicated AS (
            SELECT *
            FROM eligible
            WHERE duplicate_rank = 1
        )
        SELECT
            product_id,
            ROW_NUMBER() OVER (
                PARTITION BY product_id
                ORDER BY
                    is_representative DESC,
                    COALESCE(candidate_score, 0) DESC,
                    COALESCE(w_mm, -1) DESC,
                    COALESCE(d_mm, -1) DESC,
                    COALESCE(h_mm, -1) DESC,
                    candidate_key
            ) AS comparison_no,
            product_name,
            mid_category,
            small_category,
            is_representative,
            COALESCE(
                NULLIF(option_label, ''),
                NULLIF(source_axis_signature, ''),
                NULLIF(section_role, ''),
                '규격 후보'
            ) AS comparison_target,
            w_mm,
            d_mm,
            h_mm,
            shape_type,
            normalized_axis_mapping,
            source_type,
            raw_notation,
            context_text,
            candidate_key
        FROM deduplicated;

        DROP VIEW IF EXISTS vw_dimension_comparison_candidates_current;
        CREATE VIEW vw_dimension_comparison_candidates_current AS
        SELECT
            f.product_id,
            f.candidate_no AS comparison_no,
            l.product_name,
            l.mid_category,
            l.small_category,
            f.is_representative,
            f.comparison_target,
            f.w_mm,
            f.d_mm,
            f.h_mm,
            f.shape_type,
            f.normalized_axis_mapping,
            f.source_type,
            f.raw_notation,
            f.context_text,
            f.candidate_key
        FROM fact_dimension_comparison_candidate f
        JOIN fact_dimension_resolution_ledger l
          ON l.product_id = f.product_id
        WHERE l.resolution_status = 'COMPARISON_PROVIDED';

        DROP VIEW IF EXISTS vw_dimension_remaining_ocr_current;
        CREATE VIEW vw_dimension_remaining_ocr_current AS
        SELECT *
        FROM fact_dimension_resolution_ledger
        WHERE needs_ocr = 1
          AND is_locked = 0
        ORDER BY
            CASE resolution_status
                WHEN 'OCR_REQUIRED' THEN 1
                WHEN 'NO_CANDIDATE' THEN 2
                ELSE 3
            END,
            mid_category, small_category, product_id;

        DROP VIEW IF EXISTS vw_dimension_work_queue_current;
        CREATE VIEW vw_dimension_work_queue_current AS
        SELECT
            product_id, product_name, mid_category, small_category,
            resolution_status,
            CASE
                WHEN needs_human_review = 1 THEN 'HUMAN_REVIEW'
                WHEN resolution_status = 'OCR_REQUIRED' THEN 'TARGETED_REOCR'
                WHEN resolution_status = 'NO_CANDIDATE' THEN 'IMAGE_SCAN_AND_OCR'
                ELSE 'SOURCE_DISCOVERY'
            END AS work_type,
            candidate_w_mm, candidate_d_mm, candidate_h_mm,
            context_status, evidence_text
        FROM fact_dimension_resolution_ledger
        WHERE is_locked = 0
          AND (needs_human_review = 1 OR needs_ocr = 1)
        ORDER BY
            CASE
                WHEN needs_human_review = 1 THEN 1
                WHEN resolution_status = 'OCR_REQUIRED' THEN 2
                WHEN resolution_status = 'NO_CANDIDATE' THEN 3
                ELSE 4
            END,
            product_id;

        DROP VIEW IF EXISTS vw_dimension_progress_authoritative;
        CREATE VIEW vw_dimension_progress_authoritative AS
        SELECT
            COUNT(*) AS total_products,
            SUM(CASE WHEN is_locked = 1 THEN 1 ELSE 0 END) AS locked_resolved,
            SUM(CASE WHEN resolution_status = 'SOURCE_CONFIRMED' THEN 1 ELSE 0 END)
                AS source_confirmed,
            SUM(CASE WHEN resolution_status = 'RULE_RESOLVED' THEN 1 ELSE 0 END)
                AS rule_resolved,
            SUM(CASE WHEN resolution_status = 'MANUAL_CONFIRMED' THEN 1 ELSE 0 END)
                AS manual_confirmed,
            SUM(CASE WHEN resolution_status = 'COMPARISON_PROVIDED' THEN 1 ELSE 0 END)
                AS comparison_provided,
            SUM(CASE WHEN resolution_status = 'REVIEW_REQUIRED' THEN 1 ELSE 0 END)
                AS human_review_required,
            SUM(CASE WHEN resolution_status = 'OCR_REQUIRED' THEN 1 ELSE 0 END)
                AS targeted_reocr_required,
            SUM(CASE WHEN resolution_status = 'NO_CANDIDATE' THEN 1 ELSE 0 END)
                AS no_candidate,
            SUM(CASE WHEN resolution_status = 'UNCLASSIFIED' THEN 1 ELSE 0 END)
                AS unclassified,
            SUM(
                CASE WHEN needs_human_review = 1 OR needs_ocr = 1 THEN 1 ELSE 0 END
            ) AS total_remaining,
            SUM(CASE WHEN needs_ocr = 1 THEN 1 ELSE 0 END) AS ocr_pipeline_remaining
        FROM fact_dimension_resolution_ledger;

        DROP VIEW IF EXISTS vw_dimension_progress_by_status;
        CREATE VIEW vw_dimension_progress_by_status AS
        SELECT
            resolution_status,
            COUNT(*) AS product_count,
            SUM(is_locked) AS locked_count,
            SUM(needs_human_review) AS human_review_count,
            SUM(needs_ocr) AS ocr_count
        FROM fact_dimension_resolution_ledger
        GROUP BY resolution_status
        ORDER BY product_count DESC;
        """
    )


def upsert_rules(connection: sqlite3.Connection, timestamp: str) -> None:
    connection.executemany(
        """
        INSERT INTO ref_dimension_resolution_rule (
            rule_code, rule_group, rule_name, rule_status,
            applies_to, notes, ledger_version, confirmed_at
        ) VALUES (?,?,?,?,?,?,?,?)
        ON CONFLICT(rule_code) DO UPDATE SET
            rule_group=excluded.rule_group,
            rule_name=excluded.rule_name,
            rule_status=excluded.rule_status,
            applies_to=excluded.applies_to,
            notes=excluded.notes,
            ledger_version=excluded.ledger_version,
            confirmed_at=excluded.confirmed_at
        """,
        [(*rule, LEDGER_VERSION, timestamp) for rule in RULES],
    )


def rule_code_for(context_status: str, mapping: str, shape: str) -> str:
    if context_status == "AUTO_ACCEPT_TITLE_MATCH":
        return "TITLE_MODEL_MATCH"
    if context_status == "MULTI_OPTION_CANDIDATE":
        return "MULTI_OPTION_LABEL"
    if shape == "AREA_2D" or mapping == "2D_PAIR->W,H;D=N/A":
        return "AREA_2D"
    if shape == "ROUND" or mapping in {
        "DIAMETER,H->W,D,H",
        "D,H(DIAMETER)->W,D,H",
        "ROUND_PAIR->W,D,H",
    }:
        return "ROUND_DIAMETER"
    return {
        "W,D,H->W,D,H": "EXPLICIT_WDH",
        "ORDERED_TRIPLE->W,D,H": "ORDERED_TRIPLE",
        "L,D,H->W,D,H": "LDH_TO_WDH",
        "W,L,H->W,D,H": "WLH_TO_WDH",
        "L,W,H->W,D,H": "LWH_TO_WDH",
    }.get(mapping, "EXPLICIT_WDH")


def same_value(left: Any, right: Any) -> bool:
    if left is None and right is None:
        return True
    if left is None or right is None:
        return False
    return abs(float(left) - float(right)) < 0.001


def locked_rule_invalidation_reason(old: sqlite3.Row) -> str:
    """Return a narrow, explicit reason for invalidating an old rule lock."""

    locked_values = (
        old["locked_w_mm"],
        old["locked_d_mm"],
        old["locked_h_mm"],
    )
    if any(
        value is not None and float(value) <= 0
        for value in locked_values
    ):
        return "INVALIDATE_NONPOSITIVE_DIMENSION"
    low_value_assessment = assess_low_dimension(
        str(old["product_name"] or ""),
        str(old["mid_category"] or ""),
        str(old["small_category"] or ""),
        old["locked_w_mm"],
        old["locked_d_mm"],
        old["locked_h_mm"],
    )
    if low_value_assessment.requires_review:
        return "INVALIDATE_LOW_DIMENSION_VALUE"

    category = str(old["small_category"] or "")
    if "침대" not in category and "매트리스" not in category:
        return ""
    depth = old["locked_d_mm"]
    height = old["locked_h_mm"]
    if (
        depth is not None
        and height is not None
        and float(depth) < 1000
        and float(height) > 1000
    ):
        return "INVALIDATE_BED_AXIS_ORDER"

    product_name = str(old["product_name"] or "")
    evidence = str(old["evidence_text"] or "")
    if re.search(r"(?:매트리스|mattress)", product_name, re.I):
        return ""
    if not re.search(
        r"(?:HYBRID\s+TECH|원매트리스|투매트리스|mattress\s+type)",
        evidence,
        re.I,
    ):
        return ""
    if height is None or float(height) > 500:
        return ""
    ignored = {
        "ACE",
        "BED",
        "FRAME",
        "MATTRESS",
        "HYBRID",
        "TECH",
        "SIZE",
        "INFO",
    }
    product_tokens = {
        token.upper()
        for token in re.findall(r"[A-Za-z][A-Za-z0-9-]{3,}", product_name)
        if token.upper() not in ignored
    }
    evidence_tokens = {
        token.upper()
        for token in re.findall(r"[A-Za-z][A-Za-z0-9-]{3,}", evidence)
        if token.upper() not in ignored
    }
    if not (product_tokens & evidence_tokens):
        return "INVALIDATE_GENERIC_MATTRESS_COMPONENT"
    return ""


def preserve_new_comparison_candidates(
    connection: sqlite3.Connection, timestamp: str
) -> int:
    before = connection.total_changes
    connection.execute(
        """
        INSERT INTO fact_dimension_comparison_candidate (
            product_id, candidate_no, is_representative, comparison_target,
            w_mm, d_mm, h_mm, shape_type, normalized_axis_mapping,
            source_type, raw_notation, context_text, candidate_key,
            source_snapshot_id, captured_at, ledger_version
        )
        SELECT
            e.product_id,
            e.comparison_no,
            e.is_representative,
            e.comparison_target,
            e.w_mm,
            e.d_mm,
            e.h_mm,
            e.shape_type,
            e.normalized_axis_mapping,
            e.source_type,
            e.raw_notation,
            e.context_text,
            e.candidate_key,
            l.source_snapshot_id,
            ?,
            ?
        FROM vw_dimension_context_comparison_candidates_eligible e
        JOIN fact_dimension_resolution_ledger l
          ON l.product_id = e.product_id
        WHERE NOT EXISTS (
            SELECT 1
            FROM fact_dimension_comparison_candidate old
            WHERE old.product_id = e.product_id
        )
        ORDER BY e.product_id, e.comparison_no
        """,
        (timestamp, LEDGER_VERSION),
    )
    return connection.total_changes - before


def main() -> None:
    timestamp = now_text()
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA journal_mode=WAL")
    connection.execute("PRAGMA synchronous=NORMAL")
    init_schema(connection)
    upsert_rules(connection, timestamp)

    masters = list(
        connection.execute(
            "SELECT * FROM vw_dimension_classification_master_current ORDER BY product_id"
        )
    )
    contexts = {
        row["product_id"]: row
        for row in connection.execute(
            "SELECT * FROM vw_dimension_context_products_current"
        )
    }
    candidates = {
        row["candidate_key"]: row
        for row in connection.execute(
            "SELECT * FROM vw_dimension_context_candidates_current"
        )
    }
    options: dict[str, list[sqlite3.Row]] = defaultdict(list)
    for row in connection.execute(
        "SELECT * FROM vw_product_dimension_options_current ORDER BY product_id, option_no"
    ):
        options[row["product_id"]].append(row)
    existing = {
        row["product_id"]: row
        for row in connection.execute(
            "SELECT * FROM fact_dimension_resolution_ledger"
        )
    }

    status_counts: dict[str, int] = defaultdict(int)
    changed_products: list[str] = []

    for master in masters:
        product_id = master["product_id"]
        context = contexts.get(product_id)
        representative = (
            candidates.get(context["representative_candidate_key"])
            if context and context["representative_candidate_key"]
            else None
        )

        desired: dict[str, Any]
        confirmed_values = (
            master["confirmed_w_mm"],
            master["confirmed_d_mm"],
            master["confirmed_h_mm"],
        )
        confirmed_values_physically_valid = all(
            value is not None and float(value) > 0
            for value in confirmed_values
        )
        confirmed_low_value_assessment = assess_low_dimension(
            str(master["product_name"] or ""),
            str(master["mid_category"] or ""),
            str(master["small_category"] or ""),
            master["confirmed_w_mm"],
            master["confirmed_d_mm"],
            master["confirmed_h_mm"],
        )
        if (
            int(master["dimension_value_confirmed"] or 0) == 1
            and confirmed_values_physically_valid
            and not confirmed_low_value_assessment.requires_review
        ):
            desired = {
                "resolution_status": "SOURCE_CONFIRMED",
                "is_locked": 1,
                "pass_status": "PASS",
                "needs_human_review": 0,
                "needs_ocr": 0,
                "locked_w_mm": master["confirmed_w_mm"],
                "locked_d_mm": master["confirmed_d_mm"],
                "locked_h_mm": master["confirmed_h_mm"],
                "locked_d_applicability": "APPLICABLE",
                "locked_shape_type": "",
                "locked_option_count": 1,
                "candidate_w_mm": None,
                "candidate_d_mm": None,
                "candidate_h_mm": None,
                "context_status": "EXISTING_CONFIRMED",
                "resolution_rule_code": "SOURCE_DB_CONFIRMED",
                "resolution_source": "classification_master.confirmed",
                "source_snapshot_id": master["snapshot_id"],
                "representative_candidate_key": "",
                "evidence_text": master["evidence_text"],
            }
        elif context:
            context_status = context["context_status"]
            if context_status in {
                "AUTO_ACCEPT_CANDIDATE",
                "AUTO_ACCEPT_TITLE_MATCH",
                "MULTI_OPTION_CANDIDATE",
            }:
                mapping = (
                    representative["normalized_axis_mapping"]
                    if representative
                    else ""
                )
                shape = context["representative_shape_type"] or ""
                desired_status = "RULE_RESOLVED"
                desired = {
                    "resolution_status": desired_status,
                    "is_locked": 1,
                    "pass_status": "PASS",
                    "needs_human_review": 0,
                    "needs_ocr": 0,
                    "locked_w_mm": context["representative_w_mm"],
                    "locked_d_mm": context["representative_d_mm"],
                    "locked_h_mm": context["representative_h_mm"],
                    "locked_d_applicability": context[
                        "representative_d_applicability"
                    ],
                    "locked_shape_type": shape,
                    "locked_option_count": context["option_count"],
                    "candidate_w_mm": None,
                    "candidate_d_mm": None,
                    "candidate_h_mm": None,
                    "context_status": context_status,
                    "resolution_rule_code": rule_code_for(
                        context_status, mapping, shape
                    ),
                    "resolution_source": "dimension_context_rule_engine",
                    "source_snapshot_id": context["snapshot_id"],
                    "representative_candidate_key": context[
                        "representative_candidate_key"
                    ],
                    "evidence_text": (
                        representative["context_text"] if representative else ""
                    ),
                }
            elif context_status in {
                "HUMAN_REVIEW",
                "AMBIGUOUS_MULTI_CANDIDATE",
            }:
                desired = {
                    "resolution_status": "COMPARISON_PROVIDED",
                    "is_locked": 0,
                    "pass_status": "COMPARISON_READY",
                    "needs_human_review": 0,
                    "needs_ocr": 0,
                    "locked_w_mm": None,
                    "locked_d_mm": None,
                    "locked_h_mm": None,
                    "locked_d_applicability": "",
                    "locked_shape_type": "",
                    "locked_option_count": 0,
                    "candidate_w_mm": context["representative_w_mm"],
                    "candidate_d_mm": context["representative_d_mm"],
                    "candidate_h_mm": context["representative_h_mm"],
                    "context_status": context_status,
                    "resolution_rule_code": "COMPARISON_CANDIDATES_EXPOSED",
                    "resolution_source": "dimension_context_comparison",
                    "source_snapshot_id": context["snapshot_id"],
                    "representative_candidate_key": context[
                        "representative_candidate_key"
                    ],
                    "evidence_text": (
                        representative["context_text"] if representative else ""
                    ),
                }
            elif context_status == "REOCR_REQUIRED":
                desired = {
                    "resolution_status": "OCR_REQUIRED",
                    "is_locked": 0,
                    "pass_status": "REMAINING",
                    "needs_human_review": 0,
                    "needs_ocr": 1,
                    "locked_w_mm": None,
                    "locked_d_mm": None,
                    "locked_h_mm": None,
                    "locked_d_applicability": "",
                    "locked_shape_type": "",
                    "locked_option_count": 0,
                    "candidate_w_mm": master["candidate_w_mm"],
                    "candidate_d_mm": master["candidate_d_mm"],
                    "candidate_h_mm": master["candidate_h_mm"],
                    "context_status": context_status,
                    "resolution_rule_code": "",
                    "resolution_source": "targeted_reocr_queue",
                    "source_snapshot_id": context["snapshot_id"],
                    "representative_candidate_key": "",
                    "evidence_text": master["evidence_text"],
                }
            else:
                desired = {
                    "resolution_status": "NO_CANDIDATE",
                    "is_locked": 0,
                    "pass_status": "REMAINING",
                    "needs_human_review": 0,
                    "needs_ocr": 1,
                    "locked_w_mm": None,
                    "locked_d_mm": None,
                    "locked_h_mm": None,
                    "locked_d_applicability": "",
                    "locked_shape_type": "",
                    "locked_option_count": 0,
                    "candidate_w_mm": master["candidate_w_mm"],
                    "candidate_d_mm": master["candidate_d_mm"],
                    "candidate_h_mm": master["candidate_h_mm"],
                    "context_status": context_status,
                    "resolution_rule_code": "",
                    "resolution_source": "image_scan_queue",
                    "source_snapshot_id": context["snapshot_id"],
                    "representative_candidate_key": "",
                    "evidence_text": master["evidence_text"],
                }
        else:
            desired = {
                "resolution_status": "UNCLASSIFIED",
                "is_locked": 0,
                "pass_status": "REMAINING",
                "needs_human_review": 0,
                "needs_ocr": 1,
                "locked_w_mm": None,
                "locked_d_mm": None,
                "locked_h_mm": None,
                "locked_d_applicability": "",
                "locked_shape_type": "",
                "locked_option_count": 0,
                "candidate_w_mm": master["candidate_w_mm"],
                "candidate_d_mm": master["candidate_d_mm"],
                "candidate_h_mm": master["candidate_h_mm"],
                "context_status": master["classification_status"],
                "resolution_rule_code": "",
                "resolution_source": "source_discovery_queue",
                "source_snapshot_id": master["snapshot_id"],
                "representative_candidate_key": "",
                "evidence_text": master["evidence_text"],
            }

        old = existing.get(product_id)
        preserve_locked = False
        invalidation_reason = (
            locked_rule_invalidation_reason(old)
            if old and old["resolution_status"] == "RULE_RESOLVED"
            else ""
        )
        if old and old["resolution_status"] == "MANUAL_CONFIRMED":
            desired = {
                key: old[key]
                for key in desired
            }
            preserve_locked = True
        elif (
            old
            and old["resolution_status"] == "COMPARISON_PROVIDED"
            and desired["resolution_status"] not in {
                "SOURCE_CONFIRMED",
                "MANUAL_CONFIRMED",
            }
        ):
            # Providing all candidates is an accepted terminal policy for this
            # group. Later diagnostic snapshots cannot put it back into a
            # human-review/OCR queue unless the policy is explicitly changed.
            desired = {
                key: old[key]
                for key in desired
            }
        elif (
            old
            and int(old["is_locked"] or 0) == 1
            and old["resolution_status"] == "RULE_RESOLVED"
            and not invalidation_reason
            and desired["resolution_status"] not in {
                "SOURCE_CONFIRMED",
                "MANUAL_CONFIRMED",
            }
        ):
            # A locked rule result is cumulative.  A later diagnostic snapshot
            # cannot silently downgrade it; explicit invalidation is required.
            desired = {
                key: old[key]
                for key in desired
            }
            preserve_locked = True

        first_resolved_at = (
            old["first_resolved_at"]
            if old and old["first_resolved_at"]
            else timestamp
            if desired["is_locked"]
            else None
        )
        transition = (
            old is None
            or old["resolution_status"] != desired["resolution_status"]
            or not same_value(old["locked_w_mm"], desired["locked_w_mm"])
            or not same_value(old["locked_d_mm"], desired["locked_d_mm"])
            or not same_value(old["locked_h_mm"], desired["locked_h_mm"])
        )
        last_transition_at = (
            timestamp if transition else old["last_transition_at"]
        )

        connection.execute(
            """
            INSERT INTO fact_dimension_resolution_ledger (
                product_id, product_name, mid_category, small_category,
                resolution_status, is_locked, pass_status,
                needs_human_review, needs_ocr,
                locked_w_mm, locked_d_mm, locked_h_mm,
                locked_d_applicability, locked_shape_type, locked_option_count,
                candidate_w_mm, candidate_d_mm, candidate_h_mm,
                context_status, resolution_rule_code, resolution_source,
                source_snapshot_id, representative_candidate_key, evidence_text,
                first_resolved_at, last_transition_at, updated_at, ledger_version
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(product_id) DO UPDATE SET
                product_name=excluded.product_name,
                mid_category=excluded.mid_category,
                small_category=excluded.small_category,
                resolution_status=excluded.resolution_status,
                is_locked=excluded.is_locked,
                pass_status=excluded.pass_status,
                needs_human_review=excluded.needs_human_review,
                needs_ocr=excluded.needs_ocr,
                locked_w_mm=excluded.locked_w_mm,
                locked_d_mm=excluded.locked_d_mm,
                locked_h_mm=excluded.locked_h_mm,
                locked_d_applicability=excluded.locked_d_applicability,
                locked_shape_type=excluded.locked_shape_type,
                locked_option_count=excluded.locked_option_count,
                candidate_w_mm=excluded.candidate_w_mm,
                candidate_d_mm=excluded.candidate_d_mm,
                candidate_h_mm=excluded.candidate_h_mm,
                context_status=excluded.context_status,
                resolution_rule_code=excluded.resolution_rule_code,
                resolution_source=excluded.resolution_source,
                source_snapshot_id=excluded.source_snapshot_id,
                representative_candidate_key=excluded.representative_candidate_key,
                evidence_text=excluded.evidence_text,
                first_resolved_at=excluded.first_resolved_at,
                last_transition_at=excluded.last_transition_at,
                updated_at=excluded.updated_at,
                ledger_version=excluded.ledger_version
            """,
            (
                product_id,
                master["product_name"],
                master["mid_category"],
                master["small_category"],
                desired["resolution_status"],
                desired["is_locked"],
                desired["pass_status"],
                desired["needs_human_review"],
                desired["needs_ocr"],
                desired["locked_w_mm"],
                desired["locked_d_mm"],
                desired["locked_h_mm"],
                desired["locked_d_applicability"],
                desired["locked_shape_type"],
                desired["locked_option_count"],
                desired["candidate_w_mm"],
                desired["candidate_d_mm"],
                desired["candidate_h_mm"],
                desired["context_status"],
                desired["resolution_rule_code"],
                desired["resolution_source"],
                desired["source_snapshot_id"],
                desired["representative_candidate_key"],
                desired["evidence_text"],
                first_resolved_at,
                last_transition_at,
                timestamp,
                LEDGER_VERSION,
            ),
        )
        if transition:
            changed_products.append(product_id)
            connection.execute(
                """
                INSERT INTO hist_dimension_resolution_event (
                    event_at, product_id, old_status, new_status,
                    transition_reason, source_snapshot_id,
                    old_w_mm, old_d_mm, old_h_mm,
                    new_w_mm, new_d_mm, new_h_mm, ledger_version
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
                """,
                (
                    timestamp,
                    product_id,
                    old["resolution_status"] if old else None,
                    desired["resolution_status"],
                    "INITIAL_LOAD"
                    if old is None
                    else invalidation_reason
                    or "STATUS_OR_VALUE_CHANGED",
                    desired["source_snapshot_id"],
                    old["locked_w_mm"] if old else None,
                    old["locked_d_mm"] if old else None,
                    old["locked_h_mm"] if old else None,
                    desired["locked_w_mm"],
                    desired["locked_d_mm"],
                    desired["locked_h_mm"],
                    LEDGER_VERSION,
                ),
            )
        status_counts[desired["resolution_status"]] += 1

        if (
            desired["is_locked"]
            and desired["resolution_status"] != "MANUAL_CONFIRMED"
            and not preserve_locked
        ):
            connection.execute(
                "DELETE FROM fact_dimension_resolution_option WHERE product_id=?",
                (product_id,),
            )
            if desired["resolution_status"] == "SOURCE_CONFIRMED":
                connection.execute(
                    """
                    INSERT INTO fact_dimension_resolution_option (
                        product_id, option_no, option_label, is_primary,
                        w_mm, d_mm, h_mm, d_applicability, shape_type,
                        diameter_mm, normalized_axis_mapping,
                        resolution_rule_code, source_candidate_key,
                        source_type, evidence_text, locked_at, ledger_version
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """,
                    (
                        product_id,
                        1,
                        "대표",
                        1,
                        desired["locked_w_mm"],
                        desired["locked_d_mm"],
                        desired["locked_h_mm"],
                        desired["locked_d_applicability"],
                        desired["locked_shape_type"],
                        None,
                        "SOURCE_CONFIRMED",
                        desired["resolution_rule_code"],
                        "",
                        "SOURCE_DB",
                        desired["evidence_text"],
                        first_resolved_at,
                        LEDGER_VERSION,
                    ),
                )
            else:
                for option in options.get(product_id, []):
                    option_rule = rule_code_for(
                        context["context_status"],
                        option["normalized_axis_mapping"] or "",
                        option["shape_type"] or "",
                    )
                    connection.execute(
                        """
                        INSERT INTO fact_dimension_resolution_option (
                            product_id, option_no, option_label, is_primary,
                            w_mm, d_mm, h_mm, d_applicability, shape_type,
                            diameter_mm, normalized_axis_mapping,
                            resolution_rule_code, source_candidate_key,
                            source_type, evidence_text, locked_at, ledger_version
                        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                        """,
                        (
                            product_id,
                            option["option_no"],
                            option["option_label"],
                            option["is_primary"],
                            option["w_mm"],
                            option["d_mm"],
                            option["h_mm"],
                            option["d_applicability"],
                            option["shape_type"],
                            option["diameter_mm"],
                            option["normalized_axis_mapping"],
                            option_rule,
                            option["candidate_key"],
                            option["source_type"],
                            option["evidence_text"],
                            first_resolved_at,
                            LEDGER_VERSION,
                        ),
                    )

    comparison_candidates_preserved = preserve_new_comparison_candidates(
        connection, timestamp
    )
    connection.commit()
    progress = dict(
        connection.execute(
            "SELECT * FROM vw_dimension_progress_authoritative"
        ).fetchone()
    )
    result = {
        "ledger_version": LEDGER_VERSION,
        "updated_at": timestamp,
        "transitions_written": len(changed_products),
        "status_counts": dict(sorted(status_counts.items())),
        "authoritative_progress": progress,
        "locked_options": connection.execute(
            "SELECT COUNT(*) FROM fact_dimension_resolution_option"
        ).fetchone()[0],
        "comparison_candidates_preserved": comparison_candidates_preserved,
        "comparison_candidates_total": connection.execute(
            "SELECT COUNT(*) FROM fact_dimension_comparison_candidate"
        ).fetchone()[0],
        "integrity_check": connection.execute("PRAGMA integrity_check").fetchone()[0],
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
