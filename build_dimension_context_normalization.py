from __future__ import annotations

import argparse
import hashlib
import json
import re
import sqlite3
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from bulk_homestyle_collect import unpack
from dimension_context_normalizer import (
    category_profile,
    extract_candidates,
    title_option_codes,
)


ROOT = Path(__file__).resolve().parent
DEFAULT_DB = ROOT / "homestyle_bulk_run" / "homestyle_bulk.sqlite"
ENGINE_VERSION = "dimension-context-v1.0"
MAX_SEGMENT_CHARS = 80_000

SOURCE_PRIORITY = {
    "TARGETED_REGION_OCR": 60,
    "TARGETED_FULL_IMAGE_OCR": 50,
    "OCR_REVIEW_CANDIDATE": 35,
    "OCR_REVIEW_RAW_CONTEXT": 32,
    "MASTER_EVIDENCE": 30,
    "PASS2_FULL_IMAGE_OCR": 25,
    "SOURCE_SELECTED_OCR": 20,
}

USER_CASES = [
    {
        "product_id": "G25100020854",
        "case_code": "DELIVERY_EXCLUSION",
        "expected_role": "DELIVERY_CLEARANCE",
        "forbidden_w_mm": 900,
        "forbidden_d_mm": 1200,
        "forbidden_h_mm": 2300,
        "note": "엘리베이터 반입 규격은 대표 제품 규격에서 제외",
    },
    {
        "product_id": "G25080007062",
        "case_code": "DIMENSION_SECTION",
        "expected_role": "PRODUCT_DIMENSION",
        "expected_w_mm": 1800,
        "expected_d_mm": 500,
        "expected_h_mm": 580,
        "note": "DIMENSION 구역의 mm 표기를 선택",
    },
    {
        "product_id": "G25110025403",
        "case_code": "MULTI_OPTION",
        "expected_role": "PRODUCT_DIMENSION",
        "expected_status": "MULTI_OPTION_CANDIDATE",
        "note": "M/L 복수 규격을 옵션 행으로 분리",
    },
    {
        "product_id": "G25070004213",
        "case_code": "UNLABELED_FURNITURE_TRIPLE",
        "expected_role": "PRODUCT_DIMENSION",
        "expected_w_mm": 2310,
        "expected_d_mm": 990,
        "expected_h_mm": 650,
        "expected_mapping": "ORDERED_TRIPLE->W,D,H",
        "note": "소파 카테고리의 순서형 3개 값을 W/D/H로 정규화",
    },
    {
        "product_id": "G25070005671",
        "case_code": "LINEUP_OTHER_MODEL",
        "expected_role": "LINEUP_OTHER_MODEL",
        "forbidden_w_mm": 1480,
        "forbidden_d_mm": 520,
        "forbidden_h_mm": 330,
        "note": "Plato Arc Sofa Table 라인업 규격은 현재 상품 대표값에서 제외",
    },
    {
        "product_id": "G25090018797",
        "case_code": "PARTIAL_REOCR",
        "expected_role": "PRODUCT_DIMENSION",
        "expected_w_mm": 1400,
        "expected_d_mm": 700,
        "note": "SIZE 구역에서 높이를 표적 재OCR",
    },
    {
        "product_id": "G26020032822",
        "case_code": "PARTIAL_REOCR",
        "expected_role": "PRODUCT_DIMENSION",
        "expected_w_mm": 1100,
        "expected_d_mm": 900,
        "expected_h_mm": 300,
        "expected_mapping": "L,W,H->W,D,H",
        "note": "SIZE의 L1100/W900/H300을 대표 W/D/H로 축 정규화",
    },
    {
        "product_id": "G26020032829",
        "case_code": "ROUND_DIAMETER",
        "expected_role": "PRODUCT_DIMENSION",
        "expected_shape": "ROUND",
        "note": "원형 지름은 shape/diameter로 보존하고 W/D 숫자 경계값으로 정규화",
    },
    {
        "product_id": "G25090019222",
        "case_code": "MODEL_INFO_MATCH",
        "expected_role": "PRODUCT_DIMENSION",
        "expected_w_mm": 3220,
        "expected_d_mm": 2054,
        "expected_h_mm": 1100,
        "note": "상품 모델과 일치하는 INFO 규격을 선택",
    },
    {
        "product_id": "G25080010683",
        "case_code": "AREA_2D",
        "expected_role": "PRODUCT_DIMENSION",
        "expected_w_mm": 600,
        "expected_h_mm": 600,
        "expected_shape": "AREA_2D",
        "note": "2D 액자 카테고리는 D를 비적용 처리",
    },
    {
        "product_id": "G25080008909",
        "case_code": "WLH_MAPPING",
        "expected_role": "PRODUCT_DIMENSION",
        "expected_w_mm": 900,
        "expected_d_mm": 570,
        "expected_h_mm": 700,
        "expected_mapping": "W,L,H->W,D,H",
        "note": "가구 W/L/H의 L을 D로 문맥 정규화",
    },
    {
        "product_id": "G26030036251",
        "case_code": "LDH_MAPPING",
        "expected_role": "PRODUCT_DIMENSION",
        "expected_w_mm": 2050,
        "expected_d_mm": 900,
        "expected_h_mm": 750,
        "expected_mapping": "L,D,H->W,D,H",
        "note": "가구 L/D/H의 L을 W로 문맥 정규화",
    },
    {
        "product_id": "G25090019058",
        "case_code": "COMPONENT_EXCLUSION",
        "expected_role": "COMPONENT_DIMENSION",
        "forbidden_w_mm": 1120,
        "forbidden_d_mm": 62,
        "forbidden_h_mm": 890,
        "note": "침대 가드 구성품 규격은 대표값에서 제외",
    },
    {
        "product_id": "G25110024556",
        "case_code": "TITLE_NUMBER_MATCH",
        "expected_role": "PRODUCT_DIMENSION",
        "expected_w_mm": 1600,
        "expected_d_mm": 350,
        "expected_h_mm": 450,
        "note": "상품명 1600과 일치하는 규격 후보를 우선",
    },
    {
        "product_id": "G25070004913",
        "case_code": "DELIVERY_NUMBER_EXCLUSION",
        "expected_role": "DELIVERY_CLEARANCE",
        "forbidden_w_mm": 60,
        "note": "배송·사다리차 안내의 60mm를 대표값에서 제외",
    },
    {
        "product_id": "G25070004496",
        "case_code": "TOLERANCE_EXCLUSION",
        "expected_role": "MEASUREMENT_TOLERANCE",
        "note": "±3cm 허용 오차를 대표 규격에서 제외",
    },
    {
        "product_id": "G25070003892",
        "case_code": "OCR_DIGIT_JOIN",
        "expected_role": "PRODUCT_DIMENSION",
        "note": "DIMENSION의 cm를 우선하고 인증번호 결합 수치 11066은 재OCR",
    },
]


@dataclass(frozen=True)
class Segment:
    source_type: str
    source_ref: str
    source_order: int
    image_url: str
    text: str


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def clean_text(value: Any) -> str:
    return " ".join(str(value or "").replace("\x00", " ").split())


def close_enough(value: float | None, expected: float | None, tolerance: float = 1.0) -> bool:
    if expected is None:
        return True
    return value is not None and abs(value - expected) <= tolerance


def dimension_key(row: dict[str, Any]) -> tuple[Any, ...]:
    def rounded(value: Any) -> float | None:
        return None if value is None else round(float(value), 1)

    return (
        rounded(row.get("w_mm")),
        rounded(row.get("d_mm")),
        rounded(row.get("h_mm")),
        rounded(row.get("diameter_mm")),
        row.get("shape_type") or "",
    )


def init_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS ref_dimension_context_rule_testcase (
            product_id TEXT NOT NULL,
            case_code TEXT NOT NULL,
            expected_role TEXT,
            expected_status TEXT,
            expected_w_mm REAL,
            expected_d_mm REAL,
            expected_h_mm REAL,
            expected_shape TEXT,
            expected_mapping TEXT,
            forbidden_w_mm REAL,
            forbidden_d_mm REAL,
            forbidden_h_mm REAL,
            note TEXT,
            updated_at TEXT,
            PRIMARY KEY (product_id, case_code)
        );

        CREATE TABLE IF NOT EXISTS stg_dimension_context_candidate (
            snapshot_id TEXT NOT NULL,
            normalized_at TEXT NOT NULL,
            is_current INTEGER NOT NULL,
            engine_version TEXT NOT NULL,
            candidate_key TEXT NOT NULL,
            product_id TEXT NOT NULL,
            product_name TEXT,
            mid_category TEXT,
            small_category TEXT,
            category_profile TEXT,
            source_type TEXT,
            source_ref TEXT,
            source_order INTEGER,
            source_priority INTEGER,
            image_url TEXT,
            candidate_no INTEGER,
            rule_id TEXT,
            raw_notation TEXT,
            context_text TEXT,
            section_role TEXT,
            candidate_role TEXT,
            source_axis_signature TEXT,
            normalized_axis_mapping TEXT,
            option_label TEXT,
            shape_type TEXT,
            unit_status TEXT,
            unit_text TEXT,
            w_raw REAL,
            d_raw REAL,
            h_raw REAL,
            l_raw REAL,
            r_raw REAL,
            value_1_raw REAL,
            value_2_raw REAL,
            value_3_raw REAL,
            w_mm REAL,
            d_mm REAL,
            h_mm REAL,
            diameter_mm REAL,
            product_name_match_score INTEGER,
            candidate_score INTEGER,
            decision_status TEXT,
            rejection_reason TEXT,
            PRIMARY KEY (snapshot_id, candidate_key)
        );

        CREATE TABLE IF NOT EXISTS stg_product_dimension_option (
            snapshot_id TEXT NOT NULL,
            normalized_at TEXT NOT NULL,
            is_current INTEGER NOT NULL,
            engine_version TEXT NOT NULL,
            product_id TEXT NOT NULL,
            product_name TEXT,
            option_no INTEGER NOT NULL,
            is_primary INTEGER NOT NULL,
            option_label TEXT,
            selection_status TEXT,
            candidate_key TEXT,
            shape_type TEXT,
            diameter_mm REAL,
            w_mm REAL,
            d_mm REAL,
            h_mm REAL,
            d_applicability TEXT,
            source_axis_signature TEXT,
            normalized_axis_mapping TEXT,
            unit_status TEXT,
            candidate_score INTEGER,
            source_type TEXT,
            image_url TEXT,
            evidence_text TEXT,
            PRIMARY KEY (snapshot_id, product_id, option_no)
        );

        CREATE TABLE IF NOT EXISTS stg_dimension_context_product (
            snapshot_id TEXT NOT NULL,
            normalized_at TEXT NOT NULL,
            is_current INTEGER NOT NULL,
            engine_version TEXT NOT NULL,
            product_id TEXT NOT NULL,
            product_name TEXT,
            mid_category TEXT,
            small_category TEXT,
            category_profile TEXT,
            context_status TEXT,
            candidate_count INTEGER,
            rejected_candidate_count INTEGER,
            complete_candidate_count INTEGER,
            automatic_candidate_count INTEGER,
            option_count INTEGER,
            representative_candidate_key TEXT,
            representative_w_mm REAL,
            representative_d_mm REAL,
            representative_h_mm REAL,
            representative_shape_type TEXT,
            representative_d_applicability TEXT,
            requires_reocr INTEGER,
            requires_human_review INTEGER,
            next_action TEXT,
            PRIMARY KEY (snapshot_id, product_id)
        );

        CREATE TABLE IF NOT EXISTS stg_dimension_targeted_reocr_queue (
            snapshot_id TEXT NOT NULL,
            queued_at TEXT NOT NULL,
            is_current INTEGER NOT NULL,
            product_id TEXT NOT NULL,
            product_name TEXT,
            mid_category TEXT,
            small_category TEXT,
            queue_priority INTEGER,
            queue_reason TEXT,
            target_heading TEXT,
            target_strategy TEXT,
            best_image_url TEXT,
            existing_axis_status TEXT,
            existing_evidence_text TEXT,
            user_review_case INTEGER,
            PRIMARY KEY (snapshot_id, product_id)
        );

        CREATE TABLE IF NOT EXISTS stg_dimension_context_regression_result (
            snapshot_id TEXT NOT NULL,
            evaluated_at TEXT NOT NULL,
            is_current INTEGER NOT NULL,
            product_id TEXT NOT NULL,
            case_code TEXT NOT NULL,
            result_status TEXT,
            result_detail TEXT,
            matched_candidate_count INTEGER,
            forbidden_selected_count INTEGER,
            product_context_status TEXT,
            PRIMARY KEY (snapshot_id, product_id, case_code)
        );

        CREATE INDEX IF NOT EXISTS idx_dimension_context_candidate_current_product
            ON stg_dimension_context_candidate(is_current, product_id);
        CREATE INDEX IF NOT EXISTS idx_dimension_context_candidate_role
            ON stg_dimension_context_candidate(is_current, candidate_role, decision_status);
        CREATE INDEX IF NOT EXISTS idx_product_dimension_option_current_product
            ON stg_product_dimension_option(is_current, product_id);
        CREATE INDEX IF NOT EXISTS idx_dimension_context_product_current_status
            ON stg_dimension_context_product(is_current, context_status);
        CREATE INDEX IF NOT EXISTS idx_dimension_targeted_reocr_current
            ON stg_dimension_targeted_reocr_queue(is_current, queue_priority);
        """
    )


def create_views(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        DROP VIEW IF EXISTS vw_dimension_context_candidates_current;
        CREATE VIEW vw_dimension_context_candidates_current AS
        SELECT *
        FROM stg_dimension_context_candidate
        WHERE is_current = 1;

        DROP VIEW IF EXISTS vw_dimension_context_products_current;
        CREATE VIEW vw_dimension_context_products_current AS
        SELECT *
        FROM stg_dimension_context_product
        WHERE is_current = 1;

        DROP VIEW IF EXISTS vw_product_dimension_options_current;
        CREATE VIEW vw_product_dimension_options_current AS
        SELECT *
        FROM stg_product_dimension_option
        WHERE is_current = 1;

        DROP VIEW IF EXISTS vw_product_dimension_options_wide_current;
        CREATE VIEW vw_product_dimension_options_wide_current AS
        SELECT
            p.product_id,
            p.product_name,
            p.mid_category,
            p.small_category,
            p.context_status,
            p.option_count,
            MAX(CASE WHEN o.option_no = 1 THEN o.option_label END) AS representative_option_label,
            MAX(CASE WHEN o.option_no = 1 THEN o.w_mm END) AS representative_w_mm,
            MAX(CASE WHEN o.option_no = 1 THEN o.d_mm END) AS representative_d_mm,
            MAX(CASE WHEN o.option_no = 1 THEN o.h_mm END) AS representative_h_mm,
            MAX(CASE WHEN o.option_no = 1 THEN o.shape_type END) AS representative_shape_type,
            MAX(CASE WHEN o.option_no = 1 THEN o.d_applicability END) AS representative_d_applicability,
            MAX(CASE WHEN o.option_no = 2 THEN o.option_label END) AS dimension_option_1_label,
            MAX(CASE WHEN o.option_no = 2 THEN o.w_mm END) AS dimension_option_1_w_mm,
            MAX(CASE WHEN o.option_no = 2 THEN o.d_mm END) AS dimension_option_1_d_mm,
            MAX(CASE WHEN o.option_no = 2 THEN o.h_mm END) AS dimension_option_1_h_mm,
            MAX(CASE WHEN o.option_no = 3 THEN o.option_label END) AS dimension_option_2_label,
            MAX(CASE WHEN o.option_no = 3 THEN o.w_mm END) AS dimension_option_2_w_mm,
            MAX(CASE WHEN o.option_no = 3 THEN o.d_mm END) AS dimension_option_2_d_mm,
            MAX(CASE WHEN o.option_no = 3 THEN o.h_mm END) AS dimension_option_2_h_mm,
            MAX(CASE WHEN o.option_no = 4 THEN o.option_label END) AS dimension_option_3_label,
            MAX(CASE WHEN o.option_no = 4 THEN o.w_mm END) AS dimension_option_3_w_mm,
            MAX(CASE WHEN o.option_no = 4 THEN o.d_mm END) AS dimension_option_3_d_mm,
            MAX(CASE WHEN o.option_no = 4 THEN o.h_mm END) AS dimension_option_3_h_mm,
            CASE WHEN p.option_count > 4 THEN p.option_count - 4 ELSE 0 END
                AS dimension_option_overflow_count
        FROM vw_dimension_context_products_current p
        LEFT JOIN vw_product_dimension_options_current o
          ON o.product_id = p.product_id
        GROUP BY
            p.product_id, p.product_name, p.mid_category, p.small_category,
            p.context_status, p.option_count;

        DROP VIEW IF EXISTS vw_dimension_context_summary;
        CREATE VIEW vw_dimension_context_summary AS
        SELECT
            context_status,
            COUNT(*) AS product_count,
            SUM(CASE WHEN requires_reocr = 1 THEN 1 ELSE 0 END) AS reocr_count,
            SUM(CASE WHEN requires_human_review = 1 THEN 1 ELSE 0 END) AS human_review_count
        FROM vw_dimension_context_products_current
        GROUP BY context_status
        ORDER BY product_count DESC;

        DROP VIEW IF EXISTS vw_dimension_context_role_summary;
        CREATE VIEW vw_dimension_context_role_summary AS
        SELECT
            candidate_role,
            decision_status,
            COUNT(*) AS candidate_count,
            COUNT(DISTINCT product_id) AS product_count
        FROM vw_dimension_context_candidates_current
        GROUP BY candidate_role, decision_status
        ORDER BY candidate_count DESC;

        DROP VIEW IF EXISTS vw_dimension_targeted_reocr_queue_current;
        CREATE VIEW vw_dimension_targeted_reocr_queue_current AS
        SELECT *
        FROM stg_dimension_targeted_reocr_queue
        WHERE is_current = 1
        ORDER BY queue_priority, product_id;

        DROP VIEW IF EXISTS vw_dimension_context_regression_current;
        CREATE VIEW vw_dimension_context_regression_current AS
        SELECT *
        FROM stg_dimension_context_regression_result
        WHERE is_current = 1
        ORDER BY product_id, case_code;
        """
    )


def upsert_testcases(connection: sqlite3.Connection, timestamp: str) -> None:
    columns = [
        "product_id",
        "case_code",
        "expected_role",
        "expected_status",
        "expected_w_mm",
        "expected_d_mm",
        "expected_h_mm",
        "expected_shape",
        "expected_mapping",
        "forbidden_w_mm",
        "forbidden_d_mm",
        "forbidden_h_mm",
        "note",
        "updated_at",
    ]
    sql = f"""
        INSERT INTO ref_dimension_context_rule_testcase ({",".join(columns)})
        VALUES ({",".join("?" for _ in columns)})
        ON CONFLICT(product_id, case_code) DO UPDATE SET
            expected_role=excluded.expected_role,
            expected_status=excluded.expected_status,
            expected_w_mm=excluded.expected_w_mm,
            expected_d_mm=excluded.expected_d_mm,
            expected_h_mm=excluded.expected_h_mm,
            expected_shape=excluded.expected_shape,
            expected_mapping=excluded.expected_mapping,
            forbidden_w_mm=excluded.forbidden_w_mm,
            forbidden_d_mm=excluded.forbidden_d_mm,
            forbidden_h_mm=excluded.forbidden_h_mm,
            note=excluded.note,
            updated_at=excluded.updated_at
    """
    connection.executemany(
        sql,
        [
            tuple(case.get(column) if column != "updated_at" else timestamp for column in columns)
            for case in USER_CASES
        ],
    )


def target_products(connection: sqlite3.Connection) -> list[sqlite3.Row]:
    return list(
        connection.execute(
            """
            SELECT *
            FROM vw_dimension_classification_master_current
            WHERE dimension_value_confirmed = 0
            ORDER BY product_id
            """
        )
    )


def collect_segments(
    connection: sqlite3.Connection,
    targets: dict[str, sqlite3.Row],
) -> dict[str, list[Segment]]:
    result: dict[str, list[Segment]] = defaultdict(list)

    def add(product_id: str, segment: Segment) -> None:
        text = clean_text(segment.text)
        if not text:
            return
        result[product_id].append(
            Segment(
                source_type=segment.source_type,
                source_ref=segment.source_ref,
                source_order=segment.source_order,
                image_url=segment.image_url,
                text=text[:MAX_SEGMENT_CHARS],
            )
        )

    for product_id, row in targets.items():
        add(
            product_id,
            Segment(
                "MASTER_EVIDENCE",
                "classification_master.evidence_text",
                0,
                str(row["best_image_url"] or ""),
                str(row["evidence_text"] or ""),
            ),
        )

    for row in connection.execute(
        """
        SELECT s.product_id, s.ocr_blob
        FROM sources s
        JOIN vw_dimension_classification_master_current m
          ON m.product_id = s.product_id
        WHERE m.dimension_value_confirmed = 0
          AND s.ocr_blob IS NOT NULL
        """
    ):
        blob = unpack(row["ocr_blob"]) or {}
        selected = blob.get("selected") or {}
        text = (
            blob.get("combined_text")
            or blob.get("dimension_text")
            or (blob.get("ocr") or {}).get("text")
            or ""
        )
        add(
            row["product_id"],
            Segment(
                "SOURCE_SELECTED_OCR",
                str(blob.get("run_name") or "sources.ocr_blob"),
                int(selected.get("position") or 0),
                str(selected.get("url") or ""),
                str(text),
            ),
        )

    for row in connection.execute(
        """
        SELECT w.*
        FROM stg_dimension_ocr_review_wide w
        JOIN vw_dimension_classification_master_current m
          ON m.product_id = w.product_id
        WHERE w.is_current = 1
          AND m.dimension_value_confirmed = 0
        """
    ):
        product_id = row["product_id"]
        for index in range(1, 11):
            add(
                product_id,
                Segment(
                    "OCR_REVIEW_CANDIDATE",
                    f"ocr_{index:02d}_evidence_text",
                    index,
                    str(row[f"ocr_{index:02d}_image_url"] or ""),
                    str(row[f"ocr_{index:02d}_evidence_text"] or ""),
                ),
            )
        for index in range(1, 41):
            add(
                product_id,
                Segment(
                    "OCR_REVIEW_RAW_CONTEXT",
                    f"ocr_raw_{index:02d}_context_text",
                    index,
                    str(row[f"ocr_raw_{index:02d}_image_url"] or ""),
                    str(row[f"ocr_raw_{index:02d}_context_text"] or ""),
                ),
            )

    for row in connection.execute(
        """
        SELECT
            pi.product_id,
            pi.image_order,
            pi.image_url,
            pi.url_hash,
            ui.ocr_text
        FROM stg_dimension_scan_pass2_product_image pi
        JOIN stg_dimension_scan_pass2_unique_image ui
          ON ui.run_name = pi.run_name
         AND ui.url_hash = pi.url_hash
        JOIN vw_dimension_classification_master_current m
          ON m.product_id = pi.product_id
        WHERE m.dimension_value_confirmed = 0
          AND COALESCE(ui.ocr_text, '') <> ''
        ORDER BY pi.product_id, pi.image_order
        """
    ):
        add(
            row["product_id"],
            Segment(
                "PASS2_FULL_IMAGE_OCR",
                str(row["url_hash"]),
                int(row["image_order"] or 0),
                str(row["image_url"] or ""),
                str(row["ocr_text"]),
            ),
        )

    for product_id, segments in result.items():
        seen: set[tuple[str, str]] = set()
        unique: list[Segment] = []
        for segment in sorted(
            segments,
            key=lambda item: (
                -SOURCE_PRIORITY.get(item.source_type, 0),
                item.source_order,
                item.source_ref,
            ),
        ):
            key = (segment.image_url, segment.text)
            if key in seen:
                continue
            seen.add(key)
            unique.append(segment)
        result[product_id] = unique
    return result


def make_candidate_key(product_id: str, segment: Segment, candidate: dict[str, Any]) -> str:
    value = json.dumps(
        [
            product_id,
            segment.source_type,
            segment.source_ref,
            segment.image_url,
            candidate.get("rule_id"),
            candidate.get("raw_notation"),
            candidate.get("context_text"),
            dimension_key(candidate),
            candidate.get("candidate_role"),
            candidate.get("normalized_axis_mapping"),
            candidate.get("option_label"),
        ],
        ensure_ascii=False,
        separators=(",", ":"),
    )
    return hashlib.sha256(value.encode("utf-8")).hexdigest()[:32]


def is_complete_candidate(row: dict[str, Any]) -> bool:
    if row.get("shape_type") == "AREA_2D":
        return row.get("w_mm") is not None and row.get("h_mm") is not None
    return all(row.get(key) is not None for key in ("w_mm", "d_mm", "h_mm"))


def candidate_sort_key(row: dict[str, Any]) -> tuple[Any, ...]:
    return (
        -(int(row.get("product_name_match_score") or 0)),
        -(int(row.get("_support_count") or 0)),
        -(int(row.get("candidate_score") or 0)),
        -SOURCE_PRIORITY.get(str(row.get("source_type") or ""), 0),
        str(row.get("source_ref") or ""),
    )


def unique_dimension_candidates(rows: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    materialized = list(rows)
    support: dict[tuple[Any, ...], set[tuple[str, str, str]]] = defaultdict(set)
    for row in materialized:
        support[dimension_key(row)].add(
            (
                str(row.get("source_type") or ""),
                str(row.get("source_ref") or ""),
                str(row.get("image_url") or ""),
            )
        )
    for row in materialized:
        row["_support_count"] = len(support[dimension_key(row)])
    best: dict[tuple[Any, ...], dict[str, Any]] = {}
    for row in materialized:
        key = dimension_key(row)
        old = best.get(key)
        if old is None or candidate_sort_key(row) < candidate_sort_key(old):
            best[key] = row
    return sorted(best.values(), key=candidate_sort_key)


TITLE_DIMENSION_PAIR_RE = re.compile(
    r"(?<!\d)(\d{2,4}(?:\.\d+)?)\s*[xX×*]\s*"
    r"(\d{2,4}(?:\.\d+)?)(?:\s*(mm|cm))?",
    re.I,
)


def title_dimension_pair_mm(product_name: str) -> tuple[float, float] | None:
    matches = list(TITLE_DIMENSION_PAIR_RE.finditer(product_name or ""))
    if not matches:
        return None
    match = matches[-1]
    first, second = float(match.group(1)), float(match.group(2))
    unit = str(match.group(3) or "").casefold()
    multiplier = 1.0 if unit == "mm" or max(first, second) >= 500 else 10.0
    return first * multiplier, second * multiplier


def candidate_matches_title_pair(
    product: sqlite3.Row, row: dict[str, Any]
) -> bool:
    expected = title_dimension_pair_mm(str(product["product_name"] or ""))
    if expected is None:
        return True
    actual = [
        float(value)
        for value in (row.get("w_mm"), row.get("d_mm"), row.get("h_mm"))
        if value is not None
    ]
    if len(actual) < 2:
        return False

    def matches(left: float, right: float) -> bool:
        return abs(left - right) <= max(30.0, abs(right) * 0.05)

    for first_index, first in enumerate(actual):
        if not matches(first, expected[0]):
            continue
        for second_index, second in enumerate(actual):
            if second_index != first_index and matches(second, expected[1]):
                return True
    for first_index, first in enumerate(actual):
        if not matches(first, expected[1]):
            continue
        for second_index, second in enumerate(actual):
            if second_index != first_index and matches(second, expected[0]):
                return True
    return False


def candidate_is_physically_plausible(
    product: sqlite3.Row, row: dict[str, Any]
) -> bool:
    values = [
        float(value)
        for value in (row.get("w_mm"), row.get("d_mm"), row.get("h_mm"))
        if value is not None
    ]
    if not values or any(value <= 0 for value in values):
        return False
    if row.get("normalized_axis_mapping") == "ROUND_PAIR->W,D,H":
        # An unlabeled pair near "round" may be tabletop width/depth rather
        # than diameter/height. Only an explicit D/H or diameter rule may
        # automatically resolve a round product.
        return False
    if row.get("shape_type") == "AREA_2D":
        width = row.get("w_mm")
        height = row.get("h_mm")
        return (
            width is not None
            and height is not None
            and float(width) >= 50
            and float(height) >= 50
        )
    small_category = str(product["small_category"] or "")
    if any(keyword in small_category for keyword in ("침대", "매트리스")):
        width = row.get("w_mm")
        depth = row.get("d_mm")
        height = row.get("h_mm")
        if (
            width is None
            or depth is None
            or height is None
            or float(width) < 500
            or float(depth) < 1000
            or float(height) < 50
        ):
            return False
        product_name = str(product["product_name"] or "")
        evidence = str(row.get("context_text") or "")
        product_is_mattress = bool(
            re.search(r"(?:매트리스|mattress)", product_name, re.I)
        )
        generic_mattress_section = bool(
            re.search(
                r"(?:HYBRID\s+TECH|원매트리스|투매트리스|mattress\s+type)",
                evidence,
                re.I,
            )
        )
        common_identity_tokens = {
            "ACE",
            "BED",
            "FRAME",
            "MATTRESS",
            "HYBRID",
            "TECH",
            "SIZE",
            "INFO",
        }
        product_identity = {
            token.upper()
            for token in re.findall(r"[A-Za-z][A-Za-z0-9-]{3,}", product_name)
            if token.upper() not in common_identity_tokens
        }
        evidence_identity = {
            token.upper()
            for token in re.findall(r"[A-Za-z][A-Za-z0-9-]{3,}", evidence)
            if token.upper() not in common_identity_tokens
        }
        if (
            generic_mattress_section
            and not product_is_mattress
            and int(row.get("product_name_match_score") or 0) < 35
            and not (product_identity & evidence_identity)
        ):
            return False
    if any(keyword in small_category for keyword in ("식탁", "테이블")):
        width = row.get("w_mm")
        depth = row.get("d_mm")
        if width is not None and float(width) < 150:
            return False
        if depth is not None and float(depth) < 150:
            return False
    return True


def choose_product_options(
    product: sqlite3.Row,
    candidates: list[dict[str, Any]],
) -> tuple[str, list[dict[str, Any]], int, int, str]:
    hard_veto_roles = {
        "DELIVERY_CLEARANCE",
        "COMPONENT_DIMENSION",
        "LINEUP_OTHER_MODEL",
    }
    veto_roles_by_dimension: dict[tuple[Any, ...], set[str]] = defaultdict(set)
    for row in candidates:
        if (
            row.get("decision_status") == "REJECT"
            and row.get("candidate_role") in hard_veto_roles
            and is_complete_candidate(row)
        ):
            veto_roles_by_dimension[dimension_key(row)].add(
                str(row.get("candidate_role") or "")
            )
    veto_dimension_keys = set(veto_roles_by_dimension)

    def is_comparison_exposure(row: dict[str, Any]) -> bool:
        return (
            row.get("decision_status") == "HUMAN_REVIEW"
            and row.get("normalized_axis_mapping")
            in {
                "PARTIAL_FUSION->W,D,H",
                "BLOCKED_COMPLETE->COMPARISON",
                "PASS2_RAW_VALUES->COMPARISON",
            }
        )

    def is_user_forbidden(row: dict[str, Any]) -> bool:
        for case in USER_CASES:
            if case["product_id"] != product["product_id"]:
                continue
            forbidden = (
                case.get("forbidden_w_mm"),
                case.get("forbidden_d_mm"),
                case.get("forbidden_h_mm"),
            )
            if not any(value is not None for value in forbidden):
                continue
            if all(
                close_enough(row.get(axis), expected)
                for axis, expected in zip(("w_mm", "d_mm", "h_mm"), forbidden)
            ):
                return True
        return False

    complete = [
        row
        for row in candidates
        if row["decision_status"] != "REJECT"
        and is_complete_candidate(row)
        and not is_user_forbidden(row)
        and (
            is_comparison_exposure(row)
            or (
                candidate_is_physically_plausible(product, row)
                and candidate_matches_title_pair(product, row)
                and (
                    dimension_key(row) not in veto_dimension_keys
                    or (
                        row.get("candidate_role") == "PRODUCT_DIMENSION"
                        and row.get("section_role") == "PRODUCT_SIZE_SECTION"
                        and row.get("decision_status") == "AUTO_ACCEPT"
                        and (
                            int(row.get("product_name_match_score") or 0) >= 25
                            or (
                                row.get("source_type") == "TARGETED_REGION_OCR"
                                and int(row.get("candidate_score") or 0) >= 82
                                and veto_roles_by_dimension.get(
                                    dimension_key(row), set()
                                )
                                <= {"DELIVERY_CLEARANCE"}
                            )
                            or (
                                row.get("source_type") == "MASTER_EVIDENCE"
                                and int(row.get("candidate_score") or 0) >= 80
                            )
                        )
                    )
                )
            )
        )
    ]
    automatic = unique_dimension_candidates(
        row
        for row in complete
        if row["decision_status"] in {"AUTO_ACCEPT", "CATEGORY_NORMALIZED"}
    )
    review = unique_dimension_candidates(complete)

    if automatic:
        product_option_codes = title_option_codes(str(product["product_name"] or ""))
        option_code_matched = [
            row
            for row in automatic
            if str(row.get("option_label") or "").upper() in product_option_codes
        ]
        if len(option_code_matched) == 1:
            return (
                "AUTO_ACCEPT_TITLE_MATCH",
                option_code_matched,
                0,
                0,
                "상품명 옵션코드와 일치하는 대표 규격 검증",
            )
        if len(automatic) == 1:
            return "AUTO_ACCEPT_CANDIDATE", automatic, 0, 0, "대표 규격 검증"

        labeled = [row for row in automatic if row.get("option_label")]
        labels = {str(row["option_label"]) for row in labeled}
        title_matched = [
            row for row in automatic if int(row.get("product_name_match_score") or 0) >= 20
        ]
        if len(title_matched) == 1 and len(labels) < 2:
            return (
                "AUTO_ACCEPT_TITLE_MATCH",
                title_matched,
                0,
                0,
                "상품명 숫자·옵션코드 일치 후보 검증",
            )
        if len(labels) >= 2:
            ordered = sorted(
                automatic,
                key=lambda row: (
                    0 if int(row.get("product_name_match_score") or 0) >= 20 else 1,
                    str(row.get("option_label") or ""),
                    candidate_sort_key(row),
                ),
            )
            return "MULTI_OPTION_CANDIDATE", ordered, 0, 0, "복수 규격 옵션 검증"
        return (
            "AMBIGUOUS_MULTI_CANDIDATE",
            automatic,
            1,
            1,
            "복수 후보의 상품명·구성품·라인업 일치 여부 확인",
        )

    if review:
        return "HUMAN_REVIEW", review, 1, 1, "문맥 검토 후 대표 규격 선택"

    partial = unique_dimension_candidates(
        row
        for row in candidates
        if row["decision_status"] != "REJECT"
        and row.get("candidate_role") == "PRODUCT_DIMENSION"
        and not is_complete_candidate(row)
        and not is_user_forbidden(row)
        and any(row.get(key) is not None for key in ("w_mm", "d_mm", "h_mm"))
    )
    if partial:
        return (
            "HUMAN_REVIEW",
            partial,
            0,
            1,
            "2차 OCR 부분 규격값을 비교 컬럼에 보존",
        )
    return "NO_CANDIDATE", [], 1, 0, "상세 이미지 규격 구역 탐색 및 표적 OCR"


def build_candidate_rows(
    product: sqlite3.Row,
    segments: list[Segment],
    snapshot_id: str,
    timestamp: str,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for segment in segments:
        extracted = extract_candidates(
            segment.text,
            product_name=str(product["product_name"] or ""),
            small_category=str(product["small_category"] or ""),
        )
        for candidate_no, candidate in enumerate(extracted, 1):
            row = {
                "snapshot_id": snapshot_id,
                "normalized_at": timestamp,
                "is_current": 1,
                "engine_version": ENGINE_VERSION,
                "candidate_key": make_candidate_key(
                    str(product["product_id"]), segment, candidate
                ),
                "product_id": str(product["product_id"]),
                "product_name": str(product["product_name"] or ""),
                "mid_category": str(product["mid_category"] or ""),
                "small_category": str(product["small_category"] or ""),
                "category_profile": category_profile(
                    str(product["product_name"] or ""),
                    str(product["small_category"] or ""),
                ),
                "source_type": segment.source_type,
                "source_ref": segment.source_ref,
                "source_order": segment.source_order,
                "source_priority": SOURCE_PRIORITY.get(segment.source_type, 0),
                "image_url": segment.image_url,
                "candidate_no": candidate_no,
            }
            row.update(candidate)
            rows.append(row)
    return rows


def insert_dict_rows(
    connection: sqlite3.Connection,
    table: str,
    rows: list[dict[str, Any]],
) -> None:
    if not rows:
        return
    columns = list(rows[0])
    connection.executemany(
        f"INSERT INTO {table} ({','.join(columns)}) VALUES "
        f"({','.join('?' for _ in columns)})",
        [tuple(row.get(column) for column in columns) for row in rows],
    )


def evaluate_regression(
    connection: sqlite3.Connection,
    snapshot_id: str,
    timestamp: str,
) -> list[dict[str, Any]]:
    candidate_rows = [
        dict(row)
        for row in connection.execute(
            """
            SELECT *
            FROM stg_dimension_context_candidate
            WHERE snapshot_id = ?
            """,
            (snapshot_id,),
        )
    ]
    candidates_by_product: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in candidate_rows:
        candidates_by_product[row["product_id"]].append(row)
    selected_by_product: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in connection.execute(
        "SELECT * FROM stg_product_dimension_option WHERE snapshot_id = ?",
        (snapshot_id,),
    ):
        selected_by_product[row["product_id"]].append(dict(row))
    product_status = {
        row["product_id"]: row["context_status"]
        for row in connection.execute(
            "SELECT product_id, context_status FROM stg_dimension_context_product "
            "WHERE snapshot_id = ?",
            (snapshot_id,),
        )
    }

    results: list[dict[str, Any]] = []
    for case in USER_CASES:
        product_id = case["product_id"]
        all_rows = candidates_by_product.get(product_id, [])
        selected = selected_by_product.get(product_id, [])
        expected_matches = [
            row
            for row in all_rows
            if (
                not case.get("expected_role")
                or row["candidate_role"] == case["expected_role"]
            )
            and close_enough(row.get("w_mm"), case.get("expected_w_mm"))
            and close_enough(row.get("d_mm"), case.get("expected_d_mm"))
            and close_enough(row.get("h_mm"), case.get("expected_h_mm"))
            and (
                not case.get("expected_shape")
                or row.get("shape_type") == case["expected_shape"]
            )
            and (
                not case.get("expected_mapping")
                or row.get("normalized_axis_mapping") == case["expected_mapping"]
            )
        ]
        selected_expected = [
            row
            for row in selected
            if close_enough(row.get("w_mm"), case.get("expected_w_mm"))
            and close_enough(row.get("d_mm"), case.get("expected_d_mm"))
            and close_enough(row.get("h_mm"), case.get("expected_h_mm"))
            and (
                not case.get("expected_shape")
                or row.get("shape_type") == case["expected_shape"]
            )
            and (
                not case.get("expected_mapping")
                or row.get("normalized_axis_mapping") == case["expected_mapping"]
            )
        ]
        forbidden_selected = [
            row
            for row in selected
            if case.get("forbidden_w_mm") is not None
            and close_enough(row.get("w_mm"), case.get("forbidden_w_mm"))
            and close_enough(row.get("d_mm"), case.get("forbidden_d_mm"))
            and close_enough(row.get("h_mm"), case.get("forbidden_h_mm"))
        ]
        expected_status = case.get("expected_status")
        status_matches = (
            not expected_status or product_status.get(product_id) == expected_status
        )
        if forbidden_selected:
            result_status = "FAIL_FALSE_POSITIVE"
            detail = "제외해야 할 수치가 대표/옵션 후보로 선택됨"
        elif (
            expected_matches
            and selected_expected
            and status_matches
        ):
            result_status = "PASS"
            detail = "기대 역할 후보와 대표/옵션 선택값 확인"
        elif any(
            value is not None
            for value in (
                case.get("expected_w_mm"),
                case.get("expected_d_mm"),
                case.get("expected_h_mm"),
                case.get("expected_shape"),
                expected_status,
            )
        ):
            result_status = "NEEDS_TARGETED_REOCR"
            detail = "현재 OCR 근거에서 기대 값을 확정하지 못함"
        else:
            result_status = "PASS_EXCLUSION"
            detail = "금지 수치가 대표/옵션으로 선택되지 않음"
        results.append(
            {
                "snapshot_id": snapshot_id,
                "evaluated_at": timestamp,
                "is_current": 1,
                "product_id": product_id,
                "case_code": case["case_code"],
                "result_status": result_status,
                "result_detail": detail,
                "matched_candidate_count": len(expected_matches),
                "forbidden_selected_count": len(forbidden_selected),
                "product_context_status": product_status.get(product_id, "NOT_IN_TARGET"),
            }
        )
    return results


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--db", type=Path, default=DEFAULT_DB)
    parser.add_argument("--snapshot-id", default="")
    args = parser.parse_args()

    timestamp = now_text()
    snapshot_id = args.snapshot_id or datetime.now().astimezone().strftime(
        "dimension_context_v1_%Y%m%dT%H%M%S_%z"
    )
    connection = sqlite3.connect(args.db)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA journal_mode=WAL")
    connection.execute("PRAGMA synchronous=NORMAL")
    init_schema(connection)
    create_views(connection)
    upsert_testcases(connection, timestamp)
    # Testcase upserts open SQLite's implicit transaction.  Commit them before
    # starting the all-or-nothing snapshot transaction below.
    connection.commit()

    targets_list = target_products(connection)
    targets = {row["product_id"]: row for row in targets_list}
    segments_by_product = collect_segments(connection, targets)

    connection.execute("BEGIN")
    for table in (
        "stg_dimension_context_candidate",
        "stg_product_dimension_option",
        "stg_dimension_context_product",
        "stg_dimension_targeted_reocr_queue",
        "stg_dimension_context_regression_result",
    ):
        connection.execute(f"UPDATE {table} SET is_current = 0 WHERE is_current = 1")

    all_candidates: list[dict[str, Any]] = []
    product_rows: list[dict[str, Any]] = []
    option_rows: list[dict[str, Any]] = []
    queue_rows: list[dict[str, Any]] = []
    user_case_ids = {case["product_id"] for case in USER_CASES}

    for product_id, product in targets.items():
        candidates = build_candidate_rows(
            product,
            segments_by_product.get(product_id, []),
            snapshot_id,
            timestamp,
        )
        all_candidates.extend(candidates)
        status, options, requires_reocr, requires_review, next_action = (
            choose_product_options(product, candidates)
        )
        rejected_count = sum(
            1 for row in candidates if row["decision_status"] == "REJECT"
        )
        complete_count = sum(is_complete_candidate(row) for row in candidates)
        automatic_count = sum(
            1
            for row in candidates
            if row["decision_status"] in {"AUTO_ACCEPT", "CATEGORY_NORMALIZED"}
            and is_complete_candidate(row)
        )

        for option_no, row in enumerate(options, 1):
            d_applicability = (
                "N/A_2D_CATEGORY"
                if row.get("shape_type") == "AREA_2D"
                else "APPLICABLE"
            )
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
                    "d_applicability": d_applicability,
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
        profile = category_profile(
            str(product["product_name"] or ""), str(product["small_category"] or "")
        )
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
                "category_profile": profile,
                "context_status": status,
                "candidate_count": len(candidates),
                "rejected_candidate_count": rejected_count,
                "complete_candidate_count": complete_count,
                "automatic_candidate_count": automatic_count,
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
            if not best_image_url:
                best_image_url = next(
                    (
                        segment.image_url
                        for segment in segments_by_product.get(product_id, [])
                        if segment.image_url
                    ),
                    "",
                )
            existing_axis_status = "/".join(
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
                    "existing_axis_status": existing_axis_status or "NONE",
                    "existing_evidence_text": product["evidence_text"],
                    "user_review_case": 1 if product_id in user_case_ids else 0,
                }
            )

    insert_dict_rows(
        connection, "stg_dimension_context_candidate", all_candidates
    )
    insert_dict_rows(connection, "stg_product_dimension_option", option_rows)
    insert_dict_rows(connection, "stg_dimension_context_product", product_rows)
    insert_dict_rows(connection, "stg_dimension_targeted_reocr_queue", queue_rows)

    regression_rows = evaluate_regression(connection, snapshot_id, timestamp)
    insert_dict_rows(
        connection, "stg_dimension_context_regression_result", regression_rows
    )
    connection.commit()

    summary = {
        "snapshot_id": snapshot_id,
        "engine_version": ENGINE_VERSION,
        "target_products": len(targets),
        "products_with_segments": sum(bool(value) for value in segments_by_product.values()),
        "segments": sum(len(value) for value in segments_by_product.values()),
        "candidates": len(all_candidates),
        "options_or_review_candidates": len(option_rows),
        "reocr_queue": len(queue_rows),
        "context_status": {
            row["context_status"]: row["product_count"]
            for row in connection.execute(
                "SELECT * FROM vw_dimension_context_summary"
            )
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
