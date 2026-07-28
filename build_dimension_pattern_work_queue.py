from __future__ import annotations

import json
import re
import sqlite3
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

from bulk_homestyle_collect import DB_PATH, RUN_DIR


TABLE = "stg_dimension_pattern_work_queue"
TAXONOMY_VERSION = "dimension-pattern-taxonomy-v1.0-20260723"
PASS1_RUN_NAME = "dimension_scan_pass1_all_detail_images_v1"
SUMMARY_PATH = RUN_DIR / "dimension_pattern_work_queue_latest.json"
AXIS_ORDER = ("W", "D", "H", "L")
AXIS_RE = re.compile(r"(?i)^\s*([WDHL])")


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def axis_text(axes: set[str]) -> str:
    return ",".join(axis for axis in AXIS_ORDER if axis in axes)


def missing_wdh_text(axes: set[str]) -> str:
    return ",".join(axis for axis in ("W", "D", "H") if axis not in axes)


def current_axes(row: sqlite3.Row) -> set[str]:
    return {
        axis
        for axis, column in (
            ("W", "current_w_mm"),
            ("D", "current_d_mm"),
            ("H", "current_h_mm"),
        )
        if row[column] is not None
    }


def raw_axes(row: sqlite3.Row) -> set[str]:
    result: set[str] = set()
    for number in range(1, 41):
        value = str(row[f"ocr_raw_{number:02d}_value_text"] or "")
        match = AXIS_RE.match(value)
        if match:
            result.add(match.group(1).upper())
    return result


def unit_suffix(unit_status: str) -> str:
    return "UNIT_PRESENT" if unit_status in {"ALL_UNIT_PRESENT", "MIXED_UNIT_PRESENT"} else "UNIT_MISSING"


def current_partial_pattern(axes: set[str]) -> tuple[str, str]:
    signature = "_".join(axis for axis in ("W", "D", "H") if axis in axes) or "NONE"
    names = {"W": "폭", "D": "깊이", "H": "높이"}
    label = "·".join(names[axis] for axis in ("W", "D", "H") if axis in axes)
    return f"P10_CURRENT_{signature}_ONLY", f"기존 구조값 {label}만 확보"


def partial_axis_pattern(axes: set[str], unit_status: str) -> tuple[str, str]:
    suffix = unit_suffix(unit_status)
    count = len(axes)
    if count == 1:
        base, name = "P20_OCR_SINGLE_AXIS", "OCR 단일 축"
    elif count == 2:
        base, name = "P21_OCR_TWO_AXES", "OCR 2개 축"
    else:
        base, name = "P22_OCR_NONSTANDARD_MULTI_AXES", "OCR 비표준 다축"
    unit_name = "단위 있음" if suffix == "UNIT_PRESENT" else "단위 없음"
    return f"{base}_{suffix}", f"{name}·{unit_name}"


def complete_candidate_pattern(axes: set[str], unit_status: str) -> tuple[str, str, str, str, int, str]:
    suffix = unit_suffix(unit_status)
    unit_name = "단위 있음" if suffix == "UNIT_PRESENT" else "단위 없음"
    if {"W", "D", "H"}.issubset(axes):
        return (
            f"C01_COMPLETE_WDH_EXPLICIT_{suffix}",
            f"완전 W/D/H 후보·축 명시·{unit_name}",
            "WDH_MAPPED",
            "VALIDATE_THEN_APPLY",
            1,
            "후보 간 충돌과 제품 대표 규격 여부를 검증한 뒤 W/D/H 적용",
        )
    if axes:
        return (
            f"C02_COMPLETE_TRIPLE_PARTIAL_LABEL_{suffix}",
            f"완전 3값 후보·일부 축만 명시·{unit_name}",
            "AXIS_MAPPING_REQUIRED",
            "RULE_WITH_AXIS_INFERENCE",
            2,
            "L/비표준 축과 값 순서를 해석한 뒤 W/D/H로 매핑",
        )
    return (
        f"C03_COMPLETE_TRIPLE_UNLABELED_{suffix}",
        f"완전 3값 후보·축 미표기·{unit_name}",
        "ORDER_INFERENCE_REQUIRED",
        "RULE_WITH_ORDER_INFERENCE",
        2,
        "카테고리별 값 순서 규칙과 도면 문맥으로 W/D/H 순서 판단",
    )


def size_label_pattern(axes: set[str]) -> tuple[str, str, str, str]:
    if {"W", "D", "H"}.issubset(axes):
        return (
            "U01_SIZE_LABEL_WDH_UNIT_INFERENCE",
            "사이즈 문구·W/D/H 있음·단위 없음",
            "WDH_MAPPED",
            "INFER_UNIT_THEN_VALIDATE",
        )
    if len(axes) == 2:
        return (
            "U02_SIZE_LABEL_TWO_AXES_UNIT_INFERENCE",
            "사이즈 문구·2개 축·단위 없음",
            "MISSING_AXIS_AND_UNIT",
            "SECOND_PASS_FOR_MISSING_AXIS",
        )
    if len(axes) == 1:
        return (
            "U03_SIZE_LABEL_SINGLE_AXIS_UNIT_INFERENCE",
            "사이즈 문구·단일 축·단위 없음",
            "MISSING_AXES_AND_UNIT",
            "SECOND_PASS_FOR_MISSING_AXES",
        )
    if axes:
        return (
            "U04_SIZE_LABEL_NONSTANDARD_AXES",
            "사이즈 문구·비표준 다축·단위 없음",
            "NONSTANDARD_AXIS_MAPPING_REQUIRED",
            "CATEGORY_AXIS_MAPPING",
        )
    return (
        "U05_SIZE_LABEL_UNLABELED_NUMBERS",
        "사이즈 문구·숫자 있음·축/단위 없음",
        "ORDER_AND_UNIT_INFERENCE_REQUIRED",
        "SECOND_PASS_SPATIAL_OCR",
    )


def unclassified_pattern(axes: set[str], raw_count: int) -> tuple[str, str, str, str, int, str]:
    if {"W", "D", "H"}.issubset(axes):
        return (
            "M01_RAW_WDH_GROUPING_REQUIRED",
            "OCR W/D/H 원시값·동일 규격 그룹화 필요",
            "WDH_LABELS_DETECTED_GROUP_UNKNOWN",
            "GROUP_BY_IMAGE_AND_CONTEXT",
            1,
            "같은 이미지·근접 문맥의 W/D/H를 한 세트로 그룹화",
        )
    if axes:
        return (
            "M02_RAW_NONSTANDARD_AXES_MAPPING_REQUIRED",
            "OCR 비표준 축 원시값·의미 매핑 필요",
            "NONSTANDARD_AXIS_MAPPING_REQUIRED",
            "CATEGORY_AXIS_MAPPING",
            2,
            "L/폭/길이 등 카테고리별 축 의미 규칙 적용",
        )
    if raw_count == 3:
        return (
            "M03_UNLABELED_TRIPLET_ORDER_INFERENCE",
            "축 미표기 OCR 숫자 3개",
            "ORDER_INFERENCE_REQUIRED",
            "ORDER_PATTERN_BY_CATEGORY",
            2,
            "카테고리 및 표기 순서 패턴으로 W/D/H 후보 생성",
        )
    if raw_count == 2:
        return (
            "M04_UNLABELED_PAIR",
            "축 미표기 OCR 숫자 2개",
            "ONE_AXIS_MISSING",
            "SECOND_PASS_FOR_MISSING_AXIS",
            3,
            "2차 OCR로 누락 축 또는 원형/평면 2축 제품 여부 확인",
        )
    if raw_count == 1:
        return (
            "M05_UNLABELED_SINGLE_VALUE",
            "축 미표기 OCR 숫자 1개",
            "MULTIPLE_AXES_MISSING",
            "SECOND_PASS_FOR_MISSING_AXES",
            4,
            "2차 OCR 또는 제품 유형별 단일 규격 적용 가능성 확인",
        )
    return (
        "M06_UNLABELED_MULTI_VALUE_OPTION_SPLIT",
        "축 미표기 OCR 숫자 4개 이상",
        "OPTION_OR_COMPONENT_SPLIT_REQUIRED",
        "SPLIT_BY_IMAGE_REGION_AND_OPTION",
        3,
        "옵션/구성품/도면 영역별로 숫자 군집을 분리한 뒤 대표 규격 선택",
    )


def pass1_pattern(best_class: str) -> tuple[str, str, str, str, int, str]:
    if best_class in {"SIZE_LABEL_AXIS_UNIT", "SIZE_LABEL_WITH_AXIS", "MULTI_AXIS_FOUND"}:
        return (
            "S01_PASS1_STRONG_AXIS_IMAGE",
            "1차 OCR 강한 축 규격 이미지",
            "IMAGE_AXIS_SIGNAL_DETECTED",
            "SECOND_PASS_SPATIAL_OCR",
            2,
            "후보 이미지의 규격 영역을 탐지하고 고해상도 좌표 OCR 수행",
        )
    if best_class in {
        "SIZE_LABEL_WITH_UNIT",
        "SIZE_LABEL_UNIT_MISSING",
        "DIMENSION_PAIR_FOUND",
        "DIMENSION_TRIPLE_FOUND",
    }:
        return (
            "S02_PASS1_SIZE_PATTERN_IMAGE",
            "1차 OCR 사이즈/숫자 패턴 이미지",
            "IMAGE_SIZE_SIGNAL_DETECTED",
            "SECOND_PASS_TARGETED_OCR",
            3,
            "사이즈 문구와 숫자 군집 주변을 잘라 2차 OCR 수행",
        )
    if best_class in {"PARTIAL_AXIS_FOUND", "UNIT_NUMERIC_CLUSTER"}:
        return (
            "S03_PASS1_PARTIAL_SIGNAL_IMAGE",
            "1차 OCR 부분 축/단위 숫자 이미지",
            "IMAGE_PARTIAL_SIGNAL_DETECTED",
            "SECOND_PASS_TARGETED_OCR",
            4,
            "부분 축 또는 단위 주변 영역을 확대 OCR하고 제품 규격 여부 판정",
        )
    return (
        "S04_PASS1_URL_HINT_IMAGE",
        "이미지 URL 규격 힌트",
        "URL_HINT_ONLY",
        "SECOND_PASS_FULL_IMAGE_OCR",
        5,
        "URL 의미만 있으므로 전체 이미지를 2차 OCR해 실제 규격 신호 확인",
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
            review_classification TEXT NOT NULL,
            work_group_code TEXT NOT NULL,
            work_group_name TEXT NOT NULL,
            pattern_code TEXT NOT NULL,
            pattern_name TEXT NOT NULL,
            current_axes TEXT,
            ocr_axes TEXT,
            available_axes TEXT,
            missing_wdh_axes TEXT,
            current_value_count INTEGER NOT NULL,
            ocr_axis_count INTEGER NOT NULL,
            raw_value_count INTEGER NOT NULL,
            complete_candidate_count INTEGER NOT NULL,
            unit_status TEXT NOT NULL,
            size_label_present INTEGER NOT NULL,
            semantic_mapping_status TEXT NOT NULL,
            unit_normalization_status TEXT NOT NULL,
            pattern_readiness TEXT NOT NULL,
            apply_mode TEXT NOT NULL,
            priority INTEGER NOT NULL,
            pass1_product_classification TEXT,
            pass1_candidate_image_count INTEGER NOT NULL,
            pass1_best_image_order INTEGER,
            pass1_best_image_url TEXT,
            pass1_best_classification TEXT,
            pass1_best_score INTEGER,
            next_action TEXT NOT NULL,
            rule_notes TEXT,
            PRIMARY KEY(snapshot_id,product_id)
        );

        CREATE INDEX IF NOT EXISTS idx_dimension_pattern_queue_current_pattern
            ON {TABLE}(is_current,pattern_code,priority);
        CREATE INDEX IF NOT EXISTS idx_dimension_pattern_queue_current_group
            ON {TABLE}(is_current,work_group_code,small_category);

        CREATE VIEW IF NOT EXISTS vw_dimension_pattern_work_queue_current AS
        SELECT * FROM {TABLE} WHERE is_current=1;

        CREATE VIEW IF NOT EXISTS vw_dimension_pattern_work_queue_summary AS
        SELECT pattern_code,pattern_name,pattern_readiness,apply_mode,priority,
               COUNT(*) AS product_count
        FROM {TABLE}
        WHERE is_current=1
        GROUP BY pattern_code,pattern_name,pattern_readiness,apply_mode,priority;
        """
    )


def load_pass1(connection: sqlite3.Connection) -> tuple[dict[str, sqlite3.Row], dict[str, sqlite3.Row]]:
    products: dict[str, sqlite3.Row] = {}
    best: dict[str, sqlite3.Row] = {}
    for row in connection.execute(
        "SELECT * FROM stg_dimension_scan_pass1_product WHERE run_name=?",
        (PASS1_RUN_NAME,),
    ):
        products[row["product_id"]] = row
    for row in connection.execute(
        """
        SELECT * FROM vw_dimension_scan_pass1_candidate_images
        WHERE run_name=?
        ORDER BY product_id,candidate_score DESC,image_order
        """,
        (PASS1_RUN_NAME,),
    ):
        best.setdefault(row["product_id"], row)
    return products, best


def build() -> dict[str, Any]:
    snapshot_id = datetime.now().astimezone().strftime("%Y%m%dT%H%M%S%z")
    classified_at = now_text()
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    create_schema(connection)
    pass1_products, pass1_best = load_pass1(connection)
    rows = connection.execute(
        "SELECT * FROM stg_dimension_ocr_review_wide WHERE is_current=1 ORDER BY product_id"
    ).fetchall()
    inserts: list[tuple[Any, ...]] = []
    pattern_counts: Counter[str] = Counter()
    group_counts: Counter[str] = Counter()

    for row in rows:
        pid = row["product_id"]
        review_class = row["ocr_review_classification"]
        c_axes = current_axes(row)
        o_axes = raw_axes(row)
        available = c_axes | o_axes
        unit_status = str(row["ocr_unit_status"] or "NO_VALUES")
        raw_count = int(row["ocr_raw_token_count"] or 0)
        candidate_count = int(row["ocr_candidate_count"] or 0)
        pass1_product = pass1_products.get(pid)
        best = pass1_best.get(pid)
        best_class = str(best["classification"] if best else "")
        readiness = "CONDITIONAL"
        unit_normalization = (
            "UNIT_INFERENCE_REQUIRED"
            if unit_status == "UNIT_MISSING"
            else "UNIT_NORMALIZATION_REQUIRED"
            if unit_status == "MIXED_UNIT_PRESENT"
            else "UNIT_READY"
            if unit_status == "ALL_UNIT_PRESENT"
            else "NO_OCR_VALUE"
        )
        notes = f"기존 분류={review_class}; OCR축={axis_text(o_axes) or '-'}; 원시값={raw_count}"

        if review_class == "COMPLETE_WDH_CANDIDATE":
            work_group, work_name = "G01_COMPLETE_CANDIDATE", "완전 3값 후보"
            pattern, pattern_name, semantic, apply_mode, priority, action = complete_candidate_pattern(
                o_axes, unit_status
            )
            readiness = "RULE_VALIDATION_READY"
        elif review_class == "CURRENT_PARTIAL_ONLY":
            work_group, work_name = "G02_CURRENT_PARTIAL", "기존 구조값 일부 확보"
            pattern, pattern_name = current_partial_pattern(c_axes)
            semantic = "STRUCTURED_AXES_MAPPED"
            apply_mode = "MERGE_AFTER_MISSING_AXIS_RECOVERY"
            priority = 2 if len(c_axes) == 2 else 3
            readiness = "MISSING_AXIS_REQUIRED"
            unit_normalization = "STANDARD_MM_EXISTING_VALUE"
            action = f"누락 축({missing_wdh_text(c_axes)})만 보강 후 기존 구조값과 병합"
        elif review_class == "PARTIAL_AXES":
            work_group, work_name = "G03_OCR_PARTIAL_AXES", "OCR 일부 축 확보"
            pattern, pattern_name = partial_axis_pattern(o_axes, unit_status)
            semantic = "PARTIAL_AXES_MAPPED"
            apply_mode = "RECOVER_MISSING_AXES"
            priority = 2 if len(o_axes) >= 2 else 3
            readiness = "MISSING_AXIS_REQUIRED"
            action = f"누락 축({missing_wdh_text(o_axes)})을 같은 이미지 문맥 또는 2차 OCR로 보강"
        elif review_class == "SIZE_LABEL_UNIT_MISSING":
            work_group, work_name = "G04_SIZE_LABEL_UNIT_MISSING", "사이즈 문구·단위 미확보"
            pattern, pattern_name, semantic, apply_mode = size_label_pattern(o_axes)
            priority = 2 if {"W", "D", "H"}.issubset(o_axes) else 3
            readiness = "UNIT_OR_AXIS_INFERENCE_REQUIRED"
            action = "카테고리 기본 단위와 이미지 문맥으로 단위를 판단하고 누락 축 보강"
        elif review_class == "UNCLASSIFIED_OCR_VALUES":
            work_group, work_name = "G05_UNCLASSIFIED_OCR", "OCR 값 의미 미분류"
            pattern, pattern_name, semantic, apply_mode, priority, action = unclassified_pattern(
                o_axes, raw_count
            )
            readiness = "SEMANTIC_GROUPING_REQUIRED"
        else:
            product_class = str(
                pass1_product["product_classification"] if pass1_product else "PASS1_NOT_FOUND"
            )
            if product_class == "CANDIDATE_IMAGES_FOUND":
                work_group, work_name = "G06_PASS1_IMAGE_CANDIDATE", "1차 OCR 이미지 후보 발견"
                pattern, pattern_name, semantic, apply_mode, priority, action = pass1_pattern(best_class)
                readiness = "SECOND_PASS_OCR_REQUIRED"
                unit_normalization = "SECOND_PASS_RESULT_REQUIRED"
            elif product_class == "NO_SIZE_SIGNAL_ALL_IMAGES":
                work_group, work_name = "G07_PASS1_NO_SIGNAL", "전체 이미지 규격 신호 없음"
                pattern, pattern_name = "S90_PASS1_NO_SIZE_SIGNAL", "전체 상세 이미지 OCR 규격 신호 없음"
                semantic, apply_mode, priority = "NO_SIGNAL", "NEW_SOURCE_RESEARCH", 8
                readiness = "NOT_READY"
                unit_normalization = "NO_VALUE"
                action = "옵션/HTML/제조사 원문 등 새 소스를 조사하거나 사람 검토"
            elif product_class == "NO_DETAIL_IMAGES":
                work_group, work_name = "G08_NO_DETAIL_IMAGE", "상세 이미지 없음"
                pattern, pattern_name = "S91_NO_DETAIL_IMAGE", "상세 이미지 소스 없음"
                semantic, apply_mode, priority = "NO_IMAGE_SOURCE", "NEW_SOURCE_REQUIRED", 9
                readiness = "NOT_READY"
                unit_normalization = "NO_VALUE"
                action = "상품 API/HTML/제조사 페이지에서 상세 이미지 또는 규격 소스 확보"
            else:
                work_group, work_name = "G09_SCAN_INCOMPLETE", "1차 OCR 이미지 오류"
                pattern, pattern_name = "S92_PASS1_SOURCE_RETRY", "이미지 다운로드 오류 재시도"
                semantic, apply_mode, priority = "SOURCE_ERROR", "RETRY_OR_REPLACE_URL", 7
                readiness = "SOURCE_RETRY_REQUIRED"
                unit_normalization = "NO_VALUE"
                action = "404 URL 교체 또는 원본 판매처 이미지 재수집 후 1차 OCR 재시도"

        pass1_count = int(pass1_product["candidate_image_count"] if pass1_product else 0)
        values = (
            snapshot_id,
            classified_at,
            1,
            TAXONOMY_VERSION,
            pid,
            row["product_name"],
            row["small_category"],
            review_class,
            work_group,
            work_name,
            pattern,
            pattern_name,
            axis_text(c_axes),
            axis_text(o_axes),
            axis_text(available),
            missing_wdh_text(available),
            len(c_axes),
            len(o_axes),
            raw_count,
            candidate_count,
            unit_status,
            int(row["ocr_size_label_present"] or 0),
            semantic,
            unit_normalization,
            readiness,
            apply_mode,
            priority,
            str(pass1_product["product_classification"] if pass1_product else ""),
            pass1_count,
            int(best["image_order"]) if best else None,
            str(best["image_url"] if best else ""),
            best_class,
            int(best["candidate_score"]) if best else None,
            action,
            notes,
        )
        inserts.append(values)
        pattern_counts[pattern] += 1
        group_counts[work_group] += 1

    with connection:
        connection.execute(f"UPDATE {TABLE} SET is_current=0 WHERE is_current=1")
        connection.executemany(
            f"INSERT INTO {TABLE} VALUES ({','.join('?' for _ in range(35))})",
            inserts,
        )
    total = connection.execute(
        f"SELECT COUNT(*) FROM {TABLE} WHERE snapshot_id=?", (snapshot_id,)
    ).fetchone()[0]
    connection.close()
    if total != len(rows) or total != 5206:
        raise RuntimeError(f"Row count mismatch: source={len(rows)} inserted={total}")

    result = {
        "snapshot_id": snapshot_id,
        "classified_at": classified_at,
        "taxonomy_version": TAXONOMY_VERSION,
        "products": total,
        "work_group_counts": dict(sorted(group_counts.items())),
        "pattern_counts": dict(sorted(pattern_counts.items())),
        "table": TABLE,
        "views": [
            "vw_dimension_pattern_work_queue_current",
            "vw_dimension_pattern_work_queue_summary",
        ],
        "excel_written": False,
        "source_dimension_values_written": False,
    }
    SUMMARY_PATH.write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return result


if __name__ == "__main__":
    build()
