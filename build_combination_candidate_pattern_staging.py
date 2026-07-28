from __future__ import annotations

import json
import re
import sqlite3
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any

from bulk_homestyle_collect import DB_PATH


PATTERN_VERSION = "combination-candidate-pattern-v1.0"

LIKELY_COMBINATION = "LIKELY_COMBINATION"
LIKELY_NOT_COMBINATION = "LIKELY_NOT_COMBINATION"
REVIEW_REQUIRED = "REVIEW_REQUIRED"


PATTERN_META: dict[str, tuple[str, str, str, int]] = {
    "Y01_MULTI_BED_BUNDLE": (
        "복수 침대·사이즈 묶음",
        LIKELY_COMBINATION,
        "침대 크기별 구성품과 수량을 분리",
        1,
    ),
    "Y02_BED_NIGHTSTAND": (
        "침대+협탁",
        LIKELY_COMBINATION,
        "침대와 협탁을 별도 구성품으로 분리",
        1,
    ),
    "Y03_BED_PANEL_GUARD": (
        "침대+패널·가드",
        LIKELY_COMBINATION,
        "침대 본체와 패널·가드를 별도 구성품으로 분리",
        2,
    ),
    "Y04_BED_SLEEP_SYSTEM": (
        "침대·프레임+매트리스",
        LIKELY_COMBINATION,
        "프레임·매트리스·파운데이션을 별도 구성품으로 분리",
        1,
    ),
    "Y05_KIDS_BED_MATTRESS": (
        "유아동 침대+매트리스·토퍼",
        LIKELY_COMBINATION,
        "침대·매트리스·토퍼를 별도 구성품으로 분리",
        1,
    ),
    "Y06_KIDS_BED_ACCESSORY": (
        "유아동 침대+가드·책상",
        LIKELY_COMBINATION,
        "침대와 가드·책상 구성품을 분리",
        2,
    ),
    "Y07_TABLE_CHAIR_BENCH": (
        "테이블+의자·벤치",
        LIKELY_COMBINATION,
        "테이블·의자·벤치와 수량을 분리",
        1,
    ),
    "Y08_TABLE_TOP_LEG": (
        "상판+다리",
        LIKELY_COMBINATION,
        "상판과 다리를 별도 구성품으로 분리",
        2,
    ),
    "Y09_MODULAR_SOFA_UNITS": (
        "모듈 소파 코드 결합",
        LIKELY_COMBINATION,
        "MOD/A/B/C 모듈 코드를 구성품 순번으로 분리",
        2,
    ),
    "Y10_SOFA_ACCESSORY": (
        "소파+쿠션·헤드레스트",
        LIKELY_COMBINATION,
        "소파 본체와 동봉 액세서리를 분리",
        1,
    ),
    "Y11_CABINET_STORAGE_MODULES": (
        "거실장·수납장 모듈 결합",
        LIKELY_COMBINATION,
        "폭·형태가 다른 장 모듈을 별도 구성품으로 분리",
        2,
    ),
    "Y12_HANGER_ACCESSORY": (
        "행거+후크·바지걸이",
        LIKELY_COMBINATION,
        "행거 본체와 동봉 액세서리를 분리",
        2,
    ),
    "Y13_WARDROBE_MODULES": (
        "옷장 모듈 결합",
        LIKELY_COMBINATION,
        "행거형·서랍형·이불장형 모듈을 순서대로 분리",
        2,
    ),
    "Y14_CHAIR_BUNDLE_1PLUS1": (
        "의자 1+1",
        LIKELY_COMBINATION,
        "동일 의자 2개로 수량 정규화",
        1,
    ),
    "Y15_DESK_ACCESSORY_PACK": (
        "책상 액세서리 팩",
        LIKELY_COMBINATION,
        "책상 본체와 조명·패드·자석바 등 동봉품 분리",
        2,
    ),
    "N01_MODEL_SUFFIX_PLUS": (
        "브랜드·모델명 접미사 +",
        LIKELY_NOT_COMBINATION,
        "아카이브+/Lotus+/TERRACE+는 모델명으로 유지",
        1,
    ),
    "N02_OPTION_TYPE_JOIN": (
        "옵션 타입 연결 +",
        LIKELY_NOT_COMBINATION,
        "바퀴+글라이드, lite+ option은 구성품이 아닌 선택 옵션",
        1,
    ),
    "N03_ARTWORK_TITLE_PLUS": (
        "작품명 내부 +",
        LIKELY_NOT_COMBINATION,
        "작품명 문자열로 유지",
        1,
    ),
    "N04_COLOR_MATERIAL_JOIN": (
        "색상·소재 조합 표기",
        LIKELY_NOT_COMBINATION,
        "색상·소재 필드로 이동하고 구성품 분리 제외",
        1,
    ),
    "N05_EXTENSION_FEATURE": (
        "확장 기능 표기 (+Extension)",
        LIKELY_NOT_COMBINATION,
        "테이블 기능·옵션으로 유지",
        2,
    ),
    "N06_INTEGRATED_FUNCTION": (
        "일체형 기능 조합",
        LIKELY_NOT_COMBINATION,
        "조명+스피커가 한 본체이면 구성품 분리 제외",
        2,
    ),
    "R01_INTERNAL_CONFIGURATION": (
        "제품 내부 구성 방식",
        REVIEW_REQUIRED,
        "박스+책장형이 별도 물리 구성품인지 상세 설명 확인",
        3,
    ),
    "R99_OTHER_PLUS": (
        "기타 + 표기",
        REVIEW_REQUIRED,
        "상품 상세와 구성품 고시를 확인하여 새 패턴 정의",
        3,
    ),
}


def now_iso() -> str:
    return datetime.now().astimezone().replace(microsecond=0).isoformat()


def has(pattern: str, text: str) -> bool:
    return bool(re.search(pattern, text, flags=re.I))


def classify(row: sqlite3.Row | dict[str, Any]) -> tuple[str, str]:
    name = row["product_name"] or ""
    normalized_name = name.replace("＋", "+")
    mid = row["mid_category"] or ""
    small = row["small_category"] or ""

    if has(r"아카이브\+|Lotus\+|TERRACE\+", normalized_name):
        return "N01_MODEL_SUFFIX_PLUS", "모델명에 +가 붙어 있음"
    if has(r"바퀴\s*\+\s*글라이드|lite\+\s*option", normalized_name):
        return "N02_OPTION_TYPE_JOIN", "선택 옵션 명칭을 +로 연결"
    if has(r"Reality\s*\+\s*Image", normalized_name):
        return "N03_ARTWORK_TITLE_PLUS", "작품명 내부의 +"
    if has(
        r"블러쉬\s*\+\s*에그화이트|"
        r"그레인 마론\s*\d*\s*\+\s*그레인 네그로",
        normalized_name,
    ):
        return "N04_COLOR_MATERIAL_JOIN", "색상 또는 소재명을 +로 연결"
    if has(r"\(\s*\+\s*Extension\s*\)", normalized_name):
        return "N05_EXTENSION_FEATURE", "확장 기능을 괄호 안 +로 표기"
    if has(r"(?:조명|램프)\s*\+\s*스피커", normalized_name):
        return "N06_INTEGRATED_FUNCTION", "한 본체의 조명·스피커 기능 표기"

    if mid == "침대":
        if has(
            r"2개묶음|슈퍼싱글\s*\+\s*슈퍼싱글|"
            r"퀸\s*\+\s*(?:퀸|슈퍼싱글)|Q\s*\+\s*SS",
            normalized_name,
        ):
            return "Y01_MULTI_BED_BUNDLE", "두 침대 또는 서로 다른 사이즈 묶음"
        if has(r"침대\s*\+\s*협탁", normalized_name):
            return "Y02_BED_NIGHTSTAND", "침대와 협탁을 +로 연결"
        if has(
            r"침대.*?\+\s*패널|수납침대.*?\+\s*가드",
            normalized_name,
        ):
            return "Y03_BED_PANEL_GUARD", "침대와 패널·가드를 +로 연결"
        if has(
            r"매트리스|매트\b|매트Q|매트SS|메모리폼탑|"
            r"노뜨컴포트|파인탑|볼륨탑|프레임",
            normalized_name,
        ):
            return "Y04_BED_SLEEP_SYSTEM", "침대·프레임과 매트리스 계열을 결합"

    if mid == "유아동가구" and small == "침대":
        if has(r"매트리스|매트\b|토퍼", normalized_name):
            return "Y05_KIDS_BED_MATTRESS", "유아동 침대와 매트리스·토퍼 결합"
        if has(r"가드|책상", normalized_name):
            return "Y06_KIDS_BED_ACCESSORY", "유아동 침대와 가드·책상 결합"

    if mid == "식탁·테이블":
        if has(r"의자|체어|chair|벤치|bench", normalized_name):
            return "Y07_TABLE_CHAIR_BENCH", "테이블과 의자·벤치 결합"
        if has(r"상판\s*\+.*다리", normalized_name):
            return "Y08_TABLE_TOP_LEG", "상판과 다리를 결합"

    if mid == "소파":
        if has(
            r"MOD\.\s*\d+\s*\+|\b[ABC]\d\s*\+\s*[ABC]\d|모듈.*\+",
            normalized_name,
        ):
            return "Y09_MODULAR_SOFA_UNITS", "모듈 코드 또는 모듈 수량을 +로 연결"
        if has(r"헤드레스트|쿠션|cushion", normalized_name):
            return "Y10_SOFA_ACCESSORY", "소파와 쿠션·헤드레스트 결합"

    if mid in {"TV거실장", "책장·수납장"}:
        return "Y11_CABINET_STORAGE_MODULES", "장류 모듈·모델을 +로 연결"

    if mid == "옷장·행거":
        if has(r"후크|바지걸이|옷걸이", normalized_name):
            return "Y12_HANGER_ACCESSORY", "행거와 후크·바지걸이 결합"
        return "Y13_WARDROBE_MODULES", "옷장 내부 모듈 유형을 +로 연결"

    if mid == "의자" and has(r"1\s*\+\s*1", normalized_name):
        return "Y14_CHAIR_BUNDLE_1PLUS1", "동일 의자 1+1 판매"

    if mid == "책상":
        return "Y15_DESK_ACCESSORY_PACK", "책상 팩의 동봉품을 +로 연결"

    if mid == "유아동가구" and has(
        r"박스\s*\+\s*책장형",
        normalized_name,
    ):
        return "R01_INTERNAL_CONFIGURATION", "하나의 수납장 내부 구성 표기 가능성"

    return "R99_OTHER_PLUS", "기존 패턴에 포함되지 않은 + 표기"


def create_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS stg_combination_candidate_pattern (
            snapshot_id TEXT NOT NULL,
            built_at TEXT NOT NULL,
            is_current INTEGER NOT NULL,
            pattern_version TEXT NOT NULL,
            product_id TEXT NOT NULL,
            product_name TEXT NOT NULL,
            mid_category TEXT,
            small_category TEXT,
            pattern_code TEXT NOT NULL,
            pattern_name TEXT NOT NULL,
            provisional_decision TEXT NOT NULL,
            reason_text TEXT NOT NULL,
            split_hint TEXT NOT NULL,
            review_priority INTEGER NOT NULL,
            evidence_text TEXT NOT NULL,
            PRIMARY KEY (snapshot_id, product_id)
        );

        CREATE INDEX IF NOT EXISTS idx_combination_candidate_pattern_current
            ON stg_combination_candidate_pattern(is_current, pattern_code, product_id);

        DROP VIEW IF EXISTS vw_combination_candidate_pattern_current;
        CREATE VIEW vw_combination_candidate_pattern_current AS
        SELECT *
        FROM stg_combination_candidate_pattern
        WHERE is_current = 1;

        DROP VIEW IF EXISTS vw_combination_candidate_pattern_summary;
        CREATE VIEW vw_combination_candidate_pattern_summary AS
        SELECT
            pattern_code,
            pattern_name,
            provisional_decision,
            review_priority,
            COUNT(*) AS product_count
        FROM vw_combination_candidate_pattern_current
        GROUP BY
            pattern_code,
            pattern_name,
            provisional_decision,
            review_priority;
        """
    )


def build(db_path: Path = DB_PATH) -> dict[str, Any]:
    connection = sqlite3.connect(db_path)
    connection.row_factory = sqlite3.Row
    create_schema(connection)

    source_rows = connection.execute(
        """
        SELECT
            product_id,
            product_name,
            mid_category,
            small_category,
            detection_rule,
            evidence_text
        FROM vw_product_combination_current
        WHERE detection_status = 'CANDIDATE'
        ORDER BY product_id
        """
    ).fetchall()
    if not source_rows:
        raise RuntimeError("No current combination candidates found")

    built_at = now_iso()
    snapshot_id = (
        "combination_candidate_pattern_"
        + datetime.now().astimezone().strftime("%Y%m%d_%H%M%S")
    )
    connection.execute(
        "UPDATE stg_combination_candidate_pattern SET is_current = 0 "
        "WHERE is_current = 1"
    )

    insert_rows: list[tuple[Any, ...]] = []
    pattern_counts: Counter[str] = Counter()
    decision_counts: Counter[str] = Counter()
    examples: dict[str, list[dict[str, str]]] = {}
    for row in source_rows:
        pattern_code, reason_text = classify(row)
        pattern_name, decision, split_hint, review_priority = PATTERN_META[
            pattern_code
        ]
        pattern_counts[pattern_code] += 1
        decision_counts[decision] += 1
        examples.setdefault(pattern_code, [])
        if len(examples[pattern_code]) < 3:
            examples[pattern_code].append(
                {
                    "product_id": row["product_id"],
                    "product_name": row["product_name"],
                    "mid_category": row["mid_category"],
                }
            )
        insert_rows.append(
            (
                snapshot_id,
                built_at,
                1,
                PATTERN_VERSION,
                row["product_id"],
                row["product_name"],
                row["mid_category"],
                row["small_category"],
                pattern_code,
                pattern_name,
                decision,
                reason_text,
                split_hint,
                review_priority,
                row["evidence_text"] or row["product_name"],
            )
        )

    connection.executemany(
        """
        INSERT INTO stg_combination_candidate_pattern (
            snapshot_id,
            built_at,
            is_current,
            pattern_version,
            product_id,
            product_name,
            mid_category,
            small_category,
            pattern_code,
            pattern_name,
            provisional_decision,
            reason_text,
            split_hint,
            review_priority,
            evidence_text
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        insert_rows,
    )
    connection.commit()

    current_count = connection.execute(
        "SELECT COUNT(*) FROM vw_combination_candidate_pattern_current"
    ).fetchone()[0]
    if current_count != len(source_rows):
        raise RuntimeError(
            f"Current snapshot count mismatch: {current_count} != {len(source_rows)}"
        )
    connection.close()

    summary_rows = []
    for pattern_code in sorted(pattern_counts):
        pattern_name, decision, split_hint, review_priority = PATTERN_META[
            pattern_code
        ]
        summary_rows.append(
            {
                "pattern_code": pattern_code,
                "pattern_name": pattern_name,
                "provisional_decision": decision,
                "review_priority": review_priority,
                "product_count": pattern_counts[pattern_code],
                "split_hint": split_hint,
                "examples": examples[pattern_code],
            }
        )

    return {
        "database": str(db_path),
        "snapshot_id": snapshot_id,
        "pattern_version": PATTERN_VERSION,
        "candidate_products": len(source_rows),
        "patterns": len(summary_rows),
        "decision_counts": dict(sorted(decision_counts.items())),
        "pattern_summary": summary_rows,
    }


if __name__ == "__main__":
    print(json.dumps(build(), ensure_ascii=False, indent=2))
