from __future__ import annotations

import json
import re
import sqlite3
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

from bulk_homestyle_collect import DB_PATH, RUN_DIR


TYPE_TABLE = "ref_dimension_notation_type"
STAGING_TABLE = "stg_dimension_pattern"
LATEST_OUTPUT = RUN_DIR / "dimension_pattern_staging_latest.json"
PARSER_VERSION = "dimension-pattern-v1; target-vs-notation-separated; L-deferred"

NUMBER = r"\d{1,4}(?:[.,]\d+)?(?:\s*[-~]\s*\d{1,4}(?:[.,]\d+)?)?"
AXIS_VALUE_RE = re.compile(
    rf"(?<![A-Z])(L|W|D|H)\s*[:=]?\s*\(?\s*({NUMBER})", re.I
)
VALUE_AXIS_RE = re.compile(
    rf"({NUMBER})\s*\(\s*(L|W|D|H)\s*\)", re.I
)
TRIPLE_RE = re.compile(
    rf"(?<!\d){NUMBER}\s*(?:mm|cm)?\s*(?:x|×|\*)\s*"
    rf"{NUMBER}\s*(?:mm|cm)?\s*(?:x|×|\*)\s*{NUMBER}\s*(?:mm|cm)?",
    re.I,
)
PAIR_RE = re.compile(
    rf"(?<!\d){NUMBER}\s*(?:mm|cm)?\s*(?:x|×|\*)\s*{NUMBER}\s*(?:mm|cm)?",
    re.I,
)
DIAMETER_RE = re.compile(rf"(?:DIA\.?|Ø|Φ|⌀|지름|직경)\s*[:=]?\s*{NUMBER}", re.I)
HEIGHT_RE = re.compile(rf"(?:\bH|높이)\s*[:=]?\s*{NUMBER}", re.I)
BED_CODE_RE = re.compile(r"(?:^|[\s(/_-])(SS|GSS|Q|K|LK|S|D)(?:$|[\s)/_-])", re.I)


TYPE_ROWS = [
    (
        "W_D_H_EXPLICIT", "명시 W/D/H", "W,D,H", "W→W, D→D, H→H",
        "MAPPED", 10, "원문에 W·D·H 축과 숫자가 모두 명시된 타입",
    ),
    (
        "KOREAN_W_D_H", "한글 3축", "가로/너비/폭, 깊이/세로, 높이",
        "가로/너비/폭→W, 깊이/세로→D, 높이→H", "MAPPED", 20,
        "한글 축 라벨이 모두 있는 타입. 세로의 의미가 높이인 카테고리는 검수",
    ),
    (
        "L_D_H_EXPLICIT", "명시 L/D/H", "L,D,H",
        "후보: L→W, D→D, H→H", "INTERPRETATION_REQUIRED", 30,
        "L의 제조사 의미를 확인한 뒤 W로 매핑하는 타입",
    ),
    (
        "L_W_H_EXPLICIT", "명시 L/W/H", "L,W,H",
        "후보: L→W, 원문 W→D, H→H", "INTERPRETATION_REQUIRED", 31,
        "Length/Width/Height 관례 후보이나 3D 축 정의 확인이 필요한 타입",
    ),
    (
        "W_L_H_EXPLICIT", "명시 W/L/H", "W,L,H",
        "후보: W→W, L→D, H→H", "INTERPRETATION_REQUIRED", 32,
        "원문 순서가 W/L/H인 타입. L의 깊이 의미 확인 필요",
    ),
    (
        "DIAMETER_H", "직경+높이", "Ø/DIA/직경,H",
        "직경→W·D, H→H", "CATEGORY_POLICY_REQUIRED", 40,
        "원형 외형일 때 직경을 W와 D에 복제할 수 있는 타입",
    ),
    (
        "UNLABELED_3_AXIS", "무라벨 3축", "A×B×C",
        "후보: 순서 기반 W/D/H", "ORDER_REVIEW_REQUIRED", 50,
        "축 라벨 없는 세 숫자 타입. 카테고리/도면 순서 검수 필요",
    ),
    (
        "W_D_PAIR", "명시 W/D 2축", "W,D", "W→W, D→D; H 정책 필요",
        "INCOMPLETE_OR_NA_POLICY", 60, "폭과 깊이만 명시된 2축 타입",
    ),
    (
        "W_H_PAIR", "명시 W/H 2축", "W,H", "W→W, H→H; D 정책 필요",
        "INCOMPLETE_OR_NA_POLICY", 61, "폭과 높이만 명시된 2축 타입",
    ),
    (
        "L_W_PAIR", "명시 L/W 2축", "L,W", "L/W 의미 확인; H 정책 필요",
        "INTERPRETATION_REQUIRED", 62, "L과 W만 명시된 2축 타입",
    ),
    (
        "UNLABELED_2_AXIS", "무라벨 2축", "A×B", "카테고리별 W/D 또는 W/H",
        "CATEGORY_POLICY_REQUIRED", 63, "러그·액자 등 카테고리 축 정책이 필요한 타입",
    ),
    (
        "PARTIAL_W_D", "부분축 W/D", "W,D 확보", "H 추가 추출 또는 비적용 정책",
        "MISSING_AXIS", 70, "현재 정규화 결과에서 W와 D만 확보된 타입",
    ),
    (
        "PARTIAL_W_H", "부분축 W/H", "W,H 확보", "D 추가 추출 또는 비적용 정책",
        "MISSING_AXIS", 71, "현재 정규화 결과에서 W와 H만 확보된 타입",
    ),
    (
        "PARTIAL_D_H", "부분축 D/H", "D,H 확보", "W 추가 추출",
        "MISSING_AXIS", 72, "현재 정규화 결과에서 D와 H만 확보된 타입",
    ),
    (
        "PARTIAL_EXPLICIT_AXIS", "기타 부분축", "1~2개 명시 축",
        "누락 축 추가 추출", "MISSING_AXIS", 73, "명시 축 일부만 확인된 타입",
    ),
    (
        "STANDARD_SIZE_CODE", "표준 사이즈 코드", "SS/Q/K 등",
        "브랜드·모델 규격표 연결", "SOURCE_TABLE_REQUIRED", 80,
        "코드만으로 실제 제품 외형 W/D/H를 확정할 수 없는 타입",
    ),
    (
        "NO_CONFIRMED_PATTERN", "확정 패턴 없음", "-", "-", "NO_SIGNAL", 99,
        "현재 API/HTML/OCR/옵션에서 확정 가능한 규격 패턴을 찾지 못한 타입",
    ),
]


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def clean(value: Any, limit: int = 1000) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    return text if len(text) <= limit else text[:limit] + "…"


def target_type(value: str) -> str:
    compact = re.sub(r"\s+", "", value or "")
    if "옵션" in compact:
        return "PRODUCT_OPTION"
    if "구성품" in compact or "부속" in compact:
        return "PRODUCT_COMPONENT"
    if "포장" in compact or "박스" in compact:
        return "PACKAGE"
    if "제품" in compact or "외형" in compact:
        return "PRODUCT_EXTERIOR"
    return "UNKNOWN"


def axis_orders(text: str) -> list[tuple[str, str]]:
    matches: list[tuple[int, str, str]] = []
    for match in AXIS_VALUE_RE.finditer(text):
        matches.append((match.start(), match.group(1).upper(), match.group(0)))
    for match in VALUE_AXIS_RE.finditer(text):
        matches.append((match.start(), match.group(2).upper(), match.group(0)))
    matches.sort()
    result: list[tuple[str, str]] = []
    seen: set[str] = set()
    for _, axis, evidence in matches:
        if axis not in seen:
            seen.add(axis)
            result.append((axis, evidence))
    return result


def detect_one(text: str) -> dict[str, Any] | None:
    value = clean(text, 5000)
    if not value:
        return None
    axes = axis_orders(value)
    axis_sequence = [axis for axis, _ in axes]
    axis_set = set(axis_sequence)
    evidence = " | ".join(item for _, item in axes)

    if {"W", "D", "H"}.issubset(axis_set):
        return pattern("W_D_H_EXPLICIT", ">".join(axis_sequence), evidence or value)
    if {"L", "D", "H"}.issubset(axis_set) and "W" not in axis_set:
        return pattern("L_D_H_EXPLICIT", ">".join(axis_sequence), evidence or value)
    if {"L", "W", "H"}.issubset(axis_set) and "D" not in axis_set:
        code = "W_L_H_EXPLICIT" if axis_sequence.index("W") < axis_sequence.index("L") else "L_W_H_EXPLICIT"
        return pattern(code, ">".join(axis_sequence), evidence or value)

    has_kr_w = bool(re.search(r"가로|너비|폭", value))
    has_kr_d = bool(re.search(r"깊이|세로", value))
    has_kr_h = bool(re.search(r"높이", value))
    if has_kr_w and has_kr_d and has_kr_h:
        return pattern("KOREAN_W_D_H", "한글W>한글D>한글H", value)
    if DIAMETER_RE.search(value) and HEIGHT_RE.search(value):
        return pattern("DIAMETER_H", "DIAMETER>H", value)
    if TRIPLE_RE.search(value):
        return pattern("UNLABELED_3_AXIS", "A>B>C", TRIPLE_RE.search(value).group(0))
    if axis_set == {"W", "D"}:
        return pattern("W_D_PAIR", ">".join(axis_sequence), evidence or value)
    if axis_set == {"W", "H"}:
        return pattern("W_H_PAIR", ">".join(axis_sequence), evidence or value)
    if axis_set == {"L", "W"}:
        return pattern("L_W_PAIR", ">".join(axis_sequence), evidence or value)
    if axis_sequence:
        return pattern("PARTIAL_EXPLICIT_AXIS", ">".join(axis_sequence), evidence or value)
    if PAIR_RE.search(value):
        return pattern("UNLABELED_2_AXIS", "A>B", PAIR_RE.search(value).group(0))
    if BED_CODE_RE.search(value):
        return pattern("STANDARD_SIZE_CODE", "SIZE_CODE", BED_CODE_RE.search(value).group(0))
    return None


def pattern(code: str, order: str, evidence: str) -> dict[str, Any]:
    row = next(item for item in TYPE_ROWS if item[0] == code)
    return {
        "type_code": code,
        "type_name": row[1],
        "axis_order": order,
        "axis_mapping": row[3],
        "mapping_status": row[4],
        "sort_order": row[5],
        "evidence": clean(evidence, 1500),
    }


def select_pattern(
    texts: list[str],
    *,
    notation_rule: str,
    notation_evidence: str,
    cause_code: str,
    candidate_rule: str,
    w_mm: Any,
    d_mm: Any,
    h_mm: Any,
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    diagnostics: list[dict[str, Any]] = []
    if "L-W-H" in notation_rule:
        detected_order = detect_one(notation_evidence)
        if detected_order and detected_order["type_code"] in {
            "L_W_H_EXPLICIT", "W_L_H_EXPLICIT"
        }:
            diagnostics.append(detected_order)
        else:
            diagnostics.append(pattern("L_W_H_EXPLICIT", "L>W>H", notation_evidence))
    elif "L-D-H" in notation_rule:
        detected_order = detect_one(notation_evidence)
        if detected_order and detected_order["type_code"] == "L_D_H_EXPLICIT":
            diagnostics.append(detected_order)
        else:
            diagnostics.append(pattern("L_D_H_EXPLICIT", "L>D>H", notation_evidence))
    elif "직경" in notation_rule:
        diagnostics.append(pattern("DIAMETER_H", "DIAMETER>H", notation_evidence))
    elif cause_code == "DIAMETER_AXIS_MAPPING_REQUIRED":
        diagnostics.append(pattern("DIAMETER_H", "DIAMETER>H", notation_evidence))

    for text in texts:
        detected = detect_one(text)
        if detected and not any(
            row["type_code"] == detected["type_code"] and row["evidence"] == detected["evidence"]
            for row in diagnostics
        ):
            diagnostics.append(detected)

    if not diagnostics:
        if candidate_rule == "OPTION_FLAT_WD":
            diagnostics.append(pattern("W_D_PAIR", "W>D", "옵션 2축; 러그 W/D 후보"))
        elif candidate_rule == "OPTION_FLAT_WH":
            diagnostics.append(pattern("W_H_PAIR", "W>H", "옵션 2축; 액자/포스터 W/H 후보"))
        elif candidate_rule == "BED_STANDARD_SIZE_CODE":
            diagnostics.append(pattern("STANDARD_SIZE_CODE", "SIZE_CODE", candidate_rule))

    if not diagnostics:
        present = {
            axis for axis, value in (("W", w_mm), ("D", d_mm), ("H", h_mm))
            if isinstance(value, (int, float)) and not isinstance(value, bool)
        }
        code = {
            frozenset({"W", "D"}): "PARTIAL_W_D",
            frozenset({"W", "H"}): "PARTIAL_W_H",
            frozenset({"D", "H"}): "PARTIAL_D_H",
        }.get(frozenset(present))
        if code:
            present_order = [axis for axis in ("W", "D", "H") if axis in present]
            diagnostics.append(
                pattern(code, ">".join(present_order), f"현재 확보 축={','.join(present_order)}")
            )
        elif present and len(present) < 3:
            present_order = [axis for axis in ("W", "D", "H") if axis in present]
            diagnostics.append(
                pattern("PARTIAL_EXPLICIT_AXIS", ">".join(present_order), f"현재 확보 축={','.join(present_order)}")
            )

    if not diagnostics:
        diagnostics.append(pattern("NO_CONFIRMED_PATTERN", "", ""))
    diagnostics.sort(key=lambda row: row["sort_order"])
    return diagnostics[0], diagnostics


def create_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        f"""
        CREATE TABLE IF NOT EXISTS {TYPE_TABLE} (
            type_code TEXT PRIMARY KEY,
            type_name TEXT NOT NULL,
            source_axis_pattern TEXT,
            normalized_axis_mapping TEXT,
            default_mapping_status TEXT NOT NULL,
            sort_order INTEGER NOT NULL,
            description TEXT
        );

        CREATE TABLE IF NOT EXISTS {STAGING_TABLE} (
            snapshot_id TEXT NOT NULL,
            assessed_at TEXT NOT NULL,
            is_current INTEGER NOT NULL DEFAULT 1,
            parser_version TEXT NOT NULL,
            product_id TEXT NOT NULL,
            product_name TEXT,
            small_category TEXT,
            representative_target_value TEXT,
            representative_target_type TEXT NOT NULL,
            current_dimension_status TEXT,
            current_w_mm REAL,
            current_d_mm REAL,
            current_h_mm REAL,
            notation_type_code TEXT NOT NULL,
            notation_type_name TEXT NOT NULL,
            source_axis_order TEXT,
            normalized_axis_mapping TEXT,
            mapping_status TEXT NOT NULL,
            pattern_evidence TEXT,
            detected_patterns_json TEXT,
            PRIMARY KEY(snapshot_id,product_id),
            FOREIGN KEY(notation_type_code) REFERENCES {TYPE_TABLE}(type_code)
        );
        CREATE INDEX IF NOT EXISTS idx_dimension_pattern_current_type
            ON {STAGING_TABLE}(is_current,notation_type_code);
        CREATE INDEX IF NOT EXISTS idx_dimension_pattern_current_target
            ON {STAGING_TABLE}(is_current,representative_target_type);
        CREATE INDEX IF NOT EXISTS idx_dimension_pattern_current_product
            ON {STAGING_TABLE}(is_current,product_id);
        CREATE INDEX IF NOT EXISTS idx_reinforcement_backlog_current_product
            ON stg_reinforcement_backlog(is_current,product_id);

        DROP VIEW IF EXISTS vw_dimension_pattern_current_summary;
        CREATE VIEW vw_dimension_pattern_current_summary AS
        SELECT notation_type_code,notation_type_name,mapping_status,current_dimension_status,
               COUNT(*) AS product_count
        FROM {STAGING_TABLE}
        WHERE is_current=1
        GROUP BY notation_type_code,notation_type_name,mapping_status,current_dimension_status;

        DROP VIEW IF EXISTS vw_dimension_reinforcement_with_pattern;
        CREATE VIEW vw_dimension_reinforcement_with_pattern AS
        SELECT b.*, p.representative_target_value, p.representative_target_type,
               p.notation_type_code, p.notation_type_name, p.source_axis_order,
               p.normalized_axis_mapping, p.mapping_status AS axis_mapping_status,
               p.pattern_evidence, p.detected_patterns_json
        FROM stg_reinforcement_backlog b
        JOIN {STAGING_TABLE} p ON p.product_id=b.product_id AND p.is_current=1
        WHERE b.is_current=1 AND b.missing_field='규격(W/D/H)';
        """
    )
    with connection:
        connection.executemany(
            f"""
            INSERT INTO {TYPE_TABLE} VALUES (?,?,?,?,?,?,?)
            ON CONFLICT(type_code) DO UPDATE SET
                type_name=excluded.type_name,
                source_axis_pattern=excluded.source_axis_pattern,
                normalized_axis_mapping=excluded.normalized_axis_mapping,
                default_mapping_status=excluded.default_mapping_status,
                sort_order=excluded.sort_order,
                description=excluded.description
            """,
            TYPE_ROWS,
        )


def build_snapshot() -> dict[str, Any]:
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    # Create the product join indexes before loading the source rows. Without
    # the composite current/product indexes SQLite can choose a large nested
    # loop across the two staging tables.
    create_schema(connection)
    source_rows = connection.execute(
        """
        SELECT d.*, m.product_name AS mandatory_product_name,
               m.small_category_value AS mandatory_small_category,
               b.cause_code AS backlog_cause_code,
               b.candidate_evidence AS backlog_candidate_evidence,
               b.details_json AS backlog_details_json
        FROM stg_dimension_reinforcement d
        JOIN stg_mandatory_pass m
          ON m.product_id=d.product_id AND m.is_current=1
        LEFT JOIN stg_reinforcement_backlog b
          ON b.product_id=d.product_id AND b.is_current=1
         AND b.missing_field='규격(W/D/H)'
        WHERE d.is_current=1
        ORDER BY d.product_id
        """
    ).fetchall()

    assessed_at = now_text()
    snapshot_id = assessed_at.replace(":", "").replace("+", "_")
    inserts: list[tuple[Any, ...]] = []
    type_counts: Counter[str] = Counter()
    target_counts: Counter[str] = Counter()
    mapping_counts: Counter[str] = Counter()

    for row in source_rows:
        product_id = str(row["product_id"])
        product_name = str(row["mandatory_product_name"] or row["product_name"] or "")
        small_category = str(row["mandatory_small_category"] or row["small_category"] or "")
        dimension_status = str(row["current_status"] or "")
        w_mm = row["current_w_mm"]
        d_mm = row["current_d_mm"]
        h_mm = row["current_h_mm"]
        notation_rule = ""
        notation_evidence = ""
        candidate_rule = str(row["candidate_rule"] or "")
        cause_code = str(row["backlog_cause_code"] or "")
        texts: list[str] = []
        try:
            records = json.loads(row["dimension_records_json"] or "[]")
        except json.JSONDecodeError:
            records = []
        targets = list(
            dict.fromkeys(
                clean(item.get("target"), 200)
                for item in records
                if isinstance(item, dict) and item.get("target")
            )
        )
        target_value = "|".join(targets) or "제품 외형"
        try:
            details = json.loads(row["backlog_details_json"] or "{}")
        except json.JSONDecodeError:
            details = {}
        notation_rule = str(details.get("notation_rule") or "")
        notation_evidence = str(row["backlog_candidate_evidence"] or "")

        selected, detected = select_pattern(
            texts,
            notation_rule=notation_rule,
            notation_evidence=notation_evidence,
            cause_code=cause_code,
            candidate_rule=candidate_rule,
            w_mm=w_mm,
            d_mm=d_mm,
            h_mm=h_mm,
        )
        normalized_target = target_type(target_value)
        type_counts[selected["type_code"]] += 1
        target_counts[normalized_target] += 1
        mapping_counts[selected["mapping_status"]] += 1
        inserts.append(
            (
                snapshot_id,
                assessed_at,
                1,
                PARSER_VERSION,
                product_id,
                product_name,
                small_category,
                target_value,
                normalized_target,
                dimension_status,
                w_mm,
                d_mm,
                h_mm,
                selected["type_code"],
                selected["type_name"],
                selected["axis_order"],
                selected["axis_mapping"],
                selected["mapping_status"],
                selected["evidence"],
                json.dumps(detected, ensure_ascii=False),
            )
        )

    placeholders = ",".join("?" for _ in range(20))
    with connection:
        connection.execute(f"UPDATE {STAGING_TABLE} SET is_current=0 WHERE is_current=1")
        connection.executemany(
            f"INSERT INTO {STAGING_TABLE} VALUES ({placeholders})", inserts
        )
    reinforcement_type_counts = {
        code: count
        for code, count in connection.execute(
            """
            SELECT notation_type_code,COUNT(*)
            FROM vw_dimension_reinforcement_with_pattern
            GROUP BY notation_type_code
            ORDER BY COUNT(*) DESC
            """
        )
    }
    result = {
        "database": str(DB_PATH),
        "reference_table": TYPE_TABLE,
        "staging_table": STAGING_TABLE,
        "snapshot_id": snapshot_id,
        "assessed_at": assessed_at,
        "parser_version": PARSER_VERSION,
        "products": len(inserts),
        "type_counts": dict(type_counts.most_common()),
        "reinforcement_type_counts": reinforcement_type_counts,
        "target_type_counts": dict(target_counts.most_common()),
        "mapping_status_counts": dict(mapping_counts.most_common()),
        "views": [
            "vw_dimension_pattern_current_summary",
            "vw_dimension_reinforcement_with_pattern",
        ],
        "excel_written": False,
        "notes": [
            "요청1_대표 규격 대상은 제품/옵션/구성품 대상 구분으로 유지했다.",
            "규격 표기 타입과 W/D/H 축 매핑 상태를 별도 필드로 추가했다.",
            "L 포함 타입은 INTERPRETATION_REQUIRED이며 자동 치환하지 않는다.",
            f"현재 스냅샷은 규격 보강대상 {len(inserts):,}개를 분류한다.",
        ],
    }
    connection.close()
    LATEST_OUTPUT.write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return result


if __name__ == "__main__":
    print(json.dumps(build_snapshot(), ensure_ascii=False, indent=2))
