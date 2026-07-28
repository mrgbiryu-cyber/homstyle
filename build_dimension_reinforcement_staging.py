from __future__ import annotations

import json
import re
import sqlite3
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

import build_homestyle_bulk_workbook as workbook
from bulk_homestyle_collect import DB_PATH, RUN_DIR, unpack
from bulk_homestyle_ocr import detail_images, normalize_url


TABLE = "stg_dimension_reinforcement"
LATEST_OUTPUT = RUN_DIR / "dimension_reinforcement_staging_latest.json"
NUMBER = r"\d{1,4}(?:[.,]\d+)?"
SEP = r"(?:x|×|Ⅹ|\*)"
PAIR_RE = re.compile(
    rf"(?<![\d.])({NUMBER})\s*(mm|cm)?\s*{SEP}\s*"
    rf"({NUMBER})\s*(mm|cm)?(?![\d.])",
    re.I,
)
TRIPLE_RE = re.compile(
    rf"(?<![\d.])({NUMBER})\s*(mm|cm)?\s*{SEP}\s*"
    rf"({NUMBER})\s*(mm|cm)?\s*{SEP}\s*"
    rf"({NUMBER})\s*(mm|cm)?(?![\d.])",
    re.I,
)
BED_CODE_RE = re.compile(
    r"(?:^|[\s(/_-])(SS|GSS|Q|K|LK|S|D)(?:$|[\s)/_-])", re.I
)
FLAT_WD = {"러그"}
FLAT_WH = {"액자", "인테리어포스터"}
PAIR_WD_WITH_HEIGHT = {
    "식탁", "거실테이블", "사이드테이블", "베드테이블", "일자형책상",
    "컴퓨터책상", "좌식책상", "협탁", "콘솔", "테이블",
}


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def sheet_rows(sheets: list[tuple], name: str) -> list[list[Any]]:
    return next(sheet[1] for sheet in sheets if sheet[0] == name)


def number(value: str) -> float:
    return float(value.replace(",", "."))


def to_mm(value: str, unit: str) -> int | float:
    result = number(value) * (10 if unit.casefold() == "cm" else 1)
    return int(result) if result.is_integer() else result


def shared_unit(values: list[str], units: list[str | None]) -> str:
    explicit = next((unit for unit in reversed(units) if unit), "")
    if explicit:
        return explicit
    maximum = max(number(value) for value in values)
    return "cm" if maximum <= 300 else "mm"


def pair_candidates(value: str) -> list[tuple[int | float, int | float, str]]:
    result = []
    for match in PAIR_RE.finditer(value):
        values = [match.group(1), match.group(3)]
        units = [match.group(2), match.group(4)]
        unit = shared_unit(values, units)
        first = to_mm(values[0], units[0] or unit)
        second = to_mm(values[1], units[1] or unit)
        if min(first, second) < 50 or max(first, second) > 10000:
            continue
        result.append((first, second, match.group(0)))
    return result


def triple_candidates(value: str) -> list[tuple[int | float, int | float, int | float, str]]:
    result = []
    for match in TRIPLE_RE.finditer(value):
        values = [match.group(1), match.group(3), match.group(5)]
        units = [match.group(2), match.group(4), match.group(6)]
        unit = shared_unit(values, units)
        converted = [
            to_mm(raw, axis_unit or unit) for raw, axis_unit in zip(values, units)
        ]
        if min(converted) < 20 or max(converted) > 10000:
            continue
        result.append((converted[0], converted[1], converted[2], match.group(0)))
    return result


def create_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        f"""
        CREATE TABLE IF NOT EXISTS {TABLE} (
            snapshot_id TEXT NOT NULL,
            assessed_at TEXT NOT NULL,
            is_current INTEGER NOT NULL DEFAULT 1,
            product_id TEXT NOT NULL,
            product_name TEXT,
            small_category TEXT,
            current_status TEXT NOT NULL,
            current_w_mm REAL,
            current_d_mm REAL,
            current_h_mm REAL,
            missing_axes TEXT,
            dimension_records_json TEXT,
            size_option_count INTEGER NOT NULL DEFAULT 0,
            size_options_json TEXT,
            option_pair_count INTEGER NOT NULL DEFAULT 0,
            option_triple_count INTEGER NOT NULL DEFAULT 0,
            standard_size_code_count INTEGER NOT NULL DEFAULT 0,
            candidate_rule TEXT,
            candidate_confidence TEXT,
            candidate_dimensions_json TEXT,
            candidate_requires_policy INTEGER NOT NULL DEFAULT 0,
            ocr_status INTEGER,
            ocr_dimension_signal INTEGER NOT NULL DEFAULT 0,
            detail_image_count INTEGER NOT NULL DEFAULT 0,
            alternate_image_count INTEGER NOT NULL DEFAULT 0,
            selected_image_url TEXT,
            suggested_action TEXT NOT NULL,
            review_status TEXT NOT NULL DEFAULT '대기',
            applied INTEGER NOT NULL DEFAULT 0,
            PRIMARY KEY(snapshot_id,product_id)
        );
        CREATE INDEX IF NOT EXISTS idx_stg_dimension_current_status
            ON {TABLE}(is_current,current_status);
        CREATE INDEX IF NOT EXISTS idx_stg_dimension_current_action
            ON {TABLE}(is_current,suggested_action);

        DROP VIEW IF EXISTS vw_dimension_reinforcement_current_summary;
        CREATE VIEW vw_dimension_reinforcement_current_summary AS
        SELECT current_status,suggested_action,candidate_confidence,COUNT(*) product_count
        FROM {TABLE}
        WHERE is_current=1
        GROUP BY current_status,suggested_action,candidate_confidence;

        DROP VIEW IF EXISTS vw_dimension_reinforcement_option_candidates;
        CREATE VIEW vw_dimension_reinforcement_option_candidates AS
        SELECT * FROM {TABLE}
        WHERE is_current=1 AND candidate_rule IS NOT NULL;
        """
    )


def build_snapshot() -> dict[str, Any]:
    sheets, meta = workbook.build_rows()
    main_rows = sheet_rows(sheets, "01_상품별_요구필드")
    dimension_rows = sheet_rows(sheets, "02_규격_상세")
    option_rows = sheet_rows(sheets, "03_옵션_상세")
    main_header = {value: index for index, value in enumerate(main_rows[0])}
    dimension_header = {value: index for index, value in enumerate(dimension_rows[0])}
    option_header = {value: index for index, value in enumerate(option_rows[0])}

    records_by_id: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in dimension_rows[1:]:
        product_id = str(row[dimension_header["상품 ID"]])
        records_by_id[product_id].append(
            {
                "target": row[dimension_header["규격 대상/옵션"]],
                "w_mm": row[dimension_header["W (mm)"]] or None,
                "d_mm": row[dimension_header["D (mm)"]] or None,
                "h_mm": row[dimension_header["H (mm)"]] or None,
                "status": row[dimension_header["규격 상태"]],
                "source": row[dimension_header["원천/비고"]],
            }
        )

    options_by_id: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in option_rows[1:]:
        product_id = str(row[option_header["상품 ID"]])
        style = str(row[option_header["옵션 스타일"]] or "")
        value = str(row[option_header["옵션 값"]] or "")
        if style == "사이즈" or PAIR_RE.search(value) or TRIPLE_RE.search(value):
            options_by_id[product_id].append({"style": style, "value": value})

    connection = sqlite3.connect(DB_PATH)
    db_rows = {
        product_id: (goods_blob, int(ocr_status or 0), unpack(ocr_blob) or {})
        for product_id, goods_blob, ocr_status, ocr_blob in connection.execute(
            "SELECT product_id,goods_blob,ocr_status,ocr_blob FROM sources"
        )
    }
    assessed_at = now_text()
    snapshot_id = assessed_at.replace(":", "").replace("+", "_")
    inserts = []
    status_counts: dict[str, int] = defaultdict(int)
    action_counts: dict[str, int] = defaultdict(int)

    for row in main_rows[1:]:
        status = str(row[main_header["요청1_규격 상태"]])
        if status == "확보":
            continue
        product_id = str(row[main_header["상품 ID"]])
        product_name = str(row[main_header["상품명"]])
        small = str(row[main_header["요청1_소카테고리"]] or "")
        current = {
            "W": row[main_header["요청1_W (mm)"]] or None,
            "D": row[main_header["요청1_D (mm)"]] or None,
            "H": row[main_header["요청1_H (mm)"]] or None,
        }
        missing_axes = [axis for axis, value in current.items() if value is None]
        options = options_by_id.get(product_id, [])
        pairs = []
        triples = []
        codes = []
        for option in options:
            pairs.extend(pair_candidates(option["value"]))
            triples.extend(triple_candidates(option["value"]))
            if (
                option["style"] == "사이즈"
                and small in {"침대", "침대+매트리스", "매트리스"}
            ):
                codes.extend(
                    match.group(1).upper()
                    for match in BED_CODE_RE.finditer(option["value"])
                )

        known_heights = sorted(
            {
                record["h_mm"] for record in records_by_id.get(product_id, [])
                if isinstance(record.get("h_mm"), (int, float))
            }
        )
        candidate_rule = None
        confidence = None
        candidate_dimensions: list[dict[str, Any]] = []
        requires_policy = 0
        if triples:
            candidate_rule = "OPTION_UNLABELED_WDH"
            confidence = "HIGH"
            candidate_dimensions = [
                {"w_mm": w, "d_mm": d, "h_mm": h, "raw": raw}
                for w, d, h, raw in triples
            ]
        elif pairs and small in FLAT_WD:
            candidate_rule = "OPTION_FLAT_WD"
            confidence = "HIGH"
            requires_policy = 1
            candidate_dimensions = [
                {"w_mm": first, "d_mm": second, "h_mm": None, "raw": raw}
                for first, second, raw in pairs
            ]
        elif pairs and small in FLAT_WH:
            candidate_rule = "OPTION_FLAT_WH"
            confidence = "HIGH"
            requires_policy = 1
            candidate_dimensions = [
                {"w_mm": first, "d_mm": None, "h_mm": second, "raw": raw}
                for first, second, raw in pairs
            ]
        elif pairs and small in PAIR_WD_WITH_HEIGHT and len(known_heights) == 1:
            candidate_rule = "OPTION_WD_PLUS_SINGLE_H"
            confidence = "MEDIUM"
            candidate_dimensions = [
                {
                    "w_mm": first,
                    "d_mm": second,
                    "h_mm": known_heights[0],
                    "raw": raw,
                    "height_source": "기존 규격 레코드의 단일 H",
                }
                for first, second, raw in pairs
            ]
        elif codes:
            candidate_rule = "BED_STANDARD_SIZE_CODE"
            confidence = "LOW"
            requires_policy = 1
            candidate_dimensions = [{"size_code": code} for code in sorted(set(codes))]

        goods_blob, ocr_status, old_ocr = db_rows[product_id]
        data = (unpack(goods_blob) or {}).get("data") or {}
        images = detail_images(str(data.get("detailInfo") or ""))
        selected_url = normalize_url(str((old_ocr.get("selected") or {}).get("url") or ""))
        alternate = [item for item in images if item["url"] != selected_url]
        ocr_signal = int(bool(old_ocr.get("dimension_text")))
        if candidate_rule and confidence == "HIGH" and not requires_policy:
            action = "OPTION_AUTO_VERIFY"
        elif candidate_rule:
            action = "OPTION_POLICY_REVIEW"
        elif ocr_status == 502:
            action = "OCR_URL_RETRY"
        elif alternate:
            action = "OCR_NEXT_IMAGE"
        elif images:
            action = "OCR_RECHECK_IMAGE"
        else:
            action = "MANUAL_SOURCE_REQUIRED"

        status_counts[status] += 1
        action_counts[action] += 1
        inserts.append(
            (
                snapshot_id, assessed_at, 1, product_id, product_name, small,
                status, current["W"], current["D"], current["H"],
                "|".join(missing_axes),
                json.dumps(records_by_id.get(product_id, []), ensure_ascii=False),
                len(options), json.dumps(options, ensure_ascii=False),
                len(pairs), len(triples), len(codes), candidate_rule, confidence,
                json.dumps(candidate_dimensions, ensure_ascii=False), requires_policy,
                ocr_status, ocr_signal, len(images), len(alternate),
                selected_url or None, action, "대기", 0,
            )
        )

    create_schema(connection)
    placeholders = ",".join("?" for _ in range(29))
    with connection:
        connection.execute(f"UPDATE {TABLE} SET is_current=0 WHERE is_current=1")
        connection.executemany(
            f"INSERT INTO {TABLE} VALUES ({placeholders})", inserts
        )
    candidate_counts = {
        rule: count
        for rule, count in connection.execute(
            f"SELECT candidate_rule,COUNT(*) FROM {TABLE} "
            "WHERE is_current=1 AND candidate_rule IS NOT NULL GROUP BY candidate_rule"
        )
    }
    result = {
        "database": str(DB_PATH),
        "table": TABLE,
        "snapshot_id": snapshot_id,
        "assessed_at": assessed_at,
        "rows": len(inserts),
        "current_status_counts": dict(status_counts),
        "suggested_action_counts": dict(sorted(action_counts.items())),
        "candidate_rule_counts": candidate_counts,
        "views": [
            "vw_dimension_reinforcement_current_summary",
            "vw_dimension_reinforcement_option_candidates",
        ],
        "workbook_meta": meta,
        "excel_written": False,
        "notes": [
            "후보값은 아직 sources 원천 규격이나 고객용 엑셀에 병합하지 않았다.",
            "OPTION_POLICY_REVIEW는 카테고리 축/해당없음 정책 확인 후 적용한다.",
            "표준 침대 사이즈 코드는 실제 브랜드·모델 치수표 없이는 확정하지 않는다.",
        ],
    }
    connection.close()
    LATEST_OUTPUT.write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return result


if __name__ == "__main__":
    print(json.dumps(build_snapshot(), ensure_ascii=False, indent=2))
