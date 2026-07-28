from __future__ import annotations

import json
import re
import sqlite3
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import build_homestyle_bulk_workbook as workbook
from analyze_dimension_notations import (
    NUMBER,
    RANGE,
    UNIT,
    flat_category_sufficient,
    normalize,
    notation_flags,
    product_sources,
    strict_complete,
)
from bulk_homestyle_collect import DB_PATH, RUN_DIR, unpack


OUTPUT = RUN_DIR / "remaining_dimension_research.json"
XLSX_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def cell_value(cell: ET.Element) -> str | int | float | None:
    cell_type = cell.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.iter(XLSX_NS + "t"))
    value = cell.find(XLSX_NS + "v")
    if value is None:
        return None
    raw = value.text or ""
    if cell_type in ("str", "s"):
        return raw
    try:
        return float(raw) if "." in raw else int(raw)
    except ValueError:
        return raw


def column_number(reference: str) -> int:
    letters = re.match(r"[A-Z]+", reference).group(0)
    result = 0
    for letter in letters:
        result = result * 26 + ord(letter) - 64
    return result


def old_workbook_complete_ids() -> tuple[set[str], Path]:
    candidates = [
        path
        for path in Path(__file__).resolve().parent.glob("*.xlsx")
        if not path.name.startswith("~$")
        and "전체상품" in path.name
        and "요구필드" in path.name
    ]
    if not candidates:
        raise FileNotFoundError("기존 전체상품 결과 엑셀을 찾을 수 없습니다.")
    path = max(candidates, key=lambda item: item.stat().st_mtime)
    complete: set[str] = set()
    with ZipFile(path) as archive, archive.open("xl/worksheets/sheet2.xml") as stream:
        headers: dict[str, int] = {}
        for _, element in ET.iterparse(stream, events=("end",)):
            if element.tag != XLSX_NS + "row":
                continue
            row_number = int(element.get("r"))
            cells = {
                column_number(cell.get("r")): cell_value(cell)
                for cell in element.findall(XLSX_NS + "c")
            }
            if row_number == 1:
                headers = {str(value): column for column, value in cells.items()}
            else:
                product_id = str(cells.get(headers["상품 ID"]) or "")
                values = [
                    cells.get(headers["요청1_W (mm)"]),
                    cells.get(headers["요청1_D (mm)"]),
                    cells.get(headers["요청1_H (mm)"]),
                ]
                if product_id and all(isinstance(value, (int, float)) for value in values):
                    complete.add(product_id)
            element.clear()
    return complete, path


def short_context(text: str, pattern: str, before: int = 100, after: int = 260) -> str:
    match = re.search(pattern, text, re.I | re.S)
    if not match:
        return text[:500]
    return text[max(0, match.start() - before) : match.end() + after]


def residual_flags(
    product_name: str,
    categories: list[str],
    sources: list[tuple[str, str]],
    records: list[dict[str, Any]],
    ocr_status: int,
) -> set[str]:
    text = normalize("\n".join(value for _, value in sources))
    flags: set[str] = set()
    strict_ok, strict_rule, _ = strict_complete(sources)
    if strict_ok and "L-" in strict_rule:
        flags.add("L 축 해석 대기")
    if any(
        sum(item.get(axis) is not None for axis in ("w_mm", "d_mm", "h_mm")) >= 2
        for item in records
    ):
        flags.add("2축만 확보(비평면 카테고리 포함)")
    if re.search(
        rf"(?:DIA\.?|Ø|Φ|⌀|지름|직경|\bC)\s*[:=]?\s*{RANGE}\s*"
        rf"(?:x|×|\*)\s*(?:H\s*)?{RANGE}",
        text,
        re.I,
    ):
        flags.add("직경/C + 두번째축(높이 라벨 불완전)")
    if re.search(rf"(?<!\d){RANGE}\s*(?:x|×|\*)\s*H\s*{RANGE}", text, re.I):
        flags.add("무라벨 가로/직경 + H")
    if re.search(
        rf"{RANGE}\s*[WLDH]\s*(?:x|×|\*)\s*{RANGE}\s*[WLDH]",
        text,
        re.I,
    ):
        flags.add("축 후위표기")
    if re.search(
        rf"{RANGE}\s*(?:\)\s*\(|/|,|;)\s*{RANGE}\s*"
        rf"(?:\)\s*\(|/|,|;)\s*{RANGE}\s*{UNIT}?",
        text,
        re.I,
    ):
        flags.add("비표준 구분자 3축")
    if re.search(
        rf"{RANGE}\s*(?:x|×|\*)\s*{RANGE}\s*(?:x|×|\*)\s*{RANGE}",
        text,
        re.I,
    ):
        flags.add("무라벨 3축")
    if re.search(rf"{RANGE}\s*(?:x|×|\*)\s*{RANGE}", text, re.I):
        flags.add("무라벨/부분라벨 2축")
    # 상세 본문의 교환/반품 안내에 등장하는 일반적인 "구성품" 문구는
    # 실제 세트상품 근거가 아니므로 사용하지 않는다. 상품명과 카테고리처럼
    # 상품 단위 의미가 분명한 원천만 세트 판정에 사용한다.
    set_name_pattern = r"(?:세트|패키지|\bSET\b|2\s*IN\s*1|[+＋])"
    if (
        re.search(set_name_pattern, product_name, re.I)
        or "침대+매트리스" in categories
    ):
        flags.add("세트·구성품별 규격 필요")
    if any(category in {"침대+매트리스", "침대"} for category in categories) and re.search(
        r"(?:^|[\s(/_-])(SS|GSS|Q|K|LK|S|D)(?:$|[\s)/_-])",
        product_name + " " + text,
        re.I,
    ):
        flags.add("침대 표준 사이즈코드만 존재")
    if re.search(r"(?:SH|SD|AH|SP|THK|\bT\s*\d)", text, re.I):
        flags.add("보조축 존재")
    if re.search(r"상세\s*페이지\s*(?:참조|참고)", text):
        flags.add("상세페이지 참조")
    if not re.search(r"\d", text):
        flags.add("숫자 신호 없음")
    elif re.search(rf"{RANGE}\s*{UNIT}", text, re.I):
        flags.add("숫자·단위는 있으나 축 미확정")
    if ocr_status == 502:
        flags.add("OCR 이미지 다운로드 실패")
    elif ocr_status == 204:
        flags.add("OCR 대상 이미지 없음")
    elif ocr_status == 0:
        flags.add("OCR 미실행")
    return flags


def primary_reason(flags: set[str]) -> str:
    order = (
        "L 축 해석 대기",
        "침대 표준 사이즈코드만 존재",
        "세트·구성품별 규격 필요",
        "직경/C + 두번째축(높이 라벨 불완전)",
        "무라벨 가로/직경 + H",
        "비표준 구분자 3축",
        "축 후위표기",
        "2축만 확보(비평면 카테고리 포함)",
        "무라벨 3축",
        "무라벨/부분라벨 2축",
        "OCR 이미지 다운로드 실패",
        "OCR 대상 이미지 없음",
        "숫자 신호 없음",
        "상세페이지 참조",
        "숫자·단위는 있으나 축 미확정",
        "OCR 미실행",
    )
    return next((reason for reason in order if reason in flags), "기타")


def main() -> None:
    old_complete, workbook_path = old_workbook_complete_ids()
    connection = sqlite3.connect(DB_PATH)
    categories_by_scope = {
        row[0]: row[1] or ""
        for row in connection.execute("SELECT scope_id,small_name FROM categories")
    }
    rows = connection.execute(
        """
        SELECT p.product_id,p.category_scope_ids,s.goods_blob,s.html_blob,
               s.qna_blob,s.ocr_status,s.ocr_blob
        FROM products p JOIN sources s ON s.product_id=p.product_id
        WHERE s.goods_status=200 ORDER BY p.product_id
        """
    ).fetchall()
    current_complete: set[str] = set()
    flat_sufficient_ids: set[str] = set()
    residual_ids: set[str] = set()
    overlap_flags = Counter()
    primary = Counter()
    category = defaultdict(Counter)
    source_state = Counter()
    examples: dict[str, list[dict[str, Any]]] = defaultdict(list)
    product_flags: dict[str, list[str]] = {}

    for index, row in enumerate(rows, 1):
        product_id = row[0]
        categories = [
            categories_by_scope.get(scope, scope)
            for scope in json.loads(row[1] or "[]")
        ]
        data = (unpack(row[2]) or {}).get("data") or {}
        product = {
            "data": data,
            "html": unpack(row[3]) or {},
            "qna": unpack(row[4]) or {},
            "ocr": unpack(row[6]) or {},
        }
        sources = product_sources(product)
        records = workbook.dimension_records(sources)
        complete = any(
            item.get("w_mm") is not None
            and item.get("d_mm") is not None
            and item.get("h_mm") is not None
            for item in records
        )
        if complete:
            current_complete.add(product_id)
            continue
        flat_ok, _ = flat_category_sufficient(categories, records, sources)
        if flat_ok:
            flat_sufficient_ids.add(product_id)
            continue
        residual_ids.add(product_id)
        ocr_status = int(row[5] or 0)
        flags = residual_flags(
            str(data.get("productName") or ""), categories, sources, records, ocr_status
        )
        product_flags[product_id] = sorted(flags)
        reason = primary_reason(flags)
        primary[reason] += 1
        source_state[f"ocr_status:{ocr_status}"] += 1
        if product["ocr"].get("dimension_text"):
            source_state["ocr_dimension_text"] += 1
        if product["html"].get("dimension_signals"):
            source_state["pdp_html_dimension_signal"] += 1
        for flag in flags:
            overlap_flags[flag] += 1
        for category_name in categories or ["분류 미확보"]:
            category[category_name]["residual"] += 1
            category[category_name][f"reason:{reason}"] += 1
        if len(examples[reason]) < 15:
            joined = "\n".join(f"[{source}] {value}" for source, value in sources)
            examples[reason].append(
                {
                    "product_id": product_id,
                    "product_name": str(data.get("productName") or ""),
                    "categories": categories,
                    "flags": sorted(flags),
                    "text": short_context(
                        joined,
                        r"(?:\bL\s*\d|DIA|Ø|Φ|지름|직경|\b[WDH]\s*\d|"
                        r"가로|너비|폭|깊이|세로|높이|사이즈|규격|치수)",
                    ),
                }
            )
        if index % 1000 == 0:
            print(f"scanned={index}/{len(rows)}", flush=True)

    old_lost = old_complete - current_complete
    old_gained = current_complete - old_complete
    lost_flags = Counter(
        flag for product_id in old_lost for flag in product_flags.get(product_id, [])
    )
    category_rows = [
        {"category": name, **dict(counts)} for name, counts in category.items()
    ]
    category_rows.sort(key=lambda item: (-item["residual"], item["category"]))
    result = {
        "baseline_workbook": str(workbook_path),
        "old_complete": len(old_complete),
        "current_complete_after_confirmed_regex_and_L_defer": len(current_complete),
        "flat_2d_category_sufficient": len(flat_sufficient_ids),
        "effective_complete_or_category_sufficient": len(current_complete | flat_sufficient_ids),
        "residual": len(residual_ids),
        "old_to_current_comparison": {
            "gained": len(old_gained),
            "lost": len(old_lost),
            "lost_product_flags": dict(lost_flags.most_common()),
        },
        "residual_primary_reason": dict(primary.most_common()),
        "residual_overlapping_flags": dict(overlap_flags.most_common()),
        "residual_source_state": dict(source_state.most_common()),
        "categories": category_rows,
        "examples": dict(examples),
        "product_flags": product_flags,
        "notes": [
            "L 축은 자동 정규화하지 않고 잔여 집합에 유지했다.",
            "러그/액자/인테리어포스터의 2축 충족 상품은 잔여에서 제외했다.",
            "원인 플래그는 중복될 수 있고 primary reason만 상호 배타적이다.",
            "이 파일은 연구용이며 고객용 엑셀은 아직 재생성하지 않았다.",
        ],
    }
    OUTPUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    connection.close()
    print(
        json.dumps(
            {
                key: result[key]
                for key in (
                    "old_complete",
                    "current_complete_after_confirmed_regex_and_L_defer",
                    "flat_2d_category_sufficient",
                    "effective_complete_or_category_sufficient",
                    "residual",
                    "old_to_current_comparison",
                    "residual_primary_reason",
                    "residual_overlapping_flags",
                    "residual_source_state",
                )
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    print(f"output={OUTPUT}")


if __name__ == "__main__":
    main()
