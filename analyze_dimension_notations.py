from __future__ import annotations

import html
import json
import re
import sqlite3
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

import build_homestyle_bulk_workbook as workbook
from bulk_homestyle_collect import DB_PATH, RUN_DIR, unpack


OUTPUT = RUN_DIR / "dimension_notation_analysis.json"
NUMBER = r"\d+(?:[.,]\d+)?"
RANGE = rf"{NUMBER}(?:\s*[-~]\s*{NUMBER})?"
UNIT = r"(?:mm|cm|㎜|인치|inch(?:es)?)"


def product_sources(product: dict[str, Any]) -> list[tuple[str, str]]:
    data = product["data"]
    notifications = workbook.notification_items(data)
    colors = workbook.joined(
        workbook.notification_values(notifications, "색상", "컬러"), ""
    )
    groups = workbook.option_groups(data, colors)
    texts = [
        ("상품정보고시", value)
        for value in workbook.notification_values(
            notifications, "크기", "치수", "규격", "사이즈"
        )
    ]
    for group in groups:
        if group["style"] == "사이즈":
            texts.extend(("상품 옵션", item["name"]) for item in group["items"])
    texts.extend(
        ("PDP HTML", value)
        for value in (product["html"].get("dimension_signals") or [])
    )
    texts.extend(
        ("상품 상세 HTML", value)
        for value in workbook.dimension_keyword_snippets(data.get("detailInfo"))
    )
    texts.append(("상품명", str(data.get("productName") or "")))
    qna_dimension = workbook.qna_signal(
        product["qna"],
        product["html"],
        ("크기", "사이즈", "가로", "세로", "높이", "폭", "치수"),
    )
    if qna_dimension:
        texts.append(("FAQ/Q&A", qna_dimension))
    texts.extend(workbook.verified_dimension_texts(product["ocr"]))
    ocr_text = str(product["ocr"].get("dimension_text") or "")
    if ocr_text:
        texts.append(("상세 이미지 OCR", ocr_text))

    # The current extractor does not trigger on an L-only axis. Add short
    # windows solely for diagnosis so we can measure those missed notations.
    detail_text = workbook.clean_text(data.get("detailInfo"), 160000)
    extra_patterns = (
        rf"(?:\bL\s*[:=]?\s*{RANGE}|{RANGE}\s*\(?L\)?)(?=.{{0,100}}\b[WHDT]\b)",
        rf"(?:DIA\.?|[ØΦ⌀])\s*[:=]?\s*{RANGE}",
        rf"{RANGE}\s*(?:x|×|\*)\s*{RANGE}\s*(?:x|×|\*)\s*{RANGE}\s*{UNIT}",
    )
    for pattern in extra_patterns:
        for match in re.finditer(pattern, detail_text, re.I):
            snippet = detail_text[max(0, match.start() - 100) : match.end() + 250]
            if snippet and snippet not in [value for _, value in texts]:
                texts.append(("상품 상세 HTML(L/직경 보강 탐색)", snippet))
            if sum(source == "상품 상세 HTML(L/직경 보강 탐색)" for source, _ in texts) >= 10:
                break
    return [(source, value) for source, value in texts if str(value).strip()]


def normalize(text: str) -> str:
    return (
        html.unescape(str(text or ""))
        .replace("×", "x")
        .replace("✕", "x")
        .replace("㎜", "mm")
        .replace("Φ", "Ø")
        .replace("⌀", "Ø")
    )


def notation_flags(text: str) -> set[str]:
    t = normalize(text)
    flags: set[str] = set()
    if not re.search(r"\d", t):
        flags.add("숫자 없음")
        return flags
    if re.search(rf"\bL\s*[:=]?\s*{RANGE}.{{0,90}}\bW\s*[:=]?\s*{RANGE}.{{0,90}}\bH\s*[:=]?\s*{RANGE}", t, re.I | re.S):
        flags.add("L-W-H")
    if re.search(rf"\bL\s*[:=]?\s*{RANGE}", t, re.I):
        flags.add("L 축 포함")
    if all(re.search(rf"\b{axis}\s*[:=]?\s*{RANGE}", t, re.I) for axis in "WDH"):
        flags.add("W-D-H 변형")
    if re.search(r"(?:가로|너비|폭).{0,100}(?:깊이|세로).{0,100}높이", t, re.S):
        flags.add("한글 3축")
    if re.search(rf"(?:DIA\.?|Ø|지름|직경)\s*[:=]?\s*{RANGE}", t, re.I):
        flags.add("직경/DIA/Ø")
    if re.search(rf"\b(?:SP|T|THK|SH|SD|AH|C)\s*[:=]?\s*{RANGE}", t, re.I):
        flags.add("보조축 SP/T/SH/C")
    triples = re.findall(
        rf"{RANGE}\s*(?:x|×|\*)\s*{RANGE}\s*(?:x|×|\*)\s*{RANGE}",
        t,
        re.I,
    )
    pairs = re.findall(rf"{RANGE}\s*(?:x|×|\*)\s*{RANGE}", t, re.I)
    if triples:
        flags.add("숫자 3축 곱셈표기")
    elif pairs:
        flags.add("숫자 2축 곱셈표기")
    if re.search(rf"{RANGE}\s*{UNIT}", t, re.I):
        flags.add("숫자+단위")
    if re.search(rf"{NUMBER}\s*[-~]\s*{NUMBER}", t):
        flags.add("범위값")
    if re.search(r"상세\s*페이지\s*참조|상세페이지\s*참조", t):
        flags.add("상세페이지 참조")
    return flags


def num(value: str) -> float:
    values = re.findall(r"\d+(?:[.,]\d+)?", value)
    number = float(values[-1].replace(",", ""))
    return number


def unit_factor(text: str, values: list[float]) -> int:
    if re.search(r"\bcm\b", text, re.I):
        return 10
    if re.search(r"(?:inch|인치)", text, re.I):
        return 25.4
    if re.search(r"(?:mm|㎜)", text, re.I):
        return 1
    return 10 if values and max(values) <= 300 else 1


def experimental_complete(text: str) -> tuple[bool, str]:
    """High-confidence completeness test; values are not persisted."""
    t = normalize(text)
    # Arbitrary-order Latin axes. L-W-H means length-width-height, so L maps
    # to output W and the source W maps to output D.
    axis_matches = re.findall(
        rf"(?<![A-Z])(?:\(?\s*)(L|W|D|H|SP|T|THK)\s*\)?\s*[:=]?\s*({RANGE})",
        t,
        re.I,
    )
    axes: dict[str, float] = {}
    for axis, value in axis_matches:
        axes.setdefault(axis.upper(), num(value))
    if {"L", "W", "H"}.issubset(axes) and "D" not in axes:
        return True, "L-W-H를 W-D-H로 변환"
    if {"W", "D", "H"}.issubset(axes):
        return True, "임의 순서 W-D-H"
    if {"L", "D", "H"}.issubset(axes) and "W" not in axes:
        return True, "L-D-H에서 L을 W로 변환"
    if {"W", "H"}.issubset(axes) and any(k in axes for k in ("SP", "T", "THK")):
        return True, "SP/T/THK를 깊이로 변환"

    # Diameter and height provide all three bounding-box axes.
    diameter = re.search(rf"(?:DIA\.?|Ø|지름|직경)\s*[:=]?\s*({RANGE})", t, re.I)
    height = re.search(rf"(?:\bH|높이)\s*[:=]?\s*({RANGE})", t, re.I)
    if diameter and height:
        return True, "직경을 W/D로 복제 + H"

    # Korean semantic axes in any common order.
    kr_w = re.search(rf"(?:가로|너비|폭)\s*[:=]?\s*({RANGE})", t)
    kr_d = re.search(rf"(?:깊이|세로)\s*[:=]?\s*({RANGE})", t)
    kr_h = re.search(rf"높이\s*[:=]?\s*({RANGE})", t)
    if kr_w and kr_d and kr_h:
        return True, "한글 W-D-H"
    return False, ""


def strict_complete(sources: list[tuple[str, str]]) -> tuple[bool, str, str]:
    """Require all main axes to occur in one short source window."""
    for source, raw in sources:
        text = normalize(raw)
        latin = []
        for match in re.finditer(
            rf"(?<![A-Z])\(?\s*(L|W|D|H|SP|T|THK)\s*\)?\s*[:=]?\s*"
            rf"\(?\s*({RANGE})",
            text,
            re.I,
        ):
            latin.append((match.start(), match.end(), match.group(1).upper(), match.group(2)))
        for match in re.finditer(
            rf"({RANGE})\s*\(\s*(L|W|D|H)\s*\)", text, re.I
        ):
            latin.append((match.start(), match.end(), match.group(2).upper(), match.group(1)))
        latin.sort()
        for start_index in range(len(latin)):
            cluster = []
            start = latin[start_index][0]
            for item in latin[start_index:]:
                if item[1] - start > 180:
                    break
                cluster.append(item)
            axes = {item[2] for item in cluster}
            rule = ""
            if {"W", "D", "H"}.issubset(axes):
                rule = "근접 W-D-H"
            elif {"L", "W", "H"}.issubset(axes) and "D" not in axes:
                rule = "근접 L-W-H"
            elif {"L", "D", "H"}.issubset(axes) and "W" not in axes:
                rule = "근접 L-D-H"
            elif {"W", "H"}.issubset(axes) and axes.intersection({"SP", "T", "THK"}):
                rule = "근접 W-H-SP/T"
            if rule:
                context = text[max(0, start - 60) : min(len(text), cluster[-1][1] + 80)]
                if re.search(r"(?:mm|cm|㎜|inch|인치|\bx\b|×|\*)", context, re.I):
                    return True, rule, f"[{source}] {context}"

        diameter = list(
            re.finditer(rf"(?:DIA\.?|Ø|지름|직경)\s*[:=]?\s*\(?\s*({RANGE})", text, re.I)
        )
        heights = list(
            re.finditer(rf"(?:\bH|높이)\s*[:=]?\s*\(?\s*({RANGE})", text, re.I)
        )
        for dia in diameter:
            for height in heights:
                if abs(dia.start() - height.start()) <= 140:
                    start, end = min(dia.start(), height.start()), max(dia.end(), height.end())
                    context = text[max(0, start - 60) : min(len(text), end + 80)]
                    if re.search(r"(?:mm|cm|㎜|inch|인치|\bx\b|×|\*)", context, re.I):
                        return True, "근접 직경+H", f"[{source}] {context}"

        korean = []
        for axis, words in (
            ("W", "가로|너비|폭"),
            ("D", "깊이|세로"),
            ("H", "높이"),
        ):
            for match in re.finditer(
                rf"(?:{words})\s*[:=]?\s*\(?\s*({RANGE})", text
            ):
                korean.append((match.start(), match.end(), axis, match.group(1)))
        korean.sort()
        for start_index in range(len(korean)):
            cluster = []
            start = korean[start_index][0]
            for item in korean[start_index:]:
                if item[1] - start > 180:
                    break
                cluster.append(item)
            if {item[2] for item in cluster} >= {"W", "D", "H"}:
                context = text[max(0, start - 60) : min(len(text), cluster[-1][1] + 80)]
                if re.search(r"(?:mm|cm|㎜|inch|인치|\bx\b|×|\*)", context, re.I):
                    return True, "근접 한글 W-D-H", f"[{source}] {context}"
    return False, "", ""


def flat_category_sufficient(
    categories: list[str],
    records: list[dict[str, Any]],
    sources: list[tuple[str, str]],
) -> tuple[bool, str]:
    flat = {"러그", "액자", "인테리어포스터"}
    matched = flat.intersection(categories)
    if not matched:
        return False, ""
    if any(
        sum(item.get(axis) is not None for axis in ("w_mm", "d_mm", "h_mm")) >= 2
        for item in records
    ):
        return True, "기존 2축 숫자"
    pair = re.compile(
        rf"(?<!\d)({RANGE})\s*(?:mm|cm|㎜|inch|인치)?\s*"
        rf"(?:x|×|\*)\s*({RANGE})\s*(?:mm|cm|㎜|inch|인치)",
        re.I,
    )
    for source, raw in sources:
        if pair.search(normalize(raw)):
            return True, f"2축 표기 추가 파싱({source})"
    return False, ""


def main() -> None:
    connection = sqlite3.connect(DB_PATH)
    category_by_scope = {
        row[0]: {"large": row[1] or "", "mid": row[2] or "", "small": row[3] or ""}
        for row in connection.execute(
            "SELECT scope_id,large_name,mid_name,small_name FROM categories"
        )
    }
    rows = connection.execute(
        """
        SELECT p.product_id,p.category_scope_ids,s.goods_blob,s.html_blob,
               s.qna_blob,s.ocr_blob
        FROM products p JOIN sources s ON s.product_id=p.product_id
        WHERE s.goods_status=200 ORDER BY p.product_id
        """
    ).fetchall()
    overall = Counter()
    flags_counter = Counter()
    exclusive_counter = Counter()
    experimental = Counter()
    strict = Counter()
    flat_sufficient = Counter()
    recovery_union = Counter()
    examples: dict[str, list[dict[str, Any]]] = defaultdict(list)
    category_stats: dict[str, Counter] = defaultdict(Counter)

    for index, row in enumerate(rows, 1):
        product_id = row[0]
        scopes = json.loads(row[1] or "[]")
        data = (unpack(row[2]) or {}).get("data") or {}
        product = {
            "data": data,
            "html": unpack(row[3]) or {},
            "qna": unpack(row[4]) or {},
            "ocr": unpack(row[5]) or {},
        }
        sources = product_sources(product)
        records = workbook.dimension_records(sources)
        complete = any(
            item.get("w_mm") is not None
            and item.get("d_mm") is not None
            and item.get("h_mm") is not None
            for item in records
        )
        status = "complete" if complete else "partial" if records else "missing"
        overall[status] += 1
        categories = [
            category_by_scope.get(scope, {}).get("small") or scope for scope in scopes
        ] or ["분류 미확보"]
        for category in categories:
            category_stats[category]["total"] += 1
            category_stats[category][status] += 1
        if complete:
            continue

        joined = "\n".join(f"[{source}] {value}" for source, value in sources)
        flags = notation_flags(joined)
        if flags == {"숫자 없음"}:
            exclusive = "숫자 없음"
        elif "L-W-H" in flags:
            exclusive = "L-W-H"
        elif "직경/DIA/Ø" in flags:
            exclusive = "직경/DIA/Ø"
        elif "한글 3축" in flags:
            exclusive = "한글 3축"
        elif "보조축 SP/T/SH/C" in flags:
            exclusive = "보조축 SP/T/SH/C"
        elif "숫자 3축 곱셈표기" in flags:
            exclusive = "숫자 3축 곱셈표기"
        elif "숫자 2축 곱셈표기" in flags:
            exclusive = "숫자 2축 곱셈표기"
        elif "숫자+단위" in flags:
            exclusive = "숫자+단위"
        else:
            exclusive = "기타 숫자"
        exclusive_counter[exclusive] += 1
        for flag in flags:
            flags_counter[flag] += 1
            for category in categories:
                category_stats[category][f"flag:{flag}"] += 1
            if len(examples[flag]) < 12:
                examples[flag].append(
                    {
                        "product_id": product_id,
                        "product_name": str(data.get("productName") or ""),
                        "categories": categories,
                        "status": status,
                        "text": joined[:1800],
                    }
                )
        recovered, rule = experimental_complete(joined)
        if recovered:
            experimental["high_confidence_complete"] += 1
            experimental[f"rule:{rule}"] += 1
            for category in categories:
                category_stats[category]["experimental_complete"] += 1
            if len(examples[f"실험:{rule}"]) < 20:
                examples[f"실험:{rule}"].append(
                    {
                        "product_id": product_id,
                        "product_name": str(data.get("productName") or ""),
                        "categories": categories,
                        "status": status,
                        "text": joined[:1800],
                    }
                )
        strict_recovered, strict_rule, strict_context = strict_complete(sources)
        if strict_recovered:
            strict["high_confidence_complete"] += 1
            strict[f"rule:{strict_rule}"] += 1
            strict_source = strict_context.split("]", 1)[0].lstrip("[")
            strict[f"source:{strict_source}"] += 1
            for category in categories:
                category_stats[category]["strict_complete"] += 1
            if len(examples[f"엄격:{strict_rule}"]) < 20:
                examples[f"엄격:{strict_rule}"].append(
                    {
                        "product_id": product_id,
                        "product_name": str(data.get("productName") or ""),
                        "categories": categories,
                        "status": status,
                        "context": strict_context,
                    }
                )
        flat_ok, flat_rule = flat_category_sufficient(categories, records, sources)
        if flat_ok:
            flat_sufficient["category_sufficient"] += 1
            flat_sufficient[f"rule:{flat_rule}"] += 1
            for category in categories:
                if category in {"러그", "액자", "인테리어포스터"}:
                    category_stats[category]["flat_2d_sufficient"] += 1
        if strict_recovered or flat_ok:
            recovery_union["unique_additional_complete_or_sufficient"] += 1
            if strict_recovered and flat_ok:
                recovery_union["overlap"] += 1
        if index % 1000 == 0:
            print(f"scanned={index}/{len(rows)}", flush=True)

    categories_output = []
    for category, counts in category_stats.items():
        incomplete = counts["partial"] + counts["missing"]
        categories_output.append(
            {
                "category": category,
                **dict(counts),
                "incomplete": incomplete,
                "incomplete_rate": round(incomplete / counts["total"], 4)
                if counts["total"]
                else 0,
            }
        )
    categories_output.sort(key=lambda item: (-item["incomplete"], item["category"]))
    result = {
        "overall": dict(overall),
        "incomplete": overall["partial"] + overall["missing"],
        "overlapping_notation_flags": dict(flags_counter.most_common()),
        "exclusive_primary_reason": dict(exclusive_counter.most_common()),
        "experimental_high_confidence": dict(experimental),
        "strict_same_source_short_window": dict(strict),
        "flat_category_two_axis_sufficient": dict(flat_sufficient),
        "combined_unique_recovery": dict(recovery_union),
        "categories": categories_output,
        "examples": dict(examples),
        "notes": [
            "표기 유형 수치는 서로 중복될 수 있다.",
            "실험 파서는 숫자를 엑셀에 쓰지 않고 고신뢰 추가 회수 가능 건만 계산한다.",
            "L-W-H는 L→출력 W, 원문의 W→출력 D, H→출력 H로 해석했다.",
            "2D 상품과 직경 단독 상품은 W-D-H 완전 확보로 과대평가하지 않았다.",
        ],
    }
    OUTPUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    connection.close()
    print(json.dumps({k: result[k] for k in ("overall", "incomplete", "overlapping_notation_flags", "exclusive_primary_reason", "experimental_high_confidence", "strict_same_source_short_window", "flat_category_two_axis_sufficient", "combined_unique_recovery")}, ensure_ascii=False, indent=2))
    print(f"output={OUTPUT}")


if __name__ == "__main__":
    main()
