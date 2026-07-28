from __future__ import annotations

import json
import math
import re
from statistics import median
from typing import Any


NUMBER = r"(?:\d{2,5}(?:[.,]\d+)?|\d{1,3}[.,]\d+)"
UNIT = r"(?:mm|cm|㎜|㎝)"
AXIS_BEFORE_RE = re.compile(
    rf"(?i)(?<![A-Za-z])(?P<axis>[WDHL])\s*[\(\[\]\):=._-]*\s*"
    rf"(?P<number>{NUMBER})\s*(?P<unit>{UNIT})?"
)
AXIS_AFTER_RE = re.compile(
    rf"(?i)(?P<number>{NUMBER})\s*[\(\[]\s*(?P<axis>[WDHL])\s*[\)\]]\s*"
    rf"(?P<unit>{UNIT})?"
)
KOREAN_BEFORE_RE = re.compile(
    rf"(?i)(?P<axis>가로|폭|너비|세로|깊이|높이|길이|지름|직경|두께)\s*"
    rf"[:=]?\s*(?P<number>{NUMBER})\s*(?P<unit>{UNIT})?"
)
KOREAN_AFTER_RE = re.compile(
    rf"(?i)(?P<number>{NUMBER})\s*(?P<unit>{UNIT})?\s*[\(\[]?\s*"
    rf"(?P<axis>가로|폭|너비|세로|깊이|높이|길이|지름|직경|두께)\s*[\)\]]?"
)
TRIPLE_RE = re.compile(
    rf"(?i)(?<!\d)(?P<a>{NUMBER})\s*(?:x|×|X|＊|\*|\)\()\s*"
    rf"(?P<b>{NUMBER})\s*(?:x|×|X|＊|\*|\)\()\s*(?P<c>{NUMBER})\s*"
    rf"(?P<unit>{UNIT})?"
)
PAIR_RE = re.compile(
    rf"(?i)(?<!\d)(?P<a>{NUMBER})\s*(?:x|×|X|＊|\*|\)\()\s*"
    rf"(?P<b>{NUMBER})\s*(?P<unit>{UNIT})?"
)
UNIT_RE = re.compile(rf"(?i){UNIT}")
BARE_NUMBER_RE = re.compile(rf"(?<!\d){NUMBER}(?!\d)")
SIZE_LABEL_RE = re.compile(
    r"(?i)(?:제품\s*(?:사이즈|크기|규격|치수)|사이즈|규격|치수|"
    r"product\s+(?:size|description)|\bsize\b|dimensions?|measurements?)"
)
WORD_DIMENSION_RE = re.compile(rf"(?i)^(?P<number>{NUMBER})(?P<unit>{UNIT})$")
WORD_NUMBER_RE = re.compile(rf"(?i)^(?P<number>{NUMBER})$")
WORD_UNIT_RE = re.compile(rf"(?i)^(?P<unit>{UNIT})$")

KOREAN_AXIS = {
    "가로": "W",
    "폭": "W",
    "너비": "W",
    "세로": "D",
    "깊이": "D",
    "높이": "H",
    "두께": "H",
    "길이": "L",
    "지름": "R",
    "직경": "R",
}
AXIS_ORDER = ("W", "D", "H", "L", "R")


def parse_number(raw: str) -> float:
    value = str(raw).strip()
    if "," in value:
        left, right = value.rsplit(",", 1)
        value = left + right if len(right) == 3 else left + "." + right
    return float(value)


def normalize_unit(raw: str) -> str:
    value = str(raw or "").casefold()
    if value in {"mm", "㎜"}:
        return "mm"
    if value in {"cm", "㎝"}:
        return "cm"
    return ""


def to_mm(value: float | None, unit: str) -> float | None:
    if value is None or not unit:
        return None
    result = value * 10 if unit == "cm" else value
    if result < 5 or result > 20000:
        return None
    return int(result) if float(result).is_integer() else round(result, 3)


def compact_number(value: float | None) -> str:
    if value is None:
        return ""
    return str(int(value)) if float(value).is_integer() else str(value)


def axis_signature(values: dict[str, float]) -> str:
    return ",".join(axis for axis in AXIS_ORDER if axis in values)


def line_chunks(text: str) -> list[str]:
    lines = [re.sub(r"\s+", " ", line).strip() for line in str(text or "").splitlines()]
    lines = [line for line in lines if line]
    result: list[str] = []
    for index in range(len(lines)):
        for width in (1, 2, 3):
            chunk = " ".join(lines[index : index + width]).strip()
            if chunk and chunk not in result:
                result.append(chunk)
    if len(str(text or "")) <= 1200:
        compact = re.sub(r"\s+", " ", str(text or "")).strip()
        if compact and compact not in result:
            result.append(compact)
    return result


def labeled_values(text: str) -> tuple[dict[str, float], dict[str, str]]:
    matches: list[tuple[int, str, float, str]] = []
    for pattern in (AXIS_BEFORE_RE, AXIS_AFTER_RE, KOREAN_BEFORE_RE, KOREAN_AFTER_RE):
        for match in pattern.finditer(text):
            raw_axis = match.group("axis")
            axis = raw_axis.upper() if len(raw_axis) == 1 else KOREAN_AXIS[raw_axis]
            try:
                number = parse_number(match.group("number"))
            except ValueError:
                continue
            if number < 5 or number > 20000:
                continue
            matches.append(
                (match.start(), axis, number, normalize_unit(match.group("unit") or ""))
            )
    matches.sort()
    values: dict[str, float] = {}
    units: dict[str, str] = {}
    for _, axis, number, unit in matches:
        if axis not in values:
            values[axis] = number
            units[axis] = unit
    return values, units


def common_unit(text: str, units: dict[str, str]) -> tuple[str, str]:
    explicit = {value for value in units.values() if value}
    if len(explicit) == 1:
        unit = next(iter(explicit))
        return unit, "ALL_OR_COMMON_UNIT_PRESENT"
    if len(explicit) > 1:
        return "", "MIXED_UNITS_REVIEW"
    found = {normalize_unit(value) for value in UNIT_RE.findall(text)}
    found.discard("")
    if len(found) == 1:
        return next(iter(found)), "COMMON_UNIT_PRESENT"
    return "", "UNIT_MISSING"


def observation(
    *,
    candidate_type: str,
    evidence: str,
    values: dict[str, float] | None = None,
    ordered: list[float] | None = None,
    unit: str = "",
    unit_status: str = "UNIT_MISSING",
    mapping_status: str,
    confidence: str,
    score: int,
    bbox: dict[str, Any] | None = None,
) -> dict[str, Any]:
    values = values or {}
    ordered = ordered or []
    resolved = {axis: to_mm(values.get(axis), unit) for axis in ("W", "D", "H", "L")}
    if mapping_status not in {"EXPLICIT_WDH_MAPPED", "EXPLICIT_PARTIAL_AXES"}:
        resolved = {axis: None for axis in ("W", "D", "H", "L")}
    raw_notation = " ".join(
        f"{axis}={compact_number(values[axis])}{unit}"
        for axis in AXIS_ORDER
        if axis in values
    )
    if not raw_notation and ordered:
        raw_notation = " × ".join(compact_number(value) for value in ordered) + unit
    return {
        "candidate_type": candidate_type,
        "raw_notation": raw_notation,
        "axis_signature": axis_signature(values),
        "unit_status": unit_status,
        "unit_text": unit,
        "w_raw": values.get("W"),
        "d_raw": values.get("D"),
        "h_raw": values.get("H"),
        "l_raw": values.get("L"),
        "r_raw": values.get("R"),
        "value_1_raw": ordered[0] if len(ordered) > 0 else None,
        "value_2_raw": ordered[1] if len(ordered) > 1 else None,
        "value_3_raw": ordered[2] if len(ordered) > 2 else None,
        "w_mm": resolved["W"],
        "d_mm": resolved["D"],
        "h_mm": resolved["H"],
        "l_mm": resolved["L"],
        "mapping_status": mapping_status,
        "confidence": confidence,
        "candidate_score": score,
        "evidence_text": re.sub(r"\s+", " ", evidence).strip()[:2000],
        "bbox_json": json.dumps(bbox or {}, ensure_ascii=False),
    }


def grouped_word_lines(words: list[dict[str, Any]]) -> list[list[dict[str, Any]]]:
    clean = [word for word in words if isinstance(word, dict) and str(word.get("text") or "").strip()]
    if not clean:
        return []
    heights = [float(word.get("height") or 0) for word in clean if float(word.get("height") or 0) > 0]
    tolerance = max(8.0, (median(heights) if heights else 16.0) * 0.8)
    result: list[list[dict[str, Any]]] = []
    for word in sorted(clean, key=lambda item: (float(item.get("y") or 0), float(item.get("x") or 0))):
        cy = float(word.get("y") or 0) + float(word.get("height") or 0) / 2
        target = None
        for line in reversed(result[-4:]):
            line_cy = sum(
                float(item.get("y") or 0) + float(item.get("height") or 0) / 2
                for item in line
            ) / len(line)
            if abs(cy - line_cy) <= tolerance:
                target = line
                break
        if target is None:
            target = []
            result.append(target)
        target.append(word)
    for line in result:
        line.sort(key=lambda item: float(item.get("x") or 0))
    return sorted(result, key=lambda line: min(float(item.get("y") or 0) for item in line))


def spatial_dimension_tokens(layout: dict[str, Any]) -> list[dict[str, Any]]:
    tokens: list[dict[str, Any]] = []
    for line in grouped_word_lines(layout.get("words") or []):
        index = 0
        while index < len(line):
            word = line[index]
            raw = re.sub(r"\s+", "", str(word.get("text") or ""))
            direct = WORD_DIMENSION_RE.fullmatch(raw)
            number_match = WORD_NUMBER_RE.fullmatch(raw)
            unit = ""
            number = ""
            end = index
            if direct:
                number, unit = direct.group("number"), normalize_unit(direct.group("unit"))
            elif number_match and index + 1 < len(line):
                next_word = line[index + 1]
                next_raw = re.sub(r"\s+", "", str(next_word.get("text") or ""))
                unit_match = WORD_UNIT_RE.fullmatch(next_raw)
                horizontal_gap = float(next_word.get("x") or 0) - (
                    float(word.get("x") or 0) + float(word.get("width") or 0)
                )
                if unit_match and horizontal_gap <= max(40.0, float(word.get("height") or 0) * 2):
                    number, unit, end = (
                        number_match.group("number"),
                        normalize_unit(unit_match.group("unit")),
                        index + 1,
                    )
            if number and unit:
                value = parse_number(number)
                value_mm = to_mm(value, unit)
                if value_mm is not None:
                    x1 = float(word.get("x") or 0)
                    y1 = float(word.get("y") or 0)
                    last = line[end]
                    x2 = float(last.get("x") or 0) + float(last.get("width") or 0)
                    y2 = max(
                        float(item.get("y") or 0) + float(item.get("height") or 0)
                        for item in line[index : end + 1]
                    )
                    tokens.append(
                        {
                            "raw": f"{compact_number(value)}{unit}",
                            "value": value,
                            "value_mm": value_mm,
                            "unit": unit,
                            "x": x1,
                            "y": y1,
                            "width": x2 - x1,
                            "height": y2 - y1,
                            "cx": (x1 + x2) / 2,
                            "cy": (y1 + y2) / 2,
                        }
                    )
            index = end + 1
    return tokens


def parse_layout(layout: dict[str, Any]) -> list[dict[str, Any]]:
    text = str(layout.get("text") or "")
    observations: list[dict[str, Any]] = []
    chunks = line_chunks(text)
    word_line_text = "\n".join(
        " ".join(str(word.get("text") or "") for word in line)
        for line in grouped_word_lines(layout.get("words") or [])
    )
    for chunk in line_chunks(word_line_text):
        if chunk not in chunks:
            chunks.append(chunk)

    for chunk in chunks:
        values, units = labeled_values(chunk)
        if not values:
            continue
        unit, unit_status = common_unit(chunk, units)
        axes = set(values)
        if {"W", "D", "H"}.issubset(axes):
            candidate_type = "EXPLICIT_WDH"
            mapping = "EXPLICIT_WDH_MAPPED"
            confidence = "HIGH" if unit else "MEDIUM_UNIT_MISSING"
            score = 100 if unit else 82
        elif {"L", "W", "H"}.issubset(axes):
            candidate_type = "EXPLICIT_LWH"
            mapping = "NONSTANDARD_AXIS_MAPPING_REQUIRED"
            confidence = "HIGH_LABELS_REVIEW_MAPPING"
            score = 88 if unit else 74
        elif len(axes) >= 2:
            candidate_type = "PARTIAL_LABELED_AXES"
            mapping = "EXPLICIT_PARTIAL_AXES"
            confidence = "MEDIUM_PARTIAL"
            score = 66 if unit else 56
        else:
            candidate_type = "SINGLE_LABELED_AXIS"
            mapping = "EXPLICIT_PARTIAL_AXES"
            confidence = "LOW_PARTIAL"
            score = 42 if unit else 32
        observations.append(
            observation(
                candidate_type=candidate_type,
                evidence=chunk,
                values=values,
                unit=unit,
                unit_status=unit_status,
                mapping_status=mapping,
                confidence=confidence,
                score=score,
            )
        )

    for chunk in chunks:
        for match in TRIPLE_RE.finditer(chunk):
            ordered = [parse_number(match.group(name)) for name in ("a", "b", "c")]
            unit = normalize_unit(match.group("unit") or "")
            observations.append(
                observation(
                    candidate_type="UNLABELED_TRIPLE",
                    evidence=match.group(0),
                    ordered=ordered,
                    unit=unit,
                    unit_status="COMMON_UNIT_PRESENT" if unit else "UNIT_MISSING",
                    mapping_status="ORDER_INFERENCE_REQUIRED",
                    confidence="MEDIUM_ORDER_REVIEW" if unit else "LOW_UNIT_AND_ORDER_REVIEW",
                    score=64 if unit else 48,
                )
            )
        for match in PAIR_RE.finditer(chunk):
            # A triple also contains two pairs; keep pairs only when the chunk
            # does not already carry a three-value separator pattern.
            if TRIPLE_RE.search(chunk):
                continue
            ordered = [parse_number(match.group(name)) for name in ("a", "b")]
            unit = normalize_unit(match.group("unit") or "")
            observations.append(
                observation(
                    candidate_type="UNLABELED_PAIR",
                    evidence=match.group(0),
                    ordered=ordered,
                    unit=unit,
                    unit_status="COMMON_UNIT_PRESENT" if unit else "UNIT_MISSING",
                    mapping_status="ONE_AXIS_OR_2D_PRODUCT_REVIEW",
                    confidence="LOW_PARTIAL",
                    score=38 if unit else 28,
                )
            )

    for chunk in chunks:
        if not SIZE_LABEL_RE.search(chunk):
            continue
        numbers = []
        for raw in BARE_NUMBER_RE.findall(chunk):
            try:
                value = parse_number(raw)
            except ValueError:
                continue
            if 5 <= value <= 20000:
                numbers.append(value)
        if numbers and not labeled_values(chunk)[0] and not TRIPLE_RE.search(chunk):
            unit_values = {normalize_unit(value) for value in UNIT_RE.findall(chunk)}
            unit_values.discard("")
            unit = next(iter(unit_values)) if len(unit_values) == 1 else ""
            observations.append(
                observation(
                    candidate_type="SIZE_LABEL_NUMERIC_CLUSTER",
                    evidence=chunk,
                    ordered=numbers[:3],
                    unit=unit,
                    unit_status="COMMON_UNIT_PRESENT" if unit else "UNIT_MISSING",
                    mapping_status="SPATIAL_OR_ORDER_REVIEW_REQUIRED",
                    confidence="LOW_NUMERIC_CLUSTER",
                    score=34 if len(numbers) >= 3 else 24,
                )
            )

    spatial = spatial_dimension_tokens(layout)
    if len(spatial) == 3:
        unit = spatial[0]["unit"] if len({item["unit"] for item in spatial}) == 1 else ""
        observations.append(
            observation(
                candidate_type="SPATIAL_THREE_DIMENSION_TOKENS",
                evidence=" ".join(item["raw"] for item in spatial),
                ordered=[item["value"] for item in spatial],
                unit=unit,
                unit_status="ALL_UNIT_PRESENT" if unit else "MIXED_UNITS_REVIEW",
                mapping_status="SPATIAL_MAPPING_REQUIRED",
                confidence="MEDIUM_SPATIAL_REVIEW",
                score=70,
                bbox={"tokens": spatial},
            )
        )

    dedup: dict[tuple[Any, ...], dict[str, Any]] = {}
    for item in observations:
        key = (
            item["candidate_type"],
            item["axis_signature"],
            item["w_raw"],
            item["d_raw"],
            item["h_raw"],
            item["l_raw"],
            item["value_1_raw"],
            item["value_2_raw"],
            item["value_3_raw"],
            item["unit_text"],
        )
        old = dedup.get(key)
        if old is None or (
            item["candidate_score"] > old["candidate_score"]
            or len(item["evidence_text"]) < len(old["evidence_text"])
        ):
            dedup[key] = item
    return sorted(
        dedup.values(),
        key=lambda item: (-item["candidate_score"], item["candidate_type"], item["raw_notation"]),
    )


if __name__ == "__main__":
    import argparse
    from pathlib import Path

    parser = argparse.ArgumentParser()
    parser.add_argument("layout_json")
    args = parser.parse_args()
    data = json.loads(Path(args.layout_json).read_text(encoding="utf-8-sig"))
    if isinstance(data, list):
        data = data[0]
    print(json.dumps(parse_layout(data), ensure_ascii=False, indent=2))
