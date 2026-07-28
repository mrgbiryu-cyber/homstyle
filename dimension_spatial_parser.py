from __future__ import annotations

import itertools
import json
import re
from typing import Any


DIMENSION_TOKEN_RE = re.compile(
    r"^(?P<prefix>[ØΦ⌀])?\s*(?P<number>\d{2,5}(?:[.,]\d+)?)\s*(?P<unit>mm|cm)$",
    re.I,
)
TABLE_LIKE_TOKENS = (
    "테이블", "식탁", "책상", "협탁", "화장대", "데스크", "table", "desk"
)


def to_mm(number: str, unit: str) -> int | float:
    value = float(number.replace(",", "."))
    if unit.casefold() == "cm":
        value *= 10
    return int(value) if value.is_integer() else value


def dimension_words(layout: dict[str, Any]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for index, word in enumerate(layout.get("words") or []):
        if not isinstance(word, dict):
            continue
        raw = re.sub(r"\s+", "", str(word.get("text") or ""))
        match = DIMENSION_TOKEN_RE.fullmatch(raw)
        if not match:
            continue
        number = match.group("number")
        prefix = match.group("prefix") or ""
        leading_zero_diameter = False
        # Windows OCR frequently reads Ø300mm as 0300mm. Keep this as a
        # diameter candidate only when the leading zero precedes 2–4 digits.
        if not prefix and number.startswith("0") and len(number.split(".", 1)[0]) >= 4:
            number = number[1:]
            leading_zero_diameter = True
        value_mm = to_mm(number, match.group("unit"))
        if value_mm < 20 or value_mm > 10000:
            continue
        x = float(word.get("x") or 0)
        y = float(word.get("y") or 0)
        width = float(word.get("width") or 0)
        height = float(word.get("height") or 0)
        result.append(
            {
                "index": index,
                "raw": raw,
                "value_mm": value_mm,
                "is_diameter": bool(prefix) or leading_zero_diameter,
                "diameter_basis": (
                    "EXPLICIT_SYMBOL" if prefix else "OCR_LEADING_ZERO" if leading_zero_diameter else ""
                ),
                "x": x,
                "y": y,
                "width": width,
                "height": height,
                "cx": x + width / 2,
                "cy": y + height / 2,
            }
        )
    return result


def is_table_like(product_name: str, small_category: str) -> bool:
    text = f"{product_name} {small_category}".casefold()
    return any(token.casefold() in text for token in TABLE_LIKE_TOKENS)


def spatial_callout_candidate(
    layout: dict[str, Any],
    *,
    product_name: str = "",
    small_category: str = "",
) -> dict[str, Any] | None:
    """Recover a diagram callout shaped as two top dimensions plus one height.

    This is intentionally narrower than general visual dimension inference.
    The high-confidence path requires a table-like category, two dimension
    labels on the same upper row, and a third label lower and to the right.
    """
    image_width = float(layout.get("scaled_width") or 0)
    image_height = float(layout.get("scaled_height") or 0)
    if image_width <= 0 or image_height <= 0:
        return None
    words = dimension_words(layout)
    regular = [word for word in words if not word["is_diameter"]]
    diameters = [word for word in words if word["is_diameter"]]
    if len(regular) < 3:
        return None

    candidates: list[tuple[tuple[float, ...], dict[str, Any]]] = []
    for first, second in itertools.combinations(regular, 2):
        row_delta = abs(first["cy"] - second["cy"])
        if row_delta > max(60.0, image_height * 0.075):
            continue
        top_y = (first["cy"] + second["cy"]) / 2
        if top_y > image_height * 0.48:
            continue
        horizontal_gap = abs(first["cx"] - second["cx"])
        if horizontal_gap < image_width * 0.08:
            continue
        for height_word in regular:
            if height_word is first or height_word is second:
                continue
            if height_word["cy"] < top_y + image_height * 0.10:
                continue
            if height_word["cy"] > image_height * 0.72:
                continue
            if height_word["cx"] < max(first["cx"], second["cx"]) - image_width * 0.04:
                continue
            pair_values = [first["value_mm"], second["value_mm"]]
            w_mm = max(pair_values)
            d_mm = min(pair_values)
            h_mm = height_word["value_mm"]
            if min(w_mm, d_mm, h_mm) < 50:
                continue
            extra_regular = [
                word for word in regular
                if word not in (first, second, height_word)
                and word["cy"] <= image_height * 0.72
            ]
            table_like = is_table_like(product_name, small_category)
            confidence = (
                "HIGH_SPATIAL_TABLE_CALLOUT"
                if table_like and not extra_regular
                else "REVIEW_SPATIAL_CALLOUT"
            )
            score = (
                0.0 if table_like else 1.0,
                float(len(extra_regular)),
                row_delta / image_height,
                -height_word["cx"] / image_width,
                top_y / image_height,
            )
            candidates.append(
                (
                    score,
                    {
                        "pattern_type": "DIAGRAM_CALLOUT_3_AXIS",
                        "confidence": confidence,
                        "w_mm": w_mm,
                        "d_mm": d_mm,
                        "h_mm": h_mm,
                        "top_dimension_words": [first, second],
                        "height_word": height_word,
                        "diameter_words": diameters,
                        "extra_regular_words": extra_regular,
                        "image_width": image_width,
                        "image_height": image_height,
                        "product_name": product_name,
                        "small_category": small_category,
                    },
                )
            )
    if not candidates:
        return None
    candidates.sort(key=lambda item: item[0])
    if len(candidates) > 1 and candidates[0][0] == candidates[1][0]:
        return None
    result = candidates[0][1]
    result["evidence_text"] = (
        f"도면 좌표 검증 W={result['w_mm']} mm D={result['d_mm']} mm "
        f"H={result['h_mm']} mm; OCR={layout.get('text') or ''}"
    )
    return result


if __name__ == "__main__":
    import argparse
    from pathlib import Path

    parser = argparse.ArgumentParser()
    parser.add_argument("layout_json")
    parser.add_argument("--product-name", default="")
    parser.add_argument("--small-category", default="")
    args = parser.parse_args()
    data = json.loads(Path(args.layout_json).read_text(encoding="utf-8-sig"))
    if isinstance(data, list):
        data = data[0]
    print(
        json.dumps(
            spatial_callout_candidate(
                data,
                product_name=args.product_name,
                small_category=args.small_category,
            ),
            ensure_ascii=False,
            indent=2,
        )
    )
