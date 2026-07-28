"""Audit explicit W/D/H expressions found in spatial-diagram layout OCR.

This is a read-only staging step.  It deliberately does not update the source
database: products with set/option/common-lineup evidence remain review-only.
"""

from __future__ import annotations

import argparse
import json
import re
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

from build_homestyle_bulk_workbook import dimension_records


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def complete_triples(text: str, evidence: str) -> list[tuple[int | float, int | float, int | float]]:
    triples: list[tuple[int | float, int | float, int | float]] = []
    for record in dimension_records([(evidence, text)]):
        values = (record.get("w_mm"), record.get("d_mm"), record.get("h_mm"))
        if all(value is not None for value in values) and values not in triples:
            triples.append(values)  # type: ignore[arg-type]
    return triples


def strong_axis_triples(text: str) -> list[dict[str, Any]]:
    """Return only triples whose three values each retain an explicit W/D/H label."""
    pattern = re.compile(
        r"(?<![A-Z])\(?\s*([WDH])\s*\)?\s*[:=]?\s*\(?\s*"
        r"(\d[\d,.]*)(?:\s*[-~]\s*(\d[\d,.]*))?\s*\)?\s*"
        r"\(?\s*(mm|cm)?\s*\)?",
        flags=re.I,
    )
    matches = list(pattern.finditer(text))
    found: list[dict[str, Any]] = []
    seen: set[tuple[int | float, int | float, int | float]] = set()
    for start_index, first in enumerate(matches):
        cluster = []
        axes_seen: set[str] = set()
        for match in matches[start_index:]:
            if match.end() - first.start() > 180:
                break
            axis = match.group(1).upper()
            if axis in axes_seen:
                break
            axes_seen.add(axis)
            cluster.append(match)
            if axes_seen == {"W", "D", "H"}:
                break
        if axes_seen != {"W", "D", "H"}:
            continue
        context = text[max(0, first.start() - 45) : cluster[-1].end() + 45]
        explicit_units = [match.group(4) for match in cluster if match.group(4)]
        raw_values = [match.group(3) or match.group(2) for match in cluster]
        if explicit_units:
            common_unit = explicit_units[-1].lower()
        else:
            numeric = [float(value.replace(",", "").rstrip(".")) for value in raw_values]
            common_unit = "cm" if numeric and max(numeric) <= 300 else "mm"
        mapped: dict[str, int | float] = {}
        for match in cluster:
            raw_value = (match.group(3) or match.group(2)).replace(",", "").rstrip(".")
            value: int | float = float(raw_value)
            unit = (match.group(4) or common_unit).lower()
            if unit == "cm":
                value *= 10
            if float(value).is_integer():
                value = int(value)
            mapped[match.group(1).upper()] = value
        triple = (mapped["W"], mapped["D"], mapped["H"])
        if triple in seen:
            continue
        seen.add(triple)
        lower_context = context.lower()
        found.append(
            {
                "triple": list(triple),
                "context": context,
                "packaging_context": any(token in lower_context for token in ("포장", "package", "packing")),
            }
        )
    return found


def strict_expression_triples(text: str) -> list[dict[str, Any]]:
    """Parse only compact, ordered W-separator-D-separator-H expressions."""
    value = r"(\d[\d,.]*)(?:\s*[-~]\s*(\d[\d,.]*))?"
    unit = r"\s*\(?\s*(mm|cm)?\s*\)?"
    separator = r"\s*(?:[xX*+×'·,]|\)\s*\()+\s*"
    pattern = re.compile(
        rf"\(?\s*W\s*\)?\s*[:=]?\s*{value}{unit}{separator}"
        rf"\(?\s*D\s*\)?\s*[:=]?\s*{value}{unit}{separator}"
        rf"\(?\s*H\s*\)?\s*[:=]?\s*{value}{unit}",
        flags=re.I,
    )
    found: list[dict[str, Any]] = []
    seen: set[tuple[int | float, int | float, int | float]] = set()
    for match in pattern.finditer(text):
        raw_values = [match.group(2) or match.group(1), match.group(5) or match.group(4), match.group(8) or match.group(7)]
        units = [match.group(3), match.group(6), match.group(9)]
        explicit_unit = next((item.lower() for item in reversed(units) if item), "")
        if explicit_unit:
            common_unit = explicit_unit
        else:
            numeric = [float(item.replace(",", "").rstrip(".")) for item in raw_values]
            common_unit = "cm" if numeric and max(numeric) <= 300 else "mm"
        converted: list[int | float] = []
        for raw_value, raw_unit in zip(raw_values, units):
            number: int | float = float(raw_value.replace(",", "").rstrip("."))
            if (raw_unit or common_unit).lower() == "cm":
                number *= 10
            if float(number).is_integer():
                number = int(number)
            converted.append(number)
        triple = tuple(converted)
        if triple in seen:
            continue
        seen.add(triple)
        context = text[max(0, match.start() - 70) : match.end() + 70]
        lower_context = context.lower()
        found.append(
            {
                "triple": list(triple),
                "expression": match.group(0),
                "context": context,
                "non_product_context": any(
                    token in lower_context
                    for token in (
                        "포장",
                        "package",
                        "packing",
                        "수납함",
                        "에어드레서",
                        "스타일러",
                        "설치 가능한",
                        "호환 가능한",
                    )
                ),
            }
        )
    return found


def option_name_conflict(product_name: str, triple: tuple[int | float, int | float, int | float]) -> bool:
    """Flag a missed larger named size, e.g. a 1600-2200 option parsed as 1400-2000."""
    named = [int(value) for value in re.findall(r"(?<!\d)(\d{3,4})(?!\d)", product_name)]
    furniture_sizes = [value for value in named if 300 <= value <= 5000]
    return bool(furniture_sizes and max(furniture_sizes) > float(triple[0]) + 50)


def is_set_or_common(row: dict[str, Any], text: str) -> tuple[bool, list[str]]:
    product_name = str(row.get("product_name") or "")
    small_category = str(row.get("small_category") or "")
    combined = f"{product_name} {small_category} {text}".lower()
    reasons: list[str] = []
    if any(token in product_name for token in ("세트", "SET", "Set", "set", "+")):
        reasons.append("PRODUCT_SET_NAME")
    if any(token in combined for token in ("option", "옵션", "type a", "type b", "type c", "타입")):
        reasons.append("OPTION_OR_TYPE_IMAGE")
    if any(token in combined for token in ("2인", "4인", "6인", "체어", "chair", "스툴", "stool")) and "테이블" in combined:
        reasons.append("TABLE_COMPONENT_MIX")
    if any(token in combined for token in ("공통", "lineup", "line-up", "사이즈별", "size guide")):
        reasons.append("COMMON_LINEUP_IMAGE")
    return bool(reasons), reasons


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--run-dir",
        type=Path,
        default=Path("homestyle_bulk_run/ocr/spatial_diagram_callout_wave1"),
    )
    parser.add_argument("--output", type=Path, default=None)
    args = parser.parse_args()

    run_dir = args.run_dir
    output = args.output or run_dir / "explicit_dimension_audit.json"
    manifest = load_json(run_dir / "manifest.json")
    by_file = {
        Path(str(row.get("file") or "")).name: row
        for row in manifest.get("products", [])
        if row.get("file")
    }

    layout_rows: dict[str, dict[str, Any]] = {}
    for path in sorted(run_dir.glob("layout_ocr_[0-9][0-9].json")):
        for row in load_json(path):
            filename = str(row.get("file") or "")
            if filename:
                layout_rows[filename] = row

    attempts: list[dict[str, Any]] = []
    by_product: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for filename, ocr_row in sorted(layout_rows.items()):
        source = by_file.get(filename)
        if not source or ocr_row.get("status") != "SUCCESS":
            continue
        text = str(ocr_row.get("text") or "")
        triples = complete_triples(text, f"LAYOUT_OCR:{filename}")
        axis_triples = strong_axis_triples(text)
        strict_triples = strict_expression_triples(text)
        mixed, reasons = is_set_or_common(source, text)
        item = {
            "product_id": source.get("product_id"),
            "product_name": source.get("product_name"),
            "small_category": source.get("small_category"),
            "file": filename,
            "image_url": (source.get("image") or {}).get("url"),
            "triples": [list(values) for values in triples],
            "strong_axis_triples": axis_triples,
            "strict_expression_triples": strict_triples,
            "triple_count": len(triples),
            "mixed_product_risk": mixed,
            "risk_reasons": reasons,
            "ocr_text": text,
        }
        attempts.append(item)
        by_product[str(source.get("product_id") or "")].append(item)

    products: list[dict[str, Any]] = []
    decision_counts: Counter[str] = Counter()
    strong_decision_counts: Counter[str] = Counter()
    strict_decision_counts: Counter[str] = Counter()
    for product_id, rows in sorted(by_product.items()):
        unique = sorted({tuple(values) for row in rows for values in row["triples"]})
        risks = sorted({reason for row in rows for reason in row["risk_reasons"]})
        strong = {
            tuple(candidate["triple"])
            for row in rows
            for candidate in row["strong_axis_triples"]
            if not candidate["packaging_context"]
            and all(0 < float(value) <= 5000 for value in candidate["triple"])
        }
        strict = {
            tuple(candidate["triple"])
            for row in rows
            for candidate in row["strict_expression_triples"]
            if not candidate["non_product_context"]
            and all(0 < float(value) <= 5000 for value in candidate["triple"])
            and float(candidate["triple"][0]) >= 200
            and float(candidate["triple"][1]) >= 100
            and float(candidate["triple"][2]) >= 200
        }
        if not unique:
            decision = "NO_EXPLICIT_COMPLETE_TRIPLE"
        elif len(unique) > 1:
            decision = "REVIEW_MULTIPLE_OR_CONFLICTING_TRIPLES"
        elif risks:
            decision = "REVIEW_SET_OPTION_COMMON_IMAGE"
        else:
            decision = "HIGH_UNIQUE_EXPLICIT_TRIPLE"
        decision_counts[decision] += 1
        if not strong:
            strong_decision = "NO_STRONG_AXIS_TRIPLE"
        elif len(strong) > 1:
            strong_decision = "REVIEW_MULTIPLE_STRONG_AXIS_TRIPLES"
        elif risks:
            strong_decision = "REVIEW_STRONG_AXIS_SET_OPTION_COMMON"
        else:
            strong_decision = "HIGH_UNIQUE_STRONG_AXIS_TRIPLE"
        strong_decision_counts[strong_decision] += 1
        if not strict:
            strict_decision = "NO_STRICT_EXPRESSION_TRIPLE"
        elif len(strict) > 1:
            strict_decision = "REVIEW_MULTIPLE_STRICT_EXPRESSION_TRIPLES"
        elif risks:
            strict_decision = "REVIEW_STRICT_SET_OPTION_COMMON"
        elif option_name_conflict(str(rows[0]["product_name"] or ""), next(iter(strict))):
            strict_decision = "REVIEW_STRICT_PRODUCT_NAME_CONFLICT"
        else:
            strict_decision = "HIGH_UNIQUE_STRICT_EXPRESSION_TRIPLE"
        strict_decision_counts[strict_decision] += 1
        products.append(
            {
                "product_id": product_id,
                "product_name": rows[0]["product_name"],
                "small_category": rows[0]["small_category"],
                "unique_triples": [list(values) for values in unique],
                "strong_axis_triples": [list(values) for values in sorted(strong)],
                "strict_expression_triples": [list(values) for values in sorted(strict)],
                "candidate_image_count": len(rows),
                "risk_reasons": risks,
                "decision": decision,
                "strong_axis_decision": strong_decision,
                "strict_expression_decision": strict_decision,
                "evidence_files": [row["file"] for row in rows if row["triples"]],
            }
        )

    payload = {
        "summary": {
            "layout_images": len(attempts),
            "products": len(products),
            "images_with_explicit_complete_triple": sum(bool(row["triples"]) for row in attempts),
            "parsed_triple_frequency": dict(
                Counter(str(row["triple_count"]) for row in attempts)
            ),
            "decision_counts": dict(decision_counts),
            "strong_axis_decision_counts": dict(strong_decision_counts),
            "strict_expression_decision_counts": dict(strict_decision_counts),
        },
        "products": products,
        "attempts": attempts,
    }
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(payload["summary"], ensure_ascii=False, indent=2))
    print(f"output={output}")


if __name__ == "__main__":
    main()
