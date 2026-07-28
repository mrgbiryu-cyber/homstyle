"""Build a human-review wide table containing every complete OCR W/D/H candidate.

No semantic selection is performed.  Component, option, state, unit-error and
other conflicting triples are preserved as separate candidate column groups.
"""

from __future__ import annotations

import json
import re
import sqlite3
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parent
RUN_DIR = ROOT / "homestyle_bulk_run"
DB_PATH = RUN_DIR / "homestyle_bulk.sqlite"
AUDIT_PATH = (
    RUN_DIR
    / "ocr"
    / "spatial_diagram_callout_wave1"
    / "explicit_dimension_audit.json"
)
OUTPUT_PATH = RUN_DIR / "dimension_ocr_review_wide_staging_latest.json"
TABLE = "stg_dimension_ocr_review_wide"
RUN_NAME = "spatial_diagram_callout_wave1"
MAX_CANDIDATES = 10
MAX_RAW_TOKENS = 40
TARGETED_RECHECK_DIR = RUN_DIR / "ocr" / "targeted_recheck_20260723"
CLASSIFICATION_COLUMNS = [
    "ocr_review_classification TEXT",
    "ocr_size_label_present INTEGER NOT NULL DEFAULT 0",
    "ocr_axis_status TEXT",
    "ocr_unit_status TEXT",
    "ocr_detected_axis_count INTEGER NOT NULL DEFAULT 0",
    "ocr_unclassified_value_count INTEGER NOT NULL DEFAULT 0",
]


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def candidate_columns() -> list[str]:
    result: list[str] = []
    for number in range(1, MAX_CANDIDATES + 1):
        prefix = f"ocr_{number:02d}"
        result.extend(
            [
                f"{prefix}_w_mm REAL",
                f"{prefix}_d_mm REAL",
                f"{prefix}_h_mm REAL",
                f"{prefix}_source_type TEXT",
                f"{prefix}_image_url TEXT",
                f"{prefix}_evidence_text TEXT",
            ]
        )
    return result


def raw_token_columns() -> list[str]:
    result: list[str] = []
    for number in range(1, MAX_RAW_TOKENS + 1):
        prefix = f"ocr_raw_{number:02d}"
        result.extend(
            [
                f"{prefix}_value_text TEXT",
                f"{prefix}_image_url TEXT",
                f"{prefix}_context_text TEXT",
            ]
        )
    return result


def create_table(connection: sqlite3.Connection) -> None:
    connection.execute(
        f"""
        CREATE TABLE IF NOT EXISTS {TABLE} (
            snapshot_id TEXT NOT NULL,
            assessed_at TEXT NOT NULL,
            is_current INTEGER NOT NULL DEFAULT 1,
            product_id TEXT NOT NULL,
            product_name TEXT,
            small_category TEXT,
            current_status TEXT,
            current_w_mm REAL,
            current_d_mm REAL,
            current_h_mm REAL,
            ocr_candidate_count INTEGER NOT NULL DEFAULT 0,
            ocr_candidate_overflow_count INTEGER NOT NULL DEFAULT 0,
            ocr_raw_token_count INTEGER NOT NULL DEFAULT 0,
            ocr_raw_token_overflow_count INTEGER NOT NULL DEFAULT 0,
            has_human_review_information INTEGER NOT NULL DEFAULT 0,
            human_selected_candidate_no INTEGER,
            human_review_status TEXT NOT NULL DEFAULT '대기',
            human_review_note TEXT,
            {', '.join(candidate_columns())},
            {', '.join(raw_token_columns())},
            {', '.join(CLASSIFICATION_COLUMNS)},
            PRIMARY KEY(snapshot_id, product_id)
        )
        """
    )
    existing = {row[1] for row in connection.execute(f"PRAGMA table_info({TABLE})")}
    for definition in CLASSIFICATION_COLUMNS:
        name = definition.split()[0]
        if name not in existing:
            connection.execute(f"ALTER TABLE {TABLE} ADD COLUMN {definition}")
    connection.execute(
        f"CREATE INDEX IF NOT EXISTS idx_{TABLE}_current "
        f"ON {TABLE}(is_current,ocr_candidate_count)"
    )


def add_candidate(
    target: dict[str, list[dict[str, Any]]],
    product_id: str,
    triple: list[Any] | tuple[Any, Any, Any],
    source_type: str,
    image_url: str = "",
    evidence_text: str = "",
) -> None:
    if len(triple) != 3 or any(value is None for value in triple):
        return
    key = tuple(float(value) for value in triple)
    rows = target[product_id]
    for row in rows:
        if row["key"] == key:
            if source_type not in row["source_type"].split("+"):
                row["source_type"] += f"+{source_type}"
            if not row["image_url"] and image_url:
                row["image_url"] = image_url
            if not row["evidence_text"] and evidence_text:
                row["evidence_text"] = evidence_text
            return
    rows.append(
        {
            "key": key,
            "w_mm": key[0],
            "d_mm": key[1],
            "h_mm": key[2],
            "source_type": source_type,
            "image_url": image_url,
            "evidence_text": evidence_text,
        }
    )


def build() -> dict[str, Any]:
    audit = json.loads(AUDIT_PATH.read_text(encoding="utf-8-sig"))
    candidates: dict[str, list[dict[str, Any]]] = defaultdict(list)
    raw_tokens: dict[str, list[dict[str, str]]] = defaultdict(list)

    # Preserve OCR/image order and parser order; do not rank by meaning.
    for attempt in audit.get("attempts") or []:
        for triple in attempt.get("triples") or []:
            add_candidate(
                candidates,
                str(attempt.get("product_id") or ""),
                triple,
                "EXPLICIT_TEXT_OCR",
                str(attempt.get("image_url") or ""),
                str(attempt.get("ocr_text") or ""),
            )

    manifest = json.loads(
        (AUDIT_PATH.parent / "manifest.json").read_text(encoding="utf-8-sig")
    )
    image_by_file = {
        Path(str(row.get("file") or "")).name: {
            "product_id": str(row.get("product_id") or ""),
            "image_url": str((row.get("image") or {}).get("url") or ""),
        }
        for row in manifest.get("products") or []
        if row.get("file")
    }
    token_pattern = re.compile(
        r"(?i)(?:(?<![A-Z])([WDHL])|가로|세로|깊이|높이|너비|폭|지름|직경|[ØΦ])"
        r"\s*[:=]?\s*\(?\s*(\d[\d,.]*)(?:\s*[-~]\s*(\d[\d,.]*))?"
        r"\s*\)?\s*(mm|cm)?"
        r"|(?<![\d.])(\d[\d,.]*)\s*(mm|cm)\b"
    )
    raw_seen: dict[str, set[str]] = defaultdict(set)

    def add_raw_token(
        product_id: str,
        value_text: str,
        image_url: str,
        context_text: str,
        key_prefix: str = "",
    ) -> None:
        key = key_prefix + re.sub(r"\s+", "", value_text).lower()
        if key in raw_seen[product_id]:
            return
        raw_seen[product_id].add(key)
        raw_tokens[product_id].append(
            {
                "value_text": value_text,
                "image_url": image_url,
                "context_text": re.sub(r"[\r\n]+", " ", context_text).strip(),
            }
        )

    for ocr_path in sorted(AUDIT_PATH.parent.glob("text_ocr_[0-9][0-9].json")):
        for ocr_row in json.loads(ocr_path.read_text(encoding="utf-8-sig")):
            if ocr_row.get("status") != "SUCCESS":
                continue
            filename = str(ocr_row.get("file") or "")
            image = image_by_file.get(filename)
            if not image:
                continue
            product_id = image["product_id"]
            text = str(ocr_row.get("text") or "")
            for match in token_pattern.finditer(text):
                value_text = match.group(0).strip()
                context = text[
                    max(0, match.start() - 70) : min(len(text), match.end() + 70)
                ]
                add_raw_token(
                    product_id,
                    value_text,
                    image["image_url"],
                    context,
                )

    # Targeted full-detail scans are merged into the same review columns.  A
    # size-labelled block with bare numbers is retained even without axes/unit.
    strict_wdh = re.compile(
        r"(?i)\(?\s*W\s*\)?\s*[:=]?\s*(\d[\d,.]*)\s*(mm|cm)?\s*"
        r"(?:x|×|\*|\)\s*\()\s*"
        r"\(?\s*D\s*\)?\s*[:=]?\s*(\d[\d,.]*)\s*(mm|cm)?\s*"
        r"(?:x|×|\*|\)\s*\()\s*"
        r"\(?\s*H\s*\)?\s*[:=]?\s*(\d[\d,.]*)\s*(mm|cm)?"
    )
    if TARGETED_RECHECK_DIR.exists():
        for result_path in sorted(TARGETED_RECHECK_DIR.glob("G*_ocr.json")):
            product_id = result_path.stem.removesuffix("_ocr")
            for ocr_row in json.loads(result_path.read_text(encoding="utf-8-sig")):
                if ocr_row.get("status") != "SUCCESS":
                    continue
                filename = str(ocr_row.get("file") or "")
                image_path = str(TARGETED_RECHECK_DIR / product_id / filename)
                text = str(ocr_row.get("text") or "")
                matches = list(token_pattern.finditer(text))
                for match in matches:
                    context = text[
                        max(0, match.start() - 70) : min(len(text), match.end() + 70)
                    ]
                    add_raw_token(
                        product_id,
                        match.group(0).strip(),
                        image_path,
                        context,
                    )

                size_label = re.search(
                    r"(?i)(?:제품\s*사이즈|제품\s*크기|product\s*(?:size|description)|\bsize\b)",
                    text,
                )
                if size_label:
                    size_block = text[size_label.start() : size_label.start() + 300]
                    if not token_pattern.search(size_block):
                        for number in re.finditer(r"(?<![\d.])\d+(?:\.\d+)?(?![\d.])", size_block):
                            add_raw_token(
                                product_id,
                                number.group(0),
                                image_path,
                                size_block,
                                key_prefix=f"bare-size:{filename}:{number.start()}:",
                            )

                for match in strict_wdh.finditer(text):
                    values = [match.group(1), match.group(3), match.group(5)]
                    units = [match.group(2), match.group(4), match.group(6)]
                    common_unit = next(
                        (unit.lower() for unit in reversed(units) if unit), "mm"
                    )
                    converted = []
                    for value, unit in zip(values, units):
                        number = float(value.replace(",", ""))
                        if (unit or common_unit).lower() == "cm":
                            number *= 10
                        converted.append(int(number) if number.is_integer() else number)
                    add_candidate(
                        candidates,
                        product_id,
                        converted,
                        "FULL_DETAIL_IMAGE_OCR",
                        image_path,
                        text,
                    )

    connection = sqlite3.connect(DB_PATH)
    for product_id, w_mm, d_mm, h_mm, image_url, raw_text in connection.execute(
        """
        SELECT product_id,candidate_w_mm,candidate_d_mm,candidate_h_mm,
               image_url,raw_text
        FROM stg_dimension_spatial_attempt
        WHERE run_name=? AND spatial_candidate=1
          AND candidate_w_mm IS NOT NULL
          AND candidate_d_mm IS NOT NULL
          AND candidate_h_mm IS NOT NULL
        ORDER BY product_id,attempt_order
        """,
        (RUN_NAME,),
    ):
        add_candidate(
            candidates,
            product_id,
            [w_mm, d_mm, h_mm],
            "SPATIAL_CALLOUT_OCR",
            str(image_url or ""),
            str(raw_text or ""),
        )

    source_rows = connection.execute(
        """
        SELECT product_id,product_name,small_category,current_status,
               current_w_mm,current_d_mm,current_h_mm
        FROM stg_dimension_reinforcement
        WHERE is_current=1
        ORDER BY product_id
        """
    ).fetchall()

    snapshot_id = datetime.now().astimezone().strftime("%Y-%m-%dT%H%M%S_%z")
    assessed_at = now_text()
    create_table(connection)
    insert_columns = (
        18 + MAX_CANDIDATES * 6 + MAX_RAW_TOKENS * 3
        + len(CLASSIFICATION_COLUMNS)
    )
    placeholders = ",".join("?" for _ in range(insert_columns))
    inserts: list[list[Any]] = []
    distribution: Counter[int] = Counter()
    covered = 0
    overflow_products = 0
    raw_covered = 0
    raw_overflow_products = 0
    human_review_information = 0
    classification_counts: Counter[str] = Counter()

    for product_id, name, category, status, current_w, current_d, current_h in source_rows:
        all_candidates = candidates.get(product_id) or []
        count = len(all_candidates)
        distribution[count] += 1
        covered += int(count > 0)
        overflow = max(0, count - MAX_CANDIDATES)
        overflow_products += int(overflow > 0)
        all_raw_tokens = raw_tokens.get(product_id) or []
        raw_count = len(all_raw_tokens)
        raw_covered += int(raw_count > 0)
        raw_overflow = max(0, raw_count - MAX_RAW_TOKENS)
        raw_overflow_products += int(raw_overflow > 0)
        has_existing_axis = any(
            value is not None for value in (current_w, current_d, current_h)
        )
        has_review_information = bool(has_existing_axis or count or raw_count)
        human_review_information += int(has_review_information)
        raw_value_texts = [item["value_text"] for item in all_raw_tokens]
        raw_context = " ".join(item["context_text"] for item in all_raw_tokens)
        size_label_present = int(
            bool(
                re.search(
                    r"(?i)(?:제품\s*사이즈|제품\s*크기|product\s*(?:size|description)|\bsize\b)",
                    raw_context,
                )
            )
        )
        axes = {
            match.group(1).upper()
            for value in raw_value_texts
            for match in [re.match(r"(?i)\s*([WDHL])", value)]
            if match
        }
        if {"W", "D", "H"}.issubset(axes):
            axis_status = "W_D_H_PRESENT"
        elif axes:
            axis_status = "PARTIAL_AXES"
        elif raw_count:
            axis_status = "UNLABELED_VALUES"
        else:
            axis_status = "NO_VALUES"
        unit_flags = [bool(re.search(r"(?i)\b(?:mm|cm)\b", value)) for value in raw_value_texts]
        if not unit_flags:
            unit_status = "NO_VALUES"
        elif all(unit_flags):
            unit_status = "ALL_UNIT_PRESENT"
        elif any(unit_flags):
            unit_status = "MIXED_UNIT_PRESENT"
        else:
            unit_status = "UNIT_MISSING"
        if count:
            classification = "COMPLETE_WDH_CANDIDATE"
        elif size_label_present and raw_count and unit_status == "UNIT_MISSING":
            classification = "SIZE_LABEL_UNIT_MISSING"
        elif axis_status == "PARTIAL_AXES":
            classification = "PARTIAL_AXES"
        elif raw_count:
            classification = "UNCLASSIFIED_OCR_VALUES"
        elif has_existing_axis:
            classification = "CURRENT_PARTIAL_ONLY"
        else:
            classification = "NO_INFORMATION"
        classification_counts[classification] += 1
        row: list[Any] = [
            snapshot_id,
            assessed_at,
            1,
            product_id,
            name,
            category,
            status,
            current_w,
            current_d,
            current_h,
            count,
            overflow,
            raw_count,
            raw_overflow,
            int(has_review_information),
            None,
            "대기",
            None,
        ]
        for number in range(MAX_CANDIDATES):
            if number < count:
                item = all_candidates[number]
                row.extend(
                    [
                        item["w_mm"],
                        item["d_mm"],
                        item["h_mm"],
                        item["source_type"],
                        item["image_url"],
                        item["evidence_text"],
                    ]
                )
            else:
                row.extend([None, None, None, None, None, None])
        for number in range(MAX_RAW_TOKENS):
            if number < raw_count:
                item = all_raw_tokens[number]
                row.extend(
                    [
                        item["value_text"],
                        item["image_url"],
                        item["context_text"],
                    ]
                )
            else:
                row.extend([None, None, None])
        row.extend(
            [
                classification,
                size_label_present,
                axis_status,
                unit_status,
                len(axes),
                raw_count,
            ]
        )
        inserts.append(row)

    with connection:
        connection.execute(f"UPDATE {TABLE} SET is_current=0 WHERE is_current=1")
        connection.executemany(
            f"INSERT INTO {TABLE} VALUES ({placeholders})",
            inserts,
        )
    connection.close()

    total = len(source_rows)
    result = {
        "database": str(DB_PATH),
        "table": TABLE,
        "snapshot_id": snapshot_id,
        "assessed_at": assessed_at,
        "rows": total,
        "ocr_candidate_products": covered,
        "raw_ocr_token_products": raw_covered,
        "human_review_information_products": human_review_information,
        "remaining_without_current_axis_or_raw_ocr_token": total
        - human_review_information,
        "remaining_without_complete_ocr_candidate": total - covered,
        "ocr_candidate_coverage_rate": covered / total if total else 0,
        "candidate_count_distribution": {
            str(key): value for key, value in sorted(distribution.items())
        },
        "max_candidate_column_groups": MAX_CANDIDATES,
        "max_raw_token_column_groups": MAX_RAW_TOKENS,
        "columns_per_candidate": [
            "W_mm",
            "D_mm",
            "H_mm",
            "source_type",
            "image_url",
            "evidence_text",
        ],
        "overflow_products": overflow_products,
        "raw_token_overflow_products": raw_overflow_products,
        "classification_counts": dict(classification_counts.most_common()),
        "semantic_filtering": False,
        "excel_written": False,
    }
    OUTPUT_PATH.write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return result


if __name__ == "__main__":
    print(json.dumps(build(), ensure_ascii=False, indent=2))
