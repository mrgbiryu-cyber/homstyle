from __future__ import annotations

import json
import re
import sqlite3
from datetime import datetime

import build_homestyle_bulk_workbook as workbook
from bulk_homestyle_collect import DB_PATH, RUN_DIR, pack, unpack


RUN_NAME = "option_single_width_plus_ocr"
OUTPUT = RUN_DIR / "dimension_option_inference_apply.json"
NUMBER_WITH_OCR_SPACE = r"\d(?:\s?\d){1,3}"
OCR_TRIPLE = re.compile(
    rf"(?<!\d)({NUMBER_WITH_OCR_SPACE})\s*\)\(\s*"
    rf"({NUMBER_WITH_OCR_SPACE})\s*(?:x|\)\()\s*"
    rf"({NUMBER_WITH_OCR_SPACE})\s*(?:mm|cm)",
    re.I,
)


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def normalized_int(value: str) -> int:
    return int(re.sub(r"\s", "", value))


def numeric_size_options(data: dict) -> list[int]:
    result = []
    for group in workbook.option_groups(data, ""):
        if group["style"] != "사이즈":
            continue
        for item in group["items"]:
            match = re.fullmatch(
                r"\s*(\d{2,4})\s*(?:mm|cm)?\s*", item["name"], re.I
            )
            if match:
                result.append(int(match.group(1)))
    return sorted(set(result))


def candidates() -> list[dict]:
    connection = sqlite3.connect(DB_PATH)
    rows = connection.execute(
        "SELECT product_id,goods_blob,ocr_blob FROM sources ORDER BY product_id"
    ).fetchall()
    result = []
    for product_id, goods_blob, ocr_blob in rows:
        data = (unpack(goods_blob) or {}).get("data") or {}
        widths = numeric_size_options(data)
        if len(widths) < 2:
            continue
        ocr = unpack(ocr_blob) or {}
        dimension_text = str(ocr.get("dimension_text") or "")
        triples = [
            tuple(normalized_int(value) for value in match.groups())
            for match in OCR_TRIPLE.finditer(dimension_text)
        ]
        matched = [triple for triple in triples if triple[0] in widths]
        by_width = {triple[0]: triple for triple in matched}
        if sorted(by_width) != widths:
            continue
        common_dh = {(triple[1], triple[2]) for triple in by_width.values()}
        if len(common_dh) != 1:
            continue
        ordered = [by_width[width] for width in widths]
        result.append(
            {
                "product_id": product_id,
                "product_name": str(data.get("productName") or ""),
                "option_widths": widths,
                "dimensions": ordered,
                "image_url": str((ocr.get("selected") or {}).get("url") or ""),
                "rule": "ALL_NUMERIC_SIZE_OPTIONS_MATCH_OCR_W_PLUS_COMMON_DH",
                "confidence": "HIGH",
            }
        )
    connection.close()
    return result


def apply() -> dict:
    rows = candidates()
    applied_at = now_text()
    connection = sqlite3.connect(DB_PATH)
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS stg_dimension_option_inference (
            run_name TEXT NOT NULL,
            assessed_at TEXT NOT NULL,
            product_id TEXT NOT NULL,
            product_name TEXT,
            rule TEXT NOT NULL,
            confidence TEXT NOT NULL,
            option_widths_json TEXT NOT NULL,
            dimensions_json TEXT NOT NULL,
            evidence_image_url TEXT,
            applied INTEGER NOT NULL DEFAULT 0,
            PRIMARY KEY(run_name,product_id)
        );
        CREATE TABLE IF NOT EXISTS source_ocr_dimension_backup (
            run_name TEXT NOT NULL,
            product_id TEXT NOT NULL,
            backed_up_at TEXT NOT NULL,
            previous_ocr_status INTEGER,
            previous_ocr_blob BLOB,
            PRIMARY KEY(run_name,product_id)
        );
        """
    )
    applied = []
    skipped = []
    with connection:
        connection.execute(
            "DELETE FROM stg_dimension_option_inference WHERE run_name=?", (RUN_NAME,)
        )
        for row in rows:
            connection.execute(
                "INSERT INTO stg_dimension_option_inference VALUES (?,?,?,?,?,?,?,?,?,0)",
                (
                    RUN_NAME,
                    applied_at,
                    row["product_id"],
                    row["product_name"],
                    row["rule"],
                    row["confidence"],
                    json.dumps(row["option_widths"], ensure_ascii=False),
                    json.dumps(row["dimensions"], ensure_ascii=False),
                    row["image_url"],
                ),
            )
            source = connection.execute(
                "SELECT ocr_status,ocr_blob FROM sources WHERE product_id=?",
                (row["product_id"],),
            ).fetchone()
            old_status, old_blob = source
            ocr_blob = unpack(old_blob) or {}
            reinforcements = list(ocr_blob.get("dimension_reinforcements") or [])
            if any(item.get("run_name") == RUN_NAME for item in reinforcements):
                connection.execute(
                    "UPDATE stg_dimension_option_inference SET applied=1 "
                    "WHERE run_name=? AND product_id=?",
                    (RUN_NAME, row["product_id"]),
                )
                skipped.append({"product_id": row["product_id"], "reason": "ALREADY_APPLIED"})
                continue
            connection.execute(
                "INSERT OR IGNORE INTO source_ocr_dimension_backup VALUES (?,?,?,?,?)",
                (RUN_NAME, row["product_id"], applied_at, old_status, old_blob),
            )
            explicit_lines = [
                f"사이즈 옵션+OCR 동일모델 검증 W={w} mm D={d} mm H={h} mm"
                for w, d, h in row["dimensions"]
            ]
            old_text = str(ocr_blob.get("dimension_text") or "").strip()
            ocr_blob["dimension_text"] = "\n".join(
                value for value in (old_text, *explicit_lines) if value
            )
            reinforcements.append(
                {
                    "run_name": RUN_NAME,
                    "applied_at": applied_at,
                    "image_url": row["image_url"],
                    "confidence": row["confidence"],
                    "rule": row["rule"],
                    "option_widths": row["option_widths"],
                    "dimensions": row["dimensions"],
                }
            )
            ocr_blob["dimension_reinforcements"] = reinforcements
            connection.execute(
                "UPDATE sources SET ocr_blob=? WHERE product_id=?",
                (pack(ocr_blob), row["product_id"]),
            )
            connection.execute(
                "UPDATE stg_dimension_option_inference SET applied=1 "
                "WHERE run_name=? AND product_id=?",
                (RUN_NAME, row["product_id"]),
            )
            applied.append(row)
    result = {
        "run_name": RUN_NAME,
        "applied_at": applied_at,
        "candidate_count": len(rows),
        "applied_count": len(applied),
        "skipped_count": len(skipped),
        "applied": applied,
        "skipped": skipped,
        "staging_table": "stg_dimension_option_inference",
        "backup_table": "source_ocr_dimension_backup",
        "excel_written": False,
    }
    connection.close()
    OUTPUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result


if __name__ == "__main__":
    print(json.dumps(apply(), ensure_ascii=False, indent=2))
