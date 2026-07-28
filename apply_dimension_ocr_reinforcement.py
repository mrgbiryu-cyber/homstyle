from __future__ import annotations

import json
import sqlite3
from datetime import datetime

from bulk_homestyle_collect import DB_PATH, RUN_DIR, pack, unpack


RUN_NAME = "partial_nonflat_sequential"
OUTPUT = RUN_DIR / "dimension_ocr_reinforcement_apply.json"


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def apply() -> dict:
    connection = sqlite3.connect(DB_PATH)
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS source_ocr_dimension_backup (
            run_name TEXT NOT NULL,
            product_id TEXT NOT NULL,
            backed_up_at TEXT NOT NULL,
            previous_ocr_status INTEGER,
            previous_ocr_blob BLOB,
            PRIMARY KEY(run_name,product_id)
        )
        """
    )
    rows = connection.execute(
        """
        SELECT product_id,w_mm,d_mm,h_mm,confidence,
               evidence_image_url,evidence_text
        FROM stg_dimension_ocr_result
        WHERE run_name=? AND recovered=1
          AND confidence='HIGH_KNOWN_AXES_MATCH'
        ORDER BY product_id
        """,
        (RUN_NAME,),
    ).fetchall()
    applied_at = now_text()
    applied = []
    skipped = []
    with connection:
        for product_id, w_mm, d_mm, h_mm, confidence, image_url, evidence_text in rows:
            source = connection.execute(
                "SELECT ocr_status,ocr_blob FROM sources WHERE product_id=?",
                (product_id,),
            ).fetchone()
            if not source:
                skipped.append({"product_id": product_id, "reason": "SOURCE_NOT_FOUND"})
                continue
            old_status, old_blob = source
            ocr_blob = unpack(old_blob) or {}
            reinforcements = list(ocr_blob.get("dimension_reinforcements") or [])
            explicit_line = (
                f"추가 OCR 동일제품 검증 W={w_mm} mm D={d_mm} mm H={h_mm} mm"
            )
            if any(item.get("run_name") == RUN_NAME for item in reinforcements):
                old_dimension_text = str(ocr_blob.get("dimension_text") or "").strip()
                if explicit_line not in old_dimension_text:
                    ocr_blob["dimension_text"] = "\n".join(
                        value for value in (old_dimension_text, explicit_line) if value
                    )
                    connection.execute(
                        "UPDATE sources SET ocr_blob=? WHERE product_id=?",
                        (pack(ocr_blob), product_id),
                    )
                    reason = "ALREADY_APPLIED_EXPLICIT_LINE_ADDED"
                else:
                    reason = "ALREADY_APPLIED"
                skipped.append({"product_id": product_id, "reason": reason})
                connection.execute(
                    "UPDATE stg_dimension_ocr_result SET applied_to_sources=1 "
                    "WHERE run_name=? AND product_id=?",
                    (RUN_NAME, product_id),
                )
                continue
            connection.execute(
                "INSERT OR IGNORE INTO source_ocr_dimension_backup "
                "VALUES (?,?,?,?,?)",
                (RUN_NAME, product_id, applied_at, old_status, old_blob),
            )
            evidence_text = str(evidence_text or "").strip()
            old_dimension_text = str(ocr_blob.get("dimension_text") or "").strip()
            if evidence_text and evidence_text not in old_dimension_text:
                merged_dimension_text = "\n".join(
                    value for value in (old_dimension_text, evidence_text) if value
                )
            else:
                merged_dimension_text = old_dimension_text
            if explicit_line not in merged_dimension_text:
                merged_dimension_text = "\n".join(
                    value for value in (merged_dimension_text, explicit_line) if value
                )
            reinforcement = {
                "run_name": RUN_NAME,
                "applied_at": applied_at,
                "image_url": image_url,
                "confidence": confidence,
                "w_mm": w_mm,
                "d_mm": d_mm,
                "h_mm": h_mm,
                "evidence_text": evidence_text,
            }
            reinforcements.append(reinforcement)
            ocr_blob["dimension_text"] = merged_dimension_text
            ocr_blob["dimension_reinforcements"] = reinforcements
            connection.execute(
                "UPDATE sources SET ocr_blob=? WHERE product_id=?",
                (pack(ocr_blob), product_id),
            )
            connection.execute(
                "UPDATE stg_dimension_ocr_result SET applied_to_sources=1 "
                "WHERE run_name=? AND product_id=?",
                (RUN_NAME, product_id),
            )
            applied.append(
                {
                    "product_id": product_id,
                    "w_mm": w_mm,
                    "d_mm": d_mm,
                    "h_mm": h_mm,
                    "image_url": image_url,
                }
            )
    active_applied_count = connection.execute(
        "SELECT COUNT(*) FROM stg_dimension_ocr_result "
        "WHERE run_name=? AND recovered=1 AND applied_to_sources=1",
        (RUN_NAME,),
    ).fetchone()[0]
    result = {
        "run_name": RUN_NAME,
        "applied_at": applied_at,
        "candidate_count": len(rows),
        "applied_count": len(applied),
        "active_applied_count": active_applied_count,
        "skipped_count": len(skipped),
        "applied": applied,
        "skipped": skipped,
        "backup_table": "source_ocr_dimension_backup",
        "excel_written": False,
    }
    connection.close()
    OUTPUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result


if __name__ == "__main__":
    print(json.dumps(apply(), ensure_ascii=False, indent=2))
