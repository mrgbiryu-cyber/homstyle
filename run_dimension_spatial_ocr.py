from __future__ import annotations

import argparse
import concurrent.futures
import io
import json
import os
import re
import shutil
import sqlite3
import subprocess
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from PIL import Image, ImageFile

import bulk_homestyle_ocr as ocr
from bulk_homestyle_collect import DB_PATH, pack, request_bytes, unpack
from dimension_spatial_parser import spatial_callout_candidate


RUN_NAME = "spatial_diagram_callout_wave1"
RUN_DIR = ocr.OCR_ROOT / RUN_NAME
MANIFEST_PATH = RUN_DIR / "manifest.json"
SCREEN_PATH = RUN_DIR / "layout_screen.json"
SUMMARY_PATH = ocr.RUN_DIR / "spatial_diagram_callout_wave1.json"
LAYOUT_OCR = Path(__file__).resolve().parent / "run_windows_ocr_layout.ps1"
DIRECT_TEXT_OCR = Path(__file__).resolve().parent / "run_windows_ocr_direct.ps1"
DEFAULT_SHARDS = 12
DEFAULT_MAX_CANDIDATES = 3
STILL_EXTENSIONS = {"jpg", "jpeg", "png", "webp", "bmp", "avif"}
DIMENSION_TOKEN_RE = re.compile(r"(?<!\d)\d{2,5}(?:[.,]\d+)?\s*(?:mm|cm)\b", re.I)

ImageFile.LOAD_TRUNCATED_IMAGES = True
Image.MAX_IMAGE_PIXELS = None


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def image_extension(url: str) -> str:
    path = urlparse(url).path.casefold()
    return path.rsplit(".", 1)[-1] if "." in path else ""


def candidate_images(
    data: dict[str, Any], old_ocr: dict[str, Any], max_candidates: int
) -> list[dict[str, Any]]:
    images = ocr.detail_images(str(data.get("detailInfo") or ""))
    old_url = ocr.normalize_url(str((old_ocr.get("selected") or {}).get("url") or ""))
    attempted = {old_url} if old_url else set()
    for item in old_ocr.get("dimension_reinforcements") or []:
        if isinstance(item, dict):
            attempted.add(
                ocr.normalize_url(str(item.get("image_url") or item.get("url") or ""))
            )
    result: list[dict[str, Any]] = []
    for image in images:
        url = image["url"]
        if url in attempted or image_extension(url) not in STILL_EXTENSIONS:
            continue
        if float(image.get("score") or 0) >= 100:
            continue
        result.append({**image, "reason": "NO_PATTERN_NEXT_RANKED_STILL"})
        if len(result) >= max_candidates:
            break
    return result


def download_task(task: dict[str, Any], shards: int) -> dict[str, Any]:
    product_id = task["product_id"]
    order = int(task["attempt_order"])
    shard = int(product_id[-4:]) % shards
    destination = RUN_DIR / f"text_shard_{shard:02d}" / f"{product_id}__{order:02d}.jpg"
    try:
        if destination.exists() and destination.stat().st_size > 0:
            with Image.open(destination) as existing:
                width, height = existing.size
            return {
                **task,
                "shard": shard,
                "download_status": "SUCCESS",
                "download_reused": True,
                "file": str(destination),
                "width": width,
                "height": height,
                "download_bytes": destination.stat().st_size,
                "error": "",
            }

        status, body = request_bytes(task["image"]["url"], timeout=20, attempts=2)
        if status != 200 or not body:
            raise RuntimeError(f"HTTP {status}; bytes={len(body)}")
        with Image.open(io.BytesIO(body)) as source:
            source.seek(0)
            image = source.convert("RGB")
            if image.width > 2200:
                resized_height = max(1, round(image.height * 2200 / image.width))
                image = image.resize((2200, resized_height), Image.Resampling.LANCZOS)
            destination.parent.mkdir(parents=True, exist_ok=True)
            image.save(destination, "JPEG", quality=90, optimize=True)
            width, height = image.size
        return {
            **task,
            "shard": shard,
            "download_status": "SUCCESS",
            "download_reused": False,
            "file": str(destination),
            "width": width,
            "height": height,
            "download_bytes": len(body),
            "error": "",
        }
    except Exception as exc:
        return {
            **task,
            "shard": shard,
            "download_status": "ERROR",
            "download_reused": False,
            "file": "",
            "width": None,
            "height": None,
            "download_bytes": 0,
            "error": f"{type(exc).__name__}: {exc}",
        }


def prepare(
    *,
    product_id: str = "",
    limit: int = 0,
    max_candidates: int = DEFAULT_MAX_CANDIDATES,
    workers: int = 24,
    shards: int = DEFAULT_SHARDS,
) -> dict[str, Any]:
    connection = sqlite3.connect(DB_PATH)
    sql = """
        SELECT p.product_id,p.product_name,p.small_category,s.goods_blob,s.ocr_blob
        FROM stg_dimension_pattern p
        JOIN sources s ON s.product_id=p.product_id
        WHERE p.is_current=1 AND p.notation_type_code='NO_CONFIRMED_PATTERN'
    """
    params: list[Any] = []
    if product_id:
        sql += " AND p.product_id=?"
        params.append(product_id)
    sql += " ORDER BY p.product_id"
    if limit:
        sql += " LIMIT ?"
        params.append(limit)
    targets = connection.execute(sql, params).fetchall()
    connection.close()

    tasks: list[dict[str, Any]] = []
    without_candidate: list[str] = []
    for pid, name, category, goods_blob, ocr_blob in targets:
        data = (unpack(goods_blob) or {}).get("data") or {}
        old_ocr = unpack(ocr_blob) or {}
        images = candidate_images(data, old_ocr, max_candidates)
        if not images:
            without_candidate.append(pid)
        for order, image in enumerate(images, 1):
            tasks.append(
                {
                    "product_id": pid,
                    "product_name": name,
                    "small_category": category,
                    "attempt_order": order,
                    "image": image,
                }
            )

    RUN_DIR.mkdir(parents=True, exist_ok=True)
    for shard in range(shards):
        (RUN_DIR / f"text_shard_{shard:02d}").mkdir(parents=True, exist_ok=True)
        (RUN_DIR / f"layout_shard_{shard:02d}").mkdir(parents=True, exist_ok=True)
    results: list[dict[str, Any]] = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(download_task, task, shards) for task in tasks]
        for index, future in enumerate(concurrent.futures.as_completed(futures), 1):
            results.append(future.result())
            if index % 250 == 0 or index == len(futures):
                print(f"downloaded={index}/{len(futures)}", flush=True)
    results.sort(key=lambda row: (row["product_id"], row["attempt_order"]))
    metadata = {
        "run_name": RUN_NAME,
        "created_at": now_text(),
        "target_products": len(targets),
        "candidate_images": len(tasks),
        "download_success": sum(row["download_status"] == "SUCCESS" for row in results),
        "download_reused": sum(bool(row.get("download_reused")) for row in results),
        "download_error": sum(row["download_status"] == "ERROR" for row in results),
        "products_without_candidate": len(without_candidate),
        "products_without_candidate_ids": without_candidate,
        "max_candidates_per_product": max_candidates,
        "shards": shards,
        "sources_db_written": False,
    }
    MANIFEST_PATH.write_text(
        json.dumps({"metadata": metadata, "products": results}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(json.dumps(metadata, ensure_ascii=False, indent=2), flush=True)
    return metadata


def read_ocr_outputs(prefix: str, shards: int) -> dict[str, dict[str, Any]]:
    rows: dict[str, dict[str, Any]] = {}
    for shard in range(shards):
        path = RUN_DIR / f"{prefix}_{shard:02d}.json"
        if not path.exists():
            continue
        raw = path.read_text(encoding="utf-8-sig").strip()
        if not raw:
            continue
        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            print(f"[warn] invalid OCR output skipped: {path}")
            continue
        if isinstance(data, dict):
            data = [data]
        for row in data:
            rows[Path(row.get("file") or "").stem] = row
    return rows


def run_text_ocr(
    shards: int, ocr_workers: int = 2, shard_start: int = 0, shard_end: int = 0
) -> None:
    failed = []
    stop = min(shards, shard_end) if shard_end > 0 else shards
    for start in range(max(0, shard_start), stop, max(1, ocr_workers)):
        processes = []
        for shard in range(start, min(stop, start + max(1, ocr_workers))):
            input_dir = RUN_DIR / f"text_shard_{shard:02d}"
            output_json = RUN_DIR / f"text_ocr_{shard:02d}.json"
            command = [
                "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                "-File", str(DIRECT_TEXT_OCR),
                "-InputDirectory", str(input_dir),
                "-OutputJson", str(output_json),
                "-LanguageTag", "ko",
                "-Resume",
            ]
            processes.append((shard, subprocess.Popen(command)))
        for shard, process in processes:
            code = process.wait()
            print(f"text_ocr_shard={shard} exit={code}", flush=True)
            if code:
                failed.append(shard)
    if failed:
        raise RuntimeError(f"text OCR shards failed: {failed}")


def screen_layout_candidates(shards: int) -> dict[str, Any]:
    manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    text_rows = read_ocr_outputs("text_ocr", shards)
    screened: list[dict[str, Any]] = []
    status_counts = Counter()
    for item in manifest["products"]:
        stem = f"{item['product_id']}__{int(item['attempt_order']):02d}"
        row = text_rows.get(stem) or {}
        raw_text = str(row.get("text") or "")
        token_count = len(DIMENSION_TOKEN_RE.findall(raw_text))
        status_counts[f"ocr:{row.get('status') or 'NOT_RUN'}"] += 1
        if token_count < 3 or item["download_status"] != "SUCCESS":
            continue
        source = Path(item["file"])
        layout_dir = RUN_DIR / f"layout_shard_{int(item['shard']):02d}"
        destination = layout_dir / source.name
        if destination.exists():
            destination.unlink()
        try:
            os.link(source, destination)
        except OSError:
            shutil.copy2(source, destination)
        screened.append(
            {
                **item,
                "stem": stem,
                "text_ocr_status": row.get("status") or "NOT_RUN",
                "text_ocr": raw_text,
                "dimension_token_count": token_count,
                "layout_file": str(destination),
            }
        )
    result = {
        "metadata": {
            "run_name": RUN_NAME,
            "screened_at": now_text(),
            "input_images": len(manifest["products"]),
            "layout_candidate_images": len(screened),
            "status_counts": dict(status_counts),
            "shards": shards,
        },
        "products": screened,
    }
    SCREEN_PATH.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result["metadata"], ensure_ascii=False, indent=2), flush=True)
    return result["metadata"]


def run_layout_ocr(
    shards: int, ocr_workers: int = 2, shard_start: int = 0, shard_end: int = 0
) -> None:
    failed = []
    stop = min(shards, shard_end) if shard_end > 0 else shards
    for start in range(max(0, shard_start), stop, max(1, ocr_workers)):
        processes = []
        for shard in range(start, min(stop, start + max(1, ocr_workers))):
            input_dir = RUN_DIR / f"layout_shard_{shard:02d}"
            output_json = RUN_DIR / f"layout_ocr_{shard:02d}.json"
            command = [
                "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                "-File", str(LAYOUT_OCR),
                "-InputDirectory", str(input_dir),
                "-OutputJson", str(output_json),
                "-LanguageTag", "ko",
                "-Resume",
            ]
            processes.append((shard, subprocess.Popen(command)))
        for shard, process in processes:
            code = process.wait()
            print(f"layout_ocr_shard={shard} exit={code}", flush=True)
            if code:
                failed.append(shard)
    if failed:
        raise RuntimeError(f"layout OCR shards failed: {failed}")


def create_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS stg_dimension_spatial_attempt (
            run_name TEXT NOT NULL,
            created_at TEXT NOT NULL,
            product_id TEXT NOT NULL,
            attempt_order INTEGER NOT NULL,
            image_url TEXT,
            download_status TEXT,
            text_ocr_status TEXT,
            dimension_token_count INTEGER NOT NULL DEFAULT 0,
            layout_ocr_status TEXT,
            spatial_candidate INTEGER NOT NULL DEFAULT 0,
            confidence TEXT,
            candidate_w_mm REAL,
            candidate_d_mm REAL,
            candidate_h_mm REAL,
            raw_text TEXT,
            candidate_json TEXT,
            error TEXT,
            PRIMARY KEY(run_name,product_id,attempt_order)
        );
        CREATE TABLE IF NOT EXISTS stg_dimension_spatial_result (
            run_name TEXT NOT NULL,
            created_at TEXT NOT NULL,
            product_id TEXT NOT NULL,
            product_name TEXT,
            small_category TEXT,
            recovered INTEGER NOT NULL DEFAULT 0,
            confidence TEXT,
            w_mm REAL,
            d_mm REAL,
            h_mm REAL,
            pattern_type TEXT,
            candidate_image_count INTEGER NOT NULL DEFAULT 0,
            conflict INTEGER NOT NULL DEFAULT 0,
            evidence_image_url TEXT,
            evidence_text TEXT,
            applied_to_sources INTEGER NOT NULL DEFAULT 0,
            stop_reason TEXT NOT NULL,
            PRIMARY KEY(run_name,product_id)
        );
        """
    )


def analyze(shards: int) -> dict[str, Any]:
    manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    screen = json.loads(SCREEN_PATH.read_text(encoding="utf-8"))
    layout_rows = read_ocr_outputs("layout_ocr", shards)
    screened_by_key = {
        (item["product_id"], int(item["attempt_order"])): item
        for item in screen["products"]
    }
    all_by_product: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for item in manifest["products"]:
        all_by_product[item["product_id"]].append(item)

    created_at = now_text()
    attempt_inserts: list[tuple[Any, ...]] = []
    result_inserts: list[tuple[Any, ...]] = []
    summary = Counter()

    for product_id, items in sorted(all_by_product.items()):
        product_name = str(items[0].get("product_name") or "")
        small_category = str(items[0].get("small_category") or "")
        candidates: list[dict[str, Any]] = []
        for item in sorted(items, key=lambda row: int(row["attempt_order"])):
            order = int(item["attempt_order"])
            screened = screened_by_key.get((product_id, order))
            stem = f"{product_id}__{order:02d}"
            layout = layout_rows.get(stem) or {}
            candidate = None
            if screened and layout.get("status") == "SUCCESS":
                candidate = spatial_callout_candidate(
                    layout,
                    product_name=product_name,
                    small_category=small_category,
                )
                if candidate:
                    candidate["image_url"] = item["image"]["url"]
                    candidate["attempt_order"] = order
                    candidates.append(candidate)
            attempt_inserts.append(
                (
                    RUN_NAME, created_at, product_id, order, item["image"]["url"],
                    item["download_status"],
                    screened.get("text_ocr_status") if screened else "NOT_SCREENED",
                    int(screened.get("dimension_token_count") or 0) if screened else 0,
                    layout.get("status") or "NOT_RUN",
                    int(candidate is not None),
                    candidate.get("confidence") if candidate else None,
                    candidate.get("w_mm") if candidate else None,
                    candidate.get("d_mm") if candidate else None,
                    candidate.get("h_mm") if candidate else None,
                    str(layout.get("text") or screened.get("text_ocr") if screened else ""),
                    json.dumps(candidate, ensure_ascii=False) if candidate else None,
                    item.get("error") or layout.get("error") or "",
                )
            )

        high = [c for c in candidates if c["confidence"] == "HIGH_SPATIAL_TABLE_CALLOUT"]
        distinct = {
            (c["w_mm"], c["d_mm"], c["h_mm"])
            for c in high
        }
        conflict = len(distinct) > 1
        selected = high[0] if len(distinct) == 1 else None
        recovered = selected is not None and not conflict
        if recovered:
            stop_reason = "HIGH_SPATIAL_SINGLE_DIMENSION"
            summary["recovered_high"] += 1
        elif conflict:
            stop_reason = "HIGH_SPATIAL_CONFLICT"
            summary["conflict"] += 1
        elif candidates:
            stop_reason = "REVIEW_SPATIAL_ONLY"
            summary["review_only"] += 1
        else:
            stop_reason = "NO_SPATIAL_PATTERN"
            summary["no_pattern"] += 1
        result_inserts.append(
            (
                RUN_NAME, created_at, product_id, product_name, small_category,
                int(recovered), selected.get("confidence") if selected else None,
                selected.get("w_mm") if selected else None,
                selected.get("d_mm") if selected else None,
                selected.get("h_mm") if selected else None,
                selected.get("pattern_type") if selected else None,
                len(candidates), int(conflict),
                selected.get("image_url") if selected else None,
                selected.get("evidence_text") if selected else None,
                0, stop_reason,
            )
        )

    connection = sqlite3.connect(DB_PATH)
    create_schema(connection)
    with connection:
        connection.execute("DELETE FROM stg_dimension_spatial_attempt WHERE run_name=?", (RUN_NAME,))
        connection.execute("DELETE FROM stg_dimension_spatial_result WHERE run_name=?", (RUN_NAME,))
        connection.executemany(
            "INSERT INTO stg_dimension_spatial_attempt VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            attempt_inserts,
        )
        connection.executemany(
            "INSERT INTO stg_dimension_spatial_result VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            result_inserts,
        )
    connection.close()
    result = {
        "run_name": RUN_NAME,
        "created_at": created_at,
        "target_products": len(all_by_product),
        "attempt_rows": len(attempt_inserts),
        "layout_candidate_images": len(screen["products"]),
        "summary": dict(summary),
        "sources_db_written": False,
        "staging_tables": ["stg_dimension_spatial_attempt", "stg_dimension_spatial_result"],
    }
    SUMMARY_PATH.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2), flush=True)
    return result


def apply() -> dict[str, Any]:
    connection = sqlite3.connect(DB_PATH)
    create_schema(connection)
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
        SELECT product_id,w_mm,d_mm,h_mm,confidence,pattern_type,
               evidence_image_url,evidence_text
        FROM stg_dimension_spatial_result
        WHERE run_name=? AND recovered=1 AND conflict=0
          AND confidence='HIGH_SPATIAL_TABLE_CALLOUT'
        ORDER BY product_id
        """,
        (RUN_NAME,),
    ).fetchall()
    applied_at = now_text()
    applied: list[dict[str, Any]] = []
    skipped: list[dict[str, str]] = []
    with connection:
        for product_id, w_mm, d_mm, h_mm, confidence, pattern_type, image_url, evidence_text in rows:
            source = connection.execute(
                "SELECT ocr_status,ocr_blob FROM sources WHERE product_id=?", (product_id,)
            ).fetchone()
            if not source:
                skipped.append({"product_id": product_id, "reason": "SOURCE_NOT_FOUND"})
                continue
            old_status, old_blob = source
            ocr_blob = unpack(old_blob) or {}
            reinforcements = list(ocr_blob.get("dimension_reinforcements") or [])
            if any(item.get("run_name") == RUN_NAME for item in reinforcements):
                skipped.append({"product_id": product_id, "reason": "ALREADY_APPLIED"})
                connection.execute(
                    "UPDATE stg_dimension_spatial_result SET applied_to_sources=1 "
                    "WHERE run_name=? AND product_id=?", (RUN_NAME, product_id)
                )
                continue
            connection.execute(
                "INSERT OR IGNORE INTO source_ocr_dimension_backup VALUES (?,?,?,?,?)",
                (RUN_NAME, product_id, applied_at, old_status, old_blob),
            )
            explicit_line = (
                f"추가 공간좌표 OCR 검증 W={w_mm} mm D={d_mm} mm H={h_mm} mm"
            )
            old_dimension_text = str(ocr_blob.get("dimension_text") or "").strip()
            merged = "\n".join(
                value for value in (old_dimension_text, str(evidence_text or "").strip(), explicit_line)
                if value and value not in old_dimension_text
            )
            ocr_blob["dimension_text"] = "\n".join(
                value for value in (old_dimension_text, merged) if value
            )
            reinforcements.append(
                {
                    "run_name": RUN_NAME,
                    "applied_at": applied_at,
                    "image_url": image_url,
                    "confidence": confidence,
                    "pattern_type": pattern_type,
                    "w_mm": w_mm,
                    "d_mm": d_mm,
                    "h_mm": h_mm,
                    "evidence_text": evidence_text,
                }
            )
            ocr_blob["dimension_reinforcements"] = reinforcements
            connection.execute(
                "UPDATE sources SET ocr_status=200,ocr_blob=?,ocr_at=? WHERE product_id=?",
                (pack(ocr_blob), applied_at, product_id),
            )
            connection.execute(
                "UPDATE stg_dimension_spatial_result SET applied_to_sources=1 "
                "WHERE run_name=? AND product_id=?", (RUN_NAME, product_id)
            )
            applied.append(
                {"product_id": product_id, "w_mm": w_mm, "d_mm": d_mm, "h_mm": h_mm}
            )
    connection.close()
    result = {
        "run_name": RUN_NAME,
        "applied_at": applied_at,
        "candidate_count": len(rows),
        "applied_count": len(applied),
        "skipped_count": len(skipped),
        "applied": applied,
        "skipped": skipped,
        "backup_table": "source_ocr_dimension_backup",
        "excel_written": False,
    }
    (RUN_DIR / "apply.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(result, ensure_ascii=False, indent=2), flush=True)
    return result


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "phase",
        choices=("prepare", "text", "screen", "layout", "analyze", "apply", "all"),
    )
    parser.add_argument("--product-id", default="")
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--max-candidates", type=int, default=DEFAULT_MAX_CANDIDATES)
    parser.add_argument("--workers", type=int, default=24)
    parser.add_argument("--ocr-workers", type=int, default=2)
    parser.add_argument("--shard-start", type=int, default=0)
    parser.add_argument("--shard-end", type=int, default=0)
    parser.add_argument("--shards", type=int, default=DEFAULT_SHARDS)
    args = parser.parse_args()
    if args.phase in ("prepare", "all"):
        prepare(
            product_id=args.product_id,
            limit=args.limit,
            max_candidates=args.max_candidates,
            workers=args.workers,
            shards=args.shards,
        )
    if args.phase in ("text", "all"):
        run_text_ocr(
            args.shards, args.ocr_workers, args.shard_start, args.shard_end
        )
    if args.phase in ("screen", "all"):
        screen_layout_candidates(args.shards)
    if args.phase in ("layout", "all"):
        run_layout_ocr(
            args.shards, args.ocr_workers, args.shard_start, args.shard_end
        )
    if args.phase in ("analyze", "all"):
        analyze(args.shards)
    if args.phase in ("apply",):
        apply()


if __name__ == "__main__":
    main()
