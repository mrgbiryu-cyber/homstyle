from __future__ import annotations

import argparse
import concurrent.futures
import hashlib
import html
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

from PIL import Image, ImageFile

import bulk_homestyle_ocr as ocr
from bulk_homestyle_collect import DB_PATH, request_bytes, unpack


RUN_NAME = "dimension_scan_pass1_all_detail_images_v1"
RUN_DIR = ocr.OCR_ROOT / RUN_NAME
MANIFEST_PATH = RUN_DIR / "manifest.json"
REUSED_OCR_PATH = RUN_DIR / "reused_ocr.json"
SUMMARY_PATH = ocr.RUN_DIR / f"{RUN_NAME}.json"
DIRECT_OCR = Path(__file__).resolve().parent / "run_windows_ocr_direct.ps1"
DEFAULT_SHARDS = 8
DEFAULT_MAX_WIDTH = 1400
NON_STILL_EXTENSIONS = {"mp4", "mov", "avi", "webm", "m3u8", "mpd"}

ImageFile.LOAD_TRUNCATED_IMAGES = True
Image.MAX_IMAGE_PIXELS = None


SIZE_LABEL_RE = re.compile(
    r"(?:제품\s*(?:사이즈|크기|규격|치수)|상품\s*(?:사이즈|크기|규격|치수)|"
    r"사이즈|규격|치수|크기|한\s*눈에\s*보기|size|dimensions?|"
    r"measurements?|specifications?|info(?:rmation)?|detail|"
    r"overall\s+dimensions?|product\s+(?:description|dimensions?))",
    re.I,
)
UNIT_RE = re.compile(r"(?<![a-z])(?:mm|cm|㎜|㎝|millimet(?:er|re)s?|centimet(?:er|re)s?)(?![a-z])", re.I)
TRIPLE_RE = re.compile(
    r"(?<!\d)\d{2,5}(?:[.,]\d+)?\s*(?:x|×|X|＊|\*)\s*"
    r"\d{2,5}(?:[.,]\d+)?\s*(?:x|×|X|＊|\*)\s*\d{2,5}(?:[.,]\d+)?",
    re.I,
)
PAIR_RE = re.compile(
    r"(?<!\d)\d{2,5}(?:[.,]\d+)?\s*(?:x|×|X|＊|\*)\s*"
    r"\d{2,5}(?:[.,]\d+)?",
    re.I,
)
NUMBER_RE = re.compile(r"(?<!\d)\d{2,5}(?:[.,]\d+)?(?!\d)")
URL_HINT_RE = re.compile(r"(?:size|dimension|measure|drawing|spec|dim[_-])", re.I)
GENERIC_NOTICE_URL_RE = re.compile(
    r"(?:notice|footer|policy|delivery|shipping|customer|return|refund|faq)",
    re.I,
)
KOREAN_AXIS_WORD_RE = re.compile(r"(?:가로|세로|폭|너비|깊이|높이|길이|지름|직경|반지름)")
KOREAN_AXIS_RE = re.compile(
    r"(?:(가로|세로|폭|너비|깊이|높이|길이|지름|직경|반지름)\s*[:=]?\s*"
    r"\d{2,5}(?:[.,]\d+)?|\d{2,5}(?:[.,]\d+)?\s*(?:mm|cm|㎜|㎝)?\s*"
    r"(가로|세로|폭|너비|깊이|높이|길이|지름|직경|반지름))",
    re.I,
)
LATIN_AXIS_RE = re.compile(
    r"(?<![A-Za-z])([WDLH])\s*(?:[:=]|\)|\(|\]|\[|x|×|-)?\s*\d{2,5}(?:[.,]\d+)?",
    re.I,
)
DIAMETER_RE = re.compile(r"(?:Ø|Φ|ø|직경|지름)\s*[:=]?\s*\d{2,5}(?:[.,]\d+)?", re.I)


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def url_hash(url: str) -> str:
    return hashlib.sha1(url.encode("utf-8")).hexdigest()


def image_extension(url: str) -> str:
    path = url.split("?", 1)[0].casefold()
    return path.rsplit(".", 1)[-1] if "." in path else ""


def row_url(row: dict[str, Any]) -> str:
    for key in ("image", "selected"):
        value = row.get(key)
        if isinstance(value, dict):
            candidate = value.get("url") or value.get("image_url")
            if candidate:
                return ocr.normalize_url(str(candidate))
    return ocr.normalize_url(
        str(row.get("image_url") or row.get("url") or "")
    )


def existing_path(raw: str, manifest_path: Path) -> Path | None:
    if not raw:
        return None
    candidates = [Path(raw)]
    if not Path(raw).is_absolute():
        candidates.extend([Path.cwd() / raw, manifest_path.parent / raw])
    for candidate in candidates:
        try:
            if candidate.exists() and candidate.is_file() and candidate.stat().st_size:
                return candidate.resolve()
        except OSError:
            continue
    return None


def load_json(path: Path) -> Any:
    try:
        raw = path.read_text(encoding="utf-8-sig").strip()
        return json.loads(raw) if raw else None
    except (OSError, UnicodeError, json.JSONDecodeError):
        return None


def discover_previous_cache() -> tuple[dict[str, Path], dict[str, dict[str, Any]]]:
    """Return URL-indexed prepared image files and successful direct OCR text."""
    image_cache: dict[str, Path] = {}
    text_cache: dict[str, dict[str, Any]] = {}
    if not ocr.OCR_ROOT.exists():
        return image_cache, text_cache

    for manifest_path in ocr.OCR_ROOT.rglob("manifest.json"):
        if RUN_DIR in manifest_path.parents:
            continue
        manifest = load_json(manifest_path)
        if not isinstance(manifest, dict):
            continue
        rows = manifest.get("products") or manifest.get("images") or []
        if not isinstance(rows, list):
            continue
        url_by_file: dict[str, str] = {}
        for row in rows:
            if not isinstance(row, dict):
                continue
            url = row_url(row)
            if not url.startswith(("http://", "https://")):
                continue
            source = existing_path(str(row.get("file") or ""), manifest_path)
            if source:
                image_cache.setdefault(url, source)
                url_by_file[source.name.casefold()] = url
            raw_file = Path(str(row.get("file") or "")).name.casefold()
            if raw_file:
                url_by_file.setdefault(raw_file, url)

        for result_path in manifest_path.parent.glob("*.json"):
            lower_name = result_path.name.casefold()
            if result_path == manifest_path or "ocr" not in lower_name:
                continue
            result = load_json(result_path)
            if isinstance(result, dict):
                result = result.get("products") or result.get("results") or [result]
            if not isinstance(result, list):
                continue
            for value in result:
                if not isinstance(value, dict) or value.get("status") != "SUCCESS":
                    continue
                text = str(value.get("text") or "").strip()
                if not text:
                    continue
                name = Path(str(value.get("file") or "")).name.casefold()
                url = url_by_file.get(name)
                if not url:
                    continue
                old = text_cache.get(url)
                if old is None or len(text) > len(str(old.get("text") or "")):
                    text_cache[url] = {
                        "url": url,
                        "status": "SUCCESS",
                        "text": text,
                        "origin": str(result_path),
                        "file": name,
                    }
    return image_cache, text_cache


def discover_database_cache() -> tuple[
    dict[str, Path], dict[str, dict[str, Any]]
]:
    """Build the reusable OCR cache from DB staging instead of scanning JSON files.

    The filesystem-wide manifest scan becomes very slow after spatial/layout OCR
    creates thousands of result files. The DB already contains the URL, local
    file and OCR text needed by pass 1, so use those indexed rows directly.
    """
    image_cache: dict[str, Path] = {}
    text_cache: dict[str, dict[str, Any]] = {}
    connection = sqlite3.connect(DB_PATH)
    sources = (
        (
            "stg_dimension_scan_pass1_unique_image",
            "image_url",
            "local_file",
            "ocr_text",
            "run_name",
        ),
        (
            "stg_dimension_scan_pass2_unique_image",
            "image_url",
            "local_file",
            "ocr_text",
            "run_name",
        ),
        (
            "stg_dimension_targeted_ocr_image",
            "image_url",
            "local_file",
            "stage1_text",
            "run_name",
        ),
    )
    for table, url_column, file_column, text_column, origin_column in sources:
        exists = connection.execute(
            "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=?",
            (table,),
        ).fetchone()[0]
        if not exists:
            continue
        query = (
            f"SELECT {url_column},{file_column},{text_column},{origin_column} "
            f"FROM {table} WHERE COALESCE({url_column},'')<>''"
        )
        for url, local_file, text, origin in connection.execute(query):
            normalized = ocr.normalize_url(str(url or ""))
            if not normalized:
                continue
            path = Path(str(local_file or ""))
            if path.exists() and path.is_file() and path.stat().st_size:
                image_cache.setdefault(normalized, path)
            value = str(text or "").strip()
            if value:
                old = text_cache.get(normalized)
                if old is None or len(value) > len(str(old.get("text") or "")):
                    text_cache[normalized] = {
                        "url": normalized,
                        "status": "SUCCESS",
                        "text": value,
                        "origin": f"{table}:{origin}",
                        "file": path.name if path.name else "",
                    }

    for (ocr_blob,) in connection.execute(
        "SELECT ocr_blob FROM sources WHERE ocr_blob IS NOT NULL"
    ):
        blob = unpack(ocr_blob) or {}
        selected = blob.get("selected") or {}
        normalized = ocr.normalize_url(str(selected.get("url") or ""))
        value = str(
            blob.get("combined_text")
            or blob.get("dimension_text")
            or (blob.get("ocr") or {}).get("text")
            or ""
        ).strip()
        if normalized and value:
            old = text_cache.get(normalized)
            if old is None or len(value) > len(str(old.get("text") or "")):
                text_cache[normalized] = {
                    "url": normalized,
                    "status": "SUCCESS",
                    "text": value,
                    "origin": "sources.ocr_blob",
                    "file": "",
                }
    connection.close()
    return image_cache, text_cache


def target_rows(product_id: str = "", limit: int = 0) -> list[tuple[Any, ...]]:
    connection = sqlite3.connect(DB_PATH)
    sql = """
        SELECT
            l.product_id,
            l.product_name,
            l.small_category,
            s.goods_blob
        FROM fact_dimension_resolution_ledger l
        JOIN sources s ON s.product_id = l.product_id
        WHERE l.needs_ocr = 1
          AND l.is_locked = 0
    """
    params: list[Any] = []
    if product_id:
        sql += " AND l.product_id=?"
        params.append(product_id)
    sql += " ORDER BY l.product_id"
    if limit:
        sql += " LIMIT ?"
        params.append(limit)
    rows = connection.execute(sql, params).fetchall()
    connection.close()
    return rows


def link_or_convert(source: Path, destination: Path, max_width: int) -> tuple[int, int]:
    destination.parent.mkdir(parents=True, exist_ok=True)
    with Image.open(source) as raw:
        raw.seek(0)
        image = raw.convert("RGB")
        if image.width > max_width:
            height = max(1, round(image.height * max_width / image.width))
            image = image.resize((max_width, height), Image.Resampling.LANCZOS)
        image.save(destination, "JPEG", quality=86, optimize=True)
        return image.size


def prepare_one(
    item: dict[str, Any], image_cache: dict[str, Path], shards: int, max_width: int
) -> dict[str, Any]:
    url = item["image_url"]
    digest = item["url_hash"]
    shard = int(digest[:8], 16) % shards
    destination = RUN_DIR / f"shard_{shard:02d}" / f"{digest}.jpg"
    try:
        if destination.exists() and destination.stat().st_size:
            with Image.open(destination) as image:
                width, height = image.size
            return {
                **item,
                "shard": shard,
                "file": str(destination),
                "file_name": destination.name,
                "download_status": "REUSED_RUN_FILE",
                "width": width,
                "height": height,
                "error": "",
            }

        cached = image_cache.get(url)
        if cached:
            width, height = link_or_convert(cached, destination, max_width)
            return {
                **item,
                "shard": shard,
                "file": str(destination),
                "file_name": destination.name,
                "download_status": "REUSED_IMAGE_CACHE",
                "width": width,
                "height": height,
                "error": "",
            }

        status, body = request_bytes(url, timeout=30, attempts=2)
        if status != 200 or not body:
            raise RuntimeError(f"HTTP {status}; bytes={len(body)}")
        destination.parent.mkdir(parents=True, exist_ok=True)
        with Image.open(io.BytesIO(body)) as raw:
            raw.seek(0)
            image = raw.convert("RGB")
            if image.width > max_width:
                height = max(1, round(image.height * max_width / image.width))
                image = image.resize((max_width, height), Image.Resampling.LANCZOS)
            image.save(destination, "JPEG", quality=86, optimize=True)
            width, height = image.size
        return {
            **item,
            "shard": shard,
            "file": str(destination),
            "file_name": destination.name,
            "download_status": "DOWNLOADED",
            "width": width,
            "height": height,
            "error": "",
        }
    except Exception as exc:
        return {
            **item,
            "shard": shard,
            "file": "",
            "file_name": "",
            "download_status": "ERROR",
            "width": None,
            "height": None,
            "error": f"{type(exc).__name__}: {exc}",
        }


def prepare(
    *,
    product_id: str = "",
    limit: int = 0,
    workers: int = 8,
    shards: int = DEFAULT_SHARDS,
    max_width: int = DEFAULT_MAX_WIDTH,
) -> dict[str, Any]:
    targets = target_rows(product_id, limit)
    product_images: list[dict[str, Any]] = []
    unique: dict[str, dict[str, Any]] = {}
    products_without_images: list[str] = []

    for pid, name, category, goods_blob in targets:
        data = (unpack(goods_blob) or {}).get("data") or {}
        images = sorted(
            ocr.detail_images(str(data.get("detailInfo") or "")),
            key=lambda value: int(value.get("position") or 0),
        )
        # Image proxy URLs often end in `/optimize` without a file extension.
        # Keep those and let Pillow validate the response; exclude only known
        # video/stream extensions.
        stills = [
            value
            for value in images
            if image_extension(str(value.get("url") or "")) not in NON_STILL_EXTENSIONS
        ]
        if not stills:
            products_without_images.append(pid)
        for order, image in enumerate(stills, 1):
            url = ocr.normalize_url(str(image.get("url") or ""))
            digest = url_hash(url)
            product_images.append(
                {
                    "product_id": pid,
                    "product_name": str(name or ""),
                    "small_category": str(category or ""),
                    "image_order": order,
                    "source_position": int(image.get("position") or 0),
                    "image_url": url,
                    "url_hash": digest,
                }
            )
            if digest not in unique:
                unique[digest] = {
                    "url_hash": digest,
                    "image_url": url,
                    "reference_count": 0,
                    "first_product_id": pid,
                    "first_image_order": order,
                }
            unique[digest]["reference_count"] += 1

    RUN_DIR.mkdir(parents=True, exist_ok=True)
    for shard in range(shards):
        (RUN_DIR / f"shard_{shard:02d}").mkdir(parents=True, exist_ok=True)

    image_cache, text_cache = discover_database_cache()
    reused_ocr = [text_cache[row["image_url"]] for row in unique.values() if row["image_url"] in text_cache]
    REUSED_OCR_PATH.write_text(
        json.dumps(reused_ocr, ensure_ascii=False), encoding="utf-8"
    )
    to_prepare = [row for row in unique.values() if row["image_url"] not in text_cache]
    results: list[dict[str, Any]] = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, workers)) as pool:
        futures = [
            pool.submit(prepare_one, row, image_cache, shards, max_width)
            for row in to_prepare
        ]
        for index, future in enumerate(concurrent.futures.as_completed(futures), 1):
            results.append(future.result())
            if index % 250 == 0 or index == len(futures):
                print(f"prepared={index}/{len(futures)}", flush=True)

    prepared_by_hash = {row["url_hash"]: row for row in results}
    unique_rows: list[dict[str, Any]] = []
    for digest, row in unique.items():
        cached = text_cache.get(row["image_url"])
        if cached:
            unique_rows.append(
                {
                    **row,
                    "shard": int(digest[:8], 16) % shards,
                    "file": "",
                    "file_name": "",
                    "download_status": "OCR_TEXT_REUSED",
                    "width": None,
                    "height": None,
                    "error": "",
                    "ocr_reused": True,
                    "ocr_origin": cached.get("origin") or "",
                }
            )
        else:
            unique_rows.append(
                {
                    **prepared_by_hash[digest],
                    "ocr_reused": False,
                    "ocr_origin": "",
                }
            )
    unique_rows.sort(key=lambda value: value["url_hash"])
    product_images.sort(key=lambda value: (value["product_id"], value["image_order"]))
    metadata = {
        "run_name": RUN_NAME,
        "created_at": now_text(),
        "target_products": len(targets),
        "product_image_rows": len(product_images),
        "unique_images": len(unique_rows),
        "duplicate_references_saved": len(product_images) - len(unique_rows),
        "reused_previous_ocr": len(reused_ocr),
        "new_ocr_images": len(to_prepare),
        "prepared_success": sum(row.get("download_status") != "ERROR" for row in results),
        "prepared_error": sum(row.get("download_status") == "ERROR" for row in results),
        "products_without_images": len(products_without_images),
        "products_without_image_ids": products_without_images,
        "shards": shards,
        "max_width": max_width,
        "source_dimension_values_written": False,
        "excel_written": False,
    }
    MANIFEST_PATH.write_text(
        json.dumps(
            {
                "metadata": metadata,
                "unique_images": unique_rows,
                "product_images": product_images,
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    print(json.dumps(metadata, ensure_ascii=False, indent=2), flush=True)
    return metadata


def run_ocr(
    shards: int = DEFAULT_SHARDS,
    workers: int = 4,
    shard_start: int = 0,
    shard_end: int = 0,
) -> None:
    stop = min(shards, shard_end) if shard_end > 0 else shards
    failed: list[int] = []
    for start in range(max(0, shard_start), stop, max(1, workers)):
        processes: list[tuple[int, subprocess.Popen[Any]]] = []
        for shard in range(start, min(stop, start + max(1, workers))):
            command = [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(DIRECT_OCR),
                "-InputDirectory",
                str(RUN_DIR / f"shard_{shard:02d}"),
                "-OutputJson",
                str(RUN_DIR / f"ocr_{shard:02d}.json"),
                "-LanguageTag",
                "ko",
                "-Resume",
            ]
            processes.append((shard, subprocess.Popen(command)))
        for shard, process in processes:
            code = process.wait()
            print(f"ocr_shard={shard} exit={code}", flush=True)
            if code:
                failed.append(shard)
    if failed:
        raise RuntimeError(f"OCR shards failed: {failed}")


def load_ocr_rows(shards: int) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for shard in range(shards):
        path = RUN_DIR / f"ocr_{shard:02d}.json"
        data = load_json(path)
        if isinstance(data, dict):
            data = [data]
        if not isinstance(data, list):
            continue
        for row in data:
            if isinstance(row, dict):
                result[Path(str(row.get("file") or "")).stem] = row
    return result


def signal_excerpt(text: str) -> str:
    compact = re.sub(r"[ \t]+", " ", text or "").strip()
    lines = [line.strip() for line in compact.splitlines() if line.strip()]
    matched: list[str] = []
    for index, line in enumerate(lines):
        if (
            SIZE_LABEL_RE.search(line)
            or UNIT_RE.search(line)
            or LATIN_AXIS_RE.search(line)
            or KOREAN_AXIS_WORD_RE.search(line)
            or TRIPLE_RE.search(line)
        ):
            for value in lines[max(0, index - 1) : min(len(lines), index + 2)]:
                if value not in matched:
                    matched.append(value)
    return " | ".join(matched)[:2000]


def classify_text(text: str, url: str, ocr_status: str) -> dict[str, Any]:
    value = html.unescape(str(text or ""))
    label_matches = list(SIZE_LABEL_RE.finditer(value))
    label_count = len(label_matches)
    latin_axes = {match.upper() for match in LATIN_AXIS_RE.findall(value)}
    korean_axes = {
        next((part for part in match if part), "")
        for match in KOREAN_AXIS_RE.findall(value)
    }
    korean_axes.discard("")
    axis_count = len(latin_axes) + len(korean_axes)
    unit_count = len(UNIT_RE.findall(value))
    number_count = len(NUMBER_RE.findall(value))
    triple_count = len(TRIPLE_RE.findall(value))
    pair_count = len(PAIR_RE.findall(value))
    diameter_count = len(DIAMETER_RE.findall(value))
    max_line_numbers = max(
        (len(NUMBER_RE.findall(line)) for line in value.splitlines()), default=0
    )
    url_hint = bool(URL_HINT_RE.search(url))
    label_near_number_count = 0
    label_near_unit = False
    label_near_axis = False
    label_near_pair = False
    for match in label_matches:
        nearby = value[max(0, match.start() - 180) : min(len(value), match.end() + 240)]
        label_near_number_count = max(
            label_near_number_count, len(NUMBER_RE.findall(nearby))
        )
        label_near_unit = label_near_unit or bool(UNIT_RE.search(nearby))
        label_near_axis = label_near_axis or bool(
            LATIN_AXIS_RE.search(nearby) or KOREAN_AXIS_RE.search(nearby)
        )
        label_near_pair = label_near_pair or bool(PAIR_RE.search(nearby))
    unit_near_number = False
    for match in UNIT_RE.finditer(value):
        nearby = value[max(0, match.start() - 80) : min(len(value), match.end() + 80)]
        if NUMBER_RE.search(nearby):
            unit_near_number = True
            break
    reasons: list[str] = []
    if label_count:
        reasons.append("SIZE_LABEL")
    if axis_count:
        reasons.append("AXIS_TOKEN")
    if unit_count:
        reasons.append("DIMENSION_UNIT")
    if triple_count:
        reasons.append("THREE_VALUE_PATTERN")
    if pair_count:
        reasons.append("TWO_VALUE_PATTERN")
    if diameter_count:
        reasons.append("DIAMETER_PATTERN")
    if max_line_numbers >= 3:
        reasons.append("NUMERIC_CLUSTER")
    if url_hint:
        reasons.append("URL_SEMANTIC_HINT")

    if label_count and label_near_axis and label_near_unit and axis_count >= 2:
        classification = "SIZE_LABEL_AXIS_UNIT"
    elif label_count and label_near_axis:
        classification = "SIZE_LABEL_WITH_AXIS"
    elif label_count and label_near_unit:
        classification = "SIZE_LABEL_WITH_UNIT"
    elif label_count and (label_near_number_count >= 2 or label_near_pair):
        classification = "SIZE_LABEL_UNIT_MISSING"
    elif triple_count:
        classification = "DIMENSION_TRIPLE_FOUND"
    elif pair_count and (unit_near_number or url_hint):
        classification = "DIMENSION_PAIR_FOUND"
    elif axis_count >= 2:
        classification = "MULTI_AXIS_FOUND"
    elif axis_count or diameter_count:
        classification = "PARTIAL_AXIS_FOUND"
    elif unit_count and unit_near_number:
        classification = "UNIT_NUMERIC_CLUSTER"
    elif url_hint:
        classification = "URL_HINT_ONLY"
    elif label_count:
        classification = "WEAK_SIZE_WORD_ONLY"
    elif ocr_status not in {"REUSED_SUCCESS", "NEW_SUCCESS"}:
        classification = "SCAN_OCR_ERROR"
    else:
        classification = "NO_SIZE_SIGNAL"

    candidate = classification not in {
        "NO_SIZE_SIGNAL",
        "SCAN_OCR_ERROR",
        "WEAK_SIZE_WORD_ONLY",
    }
    score = (
        min(label_count, 2) * 5
        + min(axis_count, 4) * 2
        + min(unit_count, 2) * 3
        + min(triple_count, 1) * 5
        + min(pair_count, 1) * 3
        + min(diameter_count, 1) * 3
        + (2 if max_line_numbers >= 3 else 0)
        + (1 if url_hint else 0)
    )
    return {
        "is_candidate": int(candidate),
        "candidate_score": score,
        "classification": classification,
        "reason_codes": json.dumps(reasons, ensure_ascii=False),
        "size_label_count": label_count,
        "axis_count": axis_count,
        "axes": json.dumps(sorted(latin_axes | korean_axes), ensure_ascii=False),
        "unit_count": unit_count,
        "number_count": number_count,
        "triple_count": triple_count,
        "max_line_number_count": max_line_numbers,
        "signal_excerpt": signal_excerpt(value),
    }


def create_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS stg_dimension_scan_pass1_unique_image (
            run_name TEXT NOT NULL,
            created_at TEXT NOT NULL,
            url_hash TEXT NOT NULL,
            image_url TEXT NOT NULL,
            reference_count INTEGER NOT NULL,
            first_product_id TEXT,
            first_image_order INTEGER,
            local_file TEXT,
            download_status TEXT,
            ocr_status TEXT,
            ocr_origin TEXT,
            ocr_text TEXT,
            is_candidate INTEGER NOT NULL,
            candidate_score INTEGER NOT NULL,
            classification TEXT NOT NULL,
            reason_codes TEXT NOT NULL,
            size_label_count INTEGER NOT NULL,
            axis_count INTEGER NOT NULL,
            axes TEXT NOT NULL,
            unit_count INTEGER NOT NULL,
            number_count INTEGER NOT NULL,
            triple_count INTEGER NOT NULL,
            max_line_number_count INTEGER NOT NULL,
            signal_excerpt TEXT,
            error TEXT,
            PRIMARY KEY(run_name,url_hash)
        );

        CREATE TABLE IF NOT EXISTS stg_dimension_scan_pass1_product_image (
            run_name TEXT NOT NULL,
            created_at TEXT NOT NULL,
            product_id TEXT NOT NULL,
            product_name TEXT,
            small_category TEXT,
            image_order INTEGER NOT NULL,
            source_position INTEGER,
            url_hash TEXT NOT NULL,
            image_url TEXT NOT NULL,
            ocr_status TEXT,
            is_candidate INTEGER NOT NULL,
            candidate_score INTEGER NOT NULL,
            classification TEXT NOT NULL,
            reason_codes TEXT NOT NULL,
            signal_excerpt TEXT,
            PRIMARY KEY(run_name,product_id,image_order)
        );

        CREATE TABLE IF NOT EXISTS stg_dimension_scan_pass1_product (
            run_name TEXT NOT NULL,
            created_at TEXT NOT NULL,
            product_id TEXT NOT NULL,
            product_name TEXT,
            small_category TEXT,
            total_image_count INTEGER NOT NULL,
            ocr_success_image_count INTEGER NOT NULL,
            ocr_error_image_count INTEGER NOT NULL,
            candidate_image_count INTEGER NOT NULL,
            candidate_image_orders TEXT NOT NULL,
            candidate_image_urls TEXT NOT NULL,
            product_classification TEXT NOT NULL,
            second_pass_required INTEGER NOT NULL,
            PRIMARY KEY(run_name,product_id)
        );

        CREATE INDEX IF NOT EXISTS idx_dimension_scan_pass1_unique_candidate
            ON stg_dimension_scan_pass1_unique_image(run_name,is_candidate,classification);
        CREATE INDEX IF NOT EXISTS idx_dimension_scan_pass1_product_image_candidate
            ON stg_dimension_scan_pass1_product_image(run_name,product_id,is_candidate,image_order);
        CREATE INDEX IF NOT EXISTS idx_dimension_scan_pass1_product_class
            ON stg_dimension_scan_pass1_product(run_name,product_classification);

        CREATE VIEW IF NOT EXISTS vw_dimension_scan_pass1_candidate_images AS
        SELECT
            p.run_name,
            p.product_id,
            p.product_name,
            p.small_category,
            p.image_order,
            p.image_url,
            p.ocr_status,
            p.candidate_score,
            p.classification,
            p.reason_codes,
            p.signal_excerpt,
            u.reference_count,
            u.local_file
        FROM stg_dimension_scan_pass1_product_image p
        JOIN stg_dimension_scan_pass1_unique_image u
          ON u.run_name=p.run_name AND u.url_hash=p.url_hash
        WHERE p.is_candidate=1;
        """
    )


def classify(shards: int = DEFAULT_SHARDS) -> dict[str, Any]:
    manifest = load_json(MANIFEST_PATH)
    if not isinstance(manifest, dict):
        raise RuntimeError(f"Manifest not found or invalid: {MANIFEST_PATH}")
    created_at = now_text()
    reused_raw = load_json(REUSED_OCR_PATH)
    reused = {
        row["url"]: row
        for row in (reused_raw if isinstance(reused_raw, list) else [])
        if isinstance(row, dict) and row.get("url")
    }
    new_ocr = load_ocr_rows(shards)
    classified_by_hash: dict[str, dict[str, Any]] = {}
    unique_inserts: list[tuple[Any, ...]] = []
    unique_status = Counter()
    unique_classes = Counter()

    for item in manifest.get("unique_images") or []:
        url = item["image_url"]
        if item.get("ocr_reused") and url in reused:
            text = str(reused[url].get("text") or "")
            ocr_status = "REUSED_SUCCESS"
            origin = str(reused[url].get("origin") or item.get("ocr_origin") or "")
            error = ""
        else:
            row = new_ocr.get(item["url_hash"]) or {}
            text = str(row.get("text") or "")
            if row.get("status") == "SUCCESS":
                ocr_status = "NEW_SUCCESS"
                error = ""
            elif item.get("download_status") == "ERROR":
                ocr_status = "DOWNLOAD_ERROR"
                error = str(item.get("error") or "")
            elif row:
                ocr_status = "OCR_ERROR"
                error = str(row.get("error") or "")
            else:
                ocr_status = "OCR_NOT_RUN"
                error = "OCR output missing"
            origin = str(RUN_DIR / f"ocr_{int(item.get('shard') or 0):02d}.json")
        signals = classify_text(text, url, ocr_status)
        if (
            signals["is_candidate"]
            and GENERIC_NOTICE_URL_RE.search(url)
            and not signals["triple_count"]
            and signals["axis_count"] < 3
        ):
            reasons = json.loads(signals["reason_codes"])
            reasons.append("GENERIC_NOTICE_URL")
            signals["is_candidate"] = 0
            signals["classification"] = "GENERIC_NOTICE_IMAGE_SUPPRESSED"
            signals["reason_codes"] = json.dumps(reasons, ensure_ascii=False)
        result = {
            **signals,
            "ocr_status": ocr_status,
            "ocr_text": text,
            "ocr_origin": origin,
            "error": error,
        }
        classified_by_hash[item["url_hash"]] = result
        unique_status[ocr_status] += 1
        unique_classes[signals["classification"]] += 1
        unique_inserts.append(
            (
                RUN_NAME,
                created_at,
                item["url_hash"],
                url,
                int(item.get("reference_count") or 0),
                item.get("first_product_id"),
                item.get("first_image_order"),
                item.get("file") or "",
                item.get("download_status") or "",
                ocr_status,
                origin,
                text,
                signals["is_candidate"],
                signals["candidate_score"],
                signals["classification"],
                signals["reason_codes"],
                signals["size_label_count"],
                signals["axis_count"],
                signals["axes"],
                signals["unit_count"],
                signals["number_count"],
                signals["triple_count"],
                signals["max_line_number_count"],
                signals["signal_excerpt"],
                error,
            )
        )

    by_product: dict[str, list[dict[str, Any]]] = defaultdict(list)
    product_image_inserts: list[tuple[Any, ...]] = []
    for item in manifest.get("product_images") or []:
        signals = classified_by_hash[item["url_hash"]]
        merged = {**item, **signals}
        by_product[item["product_id"]].append(merged)
        product_image_inserts.append(
            (
                RUN_NAME,
                created_at,
                item["product_id"],
                item.get("product_name") or "",
                item.get("small_category") or "",
                int(item["image_order"]),
                int(item.get("source_position") or 0),
                item["url_hash"],
                item["image_url"],
                signals["ocr_status"],
                signals["is_candidate"],
                signals["candidate_score"],
                signals["classification"],
                signals["reason_codes"],
                signals["signal_excerpt"],
            )
        )

    product_inserts: list[tuple[Any, ...]] = []
    product_classes = Counter()
    target_meta = {
        pid: (str(name or ""), str(category or ""))
        for pid, name, category, _ in target_rows()
    }
    for pid, (name, category) in target_meta.items():
        images = sorted(by_product.get(pid, []), key=lambda value: value["image_order"])
        candidates = [value for value in images if value["is_candidate"]]
        success_count = sum(
            value["ocr_status"] in {"REUSED_SUCCESS", "NEW_SUCCESS"} for value in images
        )
        error_count = len(images) - success_count
        if not images:
            product_class = "NO_DETAIL_IMAGES"
        elif candidates:
            product_class = "CANDIDATE_IMAGES_FOUND"
        elif error_count:
            product_class = "SCAN_INCOMPLETE"
        else:
            product_class = "NO_SIZE_SIGNAL_ALL_IMAGES"
        product_classes[product_class] += 1
        product_inserts.append(
            (
                RUN_NAME,
                created_at,
                pid,
                name,
                category,
                len(images),
                success_count,
                error_count,
                len(candidates),
                json.dumps([value["image_order"] for value in candidates], ensure_ascii=False),
                json.dumps([value["image_url"] for value in candidates], ensure_ascii=False),
                product_class,
                int(bool(candidates)),
            )
        )

    connection = sqlite3.connect(DB_PATH)
    create_schema(connection)
    with connection:
        connection.execute(
            "DELETE FROM stg_dimension_scan_pass1_unique_image WHERE run_name=?", (RUN_NAME,)
        )
        connection.execute(
            "DELETE FROM stg_dimension_scan_pass1_product_image WHERE run_name=?", (RUN_NAME,)
        )
        connection.execute(
            "DELETE FROM stg_dimension_scan_pass1_product WHERE run_name=?", (RUN_NAME,)
        )
        connection.executemany(
            "INSERT INTO stg_dimension_scan_pass1_unique_image VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            unique_inserts,
        )
        connection.executemany(
            "INSERT INTO stg_dimension_scan_pass1_product_image VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            product_image_inserts,
        )
        connection.executemany(
            "INSERT INTO stg_dimension_scan_pass1_product VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
            product_inserts,
        )
    connection.close()

    summary = {
        "run_name": RUN_NAME,
        "created_at": created_at,
        "scope": "has_human_review_information=0인 상품의 전체 상세 정지 이미지",
        "target_products": len(product_inserts),
        "product_image_rows": len(product_image_inserts),
        "unique_images": len(unique_inserts),
        "ocr_success_unique_images": (
            unique_status["NEW_SUCCESS"] + unique_status["REUSED_SUCCESS"]
        ),
        "ocr_success_unique_pct": round(
            100
            * (unique_status["NEW_SUCCESS"] + unique_status["REUSED_SUCCESS"])
            / max(1, len(unique_inserts)),
            1,
        ),
        "unique_ocr_status": dict(unique_status),
        "unique_image_classification": dict(unique_classes),
        "candidate_unique_images": sum(row[12] for row in unique_inserts),
        "candidate_product_image_rows": sum(row[10] for row in product_image_inserts),
        "image_row_reduction_pct": round(
            100
            * (
                1
                - sum(row[10] for row in product_image_inserts)
                / max(1, len(product_image_inserts))
            ),
            1,
        ),
        "product_classification": dict(product_classes),
        "second_pass_target_products": product_classes["CANDIDATE_IMAGES_FOUND"],
        "no_signal_products": product_classes["NO_SIZE_SIGNAL_ALL_IMAGES"],
        "scan_incomplete_products": product_classes["SCAN_INCOMPLETE"],
        "no_detail_image_products": product_classes["NO_DETAIL_IMAGES"],
        "source_dimension_values_written": False,
        "excel_written": False,
        "staging_tables": [
            "stg_dimension_scan_pass1_unique_image",
            "stg_dimension_scan_pass1_product_image",
            "stg_dimension_scan_pass1_product",
            "vw_dimension_scan_pass1_candidate_images",
        ],
    }
    SUMMARY_PATH.write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2), flush=True)
    return summary


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("phase", choices=("prepare", "ocr", "classify", "all"))
    parser.add_argument("--product-id", default="")
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--workers", type=int, default=8)
    parser.add_argument("--ocr-workers", type=int, default=4)
    parser.add_argument("--shards", type=int, default=DEFAULT_SHARDS)
    parser.add_argument("--shard-start", type=int, default=0)
    parser.add_argument("--shard-end", type=int, default=0)
    parser.add_argument("--max-width", type=int, default=DEFAULT_MAX_WIDTH)
    args = parser.parse_args()
    if args.phase in {"prepare", "all"}:
        prepare(
            product_id=args.product_id,
            limit=args.limit,
            workers=args.workers,
            shards=args.shards,
            max_width=args.max_width,
        )
    if args.phase in {"ocr", "all"}:
        run_ocr(args.shards, args.ocr_workers, args.shard_start, args.shard_end)
    if args.phase in {"classify", "all"}:
        classify(args.shards)


if __name__ == "__main__":
    main()
