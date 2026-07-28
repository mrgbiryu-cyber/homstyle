from __future__ import annotations

import argparse
import concurrent.futures
import io
import json
import os
import shutil
import sqlite3
import subprocess
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

from PIL import Image, ImageFile

import bulk_homestyle_ocr as ocr
from bulk_homestyle_collect import DB_PATH, request_bytes
from dimension_pass2_parser import parse_layout
from run_dimension_scan_pass1 import existing_path, load_json, row_url


RUN_NAME = "dimension_scan_pass2_layout_v1"
PASS1_RUN_NAME = "dimension_scan_pass1_all_detail_images_v1"
RUN_DIR = ocr.OCR_ROOT / RUN_NAME
MANIFEST_PATH = RUN_DIR / "manifest.json"
SUMMARY_PATH = ocr.RUN_DIR / f"{RUN_NAME}.json"
LAYOUT_OCR = Path(__file__).resolve().parent / "run_windows_ocr_layout.ps1"
DEFAULT_SHARDS = 8
DEFAULT_MAX_WIDTH = 2200
MAX_OBSERVATIONS_PER_IMAGE = 100

ImageFile.LOAD_TRUNCATED_IMAGES = True
Image.MAX_IMAGE_PIXELS = None


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def discover_image_cache(target_urls: set[str]) -> dict[str, Path]:
    cache: dict[str, Path] = {}
    width_cache: dict[Path, int] = {}
    if not ocr.OCR_ROOT.exists():
        return cache
    for manifest_path in ocr.OCR_ROOT.rglob("manifest.json"):
        if RUN_DIR in manifest_path.parents:
            continue
        manifest = load_json(manifest_path)
        if not isinstance(manifest, dict):
            continue
        rows = (
            manifest.get("unique_images")
            or manifest.get("products")
            or manifest.get("images")
            or []
        )
        if not isinstance(rows, list):
            continue
        for row in rows:
            if not isinstance(row, dict):
                continue
            url = row_url(row)
            if not url:
                url = ocr.normalize_url(str(row.get("image_url") or ""))
            if url not in target_urls:
                continue
            source = existing_path(str(row.get("file") or row.get("local_file") or ""), manifest_path)
            if url.startswith(("http://", "https://")) and source:
                old = cache.get(url)
                if old is None:
                    cache[url] = source
                else:
                    try:
                        if old not in width_cache:
                            with Image.open(old) as old_image:
                                width_cache[old] = old_image.width
                        if source not in width_cache:
                            with Image.open(source) as new_image:
                                width_cache[source] = new_image.width
                        if width_cache[source] > width_cache[old]:
                            cache[url] = source
                    except OSError:
                        pass
    return cache


def save_source(source: Image.Image, destination: Path, max_width: int) -> tuple[int, int]:
    source.seek(0)
    image = source.convert("RGB")
    if image.width > max_width:
        height = max(1, round(image.height * max_width / image.width))
        image = image.resize((max_width, height), Image.Resampling.LANCZOS)
    destination.parent.mkdir(parents=True, exist_ok=True)
    image.save(destination, "JPEG", quality=92, optimize=True)
    return image.size


def prepare_one(
    item: dict[str, Any], cache: dict[str, Path], shards: int, max_width: int
) -> dict[str, Any]:
    digest = item["url_hash"]
    shard = int(digest[:8], 16) % shards
    destination = RUN_DIR / f"shard_{shard:02d}" / f"{digest}.jpg"
    if destination.exists() and destination.stat().st_size:
        try:
            with Image.open(destination) as image:
                width, height = image.size
            return {
                **item,
                "shard": shard,
                "file": str(destination),
                "file_name": destination.name,
                "prepare_status": "REUSED_RUN_FILE",
                "width": width,
                "height": height,
                "error": "",
            }
        except OSError:
            pass

    fallback_paths: list[Path] = []
    local = Path(str(item.get("pass1_local_file") or ""))
    if local.exists() and local.is_file():
        fallback_paths.append(local)
    cached = cache.get(item["image_url"])
    if cached and cached not in fallback_paths:
        fallback_paths.append(cached)

    best_source: Path | None = None
    best_width = 0
    for source in fallback_paths:
        try:
            with Image.open(source) as image:
                width = image.width
            if width > best_width:
                best_source, best_width = source, width
        except OSError:
            continue

    # Widths below 1,400 were not downscaled by pass 1 and therefore already
    # preserve the available source resolution. Widths above 1,400 are also
    # suitable. Exactly 1,400 can be a pass-1 downscale, so try the URL first.
    if best_source and best_width != 1400:
        try:
            with Image.open(best_source) as image:
                width, height = save_source(image, destination, max_width)
            return {
                **item,
                "shard": shard,
                "file": str(destination),
                "file_name": destination.name,
                "prepare_status": "REUSED_BEST_CACHE",
                "width": width,
                "height": height,
                "error": "",
            }
        except OSError:
            pass

    download_error = ""
    try:
        status, body = request_bytes(item["image_url"], timeout=35, attempts=2)
        if status != 200 or not body:
            raise RuntimeError(f"HTTP {status}; bytes={len(body)}")
        with Image.open(io.BytesIO(body)) as image:
            width, height = save_source(image, destination, max_width)
        return {
            **item,
            "shard": shard,
            "file": str(destination),
            "file_name": destination.name,
            "prepare_status": "DOWNLOADED_HIGH_RES",
            "width": width,
            "height": height,
            "error": "",
        }
    except Exception as exc:
        download_error = f"{type(exc).__name__}: {exc}"

    if best_source:
        try:
            with Image.open(best_source) as image:
                width, height = save_source(image, destination, max_width)
            return {
                **item,
                "shard": shard,
                "file": str(destination),
                "file_name": destination.name,
                "prepare_status": "FALLBACK_LOWER_RES",
                "width": width,
                "height": height,
                "error": download_error,
            }
        except OSError:
            pass
    return {
        **item,
        "shard": shard,
        "file": "",
        "file_name": "",
        "prepare_status": "ERROR",
        "width": None,
        "height": None,
        "error": download_error or "No usable source image",
    }


def prepare(
    workers: int = 8,
    shards: int = DEFAULT_SHARDS,
    max_width: int = DEFAULT_MAX_WIDTH,
) -> dict[str, Any]:
    connection = sqlite3.connect(DB_PATH)
    rows = connection.execute(
        """
        SELECT url_hash,image_url,reference_count,local_file,classification,candidate_score
        FROM stg_dimension_scan_pass1_unique_image
        WHERE run_name=? AND is_candidate=1
        ORDER BY url_hash
        """,
        (PASS1_RUN_NAME,),
    ).fetchall()
    connection.close()
    items = [
        {
            "url_hash": digest,
            "image_url": url,
            "reference_count": reference_count,
            "pass1_local_file": local_file or "",
            "pass1_classification": classification,
            "pass1_candidate_score": score,
        }
        for digest, url, reference_count, local_file, classification, score in rows
    ]
    RUN_DIR.mkdir(parents=True, exist_ok=True)
    for shard in range(shards):
        (RUN_DIR / f"shard_{shard:02d}").mkdir(parents=True, exist_ok=True)
    # Pass-1 local files cover most candidates. Reading every historical OCR
    # manifest on Windows is slower than re-fetching the remaining URLs, so
    # pass 2 deliberately uses only the direct pass-1 file plus the source URL.
    cache: dict[str, Path] = {}
    results: list[dict[str, Any]] = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, workers)) as pool:
        futures = [pool.submit(prepare_one, item, cache, shards, max_width) for item in items]
        for index, future in enumerate(concurrent.futures.as_completed(futures), 1):
            results.append(future.result())
            if index % 100 == 0 or index == len(futures):
                print(f"prepared={index}/{len(futures)}", flush=True)
    results.sort(key=lambda item: item["url_hash"])
    status = Counter(item["prepare_status"] for item in results)
    metadata = {
        "run_name": RUN_NAME,
        "created_at": now_text(),
        "pass1_run_name": PASS1_RUN_NAME,
        "unique_candidate_images": len(items),
        "prepare_status": dict(status),
        "prepare_success": sum(item["prepare_status"] != "ERROR" for item in results),
        "prepare_error": status["ERROR"],
        "shards": shards,
        "max_width": max_width,
        "source_dimension_values_written": False,
        "excel_written": False,
    }
    MANIFEST_PATH.write_text(
        json.dumps({"metadata": metadata, "images": results}, ensure_ascii=False),
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
                str(LAYOUT_OCR),
                "-InputDirectory",
                str(RUN_DIR / f"shard_{shard:02d}"),
                "-OutputJson",
                str(RUN_DIR / f"layout_ocr_{shard:02d}.json"),
                "-LanguageTag",
                "ko",
                "-Resume",
            ]
            processes.append((shard, subprocess.Popen(command)))
        for shard, process in processes:
            code = process.wait()
            print(f"layout_ocr_shard={shard} exit={code}", flush=True)
            if code:
                failed.append(shard)
    if failed:
        raise RuntimeError(f"Layout OCR shards failed: {failed}")


def load_layout_rows(shards: int) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for shard in range(shards):
        path = RUN_DIR / f"layout_ocr_{shard:02d}.json"
        rows = load_json(path)
        if isinstance(rows, dict):
            rows = [rows]
        if not isinstance(rows, list):
            continue
        for row in rows:
            if isinstance(row, dict):
                result[Path(str(row.get("file") or "")).stem] = row
    return result


def create_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        DROP VIEW IF EXISTS vw_dimension_scan_pass2_complete_wdh;
        DROP VIEW IF EXISTS vw_dimension_scan_pass2_review_queue;
        DROP VIEW IF EXISTS vw_dimension_scan_pass2_status_summary;
        DROP VIEW IF EXISTS vw_dimension_scan_pass2_best_evidence;
        DROP VIEW IF EXISTS vw_dimension_scan_pass2_prepare_errors;
        DROP VIEW IF EXISTS vw_dimension_pattern_with_pass2_current;

        CREATE TABLE IF NOT EXISTS stg_dimension_scan_pass2_unique_image (
            run_name TEXT NOT NULL,
            analyzed_at TEXT NOT NULL,
            url_hash TEXT NOT NULL,
            image_url TEXT NOT NULL,
            reference_count INTEGER NOT NULL,
            pass1_classification TEXT,
            pass1_candidate_score INTEGER,
            local_file TEXT,
            prepare_status TEXT,
            ocr_status TEXT,
            original_width INTEGER,
            original_height INTEGER,
            scaled_width INTEGER,
            scaled_height INTEGER,
            tile_count INTEGER,
            word_count INTEGER,
            character_count INTEGER,
            observation_count INTEGER NOT NULL,
            observation_overflow_count INTEGER NOT NULL,
            ocr_text TEXT,
            error TEXT,
            PRIMARY KEY(run_name,url_hash)
        );

        CREATE TABLE IF NOT EXISTS stg_dimension_scan_pass2_observation (
            run_name TEXT NOT NULL,
            analyzed_at TEXT NOT NULL,
            url_hash TEXT NOT NULL,
            observation_no INTEGER NOT NULL,
            image_url TEXT NOT NULL,
            candidate_type TEXT NOT NULL,
            raw_notation TEXT,
            axis_signature TEXT,
            unit_status TEXT,
            unit_text TEXT,
            w_raw REAL,
            d_raw REAL,
            h_raw REAL,
            l_raw REAL,
            r_raw REAL,
            value_1_raw REAL,
            value_2_raw REAL,
            value_3_raw REAL,
            w_mm REAL,
            d_mm REAL,
            h_mm REAL,
            l_mm REAL,
            mapping_status TEXT NOT NULL,
            confidence TEXT NOT NULL,
            candidate_score INTEGER NOT NULL,
            evidence_text TEXT,
            bbox_json TEXT,
            PRIMARY KEY(run_name,url_hash,observation_no)
        );

        CREATE TABLE IF NOT EXISTS stg_dimension_scan_pass2_product_image (
            run_name TEXT NOT NULL,
            analyzed_at TEXT NOT NULL,
            product_id TEXT NOT NULL,
            product_name TEXT,
            small_category TEXT,
            image_order INTEGER NOT NULL,
            url_hash TEXT NOT NULL,
            image_url TEXT NOT NULL,
            ocr_status TEXT,
            observation_count INTEGER NOT NULL,
            best_observation_no INTEGER,
            best_candidate_type TEXT,
            best_raw_notation TEXT,
            best_axis_signature TEXT,
            best_unit_status TEXT,
            best_mapping_status TEXT,
            best_confidence TEXT,
            best_candidate_score INTEGER,
            best_w_mm REAL,
            best_d_mm REAL,
            best_h_mm REAL,
            best_evidence_text TEXT,
            PRIMARY KEY(run_name,product_id,image_order)
        );

        CREATE TABLE IF NOT EXISTS stg_dimension_scan_pass2_product (
            run_name TEXT NOT NULL,
            analyzed_at TEXT NOT NULL,
            product_id TEXT NOT NULL,
            product_name TEXT,
            small_category TEXT,
            target_image_count INTEGER NOT NULL,
            ocr_success_image_count INTEGER NOT NULL,
            ocr_error_image_count INTEGER NOT NULL,
            observation_count INTEGER NOT NULL,
            complete_wdh_candidate_count INTEGER NOT NULL,
            best_image_order INTEGER,
            best_image_url TEXT,
            best_observation_no INTEGER,
            best_candidate_type TEXT,
            best_raw_notation TEXT,
            best_axis_signature TEXT,
            best_unit_status TEXT,
            best_mapping_status TEXT,
            best_confidence TEXT,
            best_candidate_score INTEGER,
            resolved_w_mm REAL,
            resolved_d_mm REAL,
            resolved_h_mm REAL,
            pass2_status TEXT NOT NULL,
            needs_human_review INTEGER NOT NULL,
            prior_pattern_code TEXT,
            next_action TEXT,
            PRIMARY KEY(run_name,product_id)
        );

        CREATE INDEX IF NOT EXISTS idx_dimension_pass2_obs_type
            ON stg_dimension_scan_pass2_observation(run_name,candidate_type,mapping_status);
        CREATE INDEX IF NOT EXISTS idx_dimension_pass2_product_status
            ON stg_dimension_scan_pass2_product(run_name,pass2_status,small_category);

        CREATE VIEW vw_dimension_scan_pass2_complete_wdh AS
        SELECT * FROM stg_dimension_scan_pass2_product
        WHERE pass2_status='COMPLETE_EXPLICIT_WDH_CANDIDATE'
          AND resolved_w_mm IS NOT NULL AND resolved_d_mm IS NOT NULL AND resolved_h_mm IS NOT NULL;

        CREATE VIEW vw_dimension_scan_pass2_review_queue AS
        SELECT * FROM stg_dimension_scan_pass2_product
        WHERE needs_human_review=1;

        CREATE VIEW vw_dimension_scan_pass2_status_summary AS
        SELECT run_name,pass2_status,COUNT(*) AS product_count
        FROM stg_dimension_scan_pass2_product
        GROUP BY run_name,pass2_status;

        CREATE VIEW vw_dimension_scan_pass2_best_evidence AS
        SELECT p.*,
               o.evidence_text AS best_evidence_text,
               o.bbox_json AS best_bbox_json,
               o.unit_text AS best_unit_text,
               o.w_raw AS best_w_raw,o.d_raw AS best_d_raw,o.h_raw AS best_h_raw,o.l_raw AS best_l_raw
        FROM stg_dimension_scan_pass2_product p
        LEFT JOIN stg_dimension_scan_pass2_product_image pi
          ON pi.run_name=p.run_name
         AND pi.product_id=p.product_id
         AND pi.image_order=p.best_image_order
        LEFT JOIN stg_dimension_scan_pass2_observation o
          ON o.run_name=pi.run_name
         AND o.url_hash=pi.url_hash
         AND o.observation_no=p.best_observation_no;

        CREATE VIEW vw_dimension_scan_pass2_prepare_errors AS
        SELECT * FROM stg_dimension_scan_pass2_unique_image
        WHERE ocr_status IN ('PREPARE_ERROR','OCR_ERROR','OCR_NOT_RUN');

        CREATE VIEW vw_dimension_pattern_with_pass2_current AS
        SELECT q.*,p.run_name AS pass2_run_name,p.pass2_status,
               p.best_candidate_type,p.best_raw_notation,p.best_axis_signature,
               p.best_unit_status,p.best_mapping_status,p.best_confidence,
               p.resolved_w_mm,p.resolved_d_mm,p.resolved_h_mm,
               p.needs_human_review AS pass2_needs_human_review,
               p.next_action AS pass2_next_action
        FROM vw_dimension_pattern_work_queue_current q
        LEFT JOIN stg_dimension_scan_pass2_product p
          ON p.product_id=q.product_id AND p.run_name='dimension_scan_pass2_layout_v1';
        """
    )


def product_status(best: dict[str, Any] | None, ocr_error_count: int) -> tuple[str, int, str]:
    if best is None:
        if ocr_error_count:
            return "OCR_INCOMPLETE_NO_CANDIDATE", 1, "오류 이미지 재수집 또는 후보 이미지 재OCR"
        return "NO_DIMENSION_OBSERVATION", 1, "OCR 원문 사람 검토 또는 다른 소스 조사"
    if best.get("w_mm") is not None and best.get("d_mm") is not None and best.get("h_mm") is not None:
        dimensions = [float(best["w_mm"]), float(best["d_mm"]), float(best["h_mm"])]
        if min(dimensions) < 20 or max(dimensions) > 5000:
            return "COMPLETE_WDH_RANGE_REVIEW", 1, "비정상 범위 또는 OCR 숫자 결합 오탐 여부 검증"
        return "COMPLETE_EXPLICIT_WDH_CANDIDATE", 1, "대표 제품 규격·옵션 충돌 검증 후 적용"
    candidate_type = best.get("candidate_type")
    if candidate_type == "EXPLICIT_WDH":
        return "COMPLETE_WDH_UNIT_MISSING", 1, "단위를 검증한 뒤 W/D/H 정규화"
    if candidate_type == "EXPLICIT_LWH":
        return "COMPLETE_NONSTANDARD_AXES_REVIEW", 1, "L과 W의 의미를 카테고리·도면 기준으로 해석"
    if candidate_type in {"UNLABELED_TRIPLE", "SPATIAL_THREE_DIMENSION_TOKENS"}:
        return "TRIPLET_MAPPING_REVIEW", 1, "3개 값의 W/D/H 순서 또는 공간 관계 판단"
    if candidate_type in {"PARTIAL_LABELED_AXES", "SINGLE_LABELED_AXIS"}:
        return "PARTIAL_AXES_REVIEW", 1, "확보 축을 유지하고 누락 축 보강"
    return "NUMERIC_CLUSTER_REVIEW", 1, "옵션·구성품·제품 규격 숫자 군집 분리"


def analyze(shards: int = DEFAULT_SHARDS) -> dict[str, Any]:
    manifest = load_json(MANIFEST_PATH)
    if not isinstance(manifest, dict):
        raise RuntimeError(f"Manifest missing or invalid: {MANIFEST_PATH}")
    analyzed_at = now_text()
    layout_rows = load_layout_rows(shards)
    image_results: dict[str, dict[str, Any]] = {}
    image_inserts: list[tuple[Any, ...]] = []
    observation_inserts: list[tuple[Any, ...]] = []
    observation_by_hash: dict[str, list[dict[str, Any]]] = {}
    ocr_status_counts = Counter()
    observation_type_counts = Counter()

    for item in manifest.get("images") or []:
        digest = item["url_hash"]
        layout = layout_rows.get(digest) or {}
        if layout.get("status") == "SUCCESS":
            ocr_status = "SUCCESS"
            observations_all = parse_layout(layout)
            observations = observations_all[:MAX_OBSERVATIONS_PER_IMAGE]
            error = str(item.get("error") or "")
        elif item.get("prepare_status") == "ERROR":
            ocr_status = "PREPARE_ERROR"
            observations_all = []
            observations = []
            error = str(item.get("error") or "")
        elif layout:
            ocr_status = "OCR_ERROR"
            observations_all = []
            observations = []
            error = str(layout.get("error") or "")
        else:
            ocr_status = "OCR_NOT_RUN"
            observations_all = []
            observations = []
            error = "Layout OCR output missing"
        ocr_status_counts[ocr_status] += 1
        for number, observation_row in enumerate(observations, 1):
            observation_row["observation_no"] = number
            observation_type_counts[observation_row["candidate_type"]] += 1
            observation_inserts.append(
                (
                    RUN_NAME,
                    analyzed_at,
                    digest,
                    number,
                    item["image_url"],
                    observation_row["candidate_type"],
                    observation_row["raw_notation"],
                    observation_row["axis_signature"],
                    observation_row["unit_status"],
                    observation_row["unit_text"],
                    observation_row["w_raw"],
                    observation_row["d_raw"],
                    observation_row["h_raw"],
                    observation_row["l_raw"],
                    observation_row["r_raw"],
                    observation_row["value_1_raw"],
                    observation_row["value_2_raw"],
                    observation_row["value_3_raw"],
                    observation_row["w_mm"],
                    observation_row["d_mm"],
                    observation_row["h_mm"],
                    observation_row["l_mm"],
                    observation_row["mapping_status"],
                    observation_row["confidence"],
                    observation_row["candidate_score"],
                    observation_row["evidence_text"],
                    observation_row["bbox_json"],
                )
            )
        observation_by_hash[digest] = observations
        result = {
            "ocr_status": ocr_status,
            "observations": observations,
            "error": error,
        }
        image_results[digest] = result
        text = str(layout.get("text") or "")
        image_inserts.append(
            (
                RUN_NAME,
                analyzed_at,
                digest,
                item["image_url"],
                int(item.get("reference_count") or 0),
                item.get("pass1_classification") or "",
                item.get("pass1_candidate_score"),
                item.get("file") or "",
                item.get("prepare_status") or "",
                ocr_status,
                layout.get("original_width") or item.get("width"),
                layout.get("original_height") or item.get("height"),
                layout.get("scaled_width") or item.get("width"),
                layout.get("scaled_height") or item.get("height"),
                int(layout.get("tile_count") or 0),
                len(layout.get("words") or []),
                len(text),
                len(observations),
                max(0, len(observations_all) - len(observations)),
                text,
                error,
            )
        )

    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    create_schema(connection)
    work_queue = {
        row["product_id"]: row
        for row in connection.execute("SELECT * FROM vw_dimension_pattern_work_queue_current")
    }
    product_refs = connection.execute(
        """
        SELECT product_id,product_name,small_category,image_order,url_hash,image_url
        FROM stg_dimension_scan_pass1_product_image
        WHERE run_name=? AND is_candidate=1
        ORDER BY product_id,image_order
        """,
        (PASS1_RUN_NAME,),
    ).fetchall()

    refs_by_product: dict[str, list[sqlite3.Row]] = defaultdict(list)
    product_image_inserts: list[tuple[Any, ...]] = []
    for ref in product_refs:
        refs_by_product[ref["product_id"]].append(ref)
        observations = observation_by_hash.get(ref["url_hash"]) or []
        best = observations[0] if observations else None
        image_result = image_results.get(ref["url_hash"]) or {"ocr_status": "NOT_FOUND"}
        product_image_inserts.append(
            (
                RUN_NAME,
                analyzed_at,
                ref["product_id"],
                ref["product_name"],
                ref["small_category"],
                int(ref["image_order"]),
                ref["url_hash"],
                ref["image_url"],
                image_result["ocr_status"],
                len(observations),
                best.get("observation_no") if best else None,
                best.get("candidate_type") if best else "",
                best.get("raw_notation") if best else "",
                best.get("axis_signature") if best else "",
                best.get("unit_status") if best else "",
                best.get("mapping_status") if best else "",
                best.get("confidence") if best else "",
                best.get("candidate_score") if best else None,
                best.get("w_mm") if best else None,
                best.get("d_mm") if best else None,
                best.get("h_mm") if best else None,
                best.get("evidence_text") if best else "",
            )
        )

    product_inserts: list[tuple[Any, ...]] = []
    product_status_counts = Counter()
    for pid, refs in sorted(refs_by_product.items()):
        candidates: list[tuple[int, dict[str, Any], sqlite3.Row]] = []
        success_count = 0
        error_count = 0
        observation_count = 0
        complete_count = 0
        for ref in refs:
            image_result = image_results.get(ref["url_hash"]) or {"ocr_status": "NOT_FOUND"}
            if image_result["ocr_status"] == "SUCCESS":
                success_count += 1
            else:
                error_count += 1
            observations = observation_by_hash.get(ref["url_hash"]) or []
            observation_count += len(observations)
            for obs in observations:
                candidates.append((int(obs["candidate_score"]), obs, ref))
                if obs.get("w_mm") is not None and obs.get("d_mm") is not None and obs.get("h_mm") is not None:
                    complete_count += 1
        candidates.sort(key=lambda value: (-value[0], int(value[2]["image_order"]), int(value[1]["observation_no"])))
        best_obs = candidates[0][1] if candidates else None
        best_ref = candidates[0][2] if candidates else None
        status, needs_review, next_action = product_status(best_obs, error_count)
        product_status_counts[status] += 1
        first = refs[0]
        prior = work_queue.get(pid)
        product_inserts.append(
            (
                RUN_NAME,
                analyzed_at,
                pid,
                first["product_name"],
                first["small_category"],
                len(refs),
                success_count,
                error_count,
                observation_count,
                complete_count,
                int(best_ref["image_order"]) if best_ref else None,
                best_ref["image_url"] if best_ref else "",
                best_obs.get("observation_no") if best_obs else None,
                best_obs.get("candidate_type") if best_obs else "",
                best_obs.get("raw_notation") if best_obs else "",
                best_obs.get("axis_signature") if best_obs else "",
                best_obs.get("unit_status") if best_obs else "",
                best_obs.get("mapping_status") if best_obs else "",
                best_obs.get("confidence") if best_obs else "",
                best_obs.get("candidate_score") if best_obs else None,
                best_obs.get("w_mm") if best_obs else None,
                best_obs.get("d_mm") if best_obs else None,
                best_obs.get("h_mm") if best_obs else None,
                status,
                needs_review,
                str(prior["pattern_code"] if prior else ""),
                next_action,
            )
        )

    with connection:
        for table in (
            "stg_dimension_scan_pass2_unique_image",
            "stg_dimension_scan_pass2_observation",
            "stg_dimension_scan_pass2_product_image",
            "stg_dimension_scan_pass2_product",
        ):
            connection.execute(f"DELETE FROM {table} WHERE run_name=?", (RUN_NAME,))
        connection.executemany(
            "INSERT INTO stg_dimension_scan_pass2_unique_image VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            image_inserts,
        )
        connection.executemany(
            "INSERT INTO stg_dimension_scan_pass2_observation VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            observation_inserts,
        )
        connection.executemany(
            "INSERT INTO stg_dimension_scan_pass2_product_image VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            product_image_inserts,
        )
        connection.executemany(
            "INSERT INTO stg_dimension_scan_pass2_product VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            product_inserts,
        )
    connection.close()

    summary = {
        "run_name": RUN_NAME,
        "analyzed_at": analyzed_at,
        "pass1_run_name": PASS1_RUN_NAME,
        "unique_target_images": len(image_inserts),
        "product_image_rows": len(product_image_inserts),
        "target_products": len(product_inserts),
        "ocr_status": dict(ocr_status_counts),
        "observation_rows": len(observation_inserts),
        "observation_type": dict(observation_type_counts),
        "product_status": dict(product_status_counts),
        "products_with_any_observation": len(product_inserts)
        - product_status_counts["NO_DIMENSION_OBSERVATION"]
        - product_status_counts["OCR_INCOMPLETE_NO_CANDIDATE"],
        "products_without_observation": product_status_counts["NO_DIMENSION_OBSERVATION"]
        + product_status_counts["OCR_INCOMPLETE_NO_CANDIDATE"],
        "complete_explicit_wdh_products": product_status_counts["COMPLETE_EXPLICIT_WDH_CANDIDATE"],
        "complete_wdh_range_review_products": product_status_counts["COMPLETE_WDH_RANGE_REVIEW"],
        "source_dimension_values_written": False,
        "excel_written": False,
        "staging_tables": [
            "stg_dimension_scan_pass2_unique_image",
            "stg_dimension_scan_pass2_observation",
            "stg_dimension_scan_pass2_product_image",
            "stg_dimension_scan_pass2_product",
        ],
        "views": [
            "vw_dimension_scan_pass2_complete_wdh",
            "vw_dimension_scan_pass2_review_queue",
            "vw_dimension_scan_pass2_status_summary",
            "vw_dimension_scan_pass2_best_evidence",
            "vw_dimension_scan_pass2_prepare_errors",
            "vw_dimension_pattern_with_pass2_current",
        ],
    }
    SUMMARY_PATH.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2), flush=True)
    return summary


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("phase", choices=("prepare", "ocr", "analyze", "all"))
    parser.add_argument("--workers", type=int, default=8)
    parser.add_argument("--ocr-workers", type=int, default=4)
    parser.add_argument("--shards", type=int, default=DEFAULT_SHARDS)
    parser.add_argument("--shard-start", type=int, default=0)
    parser.add_argument("--shard-end", type=int, default=0)
    parser.add_argument("--max-width", type=int, default=DEFAULT_MAX_WIDTH)
    args = parser.parse_args()
    if args.phase in {"prepare", "all"}:
        prepare(args.workers, args.shards, args.max_width)
    if args.phase in {"ocr", "all"}:
        run_ocr(args.shards, args.ocr_workers, args.shard_start, args.shard_end)
    if args.phase in {"analyze", "all"}:
        analyze(args.shards)


if __name__ == "__main__":
    main()
