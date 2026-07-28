from __future__ import annotations

import argparse
import concurrent.futures
import hashlib
import json
import sqlite3
import subprocess
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

from PIL import Image, ImageEnhance, ImageFile, ImageOps

from bulk_homestyle_collect import DB_PATH
from dimension_context_normalizer import extract_candidates
from run_dimension_targeted_reocr import trigger_positions


ROOT = Path(__file__).resolve().parent
RUN_NAME = "dimension_targeted_reocr_remaining_v1"
RUN_DIR = ROOT / "homestyle_bulk_run" / "ocr" / RUN_NAME
PASS2_DIR = ROOT / "homestyle_bulk_run" / "ocr" / "dimension_scan_pass2_layout_v1"
PASS2_MANIFEST = PASS2_DIR / "manifest.json"
MANIFEST_PATH = RUN_DIR / "manifest.json"
SUMMARY_PATH = RUN_DIR / "summary.json"
OCR_SCRIPT = ROOT / "run_windows_ocr_layout.ps1"
DEFAULT_SHARDS = 8

ImageFile.LOAD_TRUNCATED_IMAGES = True
Image.MAX_IMAGE_PIXELS = None


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def save_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")


def target_product_rows(connection: sqlite3.Connection) -> dict[str, dict[str, Any]]:
    master = {
        row["product_id"]: dict(row)
        for row in connection.execute(
            """
            SELECT product_id, product_name, mid_category, small_category
            FROM vw_dimension_classification_master_current
            """
        )
    }
    statuses = {
        row["product_id"]: row["resolution_status"]
        for row in connection.execute(
            """
            SELECT product_id, resolution_status
            FROM fact_dimension_resolution_ledger
            WHERE resolution_status IN ('OCR_REQUIRED', 'NO_CANDIDATE')
            """
        )
    }
    result: dict[str, dict[str, Any]] = {}
    for product_id, status in statuses.items():
        product = master.get(product_id)
        if not product:
            continue
        result[product_id] = {**product, "ledger_status": status, "images": []}
    for row in connection.execute(
        """
        SELECT product_id, image_order, url_hash, image_url,
               classification, candidate_score
        FROM stg_dimension_scan_pass1_product_image
        WHERE run_name = 'dimension_scan_pass1_all_detail_images_v1'
          AND is_candidate = 1
        ORDER BY product_id, image_order
        """
    ):
        product = result.get(row["product_id"])
        if product is None:
            continue
        product["images"].append(dict(row))
    return result


def load_pass2_layout(
    target_hashes: set[str], shards: int
) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for shard in range(shards):
        path = PASS2_DIR / f"layout_ocr_{shard:02d}.json"
        if not path.exists():
            continue
        rows = load_json(path)
        if isinstance(rows, dict):
            rows = [rows]
        for row in rows:
            digest = Path(str(row.get("file") or "")).stem
            if digest in target_hashes:
                result[digest] = row
    return result


def crop_one(
    task: dict[str, Any],
    shards: int,
) -> dict[str, Any]:
    digest = task["url_hash"]
    crop_no = task["crop_no"]
    shard = int(hashlib.sha1(f"{digest}:{crop_no}".encode()).hexdigest()[:8], 16) % shards
    crop_name = f"{digest}_crop{crop_no:02d}.jpg"
    destination = RUN_DIR / f"shard_{shard:02d}" / crop_name
    try:
        source_path = Path(task["source_file"])
        with Image.open(source_path) as raw:
            source = raw.convert("RGB")
            y = float(task["trigger_y"])
            top = max(0, int(y) - 300)
            bottom = min(source.height, int(y) + 1900)
            if bottom - top < 300:
                raise RuntimeError("crop height below 300px")
            crop = source.crop((0, top, source.width, bottom))
            scale = min(3.0, max(1.0, 2200 / max(1, crop.width)))
            if scale > 1.01:
                crop = crop.resize(
                    (round(crop.width * scale), round(crop.height * scale)),
                    Image.Resampling.LANCZOS,
                )
            enhanced = ImageOps.autocontrast(ImageOps.grayscale(crop))
            enhanced = ImageEnhance.Contrast(enhanced).enhance(1.35)
            enhanced = ImageEnhance.Sharpness(enhanced).enhance(1.8)
            enhanced.save(destination, "JPEG", quality=94)
        return {
            **task,
            "shard": shard,
            "crop_file": crop_name,
            "local_file": str(destination),
            "crop_top": top,
            "crop_bottom": bottom,
            "crop_scale": scale,
            "prepare_status": "REUSED_CROP" if task.get("reused") else "CREATED",
            "error": "",
        }
    except Exception as exc:
        return {
            **task,
            "shard": shard,
            "crop_file": crop_name,
            "local_file": "",
            "crop_top": None,
            "crop_bottom": None,
            "crop_scale": None,
            "prepare_status": "ERROR",
            "error": f"{type(exc).__name__}: {exc}",
        }


def prepare(workers: int = 8, shards: int = DEFAULT_SHARDS) -> dict[str, Any]:
    RUN_DIR.mkdir(parents=True, exist_ok=True)
    for shard in range(shards):
        (RUN_DIR / f"shard_{shard:02d}").mkdir(parents=True, exist_ok=True)

    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    products = target_product_rows(connection)
    connection.close()

    target_hashes = {
        image["url_hash"]
        for product in products.values()
        for image in product["images"]
    }
    pass2_manifest = load_json(PASS2_MANIFEST)
    pass2_images = {
        row["url_hash"]: row
        for row in pass2_manifest.get("images", [])
        if row.get("url_hash") in target_hashes
    }
    layouts = load_pass2_layout(target_hashes, shards)

    tasks: list[dict[str, Any]] = []
    image_rows: list[dict[str, Any]] = []
    for digest in sorted(target_hashes):
        layout = layouts.get(digest) or {}
        image = pass2_images.get(digest) or {}
        positions = (
            trigger_positions(layout) if layout.get("status") == "SUCCESS" else []
        )
        image_rows.append(
            {
                "url_hash": digest,
                "image_url": image.get("image_url") or "",
                "source_file": image.get("file") or "",
                "layout_status": layout.get("status") or "NOT_FOUND",
                "heading_positions": positions,
            }
        )
        if not image.get("file"):
            continue
        for crop_no, position in enumerate(positions, 1):
            tasks.append(
                {
                    "url_hash": digest,
                    "image_url": image.get("image_url") or "",
                    "source_file": image["file"],
                    "crop_no": crop_no,
                    "trigger_y": position,
                }
            )

    with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, workers)) as pool:
        crops = list(pool.map(lambda task: crop_one(task, shards), tasks))

    product_rows = []
    for product in products.values():
        product_rows.append(
            {
                key: value
                for key, value in product.items()
                if key != "images"
            }
            | {"images": product["images"]}
        )
    status = Counter(row["prepare_status"] for row in crops)
    manifest = {
        "metadata": {
            "run_name": RUN_NAME,
            "created_at": now_text(),
            "target_products": len(products),
            "products_with_candidate_images": sum(
                bool(product["images"]) for product in products.values()
            ),
            "unique_candidate_images": len(target_hashes),
            "images_with_heading": sum(bool(row["heading_positions"]) for row in image_rows),
            "crop_count": len(crops),
            "prepare_status": dict(status),
            "shards": shards,
        },
        "products": product_rows,
        "images": image_rows,
        "crops": crops,
    }
    save_json(MANIFEST_PATH, manifest)
    print(json.dumps(manifest["metadata"], ensure_ascii=False, indent=2), flush=True)
    return manifest["metadata"]


def run_ocr(
    shards: int = DEFAULT_SHARDS,
    workers: int = 8,
) -> None:
    failed: list[int] = []
    for start in range(0, shards, max(1, workers)):
        processes: list[tuple[int, subprocess.Popen[Any]]] = []
        for shard in range(start, min(shards, start + max(1, workers))):
            command = [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(OCR_SCRIPT),
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
            print(f"targeted_ocr_shard={shard} exit={code}", flush=True)
            if code:
                failed.append(shard)
    if failed:
        raise RuntimeError(f"Targeted OCR shards failed: {failed}")


def load_ocr_results(shards: int) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for shard in range(shards):
        path = RUN_DIR / f"layout_ocr_{shard:02d}.json"
        if not path.exists():
            continue
        rows = load_json(path)
        if isinstance(rows, dict):
            rows = [rows]
        for row in rows:
            result[Path(str(row.get("file") or "")).name] = row
    return result


def init_db(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS stg_dimension_targeted_ocr_crop (
            run_name TEXT NOT NULL,
            processed_at TEXT NOT NULL,
            product_id TEXT NOT NULL,
            image_url TEXT,
            crop_no INTEGER,
            crop_file TEXT,
            crop_top INTEGER,
            crop_bottom INTEGER,
            crop_scale REAL,
            stage2_status TEXT,
            stage2_text TEXT,
            stage2_word_count INTEGER,
            candidate_count INTEGER,
            PRIMARY KEY(run_name, product_id, image_url, crop_no)
        );
        CREATE TABLE IF NOT EXISTS stg_dimension_targeted_ocr_candidate (
            run_name TEXT NOT NULL,
            processed_at TEXT NOT NULL,
            product_id TEXT NOT NULL,
            image_url TEXT,
            crop_no INTEGER,
            candidate_no INTEGER,
            candidate_role TEXT,
            decision_status TEXT,
            raw_notation TEXT,
            normalized_axis_mapping TEXT,
            option_label TEXT,
            shape_type TEXT,
            w_mm REAL,
            d_mm REAL,
            h_mm REAL,
            diameter_mm REAL,
            candidate_score INTEGER,
            evidence_text TEXT,
            candidate_json TEXT,
            PRIMARY KEY(run_name, product_id, image_url, crop_no, candidate_no)
        );
        DROP VIEW IF EXISTS vw_dimension_targeted_ocr_candidates_current;
        CREATE VIEW vw_dimension_targeted_ocr_candidates_current AS
        SELECT *
        FROM stg_dimension_targeted_ocr_candidate
        WHERE run_name IN (
            'dimension_targeted_reocr_user_cases_v1',
            'dimension_targeted_reocr_remaining_v1',
            'dimension_partial_fusion_v1',
            'dimension_blocked_candidate_exposure_v1',
            'dimension_pass2_raw_exposure_v1'
        );
        """
    )


def load_results(shards: int = DEFAULT_SHARDS) -> dict[str, Any]:
    timestamp = now_text()
    manifest = load_json(MANIFEST_PATH)
    results = load_ocr_results(shards)
    crop_by_hash: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for crop in manifest["crops"]:
        crop_by_hash[crop["url_hash"]].append(crop)

    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    init_db(connection)
    connection.execute(
        "DELETE FROM stg_dimension_targeted_ocr_crop WHERE run_name=?", (RUN_NAME,)
    )
    connection.execute(
        "DELETE FROM stg_dimension_targeted_ocr_candidate WHERE run_name=?",
        (RUN_NAME,),
    )

    product_count = 0
    products_with_candidates: set[str] = set()
    total_candidates = 0
    crop_rows = 0
    for product in manifest["products"]:
        product_count += 1
        seen_urls: set[str] = set()
        for image in product["images"]:
            digest = image["url_hash"]
            image_url = image["image_url"]
            if image_url in seen_urls:
                continue
            seen_urls.add(image_url)
            for crop in crop_by_hash.get(digest, []):
                result = results.get(crop["crop_file"]) or {}
                text = str(result.get("text") or "")
                candidates = extract_candidates(
                    text,
                    product_name=product["product_name"],
                    small_category=product["small_category"],
                )
                crop_rows += 1
                connection.execute(
                    "INSERT INTO stg_dimension_targeted_ocr_crop VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    (
                        RUN_NAME,
                        timestamp,
                        product["product_id"],
                        image_url,
                        crop["crop_no"],
                        crop["crop_file"],
                        crop.get("crop_top"),
                        crop.get("crop_bottom"),
                        crop.get("crop_scale"),
                        result.get("status") or "NOT_RUN",
                        text,
                        len(result.get("words") or []),
                        len(candidates),
                    ),
                )
                for candidate_no, candidate in enumerate(candidates, 1):
                    connection.execute(
                        "INSERT INTO stg_dimension_targeted_ocr_candidate VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                        (
                            RUN_NAME,
                            timestamp,
                            product["product_id"],
                            image_url,
                            crop["crop_no"],
                            candidate_no,
                            candidate["candidate_role"],
                            candidate["decision_status"],
                            candidate["raw_notation"],
                            candidate["normalized_axis_mapping"],
                            candidate["option_label"],
                            candidate["shape_type"],
                            candidate["w_mm"],
                            candidate["d_mm"],
                            candidate["h_mm"],
                            candidate["diameter_mm"],
                            candidate["candidate_score"],
                            candidate["context_text"],
                            json.dumps(candidate, ensure_ascii=False),
                        ),
                    )
                if candidates:
                    products_with_candidates.add(product["product_id"])
                    total_candidates += len(candidates)
    connection.commit()
    connection.close()

    statuses = Counter(
        row.get("status") or "NOT_FOUND" for row in results.values()
    )
    summary = {
        "run_name": RUN_NAME,
        "processed_at": timestamp,
        "target_products": product_count,
        "unique_crops": len(manifest["crops"]),
        "product_crop_rows": crop_rows,
        "ocr_status": dict(statuses),
        "products_with_candidates": len(products_with_candidates),
        "dimension_candidates": total_candidates,
    }
    save_json(SUMMARY_PATH, summary)
    print(json.dumps(summary, ensure_ascii=False, indent=2), flush=True)
    return summary


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("stage", choices=("prepare", "ocr", "load", "all"))
    parser.add_argument("--workers", type=int, default=8)
    parser.add_argument("--ocr-workers", type=int, default=8)
    parser.add_argument("--shards", type=int, default=DEFAULT_SHARDS)
    args = parser.parse_args()
    if args.stage in {"prepare", "all"}:
        prepare(args.workers, args.shards)
    if args.stage in {"ocr", "all"}:
        run_ocr(args.shards, args.ocr_workers)
    if args.stage in {"load", "all"}:
        load_results(args.shards)


if __name__ == "__main__":
    main()
