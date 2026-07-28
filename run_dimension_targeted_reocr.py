from __future__ import annotations

import argparse
import concurrent.futures
import hashlib
import io
import json
import re
import shutil
import sqlite3
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Any

from PIL import Image, ImageEnhance, ImageFile, ImageOps

import bulk_homestyle_ocr as ocr
from bulk_homestyle_collect import DB_PATH, request_bytes, unpack
from dimension_context_normalizer import category_profile, extract_candidates
from run_dimension_scan_pass1 import discover_previous_cache


ROOT = Path(__file__).resolve().parent
RUN_NAME = "dimension_targeted_reocr_user_cases_v1"
RUN_DIR = ROOT / "homestyle_bulk_run" / "ocr" / RUN_NAME
IMAGE_DIR = RUN_DIR / "stage1_images"
CROP_DIR = RUN_DIR / "stage2_crops"
MANIFEST_PATH = RUN_DIR / "manifest.json"
STAGE1_JSON = RUN_DIR / "stage1_layout_ocr.json"
STAGE2_JSON = RUN_DIR / "stage2_layout_ocr.json"
SUMMARY_PATH = RUN_DIR / "summary.json"
OCR_SCRIPT = ROOT / "run_windows_ocr_layout.ps1"

ImageFile.LOAD_TRUNCATED_IMAGES = True
Image.MAX_IMAGE_PIXELS = None

HEADING_RE = re.compile(
    r"(?:^|[^a-z가-힣])(?:sizes?|dimensions?|measurements?|specifications?|"
    r"info(?:rmation)?|detail|제품\s*사이즈|상품\s*사이즈|사이즈|규격|치수|"
    r"제품\s*크기|한\s*눈에\s*보기)(?:$|[^a-z가-힣])",
    re.I,
)
AXIS_WORD_RE = re.compile(r"^(?:w|d|h|l|r|width|depth|height|가로|세로|깊이|높이|지름)$", re.I)


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def save_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def image_hash(url: str) -> str:
    return hashlib.sha1(url.encode("utf-8")).hexdigest()


def current_target_ids(connection: sqlite3.Connection) -> list[str]:
    return [
        row[0]
        for row in connection.execute(
            """
            SELECT DISTINCT p.product_id
            FROM vw_dimension_context_products_current p
            JOIN ref_dimension_context_rule_testcase t
              ON t.product_id = p.product_id
            LEFT JOIN vw_dimension_context_regression_current r
              ON r.product_id = p.product_id
            WHERE p.context_status IN ('NO_CANDIDATE', 'REOCR_REQUIRED')
               OR r.result_status = 'NEEDS_TARGETED_REOCR'
            ORDER BY p.product_id
            """
        )
    ]


def pass1_selected_urls(connection: sqlite3.Connection, product_id: str) -> set[str]:
    rows = list(
        connection.execute(
            """
            SELECT image_url, is_candidate, candidate_score
            FROM stg_dimension_scan_pass1_product_image
            WHERE product_id = ?
            ORDER BY is_candidate DESC, candidate_score DESC, image_order
            """,
            (product_id,),
        )
    )
    if not rows:
        return set()
    selected = [row[0] for row in rows if row[1]]
    if not selected:
        selected = [row[0] for row in rows[:3]]
    return set(selected)


def source_image_rows(connection: sqlite3.Connection) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for product_id in current_target_ids(connection):
        product = connection.execute(
            """
            SELECT m.product_id, m.product_name, m.mid_category, m.small_category,
                   s.goods_blob
            FROM vw_dimension_classification_master_current m
            JOIN sources s ON s.product_id = m.product_id
            WHERE m.product_id = ?
            """,
            (product_id,),
        ).fetchone()
        if not product:
            continue
        data = (unpack(product["goods_blob"]) or {}).get("data") or {}
        images = sorted(
            ocr.detail_images(str(data.get("detailInfo") or "")),
            key=lambda value: int(value.get("position") or 0),
        )
        selected_urls = pass1_selected_urls(connection, product_id)
        if selected_urls:
            images = [
                image
                for image in images
                if ocr.normalize_url(str(image.get("url") or "")) in selected_urls
            ]
        for image_order, image in enumerate(images, 1):
            url = ocr.normalize_url(str(image.get("url") or ""))
            if not url:
                continue
            rows.append(
                {
                    "product_id": product_id,
                    "product_name": str(product["product_name"] or ""),
                    "mid_category": str(product["mid_category"] or ""),
                    "small_category": str(product["small_category"] or ""),
                    "image_order": image_order,
                    "source_position": int(image.get("position") or 0),
                    "image_url": url,
                    "url_hash": image_hash(url),
                }
            )
    return rows


def prepare_one(row: dict[str, Any], cache: dict[str, Path]) -> dict[str, Any]:
    destination = IMAGE_DIR / (
        f"{row['product_id']}_{row['image_order']:02d}_{row['url_hash'][:12]}.jpg"
    )
    try:
        if destination.exists() and destination.stat().st_size:
            with Image.open(destination) as image:
                width, height = image.size
            return {**row, "file": destination.name, "width": width, "height": height, "prepare_status": "REUSED_RUN", "error": ""}

        body = b""
        status = 0
        try:
            status, body = request_bytes(row["image_url"], timeout=40, attempts=2)
        except Exception:
            status = 0
        if status == 200 and body:
            source: io.BytesIO | Path = io.BytesIO(body)
            prepare_status = "DOWNLOADED"
        elif row["image_url"] in cache:
            source = cache[row["image_url"]]
            prepare_status = "REUSED_CACHE"
        else:
            raise RuntimeError(f"image unavailable: HTTP {status}")

        with Image.open(source) as raw:
            raw.seek(0)
            image = raw.convert("RGB")
            if image.width > 2400:
                new_height = max(1, round(image.height * 2400 / image.width))
                image = image.resize((2400, new_height), Image.Resampling.LANCZOS)
            image.save(destination, "JPEG", quality=94, optimize=True)
            width, height = image.size
        return {**row, "file": destination.name, "width": width, "height": height, "prepare_status": prepare_status, "error": ""}
    except Exception as exc:
        return {**row, "file": "", "width": None, "height": None, "prepare_status": "ERROR", "error": f"{type(exc).__name__}: {exc}"}


def prepare(workers: int = 8) -> dict[str, Any]:
    RUN_DIR.mkdir(parents=True, exist_ok=True)
    IMAGE_DIR.mkdir(parents=True, exist_ok=True)
    CROP_DIR.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    rows = source_image_rows(connection)
    connection.close()
    cache, _ = discover_previous_cache()
    with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, workers)) as pool:
        prepared = list(pool.map(lambda row: prepare_one(row, cache), rows))
    manifest = {
        "run_name": RUN_NAME,
        "created_at": now_text(),
        "target_product_count": len({row["product_id"] for row in rows}),
        "source_image_count": len(rows),
        "prepared_success": sum(row["prepare_status"] != "ERROR" for row in prepared),
        "prepared_error": sum(row["prepare_status"] == "ERROR" for row in prepared),
        "images": prepared,
        "crops": [],
    }
    save_json(MANIFEST_PATH, manifest)
    print(json.dumps({key: value for key, value in manifest.items() if key not in {"images", "crops"}}, ensure_ascii=False, indent=2))
    return manifest


def run_layout_ocr(input_dir: Path, output_json: Path) -> None:
    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(OCR_SCRIPT),
        "-InputDirectory",
        str(input_dir),
        "-OutputJson",
        str(output_json),
    ]
    completed = subprocess.run(command, cwd=ROOT, check=False)
    if completed.returncode:
        raise RuntimeError(f"layout OCR failed: exit={completed.returncode}")


def trigger_positions(result: dict[str, Any]) -> list[float]:
    words = result.get("words") or []
    positions: list[float] = []
    for word in words:
        text = str(word.get("text") or "").strip()
        padded = f" {text} "
        if HEADING_RE.search(padded) or AXIS_WORD_RE.match(text):
            positions.append(float(word.get("y") or 0))
    if not positions and HEADING_RE.search(str(result.get("text") or "")):
        positions.append(0.0)
    clustered: list[float] = []
    for position in sorted(positions):
        if not clustered or position - clustered[-1] > 550:
            clustered.append(position)
    return clustered[:8]


def build_crops() -> dict[str, Any]:
    manifest = load_json(MANIFEST_PATH)
    results = {row.get("file"): row for row in load_json(STAGE1_JSON)}
    crops: list[dict[str, Any]] = []
    for image_row in manifest["images"]:
        if not image_row.get("file"):
            continue
        result = results.get(image_row["file"]) or {}
        if result.get("status") != "SUCCESS":
            continue
        positions = trigger_positions(result)
        source_path = IMAGE_DIR / image_row["file"]
        with Image.open(source_path) as raw:
            source = raw.convert("RGB")
            for crop_no, y in enumerate(positions, 1):
                top = max(0, int(y) - 300)
                bottom = min(source.height, int(y) + 1900)
                if bottom - top < 300:
                    continue
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
                crop_name = (
                    f"{image_row['product_id']}_{image_row['image_order']:02d}_"
                    f"crop{crop_no:02d}_{image_row['url_hash'][:10]}.png"
                )
                crop_path = CROP_DIR / crop_name
                enhanced.save(crop_path, "PNG", optimize=True)
                crops.append(
                    {
                        "product_id": image_row["product_id"],
                        "product_name": image_row["product_name"],
                        "mid_category": image_row["mid_category"],
                        "small_category": image_row["small_category"],
                        "image_order": image_row["image_order"],
                        "image_url": image_row["image_url"],
                        "source_file": image_row["file"],
                        "crop_file": crop_name,
                        "crop_no": crop_no,
                        "crop_top": top,
                        "crop_bottom": bottom,
                        "scale": scale,
                        "trigger_y": y,
                    }
                )
    manifest["crops"] = crops
    manifest["stage1_success"] = sum(row.get("status") == "SUCCESS" for row in results.values())
    manifest["stage2_crop_count"] = len(crops)
    save_json(MANIFEST_PATH, manifest)
    print(json.dumps({"stage1_success": manifest["stage1_success"], "stage2_crop_count": len(crops)}, ensure_ascii=False, indent=2))
    return manifest


def init_db(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS stg_dimension_targeted_ocr_image (
            run_name TEXT NOT NULL,
            processed_at TEXT NOT NULL,
            product_id TEXT NOT NULL,
            product_name TEXT,
            small_category TEXT,
            image_order INTEGER,
            image_url TEXT,
            local_file TEXT,
            prepare_status TEXT,
            stage1_status TEXT,
            stage1_text TEXT,
            stage1_word_count INTEGER,
            crop_count INTEGER,
            error TEXT,
            PRIMARY KEY(run_name, product_id, image_order, image_url)
        );
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


def load_results() -> dict[str, Any]:
    timestamp = now_text()
    manifest = load_json(MANIFEST_PATH)
    stage1 = {row.get("file"): row for row in load_json(STAGE1_JSON)}
    stage2 = {row.get("file"): row for row in load_json(STAGE2_JSON)}
    candidates_by_crop: dict[str, list[dict[str, Any]]] = {}
    for crop in manifest["crops"]:
        result = stage2.get(crop["crop_file"]) or {}
        text = str(result.get("text") or "")
        candidates_by_crop[crop["crop_file"]] = extract_candidates(
            text,
            product_name=crop["product_name"],
            small_category=crop["small_category"],
        )
    stage1_candidates_by_file: dict[str, list[dict[str, Any]]] = {}
    for image in manifest["images"]:
        result = stage1.get(image.get("file")) or {}
        stage1_candidates_by_file[str(image.get("file") or "")] = extract_candidates(
            str(result.get("text") or ""),
            product_name=image["product_name"],
            small_category=image["small_category"],
        )

    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    init_db(connection)
    connection.execute(
        "DELETE FROM stg_dimension_targeted_ocr_image WHERE run_name=?", (RUN_NAME,)
    )
    connection.execute(
        "DELETE FROM stg_dimension_targeted_ocr_crop WHERE run_name=?", (RUN_NAME,)
    )
    connection.execute(
        "DELETE FROM stg_dimension_targeted_ocr_candidate WHERE run_name=?", (RUN_NAME,)
    )
    for image in manifest["images"]:
        result = stage1.get(image.get("file")) or {}
        crop_count = sum(
            crop["image_url"] == image["image_url"]
            and crop["product_id"] == image["product_id"]
            for crop in manifest["crops"]
        )
        connection.execute(
            "INSERT INTO stg_dimension_targeted_ocr_image VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (
                RUN_NAME,
                timestamp,
                image["product_id"],
                image["product_name"],
                image["small_category"],
                image["image_order"],
                image["image_url"],
                image.get("file") or "",
                image["prepare_status"],
                result.get("status") or "NOT_RUN",
                result.get("text") or "",
                len(result.get("words") or []),
                crop_count,
                image.get("error") or result.get("error") or "",
            ),
        )
    total_candidates = 0
    for image in manifest["images"]:
        candidates = stage1_candidates_by_file.get(str(image.get("file") or ""), [])
        total_candidates += len(candidates)
        for candidate_no, candidate in enumerate(candidates, 1):
            connection.execute(
                "INSERT INTO stg_dimension_targeted_ocr_candidate VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (
                    RUN_NAME,
                    timestamp,
                    image["product_id"],
                    image["image_url"],
                    0,
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
    for crop in manifest["crops"]:
        result = stage2.get(crop["crop_file"]) or {}
        candidates = candidates_by_crop[crop["crop_file"]]
        total_candidates += len(candidates)
        connection.execute(
            "INSERT INTO stg_dimension_targeted_ocr_crop VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (
                RUN_NAME,
                timestamp,
                crop["product_id"],
                crop["image_url"],
                crop["crop_no"],
                crop["crop_file"],
                crop["crop_top"],
                crop["crop_bottom"],
                crop["scale"],
                result.get("status") or "NOT_RUN",
                result.get("text") or "",
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
                    crop["product_id"],
                    crop["image_url"],
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
    connection.commit()
    connection.close()
    summary = {
        "run_name": RUN_NAME,
        "processed_at": timestamp,
        "target_products": manifest["target_product_count"],
        "source_images": manifest["source_image_count"],
        "stage1_success": manifest.get("stage1_success", 0),
        "stage2_crops": len(manifest["crops"]),
        "stage2_success": sum(row.get("status") == "SUCCESS" for row in stage2.values()),
        "dimension_candidates": total_candidates,
    }
    save_json(SUMMARY_PATH, summary)
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return summary


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "stage",
        choices=("prepare", "stage1", "crops", "stage2", "load", "all"),
    )
    parser.add_argument("--workers", type=int, default=8)
    args = parser.parse_args()
    if args.stage in {"prepare", "all"}:
        prepare(args.workers)
    if args.stage in {"stage1", "all"}:
        run_layout_ocr(IMAGE_DIR, STAGE1_JSON)
    if args.stage in {"crops", "all"}:
        build_crops()
    if args.stage in {"stage2", "all"}:
        run_layout_ocr(CROP_DIR, STAGE2_JSON)
    if args.stage in {"load", "all"}:
        load_results()


if __name__ == "__main__":
    main()
