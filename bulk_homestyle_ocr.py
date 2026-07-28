from __future__ import annotations

import argparse
import concurrent.futures
import html
import io
import json
import re
import shutil
import sqlite3
import subprocess
import sys
import urllib.parse
from datetime import datetime
from pathlib import Path
from typing import Any

from PIL import Image, ImageFile

import build_homestyle_bulk_workbook as workbook
from bulk_homestyle_collect import DB_PATH, RUN_DIR, pack, request_bytes, unpack


ROOT = Path(__file__).resolve().parent
OCR_ROOT = RUN_DIR / "ocr"
WINDOWS_OCR = ROOT / "run_windows_ocr.ps1"
ImageFile.LOAD_TRUNCATED_IMAGES = True
Image.MAX_IMAGE_PIXELS = None


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def normalize_url(url: str) -> str:
    value = html.unescape(str(url or "").strip())
    if value.startswith("//"):
        value = "https:" + value
    elif value.startswith("/goods/"):
        value = "https://static-store.lge.co.kr" + value
    if not value.startswith(("http://", "https://")):
        return value
    # 공급사 상세 이미지 URL에는 공백이나 한글 경로가 그대로 들어오는 경우가
    # 많다. urllib/http.client가 이를 제어문자 또는 ASCII 오류로 거절하므로
    # 이미 인코딩된 % 값은 보존하면서 path/query를 안전하게 인코딩한다.
    parsed = urllib.parse.urlsplit(value)
    path = urllib.parse.quote(parsed.path, safe="/%:@-._~!$&'()*+,;=%")
    query = urllib.parse.quote(parsed.query, safe="=&%/:;+,-._~!$'()*@")
    return urllib.parse.urlunsplit(
        (parsed.scheme, parsed.netloc, path, query, parsed.fragment)
    )


def detail_images(markup: str) -> list[dict[str, Any]]:
    images: list[dict[str, Any]] = []
    seen: set[str] = set()
    matches = re.finditer(
        r"(?:src|data-src|data-original)\s*=\s*[\"']([^\"']+)",
        markup or "",
        re.I,
    )
    for position, match in enumerate(matches):
        url = normalize_url(match.group(1))
        if not url.startswith(("http://", "https://")) or url in seen:
            continue
        seen.add(url)
        images.append(
            {
                "url": url,
                "position": position,
                "tag": "",
                "alt": "",
            }
        )
    total = len(images)
    for image in images:
        value = (image["url"] + " " + image["alt"] + " " + image["tag"]).casefold()
        score = 50.0
        if re.search(
            r"(?:size|spec(?:ification)?|dimension|measure|drawing|dim[_-]|"
            r"치수|규격|사이즈|크기)",
            value,
            re.I,
        ):
            score -= 45
        elif re.search(r"(?:info|guide|product[_-]?info|detail[_-]?info)", value, re.I):
            score -= 22
        elif re.search(r"(?:detail|point|feature|intro)", value, re.I):
            score -= 8
        if re.search(
            r"(?:delivery|shipping|notice|banner|event|coupon|gift|warranty|"
            r"common|footer|membership|card|benefit|배송|공지|배너|이벤트)",
            value,
            re.I,
        ):
            score += 60
        # When filenames carry no semantic hint, dimensions commonly appear in
        # the latter part of the product-specific detail sequence.
        if total > 1:
            score -= 8 * (image["position"] / max(total - 1, 1))
        image["score"] = round(score, 3)
        image["total_images"] = total
    return sorted(images, key=lambda row: (row["score"], -row["position"]))


def structured_dimensions(data: dict[str, Any], html_blob: dict[str, Any]) -> list[dict[str, Any]]:
    notifications = workbook.notification_items(data)
    texts = [
        ("상품정보고시", value)
        for value in workbook.notification_values(
            notifications, "크기", "치수", "규격", "사이즈"
        )
    ]
    colors = workbook.joined(
        workbook.notification_values(notifications, "색상", "컬러"), ""
    )
    for group in workbook.option_groups(data, colors):
        if group["style"] == "사이즈":
            texts.extend(("상품 옵션", item["name"]) for item in group["items"])
    # Do not scan PDP/detail HTML here. Those sources are merged by the final
    # workbook builder; excluding them from this fast prefilter only makes the
    # OCR candidate pool wider and avoids expensive regex work over long pages.
    texts.append(("상품명", str(data.get("productName") or "")))
    return workbook.dimension_records(texts)


def candidate_products(limit: int = 0) -> list[dict[str, Any]]:
    connection = sqlite3.connect(DB_PATH)
    rows = connection.execute(
        """
        SELECT p.product_id, s.goods_blob
        FROM products p JOIN sources s ON s.product_id=p.product_id
        WHERE s.goods_status=200
        ORDER BY p.product_id
        """
    ).fetchall()
    result = []
    for scan_index, (product_id, goods_blob) in enumerate(rows, 1):
        data = (unpack(goods_blob) or {}).get("data") or {}
        records = structured_dimensions(data, {})
        if any(
            row.get("w_mm") is not None
            and row.get("d_mm") is not None
            and row.get("h_mm") is not None
            for row in records
        ):
            continue
        images = detail_images(str(data.get("detailInfo") or ""))
        result.append(
            {
                "product_id": product_id,
                "product_name": str(data.get("productName") or ""),
                "images": images[:1],
            }
        )
        if limit and len(result) >= limit:
            break
        if scan_index % 1000 == 0:
            print(
                f"candidate_scan={scan_index}/{len(rows)} selected={len(result)}",
                flush=True,
            )
    connection.close()
    return result


def prepare_image(task: dict[str, Any], output_dir: Path, shards: int) -> dict[str, Any]:
    product_id = task["product_id"]
    if not task["images"]:
        return {**task, "download_status": "NO_IMAGE", "selected": None, "file": ""}
    selected = task["images"][0]
    shard = int(product_id[-4:]) % shards
    destination = output_dir / f"shard_{shard:02d}" / f"{product_id}.jpg"
    try:
        status, body = request_bytes(selected["url"], timeout=45, attempts=3)
        if status != 200 or not body:
            raise RuntimeError(f"HTTP {status}; bytes={len(body)}")
        with Image.open(io.BytesIO(body)) as source:
            source.seek(0)
            image = source.convert("RGB")
            if image.width > 2200:
                height = max(1, round(image.height * 2200 / image.width))
                image = image.resize((2200, height), Image.Resampling.LANCZOS)
            destination.parent.mkdir(parents=True, exist_ok=True)
            image.save(destination, "JPEG", quality=92, optimize=True)
            width, height = image.size
        return {
            **task,
            "download_status": "SUCCESS",
            "selected": selected,
            "file": str(destination),
            "width": width,
            "height": height,
        }
    except Exception as exc:  # recorded per product; does not stop the batch
        return {
            **task,
            "download_status": "ERROR",
            "selected": selected,
            "file": "",
            "error": f"{type(exc).__name__}: {exc}",
        }


def prepare(run_name: str, limit: int, workers: int, shards: int) -> Path:
    run_dir = OCR_ROOT / run_name
    if run_dir.exists():
        resolved = run_dir.resolve()
        if OCR_ROOT.resolve() not in resolved.parents:
            raise RuntimeError(f"Unsafe OCR run path: {resolved}")
        shutil.rmtree(resolved)
    run_dir.mkdir(parents=True)
    for shard in range(shards):
        (run_dir / f"shard_{shard:02d}").mkdir()
    candidates = candidate_products(limit)
    results = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(prepare_image, task, run_dir, shards) for task in candidates]
        for index, future in enumerate(concurrent.futures.as_completed(futures), 1):
            results.append(future.result())
            if index % 250 == 0 or index == len(futures):
                print(f"downloaded={index}/{len(futures)}", flush=True)
    results.sort(key=lambda row: row["product_id"])
    manifest = {
        "metadata": {
            "run_name": run_name,
            "created_at": now_text(),
            "candidate_count": len(candidates),
            "download_success": sum(row["download_status"] == "SUCCESS" for row in results),
            "no_image": sum(row["download_status"] == "NO_IMAGE" for row in results),
            "download_error": sum(row["download_status"] == "ERROR" for row in results),
            "shards": shards,
        },
        "products": results,
    }
    manifest_path = run_dir / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(manifest["metadata"], ensure_ascii=False), flush=True)
    return manifest_path


def run_ocr(run_name: str, shards: int) -> None:
    run_dir = OCR_ROOT / run_name
    processes = []
    for shard in range(shards):
        input_dir = run_dir / f"shard_{shard:02d}"
        output_json = run_dir / f"ocr_{shard:02d}.json"
        command = [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(WINDOWS_OCR),
            "-InputDirectory",
            str(input_dir),
            "-OutputJson",
            str(output_json),
            "-LanguageTag",
            "ko",
        ]
        processes.append((shard, subprocess.Popen(command)))
    failed = []
    for shard, process in processes:
        code = process.wait()
        print(f"ocr_shard={shard} exit={code}", flush=True)
        if code:
            failed.append(shard)
    if failed:
        raise RuntimeError(f"OCR shards failed: {failed}")


def ocr_dimension_text(text: str) -> str:
    lines = [line.strip() for line in str(text or "").splitlines() if line.strip()]
    result = []
    signal = re.compile(
        r"(?:\d\s*(?:mm|cm|㎜)|(?:W|D|H)\s*[:=]?\s*\d|"
        r"가로|세로|깊이|높이|너비|폭|지름|직경|규격|치수|사이즈)",
        re.I,
    )
    for index, line in enumerate(lines):
        if signal.search(line):
            start, end = max(0, index - 1), min(len(lines), index + 2)
            for value in lines[start:end]:
                if value not in result:
                    result.append(value)
    return "\n".join(result)


def merge(run_name: str, shards: int, write_db: bool) -> dict[str, Any]:
    run_dir = OCR_ROOT / run_name
    manifest = json.loads((run_dir / "manifest.json").read_text(encoding="utf-8"))
    ocr_rows: dict[str, dict[str, Any]] = {}
    for shard in range(shards):
        path = run_dir / f"ocr_{shard:02d}.json"
        data = json.loads(path.read_text(encoding="utf-8-sig"))
        if isinstance(data, dict):
            data = [data]
        for row in data:
            ocr_rows[Path(row.get("file") or "").stem] = row
    summary = {
        "candidate_count": len(manifest["products"]),
        "download_success": 0,
        "ocr_success": 0,
        "ocr_nonempty": 0,
        "dimension_signal": 0,
        "new_complete_dimension": 0,
        "new_partial_dimension": 0,
        "no_image": 0,
        "download_error": 0,
        "ocr_error": 0,
    }
    updates = []
    for product in manifest["products"]:
        product_id = product["product_id"]
        download_status = product["download_status"]
        if download_status != "SUCCESS":
            summary["no_image" if download_status == "NO_IMAGE" else "download_error"] += 1
            status = 204 if download_status == "NO_IMAGE" else 502
            payload = {
                "engine": "Windows.Media.Ocr/ko",
                "run_name": run_name,
                "download_status": download_status,
                "selected": product.get("selected"),
                "error": product.get("error", ""),
                "combined_text": "",
                "dimension_text": "",
                "dimensions": [],
            }
        else:
            summary["download_success"] += 1
            row = ocr_rows.get(product_id) or {}
            success = row.get("status") == "SUCCESS"
            status = 200 if success else 500
            summary["ocr_success" if success else "ocr_error"] += 1
            raw_text = str(row.get("text") or "")
            if raw_text.strip():
                summary["ocr_nonempty"] += 1
            dimension_text = ocr_dimension_text(raw_text)
            dimensions = workbook.dimension_records([("상세 이미지 OCR", dimension_text)]) if dimension_text else []
            if dimension_text:
                summary["dimension_signal"] += 1
            complete = any(
                item.get("w_mm") is not None
                and item.get("d_mm") is not None
                and item.get("h_mm") is not None
                for item in dimensions
            )
            if complete:
                summary["new_complete_dimension"] += 1
            elif dimensions:
                summary["new_partial_dimension"] += 1
            payload = {
                "engine": "Windows.Media.Ocr/ko",
                "run_name": run_name,
                "download_status": download_status,
                "selected": product.get("selected"),
                "width": product.get("width"),
                "height": product.get("height"),
                "ocr": row,
                "combined_text": raw_text,
                "dimension_text": dimension_text,
                "dimensions": dimensions,
            }
        updates.append((status, pack(payload), now_text(), product_id))
    if write_db:
        connection = sqlite3.connect(DB_PATH)
        connection.executemany(
            "UPDATE sources SET ocr_status=?, ocr_blob=?, ocr_at=? WHERE product_id=?",
            updates,
        )
        connection.commit()
        connection.close()
    (run_dir / "summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2), flush=True)
    return summary


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("phase", choices=("prepare", "run", "merge", "all"))
    parser.add_argument("--run-name", default="pass1_all")
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--workers", type=int, default=20)
    parser.add_argument("--shards", type=int, default=4)
    parser.add_argument("--write-db", action="store_true")
    args = parser.parse_args()
    if args.phase in ("prepare", "all"):
        prepare(args.run_name, args.limit, args.workers, args.shards)
    if args.phase in ("run", "all"):
        run_ocr(args.run_name, args.shards)
    if args.phase in ("merge", "all"):
        merge(args.run_name, args.shards, args.write_db or args.phase == "all")


if __name__ == "__main__":
    main()
