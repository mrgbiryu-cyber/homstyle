from __future__ import annotations

import argparse
import concurrent.futures
import json
import sqlite3
from collections import Counter, defaultdict, deque
from pathlib import Path
from typing import Any

import build_homestyle_bulk_workbook as workbook
import bulk_homestyle_ocr as ocr
from bulk_homestyle_collect import DB_PATH, unpack
from research_remaining_dimensions import product_sources


RUN_NAME = "research_missing_sample"
RUN_DIR = ocr.OCR_ROOT / RUN_NAME
RESEARCH_PATH = ocr.RUN_DIR / "remaining_dimension_research.json"
OUTPUT_PATH = ocr.RUN_DIR / "research_missing_ocr_sample.json"
SHARDS = 8


def categories_by_scope(connection: sqlite3.Connection) -> dict[str, str]:
    return {
        row[0]: row[1] or row[0]
        for row in connection.execute("SELECT scope_id,small_name FROM categories")
    }


def evenly_select(pool: list[dict[str, Any]], limit: int) -> list[dict[str, Any]]:
    """Take a deterministic category round-robin sample."""
    buckets: dict[str, deque[dict[str, Any]]] = defaultdict(deque)
    for row in sorted(pool, key=lambda item: item["product_id"]):
        buckets[row["category"]].append(row)
    result: list[dict[str, Any]] = []
    names = sorted(buckets)
    while names and len(result) < limit:
        next_names = []
        for name in names:
            if len(result) >= limit:
                break
            if buckets[name]:
                result.append(buckets[name].popleft())
            if buckets[name]:
                next_names.append(name)
        names = next_names
    return result


def prepare() -> None:
    research = json.loads(RESEARCH_PATH.read_text(encoding="utf-8"))
    residual_ids = set(research["product_flags"])
    connection = sqlite3.connect(DB_PATH)
    category_names = categories_by_scope(connection)
    pools: dict[str, list[dict[str, Any]]] = defaultdict(list)
    rows = connection.execute(
        """
        SELECT p.product_id,p.category_scope_ids,s.goods_blob,s.ocr_status,s.ocr_blob
        FROM products p JOIN sources s ON s.product_id=p.product_id
        WHERE s.goods_status=200 ORDER BY p.product_id
        """
    ).fetchall()
    for product_id, scope_json, goods_blob, ocr_status, ocr_blob in rows:
        if product_id not in residual_ids:
            continue
        data = (unpack(goods_blob) or {}).get("data") or {}
        old_ocr = unpack(ocr_blob) or {}
        categories = [
            category_names.get(scope, scope)
            for scope in json.loads(scope_json or "[]")
        ]
        category = categories[0] if categories else "분류 미확보"
        images = ocr.detail_images(str(data.get("detailInfo") or ""))
        base = {
            "product_id": product_id,
            "product_name": str(data.get("productName") or ""),
            "category": category,
            "categories": categories,
        }
        if int(ocr_status or 0) == 502:
            error = str(old_ocr.get("error") or "")
            if "control characters" in error or "UnicodeEncodeError" in error:
                selected = old_ocr.get("selected") or {}
                selected_url = ocr.normalize_url(str(selected.get("url") or ""))
                if selected_url:
                    pools["URL 인코딩 실패 재시도"].append(
                        {**base, "images": [{**selected, "url": selected_url}]}
                    )
        elif int(ocr_status or 0) == 0 and images:
            pools["L 보류로 OCR 미실행"].append({**base, "images": images[:1]})
        elif (
            int(ocr_status or 0) == 200
            and not old_ocr.get("dimension_text")
            and len(images) >= 2
        ):
            old_url = ocr.normalize_url(
                str((old_ocr.get("selected") or {}).get("url") or "")
            )
            alternatives = [item for item in images if item["url"] != old_url]
            if alternatives:
                pools["첫 이미지 규격 없음·대체 이미지"].append(
                    {**base, "images": alternatives[:1]}
                )
    connection.close()

    limits = {
        "URL 인코딩 실패 재시도": 60,
        "L 보류로 OCR 미실행": 40,
        "첫 이미지 규격 없음·대체 이미지": 100,
    }
    tasks: list[dict[str, Any]] = []
    pool_sizes = {}
    for group, limit in limits.items():
        pool_sizes[group] = len(pools[group])
        for row in evenly_select(pools[group], limit):
            tasks.append({**row, "research_group": group})

    RUN_DIR.mkdir(parents=True, exist_ok=True)
    for shard in range(SHARDS):
        (RUN_DIR / f"shard_{shard:02d}").mkdir(parents=True, exist_ok=True)
    results = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=20) as pool:
        futures = [pool.submit(ocr.prepare_image, task, RUN_DIR, SHARDS) for task in tasks]
        for index, future in enumerate(concurrent.futures.as_completed(futures), 1):
            results.append(future.result())
            if index % 50 == 0 or index == len(futures):
                print(f"downloaded={index}/{len(futures)}", flush=True)
    results.sort(key=lambda row: row["product_id"])
    manifest = {
        "metadata": {
            "run_name": RUN_NAME,
            "candidate_count": len(tasks),
            "pool_sizes": pool_sizes,
            "sample_limits": limits,
            "download_success": sum(
                row["download_status"] == "SUCCESS" for row in results
            ),
            "download_error": sum(
                row["download_status"] == "ERROR" for row in results
            ),
            "shards": SHARDS,
            "db_written": False,
        },
        "products": results,
    }
    (RUN_DIR / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(manifest["metadata"], ensure_ascii=False, indent=2))


def analyze() -> None:
    manifest = json.loads((RUN_DIR / "manifest.json").read_text(encoding="utf-8"))
    ocr_rows: dict[str, dict[str, Any]] = {}
    for shard in range(SHARDS):
        path = RUN_DIR / f"ocr_{shard:02d}.json"
        rows = json.loads(path.read_text(encoding="utf-8-sig"))
        if isinstance(rows, dict):
            rows = [rows]
        for row in rows:
            ocr_rows[Path(row.get("file") or "").stem] = row

    connection = sqlite3.connect(DB_PATH)
    by_group: dict[str, Counter] = defaultdict(Counter)
    products = []
    for item in manifest["products"]:
        product_id = item["product_id"]
        group = item["research_group"]
        counts = by_group[group]
        counts["sample"] += 1
        if item["download_status"] != "SUCCESS":
            counts["download_error"] += 1
            products.append(
                {
                    "product_id": product_id,
                    "group": group,
                    "category": item["category"],
                    "download_status": item["download_status"],
                    "error": item.get("error", ""),
                }
            )
            continue
        counts["download_success"] += 1
        ocr_row = ocr_rows.get(product_id) or {}
        if ocr_row.get("status") == "SUCCESS":
            counts["ocr_success"] += 1
        raw_text = str(ocr_row.get("text") or "")
        if raw_text.strip():
            counts["ocr_nonempty"] += 1
        dimension_text = ocr.ocr_dimension_text(raw_text)
        if dimension_text:
            counts["new_dimension_signal"] += 1

        db_row = connection.execute(
            "SELECT goods_blob,html_blob,qna_blob,ocr_blob FROM sources WHERE product_id=?",
            (product_id,),
        ).fetchone()
        data = (unpack(db_row[0]) or {}).get("data") or {}
        product = {
            "data": data,
            "html": unpack(db_row[1]) or {},
            "qna": unpack(db_row[2]) or {},
            "ocr": unpack(db_row[3]) or {},
        }
        sources = product_sources(product)
        if dimension_text:
            sources.append(("연구용 추가 이미지 OCR", dimension_text))
        records = workbook.dimension_records(sources)
        complete = any(
            row.get("w_mm") is not None
            and row.get("d_mm") is not None
            and row.get("h_mm") is not None
            for row in records
        )
        partial = any(
            sum(row.get(axis) is not None for axis in ("w_mm", "d_mm", "h_mm"))
            >= 1
            for row in records
        )
        if complete:
            counts["would_become_complete"] += 1
        elif partial:
            counts["still_partial"] += 1
        else:
            counts["still_missing"] += 1
        products.append(
            {
                "product_id": product_id,
                "group": group,
                "category": item["category"],
                "download_status": item["download_status"],
                "ocr_status": ocr_row.get("status"),
                "dimension_signal": bool(dimension_text),
                "would_become_complete": complete,
                "selected_url": (item.get("selected") or {}).get("url", ""),
                # 연구 결과 재검증 시 뒤쪽의 실제 규격 구문이 잘리지 않도록
                # 전체 추출문을 보존한다.
                "dimension_text": dimension_text,
            }
        )
    connection.close()
    result = {
        "metadata": manifest["metadata"],
        "by_group": {key: dict(value) for key, value in by_group.items()},
        "products": products,
        "notes": [
            "연구 표본 결과이며 sources DB와 고객용 엑셀에는 병합하지 않았다.",
            "would_become_complete는 기존 API·HTML·FAQ/Q&A·OCR에 이번 추가 OCR을 합친 결과다.",
            "L은 이번에도 직접 W/D/H로 변환하지 않았다.",
        ],
    }
    OUTPUT_PATH.write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(result["by_group"], ensure_ascii=False, indent=2))
    print(f"output={OUTPUT_PATH}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("phase", choices=("prepare", "analyze"))
    args = parser.parse_args()
    if args.phase == "prepare":
        prepare()
    else:
        analyze()


if __name__ == "__main__":
    main()
