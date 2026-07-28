from __future__ import annotations

import argparse
import concurrent.futures
import io
import itertools
import json
import re
import sqlite3
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

from PIL import Image, ImageFile

import build_homestyle_bulk_workbook as workbook
import bulk_homestyle_ocr as ocr
from analyze_dimension_notations import product_sources
from bulk_homestyle_collect import DB_PATH, request_bytes, unpack


RUN_NAME = "partial_nonflat_sequential"
RUN_DIR = ocr.OCR_ROOT / RUN_NAME
OUTPUT = ocr.RUN_DIR / "partial_nonflat_sequential_ocr.json"
SHARDS = 8
MAX_CANDIDATES = 3
FLAT_CATEGORIES = {"러그", "액자", "인테리어포스터"}
ImageFile.LOAD_TRUNCATED_IMAGES = True
Image.MAX_IMAGE_PIXELS = None


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def candidate_confidence(text: str) -> str:
    normalized = str(text or "").replace("×", "x")
    axes = {
        axis
        for axis, patterns in {
            "W": (r"(?<![A-Z])W\s*[:=]?\s*\d", r"(?:가로|너비|폭)\s*[:=]?\s*\d"),
            "D": (r"(?<![A-Z])D\s*[:=]?\s*\d", r"(?:깊이|세로)\s*[:=]?\s*\d"),
            "H": (r"(?<![A-Z])H\s*[:=]?\s*\d", r"높이\s*[:=]?\s*\d"),
        }.items()
        if any(re.search(pattern, normalized, re.I) for pattern in patterns)
    }
    if axes == {"W", "D", "H"}:
        return "HIGH_EXPLICIT_AXES"
    if re.search(r"(?:DIA\.?|Ø|Φ|⌀|지름|직경)", normalized, re.I) and "H" in axes:
        return "HIGH_DIAMETER_HEIGHT"
    return "MEDIUM_UNLABELED_OR_OCR"


def record_has_matching_explicit_axes(record: dict[str, Any]) -> bool:
    """Check that this exact W/D/H record is explicitly labelled in its raw text."""
    text = str(record.get("raw") or "")
    matches = list(
        re.finditer(
            r"(?<![A-Z])\(?\s*([WDH])\s*\)?\s*[:=]?\s*\(?\s*"
            r"(\d[\d,.]*)\s*\)?\s*\(?\s*(mm|cm)?\s*\)?",
            text,
            re.I,
        )
    )
    for start_index, first in enumerate(matches):
        cluster = []
        for match in matches[start_index:]:
            if match.end() - first.start() > 180:
                break
            cluster.append(match)
        axes = {match.group(1).upper() for match in cluster}
        if axes != {"W", "D", "H"}:
            continue
        raw_values = [float(match.group(2).replace(",", "")) for match in cluster]
        common_unit = next(
            (match.group(3) for match in reversed(cluster) if match.group(3)),
            "cm" if max(raw_values) <= 300 else "mm",
        )
        mapped: dict[str, float] = {}
        for match, value in zip(cluster, raw_values):
            axis = match.group(1).upper()
            unit = match.group(3) or common_unit
            mapped.setdefault(axis, value * (10 if unit.casefold() == "cm" else 1))
        if all(
            abs(mapped[axis] - float(record[key])) <= max(2.0, float(record[key]) * 0.005)
            for axis, key in (("W", "w_mm"), ("D", "d_mm"), ("H", "h_mm"))
        ):
            return True
    return False


def matched_complete_record(
    records: list[dict[str, Any]],
    current: dict[str, float | None],
) -> dict[str, Any] | None:
    """Return a complete OCR record only when every known baseline value agrees.

    The baseline may be a generic two-value expression.  For example, a shelf
    title can expose W x H while the old parser stored it as W x D.  Therefore
    the comparison is value-based rather than requiring the same axis name.
    Candidate axes are still matched one-to-one, so one number cannot satisfy
    two known values.
    """
    known = [(axis, float(value)) for axis, value in current.items() if value is not None]
    if len(known) < 2:
        return None
    candidates: list[tuple[tuple[float, float, int], dict[str, Any]]] = []
    for record in records:
        values = [
            (axis, float(record[axis]))
            for axis in ("w_mm", "d_mm", "h_mm")
            if record.get(axis) is not None
        ]
        if len(values) != 3:
            continue
        best_assignment: tuple[float, float, int] | None = None
        for assigned in itertools.permutations(values, len(known)):
            deltas = [abs(value - assigned[index][1]) for index, (_, value) in enumerate(known)]
            tolerances = [max(15.0, value * 0.03) for _, value in known]
            if not all(delta <= tolerance for delta, tolerance in zip(deltas, tolerances)):
                continue
            normalized = [delta / tolerance for delta, tolerance in zip(deltas, tolerances)]
            same_axis = sum(
                known[index][0] == assigned[index][0]
                for index in range(len(known))
            )
            score = (max(normalized, default=0.0), sum(normalized), -same_axis)
            if best_assignment is None or score < best_assignment:
                best_assignment = score
        if best_assignment is not None:
            same_axis = -best_assignment[2]
            if same_axis < len(known) and not record_has_matching_explicit_axes(record):
                continue
            candidates.append((best_assignment, record))
    if not candidates:
        return None
    candidates.sort(key=lambda item: item[0])
    # A lineup/specification image can contain several variants that share the
    # same two known axes.  If two different complete records have the exact
    # same best score, the product/option identity is ambiguous and must not be
    # auto-selected (for example, a 5-tier and 6-tier desk in one chart).
    if len(candidates) > 1 and candidates[0][0] == candidates[1][0]:
        return None
    return candidates[0][1]


def selected_candidates(data: dict[str, Any], old_ocr: dict[str, Any], status: int) -> list[dict[str, Any]]:
    images = ocr.detail_images(str(data.get("detailInfo") or ""))
    selected = old_ocr.get("selected") or {}
    old_url = ocr.normalize_url(str(selected.get("url") or ""))
    normal = [item for item in images if float(item.get("score") or 0) < 100]
    ranked = normal or images
    result = []
    seen = set()
    if status == 502 and old_url:
        result.append({**selected, "url": old_url, "reason": "URL_RETRY"})
        seen.add(old_url)
    for item in ranked:
        if item["url"] == old_url or item["url"] in seen:
            continue
        seen.add(item["url"])
        result.append({**item, "reason": "NEXT_RANKED_IMAGE"})
        if len(result) >= MAX_CANDIDATES:
            break
    return result[:MAX_CANDIDATES]


def download_task(task: dict[str, Any]) -> dict[str, Any]:
    product_id = task["product_id"]
    order = task["attempt_order"]
    shard = int(product_id[-4:]) % SHARDS
    destination = RUN_DIR / f"shard_{shard:02d}" / f"{product_id}__{order:02d}.jpg"
    try:
        status, body = request_bytes(task["image"]["url"], timeout=45, attempts=3)
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
            "file": str(destination),
            "width": width,
            "height": height,
        }
    except Exception as exc:
        return {
            **task,
            "download_status": "ERROR",
            "file": "",
            "error": f"{type(exc).__name__}: {exc}",
        }


def prepare() -> None:
    connection = sqlite3.connect(DB_PATH)
    targets = connection.execute(
        """
        SELECT product_id,product_name,small_category,ocr_status
        FROM stg_dimension_reinforcement
        WHERE is_current=1 AND current_status='부분확보'
          AND small_category NOT IN ('러그','액자','인테리어포스터')
        ORDER BY product_id
        """
    ).fetchall()
    tasks = []
    target_without_candidate = []
    for product_id, product_name, category, staged_ocr_status in targets:
        row = connection.execute(
            "SELECT goods_blob,ocr_status,ocr_blob FROM sources WHERE product_id=?",
            (product_id,),
        ).fetchone()
        data = (unpack(row[0]) or {}).get("data") or {}
        status = int(row[1] or staged_ocr_status or 0)
        old_ocr = unpack(row[2]) or {}
        candidates = selected_candidates(data, old_ocr, status)
        if not candidates:
            target_without_candidate.append(product_id)
        for order, image in enumerate(candidates, 1):
            tasks.append(
                {
                    "product_id": product_id,
                    "product_name": product_name,
                    "small_category": category,
                    "attempt_order": order,
                    "image": image,
                }
            )
    connection.close()
    RUN_DIR.mkdir(parents=True, exist_ok=True)
    for shard in range(SHARDS):
        (RUN_DIR / f"shard_{shard:02d}").mkdir(parents=True, exist_ok=True)
    results = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=20) as pool:
        futures = [pool.submit(download_task, task) for task in tasks]
        for index, future in enumerate(concurrent.futures.as_completed(futures), 1):
            results.append(future.result())
            if index % 50 == 0 or index == len(futures):
                print(f"downloaded={index}/{len(futures)}", flush=True)
    results.sort(key=lambda row: (row["product_id"], row["attempt_order"]))
    manifest = {
        "metadata": {
            "run_name": RUN_NAME,
            "created_at": now_text(),
            "target_products": len(targets),
            "candidate_images": len(tasks),
            "download_success": sum(row["download_status"] == "SUCCESS" for row in results),
            "download_error": sum(row["download_status"] == "ERROR" for row in results),
            "target_without_candidate": len(target_without_candidate),
            "target_without_candidate_ids": target_without_candidate,
            "max_candidates_per_product": MAX_CANDIDATES,
            "shards": SHARDS,
            "sources_db_written": False,
        },
        "products": results,
    }
    (RUN_DIR / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(manifest["metadata"], ensure_ascii=False, indent=2))


def read_ocr_rows() -> dict[str, dict[str, Any]]:
    result = {}
    for shard in range(SHARDS):
        path = RUN_DIR / f"ocr_{shard:02d}.json"
        rows = json.loads(path.read_text(encoding="utf-8-sig"))
        if isinstance(rows, dict):
            rows = [rows]
        for row in rows:
            result[Path(row.get("file") or "").stem] = row
    return result


def create_result_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS stg_dimension_ocr_attempt (
            run_name TEXT NOT NULL,
            created_at TEXT NOT NULL,
            product_id TEXT NOT NULL,
            attempt_order INTEGER NOT NULL,
            image_url TEXT,
            image_reason TEXT,
            download_status TEXT,
            ocr_status TEXT,
            dimension_signal INTEGER NOT NULL DEFAULT 0,
            candidate_dimensions_json TEXT,
            used_in_sequence INTEGER NOT NULL DEFAULT 0,
            cumulative_complete INTEGER NOT NULL DEFAULT 0,
            raw_dimension_text TEXT,
            error TEXT,
            PRIMARY KEY(run_name,product_id,attempt_order)
        );
        CREATE TABLE IF NOT EXISTS stg_dimension_ocr_result (
            run_name TEXT NOT NULL,
            created_at TEXT NOT NULL,
            product_id TEXT NOT NULL,
            product_name TEXT,
            small_category TEXT,
            baseline_status TEXT NOT NULL,
            final_status TEXT NOT NULL,
            recovered INTEGER NOT NULL DEFAULT 0,
            confidence TEXT,
            w_mm REAL,
            d_mm REAL,
            h_mm REAL,
            attempts_used INTEGER NOT NULL DEFAULT 0,
            dimension_signal_images INTEGER NOT NULL DEFAULT 0,
            stop_reason TEXT NOT NULL,
            evidence_image_url TEXT,
            evidence_text TEXT,
            applied_to_sources INTEGER NOT NULL DEFAULT 0,
            PRIMARY KEY(run_name,product_id)
        );
        """
    )


def analyze() -> dict[str, Any]:
    manifest = json.loads((RUN_DIR / "manifest.json").read_text(encoding="utf-8"))
    ocr_rows = read_ocr_rows()
    by_product: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for item in manifest["products"]:
        by_product[item["product_id"]].append(item)
    connection = sqlite3.connect(DB_PATH)
    created_at = now_text()
    attempt_inserts = []
    result_inserts = []
    summary = Counter()
    target_rows = connection.execute(
        """
        SELECT product_id,product_name,small_category,
               current_w_mm,current_d_mm,current_h_mm
        FROM stg_dimension_reinforcement
        WHERE is_current=1 AND current_status='부분확보'
          AND small_category NOT IN ('러그','액자','인테리어포스터')
        ORDER BY product_id
        """
    ).fetchall()
    for product_id, product_name, category, current_w, current_d, current_h in target_rows:
        db_row = connection.execute(
            "SELECT goods_blob,html_blob,qna_blob,ocr_blob FROM sources WHERE product_id=?",
            (product_id,),
        ).fetchone()
        product = {
            "data": (unpack(db_row[0]) or {}).get("data") or {},
            "html": unpack(db_row[1]) or {},
            "qna": unpack(db_row[2]) or {},
            "ocr": unpack(db_row[3]) or {},
        }
        sources = product_sources(product)
        current_dimensions = {
            "w_mm": current_w,
            "d_mm": current_d,
            "h_mm": current_h,
        }
        attempts = sorted(by_product.get(product_id, []), key=lambda row: row["attempt_order"])
        recovered = False
        selected_record = None
        confidence = None
        evidence_url = None
        evidence_text = None
        attempts_used = 0
        signal_images = 0
        unmatched_complete = False
        stop_reason = "NO_VALID_IMAGE" if not attempts else "MAX_CANDIDATES_REACHED"
        for item in attempts:
            order = int(item["attempt_order"])
            stem = f"{product_id}__{order:02d}"
            row = ocr_rows.get(stem) or {}
            download_ok = item["download_status"] == "SUCCESS"
            ocr_ok = row.get("status") == "SUCCESS"
            raw_text = str(row.get("text") or "") if download_ok else ""
            dimension_text = ocr.ocr_dimension_text(raw_text) if ocr_ok else ""
            candidate_records = workbook.dimension_records(
                [(f"추가 이미지 OCR {order}", dimension_text)]
            ) if dimension_text else []
            used = int(not recovered)
            cumulative_complete = 0
            if not recovered:
                attempts_used += 1
                if dimension_text:
                    signal_images += 1
                complete_records = [
                    record for record in candidate_records
                    if all(record.get(axis) is not None for axis in ("w_mm", "d_mm", "h_mm"))
                ]
                matched_record = matched_complete_record(
                    complete_records, current_dimensions
                )
                if complete_records and not matched_record:
                    unmatched_complete = True
                if matched_record:
                    recovered = True
                    cumulative_complete = 1
                    selected_record = matched_record
                    confidence = "HIGH_KNOWN_AXES_MATCH"
                    evidence_url = item["image"]["url"]
                    evidence_text = dimension_text
                    stop_reason = "COMPLETE_WDH"
            attempt_inserts.append(
                (
                    RUN_NAME, created_at, product_id, order,
                    item["image"]["url"], item["image"].get("reason"),
                    item["download_status"], row.get("status") or "NOT_RUN",
                    int(bool(dimension_text)),
                    json.dumps(candidate_records, ensure_ascii=False), used,
                    cumulative_complete, dimension_text,
                    item.get("error") or row.get("error") or "",
                )
            )
        if recovered and selected_record:
            final_status = "확보"
            w_mm, d_mm, h_mm = (
                selected_record["w_mm"], selected_record["d_mm"], selected_record["h_mm"]
            )
            summary["recovered"] += 1
            summary[f"confidence:{confidence}"] += 1
        else:
            final_status = "부분확보"
            w_mm = d_mm = h_mm = None
            summary["not_recovered"] += 1
            if unmatched_complete:
                confidence = "REVIEW_COMPLETE_UNMATCHED"
                stop_reason = "COMPLETE_UNMATCHED_REVIEW"
        summary[f"stop:{stop_reason}"] += 1
        result_inserts.append(
            (
                RUN_NAME, created_at, product_id, product_name, category,
                "부분확보", final_status, int(recovered), confidence,
                w_mm, d_mm, h_mm, attempts_used, signal_images, stop_reason,
                evidence_url, evidence_text, 0,
            )
        )
    create_result_schema(connection)
    with connection:
        connection.execute(
            "DELETE FROM stg_dimension_ocr_attempt WHERE run_name=?", (RUN_NAME,)
        )
        connection.execute(
            "DELETE FROM stg_dimension_ocr_result WHERE run_name=?", (RUN_NAME,)
        )
        connection.executemany(
            "INSERT INTO stg_dimension_ocr_attempt VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            attempt_inserts,
        )
        connection.executemany(
            "INSERT INTO stg_dimension_ocr_result VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            result_inserts,
        )
        connection.execute(
            "UPDATE stg_dimension_reinforcement SET review_status='추가OCR검토완료' "
            "WHERE is_current=1 AND product_id IN "
            "(SELECT product_id FROM stg_dimension_ocr_result WHERE run_name=?)",
            (RUN_NAME,),
        )
    result = {
        "run_name": RUN_NAME,
        "created_at": created_at,
        "target_products": len(target_rows),
        "attempt_rows": len(attempt_inserts),
        "summary": dict(summary),
        "sources_db_written": False,
        "staging_tables_written": [
            "stg_dimension_ocr_attempt", "stg_dimension_ocr_result"
        ],
        "notes": [
            "추가 OCR 결과는 연구용 스테이징에만 저장했다.",
            "HIGH 명시축 후보를 원문 검증한 뒤 sources와 필수값 스냅샷을 갱신한다.",
        ],
    }
    connection.close()
    OUTPUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return result


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("phase", choices=("prepare", "run", "analyze", "all"))
    args = parser.parse_args()
    if args.phase in ("prepare", "all"):
        prepare()
    if args.phase in ("run", "all"):
        ocr.run_ocr(RUN_NAME, SHARDS)
    if args.phase in ("analyze", "all"):
        analyze()


if __name__ == "__main__":
    main()
