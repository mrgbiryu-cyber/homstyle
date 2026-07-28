from __future__ import annotations

import argparse
import concurrent.futures
import gzip
import html as html_module
import json
import re
import sqlite3
import ssl
import time
import urllib.error
import urllib.parse
import urllib.request
import zlib
from datetime import datetime
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parent
RUN_DIR = ROOT / "homestyle_bulk_run"
DB_PATH = RUN_DIR / "homestyle_bulk.sqlite"
SCOPE_PATH = ROOT / "poc_full_run" / "source_scope_cache.json"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 Chrome/150 Safari/537.36"
)
SSL_CONTEXT = ssl._create_unverified_context()
PLP_ENDPOINT = "https://livingapi.lge.co.kr/displaysvc/ajax/v1/shop/goods"
GOODS_ENDPOINT = "https://livingapi.lge.co.kr/itemsvc/ajax/v1/pdp/goods/{product_id}"
SPACE_ENDPOINT = (
    "https://livingapi.lge.co.kr/displaysvc/ajax/v1/collection/"
    "pdp/space-recommendation"
)
PACKAGES_ENDPOINT = "https://livingapi.lge.co.kr/itemsvc/ajax/v1/goods/packages"
QNA_ENDPOINT = "https://livingapi.lge.co.kr/itemsvc/ajax/v1/pdp/qna-list"
PDP_ENDPOINT = "https://homestyle.lge.co.kr/item"


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def pack(value: Any) -> bytes:
    return zlib.compress(
        json.dumps(value, ensure_ascii=False, separators=(",", ":")).encode("utf-8"),
        level=6,
    )


def unpack(value: bytes | None) -> Any:
    if not value:
        return None
    return json.loads(zlib.decompress(value).decode("utf-8"))


def connect_db() -> sqlite3.Connection:
    RUN_DIR.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(DB_PATH)
    connection.execute("PRAGMA journal_mode=WAL")
    connection.execute("PRAGMA synchronous=NORMAL")
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS categories (
            scope_id TEXT PRIMARY KEY,
            large_name TEXT,
            mid_name TEXT,
            small_name TEXT,
            category_id TEXT,
            source_count INTEGER,
            live_count INTEGER,
            http_status INTEGER,
            error TEXT,
            collected_at TEXT
        );
        CREATE TABLE IF NOT EXISTS products (
            product_id TEXT PRIMARY KEY,
            category_scope_ids TEXT NOT NULL,
            listing_blob BLOB,
            inventoried_at TEXT
        );
        CREATE TABLE IF NOT EXISTS sources (
            product_id TEXT PRIMARY KEY,
            goods_status INTEGER,
            goods_blob BLOB,
            space_status INTEGER,
            space_blob BLOB,
            packages_status INTEGER,
            packages_blob BLOB,
            qna_status INTEGER,
            qna_blob BLOB,
            html_status INTEGER,
            html_blob BLOB,
            ocr_status INTEGER,
            ocr_blob BLOB,
            structured_error TEXT,
            html_error TEXT,
            structured_at TEXT,
            html_at TEXT,
            ocr_at TEXT,
            FOREIGN KEY(product_id) REFERENCES products(product_id)
        );
        CREATE INDEX IF NOT EXISTS idx_sources_goods_status ON sources(goods_status);
        CREATE INDEX IF NOT EXISTS idx_sources_html_status ON sources(html_status);
        """
    )
    existing_columns = {
        row[1] for row in connection.execute("PRAGMA table_info(sources)")
    }
    for column_name, column_type in (
        ("ocr_status", "INTEGER"),
        ("ocr_blob", "BLOB"),
        ("ocr_at", "TEXT"),
    ):
        if column_name not in existing_columns:
            connection.execute(
                f"ALTER TABLE sources ADD COLUMN {column_name} {column_type}"
            )
    return connection


def request_bytes(
    url: str,
    *,
    params: dict[str, Any] | None = None,
    referer: str = "https://homestyle.lge.co.kr/",
    timeout: int = 60,
    attempts: int = 4,
) -> tuple[int, bytes]:
    if params:
        query = urllib.parse.urlencode(
            {key: value for key, value in params.items() if value not in (None, "")},
            doseq=True,
        )
        url += ("&" if "?" in url else "?") + query
    headers = {
        "User-Agent": USER_AGENT,
        "Accept": "*/*",
        "Accept-Encoding": "gzip, deflate",
        "Referer": referer,
    }
    last_error: Exception | None = None
    for attempt in range(attempts):
        try:
            request = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(
                request, timeout=timeout, context=SSL_CONTEXT
            ) as response:
                body = response.read()
                encoding = (response.headers.get("Content-Encoding") or "").casefold()
                if encoding == "gzip":
                    body = gzip.decompress(body)
                elif encoding == "deflate":
                    body = zlib.decompress(body)
                return int(response.status), body
        except urllib.error.HTTPError as exc:
            last_error = exc
            if exc.code < 500 and exc.code != 429:
                return exc.code, exc.read()
        except Exception as exc:  # noqa: BLE001 - network retry boundary
            last_error = exc
        if attempt + 1 < attempts:
            time.sleep(min(8, 1.5 * (2**attempt)))
    raise RuntimeError(f"request failed: {url}: {last_error!r}")


def request_json(
    url: str,
    *,
    params: dict[str, Any] | None = None,
    referer: str = "https://homestyle.lge.co.kr/",
    timeout: int = 60,
    attempts: int = 4,
) -> tuple[int, dict[str, Any]]:
    status, body = request_bytes(
        url,
        params=params,
        referer=referer,
        timeout=timeout,
        attempts=attempts,
    )
    try:
        payload = json.loads(body.decode("utf-8", errors="replace"))
    except json.JSONDecodeError:
        payload = {"status": status, "message": "NON_JSON", "raw": body[:1000].decode("utf-8", "replace")}
    return status, payload if isinstance(payload, dict) else {"data": payload}


def load_home_categories() -> list[dict[str, Any]]:
    cached = json.loads(SCOPE_PATH.read_text(encoding="utf-8"))
    return [
        row
        for row in cached["selected"]
        if row.get("source_system") == "HOMESTYLE"
    ]


def fetch_category(category: dict[str, Any]) -> dict[str, Any]:
    page_size = 2000
    params = {
        "filterSelType": "SHOP",
        "pageNum": 1,
        "pageSize": page_size,
        "category3": category["category_id"],
    }
    try:
        status, payload = request_json(PLP_ENDPOINT, params=params, timeout=120)
    except Exception:
        page_size = 500
        params["pageSize"] = page_size
        status, payload = request_json(PLP_ENDPOINT, params=params, timeout=120)
    data = payload.get("data") or {}
    total = int(data.get("totalCount") or 0)
    rows = list(data.get("list") or [])
    pages = (total + page_size - 1) // page_size
    for page in range(2, pages + 1):
        page_status, page_payload = request_json(
            PLP_ENDPOINT,
            params={**params, "pageNum": page},
            timeout=120,
        )
        status = page_status if page_status != 200 else status
        rows.extend((page_payload.get("data") or {}).get("list") or [])
    return {
        "category": category,
        "status": status,
        "total": total,
        "rows": rows,
        "error": "" if status == 200 else str(payload.get("message") or "HTTP error"),
    }


def collect_inventory(workers: int) -> None:
    categories = load_home_categories()
    started = time.time()
    results: list[dict[str, Any]] = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {executor.submit(fetch_category, category): category for category in categories}
        for index, future in enumerate(concurrent.futures.as_completed(futures), start=1):
            category = futures[future]
            try:
                results.append(future.result())
            except Exception as exc:  # noqa: BLE001 - record category failure
                results.append(
                    {
                        "category": category,
                        "status": 0,
                        "total": 0,
                        "rows": [],
                        "error": repr(exc),
                    }
                )
            if index % 10 == 0 or index == len(categories):
                print(
                    f"inventory {index}/{len(categories)} "
                    f"elapsed={time.time() - started:.1f}s",
                    flush=True,
                )

    memberships: dict[str, set[str]] = {}
    listings: dict[str, dict[str, Any]] = {}
    connection = connect_db()
    collected_at = now_text()
    with connection:
        for result in results:
            category = result["category"]
            connection.execute(
                """
                INSERT INTO categories (
                    scope_id, large_name, mid_name, small_name, category_id,
                    source_count, live_count, http_status, error, collected_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(scope_id) DO UPDATE SET
                    large_name=excluded.large_name,
                    mid_name=excluded.mid_name,
                    small_name=excluded.small_name,
                    category_id=excluded.category_id,
                    source_count=excluded.source_count,
                    live_count=excluded.live_count,
                    http_status=excluded.http_status,
                    error=excluded.error,
                    collected_at=excluded.collected_at
                """,
                (
                    category["scope_id"],
                    category.get("large", ""),
                    category.get("mid", ""),
                    category.get("small", ""),
                    category.get("category_id", ""),
                    int(category.get("public_count") or 0),
                    int(result["total"]),
                    int(result["status"] or 0),
                    result["error"],
                    collected_at,
                ),
            )
            for item in result["rows"]:
                product_id = str(item.get("productId") or item.get("goodsNo") or "").strip()
                if not product_id:
                    continue
                memberships.setdefault(product_id, set()).add(category["scope_id"])
                listings.setdefault(product_id, item)

        for product_id in sorted(memberships):
            connection.execute(
                """
                INSERT INTO products (product_id, category_scope_ids, listing_blob, inventoried_at)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(product_id) DO UPDATE SET
                    category_scope_ids=excluded.category_scope_ids,
                    listing_blob=excluded.listing_blob,
                    inventoried_at=excluded.inventoried_at
                """,
                (
                    product_id,
                    json.dumps(sorted(memberships[product_id]), ensure_ascii=False),
                    pack(listings[product_id]),
                    collected_at,
                ),
            )
            connection.execute(
                "INSERT OR IGNORE INTO sources (product_id) VALUES (?)",
                (product_id,),
            )
    connection.close()

    failed = sum(result["status"] != 200 for result in results)
    total_rows = sum(len(result["rows"]) for result in results)
    duplicates = total_rows - len(memberships)
    print(
        f"inventory_complete categories={len(categories)} failed={failed} "
        f"category_rows={total_rows} unique_products={len(memberships)} "
        f"duplicate_memberships={duplicates}",
        flush=True,
    )


def qna_records(product_id: str, referer: str) -> tuple[int, dict[str, Any]]:
    base_params = {
        "goodsId": product_id,
        "qnaType": "",
        "isMyInquiry": "false",
        "excludeSecret": "true",
        "pageSize": 100,
    }
    status, payload = request_json(
        QNA_ENDPOINT,
        params={**base_params, "pageNum": 1},
        referer=referer,
    )
    first = ((payload.get("data") or {}).get("qnaListData") or {})
    pages = int(first.get("pages") or 0)
    rows = list(first.get("list") or [])
    for page in range(2, pages + 1):
        page_status, page_payload = request_json(
            QNA_ENDPOINT,
            params={**base_params, "pageNum": page},
            referer=referer,
        )
        if page_status != 200:
            status = page_status
        rows.extend(
            ((page_payload.get("data") or {}).get("qnaListData") or {}).get("list")
            or []
        )

    public_rows = [row for row in rows if not row.get("isSecret") and not row.get("isNotice")]
    records: list[dict[str, Any]] = []
    # The list payload already contains the public title/answer summary. Fetch a
    # small number of detail records per product to recover full question text
    # without turning a high-Q&A product into hundreds of extra requests.
    detail_fetch_limit = 3
    for row_index, row in enumerate(public_rows):
        inquiry_id = str(row.get("inquiryId") or "")
        detail: dict[str, Any] = {}
        if inquiry_id and row_index < detail_fetch_limit:
            try:
                _, detail_payload = request_json(
                    f"{QNA_ENDPOINT}/{inquiry_id}",
                    params={"isNotice": "false"},
                    referer=referer,
                )
                detail = detail_payload.get("data") or {}
            except Exception:
                detail = {}
        records.append(
            {
                "inquiry_id": inquiry_id,
                "qna_type": detail.get("qnaTypeName") or row.get("qnaTypeName") or "",
                "registered_date": detail.get("registeredDate") or row.get("registeredDate") or "",
                "question_title": detail.get("inquiryTitle") or row.get("inquiryTitle") or "",
                "question_text": detail.get("inquiryContent") or "",
                "answer_text": detail.get("answerContent") or row.get("answerContent") or "",
            }
        )
    return status, {
        "reported_public_total": int(first.get("total") or 0),
        "pages": pages,
        "records": records,
        "privacy_filter": "excludeSecret=true; isMyInquiry=false; author/files omitted",
    }


def collect_structured_one(product_id: str) -> dict[str, Any]:
    referer = f"https://homestyle.lge.co.kr/item?productId={product_id}"
    result: dict[str, Any] = {"product_id": product_id, "error": ""}
    errors = []
    endpoints = (
        (
            "goods",
            GOODS_ENDPOINT.format(product_id=product_id),
            {"epFlagYn": "N"},
        ),
        ("space", SPACE_ENDPOINT, {"goodsId": product_id}),
        ("packages", PACKAGES_ENDPOINT, {"goodsId": product_id}),
    )
    for name, endpoint, params in endpoints:
        try:
            status, payload = request_json(endpoint, params=params, referer=referer)
        except Exception as exc:  # noqa: BLE001 - per-source recovery
            status, payload = 0, {"error": repr(exc)}
            errors.append(f"{name}={exc!r}")
        result[f"{name}_status"] = status
        result[f"{name}_blob"] = pack(payload)
    result["error"] = " | ".join(errors)
    return result


def collect_structured(workers: int, retry_errors: bool) -> None:
    connection = connect_db()
    if retry_errors:
        where = (
            "goods_status IS NULL OR goods_status != 200 OR "
            "space_status IS NULL OR space_status != 200 OR "
            "packages_status IS NULL OR packages_status != 200"
        )
    else:
        where = "goods_status IS NULL"
    product_ids = [
        row[0]
        for row in connection.execute(
            f"SELECT product_id FROM sources WHERE {where} ORDER BY product_id"
        )
    ]
    print(f"structured_pending={len(product_ids)} workers={workers}", flush=True)
    if not product_ids:
        connection.close()
        return

    started = time.time()
    complete = 0
    failed = 0
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {
            executor.submit(collect_structured_one, product_id): product_id
            for product_id in product_ids
        }
        for future in concurrent.futures.as_completed(futures):
            product_id = futures[future]
            try:
                result = future.result()
            except Exception as exc:  # noqa: BLE001 - persist worker crash
                result = {
                    "product_id": product_id,
                    "goods_status": 0,
                    "goods_blob": pack({"error": repr(exc)}),
                    "space_status": 0,
                    "space_blob": pack({"error": repr(exc)}),
                    "packages_status": 0,
                    "packages_blob": pack({"error": repr(exc)}),
                    "error": repr(exc),
                }
            connection.execute(
                """
                UPDATE sources SET
                    goods_status=?, goods_blob=?,
                    space_status=?, space_blob=?,
                    packages_status=?, packages_blob=?,
                    structured_error=?, structured_at=?
                WHERE product_id=?
                """,
                (
                    result["goods_status"],
                    result["goods_blob"],
                    result["space_status"],
                    result["space_blob"],
                    result["packages_status"],
                    result["packages_blob"],
                    result["error"],
                    now_text(),
                    product_id,
                ),
            )
            complete += 1
            failed += any(
                result[key] != 200
                for key in (
                    "goods_status",
                    "space_status",
                    "packages_status",
                )
            )
            if complete % 25 == 0:
                connection.commit()
            if complete % 100 == 0 or complete == len(product_ids):
                rate = complete / max(0.001, time.time() - started)
                print(
                    f"structured {complete}/{len(product_ids)} failed_rows={failed} "
                    f"rate={rate:.1f}/s elapsed={time.time() - started:.1f}s",
                    flush=True,
                )
    connection.commit()
    connection.close()


def collect_qna_one(product_id: str) -> dict[str, Any]:
    referer = f"https://homestyle.lge.co.kr/item?productId={product_id}"
    try:
        status, qna = qna_records(product_id, referer)
        return {
            "product_id": product_id,
            "status": status,
            "blob": pack(qna),
            "error": "" if status == 200 else f"HTTP {status}",
        }
    except Exception as exc:  # noqa: BLE001 - per-source recovery
        return {
            "product_id": product_id,
            "status": 0,
            "blob": pack({"error": repr(exc), "records": []}),
            "error": repr(exc),
        }


def collect_qna(workers: int, retry_errors: bool) -> None:
    connection = connect_db()
    where = "qna_status IS NULL OR qna_status != 200" if retry_errors else "qna_status IS NULL"
    product_ids = [
        row[0]
        for row in connection.execute(
            f"SELECT product_id FROM sources WHERE {where} ORDER BY product_id"
        )
    ]
    print(f"qna_pending={len(product_ids)} workers={workers}", flush=True)
    if not product_ids:
        connection.close()
        return
    started = time.time()
    complete = 0
    failed = 0
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {
            executor.submit(collect_qna_one, product_id): product_id
            for product_id in product_ids
        }
        for future in concurrent.futures.as_completed(futures):
            result = future.result()
            connection.execute(
                """
                UPDATE sources SET qna_status=?, qna_blob=?, structured_error=
                    CASE WHEN ?='' THEN structured_error ELSE TRIM(COALESCE(structured_error,'') || ' | ' || ?) END,
                    structured_at=?
                WHERE product_id=?
                """,
                (
                    result["status"],
                    result["blob"],
                    result["error"],
                    result["error"],
                    now_text(),
                    result["product_id"],
                ),
            )
            complete += 1
            failed += result["status"] != 200
            if complete % 25 == 0:
                connection.commit()
            if complete % 100 == 0 or complete == len(product_ids):
                rate = complete / max(0.001, time.time() - started)
                print(
                    f"qna {complete}/{len(product_ids)} failed_rows={failed} "
                    f"rate={rate:.1f}/s elapsed={time.time() - started:.1f}s",
                    flush=True,
                )
    connection.commit()
    connection.close()


def extract_jsonld(html_text: str) -> list[Any]:
    blocks = re.findall(
        r"<script[^>]+type=[\"']application/ld\+json[\"'][^>]*>(.*?)</script>",
        html_text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    result = []
    for block in blocks:
        try:
            result.append(json.loads(html_module.unescape(block.strip())))
        except json.JSONDecodeError:
            continue
    return result


def html_signals(html_text: str) -> dict[str, Any]:
    title_match = re.search(r"<title>(.*?)</title>", html_text, flags=re.I | re.S)
    meta: dict[str, str] = {}
    for tag in re.findall(r"<meta\b[^>]*>", html_text, flags=re.I):
        key_match = re.search(r"(?:name|property)=[\"']([^\"']+)", tag, flags=re.I)
        value_match = re.search(r"content=[\"']([^\"']*)", tag, flags=re.I)
        if key_match and value_match:
            meta[key_match.group(1).lower()] = html_module.unescape(value_match.group(1))

    jsonld = extract_jsonld(html_text)
    faq_records = []
    for value in jsonld:
        values = value if isinstance(value, list) else [value]
        for item in values:
            if not isinstance(item, dict) or item.get("@type") != "FAQPage":
                continue
            for entity in item.get("mainEntity") or []:
                if not isinstance(entity, dict):
                    continue
                answer = entity.get("acceptedAnswer") or {}
                faq_records.append(
                    {
                        "question": entity.get("name") or "",
                        "answer": answer.get("text") if isinstance(answer, dict) else answer,
                    }
                )

    text = re.sub(r"<script\b.*?</script>", " ", html_text, flags=re.I | re.S)
    text = re.sub(r"<style\b.*?</style>", " ", text, flags=re.I | re.S)
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"\s+", " ", html_module.unescape(text)).strip()
    dimension_signals = []
    for match in re.finditer(
        r"(?:크기|규격|사이즈|가로|세로|높이|폭|깊이|W\s*×?\s*D\s*×?\s*H)",
        text,
        flags=re.I,
    ):
        snippet = text[max(0, match.start() - 100) : match.start() + 350]
        if any(char.isdigit() for char in snippet) and snippet not in dimension_signals:
            dimension_signals.append(snippet)
        if len(dimension_signals) >= 20:
            break
    return {
        "title": re.sub(r"\s+", " ", html_module.unescape(title_match.group(1))).strip()
        if title_match
        else "",
        "description": meta.get("description") or meta.get("og:description") or "",
        "og_image": meta.get("og:image") or "",
        "jsonld_count": len(jsonld),
        "faq_records": faq_records,
        "dimension_signals": dimension_signals,
        "visible_text_length": len(text),
    }


def collect_html_one(product_id: str) -> dict[str, Any]:
    referer = "https://homestyle.lge.co.kr/"
    try:
        status, body = request_bytes(
            PDP_ENDPOINT,
            params={"productId": product_id},
            referer=referer,
            timeout=90,
        )
        text = body.decode("utf-8", errors="replace")
        return {
            "product_id": product_id,
            "status": status,
            "blob": pack(html_signals(text)),
            "error": "" if status == 200 else f"HTTP {status}",
        }
    except Exception as exc:  # noqa: BLE001 - per-source recovery
        return {
            "product_id": product_id,
            "status": 0,
            "blob": pack({"error": repr(exc)}),
            "error": repr(exc),
        }


def collect_html(workers: int, retry_errors: bool) -> None:
    connection = connect_db()
    where = "html_status IS NULL OR html_status != 200" if retry_errors else "html_status IS NULL"
    product_ids = [
        row[0]
        for row in connection.execute(
            f"SELECT product_id FROM sources WHERE {where} ORDER BY product_id"
        )
    ]
    print(f"html_pending={len(product_ids)} workers={workers}", flush=True)
    if not product_ids:
        connection.close()
        return
    started = time.time()
    complete = 0
    failed = 0
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {
            executor.submit(collect_html_one, product_id): product_id
            for product_id in product_ids
        }
        for future in concurrent.futures.as_completed(futures):
            result = future.result()
            connection.execute(
                """
                UPDATE sources SET html_status=?, html_blob=?, html_error=?, html_at=?
                WHERE product_id=?
                """,
                (
                    result["status"],
                    result["blob"],
                    result["error"],
                    now_text(),
                    result["product_id"],
                ),
            )
            complete += 1
            failed += result["status"] != 200
            if complete % 25 == 0:
                connection.commit()
            if complete % 100 == 0 or complete == len(product_ids):
                rate = complete / max(0.001, time.time() - started)
                print(
                    f"html {complete}/{len(product_ids)} failed_rows={failed} "
                    f"rate={rate:.1f}/s elapsed={time.time() - started:.1f}s",
                    flush=True,
                )
    connection.commit()
    connection.close()


def report_status() -> None:
    connection = connect_db()
    category_count = connection.execute("SELECT COUNT(*) FROM categories").fetchone()[0]
    product_count = connection.execute("SELECT COUNT(*) FROM products").fetchone()[0]
    source_stats = connection.execute(
        """
        SELECT
            SUM(goods_status=200), SUM(space_status=200),
            SUM(packages_status=200), SUM(qna_status=200),
            SUM(html_status=200), COUNT(*)
        FROM sources
        """
    ).fetchone()
    print(
        f"status db={DB_PATH} categories={category_count} products={product_count} "
        f"goods={source_stats[0] or 0}/{source_stats[5]} "
        f"space={source_stats[1] or 0}/{source_stats[5]} "
        f"packages={source_stats[2] or 0}/{source_stats[5]} "
        f"qna={source_stats[3] or 0}/{source_stats[5]} "
        f"html={source_stats[4] or 0}/{source_stats[5]}"
    )
    connection.close()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "phase",
        choices=("inventory", "structured", "qna", "html", "all", "status"),
        nargs="?",
        default="all",
    )
    parser.add_argument("--workers", type=int, default=16)
    parser.add_argument("--retry-errors", action="store_true")
    args = parser.parse_args()

    if args.phase in ("inventory", "all"):
        collect_inventory(min(args.workers, 10))
    if args.phase in ("structured", "all"):
        collect_structured(args.workers, args.retry_errors)
    if args.phase in ("qna", "all"):
        collect_qna(args.workers, args.retry_errors)
    if args.phase in ("html", "all"):
        collect_html(args.workers, args.retry_errors)
    report_status()


if __name__ == "__main__":
    main()
