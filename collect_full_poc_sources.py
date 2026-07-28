from __future__ import annotations

import hashlib
import html
import json
import re
import shutil
import ssl
import urllib.error
import urllib.parse
import urllib.request
from html.parser import HTMLParser
from pathlib import Path


ROOT = Path(__file__).resolve().parent
RUN_DIR = ROOT / "poc_full_run"
RAW_API_DIR = RUN_DIR / "raw_api"
RAW_HTML_DIR = RUN_DIR / "raw_html"
IMAGE_DIR = RUN_DIR / "images_flat"

USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/150 Safari/537.36"
SSL_CONTEXT = ssl._create_unverified_context()

PRODUCTS = [
    {"product_id": "G25070005743", "source_system": "HOMESTYLE", "url": "https://homestyle.lge.co.kr/item?productId=G25070005743"},
    {"product_id": "G25100020496", "source_system": "HOMESTYLE", "url": "https://homestyle.lge.co.kr/item?productId=G25100020496"},
    {"product_id": "G25070001871", "source_system": "HOMESTYLE", "url": "https://homestyle.lge.co.kr/item?productId=G25070001871"},
    {"product_id": "G25070006112", "source_system": "HOMESTYLE", "url": "https://homestyle.lge.co.kr/item?productId=G25070006112"},
    {"product_id": "OLED48C6KNA", "source_system": "LGE_APPLIANCE", "model_id": "MD10770851", "url": "https://www.lge.co.kr/tvs/oled48c6kna-wall"},
    {"product_id": "G646GBB031", "source_system": "LGE_APPLIANCE", "model_id": "MD10780848", "url": "https://www.lge.co.kr/refrigerators/g646gbb031"},
    {"product_id": "WA2525YMZF", "source_system": "LGE_APPLIANCE", "model_id": "MD10576829", "url": "https://www.lge.co.kr/wash-tower/wa2525ymzf"},
    {"product_id": "SQ06GJ1WFS", "source_system": "LGE_APPLIANCE", "model_id": "MD10766829", "url": "https://www.lge.co.kr/air-conditioners/sq06gj1wfs"},
]


class ProductHTMLParser(HTMLParser):
    IMAGE_ATTRS = ("src", "data-src", "data-original", "data-pc-src", "data-mobile-src", "data-lazy")

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.title_parts: list[str] = []
        self.text_parts: list[str] = []
        self.meta: dict[str, str] = {}
        self.images: list[dict[str, str]] = []
        self.jsonld: list[str] = []
        self._in_title = False
        self._skip_depth = 0
        self._jsonld_depth = 0
        self._jsonld_parts: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        values = {key.lower(): value or "" for key, value in attrs}
        tag = tag.lower()
        if tag == "title":
            self._in_title = True
        if tag in ("script", "style", "noscript", "svg"):
            self._skip_depth += 1
        if tag == "script" and values.get("type", "").lower() == "application/ld+json":
            self._jsonld_depth = self._skip_depth
            self._jsonld_parts = []
        if tag == "meta":
            key = values.get("property") or values.get("name")
            if key and values.get("content"):
                self.meta[key.lower()] = values["content"].strip()
        if tag == "img":
            alt = values.get("alt", "").strip()
            for attr in self.IMAGE_ATTRS:
                src = values.get(attr, "").strip()
                if src and not src.startswith("data:"):
                    self.images.append({"url": src, "alt": alt, "attribute": attr})

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag == "title":
            self._in_title = False
        if tag == "script" and self._jsonld_depth:
            self.jsonld.append("".join(self._jsonld_parts).strip())
            self._jsonld_depth = 0
            self._jsonld_parts = []
        if tag in ("script", "style", "noscript", "svg") and self._skip_depth:
            self._skip_depth -= 1

    def handle_data(self, data: str) -> None:
        if self._in_title:
            self.title_parts.append(data)
        if self._jsonld_depth:
            self._jsonld_parts.append(data)
        if not self._skip_depth:
            value = " ".join(data.split())
            if value:
                self.text_parts.append(value)


class DetailImageParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.images: list[dict[str, str]] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() != "img":
            return
        values = {key.lower(): value or "" for key, value in attrs}
        src = values.get("src") or values.get("data-src") or values.get("data-original")
        if src:
            self.images.append({"url": src.strip(), "alt": values.get("alt", "").strip(), "role": "detail"})


def request_bytes(url: str, *, method: str = "GET", data: bytes | None = None, referer: str = "") -> tuple[int, bytes, dict[str, str]]:
    headers = {"User-Agent": USER_AGENT, "Accept": "*/*"}
    if referer:
        headers["Referer"] = referer
    request = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(request, timeout=45, context=SSL_CONTEXT) as response:
            return response.status, response.read(), dict(response.headers.items())
    except urllib.error.HTTPError as exc:
        return exc.code, exc.read(), dict(exc.headers.items())


def fetch_json(url: str, *, method: str = "GET", form: dict[str, str] | None = None, referer: str = "") -> tuple[int, object, str]:
    data = None
    if form is not None:
        data = urllib.parse.urlencode(form).encode("utf-8")
    status, body, _ = request_bytes(url, method=method, data=data, referer=referer)
    text = body.decode("utf-8", errors="replace")
    try:
        return status, json.loads(text), text
    except json.JSONDecodeError:
        return status, None, text


def normalize_url(url: str, base_url: str, source_system: str) -> str:
    url = html.unescape(url.strip())
    if url.startswith("//"):
        return "https:" + url
    if source_system == "HOMESTYLE" and url.startswith("/goods/"):
        return "https://static-store.lge.co.kr" + url
    return urllib.parse.urljoin(base_url, url)


def unique_images(images: list[dict[str, str]], base_url: str, source_system: str) -> list[dict[str, str]]:
    seen: set[str] = set()
    result: list[dict[str, str]] = []
    for image in images:
        url = normalize_url(image.get("url", ""), base_url, source_system)
        if not url.startswith(("http://", "https://")) or url in seen:
            continue
        seen.add(url)
        result.append({**image, "url": url})
    return result


def image_score(image: dict[str, str]) -> tuple[int, str]:
    value = (image.get("url", "") + " " + image.get("alt", "")).lower()
    if re.search(r"size|spec|dimension|info|규격|크기|폭|높이|두께|\bmm\b", value):
        return 0, value
    if re.search(r"notice|guide|manual|txt|text|bottom|상세", value):
        return 1, value
    if image.get("role") == "gallery" or re.search(r"gallery|main|medium|메인이미지", value):
        return 2, value
    if re.search(r"detail|point|intro|feature", value):
        return 3, value
    return 4, value


def select_ocr_images(images: list[dict[str, str]], limit: int = 8) -> list[dict[str, str]]:
    if not images:
        return []
    gallery = [item for item in images if item.get("role") == "gallery"][:2]
    ranked = sorted(images, key=lambda item: image_score(item)[0])
    selected: list[dict[str, str]] = []
    for item in gallery + ranked:
        if item["url"] not in {row["url"] for row in selected}:
            selected.append(item)
        if len(selected) >= limit:
            break
    return selected


def safe_extension(url: str, content_type: str) -> str:
    extension = Path(urllib.parse.urlparse(url).path).suffix.lower()
    if extension in (".jpg", ".jpeg", ".png", ".gif", ".bmp"):
        return extension
    return {"image/jpeg": ".jpg", "image/png": ".png", "image/gif": ".gif", "image/bmp": ".bmp"}.get(content_type.split(";")[0].lower(), ".jpg")


def main() -> None:
    if RUN_DIR.exists():
        shutil.rmtree(RUN_DIR)
    RAW_API_DIR.mkdir(parents=True)
    RAW_HTML_DIR.mkdir(parents=True)
    IMAGE_DIR.mkdir(parents=True)

    source_results: list[dict] = []
    image_manifest: list[dict] = []

    for product in PRODUCTS:
        product_id = product["product_id"]
        source_system = product["source_system"]
        source_url = product["url"]
        api_summary: dict[str, object] = {}
        api_images: list[dict[str, str]] = []

        if source_system == "HOMESTYLE":
            api_base = "https://livingapi.lge.co.kr/itemsvc/ajax/v1/pdp"
            goods_url = f"{api_base}/goods/{product_id}?epFlagYn=N"
            space_url = f"{api_base}/space-recommendation?goodsId={product_id}"
            package_url = f"{api_base}/packages?goodsId={product_id}"
            goods_status, goods_json, goods_raw = fetch_json(goods_url, referer=source_url)
            space_status, space_json, space_raw = fetch_json(space_url, referer=source_url)
            package_status, package_json, package_raw = fetch_json(package_url, referer=source_url)
            (RAW_API_DIR / f"{product_id}_goods.json").write_text(goods_raw, encoding="utf-8")
            (RAW_API_DIR / f"{product_id}_space.json").write_text(space_raw, encoding="utf-8")
            (RAW_API_DIR / f"{product_id}_packages.json").write_text(package_raw, encoding="utf-8")
            data = goods_json.get("data", {}) if isinstance(goods_json, dict) else {}
            notifications = (data.get("productNotification") or {}).get("items") or []
            detail_parser = DetailImageParser()
            detail_parser.feed(str(data.get("detailInfo") or ""))
            for image in data.get("images") or []:
                if image.get("imageUrl"):
                    api_images.append({"url": image["imageUrl"], "alt": image.get("imageAlt") or "", "role": "gallery"})
            api_images.extend(detail_parser.images)
            api_summary = {
                "attempted": True,
                "goods_http_status": goods_status,
                "goods_api_status": goods_json.get("status") if isinstance(goods_json, dict) else None,
                "space_http_status": space_status,
                "space_api_status": space_json.get("status") if isinstance(space_json, dict) else None,
                "packages_http_status": package_status,
                "packages_api_status": package_json.get("status") if isinstance(package_json, dict) else None,
                "product_name": data.get("productName"),
                "brand_name": data.get("brandName"),
                "category": data.get("category"),
                "purchase_options": data.get("purchaseOptions"),
                "notification_items": notifications,
                "gallery_image_count": len(data.get("images") or []),
                "detail_image_count": len(detail_parser.images),
                "detail_html_length": len(str(data.get("detailInfo") or "")),
                "field_signal_count": sum(bool(value) for value in (data.get("productName"), data.get("brandName"), data.get("category"), data.get("purchaseOptions"), notifications)),
            }
        else:
            endpoint = "https://apiv2.lge.co.kr/itemsvc/ajax/v1/product/retrieveDealProduct"
            api_status, api_json, api_raw = fetch_json(
                endpoint,
                method="POST",
                form={"modelId": product["model_id"], "modelStatusCode": "ACTIVE"},
                referer=source_url,
            )
            (RAW_API_DIR / f"{product_id}_deal_product.json").write_text(api_raw, encoding="utf-8")
            data = api_json.get("data") if isinstance(api_json, dict) else None
            api_summary = {
                "attempted": True,
                "endpoint": endpoint,
                "http_status": api_status,
                "api_status": api_json.get("status") if isinstance(api_json, dict) else None,
                "model_id": product["model_id"],
                "deal_product_model_present": bool(isinstance(data, dict) and data.get("dealProductModel")),
                "field_signal_count": 1 if isinstance(data, dict) and data.get("dealProductModel") else 0,
                "note": "API는 정상 응답했으나 핵심 제품 스펙은 공개 HTML에 포함",
            }

        html_status, html_bytes, html_headers = request_bytes(source_url, referer="https://www.lge.co.kr/")
        html_text = html_bytes.decode("utf-8", errors="replace")
        (RAW_HTML_DIR / f"{product_id}.html").write_text(html_text, encoding="utf-8")
        parser = ProductHTMLParser()
        parser.feed(html_text)
        visible_text = " ".join(parser.text_parts)
        dimension_matches = sorted(set(re.findall(r"\b\d{2,4}(?:\.\d+)?\s*(?:mm|cm|×|x|X)\s*\d{0,4}(?:\.\d+)?", visible_text)))[:30]
        keyword_snippets = []
        for match in re.finditer(r"(?:크기|규격|색상|설치|재질|사이즈|제품 크기)", visible_text, re.IGNORECASE):
            keyword_snippets.append(visible_text[max(0, match.start() - 80): match.start() + 220])
            if len(keyword_snippets) >= 12:
                break
        html_summary = {
            "attempted": True,
            "http_status": html_status,
            "content_length": len(html_bytes),
            "sha256": hashlib.sha256(html_bytes).hexdigest(),
            "title": " ".join(parser.title_parts).strip(),
            "meta_description": parser.meta.get("description") or parser.meta.get("og:description") or "",
            "og_image": parser.meta.get("og:image") or "",
            "jsonld_count": len([value for value in parser.jsonld if value]),
            "visible_character_count": len(visible_text),
            "dimension_matches": dimension_matches,
            "keyword_snippets": keyword_snippets,
            "html_image_reference_count": len(parser.images),
        }

        html_images = [{**image, "role": "html"} for image in parser.images]
        if source_system == "LGE_APPLIANCE":
            model_key = product["model_id"].lower()
            html_images = [
                image for image in html_images
                if model_key in image.get("url", "").lower()
                or re.search(r"규격|크기|폭|높이|두께|\bmm\b", image.get("alt", ""), re.IGNORECASE)
            ]
            if parser.meta.get("og:image"):
                html_images.append({"url": parser.meta["og:image"], "alt": parser.title_parts[0] if parser.title_parts else product_id, "role": "gallery"})
        else:
            html_images = [
                image for image in html_images
                if "/common/footer/" not in image.get("url", "").lower()
                and "common%2ffooter" not in image.get("url", "").lower()
                and not image.get("url", "").lower().endswith(".svg")
            ]

        all_images = unique_images(api_images + html_images, source_url, source_system)
        selected_images = select_ocr_images(all_images, limit=8)
        download_success = 0
        for image_index, image in enumerate(selected_images, start=1):
            url = image["url"]
            try:
                status, body, headers = request_bytes(url, referer=source_url)
                content_type = headers.get("Content-Type", headers.get("content-type", ""))
                if status != 200 or not body or not content_type.lower().startswith("image/"):
                    raise RuntimeError(f"HTTP {status} content-type={content_type}")
                extension = safe_extension(url, content_type)
                filename = f"{product_id}__{image_index:02d}{extension}"
                (IMAGE_DIR / filename).write_bytes(body)
                manifest_row = {
                    "product_id": product_id,
                    "source_system": source_system,
                    "source_url": source_url,
                    "file": filename,
                    "image_url": url,
                    "role": image.get("role", ""),
                    "alt": image.get("alt", ""),
                    "http_status": status,
                    "content_type": content_type,
                    "byte_count": len(body),
                    "sha256": hashlib.sha256(body).hexdigest(),
                    "download_status": "SUCCESS",
                }
                download_success += 1
            except Exception as exc:
                manifest_row = {
                    "product_id": product_id,
                    "source_system": source_system,
                    "source_url": source_url,
                    "file": "",
                    "image_url": url,
                    "role": image.get("role", ""),
                    "alt": image.get("alt", ""),
                    "http_status": "",
                    "content_type": "",
                    "byte_count": 0,
                    "sha256": "",
                    "download_status": "ERROR",
                    "error": str(exc),
                }
            image_manifest.append(manifest_row)

        source_results.append({
            "product_id": product_id,
            "source_system": source_system,
            "source_url": source_url,
            "api": api_summary,
            "html": html_summary,
            "image_discovered_count": len(all_images),
            "ocr_selected_count": len(selected_images),
            "ocr_download_success_count": download_success,
        })
        print(f"{product_id}: api=done html={html_status} discovered={len(all_images)} ocr_selected={len(selected_images)} downloaded={download_success}")

    (RUN_DIR / "source_results.json").write_text(json.dumps(source_results, ensure_ascii=False, indent=2), encoding="utf-8")
    (RUN_DIR / "image_manifest.json").write_text(json.dumps(image_manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"source_results={RUN_DIR / 'source_results.json'}")
    print(f"image_manifest={RUN_DIR / 'image_manifest.json'} rows={len(image_manifest)}")


if __name__ == "__main__":
    main()
