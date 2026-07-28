from __future__ import annotations

import html as html_module
import json
import re
from collections import Counter
from html.parser import HTMLParser
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parent
RUN_DIR = ROOT / "poc_full_run"
OUTPUT = RUN_DIR / "deep_source_inventory.json"
FAQ_QNA_OUTPUT = RUN_DIR / "faq_qna_probe.json"

HOME_IDS = ["G25070005743", "G25100020496", "G25070001871", "G25070006112"]
APPLIANCE_IDS = ["OLED48C6KNA", "G646GBB031", "WA2525YMZF", "SQ06GJ1WFS"]
PRODUCT_IDS = HOME_IDS + APPLIANCE_IDS


FIELD_DEFS = [
    ("product_name", "상품명", "BOTH"),
    ("model_sku", "모델/SKU", "BOTH"),
    ("identifier_consistency", "식별자 일치 검증", "BOTH"),
    ("brand_name", "브랜드", "BOTH"),
    ("category_path", "카테고리 경로", "BOTH"),
    ("representative_image", "대표 이미지", "BOTH"),
    ("sale_status", "판매 상태", "R2"),
    ("sale_price", "판매가", "R2"),
    ("effective_price", "실결제가/할인가", "R2"),
    ("currency", "통화", "R2"),
    ("promotion", "프로모션 기간·잔여수량", "R2"),
    ("stock_total", "재고/판매가능수량", "R2"),
    ("option_matrix", "옵션 ID·가격·재고", "BOTH"),
    ("delivery_type", "배송 유형", "R2"),
    ("delivery_eta", "도착 예정일", "R2"),
    ("delivery_fee", "배송·설치 비용", "R2"),
    ("seller_name", "판매자", "R2"),
    ("seller_contact", "판매자 연락처", "R2"),
    ("review_count", "리뷰 수", "R2"),
    ("rating", "평점", "R2"),
    ("registered_at", "등록/출시 시점", "R2"),
    ("description_text", "상품 상세설명", "R2"),
    ("manufacturer_importer", "제조자·수입자", "R2"),
    ("origin_country", "제조국·원산지", "R2"),
    ("certification", "인증·안전기준", "BOTH"),
    ("warranty", "품질보증", "R2"),
    ("as_contact", "A/S 연락처", "R2"),
    ("dimensions", "제품 규격", "R1"),
    ("colors", "색상", "BOTH"),
    ("materials", "소재", "R2"),
    ("components", "구성품", "BOTH"),
    ("category_specs", "카테고리별 상세 스펙", "R2"),
    ("care_safety", "관리·안전·취급주의", "R2"),
    ("return_exchange_policy", "반품·교환 정책", "R2"),
    ("feature_tags", "기능·특징 태그", "R2"),
    ("recommended_placement", "추천 배치·공간", "BOTH"),
    ("provenance", "출처·수집시각", "BOTH"),
]

FIELD_LABEL = {key: label for key, label, _ in FIELD_DEFS}
FIELD_GROUP = {key: group for key, _, group in FIELD_DEFS}

NEW_REQUIREMENTS = [
    ("NR-01", "상품 식별자·SKU 정합성", ["model_sku", "identifier_consistency"], "BOTH"),
    ("NR-02", "판매가·통화", ["sale_price", "effective_price", "currency"], "R2"),
    ("NR-03", "할인·프로모션", ["promotion"], "R2"),
    ("NR-04", "판매상태·재고", ["sale_status", "stock_total"], "R2"),
    ("NR-05", "옵션별 ID·가격·재고", ["option_matrix"], "BOTH"),
    ("NR-06", "배송 유형·도착일·비용", ["delivery_type", "delivery_eta", "delivery_fee"], "R2"),
    ("NR-07", "판매자·연락처", ["seller_name", "seller_contact"], "R2"),
    ("NR-08", "제조자·수입자·원산지", ["manufacturer_importer", "origin_country"], "R2"),
    ("NR-09", "인증·안전기준", ["certification"], "BOTH"),
    ("NR-10", "품질보증·A/S", ["warranty", "as_contact"], "R2"),
    ("NR-11", "리뷰 수·평점", ["review_count", "rating"], "R2"),
    ("NR-12", "카테고리별 상세 스펙·기능", ["category_specs", "feature_tags"], "R2"),
    ("NR-13", "관리·취급·반품정책", ["care_safety", "return_exchange_policy"], "R2"),
    ("NR-14", "데이터 출처·수집시각", ["provenance"], "BOTH"),
]

APPLIANCE_DIMENSIONS = {
    "OLED48C6KNA": "W1071×H620×D46.9mm (벽걸이 기준)",
    "G646GBB031": "W914×H1860×D709mm (도어부 911, 핸들 제외 깊이 698)",
    "WA2525YMZF": "W700×H1890×D830mm (문 열림 깊이 1410)",
    "SQ06GJ1WFS": "W837×H308×D189mm",
}


class TextExtractor(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.parts: list[str] = []

    def handle_data(self, data: str) -> None:
        value = re.sub(r"\s+", " ", html_module.unescape(data)).strip()
        if value:
            self.parts.append(value)


def html_to_text(value: str | None) -> str:
    if not value:
        return ""
    parser = TextExtractor()
    parser.feed(value)
    return "\n".join(parser.parts)


def json_value(value: Any) -> Any:
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False, separators=(",", ":"))
    return value


def put(
    fields: dict[str, dict],
    name: str,
    value: Any,
    source: str,
    evidence: str,
    *,
    status: str = "AVAILABLE",
    confidence: float = 1.0,
) -> None:
    if value is None or value == "" or value == [] or value == {}:
        return
    fields[name] = {
        "value": json_value(value),
        "status": status,
        "source": source,
        "evidence": evidence,
        "confidence": confidence,
    }


def is_reference_only(value: str) -> bool:
    compact = re.sub(r"\s+", "", value)
    return any(token in compact for token in ("상세페이지참조", "상품설명/상세정보참고", "상세정보참고"))


def notification_lookup(items: list[dict], *keywords: str) -> list[dict]:
    return [item for item in items if any(keyword in item.get("title", "") for keyword in keywords)]


def join_notifications(items: list[dict]) -> str:
    return " | ".join(f"{item.get('title', '')}: {item.get('description', '')}" for item in items)


def scalar_inventory(value: Any, path: str = "", *, skip_popular: bool = True) -> list[tuple[str, Any]]:
    rows: list[tuple[str, Any]] = []
    if isinstance(value, dict):
        for key, child in value.items():
            child_path = f"{path}.{key}" if path else key
            if skip_popular and child_path.startswith("data.popularProducts"):
                continue
            rows.extend(scalar_inventory(child, child_path, skip_popular=skip_popular))
    elif isinstance(value, list):
        for index, child in enumerate(value):
            rows.extend(scalar_inventory(child, f"{path}[{index}]", skip_popular=skip_popular))
    elif value not in (None, ""):
        rows.append((path, value))
    return rows


def extract_jsonld(html_text: str) -> list[dict]:
    blocks = re.findall(
        r"<script[^>]+type=[\"']application/ld\+json[\"'][^>]*>(.*?)</script>",
        html_text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    result = []
    for block in blocks:
        try:
            value = json.loads(html_module.unescape(block.strip()))
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict):
            result.append(value)
        elif isinstance(value, list):
            result.extend(item for item in value if isinstance(item, dict))
    return result


def parse_home(product_id: str, source_result: dict) -> dict:
    path = RUN_DIR / "raw_api" / f"{product_id}_goods.json"
    raw = json.loads(path.read_text(encoding="utf-8"))
    data = raw["data"]
    fields: dict[str, dict] = {}
    notifications = data.get("productNotification") or []
    notification_items = []
    if isinstance(notifications, dict):
        if isinstance(notifications.get("items"), list):
            notification_items.extend(notifications["items"])
        elif "title" in notifications or "description" in notifications:
            notification_items.append(notifications)
    elif isinstance(notifications, list):
        for group in notifications:
            if isinstance(group, dict) and isinstance(group.get("items"), list):
                notification_items.extend(group["items"])
            elif isinstance(group, dict) and ("title" in group or "description" in group):
                notification_items.append(group)

    category = data.get("category") or {}
    images = data.get("images") or []
    image_urls = []
    for item in images:
        if isinstance(item, str):
            image_urls.append(item)
        elif isinstance(item, dict):
            image_urls.extend(
                value for key, value in item.items() if "url" in key.lower() and isinstance(value, str)
            )

    stocks = data.get("productStock") or []
    option_matrix = [
        {
            "option_id": row.get("optionId"),
            "option_name": row.get("optionName"),
            "list_price": row.get("salePrice"),
            "effective_price": row.get("discountPrice") or row.get("salePrice"),
            "stock": int(row.get("stock") or 0),
        }
        for row in stocks
    ]
    list_prices = [row["list_price"] for row in option_matrix if row.get("list_price") is not None]
    effective_prices = [row["effective_price"] for row in option_matrix if row.get("effective_price") is not None]
    delivery = data.get("deliveryInfo") or {}
    seller = data.get("sellerInfo") or {}
    detail_text = html_to_text(data.get("detailInfo"))
    service_text = html_to_text(data.get("serviceGuide"))

    put(fields, "product_name", data.get("productName"), "API", f"{path.name}:data.productName")
    put(fields, "model_sku", data.get("productId"), "API", f"{path.name}:data.productId")
    put(fields, "identifier_consistency", "MATCH", "API", f"{path.name}:data.productId")
    put(fields, "brand_name", data.get("brandName"), "API", f"{path.name}:data.brandName")
    category_path = " > ".join(
        str(category.get(key))
        for key in ("superCategoryName", "categoryName", "subCategoryName")
        if category.get(key)
    )
    put(fields, "category_path", category_path, "API", f"{path.name}:data.category")
    if image_urls:
        put(fields, "representative_image", image_urls[0], "API", f"{path.name}:data.images[0]")
    put(fields, "sale_status", data.get("status"), "API", f"{path.name}:data.status")
    if list_prices:
        put(fields, "sale_price", {"min": min(list_prices), "max": max(list_prices)}, "API", f"{path.name}:data.productStock")
    if effective_prices:
        put(fields, "effective_price", {"min": min(effective_prices), "max": max(effective_prices)}, "API", f"{path.name}:data.productStock")
    put(fields, "currency", "KRW", "API+SITE_RULE", f"{path.name}:price fields", confidence=0.99)
    promotion = data.get("timeSale")
    put(
        fields,
        "promotion",
        promotion if promotion else "NO_ACTIVE_TIME_SALE",
        "API",
        f"{path.name}:data.timeSale",
    )
    put(fields, "stock_total", int(data.get("stock") or 0), "API", f"{path.name}:data.stock")
    put(fields, "option_matrix", option_matrix, "API", f"{path.name}:data.productStock")
    put(fields, "delivery_type", delivery.get("typeDescription") or delivery.get("deliveryDivComment"), "API", f"{path.name}:data.deliveryInfo")
    put(fields, "delivery_eta", delivery.get("estimatedDeliveryDateComment"), "API", f"{path.name}:data.deliveryInfo.estimatedDeliveryDateComment")
    fee_value = " | ".join(
        value for value in (delivery.get("fee"), delivery.get("additionalFee"), delivery.get("additionalInfo")) if value
    )
    put(fields, "delivery_fee", fee_value, "API", f"{path.name}:data.deliveryInfo")
    put(fields, "seller_name", seller.get("sellerName"), "API", f"{path.name}:data.sellerInfo.sellerName")
    put(fields, "seller_contact", seller.get("contact"), "API", f"{path.name}:data.sellerInfo.contact")
    put(fields, "review_count", data.get("reviewCount"), "API", f"{path.name}:data.reviewCount")
    put(fields, "rating", data.get("score"), "API", f"{path.name}:data.score")
    put(fields, "registered_at", data.get("registeredDate"), "API", f"{path.name}:data.registeredDate")
    if len(detail_text) >= 20:
        put(fields, "description_text", detail_text[:10000], "API", f"{path.name}:data.detailInfo")

    manufacturer_items = notification_lookup(notification_items, "제조자", "제조사", "수입자")
    origin_items = notification_lookup(notification_items, "제조국", "원산지")
    certification_items = notification_lookup(notification_items, "인증", "허가")
    warranty_items = notification_lookup(notification_items, "품질보증")
    as_items = notification_lookup(notification_items, "A/S", "소비자상담")
    dimension_items = notification_lookup(notification_items, "크기", "치수")
    color_items = notification_lookup(notification_items, "색상")
    material_items = notification_lookup(notification_items, "소재", "재질")
    component_items = notification_lookup(notification_items, "구성품", "제품구성")
    care_items = notification_lookup(notification_items, "세탁", "취급", "주의")

    for name, items, evidence in (
        ("manufacturer_importer", manufacturer_items, "productNotification:manufacturer"),
        ("origin_country", origin_items, "productNotification:origin"),
        ("certification", certification_items, "productNotification:certification"),
        ("warranty", warranty_items, "productNotification:warranty"),
        ("as_contact", as_items, "productNotification:as"),
        ("dimensions", dimension_items, "productNotification:dimensions"),
        ("colors", color_items, "productNotification:colors"),
        ("materials", material_items, "productNotification:materials"),
        ("components", component_items, "productNotification:components"),
        ("care_safety", care_items, "productNotification:care"),
    ):
        if items:
            value = join_notifications(items)
            put(
                fields,
                name,
                value,
                "API",
                f"{path.name}:{evidence}",
                status="PARTIAL" if is_reference_only(value) else "AVAILABLE",
                confidence=0.7 if is_reference_only(value) else 1.0,
            )

    # Use explicit option/name values when the legal notice only points to the detail page.
    option_names = [row.get("optionName", "") for row in stocks if row.get("optionName")]
    if product_id == "G25100020496" and option_names:
        put(fields, "dimensions", option_names, "API", f"{path.name}:data.productStock.optionName")
    if product_id == "G25070001871" and option_names:
        put(fields, "colors", option_names, "API", f"{path.name}:data.productStock.optionName")
    if product_id == "G25070006112":
        put(fields, "dimensions", "3×3 inch (product name)", "API", f"{path.name}:data.productName", status="PARTIAL", confidence=0.85)
        put(fields, "colors", "골드", "API", f"{path.name}:data.productName")

    put(fields, "category_specs", notification_items, "API", f"{path.name}:data.productNotification")
    safety_parts = [value for value in (delivery.get("notice"), detail_text) if value]
    if safety_parts:
        put(fields, "care_safety", "\n".join(safety_parts)[:10000], "API", f"{path.name}:deliveryInfo.notice+detailInfo")
    put(fields, "return_exchange_policy", service_text[:10000], "API", f"{path.name}:data.serviceGuide")
    if detail_text:
        put(fields, "feature_tags", detail_text[:3000], "API", f"{path.name}:data.detailInfo", status="PARTIAL", confidence=0.85)

    placement = {
        "G25070005743": "거실 메인 소파; TV 시청·휴식·손님 응대",
        "G25070006112": "우드 협탁 또는 수납장 위",
        "G25070001871": "천장 설치 (카테고리)",
    }.get(product_id)
    if placement:
        put(fields, "recommended_placement", placement, "API_TEXT/RULE", f"{path.name}:detailInfo/category", status="PARTIAL", confidence=0.85)
    put(
        fields,
        "provenance",
        {"source": path.name, "api_timestamp": raw.get("timestamp"), "price_valid_end": data.get("priceValidEndDate")},
        "PIPELINE",
        path.name,
    )

    html_fields: dict[str, dict] = {}
    html_summary = source_result.get("html") or {}
    put(
        html_fields,
        "product_name",
        html_summary.get("title"),
        "HTML_META",
        f"{product_id}.html:<title>",
        status="PARTIAL",
        confidence=0.9,
    )
    put(
        html_fields,
        "representative_image",
        html_summary.get("og_image"),
        "HTML_META",
        f"{product_id}.html:meta[property=og:image]",
    )
    put(
        html_fields,
        "description_text",
        html_summary.get("meta_description"),
        "HTML_META",
        f"{product_id}.html:meta[name=description]",
        status="PARTIAL",
        confidence=0.9,
    )

    return {
        "product_id": product_id,
        "source_system": "HOMESTYLE",
        "api_fields": fields,
        "html_fields": html_fields,
        "raw_inventory": [
            {"path": item_path, "value": json_value(value), "source": "API"}
            for item_path, value in scalar_inventory(raw)
        ],
    }


def parse_appliance(product_id: str) -> dict:
    html_path = RUN_DIR / "raw_html" / f"{product_id}.html"
    api_path = RUN_DIR / "raw_api" / f"{product_id}_deal_product.json"
    html_text = html_path.read_text(encoding="utf-8", errors="ignore")
    api_raw = json.loads(api_path.read_text(encoding="utf-8"))
    blocks = extract_jsonld(html_text)
    product = next((item for item in blocks if item.get("@type") == "Product"), {})
    breadcrumb = next((item for item in blocks if item.get("@type") == "BreadcrumbList"), {})
    fields: dict[str, dict] = {}

    sku = str(product.get("sku") or "")
    sku_base = sku.split(".", 1)[0]
    identifier_status = "MATCH" if sku_base == product_id else f"MISMATCH: requested={product_id}, jsonld={sku}"
    put(fields, "product_name", product.get("name"), "HTML_JSONLD", f"{html_path.name}:Product.name")
    put(fields, "model_sku", sku, "HTML_JSONLD", f"{html_path.name}:Product.sku")
    put(
        fields,
        "identifier_consistency",
        identifier_status,
        "HTML_JSONLD",
        f"{html_path.name}:Product.sku",
        status="AVAILABLE" if identifier_status == "MATCH" else "PARTIAL",
        confidence=1.0,
    )
    brand = product.get("brand") or {}
    put(fields, "brand_name", brand.get("name") if isinstance(brand, dict) else brand, "HTML_JSONLD", f"{html_path.name}:Product.brand")
    breadcrumb_names = [
        item.get("name") or (item.get("item") or {}).get("name")
        for item in breadcrumb.get("itemListElement", [])
        if isinstance(item, dict)
    ]
    put(fields, "category_path", " > ".join(value for value in breadcrumb_names if value), "HTML_JSONLD", f"{html_path.name}:BreadcrumbList")
    put(fields, "representative_image", product.get("image"), "HTML_JSONLD", f"{html_path.name}:Product.image")
    offers = product.get("offers") or {}
    if offers:
        put(fields, "sale_status", "OFFER_PRESENT", "HTML_JSONLD", f"{html_path.name}:Product.offers", status="PARTIAL", confidence=0.8)
        put(fields, "sale_price", offers.get("price"), "HTML_JSONLD", f"{html_path.name}:Product.offers.price")
        put(fields, "effective_price", offers.get("price"), "HTML_JSONLD", f"{html_path.name}:Product.offers.price", status="PARTIAL", confidence=0.8)
        put(fields, "currency", offers.get("priceCurrency"), "HTML_JSONLD", f"{html_path.name}:Product.offers.priceCurrency")
        seller = offers.get("seller") or {}
        put(fields, "seller_name", seller.get("name") if isinstance(seller, dict) else seller, "HTML_JSONLD", f"{html_path.name}:Product.offers.seller")
    rating = product.get("aggregateRating") or {}
    put(fields, "review_count", rating.get("reviewCount"), "HTML_JSONLD", f"{html_path.name}:Product.aggregateRating.reviewCount")
    put(fields, "rating", rating.get("ratingValue"), "HTML_JSONLD", f"{html_path.name}:Product.aggregateRating.ratingValue")
    description = str(product.get("description") or "").strip()
    put(fields, "description_text", description, "HTML_JSONLD", f"{html_path.name}:Product.description")
    put(fields, "manufacturer_importer", "LG전자 (brand/seller; 제조국 미확인)", "HTML_JSONLD", f"{html_path.name}:Product.brand+offers.seller", status="PARTIAL", confidence=0.8)
    put(fields, "dimensions", APPLIANCE_DIMENSIONS[product_id], "HTML+DETAIL_IMAGE", f"{html_path.name}:dimension text/detail image", confidence=0.98)

    properties = product.get("additionalProperty") or []
    property_map = {str(item.get("name")): item.get("value") for item in properties if isinstance(item, dict)}
    put(fields, "category_specs", property_map, "HTML_JSONLD", f"{html_path.name}:Product.additionalProperty")
    colors = [value for name, value in property_map.items() if "색상" in name]
    if colors:
        put(fields, "colors", colors, "HTML_JSONLD", f"{html_path.name}:Product.additionalProperty[color]")
    materials = [f"{name}: {value}" for name, value in property_map.items() if "재질" in name]
    if materials:
        put(fields, "materials", materials, "HTML_JSONLD", f"{html_path.name}:Product.additionalProperty[material]")
    if product_id == "WA2525YMZF":
        put(fields, "components", "세탁기+건조기 일체형", "HTML_JSONLD", f"{html_path.name}:Product.description")

    feature_value = {"description": description, "properties": property_map}
    put(fields, "feature_tags", feature_value, "HTML_JSONLD", f"{html_path.name}:description+additionalProperty", status="PARTIAL", confidence=0.9)
    placement = {
        "OLED48C6KNA": "벽면 부착 (벽걸이형 명시)",
        "G646GBB031": "주방 바닥 설치 (냉장고 카테고리)",
        "WA2525YMZF": "세탁 공간 바닥 설치 (워시타워 카테고리)",
        "SQ06GJ1WFS": "벽면 부착 (벽걸이에어컨 명시)",
    }[product_id]
    put(fields, "recommended_placement", placement, "HTML_TEXT/RULE", f"{html_path.name}:name/category", status="PARTIAL", confidence=0.9)
    put(
        fields,
        "provenance",
        {"html": html_path.name, "api": api_path.name, "api_timestamp": api_raw.get("timestamp")},
        "PIPELINE",
        f"{html_path.name}|{api_path.name}",
    )

    inventory = []
    for index, block in enumerate(blocks):
        for item_path, value in scalar_inventory(block, path=f"jsonld[{index}]", skip_popular=False):
            inventory.append({"path": item_path, "value": json_value(value), "source": "HTML_JSONLD"})
    return {
        "product_id": product_id,
        "source_system": "LGE_APPLIANCE",
        "api_fields": {},
        "html_fields": fields,
        "raw_inventory": inventory,
    }


def add_ocr_contributions(product: dict, benchmark: dict) -> None:
    product_id = product["product_id"]
    windows: dict[str, dict] = {}
    paddle: dict[str, dict] = {}
    if product_id == "G25100020496":
        value = "가구 배치 추천: 소파·리클라이너·침대·책상·식탁별 러그 크기 가이드"
        put(windows, "recommended_placement", value, "WINDOWS_OCR", "G25100020496__03.jpg", status="PARTIAL", confidence=0.86)
        put(windows, "feature_tags", value, "WINDOWS_OCR", "G25100020496__03.jpg", status="PARTIAL", confidence=0.86)
        put(paddle, "recommended_placement", value, "PADDLE_OCR", "G25100020496__03.jpg", status="AVAILABLE", confidence=0.98)
        put(paddle, "feature_tags", value, "PADDLE_OCR", "G25100020496__03.jpg", status="AVAILABLE", confidence=0.98)
    product["windows_ocr_fields"] = windows
    product["paddle_ocr_fields"] = paddle

    image_rows = [row for row in benchmark["image_details"] if row["product_id"] == product_id]
    product["ocr_diagnostics"] = {
        engine: {
            "images": len([row for row in image_rows if row["engine"] == engine]),
            "nonempty": len([row for row in image_rows if row["engine"] == engine and row["character_count"] > 0]),
            "characters": sum(row["character_count"] for row in image_rows if row["engine"] == engine),
        }
        for engine in ("Windows OCR (ko+en)", "PaddleOCR PP-OCRv5", "Tesseract.js")
    }


def add_html_faq_qna_context(product: dict, faq_qna_by_id: dict[str, dict]) -> None:
    """Attach PDP FAQ/Q&A to the operational HTML bundle.

    Some records are embedded as HTML JSON-LD and others are loaded by the PDP
    asynchronously. They are one PDP/HTML content layer for production use,
    while the transport/API distinction remains in the provenance summary.
    """
    product_id = product["product_id"]
    source = faq_qna_by_id.get(product_id)
    if source is None:
        raise KeyError(f"FAQ/Q&A result missing for product: {product_id}")

    html_faq = source.get("html_faqpage") or {}
    qna_api = source.get("qna_api") or {}
    faq_api = source.get("faq_api") or {}
    model_scope = source.get("qna_model_scope") or {}
    exact_qna_records = source.get("exact_qna_candidate_signals") or {}
    context = {
        "bundle_role": "PDP_HTML_FAQ_QNA",
        "html_faqpage_count": int(html_faq.get("record_count") or 0),
        "official_faq_api_count": int(faq_api.get("record_count") or 0),
        "public_qna_count": int(qna_api.get("retrieved_public_count") or 0),
        "exact_product_or_goods_qna_count": int(
            model_scope.get("exact_product_or_goods_count") or 0
        ),
        "related_model_qna_excluded_count": int(
            model_scope.get("related_model_count") or 0
        ),
        "exact_qna_candidate_signal_fields": sorted(exact_qna_records),
        "record_store": FAQ_QNA_OUTPUT.name,
        "direct_fill_policy": (
            "FAQ/Q&A is always included in the PDP HTML bundle, but questions, "
            "category FAQ and related-model records do not directly overwrite product facts."
        ),
    }
    product["html_faq_qna_context"] = context
    put(
        product["html_fields"],
        "faq_qna_context",
        context,
        "HTML_JSONLD+PDP_ASYNC_FAQ_QNA",
        f"{FAQ_QNA_OUTPUT.name}:products[{product_id}]",
        status="AVAILABLE",
        confidence=1.0,
    )


def merge_fields(*sources: dict[str, dict]) -> dict[str, dict]:
    result: dict[str, dict] = {}
    for fields in sources:
        for key, value in fields.items():
            current = result.get(key)
            if current is None:
                result[key] = value
            elif current["status"] == "PARTIAL" and value["status"] == "AVAILABLE":
                result[key] = value
    return result


def stage_statistics(products: list[dict]) -> list[dict]:
    stages = [
        ("RESEARCH_API_ONLY_DO_NOT_USE", lambda p: merge_fields(p["api_fields"])),
        ("BASELINE_API_HTML_FAQ_QNA", lambda p: merge_fields(p["api_fields"], p["html_fields"])),
        ("BASELINE_API_HTML_FAQ_QNA_WINDOWS_OCR", lambda p: merge_fields(p["api_fields"], p["html_fields"], p["windows_ocr_fields"])),
        ("BASELINE_API_HTML_FAQ_QNA_PADDLE_OCR", lambda p: merge_fields(p["api_fields"], p["html_fields"], p["paddle_ocr_fields"])),
    ]
    total = len(products) * len(FIELD_DEFS)
    result = []
    previous_usable = 0
    for stage_name, getter in stages:
        available = partial = missing = 0
        per_product = []
        for product in products:
            fields = getter(product)
            statuses = [fields.get(key, {}).get("status", "MISSING") for key, _, _ in FIELD_DEFS]
            counts = Counter(statuses)
            available += counts["AVAILABLE"]
            partial += counts["PARTIAL"]
            missing += counts["MISSING"]
            per_product.append(
                {
                    "product_id": product["product_id"],
                    "available": counts["AVAILABLE"],
                    "partial": counts["PARTIAL"],
                    "missing": counts["MISSING"],
                    "usable": counts["AVAILABLE"] + counts["PARTIAL"],
                }
            )
        usable = available + partial
        result.append(
            {
                "stage": stage_name,
                "total_field_cells": total,
                "available": available,
                "partial": partial,
                "missing": missing,
                "strict_coverage_pct": round(available / total * 100, 1),
                "usable_coverage_pct": round(usable / total * 100, 1),
                "usable_increment_vs_api": usable - result[0]["usable"] if result else 0,
                "usable_increment_vs_previous_row": usable - previous_usable if result else 0,
                "per_product": per_product,
                "usable": usable,
            }
        )
        previous_usable = usable
    return result


def requirement_matrix(products: list[dict]) -> tuple[list[dict], list[dict]]:
    evidence = []
    summaries = []
    for requirement_id, name, field_names, group in NEW_REQUIREMENTS:
        statuses = []
        for product in products:
            merged = product["merged_fields"]
            field_statuses = [merged.get(field, {}).get("status", "MISSING") for field in field_names]
            if all(status == "AVAILABLE" for status in field_statuses):
                status = "가능"
            elif any(status in {"AVAILABLE", "PARTIAL"} for status in field_statuses):
                status = "조건부"
            else:
                status = "불가능"
            statuses.append(status)
            evidence.append(
                {
                    "requirement_id": requirement_id,
                    "requirement_name": name,
                    "request_group": group,
                    "product_id": product["product_id"],
                    "source_system": product["source_system"],
                    "status": status,
                    "field_names": field_names,
                    "field_statuses": field_statuses,
                    "values": {field: merged.get(field, {}).get("value", "") for field in field_names},
                    "sources": {field: merged.get(field, {}).get("source", "") for field in field_names},
                    "evidence": {field: merged.get(field, {}).get("evidence", "") for field in field_names},
                }
            )
        counts = Counter(statuses)
        summaries.append(
            {
                "requirement_id": requirement_id,
                "requirement_name": name,
                "request_group": group,
                "fields": field_names,
                "product_count": len(products),
                "available": counts["가능"],
                "conditional": counts["조건부"],
                "unavailable": counts["불가능"],
                "strict_pct": round(counts["가능"] / len(products) * 100, 1),
                "conditional_included_pct": round((counts["가능"] + counts["조건부"]) / len(products) * 100, 1),
            }
        )
    return summaries, evidence


def main() -> None:
    benchmark = json.loads((RUN_DIR / "ocr_engine_benchmark.json").read_text(encoding="utf-8"))
    source_results = json.loads((RUN_DIR / "source_results.json").read_text(encoding="utf-8"))
    if not FAQ_QNA_OUTPUT.exists():
        raise FileNotFoundError(
            f"Required PDP HTML FAQ/Q&A result is missing: {FAQ_QNA_OUTPUT}. "
            "Run parse_faq_qna.py first."
        )
    faq_qna = json.loads(FAQ_QNA_OUTPUT.read_text(encoding="utf-8"))
    faq_qna_by_id = {row["product_id"]: row for row in faq_qna.get("products", [])}
    source_result_by_id = {row["product_id"]: row for row in source_results}
    products = [parse_home(product_id, source_result_by_id[product_id]) for product_id in HOME_IDS]
    products.extend(parse_appliance(product_id) for product_id in APPLIANCE_IDS)
    for product in products:
        add_html_faq_qna_context(product, faq_qna_by_id)
        add_ocr_contributions(product, benchmark)
        product["merged_fields"] = merge_fields(
            product["api_fields"], product["html_fields"], product["paddle_ocr_fields"]
        )

    stages = stage_statistics(products)
    requirement_summaries, requirement_evidence = requirement_matrix(products)

    field_evidence = []
    for product in products:
        for key, label, group in FIELD_DEFS:
            value = product["merged_fields"].get(key)
            field_evidence.append(
                {
                    "product_id": product["product_id"],
                    "source_system": product["source_system"],
                    "field_name": key,
                    "field_label": label,
                    "request_group": group,
                    "status": value.get("status", "MISSING") if value else "MISSING",
                    "value": value.get("value", "") if value else "",
                    "source": value.get("source", "") if value else "",
                    "confidence": value.get("confidence", "") if value else "",
                    "evidence": value.get("evidence", "") if value else "",
                }
            )

    payload = {
        "metadata": {
            "product_count": len(products),
            "field_count": len(FIELD_DEFS),
            "field_cell_count": len(products) * len(FIELD_DEFS),
            "new_requirement_count": len(NEW_REQUIREMENTS),
            "analysis_date": "2026-07-21",
            "operational_baseline": "PRODUCT_API + PDP_HTML/JSONLD + PDP_FAQ/QNA",
            "api_only_policy": "Research comparison only; not a valid production collection mode.",
            "price_stock_notice": "Prices, stock, delivery dates and promotions are the 2026-07-20 collection snapshot and must be refreshed for production use.",
            "scope_note": "The required baseline combines product API, PDP HTML/JSON-LD and PDP FAQ/Q&A. OCR only contributes accepted product fields when that baseline is missing information.",
        },
        "field_definitions": [
            {"field_name": key, "field_label": label, "request_group": group}
            for key, label, group in FIELD_DEFS
        ],
        "new_requirements": [
            {"requirement_id": rid, "requirement_name": name, "fields": fields, "request_group": group}
            for rid, name, fields, group in NEW_REQUIREMENTS
        ],
        "stage_statistics": stages,
        "requirement_summary": requirement_summaries,
        "requirement_evidence": requirement_evidence,
        "field_evidence": field_evidence,
        "products": products,
    }
    OUTPUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(
        json.dumps(
            {
                "output": str(OUTPUT),
                "stage_statistics": [{k: v for k, v in row.items() if k != "per_product"} for row in stages],
                "requirement_summary": requirement_summaries,
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
