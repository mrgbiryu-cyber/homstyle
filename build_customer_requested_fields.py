from __future__ import annotations

import json
import re
import html as html_module
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET

import build_feasibility_sheet as xlsx_writer


ROOT = Path(__file__).resolve().parent
RUN_DIR = ROOT / "poc_full_run"
OUTPUT = ROOT / "상품ID별_17개_요구필드_확장양식_롤링이미지_숫자규격_옵션분리.xlsx"
M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
MISSING = "미확보"
NOT_APPLICABLE = "해당없음"


SPACE_MAP = {
    "LIVING_ROOM": "리빙룸",
    "BEDROOM": "베드룸",
    "DINING_ROOM": "다이닝룸",
    "ENTRY": "현관",
    "KITCHEN": "주방",
    "LAUNDRY_ROOM": "세탁실",
    "UTILITY_ROOM": "다용도실",
    "STUDY": "서재",
}

PLACEMENT_MAP = {
    "FLOOR": "바닥",
    "WALL": "벽",
    "CEILING": "천장",
    "SHELF": "선반",
    "CONSOLE": "콘솔",
    "TABLETOP": "테이블 위",
}

STYLE_MAP = {
    "CONTEMPORARY": "컨템포러리",
    "MODERN": "모던",
    "RECTILINEAR": "직선형",
    "GEOMETRIC": "기하학형",
    "MINIMAL": "미니멀",
    "SCULPTURAL": "조형적",
    "CLASSIC": "클래식",
    "DECORATIVE": "장식형",
    "VINTAGE": "빈티지",
    "VERTICAL": "수직형",
}

PURPOSE_MAP = {
    "REST": "휴식",
    "TV": "TV 시청",
    "GUEST": "손님 응대",
    "LIVING_ROOM_FLOOR": "거실 바닥 연출",
    "BEDROOM_FLOOR": "침실 바닥 연출",
    "AMBIENT_LIGHTING": "공간 조명",
    "PHOTO_DISPLAY": "사진 전시",
    "MEDIA_VIEWING": "미디어 시청",
    "FOOD_STORAGE": "식품 보관",
    "LAUNDRY": "세탁",
    "DRYING": "건조",
    "COOLING": "냉방",
    "DEHUMIDIFICATION": "제습",
}

INSTALLATION_MAP = {
    "WALL_MOUNTED": "벽걸이형",
    "FIT_AND_MAX_FLOOR_STANDING": "스탠딩형(Fit & Max)",
    "FLOOR_STANDING_STACKED": "스탠딩 일체형",
    "WALL_MOUNTED_PRO_INSTALL": "벽걸이형(전문 설치)",
}

MOOD_BY_PRODUCT = {
    "G25070005743": "편안함·따뜻함 (추론)",
    "G25100020496": "편안함·캐주얼 (추론)",
    "G25070001871": "모던·포인트 (추론)",
    "G25070006112": "클래식·장식적 (추론)",
    "OLED48C6KNA": "모던·몰입감 (추론)",
    "G646GBB031": "차분함·정돈감 (추론)",
    "WA2525YMZF": "깔끔함·실용적 (추론)",
    "SQ06GJ1WFS": "깔끔함·쾌적함 (추론)",
}


def read_result_rows() -> list[dict[str, str]]:
    matches = list(ROOT.glob("*v6_API_HTML*.xlsx"))
    if len(matches) != 1:
        raise FileNotFoundError(f"v6 result workbook not found or ambiguous: {matches}")
    with ZipFile(matches[0]) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {item.attrib["Id"]: item.attrib["Target"] for item in relationships}
        sheet = next(
            item
            for item in workbook.find(M + "sheets")
            if item.attrib["name"].startswith("04_") and item.attrib["name"].endswith("RESULT")
        )
        target = targets[sheet.attrib[R + "id"]]
        target = target if target.startswith("xl/") else "xl/" + target.lstrip("/")
        root = ET.fromstring(archive.read(target))
        values: list[list[str]] = []
        for row in root.findall(".//" + M + "sheetData/" + M + "row"):
            row_values = []
            for cell in row.findall(M + "c"):
                if cell.attrib.get("t") == "inlineStr":
                    row_values.append("".join(node.text or "" for node in cell.iter(M + "t")))
                else:
                    value = cell.find(M + "v")
                    row_values.append(value.text or "" if value is not None else "")
            values.append(row_values)
    headers = values[0]
    return [dict(zip(headers, row)) for row in values[1:]]


def translated_list(value: str, mapping: dict[str, str]) -> str:
    if not value:
        return MISSING
    return "|".join(mapping.get(item, item) for item in value.split("|") if item)


def deep_value(product: dict, field: str, *, allow_reference: bool = False) -> str:
    entry = product["merged_fields"].get(field) or {}
    if entry.get("status") == "MISSING" or not entry.get("value"):
        return ""
    text = str(entry["value"]).strip()
    compact = re.sub(r"\s+", "", text)
    reference_only = any(
        token in compact
        for token in ("상품설명/상세정보참고", "상세페이지참조", "상세정보참고")
    )
    if reference_only and entry.get("status") == "PARTIAL" and not allow_reference:
        return ""
    return text


def concise(value: str, limit: int = 500) -> str:
    value = re.sub(r"\s+", " ", value).strip()
    return value if len(value) <= limit else value[:limit] + "…"


def price_text(value: str) -> str:
    if not value:
        return ""
    try:
        parsed = json.loads(value)
    except json.JSONDecodeError:
        return value
    if not isinstance(parsed, dict):
        return str(parsed)
    low, high = parsed.get("min"), parsed.get("max")
    if low is None:
        return ""
    if low == high:
        return f"{int(low):,}원"
    return f"{int(low):,}~{int(high):,}원"


def dimension_value(row: dict[str, str], deep_product: dict) -> str:
    product_id = row["product_id"]
    if product_id == "G25070005743":
        return "소파 W2910×D1020×H910mm | 스툴 W740×D660×H410mm"
    if product_id == "G25100020496":
        return "W1400×D2000 / W1600×D2300 / W2000×D3000mm | H 미확보"
    if row["width_mm"] and row["depth_mm"] and row["height_mm"]:
        return f"W{row['width_mm']}×D{row['depth_mm']}×H{row['height_mm']}mm"
    value = deep_value(deep_product, "dimensions")
    if value and any(char.isdigit() for char in value) and product_id == "G25070006112":
        return "3×3인치 사진 규격 | 제품 외형 W×D×H 미확보"
    return MISSING


def numeric_value(value: str | int | float | None) -> int | float | None:
    if value in (None, ""):
        return None
    text = str(value).strip().replace(",", "")
    match = re.fullmatch(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    number = float(text)
    return int(number) if number.is_integer() else number


def rolling_image_urls(product_id: str, is_home: bool) -> list[str]:
    urls: list[str] = []
    if is_home:
        path = RUN_DIR / "raw_api" / f"{product_id}_goods.json"
        raw = json.loads(path.read_text(encoding="utf-8"))
        images = (raw.get("data") or {}).get("images") or []
        images = sorted(
            (row for row in images if isinstance(row, dict)),
            key=lambda row: int(row.get("sortSeq") or 9999),
        )
        for image in images:
            if image.get("type") not in (None, "", "IMAGE"):
                continue
            url = str(image.get("imageUrl") or "").strip()
            if url.startswith("/goods/"):
                url = "https://static-store.lge.co.kr" + url
            if url.startswith(("http://", "https://")) and url not in urls:
                urls.append(url)
    else:
        path = RUN_DIR / "raw_html" / f"{product_id}.html"
        source = html_module.unescape(path.read_text(encoding="utf-8", errors="ignore"))
        # One canonical medium image per PDP carousel slide. This intentionally
        # excludes duplicate small/large renditions of the same slide.
        pattern = re.compile(
            r'<img\s+data-lazy="([^"]+/gallery/medium[0-9]+\.(?:jpg|jpeg|png|webp))"[^>]*>',
            flags=re.IGNORECASE,
        )
        for url in pattern.findall(source):
            if url.startswith("//"):
                url = "https:" + url
            elif url.startswith("/"):
                url = "https://www.lge.co.kr" + url
            if url.startswith(("http://", "https://")) and url not in urls:
                urls.append(url)
    return urls


def normalize_option_style(raw_style: str, values: list[str], colors_raw: str) -> str:
    """Convert loose PDP labels into a small, human-readable option style set."""
    compact = re.sub(r"\s+", "", raw_style).casefold()
    if "사이즈" in compact or "size" in compact:
        return "사이즈"
    if "단일" in compact or (
        len(values) == 1 and re.sub(r"\s+", "", values[0]).casefold() in {"단일", "단일옵션"}
    ):
        return "단일옵션"

    color_values = {
        re.sub(r"\s+", "", value).casefold()
        for value in (colors_raw or "").split("|")
        if value.strip()
    }
    option_values = {
        re.sub(r"\s+", "", value).casefold()
        for value in values
        if value.strip()
    }
    if color_values and option_values and option_values.issubset(color_values):
        return "색상"
    if any(token in compact for token in ("색상", "컬러", "color", "colour")):
        return "색상"
    return raw_style.strip() or "옵션"


def appliance_selected_option(product_id: str) -> tuple[str, str, str]:
    """Read the currently selected variant from the appliance PDP HTML."""
    path = RUN_DIR / "raw_html" / f"{product_id}.html"
    source = path.read_text(encoding="utf-8", errors="ignore")
    match = re.search(
        r'class="lists Spec spec-type-2 on"[^>]*>.*?옵션\s*선택(.*?)정상가',
        source,
        flags=re.DOTALL | re.IGNORECASE,
    )
    if not match:
        return "단일옵션", "단일옵션", "PDP HTML에 옵션 선택 영역 없음; 단일 SKU로 분류"

    selected = re.sub(r"<[^>]+>", " ", match.group(1))
    selected = re.sub(r"\s+", " ", html_module.unescape(selected)).strip()
    if product_id == "OLED48C6KNA":
        style = "화면크기/설치형태"
    elif product_id == "G646GBB031":
        style = "도어재질/색상"
    elif product_id == "WA2525YMZF":
        style = "색상"
    else:
        style = "옵션"
    return style, selected or MISSING, "가전 PDP HTML 옵션 선택 영역의 현재 선택값"


def option_groups(row: dict[str, str], is_home: bool) -> list[dict]:
    product_id = row["product_id"]
    if not is_home:
        style, selected, evidence = appliance_selected_option(product_id)
        return [
            {
                "style": style,
                "raw_style": "옵션 선택" if style != "단일옵션" else "",
                "items": [
                    {
                        "name": selected,
                        "option_id": "",
                        "image_url": "",
                    }
                ],
                "evidence": evidence,
            }
        ]

    path = RUN_DIR / "raw_api" / f"{product_id}_goods.json"
    raw = json.loads(path.read_text(encoding="utf-8"))
    data = raw.get("data") or {}
    stocks = {
        str(item.get("optionName") or "").strip(): item
        for item in (data.get("productStock") or [])
        if isinstance(item, dict)
    }
    groups: list[dict] = []
    for group in data.get("purchaseOptions") or []:
        if not isinstance(group, dict):
            continue
        raw_style = str(group.get("title") or "").strip()
        raw_items = [item for item in (group.get("items") or []) if isinstance(item, dict)]
        values = [str(item.get("name") or "").strip() for item in raw_items]
        values = [value for value in values if value]
        items = []
        for raw_item in raw_items:
            name = str(raw_item.get("name") or "").strip()
            if not name:
                continue
            stock = stocks.get(name) or {}
            image_url = str(raw_item.get("imageUrl") or "").strip()
            if image_url.startswith("/"):
                image_url = "https://static-store.lge.co.kr" + image_url
            items.append(
                {
                    "name": name,
                    "option_id": str(stock.get("optionId") or ""),
                    "image_url": image_url,
                }
            )
        if items:
            groups.append(
                {
                    "style": normalize_option_style(raw_style, values, row.get("colors_raw", "")),
                    "raw_style": raw_style,
                    "items": items,
                    "evidence": "홈스타일 goods API purchaseOptions + productStock",
                }
            )
    if groups:
        return groups
    return [
        {
            "style": MISSING,
            "raw_style": "",
            "items": [{"name": MISSING, "option_id": "", "image_url": ""}],
            "evidence": "홈스타일 goods API에 옵션 구조 없음",
        }
    ]


def dimension_records(row: dict[str, str], deep_product: dict) -> list[dict]:
    product_id = row["product_id"]
    if product_id == "G25070005743":
        return [
            {
                "target": "소파 본체",
                "w_mm": 2910,
                "d_mm": 1020,
                "h_mm": 910,
                "status": "확보",
                "evidence": "상품 API 상품정보고시 크기",
            },
            {
                "target": "스툴",
                "w_mm": 740,
                "d_mm": 660,
                "h_mm": 410,
                "status": "확보",
                "evidence": "상품 API 상품정보고시 크기",
            },
        ]
    if product_id == "G25100020496":
        return [
            {
                "target": f"러그 옵션 {w}×{d}",
                "w_mm": w,
                "d_mm": d,
                "h_mm": None,
                "status": "부분확보",
                "evidence": "상품 옵션명; 두께/높이 미확보",
            }
            for w, d in ((1400, 2000), (1600, 2300), (2000, 3000))
        ]
    if product_id == "G25070001871":
        return [
            {
                "target": "제품 외형",
                "w_mm": None,
                "d_mm": None,
                "h_mm": None,
                "status": MISSING,
                "evidence": "PDP/API/FAQ·Q&A에 제품 외형 치수 없음",
            }
        ]
    if product_id == "G25070006112":
        return [
            {
                "target": "제품 외형",
                "w_mm": None,
                "d_mm": None,
                "h_mm": None,
                "status": MISSING,
                "evidence": "3×3인치는 사진 규격이며 제품 외형 치수가 아님",
            }
        ]

    w_mm = numeric_value(row.get("width_mm"))
    d_mm = numeric_value(row.get("depth_mm"))
    h_mm = numeric_value(row.get("height_mm"))
    status = "확보" if all(value is not None for value in (w_mm, d_mm, h_mm)) else "부분확보"
    if all(value is None for value in (w_mm, d_mm, h_mm)):
        status = MISSING
    return [
        {
            "target": "제품 본체",
            "w_mm": w_mm,
            "d_mm": d_mm,
            "h_mm": h_mm,
            "status": status,
            "evidence": "PDP HTML/JSON-LD/상세 규격",
        }
    ]


def sales_tag(deep_product: dict) -> str:
    status = deep_value(deep_product, "sale_status")
    sale_price = price_text(deep_value(deep_product, "sale_price"))
    effective = price_text(deep_value(deep_product, "effective_price"))
    seller = deep_value(deep_product, "seller_name")
    parts = []
    if status:
        parts.append("판매중" if status == "ON_SALE" else status)
    if sale_price:
        parts.append("판매가 " + sale_price)
    if effective:
        parts.append("실결제가 " + effective)
    if seller:
        parts.append("판매자 " + seller)
    if parts:
        parts.append("2026-07-20 수집 기준")
    return " | ".join(parts) if parts else MISSING


def handling_tag(product_id: str, deep_product: dict) -> str:
    care = deep_value(deep_product, "care_safety")
    if not care:
        return MISSING
    if product_id == "G25070005743":
        return "설치·사용 전 제품 상태 확인; 설치 완료 후 단순변심 교환·반품 불가"
    if product_id == "G25070006112":
        marker = care.find("주의사항")
        return concise(care[marker:] if marker >= 0 else care)
    return concise(care)


def safety_tag(product_id: str, deep_product: dict) -> str:
    certification = deep_value(deep_product, "certification")
    if certification:
        return concise(certification)
    if product_id == "G25070006112":
        return "날카로운 물체의 긁힘 및 외부 충격 주의"
    return MISSING


def build_rows() -> tuple[
    list[list],
    list[str],
    list[list],
    list[str],
    list[list],
    list[str],
    list[list],
    list[str],
]:
    source_rows = read_result_rows()
    deep = json.loads((RUN_DIR / "deep_source_inventory.json").read_text(encoding="utf-8"))
    deep_by_id = {row["product_id"]: row for row in deep["products"]}

    rolling_by_id: dict[str, list[str]] = {}
    dimensions_by_id: dict[str, list[dict]] = {}
    options_by_id: dict[str, list[dict]] = {}
    for row in source_rows:
        product_id = row["product_id"]
        deep_product = deep_by_id[product_id]
        is_home = deep_product["source_system"] == "HOMESTYLE"
        rolling_by_id[product_id] = rolling_image_urls(product_id, is_home)
        dimensions_by_id[product_id] = dimension_records(row, deep_product)
        options_by_id[product_id] = option_groups(row, is_home)
    max_rolling_images = max((len(urls) for urls in rolling_by_id.values()), default=1)
    max_options = max(
        (
            sum(len(group["items"]) for group in groups)
            for groups in options_by_id.values()
        ),
        default=1,
    )

    home_ids = [row["product_id"] for row in source_rows if row["product_id"].startswith("G25")]
    appliance_ids = [row["product_id"] for row in source_rows if row["product_id"] not in home_ids]
    summary_headers = ["구분", "상품수/필드수", "상품 ID / 설명"]
    summary_rows = [summary_headers,
        ["홈스타일", len(home_ids), " | ".join(home_ids)],
        ["가전", len(appliance_ids), " | ".join(appliance_ids)],
        ["합계", len(source_rows), "대표 상품 8개"],
        [
            "요청 1",
            "10개 항목",
            f"롤링 이미지 {max_rolling_images}열, 옵션 {max_options}열, W/D/H 숫자 분리, 중·소카테고리 분리를 포함해 출력 필드 {max_rolling_images + max_options + 18}개",
        ],
        ["요청 2", "7개 항목", "상품 설명서 태그 13개와 공간 콘텐츠 태그 8개를 개별 필드로 분리하여 출력 필드 26개"],
        ["롤링 이미지", f"최대 {max_rolling_images}장", "대표 이미지를 포함한 PDP 롤링 순서; URL마다 별도 컬럼"],
        ["숫자 규격", "mm", "W/D/H는 숫자 셀; 단위는 컬럼명에 표시; 복수 규격은 02_규격_상세 시트"],
        [
            "옵션 분리",
            f"최대 {max_options}개",
            "옵션 스타일과 옵션 01~N을 별도 컬럼으로 분리; 복수 스타일·원문·옵션 ID는 03_옵션_상세 시트",
        ],
        ["요청 1 색상", "파란색", "요청 1 필드"],
        ["요청 2 색상", "초록색", "요청 2 필드"],
        ["빈 필드 색상", "분홍색", MISSING],
        ["비적용 표시", "회색", NOT_APPLICABLE],
    ]
    summary_row_groups = [
        "COMMON", "COMMON", "COMMON", "R1", "R2", "R1", "R1", "R1", "R1", "R2", "COMMON", "COMMON"
    ]

    common_headers = ["구분", "상품 ID", "상품명"]
    rolling_headers = [
        f"요청1_롤링 이미지 URL {index:02d}"
        for index in range(1, max_rolling_images + 1)
    ]
    option_headers = [
        f"요청1_옵션 {index:02d}"
        for index in range(1, max_options + 1)
    ]
    request1_headers = [
        "요청1_대표 이미지 URL",
        "요청1_롤링 이미지 수",
        *rolling_headers,
        "요청1_대표 규격 대상",
        "요청1_W (mm)",
        "요청1_D (mm)",
        "요청1_H (mm)",
        "요청1_규격 상태",
        "요청1_중카테고리",
        "요청1_소카테고리",
        "요청1_배치 추천 공간 리스트",
        "요청1_브랜드명",
        "요청1_제품 색상",
        "요청1_설치 타입 구분",
        "요청1_세트 구성 ID 리스트",
        "요청1_배치 가능 위치",
        "요청1_벽면부착 추천 높이 (mm)",
        "요청1_옵션 스타일",
        "요청1_옵션 수",
        *option_headers,
    ]
    request2_headers = [
        "요청2_설명서태그_기본",
        "요청2_설명서태그_제조",
        "요청2_설명서태그_재질",
        "요청2_설명서태그_규격",
        "요청2_설명서태그_구성품",
        "요청2_설명서태그_색상",
        "요청2_설명서태그_사용목적",
        "요청2_설명서태그_조립",
        "요청2_설명서태그_안전",
        "요청2_설명서태그_취급주의",
        "요청2_설명서태그_품질보증",
        "요청2_설명서태그_인증",
        "요청2_설명서태그_판매정보",
        "요청2_디자인 스타일 추론",
        "요청2_공간태그_인테리어스타일",
        "요청2_공간태그_분위기",
        "요청2_공간태그_공간명",
        "요청2_공간태그_공간목적",
        "요청2_공간태그_색상톤",
        "요청2_공간태그_공간크기",
        "요청2_공간태그_포함제품목록",
        "요청2_공간태그_배치사유·위치",
        "요청2_공간스타일↔제품 관계정보",
        "요청2_제품간 관계정보",
        "요청2_개인별 구매·선호·보유·CRM",
        "요청2_의미기반 검색·추천",
    ]
    headers = common_headers + request1_headers + request2_headers
    data_rows = [headers]
    dimension_headers = [
        "구분",
        "상품 ID",
        "상품명",
        "규격 순번",
        "규격 대상/옵션",
        "W (mm)",
        "D (mm)",
        "H (mm)",
        "규격 상태",
        "원천/비고",
    ]
    dimension_rows = [dimension_headers]
    option_detail_headers = [
        "구분",
        "상품 ID",
        "상품명",
        "옵션 스타일 순번",
        "옵션 스타일",
        "옵션 스타일 원문",
        "옵션 순번",
        "옵션 값",
        "옵션 ID",
        "옵션 이미지 URL",
        "원천/비고",
    ]
    option_detail_rows = [option_detail_headers]

    for row in source_rows:
        product_id = row["product_id"]
        deep_product = deep_by_id[product_id]
        is_home = deep_product["source_system"] == "HOMESTYLE"
        spaces = translated_list(row["recommended_spaces"], SPACE_MAP)
        placements = translated_list(row["placement_positions"], PLACEMENT_MAP)
        styles = translated_list(row["design_style_tags"], STYLE_MAP)
        purpose = translated_list(row["tag_use_purpose"], PURPOSE_MAP)
        installation = NOT_APPLICABLE if is_home else INSTALLATION_MAP.get(row["installation_type"], row["installation_type"] or MISSING)
        set_ids = (
            row["set_component_product_ids"] or MISSING
            if row["set_yn"] == "Y"
            else NOT_APPLICABLE
        )
        wall_height_raw = row["wall_mount_height_mm"]
        wall_height = numeric_value(wall_height_raw)
        if wall_height is None:
            wall_height = MISSING if "WALL" in row["placement_positions"].split("|") else NOT_APPLICABLE

        rolling_urls = rolling_by_id[product_id]
        representative_image = rolling_urls[0] if rolling_urls else row["pdp_image_url"] or MISSING
        rolling_cells = rolling_urls + [""] * (max_rolling_images - len(rolling_urls))
        dimension_list = dimensions_by_id[product_id]
        primary_dimension = dimension_list[0]
        option_group_list = options_by_id[product_id]
        option_styles = " | ".join(group["style"] for group in option_group_list)
        flattened_options = []
        for group in option_group_list:
            for item in group["items"]:
                option_name = item["name"]
                if len(option_group_list) > 1:
                    option_name = f"{group['style']}: {option_name}"
                flattened_options.append(option_name)
        option_cells = flattened_options + [NOT_APPLICABLE] * (max_options - len(flattened_options))

        request1 = [
            representative_image,
            len(rolling_urls),
            *rolling_cells,
            primary_dimension["target"],
            primary_dimension["w_mm"] if primary_dimension["w_mm"] is not None else "",
            primary_dimension["d_mm"] if primary_dimension["d_mm"] is not None else "",
            primary_dimension["h_mm"] if primary_dimension["h_mm"] is not None else "",
            primary_dimension["status"],
            row["mid_category"] or MISSING,
            row["small_category"] or MISSING,
            spaces,
            row["brand_name"] or MISSING,
            row["colors_raw"] or MISSING,
            installation,
            set_ids,
            placements,
            wall_height,
            option_styles or MISSING,
            len(flattened_options),
            *option_cells,
        ]

        if is_home:
            manufacturing = " | ".join(
                value
                for value in (
                    deep_value(deep_product, "manufacturer_importer"),
                    deep_value(deep_product, "origin_country"),
                )
                if value
            ) or MISSING
            specification = dimension_value(row, deep_product)
            assembly = {
                "G25070005743": "업체 방문 조립·설치",
                "G25100020496": NOT_APPLICABLE,
                "G25070001871": MISSING,
                "G25070006112": NOT_APPLICABLE,
            }[product_id]
            manual_tags = [
                row["tag_basic"] or MISSING,
                concise(manufacturing),
                concise(deep_value(deep_product, "materials")) or MISSING,
                specification,
                concise(deep_value(deep_product, "components")) or MISSING,
                concise(deep_value(deep_product, "colors")) or row["colors_raw"] or MISSING,
                purpose,
                assembly,
                safety_tag(product_id, deep_product),
                handling_tag(product_id, deep_product),
                concise(deep_value(deep_product, "warranty")) or MISSING,
                concise(deep_value(deep_product, "certification")) or MISSING,
                sales_tag(deep_product),
            ]
        else:
            manual_tags = [NOT_APPLICABLE] * 13

        relation = f"{styles} / {spaces} ↔ {product_id} (추론)"
        semantic = row["semantic_search_text"] or MISSING
        if row["semantic_synonyms"]:
            semantic += " | 동의어=" + row["semantic_synonyms"]
        request2 = manual_tags + [
            styles,
            styles + " (추론)",
            MOOD_BY_PRODUCT[product_id],
            spaces,
            purpose,
            row["colors_raw"] + " (제품 색상 기반 추론)" if row["colors_raw"] else MISSING,
            MISSING,
            product_id,
            f"제품 카테고리 기준 추천 | 위치={placements} (추론)",
            relation,
            MISSING,
            MISSING,
            semantic,
        ]
        data_rows.append([
            "홈스타일" if is_home else "가전",
            product_id,
            row["product_name"],
        ] + request1 + request2)

        for dimension_index, dimension in enumerate(dimension_list, start=1):
            dimension_rows.append(
                [
                    "홈스타일" if is_home else "가전",
                    product_id,
                    row["product_name"],
                    dimension_index,
                    dimension["target"],
                    dimension["w_mm"] if dimension["w_mm"] is not None else "",
                    dimension["d_mm"] if dimension["d_mm"] is not None else "",
                    dimension["h_mm"] if dimension["h_mm"] is not None else "",
                    dimension["status"],
                    dimension["evidence"],
                ]
            )

        for style_index, group in enumerate(option_group_list, start=1):
            for option_index, item in enumerate(group["items"], start=1):
                option_detail_rows.append(
                    [
                        "홈스타일" if is_home else "가전",
                        product_id,
                        row["product_name"],
                        style_index,
                        group["style"],
                        group["raw_style"] or NOT_APPLICABLE,
                        option_index,
                        item["name"],
                        item["option_id"] or MISSING,
                        item["image_url"] or NOT_APPLICABLE,
                        group["evidence"],
                    ]
                )

    column_groups = ["COMMON"] * len(common_headers) + ["R1"] * len(request1_headers) + ["R2"] * len(request2_headers)
    dimension_column_groups = ["COMMON"] * 3 + ["R1"] * (len(dimension_headers) - 3)
    option_detail_column_groups = ["COMMON"] * 3 + ["R1"] * (len(option_detail_headers) - 3)
    return (
        summary_rows,
        summary_row_groups,
        data_rows,
        column_groups,
        dimension_rows,
        dimension_column_groups,
        option_detail_rows,
        option_detail_column_groups,
    )


def patch_special_cell_styles(path: Path) -> None:
    with ZipFile(path) as source:
        files = {name: source.read(name) for name in source.namelist()}

    styles = files["xl/styles.xml"].decode("utf-8")
    styles = styles.replace('<fills count="9">', '<fills count="11">', 1)
    styles = styles.replace(
        "</fills>",
        '<fill><patternFill patternType="solid"><fgColor rgb="FFF4CCCC"/><bgColor indexed="64"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFE7E6E6"/><bgColor indexed="64"/></patternFill></fill>'
        "</fills>",
        1,
    )
    styles = styles.replace('<cellXfs count="9">', '<cellXfs count="11">', 1)
    special_xfs = (
        '<xf numFmtId="0" fontId="0" fillId="9" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="0" fillId="10" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>'
    )
    styles = styles.replace("</cellXfs>", special_xfs + "</cellXfs>", 1)
    files["xl/styles.xml"] = styles.encode("utf-8")

    for name in list(files):
        if not name.startswith("xl/worksheets/sheet") or not name.endswith(".xml"):
            continue
        root = ET.fromstring(files[name])
        first_row = root.find(".//" + M + "sheetData/" + M + "row")
        numeric_headers = {
            "요청1_W (mm)",
            "요청1_D (mm)",
            "요청1_H (mm)",
            "W (mm)",
            "D (mm)",
            "H (mm)",
        }
        numeric_columns: set[str] = set()
        if first_row is not None:
            for cell in first_row.findall(M + "c"):
                text = "".join(node.text or "" for node in cell.iter(M + "t"))
                if text in numeric_headers:
                    match = re.match(r"[A-Z]+", cell.attrib.get("r", ""))
                    if match:
                        numeric_columns.add(match.group(0))
        for cell in root.findall(".//" + M + "c"):
            text = "".join(node.text or "" for node in cell.iter(M + "t"))
            if text == MISSING:
                cell.set("s", "9")
            elif text == NOT_APPLICABLE:
                cell.set("s", "10")
            else:
                reference = cell.attrib.get("r", "")
                match = re.match(r"([A-Z]+)([0-9]+)", reference)
                value = cell.find(M + "v")
                if (
                    match
                    and int(match.group(2)) > 1
                    and match.group(1) in numeric_columns
                    and not text
                    and (value is None or not (value.text or "").strip())
                ):
                    cell.set("s", "9")
        files[name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    temp_path = path.with_suffix(".tmp.xlsx")
    with ZipFile(temp_path, "w", ZIP_DEFLATED) as target:
        for name, data in files.items():
            target.writestr(name, data)
    temp_path.replace(path)


def validate(path: Path) -> dict:
    with ZipFile(path) as archive:
        if archive.testzip() is not None:
            raise ValueError("corrupt XLSX archive")
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        sheet_count = len(list(workbook.find(M + "sheets")))
        missing_cells = 0
        not_applicable_cells = 0
        numeric_dimension_cells = 0
        rolling_image_columns = 0
        option_columns = 0
        option_detail_rows = 0
        for name in archive.namelist():
            if name.endswith(".xml"):
                root = ET.fromstring(archive.read(name))
                if name.startswith("xl/worksheets/"):
                    rows = root.findall(".//" + M + "sheetData/" + M + "row")
                    header_by_column: dict[str, str] = {}
                    if rows:
                        for cell in rows[0].findall(M + "c"):
                            match = re.match(r"[A-Z]+", cell.attrib.get("r", ""))
                            if match:
                                header_by_column[match.group(0)] = "".join(
                                    node.text or "" for node in cell.iter(M + "t")
                                )
                    rolling_image_columns += sum(
                        header.startswith("요청1_롤링 이미지 URL")
                        for header in header_by_column.values()
                    )
                    option_columns += sum(
                        re.fullmatch(r"요청1_옵션 \d{2}", header) is not None
                        for header in header_by_column.values()
                    )
                    if "옵션 스타일 순번" in header_by_column.values():
                        option_detail_rows += max(0, len(rows) - 1)
                    for cell in root.findall(".//" + M + "c"):
                        if cell.attrib.get("s") == "9":
                            missing_cells += 1
                        elif cell.attrib.get("s") == "10":
                            not_applicable_cells += 1
                        match = re.match(r"([A-Z]+)([0-9]+)", cell.attrib.get("r", ""))
                        if not match or int(match.group(2)) == 1:
                            continue
                        header = header_by_column.get(match.group(1), "")
                        if header in {
                            "요청1_W (mm)",
                            "요청1_D (mm)",
                            "요청1_H (mm)",
                            "W (mm)",
                            "D (mm)",
                            "H (mm)",
                        }:
                            value = cell.find(M + "v")
                            if value is not None and (value.text or "").strip():
                                if cell.attrib.get("t") != "n":
                                    raise ValueError(
                                        f"dimension cell is not numeric: {name}:{cell.attrib.get('r')}"
                                    )
                                numeric_dimension_cells += 1
    if sheet_count != 4:
        raise ValueError(f"unexpected sheet count: {sheet_count}")
    if rolling_image_columns < 1:
        raise ValueError("rolling image URL columns are missing")
    if numeric_dimension_cells < 1:
        raise ValueError("numeric W/D/H cells are missing")
    if option_columns < 1:
        raise ValueError("split option columns are missing")
    if option_detail_rows < 1:
        raise ValueError("option detail rows are missing")
    return {
        "sheets": sheet_count,
        "missing_cells": missing_cells,
        "not_applicable_cells": not_applicable_cells,
        "rolling_image_columns": rolling_image_columns,
        "numeric_dimension_cells": numeric_dimension_cells,
        "option_columns": option_columns,
        "option_detail_rows": option_detail_rows,
    }


def main() -> None:
    (
        summary_rows,
        summary_row_groups,
        data_rows,
        column_groups,
        dimension_rows,
        dimension_column_groups,
        option_detail_rows,
        option_detail_column_groups,
    ) = build_rows()
    data_widths = [14, 20, 42] + [
        54
        if "URL" in header
        else 16
        if header.endswith("(mm)") or header.endswith("이미지 수")
        else 28
        for header in data_rows[0][3:]
    ]
    xlsx_writer.OUTPUT = OUTPUT
    xlsx_writer.SHEETS = [
        ("00_요약", summary_rows, [22, 22, 110], ["COMMON"] * 3, summary_row_groups),
        ("01_상품별_요구필드", data_rows, data_widths, column_groups),
        (
            "02_규격_상세",
            dimension_rows,
            [14, 20, 42, 12, 28, 14, 14, 14, 16, 48],
            dimension_column_groups,
        ),
        (
            "03_옵션_상세",
            option_detail_rows,
            [14, 20, 42, 16, 22, 22, 12, 32, 16, 54, 48],
            option_detail_column_groups,
        ),
    ]
    xlsx_writer.build_xlsx()
    patch_special_cell_styles(OUTPUT)
    result = validate(OUTPUT)
    print(f"output={OUTPUT}")
    print(f"products={len(data_rows) - 1} fields={len(data_rows[0]) - 3} sheets={result['sheets']}")
    print(f"missing_cells={result['missing_cells']} not_applicable_cells={result['not_applicable_cells']}")
    print(
        f"rolling_image_columns={result['rolling_image_columns']} "
        f"numeric_dimension_cells={result['numeric_dimension_cells']}"
    )
    print(
        f"option_columns={result['option_columns']} "
        f"option_detail_rows={result['option_detail_rows']}"
    )


if __name__ == "__main__":
    main()
