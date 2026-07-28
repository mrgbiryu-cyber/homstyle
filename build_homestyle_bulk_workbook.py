from __future__ import annotations

import html as html_module
import json
import re
import sqlite3
from collections import Counter
from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET

import build_feasibility_sheet as xlsx_writer
from build_customer_requested_fields import patch_special_cell_styles, validate as base_validate
from bulk_homestyle_collect import DB_PATH, unpack


ROOT = Path(__file__).resolve().parent
OUTPUT = (
    ROOT
    / "홈스타일_비음영대상군_전체상품_요구필드_대량결과_패턴상태_규격품질보정_20mm이하재검증.xlsx"
)
MISSING = "미확보"
NOT_APPLICABLE = "해당없음"
MANDATORY_PASS = "PASS"
MANDATORY_REINFORCEMENT = "보강대상"
DIMENSION_STATUS_LABELS = {
    "SOURCE_CONFIRMED": "확정(DB)",
    "RULE_RESOLVED": "확정(규칙)",
    "MANUAL_CONFIRMED": "확정(수동)",
    "COMPARISON_PROVIDED": "비교정보 제공",
    "REVIEW_REQUIRED": "사람 확인 필요",
    "OCR_REQUIRED": "OCR 보강대상",
    "NO_CANDIDATE": "규격 후보 없음",
    "UNCLASSIFIED": "규격 원천 미분류",
}
XLSX_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
HUMAN_REVIEW_STYLE_ID = "11"
LOCKED_DIMENSION_STATUSES = {"SOURCE_CONFIRMED", "RULE_RESOLVED", "MANUAL_CONFIRMED"}
HUMAN_REVIEW_DIMENSION_LABELS = {
    "비교정보 제공",
    "사람 확인 필요",
    "OCR 보강대상",
    "규격 후보 없음",
    "규격 원천 미분류",
}
ATTACHMENT_0715_STATUS_LABELS = {
    "SOURCE_CONFIRMED": "API/HTML 확정",
    "RULE_RESOLVED": "OCR·규칙확정",
    "MANUAL_CONFIRMED": "OCR·규칙확정",
    "COMPARISON_PROVIDED": "비교정보제공",
    "REVIEW_REQUIRED": "비교정보제공",
    "OCR_REQUIRED": "최종무후보",
    "NO_CANDIDATE": "최종무후보",
    "UNCLASSIFIED": "최종무후보",
}
COMBINATION_STATUS_LABELS = {
    "CONFIRMED": "Y",
    "CANDIDATE": "검토필요",
    "NOT_COMBINATION": "N",
}
COMPONENT_OUTPUT_STATUS_LABELS = {
    "ALL_COMPONENT_DIMENSIONS_CONFIRMED": "구성품 규격 확정",
    "COMPONENT_DIMENSIONS_CANDIDATE": "구성품 규격 후보",
    "PARTIAL_COMPONENT_DIMENSIONS": "구성품 규격 부분확보",
    "COMPONENT_DIMENSIONS_MISSING": "구성품 규격 미확보",
    "COMPONENT_PARSE_REQUIRED": "구성품명 분리 필요",
    "COMBINATION_REVIEW_REQUIRED": "조합 여부 검토",
    "NOT_APPLICABLE": NOT_APPLICABLE,
}
COMPONENT_RESOLUTION_STATUS_LABELS = {
    "API_CONFIRMED": "API 확정",
    "API_UNIT_INFERRED": "단위 추론 후보",
    "DIMENSION_MISSING": "규격 미확보",
    "COMPONENT_NAME_REQUIRED": "구성품명 미확보",
    "COMBINATION_REVIEW_REQUIRED": "조합 여부 검토",
}
OUTPUT_STATUS_LABELS = {
    "SOURCE_CONFIRMED": "완료_확정값",
    "RULE_RESOLVED": "완료_확정값",
    "MANUAL_CONFIRMED": "완료_확정값",
    "COMPARISON_PROVIDED": "완료_RAW분리",
    "REVIEW_REQUIRED": "완료_RAW분리",
    "OCR_REQUIRED": "완료_무후보분류",
    "NO_CANDIDATE": "완료_무후보분류",
    "UNCLASSIFIED": "완료_무후보분류",
}
DIMENSION_PATTERN_NAMES = {
    "D00": "규격+조합 패턴 동시 확인",
    "D01": "완전 W/D/H 후보 1세트",
    "D02": "완전 W/D/H 1세트+추가 부분 후보",
    "D03": "완전 W/D/H 복수 후보",
    "D04": "부분 규격 후보 1세트",
    "D05": "부분 규격 복수 후보",
}
PATTERN_GROUP_ORDER = ("N", "S", "Q", "M", "A", "O")

SPACE_WORDS = {
    "거실": "리빙룸",
    "침실": "베드룸",
    "주방": "주방",
    "다이닝": "다이닝룸",
    "식당": "다이닝룸",
    "서재": "서재",
    "공부방": "서재",
    "아이방": "아이방",
    "현관": "현관",
    "욕실": "욕실",
    "베란다": "베란다",
    "테라스": "테라스",
    "야외": "야외",
}

STYLE_WORDS = {
    "미니멀": "미니멀",
    "모던": "모던",
    "컨템포러리": "컨템포러리",
    "스칸디나비안": "스칸디나비안",
    "북유럽": "스칸디나비안",
    "클래식": "클래식",
    "빈티지": "빈티지/레트로",
    "레트로": "빈티지/레트로",
    "내추럴": "내추럴",
    "우드": "내추럴",
    "인더스트리얼": "인더스트리얼",
    "곡선": "곡선형",
    "라운드": "곡선형",
    "아치": "곡선형",
    "직선": "직선형",
    "모듈": "모듈형",
    "장식": "장식형",
    "앤틱": "클래식",
}

COLOR_SYNONYMS = {
    "빨강": "레드|적색|붉은색",
    "레드": "빨강|적색|붉은색",
    "파랑": "블루|청색|푸른색",
    "블루": "파랑|청색|푸른색",
    "초록": "그린|녹색",
    "그린": "초록|녹색",
    "검정": "블랙|흑색",
    "블랙": "검정|흑색",
    "흰색": "화이트|백색",
    "화이트": "흰색|백색",
    "회색": "그레이|그레이색",
    "그레이": "회색|Gray|Grey",
    "베이지": "아이보리|크림",
    "아이보리": "베이지|크림",
    "갈색": "브라운|Brown",
    "브라운": "갈색|Brown",
}


def clean_text(value: Any, limit: int = 2000) -> str:
    if value in (None, ""):
        return ""
    text = str(value)
    text = re.sub(r"<script\b.*?</script>", " ", text, flags=re.I | re.S)
    text = re.sub(r"<style\b.*?</style>", " ", text, flags=re.I | re.S)
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"\s+", " ", html_module.unescape(text)).strip()
    return text if len(text) <= limit else text[:limit] + "…"


def normalize_url(value: str) -> str:
    value = html_module.unescape(str(value or "").strip())
    if value.startswith("//"):
        return "https:" + value
    if value.startswith("/goods/") or value.startswith("/banner/"):
        return "https://static-store.lge.co.kr" + value
    if value.startswith("/"):
        return "https://homestyle.lge.co.kr" + value
    return value


def notification_items(data: dict[str, Any]) -> list[dict[str, Any]]:
    raw = data.get("productNotification") or []
    if isinstance(raw, dict):
        if isinstance(raw.get("items"), list):
            return [item for item in raw["items"] if isinstance(item, dict)]
        return [raw] if "title" in raw or "description" in raw else []
    result = []
    for group in raw if isinstance(raw, list) else []:
        if not isinstance(group, dict):
            continue
        if isinstance(group.get("items"), list):
            result.extend(item for item in group["items"] if isinstance(item, dict))
        elif "title" in group or "description" in group:
            result.append(group)
    return result


def notification_values(items: list[dict[str, Any]], *keywords: str) -> list[str]:
    result = []
    for item in items:
        title = clean_text(item.get("title"), 200)
        if any(keyword.casefold() in title.casefold() for keyword in keywords):
            value = clean_text(item.get("description"))
            if value and value not in result:
                result.append(value)
    return result


def joined(values: list[str], default: str = MISSING, limit: int = 3000) -> str:
    text = " | ".join(value for value in values if value)
    if not text:
        return default
    return text if len(text) <= limit else text[:limit] + "…"


def has_required_value(value: Any) -> bool:
    return value not in (None, "", MISSING, NOT_APPLICABLE)


def mandatory_asset_assessment(
    *,
    system: str,
    product_id: str,
    brand: str,
    mid_category: str,
    small_category: str,
    representative_image_url: str,
    rolling_image_urls: list[str],
    color: str,
    w_mm: int | float | None,
    d_mm: int | float | None,
    h_mm: int | float | None,
    is_set: bool,
    set_component_ids: list[str],
    appliance_rear_image_url: str = "",
    appliance_additional: str = "",
) -> dict[str, Any]:
    """Apply the mandatory marks from the 2026-07-15 3D asset attachment.

    Per the agreed substitute rule, the image requirement is fulfilled when
    both the representative-image field and at least one rolling-image URL
    field are populated. Visual angle/background quality is not part of PASS.
    """
    checks: dict[str, bool | None] = {
        "브랜드": has_required_value(brand),
        "카테고리(중·소)": (
            has_required_value(mid_category) and has_required_value(small_category)
        ),
        "ID": has_required_value(product_id),
        "대표·롤링 이미지 URL": (
            has_required_value(representative_image_url)
            and any(has_required_value(value) for value in rolling_image_urls)
        ),
        "색상": has_required_value(color),
        "사이즈(W/D/H)": all(
            isinstance(value, (int, float))
            and not isinstance(value, bool)
            for value in (w_mm, d_mm, h_mm)
        ),
        "가전 후면 이미지": None,
        "가전 추가사항": None,
        "세트 구성 실제 ID": None,
    }
    if system == "가전":
        checks["가전 후면 이미지"] = has_required_value(appliance_rear_image_url)
        checks["가전 추가사항"] = has_required_value(appliance_additional)
    if is_set:
        checks["세트 구성 실제 ID"] = bool(set_component_ids)
    # Policy override (2026-07-22): set component IDs remain visible as a
    # reference/data-quality field, but do not participate in mandatory PASS.
    required = {
        name: value
        for name, value in checks.items()
        if value is not None and name != "세트 구성 실제 ID"
    }
    missing = [name for name, value in required.items() if not value]
    return {
        "status": MANDATORY_PASS if not missing else MANDATORY_REINFORCEMENT,
        "required_count": len(required),
        "fulfilled_count": len(required) - len(missing),
        "missing": missing,
        "checks": checks,
        "image_quality_status": (
            "URL 대체충족" if checks["대표·롤링 이미지 URL"] else MISSING
        ),
    }


def image_urls(data: dict[str, Any]) -> list[str]:
    rows = [row for row in (data.get("images") or []) if isinstance(row, dict)]
    rows.sort(key=lambda row: int(row.get("sortSeq") or 999999))
    result = []
    for row in rows:
        if row.get("type") not in (None, "", "IMAGE"):
            continue
        value = normalize_url(row.get("imageUrl") or "")
        if value.startswith(("http://", "https://")) and value not in result:
            result.append(value)
    return result


def normalize_option_style(raw_style: str, values: list[str], color_text: str) -> str:
    compact = re.sub(r"\s+", "", raw_style).casefold()
    if "사이즈" in compact or "size" in compact:
        return "사이즈"
    if "단일" in compact or (
        len(values) == 1 and re.sub(r"\s+", "", values[0]).casefold() in {"단일", "단일옵션"}
    ):
        return "단일옵션"
    colors = {
        re.sub(r"\s+", "", value).casefold()
        for value in re.split(r"[,/|]", color_text or "")
        if value.strip()
    }
    option_values = {
        re.sub(r"\s+", "", value).casefold() for value in values if value.strip()
    }
    if colors and option_values and option_values.issubset(colors):
        return "색상"
    if any(token in compact for token in ("색상", "컬러", "color", "colour")):
        return "색상"
    color_words = (
        "black|white|grey|gray|brown|beige|ivory|orange|red|blue|green|pink|yellow|purple|"
        "블랙|화이트|그레이|브라운|베이지|아이보리|오렌지|레드|블루|그린|핑크|옐로|퍼플|"
        "투명|크림|카멜|골드|실버"
    )
    if values and all(re.search(color_words, value, flags=re.I) for value in values):
        return "색상"
    return raw_style.strip() or "옵션"


def option_groups(data: dict[str, Any], color_text: str) -> list[dict[str, Any]]:
    stock_by_name = {
        str(row.get("optionName") or "").strip(): row
        for row in (data.get("productStock") or [])
        if isinstance(row, dict)
    }
    groups = []
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
            stock = stock_by_name.get(name) or {}
            items.append(
                {
                    "name": name,
                    "option_id": str(stock.get("optionId") or ""),
                    "image_url": normalize_url(raw_item.get("imageUrl") or ""),
                }
            )
        if items:
            groups.append(
                {
                    "style": normalize_option_style(raw_style, values, color_text),
                    "raw_style": raw_style,
                    "items": items,
                    "evidence": "goods API purchaseOptions + productStock",
                }
            )
    if groups:
        return groups
    stock_names = [name for name in stock_by_name if name]
    if stock_names:
        style = normalize_option_style("옵션", stock_names, color_text)
        return [
            {
                "style": style,
                "raw_style": "상품재고 옵션",
                "items": [
                    {
                        "name": name,
                        "option_id": str(stock_by_name[name].get("optionId") or ""),
                        "image_url": "",
                    }
                    for name in stock_names
                ],
                "evidence": "goods API productStock",
            }
        ]
    return [
        {
            "style": "단일옵션",
            "raw_style": "",
            "items": [{"name": "단일옵션", "option_id": "", "image_url": ""}],
            "evidence": "purchaseOptions/productStock 없음; 단일 SKU로 분류",
        }
    ]


def option_pattern_codes(groups: list[dict[str, Any]]) -> list[str]:
    """Classify API option groups using the O01~O06 review taxonomy."""
    codes: set[str] = set()
    color_words = (
        "black|white|grey|gray|brown|beige|ivory|orange|red|blue|green|pink|"
        "yellow|purple|graphite|mineral|cream|camel|gold|silver|deepgreen|"
        "mustard|burgundy|블랙|화이트|그레이|브라운|베이지|아이보리|"
        "오렌지|레드|블루|그린|핑크|크림|카멜|골드|실버|버건디|머스타드"
    )
    for group in groups:
        if group["style"] == "단일옵션":
            continue
        raw_style = str(group.get("raw_style") or group.get("style") or "")
        item_text = " | ".join(
            str(item.get("name") or "") for item in group.get("items") or []
        )
        raw_lower = raw_style.casefold()
        if (
            "사이즈" in raw_style
            or "size" in raw_lower
            or "사용인원" in raw_style
            or re.search(r"베이스\s*\d+(?:\.\d+)?\s*cm", raw_style, flags=re.I)
        ):
            codes.add("O01")
        quantities = set(
            re.findall(
                r"(?<!\d)(\d+(?:\.\d+)?)\s*(?:인|p|ea)(?![a-z])",
                item_text,
                flags=re.I,
            )
        )
        if len(quantities) >= 2:
            codes.add("O02")
        if (
            any(
                token in raw_lower
                for token in (
                    "옵션",
                    "선택",
                    "내부구성",
                    "스툴",
                    "오토만",
                    "extension",
                )
            )
            or (
                "프레임" in raw_style
                and ("유무" in raw_style or "하부" in raw_style)
            )
        ):
            codes.add("O03")
        if "방향" in raw_style or (
            "팔걸이" in raw_style
            and ("좌형" in item_text or "우형" in item_text)
        ):
            codes.add("O04")
        if "토퍼" in raw_style or "매트리스" in raw_style:
            codes.add("O05")
        if (
            group["style"] == "색상"
            or "패브릭" in raw_style
            or "fabric" in raw_lower
            or re.search(color_words, item_text, flags=re.I)
        ):
            codes.add("O06")
    return sorted(codes)


def combination_review_codes(
    legacy_pattern_code: str,
    product_name: str,
) -> list[str]:
    """Map the current 149-row legacy staging to N/S/Q/M/A review codes."""
    direct = {
        "N01_MODEL_SUFFIX_PLUS": ["N01"],
        "N02_OPTION_TYPE_JOIN": ["N04"],
        "N03_ARTWORK_TITLE_PLUS": ["N02"],
        "N04_COLOR_MATERIAL_JOIN": ["N03"],
        "N05_EXTENSION_FEATURE": ["N05"],
        "N06_INTEGRATED_FUNCTION": ["N06"],
        "R01_INTERNAL_CONFIGURATION": ["S06"],
        "Y02_BED_NIGHTSTAND": ["M04"],
        "Y03_BED_PANEL_GUARD": ["S05", "A02"],
        "Y06_KIDS_BED_ACCESSORY": ["S05", "A02"],
        "Y08_TABLE_TOP_LEG": ["S02"],
        "Y09_MODULAR_SOFA_UNITS": ["S03"],
        "Y10_SOFA_ACCESSORY": ["A01"],
        "Y11_CABINET_STORAGE_MODULES": ["S04", "M12"],
        "Y12_HANGER_ACCESSORY": ["A03"],
        "Y13_WARDROBE_MODULES": ["S04", "M12"],
        "Y14_CHAIR_BUNDLE_1PLUS1": ["Q01"],
        "Y15_DESK_ACCESSORY_PACK": ["A04"],
    }
    if legacy_pattern_code in direct:
        return direct[legacy_pattern_code]

    normalized_name = product_name.replace("＋", "+")
    if legacy_pattern_code == "Y01_MULTI_BED_BUNDLE":
        result = []
        if re.search(
            r"퀸\s*\+\s*퀸|슈퍼싱글\s*\+\s*슈퍼싱글",
            normalized_name,
            flags=re.I,
        ):
            result.append("Q02")
        if re.search(
            r"퀸\s*\+\s*슈퍼싱글|Q\s*\+\s*SS",
            normalized_name,
            flags=re.I,
        ):
            result.append("M03")
        if "협탁" in normalized_name:
            result.append("M04")
        return result or ["M03"]
    if legacy_pattern_code == "Y05_KIDS_BED_MATTRESS":
        result = ["S01", "A05"]
        if re.search(r"수납장|책상", normalized_name):
            result.append("M07")
        return result
    if legacy_pattern_code == "Y07_TABLE_CHAIR_BENCH":
        return ["M02"] if re.search(
            r"벤치|bench", normalized_name, flags=re.I
        ) else ["M01"]
    return []


def dimension_review_pattern_code(
    resolution_status: str,
    comparison_rows: list[dict[str, Any]],
    combination_status: str,
) -> str:
    if resolution_status != "COMPARISON_PROVIDED":
        return ""
    if combination_status == "CANDIDATE":
        return "D00"
    all_sets = {
        (row.get("w_mm"), row.get("d_mm"), row.get("h_mm"))
        for row in comparison_rows
    }
    full_sets = {
        values
        for values in all_sets
        if all(value is not None for value in values)
    }
    if len(full_sets) == 1 and len(all_sets) == 1:
        return "D01"
    if len(full_sets) == 1:
        return "D02"
    if len(full_sets) >= 2:
        return "D03"
    if len(all_sets) == 1:
        return "D04"
    return "D05"


def product_pattern_status(
    resolution_status: str,
    combination_status: str,
) -> str:
    dimension_check = resolution_status == "COMPARISON_PROVIDED"
    combination_check = combination_status == "CANDIDATE"
    if dimension_check and combination_check:
        return "체크필요_규격+조합패턴"
    if dimension_check:
        return "체크필요_규격패턴"
    if combination_check:
        return "체크필요_조합패턴"
    return "완료_패턴체크불필요"


def to_mm(value: str, unit: str) -> int | float:
    cleaned = value.replace(",", "").strip().rstrip(".")
    number = float(cleaned)
    if unit.casefold() == "cm":
        number *= 10
    return int(number) if number.is_integer() else number


def dimension_records(texts: list[tuple[str, str]]) -> list[dict[str, Any]]:
    records = []
    seen: set[tuple[Any, Any, Any]] = set()
    for source, raw_text in texts:
        text = clean_text(raw_text, 10000).replace("㎜", "mm").replace("×", "x")
        # OCR substitutions such as `)(` for multiplication and `I/l` for the
        # digit 1 are not normalized globally.  Long lineup charts can combine
        # numbers from different variants.  Those repairs are allowed only in
        # the product/option identity-validation layer, which writes an explicit
        # verified W/D/H evidence line back to the OCR source.

        def inferred_unit(values: list[str], explicit: str = "") -> str:
            if explicit:
                return explicit
            numbers = [float(value.replace(",", "").rstrip(".")) for value in values]
            # Furniture notices that omit a unit overwhelmingly use centimetres
            # for values below 300; larger values are normally millimetres.
            return "cm" if numbers and max(numbers) <= 300 else "mm"

        def append_record(
            w: int | float | None,
            d: int | float | None,
            h: int | float | None,
            status: str,
        ) -> None:
            # OCR strings such as `W1격00` used to be shortened to `00` and
            # silently accepted as W=0. A physical exterior dimension cannot
            # be zero or negative; retain any positive partial axes instead.
            if any(
                value is not None and float(value) <= 0
                for value in (w, d, h)
            ):
                return
            key = (w, d, h)
            if key in seen:
                return
            seen.add(key)
            records.append(
                {
                    "target": "제품/구성품",
                    "w_mm": w,
                    "d_mm": d,
                    "h_mm": h,
                    "status": status,
                    "evidence": source,
                    "raw": text,
                }
            )

        def range_value(first: str, second: str | None = None) -> str:
            # Use the maximum end of an adjustable range for the placement
            # bounding box; the untouched source stays available in raw.
            return second or first

        # Confirmed W/D/H labels in arbitrary order and within one short
        # dimension expression. SH/SD/AH are deliberately not exterior axes.
        axis_matches = list(
            re.finditer(
                r"(?<![A-Z])\(?\s*([WDH])\s*\)?\s*[:=]?\s*\(?\s*"
                r"(\d[\d,.]*)(?:\s*[-~]\s*(\d[\d,.]*))?\s*\)?\s*"
                r"\(?\s*(mm|cm)?\s*\)?",
                text,
                flags=re.I,
            )
        )
        for start_index, first_match in enumerate(axis_matches):
            cluster = []
            for match in axis_matches[start_index:]:
                if match.end() - first_match.start() > 180:
                    break
                cluster.append(match)
            axes = {match.group(1).upper() for match in cluster}
            if not {"W", "D", "H"}.issubset(axes):
                continue
            context = text[first_match.start() : cluster[-1].end() + 30]
            common_unit_match = re.search(r"(?:mm|cm)", context, flags=re.I)
            raw_values = [range_value(match.group(2), match.group(3)) for match in cluster]
            common_unit = common_unit_match.group(0) if common_unit_match else inferred_unit(raw_values)
            mapped: dict[str, int | float] = {}
            for match in cluster:
                axis = match.group(1).upper()
                if axis not in mapped:
                    mapped[axis] = to_mm(
                        range_value(match.group(2), match.group(3)),
                        match.group(4) or common_unit,
                    )
            append_record(mapped["W"], mapped["D"], mapped["H"], "확보")
            break

        # Confirmed Korean semantic axes and range notation.
        korean_axis_map = {
            "가로": "W", "너비": "W", "폭": "W",
            "깊이": "D", "세로": "D", "높이": "H",
        }
        korean_matches = list(
            re.finditer(
                r"(가로|너비|폭|깊이|세로|높이)\s*[:=]?\s*\(?\s*"
                r"(\d[\d,.]*)(?:\s*[-~]\s*(\d[\d,.]*))?\s*\)?\s*"
                r"\(?\s*(mm|cm)?\s*\)?",
                text,
                flags=re.I,
            )
        )
        for start_index, first_match in enumerate(korean_matches):
            cluster = []
            for match in korean_matches[start_index:]:
                if match.end() - first_match.start() > 180:
                    break
                cluster.append(match)
            axes = {korean_axis_map[match.group(1)] for match in cluster}
            if not {"W", "D", "H"}.issubset(axes):
                continue
            context = text[first_match.start() : cluster[-1].end() + 30]
            common_unit_match = re.search(r"(?:mm|cm)", context, flags=re.I)
            raw_values = [range_value(match.group(2), match.group(3)) for match in cluster]
            common_unit = common_unit_match.group(0) if common_unit_match else inferred_unit(raw_values)
            mapped: dict[str, int | float] = {}
            for match in cluster:
                axis = korean_axis_map[match.group(1)]
                if axis not in mapped:
                    mapped[axis] = to_mm(
                        range_value(match.group(2), match.group(3)),
                        match.group(4) or common_unit,
                    )
            append_record(mapped["W"], mapped["D"], mapped["H"], "확보")
            break

        # Circular bounding box: Dia/Ø/직경 plus H means W=D=diameter.
        diameter_matches = list(
            re.finditer(
                r"(?:DIA\.?|Ø|Φ|⌀|지름|직경)\s*[:=]?\s*\(?\s*"
                r"(\d[\d,.]*)(?:\s*[-~]\s*(\d[\d,.]*))?\s*\)?"
                r"(?:\s*\(\s*\d[\d,.]*\s*\))?\s*(mm|cm)?",
                text,
                flags=re.I,
            )
        )
        height_matches = list(
            re.finditer(
                r"(?:\bH|높이)\s*[:=]?\s*\(?\s*"
                r"(\d[\d,.]*)(?:\s*[-~]\s*(\d[\d,.]*))?\s*\)?\s*"
                r"\(?\s*(mm|cm)?\s*\)?",
                text,
                flags=re.I,
            )
        )
        circular_done = False
        for diameter_match in diameter_matches:
            for height_match in height_matches:
                if abs(diameter_match.start() - height_match.start()) > 140:
                    continue
                start = min(diameter_match.start(), height_match.start())
                end = max(diameter_match.end(), height_match.end())
                context = text[start : end + 20]
                common_unit_match = re.search(r"(?:mm|cm)", context, flags=re.I)
                raw_values = [
                    range_value(diameter_match.group(1), diameter_match.group(2)),
                    range_value(height_match.group(1), height_match.group(2)),
                ]
                common_unit = common_unit_match.group(0) if common_unit_match else inferred_unit(raw_values)
                diameter = to_mm(raw_values[0], diameter_match.group(3) or common_unit)
                height = to_mm(raw_values[1], height_match.group(3) or common_unit)
                append_record(diameter, diameter, height, "확보")
                circular_done = True
                break
            if circular_done:
                break

        # Circular products: diameter becomes both W and D.
        for match in re.finditer(
            r"(?:Ø|지름|직경|갓지름)\s*[:=]?\s*([\d,.]+)\s*(mm|cm)?"
            r"(?:\s*/\s*(?:Ø|지름|직경)?\s*([\d,.]+)\s*(mm|cm)?)?"
            r"\s*(?:x|\*)\s*H?\s*[:=]?\s*([\d,.]+)\s*(mm|cm)?",
            text,
            flags=re.I,
        ):
            height_unit = match.group(6) or match.group(4) or match.group(2)
            values = [match.group(1), match.group(5)]
            if match.group(3):
                values.append(match.group(3))
            shared_unit = inferred_unit(values, height_unit or "")
            height = to_mm(match.group(5), match.group(6) or shared_unit)
            for diameter_value, diameter_unit in (
                (match.group(1), match.group(2)),
                (match.group(3), match.group(4)),
            ):
                if diameter_value:
                    diameter = to_mm(diameter_value, diameter_unit or shared_unit)
                    append_record(diameter, diameter, height, "확보")

        # Height followed by a parenthesized diameter, common for floor lamps.
        for match in re.finditer(
            r"([\d,.]+)\s*(mm|cm)\s*\([^)]*(?:지름|직경)\s*[:=]?\s*"
            r"([\d,.]+)\s*(mm|cm)",
            text,
            flags=re.I,
        ):
            height = to_mm(match.group(1), match.group(2))
            diameter = to_mm(match.group(3), match.group(4))
            append_record(diameter, diameter, height, "확보")

        # Flexible 3-axis notation. Labels may be prefix/postfix, axes may be
        # ordered W-D-H or W-H-D, and the unit may appear only at the end.
        flexible_triple = re.compile(
            r"(?:\(?\s*(W|D|H)\s*\)?\s*[:=]?)?\s*"
            r"(?:\d[\d,.]*\s*[-~]\s*)?(\d[\d,.]*)\s*(?:\((W|D|H|L)\))?"
            r"(?:\s*\(\d[\d,.]*\))?\s*(mm|cm)?\s*"
            r"(?:x|\*|\s+(?=\(?\s*[WDH]\s*\)?))\s*"
            r"(?:\(?\s*(W|D|H)\s*\)?\s*[:=]?)?\s*"
            r"(?:\d[\d,.]*\s*[-~]\s*)?(\d[\d,.]*)\s*(?:\((W|D|H|L)\))?"
            r"(?:\s*\(\d[\d,.]*\))?\s*(mm|cm)?\s*"
            r"(?:x|\*|\s+(?=\(?\s*[WDH]\s*\)?))\s*"
            r"(?:\(?\s*(W|D|H)\s*\)?\s*[:=]?)?\s*"
            r"(?:\d[\d,.]*\s*[-~]\s*)?(\d[\d,.]*)\s*(?:\((W|D|H|L)\))?"
            r"(?:\s*\([A-Z]{1,3}\s*\d[\d,.]*\))?\s*\(?\s*(mm|cm)?\s*\)?",
            flags=re.I,
        )
        for match in flexible_triple.finditer(text):
            prefix = text[max(0, match.start() - 8) : match.start()]
            if re.search(r"\bL\s*$", prefix, flags=re.I):
                continue
            raw_values = [match.group(2), match.group(6), match.group(10)]
            labels = [
                (match.group(1) or match.group(3) or "").upper(),
                (match.group(5) or match.group(7) or "").upper(),
                (match.group(9) or match.group(11) or "").upper(),
            ]
            # L is deliberately deferred: its meaning depends on the vendor
            # and category coordinate convention. Keep only the raw evidence.
            if "L" in labels:
                continue
            units = [match.group(4), match.group(8), match.group(12)]
            shared_unit = inferred_unit(
                raw_values,
                next((unit for unit in reversed(units) if unit), ""),
            )
            numeric_values = [
                to_mm(value, unit or shared_unit)
                for value, unit in zip(raw_values, units)
            ]
            axes: dict[str, int | float] = {}
            unlabeled_values = []
            duplicate_label = False
            for label, value in zip(labels, numeric_values):
                if label:
                    if label in axes:
                        # W-D-W and W-D-D occur as obvious height-axis typos.
                        if "W" in axes and "D" in axes and "H" not in axes:
                            axes["H"] = value
                            continue
                        duplicate_label = True
                        break
                    axes[label] = value
                else:
                    unlabeled_values.append(value)
            if duplicate_label:
                continue
            for axis in ("W", "D", "H"):
                if axis not in axes and unlabeled_values:
                    axes[axis] = unlabeled_values.pop(0)
            if all(axis in axes for axis in ("W", "D", "H")):
                append_record(axes["W"], axes["D"], axes["H"], "확보")
        # Explicit two-axis notices must keep their semantic axes. This avoids
        # treating W x H wall objects as W x D in the generic 2D fallback.
        labeled_pair = re.compile(
            r"\(?\s*(W|D|H)\s*\)?\s*[:=]?\s*([\d,.]+)\s*(mm|cm)?\s*"
            r"(?:x|\*)\s*\(?\s*(W|D|H)\s*\)?\s*[:=]?\s*([\d,.]+)\s*(mm|cm)?",
            flags=re.I,
        )
        for match in labeled_pair.finditer(text):
            first_axis, second_axis = match.group(1).upper(), match.group(4).upper()
            if first_axis == second_axis:
                continue
            unit = match.group(6) or match.group(3) or "mm"
            axes = {
                first_axis: to_mm(match.group(2), match.group(3) or unit),
                second_axis: to_mm(match.group(5), match.group(6) or unit),
            }
            append_record(axes.get("W"), axes.get("D"), axes.get("H"), "부분확보")

        # Explicit W/D/H or Korean width/depth/height labels.
        labeled_patterns = [
            re.compile(
                r"W(?:IDTH)?\s*[:=]?\s*([\d,.]+)\s*(mm|cm)?\D{0,30}?"
                r"D(?:EPTH)?\s*[:=]?\s*([\d,.]+)\s*(mm|cm)?\D{0,30}?"
                r"H(?:EIGHT)?\s*[:=]?\s*([\d,.]+)\s*(mm|cm)?",
                re.I,
            ),
            re.compile(
                r"(?:가로|폭)\s*[:=]?\s*([\d,.]+)\s*(mm|cm)?\D{0,30}?"
                r"(?:세로|깊이)\s*[:=]?\s*([\d,.]+)\s*(mm|cm)?\D{0,30}?"
                r"(?:높이)\s*[:=]?\s*([\d,.]+)\s*(mm|cm)?",
                re.I,
            ),
        ]
        for pattern in labeled_patterns:
            for match in pattern.finditer(text):
                units = [match.group(2), match.group(4), match.group(6)]
                fallback_unit = next((unit for unit in reversed(units) if unit), "mm")
                w = to_mm(match.group(1), units[0] or fallback_unit)
                d = to_mm(match.group(3), units[1] or fallback_unit)
                h = to_mm(match.group(5), units[2] or fallback_unit)
                key = (w, d, h)
                append_record(w, d, h, "확보")

        # Ordered 3D dimensions with a shared unit.
        for match in re.finditer(
            r"(?<![\d.])([\d,.]{2,})\s*(?:mm|cm)?\s*x\s*"
            r"([\d,.]{2,})\s*(?:mm|cm)?\s*x\s*"
            r"([\d,.]{2,})\s*(mm|cm)(?![a-z])",
            text,
            flags=re.I,
        ):
            unit = match.group(4)
            key = tuple(to_mm(match.group(index), unit) for index in (1, 2, 3))
            if key not in seen:
                append_record(key[0], key[1], key[2], "확보")

        # 2D size options such as rugs; height remains intentionally blank.
        for match in re.finditer(
            r"(?<![\d.])([\d,.]{2,})\s*(mm|cm)?\s*x\s*([\d,.]{2,})\s*(mm|cm)(?![a-z])",
            text,
            flags=re.I,
        ):
            unit = match.group(4) or match.group(2) or "mm"
            key = (to_mm(match.group(1), match.group(2) or unit), to_mm(match.group(3), unit), None)
            if key not in seen:
                seen.add(key)
                records.append(
                    {
                        "target": "제품/옵션",
                        "w_mm": key[0],
                        "d_mm": key[1],
                        "h_mm": None,
                        "status": "부분확보",
                        "evidence": source,
                        "raw": text,
                    }
                )
    return records


def verified_dimension_texts(ocr_blob: dict[str, Any]) -> list[tuple[str, str]]:
    """Materialize identity-validated OCR/option dimensions before raw OCR."""
    result: list[tuple[str, str]] = []
    for item in ocr_blob.get("dimension_reinforcements") or []:
        run_name = str(item.get("run_name") or "")
        source = (
            "검증된 옵션+OCR 규격"
            if run_name == "option_single_width_plus_ocr"
            else "검증된 추가 이미지 OCR 규격"
        )
        dimensions = item.get("dimensions") or []
        if not dimensions and all(item.get(key) is not None for key in ("w_mm", "d_mm", "h_mm")):
            dimensions = [[item["w_mm"], item["d_mm"], item["h_mm"]]]
        for values in dimensions:
            if len(values) != 3 or any(value is None for value in values):
                continue
            context = clean_text(item.get("evidence_text"), 1000)
            dimension_line = (
                f"W={values[0]} mm D={values[1]} mm H={values[2]} mm"
            )
            result.append(
                (
                    source,
                    f"{context} | {dimension_line}" if context else dimension_line,
                )
            )
    return result


def dimension_keyword_snippets(value: Any, limit: int = 40) -> list[str]:
    text = clean_text(value, 120000).replace("㎜", "mm")
    snippets = []
    patterns = (
        r"(?:크기|치수|규격|사이즈|가로|세로|높이|깊이|폭|너비)",
        r"(?:\bW\s*\d|\bD\s*\d|\bH\s*\d)",
    )
    for pattern in patterns:
        for match in re.finditer(pattern, text, flags=re.I):
            snippet = text[max(0, match.start() - 120) : match.start() + 500]
            if any(char.isdigit() for char in snippet) and snippet not in snippets:
                snippets.append(snippet)
            if len(snippets) >= limit:
                return snippets
    return snippets


def category_spaces(large: str, mid: str, small: str) -> list[str]:
    text = f"{large} {mid} {small}"
    result = []
    rules = (
        (("침대", "매트리스", "옷장", "행거", "화장대"), "베드룸"),
        (("소파", "거실장", "거실테이블", "러그"), "리빙룸"),
        (("식탁", "다이닝"), "다이닝룸"),
        (("책상", "책장", "오피스"), "서재"),
        (("유아", "아동", "키즈"), "아이방"),
        (("아웃도어", "야외"), "테라스"),
        (("조명", "액자", "거울", "소품"), "리빙룸"),
    )
    for keywords, space in rules:
        if any(keyword in text for keyword in keywords) and space not in result:
            result.append(space)
    return result or ["리빙룸"]


def placement_values(large: str, mid: str, small: str) -> list[str]:
    text = f"{large} {mid} {small}"
    if any(word in text for word in ("팬던트", "샹들리에", "천장", "실링")):
        return ["천장"]
    if any(word in text for word in ("벽등", "벽조명", "벽걸이", "벽거울", "벽장식")):
        return ["벽"]
    if any(word in text for word in ("테이블조명", "스탠드조명", "액자", "소품")):
        return ["테이블·선반 위"]
    return ["바닥"]


def space_context(space_payload: dict[str, Any], fallback_spaces: list[str]) -> dict[str, str]:
    data = space_payload.get("data") or {}
    collections = data.get("collections") or []
    spaces = []
    styles = []
    moods = []
    content_products = []
    placement_reasons = []
    relationships = []
    for collection in collections[:20]:
        if not isinstance(collection, dict):
            continue
        labels = [
            clean_text(tag.get("tagLabel"), 100)
            for tag in (collection.get("hashtags") or [])
            if isinstance(tag, dict)
        ]
        collection_text = " ".join(
            [
                clean_text(collection.get("title"), 300),
                clean_text(collection.get("subTitle"), 500),
                clean_text(collection.get("category"), 200),
                *labels,
            ]
        )
        for word, normalized in SPACE_WORDS.items():
            if word in collection_text and normalized not in spaces:
                spaces.append(normalized)
        for word, normalized in STYLE_WORDS.items():
            if word.casefold() in collection_text.casefold() and normalized not in styles:
                styles.append(normalized)
        title = clean_text(collection.get("title"), 300)
        subtitle = clean_text(collection.get("subTitle"), 500)
        if title and title not in moods:
            moods.append(title)
        if subtitle and subtitle not in placement_reasons:
            placement_reasons.append(subtitle)
        product_ids = [
            str(product.get("productId") or "")
            for product in (collection.get("products") or [])
            if isinstance(product, dict) and product.get("productId")
        ]
        for product_id in product_ids:
            if product_id not in content_products:
                content_products.append(product_id)
        if product_ids:
            relationships.append("↔".join(product_ids[:20]))
    return {
        "spaces": "|".join(spaces or fallback_spaces),
        "styles": "|".join(styles),
        "moods": joined(moods, MISSING, 1000),
        "content_products": "|".join(content_products[:100]) or MISSING,
        "placement_reasons": joined(placement_reasons, MISSING, 1200),
        "relationships": joined(relationships, MISSING, 2000),
        "collection_count": str(len(collections)),
    }


def style_tags(name: str, detail_text: str, space_styles: str) -> str:
    result = [value for value in space_styles.split("|") if value]
    text = f"{name} {detail_text[:3000]}"
    for word, normalized in STYLE_WORDS.items():
        if word.casefold() in text.casefold() and normalized not in result:
            result.append(normalized)
    if not result:
        result.append("컨템포러리 (카테고리 기반 추론)")
    return "|".join(result)


def qna_signal(qna: dict[str, Any], html: dict[str, Any], keywords: tuple[str, ...]) -> str:
    records = list(qna.get("records") or []) + list(html.get("faq_records") or [])
    snippets = []
    for row in records:
        text = clean_text(
            " ".join(
                str(row.get(key) or "")
                for key in ("question_title", "question_text", "answer_text", "question", "answer")
            ),
            1000,
        )
        if text and any(keyword.casefold() in text.casefold() for keyword in keywords):
            if text not in snippets:
                snippets.append(text)
        if len(snippets) >= 3:
            break
    return joined(snippets, "", 1500)


def recursive_product_ids(value: Any, self_id: str) -> list[str]:
    result = []
    if isinstance(value, dict):
        for key, child in value.items():
            if key.casefold() in {"productid", "goodsid", "goodsno"}:
                product_id = str(child or "").strip()
                if product_id and product_id != self_id and product_id not in result:
                    result.append(product_id)
            else:
                for product_id in recursive_product_ids(child, self_id):
                    if product_id not in result:
                        result.append(product_id)
    elif isinstance(value, list):
        for child in value:
            for product_id in recursive_product_ids(child, self_id):
                if product_id not in result:
                    result.append(product_id)
    return result


def category_record(connection: sqlite3.Connection, scope_ids: list[str]) -> dict[str, str]:
    if not scope_ids:
        return {"large": "", "mid": "", "small": ""}
    rows = connection.execute(
        f"SELECT large_name, mid_name, small_name FROM categories WHERE scope_id IN ({','.join('?' for _ in scope_ids)})",
        scope_ids,
    ).fetchall()
    return {
        "large": "|".join(dict.fromkeys(row[0] or "" for row in rows if row[0])),
        "mid": "|".join(dict.fromkeys(row[1] or "" for row in rows if row[1])),
        "small": "|".join(dict.fromkeys(row[2] or "" for row in rows if row[2])),
    }


def extract_products() -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    connection = sqlite3.connect(DB_PATH)
    rows = connection.execute(
        """
        SELECT p.product_id, p.category_scope_ids, p.listing_blob,
               s.goods_status, s.goods_blob, s.space_status, s.space_blob,
               s.packages_status, s.packages_blob, s.qna_status, s.qna_blob,
               s.html_status, s.html_blob, s.ocr_status, s.ocr_blob
        FROM products p JOIN sources s ON s.product_id=p.product_id
        ORDER BY p.product_id
        """
    ).fetchall()
    products = []
    errors = []
    for row in rows:
        product_id = row[0]
        scope_ids = json.loads(row[1] or "[]")
        category = category_record(connection, scope_ids)
        listing = unpack(row[2]) or {}
        goods_payload = unpack(row[4]) or {}
        data = goods_payload.get("data") or {}
        if row[3] != 200 or not data:
            errors.append(
                {
                    "product_id": product_id,
                    "goods_status": row[3] or 0,
                    "space_status": row[5] or 0,
                    "packages_status": row[7] or 0,
                    "qna_status": row[9] or 0,
                    "html_status": row[11] or 0,
                    "ocr_status": row[13] or 0,
                    "note": clean_text(goods_payload.get("message") or goods_payload.get("error")),
                }
            )
            continue
        products.append(
            {
                "product_id": product_id,
                "scope_ids": scope_ids,
                "category_scope": category,
                "listing": listing,
                "data": data,
                "space": unpack(row[6]) or {},
                "packages": unpack(row[8]) or {},
                "qna": unpack(row[10]) or {},
                "html": unpack(row[12]) or {},
                "ocr": unpack(row[14]) or {},
                "statuses": {
                    "goods": row[3] or 0,
                    "space": row[5] or 0,
                    "packages": row[7] or 0,
                    "qna": row[9] or 0,
                    "html": row[11] or 0,
                    "ocr": row[13] or 0,
                },
            }
        )
    connection.close()
    return products, errors


def load_dimension_resolution_export() -> tuple[
    dict[str, dict[str, Any]],
    dict[str, list[dict[str, Any]]],
    dict[str, Any],
    list[dict[str, Any]],
]:
    """Load the authoritative dimension ledger and comparison candidates."""
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    ledgers = {
        row["product_id"]: dict(row)
        for row in connection.execute(
            "SELECT * FROM vw_dimension_resolution_ledger_current"
        )
    }
    comparisons: dict[str, list[dict[str, Any]]] = {}
    for row in connection.execute(
        """
        SELECT *
        FROM vw_dimension_comparison_candidates_current
        ORDER BY product_id, comparison_no
        """
    ):
        comparisons.setdefault(row["product_id"], []).append(dict(row))
    progress_row = connection.execute(
        "SELECT * FROM vw_dimension_progress_authoritative"
    ).fetchone()
    progress = dict(progress_row) if progress_row else {}
    reason_view_exists = connection.execute(
        """
        SELECT COUNT(*)
        FROM sqlite_master
        WHERE type = 'view'
          AND name = 'vw_dimension_remaining_reason_summary'
        """
    ).fetchone()[0]
    remaining_reasons = (
        [
            dict(row)
            for row in connection.execute(
                "SELECT * FROM vw_dimension_remaining_reason_summary"
            )
        ]
        if reason_view_exists
        else []
    )
    connection.close()
    return ledgers, comparisons, progress, remaining_reasons


def load_product_component_export() -> tuple[
    dict[str, dict[str, Any]],
    list[dict[str, Any]],
    dict[str, int],
    dict[str, int],
]:
    """Load the current product-combination and component-dimension snapshot."""
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    required_views = {
        row[0]
        for row in connection.execute(
            """
            SELECT name
            FROM sqlite_master
            WHERE type='view'
              AND name IN (
                  'vw_product_combination_current',
                  'vw_product_component_dimensions_current'
              )
            """
        )
    }
    expected_views = {
        "vw_product_combination_current",
        "vw_product_component_dimensions_current",
    }
    if required_views != expected_views:
        connection.close()
        raise RuntimeError(
            "조합상품 스테이징이 없습니다. "
            "먼저 python build_product_component_staging.py 를 실행하세요."
        )
    combinations = {
        row["product_id"]: dict(row)
        for row in connection.execute(
            "SELECT * FROM vw_product_combination_current"
        )
    }
    components = [
        dict(row)
        for row in connection.execute(
            """
            SELECT *
            FROM vw_product_component_dimensions_current
            ORDER BY product_id, component_seq
            """
        )
    ]
    detection_counts = {
        row["detection_status"]: int(row["product_count"])
        for row in connection.execute(
            """
            SELECT detection_status,COUNT(*) AS product_count
            FROM vw_product_combination_current
            GROUP BY detection_status
            """
        )
    }
    output_counts = {
        row["component_output_status"]: int(row["product_count"])
        for row in connection.execute(
            """
            SELECT component_output_status,COUNT(*) AS product_count
            FROM vw_product_combination_current
            GROUP BY component_output_status
            """
        )
    }
    connection.close()
    return combinations, components, detection_counts, output_counts


def load_combination_review_patterns() -> dict[str, dict[str, Any]]:
    """Load the non-final 149-product pattern staging without changing it."""
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    exists = connection.execute(
        """
        SELECT COUNT(*)
        FROM sqlite_master
        WHERE type='view'
          AND name='vw_combination_candidate_pattern_current'
        """
    ).fetchone()[0]
    if not exists:
        connection.close()
        raise RuntimeError(
            "조합 패턴 스테이징이 없습니다. "
            "먼저 python build_combination_candidate_pattern_staging.py 를 실행하세요."
        )
    result = {
        row["product_id"]: dict(row)
        for row in connection.execute(
            "SELECT * FROM vw_combination_candidate_pattern_current"
        )
    }
    connection.close()
    return result


def compact_dimension_number(value: Any, empty_value: str = MISSING) -> str:
    if value is None:
        return empty_value
    number = float(value)
    if number.is_integer():
        return str(int(number))
    return f"{number:.3f}".rstrip("0").rstrip(".")


def comparison_cells(rows: list[dict[str, Any]]) -> list[Any]:
    """Return positionally aligned comma-separated comparison columns."""
    if not rows:
        return [0, NOT_APPLICABLE, NOT_APPLICABLE, NOT_APPLICABLE,
                NOT_APPLICABLE, NOT_APPLICABLE]

    def joined_values(values: list[str]) -> str:
        return ", ".join(value.replace(",", " / ") for value in values)

    return [
        len(rows),
        joined_values([compact_dimension_number(row.get("w_mm")) for row in rows]),
        joined_values(
            [
                compact_dimension_number(row.get("d_mm"), NOT_APPLICABLE)
                for row in rows
            ]
        ),
        joined_values([compact_dimension_number(row.get("h_mm")) for row in rows]),
        joined_values(
            [
                clean_text(row.get("comparison_target"), 100) or "규격 후보"
                for row in rows
            ]
        ),
        joined_values(
            [
                (
                    clean_text(row.get("raw_notation"), 180)
                    or clean_text(row.get("context_text"), 180)
                    or MISSING
                )
                for row in rows
            ]
        ),
    ]


def build_rows() -> tuple[list[tuple], dict[str, Any]]:
    products, errors = extract_products()
    (
        resolution_ledgers,
        dimension_comparisons,
        dimension_progress,
        dimension_remaining_reasons,
    ) = (
        load_dimension_resolution_export()
    )
    (
        product_combinations,
        product_components,
        combination_detection_counts,
        component_output_counts,
    ) = load_product_component_export()
    combination_review_patterns = load_combination_review_patterns()
    if len(product_combinations) != len(products) + len(errors):
        raise RuntimeError(
            "조합상품 스테이징 상품 수 불일치: "
            f"DB={len(product_combinations)}, 대상={len(products) + len(errors)}"
        )
    expected_combination_candidates = combination_detection_counts.get(
        "CANDIDATE", 0
    )
    if len(combination_review_patterns) != expected_combination_candidates:
        raise RuntimeError(
            "조합 패턴 스테이징 상품 수 불일치: "
            f"DB={len(combination_review_patterns)}, "
            f"후보={expected_combination_candidates}"
        )
    parsed = []
    max_images = 1
    max_options = 1
    dimension_stats = Counter()
    ocr_impact = Counter()
    for product in products:
        data = product["data"]
        notifications = notification_items(data)
        color_values = notification_values(notifications, "색상", "컬러")
        colors = joined(color_values, "")
        groups = option_groups(data, colors)
        if not colors:
            color_options = [
                item["name"]
                for group in groups
                if group["style"] == "색상"
                for item in group["items"]
            ]
            colors = "|".join(color_options)
        images = image_urls(data)
        dimension_texts = [
            ("상품정보고시", value)
            for value in notification_values(notifications, "크기", "치수", "규격", "사이즈")
        ]
        for group in groups:
            if group["style"] == "사이즈":
                dimension_texts.extend(("상품 옵션", item["name"]) for item in group["items"])
        html_dimension_signals = product["html"].get("dimension_signals") or []
        dimension_texts.extend(("PDP HTML", value) for value in html_dimension_signals)
        dimension_texts.extend(
            ("상품 상세 HTML", value)
            for value in dimension_keyword_snippets(data.get("detailInfo"))
        )
        dimension_texts.append(("상품명", str(data.get("productName") or "")))
        qna_dimension = qna_signal(
            product["qna"],
            product["html"],
            ("크기", "사이즈", "가로", "세로", "높이", "폭", "치수"),
        )
        if qna_dimension:
            dimension_texts.append(("FAQ/Q&A", qna_dimension))
        base_dimensions = dimension_records(dimension_texts)
        dimension_texts.extend(verified_dimension_texts(product["ocr"]))
        ocr_dimension_text = str(product["ocr"].get("dimension_text") or "")
        if ocr_dimension_text:
            dimension_texts.append(("상세 이미지 OCR", ocr_dimension_text))
        dimensions = dimension_records(dimension_texts)
        base_complete = any(
            row.get("w_mm") is not None
            and row.get("d_mm") is not None
            and row.get("h_mm") is not None
            for row in base_dimensions
        )
        final_complete = any(
            row.get("w_mm") is not None
            and row.get("d_mm") is not None
            and row.get("h_mm") is not None
            for row in dimensions
        )
        if not base_complete and final_complete:
            ocr_impact["complete"] += 1
        elif ocr_dimension_text and len(dimensions) > len(base_dimensions):
            ocr_impact["partial"] += 1
        if final_complete:
            dimension_stats["complete"] += 1
        elif dimensions:
            dimension_stats["partial"] += 1
        else:
            dimension_stats["missing"] += 1
        if not dimensions:
            dimensions = [
                {
                    "target": "제품 외형",
                    "w_mm": None,
                    "d_mm": None,
                    "h_mm": None,
                    "status": MISSING,
                    "evidence": "API/HTML/FAQ·Q&A에서 숫자 규격 미확보",
                    "raw": "",
                }
            ]
        max_images = max(max_images, len(images))
        max_options = max(max_options, sum(len(group["items"]) for group in groups))
        parsed.append(
            {
                **product,
                "notifications": notifications,
                "colors": colors or MISSING,
                "groups": groups,
                "images": images,
                "dimensions": dimensions,
                "dimension_resolution": resolution_ledgers.get(
                    product["product_id"], {}
                ),
                "dimension_comparisons": dimension_comparisons.get(
                    product["product_id"], []
                ),
            }
        )

    common_headers = ["구분", "상품 ID", "상품명"]
    pattern_headers = [
        "산출 상태",
        "패턴 상태",
        "규격 패턴 코드",
        "규격 패턴명",
        "조합 패턴군",
        "조합 세부 패턴 코드",
        "조합 원천 패턴 코드",
        "패턴 확인 순서",
        "상품 페이지 URL",
    ]
    rolling_headers = [
        f"요청1_롤링 이미지 URL {index:02d}" for index in range(1, max_images + 1)
    ]
    option_headers = [f"요청1_옵션 {index:02d}" for index in range(1, max_options + 1)]
    request1_headers = [
        "별첨0715_필수값 판정",
        "별첨0715_필수값 충족수",
        "별첨0715_필수값 대상수",
        "별첨0715_보강대상 필드",
        "요청1_대표 이미지 URL",
        "요청1_롤링 이미지 수",
        *rolling_headers,
        "요청1_대표 규격 대상",
        "요청1_W (mm)",
        "요청1_D (mm)",
        "요청1_H (mm)",
        "요청1_규격 비교후보 수",
        "요청1_규격 비교 W (mm)",
        "요청1_규격 비교 D (mm)",
        "요청1_규격 비교 H (mm)",
        "요청1_규격 비교 대상/옵션",
        "요청1_규격 비교 원문",
        "요청1_규격 상태",
        "요청1_중카테고리",
        "요청1_소카테고리",
        "요청1_배치 추천 공간 리스트",
        "요청1_브랜드명",
        "요청1_제품 색상",
        "요청1_설치 타입 구분",
        "요청1_세트 구성 ID 리스트",
        "요청1_조합상품 여부",
        "요청1_조합상품 판정근거",
        "요청1_조합 구성품 수",
        "요청1_구성품 규격 확정 수",
        "요청1_구성품 분리 상태",
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
    headers = common_headers + pattern_headers + request1_headers + request2_headers
    main_rows = [headers]
    pattern_rows = [[
        "구분",
        "상품 ID",
        "상품명",
        *pattern_headers,
    ]]
    dimension_rows = [[
        "구분", "상품 ID", "상품명", "규격 순번", "규격 대상/옵션",
        "W (mm)", "D (mm)", "H (mm)", "규격 상태", "원천/비고", "원문",
    ]]
    option_rows = [[
        "구분", "상품 ID", "상품명", "옵션 스타일 순번", "옵션 스타일",
        "옵션 스타일 원문", "옵션 순번", "옵션 값", "옵션 ID",
        "옵션 이미지 URL", "원천/비고",
    ]]
    component_rows = [[
        "구분", "상품 ID", "상품명", "조합상품 판정", "구성품 순번",
        "구성품명", "구성품 유형", "구성품 수량", "대표 구성품",
        "W (mm)", "D (mm)", "H (mm)", "구성품 규격 상태",
        "원천 유형", "원천 위치", "근거 원문", "사람 확인 필요",
    ]]
    mandatory_check_names = [
        "브랜드", "카테고리(중·소)", "ID", "대표·롤링 이미지 URL",
        "색상", "사이즈(W/D/H)", "가전 후면 이미지", "가전 추가사항",
        "세트 구성 실제 ID",
    ]
    mandatory_rows = [[
        "구분", "상품 ID", "상품명", "필수값 판정", "충족수", "대상수",
        "보강대상 필드", "이미지 필수 대체판정", *mandatory_check_names,
    ]]

    availability = {"R1": Counter(), "R2": Counter()}
    mandatory_stats = Counter()
    mandatory_missing_fields = Counter()
    request1_complete_count = 0
    request1_review_count = 0
    output_status_counts: Counter[str] = Counter()
    pattern_status_counts: Counter[str] = Counter()
    dimension_pattern_counts: Counter[str] = Counter()
    combination_pattern_group_counts: Counter[str] = Counter()
    for product in parsed:
        data = product["data"]
        product_id = product["product_id"]
        notifications = product["notifications"]
        category = data.get("category") or {}
        large = category.get("superCategoryName") or product["category_scope"]["large"]
        mid = category.get("categoryName") or product["category_scope"]["mid"]
        small = category.get("subCategoryName") or product["category_scope"]["small"]
        name = clean_text(data.get("productName"), 500) or product_id
        brand = clean_text(data.get("brandName"), 300) or MISSING
        images = product["images"]
        rolling_cells = images + [NOT_APPLICABLE] * (max_images - len(images))
        # Prefer the most complete record for the product-level W/D/H fields.
        # The previous first-record rule could expose a partial option record
        # even when a later source contained complete exterior dimensions.
        primary_dimension = dict(max(
            product["dimensions"],
            key=lambda record: (
                sum(record.get(axis) is not None for axis in ("w_mm", "d_mm", "h_mm")),
                int("검증된" in str(record.get("evidence") or "")),
            ),
        ))
        resolution = product["dimension_resolution"]
        resolution_status = resolution.get("resolution_status")
        if resolution:
            if int(resolution.get("is_locked") or 0) == 1:
                primary_dimension["w_mm"] = resolution.get("locked_w_mm")
                primary_dimension["d_mm"] = resolution.get("locked_d_mm")
                primary_dimension["h_mm"] = resolution.get("locked_h_mm")
                primary_dimension["target"] = "대표 규격"
            else:
                primary_dimension["w_mm"] = resolution.get("candidate_w_mm")
                primary_dimension["d_mm"] = resolution.get("candidate_d_mm")
                primary_dimension["h_mm"] = resolution.get("candidate_h_mm")
                if resolution_status == "COMPARISON_PROVIDED":
                    primary_dimension["target"] = "비교 대표후보"
            primary_dimension["status"] = DIMENSION_STATUS_LABELS.get(
                str(resolution_status), str(resolution_status or MISSING)
            )
            primary_dimension["evidence"] = (
                clean_text(resolution.get("evidence_text"), 1000)
                or primary_dimension.get("evidence")
                or ""
            )
        dimension_comparison_cells = comparison_cells(
            product["dimension_comparisons"]
            if resolution_status == "COMPARISON_PROVIDED"
            else []
        )
        fallback_spaces = category_spaces(large, mid, small)
        space_info = space_context(product["space"], fallback_spaces)
        placements = placement_values(large, mid, small)
        package_ids = recursive_product_ids(product["packages"].get("data"), product_id)
        combination = product_combinations[product_id]
        combination_detection_status = combination["detection_status"]
        combination_output_status = combination["component_output_status"]
        is_set = combination_detection_status != "NOT_COMBINATION"
        set_value = (
            "|".join(package_ids)
            if package_ids
            else MISSING
            if is_set
            else NOT_APPLICABLE
        )
        combination_label = COMBINATION_STATUS_LABELS.get(
            combination_detection_status,
            "검토필요",
        )
        component_count_value = (
            MISSING
            if combination_detection_status == "CANDIDATE"
            else combination["component_count"]
            if combination_detection_status == "CONFIRMED"
            else 0
        )
        component_confirmed_count_value = (
            combination["complete_component_count"]
            if combination_detection_status == "CONFIRMED"
            else MISSING
            if combination_detection_status == "CANDIDATE"
            else 0
        )
        component_output_label = COMPONENT_OUTPUT_STATUS_LABELS.get(
            combination_output_status,
            combination_output_status or MISSING,
        )
        output_status = OUTPUT_STATUS_LABELS.get(
            str(resolution_status),
            "완료_무후보분류",
        )
        pattern_status = product_pattern_status(
            str(resolution_status),
            combination_detection_status,
        )
        dimension_pattern_code = dimension_review_pattern_code(
            str(resolution_status),
            product["dimension_comparisons"],
            combination_detection_status,
        )
        dimension_pattern_name = DIMENSION_PATTERN_NAMES.get(
            dimension_pattern_code,
            NOT_APPLICABLE,
        )
        combination_review = combination_review_patterns.get(product_id) or {}
        legacy_combination_pattern = str(
            combination_review.get("pattern_code") or ""
        )
        combination_pattern_codes = (
            combination_review_codes(
                legacy_combination_pattern,
                name,
            )
            if combination_detection_status == "CANDIDATE"
            else []
        )
        if combination_detection_status == "CANDIDATE":
            combination_pattern_codes.extend(
                option_pattern_codes(product["groups"])
            )
        combination_pattern_codes = sorted(
            set(combination_pattern_codes),
            key=lambda code: (
                PATTERN_GROUP_ORDER.index(code[0])
                if code and code[0] in PATTERN_GROUP_ORDER
                else len(PATTERN_GROUP_ORDER),
                code,
            ),
        )
        combination_pattern_groups = [
            group
            for group in PATTERN_GROUP_ORDER
            if any(code.startswith(group) for code in combination_pattern_codes)
        ]
        if pattern_status == "체크필요_규격+조합패턴":
            pattern_next_action = "조합 패턴 확인 → 규격 패턴 확인"
        elif pattern_status == "체크필요_규격패턴":
            pattern_next_action = "규격 패턴 확인"
        elif pattern_status == "체크필요_조합패턴":
            pattern_next_action = "조합 패턴 확인"
        else:
            pattern_next_action = "추가 확인 없음"
        product_page_url = (
            "https://homestyle.lge.co.kr/item?productId=" + product_id
        )
        pattern_cells = [
            output_status,
            pattern_status,
            dimension_pattern_code or NOT_APPLICABLE,
            dimension_pattern_name,
            "|".join(combination_pattern_groups) or NOT_APPLICABLE,
            "|".join(combination_pattern_codes) or NOT_APPLICABLE,
            legacy_combination_pattern or NOT_APPLICABLE,
            pattern_next_action,
            product_page_url,
        ]
        output_status_counts[output_status] += 1
        pattern_status_counts[pattern_status] += 1
        if dimension_pattern_code:
            dimension_pattern_counts[dimension_pattern_code] += 1
        for pattern_group in combination_pattern_groups:
            combination_pattern_group_counts[pattern_group] += 1
        wall_height = MISSING if "벽" in placements else NOT_APPLICABLE

        mandatory = mandatory_asset_assessment(
            system="홈스타일",
            product_id=product_id,
            brand=brand,
            mid_category=mid or "",
            small_category=small or "",
            representative_image_url=images[0] if images else "",
            rolling_image_urls=images,
            color=product["colors"],
            w_mm=primary_dimension["w_mm"],
            d_mm=primary_dimension["d_mm"],
            h_mm=primary_dimension["h_mm"],
            is_set=is_set,
            set_component_ids=package_ids,
        )
        if (
            mandatory["status"] == MANDATORY_PASS
            and resolution_status in LOCKED_DIMENSION_STATUSES
            and int(combination.get("needs_human_review") or 0) == 0
        ):
            request1_complete_count += 1
        else:
            request1_review_count += 1

        groups = product["groups"]
        flattened_options = []
        for group in groups:
            for item in group["items"]:
                value = item["name"]
                if len(groups) > 1:
                    value = f"{group['style']}: {value}"
                flattened_options.append(value)
        option_cells = flattened_options + [NOT_APPLICABLE] * (max_options - len(flattened_options))
        option_styles = " | ".join(group["style"] for group in groups)

        request1 = [
            ATTACHMENT_0715_STATUS_LABELS.get(
                str(resolution_status), "최종무후보"
            ),
            mandatory["fulfilled_count"],
            mandatory["required_count"],
            "|".join(mandatory["missing"]) or NOT_APPLICABLE,
            images[0] if images else MISSING,
            len(images),
            *rolling_cells,
            primary_dimension["target"],
            primary_dimension["w_mm"] if primary_dimension["w_mm"] is not None else "",
            primary_dimension["d_mm"] if primary_dimension["d_mm"] is not None else "",
            primary_dimension["h_mm"] if primary_dimension["h_mm"] is not None else "",
            *dimension_comparison_cells,
            primary_dimension["status"],
            mid or MISSING,
            small or MISSING,
            space_info["spaces"] or MISSING,
            brand,
            product["colors"],
            NOT_APPLICABLE,
            set_value,
            combination_label,
            combination["detection_rule"]
            if combination_detection_status != "NOT_COMBINATION"
            else NOT_APPLICABLE,
            component_count_value,
            component_confirmed_count_value,
            component_output_label,
            "|".join(placements),
            wall_height,
            option_styles or MISSING,
            len(flattened_options),
            *option_cells,
        ]

        detail_text = clean_text(data.get("detailInfo"), 10000)
        service_text = clean_text(data.get("serviceGuide"), 5000)
        manufacturer = notification_values(notifications, "제조자", "제조사", "수입자")
        origin = notification_values(notifications, "제조국", "원산지")
        materials = notification_values(notifications, "소재", "재질")
        components = notification_values(notifications, "구성품", "제품구성")
        assembly = notification_values(notifications, "조립", "설치")
        certifications = notification_values(notifications, "인증", "허가", "KC")
        warranty = notification_values(notifications, "품질보증", "보증")
        safety = notification_values(notifications, "안전", "주의")
        care = notification_values(notifications, "세탁", "취급", "관리", "주의")
        if not materials:
            signal = qna_signal(product["qna"], product["html"], ("소재", "재질"))
            if signal:
                materials.append("FAQ/Q&A: " + signal)
        dimension_text = " | ".join(
            f"W={record['w_mm']}, D={record['d_mm']}, H={record['h_mm']} mm"
            for record in product["dimensions"]
            if any(record[key] is not None for key in ("w_mm", "d_mm", "h_mm"))
        ) or MISSING
        purpose = "|".join(fallback_spaces) + " 공간에서 " + (small or mid or "제품") + " 용도"
        sale_prices = [
            row.get("discountPrice") or row.get("salePrice")
            for row in (data.get("productStock") or [])
            if isinstance(row, dict) and (row.get("discountPrice") or row.get("salePrice"))
        ]
        sales = ["판매상태=" + str(data.get("status") or MISSING)]
        if sale_prices:
            sales.append(f"판매가={min(sale_prices):,}~{max(sale_prices):,}원")
        sales.append("수집시점 기준")
        styles = style_tags(name, detail_text, space_info["styles"])
        synonyms = []
        for color in re.split(r"[,/|]", product["colors"]):
            for key, value in COLOR_SYNONYMS.items():
                if key in color and value not in synonyms:
                    synonyms.append(value)
        semantic = f"{brand} {name} {large} {mid} {small} 색상={product['colors']}"
        if synonyms:
            semantic += " | 동의어=" + "|".join(synonyms)

        request2 = [
            f"상품명={name} | 브랜드={brand} | 분류={large}>{mid}>{small}",
            joined(manufacturer + origin),
            joined(materials),
            dimension_text,
            joined(components),
            product["colors"],
            purpose,
            joined(assembly, NOT_APPLICABLE if any(word in small for word in ("러그", "소품", "액자")) else MISSING),
            joined(safety),
            joined(care, clean_text(service_text, 1200) or MISSING),
            joined(warranty),
            joined(certifications),
            " | ".join(sales),
            styles,
            space_info["styles"] or styles + " (제품 기반 추론)",
            space_info["moods"],
            space_info["spaces"],
            purpose,
            product["colors"] + " (제품 색상 기반)",
            MISSING,
            space_info["content_products"],
            space_info["placement_reasons"] if space_info["placement_reasons"] != MISSING else "위치=" + "|".join(placements),
            f"{styles} / {space_info['spaces']} ↔ {product_id} (API+추론)",
            space_info["relationships"],
            MISSING,
            semantic,
        ]
        main_rows.append(
            ["홈스타일", product_id, name]
            + pattern_cells
            + request1
            + request2
        )
        pattern_rows.append(
            ["홈스타일", product_id, name] + pattern_cells
        )

        mandatory_stats[mandatory["status"]] += 1
        mandatory_missing_fields.update(mandatory["missing"])
        mandatory_rows.append(
            [
                "홈스타일", product_id, name, mandatory["status"],
                mandatory["fulfilled_count"], mandatory["required_count"],
                "|".join(mandatory["missing"]) or NOT_APPLICABLE,
                mandatory["image_quality_status"],
                *[
                    NOT_APPLICABLE
                    if mandatory["checks"][check_name] is None
                    else "충족"
                    if mandatory["checks"][check_name]
                    else MISSING
                    for check_name in mandatory_check_names
                ],
            ]
        )

        # The first four request-1 cells are assessment metadata, not customer
        # source fields, so exclude them from the legacy cell availability rate.
        for value in request1[4:]:
            availability["R1"]["missing" if value in (MISSING, "") else "na" if value == NOT_APPLICABLE else "available"] += 1
        for value in request2:
            availability["R2"]["missing" if value in (MISSING, "") else "na" if value == NOT_APPLICABLE else "available"] += 1

        for dimension_index, dimension in enumerate(product["dimensions"], start=1):
            dimension_rows.append(
                [
                    "홈스타일", product_id, name, dimension_index, dimension["target"],
                    dimension["w_mm"] if dimension["w_mm"] is not None else "",
                    dimension["d_mm"] if dimension["d_mm"] is not None else "",
                    dimension["h_mm"] if dimension["h_mm"] is not None else "",
                    dimension["status"], dimension["evidence"], dimension["raw"],
                ]
            )
        for style_index, group in enumerate(groups, start=1):
            for option_index, item in enumerate(group["items"], start=1):
                option_rows.append(
                    [
                        "홈스타일", product_id, name, style_index, group["style"],
                        group["raw_style"] or NOT_APPLICABLE, option_index, item["name"],
                        item["option_id"] or MISSING, item["image_url"] or NOT_APPLICABLE,
                        group["evidence"],
                    ]
                )

    for component in product_components:
        combination = product_combinations[component["product_id"]]
        component_rows.append(
            [
                "홈스타일",
                component["product_id"],
                component["product_name"],
                COMBINATION_STATUS_LABELS.get(
                    combination["detection_status"],
                    "검토필요",
                ),
                component["component_seq"],
                component["component_name"],
                component["component_type"],
                component["component_quantity"],
                "Y" if int(component["is_primary"] or 0) == 1 else "N",
                component["w_mm"] if component["w_mm"] is not None else "",
                component["d_mm"] if component["d_mm"] is not None else "",
                component["h_mm"] if component["h_mm"] is not None else "",
                COMPONENT_RESOLUTION_STATUS_LABELS.get(
                    component["resolution_status"],
                    component["resolution_status"],
                ),
                component["source_type"] or MISSING,
                component["source_ref"] or MISSING,
                component["evidence_text"] or MISSING,
                "Y" if int(component["needs_human_review"] or 0) == 1 else "N",
            ]
        )

    connection = sqlite3.connect(DB_PATH)
    category_stats_rows = [[
        "scope_id", "대분류", "중분류", "소분류", "category_id",
        "원본 상품수", "현재 노출행수", "고유 상품수", "API 상태", "비고",
    ]]
    for category in connection.execute(
        "SELECT scope_id, large_name, mid_name, small_name, category_id, source_count, live_count, http_status, error FROM categories ORDER BY scope_id"
    ):
        unique_count = connection.execute(
            "SELECT COUNT(*) FROM products WHERE category_scope_ids LIKE ?",
            (f'%"{category[0]}"%',),
        ).fetchone()[0]
        category_stats_rows.append(
            [*category[:7], unique_count, category[7], category[8] or ""]
        )
    connection.close()

    error_rows = [[
        "상품 ID", "goods API", "공간 API", "세트 API", "FAQ/Q&A", "PDP HTML", "이미지 OCR", "비고",
    ]]
    for error in errors:
        error_rows.append(
            [
                error["product_id"], error["goods_status"], error["space_status"],
                error["packages_status"], error["qna_status"], error["html_status"],
                error["ocr_status"],
                error["note"],
            ]
        )
    for product in products:
        ocr_status = product["statuses"].get("ocr") or 0
        if ocr_status not in (204, 502):
            continue
        selected = product["ocr"].get("selected") or {}
        note = (
            "상세 이미지 없음"
            if ocr_status == 204
            else clean_text(product["ocr"].get("error") or "외부 이미지 다운로드 실패")
        )
        if selected.get("url"):
            note += " | " + selected["url"]
        error_rows.append(
            [
                product["product_id"],
                product["statuses"].get("goods") or 0,
                product["statuses"].get("space") or 0,
                product["statuses"].get("packages") or 0,
                product["statuses"].get("qna") or 0,
                product["statuses"].get("html") or 0,
                ocr_status,
                note,
            ]
        )

    source_counts = Counter()
    for product in products:
        for name, status in product["statuses"].items():
            source_counts[f"{name}_ok" if status == 200 else f"{name}_fail"] += 1

    summary_rows = [["구분", "수치", "설명"]]
    mandatory_total = len(products) + len(errors)
    mandatory_pass_count = mandatory_stats[MANDATORY_PASS]
    mandatory_reinforcement_count = (
        mandatory_stats[MANDATORY_REINFORCEMENT] + len(errors)
    )
    request1_review_count += len(errors)
    mandatory_pass_rate = (
        mandatory_pass_count / mandatory_total if mandatory_total else 0
    )
    dimension_analyzable = (
        dimension_progress.get("locked_resolved", 0)
        + dimension_progress.get("comparison_provided", 0)
    )
    dimension_total = dimension_progress.get("total_products", 0)
    dimension_analyzable_rate = (
        dimension_analyzable / dimension_total if dimension_total else 0
    )
    summary_rows.extend(
        [
            ["대상 시스템", "홈스타일", "가전 제외"],
            ["대상 카테고리", len(category_stats_rows) - 1, "제품군.xlsx 비음영 홈스타일 소분류"],
            ["카테고리 노출행", sum(row[6] or 0 for row in category_stats_rows[1:]), "카테고리 중복 포함"],
            ["고유 상품 ID", len(products) + len(errors), "상품 ID 기준 중복 제거"],
            ["결과 생성 상품", len(products), "goods API 정상 응답 상품"],
            ["결과 제외/오류", len(errors), "06_수집오류 시트 참조"],
            [
                "산출 상태: 완료_확정값",
                output_status_counts["완료_확정값"],
                "사용할 규격값 산출 완료",
            ],
            [
                "산출 상태: 완료_RAW분리",
                output_status_counts["완료_RAW분리"],
                "규격 비교 후보를 RAW 컬럼으로 분리·저장 완료",
            ],
            [
                "산출 상태: 완료_무후보분류",
                output_status_counts["완료_무후보분류"],
                "유효 후보 없음 상태까지 분류 완료",
            ],
            [
                "패턴 상태: 완료_패턴체크불필요",
                pattern_status_counts["완료_패턴체크불필요"],
                "현재 추가 패턴 확인 불필요",
            ],
            [
                "패턴 상태: 체크필요_규격패턴",
                pattern_status_counts["체크필요_규격패턴"],
                "규격 RAW 확정 패턴 확인 필요",
            ],
            [
                "패턴 상태: 체크필요_조합패턴",
                pattern_status_counts["체크필요_조합패턴"],
                "조합·옵션 구조 패턴 확인 필요",
            ],
            [
                "패턴 상태: 체크필요_규격+조합패턴",
                pattern_status_counts["체크필요_규격+조합패턴"],
                "조합 구조 확인 후 규격 패턴 확인",
            ],
            [
                "패턴 체크필요 합계",
                sum(
                    count
                    for status, count in pattern_status_counts.items()
                    if status.startswith("체크필요_")
                ),
                "규격·조합 패턴 체크 대상의 고유 상품 수",
            ],
            ["별첨0715 필수값 PASS", mandatory_pass_count, "적용되는 모든 필수값이 존재하는 상품"],
            ["별첨0715 필수값 보강대상", mandatory_reinforcement_count, "필수값 하나 이상 미확보 또는 상품 수집오류"],
            ["별첨0715 필수값 PASS율", f"{mandatory_pass_rate:.1%}", f"PASS {mandatory_pass_count:,} / 전체 {mandatory_total:,}"],
            ["별첨0715 공통 필수", 6, "브랜드, 중·소카테고리, ID, 대표+롤링 이미지 URL, 색상, W/D/H"],
            ["별첨0715 조건부 필수", "가전", "가전=후면 이미지+가전 추가사항"],
            [
                "기존 요청1 완전완료 판정(참고)",
                request1_complete_count,
                "별첨 필수값 PASS + 대표 규격이 DB/규칙/수동으로 잠긴 상품",
            ],
            [
                "기존 요청1 사람판단·보강 판정(참고)",
                request1_review_count,
                "규격 비교·무후보 또는 필수값 보강; 노란색/분홍색 셀 확인",
            ],
            ["세트 구성 실제 ID", "참고 필드", "2026-07-22 정책 변경: PASS 필수 판정에서 제외"],
            [
                "조합상품 확정",
                combination_detection_counts.get("CONFIRMED", 0),
                "상품명 세트·패키지 또는 소파+스툴/오토만 고신뢰 규칙",
            ],
            [
                "조합상품 검토필요",
                combination_detection_counts.get("CANDIDATE", 0),
                "+ 표기는 있으나 모델명·색상·옵션 결합 과탐 가능",
            ],
            [
                "조합상품 구성품 행",
                len(component_rows) - 1,
                "07_조합상품_구성품 시트의 product_id+component_seq 행 수",
            ],
            [
                "구성품 규격 전체 확정 상품",
                component_output_counts.get(
                    "ALL_COMPONENT_DIMENSIONS_CONFIRMED",
                    0,
                ),
                "모든 구성품의 라벨·W/D/H·단위가 상품 API에 명시",
            ],
            [
                "구성품 규격 후보·부분확보",
                component_output_counts.get("COMPONENT_DIMENSIONS_CANDIDATE", 0)
                + component_output_counts.get("PARTIAL_COMPONENT_DIMENSIONS", 0),
                "단위 추론 또는 일부 구성품만 규격 확보; 노란색 검수",
            ],
            ["필수 이미지 대체 기준", "대표+롤링 URL", "대표 이미지와 롤링 이미지 URL이 모두 입력되면 필수 충족"],
            ["롤링 이미지 열", max_images, "상품별 모든 이미지 URL을 개별 열로 분리"],
            ["옵션 열", max_options, "옵션 스타일/옵션 01~N 분리"],
            ["규격 단위", "mm", "W/D/H는 숫자 셀; 단위는 컬럼명에 표시"],
            [
                "규격 확정·잠금 상품",
                dimension_progress.get("locked_resolved", 0),
                "기존 DB 확정 또는 누적 규칙으로 대표 규격 확정",
            ],
            [
                "별첨0715 판정: API/HTML 확정",
                dimension_progress.get("source_confirmed", 0),
                "상품 API·PDP HTML 등 기존 원천에서 대표 규격 확정",
            ],
            [
                "별첨0715 판정: OCR·규칙확정",
                dimension_progress.get("rule_resolved", 0)
                + dimension_progress.get("manual_confirmed", 0),
                "OCR·문맥 규칙 또는 수동 확정으로 대표 규격 잠금",
            ],
            [
                "별첨0715 판정: 비교정보제공",
                dimension_progress.get("comparison_provided", 0),
                "자동 확정하지 않고 사람이 선택할 비교 후보 제공",
            ],
            [
                "별첨0715 판정: 최종무후보",
                dimension_progress.get("no_candidate", 0)
                + dimension_progress.get("unclassified", 0),
                "전체 원천·OCR·규칙 적용 후 대표 규격 후보 없음",
            ],
            [
                "규격 비교정보 제공 상품",
                dimension_progress.get("comparison_provided", 0),
                "자동 확정하지 않고 사람이 대표값을 선택하도록 모든 후보 제공; 노란색 검수",
            ],
            [
                "규격 사람 확인 필요",
                dimension_progress.get("human_review_required", 0),
                "비교정보 제공 정책 적용 후 잔여",
            ],
            [
                "규격 OCR 보강 잔여",
                dimension_progress.get("ocr_pipeline_remaining", 0),
                "전체/영역 OCR 및 원시 숫자 비교정보 제공 후에도 후보가 없는 상품",
            ],
            [
                "요청 1 규격 분석 가능",
                dimension_analyzable,
                "확정·잠금 + 비교정보 제공",
            ],
            [
                "요청 1 규격 분석 가능률",
                f"{dimension_analyzable_rate:.1%}",
                "상품 ID 기준 전체 대비 확정 또는 비교 가능한 비율",
            ],
            ["OCR 추가 완전 확보", ocr_impact["complete"], "API·HTML·FAQ/Q&A 대비 OCR로 W·D·H 완성"],
            ["OCR 추가 부분 확보", ocr_impact["partial"], "OCR로 새 규격 레코드 추가"],
            ["요청 1 색상", "파란색", "요청 1 필드"],
            ["요청 2 색상", "초록색", "요청 2 필드"],
            [
                "사람 판단 필요 색상",
                "노란색",
                "비교정보·부분확보·규칙/제품 기반 추론값·필수 보강·수집오류",
            ],
            ["빈 필드 색상", "분홍색", MISSING],
            ["비적용 표시", "회색", NOT_APPLICABLE],
        ]
    )
    for reason in dimension_remaining_reasons:
        summary_rows.append(
            [
                f"요청 1 규격 잔여: {reason['reason_name']}",
                reason["product_count"],
                reason["next_action"],
            ]
        )
    for source_name, label in (
        ("goods", "상품 API"),
        ("space", "공간 추천 API"),
        ("packages", "세트 구성 API"),
        ("qna", "FAQ/Q&A API"),
        ("html", "PDP HTML"),
        ("ocr", "상세 이미지 OCR(선별)"),
    ):
        summary_rows.append(
            [
                label,
                source_counts[f"{source_name}_ok"],
                f"정상 {source_counts[f'{source_name}_ok']} / 실패·미실행 {source_counts[f'{source_name}_fail']}",
            ]
        )
    for field_name, count in mandatory_missing_fields.most_common():
        summary_rows.append(
            [
                f"필수값 보강: {field_name}",
                count,
                "04_필수값_판정 시트에서 대상 상품 확인",
            ]
        )
    if errors:
        summary_rows.append(
            ["필수값 보강: 상품 수집오류", len(errors), "필수값 판정 불가이므로 보강대상에 포함"]
        )
    for group, label in (("R1", "요청 1 출력 셀"), ("R2", "요청 2 출력 셀")):
        stats = availability[group]
        denominator = stats["available"] + stats["missing"]
        rate = stats["available"] / denominator if denominator else 0
        summary_rows.append(
            [
                label,
                f"{rate:.1%}",
                f"확보 {stats['available']:,} / 미확보 {stats['missing']:,} / 해당없음 {stats['na']:,}",
            ]
        )

    summary_groups = ["COMMON"] * (len(summary_rows) - 1)
    for index, row in enumerate(summary_rows[1:]):
        if row[0].startswith(
            (
                "요청 1",
                "별첨0715",
                "필수값",
                "조합상품",
                "구성품",
                "산출 상태",
                "패턴 상태",
                "패턴 체크",
                "기존 요청1",
            )
        ):
            summary_groups[index] = "R1"
        elif row[0].startswith("요청 2"):
            summary_groups[index] = "R2"

    sheets = [
        (
            "00_요약",
            summary_rows,
            [24, 24, 105],
            ["COMMON"] * 3,
            summary_groups,
        ),
        (
            "01_상품별_요구필드",
            main_rows,
            [14, 20, 42]
            + [20, 30, 16, 34, 18, 38, 34, 34, 54]
            + [
                70
                if header == "요청1_규격 비교 원문"
                else 38
                if header == "요청1_규격 비교 대상/옵션"
                else 54
                if "URL" in header
                else 16
                if header.endswith("(mm)") or header.endswith(" 수")
                else 28
                for header in request1_headers + request2_headers
            ],
            ["COMMON"] * (3 + len(pattern_headers))
            + ["R1"] * len(request1_headers)
            + ["R2"] * len(request2_headers),
        ),
        (
            "02_규격_상세",
            dimension_rows,
            [14, 20, 42, 12, 28, 14, 14, 14, 16, 38, 70],
            ["COMMON"] * 3 + ["R1"] * 8,
        ),
        (
            "03_옵션_상세",
            option_rows,
            [14, 20, 42, 16, 22, 22, 12, 32, 16, 54, 42],
            ["COMMON"] * 3 + ["R1"] * 8,
        ),
        (
            "04_필수값_판정",
            mandatory_rows,
            [14, 20, 42, 16, 12, 12, 48, 18] + [19] * len(mandatory_check_names),
            ["COMMON"] * 3 + ["R1"] * (len(mandatory_rows[0]) - 3),
        ),
        (
            "05_카테고리_통계",
            category_stats_rows,
            [14, 18, 22, 24, 18, 16, 16, 16, 14, 40],
            ["COMMON"] * 10,
        ),
        (
            "06_수집오류",
            error_rows,
            [20, 14, 14, 14, 14, 14, 14, 60],
            ["COMMON"] * 8,
        ),
        (
            "07_조합상품_구성품",
            component_rows,
            [14, 20, 44, 16, 12, 22, 20, 13, 13, 14, 14, 14, 20, 24, 38, 70, 16],
            ["COMMON"] * 3 + ["R1"] * 14,
        ),
        (
            "08_패턴_상태",
            pattern_rows,
            [14, 20, 44, 20, 30, 16, 34, 18, 38, 34, 34, 54],
            ["COMMON"] * len(pattern_rows[0]),
        ),
    ]
    meta = {
        "products": len(products),
        "errors": len(errors),
        "fields": len(main_rows[0]) - 3,
        "max_images": max_images,
        "max_options": max_options,
        "dimension_rows": len(dimension_rows) - 1,
        "option_rows": len(option_rows) - 1,
        "combination_confirmed": combination_detection_counts.get(
            "CONFIRMED",
            0,
        ),
        "combination_candidate": combination_detection_counts.get(
            "CANDIDATE",
            0,
        ),
        "component_rows": len(component_rows) - 1,
        "pattern_rows": len(pattern_rows) - 1,
        "output_status_counts": dict(output_status_counts),
        "pattern_status_counts": dict(pattern_status_counts),
        "dimension_pattern_counts": dict(dimension_pattern_counts),
        "combination_pattern_group_counts": dict(
            combination_pattern_group_counts
        ),
        "component_output_status_counts": component_output_counts,
        "dimension_complete": dimension_stats["complete"],
        "dimension_partial": dimension_stats["partial"],
        "dimension_missing": dimension_stats["missing"],
        "ocr_added_complete": ocr_impact["complete"],
        "ocr_added_partial": ocr_impact["partial"],
        "mandatory_pass": mandatory_pass_count,
        "mandatory_reinforcement": mandatory_reinforcement_count,
        "mandatory_pass_rate": mandatory_pass_rate,
        "mandatory_missing_fields": dict(mandatory_missing_fields.most_common()),
        "request1_complete": request1_complete_count,
        "request1_review_or_reinforcement": request1_review_count,
        "issue_rows": len(error_rows) - 1,
        "dimension_locked_resolved": dimension_progress.get("locked_resolved", 0),
        "dimension_comparison_provided": dimension_progress.get(
            "comparison_provided", 0
        ),
        "dimension_human_review_required": dimension_progress.get(
            "human_review_required", 0
        ),
        "dimension_ocr_remaining": dimension_progress.get(
            "ocr_pipeline_remaining", 0
        ),
        "dimension_remaining_reasons": {
            row["reason_code"]: row["product_count"]
            for row in dimension_remaining_reasons
        },
    }
    return sheets, meta


def _xlsx_cell_text(cell: ET.Element) -> str:
    if cell.attrib.get("t") == "inlineStr":
        return "".join(node.text or "" for node in cell.iter(XLSX_NS + "t"))
    value = cell.find(XLSX_NS + "v")
    return value.text or "" if value is not None else ""


def _xlsx_column(reference: str) -> str:
    match = re.match(r"[A-Z]+", reference)
    return match.group(0) if match else ""


def patch_human_review_cell_styles(path: Path) -> dict[str, Any]:
    """Color only cells whose value still requires a person to judge.

    Priority is deliberate:
    - missing/blank dimension cells keep the pink style added by
      patch_special_cell_styles();
    - not-applicable cells keep the gray style;
    - comparison, partial, inferred, reinforcement, and collection-error
      cells receive the yellow review style.
    """
    with ZipFile(path) as source:
        files = {name: source.read(name) for name in source.namelist()}

    styles = files["xl/styles.xml"].decode("utf-8")
    if '<fills count="11">' not in styles or '<cellXfs count="11">' not in styles:
        raise RuntimeError("expected missing/N/A styles were not found before review styling")
    styles = styles.replace('<fills count="11">', '<fills count="12">', 1)
    styles = styles.replace(
        "</fills>",
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFE699"/>'
        '<bgColor indexed="64"/></patternFill></fill></fills>',
        1,
    )
    styles = styles.replace('<cellXfs count="11">', '<cellXfs count="12">', 1)
    styles = styles.replace(
        "</cellXfs>",
        '<xf numFmtId="0" fontId="0" fillId="11" borderId="1" xfId="0" '
        'applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>'
        "</cellXfs>",
        1,
    )
    files["xl/styles.xml"] = styles.encode("utf-8")

    workbook = ET.fromstring(files["xl/workbook.xml"])
    sheet_names = [
        sheet.attrib["name"] for sheet in workbook.find(XLSX_NS + "sheets")
    ]
    review_cells_by_sheet: Counter[str] = Counter()
    request1_review_products: set[str] = set()
    inferred_markers = (
        "추론",
        "제품 기반",
        "카테고리 기반",
        "API+추론",
        "규칙 기반",
    )
    comparison_headers = {
        "요청1_대표 규격 대상",
        "요청1_W (mm)",
        "요청1_D (mm)",
        "요청1_H (mm)",
        "요청1_규격 비교후보 수",
        "요청1_규격 비교 W (mm)",
        "요청1_규격 비교 D (mm)",
        "요청1_규격 비교 H (mm)",
        "요청1_규격 비교 대상/옵션",
        "요청1_규격 비교 원문",
        "요청1_규격 상태",
    }
    combination_headers = {
        "요청1_조합상품 여부",
        "요청1_조합상품 판정근거",
        "요청1_조합 구성품 수",
        "요청1_구성품 규격 확정 수",
        "요청1_구성품 분리 상태",
    }
    pattern_review_headers = {
        "패턴 상태",
        "규격 패턴 코드",
        "규격 패턴명",
        "조합 패턴군",
        "조합 세부 패턴 코드",
        "조합 원천 패턴 코드",
        "패턴 확인 순서",
        "상품 페이지 URL",
    }

    def apply_review(cell: ET.Element | None, sheet_name: str) -> None:
        if cell is None or cell.attrib.get("s") in {"9", "10", HUMAN_REVIEW_STYLE_ID}:
            return
        if not _xlsx_cell_text(cell).strip():
            return
        cell.set("s", HUMAN_REVIEW_STYLE_ID)
        review_cells_by_sheet[sheet_name] += 1

    for sheet_index, sheet_name in enumerate(sheet_names, start=1):
        xml_name = f"xl/worksheets/sheet{sheet_index}.xml"
        root = ET.fromstring(files[xml_name])
        rows = root.findall(".//" + XLSX_NS + "sheetData/" + XLSX_NS + "row")
        if not rows:
            continue
        header_by_column = {
            _xlsx_column(cell.attrib.get("r", "")): _xlsx_cell_text(cell)
            for cell in rows[0].findall(XLSX_NS + "c")
        }
        column_by_header = {
            header: column for column, header in header_by_column.items()
        }

        for row in rows[1:]:
            cell_by_column = {
                _xlsx_column(cell.attrib.get("r", "")): cell
                for cell in row.findall(XLSX_NS + "c")
            }

            def cell_for(header: str) -> ET.Element | None:
                column = column_by_header.get(header)
                return cell_by_column.get(column) if column else None

            def value_for(header: str) -> str:
                cell = cell_for(header)
                return _xlsx_cell_text(cell).strip() if cell is not None else ""

            if sheet_name == "00_요약":
                label = value_for("구분")
                if (
                    label.startswith("패턴 상태: 체크필요")
                    or label in {
                        "패턴 체크필요 합계",
                        "기존 요청1 사람판단·보강 판정(참고)",
                        "사람 판단 필요 색상",
                    }
                ):
                    for cell in cell_by_column.values():
                        apply_review(cell, sheet_name)

            elif sheet_name == "01_상품별_요구필드":
                product_id = value_for("상품 ID")
                pattern_status = value_for("패턴 상태")
                if pattern_status.startswith("체크필요_"):
                    for header in pattern_review_headers:
                        apply_review(cell_for(header), sheet_name)
                attachment_status = value_for("별첨0715_필수값 판정")
                reinforcement_fields = value_for("별첨0715_보강대상 필드")
                dimension_status = value_for("요청1_규격 상태")
                request1_needs_review = False
                if reinforcement_fields not in {"", NOT_APPLICABLE}:
                    request1_needs_review = True
                    apply_review(
                        cell_for("별첨0715_보강대상 필드"), sheet_name
                    )
                if attachment_status in {"비교정보제공", "최종무후보"}:
                    request1_needs_review = True
                    apply_review(
                        cell_for("별첨0715_필수값 판정"), sheet_name
                    )
                if dimension_status in HUMAN_REVIEW_DIMENSION_LABELS:
                    request1_needs_review = True
                    for header in comparison_headers:
                        apply_review(cell_for(header), sheet_name)
                combination_value = value_for("요청1_조합상품 여부")
                component_split_status = value_for("요청1_구성품 분리 상태")
                if (
                    combination_value == "검토필요"
                    or (
                        combination_value == "Y"
                        and component_split_status != "구성품 규격 확정"
                    )
                ):
                    request1_needs_review = True
                    for header in combination_headers:
                        apply_review(cell_for(header), sheet_name)
                if request1_needs_review and product_id:
                    request1_review_products.add(product_id)

                for header, column in column_by_header.items():
                    if not header.startswith("요청2_"):
                        continue
                    cell = cell_by_column.get(column)
                    text = _xlsx_cell_text(cell).strip() if cell is not None else ""
                    if (
                        header == "요청2_디자인 스타일 추론"
                        or any(marker in text for marker in inferred_markers)
                    ):
                        apply_review(cell, sheet_name)

            elif sheet_name == "02_규격_상세":
                if value_for("규격 상태") == "부분확보":
                    for header in (
                        "규격 대상/옵션",
                        "W (mm)",
                        "D (mm)",
                        "H (mm)",
                        "규격 상태",
                        "원천/비고",
                        "원문",
                    ):
                        apply_review(cell_for(header), sheet_name)

            elif sheet_name == "04_필수값_판정":
                if value_for("필수값 판정") == MANDATORY_REINFORCEMENT:
                    for header in (
                        "필수값 판정",
                        "보강대상 필드",
                        "이미지 필수 대체판정",
                    ):
                        apply_review(cell_for(header), sheet_name)

            elif sheet_name == "06_수집오류":
                for cell in cell_by_column.values():
                    apply_review(cell, sheet_name)

            elif sheet_name == "07_조합상품_구성품":
                if value_for("사람 확인 필요") == "Y":
                    for header in (
                        "조합상품 판정",
                        "구성품 순번",
                        "구성품명",
                        "구성품 유형",
                        "구성품 수량",
                        "대표 구성품",
                        "W (mm)",
                        "D (mm)",
                        "H (mm)",
                        "구성품 규격 상태",
                        "원천 유형",
                        "원천 위치",
                        "근거 원문",
                        "사람 확인 필요",
                    ):
                        apply_review(cell_for(header), sheet_name)

            elif sheet_name == "08_패턴_상태":
                if value_for("패턴 상태").startswith("체크필요_"):
                    for header in pattern_review_headers:
                        apply_review(cell_for(header), sheet_name)

        files[xml_name] = ET.tostring(
            root, encoding="utf-8", xml_declaration=True
        )

    temp_path = path.with_suffix(".review.tmp.xlsx")
    with ZipFile(temp_path, "w", ZIP_DEFLATED) as target:
        for name, data in files.items():
            target.writestr(name, data)
    temp_path.replace(path)
    return {
        "style_id": int(HUMAN_REVIEW_STYLE_ID),
        "fill_rgb": "FFFFE699",
        "review_cells": sum(review_cells_by_sheet.values()),
        "review_cells_by_sheet": dict(review_cells_by_sheet),
        "request1_review_products": len(request1_review_products),
    }


def validate_mandatory_staging(meta: dict[str, Any]) -> str:
    """Block Excel output when the current DB staging snapshot is stale."""
    connection = sqlite3.connect(DB_PATH)
    exists = connection.execute(
        "SELECT COUNT(*) FROM sqlite_master "
        "WHERE type='table' AND name='stg_mandatory_pass'"
    ).fetchone()[0]
    if not exists:
        connection.close()
        raise RuntimeError(
            "필수값 스테이징이 없습니다. "
            "먼저 python build_mandatory_pass_staging.py 를 실행하세요."
        )
    snapshots = connection.execute(
        "SELECT snapshot_id,COUNT(*) FROM stg_mandatory_pass "
        "WHERE is_current=1 GROUP BY snapshot_id"
    ).fetchall()
    if len(snapshots) != 1:
        connection.close()
        raise RuntimeError(f"현재 필수값 스냅샷이 1개가 아닙니다: {snapshots}")
    snapshot_id, row_count = snapshots[0]
    status_counts = {
        status: count
        for status, count in connection.execute(
            "SELECT final_status,COUNT(*) FROM stg_mandatory_pass "
            "WHERE is_current=1 GROUP BY final_status"
        )
    }
    missing_counts = {
        field: count
        for field, count in connection.execute(
            "SELECT required_field,product_count "
            "FROM vw_mandatory_pass_current_missing WHERE product_count>0"
        )
    }
    connection.close()
    expected_rows = int(meta["products"]) + int(meta["errors"])
    expected_status = {
        MANDATORY_PASS: int(meta["mandatory_pass"]),
        MANDATORY_REINFORCEMENT: int(meta["mandatory_reinforcement"]),
    }
    expected_missing = dict(meta["mandatory_missing_fields"])
    if row_count != expected_rows:
        raise RuntimeError(
            f"필수값 스테이징 행 수 불일치: DB={row_count}, 계산={expected_rows}"
        )
    if status_counts != expected_status:
        raise RuntimeError(
            f"필수값 스테이징 판정 불일치: DB={status_counts}, 계산={expected_status}"
        )
    if missing_counts != expected_missing:
        raise RuntimeError(
            f"필수값 스테이징 누락 통계 불일치: DB={missing_counts}, 계산={expected_missing}"
        )
    return str(snapshot_id)


def main() -> None:
    if not DB_PATH.exists():
        raise FileNotFoundError(DB_PATH)
    sheets, meta = build_rows()
    mandatory_snapshot_id = validate_mandatory_staging(meta)
    # Record the exact DB snapshot used by this Excel output.
    sheets[0][1].append(
        ["필수값 DB 스냅샷", mandatory_snapshot_id, "stg_mandatory_pass is_current=1"]
    )
    sheets[0][4].append("R1")
    meta["mandatory_snapshot_id"] = mandatory_snapshot_id
    xlsx_writer.OUTPUT = OUTPUT
    xlsx_writer.SHEETS = sheets
    xlsx_writer.build_xlsx()
    patch_special_cell_styles(OUTPUT)
    human_review_style = patch_human_review_cell_styles(OUTPUT)
    meta["human_review_style"] = human_review_style
    # The base validator expects the four-sheet sample, so use its low-level ZIP
    # safety checks indirectly and validate the bulk-specific sheet count here.
    from zipfile import ZipFile
    from xml.etree import ElementTree as ET

    with ZipFile(OUTPUT) as archive:
        if archive.testzip() is not None:
            raise ValueError("corrupt XLSX archive")
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        sheet_count = len(list(workbook.find(namespace + "sheets")))
    if sheet_count != 9:
        raise ValueError(f"unexpected sheet count: {sheet_count}")
    print(f"output={OUTPUT}")
    print(json.dumps({**meta, "sheets": sheet_count}, ensure_ascii=False))


if __name__ == "__main__":
    main()
