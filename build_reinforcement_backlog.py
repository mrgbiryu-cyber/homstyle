from __future__ import annotations

import json
import re
import sqlite3
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any

import build_homestyle_bulk_workbook as workbook
import analyze_dimension_notations as notation_analysis
from bulk_homestyle_collect import DB_PATH, RUN_DIR, unpack
from bulk_homestyle_ocr import detail_images, normalize_url


TABLE = "stg_reinforcement_backlog"
LATEST_OUTPUT = RUN_DIR / "reinforcement_backlog_latest.json"
DOCUMENT_OUTPUT = Path("보강대상_분류_및_보강방법_2026-07-22.md")
PARSER_VERSION = "reinforcement-backlog-v1; color-source-depth; dimension-staging; L-deferred"


# Color and finish terms are deliberately separated. A strict color can often be
# normalized automatically. A material/finish term such as oak or chrome may be
# the commercial color value, but still needs a product/option review.
STRICT_COLOR_ALIASES: dict[str, tuple[str, ...]] = {
    "화이트": ("화이트", "흰색", "백색", "white"),
    "블랙": ("블랙", "검정", "검은색", "흑색", "black"),
    "그레이": ("그레이", "회색", "grey", "gray"),
    "차콜": ("차콜", "charcoal"),
    "베이지": ("베이지", "beige"),
    "아이보리": ("아이보리", "ivory"),
    "크림": ("크림", "cream"),
    "브라운": ("브라운", "갈색", "brown"),
    "레드": ("레드", "빨강", "빨간색", "적색", "red"),
    "오렌지": ("오렌지", "주황", "주황색", "orange"),
    "옐로우": ("옐로우", "노랑", "노란색", "yellow"),
    "그린": ("그린", "초록", "초록색", "녹색", "green"),
    "카키": ("카키", "khaki"),
    "민트": ("민트", "mint"),
    "블루": ("블루", "파랑", "파란색", "blue"),
    "네이비": ("네이비", "navy"),
    "퍼플": ("퍼플", "보라", "보라색", "purple"),
    "핑크": ("핑크", "분홍", "분홍색", "pink"),
    "버건디": ("버건디", "burgundy"),
    "코랄": ("코랄", "coral"),
    "골드": ("로즈골드", "골드", "gold"),
    "실버": ("실버", "silver"),
    "투명": ("투명", "transparent", "clear"),
    "멀티컬러": ("멀티컬러", "멀티 칼라", "multi color", "multicolor"),
}

FINISH_COLOR_ALIASES: dict[str, tuple[str, ...]] = {
    "알루미늄": ("알루미늄", "aluminium", "aluminum"),
    "크롬": ("크롬", "chrome"),
    "니켈": ("니켈", "nickel"),
    "브라스": ("브라스", "brass"),
    "브론즈": ("브론즈", "bronze"),
    "코퍼": ("코퍼", "copper"),
    "오크": ("오크", "oak"),
    "월넛": ("월넛", "walnut"),
    "내추럴": ("내추럴", "natural"),
    "티크": ("티크", "teak"),
    "마호가니": ("마호가니", "mahogany"),
    "메이플": ("메이플", "maple"),
    "애쉬": ("애쉬", "ash"),
}

COLOR_STYLE_RE = re.compile(r"색상|색깔|컬러|colou?r", re.I)
COLOR_LABEL_RE = re.compile(
    r"(?:색상|색깔|컬러|colou?r)(?=\s|[:：/·,\-]|$)", re.I
)
COLOR_COUNT_RE = re.compile(r"(?:\d+\s*colou?rs?|\d+\s*(?:색상|컬러))", re.I)
L_W_H_RE = re.compile(
    r"(?:\bL\s*[:=]?\s*\d{2,4}(?:\.\d+)?\s*(?:mm|cm)?\s*[x×]\s*"
    r"W\s*[:=]?\s*\d{2,4}(?:\.\d+)?\s*(?:mm|cm)?\s*[x×]\s*"
    r"H\s*[:=]?\s*\d{2,4}(?:\.\d+)?|\bL\s*[x×]\s*W\s*[x×]\s*H\b)",
    re.I,
)
DIAMETER_RE = re.compile(r"(?:Ø|Φ|⌀|직경)\s*[:=]?\s*\d{2,4}(?:\.\d+)?", re.I)
FLAT_CATEGORIES = {"러그", "액자", "인테리어포스터"}


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def clipped(value: Any, limit: int = 1200) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    return text if len(text) <= limit else text[:limit] + "…"


def text_matches_alias(text: str, alias: str) -> bool:
    if re.fullmatch(r"[A-Za-z ]+", alias):
        return bool(re.search(rf"(?<![A-Za-z]){re.escape(alias)}(?![A-Za-z])", text, re.I))
    return alias.casefold() in text.casefold()


def extract_color_terms(text: str, *, include_finish: bool = True) -> tuple[list[str], list[str]]:
    strict: list[str] = []
    finish: list[str] = []
    for canonical, aliases in STRICT_COLOR_ALIASES.items():
        if any(text_matches_alias(text, alias) for alias in aliases):
            strict.append(canonical)
    if include_finish:
        for canonical, aliases in FINISH_COLOR_ALIASES.items():
            if any(finish_alias_match(text, alias) for alias in aliases):
                finish.append(canonical)
    return strict, finish


def finish_alias_match(text: str, alias: str) -> bool:
    if re.fullmatch(r"[A-Za-z ]+", alias):
        return text_matches_alias(text, alias)
    # Avoid substring false positives such as 티크 in 앤티크. Common finish
    # suffixes are still allowed so that 티크우드/월넛색 can be recognized.
    return bool(
        re.search(
            rf"(?<![가-힣]){re.escape(alias)}(?=$|[^가-힣]|색|컬러|마감|우드)",
            text,
            re.I,
        )
    )


def label_color_terms(text: str) -> tuple[list[str], list[str], str]:
    strict: list[str] = []
    finish: list[str] = []
    evidence: list[str] = []
    for match in COLOR_LABEL_RE.finditer(text or ""):
        window = text[match.start() : match.start() + 140]
        # Light-source phrases are not the body color of a lighting product.
        if re.match(r"색온도|주광색|전구색", window, re.I):
            continue
        window_strict, window_finish = extract_color_terms(window)
        if window_strict or window_finish:
            strict.extend(value for value in window_strict if value not in strict)
            finish.extend(value for value in window_finish if value not in finish)
            evidence.append(clipped(window, 180))
    return strict, finish, " | ".join(evidence[:5])


def ocr_text(ocr_blob: dict[str, Any]) -> str:
    values = [
        str((ocr_blob.get("ocr") or {}).get("text") or ""),
        str(ocr_blob.get("combined_text") or ""),
        str(ocr_blob.get("dimension_text") or ""),
    ]
    unique = list(dict.fromkeys(value for value in values if value))
    return "\n".join(unique)


def raw_option_groups(data: dict[str, Any]) -> list[dict[str, Any]]:
    groups: list[dict[str, Any]] = []
    for group in data.get("purchaseOptions") or []:
        if not isinstance(group, dict):
            continue
        items = [item for item in group.get("items") or [] if isinstance(item, dict)]
        values = [clipped(item.get("name"), 300) for item in items if item.get("name")]
        if not values:
            continue
        groups.append(
            {
                "title": clipped(group.get("title"), 120),
                "values": values,
                "image_urls": [
                    normalize_url(str(item.get("imageUrl") or ""))
                    for item in items
                    if item.get("imageUrl")
                ],
            }
        )
    return groups


def notification_color_candidates(data: dict[str, Any]) -> list[dict[str, str]]:
    result: list[dict[str, str]] = []
    for item in workbook.notification_items(data):
        title = workbook.clean_text(item.get("title"), 200)
        description = workbook.clean_text(item.get("description"), 2000)
        if title and description and COLOR_STYLE_RE.search(title):
            result.append({"title": title, "value": description})
    return result


def color_backlog_record(
    mandatory: sqlite3.Row,
    data: dict[str, Any],
    html_blob: dict[str, Any],
    ocr_blob: dict[str, Any],
    ocr_status: int,
) -> dict[str, Any]:
    name = str(mandatory["product_name"] or data.get("productName") or "")
    notifications = notification_color_candidates(data)
    groups = raw_option_groups(data)
    option_values = [value for group in groups for value in group["values"]]
    option_images = list(
        dict.fromkeys(url for group in groups for url in group["image_urls"] if url)
    )
    option_terms = [extract_color_terms(value) for value in option_values]
    option_all_strict = bool(option_values) and all(strict for strict, _ in option_terms)
    option_all_any = bool(option_values) and all(strict or finish for strict, finish in option_terms)
    option_some_any = any(strict or finish for strict, finish in option_terms)
    option_strict = list(dict.fromkeys(value for strict, _ in option_terms for value in strict))
    option_finish = list(dict.fromkeys(value for _, finish in option_terms for value in finish))
    explicit_option_style = any(COLOR_STYLE_RE.search(group["title"]) for group in groups)

    detail_markup = str(data.get("detailInfo") or "")
    detail_text = workbook.clean_text(detail_markup, 100000)
    detail_strict, detail_finish, detail_evidence = label_color_terms(detail_text)
    ocr_value = ocr_text(ocr_blob)
    ocr_strict, ocr_finish, ocr_evidence = label_color_terms(ocr_value)
    name_strict, name_finish = extract_color_terms(name)
    images = detail_images(detail_markup)
    attempted_urls = {
        normalize_url(str((ocr_blob.get("selected") or {}).get("url") or ""))
    }
    for row in ocr_blob.get("dimension_reinforcements") or []:
        if isinstance(row, dict):
            attempted_urls.add(normalize_url(str(row.get("image_url") or row.get("url") or "")))
    attempted_urls.discard("")

    candidate_value = ""
    candidate_evidence = ""
    source_channel = ""
    confidence = "LOW"
    review_required = 1

    if notifications:
        cause_code = "STRUCTURED_COLOR_FIELD_MISSED"
        cause_label = "상품정보고시 색상 필드가 있으나 기존 키워드/구조가 놓침"
        method = "고시 제목 키워드와 중첩 구조를 확장해 description 값을 재파싱"
        automation = "AUTO_HIGH"
        confidence = "HIGH"
        review_required = 0
        candidate_value = "|".join(dict.fromkeys(row["value"] for row in notifications))
        candidate_evidence = json.dumps(notifications, ensure_ascii=False)
        source_channel = "API_상품정보고시"
        priority = 10
    elif explicit_option_style and option_values:
        cause_code = "OPTION_STYLE_COLOR_MISSED"
        cause_label = "옵션 스타일이 색상인데 기존 옵션 정규화가 놓침"
        method = "색상 옵션 스타일을 인식하고 옵션별 색상값을 병합"
        automation = "AUTO_HIGH"
        confidence = "HIGH"
        review_required = 0
        candidate_value = "|".join(option_values)
        candidate_evidence = json.dumps(groups, ensure_ascii=False)
        source_channel = "API_옵션"
        priority = 10
    elif option_all_strict:
        cause_code = "OPTION_COLOR_DICTIONARY_GAP"
        cause_label = "색상 옵션값은 있으나 옵션 제목이 일반명이고 색상 사전이 부족함"
        method = "색상 동의어 사전을 확장해 일반 Option/제품선택 그룹을 색상으로 재분류"
        automation = "AUTO_HIGH"
        confidence = "HIGH"
        review_required = 0
        candidate_value = "|".join(option_strict)
        candidate_evidence = " | ".join(option_values)
        source_channel = "API_옵션"
        priority = 10
    elif option_all_any:
        cause_code = "OPTION_FINISH_AS_COLOR_REVIEW"
        cause_label = "옵션값에 색상/마감명이 있으나 재질과 색상의 경계 확인이 필요함"
        method = "색상+마감 사전으로 후보를 만들고 상품 옵션 단위로 1회 검수"
        automation = "AUTO_REVIEW"
        confidence = "MEDIUM"
        candidate_value = "|".join(option_strict + option_finish)
        candidate_evidence = " | ".join(option_values)
        source_channel = "API_옵션"
        priority = 20
    elif option_some_any:
        cause_code = "OPTION_COMPOSITE_COLOR_PARSE"
        cause_label = "모델명·사이즈·색상이 한 옵션값에 혼합되어 일부만 색상으로 인식됨"
        method = "복합 옵션값에서 색상/마감 토큰만 분리하고 옵션별 대응관계를 유지"
        automation = "AUTO_REVIEW"
        confidence = "MEDIUM"
        candidate_value = "|".join(option_strict + option_finish)
        candidate_evidence = " | ".join(option_values)
        source_channel = "API_옵션"
        priority = 20
    elif detail_strict or detail_finish:
        cause_code = "HTML_COLOR_LABEL_AVAILABLE"
        cause_label = "상세 HTML 텍스트에 색상 라벨과 값이 있으나 요구필드에 미반영"
        method = "색상/컬러 라벨 주변 값을 파싱하고 옵션값과 교차검증"
        automation = "AUTO_REVIEW"
        confidence = "MEDIUM"
        candidate_value = "|".join(detail_strict + detail_finish)
        candidate_evidence = detail_evidence
        source_channel = "HTML_상세텍스트"
        priority = 20
    elif ocr_strict or ocr_finish:
        cause_code = "OCR_COLOR_LABEL_AVAILABLE"
        cause_label = "상세 이미지 OCR에 색상 라벨과 값이 있으나 색상 파서가 없음"
        method = "OCR의 색상 라벨 주변 토큰을 추출하고 상품명/옵션과 교차검증"
        automation = "OCR_REVIEW"
        confidence = "MEDIUM"
        candidate_value = "|".join(ocr_strict + ocr_finish)
        candidate_evidence = ocr_evidence
        source_channel = "OCR_상세이미지"
        priority = 30
    elif name_strict or name_finish:
        cause_code = "PRODUCT_NAME_COLOR_AVAILABLE"
        cause_label = "상품명에 색상/마감명이 있으나 색상 필드에 미반영"
        method = "상품명 색상 후보를 옵션·대표 이미지와 교차검증 후 반영"
        automation = "AUTO_REVIEW"
        confidence = "MEDIUM"
        candidate_value = "|".join(name_strict + name_finish)
        candidate_evidence = name
        source_channel = "API_상품명"
        priority = 30
    elif option_values and (
        COLOR_COUNT_RE.search(name) or len(option_images) > 1
    ):
        cause_code = "OPTION_COLOR_MEANING_UNRESOLVED"
        cause_label = "색상 옵션으로 보이지만 장문/외국어 옵션명의 의미를 현재 사전이 해석하지 못함"
        method = "외국어 색상·마감 사전 또는 제한형 LLM으로 옵션 의미를 분류하고 이미지로 검수"
        automation = "SEMANTIC_REVIEW"
        confidence = "LOW"
        candidate_value = "|".join(option_values)
        candidate_evidence = json.dumps(groups, ensure_ascii=False)
        source_channel = "API_옵션"
        priority = 40
    elif images and len(attempted_urls) < len(images):
        if option_values:
            cause_code = "NON_COLOR_OPTION_OCR_REMAINING"
            cause_label = "옵션은 있으나 색상 옵션으로 판정할 근거가 없고 미처리 상세 이미지가 있음"
            method = "옵션 의미를 보존한 채 상세 이미지 전체 OCR에서 색상/컬러/마감 라벨을 탐색"
        else:
            cause_code = "NO_COLOR_OPTION_OCR_REMAINING"
            cause_label = "색상 옵션 자체가 없고 아직 OCR하지 않은 상세 이미지가 있음"
            method = "상세 이미지 전체를 순차 OCR해 색상/컬러/마감 라벨이 있는 이미지를 탐색"
        automation = "OCR_REQUIRED"
        confidence = "UNKNOWN"
        candidate_evidence = f"상세 이미지 {len(images)}장, 기존 OCR URL {len(attempted_urls)}개"
        source_channel = "상세이미지"
        priority = 50
    elif images:
        if option_values:
            cause_code = "NON_COLOR_OPTION_VISUAL_INFERENCE"
            cause_label = "옵션은 있으나 색상 근거가 아니며 텍스트/OCR에도 색상값이 없음"
            method = "옵션 이미지와 대표 이미지를 배경 제거 후 시각 색상 모델로 후보화하고 사람 검수"
        else:
            cause_code = "NO_COLOR_OPTION_VISUAL_INFERENCE"
            cause_label = "색상 옵션이 없고 텍스트/OCR에도 색상값이 없어 이미지 픽셀 추론이 필요함"
            method = "대표·상세 이미지를 배경 제거 후 시각 색상 모델로 후보화하고 사람 검수"
        automation = "VISION_REVIEW"
        confidence = "LOW"
        candidate_evidence = f"상세 이미지 {len(images)}장, OCR 상태 {ocr_status}"
        source_channel = "대표/상세이미지"
        priority = 60
    else:
        cause_code = "NO_COLOR_SOURCE"
        cause_label = "옵션·고시·HTML·OCR·상세 이미지에서 색상 근거를 찾지 못함"
        method = "브랜드/공급사 원천 데이터 또는 수기 확인"
        automation = "MANUAL_SOURCE"
        confidence = "UNKNOWN"
        candidate_evidence = f"OCR 상태 {ocr_status}"
        source_channel = "원천부재"
        priority = 90

    details = {
        "has_color_option_style": explicit_option_style,
        "option_group_count": len(groups),
        "option_value_count": len(option_values),
        "option_image_count": len(option_images),
        "notification_color_count": len(notifications),
        "detail_image_count": len(images),
        "attempted_ocr_url_count": len(attempted_urls),
        "ocr_status": ocr_status,
        "name_color_terms": name_strict + name_finish,
        "option_color_terms": option_strict + option_finish,
        "detail_label_color_terms": detail_strict + detail_finish,
        "ocr_label_color_terms": ocr_strict + ocr_finish,
        "no_option": not bool(option_values),
    }
    return {
        "cause_code": cause_code,
        "cause_label": cause_label,
        "source_channel": source_channel,
        "candidate_value": clipped(candidate_value, 3000),
        "candidate_evidence": clipped(candidate_evidence, 4000),
        "proposed_method": method,
        "automation_level": automation,
        "confidence": confidence,
        "priority": priority,
        "review_required": review_required,
        "source_url": f"https://homestyle.lge.co.kr/item?productId={mandatory['product_id']}",
        "details_json": json.dumps(details, ensure_ascii=False),
    }


def dimension_backlog_record(
    mandatory: sqlite3.Row,
    dimension: sqlite3.Row,
    data: dict[str, Any],
    ocr_blob: dict[str, Any],
    notation_rule: str = "",
    notation_context: str = "",
) -> dict[str, Any]:
    current_status = str(dimension["current_status"] or "")
    missing_axes = str(dimension["missing_axes"] or "")
    action = str(dimension["suggested_action"] or "")
    rule = str(dimension["candidate_rule"] or "")
    small = str(mandatory["small_category_value"] or dimension["small_category"] or "")
    source_text = "\n".join(
        [
            workbook.clean_text(data.get("detailInfo"), 100000),
            ocr_text(ocr_blob),
            str(dimension["dimension_records_json"] or ""),
            str(dimension["size_options_json"] or ""),
        ]
    )
    has_lwh = bool(L_W_H_RE.search(source_text)) or "L-W-H" in notation_rule or "L-D-H" in notation_rule
    has_diameter = bool(DIAMETER_RE.search(source_text)) or "직경" in notation_rule

    raw_candidate_value = str(dimension["candidate_dimensions_json"] or "")
    candidate_value = (
        raw_candidate_value
        if raw_candidate_value.strip() not in {"", "[]", "{}", "null"}
        else ""
    )
    candidate_evidence = ""
    confidence = str(dimension["candidate_confidence"] or "UNKNOWN")
    review_required = 1

    if rule == "OPTION_UNLABELED_WDH":
        cause_code = "OPTION_WDH_AVAILABLE"
        cause_label = "옵션에 3축 숫자가 있으나 규격 필드에 미반영"
        method = "옵션별 W×D×H를 숫자(mm)로 정규화하고 원문과 자동 대조"
        automation = "AUTO_HIGH"
        priority = 10
        review_required = 0
    elif rule in {"OPTION_FLAT_WD", "OPTION_FLAT_WH"}:
        cause_code = "FLAT_PRODUCT_AXIS_POLICY_REQUIRED"
        cause_label = "러그/액자/포스터 등 평면형 상품은 2축 규격만 있고 세 번째 축 정책이 없음"
        method = "카테고리별 W/D/H 축과 비적용 축 정책을 확정한 뒤 2축 옵션값을 반영"
        automation = "POLICY_REQUIRED"
        priority = 20
    elif rule == "OPTION_WD_PLUS_SINGLE_H":
        cause_code = "OPTION_PAIR_PLUS_SHARED_AXIS"
        cause_label = "옵션 2축과 공통 높이가 분리되어 있어 결합 규칙이 필요함"
        method = "옵션별 W×D와 단일 H의 동일 모델 근거를 확인 후 결합"
        automation = "AUTO_REVIEW"
        priority = 20
    elif rule == "BED_STANDARD_SIZE_CODE":
        cause_code = "BED_SIZE_CODE_NEEDS_MODEL_TABLE"
        cause_label = "SS/Q/K 등 표준 코드만 있어 실제 제품 외형 치수를 확정할 수 없음"
        method = "브랜드·모델별 공식 규격표를 연결해 옵션 코드를 실제 W/D/H로 치환"
        automation = "SOURCE_LOOKUP"
        priority = 30
    elif has_lwh:
        cause_code = "L_NOTATION_INTERPRETATION_REQUIRED"
        cause_label = "L/W/H 표기가 있으나 L의 의미를 확인하지 않고 W로 치환할 수 없음"
        method = "카테고리·도면 방향으로 L이 길이인지 깊이인지 판정한 후 W/D/H에 매핑"
        automation = "SEMANTIC_REVIEW"
        priority = 30
        direct_match = L_W_H_RE.search(source_text)
        candidate_evidence = clipped(
            notation_context or (direct_match.group(0) if direct_match else ""), 500
        )
    elif has_diameter:
        cause_code = "DIAMETER_AXIS_MAPPING_REQUIRED"
        cause_label = "Ø/Φ/직경 표기가 있으나 원형 제품의 W/D 축 매핑 규칙이 없음"
        method = "원형 카테고리 규칙으로 직경을 W와 D에 적용하고 H를 별도 추출"
        automation = "AUTO_REVIEW"
        priority = 30
        direct_match = DIAMETER_RE.search(source_text)
        candidate_evidence = clipped(
            notation_context or (direct_match.group(0) if direct_match else ""), 500
        )
    elif current_status == "부분확보" and small in FLAT_CATEGORIES:
        cause_code = "FLAT_PRODUCT_PARTIAL_DIMENSION"
        cause_label = "평면형 상품에서 2축만 확보되어 현행 3축 필수 기준을 충족하지 못함"
        method = "카테고리 축/비적용 정책을 먼저 정한 뒤 확보된 숫자를 재사용"
        automation = "POLICY_REQUIRED"
        priority = 30
    elif current_status == "부분확보":
        cause_code = "PARTIAL_DIMENSION_AXES"
        cause_label = f"규격 일부 축만 확보됨(누락: {missing_axes})"
        method = "옵션·상세 이미지 전체 OCR에서 누락 축만 추가 추출하고 기존 축과 교차검증"
        automation = "OCR_REVIEW"
        priority = 40
    else:
        cause_code = "NO_PARSED_DIMENSION"
        cause_label = "API/HTML/현재 OCR에서 유효한 W/D/H 숫자 조합을 파싱하지 못함"
        method = "상세 이미지 전체 OCR 및 옵션/특수표기 정규화 후 재파싱"
        automation = "OCR_REQUIRED"
        priority = 50

    action_note = {
        "OCR_NEXT_IMAGE": "미처리 상세 이미지를 순차 OCR",
        "OCR_URL_RETRY": "이미지 URL 정규화·재다운로드 후 OCR",
        "OCR_RECHECK_IMAGE": "전처리와 Windows/Paddle/Tesseract 등 다중 OCR 재판독",
        "MANUAL_SOURCE_REQUIRED": "브랜드/공급사 원천 규격표 또는 수기 확인",
        "OPTION_POLICY_REVIEW": "옵션 축·카테고리 정책 검토",
        "OPTION_AUTO_VERIFY": "옵션 숫자 자동 대조",
    }.get(action, action)
    if action_note and action_note not in method:
        method = f"{method}; 다음 실행: {action_note}"
    if action == "OCR_URL_RETRY":
        automation = "OCR_RETRY"
    elif action == "OCR_RECHECK_IMAGE" and automation == "OCR_REQUIRED":
        automation = "OCR_ENGINE_REVIEW"
    elif action == "MANUAL_SOURCE_REQUIRED":
        automation = "MANUAL_SOURCE"
        priority = 90

    details = {
        "current_status": current_status,
        "current_w_mm": dimension["current_w_mm"],
        "current_d_mm": dimension["current_d_mm"],
        "current_h_mm": dimension["current_h_mm"],
        "missing_axes": missing_axes,
        "candidate_rule": rule or None,
        "candidate_confidence": dimension["candidate_confidence"],
        "size_option_count": dimension["size_option_count"],
        "detail_image_count": dimension["detail_image_count"],
        "alternate_image_count": dimension["alternate_image_count"],
        "ocr_status": dimension["ocr_status"],
        "suggested_action": action,
        "has_lwh_notation": has_lwh,
        "has_diameter_notation": has_diameter,
        "notation_rule": notation_rule or None,
    }
    return {
        "cause_code": cause_code,
        "cause_label": cause_label,
        "source_channel": "규격_통합스테이징",
        "candidate_value": clipped(candidate_value, 3000),
        "candidate_evidence": candidate_evidence,
        "proposed_method": method,
        "automation_level": automation,
        "confidence": confidence,
        "priority": priority,
        "review_required": review_required,
        "source_url": f"https://homestyle.lge.co.kr/item?productId={mandatory['product_id']}",
        "details_json": json.dumps(details, ensure_ascii=False),
    }


def create_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        f"""
        CREATE TABLE IF NOT EXISTS {TABLE} (
            snapshot_id TEXT NOT NULL,
            assessed_at TEXT NOT NULL,
            is_current INTEGER NOT NULL DEFAULT 1,
            parser_version TEXT NOT NULL,
            product_id TEXT NOT NULL,
            product_name TEXT,
            small_category TEXT,
            missing_field TEXT NOT NULL,
            missing_combination TEXT,
            cause_code TEXT NOT NULL,
            cause_label TEXT NOT NULL,
            source_channel TEXT,
            candidate_value TEXT,
            candidate_evidence TEXT,
            proposed_method TEXT NOT NULL,
            automation_level TEXT NOT NULL,
            confidence TEXT,
            priority INTEGER NOT NULL,
            review_required INTEGER NOT NULL DEFAULT 1,
            source_url TEXT,
            details_json TEXT,
            work_status TEXT NOT NULL DEFAULT '대기',
            applied INTEGER NOT NULL DEFAULT 0,
            PRIMARY KEY(snapshot_id,product_id,missing_field)
        );
        CREATE INDEX IF NOT EXISTS idx_reinforcement_current_field
            ON {TABLE}(is_current,missing_field);
        CREATE INDEX IF NOT EXISTS idx_reinforcement_current_cause
            ON {TABLE}(is_current,missing_field,cause_code);
        CREATE INDEX IF NOT EXISTS idx_reinforcement_current_priority
            ON {TABLE}(is_current,priority,automation_level);

        DROP VIEW IF EXISTS vw_reinforcement_backlog_current_summary;
        CREATE VIEW vw_reinforcement_backlog_current_summary AS
        SELECT missing_field,cause_code,cause_label,automation_level,confidence,
               COUNT(*) AS product_count
        FROM {TABLE}
        WHERE is_current=1
        GROUP BY missing_field,cause_code,cause_label,automation_level,confidence;

        DROP VIEW IF EXISTS vw_reinforcement_backlog_current_products;
        CREATE VIEW vw_reinforcement_backlog_current_products AS
        SELECT product_id,product_name,small_category,missing_combination,
               COUNT(*) AS missing_field_count,
               GROUP_CONCAT(missing_field,'|') AS missing_fields,
               MIN(priority) AS first_priority
        FROM {TABLE}
        WHERE is_current=1
        GROUP BY product_id,product_name,small_category,missing_combination;
        """
    )


def build_snapshot() -> dict[str, Any]:
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    mandatory_rows = connection.execute(
        """
        SELECT * FROM stg_mandatory_pass
        WHERE is_current=1 AND final_status='보강대상'
        ORDER BY product_id
        """
    ).fetchall()
    source_by_id = {
        row["product_id"]: row
        for row in connection.execute(
            "SELECT product_id,goods_blob,html_blob,qna_blob,ocr_status,ocr_blob FROM sources"
        )
    }
    dimension_by_id = {
        row["product_id"]: row
        for row in connection.execute(
            "SELECT * FROM stg_dimension_reinforcement WHERE is_current=1"
        )
    }

    assessed_at = now_text()
    snapshot_id = assessed_at.replace(":", "").replace("+", "_")
    inserts: list[tuple[Any, ...]] = []
    cause_counts: Counter[tuple[str, str]] = Counter()
    automation_counts: Counter[tuple[str, str]] = Counter()
    combination_counts: Counter[str] = Counter()

    for mandatory in mandatory_rows:
        product_id = mandatory["product_id"]
        source = source_by_id[product_id]
        goods_payload = unpack(source["goods_blob"]) or {}
        data = goods_payload.get("data") or {}
        html_blob = unpack(source["html_blob"]) or {}
        qna_blob = unpack(source["qna_blob"]) or {}
        ocr_blob = unpack(source["ocr_blob"]) or {}
        color_missing = int(mandatory["color_ok"] or 0) == 0
        size_missing = int(mandatory["size_wdh_ok"] or 0) == 0
        missing_fields = []
        if color_missing:
            missing_fields.append("색상")
        if size_missing:
            missing_fields.append("규격(W/D/H)")
        combination = "+".join(missing_fields)
        combination_counts[combination] += 1

        records: list[tuple[str, dict[str, Any]]] = []
        if color_missing:
            records.append(
                (
                    "색상",
                    color_backlog_record(
                        mandatory, data, html_blob, ocr_blob, int(source["ocr_status"] or 0)
                    ),
                )
            )
        if size_missing:
            dimension = dimension_by_id.get(product_id)
            if dimension is None:
                raise RuntimeError(f"dimension staging row missing: {product_id}")
            diagnostic_sources = notation_analysis.product_sources(
                {
                    "data": data,
                    "html": html_blob,
                    "qna": qna_blob,
                    "ocr": ocr_blob,
                }
            )
            _, notation_rule, notation_context = notation_analysis.strict_complete(
                diagnostic_sources
            )
            records.append(
                (
                    "규격(W/D/H)",
                    dimension_backlog_record(
                        mandatory,
                        dimension,
                        data,
                        ocr_blob,
                        notation_rule,
                        notation_context,
                    ),
                )
            )

        for missing_field, record in records:
            cause_counts[(missing_field, record["cause_code"])] += 1
            automation_counts[(missing_field, record["automation_level"])] += 1
            inserts.append(
                (
                    snapshot_id,
                    assessed_at,
                    1,
                    PARSER_VERSION,
                    product_id,
                    mandatory["product_name"],
                    mandatory["small_category_value"],
                    missing_field,
                    combination,
                    record["cause_code"],
                    record["cause_label"],
                    record["source_channel"],
                    record["candidate_value"] or None,
                    record["candidate_evidence"] or None,
                    record["proposed_method"],
                    record["automation_level"],
                    record["confidence"],
                    record["priority"],
                    record["review_required"],
                    record["source_url"],
                    record["details_json"],
                    "대기",
                    0,
                )
            )

    create_schema(connection)
    placeholders = ",".join("?" for _ in range(23))
    with connection:
        connection.execute(f"UPDATE {TABLE} SET is_current=0 WHERE is_current=1")
        connection.executemany(f"INSERT INTO {TABLE} VALUES ({placeholders})", inserts)

    summary_rows = [
        dict(row)
        for row in connection.execute(
            """
            SELECT missing_field,cause_code,cause_label,automation_level,confidence,product_count
            FROM vw_reinforcement_backlog_current_summary
            ORDER BY missing_field,product_count DESC,cause_code
            """
        )
    ]
    result = {
        "database": str(DB_PATH),
        "table": TABLE,
        "snapshot_id": snapshot_id,
        "assessed_at": assessed_at,
        "parser_version": PARSER_VERSION,
        "reinforcement_products": len(mandatory_rows),
        "backlog_rows": len(inserts),
        "missing_combination_counts": dict(combination_counts),
        "cause_counts": [
            {"missing_field": field, "cause_code": cause, "product_count": count}
            for (field, cause), count in sorted(cause_counts.items())
        ],
        "automation_counts": [
            {"missing_field": field, "automation_level": level, "product_count": count}
            for (field, level), count in sorted(automation_counts.items())
        ],
        "summary": summary_rows,
        "views": [
            "vw_reinforcement_backlog_current_summary",
            "vw_reinforcement_backlog_current_products",
        ],
        "excel_written": False,
        "notes": [
            "OCR은 이미지 안에 적힌 색상명을 읽는 단계이며, 픽셀 색상 추론은 VISION_REVIEW로 분리했다.",
            "후보값은 아직 고객용 필드와 필수값 PASS 통계에 적용하지 않았다.",
            "L 표기는 의미 확인 전 W/D/H에 자동 치환하지 않는다.",
        ],
    }
    connection.close()
    LATEST_OUTPUT.write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    write_document(result)
    return result


def write_document(result: dict[str, Any]) -> None:
    by_field: dict[str, list[dict[str, Any]]] = {}
    for row in result["summary"]:
        by_field.setdefault(row["missing_field"], []).append(row)

    lines = [
        "# 보강대상 분류 및 보강방법 (2026-07-22)",
        "",
        "현재 필수값 판정에서 세트 구성 실제 ID는 제외되어 있으며, 보강대상은 색상과 규격만 집계한다.",
        "",
        "## 전체 현황",
        "",
        f"- 보강대상 상품: {result['reinforcement_products']:,}개",
        f"- 보강 작업 행: {result['backlog_rows']:,}행(상품×누락 필드)",
    ]
    for combination, count in result["missing_combination_counts"].items():
        lines.append(f"- {combination}: {count:,}개")

    for field in ("색상", "규격(W/D/H)"):
        rows = by_field.get(field, [])
        lines.extend(
            [
                "",
                f"## {field} 보강 분류",
                "",
                "| 원인 코드 | 실제 원인 | 보강 방식 | 자동화 구분 | 건수 |",
                "|---|---|---|---:|---:|",
            ]
        )
        for row in rows:
            method = next(
                (
                    item["proposed_method"]
                    for item in current_records_for_document(
                        result,
                        field,
                        row["cause_code"],
                        row["automation_level"],
                    )
                ),
                "DB 백로그 참조",
            )
            lines.append(
                f"| {row['cause_code']} | {row['cause_label']} | {method} | "
                f"{row['automation_level']} | {row['product_count']:,} |"
            )

    lines.extend(
        [
            "",
            "## 판정 원칙",
            "",
            "- OCR: 이미지에 글자로 적힌 `색상: 크림`, `Color: Black` 같은 문구를 읽는다.",
            "- 이미지 시각 추론: 글자가 없는 제품 사진의 픽셀에서 색상 후보를 만든다. 조명·그림자·배경 영향 때문에 자동 확정하지 않는다.",
            "- 옵션/고시/API 후보는 원문 값을 보존하고, 색상 동의어 정규화 값은 별도 관리한다.",
            "- L/W/H의 L은 카테고리와 도면 방향을 확인한 뒤 W 또는 D로 매핑한다.",
            "- 이 문서는 보강 방법 분류이며 후보값을 PASS 통계에 적용한 결과가 아니다.",
            "",
            "## DB 사용",
            "",
            f"- 상세 테이블: `{TABLE}`",
            "- 원인별 요약: `vw_reinforcement_backlog_current_summary`",
            "- 상품별 누락 요약: `vw_reinforcement_backlog_current_products`",
        ]
    )
    DOCUMENT_OUTPUT.write_text("\n".join(lines) + "\n", encoding="utf-8")


def current_records_for_document(
    result: dict[str, Any], field: str, cause_code: str, automation_level: str
) -> list[dict[str, str]]:
    # The document needs one representative method per cause. Read-only access
    # keeps the JSON summary compact while the DB remains the source of detail.
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    rows = [
        dict(row)
        for row in connection.execute(
            f"""
            SELECT proposed_method FROM {TABLE}
            WHERE is_current=1 AND missing_field=? AND cause_code=?
              AND automation_level=?
            LIMIT 1
            """,
            (field, cause_code, automation_level),
        )
    ]
    connection.close()
    return rows


if __name__ == "__main__":
    print(json.dumps(build_snapshot(), ensure_ascii=False, indent=2))
