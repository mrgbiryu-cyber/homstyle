from __future__ import annotations

import re
from dataclasses import dataclass, asdict
from typing import Any, Iterable

from low_dimension_quality_policy import assess_low_dimension


NUMBER = r"\d{1,5}(?:[.,]\d+)?"
UNIT = (
    r"(?:(?:mm|cm|㎜|㎝|inch(?:es)?|in)(?![A-Za-z])|"
    r"(?<![A-Za-z])m(?![A-Za-z])|[″”\"])"
)
MULTIPLY = r"(?:x|×|X|＊|\*|\)\()"

POSITIVE_SECTION_RE = re.compile(
    r"(?:한\s*눈에\s*보기|제품\s*(?:사이즈|크기|규격|치수)|"
    r"상품\s*(?:사이즈|크기|규격|치수)|사이즈\s*/?\s*규격|"
    r"(?:^|[\s|:/(\[])(?:사이즈|규격|치수)(?:[\s|:/(\[]|$)|"
    r"dimensions?|measurements?|product\s+(?:info(?:rmation)?|description)|"
    r"specifications?|detail\s*(?:size|dimensions?)?|size)",
    re.I,
)
DELIVERY_RE = re.compile(
    r"(?:배송|반입|엘리베이터|엘레베이터|현관문|출입문|사다리차|"
    r"설치\s*(?:안내|조건|공간|기사|비용|불가)|배송\s*(?:및|/)\s*설치|"
    r"포장|박스|통로|계단|해피콜|배송조건|배송\s*가이드|shipping|delivery|elevator)",
    re.I,
)
COMPONENT_RE = re.compile(
    r"(?:가드|발통|다리\s*높이|헤드보드|프레임|구성품|부품|선반\s*내부|"
    r"수납\s*내부|포장\s*크기|package|component)",
    re.I,
)
TOLERANCE_RE = re.compile(
    r"(?:오차\s*범위|오차범위|측정\s*오차|제작\s*오차|±|\+/-|허용\s*오차)",
    re.I,
)
LINEUP_RE = re.compile(r"(?:라인업|line\s*up|collection|series)", re.I)
LOCAL_COMPONENT_LABEL_RE = re.compile(
    r"(?:상부\s*도어|다용도\s*선반|다용도\s*꽂이|(?:원형\s*)?자석|"
    r"액세서리|발걸이|가방걸이|LED\s*책상\s*조명)",
    re.I,
)
LOCAL_MAIN_PRODUCT_LABEL_RE = re.compile(
    r"(?:다리형\s*책상\s*\+\s*\d+\s*단\s*책상장|"
    r"\d+\s*단\s*다리형\s*책상세트|제품\s*(?:사이즈|규격)|MODEL\s+SIZE)",
    re.I,
)

AXIS_VALUE_RE = re.compile(
    rf"(?P<label>헤드\s*높이|가로|너비|폭|세로|깊이|높이|길이|지름|직경|반지름|"
    rf"(?<![A-Za-z])[WDLHR](?![A-Za-z]))"
    rf"\s*[:=]?\s*\(?\s*(?P<number>{NUMBER})"
    rf"(?:\s*(?:-|~|–|—|to)\s*(?P<range_end>{NUMBER}))?"
    rf"\s*(?:\(\s*(?P<paren_unit>{UNIT})\s*\)|"
    rf"(?P<unit>{UNIT})|"
    rf"\((?!\s*(?:mm|cm|㎜|㎝|m|inch|in)\s*\))[^)]{{0,40}}\))?",
    re.I,
)
TRIPLE_RE = re.compile(
    rf"(?<![\d.])(?P<a>{NUMBER})(?:\s*(?:-|~|–|—|to)\s*(?P<ae>{NUMBER}))?"
    rf"\s*(?P<ua>{UNIT})?\s*{MULTIPLY}\s*"
    rf"(?P<b>{NUMBER})(?:\s*(?:-|~|–|—|to)\s*(?P<be>{NUMBER}))?"
    rf"\s*(?P<ub>{UNIT})?\s*{MULTIPLY}\s*"
    rf"(?P<c>{NUMBER})(?:\s*(?:-|~|–|—|to)\s*(?P<ce>{NUMBER}))?"
    rf"\s*(?P<uc>{UNIT})?(?![\d.])",
    re.I,
)
PAIR_RE = re.compile(
    rf"(?<![\d.])(?P<a>{NUMBER})(?:\s*(?:-|~|–|—|to)\s*(?P<ae>{NUMBER}))?"
    rf"\s*(?P<ua>{UNIT})?\s*{MULTIPLY}\s*"
    rf"(?P<b>{NUMBER})(?:\s*(?:-|~|–|—|to)\s*(?P<be>{NUMBER}))?"
    rf"\s*(?P<ub>{UNIT})?(?!\s*{MULTIPLY})(?![\d.])",
    re.I,
)
OPTION_LABEL_RE = re.compile(
    r"(?:^|[\s(\[/])(?P<label>LQ|DD|SS|DS|LK|K3|CK|RK|KI|TW|Q|K|M|L|S)"
    r"(?:[\s)\]/:]|$)",
    re.I,
)

AREA_2D_TERMS = ("액자", "포스터", "아트웍", "러그", "카펫", "매트", "스프레드")
FURNITURE_TERMS = (
    "소파", "테이블", "식탁", "책상", "의자", "체어", "스툴", "벤치",
    "수납장", "거실장", "협탁", "침대", "선반", "트롤리", "장",
)
ROUND_TERMS = ("원형", "round", "라운드", "pedestal")
OVAL_TERMS = ("타원", "oval")

TYPE_TOKENS = {
    "side": ("side", "사이드"),
    "sofa": ("sofa", "소파"),
    "table": ("table", "테이블", "식탁", "책상"),
    "chair": ("chair", "체어", "의자"),
    "bench": ("bench", "벤치"),
    "bed": ("bed", "침대"),
    "guard": ("guard", "가드"),
    "stool": ("stool", "스툴"),
}


def compact_text(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def normalized_ocr_text(value: str) -> str:
    text = str(value or "").replace("㎜", "mm").replace("㎝", "cm")
    text = text.replace(")( ", " x ").replace(" )(", " x ").replace(")(", " x ")
    # Layout OCR can split thousands groups at a visual gap while reading
    # specification tables, for example "폭 1 500" or "길이2 000".
    # Join only a leading 1-2 digit group followed by exactly three digits
    # immediately after a dimension axis label.
    text = re.sub(
        r"(?i)(?P<label>width|depth|height|[WDLH]|가로|세로|폭|길이|깊이|높이)"
        r"\s*(?P<head>\d{1,2})\s+(?P<tail>\d{3})(?!\d)",
        lambda match: (
            f"{match.group('label')}{match.group('head')}{match.group('tail')}"
        ),
        text,
    )
    text = re.sub(r"\(\s*니\s*(?=W\s*[I1\d])", "(L) ", text)
    text = re.sub(
        r"(?i)\b([WDLH])\s*[I|l]\s*(?=\d{2,4})",
        lambda match: f"{match.group(1)}1",
        text,
    )
    text = re.sub(r"(?<=\d)\s+[Oo](?=\s*(?:mm|cm)\b)", "0", text)
    text = text.replace("해드", "헤드")
    text = re.sub(r"세로\s*[기I]\s*,?(\d{3})", r"세로2,\1", text)
    text = re.sub(
        r"헤드\s*높[0Oo]?\s*(?:II|11|I|l)?\s*,?(\d{3})",
        r"헤드높이1,\1",
        text,
    )
    text = re.sub(r"2\s*卜로\s*(\d[\d,]{3,5})", r"가로\1", text)
    text = re.sub(r"가뢰\s*,?(\d{3})", r"가로1,\1", text)
    # In ACE bed INFO drawings, Windows OCR repeatedly reads the leading
    # thousands digit 2 as the Hangul glyph "기" (for example 가로기600).
    # Restrict the repair to 가로 + 기 + exactly three digits so ordinary
    # occurrences of 기 are not converted.
    text = re.sub(r"가로\s*기\s*,?(\d{3})(?!\d)", r"가로2,\1", text)
    text = re.sub(r"\bDetaiI\b", "Detail", text, flags=re.I)
    return text


def parse_number(value: str) -> float:
    raw = value.strip()
    if "," in raw and "." not in raw:
        tail = raw.rsplit(",", 1)[1]
        raw = raw.replace(",", "") if len(tail) == 3 else raw.replace(",", ".")
    else:
        raw = raw.replace(",", "")
    return float(raw)


def canonical_unit(value: str) -> str:
    unit = str(value or "").strip().casefold()
    return {
        "㎜": "mm", "㎝": "cm", "inches": "inch", "in": "inch",
        '"': "inch", "″": "inch", "”": "inch",
    }.get(unit, unit)


def to_mm(value: float | None, unit: str) -> float | None:
    if value is None:
        return None
    factor = {"mm": 1.0, "cm": 10.0, "m": 1000.0, "inch": 25.4}.get(
        canonical_unit(unit)
    )
    return value * factor if factor is not None else None


def inferred_unit(values: Iterable[float]) -> str:
    numbers = list(values)
    return "cm" if numbers and max(numbers) <= 300 else "mm"


def category_profile(product_name: str, small_category: str) -> str:
    text = f"{product_name} {small_category}".casefold()
    if any(term.casefold() in text for term in ROUND_TERMS):
        return "ROUND_OBJECT"
    if any(term.casefold() in text for term in OVAL_TERMS):
        return "OVAL_OBJECT"
    if any(term.casefold() in text for term in FURNITURE_TERMS):
        return "FURNITURE_3D"
    if any(term.casefold() in text for term in AREA_2D_TERMS):
        return "AREA_2D"
    return "GENERIC_OBJECT"


def axis_name(label: str, profile: str) -> str:
    clean = re.sub(r"\s+", "", label).upper()
    mapping = {
        "W": "W", "D": "D", "H": "H", "L": "L", "R": "R",
        "가로": "W", "너비": "W", "폭": "W",
        "깊이": "D", "높이": "H", "헤드높이": "H",
        "길이": "L", "지름": "R", "직경": "R", "반지름": "RADIUS",
    }
    if clean == "세로":
        return "H" if profile == "AREA_2D" else "D"
    return mapping.get(clean, clean)


def context_window(text: str, start: int, end: int, radius: int = 260) -> str:
    return compact_text(text[max(0, start - radius) : min(len(text), end + radius)])


def nearest_heading_context(text: str, start: int, end: int) -> str:
    before = text[max(0, start - 500) : start]
    after = text[end : min(len(text), end + 140)]
    # Keep enough text before the nearest heading to retain model codes such as
    # BMA1165.  Section ownership is decided independently by
    # nearest_section_signal(), so retaining this evidence does not weaken the
    # delivery/component exclusions.
    return compact_text(before + " " + text[start:end] + " " + after)


def nearest_section_signal(text: str, start: int, end: int) -> tuple[str, bool]:
    """Return the closest preceding section signal and whether it is positive.

    OCR text often concatenates product and delivery blocks.  Looking for any
    keyword in a wide context therefore rejects valid product dimensions.  The
    nearest preceding heading is a better approximation of the image layout.
    """

    left = text[max(0, start - 500) : start]
    signals: list[tuple[int, str, str]] = []
    signals.extend(
        (match.start(), "POSITIVE", compact_text(match.group()))
        for match in POSITIVE_SECTION_RE.finditer(left)
    )
    signals.extend(
        (match.start(), "DELIVERY", compact_text(match.group()))
        for match in DELIVERY_RE.finditer(left)
    )
    signals.extend(
        (match.start(), "COMPONENT", compact_text(match.group()))
        for match in COMPONENT_RE.finditer(left)
    )
    signals.extend(
        (match.start(), "TOLERANCE", compact_text(match.group()))
        for match in TOLERANCE_RE.finditer(left)
    )
    if signals:
        position, role, token = max(signals, key=lambda item: item[0])
        # A bare "규격/치수" inside "엘리베이터 규격" is not a product-size
        # heading.  Let a nearby hard-negative heading own that block.
        if role == "POSITIVE" and token in {"사이즈", "규격", "치수"}:
            negatives = [
                item
                for item in signals
                if item[1] in {"DELIVERY", "COMPONENT"} and 0 <= position - item[0] <= 100
            ]
            if negatives:
                _, role, _ = max(negatives, key=lambda item: item[0])
        return role, role == "POSITIVE"

    right = text[end : min(len(text), end + 100)]
    right_signals: list[tuple[int, str]] = []
    right_signals.extend((match.start(), "DELIVERY") for match in DELIVERY_RE.finditer(right))
    right_signals.extend((match.start(), "TOLERANCE") for match in TOLERANCE_RE.finditer(right))
    if right_signals:
        _, role = min(right_signals, key=lambda item: item[0])
        return role, False
    return "UNSCOPED", False


def product_type_tokens(text: str) -> set[str]:
    lower = text.casefold()
    return {
        canonical
        for canonical, variants in TYPE_TOKENS.items()
        if any(variant.casefold() in lower for variant in variants)
    }


def title_numbers(product_name: str) -> set[int]:
    return {
        int(value.replace(",", ""))
        for value in re.findall(r"(?<!\d)(\d{3,4})(?!\d)", product_name)
    }


def title_option_codes(product_name: str) -> set[str]:
    return {
        match.group(1).upper()
        for match in re.finditer(
            r"(?:^|[\s(/])(LQ|DD|SS|DS|LK|K3|CK|RK|KI|TW|Q|K)(?:[\s()/]|$)",
            product_name,
            re.I,
        )
    }


def product_model_tokens(value: str) -> set[str]:
    tokens: set[str] = set()
    for match in re.finditer(
        r"(?<![A-Za-z0-9])[A-Za-z]{2,}\s*[-/]?\s*\d{2,}"
        r"(?:\s*[-/]\s*[A-Za-z0-9]+){0,3}",
        value,
    ):
        token = re.sub(r"[^A-Za-z0-9]", "", match.group()).upper()
        if len(token) >= 5:
            tokens.add(token)
            # The base model is also useful when OCR changes LC to L/C.
            base = re.match(r"[A-Z]{2,}\d{2,}", token)
            if base:
                tokens.add(base.group())
    return tokens


def option_label(context: str, match_start: int = 0) -> str:
    prefix = context[max(0, match_start - 80) : match_start]
    matches: list[tuple[int, int, str]] = [
        (match.start(), match.end(), match.group("label").upper())
        for match in OPTION_LABEL_RE.finditer(prefix)
    ]
    korean_labels = {
        "슈퍼싱글": "SS",
        "레귤러킹": "RK",
        "칼킹": "CK",
        "라지킹": "LK",
        "퀸": "Q",
        "트윈": "TW",
    }
    for token, label in korean_labels.items():
        for match in re.finditer(re.escape(token), prefix, re.I):
            matches.append((match.start(), match.end(), label))
    if not matches:
        return ""
    _, label_end, label = max(matches, key=lambda item: item[0])
    between = prefix[label_end:]
    # OCR flattens option tables. Do not carry an option label past a
    # complete preceding dimension row into the following row.
    if re.search(
        r"(?i)(?:width|[W]|가로|폭)\s*[:=]?\s*\d",
        between,
    ):
        return ""
    return label


def title_option_width_match(product_name: str, width_mm: float | None) -> bool:
    if width_mm is None:
        return False
    expected = {
        "SS": 1100,
        "DS": 1200,
        "DD": 1400,
        "Q": 1500,
        "LQ": 1500,
        "K": 1650,
        "KI": 1650,
        "K3": 1650,
        "LK": 1800,
        "CK": 1830,
        "RK": 1930,
    }
    return any(
        code in expected and abs(float(width_mm) - expected[code]) <= 70
        for code in title_option_codes(product_name)
    )


def type_match_status(product_name: str, context: str) -> str:
    product_types = product_type_tokens(product_name)
    context_types = product_type_tokens(context)
    if not product_types or not context_types:
        return "UNKNOWN"
    if product_types & context_types:
        if "side" in product_types and "sofa" in context_types and "side" not in context_types:
            return "MISMATCH"
        if "sofa" in product_types and "side" in context_types and "sofa" not in context_types:
            return "MISMATCH"
        return "MATCH"
    if product_types <= {"table"} or context_types <= {"table"}:
        return "UNKNOWN"
    return "MISMATCH"


@dataclass
class DimensionCandidate:
    rule_id: str
    raw_notation: str
    context_text: str
    section_role: str
    candidate_role: str
    source_axis_signature: str
    normalized_axis_mapping: str
    option_label: str
    shape_type: str
    unit_status: str
    unit_text: str
    w_raw: float | None = None
    d_raw: float | None = None
    h_raw: float | None = None
    l_raw: float | None = None
    r_raw: float | None = None
    value_1_raw: float | None = None
    value_2_raw: float | None = None
    value_3_raw: float | None = None
    w_mm: float | None = None
    d_mm: float | None = None
    h_mm: float | None = None
    diameter_mm: float | None = None
    product_name_match_score: int = 0
    candidate_score: int = 0
    decision_status: str = "HUMAN_REVIEW"
    rejection_reason: str = ""

    def as_dict(self) -> dict[str, Any]:
        return asdict(self)


def context_role(
    product_name: str,
    context: str,
    *,
    has_positive_heading: bool,
    nearest_signal: str = "UNSCOPED",
) -> tuple[str, str, str]:
    type_status = type_match_status(product_name, context)
    if nearest_signal == "DELIVERY":
        return "DELIVERY_OR_INSTALLATION", "DELIVERY_CLEARANCE", "배송·설치 문맥"
    if nearest_signal == "COMPONENT" and not COMPONENT_RE.search(product_name):
        return "COMPONENT", "COMPONENT_DIMENSION", "구성품 문맥"
    if type_status == "MISMATCH" and (
        LINEUP_RE.search(context) or len(product_type_tokens(context)) >= 2
    ):
        return "LINEUP", "LINEUP_OTHER_MODEL", "상품명과 다른 라인업"
    if nearest_signal == "TOLERANCE" and not has_positive_heading:
        return "TOLERANCE", "MEASUREMENT_TOLERANCE", "오차·허용범위 문맥"
    return (
        "PRODUCT_SIZE_SECTION" if has_positive_heading else "UNSCOPED",
        "PRODUCT_DIMENSION",
        "",
    )


def apply_candidate_local_role(
    candidate: DimensionCandidate,
    *,
    product_name: str,
) -> DimensionCandidate:
    """Detect another lineup model from the text immediately before a value."""

    context = candidate.context_text
    raw = candidate.raw_notation
    position = context.casefold().find(raw.casefold())
    if position < 0:
        axis_parts = re.findall(
            rf"(가로|너비|폭|세로|깊이|높이|(?<![A-Za-z])[WDLH](?![A-Za-z]))"
            rf"\s*[:=]?\s*({NUMBER})",
            raw,
            re.I,
        )
        if len(axis_parts) >= 2:
            flexible_axis_pattern = (
                r"[\s:=×xX＊*()]*".join(
                    rf"{re.escape(label)}\s*[:=]?\s*{re.escape(number)}"
                    for label, number in axis_parts
                )
            )
            flexible_match = re.search(
                flexible_axis_pattern,
                context,
                re.I,
            )
            if flexible_match:
                position = flexible_match.start()
    if position < 0:
        compact_raw = re.sub(r"\s+", "", raw).casefold()
        compact_context_chars: list[str] = []
        compact_to_original: list[int] = []
        for original_index, character in enumerate(context):
            if character.isspace():
                continue
            compact_context_chars.append(character.casefold())
            compact_to_original.append(original_index)
        compact_context = "".join(compact_context_chars)
        compact_position = compact_context.find(compact_raw)
        position = (
            compact_to_original[compact_position]
            if compact_position >= 0
            else -1
        )
    if position < 0:
        return candidate
    local_before = context[max(0, position - 130) : position]
    lineup_matches = list(LINEUP_RE.finditer(local_before))
    local_model = (
        local_before[lineup_matches[-1].end() :]
        if lineup_matches
        else local_before
    )
    component_matches = list(LOCAL_COMPONENT_LABEL_RE.finditer(local_before))
    main_product_matches = list(LOCAL_MAIN_PRODUCT_LABEL_RE.finditer(local_before))
    component_is_nearest_label = bool(component_matches) and (
        not main_product_matches
        or component_matches[-1].start() > main_product_matches[-1].start()
    )
    if component_is_nearest_label:
        candidate.section_role = "COMPONENT"
        candidate.candidate_role = "COMPONENT_DIMENSION"
        candidate.rejection_reason = "후보 바로 앞 구성품명이 대표 제품과 구분됨"
    elif (
        lineup_matches
        and type_match_status(product_name, local_model) == "MISMATCH"
    ):
        candidate.section_role = "LINEUP"
        candidate.candidate_role = "LINEUP_OTHER_MODEL"
        candidate.rejection_reason = "후보 바로 앞 라인업 모델이 상품명과 불일치"
    return candidate


def score_and_decide(
    candidate: DimensionCandidate,
    *,
    product_name: str,
    profile: str,
    small_category: str,
) -> DimensionCandidate:
    score = 0
    if candidate.section_role == "PRODUCT_SIZE_SECTION":
        score += 28
    if candidate.source_axis_signature in {
        "W,D,H",
        "L,D,H",
        "W,L,H",
        "L,W,H",
        "D,H",
    }:
        score += 28
    elif candidate.source_axis_signature:
        score += 12
    if candidate.unit_status == "UNIT_PRESENT":
        score += 14
    if all(value is not None for value in (candidate.w_mm, candidate.d_mm, candidate.h_mm)):
        score += 20
    values = [
        value for value in (candidate.w_mm, candidate.d_mm, candidate.h_mm)
        if value is not None
    ]
    low_value_assessment = assess_low_dimension(
        product_name,
        "",
        small_category,
        candidate.w_mm,
        candidate.d_mm,
        candidate.h_mm,
    )
    physically_plausible = bool(values) and min(values) > 0 and max(values) <= 5000
    if physically_plausible and not low_value_assessment.requires_review:
        score += 8
    elif values:
        score -= 30

    name_score = 0
    numbers = title_numbers(product_name)
    candidate_values = {
        round(value)
        for value in (
            candidate.w_mm,
            candidate.d_mm,
            candidate.h_mm,
            candidate.value_1_raw,
            candidate.value_2_raw,
            candidate.value_3_raw,
        )
        if value is not None
    }
    if numbers and any(
        number in candidate_values or number * 10 in candidate_values for number in numbers
    ):
        name_score += 20
    codes = title_option_codes(product_name)
    selection_only_name_score = 0
    if codes and candidate.option_label in codes:
        name_score += 25
    elif codes and title_option_width_match(product_name, candidate.w_mm):
        name_score += 22
        # A standard mattress width is useful for choosing the likely row,
        # but OCR-flattened tables can pair that width with the adjacent
        # row's depth/height. Keep most of this boost out of the automatic
        # acceptance threshold unless an explicit option label also matches.
        selection_only_name_score = 15
    model_tokens = product_model_tokens(product_name)
    context_model_tokens = product_model_tokens(candidate.context_text)
    model_match = bool(model_tokens and model_tokens & context_model_tokens)
    if model_match:
        name_score += 35
    type_status = type_match_status(product_name, candidate.context_text)
    if type_status == "MATCH":
        name_score += 12
    elif type_status == "MISMATCH":
        name_score -= 35
    candidate.product_name_match_score = name_score
    score += name_score - selection_only_name_score

    if (
        candidate.candidate_role == "COMPONENT_DIMENSION"
        and model_match
        and (
            not candidate.option_label
            or candidate.option_label in title_option_codes(product_name)
        )
        and POSITIVE_SECTION_RE.search(candidate.context_text)
    ):
        candidate.section_role = "PRODUCT_SIZE_SECTION"
        candidate.candidate_role = "PRODUCT_DIMENSION"
        candidate.rejection_reason = ""

    if candidate.candidate_role in {
        "DELIVERY_CLEARANCE",
        "LINEUP_OTHER_MODEL",
        "COMPONENT_DIMENSION",
        "MEASUREMENT_TOLERANCE",
    }:
        candidate.candidate_score = score - 100
        candidate.decision_status = "REJECT"
        if not candidate.rejection_reason:
            candidate.rejection_reason = candidate.candidate_role
        return candidate

    if profile == "AREA_2D" and candidate.shape_type == "AREA_2D":
        score += 12
    candidate.candidate_score = score
    complete = all(value is not None for value in (candidate.w_mm, candidate.d_mm, candidate.h_mm))
    area_complete = (
        profile == "AREA_2D"
        and candidate.w_mm is not None
        and candidate.h_mm is not None
    )
    required_dimensions_present = complete or area_complete
    plausible = physically_plausible and not low_value_assessment.requires_review
    if required_dimensions_present and not plausible:
        if low_value_assessment.code == "LOW_DIMENSION_REVIEW_REQUIRED":
            candidate.decision_status = "HUMAN_REVIEW"
            candidate.rejection_reason = "LOW_DIMENSION_REVIEW_REQUIRED"
        else:
            candidate.decision_status = "REOCR_REQUIRED"
            candidate.rejection_reason = (
                candidate.rejection_reason or "OCR 숫자 결합·단위 이상"
            )
    elif complete and score >= 76:
        candidate.decision_status = "AUTO_ACCEPT"
    elif complete or area_complete:
        candidate.decision_status = "CATEGORY_NORMALIZED" if area_complete else "HUMAN_REVIEW"
    else:
        candidate.decision_status = "REOCR_REQUIRED"
    return candidate


def raw_notation(values: dict[str, float], unit: str) -> str:
    parts = []
    for axis in ("W", "D", "H", "L", "R", "RADIUS"):
        if axis in values:
            number = values[axis]
            value_text = str(int(number)) if float(number).is_integer() else str(number)
            parts.append(f"{axis}={value_text}{unit}")
    return " ".join(parts)


def extract_labeled_candidates(
    text: str,
    *,
    product_name: str,
    small_category: str,
) -> list[DimensionCandidate]:
    profile = category_profile(product_name, small_category)
    matches = list(AXIS_VALUE_RE.finditer(text))
    result: list[DimensionCandidate] = []
    index = 0
    while index < len(matches):
        first = matches[index]
        cluster = []
        seen_axes: set[str] = set()
        for match in matches[index:]:
            if match.end() - first.start() > 220:
                break
            axis = axis_name(match.group("label"), profile)
            if axis in seen_axes:
                break
            cluster.append(match)
            seen_axes.add(axis)
            if {"W", "D", "H"}.issubset(seen_axes):
                break
            if {"L", "D", "H"}.issubset(seen_axes):
                break
            if {"W", "L", "H"}.issubset(seen_axes):
                break
            if (
                ("R" in seen_axes or "RADIUS" in seen_axes)
                and "H" in seen_axes
            ):
                break
        index += max(1, len(cluster))
        values: dict[str, float] = {}
        has_dimension_range = False
        ordered_axes: list[str] = []
        units: list[str] = []
        for match in cluster:
            axis = axis_name(match.group("label"), profile)
            if axis in values:
                continue
            ordered_axes.append(axis)
            # A range is a real product configuration, not an OCR failure.
            # Keep the first endpoint in the single-value compatibility field
            # and preserve the complete range in raw_notation/context_text.
            # The min/max pair is resolved in the option/range output layer.
            range_end = match.group("range_end")
            values[axis] = parse_number(match.group("number"))
            has_dimension_range = has_dimension_range or bool(range_end)
            unit = canonical_unit(
                match.group("paren_unit") or match.group("unit") or ""
            )
            if unit:
                units.append(unit)
        if len(values) < 2:
            continue
        unit = units[-1] if units else ""
        if not unit:
            nearby_unit = re.search(
                rf"\(?\s*(?P<unit>{UNIT})\s*\)?",
                text[cluster[-1].end() : min(len(text), cluster[-1].end() + 24)],
                re.I,
            )
            if nearby_unit:
                unit = canonical_unit(nearby_unit.group("unit"))
        common_unit = unit or inferred_unit(values.values())
        start, end = first.start(), cluster[-1].end()
        context = nearest_heading_context(text, start, end)
        nearest_signal, has_positive = nearest_section_signal(text, start, end)
        section, role, rejection = context_role(
            product_name,
            context,
            has_positive_heading=has_positive,
            nearest_signal=nearest_signal,
        )
        axes = set(values)
        mapping = "PARTIAL_AXES"
        shape = "RECTANGLE"
        normalized: dict[str, float] = {}
        if profile == "AREA_2D" and {"W", "H"}.issubset(axes):
            mapping = "W,H->W,H;D=N/A"
            shape = "AREA_2D"
            normalized = {"W": values["W"], "H": values["H"]}
        elif {"W", "D", "H"}.issubset(axes):
            mapping = "W,D,H->W,D,H"
            normalized = {axis: values[axis] for axis in ("W", "D", "H")}
        elif {"L", "D", "H"}.issubset(axes) and profile in {"FURNITURE_3D", "GENERIC_OBJECT"}:
            mapping = "L,D,H->W,D,H"
            normalized = {"W": values["L"], "D": values["D"], "H": values["H"]}
        elif {"W", "L", "H"}.issubset(axes) and profile in {
            "FURNITURE_3D",
            "GENERIC_OBJECT",
            "OVAL_OBJECT",
        }:
            if ordered_axes[:3] == ["L", "W", "H"]:
                mapping = "L,W,H->W,D,H"
                normalized = {"W": values["L"], "D": values["W"], "H": values["H"]}
            else:
                mapping = "W,L,H->W,D,H"
                normalized = {"W": values["W"], "D": values["L"], "H": values["H"]}
        elif ({"R", "H"}.issubset(axes) or {"RADIUS", "H"}.issubset(axes)):
            diameter = values.get("R")
            if diameter is None and values.get("RADIUS") is not None:
                diameter = values["RADIUS"] * 2
            mapping = "DIAMETER,H->W,D,H"
            shape = "ROUND"
            normalized = {"W": diameter, "D": diameter, "H": values["H"]}  # type: ignore[dict-item]
        elif profile == "ROUND_OBJECT" and {"D", "H"}.issubset(axes):
            mapping = "D,H(DIAMETER)->W,D,H"
            shape = "ROUND"
            normalized = {"W": values["D"], "D": values["D"], "H": values["H"]}
        else:
            normalized = {axis: values[axis] for axis in ("W", "D", "H") if axis in values}
        mm = {axis: to_mm(value, common_unit) for axis, value in normalized.items()}
        candidate = DimensionCandidate(
            rule_id="CTX_AXIS_CLUSTER_V1",
            raw_notation=" ".join(compact_text(match.group(0)) for match in cluster),
            context_text=context,
            section_role=section,
            candidate_role=role,
            source_axis_signature=",".join(ordered_axes),
            normalized_axis_mapping=mapping,
            option_label=option_label(text, start),
            shape_type=shape,
            unit_status="UNIT_PRESENT" if unit else "UNIT_INFERRED",
            unit_text=common_unit,
            w_raw=values.get("W"),
            d_raw=values.get("D"),
            h_raw=values.get("H"),
            l_raw=values.get("L"),
            r_raw=values.get("R") or values.get("RADIUS"),
            w_mm=mm.get("W"),
            d_mm=mm.get("D"),
            h_mm=mm.get("H"),
            diameter_mm=mm.get("W") if shape == "ROUND" else None,
            rejection_reason=rejection,
        )
        candidate = apply_candidate_local_role(
            candidate,
            product_name=product_name,
        )
        candidate = score_and_decide(
            candidate,
            product_name=product_name,
            profile=profile,
            small_category=small_category,
        )
        if has_dimension_range and candidate.decision_status != "REJECT":
            candidate.decision_status = "HUMAN_REVIEW"
            candidate.rejection_reason = "DIMENSION_RANGE_REVIEW"
        result.append(candidate)
    return result


def extract_ordered_candidates(
    text: str,
    *,
    product_name: str,
    small_category: str,
) -> list[DimensionCandidate]:
    profile = category_profile(product_name, small_category)
    result: list[DimensionCandidate] = []
    triple_spans = [match.span() for match in TRIPLE_RE.finditer(text)]
    for pattern, count in ((TRIPLE_RE, 3), (PAIR_RE, 2)):
        for match in pattern.finditer(text):
            if count == 2 and any(
                match.start() < triple_end and match.end() > triple_start
                for triple_start, triple_end in triple_spans
            ):
                continue
            value_names = ("a", "b", "c")[:count]
            end_names = ("ae", "be", "ce")[:count]
            has_dimension_range = any(match.group(name) for name in end_names)
            raw_values = [
                parse_number(match.group(value_name))
                for value_name in value_names
            ]
            raw_units = [
                canonical_unit(match.group(name) or "")
                for name in ("ua", "ub", "uc")[:count]
            ]
            explicit_unit = next((value for value in reversed(raw_units) if value), "")
            if not explicit_unit:
                nearby_unit = re.match(
                    rf"\s*\(\s*(?P<unit>{UNIT})\s*\)",
                    text[match.end() : min(len(text), match.end() + 20)],
                    re.I,
                )
                if nearby_unit:
                    explicit_unit = canonical_unit(nearby_unit.group("unit"))
            common_unit = explicit_unit or inferred_unit(raw_values)
            start, end = match.start(), match.end()
            context = nearest_heading_context(text, start, end)
            nearest_signal, has_positive = nearest_section_signal(text, start, end)
            section, role, rejection = context_role(
                product_name,
                context,
                has_positive_heading=has_positive,
                nearest_signal=nearest_signal,
            )
            label = option_label(text, start)
            mm_values = [to_mm(value, unit or common_unit) for value, unit in zip(raw_values, raw_units)]
            shape = "RECTANGLE"
            mapping = "ORDER_INFERENCE_REQUIRED"
            w_mm = d_mm = h_mm = None
            if count == 3:
                w_mm, d_mm, h_mm = mm_values
                mapping = "ORDERED_TRIPLE->W,D,H"
            elif profile == "AREA_2D":
                w_mm, h_mm = mm_values
                shape = "AREA_2D"
                mapping = "2D_PAIR->W,H;D=N/A"
            elif profile == "ROUND_OBJECT":
                diameter, h_mm = mm_values
                w_mm = d_mm = diameter
                shape = "ROUND"
                mapping = "ROUND_PAIR->W,D,H"
            candidate = DimensionCandidate(
                rule_id="CTX_ORDERED_VALUES_V1",
                raw_notation=compact_text(match.group(0)),
                context_text=context,
                section_role=section,
                candidate_role=role,
                source_axis_signature="",
                normalized_axis_mapping=mapping,
                option_label=label,
                shape_type=shape,
                unit_status="UNIT_PRESENT" if explicit_unit else "UNIT_INFERRED",
                unit_text=common_unit,
                value_1_raw=raw_values[0],
                value_2_raw=raw_values[1],
                value_3_raw=raw_values[2] if count == 3 else None,
                w_mm=w_mm,
                d_mm=d_mm,
                h_mm=h_mm,
                diameter_mm=w_mm if shape == "ROUND" else None,
                rejection_reason=rejection,
            )
            candidate = apply_candidate_local_role(
                candidate,
                product_name=product_name,
            )
            candidate = score_and_decide(
                candidate,
                product_name=product_name,
                profile=profile,
                small_category=small_category,
            )
            if has_dimension_range and candidate.decision_status != "REJECT":
                candidate.decision_status = "HUMAN_REVIEW"
                candidate.rejection_reason = "DIMENSION_RANGE_REVIEW"
            result.append(candidate)
    return result


def deduplicate_candidates(candidates: list[DimensionCandidate]) -> list[DimensionCandidate]:
    best: dict[tuple[Any, ...], DimensionCandidate] = {}
    for candidate in candidates:
        key = (
            candidate.w_mm,
            candidate.d_mm,
            candidate.h_mm,
            candidate.option_label,
            candidate.candidate_role,
            candidate.normalized_axis_mapping,
        )
        old = best.get(key)
        if old is None or candidate.candidate_score > old.candidate_score:
            best[key] = candidate
    return sorted(
        best.values(),
        key=lambda item: (-item.candidate_score, item.option_label, item.raw_notation),
    )


def extract_candidates(
    text: str,
    *,
    product_name: str,
    small_category: str,
) -> list[dict[str, Any]]:
    normalized = normalized_ocr_text(text)
    candidates = extract_labeled_candidates(
        normalized,
        product_name=product_name,
        small_category=small_category,
    )
    candidates.extend(
        extract_ordered_candidates(
            normalized,
            product_name=product_name,
            small_category=small_category,
        )
    )
    return [candidate.as_dict() for candidate in deduplicate_candidates(candidates)]
