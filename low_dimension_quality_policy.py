from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable


LOW_DIMENSION_THRESHOLD_MM = 20.0


@dataclass(frozen=True)
class LowDimensionAssessment:
    requires_review: bool
    code: str
    reason: str
    low_axes: tuple[str, ...]


def _numeric_dimensions(
    w_mm: float | None,
    d_mm: float | None,
    h_mm: float | None,
) -> dict[str, float]:
    return {
        axis: float(value)
        for axis, value in (("W", w_mm), ("D", d_mm), ("H", h_mm))
        if value is not None
    }


def low_axes(
    w_mm: float | None,
    d_mm: float | None,
    h_mm: float | None,
) -> tuple[str, ...]:
    values = _numeric_dimensions(w_mm, d_mm, h_mm)
    return tuple(
        axis
        for axis in ("W", "D", "H")
        if axis in values and values[axis] <= LOW_DIMENSION_THRESHOLD_MM
    )


def _contains_any(text: str, terms: Iterable[str]) -> bool:
    folded = text.casefold()
    return any(term.casefold() in folded for term in terms)


def is_thin_product_exception(
    product_name: str,
    mid_category: str,
    small_category: str,
    w_mm: float | None,
    d_mm: float | None,
    h_mm: float | None,
) -> bool:
    """Return whether a <=20 mm value is structurally plausible.

    This is only an exception to automatic rejection. The source still needs
    an explicit product-size field or a product-size OCR/HTML section.
    """

    values = _numeric_dimensions(w_mm, d_mm, h_mm)
    axes = low_axes(w_mm, d_mm, h_mm)
    if not axes or any(value <= 0 for value in values.values()):
        return False
    if len(axes) != 1:
        return False

    text = f"{product_name} {mid_category} {small_category}"
    other_values = [
        value for axis, value in values.items() if axis not in set(axes)
    ]
    if len(other_values) < 2 or min(other_values) < 100:
        return False

    axis = axes[0]
    if _contains_any(text, ("액자", "frame", "picture frame", "메모보드", "notice board")):
        return axis == "D"
    if _contains_any(text, ("러그", "매트", "rug", "mat")):
        return axis in {"D", "H"}
    if _contains_any(text, ("거울", "mirror")):
        return axis == "D"
    if _contains_any(text, ("옷걸이", "넥타이걸이", "hanger", "tie rack")):
        return axis in {"W", "D"}
    if _contains_any(text, ("벽 선반", "wall shelf")):
        return axis == "D"
    return False


def assess_low_dimension(
    product_name: str,
    mid_category: str,
    small_category: str,
    w_mm: float | None,
    d_mm: float | None,
    h_mm: float | None,
) -> LowDimensionAssessment:
    values = _numeric_dimensions(w_mm, d_mm, h_mm)
    axes = low_axes(w_mm, d_mm, h_mm)
    if not axes:
        return LowDimensionAssessment(False, "NOT_LOW_VALUE", "", ())
    if any(value <= 0 for value in values.values()):
        return LowDimensionAssessment(
            True,
            "NONPOSITIVE_DIMENSION",
            "0 이하 규격값이 있어 완료 처리할 수 없음",
            axes,
        )
    if is_thin_product_exception(
        product_name,
        mid_category,
        small_category,
        w_mm,
        d_mm,
        h_mm,
    ):
        return LowDimensionAssessment(
            False,
            "THIN_PRODUCT_EXCEPTION_REQUIRES_EVIDENCE",
            "20mm 이하 값이지만 박형 제품군으로 구조상 가능하며 제품 규격 근거 확인 필요",
            axes,
        )
    return LowDimensionAssessment(
        True,
        "LOW_DIMENSION_REVIEW_REQUIRED",
        "20mm 이하 규격값은 OCR 숫자 유실·단위 오류·배송 규격 혼입 가능성이 있어 재검증 필요",
        axes,
    )

