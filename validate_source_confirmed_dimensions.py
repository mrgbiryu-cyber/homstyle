from __future__ import annotations

import json
import hashlib
import re
import sqlite3
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

import xlsxwriter

import build_dimension_locked_evidence_workbook as evidence_workbook
import build_homestyle_bulk_workbook as bulk_workbook
from bulk_homestyle_collect import unpack
from dimension_context_normalizer import (
    COMPONENT_RE,
    DELIVERY_RE,
    LINEUP_RE,
    POSITIVE_SECTION_RE,
    TOLERANCE_RE,
    extract_candidates,
    product_model_tokens,
    product_type_tokens,
    title_numbers,
    title_option_codes,
    type_match_status,
)
from low_dimension_quality_policy import assess_low_dimension


ROOT = Path(__file__).resolve().parent
DB_PATH = ROOT / "homestyle_bulk_run" / "homestyle_bulk.sqlite"
OUTPUT = (
    ROOT
    / "홈스타일_기존확정값_4152개_RAW증빙_상품명문맥재검증_비정상값16건보정_20mm이하5건추가보정.xlsx"
)
SUMMARY_OUTPUT = (
    ROOT / "homestyle_bulk_run" / "source_confirmed_4152_validation_summary.json"
)
EXPECTED_PRODUCTS = 4_152
QUALITY_REPAIR_RUN = "locked_dimension_quality_repair_20260728"
LOW_VALUE_REPAIR_RUN = "low_dimension_quality_repair_20260728"
QUALITY_REPAIR_RUNS = {QUALITY_REPAIR_RUN, LOW_VALUE_REPAIR_RUN}

PASS = "자동검증_PASS"
CHECK_RISK = "수기검증_위험문맥"
CHECK_ALTERNATIVE = "수기검증_다른후보우선"
CHECK_MULTI = "수기검증_복수후보"
CHECK_WEAK = "수기검증_상품명근거약함"
RESEARCH = "원천재탐색"

STATUS_ORDER = {
    PASS: 0,
    CHECK_RISK: 1,
    CHECK_ALTERNATIVE: 2,
    CHECK_MULTI: 3,
    CHECK_WEAK: 4,
    RESEARCH: 5,
}

PROOF_ORDER = {
    "구조화API규격": 0,
    "수기원본검증": 1,
    "상품명직접일치": 2,
    "제품규격문맥": 3,
    "원천숫자일치": 4,
    "원천미연결": 5,
}


@dataclass(frozen=True)
class Segment:
    source_type: str
    source_detail: str
    source_ref: str
    source_url: str
    text: str


@dataclass
class EvidenceAssessment:
    source_kind: str
    source_detail: str
    source_url: str
    raw_locator: str
    proof_grade: str
    safe: bool
    structured_api_size: bool
    positive_size_context: bool
    title_number_match: bool
    option_code_match: bool
    model_token_match: bool
    product_type_status: str
    product_name_match_score: int
    candidate_score: int
    candidate_role: str
    section_role: str
    warning: str
    judged_notation: str
    raw_text: str
    ocr_region: str
    ocr_bbox_json: str
    payload_sha256: str


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def clean(value: Any, limit: int = 32_000) -> str:
    return evidence_workbook.clean_cell(value, limit)


def same_number(
    left: float | None,
    right: float | None,
    tolerance: float = 0.11,
) -> bool:
    if left is None or right is None:
        return left is None and right is None
    return abs(float(left) - float(right)) <= tolerance


def candidate_matches_locked(candidate: dict[str, Any], row: sqlite3.Row) -> bool:
    return (
        same_number(candidate.get("w_mm"), row["locked_w_mm"])
        and same_number(candidate.get("d_mm"), row["locked_d_mm"])
        and same_number(candidate.get("h_mm"), row["locked_h_mm"])
    )


def dimension_key(candidate: dict[str, Any]) -> tuple[float | None, ...]:
    return tuple(
        None if candidate.get(axis) is None else round(float(candidate[axis]), 3)
        for axis in ("w_mm", "d_mm", "h_mm")
    )


def format_dimension(
    w_mm: float | None,
    d_mm: float | None,
    h_mm: float | None,
) -> str:
    return (
        f"W={evidence_workbook.fmt_number(w_mm)} / "
        f"D={evidence_workbook.fmt_number(d_mm)} / "
        f"H={evidence_workbook.fmt_number(h_mm)} mm"
    )


def matching_title_number(product_name: str, row: sqlite3.Row) -> bool:
    numbers = title_numbers(product_name)
    values = {
        round(float(value))
        for value in (
            row["locked_w_mm"],
            row["locked_d_mm"],
            row["locked_h_mm"],
        )
        if value is not None
    }
    return bool(
        numbers
        and any(number in values or number * 10 in values for number in numbers)
    )


def local_evidence_text(item: evidence_workbook.Evidence) -> str:
    if item.judged_notation:
        match = re.search(
            r"\[(?:OCR\s*)?판단 구간\]\s*(.*?)\s*\n\n\[정규화 결과\]",
            item.judged_notation,
            re.S,
        )
        if match:
            return match.group(1).strip()
    return "\n".join(
        value for value in (item.raw_text, item.judged_notation) if value
    )


def warning_labels(
    product_name: str,
    text: str,
    candidate_roles: Iterable[str],
    *,
    structured_size_field: bool,
) -> list[str]:
    roles = set(candidate_roles)
    role_labels = {
        "DELIVERY_CLEARANCE": "배송·반입 규격 후보",
        "COMPONENT_DIMENSION": "구성품 규격 후보",
        "LINEUP_OTHER_MODEL": "다른 모델·라인업 규격 후보",
        "MEASUREMENT_TOLERANCE": "허용오차 후보",
    }
    labels = []
    for role in sorted(roles):
        if role not in role_labels:
            continue
        if structured_size_field and role in {
            "COMPONENT_DIMENSION",
            "MEASUREMENT_TOLERANCE",
        }:
            continue
        labels.append(role_labels[role])

    # If the context parser could not bind the exact values to a role, retain a
    # conservative keyword fallback. Once an exact PRODUCT_DIMENSION or hard
    # negative role exists, the nearest-section role owns the decision and text
    # from the following delivery/tolerance paragraph must not contaminate it.
    if not roles:
        compact = re.sub(r"\s+", " ", text)
        if DELIVERY_RE.search(compact):
            labels.append("배송·포장 문맥 포함")
        if TOLERANCE_RE.search(compact) and not structured_size_field:
            labels.append("오차 문맥 포함")
        if (
            COMPONENT_RE.search(compact)
            and not COMPONENT_RE.search(product_name)
            and not structured_size_field
        ):
            labels.append("구성품 문맥 포함")
        if (
            LINEUP_RE.search(compact)
            and type_match_status(product_name, compact) == "MISMATCH"
        ):
            labels.append("라인업 제품유형 불일치")
    return list(dict.fromkeys(labels))


def structured_api_size(item: evidence_workbook.Evidence) -> bool:
    return bool(
        item.source_kind == "API"
        and item.source_detail == "상품 API JSON"
        and "productNotification" in item.raw_locator
        and re.search(
            r"""["']title["']\s*:\s*["'](?:크기|치수|규격|사이즈)["']""",
            item.raw_text,
            re.I,
        )
    )


def best_exact_candidate(
    candidates: list[dict[str, Any]],
) -> dict[str, Any] | None:
    if not candidates:
        return None
    return max(
        candidates,
        key=lambda candidate: (
            candidate.get("decision_status") != "REJECT",
            candidate.get("candidate_role") == "PRODUCT_DIMENSION",
            int(candidate.get("product_name_match_score") or 0),
            int(candidate.get("candidate_score") or 0),
        ),
    )


def assess_evidence(
    row: sqlite3.Row,
    item: evidence_workbook.Evidence,
) -> EvidenceAssessment:
    product_name = str(row["product_name"] or "")
    text = local_evidence_text(item)
    candidates = extract_candidates(
        text,
        product_name=product_name,
        small_category=str(row["small_category"] or ""),
    )
    exact = [candidate for candidate in candidates if candidate_matches_locked(candidate, row)]
    best = best_exact_candidate(exact)
    roles = [str(candidate.get("candidate_role") or "") for candidate in exact]
    structured = structured_api_size(item)
    warnings = warning_labels(
        product_name,
        text,
        roles,
        structured_size_field=structured,
    )
    positive = bool(
        POSITIVE_SECTION_RE.search(text)
        or any(
            candidate.get("section_role") == "PRODUCT_SIZE_SECTION"
            and candidate.get("candidate_role") == "PRODUCT_DIMENSION"
            for candidate in exact
        )
    )
    title_number = matching_title_number(product_name, row)
    model_match = bool(
        product_model_tokens(product_name) & product_model_tokens(text)
    )
    type_status = type_match_status(product_name, text)
    option_codes = title_option_codes(product_name)
    option_match = any(
        str(candidate.get("option_label") or "").upper() in option_codes
        for candidate in exact
        if candidate.get("option_label")
    )
    name_score = max(
        (int(candidate.get("product_name_match_score") or 0) for candidate in exact),
        default=0,
    )
    candidate_score = max(
        (int(candidate.get("candidate_score") or 0) for candidate in exact),
        default=0,
    )
    hard_roles = {
        "DELIVERY_CLEARANCE",
        "LINEUP_OTHER_MODEL",
    }
    if not structured:
        hard_roles.update(
            {
                "COMPONENT_DIMENSION",
                "MEASUREMENT_TOLERANCE",
            }
        )
    safe = not warnings and not any(role in hard_roles for role in roles)
    direct_name = bool(
        name_score >= 20 or model_match or option_match or title_number
    )
    manually_verified_repair = bool(
        row["resolution_status"] == "MANUAL_CONFIRMED"
        and row["source_snapshot_id"] in QUALITY_REPAIR_RUNS
        and item.source_detail.startswith("검증된 추가 이미지 OCR 규격")
    )
    if structured and safe:
        proof_grade = "구조화API규격"
    elif manually_verified_repair and safe:
        proof_grade = "수기원본검증"
    elif direct_name and safe:
        proof_grade = "상품명직접일치"
    elif positive and safe:
        proof_grade = "제품규격문맥"
    else:
        proof_grade = "원천숫자일치"

    return EvidenceAssessment(
        source_kind=item.source_kind,
        source_detail=item.source_detail,
        source_url=item.source_url,
        raw_locator=item.raw_locator,
        proof_grade=proof_grade,
        safe=safe,
        structured_api_size=structured,
        positive_size_context=positive,
        title_number_match=title_number,
        option_code_match=option_match,
        model_token_match=model_match,
        product_type_status=type_status,
        product_name_match_score=name_score,
        candidate_score=candidate_score,
        candidate_role=str((best or {}).get("candidate_role") or ""),
        section_role=str((best or {}).get("section_role") or ""),
        warning="|".join(warnings),
        judged_notation=item.judged_notation,
        raw_text=item.raw_text,
        ocr_region=item.ocr_region,
        ocr_bbox_json=item.ocr_bbox_json,
        payload_sha256=item.payload_sha256,
    )


def raw_segments(row: sqlite3.Row) -> list[Segment]:
    product_id = row["product_id"]
    urls = evidence_workbook.source_urls(product_id)
    result: list[Segment] = []

    def add(
        source_type: str,
        source_detail: str,
        source_ref: str,
        source_url: str,
        text: Any,
    ) -> None:
        cleaned = bulk_workbook.clean_text(text, 120_000)
        if not cleaned:
            return
        chunks = [cleaned]
        if len(cleaned) > 10_000:
            chunks = bulk_workbook.dimension_keyword_snippets(cleaned, limit=30)
            span = evidence_workbook.exact_dimension_span(
                cleaned,
                row["locked_w_mm"],
                row["locked_d_mm"],
                row["locked_h_mm"],
            )
            if span:
                start = max(0, span[0] - 400)
                end = min(len(cleaned), span[1] + 500)
                chunks.insert(0, cleaned[start:end])
            if not chunks:
                chunks = [cleaned[:10_000]]
        for chunk_no, chunk in enumerate(chunks, start=1):
            result.append(
                Segment(
                    source_type=source_type,
                    source_detail=source_detail,
                    source_ref=(
                        source_ref
                        if len(chunks) == 1
                        else f"{source_ref}:chunk={chunk_no}"
                    ),
                    source_url=source_url,
                    text=chunk,
                )
            )

    goods_payload = unpack(row["goods_blob"]) or {}
    goods_data = goods_payload.get("data") or {}
    for path, value, parent in evidence_workbook.walk_strings(
        goods_data,
        "$.data",
        excluded_keys=frozenset({"detailInfo"}),
    ):
        if evidence_workbook.has_dimension_hint(value):
            add(
                "API_STRUCTURED",
                "상품 API JSON",
                path,
                urls["goods"],
                evidence_workbook.raw_parent_fragment(path, value, parent)
                if "productNotification" in path
                else value,
            )

    detail_html = str(goods_data.get("detailInfo") or "")
    for index, snippet in enumerate(
        bulk_workbook.dimension_keyword_snippets(detail_html),
        start=1,
    ):
        add(
            "HTML_DETAIL",
            "상품 상세 HTML(detailInfo)",
            f"detailInfo[{index}]",
            urls["pdp"],
            snippet,
        )

    html_payload = unpack(row["html_blob"]) or {}
    for index, signal in enumerate(html_payload.get("dimension_signals") or [], start=1):
        add(
            "HTML_PDP",
            "PDP HTML 규격 문맥",
            f"dimension_signals[{index}]",
            urls["pdp"],
            signal,
        )

    qna_payload = unpack(row["qna_blob"]) or {}
    for path, value, parent in evidence_workbook.walk_strings(qna_payload):
        if evidence_workbook.has_dimension_hint(value):
            add(
                "API_QNA",
                "FAQ/Q&A API JSON",
                path,
                urls["qna"],
                value,
            )

    ocr_payload = unpack(row["ocr_blob"]) or {}
    selected = ocr_payload.get("selected") or {}
    image_url = str(selected.get("url") or urls["pdp"])
    for index, (label, value) in enumerate(
        bulk_workbook.verified_dimension_texts(ocr_payload),
        start=1,
    ):
        add(
            "SOURCE_SELECTED_OCR",
            label,
            f"dimension_reinforcements[{index}]",
            image_url,
            value,
        )
    dimension_text = str(ocr_payload.get("dimension_text") or "")
    if dimension_text:
        add(
            "SOURCE_SELECTED_OCR",
            "상세 이미지 OCR",
            "sources.ocr_blob.dimension_text",
            image_url,
            dimension_text,
        )
    full_ocr_text = str(
        (ocr_payload.get("ocr") or {}).get("text")
        or ocr_payload.get("combined_text")
        or ""
    )
    if full_ocr_text and full_ocr_text != dimension_text:
        add(
            "SOURCE_SELECTED_OCR",
            "상세 이미지 전체 OCR",
            "sources.ocr_blob.full_text",
            image_url,
            full_ocr_text,
        )

    unique: list[Segment] = []
    seen: set[tuple[str, str, str]] = set()
    for segment in result:
        key = (segment.source_type, segment.source_url, segment.text)
        if key in seen:
            continue
        seen.add(key)
        unique.append(segment)
    return unique


def evidence_from_segments(
    row: sqlite3.Row,
    segments: list[Segment],
) -> list[evidence_workbook.Evidence]:
    result: list[evidence_workbook.Evidence] = []
    for segment in segments:
        for record in bulk_workbook.dimension_records(
            [(segment.source_detail, segment.text)]
        ):
            if not evidence_workbook.record_matches(
                record,
                row["locked_w_mm"],
                row["locked_d_mm"],
                row["locked_h_mm"],
            ):
                continue
            excerpt, location = evidence_workbook.tight_ocr_excerpt(
                segment.text,
                row["locked_w_mm"],
                row["locked_d_mm"],
                row["locked_h_mm"],
            )
            source_kind = (
                "OCR"
                if "OCR" in segment.source_type
                else "HTML"
                if segment.source_type.startswith("HTML")
                else "API"
            )
            judgment = (
                f"[판단 구간]\n{excerpt}\n\n"
                f"[정규화 결과]\n"
                f"{format_dimension(row['locked_w_mm'], row['locked_d_mm'], row['locked_h_mm'])}"
            )
            digest = hashlib.sha256(segment.text.encode("utf-8")).hexdigest()
            result.append(
                evidence_workbook.Evidence(
                    product_id=row["product_id"],
                    source_kind=source_kind,
                    source_detail=segment.source_detail,
                    source_url=segment.source_url,
                    raw_locator=segment.source_ref,
                    raw_text=clean(segment.text),
                    parsed_w_mm=record.get("w_mm"),
                    parsed_d_mm=record.get("d_mm"),
                    parsed_h_mm=record.get("h_mm"),
                    judged_notation=clean(judgment),
                    source_snapshot_id=row["source_snapshot_id"] or "",
                    collected_at="",
                    payload_sha256=digest,
                    payload_character_count=len(segment.text),
                    ocr_full_context=clean(segment.text)
                    if source_kind == "OCR"
                    else "",
                    ocr_region=location if source_kind == "OCR" else "",
                )
            )
    return list(
        {
            evidence_workbook.evidence_key(item): item
            for item in result
        }.values()
    )


def all_context_candidates(
    row: sqlite3.Row,
    segments: list[Segment],
) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for segment in segments:
        for candidate in extract_candidates(
            segment.text,
            product_name=str(row["product_name"] or ""),
            small_category=str(row["small_category"] or ""),
        ):
            candidate.update(
                {
                    "source_type": segment.source_type,
                    "source_ref": segment.source_ref,
                    "source_detail": segment.source_detail,
                    "image_url": segment.source_url,
                }
            )
            result.append(candidate)
    return result


def preferred_candidates(
    candidates: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    eligible = [
        candidate
        for candidate in candidates
        if candidate.get("decision_status") in {"AUTO_ACCEPT", "CATEGORY_NORMALIZED"}
        and candidate.get("candidate_role") == "PRODUCT_DIMENSION"
        and (
            candidate.get("shape_type") == "AREA_2D"
            and candidate.get("w_mm") is not None
            and candidate.get("h_mm") is not None
            or all(candidate.get(axis) is not None for axis in ("w_mm", "d_mm", "h_mm"))
        )
    ]
    support: dict[tuple[float | None, ...], set[tuple[str, str]]] = {}
    for candidate in eligible:
        support.setdefault(dimension_key(candidate), set()).add(
            (
                str(candidate.get("source_type") or ""),
                str(candidate.get("source_ref") or ""),
            )
        )
    best_by_dimension: dict[tuple[float | None, ...], dict[str, Any]] = {}
    for candidate in eligible:
        key = dimension_key(candidate)
        candidate["_support_count"] = len(support[key])
        old = best_by_dimension.get(key)
        rank = (
            int(candidate.get("product_name_match_score") or 0),
            int(candidate.get("_support_count") or 0),
            int(candidate.get("candidate_score") or 0),
            candidate.get("section_role") == "PRODUCT_SIZE_SECTION",
        )
        old_rank = (
            int(old.get("product_name_match_score") or 0),
            int(old.get("_support_count") or 0),
            int(old.get("candidate_score") or 0),
            old.get("section_role") == "PRODUCT_SIZE_SECTION",
        ) if old else None
        if old is None or rank > old_rank:
            best_by_dimension[key] = candidate
    return sorted(
        best_by_dimension.values(),
        key=lambda candidate: (
            -int(candidate.get("product_name_match_score") or 0),
            -int(candidate.get("_support_count") or 0),
            -int(candidate.get("candidate_score") or 0),
        ),
    )


def product_decision(
    row: sqlite3.Row,
    assessments: list[EvidenceAssessment],
    candidates: list[dict[str, Any]],
) -> tuple[str, str, str, list[dict[str, Any]], bool]:
    locked_values = (
        row["locked_w_mm"],
        row["locked_d_mm"],
        row["locked_h_mm"],
    )
    if any(
        value is None or float(value) <= 0
        for value in locked_values
    ):
        return (
            CHECK_ALTERNATIVE,
            "비정상수치",
            "확정 W/D/H 중 누락·0·음수 값이 있어 완료 처리할 수 없음; 원문 재파싱 필요",
            preferred_candidates(candidates),
            False,
        )
    low_value_assessment = assess_low_dimension(
        str(row["product_name"] or ""),
        str(row["mid_category"] or ""),
        str(row["small_category"] or ""),
        row["locked_w_mm"],
        row["locked_d_mm"],
        row["locked_h_mm"],
    )
    if low_value_assessment.requires_review:
        return (
            CHECK_ALTERNATIVE,
            "20mm 이하 의심값",
            (
                f"{','.join(low_value_assessment.low_axes)}축 값이 20mm 이하이므로 "
                "OCR 숫자 유실·단위 오류·배송 규격 혼입 여부를 재검증해야 함"
            ),
            preferred_candidates(candidates),
            False,
        )
    if not assessments:
        return RESEARCH, "원천미연결", "확정 W/D/H와 일치하는 RAW를 찾지 못함", [], False

    safe_pass = [
        assessment
        for assessment in assessments
        if assessment.safe
        and assessment.proof_grade
        in {
            "구조화API규격",
            "수기원본검증",
            "상품명직접일치",
            "제품규격문맥",
        }
    ]
    risky = [assessment for assessment in assessments if not assessment.safe]
    preferred = preferred_candidates(candidates)
    preferred_locked = [
        candidate for candidate in preferred if candidate_matches_locked(candidate, row)
    ]
    competing = [
        candidate for candidate in preferred if not candidate_matches_locked(candidate, row)
    ]

    best_locked_name_score = max(
        (
            int(candidate.get("product_name_match_score") or 0)
            for candidate in preferred_locked
        ),
        default=0,
    )
    stronger_competing = [
        candidate
        for candidate in competing
        if int(candidate.get("product_name_match_score") or 0)
        > best_locked_name_score
        and (
            int(candidate.get("product_name_match_score") or 0) >= 20
            or int(candidate.get("_support_count") or 0) >= 2
        )
    ]

    if stronger_competing:
        return (
            CHECK_ALTERNATIVE,
            "원천숫자일치",
            "현재 확정값보다 상품명 매칭 또는 원천 지지가 강한 다른 규격 후보 존재",
            stronger_competing,
            bool(risky),
        )

    option_labels = {
        str(candidate.get("option_label") or "")
        for candidate in preferred
        if candidate.get("option_label")
    }
    if len(option_labels) >= 2 and not preferred_locked:
        return (
            CHECK_MULTI,
            "원천숫자일치",
            "복수 옵션 규격이 있으나 현재 확정값과 일치하는 대표 옵션 미확정",
            preferred,
            bool(risky),
        )

    if safe_pass:
        primary = min(
            safe_pass,
            key=lambda assessment: (
                PROOF_ORDER[assessment.proof_grade],
                -assessment.product_name_match_score,
                -assessment.candidate_score,
            ),
        )
        detail = {
            "구조화API규격": "동일 productId의 상품정보고시 크기 필드와 확정값 일치",
            "수기원본검증": "저장된 상세 이미지의 규격 영역과 OCR 원문을 대조해 교정값 확정",
            "상품명직접일치": "상품명 숫자·옵션·모델·제품유형 중 직접 일치 근거 보유",
            "제품규격문맥": "제품 사이즈/규격 구역의 완전한 W/D/H와 확정값 일치",
        }[primary.proof_grade]
        if risky:
            detail += "; 동일 수치의 위험 문맥 증빙도 있어 경고 컬럼에 보존"
        return PASS, primary.proof_grade, detail, competing, bool(risky)

    if risky:
        return (
            CHECK_RISK,
            "원천숫자일치",
            "확정값과 같은 수치는 있으나 배송·포장·구성품·라인업 등 위험 문맥",
            competing,
            True,
        )

    return (
        CHECK_WEAK,
        "원천숫자일치",
        "확정값과 RAW 숫자는 일치하지만 상품명 직접 일치 또는 제품 규격 제목 근거가 약함",
        competing,
        False,
    )


def load_rows(connection: sqlite3.Connection) -> Iterable[sqlite3.Row]:
    return connection.execute(
        """
        SELECT
            l.*,
            s.goods_blob,
            s.qna_blob,
            s.html_blob,
            s.ocr_blob,
            s.structured_at,
            s.html_at,
            s.ocr_at
        FROM fact_dimension_resolution_ledger l
        JOIN sources s ON s.product_id = l.product_id
        WHERE l.is_locked = 1
          AND (
                l.resolution_status = 'SOURCE_CONFIRMED'
                OR (
                    l.resolution_status = 'MANUAL_CONFIRMED'
                    AND EXISTS (
                        SELECT 1
                        FROM hist_dimension_quality_repair q
                        WHERE q.product_id = l.product_id
                          AND q.old_resolution_status = 'SOURCE_CONFIRMED'
                    )
                )
          )
        ORDER BY l.product_id
        """
    )


def primary_assessment(
    assessments: list[EvidenceAssessment],
) -> EvidenceAssessment | None:
    return min(
        assessments,
        key=lambda assessment: (
            not assessment.safe,
            PROOF_ORDER.get(assessment.proof_grade, 9),
            -assessment.product_name_match_score,
            -assessment.candidate_score,
            {"API": 0, "HTML": 1, "OCR": 2}.get(assessment.source_kind, 9),
        ),
        default=None,
    )


def candidate_summary(candidates: list[dict[str, Any]], limit: int = 8) -> str:
    lines = []
    for candidate in candidates[:limit]:
        lines.append(
            " | ".join(
                [
                    format_dimension(
                        candidate.get("w_mm"),
                        candidate.get("d_mm"),
                        candidate.get("h_mm"),
                    ),
                    f"nameScore={int(candidate.get('product_name_match_score') or 0)}",
                    f"score={int(candidate.get('candidate_score') or 0)}",
                    f"support={int(candidate.get('_support_count') or 0)}",
                    f"role={candidate.get('candidate_role') or ''}",
                    f"section={candidate.get('section_role') or ''}",
                    f"option={candidate.get('option_label') or ''}",
                    f"source={candidate.get('source_type') or ''}",
                    clean(candidate.get("raw_notation") or "", 500),
                ]
            )
        )
    return clean("\n".join(lines))


def combined_evidence_status(status: str, proof: str) -> str:
    base = (
        "증빙완료"
        if status == PASS
        else "원천미연결"
        if status == RESEARCH
        else "수기검증필요"
    )
    return f"{base} | {proof}"


def combined_reason(
    status: str,
    reason: str,
    warnings: list[str],
    competing: list[dict[str, Any]],
) -> str:
    status_label = {
        PASS: "자동검증",
        CHECK_RISK: "위험문맥",
        CHECK_ALTERNATIVE: "다른후보우선",
        CHECK_MULTI: "복수후보",
        CHECK_WEAK: "상품명근거약함",
        RESEARCH: "원천미연결",
    }.get(status, status)
    parts = [f"[{status_label}] {reason}"]
    if warnings:
        parts.append("[문맥 경고] " + " | ".join(warnings))
    if competing:
        parts.append(
            f"[경쟁 후보 {len(competing)}개]\n{candidate_summary(competing)}"
        )
    return clean("\n".join(parts))


def first_evidence_url(
    items: list[evidence_workbook.Evidence],
    kind: str,
) -> str:
    return next(
        (
            item.source_url
            for item in items
            if item.source_kind == kind and item.source_url
        ),
        "",
    )


def combined_evidence_field(
    items: list[evidence_workbook.Evidence],
    kind: str,
    field: str,
) -> str:
    values = []
    for index, item in enumerate(
        [item for item in items if item.source_kind == kind],
        start=1,
    ):
        value = str(getattr(item, field) or "")
        if value:
            values.append(f"[{kind} RAW {index}]\n{value}")
    return clean("\n\n--------------------\n\n".join(values))


def write_sheet(
    workbook: xlsxwriter.Workbook,
    name: str,
    headers: list[str],
    rows: list[list[Any]],
    widths: list[int],
    *,
    status_column: int | None = None,
    url_columns: set[int] | None = None,
) -> None:
    sheet = workbook.add_worksheet(name)
    url_columns = url_columns or set()
    header = workbook.add_format(
        {
            "bold": True,
            "font_color": "white",
            "bg_color": "#1F4E78",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
        }
    )
    body = workbook.add_format(
        {"border": 1, "valign": "top", "text_wrap": True}
    )
    pass_format = workbook.add_format(
        {
            "border": 1,
            "valign": "top",
            "text_wrap": True,
            "bg_color": "#E2F0D9",
        }
    )
    check_format = workbook.add_format(
        {
            "border": 1,
            "valign": "top",
            "text_wrap": True,
            "bg_color": "#FFF2CC",
        }
    )
    research_format = workbook.add_format(
        {
            "border": 1,
            "valign": "top",
            "text_wrap": True,
            "bg_color": "#F4CCCC",
        }
    )
    url_format = workbook.add_format(
        {
            "border": 1,
            "valign": "top",
            "text_wrap": True,
            "font_color": "#0563C1",
            "underline": True,
        }
    )
    for col, header_value in enumerate(headers):
        sheet.write(0, col, header_value, header)
    for row_no, values in enumerate(rows, start=1):
        row_status = (
            str(values[status_column]) if status_column is not None else ""
        )
        row_format = (
            pass_format
            if row_status.startswith("증빙완료")
            else research_format
            if row_status.startswith("원천미연결")
            else check_format
            if row_status
            else body
        )
        for col, value in enumerate(values):
            if (
                col in url_columns
                and str(value or "").startswith(("http://", "https://"))
                and "\n" not in str(value)
            ):
                sheet.write_url(row_no, col, str(value), url_format, str(value))
            elif isinstance(value, (int, float)) and not isinstance(value, bool):
                sheet.write_number(row_no, col, value, row_format)
            else:
                sheet.write(row_no, col, clean(value), row_format)
    sheet.freeze_panes(1, 3)
    sheet.autofilter(0, 0, len(rows), len(headers) - 1)
    for col, width in enumerate(widths):
        sheet.set_column(col, col, width)
    sheet.set_row(0, 36)


def build() -> dict[str, Any]:
    assessed_at = now_text()
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    total = connection.execute(
        """
        SELECT COUNT(*)
        FROM fact_dimension_resolution_ledger
        WHERE is_locked=1
          AND (
                resolution_status='SOURCE_CONFIRMED'
                OR (
                    resolution_status='MANUAL_CONFIRMED'
                    AND EXISTS (
                        SELECT 1
                        FROM hist_dimension_quality_repair q
                        WHERE q.product_id=fact_dimension_resolution_ledger.product_id
                          AND q.old_resolution_status='SOURCE_CONFIRMED'
                    )
                )
          )
        """
    ).fetchone()[0]
    if total != EXPECTED_PRODUCTS:
        raise RuntimeError(
            f"기존 확정값 검증 코호트가 {EXPECTED_PRODUCTS:,}개가 아닙니다: {total:,}"
        )
    source_confirmed_count = connection.execute(
        """
        SELECT COUNT(*)
        FROM fact_dimension_resolution_ledger
        WHERE is_locked=1 AND resolution_status='SOURCE_CONFIRMED'
        """
    ).fetchone()[0]
    manual_source_repair_count = total - source_confirmed_count

    product_rows: list[list[Any]] = []
    raw_rows: list[list[Any]] = []
    unlinked_rows: list[list[Any]] = []
    status_counts: Counter[str] = Counter()
    proof_counts: Counter[str] = Counter()
    warning_counts: Counter[str] = Counter()
    source_counts: Counter[str] = Counter()
    raw_source_counts: Counter[str] = Counter()
    direct_signal_counts: Counter[str] = Counter()

    for index, row in enumerate(load_rows(connection), start=1):
        segments = raw_segments(row)
        exact_items = evidence_from_segments(row, segments)
        assessments = [assess_evidence(row, item) for item in exact_items]
        candidates = all_context_candidates(row, segments)
        status, proof, reason, competing, has_warning_evidence = product_decision(
            row,
            assessments,
            candidates,
        )
        primary = primary_assessment(assessments)
        warnings = sorted(
            {
                warning
                for assessment in assessments
                for warning in assessment.warning.split("|")
                if warning
            }
        )
        warning_text = "|".join(warnings)
        source_kinds = sorted({assessment.source_kind for assessment in assessments})
        exact_candidates = [
            candidate
            for candidate in candidates
            if candidate_matches_locked(candidate, row)
        ]
        max_name_score = max(
            (
                int(candidate.get("product_name_match_score") or 0)
                for candidate in exact_candidates
            ),
            default=max(
                (
                    assessment.product_name_match_score
                    for assessment in assessments
                ),
                default=0,
            ),
        )
        max_candidate_score = max(
            (
                int(candidate.get("candidate_score") or 0)
                for candidate in exact_candidates
            ),
            default=max(
                (assessment.candidate_score for assessment in assessments),
                default=0,
            ),
        )

        title_number_flag = any(
            assessment.title_number_match for assessment in assessments
        )
        option_flag = any(
            assessment.option_code_match for assessment in assessments
        )
        model_flag = any(
            assessment.model_token_match for assessment in assessments
        )
        type_match_flag = any(
            assessment.product_type_status == "MATCH"
            for assessment in assessments
        )
        positive_flag = any(
            assessment.positive_size_context for assessment in assessments
        )
        structured_flag = any(
            assessment.structured_api_size for assessment in assessments
        )
        for key, flag in (
            ("상품명 숫자", title_number_flag),
            ("옵션 코드", option_flag),
            ("모델 코드", model_flag),
            ("제품 유형", type_match_flag),
            ("제품 규격 문맥", positive_flag),
            ("구조화 API 크기필드", structured_flag),
        ):
            if flag:
                direct_signal_counts[key] += 1

        status_counts[status] += 1
        proof_counts[proof] += 1
        for warning in warnings:
            warning_counts[warning] += 1
        for source_kind in source_kinds:
            source_counts[source_kind] += 1

        evidence_status = combined_evidence_status(status, proof)
        reason_value = combined_reason(status, reason, warnings, competing)
        pdp_url = evidence_workbook.source_urls(row["product_id"])["pdp"]
        product_values = [
            row["product_id"],
            row["product_name"],
            row["mid_category"],
            row["small_category"],
            row["locked_w_mm"],
            row["locked_d_mm"],
            row["locked_h_mm"],
            evidence_status,
            reason_value,
            ",".join(source_kinds),
            first_evidence_url(exact_items, "API"),
            combined_evidence_field(exact_items, "API", "raw_text"),
            first_evidence_url(exact_items, "HTML"),
            combined_evidence_field(exact_items, "HTML", "raw_text"),
            first_evidence_url(exact_items, "OCR"),
            combined_evidence_field(exact_items, "OCR", "judged_notation"),
            combined_evidence_field(exact_items, "OCR", "ocr_region"),
            len(exact_items),
            (
                f"{primary.source_kind} / {primary.raw_locator}"
                if primary
                else ""
            ),
            (
                primary.judged_notation or primary.raw_text
                if primary
                else ""
            ),
            row["resolution_status"],
            row["resolution_rule_code"],
            row["resolution_source"],
            row["source_snapshot_id"],
            pdp_url,
        ]
        product_rows.append(product_values)
        if status == RESEARCH:
            unlinked_rows.append(
                [
                    row["product_id"],
                    row["product_name"],
                    row["mid_category"],
                    row["small_category"],
                    row["locked_w_mm"],
                    row["locked_d_mm"],
                    row["locked_h_mm"],
                    row["resolution_status"],
                    row["resolution_rule_code"],
                    row["resolution_source"],
                    reason_value,
                    pdp_url,
                ]
            )

        for raw_no, assessment in enumerate(assessments, start=1):
            raw_source_counts[assessment.source_kind] += 1
            collected_at = {
                "API": row["structured_at"],
                "HTML": row["html_at"],
                "OCR": row["ocr_at"],
            }.get(assessment.source_kind, "")
            raw_rows.append(
                [
                    row["product_id"],
                    raw_no,
                    row["product_name"],
                    row["locked_w_mm"],
                    row["locked_d_mm"],
                    row["locked_h_mm"],
                    evidence_status,
                    reason_value,
                    assessment.source_kind,
                    assessment.source_detail,
                    assessment.source_url,
                    assessment.raw_locator,
                    assessment.raw_text
                    if assessment.source_kind == "API"
                    else "",
                    assessment.raw_text
                    if assessment.source_kind == "HTML"
                    else "",
                    assessment.judged_notation
                    if assessment.source_kind == "OCR"
                    else "",
                    assessment.ocr_region,
                    assessment.raw_text
                    if assessment.source_kind == "OCR"
                    else "",
                    assessment.ocr_bbox_json,
                    row["locked_w_mm"],
                    row["locked_d_mm"],
                    row["locked_h_mm"],
                    "LOCKED_VALUE_EXACT",
                    collected_at or "",
                    row["source_snapshot_id"],
                    assessment.payload_sha256,
                    len(assessment.raw_text),
                ]
            )

        if index % 250 == 0 or index == total:
            print(
                f"validated={index:,}/{total:,} "
                f"pass={status_counts[PASS]:,} "
                f"review={index - status_counts[PASS]:,}",
                flush=True,
            )

    connection.close()

    def output_status_rank(value: str) -> int:
        if value.startswith("수기검증필요"):
            return 0
        if value.startswith("원천미연결"):
            return 1
        return 2

    product_rows.sort(
        key=lambda values: (
            output_status_rank(str(values[7])),
            str(values[8]),
            str(values[0]),
        )
    )
    unlinked_rows.sort(key=lambda values: (str(values[10]), str(values[0])))
    raw_rows.sort(
        key=lambda values: (
            output_status_rank(str(values[6])),
            str(values[7]),
            str(values[0]),
            int(values[1]),
        )
    )

    summary = {
        "assessed_at": assessed_at,
        "database": str(DB_PATH),
        "output": str(OUTPUT),
        "target_resolution_status": (
            "SOURCE_CONFIRMED + 수기보정"
            f"(원 SOURCE_CONFIRMED) {manual_source_repair_count}건"
        ),
        "target_products": total,
        "validation_status_counts": dict(status_counts),
        "proof_grade_counts": dict(proof_counts),
        "warning_product_counts": dict(warning_counts),
        "source_product_counts": dict(source_counts),
        "source_raw_counts": dict(raw_source_counts),
        "direct_signal_product_counts": dict(direct_signal_counts),
        "raw_evidence_rows": len(raw_rows),
        "database_mutated": False,
    }

    workbook = xlsxwriter.Workbook(OUTPUT)
    summary_rows = [
        [
            "전체 검증 대상",
            total,
            "기존 확정값으로 승계된 4,152개 상품의 W/D/H를 다시 검증했습니다.",
        ],
        [
            "1. 상품명에 있는 정보와 다양한 방식으로 추출한 규격 후보를 매칭",
            "",
            "상품명에 포함된 숫자·옵션·모델·제품 종류를 API·HTML·OCR 규격 후보와 대조했습니다.",
        ],
        [
            "1-1. 상품명 숫자와 후보 W/D/H 대조",
            direct_signal_counts["상품명 숫자"],
            "예: 상품명에 1600이 있으면 추출 후보 W/D/H 중 1600이 포함된 후보인지 확인했습니다.",
        ],
        [
            "1-2. 상품명 옵션 코드와 OCR 옵션 행 대조",
            direct_signal_counts["옵션 코드"],
            "상품명의 SS/Q/LQ/K 등 사이즈 코드를 OCR에서 추출한 옵션 행과 대조해 같은 옵션 규격인지 확인했습니다.",
        ],
        [
            "1-3. 상품명 모델 코드와 상세 문맥 모델 코드 대조",
            direct_signal_counts["모델 코드"],
            "상품명에 모델 코드가 있으면 API·HTML·OCR 규격 주변의 모델 코드와 대조해 같은 모델인지 확인했습니다.",
        ],
        [
            "1-4. 상품 종류와 후보 문맥의 제품 종류 대조",
            direct_signal_counts["제품 유형"],
            "상품명의 소파·테이블·침대·벤치·스툴 등 제품 종류와 규격 주변 문맥의 제품 종류가 같은지 확인했습니다.",
        ],
        [
            "2. API·HTML·OCR에서 제품 규격을 다시 확인",
            "",
            "확정값과 같은 숫자가 존재하는 것뿐 아니라, 그 값이 제품 크기를 설명하는 위치에 있는지 확인했습니다.",
        ],
        [
            "2-1. 상품 API의 구조화된 크기 필드 확인",
            direct_signal_counts["구조화 API 크기필드"],
            "동일 productId의 상품정보고시에서 제목이 크기·치수·규격·사이즈인 필드와 확정 W/D/H를 대조했습니다.",
        ],
        [
            "2-2. 제품 사이즈·규격 제목 아래의 값 확인",
            direct_signal_counts["제품 규격 문맥"],
            "상세 HTML·OCR에서 제품 사이즈, 규격, 치수, Dimension 등 제목 아래에 있는 후보인지 확인했습니다.",
        ],
        [
            "2-3. API에서 확정값과 같은 RAW를 연결한 상품",
            source_counts["API"],
            f"상품 API·FAQ/Q&A에서 확정값과 같은 RAW를 연결했습니다. RAW 행은 {raw_source_counts['API']:,}건입니다.",
        ],
        [
            "2-4. HTML에서 확정값과 같은 RAW를 연결한 상품",
            source_counts["HTML"],
            f"상품 상세 HTML·PDP HTML에서 확정값과 같은 RAW를 연결했습니다. RAW 행은 {raw_source_counts['HTML']:,}건입니다.",
        ],
        [
            "2-5. OCR에서 확정값과 같은 RAW를 연결한 상품",
            source_counts["OCR"],
            f"상세 이미지 OCR에서 확정값과 같은 판단 구간을 연결했습니다. RAW 행은 {raw_source_counts['OCR']:,}건입니다.",
        ],
        [
            "3. 제품 규격이 아닌 후보와 더 강한 경쟁 후보를 제외",
            "",
            "배송·포장·구성품·라인업 규격과 다른 제품 후보를 자동 완료에서 제외했습니다.",
        ],
        [
            "3-1. 배송·포장·구성품·라인업 등 위험 문맥 확인",
            status_counts[CHECK_RISK],
            "확정값과 같은 숫자는 있지만 제품 규격이 아닐 가능성이 있어 수기검증 대상으로 분리했습니다.",
        ],
        [
            "3-2. 현재 확정값보다 더 강한 다른 규격 후보 확인",
            status_counts[CHECK_ALTERNATIVE],
            "같은 상품 안에서 상품명 매칭점수 또는 복수 원천 지지가 더 높은 다른 규격 후보를 수기검증 대상으로 분리했습니다.",
        ],
        [
            "3-3. 상품명·제품 규격 문맥 근거가 약한 후보 확인",
            status_counts[CHECK_WEAK],
            "RAW 숫자는 일치하지만 상품명 직접 매칭이나 제품 규격 제목 근거가 약한 상품을 수기검증 대상으로 분리했습니다.",
        ],
        [
            "4. 최종 판정",
            "",
            "상품명 매칭, 원천 위치, 제품 규격 문맥과 오탐 제외 결과를 합쳐 최종 상태를 결정했습니다.",
        ],
        [
            "4-1. 자동검증 완료",
            status_counts[PASS],
            "구조화 API 규격, 제품 규격 문맥 또는 상품명 직접 일치 근거로 자동검증을 완료했습니다.",
        ],
        [
            "└ 구조화 API 규격으로 완료",
            proof_counts["구조화API규격"],
            "동일 productId의 상품정보고시 크기 필드와 확정값이 일치했습니다.",
        ],
        [
            "└ 제품 규격 문맥으로 완료",
            proof_counts["제품규격문맥"],
            "제품 사이즈·규격 제목 구역의 값과 확정값이 일치했습니다.",
        ],
        [
            "└ 상품명 직접 일치로 완료",
            proof_counts["상품명직접일치"],
            "상품명 숫자·옵션·모델·제품 종류와 규격 후보가 직접 일치했습니다.",
        ],
        [
            "4-2. 수기검증 필요",
            total - status_counts[PASS] - status_counts[RESEARCH],
            "위험 문맥, 다른 후보 우선 또는 상품명 근거 약함 사유를 사람이 확인해야 합니다.",
        ],
        [
            "└ 위험 문맥 확인 필요",
            status_counts[CHECK_RISK],
            "배송·포장·구성품·라인업 규격 가능성이 있습니다.",
        ],
        [
            "└ 다른 후보 우선 확인 필요",
            status_counts[CHECK_ALTERNATIVE],
            "현재 확정값보다 더 강하게 매칭되는 규격 후보가 있습니다.",
        ],
        [
            "└ 상품명 근거 약함 확인 필요",
            status_counts[CHECK_WEAK],
            "RAW 숫자 일치 외에 같은 제품임을 확정할 근거가 부족합니다.",
        ],
        [
            "4-3. 원천 미연결",
            status_counts[RESEARCH],
            "현재 DB RAW에서 확정값과 같은 규격을 연결하지 못해 재수집이 필요한 상품입니다.",
        ],
        ["생성 시각", assessed_at, ""],
        ["DB", str(DB_PATH), "읽기 전용 검증; 확정값 변경 없음"],
    ]
    write_sheet(
        workbook,
        "00_통계",
        ["검증 내용", "상품 수", "판정 방식 및 설명"],
        summary_rows,
        [54, 18, 110],
    )

    product_headers = [
        "상품 ID",
        "상품명",
        "중카테고리",
        "소카테고리",
        "W (mm)",
        "D (mm)",
        "H (mm)",
        "증빙 상태",
        "사유",
        "근거 유형",
        "API URL",
        "API 응답 RAW",
        "HTML URL",
        "사이즈 HTML RAW",
        "OCR 이미지 URL",
        "OCR 판단 구간",
        "OCR 추출 위치",
        "RAW 증빙 건수",
        "대표 근거 위치",
        "대표 근거 원문",
        "확정 상태",
        "확정 규칙",
        "확정 원천",
        "원천 snapshot ID",
        "상품 PDP URL",
    ]
    product_widths = [
        17, 34, 18, 22, 12, 12, 12, 32, 90, 16, 48, 100, 48, 100, 48,
        100, 70, 13, 46, 100, 20, 24, 34, 34, 48,
    ]
    write_sheet(
        workbook,
        "01_확정규격_4152",
        product_headers,
        product_rows,
        product_widths,
        status_column=7,
        url_columns={10, 12, 14, 24},
    )

    raw_headers = [
        "상품 ID",
        "증빙 SEQ",
        "상품명",
        "확정 W (mm)",
        "확정 D (mm)",
        "확정 H (mm)",
        "증빙 상태",
        "사유",
        "근거 유형",
        "원천 상세",
        "원천 URL",
        "RAW 위치",
        "API 응답 RAW",
        "사이즈 HTML RAW",
        "OCR 판단 구간",
        "OCR 추출 위치",
        "OCR 전체 문맥",
        "OCR bbox/crop",
        "파싱 W (mm)",
        "파싱 D (mm)",
        "파싱 H (mm)",
        "일치 판정",
        "수집 시각",
        "원천 snapshot ID",
        "원본 payload SHA-256",
        "원본 payload 문자수",
    ]
    raw_widths = [
        17, 10, 34, 12, 12, 12, 32, 90, 13, 26, 48, 46, 100, 100,
        100, 70, 100, 56, 12, 12, 12, 22, 24, 34, 68, 18,
    ]
    write_sheet(
        workbook,
        "02_증빙_RAW_건별",
        raw_headers,
        raw_rows,
        raw_widths,
        status_column=6,
        url_columns={10},
    )

    field_rows = [
        ["01_확정규격_4152", "5,713개 RAW 증빙 파일과 같은 구조입니다. 수기검증 행을 먼저, 통합 사유와 상품 ID 순으로 정렬했습니다."],
        ["증빙 상태", "`증빙완료/수기검증필요/원천미연결 | 증빙 등급`을 한 셀에 함께 표시합니다."],
        ["사유", "판정 사유, OCR 문맥 경고, 배송·반입 재검토, 경쟁 후보를 한 셀로 통합했습니다."],
        ["API 응답 RAW", "확정 규격을 포함한 상품 API·FAQ/Q&A JSON 원문입니다."],
        ["사이즈 HTML RAW", "상품 상세 HTML 또는 PDP HTML의 규격 원문입니다."],
        ["OCR 판단 구간", "확정값 주변 OCR 문맥과 정규화 결과입니다."],
        ["OCR 추출 위치", "전체 이미지 또는 OCR 문자열 내 채택 위치입니다."],
        ["02_증빙_RAW_건별", "한 상품에 증빙이 여러 개이면 RAW 하나당 한 행으로 분리합니다."],
        ["04_원천미연결", "확정값과 같은 RAW를 연결하지 못한 상품입니다."],
        ["검증 정책", "상품명 숫자·옵션·모델·제품유형, 제품 규격 제목, 배송·포장·구성품·라인업 문맥과 경쟁 후보를 함께 비교합니다."],
    ]
    write_sheet(
        workbook,
        "03_필드설명",
        ["컬럼/시트", "사용 방법"],
        field_rows,
        [32, 110],
    )
    write_sheet(
        workbook,
        "04_원천미연결",
        [
            "상품 ID",
            "상품명",
            "중카테고리",
            "소카테고리",
            "W (mm)",
            "D (mm)",
            "H (mm)",
            "확정 상태",
            "확정 규칙",
            "확정 원천",
            "사유",
            "상품 PDP URL",
        ],
        unlinked_rows,
        [17, 34, 18, 22, 12, 12, 12, 20, 24, 34, 90, 48],
        url_columns={11},
    )
    workbook.close()

    SUMMARY_OUTPUT.write_text(
        json.dumps(summary, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2), flush=True)
    return summary


if __name__ == "__main__":
    build()
