from __future__ import annotations

import hashlib
import html as html_module
import itertools
import json
import re
import sqlite3
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

import xlsxwriter

import build_homestyle_bulk_workbook as workbook
from bulk_homestyle_collect import (
    DB_PATH,
    GOODS_ENDPOINT,
    PDP_ENDPOINT,
    QNA_ENDPOINT,
    unpack,
)


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "홈스타일_확정규격_5713개_RAW증빙_OCR판단구간_수기검증상태_20mm이하재검증.xlsx"
LOW_VALUE_AUDIT_LABELS = {
    "VALID_THIN_DIMENSION_EVIDENCE_CONFIRMED": "정상 박형 규격_근거확인",
    "INVALID_LOW_DIMENSION_CORRECTED": "오류값_교정완료",
}
EXCEL_TEXT_LIMIT = 32_000

DIMENSION_HINT = re.compile(
    r"(?:크기|치수|규격|사이즈|가로|세로|높이|깊이|폭|지름|직경|"
    r"dimension|dimensions|size|width|depth|height|diameter|"
    r"\(?[WDLH]\)?\s*[:=]?\s*\d|"
    r"\d+(?:\.\d+)?\s*\(\s*[WDLH]\s*\)|"
    r"\d+(?:\.\d+)?\s*(?:mm|cm)?\s*[x×*]\s*\d+)",
    re.IGNORECASE,
)


@dataclass
class Evidence:
    product_id: str
    source_kind: str
    source_detail: str
    source_url: str
    raw_locator: str
    raw_text: str
    parsed_w_mm: float | None
    parsed_d_mm: float | None
    parsed_h_mm: float | None
    judged_notation: str
    source_snapshot_id: str
    collected_at: str
    payload_sha256: str
    payload_character_count: int
    match_type: str = "LOCKED_VALUE_EXACT"
    ocr_full_context: str = ""
    ocr_region: str = ""
    ocr_bbox_json: str = ""
    ocr_context_warning: str = ""
    delivery_context_warning: str = ""


def clean_cell(value: Any, limit: int = EXCEL_TEXT_LIMIT) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\x00", "")
    if len(text) <= limit:
        return text
    return text[: limit - 80] + f"\n…[Excel 셀 제한으로 {len(text) - limit + 80:,}자 생략]"


def compact_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"), default=str)


def payload_metadata(value: Any) -> tuple[str, int]:
    raw = compact_json(value)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest(), len(raw)


def fmt_number(value: float | None) -> str:
    if value is None:
        return ""
    number = float(value)
    return str(int(number)) if number.is_integer() else f"{number:g}"


def number_pattern(value: float | None) -> str:
    if value is None:
        return ""
    number = float(value)
    if number.is_integer():
        token = str(int(number))
        return rf"(?<!\d)0*{re.escape(token)}(?:\.0+)?(?!\d)"
    token = f"{number:g}"
    return rf"(?<!\d)0*{re.escape(token)}(?!\d)"


def best_number_span(
    text: str,
    values: list[float],
) -> tuple[int, int] | None:
    unique_values = list(dict.fromkeys(float(value) for value in values))
    positions: list[list[tuple[int, int]]] = []
    for value in unique_values:
        matches = [
            (match.start(), match.end())
            for match in re.finditer(number_pattern(value), text, re.IGNORECASE)
        ][:30]
        if not matches:
            return None
        positions.append(matches)
    best: tuple[int, int] | None = None
    best_score: float | None = None
    hint_positions = [match.start() for match in DIMENSION_HINT.finditer(text)]
    for combination in itertools.product(*positions):
        start = min(item[0] for item in combination)
        end = max(item[1] for item in combination)
        span = end - start
        hint_distance = min(
            (abs(position - start) for position in hint_positions),
            default=500,
        )
        score = span + min(hint_distance, 500) * 0.25
        if best_score is None or score < best_score:
            best = (start, end)
            best_score = score
    return best


def exact_dimension_span(
    text: str,
    w_mm: float | None,
    d_mm: float | None,
    h_mm: float | None,
) -> tuple[int, int] | None:
    axes = [
        ("W", w_mm),
        ("D", d_mm),
        ("H", h_mm),
    ]
    present = [(axis, value) for axis, value in axes if value is not None]
    if not present:
        return None
    axis_parts = [
        rf"\(?{axis}\)?\s*[:=]?\s*{number_pattern(value)}"
        for axis, value in present
    ]
    axis_pattern = r".{0,28}?".join(axis_parts)
    match = re.search(axis_pattern, text, re.IGNORECASE | re.DOTALL)
    if match:
        return match.start(), match.end()

    postfix_parts = [
        rf"{number_pattern(value)}\s*\(\s*{axis}\s*\)"
        for axis, value in present
    ]
    postfix_pattern = r".{0,18}?".join(postfix_parts)
    match = re.search(postfix_pattern, text, re.IGNORECASE | re.DOTALL)
    if match:
        return match.start(), match.end()

    if len(present) >= 2:
        value_parts = [number_pattern(value) for _, value in present]
        ordered_pattern = (
            r"\s*(?:mm|cm)?\s*(?:[x×X*]|[-–—])\s*".join(value_parts)
        )
        match = re.search(ordered_pattern, text, re.IGNORECASE)
        if match:
            return match.start(), match.end()

    return best_number_span(
        text,
        [float(value) for _, value in present],
    )


def tight_ocr_excerpt(
    text: str,
    w_mm: float | None,
    d_mm: float | None,
    h_mm: float | None,
    *,
    context_radius: int = 120,
) -> tuple[str, str]:
    normalized = re.sub(r"\s+", " ", str(text or "")).strip()
    if not normalized:
        return "", "OCR 문맥 없음"
    span = exact_dimension_span(normalized, w_mm, d_mm, h_mm)
    if not span:
        excerpt = clean_cell(normalized, context_radius * 2)
        return excerpt, f"전체 OCR {len(normalized):,}자 중 자동 위치 미확정"
    match_start, match_end = span
    start = max(0, match_start - context_radius)
    end = min(len(normalized), match_end + context_radius)
    before = normalized[start:match_start].strip()
    matched = normalized[match_start:match_end].strip()
    after = normalized[match_end:end].strip()
    excerpt = " ".join(
        part for part in (before, f"【{matched}】", after) if part
    )
    location = (
        f"전체 OCR {len(normalized):,}자 중 문맥 {start + 1:,}~{end:,}자 / "
        f"채택 표기 {match_start + 1:,}~{match_end:,}자"
    )
    return excerpt, location


def ocr_context_warnings(excerpt: str) -> tuple[str, str]:
    compact = re.sub(r"\s+", "", excerpt).casefold()
    general_warning_rules = [
        (("포장", "박스포함", "박스크기"), "포장 규격 가능성"),
        (("전구", "소켓", "램프크기"), "구성품 규격 가능성"),
        (("허용오차", "오차범위", "±"), "허용오차 표기 가능성"),
        (("라인업", "다른모델", "별도모델"), "다른 모델·라인업 규격 가능성"),
    ]
    general = "|".join(
        label
        for keywords, label in general_warning_rules
        if any(keyword.casefold() in compact for keyword in keywords)
    )
    delivery_keywords = ("배송", "반입", "엘리베이터", "사다리차", "문폭")
    delivery = (
        "배송·반입 규격 재검토"
        if any(keyword.casefold() in compact for keyword in delivery_keywords)
        else ""
    )
    return general, delivery


def same_number(left: Any, right: Any, tolerance: float = 0.11) -> bool:
    if left is None or right is None:
        return left is None and right is None
    return abs(float(left) - float(right)) <= tolerance


def record_matches(
    record: dict[str, Any],
    locked_w: float | None,
    locked_d: float | None,
    locked_h: float | None,
) -> bool:
    return (
        same_number(record.get("w_mm"), locked_w)
        and same_number(record.get("d_mm"), locked_d)
        and same_number(record.get("h_mm"), locked_h)
    )


def has_dimension_hint(value: str) -> bool:
    return bool(re.search(r"\d", value) and DIMENSION_HINT.search(value))


def dimension_records(value: str, label: str) -> list[dict[str, Any]]:
    if not value or not has_dimension_hint(value):
        return []
    return workbook.dimension_records([(label, value)])


def walk_strings(
    value: Any,
    path: str = "$",
    parent: Any = None,
    *,
    excluded_keys: frozenset[str] = frozenset(),
) -> Iterable[tuple[str, str, Any]]:
    if isinstance(value, dict):
        for key, child in value.items():
            if key in excluded_keys:
                continue
            child_path = f"{path}.{key}"
            yield from walk_strings(
                child,
                child_path,
                value,
                excluded_keys=excluded_keys,
            )
    elif isinstance(value, list):
        for index, child in enumerate(value):
            yield from walk_strings(
                child,
                f"{path}[{index}]",
                child,
                excluded_keys=excluded_keys,
            )
    elif isinstance(value, str):
        yield path, value, parent


def raw_parent_fragment(path: str, value: str, parent: Any) -> str:
    candidate = parent if isinstance(parent, (dict, list)) else {"value": value}
    raw = compact_json(candidate)
    if len(raw) <= EXCEL_TEXT_LIMIT:
        return raw
    return compact_json({"raw_locator": path, "raw_value": value})


def html_fragment(
    raw_html: str,
    locked_w: float | None,
    locked_d: float | None,
    locked_h: float | None,
) -> str:
    if len(raw_html) <= EXCEL_TEXT_LIMIT:
        return raw_html
    decoded = html_module.unescape(raw_html)
    search_terms = [
        fmt_number(value)
        for value in (locked_w, locked_d, locked_h)
        if value is not None
    ]
    positions: list[int] = []
    for term in search_terms:
        match = re.search(rf"(?<!\d){re.escape(term)}(?:\.0)?(?!\d)", decoded)
        if match:
            positions.append(match.start())
    if positions:
        center = int(sum(positions) / len(positions))
    else:
        match = DIMENSION_HINT.search(decoded)
        center = match.start() if match else 0
    start = max(0, center - 12_000)
    end = min(len(decoded), start + EXCEL_TEXT_LIMIT)
    return decoded[start:end]


def evidence_key(item: Evidence) -> tuple[str, str, str, str, str]:
    return (
        item.source_kind,
        item.source_detail,
        item.source_url,
        item.raw_locator,
        item.judged_notation,
    )


def source_urls(product_id: str) -> dict[str, str]:
    return {
        "goods": GOODS_ENDPOINT.format(product_id=product_id) + "?epFlagYn=N",
        "qna": (
            QNA_ENDPOINT
            + "?goodsId="
            + product_id
            + "&qnaType=&isMyInquiry=false&excludeSecret=true&pageSize=100&pageNum=1"
        ),
        "pdp": PDP_ENDPOINT + "?productId=" + product_id,
    }


def add_matching_records(
    result: list[Evidence],
    *,
    product_id: str,
    source_kind: str,
    source_detail: str,
    source_url: str,
    raw_locator: str,
    raw_text: str,
    parse_text: str,
    locked_w: float | None,
    locked_d: float | None,
    locked_h: float | None,
    source_snapshot_id: str,
    collected_at: str,
    payload_sha256: str,
    payload_character_count: int,
) -> None:
    for record in dimension_records(parse_text, source_detail):
        if not record_matches(record, locked_w, locked_d, locked_h):
            continue
        result.append(
            Evidence(
                product_id=product_id,
                source_kind=source_kind,
                source_detail=source_detail,
                source_url=source_url,
                raw_locator=raw_locator,
                raw_text=clean_cell(raw_text),
                parsed_w_mm=record.get("w_mm"),
                parsed_d_mm=record.get("d_mm"),
                parsed_h_mm=record.get("h_mm"),
                judged_notation=clean_cell(record.get("raw") or parse_text, 4_000),
                source_snapshot_id=source_snapshot_id,
                collected_at=collected_at,
                payload_sha256=payload_sha256,
                payload_character_count=payload_character_count,
            )
        )


def api_and_html_evidence(row: sqlite3.Row) -> list[Evidence]:
    product_id = row["product_id"]
    locked_w = row["locked_w_mm"]
    locked_d = row["locked_d_mm"]
    locked_h = row["locked_h_mm"]
    urls = source_urls(product_id)
    result: list[Evidence] = []

    goods_payload = unpack(row["goods_blob"]) or {}
    goods_hash, goods_length = payload_metadata(goods_payload)
    goods_data = goods_payload.get("data") or {}
    for path, value, parent in walk_strings(
        goods_data,
        "$.data",
        excluded_keys=frozenset({"detailInfo"}),
    ):
        if not has_dimension_hint(value):
            continue
        add_matching_records(
            result,
            product_id=product_id,
            source_kind="API",
            source_detail="상품 API JSON",
            source_url=urls["goods"],
            raw_locator=path,
            raw_text=raw_parent_fragment(path, value, parent),
            parse_text=value,
            locked_w=locked_w,
            locked_d=locked_d,
            locked_h=locked_h,
            source_snapshot_id=row["source_snapshot_id"] or "",
            collected_at=row["structured_at"] or "",
            payload_sha256=goods_hash,
            payload_character_count=goods_length,
        )

    detail_html = str(goods_data.get("detailInfo") or "")
    if detail_html:
        html_hash = hashlib.sha256(detail_html.encode("utf-8")).hexdigest()
        snippets = workbook.dimension_keyword_snippets(detail_html)
        if not snippets and has_dimension_hint(workbook.clean_text(detail_html, 30_000)):
            snippets = [workbook.clean_text(detail_html, 30_000)]
        for index, snippet in enumerate(snippets, start=1):
            add_matching_records(
                result,
                product_id=product_id,
                source_kind="HTML",
                source_detail="상품 상세 HTML(detailInfo)",
                source_url=urls["pdp"],
                raw_locator=f"goods API $.data.detailInfo / 규격 문맥 {index}",
                raw_text=html_fragment(detail_html, locked_w, locked_d, locked_h),
                parse_text=snippet,
                locked_w=locked_w,
                locked_d=locked_d,
                locked_h=locked_h,
                source_snapshot_id=row["source_snapshot_id"] or "",
                collected_at=row["structured_at"] or "",
                payload_sha256=html_hash,
                payload_character_count=len(detail_html),
            )

    html_payload = unpack(row["html_blob"]) or {}
    html_hash, html_length = payload_metadata(html_payload)
    for index, signal in enumerate(html_payload.get("dimension_signals") or [], start=1):
        signal_text = str(signal or "")
        add_matching_records(
            result,
            product_id=product_id,
            source_kind="HTML",
            source_detail="PDP HTML 규격 문맥",
            source_url=urls["pdp"],
            raw_locator=f"sources.html_blob.dimension_signals[{index - 1}]",
            raw_text=signal_text,
            parse_text=signal_text,
            locked_w=locked_w,
            locked_d=locked_d,
            locked_h=locked_h,
            source_snapshot_id=row["source_snapshot_id"] or "",
            collected_at=row["html_at"] or "",
            payload_sha256=html_hash,
            payload_character_count=html_length,
        )

    qna_payload = unpack(row["qna_blob"]) or {}
    qna_hash, qna_length = payload_metadata(qna_payload)
    for path, value, parent in walk_strings(qna_payload):
        if not has_dimension_hint(value):
            continue
        add_matching_records(
            result,
            product_id=product_id,
            source_kind="API",
            source_detail="FAQ/Q&A API JSON",
            source_url=urls["qna"],
            raw_locator=path,
            raw_text=raw_parent_fragment(path, value, parent),
            parse_text=value,
            locked_w=locked_w,
            locked_d=locked_d,
            locked_h=locked_h,
            source_snapshot_id=row["source_snapshot_id"] or "",
            collected_at=row["structured_at"] or "",
            payload_sha256=qna_hash,
            payload_character_count=qna_length,
        )

    return list({evidence_key(item): item for item in result}.values())


def ocr_blob_evidence(row: sqlite3.Row) -> list[Evidence]:
    product_id = row["product_id"]
    locked_w = row["locked_w_mm"]
    locked_d = row["locked_d_mm"]
    locked_h = row["locked_h_mm"]
    ocr_payload = unpack(row["ocr_blob"]) or {}
    if not ocr_payload:
        return []
    ocr_hash, ocr_length = payload_metadata(ocr_payload)
    selected = ocr_payload.get("selected") or {}
    image_url = str(selected.get("url") or "")
    full_ocr_text = str(
        (ocr_payload.get("ocr") or {}).get("text")
        or ocr_payload.get("combined_text")
        or ""
    )
    result: list[Evidence] = []
    candidates: list[tuple[str, str, str]] = []
    for index, (label, value) in enumerate(
        workbook.verified_dimension_texts(ocr_payload),
        start=1,
    ):
        candidates.append(
            (
                f"sources.ocr_blob.dimension_reinforcements[{index - 1}]",
                label,
                value,
            )
        )
    dimension_text = str(ocr_payload.get("dimension_text") or "")
    if dimension_text:
        candidates.append(("sources.ocr_blob.dimension_text", "상세 이미지 OCR", dimension_text))
    for locator, detail, value in candidates:
        for record in dimension_records(value, detail):
            if not record_matches(record, locked_w, locked_d, locked_h):
                continue
            excerpt, text_location = tight_ocr_excerpt(
                full_ocr_text or value,
                locked_w,
                locked_d,
                locked_h,
            )
            general_warning, delivery_warning = ocr_context_warnings(excerpt)
            judgment = (
                f"[OCR 판단 구간]\n{excerpt}\n\n"
                f"[정규화 결과]\n"
                f"W={fmt_number(locked_w)} / D={fmt_number(locked_d)} / "
                f"H={fmt_number(locked_h)} mm"
            )
            image_scope = (
                f"전체 이미지 OCR"
                f" ({ocr_payload.get('width') or '?'}×{ocr_payload.get('height') or '?'} px,"
                f" tile={((ocr_payload.get('ocr') or {}).get('tile_count') or '?')}); "
                f"{text_location}; 픽셀 bbox 미저장"
            )
            result.append(
                Evidence(
                    product_id=product_id,
                    source_kind="OCR",
                    source_detail=detail,
                    source_url=image_url or source_urls(product_id)["pdp"],
                    raw_locator=locator,
                    raw_text=clean_cell(compact_json(ocr_payload)),
                    parsed_w_mm=record.get("w_mm"),
                    parsed_d_mm=record.get("d_mm"),
                    parsed_h_mm=record.get("h_mm"),
                    judged_notation=clean_cell(judgment),
                    source_snapshot_id=row["source_snapshot_id"] or "",
                    collected_at=row["ocr_at"] or "",
                    payload_sha256=ocr_hash,
                    payload_character_count=ocr_length,
                    ocr_full_context=clean_cell(full_ocr_text or value),
                    ocr_region=image_scope,
                    ocr_context_warning=general_warning,
                    delivery_context_warning=delivery_warning,
                )
            )
    return list({evidence_key(item): item for item in result}.values())


def selected_candidate_evidence(
    row: sqlite3.Row,
    candidate: dict[str, Any] | None,
) -> list[Evidence]:
    if not candidate:
        return []
    raw_notation = str(candidate.get("raw_notation") or "")
    context_text = str(candidate.get("context_text") or "")
    image_url = str(candidate.get("image_url") or "")
    raw = compact_json(
        {
            "source_type": candidate.get("source_type"),
            "source_ref": candidate.get("source_ref"),
            "image_url": image_url,
            "raw_notation": raw_notation,
            "context_text": context_text,
            "normalized_axis_mapping": candidate.get("normalized_axis_mapping"),
            "decision_status": candidate.get("decision_status"),
            "candidate_score": candidate.get("candidate_score"),
        }
    )
    digest = hashlib.sha256(raw.encode("utf-8")).hexdigest()
    locked_w = row["locked_w_mm"]
    locked_d = row["locked_d_mm"]
    locked_h = row["locked_h_mm"]
    excerpt, text_location = tight_ocr_excerpt(
        context_text or raw_notation,
        locked_w,
        locked_d,
        locked_h,
    )
    general_warning, delivery_warning = ocr_context_warnings(excerpt)
    judgment = (
        f"[OCR 판단 구간]\n{excerpt}\n\n"
        f"[정규화 결과]\n"
        f"W={fmt_number(locked_w)} / D={fmt_number(locked_d)} / "
        f"H={fmt_number(locked_h)} mm"
    )
    source_type = str(candidate.get("source_type") or "OCR")
    source_ref = str(candidate.get("source_ref") or "")
    region_hint = str(candidate.get("_ocr_region_hint") or "")
    if "TARGETED_REGION_OCR" in source_type:
        scope = region_hint or f"표적 영역 OCR ({source_ref})"
    elif "FULL_IMAGE" in source_type:
        scope = f"전체 이미지 OCR ({source_ref})"
    else:
        scope = f"OCR 후보 문맥 ({source_ref})"
    region = f"{scope}; {text_location}"
    return [
        Evidence(
            product_id=row["product_id"],
            source_kind="OCR",
            source_detail=str(candidate.get("source_type") or "OCR") + " / 선택 후보",
            source_url=image_url or source_urls(row["product_id"])["pdp"],
            raw_locator=(
                "stg_dimension_context_candidate:"
                + str(candidate.get("candidate_key") or "")
            ),
            raw_text=clean_cell(raw),
            parsed_w_mm=candidate.get("w_mm"),
            parsed_d_mm=candidate.get("d_mm"),
            parsed_h_mm=candidate.get("h_mm"),
            judged_notation=clean_cell(judgment),
            source_snapshot_id=str(candidate.get("snapshot_id") or ""),
            collected_at=str(candidate.get("normalized_at") or ""),
            payload_sha256=digest,
            payload_character_count=len(raw),
            ocr_full_context=clean_cell(context_text or raw_notation),
            ocr_region=region,
            ocr_bbox_json=clean_cell(candidate.get("_ocr_bbox_json") or ""),
            ocr_context_warning=general_warning,
            delivery_context_warning=delivery_warning,
        )
    ]


def fallback_candidate_evidence(
    row: sqlite3.Row,
    candidates: list[dict[str, Any]],
) -> list[Evidence]:
    result: list[Evidence] = []
    for candidate in candidates[:3]:
        result.extend(selected_candidate_evidence(row, candidate))
    return result


def load() -> tuple[
    list[sqlite3.Row],
    dict[str, dict[str, Any]],
    dict[str, list[dict[str, Any]]],
]:
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    rows = connection.execute(
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
        ORDER BY l.product_id
        """
    ).fetchall()

    locked_by_product = {row["product_id"]: row for row in rows}
    pass2_bbox = {
        (row["url_hash"], row["raw_notation"]): row["bbox_json"] or ""
        for row in connection.execute(
            """
            SELECT url_hash, raw_notation, bbox_json
            FROM stg_dimension_scan_pass2_observation
            WHERE COALESCE(bbox_json, '') != ''
            """
        )
    }
    crop_metadata = {
        (
            row["run_name"],
            row["product_id"],
            row["image_url"],
            int(row["crop_no"]),
        ): dict(row)
        for row in connection.execute(
            """
            SELECT
                run_name, product_id, image_url, crop_no, crop_file,
                crop_top, crop_bottom, crop_scale
            FROM stg_dimension_targeted_ocr_crop
            """
        )
    }
    selected_candidates: dict[str, dict[str, Any]] = {}
    fallback_candidates: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for candidate in connection.execute(
        """
        SELECT *
        FROM stg_dimension_context_candidate
        ORDER BY
            product_id,
            is_current DESC,
            COALESCE(candidate_score, 0) DESC,
            normalized_at DESC
        """
    ):
        product_id = candidate["product_id"]
        locked = locked_by_product.get(product_id)
        if not locked:
            continue
        candidate_data = dict(candidate)
        source_type = str(candidate["source_type"] or "")
        source_ref = str(candidate["source_ref"] or "")
        if source_type == "PASS2_FULL_IMAGE_OCR":
            candidate_data["_ocr_bbox_json"] = pass2_bbox.get(
                (source_ref, candidate["raw_notation"]),
                "",
            )
        if "TARGETED_REGION_OCR" in source_type:
            ref_match = re.match(
                r"(?P<run>[^:]+):crop=(?P<crop>\d+):candidate=\d+",
                source_ref,
            )
            if ref_match:
                crop_no = int(ref_match.group("crop"))
                crop = crop_metadata.get(
                    (
                        ref_match.group("run"),
                        product_id,
                        candidate["image_url"],
                        crop_no,
                    )
                )
                if crop:
                    candidate_data["_ocr_region_hint"] = (
                        f"표적 영역 OCR crop {crop_no}; "
                        f"세로 pixel {crop['crop_top']}~{crop['crop_bottom']}; "
                        f"확대 {crop['crop_scale']}배; 파일 {crop['crop_file']}"
                    )
                    candidate_data["_ocr_bbox_json"] = compact_json(
                        {
                            "crop_no": crop_no,
                            "crop_file": crop["crop_file"],
                            "crop_top": crop["crop_top"],
                            "crop_bottom": crop["crop_bottom"],
                            "crop_scale": crop["crop_scale"],
                        }
                    )
        if (
            locked["resolution_status"] == "RULE_RESOLVED"
            and candidate["candidate_key"] == locked["representative_candidate_key"]
            and candidate["snapshot_id"] == locked["source_snapshot_id"]
        ):
            selected_candidates[product_id] = candidate_data
        if not (
            same_number(candidate["w_mm"], locked["locked_w_mm"])
            and same_number(candidate["d_mm"], locked["locked_d_mm"])
            and same_number(candidate["h_mm"], locked["locked_h_mm"])
        ):
            continue
        product_candidates = fallback_candidates[product_id]
        key = (
            candidate["source_type"],
            candidate["source_ref"],
            candidate["raw_notation"],
            candidate["image_url"],
        )
        if all(
            (
                item.get("source_type"),
                item.get("source_ref"),
                item.get("raw_notation"),
                item.get("image_url"),
            )
            != key
            for item in product_candidates
        ):
            product_candidates.append(candidate_data)
    connection.close()
    return rows, selected_candidates, fallback_candidates


def primary_evidence(items: list[Evidence]) -> Evidence | None:
    priority = {"API": 0, "HTML": 1, "OCR": 2}
    return min(
        items,
        key=lambda item: (
            priority.get(item.source_kind, 9),
            0 if item.source_url else 1,
            len(item.raw_text),
        ),
        default=None,
    )


def product_evidence_status(items: list[Evidence]) -> str:
    if not items:
        return "원천미연결"
    if any(
        item.ocr_context_warning or item.delivery_context_warning
        for item in items
    ):
        return "수기검증필요"
    return "증빙완료"


def combined_by_kind(items: list[Evidence], kind: str, field: str) -> str:
    values = []
    for index, item in enumerate(
        [item for item in items if item.source_kind == kind],
        start=1,
    ):
        value = getattr(item, field)
        if not value:
            continue
        values.append(f"[{kind} RAW {index}]\n{value}")
    return clean_cell("\n\n--------------------\n\n".join(values))


def write_url_or_text(
    worksheet: xlsxwriter.worksheet.Worksheet,
    row: int,
    col: int,
    value: str,
    cell_format: xlsxwriter.format.Format,
) -> None:
    if (
        value.startswith(("http://", "https://"))
        and "\n" not in value
        and len(value) <= 2_000
    ):
        worksheet.write_url(row, col, value, cell_format, value)
    else:
        worksheet.write(row, col, clean_cell(value), cell_format)


def add_table_sheet(
    xlsx: xlsxwriter.Workbook,
    name: str,
    headers: list[str],
    rows: list[list[Any]],
    widths: list[int],
    *,
    url_columns: frozenset[int] = frozenset(),
    warning_column: int | None = None,
    alert_columns: frozenset[int] = frozenset(),
) -> None:
    sheet = xlsx.add_worksheet(name)
    header = xlsx.add_format(
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
    body = xlsx.add_format(
        {"border": 1, "valign": "top", "text_wrap": True}
    )
    url_format = xlsx.add_format(
        {
            "border": 1,
            "valign": "top",
            "text_wrap": True,
            "font_color": "blue",
            "underline": True,
        }
    )
    warning = xlsx.add_format(
        {
            "border": 1,
            "valign": "top",
            "text_wrap": True,
            "bg_color": "#FCE4D6",
            "font_color": "#9C0006",
        }
    )
    for col, value in enumerate(headers):
        sheet.write(0, col, value, header)
    for row_index, values in enumerate(rows, start=1):
        for col, value in enumerate(values):
            selected_format = body
            if (
                warning_column is not None
                and col == warning_column
                and str(value) != "증빙완료"
            ):
                selected_format = warning
            if col in alert_columns and str(value or "").strip():
                selected_format = warning
            if col in url_columns:
                write_url_or_text(
                    sheet,
                    row_index,
                    col,
                    "" if value is None else str(value),
                    url_format if value else selected_format,
                )
            elif isinstance(value, (int, float)) and not isinstance(value, bool):
                sheet.write_number(row_index, col, value, selected_format)
            else:
                sheet.write(
                    row_index,
                    col,
                    clean_cell(value),
                    selected_format,
                )
    sheet.freeze_panes(1, 3)
    sheet.autofilter(0, 0, len(rows), len(headers) - 1)
    for col, width in enumerate(widths):
        sheet.set_column(col, col, width)
    sheet.set_row(0, 34)


def build() -> dict[str, Any]:
    locked_rows, selected_candidates, fallback_candidates = load()
    if len(locked_rows) != 5_713:
        raise RuntimeError(f"확정 규격 건수가 5,713개가 아닙니다: {len(locked_rows):,}")
    audit_connection = sqlite3.connect(DB_PATH)
    audit_connection.row_factory = sqlite3.Row
    low_value_audits = {
        row["product_id"]: row
        for row in audit_connection.execute(
            """
            SELECT *
            FROM stg_dimension_low_value_audit
            WHERE is_current=1
            ORDER BY product_id
            """
        )
    }
    audit_connection.close()
    low_value_audit_counts = Counter(
        row["audit_status"] for row in low_value_audits.values()
    )

    evidence_by_product: dict[str, list[Evidence]] = {}
    evidence_status_counts: Counter[str] = Counter()
    source_product_counts: Counter[str] = Counter()
    source_raw_counts: Counter[str] = Counter()
    ocr_warning_product_counts: Counter[str] = Counter()
    ocr_warning_products: set[str] = set()
    delivery_review_products: set[str] = set()

    for row_number, row in enumerate(locked_rows, start=1):
        product_id = row["product_id"]
        items: list[Evidence] = []
        if row["resolution_status"] == "RULE_RESOLVED":
            items.extend(
                selected_candidate_evidence(
                    row,
                    selected_candidates.get(product_id),
                )
            )
        items.extend(api_and_html_evidence(row))
        items.extend(ocr_blob_evidence(row))
        if not items:
            items.extend(
                fallback_candidate_evidence(
                    row,
                    fallback_candidates.get(product_id, []),
                )
            )
        items = list({evidence_key(item): item for item in items}.values())
        evidence_by_product[product_id] = items
        status = product_evidence_status(items)
        evidence_status_counts[status] += 1
        for kind in {item.source_kind for item in items}:
            source_product_counts[kind] += 1
        source_raw_counts.update(item.source_kind for item in items)
        product_warnings = {
            warning
            for item in items
            for warning in item.ocr_context_warning.split("|")
            if warning
        }
        if product_warnings:
            ocr_warning_products.add(product_id)
        for warning in product_warnings:
            ocr_warning_product_counts[warning] += 1
        if any(item.delivery_context_warning for item in items):
            delivery_review_products.add(product_id)
        if row_number % 500 == 0 or row_number == len(locked_rows):
            print(
                f"evidence_parse={row_number:,}/{len(locked_rows):,} "
                f"complete={evidence_status_counts['증빙완료']:,} "
                f"unlinked={evidence_status_counts['원천미연결']:,}",
                flush=True,
            )

    main_headers = [
        "상품 ID",
        "상품명",
        "중카테고리",
        "소카테고리",
        "W (mm)",
        "D (mm)",
        "H (mm)",
        "20mm 이하 재검증 상태",
        "20mm 이하 재검증 사유",
        "증빙 상태",
        "근거 유형",
        "API URL",
        "API 응답 RAW",
        "HTML URL",
        "사이즈 HTML RAW",
        "OCR 이미지 URL",
        "OCR 판단 구간",
        "OCR 추출 위치",
        "OCR 문맥 경고(배송·반입 제외)",
        "배송·반입 규격 재검토",
        "RAW 증빙 건수",
        "대표 근거 위치",
        "대표 근거 원문",
        "확정 상태",
        "확정 규칙",
        "확정 원천",
        "원천 snapshot ID",
        "상품 PDP URL",
    ]
    main_rows: list[list[Any]] = []
    raw_headers = [
        "상품 ID",
        "증빙 SEQ",
        "상품명",
        "확정 W (mm)",
        "확정 D (mm)",
        "확정 H (mm)",
        "근거 유형",
        "원천 상세",
        "원천 URL",
        "RAW 위치",
        "API 응답 RAW",
        "사이즈 HTML RAW",
        "OCR 판단 구간",
        "OCR 추출 위치",
        "OCR 문맥 경고(배송·반입 제외)",
        "배송·반입 규격 재검토",
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
    raw_rows: list[list[Any]] = []
    missing_rows: list[list[Any]] = []

    for row in locked_rows:
        product_id = row["product_id"]
        items = evidence_by_product[product_id]
        primary = primary_evidence(items)
        kinds = sorted(
            {item.source_kind for item in items},
            key=lambda kind: ("API", "HTML", "OCR").index(kind),
        )
        api_urls = "\n".join(
            dict.fromkeys(
                item.source_url for item in items
                if item.source_kind == "API" and item.source_url
            )
        )
        html_urls = "\n".join(
            dict.fromkeys(
                item.source_url for item in items
                if item.source_kind == "HTML" and item.source_url
            )
        )
        ocr_urls = "\n".join(
            dict.fromkeys(
                item.source_url for item in items
                if item.source_kind == "OCR" and item.source_url
            )
        )
        status = product_evidence_status(items)
        low_value_audit = low_value_audits.get(product_id)
        main_rows.append(
            [
                product_id,
                row["product_name"] or "",
                row["mid_category"] or "",
                row["small_category"] or "",
                row["locked_w_mm"],
                row["locked_d_mm"],
                row["locked_h_mm"],
                (
                    LOW_VALUE_AUDIT_LABELS.get(
                        low_value_audit["audit_status"],
                        low_value_audit["audit_status"],
                    )
                    if low_value_audit
                    else ""
                ),
                (
                    low_value_audit["audit_reason"]
                    if low_value_audit
                    else ""
                ),
                status,
                "|".join(kinds),
                api_urls,
                combined_by_kind(items, "API", "raw_text"),
                html_urls,
                combined_by_kind(items, "HTML", "raw_text"),
                ocr_urls,
                combined_by_kind(items, "OCR", "judged_notation"),
                combined_by_kind(items, "OCR", "ocr_region"),
                combined_by_kind(items, "OCR", "ocr_context_warning"),
                combined_by_kind(items, "OCR", "delivery_context_warning"),
                len(items),
                primary.raw_locator if primary else "",
                primary.judged_notation if primary else "",
                row["resolution_status"],
                row["resolution_rule_code"] or "",
                row["resolution_source"] or "",
                row["source_snapshot_id"] or "",
                source_urls(product_id)["pdp"],
            ]
        )
        if not items:
            missing_rows.append(
                [
                    product_id,
                    row["product_name"] or "",
                    row["mid_category"] or "",
                    row["small_category"] or "",
                    row["locked_w_mm"],
                    row["locked_d_mm"],
                    row["locked_h_mm"],
                    row["resolution_status"],
                    row["resolution_rule_code"] or "",
                    row["resolution_source"] or "",
                    "현재 DB RAW에서 확정값과 일치하는 원문을 자동 연결하지 못함",
                    source_urls(product_id)["pdp"],
                ]
            )
        for seq, item in enumerate(items, start=1):
            raw_rows.append(
                [
                    product_id,
                    seq,
                    row["product_name"] or "",
                    row["locked_w_mm"],
                    row["locked_d_mm"],
                    row["locked_h_mm"],
                    item.source_kind,
                    item.source_detail,
                    item.source_url,
                    item.raw_locator,
                    item.raw_text if item.source_kind == "API" else "",
                    item.raw_text if item.source_kind == "HTML" else "",
                    item.judged_notation if item.source_kind == "OCR" else "",
                    item.ocr_region if item.source_kind == "OCR" else "",
                    item.ocr_context_warning if item.source_kind == "OCR" else "",
                    item.delivery_context_warning if item.source_kind == "OCR" else "",
                    item.ocr_full_context if item.source_kind == "OCR" else "",
                    item.ocr_bbox_json if item.source_kind == "OCR" else "",
                    item.parsed_w_mm,
                    item.parsed_d_mm,
                    item.parsed_h_mm,
                    item.match_type,
                    item.collected_at,
                    item.source_snapshot_id,
                    item.payload_sha256,
                    item.payload_character_count,
                ]
            )

    generated_at = datetime.now().astimezone().isoformat(timespec="seconds")
    summary_rows = [
        ["항목", "건수", "설명"],
        ["확정 규격 전체", len(locked_rows), "DB is_locked=1"],
        [
            "20mm 이하 정상 박형 규격",
            low_value_audit_counts[
                "VALID_THIN_DIMENSION_EVIDENCE_CONFIRMED"
            ],
            "제품 SIZE/API/HTML 근거를 확인하여 값 유지",
        ],
        [
            "20mm 이하 오류 보정",
            low_value_audit_counts["INVALID_LOW_DIMENSION_CORRECTED"],
            "배송 규격 혼입 또는 원천 충돌을 제품별 상세 규격으로 교체",
        ],
        ["증빙완료", evidence_status_counts["증빙완료"], "RAW 연결 완료 및 OCR 문맥 경고 없음"],
        [
            "수기검증필요",
            evidence_status_counts["수기검증필요"],
            "RAW는 연결됐으나 OCR 문맥 경고 또는 배송·반입 재검토 신호 존재",
        ],
        ["원천미연결", evidence_status_counts["원천미연결"], "재수집 또는 원천 매핑 보강 필요"],
        ["API 증빙 상품", source_product_counts["API"], "상품·Q&A API 응답의 정확한 JSON 노드"],
        ["HTML 증빙 상품", source_product_counts["HTML"], "상품 상세 HTML 또는 PDP HTML 규격 문맥"],
        ["OCR 증빙 상품", source_product_counts["OCR"], "판단 규격과 판단에 사용한 OCR 문맥"],
        ["API RAW 건수", source_raw_counts["API"], "02_증빙_RAW_건별의 API 행"],
        ["HTML RAW 건수", source_raw_counts["HTML"], "02_증빙_RAW_건별의 HTML 행"],
        ["OCR RAW 건수", source_raw_counts["OCR"], "02_증빙_RAW_건별의 OCR 행"],
        [
            "OCR 문맥 경고 상품",
            len(ocr_warning_products),
            "포장·구성품·허용오차·다른 모델 문맥을 가진 고유 상품",
        ],
        *[
            [f"└ {warning}", count, "확정값 주변의 좁은 OCR 문맥에서 탐지"]
            for warning, count in sorted(ocr_warning_product_counts.items())
        ],
        [
            "배송·반입 규격 재검토 상품",
            len(delivery_review_products),
            "일반 OCR 경고와 분리한 독립 재검토 대상",
        ],
        ["생성 시각", generated_at, ""],
        ["DB", str(DB_PATH), ""],
    ]
    guide_rows = [
        ["컬럼/시트", "사용 방법"],
        [
            "01_확정규격_5713",
            "상품별 확정 W/D/H 바로 옆에서 API·HTML·OCR 증빙을 확인합니다.",
        ],
        [
            "02_증빙_RAW_건별",
            "한 상품에 증빙이 여러 개이면 RAW 하나당 한 행으로 분리합니다.",
        ],
        [
            "API 응답 RAW",
            "API 전체 응답을 반복 저장하지 않고, 확정 규격을 포함한 정확한 JSON 노드와 JSONPath를 저장합니다. 전체 payload 해시는 별도 컬럼에 보존합니다.",
        ],
        [
            "사이즈 HTML RAW",
            "goods API의 detailInfo HTML 또는 PDP HTML에서 규격이 존재하는 원문만 저장합니다.",
        ],
        [
            "OCR 판단 구간",
            "대표 시트에는 확정값 주변의 좁은 OCR 문맥만 표시합니다. 정규화 결과와 함께 확인합니다.",
        ],
        [
            "OCR 추출 위치",
            "전체 이미지·표적 crop 구분과 OCR 문자열 내 채택 구간 위치를 표시합니다. 픽셀 bbox가 저장된 원천은 별도 bbox/crop 컬럼에 보존합니다.",
        ],
        [
            "OCR 전체 문맥",
            "대표 시트에서는 숨기고 02_증빙_RAW_건별 시트에서만 감사를 위해 제공합니다.",
        ],
        [
            "OCR 문맥 경고",
            "채택 표기 주변에서 포장·구성품·허용오차·다른 모델 표현이 발견되면 검토 신호를 표시합니다.",
        ],
        [
            "배송·반입 규격 재검토",
            "배송·반입·엘리베이터·사다리차·문폭 표현은 일반 경고에서 제외하고 독립 컬럼으로 이동했습니다.",
        ],
        [
            "04_원천미연결",
            "확정값은 있으나 현재 DB RAW에서 같은 값을 자동 연결하지 못한 상품입니다. 완료 증빙으로 사용하지 않고 재수집 대상으로 분리합니다.",
        ],
    ]

    with xlsxwriter.Workbook(
        OUTPUT,
        {"constant_memory": True, "strings_to_urls": False},
    ) as xlsx:
        add_table_sheet(
            xlsx,
            "00_통계",
            summary_rows[0],
            summary_rows[1:],
            [25, 18, 80],
        )
        add_table_sheet(
            xlsx,
            "01_확정규격_5713",
            main_headers,
            main_rows,
            [
                16, 38, 18, 20, 12, 12, 12, 34, 70, 14, 14, 46, 90, 46,
                90, 46, 70, 55, 28, 25, 13, 42, 55, 19, 24, 30, 35, 46,
            ],
            url_columns=frozenset({11, 13, 15, 27}),
            warning_column=9,
            alert_columns=frozenset({18, 19}),
        )
        add_table_sheet(
            xlsx,
            "02_증빙_RAW_건별",
            raw_headers,
            raw_rows,
            [
                16, 10, 38, 12, 12, 12, 12, 28, 50, 48, 100, 100,
                70, 55, 28, 25, 100, 55, 12, 12, 12, 20, 25, 34, 68, 18,
            ],
            url_columns=frozenset({8}),
            alert_columns=frozenset({14, 15}),
        )
        add_table_sheet(
            xlsx,
            "03_필드설명",
            guide_rows[0],
            guide_rows[1:],
            [30, 110],
        )
        add_table_sheet(
            xlsx,
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
            missing_rows,
            [16, 38, 18, 20, 12, 12, 12, 20, 24, 30, 58, 50],
            url_columns=frozenset({11}),
        )

    result = {
        "output": str(OUTPUT),
        "locked_products": len(locked_rows),
        "evidence_status_counts": dict(evidence_status_counts),
        "source_product_counts": dict(source_product_counts),
        "source_raw_counts": dict(source_raw_counts),
        "ocr_warning_product_counts": dict(ocr_warning_product_counts),
        "ocr_warning_products": len(ocr_warning_products),
        "delivery_review_products": len(delivery_review_products),
        "low_value_audit_counts": dict(low_value_audit_counts),
        "raw_rows": len(raw_rows),
        "missing_rows": len(missing_rows),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return result


if __name__ == "__main__":
    build()
