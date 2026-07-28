from __future__ import annotations

import html
import json
import re
from collections import Counter
from datetime import datetime, timedelta, timezone
from html.parser import HTMLParser
from pathlib import Path
from typing import Any

import requests
import urllib3


ROOT = Path(__file__).resolve().parent
RUN_DIR = ROOT / "poc_full_run"
RAW_HTML_DIR = RUN_DIR / "raw_html"
OUTPUT = RUN_DIR / "faq_qna_probe.json"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 Chrome/150 Safari/537.36"
)

# The local corporate proxy certificate is not available to Python requests in
# this workspace. These calls are limited to the official lge.co.kr hosts below.
VERIFY_TLS = False

PRODUCTS = [
    {
        "product_id": "G25070005743",
        "source_system": "HOMESTYLE",
        "url": "https://homestyle.lge.co.kr/item?productId=G25070005743",
    },
    {
        "product_id": "G25100020496",
        "source_system": "HOMESTYLE",
        "url": "https://homestyle.lge.co.kr/item?productId=G25100020496",
    },
    {
        "product_id": "G25070001871",
        "source_system": "HOMESTYLE",
        "url": "https://homestyle.lge.co.kr/item?productId=G25070001871",
    },
    {
        "product_id": "G25070006112",
        "source_system": "HOMESTYLE",
        "url": "https://homestyle.lge.co.kr/item?productId=G25070006112",
    },
    {
        "product_id": "OLED48C6KNA",
        "source_system": "LGE_APPLIANCE",
        "model_id": "MD10770851",
        "url": "https://www.lge.co.kr/tvs/oled48c6kna-wall",
    },
    {
        "product_id": "G646GBB031",
        "source_system": "LGE_APPLIANCE",
        "model_id": "MD10780848",
        "url": "https://www.lge.co.kr/refrigerators/g646gbb031",
    },
    {
        "product_id": "WA2525YMZF",
        "source_system": "LGE_APPLIANCE",
        "model_id": "MD10576829",
        "url": "https://www.lge.co.kr/wash-tower/wa2525ymzf",
    },
    {
        "product_id": "SQ06GJ1WFS",
        "source_system": "LGE_APPLIANCE",
        "model_id": "MD10766829",
        "url": "https://www.lge.co.kr/air-conditioners/sq06gj1wfs",
    },
]


SIGNAL_PATTERNS = {
    "R1_사이즈_WDH": r"(?:사이즈|크기|규격|가로|세로|폭|너비|높이|깊이|\b[WHD]\b|\d\s*(?:mm|cm))",
    "R1_배치추천공간": r"(?:거실|리빙룸|침실|베드룸|주방|다이닝|서재|현관|베란다|방|영업장|공간)",
    "R1_제품색상": r"(?:색상|컬러|색깔|화이트|블랙|베이지|그레이|실버|골드|네이비|브라운)",
    "R1_설치타입": r"(?:빌트인|스탠딩|스탠드형|벽걸이|매립|천장형|설치형|설치 타입|설치방법)",
    "R1_세트구성": r"(?:구성품|구성 내역|세트 구성|포함(?:되| 여부)|동봉|전구|전선|실외기|리모컨)",
    "R1_배치가능위치": r"(?:벽면|벽걸이|천장|바닥|상판|테이블|탁상|선반|매립)",
    "R1_벽면추천높이": r"(?:(?:벽면|벽걸이).{0,30}(?:높이|위치)|(?:높이|위치).{0,30}(?:벽면|벽걸이))",
    "R2_설명서_재질": r"(?:재질|소재|원단|패브릭|가죽|스테인리스|플라스틱|유리|금속)",
    "R2_설명서_규격": r"(?:규격|사이즈|크기|가로|세로|폭|높이|깊이|용량|\d\s*(?:mm|cm|L))",
    "R2_설명서_구성품": r"(?:구성품|구성 내역|포함|동봉|세트|전구|전선|리모컨|실외기)",
    "R2_설명서_색상": r"(?:색상|컬러|색깔|화이트|블랙|베이지|그레이|실버|골드|네이비|브라운)",
    "R2_설명서_사용목적": r"(?:사용 목적|용도|사용할|사용 시|사용법|활용)",
    "R2_설명서_조립설치": r"(?:조립|설치|시공|타공|배관|연결|장착)",
    "R2_설명서_안전주의": r"(?:안전|주의|위험|화재|감전|금지|유의)",
    "R2_설명서_취급관리": r"(?:청소|관리|세척|보관|취급|관리법|필터)",
    "R2_설명서_품질보증": r"(?:보증|무상|유상|A/S|AS기간|서비스 기간)",
    "R2_설명서_인증": r"(?:인증|KC|에너지효율|등급)",
    "R2_설명서_판매정보": r"(?:구매|주문|가격|할인|재고|배송|반품|교환)",
    "R2_디자인스타일": r"(?:미니멀|직선형|곡선형|유기적|장식형|클래식|모듈형|스칸디나비안|컨템포러리)",
    "R2_공간콘텐츠": r"(?:인테리어|분위기|무드|거실|침실|주방|서재|공간|배치|위치)",
    "R2_공간제품관계": r"(?:공간|거실|침실|주방|서재).{0,40}(?:제품|설치|배치|사용)",
    "R2_제품간관계": r"(?:호환|조합|함께|동시|연결|결합|세트|실내기|실외기)",
}


class TextParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []

    def handle_data(self, data: str) -> None:
        value = re.sub(r"\s+", " ", data).strip()
        if value:
            self.parts.append(value)


def html_to_text(value: Any) -> str:
    if value in (None, ""):
        return ""
    parser = TextParser()
    parser.feed(str(value))
    return " ".join(parser.parts)


def redact_public_text(value: Any) -> str:
    text = html_to_text(value)
    text = re.sub(
        r"(?<![\w.])[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}(?![\w.])",
        "[이메일 마스킹]",
        text,
    )
    text = re.sub(
        r"(?<!\d)(?:(?:\+?82[-\s]?)?0\d{1,2}[-\s]?\d{3,4}[-\s]?\d{4}|1[568]\d{2}[-\s]?\d{4})(?!\d)",
        "[전화번호 마스킹]",
        text,
    )
    return re.sub(r"\s+", " ", text).strip()


def request_json(
    session: requests.Session,
    method: str,
    url: str,
    *,
    params: dict[str, Any] | None = None,
    data: dict[str, Any] | None = None,
    referer: str,
) -> tuple[int, dict[str, Any]]:
    response = session.request(
        method,
        url,
        params=params,
        data=data,
        timeout=30,
        headers={"Accept": "application/json", "Referer": referer},
    )
    try:
        payload = response.json()
    except requests.JSONDecodeError:
        payload = {"_parse_error": response.text[:500]}
    return response.status_code, payload if isinstance(payload, dict) else {"data": payload}


def flatten_jsonld(value: Any) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    if isinstance(value, dict):
        result.append(value)
        if isinstance(value.get("@graph"), list):
            for child in value["@graph"]:
                result.extend(flatten_jsonld(child))
    elif isinstance(value, list):
        for child in value:
            result.extend(flatten_jsonld(child))
    return result


def extract_faq_jsonld(product_id: str) -> list[dict[str, Any]]:
    html_path = RAW_HTML_DIR / f"{product_id}.html"
    source = html_path.read_text(encoding="utf-8", errors="ignore")
    scripts = re.findall(
        r"<script[^>]+type=[\"']application/ld\+json[\"'][^>]*>(.*?)</script>",
        source,
        flags=re.IGNORECASE | re.DOTALL,
    )
    faq_pages: list[dict[str, Any]] = []
    for script in scripts:
        try:
            parsed = json.loads(html.unescape(script.strip()))
        except (json.JSONDecodeError, TypeError):
            continue
        faq_pages.extend(
            item for item in flatten_jsonld(parsed) if item.get("@type") == "FAQPage"
        )

    rows: list[dict[str, Any]] = []
    seen: set[tuple[str, str]] = set()
    for page in faq_pages:
        for index, entity in enumerate(page.get("mainEntity") or [], start=1):
            if not isinstance(entity, dict):
                continue
            answer = entity.get("acceptedAnswer") or {}
            question = redact_public_text(entity.get("name"))
            answer_text = redact_public_text(
                answer.get("text") if isinstance(answer, dict) else answer
            )
            key = (question, answer_text)
            if not question or key in seen:
                continue
            seen.add(key)
            rows.append(
                {
                    "record_id": f"HTML-FAQ-{index:03d}",
                    "source": "HTML_JSONLD_FAQPAGE",
                    "scope": "공식_제품카테고리_FAQ"
                    if product_id[0] != "G" or not product_id.startswith("G25")
                    else "고객_상품_QNA_스냅샷",
                    "question_title": question,
                    "question_text": question,
                    "answer_text": answer_text,
                }
            )
    return rows


def fetch_home_qna(
    session: requests.Session, product: dict[str, str]
) -> dict[str, Any]:
    endpoint = "https://livingapi.lge.co.kr/itemsvc/ajax/v1/pdp/qna-list"
    base_params = {
        "goodsId": product["product_id"],
        "qnaType": "",
        "isMyInquiry": "false",
        "excludeSecret": "true",
        "pageSize": 10,
    }
    first_status, first_payload = request_json(
        session,
        "GET",
        endpoint,
        params={**base_params, "pageNum": 1},
        referer=product["url"],
    )
    first_data = ((first_payload.get("data") or {}).get("qnaListData") or {})
    pages = int(first_data.get("pages") or 0)
    list_rows: list[dict[str, Any]] = list(first_data.get("list") or [])
    statuses = [first_status]
    for page in range(2, pages + 1):
        status, payload = request_json(
            session,
            "GET",
            endpoint,
            params={**base_params, "pageNum": page},
            referer=product["url"],
        )
        statuses.append(status)
        data = ((payload.get("data") or {}).get("qnaListData") or {})
        list_rows.extend(data.get("list") or [])

    records: list[dict[str, Any]] = []
    detail_statuses: list[int] = []
    for row in list_rows:
        if row.get("isSecret") or row.get("isNotice"):
            continue
        inquiry_id = str(row.get("inquiryId") or "")
        detail_status, detail_payload = request_json(
            session,
            "GET",
            f"{endpoint}/{inquiry_id}",
            params={"isNotice": "false"},
            referer=product["url"],
        )
        detail_statuses.append(detail_status)
        detail = detail_payload.get("data") or {}
        records.append(
            {
                "record_id": inquiry_id,
                "source": "HOMESTYLE_QNA_API",
                "scope": "고객_상품_QNA_공개건",
                "qna_type": detail.get("qnaTypeName") or row.get("qnaTypeName") or "",
                "registered_date": detail.get("registeredDate") or row.get("registeredDate") or "",
                "question_title": redact_public_text(
                    detail.get("inquiryTitle") or row.get("inquiryTitle")
                ),
                "question_text": redact_public_text(detail.get("inquiryContent")),
                "answer_text": redact_public_text(
                    detail.get("answerContent") or row.get("answerContent")
                ),
                "answered": bool(
                    detail.get("answerContent") or row.get("answerContent")
                ),
            }
        )

    return {
        "list_endpoint": endpoint,
        "list_method": "GET",
        "detail_endpoint_pattern": endpoint + "/{inquiryId}",
        "detail_method": "GET",
        "http_statuses": sorted(set(statuses)),
        "detail_http_statuses": sorted(set(detail_statuses)),
        "privacy_filter": "excludeSecret=true; isMyInquiry=false; author/files omitted; phone/email masked",
        "reported_public_total": int(first_data.get("total") or 0),
        "pages": pages,
        "retrieved_public_count": len(records),
        "records": records,
    }


def normalize_lge_qna(row: dict[str, Any]) -> dict[str, Any]:
    return {
        "record_id": str(row.get("questionNo") or ""),
        "source": "LGE_APPLIANCE_QNA_API",
        "scope": "고객_모델그룹_QNA_공개건",
        "model_id_in_record": row.get("modelId") or "",
        "option_text": redact_public_text(row.get("modeloptinfostr")),
        "qna_type": row.get("questionTypeName") or row.get("questionTypeCode") or "",
        "registered_date": row.get("creationDate") or "",
        "question_title": redact_public_text(row.get("questionTitle")),
        "question_text": redact_public_text(row.get("questionContent")),
        "answer_text": redact_public_text(row.get("answerContent")),
        "answered": bool(row.get("answerContent") or row.get("answerUseFlag") == "Y"),
    }


def fetch_appliance_sources(
    session: requests.Session, product: dict[str, str]
) -> dict[str, Any]:
    model_id = product["model_id"]
    faq_endpoint = f"https://apiv2.lge.co.kr/itemsvc/ajax/v1/model/faq/{model_id}"
    faq_status, faq_payload = request_json(
        session, "GET", faq_endpoint, referer=product["url"]
    )
    faq_records = []
    for item in faq_payload.get("data") or []:
        if not isinstance(item, dict):
            continue
        faq_records.append(
            {
                "record_id": "FAQ-" + str(item.get("faqDetailId") or item.get("faqId") or ""),
                "source": "LGE_APPLIANCE_FAQ_API",
                "scope": "공식_제품카테고리_FAQ",
                "faq_type": item.get("faqTypeCode") or "",
                "question_title": redact_public_text(
                    item.get("title") or item.get("question")
                ),
                "question_text": redact_public_text(
                    item.get("title") or item.get("question")
                ),
                "answer_text": redact_public_text(
                    item.get("content") or item.get("answer")
                ),
            }
        )

    qna_endpoint = "https://www.lge.co.kr/mkt/api/qna/retrieveQnaList"
    base_body = {
        "questionTypeCode": "ALL",
        "excludePrivate": "Y",
        "myQna": "N",
    }
    qna_statuses: list[int] = []
    qna_rows: list[dict[str, Any]] = []
    reported_total = 0
    total_pages = 0
    page = 1
    while True:
        status, payload = request_json(
            session,
            "POST",
            qna_endpoint,
            params={"modelId": model_id, "page": page},
            data={**base_body, "page": str(page)},
            referer=product["url"],
        )
        qna_statuses.append(status)
        data = payload.get("data") or {}
        if page == 1:
            reported_total = int(data.get("qnaTotalCount") or 0)
            total_pages = int((data.get("pagination") or {}).get("totalCount") or 0)
        qna_rows.extend(data.get("qnaList") or [])
        if page >= total_pages or total_pages == 0:
            break
        page += 1

    public_rows = [
        row
        for row in qna_rows
        if str(row.get("secret") or "N") == "N"
        and str(row.get("blocked") or "N") == "N"
    ]
    qna_records = [normalize_lge_qna(row) for row in public_rows]
    return {
        "faq_api": {
            "endpoint": faq_endpoint,
            "method": "GET",
            "http_status": faq_status,
            "record_count": len(faq_records),
            "records": faq_records,
        },
        "qna_api": {
            "endpoint": qna_endpoint,
            "method": "POST",
            "http_statuses": sorted(set(qna_statuses)),
            "privacy_filter": "excludePrivate=Y; myQna=N; secret/blocked omitted; author/files omitted; phone/email masked",
            "reported_public_total": reported_total,
            "pages": total_pages,
            "retrieved_public_count": len(qna_records),
            "records": qna_records,
        },
    }


def title_key(value: str) -> str:
    return re.sub(r"[^0-9A-Za-z가-힣]", "", value).lower()


def compare_records(
    html_records: list[dict[str, Any]], api_records: list[dict[str, Any]]
) -> dict[str, int]:
    html_titles = {title_key(row.get("question_title", "")) for row in html_records}
    api_titles = {title_key(row.get("question_title", "")) for row in api_records}
    html_titles.discard("")
    api_titles.discard("")
    return {
        "html_unique_titles": len(html_titles),
        "api_unique_titles": len(api_titles),
        "exact_title_overlap": len(html_titles & api_titles),
        "api_only_titles": len(api_titles - html_titles),
        "html_only_titles": len(html_titles - api_titles),
    }


def signal_matches(records: list[dict[str, Any]]) -> dict[str, Any]:
    matches: dict[str, list[dict[str, str]]] = {key: [] for key in SIGNAL_PATTERNS}
    for row in records:
        text = " ".join(
            str(row.get(key) or "")
            for key in ("question_title", "question_text", "answer_text", "option_text")
        )
        for field, pattern in SIGNAL_PATTERNS.items():
            if re.search(pattern, text, flags=re.IGNORECASE | re.DOTALL):
                excerpt = re.sub(r"\s+", " ", text).strip()[:180]
                matches[field].append(
                    {"record_id": str(row.get("record_id") or ""), "excerpt": excerpt}
                )
    return {
        key: {"matched_record_count": len(rows), "samples": rows[:3]}
        for key, rows in matches.items()
        if rows
    }


def main() -> None:
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    session = requests.Session()
    session.verify = VERIFY_TLS
    session.headers.update({"User-Agent": USER_AGENT})

    products_output = []
    endpoint_success = Counter()
    all_scope_counts = Counter()
    all_signal_counts = Counter()
    exact_model_qna_total = 0
    related_model_qna_total = 0

    for product in PRODUCTS:
        product_id = product["product_id"]
        html_records = extract_faq_jsonld(product_id)
        if product["source_system"] == "HOMESTYLE":
            qna_api = fetch_home_qna(session, product)
            api_records = qna_api["records"]
            comparison = compare_records(html_records, api_records)
            api_output: dict[str, Any] = {"qna_api": qna_api}
            endpoint_success["HOMESTYLE_QNA_API_200"] += int(
                qna_api["http_statuses"] == [200]
                and qna_api["detail_http_statuses"] in ([200], [])
            )
        else:
            api_output = fetch_appliance_sources(session, product)
            api_records = api_output["faq_api"]["records"]
            comparison = compare_records(html_records, api_records)
            endpoint_success["LGE_FAQ_API_200"] += int(
                api_output["faq_api"]["http_status"] == 200
            )
            endpoint_success["LGE_QNA_API_200"] += int(
                api_output["qna_api"]["http_statuses"] == [200]
            )

        qna_records = list(api_output["qna_api"]["records"])
        if product["source_system"] == "HOMESTYLE":
            exact_qna_records = qna_records
            related_qna_records: list[dict[str, Any]] = []
            faq_api_records: list[dict[str, Any]] = []
        else:
            exact_qna_records = [
                row
                for row in qna_records
                if row.get("model_id_in_record") == product.get("model_id")
            ]
            related_qna_records = [
                row
                for row in qna_records
                if row.get("model_id_in_record") != product.get("model_id")
            ]
            faq_api_records = list(api_output["faq_api"]["records"])
        exact_model_qna_total += len(exact_qna_records)
        related_model_qna_total += len(related_qna_records)

        all_records = list(html_records)
        if product["source_system"] == "HOMESTYLE":
            all_records.extend(api_output["qna_api"]["records"])
        else:
            all_records.extend(api_output["faq_api"]["records"])
            all_records.extend(api_output["qna_api"]["records"])
        signals = signal_matches(all_records)
        html_signals = signal_matches(html_records)
        faq_api_signals = signal_matches(faq_api_records)
        exact_qna_signals = signal_matches(exact_qna_records)
        related_qna_signals = signal_matches(related_qna_records)
        html_signal_fields = set(html_signals)
        exact_qna_signal_fields = set(exact_qna_signals)
        for row in all_records:
            all_scope_counts[row["scope"]] += 1
        for key, value in signals.items():
            all_signal_counts[key] += value["matched_record_count"]

        products_output.append(
            {
                "product_id": product_id,
                "source_system": product["source_system"],
                "model_id": product.get("model_id", ""),
                "html_faqpage": {
                    "source_file": f"raw_html/{product_id}.html",
                    "record_count": len(html_records),
                    "records": html_records,
                },
                **api_output,
                "html_vs_corresponding_api": comparison,
                "qna_model_scope": {
                    "exact_product_or_goods_count": len(exact_qna_records),
                    "related_model_count": len(related_qna_records),
                    "rule": (
                        "홈스타일은 goodsId 단위라 전부 해당 상품으로 취급. "
                        "가전은 API가 같은 모델 그룹의 형제 modelId도 반환하므로 exact modelId만 상품값 후보로 사용."
                    ),
                },
                "source_signal_fields": {
                    "html_faqpage": sorted(html_signals),
                    "official_faq_api": sorted(faq_api_signals),
                    "qna_api_exact_product_or_goods": sorted(exact_qna_signals),
                    "qna_api_related_models_excluded_from_direct_fill": sorted(
                        related_qna_signals
                    ),
                    "exact_qna_added_vs_html": sorted(
                        exact_qna_signal_fields - html_signal_fields
                    ),
                },
                "requested_field_candidate_signals": signals,
                "exact_qna_candidate_signals": exact_qna_signals,
            }
        )
        api_counts = (
            f"qna={api_output['qna_api']['retrieved_public_count']}"
            if product["source_system"] == "HOMESTYLE"
            else (
                f"faq={api_output['faq_api']['record_count']} "
                f"qna={api_output['qna_api']['retrieved_public_count']}"
            )
        )
        print(f"{product_id}: html_faq={len(html_records)} {api_counts}")

    payload = {
        "metadata": {
            "collected_at": datetime.now(timezone(timedelta(hours=9))).isoformat(),
            "product_count": len(PRODUCTS),
            "purpose": "8개 PDP의 FAQPage 및 추가 공개 FAQ/Q&A 조회 API 파싱 가능성 검증",
            "excel_modified": False,
            "collection_scope": "공개·조회 전용 API. 비공개/내 문의 제외.",
            "pii_policy": "작성자와 첨부파일은 저장하지 않고 전화번호·이메일은 마스킹함.",
            "interpretation_note": (
                "키워드 신호는 요구 필드의 후보 근거이며 확정값이 아니다. "
                "고객 질문은 오기·추측이 있을 수 있으므로 공식 답변이나 상품 원천과 교차검증해야 한다."
            ),
            "tls_note": (
                "공식 lge.co.kr 계열 호스트만 호출했으며, 현재 Python 환경에 사내 프록시 CA가 없어 "
                "검증 실행에서는 requests TLS 검증을 비활성화했다. 운영 환경에서는 사내 CA를 설치해야 한다."
            ),
        },
        "summary": {
            "endpoint_success_product_counts": dict(endpoint_success),
            "record_counts_by_scope_including_html_api_overlap": dict(all_scope_counts),
            "qna_model_scope_counts": {
                "exact_product_or_goods": exact_model_qna_total,
                "related_appliance_models_excluded_from_direct_fill": related_model_qna_total,
            },
            "candidate_signal_record_counts_including_overlap": dict(all_signal_counts),
            "request_2_semantic_search": (
                "FAQ/Q&A 질문-답변 문장을 동의어·자연어 검색용 보조 코퍼스로 활용 가능. "
                "개인화/CRM 데이터로는 사용하지 않음."
            ),
            "direct_fill_rule": (
                "미답변 질문, 형제 모델 Q&A, 카테고리 공통 FAQ는 상품별 확정값으로 자동 입력하지 않는다. "
                "해당 goodsId/modelId의 답변 완료 Q&A도 보조 근거로만 사용하고 공식 상품 API/스펙과 교차검증한다."
            ),
            "impact_against_current_8_product_requested_field_sample": {
                "confirmed_newly_fillable_blank_cells": 0,
                "mapping_rate_change": "없음",
                "reason": (
                    "추가 Q&A는 기존에 채워진 설명서 태그의 근거를 보강하지만, 현재 비어 있는 고객 요청 필드의 "
                    "확정값은 제공하지 않는다. 질문만 있고 답변이 없거나, 카테고리/형제 모델 공통 정보이기 때문이다."
                ),
                "quality_enhancement_examples": [
                    "G25070005743: 배송·취소·기존 소파 내림/회수 정책 보조 근거",
                    "G646GBB031: 정확한 모델의 치수 기준 차이, 구성·기능·에너지등급 비교 근거",
                    "WA2525YMZF: 건조선반 미포함, 급수·설치, 사용·관리 보조 근거",
                    "SQ06GJ1WFS: W837×H308×D189mm 및 화이트 색상 공식 답변 재확인",
                ],
                "still_unfilled_examples": [
                    "G25070001871 조명 W×D×H: 전구/전선 포함 질문은 있으나 답변이 고객센터 문의 안내라 치수 없음",
                    "OLED48C6KNA/SQ06GJ1WFS 벽면부착 추천 높이: 일반 설치·시청거리 FAQ는 있으나 해당 모델 권장 부착 높이 없음",
                    "세트 구성 실제 제품 ID 리스트: FAQ/Q&A의 구성품 문장은 ID 목록을 제공하지 않음",
                    "공간 콘텐츠의 실제 공간크기·공간 내 제품관계·CRM: 상품 Q&A로 확정할 수 없음",
                ],
            },
        },
        "products": products_output,
    }
    OUTPUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"output={OUTPUT}")


if __name__ == "__main__":
    main()
