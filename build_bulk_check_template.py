from __future__ import annotations

import json
import re
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

import build_feasibility_sheet as xlsx_writer


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"m": NS_MAIN, "r": NS_REL}

SOURCE = Path(__file__).with_name("제품군.xlsx")
SOURCE_CACHE = Path(__file__).with_name("poc_full_run") / "source_scope_cache.json"
OUTPUT = Path(__file__).with_name("제품군_대량점검_양식_v6_API_HTML심층분석.xlsx")


def _cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    value_node = cell.find("m:v", NS)
    cell_type = cell.attrib.get("t")
    if cell_type == "s" and value_node is not None:
        return shared_strings[int(value_node.text or "0")]
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.iter(f"{{{NS_MAIN}}}t"))
    return "" if value_node is None else value_node.text or ""


def _read_source() -> tuple[list[dict], list[dict]]:
    if not SOURCE.exists() and SOURCE_CACHE.exists():
        cached = json.loads(SOURCE_CACHE.read_text(encoding="utf-8"))
        return cached["selected"], cached["excluded"]
    with ZipFile(SOURCE) as archive:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            shared_strings = [
                "".join(node.text or "" for node in item.iter(f"{{{NS_MAIN}}}t"))
                for item in root.findall("m:si", NS)
            ]

        styles = ET.fromstring(archive.read("xl/styles.xml"))
        style_fill_ids = [
            int(xf.attrib.get("fillId", "0")) for xf in styles.find("m:cellXfs", NS)
        ]

        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relations = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relation_targets = {item.attrib["Id"]: item.attrib["Target"] for item in relations}

        selected: list[dict] = []
        excluded: list[dict] = []

        for sheet_index, sheet in enumerate(workbook.find("m:sheets", NS), start=1):
            sheet_name = sheet.attrib["name"]
            relation_id = sheet.attrib[f"{{{NS_REL}}}id"]
            target = relation_targets[relation_id]
            target = target if target.startswith("xl/") else f"xl/{target.lstrip('/')}"
            root = ET.fromstring(archive.read(target))

            current: dict[str, str] = {}
            for row in root.findall(".//m:sheetData/m:row", NS):
                row_number = int(row.attrib["r"])
                cells: dict[str, tuple[str, int]] = {}
                for cell in row.findall("m:c", NS):
                    reference = cell.attrib["r"]
                    column = "".join(char for char in reference if char.isalpha())
                    value = _cell_value(cell, shared_strings)
                    style_id = int(cell.attrib.get("s", "0"))
                    fill_id = style_fill_ids[style_id] if style_id < len(style_fill_ids) else 0
                    cells[column] = (value, fill_id)

                if row_number == 1:
                    continue

                if sheet_index == 1:
                    number = cells.get("A", ("", 0))[0]
                    small = cells.get("D", ("", 0))[0]
                    if not number or not small:
                        continue
                    if cells.get("B", ("", 0))[0]:
                        current["large"] = cells["B"][0]
                    if cells.get("C", ("", 0))[0]:
                        current["mid"] = cells["C"][0]
                    # 원본은 러그·매트 구간의 대분류 셀이 비어 있다. 사이트 분류상 패브릭으로 보정한다.
                    if current.get("mid") == "러그·매트":
                        current["large"] = "패브릭"

                    record = {
                        "source_system": "HOMESTYLE",
                        "source_sheet": sheet_name,
                        "source_row": row_number,
                        "source_no": number,
                        "large": current.get("large", ""),
                        "mid": current.get("mid", ""),
                        "small": small,
                        "category_id": cells.get("F", ("", 0))[0],
                        "public_count": _to_int(cells.get("E", ("", 0))[0]),
                        "admin_count": None,
                        "bundle_count": None,
                        "category_fill_id": cells.get("D", ("", 0))[1],
                    }
                else:
                    if row_number >= 96:
                        continue
                    small = cells.get("C", ("", 0))[0]
                    if not small:
                        continue
                    if cells.get("A", ("", 0))[0]:
                        current["large"] = cells["A"][0]
                    if cells.get("B", ("", 0))[0]:
                        current["mid"] = cells["B"][0]

                    record = {
                        "source_system": "LGE_APPLIANCE",
                        "source_sheet": sheet_name,
                        "source_row": row_number,
                        "source_no": "",
                        "large": current.get("large", ""),
                        "mid": current.get("mid", ""),
                        "small": small,
                        "category_id": "",
                        "public_count": _to_int(cells.get("E", ("", 0))[0]),
                        "admin_count": _to_optional_int(cells.get("F", ("", 0))[0]),
                        "bundle_count": _to_optional_int(cells.get("G", ("", 0))[0]),
                        "category_fill_id": cells.get("C", ("", 0))[1],
                    }

                if record["category_fill_id"] == 3:
                    record["exclude_reason"] = "소분류 셀 회색 음영"
                    excluded.append(record)
                else:
                    selected.append(record)

    scope_counters = {"HOMESTYLE": 0, "LGE_APPLIANCE": 0}
    for item in selected:
        scope_counters[item["source_system"]] += 1
        prefix = "HS" if item["source_system"] == "HOMESTYLE" else "AP"
        item["scope_id"] = f"{prefix}-{scope_counters[item['source_system']]:04d}"
        item.update(_derive_profiles(item))
    SOURCE_CACHE.write_text(
        json.dumps({"selected": selected, "excluded": excluded}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return selected, excluded


def _to_int(value: str) -> int:
    return int(float(value)) if value not in ("", None) else 0


def _to_optional_int(value: str) -> int | None:
    return int(float(value)) if value not in ("", None) else None


def _derive_profiles(item: dict) -> dict:
    large = item["large"]
    mid = item["mid"]
    small = item["small"]
    source = item["source_system"]

    if source == "HOMESTYLE":
        if large == "조명":
            parse_profile = "HOME_LIGHTING"
        elif large == "인테리어소품":
            parse_profile = "HOME_DECOR"
        elif large == "패브릭":
            parse_profile = "HOME_TEXTILE"
        else:
            parse_profile = "HOME_FURNITURE"
    else:
        parse_profile = {
            "TV/오디오": "APPLIANCE_DISPLAY_AUDIO",
            "PC/모니터": "APPLIANCE_PC_MONITOR",
            "주방가전": "APPLIANCE_KITCHEN",
            "생활가전": "APPLIANCE_LIVING",
            "에어컨/에어케어": "APPLIANCE_AIR",
        }.get(large, "APPLIANCE_STANDARD")

    combined = f"{mid} {small}"
    set_candidate = any(token in small for token in ("세트", "+", "2in1", "패키지 세트"))
    wall_ceiling_candidate = any(
        token in combined
        for token in ("벽", "팬던트", "샹들리에", "천장", "시스템 에어컨")
    )
    builtin_candidate = source == "LGE_APPLIANCE" and any(
        token in combined for token in ("식기세척기", "광파오븐", "와인셀러", "시스템 에어컨")
    )

    count_for_sampling = (
        (item.get("admin_count") or 0) + (item.get("bundle_count") or 0)
        if source == "LGE_APPLIANCE" and ((item.get("admin_count") or 0) + (item.get("bundle_count") or 0)) > 0
        else item.get("public_count", 0)
    )
    if count_for_sampling == 0:
        sample_count = 0
    elif count_for_sampling <= 20:
        sample_count = count_for_sampling
    elif count_for_sampling <= 200:
        sample_count = 10
    elif count_for_sampling <= 1000:
        sample_count = 20
    else:
        sample_count = 30

    required_profile = "REQ_3D_APPLIANCE" if source == "LGE_APPLIANCE" else "REQ_3D_HOME"
    if set_candidate:
        required_profile += "|REQ_SET_COMPONENT"
    if wall_ceiling_candidate:
        required_profile += "|REQ_WALL_CEILING"

    return {
        "parse_profile": parse_profile,
        "required_profile": required_profile,
        "set_candidate_yn": "Y" if set_candidate else "N",
        "builtin_candidate_yn": "Y" if builtin_candidate else "N",
        "wall_ceiling_candidate_yn": "Y" if wall_ceiling_candidate else "N",
        "zero_count_yn": "Y" if item.get("public_count", 0) == 0 else "N",
        "poc_sample_count": sample_count,
    }


def _scope_rows(selected: list[dict]) -> list[list]:
    rows = [[
        "scope_id", "포함여부", "source_system", "원본시트", "원본행", "원본No",
        "대분류", "중분류", "소분류", "category_id", "공개 판매중 수",
        "admin 판매중 수", "admin 묶음 수", "parse_profile", "required_profile",
        "요청1_대상", "요청1_requirement_ids", "요청2_대상", "요청2_requirement_ids",
        "세트후보", "빌트인후보", "벽/천장후보", "0건여부", "1차 PoC 샘플수",
        "전체진행여부", "담당자", "비고",
    ]]
    for item in selected:
        request1_ids = ["R1-01", "R1-02", "R1-03", "R1-04", "R1-05", "R1-06", "R1-09"]
        if item["source_system"] == "LGE_APPLIANCE":
            request1_ids.append("R1-07")
        if item["set_candidate_yn"] == "Y":
            request1_ids.append("R1-08")
        if item["wall_ceiling_candidate_yn"] == "Y":
            request1_ids.append("R1-10")
        rows.append([
            item["scope_id"], "Y", item["source_system"], item["source_sheet"], item["source_row"], item["source_no"],
            item["large"], item["mid"], item["small"], item["category_id"], item["public_count"],
            item["admin_count"], item["bundle_count"], item["parse_profile"], item["required_profile"],
            "Y", "|".join(request1_ids), "Y", "R2-01|R2-02|R2-03|R2-04|R2-05|R2-06|R2-07",
            item["set_candidate_yn"], item["builtin_candidate_yn"], item["wall_ceiling_candidate_yn"],
            item["zero_count_yn"], item["poc_sample_count"], "N", "", "",
        ])
    return rows


def _excluded_rows(excluded: list[dict]) -> list[list]:
    rows = [[
        "source_system", "원본시트", "원본행", "원본No", "대분류", "중분류", "소분류",
        "category_id", "공개 판매중 수", "admin 판매중 수", "admin 묶음 수", "제외사유",
    ]]
    for item in excluded:
        rows.append([
            item["source_system"], item["source_sheet"], item["source_row"], item["source_no"],
            item["large"], item["mid"], item["small"], item["category_id"], item["public_count"],
            item["admin_count"], item["bundle_count"], item["exclude_reason"],
        ])
    return rows


def build_template() -> None:
    selected, excluded = _read_source()
    homestyle = [item for item in selected if item["source_system"] == "HOMESTYLE"]
    appliance = [item for item in selected if item["source_system"] == "LGE_APPLIANCE"]
    run_dir = Path(__file__).with_name("poc_full_run")
    deep_source = json.loads((run_dir / "deep_source_inventory.json").read_text(encoding="utf-8"))
    deep_product_by_id = {row["product_id"]: row for row in deep_source["products"]}
    deep_stage_by_name = {row["stage"]: row for row in deep_source["stage_statistics"]}
    deep_api = deep_stage_by_name["API_ONLY"]
    deep_html = deep_stage_by_name["API_HTML"]
    deep_paddle = deep_stage_by_name["API_HTML_PADDLE_OCR"]

    def groups(headers: list[str], *, r1=(), r2=(), both=()) -> list[str]:
        r1, r2, both = set(r1), set(r2), set(both)
        return ["BOTH" if h in both else "R1" if h in r1 else "R2" if h in r2 else "COMMON" for h in headers]

    def widths(headers: list[str]) -> list[int]:
        result = []
        for header in headers:
            if "url" in header.lower():
                result.append(72)
            elif "json" in header.lower() or "text" in header.lower() or "fields" in header.lower():
                result.append(58)
            elif "requirement_ids" in header:
                result.append(58)
            elif "message" in header.lower() or "note" in header.lower() or "설명" in header:
                result.append(48)
            else:
                result.append(max(13, min(34, len(header) * 2 + 6)))
        return result

    def compact(value, limit: int = 5000) -> str:
        if value is None:
            return ""
        if isinstance(value, (dict, list)):
            text = json.dumps(value, ensure_ascii=False, separators=(",", ":"))
        else:
            text = str(value)
        return text if len(text) <= limit else text[:limit] + "…"

    guide_rows = [
        ["항목", "내용"],
        ["목적", "음영 없는 제품군만 대상으로 상품 URL을 적재하고, 요청 1·2의 반영 여부와 API/HTML/OCR/VISION 근거를 함께 관리"],
        ["원본 파일", SOURCE.name],
        ["요청 1 색상", "파란색: 3D Asset 제작 기본정보 전용 필드"],
        ["요청 2 색상", "초록색: 제품 추천 분석 전용 필드"],
        ["공통 색상", "보라색: 요청 1과 요청 2가 함께 사용하는 상품 식별·기본정보 필드"],
        ["대상 제품군", f"총 {len(selected)}개: 홈스타일 {len(homestyle)}개 + 가전 {len(appliance)}개"],
        ["대표 테스트", "홈스타일 4프로필 + 가전 4프로필에서 1상품씩, 총 8상품을 실제 공개 PDP로 점검"],
        ["방법 실행", "8상품 모두 API 호출, 원본 HTML 저장, OCR 실행. API는 가전 4상품에서 HTTP 200이지만 핵심 제품필드가 제한됨"],
        ["OCR 실행", "동일 원본 이미지 62장에 Windows OCR(한/영 124회), PaddleOCR 62회, Tesseract.js 62회를 모두 실행"],
        ["OCR 비교 결과", "PaddleOCR 숫자 28/28·문구 13/13, Windows OCR 27/28·10/13, Tesseract.js 7/28·9/13"],
        ["OCR 적용 권장", "API/HTML 우선, PaddleOCR은 상세·규격 이미지에 선택 적용, Windows OCR은 빠른 1차 탐색에 사용"],
        ["OCR 평가 주의", "전체 전문 CER가 아니라 이미지/API로 교차확인한 핵심 숫자 28개·문구 13개의 정확 회수율"],
        ["심층 파싱 범위", f"대표 8상품의 원본 API/HTML을 37개 정규화 필드로 전수 매핑: 총 {deep_source['metadata']['field_cell_count']}개 필드 셀"],
        ["심층 파싱 개선", f"사용 가능 필드 셀: API 단독 {deep_api['usable']}/{deep_api['total_field_cells']}({deep_api['usable_coverage_pct']}%) → API+HTML {deep_html['usable']}/{deep_html['total_field_cells']}({deep_html['usable_coverage_pct']}%), +{deep_html['usable_increment_vs_api']}칸"],
        ["OCR의 추가 효과", f"API+HTML 뒤 PaddleOCR 추가 시 {deep_paddle['usable']}/{deep_paddle['total_field_cells']}({deep_paddle['usable_coverage_pct']}%), HTML 단계 대비 +{deep_paddle['usable'] - deep_html['usable']}칸. OCR 품질 개선은 주로 신뢰도 향상"],
        ["심층 결과 확인", "00_통계_소스증분 → 00_소스증분_상품별 → 00_통계_추가요구 → 04_상품확장RESULT → 05_추가필드_EVIDENCE → 11_RAW_FIELD_INVENTORY 순서로 확인"],
        ["방법 확인", "00_METHOD_EXECUTION에서 상품×방법 40행, 00_OCR엔진_요약·00_OCR정답_검증·05_OCR엔진_이미지에서 비교"],
        ["통계 산식", "엄격가능률=가능/(가능+조건부+불가능), 조건포함가능률=(가능+조건부)/(가능+조건부+불가능); 해당없음은 분모 제외"],
        ["제외 제품군", f"총 {len(excluded)}개: 소분류 셀 회색 음영 기준"],
        ["수량 참고", f"홈스타일 공개 판매중 합계 {sum(x['public_count'] for x in homestyle):,}; 가전 공개 판매중 합계 {sum(x['public_count'] for x in appliance):,}"],
        ["가전 admin 참고", f"선택 제품군 admin 판매중 {sum((x['admin_count'] or 0) for x in appliance):,}; admin 묶음 {sum((x['bundle_count'] or 0) for x in appliance):,}"],
        ["수량 주의", "공개 판매중 수는 대표/옵션 중복 가능성이 있어 최종 URL 수와 다를 수 있음. product_id 기준 중복 제거 필요"],
        ["권장 배치 단위", "1개 batch 최대 500 URL. source_system·parse_profile별로 분리"],
        ["진행 순서 1", "00_요청사항_COVERAGE에서 17개 요구사항의 양식 반영 여부와 실제 제공 판정을 확인"],
        ["진행 순서 2", "01_SCOPE_대상에서 요청별 적용 requirement_id와 PoC 샘플수를 검토"],
        ["진행 순서 3", "03_상품URL_INPUT에 product_id와 source_url을 1상품 1행으로 적재"],
        ["진행 순서 4", "04_상품RESULT는 상품 1행 요약, 05_FIELD_EVIDENCE는 필드 1행 근거로 저장"],
        ["검수 기준", "confidence<0.85, 필수값 MISSING, 세트 ID 없음, 파싱 ERROR는 07_QA_BATCH에서 요청별로 집계"],
        ["CRM 주의", "R2-06은 양식에는 포함되지만 공개 PDP에서 취득 불가. 내부 CRM과 동의·목적 기준으로만 결합"],
    ]
    guide_row_groups = [{"요청 1 색상": "R1", "요청 2 색상": "R2", "공통 색상": "BOTH"}.get(row[0], "COMMON") for row in guide_rows[1:]]

    requirement_headers = ["requirement_id", "요청구분", "구분/번호", "항목", "세부설명", "양식반영여부", "관련시트", "관련필드", "데이터제공판정", "2026년 8월 가능성", "판정근거", "셀색상"]
    requirement_data = [
        ["R1-01", "요청 1", "필수", "PDP 이미지", "정면45도 또는 정면 이미지", "Y", "04_상품RESULT", "pdp_image_url|image_view", "O", "8/7 PoC", "API 이미지 목록+VISION 방향 판정", "파랑"],
        ["R1-02", "요청 1", "필수", "사이즈 정보", "W×D×H", "Y", "04_상품RESULT|05_FIELD_EVIDENCE", "width_mm|depth_mm|height_mm|dimension_components_json", "O", "8/7 PoC", "API 우선, OCR 보강", "파랑/공통"],
        ["R1-03", "요청 1", "필수", "분류", "중카테고리, 소카테고리", "Y", "01_SCOPE_대상|04_상품RESULT", "mid_category|small_category", "O", "8/7 PoC", "API 명시값", "파랑/공통"],
        ["R1-04", "요청 1", "필수", "배치 추천 공간 리스트", "리빙룸, 베드룸 등", "Y", "04_상품RESULT", "recommended_spaces", "O(조건부)", "8/14 PoC", "공간 API 우선, 카테고리 규칙 보조", "파랑/공통"],
        ["R1-05", "요청 1", "필수", "브랜드명", "", "Y", "04_상품RESULT", "brand_id|brand_name", "O", "8/7 PoC", "API 명시값", "파랑/공통"],
        ["R1-06", "요청 1", "필수", "제품 색상", "", "Y", "04_상품RESULT", "colors_raw|colors_normalized", "O", "8/7 PoC", "옵션/고시 우선, 이미지 보조", "파랑/공통"],
        ["R1-07", "요청 1", "가전 추가 필수", "설치 타입 구분", "빌트인/스탠딩", "Y", "04_상품RESULT", "installation_type", "O(조건부)", "8/14 PoC", "가전 명시값 또는 카테고리 규칙", "파랑"],
        ["R1-08", "요청 1", "세트상품 필수", "세트 구성 ID 리스트", "실제 제품 ID 리스트", "Y", "04_상품RESULT", "set_component_product_ids", "O(원천 ID 있을 때)", "조건부", "원천 패키지/API가 없으면 MISSING", "파랑"],
        ["R1-09", "요청 1", "옵션", "배치 가능 위치", "벽, 천장, 바닥 등", "Y", "04_상품RESULT", "placement_positions", "O(조건부)", "8/14 PoC", "카테고리 규칙+VISION", "파랑"],
        ["R1-10", "요청 1", "옵션", "벽면부착 추천 높이", "", "Y", "04_상품RESULT", "wall_mount_height_mm", "O(매뉴얼 명시값 있을 때)", "조건부", "안전상 이미지 추정값은 사용하지 않음", "파랑"],
        ["R2-01", "요청 2", "1", "가구 상품 설명서 자동 태깅", "기본/제조/재질/규격/구성품/색상/사용목적/조립/안전/취급주의/품질보증/인증/판매정보", "Y", "04_상품RESULT", "tag_basic~tag_sales 13개 필드", "O", "8/14 PoC", "API/HTML/OCR 태깅", "초록"],
        ["R2-02", "요청 2", "2", "디자인 스타일 추론", "미니멀/직선형/곡선형/유기적/장식형/클래식/모듈형/스칸디나비안/컨템포러리 등", "Y", "04_상품RESULT", "design_style_tags", "O(추론)", "8/21 PoC", "멀티라벨+confidence", "초록"],
        ["R2-03", "요청 2", "3", "공간 콘텐츠 자동태깅", "스타일/Mood/공간명/목적/색상톤/크기/포함제품/배치사유·위치", "Y", "06_SPACE_RELATION", "space_*|room_*|product_id|position_*", "O(조건부)", "8/21 PoC", "공간 콘텐츠 URL 필요", "초록"],
        ["R2-04", "요청 2", "4", "공간스타일 ↔ 제품 관계정보", "", "Y", "06_SPACE_RELATION", "space_content_id|product_id|relation_type", "O", "8/21 PoC", "명시 관계 우선, 유사도 보조", "초록"],
        ["R2-05", "요청 2", "5", "공간 내 제품 간 관계정보", "", "Y", "06_SPACE_RELATION", "product_id|related_product_id|relation_type", "O(추론)", "8/31 PoC", "좌표/이미지 분석", "초록"],
        ["R2-06", "요청 2", "6", "개인별 구매/선호/보유/CRM 정보", "", "Y", "08_CRM_INPUT", "customer_key_hash|purchase|preference|owned|crm_segment", "X(PDP)/O(내부연계)", "외부 데이터 필요", "적법한 CRM 원천과 동의·목적 기준 필요", "초록"],
        ["R2-07", "요청 2", "7", "비표준화 데이터 의미기반 검색·추천", "적색/레드/붉은색 등 대응", "Y", "09_SEMANTIC_SEARCH", "search_document|normalized_terms|synonyms|embedding_id", "O", "8/31 PoC", "용어사전+임베딩/하이브리드 인덱스", "초록"],
    ]
    requirement_rows = [requirement_headers] + requirement_data
    requirement_row_groups = ["R1"] * 10 + ["R2"] * 7

    # 2026-07-20 대표 카테고리 실측 PoC: 홈스타일 4개 + 가전 4개.
    # 홈스타일은 공개 PDP API/HTML/상세이미지, 가전은 공식 LGE PDP의 스펙/대체텍스트를 근거로 작성한다.
    requirement_ids = [row[0] for row in requirement_data]
    requirement_name = {row[0]: row[3] for row in requirement_data}
    scope_lookup = {(item["source_system"], item["small"]): item for item in selected}

    def scope_value(source_system: str, small: str, key: str) -> str:
        return scope_lookup.get((source_system, small), {}).get(key, "")

    test_products = [
        {
            "profile": "HOME_FURNITURE", "source_system": "HOMESTYLE", "small": "일반소파",
            "product_id": "G25070005743", "model_id": "", "name": "벨로씨 레보르 천연면피 가죽 소파 4인＋스툴",
            "brand": "까사미아", "url": "https://homestyle.lge.co.kr/item?productId=G25070005743",
            "image": "https://static-store.lge.co.kr/goods/org/073251020000078073.jpg?aw=1&ah=1&rw=800&rh=800",
            "width": 2910, "depth": 1020, "height": 910,
            "dimension_json": '{"sofa":{"w":2910,"d":1020,"h":910},"stool":{"w":740,"d":660,"h":410}}',
            "dimension_note": "고시정보 명시값; 4인 소파와 스툴 규격 분리", "spaces": "LIVING_ROOM",
            "colors_raw": "카멜브라운|크림아이보리", "colors_normalized": "BROWN_WARM|IVORY_OFFWHITE",
            "installation": "FLOOR_STANDING", "set_yn": "Y", "set_components": "", "placement": "FLOOR",
            "tag_basic": "가구|소파|4인용|스툴세트", "tag_material": "천연면피소가죽|합성가죽|E0합판|HR폼|스틸",
            "tag_spec": "SOFA_2910x1020x910|STOOL_740x660x410", "tag_components": "4인소파|스툴",
            "tag_use": "REST|TV|GUEST", "style": "CONTEMPORARY|MODERN|RECTILINEAR",
            "semantic": "웜브라운 또는 크림아이보리 천연가죽 4인용 거실 소파와 스툴",
        },
        {
            "profile": "HOME_TEXTILE", "source_system": "HOMESTYLE", "small": "러그",
            "product_id": "G25100020496", "model_id": "", "name": "모듈로 먼지없는 거실러그 _3 Size",
            "brand": "데이드리머", "url": "https://homestyle.lge.co.kr/item?productId=G25100020496",
            "image": "https://static-store.lge.co.kr/goods/org/916/251021000084916.jpg?aw=1&ah=1&rw=800&rh=800",
            "width": "1400|1600|2000", "depth": "2000|2300|3000", "height": "",
            "dimension_json": '{"options":["1400x2000","1600x2300","2000x3000"],"height_mm":null}',
            "dimension_note": "옵션에서 W×D 확보; 러그 두께/H는 PDP에서 미확보", "spaces": "LIVING_ROOM|BEDROOM",
            "colors_raw": "아이보리|라이트블루|브라운(이미지 추론)", "colors_normalized": "IVORY|LIGHT_BLUE|BROWN",
            "installation": "FLOOR_LAY", "set_yn": "N", "set_components": "", "placement": "FLOOR",
            "tag_basic": "패브릭|러그|3사이즈", "tag_material": "PDP 고시값 없음; 상세이미지 OCR 필요",
            "tag_spec": "1400x2000|1600x2300|2000x3000", "tag_components": "러그 1매",
            "tag_use": "LIVING_ROOM_FLOOR|BEDROOM_FLOOR", "style": "GEOMETRIC|MODERN|CONTEMPORARY",
            "semantic": "아이보리 바탕에 라이트블루와 브라운 사각 패턴의 거실 러그",
        },
        {
            "profile": "HOME_LIGHTING", "source_system": "HOMESTYLE", "small": "팬던트조명",
            "product_id": "G25070001871", "model_id": "Teti", "name": "Teti _5colors",
            "brand": "아르떼미데", "url": "https://homestyle.lge.co.kr/item?productId=G25070001871",
            "image": "https://static-store.lge.co.kr/goods/org/897/250723000005897.png?aw=1&ah=1&rw=800&rh=800",
            "width": "", "depth": "", "height": "", "dimension_json": "",
            "dimension_note": "상품고시와 상세이미지 모두 크기 수치 없음", "spaces": "LIVING_ROOM|DINING_ROOM|ENTRY",
            "colors_raw": "Anthracite grey|White|Orange|Transparent|Transparent Orange",
            "colors_normalized": "DARK_GRAY|WHITE|ORANGE|TRANSPARENT|TRANSPARENT_ORANGE",
            "installation": "CEILING_MOUNTED", "set_yn": "N", "set_components": "", "placement": "CEILING",
            "tag_basic": "조명|팬던트조명|Teti", "tag_material": "PDP 명시값 없음", "tag_spec": "PDP 규격 미제공",
            "tag_components": "조명 본체", "tag_use": "AMBIENT_LIGHTING", "style": "MINIMAL|MODERN|SCULPTURAL",
            "semantic": "오렌지 화이트 투명 다크그레이 옵션의 미니멀 천장 조명",
        },
        {
            "profile": "HOME_DECOR", "source_system": "HOMESTYLE", "small": "액자",
            "product_id": "G25070006112", "model_id": "", "name": "마르코 사진액자 3×3인치_골드",
            "brand": "까사미아", "url": "https://homestyle.lge.co.kr/item?productId=G25070006112",
            "image": "https://static-store.lge.co.kr/goods/org/777250729000030777.jpg?aw=1&ah=1&rw=800&rh=800",
            "width": "", "depth": "", "height": "", "dimension_json": '{"photo_opening_inch":"3x3","outer_wdh":null}',
            "dimension_note": "3×3인치는 사진 규격이며 외형 W×D×H는 PDP에서 미확보", "spaces": "LIVING_ROOM|BEDROOM|ENTRY",
            "colors_raw": "골드", "colors_normalized": "GOLD", "installation": "TABLETOP",
            "set_yn": "N", "set_components": "", "placement": "SHELF|CONSOLE|TABLETOP",
            "tag_basic": "인테리어소품|사진액자|3x3인치", "tag_material": "PDP 명시값 없음",
            "tag_spec": "PHOTO_OPENING_3x3_INCH", "tag_components": "액자 1개", "tag_use": "PHOTO_DISPLAY",
            "style": "CLASSIC|DECORATIVE|VINTAGE", "semantic": "골드 라탄 디테일의 클래식 탁상용 사진액자",
        },
        {
            "profile": "APPLIANCE_DISPLAY_AUDIO", "source_system": "LGE_APPLIANCE", "small": "올레드",
            "product_id": "OLED48C6KNA", "model_id": "OLED48C6KNA", "name": "LG 올레드 evo AI (벽걸이형) 120cm",
            "brand": "LG전자", "url": "https://www.lge.co.kr/tvs/oled48c6kna-wall",
            "image": "https://www.lge.co.kr/kr/images/tvs/md10770851/gallery/medium-interior01.jpg",
            "width": 1071, "depth": 46.9, "height": 620, "dimension_json": '{"without_stand":{"w":1071,"d":46.9,"h":620},"stand_depth":230}',
            "dimension_note": "공식 제품 상세 스펙; 스탠드 미포함 본체 기준", "spaces": "LIVING_ROOM|BEDROOM",
            "colors_raw": "블랙", "colors_normalized": "BLACK", "installation": "WALL_MOUNTED",
            "set_yn": "N", "set_components": "", "placement": "WALL", "tag_basic": "가전|TV|OLED|120cm|4K",
            "tag_material": "", "tag_spec": "1071x46.9x620|4K_UHD", "tag_components": "TV|OLW480A 벽걸이 부품",
            "tag_use": "MEDIA_VIEWING", "style": "MINIMAL|MODERN|RECTILINEAR",
            "semantic": "120cm 4K OLED evo AI 벽걸이형 거실 TV",
        },
        {
            "profile": "APPLIANCE_KITCHEN", "source_system": "LGE_APPLIANCE", "small": "상냉장/하냉동",
            "product_id": "G646GBB031", "model_id": "G646GBB031", "name": "LG 디오스 AI 오브제컬렉션 냉장고 Fit & Max 650L",
            "brand": "LG전자", "url": "https://www.lge.co.kr/refrigerators/g646gbb031",
            "image": "https://www.lge.co.kr/kr/images/refrigerators/md10780848/gallery/medium-interior01.jpg",
            "width": 914, "depth": 709, "height": 1860, "dimension_json": '{"w":914,"d":709,"h":1860,"rear_handle_excluded_d":698}',
            "dimension_note": "공식 제품 상세 스펙", "spaces": "KITCHEN", "colors_raw": "베이지/베이지",
            "colors_normalized": "BEIGE|BEIGE", "installation": "FIT_AND_MAX_FLOOR_STANDING",
            "set_yn": "N", "set_components": "", "placement": "FLOOR", "tag_basic": "가전|냉장고|650L|4도어|상냉장하냉동",
            "tag_material": "미스트 글라스", "tag_spec": "914x709x1860|650L", "tag_components": "냉장실|냉동실",
            "tag_use": "FOOD_STORAGE", "style": "MINIMAL|MODERN|CONTEMPORARY",
            "semantic": "베이지 미스트 글라스 650L 상냉장 하냉동 Fit & Max 냉장고",
        },
        {
            "profile": "APPLIANCE_LIVING", "source_system": "LGE_APPLIANCE", "small": "워시타워",
            "product_id": "WA2525YMZF", "model_id": "WA2525YMZF", "name": "LG 트롬 AI 오브제컬렉션 워시타워",
            "brand": "LG전자", "url": "https://www.lge.co.kr/wash-tower/wa2525ymzf",
            "image": "https://www.lge.co.kr/kr/images/wash-tower/md10576829/gallery/medium-interior01.jpg",
            "width": 700, "depth": 830, "height": 1890, "dimension_json": '{"w":700,"d":830,"h":1890,"door_open_d":1410}',
            "dimension_note": "공식 제품 상세 스펙", "spaces": "LAUNDRY_ROOM|UTILITY_ROOM",
            "colors_raw": "네이처 네이비/네이처 크림 그레이", "colors_normalized": "NAVY|CREAM_GRAY",
            "installation": "FLOOR_STANDING_STACKED", "set_yn": "N", "set_components": "", "placement": "FLOOR",
            "tag_basic": "가전|워시타워|세탁25kg|건조25kg", "tag_material": "", "tag_spec": "700x830x1890|WASH25KG|DRY25KG",
            "tag_components": "세탁기능|건조기능|통합조작부", "tag_use": "LAUNDRY|DRYING",
            "style": "MINIMAL|MODERN|VERTICAL", "semantic": "네이비와 크림그레이 조합 25kg 세탁 건조 일체형 워시타워",
        },
        {
            "profile": "APPLIANCE_AIR", "source_system": "LGE_APPLIANCE", "small": "벽걸이형",
            "product_id": "SQ06GJ1WFS", "model_id": "SQ06GJ1WFS", "name": "LG 휘센 벽걸이에어컨 1등급",
            "brand": "LG전자", "url": "https://www.lge.co.kr/air-conditioners/sq06gj1wfs",
            "image": "https://www.lge.co.kr/kr/images/air-conditioners/md10766829/gallery/medium-interior01.jpg",
            "width": 837, "depth": 189, "height": 308, "dimension_json": '{"indoor_unit":{"w":837,"d":189,"h":308}}',
            "dimension_note": "공식 상세 이미지 대체텍스트의 폭·두께·높이", "spaces": "LIVING_ROOM|BEDROOM|STUDY",
            "colors_raw": "화이트", "colors_normalized": "WHITE", "installation": "WALL_MOUNTED_PRO_INSTALL",
            "set_yn": "N", "set_components": "", "placement": "WALL", "tag_basic": "가전|벽걸이에어컨|18.7m²|1등급",
            "tag_material": "", "tag_spec": "837x189x308|18.7M2", "tag_components": "실내기 SQ06GJ1WFN|실외기",
            "tag_use": "COOLING|DEHUMIDIFICATION", "style": "MINIMAL|MODERN|RECTILINEAR",
            "semantic": "화이트 1등급 18.7제곱미터 벽걸이형 에어컨",
        },
    ]

    for product in test_products:
        scope = scope_lookup.get((product["source_system"], product["small"]), {})
        product.update({
            "scope_id": scope.get("scope_id", ""), "large": scope.get("large", ""),
            "mid": scope.get("mid", ""), "category_id": scope.get("category_id", ""),
        })

    # requirement_data 순서(R1 10개 + R2 7개)에 맞춘 상품별 판정.
    status_matrix = {
        "G25070005743": ["가능", "가능", "가능", "조건부", "가능", "가능", "해당없음", "불가능", "가능", "해당없음", "가능", "가능", "가능", "가능", "조건부", "불가능", "가능"],
        "G25100020496": ["가능", "조건부", "가능", "조건부", "가능", "조건부", "해당없음", "해당없음", "가능", "해당없음", "조건부", "가능", "가능", "가능", "조건부", "불가능", "가능"],
        "G25070001871": ["가능", "불가능", "가능", "조건부", "가능", "가능", "해당없음", "해당없음", "가능", "해당없음", "조건부", "가능", "조건부", "조건부", "조건부", "불가능", "가능"],
        "G25070006112": ["가능", "불가능", "가능", "조건부", "가능", "가능", "해당없음", "해당없음", "가능", "해당없음", "조건부", "가능", "가능", "가능", "조건부", "불가능", "가능"],
        "OLED48C6KNA": ["가능", "가능", "가능", "조건부", "가능", "가능", "가능", "해당없음", "가능", "불가능", "해당없음", "가능", "가능", "가능", "조건부", "불가능", "가능"],
        "G646GBB031": ["가능", "가능", "가능", "조건부", "가능", "가능", "가능", "해당없음", "가능", "해당없음", "해당없음", "가능", "가능", "가능", "조건부", "불가능", "가능"],
        "WA2525YMZF": ["가능", "가능", "가능", "조건부", "가능", "가능", "가능", "해당없음", "가능", "해당없음", "해당없음", "가능", "가능", "가능", "조건부", "불가능", "가능"],
        "SQ06GJ1WFS": ["가능", "가능", "가능", "조건부", "가능", "가능", "가능", "해당없음", "가능", "불가능", "해당없음", "가능", "가능", "가능", "조건부", "불가능", "가능"],
    }

    source_results = json.loads((run_dir / "source_results.json").read_text(encoding="utf-8"))
    image_manifest = json.loads((run_dir / "image_manifest.json").read_text(encoding="utf-8"))
    ocr_ko_rows = json.loads((run_dir / "ocr_ko.json").read_text(encoding="utf-8-sig"))
    ocr_en_rows = json.loads((run_dir / "ocr_en.json").read_text(encoding="utf-8-sig"))
    ocr_benchmark = json.loads((run_dir / "ocr_engine_benchmark.json").read_text(encoding="utf-8"))
    source_by_product = {row["product_id"]: row for row in source_results}
    manifest_by_file = {row["file"]: row for row in image_manifest if row.get("file")}
    ocr_ko_by_file = {row["file"]: row for row in ocr_ko_rows}
    ocr_en_by_file = {row["file"]: row for row in ocr_en_rows}
    benchmark_summary_by_engine = {row["engine"]: row for row in ocr_benchmark["summary"]}
    benchmark_images_by_engine_product: dict[tuple[str, str], list[dict]] = {}
    for row in ocr_benchmark["image_details"]:
        benchmark_images_by_engine_product.setdefault((row["engine"], row["product_id"]), []).append(row)
    benchmark_tokens_by_engine_product: dict[tuple[str, str], list[dict]] = {}
    for row in ocr_benchmark["token_details"]:
        benchmark_tokens_by_engine_product.setdefault((row["engine"], row["product_id"]), []).append(row)
    benchmark_phrases_by_engine_product: dict[tuple[str, str], list[dict]] = {}
    for row in ocr_benchmark["phrase_details"]:
        benchmark_phrases_by_engine_product.setdefault((row["engine"], row["product_id"]), []).append(row)
    dimension_pattern = re.compile(r"\b\d{2,4}(?:[,.]\d+)?\s*(?:mm|cm|x|X|×)\s*\d{0,4}", re.IGNORECASE)

    ocr_summary_by_product: dict[str, dict] = {}
    for product in test_products:
        product_id = product["product_id"]
        files = [name for name in manifest_by_file if name.startswith(product_id + "__")]
        ko = [ocr_ko_by_file[name] for name in files if name in ocr_ko_by_file]
        en = [ocr_en_by_file[name] for name in files if name in ocr_en_by_file]
        merged_text = "\n".join([row.get("text", "") for row in ko + en])
        ocr_summary_by_product[product_id] = {
            "downloaded_images": len(files),
            "ko_success": sum(row.get("status") == "SUCCESS" for row in ko),
            "en_success": sum(row.get("status") == "SUCCESS" for row in en),
            "ko_nonempty": sum(bool(row.get("text", "").strip()) for row in ko),
            "en_nonempty": sum(bool(row.get("text", "").strip()) for row in en),
            "ko_characters": sum(int(row.get("character_count") or 0) for row in ko),
            "en_characters": sum(int(row.get("character_count") or 0) for row in en),
            "dimension_signals": sorted(set(dimension_pattern.findall(merged_text)))[:20],
            "merged_text": merged_text,
        }

    def assessment_reason(product: dict, requirement_id: str, status: str) -> str:
        if status == "해당없음":
            return "해당 상품 유형에 적용되지 않는 요구사항"
        if requirement_id == "R1-01":
            return "대표 이미지 URL 확보; 정면/45도 여부는 VISION 판정 가능"
        if requirement_id == "R1-02":
            return product["dimension_note"]
        if requirement_id == "R1-03":
            return f"카테고리 명시값: {product['mid']} > {product['small']}"
        if requirement_id == "R1-04":
            return f"명시 공간목록 없음; 카테고리와 연출 이미지로 {product['spaces']} 규칙 추론"
        if requirement_id == "R1-05":
            return f"브랜드 명시값: {product['brand']}"
        if requirement_id == "R1-06":
            return f"옵션/상품명/이미지에서 색상 확보: {product['colors_raw']}"
        if requirement_id == "R1-07":
            return f"공식 카테고리·설치 안내 기반: {product['installation']}"
        if requirement_id == "R1-08":
            return "소파+스툴 세트이나 packages API 404로 실제 구성 product_id 미제공"
        if requirement_id == "R1-09":
            return f"카테고리·설치 안내 기반 배치 위치: {product['placement']}"
        if requirement_id == "R1-10":
            return "벽걸이 제품이나 권장 설치 높이 수치 없음; 안전상 이미지 추정값은 사용하지 않음"
        if requirement_id == "R2-01":
            if product["product_id"] == "G25070005743":
                return "고시정보와 HTML에서 제조·재질·규격·구성품·색상·주의·보증 태깅 가능"
            return "API·HTML·OCR을 병합했지만 일부 제조/재질/규격 필드가 비어 있어 추가 원천 또는 검수 필요"
        if requirement_id == "R2-02":
            return f"대표/연출 이미지 VISION 다중라벨 추론: {product['style']}"
        if requirement_id == "R2-03":
            return "PDP 연출 이미지의 공간명·Mood·색상톤·포함 객체를 VISION으로 태깅"
        if requirement_id == "R2-04":
            return "상품 PDP의 연출 이미지이므로 해당 상품과 추론 공간스타일의 관계키 생성 가능"
        if requirement_id == "R2-05":
            return "동시 등장 객체의 유형·위치 관계는 추론 가능하나 다른 상품의 실제 product_id는 없음"
        if requirement_id == "R2-06":
            return "공개 PDP에는 개인 구매·선호·보유·CRM 데이터가 없음; 내부 동의 데이터 연계 필요"
        if requirement_id == "R2-07":
            return "상품명·카테고리·색상·소재·스타일 태그를 검색문서와 동의어로 정규화 가능"
        return status

    method_by_requirement = {
        "R1-01": "API/HTML+VISION", "R1-02": "API/HTML/OCR", "R1-03": "API/HTML",
        "R1-04": "RULE+VISION", "R1-05": "API/HTML", "R1-06": "API/HTML+VISION",
        "R1-07": "HTML+RULE", "R1-08": "PACKAGE_API", "R1-09": "HTML+RULE",
        "R1-10": "MANUAL", "R2-01": "API/HTML/OCR", "R2-02": "VISION",
        "R2-03": "VISION", "R2-04": "PDP_RELATION+VISION", "R2-05": "VISION",
        "R2-06": "EXTERNAL_CRM", "R2-07": "NORMALIZE+SEMANTIC",
    }
    special_evidence = {
        ("G25070005743", "R1-02"): "https://image.guud.com/mall/DESIGN/PRODUCT/casamia/2023/10/levor/02_levor_Info_4people+stool.jpg",
        ("G25100020496", "R1-02"): "https://www.daydreamer.co.kr/product/detail_pg/2025/10/sizeguide_rug.jpg",
        ("G25100020496", "R2-03"): "https://www.daydreamer.co.kr/product/2025/10/dtp/modulo/03.jpg",
        ("G25070001871", "R1-02"): "https://storage001.daousync.com/v1/AUTH_c55e59a8ec2149e7ad5e61ca73645bd9/mw155231/image/1753067270913.png",
    }

    def method_check(product: dict, requirement_id: str, final_status: str, method: str) -> tuple[str, str]:
        if final_status == "해당없음":
            return "N_A", "해당 상품 유형에 적용되지 않음"
        product_id = product["product_id"]
        source = source_by_product[product_id]
        ocr = ocr_summary_by_product[product_id]

        if method == "API":
            if product["source_system"] == "LGE_APPLIANCE":
                return "NOT_FOUND", "API HTTP 200 실행; dealProductModel 없음으로 핵심 제품 필드 0건"
            found = {"R1-01", "R1-03", "R1-05", "R2-07"}
            if requirement_id in found:
                return "FOUND", f"goods API 200; 구조화 필드 신호 {source['api'].get('field_signal_count', 0)}개"
            if requirement_id == "R1-02":
                if product_id == "G25070005743":
                    return "FOUND", "productNotification 크기에서 소파·스툴 W×D×H 확보"
                if product_id == "G25100020496":
                    return "PARTIAL", "purchaseOptions에서 W×D 3개 확보; 두께/H 없음"
                return "NOT_FOUND", "상품고시 크기가 '상세페이지 참조'이거나 외형 규격 미제공"
            if requirement_id == "R1-06":
                if product_id in {"G25070005743", "G25070001871", "G25070006112"}:
                    return "FOUND", "옵션·고시·상품명에서 색상 확보"
                return "NOT_FOUND", "고시 색상이 '상세페이지 참조'이고 옵션에 색상 없음"
            if requirement_id == "R1-08":
                return "NOT_FOUND", f"packages API 상태 {source['api'].get('packages_api_status')}"
            if requirement_id == "R2-01":
                return ("FOUND", "고시정보의 제조·재질·규격·구성·색상·보증 필드 확보") if product_id == "G25070005743" else ("PARTIAL", "고시정보 일부만 구조화; 상세페이지 참조값 존재")
            return "NOT_FOUND", "해당 요구사항을 직접 제공하는 API 필드 없음"

        if method == "HTML":
            base = f"HTML HTTP {source['html'].get('http_status')}; title={source['html'].get('title', '')[:80]}"
            if product["source_system"] == "HOMESTYLE":
                if requirement_id in {"R1-03", "R1-05"}:
                    return "FOUND", base
                if requirement_id in {"R1-06", "R2-07"}:
                    return "PARTIAL", base + "; 서버 HTML 제목/메타데이터 중심"
                return "NOT_FOUND", base + "; 핵심 상세값은 클라이언트 API로 로딩"
            found = {"R1-01", "R1-02", "R1-03", "R1-05", "R1-06", "R1-07", "R1-09", "R2-03", "R2-04", "R2-07"}
            partial = {"R1-04", "R2-02", "R2-05"}
            if requirement_id in found:
                return "FOUND", base + f"; 규격신호 {len(source['html'].get('dimension_matches') or [])}개"
            if requirement_id in partial:
                return "PARTIAL", base + "; 연출 이미지/대체텍스트에서 추론 가능"
            return "NOT_FOUND", base + "; 직접 명시값 없음"

        if method == "OCR":
            base = (
                f"KO/EN OCR {ocr['downloaded_images']}장 실행; "
                f"KO {ocr['ko_success']}성공/{ocr['ko_nonempty']}텍스트, "
                f"EN {ocr['en_success']}성공/{ocr['en_nonempty']}텍스트"
            )
            if requirement_id == "R1-02":
                if product_id in {"G25070005743", "OLED48C6KNA", "G646GBB031", "WA2525YMZF", "SQ06GJ1WFS"}:
                    return "FOUND", base + "; 규격 OCR=" + "|".join(ocr["dimension_signals"][:8])
                if product_id == "G25100020496":
                    return "PARTIAL", base + "; W×D 배치 가이드 인식, 두께/H 없음"
                return "NOT_FOUND", base + "; 제품 외형 규격 텍스트 미검출"
            if requirement_id == "R1-04" and product_id == "G25100020496":
                return "PARTIAL", base + "; '가구 배치 추천'과 가구별 러그 크기 인식"
            if requirement_id == "R2-01" and product["source_system"] == "HOMESTYLE":
                return "PARTIAL", base + f"; OCR 문자 {ocr['ko_characters'] + ocr['en_characters']}자"
            if requirement_id == "R2-03" and product_id == "G25100020496":
                return "PARTIAL", base + "; 소파·침대·책상·식탁 배치 가이드 인식"
            if requirement_id == "R2-07" and ocr["ko_nonempty"] + ocr["en_nonempty"] > 0:
                return "PARTIAL", base + "; 의미검색 보강 텍스트 확보"
            return "NOT_FOUND", base + "; 해당 요구사항의 직접 문자값 미검출"

        raise ValueError(method)

    def merged_method(product: dict, requirement_id: str, final_status: str) -> str:
        contributors = []
        for method in ("API", "HTML", "OCR"):
            method_status, _ = method_check(product, requirement_id, final_status, method)
            if method_status in ("FOUND", "PARTIAL"):
                contributors.append(method)
        if requirement_id in {"R1-01", "R1-04", "R1-06", "R1-09", "R2-02", "R2-03", "R2-04", "R2-05"}:
            contributors.append("VISION/RULE")
        if requirement_id == "R2-07":
            contributors.append("SEMANTIC")
        if requirement_id == "R2-06":
            contributors.append("EXTERNAL_CRM_REQUIRED")
        return "+".join(dict.fromkeys(contributors)) or "NO_SOURCE"

    method_execution_headers = [
        "product_id", "source_system", "profile", "method", "실행여부", "실행결과", "HTTP/엔진상태",
        "입력건수", "성공건수", "텍스트검출건수", "문자수", "규격신호수", "핵심필드신호수", "근거/비고",
    ]
    method_execution_rows = [method_execution_headers]
    for product in test_products:
        source = source_by_product[product["product_id"]]
        deep_product = deep_product_by_id[product["product_id"]]
        ocr = ocr_summary_by_product[product["product_id"]]
        if product["source_system"] == "HOMESTYLE":
            api_http = source["api"].get("goods_http_status")
            api_result = "SUCCESS_CORE"
            api_note = f"goods={source['api'].get('goods_api_status')}, space={source['api'].get('space_api_status')}, packages={source['api'].get('packages_api_status')}"
        else:
            api_http = source["api"].get("http_status")
            api_result = "SUCCESS_LIMITED"
            api_note = "retrieveDealProduct 응답 성공, dealProductModel 없음"
        method_execution_rows.append([
            product["product_id"], product["source_system"], product["profile"], "API", "Y", api_result, api_http,
            1, 1 if api_http == 200 else 0, "", "", "", len(deep_product["api_fields"]),
            api_note + "; 심층 정규화 필드수=" + str(len(deep_product["api_fields"])),
        ])
        method_execution_rows.append([
            product["product_id"], product["source_system"], product["profile"], "HTML", "Y", "SUCCESS",
            source["html"].get("http_status"), 1, 1 if source["html"].get("http_status") == 200 else 0,
            1 if source["html"].get("visible_character_count", 0) > 0 else 0, source["html"].get("visible_character_count", 0),
            len(source["html"].get("dimension_matches") or []), len(deep_product["html_fields"]),
            f"원본 HTML 저장; 심층 정규화 필드수={len(deep_product['html_fields'])}; SHA256={source['html'].get('sha256')}"
        ])
        method_execution_rows.append([
            product["product_id"], product["source_system"], product["profile"], "WINDOWS_OCR", "Y", "SUCCESS",
            "Windows OCR ko+en-US", ocr["downloaded_images"] * 2, ocr["ko_success"] + ocr["en_success"],
            ocr["ko_nonempty"] + ocr["en_nonempty"], ocr["ko_characters"] + ocr["en_characters"],
            len(ocr["dimension_signals"]), len(deep_product["windows_ocr_fields"]),
            "한국어와 영어 OCR을 동일 이미지에 각각 실행; 최종 채택 필드수=" + str(len(deep_product["windows_ocr_fields"])),
        ])
        for engine_name, method_name in (
            ("PaddleOCR PP-OCRv5", "PADDLE_OCR"),
            ("Tesseract.js", "TESSERACT_JS"),
        ):
            engine_images = benchmark_images_by_engine_product.get((engine_name, product["product_id"]), [])
            engine_tokens = benchmark_tokens_by_engine_product.get((engine_name, product["product_id"]), [])
            engine_phrases = benchmark_phrases_by_engine_product.get((engine_name, product["product_id"]), [])
            token_hits = sum(bool(row["found"]) for row in engine_tokens)
            phrase_hits = sum(bool(row["found"]) for row in engine_phrases)
            method_execution_rows.append([
                product["product_id"], product["source_system"], product["profile"], method_name, "Y", "SUCCESS",
                benchmark_summary_by_engine[engine_name]["version"], len(engine_images),
                sum(row["status"] in ("SUCCESS", "SKIPPED_TINY") for row in engine_images),
                sum(row["character_count"] > 0 for row in engine_images),
                sum(row["character_count"] for row in engine_images), token_hits,
                len(deep_product["paddle_ocr_fields"]) if method_name == "PADDLE_OCR" else 0,
                f"정답 숫자 {token_hits}/{len(engine_tokens) if engine_tokens else 0}; "
                f"핵심 문구 {phrase_hits}/{len(engine_phrases) if engine_phrases else 0}; "
                f"최종 채택 필드수={len(deep_product['paddle_ocr_fields']) if method_name == 'PADDLE_OCR' else 0}",
            ])

    ocr_image_headers = [
        "product_id", "file", "role", "image_url", "alt_text", "byte_count", "KO_status", "KO_chars", "KO_text",
        "EN_status", "EN_chars", "EN_text", "dimension_signals",
    ]
    ocr_image_rows = [ocr_image_headers]
    for manifest in image_manifest:
        filename = manifest.get("file")
        if not filename:
            continue
        ko = ocr_ko_by_file.get(filename, {})
        en = ocr_en_by_file.get(filename, {})
        merged = (ko.get("text", "") + "\n" + en.get("text", "")).strip()
        ocr_image_rows.append([
            manifest["product_id"], filename, manifest.get("role", ""), manifest.get("image_url", ""), manifest.get("alt", ""),
            manifest.get("byte_count", 0), ko.get("status", ""), ko.get("character_count", 0), ko.get("text", "")[:5000],
            en.get("status", ""), en.get("character_count", 0), en.get("text", "")[:5000],
            "|".join(sorted(set(dimension_pattern.findall(merged)))[:20]),
        ])

    ocr_compare_headers = [
        "quality_rank", "engine", "version", "language", "처리이미지", "성공", "오류", "텍스트검출",
        "규격정답수", "규격적중", "규격회수율", "문구정답수", "문구적중", "문구회수율",
        "무필드확인이미지", "오탐이미지", "오탐률", "총처리초", "장당ms", "판정", "설정/비고",
    ]
    ocr_compare_rows = [ocr_compare_headers]
    for row in sorted(ocr_benchmark["summary"], key=lambda item: item["quality_rank"]):
        if row["quality_rank"] == 1:
            judgment = "정확도 1순위 / 규격·상세이미지 선택 적용"
        elif row["engine"].startswith("Windows"):
            judgment = "속도 1순위 / 빠른 1차 탐색"
        else:
            judgment = "수치 추출 비권장 / 문구 보조"
        ocr_compare_rows.append([
            row["quality_rank"], row["engine"], row["version"], row["language"], row["processed_images"],
            row["successful_images"], row["error_images"], row["nonempty_images"], row["critical_token_total"],
            row["critical_token_hits"], f"{row['critical_token_recall_pct']:.1f}%", row["critical_phrase_total"],
            row["critical_phrase_hits"], f"{row['critical_phrase_recall_pct']:.1f}%", row["confirmed_blank_images"],
            row["false_positive_images"], f"{row['false_positive_rate_pct']:.1f}%", round(row["elapsed_ms"] / 1000, 3),
            row["average_ms_per_image"], judgment, row["configuration_note"],
        ])

    gold_lookup = {
        (row["gold_type"], row["product_id"], row["gold_value"], row["engine"]): row["found"]
        for row in ocr_benchmark["token_details"] + ocr_benchmark["phrase_details"]
    }
    gold_items = []
    seen_gold = set()
    for row in ocr_benchmark["token_details"] + ocr_benchmark["phrase_details"]:
        key = (row["gold_type"], row["product_id"], row["product_group"], row["gold_value"])
        if key not in seen_gold:
            seen_gold.add(key)
            gold_items.append(key)

    def gold_request_group(gold_type: str, product_id: str, value: str) -> tuple[str, str]:
        if gold_type == "CRITICAL_NUMERIC_TOKEN":
            return "R1-02", "R1"
        if product_id == "G25070005743" and value in {"4인 소파+스툴", "카멜브라운", "크림 아이보리"}:
            return "R1-06/R1-08|R2-01", "BOTH"
        if product_id == "G25100020496":
            return "R1-04|R2-03", "BOTH"
        if product_id == "G25070006112" and value == "10X15CM":
            return "R1-02", "R1"
        if product_id == "WA2525YMZF":
            return "R1-02", "R1"
        return "R2-01", "R2"

    ocr_gold_headers = [
        "gold_type", "요청연결", "product_id", "대표제품군", "정답값",
        "Windows_OCR", "PaddleOCR", "Tesseract_js", "비고",
    ]
    ocr_gold_rows = [ocr_gold_headers]
    ocr_gold_row_groups = []
    for gold_type, product_id, product_group, gold_value in gold_items:
        request_link, row_group = gold_request_group(gold_type, product_id, gold_value)
        windows_found = gold_lookup.get((gold_type, product_id, gold_value, "Windows OCR (ko+en)"), False)
        paddle_found = gold_lookup.get((gold_type, product_id, gold_value, "PaddleOCR PP-OCRv5"), False)
        tesseract_found = gold_lookup.get((gold_type, product_id, gold_value, "Tesseract.js"), False)
        ocr_gold_rows.append([
            gold_type, request_link, product_id, product_group, gold_value,
            "O" if windows_found else "X", "O" if paddle_found else "X", "O" if tesseract_found else "X",
            "정확 일치 기준; 숫자는 앞뒤 다른 숫자가 붙으면 오인식으로 판정",
        ])
        ocr_gold_row_groups.append(row_group)

    ocr_engine_image_headers = [
        "engine", "product_id", "대표제품군", "file", "role", "status", "line_count", "character_count",
        "mean_confidence", "elapsed_ms", "무필드확인이미지", "오탐판정", "text_excerpt", "error",
    ]
    ocr_engine_image_rows = [ocr_engine_image_headers]
    for row in sorted(ocr_benchmark["image_details"], key=lambda item: (item["engine"], item["product_id"], item["file"])):
        ocr_engine_image_rows.append([
            row["engine"], row["product_id"], row["product_group"], row["file"], row["role"], row["status"],
            row["line_count"], row["character_count"], row["mean_confidence"] if row["mean_confidence"] is not None else "",
            row["elapsed_ms"], "Y" if row["confirmed_no_field_text"] else "N",
            "Y" if row["false_positive_on_confirmed_blank"] else "N", row["text_excerpt"], row["error"],
        ])

    source_increment_headers = [
        "단계", "전체필드셀", "완전확보", "조건부확보", "사용가능합계", "결측",
        "엄격커버리지", "조건포함커버리지", "API대비증가", "HTML단계대비증가", "해석",
    ]
    source_increment_rows = [source_increment_headers]
    stage_explanations = {
        "API_ONLY": "홈스타일 API의 구조화 상품·가격·재고·옵션·배송·고시정보를 정규화. 가전 API 응답에는 핵심 제품필드가 없어 결측으로 계산",
        "API_HTML": "가전 HTML JSON-LD의 상품명·SKU·가격·평점·리뷰·상세스펙을 병합한 주 개선 단계",
        "API_HTML_WINDOWS_OCR": "API/HTML 결측에만 Windows OCR 채택값을 추가. 러그의 추천배치·특징 2필드 보강",
        "API_HTML_PADDLE_OCR": "동일 결측 2필드를 PaddleOCR로 더 높은 신뢰도로 채택. 사용가능 수는 같지만 완전확보 판정이 2칸 증가",
    }
    for stage in deep_source["stage_statistics"]:
        source_increment_rows.append([
            stage["stage"], stage["total_field_cells"], stage["available"], stage["partial"], stage["usable"], stage["missing"],
            f"{stage['strict_coverage_pct']}%", f"{stage['usable_coverage_pct']}%",
            stage["usable"] - deep_api["usable"], "" if stage["stage"] == "API_ONLY" else stage["usable"] - deep_html["usable"],
            stage_explanations[stage["stage"]],
        ])

    source_product_headers = [
        "단계", "product_id", "source_system", "전체필드", "완전확보", "조건부확보", "사용가능합계",
        "결측", "조건포함커버리지", "상품별_API대비증가", "직전단계대비증가",
    ]
    source_product_rows = [source_product_headers]
    api_product_usable = {row["product_id"]: row["usable"] for row in deep_api["per_product"]}
    previous_product_usable: dict[str, int] = {}
    for stage in deep_source["stage_statistics"]:
        for row in stage["per_product"]:
            product_id = row["product_id"]
            previous = previous_product_usable.get(product_id, row["usable"])
            source_product_rows.append([
                stage["stage"], product_id, deep_product_by_id[product_id]["source_system"],
                deep_source["metadata"]["field_count"], row["available"], row["partial"], row["usable"], row["missing"],
                f"{row['usable'] / deep_source['metadata']['field_count']:.1%}",
                row["usable"] - api_product_usable[product_id], row["usable"] - previous,
            ])
            previous_product_usable[product_id] = row["usable"]

    new_requirement_headers = [
        "추가요구ID", "요구항목", "요청구분", "구성필드", "상품수", "가능", "조건부", "불가능",
        "엄격가능률", "조건포함가능률", "판정메모",
    ]
    new_requirement_rows = [new_requirement_headers]
    new_requirement_row_groups = []
    for row in deep_source["requirement_summary"]:
        new_requirement_rows.append([
            row["requirement_id"], row["requirement_name"], row["request_group"], "|".join(row["fields"]),
            row["product_count"], row["available"], row["conditional"], row["unavailable"],
            f"{row['strict_pct']}%", f"{row['conditional_included_pct']}%",
            "8개 대표상품 실측. 가격·재고·배송은 수집시점 스냅샷이며 운영 시 재수집 필요",
        ])
        new_requirement_row_groups.append(row["request_group"])

    new_requirement_detail_headers = [
        "추가요구ID", "요구항목", "요청구분", "product_id", "source_system", "판정",
        "구성필드", "필드상태", "값", "채택소스", "근거",
    ]
    new_requirement_detail_rows = [new_requirement_detail_headers]
    new_requirement_detail_row_groups = []
    for row in deep_source["requirement_evidence"]:
        new_requirement_detail_rows.append([
            row["requirement_id"], row["requirement_name"], row["request_group"], row["product_id"],
            row["source_system"], row["status"], "|".join(row["field_names"]),
            compact(row["field_statuses"]), compact(row["values"]), compact(row["sources"]), compact(row["evidence"]),
        ])
        new_requirement_detail_row_groups.append(row["request_group"])

    deep_field_defs = deep_source["field_definitions"]
    extended_result_headers = [
        "product_id", "source_system", "완전확보수", "조건부수", "결측수", "검수필요필드",
    ] + [f"{row['field_label']}\n({row['field_name']})" for row in deep_field_defs]
    extended_result_rows = [extended_result_headers]
    for product in deep_source["products"]:
        merged = product["merged_fields"]
        statuses = {row["field_name"]: merged.get(row["field_name"], {}).get("status", "MISSING") for row in deep_field_defs}
        extended_result_rows.append([
            product["product_id"], product["source_system"],
            sum(value == "AVAILABLE" for value in statuses.values()),
            sum(value == "PARTIAL" for value in statuses.values()),
            sum(value == "MISSING" for value in statuses.values()),
            "|".join(key for key, value in statuses.items() if value != "AVAILABLE"),
        ] + [compact(merged.get(row["field_name"], {}).get("value", "")) for row in deep_field_defs])
    extended_result_groups = ["COMMON"] * 6 + [row["request_group"] for row in deep_field_defs]

    deep_evidence_headers = [
        "product_id", "source_system", "field_name", "필드명", "요청구분", "상태", "값",
        "채택소스", "confidence", "근거",
    ]
    deep_evidence_rows = [deep_evidence_headers]
    deep_evidence_row_groups = []
    for row in deep_source["field_evidence"]:
        deep_evidence_rows.append([
            row["product_id"], row["source_system"], row["field_name"], row["field_label"], row["request_group"],
            row["status"], compact(row["value"]), row["source"], row["confidence"], row["evidence"],
        ])
        deep_evidence_row_groups.append(row["request_group"])

    raw_inventory_headers = ["product_id", "source_system", "원천", "원천경로", "원천값"]
    raw_inventory_rows = [raw_inventory_headers]
    for product in deep_source["products"]:
        for row in product["raw_inventory"]:
            raw_inventory_rows.append([
                product["product_id"], product["source_system"], row["source"], row["path"], compact(row["value"]),
            ])

    detail_headers = [
        "batch_id", "product_id", "profile", "source_system", "대분류", "중분류", "소분류",
        "요청구분", "requirement_id", "항목",
        "API_실행", "API_결과", "API_근거", "HTML_실행", "HTML_결과", "HTML_근거",
        "OCR_실행", "OCR_결과", "OCR_근거", "최종병합소스",
        "판정", "판정근거", "confidence", "evidence_url", "검수필요", "테스트일",
    ]
    detail_rows = [detail_headers]
    detail_row_groups = []
    for product in test_products:
        statuses = status_matrix[product["product_id"]]
        for requirement_id, status in zip(requirement_ids, statuses):
            confidence = "" if status == "해당없음" else 0 if status == "불가능" else 0.75 if status == "조건부" else 0.95
            api_status, api_evidence = method_check(product, requirement_id, status, "API")
            html_status, html_evidence = method_check(product, requirement_id, status, "HTML")
            ocr_status, ocr_evidence = method_check(product, requirement_id, status, "OCR")
            detail_rows.append([
                "POC-REP-20260720", product["product_id"], product["profile"], product["source_system"],
                product["large"], product["mid"], product["small"], "요청 1" if requirement_id.startswith("R1") else "요청 2",
                requirement_id, requirement_name[requirement_id],
                "Y", api_status, api_evidence, "Y", html_status, html_evidence, "Y", ocr_status, ocr_evidence,
                merged_method(product, requirement_id, status), status, assessment_reason(product, requirement_id, status), confidence,
                special_evidence.get((product["product_id"], requirement_id), product["url"]),
                "Y" if status in ("조건부", "불가능") else "N", "2026-07-20",
            ])
            detail_row_groups.append("R1" if requirement_id.startswith("R1") else "R2")

    status_names = ("가능", "조건부", "불가능", "해당없음")

    def status_counts(statuses: list[str]) -> dict[str, int]:
        return {name: statuses.count(name) for name in status_names}

    all_statuses = [status for product in test_products for status in status_matrix[product["product_id"]]]
    overall = status_counts(all_statuses)
    overall_applicable = len(all_statuses) - overall["해당없음"]
    dashboard_headers = [
        "통계구분", "ID/상품", "요청구분/프로필", "제품명/설명", "전체판정수", "적용항목수",
        "가능", "조건부", "불가능", "해당없음", "엄격가능률", "조건포함가능률", "주요 미확보/판정",
    ]
    dashboard_rows = [dashboard_headers, [
        "전체", "POC-REP-20260720", "홈스타일4+가전4", "대표 8상품 × 17요구사항; API·HTML·OCR 모두 실행",
        len(all_statuses), overall_applicable, overall["가능"], overall["조건부"], overall["불가능"], overall["해당없음"],
        f"{overall['가능'] / overall_applicable:.1%}", f"{(overall['가능'] + overall['조건부']) / overall_applicable:.1%}",
        "API 8/8·HTML 8/8; Windows/Paddle/Tesseract.js 동일 62이미지 실행. Paddle 규격 28/28, Windows 27/28; 최종 불가 원천은 CRM·높이·규격·세트ID",
    ]]
    dashboard_rows.extend([
        [
            "심층필드", "API_ONLY", "37필드×8상품", "원본 API 구조를 전체 필드 단위로 정규화",
            deep_api["total_field_cells"], deep_api["total_field_cells"], deep_api["available"], deep_api["partial"],
            deep_api["missing"], 0, f"{deep_api['strict_coverage_pct']}%", f"{deep_api['usable_coverage_pct']}%",
            "홈스타일 API는 가격·재고·옵션·배송·고시정보까지 확보; 가전 API는 핵심 제품필드 제한",
        ],
        [
            "심층필드", "API_HTML_PADDLE", "37필드×8상품", "API+HTML 심층 병합 후 PaddleOCR 결측 보강",
            deep_paddle["total_field_cells"], deep_paddle["total_field_cells"], deep_paddle["available"], deep_paddle["partial"],
            deep_paddle["missing"], 0, f"{deep_paddle['strict_coverage_pct']}%", f"{deep_paddle['usable_coverage_pct']}%",
            f"API 단독 대비 사용가능 +{deep_paddle['usable'] - deep_api['usable']}칸; 이 중 HTML +{deep_html['usable'] - deep_api['usable']}칸, OCR +{deep_paddle['usable'] - deep_html['usable']}칸",
        ],
    ])
    for product in test_products:
        counts = status_counts(status_matrix[product["product_id"]])
        applicable = 17 - counts["해당없음"]
        missing = [requirement_name[req] for req, status in zip(requirement_ids, status_matrix[product["product_id"]]) if status == "불가능"]
        dashboard_rows.append([
            "상품", product["product_id"], product["profile"], product["name"], 17, applicable,
            counts["가능"], counts["조건부"], counts["불가능"], counts["해당없음"],
            f"{counts['가능'] / applicable:.1%}", f"{(counts['가능'] + counts['조건부']) / applicable:.1%}",
            "|".join(missing) if missing else "없음",
        ])

    request_stat_headers = [
        "요청구분", "requirement_id", "항목", "전체상품수", "적용상품수", "가능", "조건부", "불가능", "해당없음",
        "엄격가능률", "조건포함가능률", "대표 제약",
    ]
    request_stat_rows = [request_stat_headers]
    request_stat_row_groups = []
    for idx, requirement_id in enumerate(requirement_ids):
        statuses = [status_matrix[product["product_id"]][idx] for product in test_products]
        counts = status_counts(statuses)
        applicable = 8 - counts["해당없음"]
        failed_products = [product["product_id"] for product, status in zip(test_products, statuses) if status == "불가능"]
        request_stat_rows.append([
            "요청 1" if requirement_id.startswith("R1") else "요청 2", requirement_id, requirement_name[requirement_id], 8, applicable,
            counts["가능"], counts["조건부"], counts["불가능"], counts["해당없음"],
            f"{counts['가능'] / applicable:.1%}" if applicable else "N/A",
            f"{(counts['가능'] + counts['조건부']) / applicable:.1%}" if applicable else "N/A",
            "불가: " + "|".join(failed_products) if failed_products else "불가 없음",
        ])
        request_stat_row_groups.append("R1" if requirement_id.startswith("R1") else "R2")

    scope_rows = _scope_rows(selected)
    scope_headers = scope_rows[0]
    scope_groups = groups(scope_headers, r1=("요청1_대상", "요청1_requirement_ids"), r2=("요청2_대상", "요청2_requirement_ids"), both=("scope_id", "대분류", "중분류", "소분류", "category_id"))

    input_headers = [
        "batch_id", "scope_id", "source_system", "parse_profile", "대분류", "중분류", "소분류", "category_id",
        "product_id", "model_id", "option_id", "source_url", "request1_target_yn", "request1_requirement_ids",
        "request2_target_yn", "request2_requirement_ids", "set_yn", "priority", "active_yn", "input_status",
        "input_error", "requested_at", "비고",
    ]
    input_sample = {
        "batch_id": "SAMPLE-202607-001", "scope_id": next(item["scope_id"] for item in selected if item["small"] == "일반소파"),
        "source_system": "HOMESTYLE", "parse_profile": "HOME_FURNITURE", "대분류": "가구", "중분류": "소파",
        "소분류": "일반소파", "category_id": "2506000070", "product_id": "G25070005743",
        "source_url": "https://homestyle.lge.co.kr/item?productId=G25070005743", "request1_target_yn": "Y",
        "request1_requirement_ids": "R1-01|R1-02|R1-03|R1-04|R1-05|R1-06|R1-09",
        "request2_target_yn": "Y", "request2_requirement_ids": "R2-01|R2-02|R2-03|R2-04|R2-05|R2-06|R2-07",
        "set_yn": "Y", "priority": "P1", "active_yn": "N", "input_status": "SAMPLE",
        "requested_at": "2026-07-20T00:00:00+09:00", "비고": "샘플 행; 실제 배치 전 삭제 또는 active_yn=Y로 변경",
    }
    input_rows = [input_headers]
    for product in test_products:
        input_row = {
            "batch_id": "POC-REP-20260720", "scope_id": product["scope_id"],
            "source_system": product["source_system"], "parse_profile": product["profile"],
            "대분류": product["large"], "중분류": product["mid"], "소분류": product["small"],
            "category_id": product["category_id"], "product_id": product["product_id"],
            "model_id": product["model_id"], "source_url": product["url"],
            "request1_target_yn": "Y", "request1_requirement_ids": "|".join(requirement_ids[:10]),
            "request2_target_yn": "Y", "request2_requirement_ids": "|".join(requirement_ids[10:]),
            "set_yn": product["set_yn"], "priority": "P1", "active_yn": "Y", "input_status": "POC_DONE",
            "requested_at": "2026-07-20T16:00:00+09:00", "비고": "대표 카테고리 실제 PDP 점검",
        }
        input_rows.append([input_row.get(header, "") for header in input_headers])
    input_groups = groups(input_headers, r1=("request1_target_yn", "request1_requirement_ids"), r2=("request2_target_yn", "request2_requirement_ids"), both=("scope_id", "대분류", "중분류", "소분류", "category_id", "product_id", "model_id", "source_url"))

    result_headers = [
        "run_id", "batch_id", "scope_id", "product_id", "model_id", "source_url", "parse_status", "error_code", "error_message",
        "request1_check_status", "request1_missing_fields", "request1_evidence_count",
        "request2_check_status", "request2_missing_fields", "request2_evidence_count",
        "product_name", "brand_id", "brand_name", "large_category", "mid_category", "small_category",
        "pdp_image_url", "image_view", "width_mm", "depth_mm", "height_mm", "dimension_components_json", "dimension_note",
        "recommended_spaces", "colors_raw", "colors_normalized", "installation_type", "set_yn", "set_component_product_ids",
        "placement_positions", "wall_mount_height_mm",
        "tag_basic", "tag_manufacturer", "tag_material", "tag_spec", "tag_components", "tag_color", "tag_use_purpose",
        "tag_assembly", "tag_safety", "tag_handling", "tag_warranty", "tag_certification", "tag_sales",
        "design_style_tags", "space_content_ids", "space_relation_count", "product_relation_count", "crm_link_status",
        "semantic_search_text", "semantic_synonyms", "embedding_id", "review_required_yn", "parser_version", "extracted_at",
    ]
    result_sample = {
        "run_id": "RUN-SAMPLE-001", "batch_id": "SAMPLE-202607-001", "scope_id": input_sample["scope_id"],
        "product_id": "G25070005743", "source_url": input_sample["source_url"], "parse_status": "PARTIAL",
        "request1_check_status": "PARTIAL", "request1_missing_fields": "set_component_product_ids", "request1_evidence_count": 9,
        "request2_check_status": "PARTIAL", "request2_missing_fields": "CRM_EXTERNAL|embedding_id", "request2_evidence_count": 14,
        "product_name": "벨로씨 레보르 천연면피 가죽 소파 4인+스툴", "brand_id": "2500000123", "brand_name": "까사미아",
        "large_category": "가구", "mid_category": "소파", "small_category": "일반소파",
        "pdp_image_url": "https://static-store.lge.co.kr/goods/org/073251020000078073.jpg", "image_view": "FRONT",
        "width_mm": 2910, "depth_mm": 1020, "height_mm": "730~910", "dimension_components_json": "{\"stool\":{\"w\":740,\"d\":660,\"h\":410}}",
        "dimension_note": "가변 헤드레스트", "recommended_spaces": "LIVING_ROOM", "colors_raw": "카멜브라운|크림아이보리",
        "colors_normalized": "BROWN_WARM|IVORY_OFFWHITE", "installation_type": "FLOOR_STANDING", "set_yn": "Y",
        "placement_positions": "FLOOR", "tag_basic": "가구|소파|4인용|스툴세트", "tag_manufacturer": "벨로씨OEM|중국",
        "tag_material": "천연면피소가죽|PVC합성가죽|E0합판|HR폼|스틸", "tag_spec": "SOFA_2910x1020x730~910|STOOL_740x660x410",
        "tag_components": "4인소파|스툴", "tag_color": "카멜브라운|크림아이보리", "tag_use_purpose": "REST|TV|GUEST",
        "tag_assembly": "VENDOR_INSTALL", "tag_safety": "SAFETY_STANDARD_COMPLIANT", "tag_handling": "INSPECT_BEFORE_INSTALL",
        "tag_warranty": "1_YEAR", "tag_certification": "NO_CERT_NUMBER", "tag_sales": "ON_SALE",
        "design_style_tags": "CONTEMPORARY|MODERN|RECTILINEAR", "space_content_ids": "2605001190",
        "space_relation_count": 1, "product_relation_count": 6, "crm_link_status": "EXTERNAL_REQUIRED",
        "semantic_search_text": "웜브라운 천연가죽 4인용 거실 소파와 스툴", "semantic_synonyms": "카멜브라운=웜브라운|크림아이보리=오프화이트",
        "review_required_yn": "Y", "parser_version": "pdp-parser-0.2.0", "extracted_at": "2026-07-20T00:00:00+09:00",
    }
    result_rows = [result_headers]
    for index, product in enumerate(test_products, start=1):
        statuses = status_matrix[product["product_id"]]
        r1_statuses, r2_statuses = statuses[:10], statuses[10:]
        r1_issues = [f"{req}:{status}" for req, status in zip(requirement_ids[:10], r1_statuses) if status in ("조건부", "불가능")]
        r2_issues = [f"{req}:{status}" for req, status in zip(requirement_ids[10:], r2_statuses) if status in ("조건부", "불가능")]
        result = {
            "run_id": f"RUN-REP-{index:02d}", "batch_id": "POC-REP-20260720", "scope_id": product["scope_id"],
            "product_id": product["product_id"], "model_id": product["model_id"], "source_url": product["url"],
            "parse_status": "PARTIAL", "error_code": "", "error_message": "",
            "request1_check_status": "PARTIAL" if r1_issues else "PASS", "request1_missing_fields": "|".join(r1_issues),
            "request1_evidence_count": sum(status in ("가능", "조건부") for status in r1_statuses),
            "request2_check_status": "PARTIAL" if r2_issues else "PASS", "request2_missing_fields": "|".join(r2_issues),
            "request2_evidence_count": sum(status in ("가능", "조건부") for status in r2_statuses),
            "product_name": product["name"], "brand_name": product["brand"], "large_category": product["large"],
            "mid_category": product["mid"], "small_category": product["small"], "pdp_image_url": product["image"],
            "image_view": "VISION_REVIEW", "width_mm": product["width"], "depth_mm": product["depth"],
            "height_mm": product["height"], "dimension_components_json": product["dimension_json"],
            "dimension_note": product["dimension_note"], "recommended_spaces": product["spaces"],
            "colors_raw": product["colors_raw"], "colors_normalized": product["colors_normalized"],
            "installation_type": product["installation"], "set_yn": product["set_yn"],
            "set_component_product_ids": product["set_components"], "placement_positions": product["placement"],
            "wall_mount_height_mm": "", "tag_basic": product["tag_basic"], "tag_manufacturer": product["brand"],
            "tag_material": product["tag_material"], "tag_spec": product["tag_spec"], "tag_components": product["tag_components"],
            "tag_color": product["colors_raw"], "tag_use_purpose": product["tag_use"],
            "tag_assembly": product["installation"], "tag_safety": "PDP_SAFETY_TEXT", "tag_handling": "PDP_HANDLING_TEXT",
            "tag_warranty": "PDP_WARRANTY_TEXT", "tag_certification": "PDP_CERT_TEXT", "tag_sales": "PUBLIC_PDP_ACTIVE",
            "design_style_tags": product["style"], "space_content_ids": "PDP_SCENE_IMAGE",
            "space_relation_count": 1, "product_relation_count": "VISION_CONDITIONAL", "crm_link_status": "EXTERNAL_REQUIRED",
            "semantic_search_text": product["semantic"], "semantic_synonyms": product["colors_raw"],
            "embedding_id": "", "review_required_yn": "Y", "parser_version": "pdp-parser-poc-0.4.0",
            "extracted_at": "2026-07-20T16:30:00+09:00",
        }
        result_rows.append([result.get(header, "") for header in result_headers])
    r1_result_fields = {"request1_check_status", "request1_missing_fields", "request1_evidence_count", "image_view", "dimension_note", "installation_type", "set_yn", "set_component_product_ids", "placement_positions", "wall_mount_height_mm"}
    r2_result_fields = {"request2_check_status", "request2_missing_fields", "request2_evidence_count", "tag_basic", "tag_manufacturer", "tag_material", "tag_spec", "tag_components", "tag_color", "tag_use_purpose", "tag_assembly", "tag_safety", "tag_handling", "tag_warranty", "tag_certification", "tag_sales", "design_style_tags", "space_content_ids", "space_relation_count", "product_relation_count", "crm_link_status", "semantic_search_text", "semantic_synonyms", "embedding_id"}
    both_result_fields = {"scope_id", "product_id", "model_id", "source_url", "product_name", "brand_id", "brand_name", "large_category", "mid_category", "small_category", "pdp_image_url", "width_mm", "depth_mm", "height_mm", "dimension_components_json", "recommended_spaces", "colors_raw", "colors_normalized"}
    result_groups = groups(result_headers, r1=r1_result_fields, r2=r2_result_fields, both=both_result_fields)

    evidence_headers = [
        "run_id", "batch_id", "scope_id", "product_id", "request_type", "requirement_id", "field_name", "required_yn",
        "raw_value", "normalized_value", "unit", "value_status", "extract_method", "confidence", "evidence_url",
        "image_index", "bbox", "evidence_text", "error_code", "review_status", "reviewer", "reviewed_at", "review_note",
    ]
    evidence_sample_1 = {
        "run_id": "RUN-SAMPLE-001", "batch_id": "SAMPLE-202607-001", "scope_id": input_sample["scope_id"], "product_id": "G25070005743",
        "request_type": "REQUEST_1", "requirement_id": "R1-02", "field_name": "size_wdh", "required_yn": "Y",
        "raw_value": "[4인] W2910 X D1020 X H910mm", "normalized_value": "W2910×D1020×H730~910", "unit": "mm",
        "value_status": "EXACT_MERGED", "extract_method": "API+OCR", "confidence": 0.98,
        "evidence_url": "https://image.guud.com/mall/DESIGN/PRODUCT/casamia/2023/10/levor/02_levor_Info_4people+stool.jpg",
        "image_index": 1, "evidence_text": "상세 규격 이미지에 가변 높이 H730~910mm", "review_status": "UNREVIEWED",
    }
    evidence_sample_2 = {
        "run_id": "RUN-SAMPLE-001", "batch_id": "SAMPLE-202607-001", "scope_id": input_sample["scope_id"], "product_id": "G25070005743",
        "request_type": "REQUEST_2", "requirement_id": "R2-02", "field_name": "design_style_tags", "required_yn": "Y",
        "raw_value": "제품 대표 이미지", "normalized_value": "CONTEMPORARY|MODERN|RECTILINEAR", "value_status": "INFERRED",
        "extract_method": "VISION", "confidence": 0.82, "evidence_url": input_sample["source_url"], "evidence_text": "제품 형태 및 연출 이미지 기반",
        "review_status": "UNREVIEWED",
    }
    evidence_rows = [evidence_headers]
    evidence_row_groups = []
    for product_index, product in enumerate(test_products, start=1):
        for requirement_id, status in zip(requirement_ids, status_matrix[product["product_id"]]):
            for method in ("API", "HTML", "OCR"):
                method_status, method_evidence = method_check(product, requirement_id, status, method)
                confidence = "" if method_status == "N_A" else 0 if method_status == "NOT_FOUND" else 0.75 if method_status == "PARTIAL" else 0.95
                if method == "API" and product["source_system"] == "HOMESTYLE":
                    evidence_url = f"https://livingapi.lge.co.kr/itemsvc/ajax/v1/pdp/goods/{product['product_id']}?epFlagYn=N"
                elif method == "API":
                    evidence_url = "https://apiv2.lge.co.kr/itemsvc/ajax/v1/product/retrieveDealProduct"
                elif method == "OCR":
                    evidence_url = special_evidence.get((product["product_id"], requirement_id), product["image"])
                else:
                    evidence_url = product["url"]
                evidence = {
                    "run_id": f"RUN-REP-{product_index:02d}", "batch_id": "POC-REP-20260720", "scope_id": product["scope_id"],
                    "product_id": product["product_id"], "request_type": "REQUEST_1" if requirement_id.startswith("R1") else "REQUEST_2",
                    "requirement_id": requirement_id, "field_name": requirement_name[requirement_id], "required_yn": "Y",
                    "raw_value": method_evidence, "normalized_value": status, "value_status": method_status,
                    "extract_method": method, "confidence": confidence, "evidence_url": evidence_url,
                    "evidence_text": method_evidence, "error_code": "METHOD_NOT_FOUND" if method_status == "NOT_FOUND" else "",
                    "review_status": "UNREVIEWED" if status in ("조건부", "불가능") else "APPROVED",
                    "review_note": f"방법={method}; 최종판정={status}; 병합={merged_method(product, requirement_id, status)}",
                }
                evidence_rows.append([evidence.get(header, "") for header in evidence_headers])
                evidence_row_groups.append("R1" if requirement_id.startswith("R1") else "R2")

    relation_headers = [
        "run_id", "requirement_ids", "space_content_id", "space_source_url", "space_style_tags", "mood_tags", "room_name",
        "room_purpose", "color_tone", "room_size_band", "product_id", "related_product_id", "relation_type",
        "placement_reason", "position_label", "x", "y", "bbox", "value_status", "extract_method", "confidence", "evidence_url",
    ]
    relation_rows = [relation_headers]
    for index, product in enumerate(test_products, start=1):
        relation = {
            "run_id": f"RUN-REP-{index:02d}", "requirement_ids": "R2-03|R2-04|R2-05",
            "space_content_id": f"PDP-SCENE-{product['product_id']}-01", "space_source_url": product["url"],
            "space_style_tags": product["style"], "mood_tags": "CALM|MODERN", "room_name": product["spaces"].split("|")[0],
            "room_purpose": product["tag_use"], "color_tone": product["colors_normalized"], "room_size_band": "UNKNOWN",
            "product_id": product["product_id"], "related_product_id": "", "relation_type": "PDP_SCENE_CONTAINS_PRODUCT",
            "placement_reason": "카테고리와 설치형태 및 PDP 연출 이미지", "position_label": product["placement"],
            "value_status": "INFERRED", "extract_method": "VISION+PDP_RELATION", "confidence": 0.80,
            "evidence_url": special_evidence.get((product["product_id"], "R2-03"), product["url"]),
        }
        relation_rows.append([relation.get(header, "") for header in relation_headers])

    qa_headers = [
        "run_id", "batch_id", "scope_id", "source_system", "parse_profile", "input_count", "success_count", "partial_count", "error_count",
        "request1_expected_fields", "request1_pass_fields", "request1_missing_count", "request1_low_confidence_count", "request1_status",
        "request2_expected_fields", "request2_pass_fields", "request2_missing_count", "request2_low_confidence_count", "request2_status",
        "duplicate_product_count", "set_id_missing_count", "review_sample_count", "review_pass_count", "review_pass_rate", "batch_status", "owner", "updated_at", "note",
    ]
    qa_row = {
        "run_id": "RUN-REP-ALL", "batch_id": "POC-REP-20260720", "scope_id": "8_REPRESENTATIVE_PRODUCTS",
        "source_system": "HOMESTYLE+LGE_APPLIANCE", "parse_profile": "8_PROFILES", "input_count": 8,
        "success_count": 0, "partial_count": 8, "error_count": 0,
        "request1_expected_fields": 80, "request1_pass_fields": sum(s == "가능" for p in test_products for s in status_matrix[p["product_id"]][:10]),
        "request1_missing_count": sum(s == "불가능" for p in test_products for s in status_matrix[p["product_id"]][:10]),
        "request1_low_confidence_count": sum(s == "조건부" for p in test_products for s in status_matrix[p["product_id"]][:10]),
        "request1_status": "PARTIAL", "request2_expected_fields": 56,
        "request2_pass_fields": sum(s == "가능" for p in test_products for s in status_matrix[p["product_id"]][10:]),
        "request2_missing_count": sum(s == "불가능" for p in test_products for s in status_matrix[p["product_id"]][10:]),
        "request2_low_confidence_count": sum(s == "조건부" for p in test_products for s in status_matrix[p["product_id"]][10:]),
        "request2_status": "PARTIAL", "duplicate_product_count": 0, "set_id_missing_count": 1,
        "review_sample_count": 8, "review_pass_count": 0, "review_pass_rate": "0.0%", "batch_status": "POC_COMPLETE_REVIEW_REQUIRED",
        "owner": "", "updated_at": "2026-07-20T18:00:00+09:00", "note": "API·HTML·OCR 전부 실행; 17개 요구사항×8상품=136 최종판정, 방법근거 408행",
    }
    qa_rows = [qa_headers, [qa_row.get(header, "") for header in qa_headers]]
    qa_groups = groups(qa_headers,
        r1=("request1_expected_fields", "request1_pass_fields", "request1_missing_count", "request1_low_confidence_count", "request1_status"),
        r2=("request2_expected_fields", "request2_pass_fields", "request2_missing_count", "request2_low_confidence_count", "request2_status"),
        both=("scope_id",))

    crm_headers = [
        "customer_key_hash", "consent_yn", "consent_purpose", "purchase_product_ids", "preferred_product_ids",
        "preference_tags", "owned_product_ids", "crm_segment", "source_system", "effective_at", "expires_at", "updated_at", "note",
    ]
    crm_rows = [crm_headers]

    semantic_headers = [
        "product_id", "source_url", "search_document", "normalized_terms", "synonyms", "color_terms", "material_terms",
        "style_terms", "room_terms", "use_purpose_terms", "embedding_model", "embedding_id", "index_name", "index_status",
        "parser_version", "indexed_at",
    ]
    semantic_rows = [semantic_headers]
    for product in test_products:
        semantic = {
            "product_id": product["product_id"], "source_url": product["url"], "search_document": product["semantic"],
            "normalized_terms": "|".join(filter(None, [product["small"], product["brand"], product["colors_normalized"], product["style"]])),
            "synonyms": product["colors_raw"], "color_terms": product["colors_normalized"],
            "material_terms": product["tag_material"], "style_terms": product["style"], "room_terms": product["spaces"],
            "use_purpose_terms": product["tag_use"], "embedding_model": "TBD", "embedding_id": "",
            "index_name": "homestyle_poc", "index_status": "DOCUMENT_READY", "parser_version": "pdp-parser-poc-0.4.0",
            "indexed_at": "",
        }
        semantic_rows.append([semantic.get(header, "") for header in semantic_headers])

    codebook_rows = [
        ["code_group", "code", "설명"],
        ["색상", "REQUEST_1_BLUE", "파란색: 요청 1 전용"],
        ["색상", "REQUEST_2_GREEN", "초록색: 요청 2 전용"],
        ["색상", "BOTH_PURPLE", "보라색: 요청 1·2 공통"],
        ["coverage", "Y", "요구사항에 대응하는 시트/필드가 양식에 존재"],
        ["coverage", "N", "양식에 필드가 없음"],
        ["assessment", "가능", "현재 공개 PDP/API/이미지에서 값을 확보하거나 요구된 추론을 수행 가능"],
        ["assessment", "조건부", "OCR/VISION/규칙/외부 문서 또는 검수가 있어야 제공 가능"],
        ["assessment", "불가능", "현재 공개 원천에 값이 없거나 개인정보 등 별도 권한 원천이 필요"],
        ["assessment", "해당없음", "해당 상품 유형에는 요구사항이 적용되지 않음; 성공률 분모에서 제외"],
        ["method_result", "FOUND", "해당 방법에서 요구 필드를 직접 확보"],
        ["method_result", "PARTIAL", "해당 방법에서 일부 값 또는 보강 텍스트만 확보"],
        ["method_result", "NOT_FOUND", "방법은 실행했으나 해당 요구 필드가 검출되지 않음"],
        ["method_result", "N_A", "상품 유형에 요구사항이 적용되지 않음"],
        ["value_status", "EXACT", "API/HTML/문서에 명시된 원천값"],
        ["value_status", "EXACT_MERGED", "둘 이상의 원천값을 병합한 명시값"],
        ["value_status", "INFERRED", "VISION/규칙/모델 추론값"],
        ["value_status", "MISSING", "원천에 값이 없음"],
        ["value_status", "N_A", "해당 상품에 적용되지 않음"],
        ["value_status", "ERROR", "수집 또는 파싱 실패"],
        ["extract_method", "API", "구조화 JSON API"],
        ["extract_method", "HTML", "PDP HTML 텍스트/메타데이터"],
        ["extract_method", "OCR", "상세 이미지 문자 인식"],
        ["extract_method", "VISION", "제품/공간 이미지 분석"],
        ["extract_method", "RULE", "카테고리/용어 규칙"],
        ["extract_method", "EXTERNAL", "CRM 등 별도 원천"],
        ["check_status", "PASS", "해당 요청의 필수 필드 충족"],
        ["check_status", "PARTIAL", "일부 결측/저신뢰/외부연계 필요"],
        ["check_status", "FAIL", "필수 파싱 실패"],
        ["parse_status", "SUCCESS", "필수 파싱 완료"],
        ["parse_status", "PARTIAL", "일부 필드 결측 또는 저신뢰"],
        ["parse_status", "ERROR", "상품 단위 파싱 실패"],
        ["input_status", "READY", "파싱 가능"],
        ["input_status", "SKIP_ZERO", "현재 제품 수 0"],
        ["input_status", "SAMPLE", "양식 예시 행"],
        ["review_status", "UNREVIEWED", "미검수"],
        ["review_status", "APPROVED", "승인"],
        ["review_status", "REJECTED", "반려"],
        ["confidence", "0.95~1.00", "명시값/높은 신뢰"],
        ["confidence", "0.85~0.94", "자동 승인 후보"],
        ["confidence", "0.00~0.84", "수동 검수 대상"],
    ]

    xlsx_writer.OUTPUT = OUTPUT
    xlsx_writer.SHEETS = [
        ("00_통계_DASHBOARD", dashboard_rows, widths(dashboard_headers)),
        ("00_통계_소스증분", source_increment_rows, widths(source_increment_headers)),
        ("00_소스증분_상품별", source_product_rows, widths(source_product_headers)),
        ("00_통계_추가요구", new_requirement_rows, widths(new_requirement_headers), ["COMMON"] * len(new_requirement_headers), new_requirement_row_groups),
        ("00_추가요구_상품판정", new_requirement_detail_rows, widths(new_requirement_detail_headers), ["COMMON"] * len(new_requirement_detail_headers), new_requirement_detail_row_groups),
        ("00_OCR엔진_요약", ocr_compare_rows, widths(ocr_compare_headers)),
        ("00_OCR정답_검증", ocr_gold_rows, widths(ocr_gold_headers), ["COMMON"] * len(ocr_gold_headers), ocr_gold_row_groups),
        ("00_METHOD_EXECUTION", method_execution_rows, widths(method_execution_headers)),
        ("00_통계_요청별", request_stat_rows, widths(request_stat_headers), ["COMMON"] * len(request_stat_headers), request_stat_row_groups),
        ("00_테스트_판정상세", detail_rows, widths(detail_headers), ["COMMON"] * len(detail_headers), detail_row_groups),
        ("00_사용안내", guide_rows, [24, 120], ["COMMON", "COMMON"], guide_row_groups),
        ("00_요청사항_COVERAGE", requirement_rows, widths(requirement_headers), ["COMMON"] * len(requirement_headers), requirement_row_groups),
        ("01_SCOPE_대상", scope_rows, widths(scope_headers), scope_groups),
        ("02_SCOPE_제외검증", _excluded_rows(excluded), [20, 34, 10, 10, 22, 28, 32, 18, 16, 16, 16, 28]),
        ("03_상품URL_INPUT", input_rows, widths(input_headers), input_groups),
        ("04_상품RESULT", result_rows, widths(result_headers), result_groups),
        ("04_상품확장RESULT", extended_result_rows, widths(extended_result_headers), extended_result_groups),
        ("05_FIELD_EVIDENCE", evidence_rows, widths(evidence_headers), ["COMMON"] * len(evidence_headers), evidence_row_groups),
        ("05_추가필드_EVIDENCE", deep_evidence_rows, widths(deep_evidence_headers), ["COMMON"] * len(deep_evidence_headers), deep_evidence_row_groups),
        ("05_OCR_IMAGE_LOG", ocr_image_rows, widths(ocr_image_headers)),
        ("05_OCR엔진_이미지", ocr_engine_image_rows, widths(ocr_engine_image_headers)),
        ("06_SPACE_RELATION", relation_rows, widths(relation_headers), ["R2"] * len(relation_headers)),
        ("07_QA_BATCH", qa_rows, widths(qa_headers), qa_groups),
        ("08_CRM_INPUT", crm_rows, widths(crm_headers), ["R2"] * len(crm_headers)),
        ("09_SEMANTIC_SEARCH", semantic_rows, widths(semantic_headers), ["R2"] * len(semantic_headers)),
        ("10_CODEBOOK", codebook_rows, [24, 34, 88]),
        ("11_RAW_FIELD_INVENTORY", raw_inventory_rows, widths(raw_inventory_headers)),
    ]
    xlsx_writer.build_xlsx()

    print(f"output={OUTPUT}")
    print(f"selected={len(selected)} excluded={len(excluded)}")
    print(f"homestyle_selected={len(homestyle)} public_count={sum(x['public_count'] for x in homestyle)}")
    print(f"appliance_selected={len(appliance)} public_count={sum(x['public_count'] for x in appliance)}")
    print(f"appliance_admin={sum((x['admin_count'] or 0) for x in appliance)} appliance_bundle={sum((x['bundle_count'] or 0) for x in appliance)}")
    print(f"representative_test=8 products x 17 requirements = {len(all_statuses)}")
    print(f"assessment={overall} applicable={overall_applicable} strict={overall['가능'] / overall_applicable:.1%} conditional_included={(overall['가능'] + overall['조건부']) / overall_applicable:.1%}")
    print(f"deep_fields=API {deep_api['usable']}/{deep_api['total_field_cells']} ({deep_api['usable_coverage_pct']}%) -> API+HTML {deep_html['usable']}/{deep_html['total_field_cells']} ({deep_html['usable_coverage_pct']}%) -> final {deep_paddle['usable']}/{deep_paddle['total_field_cells']} ({deep_paddle['usable_coverage_pct']}%)")


if __name__ == "__main__":
    build_template()
