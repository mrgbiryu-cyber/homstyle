from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile


OUTPUT = Path(__file__).with_name("homestyle_PDP_파싱_제공가능성_검토.xlsx")

PDP_URL = "https://homestyle.lge.co.kr/item?productId=G25070005743"
GOODS_API = "https://livingapi.lge.co.kr/itemsvc/ajax/v1/pdp/goods/G25070005743?epFlagYn=N"
SPACE_API = "https://livingapi.lge.co.kr/displaysvc/ajax/v1/collection/pdp/space-recommendation?goodsId=G25070005743"
PACKAGE_API = "https://livingapi.lge.co.kr/itemsvc/ajax/v1/goods/packages?goodsId=G25070005743"
DETAIL_IMAGE = "https://image.guud.com/mall/DESIGN/PRODUCT/casamia/2023/10/levor/02_levor_Info_4people+stool.jpg"
MAIN_IMAGE = "https://static-store.lge.co.kr/goods/org/073251020000078073.jpg?aw=1&ah=1&rw=800&rh=800"


SHEETS = [
    (
        "검토요약",
        [
            ["항목", "내용"],
            ["검토일", "2026-07-20"],
            ["샘플 product_id", "G25070005743"],
            ["샘플 PDP", PDP_URL],
            ["결론", "구조화 API + HTML 파싱 + OCR/이미지 분석을 조합하면 대부분 시트화 가능"],
            ["핵심 제한", "세트 구성 실제 제품 ID, 개인 CRM, 원천에 없는 벽 부착 높이는 URL만으로 생성 불가"],
            ["권장 우선순위", "API(JSON) → HTML 텍스트 → 상세 이미지 OCR → 이미지/규칙 추론"],
            ["3D 제작 주의", "기본 카탈로그형 모델링 자료로 활용 가능. 제조/CAD 정밀도에는 다각도·부품별 치수·도면 추가 필요"],
            ["일정 전제", "8월 일정은 소량 PoC 기준. 전체 제공일은 상품 수, 카테고리, 호출 승인, 목표 정확도 확정 후 산정"],
            ["운영 주의", "AJAX API는 공식 외부 연동 명세가 확인되지 않아 변경 가능. 대량 수집 전 호출/콘텐츠 이용 권한 확인"],
        ],
        [22, 110],
    ),
    (
        "요청1_3DAsset",
        [
            ["구분", "항목", "설명", "제공가능여부(O/X)", "2026년 8월 일정", "샘플 추출값", "추출방식", "신뢰도", "비고"],
            ["필수", "PDP 이미지", "정면45도 또는 정면 이미지", "O", "8/7 PoC", MAIN_IMAGE, "API+VISION", "높음", "대표 이미지 정면 판정. 배경/소품 포함 여부 별도 품질검사"],
            ["필수", "사이즈 정보", "W x D x H", "O", "8/7 PoC", "소파 W2910×D1020×H730~910mm; 스툴 W740×D660×H410mm", "API+OCR", "높음", "API 고시는 소파 H910, 상세 이미지에는 가변 높이 H730~910"],
            ["필수", "분류", "중카테고리, 소카테고리", "O", "8/7 PoC", "가구 > 소파 > 일반소파", "API", "높음", "카테고리 ID도 제공"],
            ["필수", "배치 추천 공간 리스트", "리빙룸, 베드룸 등", "O(조건부)", "8/14 PoC", "거실", "API+RULE", "높음", "공간 추천 API 우선, 미존재 시 카테고리-공간 매핑"],
            ["필수", "브랜드명", "", "O", "8/7 PoC", "까사미아 (2500000123)", "API", "높음", ""],
            ["필수", "제품 색상", "", "O", "8/7 PoC", "카멜브라운, 크림아이보리", "API+OCR", "높음", "옵션/고시 우선, 이미지 색상은 보조"],
            ["가전 추가 필수", "설치 타입 구분", "TV, 식기세척기 등 빌트인/스탠딩", "O(조건부)", "8/14 PoC", "샘플은 바닥 배치형 가구/업체 설치상품", "API+RULE", "중간", "가전 샘플 검증 필요; 명시값이 없으면 추론값 표시"],
            ["세트상품 필수", "세트 구성 ID 리스트", "실제 제품 ID 리스트", "X", "제공 불가", "구성명: 4인 소파+스툴 / 실제 구성 ID: 없음", "API", "높음", "패키지 API가 빈 배열. 원천 ID 매핑 필요"],
            ["옵션", "배치 가능 위치", "벽, 천장 등", "O(조건부)", "8/14 PoC", "바닥", "RULE+VISION", "중간", "카테고리 규칙+이미지 추론"],
            ["옵션", "벽면부착 시 추천 높이", "", "X(원천 없을 때)", "조건부", "해당 없음", "MANUAL", "높음", "매뉴얼의 명시 치수만 사용; 안전상 이미지 추정 금지"],
        ],
        [16, 26, 34, 21, 20, 62, 18, 12, 58],
    ),
    (
        "요청2_추천분석",
        [
            ["번호", "구축 데이터", "설명", "제공가능여부(O/X)", "2026년 8월 일정", "샘플 결과", "데이터 성격", "비고"],
            [1, "가구 상품 설명서 자동 태깅", "기본/제조/재질/규격/구성품/색상/사용목적/조립/안전/취급주의/품질보증/인증/판매정보", "O", "8/14 PoC", "천연면피 소가죽, PVC 합성가죽, E0 합판, HR폼, 스틸; 1년 무상 A/S 등", "명시+일부추론", "조립·안전·취급은 상품별 결측 가능"],
            [2, "디자인 스타일 추론", "미니멀/직선형/곡선형/유기적/장식형/클래식/모듈형/스칸디나비안/컨템포러리 등", "O(추론)", "8/21 PoC", "컨템포러리/모던, 직선형 중심, 완만한 라운드, 웜 뉴트럴", "추론", "멀티라벨+확률 제공, 정답 데이터로 캘리브레이션"],
            [3, "공간 콘텐츠 자동태깅", "스타일/Mood/공간명/목적/색상톤/크기/포함제품/배치사유·위치", "O(조건부)", "8/21 PoC", "거실, 빈티지/레트로, 밝은 톤, 편안한 휴식; 포함제품 ID 목록", "명시+추론", "콘텐츠 이미지/URL 필요; 공간 크기는 실측 없으면 범주형 추정"],
            [4, "공간스타일 ↔ 제품 관계정보", "", "O", "8/21 PoC", "exhibitionId 2605001190 ↔ G25070005743; Vintage Living Room Style", "명시+추론", "기존 큐레이션 관계는 명시값, 미연결 상품은 유사도"],
            [5, "공간내 배치된 제품간 관계정보", "", "O(추론)", "8/31 PoC", "소파-티테이블-러그-실링팬-수납장-식물의 공존/공간 관계", "추론", "일부 상품은 좌표값 존재. 가림/누락 검수 필요"],
            [6, "개인별 구매/선호/보유/CRM 정보", "", "X(URL 기준)", "제공 불가", "없음", "내부 원천 필요", "로그인/주문/CRM과 적법한 결합 기준 필요"],
            [7, "비표준화 데이터 의미기반 검색·추천", "예: 적색/레드/붉은색 등", "O", "8/31 PoC", "카멜브라운→브라운/웜브라운; 크림아이보리→아이보리/오프화이트/베이지", "구축", "용어사전+임베딩/하이브리드 인덱스는 파싱과 별도 작업"],
        ],
        [10, 34, 70, 21, 20, 58, 18, 60],
    ),
    (
        "샘플_필드근거",
        [
            ["product_id", "field_name", "raw_value", "normalized_value", "value_status", "extract_method", "confidence", "evidence_url", "evidence_note"],
            ["G25070005743", "product_name", "[LG 구매혜택 상품권 5만증정] [by CASAMIA] 벨로씨 레보르 천연면피 가죽 소파 4인＋스툴", "벨로씨 레보르 천연면피 가죽 소파 4인+스툴", "EXACT", "API", 1.0, GOODS_API, "프로모션 문구 제거 가능"],
            ["G25070005743", "brand", "까사미아", "까사미아", "EXACT", "API", 1.0, GOODS_API, "brand_id=2500000123"],
            ["G25070005743", "category", "가구/소파/일반소파", "가구 > 소파 > 일반소파", "EXACT", "API", 1.0, GOODS_API, "3단계 카테고리"],
            ["G25070005743", "color", "카멜브라운, 크림아이보리", "카멜브라운|크림아이보리", "EXACT", "API", 1.0, GOODS_API, "구매 옵션과 고시 일치"],
            ["G25070005743", "sofa_size", "W2910 X D1020 X H910mm", "W2910×D1020×H730~910mm", "EXACT_MERGED", "API+OCR", 0.98, DETAIL_IMAGE, "상세 규격 이미지가 가변 높이를 보강"],
            ["G25070005743", "stool_size", "W740 X D660 X H410mm", "W740×D660×H410mm", "EXACT", "API+OCR", 1.0, DETAIL_IMAGE, ""],
            ["G25070005743", "recommended_space", "거실", "LIVING_ROOM", "EXACT", "API", 1.0, SPACE_API, "tagCd=DS0062"],
            ["G25070005743", "placement_surface", "소파 이미지/일반소파", "FLOOR", "INFERRED", "RULE+VISION", 0.99, PDP_URL, "가구 분류 및 이미지"],
            ["G25070005743", "design_style", "제품/공간 이미지", "CONTEMPORARY|MODERN|RECTILINEAR|SOFT_ROUNDED", "INFERRED", "VISION", 0.82, PDP_URL, "멀티라벨 추론 예시"],
            ["G25070005743", "space_style", "Vintage Living Room Style; 빈티지/레트로", "VINTAGE_RETRO", "EXACT", "API", 1.0, SPACE_API, "공간 추천 콘텐츠 명시값"],
            ["G25070005743", "set_component_names", "4인 소파 + 스툴", "SOFA_4SEAT|STOOL", "EXACT", "API", 1.0, GOODS_API, "구성명만 존재"],
            ["G25070005743", "set_component_product_ids", "[]", "", "MISSING", "API", 1.0, PACKAGE_API, "패키지 API 빈 배열"],
            ["G25070005743", "wall_mount_height", "", "", "N_A", "RULE", 1.0, PDP_URL, "소파는 벽 부착 제품이 아님"],
            ["G25070005743", "crm", "", "", "MISSING", "NONE", 1.0, PDP_URL, "공개 PDP에서 취득 불가"],
        ],
        [18, 28, 52, 52, 18, 18, 13, 95, 48],
    ),
    (
        "권장_출력스키마",
        [
            ["컬럼", "설명", "예시"],
            ["product_id", "상품 ID", "G25070005743"],
            ["source_url", "원본 PDP URL", PDP_URL],
            ["field_name", "필드명", "sofa_size"],
            ["raw_value", "원문 값", "W2910 X D1020 X H910mm"],
            ["normalized_value", "정규화 값", "W2910×D1020×H730~910mm"],
            ["unit", "단위", "mm"],
            ["value_status", "EXACT/EXACT_MERGED/INFERRED/MISSING/N_A", "EXACT_MERGED"],
            ["confidence", "0~1 신뢰도", "0.98"],
            ["extract_method", "API/HTML/OCR/VISION/RULE", "API+OCR"],
            ["evidence_url", "근거 URL", DETAIL_IMAGE],
            ["evidence_text", "근거 원문/요약", "상세 규격 이미지에 H730~910mm 표기"],
            ["extracted_at", "추출 시각", "2026-07-20T11:30:43+09:00"],
            ["parser_version", "파서 버전", "pdp-parser-0.1.0"],
            ["review_status", "미검수/승인/반려", "미검수"],
        ],
        [25, 58, 100],
    ),
]


def col_name(index: int) -> str:
    result = ""
    while index:
        index, rem = divmod(index - 1, 26)
        result = chr(65 + rem) + result
    return result


def cell_xml(row: int, col: int, value, style: int) -> str:
    ref = f"{col_name(col)}{row}"
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}" s="{style}" t="n"><v>{value}</v></c>'
    text = "" if value is None else str(value)
    preserve = ' xml:space="preserve"' if text[:1].isspace() or text[-1:].isspace() else ""
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t{preserve}>{escape(text)}</t></is></c>'


def sheet_xml(
    rows: list[list],
    widths: list[int],
    column_groups: list[str] | None = None,
    row_groups: list[str] | None = None,
) -> str:
    max_cols = max(len(row) for row in rows)
    header_styles = {"COMMON": 1, "R1": 3, "R2": 5, "BOTH": 7}
    body_styles = {"COMMON": 2, "R1": 4, "R2": 6, "BOTH": 8}
    row_xml = []
    for r_idx, row in enumerate(rows, start=1):
        height = 30 if r_idx == 1 else 45
        cells_list = []
        for c_idx, value in enumerate(row, start=1):
            column_group = (
                column_groups[c_idx - 1]
                if column_groups and c_idx - 1 < len(column_groups)
                else "COMMON"
            )
            if r_idx == 1:
                style = header_styles.get(column_group, header_styles["COMMON"])
            else:
                row_group = (
                    row_groups[r_idx - 2]
                    if row_groups and r_idx - 2 < len(row_groups)
                    else "COMMON"
                )
                effective_group = row_group if row_group != "COMMON" else column_group
                style = body_styles.get(effective_group, body_styles["COMMON"])
            cells_list.append(cell_xml(r_idx, c_idx, value, style))
        cells = "".join(cells_list)
        row_xml.append(f'<row r="{r_idx}" ht="{height}" customHeight="1">{cells}</row>')
    cols_xml = "".join(
        f'<col min="{idx}" max="{idx}" width="{width}" customWidth="1"/>'
        for idx, width in enumerate(widths, start=1)
    )
    end_ref = f"{col_name(max_cols)}{len(rows)}"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<dimension ref="A1:{end_ref}"/>'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        f'<cols>{cols_xml}</cols>'
        f'<sheetData>{"".join(row_xml)}</sheetData>'
        f'<autoFilter ref="A1:{end_ref}"/>'
        '<pageMargins left="0.3" right="0.3" top="0.5" bottom="0.5" header="0.2" footer="0.2"/>'
        '<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>'
        '</worksheet>'
    )


def build_xlsx() -> None:
    sheet_count = len(SHEETS)
    content_types = ''.join(
        f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i in range(1, sheet_count + 1)
    )
    workbook_sheets = ''.join(
        f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>'
        for i, (name, *_) in enumerate(SHEETS, start=1)
    )
    workbook_rels = ''.join(
        f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
        for i in range(1, sheet_count + 1)
    )
    workbook_rels += f'<Relationship Id="rId{sheet_count + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'

    styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="10"/><name val="맑은 고딕"/><family val="2"/></font>
    <font><b/><color rgb="FFFFFFFF"/><sz val="10"/><name val="맑은 고딕"/><family val="2"/></font>
  </fonts>
  <fills count="9">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF595959"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF2F75B5"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9EAF7"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF548235"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE2F0D9"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF7030A0"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE4DFEC"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFD9D9D9"/></left><right style="thin"><color rgb="FFD9D9D9"/></right><top style="thin"><color rgb="FFD9D9D9"/></top><bottom style="thin"><color rgb="FFD9D9D9"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="9">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="1" fillId="5" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="1" fillId="7" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="8" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    with ZipFile(OUTPUT, "w", ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
            '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
            f'{content_types}</Types>',
        )
        archive.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            '</Relationships>',
        )
        archive.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<sheets>{workbook_sheets}</sheets><calcPr calcId="191029"/></workbook>',
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{workbook_rels}</Relationships>',
        )
        archive.writestr("xl/styles.xml", styles)
        for index, sheet in enumerate(SHEETS, start=1):
            _, rows, widths, *style_config = sheet
            column_groups = style_config[0] if len(style_config) >= 1 else None
            row_groups = style_config[1] if len(style_config) >= 2 else None
            archive.writestr(
                f"xl/worksheets/sheet{index}.xml",
                sheet_xml(rows, widths, column_groups, row_groups),
            )
        archive.writestr(
            "docProps/core.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:title>홈스타일 PDP 파싱 제공 가능성 검토</dc:title><dc:creator>Codex</dc:creator>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>'
            '</cp:coreProperties>',
        )
        archive.writestr(
            "docProps/app.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            '<Application>Codex</Application><AppVersion>1.0</AppVersion>'
            f'<TitlesOfParts><vt:vector size="{sheet_count}" baseType="lpstr">'
            + ''.join(f'<vt:lpstr>{escape(name)}</vt:lpstr>' for name, *_ in SHEETS)
            + '</vt:vector></TitlesOfParts></Properties>',
        )


if __name__ == "__main__":
    build_xlsx()
    print(OUTPUT)
