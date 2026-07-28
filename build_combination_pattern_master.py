from __future__ import annotations

import re
import sqlite3
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from zipfile import ZipFile


ROOT = Path(__file__).resolve().parent
DB_PATH = ROOT / "homestyle_bulk_run" / "homestyle_bulk.sqlite"
WORKBOOK_PATH = (
    ROOT
    / "홈스타일_비음영대상군_전체상품_요구필드_대량결과_패턴상태.xlsx"
)
OUTPUT_PATH = ROOT / "조합상품_3D변환_패턴목록_2026-07-27.md"

SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


PATTERN_DEFINITIONS: list[tuple[str, str, str, str]] = [
    ("N01", "단일상품 처리 — 상품명의 `+`가 모델명 일부", "비조합 단일 Asset", "`아카이브+`, `Lotus+`, `TERRACE+`"),
    ("N02", "단일상품 처리 — `+`가 작품명·고유명에 포함", "비조합 단일 Asset", "작품명·컬렉션명 내부 기호"),
    ("N03", "단일상품 처리 — `+`가 색상·소재를 연결", "비조합 단일 Asset", "색상+색상, 소재+소재"),
    ("N04", "단일상품 처리 — `+`가 옵션 방식을 연결", "비조합 단일 Asset", "바퀴+글라이드, lite+ option"),
    ("N05", "단일상품 처리 — `+`가 확장 기능을 표시", "비조합 단일 Asset", "`(+Extension)`, 확장형 기능"),
    ("N06", "단일상품 처리 — 여러 기능이 하나의 본체에 통합", "비조합 단일 Asset", "조명+스피커처럼 한 본체"),
    ("N07", "상품 구성 결정 필요 — 세트·패키지가 마케팅 명칭", "구성 근거 없으면 비조합", "제목에만 세트·패키지"),
    ("S01", "단일상품 통합처리 — 침대 프레임과 매트리스를 하나로 제작", "단일 Asset", "침대/프레임+매트리스·토퍼"),
    ("S02", "단일상품 결합처리 — 상판과 다리를 조립해 하나로 사용", "단일 Asset", "조립 후 하나의 제품"),
    ("S03", "단일상품 결합처리 — 소파 모듈을 결합해 하나로 사용", "단일 Asset", "MOD.10+25, A1+A0+A1"),
    ("S04", "상품 구성 결정 필요 — 수납장·옷장 모듈의 결합 여부 확인", "단일 Asset 후보", "연속 배치·고정 결합 모듈"),
    ("S05", "상품 구성 결정 필요 — 패널·가드가 본체에 고정되는지 확인", "단일 Asset 후보", "본체에 고정되는 구조"),
    ("S06", "단일상품 결합처리 — 여러 요소가 제품 내부에 결합", "단일 Asset", "PL박스+책장형, 도어+서랍"),
    ("Q01", "복수상품 수량처리 — 동일 가구를 여러 개 제공", "Asset 1종+quantity", "1+1, 2P/4P/6P"),
    ("Q02", "복수상품 수량처리 — 동일 사이즈 침대를 여러 개 제공", "Asset 1종+quantity", "Q+Q, SS+SS"),
    ("Q03", "복수상품 수량처리 — 동일 월데코·선반을 여러 개 제공", "Asset 1종+quantity", "동일 패널·선반 N개"),
    ("M01", "복수상품 분리처리 — 테이블과 의자를 함께 제공", "`component_seq` 분리", "식탁/테이블과 의자"),
    ("M02", "복수상품 분리처리 — 테이블·벤치·의자를 함께 제공", "`component_seq` 분리", "테이블·벤치·의자 혼합"),
    ("M03", "복수상품 분리처리 — 서로 다른 크기의 침대를 함께 제공", "`component_seq` 분리", "Q+SS 등 크기가 다른 침대"),
    ("M04", "복수상품 분리처리 — 침대와 협탁을 함께 제공", "`component_seq` 분리", "독립 배치 협탁"),
    ("M05", "복수상품 분리처리 — 소파·체어와 스툴을 함께 제공", "`component_seq` 분리", "독립 이동 가능한 스툴"),
    ("M06", "복수상품 분리처리 — 책상과 의자를 함께 제공", "`component_seq` 분리", "책상과 독립 의자"),
    ("M07", "복수상품 분리처리 — 책상·책장·수납장·침대를 함께 제공", "`component_seq` 분리", "학생방·서재 풀패키지"),
    ("M08", "복수상품 분리처리 — 유아동 테이블·의자·수납장을 함께 제공", "`component_seq` 분리", "키즈 가구 패키지"),
    ("M09", "복수상품 분리처리 — 화장대·거울·스툴·수납장을 함께 제공", "`component_seq` 분리", "독립 배치 화장대 구성"),
    ("M10", "복수상품 분리처리 — 테이블램프와 플로어램프를 함께 제공", "`component_seq` 분리", "배치 위치가 다른 조명"),
    ("M11", "복수상품 분리처리 — 크기가 다른 테이블을 함께 제공", "`component_seq` 분리", "네스팅 테이블"),
    ("M12", "상품 구성 결정 필요 — 수납장·옷장 모듈의 독립 배치 여부 확인", "`component_seq` 분리 후보", "독립 장 모듈 나란히 배치"),
    ("M13", "복수상품 분리처리 — 야외 테이블·의자·벤치를 함께 제공", "`component_seq` 분리", "야외 테이블·의자·벤치"),
    ("A01", "부속품 정책 결정 필요 — 쿠션·헤드레스트를 별도 제작할지 결정", "별도 제작 정책 확인", "소파+쿠션, 헤드레스트"),
    ("A02", "부속품 정책 결정 필요 — 가드·패널·도어·서랍을 별도 제작할지 결정", "별도 제작 정책 확인", "본체 부착형 구조물"),
    ("A03", "부속품 정책 결정 필요 — 후크·옷걸이·파우치를 별도 제작할지 결정", "별도 제작 정책 확인", "행거+후크, 체어+파우치"),
    ("A04", "부속품 정책 결정 필요 — 전구·조명·콘센트를 별도 제작할지 결정", "별도 제작 정책 확인", "책상 액세서리 팩"),
    ("A05", "단일상품 통합처리 — 매트리스·토퍼를 침대 본체에 포함", "침대 정책상 본체 통합", "침대와 함께 판매"),
    ("O01", "옵션 분리처리 — 옵션에 따라 상품 크기가 달라짐", "`option_seq` 분리", "SS/Q/K, 1200/1600/1800"),
    ("O02", "옵션 수량처리 — 옵션에 따라 제공 수량이 달라짐", "quantity 또는 구성 행", "2인/4인/6인, 2P/4P"),
    ("O03", "옵션 분리처리 — 옵션에 따라 상품 구성이 달라짐", "형상별 `option_seq`", "벤치형/체어형, 도어형/서랍형"),
    ("O04", "옵션 분리처리 — 옵션에 따라 방향·배치 형태가 달라짐", "`option_seq` 분리", "좌형/우형, 코너형 A/B"),
    ("O05", "옵션 분리 여부 확인 — 매트리스 종류·경도에 따른 외형 변경 확인", "옵션 메타, 외형 변경 시만 분리", "하드/소프트, 토퍼 종류"),
    ("O06", "옵션 통합처리 — 색상·소재만 다르고 외형은 동일", "동일 형상, 규격 행 미분리", "색상·마감만 변경"),
]


LEGACY_TO_FINAL = {
    "N01_MODEL_SUFFIX_PLUS": "N01",
    "N02_OPTION_TYPE_JOIN": "N04",
    "N03_ARTWORK_TITLE_PLUS": "N02",
    "N04_COLOR_MATERIAL_JOIN": "N03",
    "N05_EXTENSION_FEATURE": "N05",
    "N06_INTEGRATED_FUNCTION": "N06",
    "R01_INTERNAL_CONFIGURATION": "S06",
    "Y01_MULTI_BED_BUNDLE": "Q02 / M03 / M04 (상품명별)",
    "Y02_BED_NIGHTSTAND": "M04",
    "Y03_BED_PANEL_GUARD": "S05 + A02",
    "Y04_BED_SLEEP_SYSTEM": "S01 + A05",
    "Y05_KIDS_BED_MATTRESS": "S01 + A05 (+M07)",
    "Y06_KIDS_BED_ACCESSORY": "S05 + A02",
    "Y07_TABLE_CHAIR_BENCH": "M01 / M02 (벤치 유무)",
    "Y08_TABLE_TOP_LEG": "S02",
    "Y09_MODULAR_SOFA_UNITS": "S03",
    "Y10_SOFA_ACCESSORY": "A01",
    "Y11_CABINET_STORAGE_MODULES": "S04 + M12",
    "Y12_HANGER_ACCESSORY": "A03",
    "Y13_WARDROBE_MODULES": "S04 + M12",
    "Y14_CHAIR_BUNDLE_1PLUS1": "Q01",
    "Y15_DESK_ACCESSORY_PACK": "A04",
}


LEGACY_DISPLAY_NAMES = {
    "N01_MODEL_SUFFIX_PLUS": "단일상품 처리 — 상품명의 +가 모델명 일부",
    "N02_OPTION_TYPE_JOIN": "단일상품 처리 — +가 옵션 방식을 연결",
    "N03_ARTWORK_TITLE_PLUS": "단일상품 처리 — +가 작품명·고유명에 포함",
    "N04_COLOR_MATERIAL_JOIN": "단일상품 처리 — +가 색상·소재를 연결",
    "N05_EXTENSION_FEATURE": "단일상품 처리 — +가 확장 기능을 표시",
    "N06_INTEGRATED_FUNCTION": "단일상품 처리 — 여러 기능이 하나의 본체에 통합",
    "R01_INTERNAL_CONFIGURATION": "단일상품 결합처리 — 여러 요소가 제품 내부에 결합",
    "Y01_MULTI_BED_BUNDLE": "상품 구성 세부결정 필요 — 복수 침대의 크기와 구성 관계 확인",
    "Y02_BED_NIGHTSTAND": "복수상품 분리처리 — 침대와 협탁을 함께 제공",
    "Y03_BED_PANEL_GUARD": "부속품 정책 결정 필요 — 침대와 패널·가드를 함께 제공",
    "Y04_BED_SLEEP_SYSTEM": "단일상품 통합처리 — 침대 프레임과 매트리스를 하나로 제작",
    "Y05_KIDS_BED_MATTRESS": "상품 구성 세부결정 필요 — 유아동 침대와 매트리스·토퍼 구성 확인",
    "Y06_KIDS_BED_ACCESSORY": "상품 구성 세부결정 필요 — 유아동 침대와 가드·책상 구성 확인",
    "Y07_TABLE_CHAIR_BENCH": "복수상품 분리처리 — 테이블과 의자·벤치를 함께 제공",
    "Y08_TABLE_TOP_LEG": "단일상품 결합처리 — 상판과 다리를 조립해 하나로 사용",
    "Y09_MODULAR_SOFA_UNITS": "단일상품 결합처리 — 소파 모듈을 결합해 하나로 사용",
    "Y10_SOFA_ACCESSORY": "부속품 정책 결정 필요 — 소파와 쿠션·헤드레스트를 함께 제공",
    "Y11_CABINET_STORAGE_MODULES": "상품 구성 결정 필요 — 수납장 모듈의 결합·독립 배치 여부 확인",
    "Y12_HANGER_ACCESSORY": "부속품 정책 결정 필요 — 행거와 후크·바지걸이를 함께 제공",
    "Y13_WARDROBE_MODULES": "상품 구성 결정 필요 — 옷장 모듈의 결합·독립 배치 여부 확인",
    "Y14_CHAIR_BUNDLE_1PLUS1": "복수상품 수량처리 — 동일 의자를 1+1로 제공",
    "Y15_DESK_ACCESSORY_PACK": "부속품 정책 결정 필요 — 책상과 전구·콘센트 부속품을 함께 제공",
}


DIMENSION_PATTERN_DEFINITIONS = {
    "D00": {
        "name": "상품 구성 선결 필요 — 상품 구성 확정 후 규격 행을 결정",
        "example": "침대+협탁+침대, 식탁+의자, 소파+스툴",
        "candidate_structure": "규격과 상품 구성 패턴을 모두 확인해야 함",
        "decision": "N/S/Q/M/A/O를 먼저 정한 뒤 대표·옵션·구성상품 규격 확정",
    },
    "D01": {
        "name": "규격 결정 필요 — 완전한 W/D/H 후보 1세트를 제품 규격으로 확인",
        "example": "의자 W/D/H 1세트, 선반 L/D/H 1세트",
        "candidate_structure": "완전 W/D/H 후보가 1세트만 존재",
        "decision": "제품명·모델명과 일치하고 배송 수치가 아니면 대표 규격으로 확정",
    },
    "D02": {
        "name": "규격 결정 필요 — 완전 후보와 추가 부분 후보를 구분",
        "example": "제품 W/D/H 1세트+두께·오차 등 부분 숫자",
        "candidate_structure": "완전 W/D/H 1세트와 추가 부분 후보 존재",
        "decision": "완전 후보를 우선하고 중복·부속품·배송·오탐 후보를 제외",
    },
    "D03": {
        "name": "규격 결정 필요 — 여러 완전 후보를 대표·옵션·타제품으로 구분",
        "example": "A/B/C 사이즈 옵션, 2~3개 규격의 테이블",
        "candidate_structure": "완전 W/D/H가 여러 세트 존재",
        "decision": "대표 규격과 옵션별 규격을 선택하고 라인업·권장치·배송 수치를 제외",
    },
    "D04": {
        "name": "규격 보강 필요 — 일부 축 후보 1세트의 비적용 축·재OCR 판단",
        "example": "원형 테이블 지름+높이, 월데코 W×H, 일부 축만 인식",
        "candidate_structure": "부분 규격 후보가 1세트만 존재",
        "decision": "2D·원형 비적용 축을 정하거나 실제 SIZE 영역을 재OCR",
    },
    "D05": {
        "name": "규격 보강 필요 — 여러 부분 후보를 이미지·옵션별로 그룹화",
        "example": "러그 여러 사이즈, 옵션별 W/H 부분 후보",
        "candidate_structure": "부분 규격 후보가 여러 세트 존재",
        "decision": "같은 이미지·옵션의 축만 결합하고 필요한 영역을 재OCR",
    },
}


LEGACY_PATTERN_EXAMPLES = {
    "N01_MODEL_SUFFIX_PLUS": "아카이브+, Lotus+, TERRACE+",
    "N02_OPTION_TYPE_JOIN": "바퀴+글라이드, lite+ option",
    "N03_ARTWORK_TITLE_PLUS": "작품명 A+B, 컬렉션명 내부의 +",
    "N04_COLOR_MATERIAL_JOIN": "레드+화이트, 패브릭+가죽",
    "N05_EXTENSION_FEATURE": "테이블(+Extension)",
    "N06_INTEGRATED_FUNCTION": "조명+스피커 일체형",
    "R01_INTERNAL_CONFIGURATION": "PL박스+책장형, 도어+서랍",
    "Y01_MULTI_BED_BUNDLE": "침대+협탁+침대, Q+SS, Q+Q",
    "Y02_BED_NIGHTSTAND": "침대+협탁",
    "Y03_BED_PANEL_GUARD": "침대+패널, 침대+가드",
    "Y04_BED_SLEEP_SYSTEM": "침대 프레임+매트리스",
    "Y05_KIDS_BED_MATTRESS": "유아동 침대+매트리스, 침대+토퍼",
    "Y06_KIDS_BED_ACCESSORY": "유아동 침대+가드+책상",
    "Y07_TABLE_CHAIR_BENCH": "식탁+의자, 테이블+벤치+의자",
    "Y08_TABLE_TOP_LEG": "상판+다리",
    "Y09_MODULAR_SOFA_UNITS": "소파 모듈 A1+A0+A1",
    "Y10_SOFA_ACCESSORY": "소파+쿠션, 소파+헤드레스트",
    "Y11_CABINET_STORAGE_MODULES": "수납장 80cm+수납장 120cm",
    "Y12_HANGER_ACCESSORY": "행거+후크, 행거+바지걸이",
    "Y13_WARDROBE_MODULES": "이불장+서랍장+행거장",
    "Y14_CHAIR_BUNDLE_1PLUS1": "의자 1+1, 스툴 1+1",
    "Y15_DESK_ACCESSORY_PACK": "책상+콘센트+조명+데스크패드",
}


REFERENCE_EXAMPLES = {
    "N07": "G25070001678",
    "Q03": "G25120026341",
    "M05": "G25070005743",
    "M06": "G26070048007",
    "M08": "G26030035346",
    "M09": "G26040040410",
    "M10": "G26040040410",
    "M11": "G26070047137",
    "M13": "G25080008909",
}


def excel_col_index(cell_ref: str) -> int:
    letters = re.match(r"[A-Z]+", cell_ref).group(0)
    value = 0
    for letter in letters:
        value = value * 26 + ord(letter) - 64
    return value - 1


def read_pattern_sheet() -> list[dict[str, str]]:
    ns = {"m": SHEET_NS, "r": REL_NS, "p": PKG_REL_NS}
    with ZipFile(WORKBOOK_PATH) as archive:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall("m:si", ns):
                shared.append(
                    "".join(node.text or "" for node in item.iter(f"{{{SHEET_NS}}}t"))
                )

        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relmap = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rels.findall("p:Relationship", ns)
        }
        target = ""
        for sheet in workbook.find("m:sheets", ns):
            if sheet.attrib["name"] == "08_패턴_상태":
                target = relmap[sheet.attrib[f"{{{REL_NS}}}id"]]
                break
        if not target:
            raise RuntimeError("08_패턴_상태 sheet not found")
        if not target.startswith("xl/"):
            target = "xl/" + target.lstrip("/")

        root = ET.fromstring(archive.read(target))
        raw_rows: list[dict[int, str]] = []
        max_index = 0
        for row in root.findall(".//m:sheetData/m:row", ns):
            values: dict[int, str] = {}
            for cell in row.findall("m:c", ns):
                index = excel_col_index(cell.attrib["r"])
                max_index = max(max_index, index)
                cell_type = cell.attrib.get("t")
                value_node = cell.find("m:v", ns)
                value = "" if value_node is None else (value_node.text or "")
                if cell_type == "s" and value:
                    value = shared[int(value)]
                elif cell_type == "inlineStr":
                    value = "".join(
                        node.text or "" for node in cell.findall(".//m:t", ns)
                    )
                values[index] = value
            raw_rows.append(values)

    header = [raw_rows[0].get(index, "") for index in range(max_index + 1)]
    return [
        {
            header[index]: row.get(index, "")
            for index in range(len(header))
            if header[index]
        }
        for row in raw_rows[1:]
    ]


def md_text(value: Any) -> str:
    return (
        str(value or "")
        .replace("|", "\\|")
        .replace("\r", " ")
        .replace("\n", " ")
        .strip()
    )


def codes(value: str) -> list[str]:
    if not value or value == "해당없음":
        return []
    return [part.strip() for part in value.split("|") if part.strip()]


def product_link(product_id: str) -> str:
    return (
        f"[{product_id}]"
        f"(https://homestyle.lge.co.kr/item?productId={product_id})"
    )


def number_text(value: Any) -> str:
    if value is None or value == "":
        return "-"
    number = float(value)
    return f"{number:g}"


def dimension_candidate_summary(
    candidate_rows: list[dict[str, Any]],
) -> tuple[str, str]:
    value_sets: list[tuple[Any, Any, Any]] = []
    labels: list[str] = []
    for row in candidate_rows:
        values = (row.get("w_mm"), row.get("d_mm"), row.get("h_mm"))
        if values in value_sets:
            continue
        value_sets.append(values)
        labels.append(
            f"W {number_text(values[0])} / "
            f"D {number_text(values[1])} / "
            f"H {number_text(values[2])}"
        )

    full_count = sum(
        all(value is not None for value in values) for values in value_sets
    )
    partial_count = len(value_sets) - full_count
    structure = f"완전 {full_count}세트 / 부분 {partial_count}세트"
    return structure, "; ".join(labels) if labels else "후보값 없음"


def action_for(pattern_codes: list[str], bed_policy: bool) -> str:
    if bed_policy:
        return "프레임+매트리스 단일 Asset; 실제 외형 사이즈 옵션만 `option_seq`"
    groups = {code[0] for code in pattern_codes if code}
    actions = []
    if "N" in groups:
        actions.append("비조합 단일 행")
    if "S" in groups:
        actions.append("조립 후 단일 Asset")
    if "Q" in groups:
        actions.append("동일 Asset+quantity")
    if "M" in groups:
        actions.append("`component_seq` 분리")
    if "A" in groups:
        actions.append("부속품 제작 정책 확인")
    if "O" in groups:
        if any(code in {"O01", "O03", "O04"} for code in pattern_codes):
            actions.append("형상 옵션 `option_seq`")
        if "O02" in pattern_codes:
            actions.append("수량 옵션 확인")
        if any(code in {"O05", "O06"} for code in pattern_codes):
            actions.append("비형상 옵션은 규격 행 미분리")
    return "; ".join(actions) or "패턴 확인 필요"


def build_document() -> tuple[str, dict[str, Any]]:
    pattern_rows = read_pattern_sheet()
    pattern_by_id = {row["상품 ID"]: row for row in pattern_rows}

    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    legacy_snapshot = connection.execute(
        """
        SELECT snapshot_id
        FROM stg_combination_candidate_pattern
        GROUP BY snapshot_id
        HAVING COUNT(*)=270
        ORDER BY MAX(built_at) DESC
        LIMIT 1
        """
    ).fetchone()[0]
    legacy_rows = [
        dict(row)
        for row in connection.execute(
            """
            SELECT *
            FROM stg_combination_candidate_pattern
            WHERE snapshot_id=?
            ORDER BY pattern_code, product_id
            """,
            (legacy_snapshot,),
        )
    ]
    bed_rows = {
        row["product_id"]: dict(row)
        for row in connection.execute(
            "SELECT * FROM vw_bed_asset_policy_current"
        )
    }
    current_candidates = {
        row["product_id"]: dict(row)
        for row in connection.execute(
            """
            SELECT *
            FROM vw_product_combination_current
            WHERE detection_status='CANDIDATE'
            """
        )
    }
    comparison_candidates: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in connection.execute(
        """
        SELECT *
        FROM vw_dimension_comparison_candidates_current
        ORDER BY product_id, comparison_no
        """
    ):
        comparison_candidates[row["product_id"]].append(dict(row))

    dimension_review_rows = [
        row
        for row in pattern_rows
        if row["규격 패턴 코드"] in DIMENSION_PATTERN_DEFINITIONS
    ]
    dimension_pattern_counts = Counter(
        row["규격 패턴 코드"] for row in dimension_review_rows
    )
    expected_dimension_patterns = {
        "D00": 76,
        "D01": 556,
        "D02": 388,
        "D03": 1476,
        "D04": 299,
        "D05": 616,
    }
    assert len(dimension_review_rows) == 3411
    assert dict(dimension_pattern_counts) == expected_dimension_patterns
    assert all(
        comparison_candidates[row["상품 ID"]] for row in dimension_review_rows
    )

    confirmed_summary = [
        dict(row)
        for row in connection.execute(
            """
            SELECT detection_rule, component_output_status, COUNT(*) AS n
            FROM vw_product_combination_current
            WHERE detection_status='CONFIRMED'
            GROUP BY detection_rule, component_output_status
            ORDER BY detection_rule, n DESC
            """
        )
    ]
    confirmed_examples = defaultdict(list)
    for row in connection.execute(
        """
        SELECT detection_rule, product_id, product_name
        FROM vw_product_combination_current
        WHERE detection_status='CONFIRMED'
        ORDER BY detection_rule, product_id
        """
    ):
        if len(confirmed_examples[row["detection_rule"]]) < 4:
            confirmed_examples[row["detection_rule"]].append(dict(row))
    connection.close()

    legacy_by_id = {row["product_id"]: row for row in legacy_rows}
    legacy_ids = set(legacy_by_id)
    bed_ids = set(bed_rows)
    candidate_ids = set(current_candidates)
    assert len(legacy_ids) == 270
    assert len(bed_ids) == 121
    assert len(candidate_ids) == 149
    assert not bed_ids & candidate_ids
    assert legacy_ids == bed_ids | candidate_ids

    master_rows = []
    code_products: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for product_id in sorted(legacy_ids):
        legacy = legacy_by_id[product_id]
        workbook = pattern_by_id[product_id]
        is_bed = product_id in bed_ids
        final_codes = (
            ["S01", "A05"]
            if is_bed
            else codes(workbook["조합 세부 패턴 코드"])
        )
        status = (
            "완료_침대단일Asset"
            if is_bed
            else "체크필요_조합패턴"
        )
        row = {
            "product_id": product_id,
            "product_name": legacy["product_name"],
            "legacy_code": legacy["pattern_code"],
            "legacy_name": legacy["pattern_name"],
            "display_pattern_name": LEGACY_DISPLAY_NAMES[legacy["pattern_code"]],
            "display_pattern_example": LEGACY_PATTERN_EXAMPLES[
                legacy["pattern_code"]
            ],
            "status": status,
            "final_codes": final_codes,
            "action": action_for(final_codes, is_bed),
            "output_status": workbook["산출 상태"],
            "dimension_pattern": workbook["규격 패턴 코드"],
        }
        master_rows.append(row)
        for code in final_codes:
            code_products[code].append(row)

    current_group_counts = Counter()
    for row in master_rows:
        if row["status"] != "체크필요_조합패턴":
            continue
        for group in {code[0] for code in row["final_codes"]}:
            current_group_counts[group] += 1
    expected_groups = {"N": 34, "S": 63, "Q": 9, "M": 59, "A": 29, "O": 127}
    assert dict(current_group_counts) == expected_groups

    current_code_counts = Counter()
    completed_code_counts = Counter()
    for row in master_rows:
        target = (
            completed_code_counts
            if row["status"] == "완료_침대단일Asset"
            else current_code_counts
        )
        target.update(row["final_codes"])

    legacy_counts = Counter(row["pattern_code"] for row in legacy_rows)
    legacy_decisions = Counter(row["provisional_decision"] for row in legacy_rows)

    lines: list[str] = []
    add = lines.append
    add("# 3D 전달 준비 현황 및 패턴 결정서")
    add("")
    add("> 전체 홈스타일 9,358개 중 3D 전달 전에 결정하거나 보강해야 할")
    add("> 규격 패턴과 상품 구성 패턴을 한 문서에서 확인한다. 결정자는 모든 상품을")
    add("> 개별 확인하지 않고 패턴 설명·대표 사례·상세 목록을 보고 일괄 처리 규칙을 승인한다.")
    add("")
    add("## 1. 3D 전달 준비 선결 현황")
    add("")
    add("| 구분 | 상품 수 | 현재 의미 |")
    add("|---|---:|---|")
    add("| 전체 대상 | 9,358 | 홈스타일 비음영 대상 상품 |")
    add("| 규격 확정 | 5,713 | API·HTML 또는 OCR·규칙으로 대표 규격 잠금 |")
    add("| 규격 패턴 결정 필요 | 3,411 | 후보 추출 완료, 대표·옵션·타제품 구분 필요 |")
    add("| 최종 무후보 보강 필요 | 234 | 추가 원천 또는 대상 이미지 OCR 필요 |")
    add("| 상품 구성 결정 필요 | 149 | 단일상품·복수상품 및 3D 생성 방식 결정 필요 |")
    add("| 규격·상품 구성 동시 결정 | 76 | 3,411개와 149개에 중복 포함 |")
    add("| 고유 패턴 결정 대상 | 3,484 | 3,411 + 149 - 76 |")
    add("")
    add("규격 패턴 3,411개가 가장 큰 선결 대상이다. 상품 구성 결정 필요")
    add("149개 중 76개는 규격 패턴에도 포함되므로 두 숫자를 단순 합산하지 않는다.")
    add("")
    add("패턴은 서로 다른 목적의 세 계층으로 관리한다. 규격 패턴 6개는")
    add("후보 구조 분류이고, 상품 구성 패턴 22개와 최종 3D 처리 코드 40개는")
    add("뒤쪽에 별도로 유지된다.")
    add("")
    add("| 패턴 계층 | 패턴 수 | 용도 |")
    add("|---|---:|---|")
    add("| 규격 후보 구조 D00~D05 | 6 | 3,411개 규격 후보의 선택·보강 방식 결정 |")
    add("| 상품 구성 원천 패턴 | 22 | 과거 270개를 상품명의 구성 형태로 분류 |")
    add("| 최종 3D 처리 코드 N/S/Q/M/A/O | 40 | 단일·수량·분리·부속품·옵션 처리 방식 |")
    add("")
    add("### 1.1 상품 구성 패턴 이력 — 과거 270개에서 현재 149개")
    add("")
    add("| 단계 | 상품 수 | 의미 |")
    add("|---|---:|---|")
    add("| 최초 제목 후보 | 1,152 | `세트/패키지/+` 표기 상품 |")
    add("| 탐지 확정 | 882 | 세트·패키지 752 + 소파·스툴 130 |")
    add("| 과거 미확정 | 270 | 2026-07-24 패턴 연구 대상 |")
    add("| 침대 정책 확정 | 121 | 프레임+매트리스 단일 3D Asset |")
    add("| 현재 조합 패턴 체크 | 149 | 270 - 121 |")
    add("| 현재 규격+조합 동시 체크 | 76 | 149 중 규격도 RAW인 상품 |")
    add("| 현재 조합만 체크 | 73 | 149 중 규격은 확정된 상품 |")
    add("")
    add("```text")
    add("최초 1,152 = 탐지 확정 882 + 과거 미확정 270")
    add("과거 미확정 270 = 침대 정책 완료 121 + 현재 체크필요 149")
    add("현재 체크필요 149 = 규격 확정 73 + 규격 RAW 76")
    add("```")
    add("")
    add("`탐지 확정`은 제목상 조합상품임을 확정했다는 뜻이다. 모든 구성품 이름과")
    add("규격이 3D 행으로 확정됐다는 뜻은 아니다. 이 둘을 같은 완료 상태로 사용하면 안 된다.")
    add("")
    add("## 2. 선결 과제 — 패턴 결정 구조")
    add("")
    add("선결 패턴은 규격 후보 구조, 상품 구성 원천 패턴, 최종 3D 처리 코드의")
    add("세 계층으로 확인한다. 각 계층은 같은 상품을 다른 관점에서 분류하므로")
    add("패턴 수를 서로 합산하지 않는다.")
    add("")
    add("### 2.1 규격 후보 구조 D00~D05 — 3,411개")
    add("")
    add("`비교정보 제공`은 규격 후보 추출이 완료됐다는 뜻이며, 3D 전달용")
    add("대표 W/D/H가 확정됐다는 뜻이 아니다. 결정자는 상품 3,411개를")
    add("하나씩 보는 대신 D00~D05 패턴의 대표 사례와 규칙을 승인한다.")
    add("")
    add("D00은 상품 구성과 규격을 함께 결정해야 하므로 가장 먼저 확인한다.")
    add("D01~D03은 후보 선택·분류 패턴이고, D04~D05는 비적용 축 또는")
    add("대상 이미지 재OCR를 포함하는 보강 패턴이다.")
    add("")
    add("| 규격 패턴 코드 | 결정자용 패턴명 | 예시 | 패턴 처리 대상 | 현재 후보 구조 | 확인 방향 |")
    add("|---|---|---|---:|---|---|")
    for code in DIMENSION_PATTERN_DEFINITIONS:
        definition = DIMENSION_PATTERN_DEFINITIONS[code]
        add(
            f"| `{code}` | {md_text(definition['name'])} | "
            f"{md_text(definition['example'])} | "
            f"{dimension_pattern_counts[code]:,} | "
            f"{md_text(definition['candidate_structure'])} | "
            f"{md_text(definition['decision'])} |"
        )
    add("| **합계** |  |  | **3,411** |  |  |")
    add("")
    add("아래 상품 목록은 같은 규격 패턴별로 접혀 있다. `상세 보기`를 누르면")
    add("해당 패턴의 전체 상품과 현재 규격 후보를 확인할 수 있다.")
    add("")
    add("| 규격 패턴 코드 | 패턴명 | 패턴 예시 | 현재 상태 | 상품 ID | 상품명 | 중·소카테고리 | 후보 구조 | 규격 후보(mm) | 확인 방향 |")
    add("|---|---|---|---|---|---|---|---|---|---|")
    for row in sorted(
        dimension_review_rows,
        key=lambda item: (item["규격 패턴 코드"], item["상품 ID"]),
    ):
        code = row["규격 패턴 코드"]
        definition = DIMENSION_PATTERN_DEFINITIONS[code]
        structure, candidates = dimension_candidate_summary(
            comparison_candidates[row["상품 ID"]]
        )
        candidate = comparison_candidates[row["상품 ID"]][0]
        categories = (
            f"{candidate.get('mid_category') or '-'} / "
            f"{candidate.get('small_category') or '-'}"
        )
        add(
            f"| `{code}` | {md_text(definition['name'])} | "
            f"{md_text(definition['example'])} | "
            f"후보 추출 완료 · 규격 결정 필요 | "
            f"{product_link(row['상품 ID'])} | {md_text(row['상품명'])} | "
            f"{md_text(categories)} | {md_text(structure)} | "
            f"{md_text(candidates)} | {md_text(definition['decision'])} |"
        )
    add("")
    add("패턴 승인 후 제품 규격으로 채택된 값은 `API·HTML 확정` 또는")
    add("`OCR·규칙 확정`으로 이동한다. 배송·타제품·오탐뿐이면")
    add("`최종 무후보`로 이동하며 추가 원천 또는 OCR 보강 대상으로 남긴다.")
    add("")
    add("### 2.2 상품 구성 원천 패턴 — 22개")
    add("")
    add("상품명의 `+`, 세트, 패키지 표기를 기준으로 과거 270개를 22개")
    add("원천 패턴으로 분류했다. 침대 정책으로 121개는 이미 단일상품 처리됐고,")
    add("현재 상품 구성 결정이 필요한 상품은 149개다.")
    add("")
    add("#### 2.2.1 원천 패턴별 상태 요약")
    add("")
    add("| 기존 패턴 코드 | 결정자용 패턴명 | 구성 예시 | 270 건수 | 상품 구성 패턴 코드 | 현재 상태 |")
    add("|---|---|---|---:|---|---|")
    for legacy_code in sorted(legacy_counts):
        final = LEGACY_TO_FINAL[legacy_code]
        state = (
            "121개 완료"
            if legacy_code == "Y04_BED_SLEEP_SYSTEM"
            else "현재 149개에서 체크"
        )
        add(
            f"| `{legacy_code}` | {md_text(LEGACY_DISPLAY_NAMES[legacy_code])} | "
            f"{md_text(LEGACY_PATTERN_EXAMPLES[legacy_code])} | "
            f"{legacy_counts[legacy_code]:,} | `{final}` | {state} |"
        )
    add("")
    add("과거 `N02/N03/N04`의 번호와 최종 N 코드 번호가 달라졌다.")
    add("예를 들어 과거 `N02_OPTION_TYPE_JOIN`은 최종 `N04`다.")
    add("DB에는 기존 원천 코드와 최종 3D 처리 코드를 별도 컬럼으로 보존한다.")
    add("")
    add("#### 2.2.2 상품 구성 원천 패턴 전체 상품")
    add("")
    add("아래 표는 과거 270개를 누락 없이 한 번씩 나열한다. 같은 원천 패턴별로")
    add("접혀 있으며 `상세 보기`에서 상품과 현재 처리 상태를 확인할 수 있다.")
    add("")
    add("| 현재 상태 | 기존 패턴 코드 | 패턴명 | 패턴 예시 | 상품 구성 패턴 코드 | 상품 ID | 상품명 | 산출 상태 | D코드 | 3D 처리 |")
    add("|---|---|---|---|---|---|---|---|---|---|")
    for row in sorted(
        master_rows,
        key=lambda item: (
            0 if item["status"] == "완료_침대단일Asset" else 1,
            item["legacy_code"],
            item["product_id"],
        ),
    ):
        final = " \\| ".join(row["final_codes"])
        add(
            f"| {row['status']} | `{row['legacy_code']}` | "
            f"{md_text(row['display_pattern_name'])} | "
            f"{md_text(row['display_pattern_example'])} | `{final}` | "
            f"{product_link(row['product_id'])} | {md_text(row['product_name'])} | "
            f"{md_text(row['output_status'])} | `{md_text(row['dimension_pattern'])}` | "
            f"{md_text(row['action'])} |"
        )
    add("")
    add("### 2.3 최종 3D 처리 코드 N/S/Q/M/A/O — 40개")
    add("")
    add("원천 패턴을 실제 3D 생성 방식으로 세분한 최종 코드다. 한 상품에는")
    add("`M03+M04+O06`처럼 여러 코드가 동시에 적용될 수 있다.")
    add("")
    add("| 코드 | 결정자용 패턴명 | 270 내 완료 | 현재 체크 | 270 통합 발생 | 3D 처리 | 대표 예제 |")
    add("|---|---|---:|---:|---:|---|---|")
    for code, name, action, signal in PATTERN_DEFINITIONS:
        completed = completed_code_counts[code]
        current = current_code_counts[code]
        total = completed + current
        examples = [row["product_id"] for row in code_products.get(code, [])[:3]]
        if not examples and code in REFERENCE_EXAMPLES:
            examples = [REFERENCE_EXAMPLES[code]]
        example_text = ", ".join(product_link(pid) for pid in examples) or "-"
        add(
            f"| `{code}` | {name}<br>{signal} | {completed:,} | {current:,} | "
            f"{total:,} | {action} | {example_text} |"
        )
    add("")
    add("#### 2.3.1 현재 상품 구성 결정 필요 149개 그룹 통계")
    add("")
    add("| 그룹 | 고유 상품 수 | 판단 결과 |")
    add("|---|---:|---|")
    group_meaning = {
        "N": "조합 탐지에서 제외",
        "S": "조립 후 단일 Asset",
        "Q": "동일 Asset 수량",
        "M": "서로 다른 Asset 행 분리",
        "A": "부속품 별도 제작 정책",
        "O": "옵션에 따른 형상·규격 분기",
    }
    for group in ("N", "S", "Q", "M", "A", "O"):
        add(
            f"| `{group}` | {current_group_counts[group]:,} | "
            f"{group_meaning[group]} |"
        )
    add("")
    add("O군 127개는 N/S/Q/M/A와 독립적인 옵션 축이다. O01~O06 행별 발생")
    add("합계는 150개지만 한 상품에 여러 옵션 패턴이 있어 고유 상품은 127개다.")
    add("")
    add("## 3. 분류 축")
    add("")
    add("| 축 | 질문 | 값 |")
    add("|---|---|---|")
    add("| 조합 탐지 | 실제 조합상품인가? | `CONFIRMED / CANDIDATE / NOT_COMBINATION` |")
    add("| 3D 구조 | 몇 개 Asset 행인가? | `N / S / Q / M / A` |")
    add("| 옵션 구조 | 옵션이 형상·규격을 바꾸는가? | `O01~O06` |")
    add("| 규격 후보 | W/D/H 후보가 어떤 구조인가? | `D00~D05` |")
    add("| 산출 상태 | 현재 값을 어떤 형태로 냈는가? | 확정값 / RAW분리 / 무후보분류 |")
    add("")
    add("한 상품은 `M03+M04+O06`처럼 여러 최종 코드를 동시에 가질 수 있다.")
    add("따라서 패턴군 발생 건수는 서로 더하지 않으며 상품 수는 항상")
    add("`COUNT(DISTINCT product_id)`로 계산한다.")
    add("")
    add("## 4. 침대 121개 확정 정책")
    add("")
    add("- 최종 코드: `S01 + A05`")
    add("- 상태: `완료_침대단일Asset`")
    add("- 프레임과 매트리스는 `component_seq`로 분리하지 않는다.")
    add("- 실제 구매 옵션에서 침대 외형 사이즈가 달라질 때만 `option_seq`로 분리한다.")
    add("- 발통, 매트리스 종류·등급, 색상 옵션은 외형 사이즈 분리에서 제외한다.")
    add("- 121개 상품의 구매 옵션 347행을 검사했으며 현재 명시적 외형 사이즈")
    add("  분리 대상은 0개다.")
    add("")
    add("이 정책 적용으로 과거 구성품 행 2,111행은 현재 1,990행이 됐다.")
    add("121개가 사라진 것이 아니라 단일 Asset 정책으로 완료 처리된 것이다.")
    add("")
    add("## 5. 탐지 확정 882개의 해석")
    add("")
    add("| 탐지 규칙 | 구성품 상태 | 상품 수 | 의미 |")
    add("|---|---|---:|---|")
    status_names = {
        "ALL_COMPONENT_DIMENSIONS_CONFIRMED": "모든 구성품 규격 확정",
        "COMPONENT_DIMENSIONS_CANDIDATE": "구성품 규격 후보",
        "PARTIAL_COMPONENT_DIMENSIONS": "구성품 규격 부분확보",
        "COMPONENT_DIMENSIONS_MISSING": "구성품 규격 미확보",
        "COMPONENT_PARSE_REQUIRED": "구성품명 분리 필요",
    }
    rule_names = {
        "TITLE_SET_OR_PACKAGE": "세트·패키지 제목",
        "TITLE_SOFA_PLUS_STOOL": "소파+스툴/오토만",
    }
    for row in confirmed_summary:
        add(
            f"| `{row['detection_rule']}`<br>{rule_names[row['detection_rule']]} | "
            f"`{row['component_output_status']}`<br>"
            f"{status_names[row['component_output_status']]} | {row['n']:,} | "
            f"조합 탐지는 확정, 구성품 규격 상태는 별도 |"
        )
    add("")
    add("- 소파+스툴 130개는 최종 구조상 `M05`다. 그중 구성품 규격 전체")
    add("  확정은 7개이고 123개는 규격 보강 대상이다.")
    add("- 세트·패키지 752개는 조합 탐지는 확정했지만 상품별 3D 세부 구조가")
    add("  모두 N/S/Q/M/A로 승인된 것은 아니다. `완료_조합탐지`와")
    add("  `완료_3D구조확정`을 별도 상태로 관리해야 한다.")
    add("")
    add("대표 탐지 확정 예제:")
    add("")
    for rule in ("TITLE_SET_OR_PACKAGE", "TITLE_SOFA_PLUS_STOOL"):
        add(f"- `{rule}`")
        for row in confirmed_examples[rule]:
            add(
                f"  - {product_link(row['product_id'])} "
                f"{md_text(row['product_name'])}"
            )
    add("")
    add("## 6. 3D 행 생성 규칙")
    add("")
    add("| 최종 그룹 | 저장 방식 |")
    add("|---|---|")
    add("| N, S | `product_id + seq=1` 단일 Asset |")
    add("| Q | Asset 1종과 `quantity`; 동일 형상 행을 중복 생성하지 않음 |")
    add("| M | 같은 `product_id`, 서로 다른 `component_seq` |")
    add("| A | 별도 제작 승인 시에만 `component_seq`; 미승인 시 본체 포함 |")
    add("| O01/O03/O04 | 외형이 바뀌므로 `option_seq`별 규격 행 |")
    add("| O02 | 수량 또는 구성 옵션으로 분리 |")
    add("| O05/O06 | 외형이 같으면 규격 행을 분리하지 않고 옵션 메타만 저장 |")
    add("")
    add("공식 세트 구성 실제 ID는 현재 원천에서 미확보다.")
    add("")
    add("```text")
    add("set_id            = NULL")
    add("source_product_id = product_id")
    add("component_key     = product_id + component_seq")
    add("sales_option_key  = product_id + option_id")
    add("asset_option_key  = product_id + option_seq")
    add("```")
    add("")
    add("## 7. 패턴 승인 후 상태 전이")
    add("")
    add("1. 상품 페이지와 API 옵션을 확인한다.")
    add("2. 과거 코드와 최종 N/S/Q/M/A/O 코드를 함께 저장한다.")
    add("3. `M`이면 구성품, 형상 변경 `O`이면 옵션 행을 생성한다.")
    add("4. 규격 후보 D00~D05에서 각 행에 연결할 W/D/H를 선택한다.")
    add("5. 패턴·예외와 `rule_version`을 저장한다.")
    add("6. ledger → component staging → mandatory → Excel 순으로 재산출한다.")
    add("7. 패턴별 대표 상품을 회귀 테스트에 추가한다.")
    add("")
    add("## 8. DB·Excel 필드")
    add("")
    add("필수 보존 필드:")
    add("")
    add("```text")
    add("product_id")
    add("legacy_pattern_code")
    add("final_pattern_codes")
    add("pattern_groups")
    add("pattern_status")
    add("asset_structure_status")
    add("component_seq / option_seq / quantity")
    add("rule_version")
    add("source_type / source_ref / evidence_text")
    add("```")
    add("")
    add("현재 Excel `08_패턴_상태`는 산출 상태, 패턴 상태, D코드, 조합 패턴군,")
    add("세부 코드, 과거 원천 코드, 상품 URL을 필터링할 수 있다.")
    add("")
    add("## 9. 문서와 실행 파일")
    add("")
    add("- 이 문서: 최종 패턴 사전·교차표·270개 전체 매핑")
    add("- `조합상품_미확정270_패턴분류_2026-07-24.md`: 과거 snapshot 기록")
    add("- `조합상품_다중행_검토_2026-07-24.md`: `component_seq` 구현 기록")
    add("- `세트ID_옵션ID_검증_2026-07-24.md`: 공식 세트 ID 미확보 근거")
    add("- `build_combination_candidate_pattern_staging.py`: 과거 패턴 staging")
    add("- `build_bed_asset_size_option_staging.py`: 침대 단일 Asset 정책")
    add("- `build_product_component_staging.py`: 구성품 행 생성")
    add("- `build_homestyle_bulk_workbook.py`: 최종 코드·상태 Excel 출력")
    add("")
    add("재생성:")
    add("")
    add("```powershell")
    add("python build_combination_pattern_master.py")
    add("```")

    meta = {
        "initial_candidates": 1152,
        "confirmed_detection": 882,
        "legacy_270": len(legacy_ids),
        "bed_completed": len(bed_ids),
        "current_review": len(candidate_ids),
        "legacy_decisions": dict(legacy_decisions),
        "current_group_counts": dict(current_group_counts),
        "master_rows": len(master_rows),
        "legacy_snapshot": legacy_snapshot,
    }
    return "\n".join(lines) + "\n", meta


def main() -> None:
    document, meta = build_document()
    OUTPUT_PATH.write_text(document, encoding="utf-8")
    print(OUTPUT_PATH)
    print(meta)


if __name__ == "__main__":
    main()
