from __future__ import annotations

import json
import re
import sqlite3
from dataclasses import dataclass, replace
from datetime import datetime
from pathlib import Path
from typing import Any

from bulk_homestyle_collect import DB_PATH, unpack
from build_homestyle_bulk_workbook import (
    clean_text,
    notification_items,
    notification_values,
)


PARSER_VERSION = "product-component-v1.1"
SEOUL_OFFSET = "+09:00"

COMBINATION_CONFIRMED = "CONFIRMED"
COMBINATION_CANDIDATE = "CANDIDATE"
NOT_COMBINATION = "NOT_COMBINATION"

COMPONENT_API_CONFIRMED = "API_CONFIRMED"
COMPONENT_API_UNIT_INFERRED = "API_UNIT_INFERRED"
COMPONENT_DIMENSION_MISSING = "DIMENSION_MISSING"
COMPONENT_NAME_REQUIRED = "COMPONENT_NAME_REQUIRED"
COMBINATION_REVIEW_REQUIRED = "COMBINATION_REVIEW_REQUIRED"

SIZE_KEYWORDS = ("크기", "치수", "규격", "사이즈")
PLUS_RE = re.compile(r"[+＋]")
SET_RE = re.compile(r"세트|패키지|package", re.I)
SOFA_STOOL_RE = re.compile(r"스툴|오토만|ottoman", re.I)
BED_SLEEP_SYSTEM_RE = re.compile(
    r"매트리스|매트\b|매트Q|매트SS|메모리폼탑|"
    r"노뜨컴포트|파인탑|볼륨탑|프레임",
    re.I,
)
BED_MULTI_BUNDLE_RE = re.compile(
    r"2개묶음|슈퍼싱글\s*[+＋]\s*슈퍼싱글|"
    r"퀸\s*[+＋]\s*(?:퀸|슈퍼싱글)|Q\s*[+＋]\s*SS",
    re.I,
)

NUMBER = r"\d{1,5}(?:,\d{3})*(?:\.\d+)?"
RANGE_NUMBER = rf"({NUMBER})(?:\s*[-~]\s*({NUMBER}))?"
PLAIN_COMPONENT_LABEL = (
    r"(?:(?:\d+(?:\.\d+)?\s*인)\s*)?"
    r"(?:소파|스툴|오토만|ottoman|식탁|테이블|table|의자|체어|chair|"
    r"벤치|bench|침대|bed|매트리스|mattress|협탁|nightstand|책상|desk|"
    r"책장|bookcase|수납장|거실장|옷장|장식장|cabinet|화장대|dresser|"
    r"거울|mirror|행거|hanger|선반|shelf|램프|조명|lamp|쿠션|cushion|"
    r"헤드레스트|headrest|가드|guard|패널|panel)"
)
LABELED_WDH_RE = re.compile(
    rf"""
    (?:
        \[(?P<bracket>[^\]\r\n]{{1,50}})\]
        |
        (?P<plain>{PLAIN_COMPONENT_LABEL})(?:\s*사이즈)?
    )
    \s*(?:[:：\-–]\s*)?
    \(?\s*W\s*\)?\s*[:=]?\s*(?P<w1>{NUMBER})(?:\s*[-~]\s*(?P<w2>{NUMBER}))?
    \s*(?P<wu>mm|cm)?\s*(?:[xX×*]\s*)?
    \(?\s*D\s*\)?\s*[:=]?\s*(?P<d1>{NUMBER})(?:\s*[-~]\s*(?P<d2>{NUMBER}))?
    \s*(?P<du>mm|cm)?\s*(?:[xX×*]\s*)?
    \(?\s*H\s*\)?\s*[:=]?\s*(?P<h1>{NUMBER})(?:\s*[-~]\s*(?P<h2>{NUMBER}))?
    \s*(?P<hu>mm|cm)?
    """,
    re.I | re.X,
)


@dataclass(frozen=True)
class Component:
    component_name: str
    component_type: str
    component_quantity: int = 1
    is_primary: int = 0
    w_mm: float | None = None
    d_mm: float | None = None
    h_mm: float | None = None
    resolution_status: str = COMPONENT_DIMENSION_MISSING
    source_type: str = "PRODUCT_TITLE"
    source_ref: str = "fact_dimension_resolution_ledger.product_name"
    evidence_text: str = ""

    @property
    def complete(self) -> bool:
        return all(value is not None for value in (self.w_mm, self.d_mm, self.h_mm))

    @property
    def confirmed(self) -> bool:
        return self.complete and self.resolution_status == COMPONENT_API_CONFIRMED


COMPONENT_PATTERNS: list[tuple[str, str, re.Pattern[str]]] = [
    ("MATTRESS", "매트리스", re.compile(r"매트리스|mattress", re.I)),
    ("NIGHTSTAND", "협탁", re.compile(r"협탁|nightstand", re.I)),
    ("STOOL", "스툴", re.compile(r"스툴|오토만|ottoman", re.I)),
    ("SOFA", "소파", re.compile(r"소파|sofa", re.I)),
    ("DINING_TABLE", "식탁", re.compile(r"식탁", re.I)),
    ("TABLE", "테이블", re.compile(r"테이블|table", re.I)),
    ("CHAIR", "의자", re.compile(r"의자|체어|chair", re.I)),
    ("BENCH", "벤치", re.compile(r"벤치|bench", re.I)),
    ("BED_FRAME", "침대", re.compile(r"침대|\bbed\b", re.I)),
    ("DESK", "책상", re.compile(r"책상|\bdesk\b", re.I)),
    ("BOOKCASE", "책장", re.compile(r"책장|bookcase", re.I)),
    (
        "CABINET",
        "수납장",
        re.compile(r"수납장|거실장|옷장|장식장|콘솔장|cabinet", re.I),
    ),
    ("DRESSER", "화장대", re.compile(r"화장대|dresser", re.I)),
    ("MIRROR", "거울", re.compile(r"거울|mirror", re.I)),
    ("HANGER", "행거", re.compile(r"행거|hanger", re.I)),
    ("SHELF", "선반", re.compile(r"선반|shelf", re.I)),
    ("LAMP", "조명", re.compile(r"램프|조명|lamp", re.I)),
    ("HEADREST", "헤드레스트", re.compile(r"헤드레스트|headrest", re.I)),
    ("CUSHION", "쿠션", re.compile(r"쿠션|cushion", re.I)),
    ("GUARD", "가드", re.compile(r"가드|guard", re.I)),
    ("PANEL", "패널", re.compile(r"패널|panel", re.I)),
]


def now_iso() -> str:
    return datetime.now().astimezone().replace(microsecond=0).isoformat()


def is_bed_frame_mattress_single_asset(
    product_name: str,
    mid_category: str,
    small_category: str = "",
) -> bool:
    normalized = (product_name or "").replace("＋", "+")
    return (
        mid_category == "침대"
        and small_category == "침대+매트리스"
        and "+" in normalized
        and bool(BED_SLEEP_SYSTEM_RE.search(normalized))
        and not bool(BED_MULTI_BUNDLE_RE.search(normalized))
    )


def detect_combination(
    product_name: str,
    mid_category: str,
    small_category: str = "",
) -> tuple[str, int, str]:
    name = product_name or ""
    is_set = bool(SET_RE.search(name))
    is_sofa_stool = (
        mid_category == "소파"
        and bool(PLUS_RE.search(name))
        and bool(SOFA_STOOL_RE.search(name))
    )
    if is_set:
        return COMBINATION_CONFIRMED, 1, "TITLE_SET_OR_PACKAGE"
    if is_bed_frame_mattress_single_asset(
        product_name,
        mid_category,
        small_category,
    ):
        return (
            NOT_COMBINATION,
            0,
            "POLICY_BED_FRAME_MATTRESS_SINGLE_3D_ASSET",
        )
    if is_sofa_stool:
        return COMBINATION_CONFIRMED, 1, "TITLE_SOFA_PLUS_STOOL"
    if PLUS_RE.search(name):
        return COMBINATION_CANDIDATE, 0, "TITLE_PLUS_REVIEW"
    return NOT_COMBINATION, 0, "NONE"


def component_type_for_label(label: str, mid_category: str) -> tuple[str, str]:
    normalized = clean_text(label, 100)
    for type_code, default_name, pattern in COMPONENT_PATTERNS:
        if pattern.search(normalized):
            return type_code, default_name
    if re.fullmatch(r"\d+(?:\.\d+)?\s*인", normalized):
        if mid_category == "소파":
            return "SOFA", "소파"
        if "식탁" in mid_category or "테이블" in mid_category:
            return "DINING_TABLE", "식탁"
    return "UNRESOLVED", normalized or "구성품명 미확보"


def component_quantity(text: str) -> int:
    match = re.search(r"(\d+)\s*(?:EA|개|P)(?![A-Za-z])", text, flags=re.I)
    return max(1, int(match.group(1))) if match else 1


def sofa_component_name(product_name: str) -> str:
    seat_matches = list(re.finditer(r"(\d+(?:\.\d+)?)\s*인", product_name))
    if seat_matches:
        return f"{seat_matches[-1].group(1)}인 소파"
    return "소파"


def title_components(
    product_name: str,
    mid_category: str,
    detection_rule: str,
) -> list[Component]:
    if detection_rule == "TITLE_SOFA_PLUS_STOOL":
        secondary = "오토만" if re.search(r"오토만|ottoman", product_name, re.I) else "스툴"
        secondary_type = "STOOL"
        return [
            Component(
                component_name=sofa_component_name(product_name),
                component_type="SOFA",
                is_primary=1,
                evidence_text=product_name,
            ),
            Component(
                component_name=secondary,
                component_type=secondary_type,
                component_quantity=component_quantity(
                    next(
                        (
                            part
                            for part in PLUS_RE.split(product_name)
                            if SOFA_STOOL_RE.search(part)
                        ),
                        secondary,
                    )
                ),
                evidence_text=product_name,
            ),
        ]

    matches: list[tuple[int, Component]] = []
    occupied: list[tuple[int, int]] = []
    for type_code, default_name, pattern in COMPONENT_PATTERNS:
        for match in pattern.finditer(product_name):
            if any(start <= match.start() < end for start, end in occupied):
                continue
            nearby = product_name[match.start() : min(len(product_name), match.end() + 12)]
            matches.append(
                (
                    match.start(),
                    Component(
                        component_name=default_name,
                        component_type=type_code,
                        component_quantity=component_quantity(nearby),
                        evidence_text=product_name,
                    ),
                )
            )
            occupied.append((match.start(), match.end()))
            break

    matches.sort(key=lambda item: item[0])
    result: list[Component] = []
    seen_types: set[str] = set()
    for _, component in matches:
        if component.component_type in seen_types:
            continue
        result.append(replace(component, is_primary=1 if not result else 0))
        seen_types.add(component.component_type)
    return result


def _number(value: str | None) -> float:
    return float(str(value or "0").replace(",", ""))


def _axis_value(first: str, second: str | None) -> float:
    return _number(second or first)


def _to_mm(value: float, unit: str) -> float:
    result = value * 10 if unit.casefold() == "cm" else value
    return float(int(result)) if result.is_integer() else result


def labeled_components_from_text(
    text: str,
    mid_category: str,
    product_name: str,
    source_ref: str,
) -> list[Component]:
    cleaned = clean_text(text, 10000).replace("㎜", "mm")
    result: list[Component] = []
    seen: set[tuple[str, float, float, float]] = set()
    for match in LABELED_WDH_RE.finditer(cleaned):
        label = clean_text(match.group("bracket") or match.group("plain"), 100)
        type_code, default_name = component_type_for_label(label, mid_category)
        if type_code == "UNRESOLVED":
            continue

        # A size label such as [3인] in a 4인 product is usually a lineup row,
        # not the sold component. Keep only the title-matching primary variant.
        seat_match = re.fullmatch(r"(\d+(?:\.\d+)?)\s*인", label)
        if seat_match and seat_match.group(1) + "인" not in product_name:
            continue

        explicit_units = [
            unit
            for unit in (match.group("wu"), match.group("du"), match.group("hu"))
            if unit
        ]
        raw_values = [
            _axis_value(match.group("w1"), match.group("w2")),
            _axis_value(match.group("d1"), match.group("d2")),
            _axis_value(match.group("h1"), match.group("h2")),
        ]
        if explicit_units:
            common_unit = explicit_units[-1]
            resolution_status = COMPONENT_API_CONFIRMED
        else:
            common_unit = "cm" if max(raw_values) <= 300 else "mm"
            resolution_status = COMPONENT_API_UNIT_INFERRED

        units = [
            match.group("wu") or common_unit,
            match.group("du") or common_unit,
            match.group("hu") or common_unit,
        ]
        values = tuple(_to_mm(value, unit) for value, unit in zip(raw_values, units))
        if any(value <= 0 or value > 10000 for value in values):
            continue
        key = (type_code, values[0], values[1], values[2])
        if key in seen:
            continue
        seen.add(key)

        component_name = label
        if type_code == "SOFA" and re.fullmatch(r"\d+(?:\.\d+)?\s*인", label):
            component_name = f"{label.replace(' ', '')} 소파"
        elif label.casefold() in {"ottoman"}:
            component_name = "오토만"
        elif not label:
            component_name = default_name
        result.append(
            Component(
                component_name=component_name,
                component_type=type_code,
                component_quantity=component_quantity(label),
                w_mm=values[0],
                d_mm=values[1],
                h_mm=values[2],
                resolution_status=resolution_status,
                source_type="GOODS_API_NOTIFICATION",
                source_ref=source_ref,
                evidence_text=match.group(0),
            )
        )
    return result


def merge_components(
    title_rows: list[Component],
    api_rows: list[Component],
) -> list[Component]:
    result = list(title_rows)
    for api_component in api_rows:
        matching_indexes = [
            index
            for index, existing in enumerate(result)
            if existing.component_type == api_component.component_type
            and not existing.complete
        ]
        if matching_indexes:
            index = matching_indexes[0]
            result[index] = replace(
                api_component,
                component_quantity=max(
                    result[index].component_quantity,
                    api_component.component_quantity,
                ),
                is_primary=result[index].is_primary,
            )
        elif not any(
            existing.component_type == api_component.component_type
            and existing.w_mm == api_component.w_mm
            and existing.d_mm == api_component.d_mm
            and existing.h_mm == api_component.h_mm
            for existing in result
        ):
            result.append(api_component)

    if result and not any(component.is_primary for component in result):
        result[0] = replace(result[0], is_primary=1)
    return result


def ensure_minimum_component_rows(
    components: list[Component],
    product_name: str,
    detection_status: str,
) -> list[Component]:
    if detection_status == COMBINATION_CANDIDATE:
        return [
            Component(
                component_name="조합 여부 검토필요",
                component_type="UNRESOLVED",
                resolution_status=COMBINATION_REVIEW_REQUIRED,
                source_type="PRODUCT_TITLE",
                evidence_text=product_name,
            )
        ]
    if detection_status != COMBINATION_CONFIRMED:
        return []
    result = list(components)
    while len(result) < 2:
        result.append(
            Component(
                component_name="추가 구성품명 미확보",
                component_type="UNRESOLVED",
                resolution_status=COMPONENT_NAME_REQUIRED,
                source_type="PRODUCT_TITLE",
                evidence_text=product_name,
            )
        )
    if result and not any(component.is_primary for component in result):
        result[0] = replace(result[0], is_primary=1)
    return result


def output_status(
    detection_status: str,
    components: list[Component],
) -> tuple[str, int, str]:
    if detection_status == NOT_COMBINATION:
        return "NOT_APPLICABLE", 0, "조합상품 아님"
    if detection_status == COMBINATION_CANDIDATE:
        return "COMBINATION_REVIEW_REQUIRED", 1, "상품명 + 의미 확인"
    numeric_complete = sum(component.complete for component in components)
    confirmed = sum(component.confirmed for component in components)
    unresolved = any(component.component_type == "UNRESOLVED" for component in components)
    if unresolved:
        return "COMPONENT_PARSE_REQUIRED", 1, "구성품명 분리 후 규격 탐색"
    if confirmed == len(components) and components:
        return "ALL_COMPONENT_DIMENSIONS_CONFIRMED", 0, "완료"
    if numeric_complete == len(components) and components:
        return "COMPONENT_DIMENSIONS_CANDIDATE", 1, "단위·구성 문맥 확인"
    if numeric_complete:
        return "PARTIAL_COMPONENT_DIMENSIONS", 1, "미확보 구성품 규격 보강"
    return "COMPONENT_DIMENSIONS_MISSING", 1, "API→HTML/Q&A→OCR 구성품 규격 보강"


def create_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS stg_product_combination (
            snapshot_id TEXT NOT NULL,
            built_at TEXT NOT NULL,
            is_current INTEGER NOT NULL CHECK (is_current IN (0,1)),
            parser_version TEXT NOT NULL,
            product_id TEXT NOT NULL,
            product_name TEXT NOT NULL,
            mid_category TEXT,
            small_category TEXT,
            detection_status TEXT NOT NULL,
            is_combination INTEGER NOT NULL CHECK (is_combination IN (0,1)),
            detection_rule TEXT NOT NULL,
            component_count INTEGER NOT NULL,
            complete_component_count INTEGER NOT NULL,
            component_output_status TEXT NOT NULL,
            needs_human_review INTEGER NOT NULL CHECK (needs_human_review IN (0,1)),
            next_action TEXT NOT NULL,
            evidence_text TEXT,
            PRIMARY KEY (snapshot_id, product_id)
        );

        CREATE INDEX IF NOT EXISTS idx_product_combination_current_status
        ON stg_product_combination(is_current, detection_status, product_id);

        CREATE TABLE IF NOT EXISTS stg_product_component_dimension (
            snapshot_id TEXT NOT NULL,
            built_at TEXT NOT NULL,
            is_current INTEGER NOT NULL CHECK (is_current IN (0,1)),
            parser_version TEXT NOT NULL,
            product_id TEXT NOT NULL,
            product_name TEXT NOT NULL,
            component_seq INTEGER NOT NULL,
            component_name TEXT NOT NULL,
            component_type TEXT NOT NULL,
            component_quantity INTEGER NOT NULL DEFAULT 1,
            is_primary INTEGER NOT NULL CHECK (is_primary IN (0,1)),
            w_mm REAL,
            d_mm REAL,
            h_mm REAL,
            resolution_status TEXT NOT NULL,
            source_type TEXT,
            source_ref TEXT,
            evidence_text TEXT,
            needs_human_review INTEGER NOT NULL CHECK (needs_human_review IN (0,1)),
            PRIMARY KEY (snapshot_id, product_id, component_seq)
        );

        CREATE INDEX IF NOT EXISTS idx_component_dimension_current_product
        ON stg_product_component_dimension(is_current, product_id, component_seq);

        DROP VIEW IF EXISTS vw_product_combination_current;
        CREATE VIEW vw_product_combination_current AS
        SELECT *
        FROM stg_product_combination
        WHERE is_current=1;

        DROP VIEW IF EXISTS vw_product_component_dimensions_current;
        CREATE VIEW vw_product_component_dimensions_current AS
        SELECT *
        FROM stg_product_component_dimension
        WHERE is_current=1;

        DROP VIEW IF EXISTS vw_product_combination_summary;
        CREATE VIEW vw_product_combination_summary AS
        SELECT
            detection_status,
            detection_rule,
            COUNT(*) AS product_count,
            SUM(component_count) AS component_rows,
            SUM(complete_component_count) AS complete_component_rows,
            SUM(needs_human_review) AS review_products
        FROM vw_product_combination_current
        GROUP BY detection_status, detection_rule;
        """
    )


def build_staging(db_path: Path = DB_PATH) -> dict[str, Any]:
    built_at = now_iso()
    snapshot_id = "product-component-v1-" + re.sub(r"[^0-9]", "", built_at)
    connection = sqlite3.connect(db_path)
    connection.row_factory = sqlite3.Row
    create_schema(connection)
    products = connection.execute(
        """
        SELECT
            l.product_id,
            l.product_name,
            l.mid_category,
            l.small_category,
            s.goods_blob
        FROM fact_dimension_resolution_ledger l
        JOIN sources s USING(product_id)
        ORDER BY l.product_id
        """
    ).fetchall()

    combination_rows: list[tuple[Any, ...]] = []
    component_rows: list[tuple[Any, ...]] = []
    for row in products:
        product_id = row["product_id"]
        product_name = row["product_name"] or product_id
        mid_category = row["mid_category"] or ""
        detection_status, is_combination, detection_rule = detect_combination(
            product_name,
            mid_category,
            row["small_category"] or "",
        )

        components: list[Component] = []
        if detection_status != NOT_COMBINATION:
            components = title_components(
                product_name,
                mid_category,
                detection_rule,
            )
            goods_payload = unpack(row["goods_blob"]) or {}
            data = goods_payload.get("data") or {}
            size_texts = notification_values(
                notification_items(data),
                *SIZE_KEYWORDS,
            )
            api_components: list[Component] = []
            for index, text in enumerate(size_texts, start=1):
                api_components.extend(
                    labeled_components_from_text(
                        text,
                        mid_category,
                        product_name,
                        f"goods.productNotification.size[{index}]",
                    )
                )
            components = merge_components(components, api_components)
            components = ensure_minimum_component_rows(
                components,
                product_name,
                detection_status,
            )

        component_output_status, needs_review, next_action = output_status(
            detection_status,
            components,
        )
        complete_component_count = sum(component.confirmed for component in components)
        combination_rows.append(
            (
                snapshot_id,
                built_at,
                1,
                PARSER_VERSION,
                product_id,
                product_name,
                mid_category,
                row["small_category"] or "",
                detection_status,
                is_combination,
                detection_rule,
                len(components),
                complete_component_count,
                component_output_status,
                needs_review,
                next_action,
                product_name if detection_status != NOT_COMBINATION else "",
            )
        )
        for component_seq, component in enumerate(components, start=1):
            component_needs_review = int(
                component.resolution_status != COMPONENT_API_CONFIRMED
            )
            component_rows.append(
                (
                    snapshot_id,
                    built_at,
                    1,
                    PARSER_VERSION,
                    product_id,
                    product_name,
                    component_seq,
                    component.component_name,
                    component.component_type,
                    component.component_quantity,
                    component.is_primary,
                    component.w_mm,
                    component.d_mm,
                    component.h_mm,
                    component.resolution_status,
                    component.source_type,
                    component.source_ref,
                    component.evidence_text,
                    component_needs_review,
                )
            )

    with connection:
        connection.execute(
            "UPDATE stg_product_combination SET is_current=0 WHERE is_current=1"
        )
        connection.execute(
            "UPDATE stg_product_component_dimension SET is_current=0 WHERE is_current=1"
        )
        connection.executemany(
            """
            INSERT INTO stg_product_combination VALUES (
                ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?
            )
            """,
            combination_rows,
        )
        connection.executemany(
            """
            INSERT INTO stg_product_component_dimension VALUES (
                ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?
            )
            """,
            component_rows,
        )

    status_counts = {
        status: count
        for status, count in connection.execute(
            """
            SELECT detection_status,COUNT(*)
            FROM vw_product_combination_current
            GROUP BY detection_status
            """
        )
    }
    output_counts = {
        status: count
        for status, count in connection.execute(
            """
            SELECT component_output_status,COUNT(*)
            FROM vw_product_combination_current
            GROUP BY component_output_status
            """
        )
    }
    component_status_counts = {
        status: count
        for status, count in connection.execute(
            """
            SELECT resolution_status,COUNT(*)
            FROM vw_product_component_dimensions_current
            GROUP BY resolution_status
            """
        )
    }
    example = connection.execute(
        """
        SELECT component_seq,component_name,w_mm,d_mm,h_mm,resolution_status
        FROM vw_product_component_dimensions_current
        WHERE product_id='G25070005743'
        ORDER BY component_seq
        """
    ).fetchall()
    connection.close()

    expected_example = [
        (1, "4인 소파", 2910.0, 1020.0, 910.0, COMPONENT_API_CONFIRMED),
        (2, "스툴", 740.0, 660.0, 410.0, COMPONENT_API_CONFIRMED),
    ]
    actual_example = [tuple(row) for row in example]
    if actual_example != expected_example:
        raise RuntimeError(
            f"G25070005743 component split mismatch: {actual_example}"
        )
    if sum(status_counts.values()) != len(products):
        raise RuntimeError("combination staging product count mismatch")

    return {
        "snapshot_id": snapshot_id,
        "products": len(products),
        "combination_status_counts": status_counts,
        "component_output_status_counts": output_counts,
        "component_rows": len(component_rows),
        "component_status_counts": component_status_counts,
        "example_G25070005743": actual_example,
    }


def main() -> None:
    result = build_staging()
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
