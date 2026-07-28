from __future__ import annotations

import json
import re
import sqlite3
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from bulk_homestyle_collect import DB_PATH


ROOT = Path(__file__).resolve().parent
OUTPUT = (
    ROOT
    / "홈스타일_비음영대상군_전체상품_요구필드_대량결과_패턴상태.xlsx"
)
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
COMPARISON_HEADERS = (
    "요청1_규격 비교 W (mm)",
    "요청1_규격 비교 D (mm)",
    "요청1_규격 비교 H (mm)",
    "요청1_규격 비교 대상/옵션",
    "요청1_규격 비교 원문",
)
ORDERED_DIMENSION_HEADERS = (
    "요청1_W (mm)",
    "요청1_D (mm)",
    "요청1_H (mm)",
    "요청1_규격 비교후보 수",
    *COMPARISON_HEADERS,
    "요청1_규격 상태",
)
COMBINATION_HEADERS = (
    "요청1_조합상품 여부",
    "요청1_조합상품 판정근거",
    "요청1_조합 구성품 수",
    "요청1_구성품 규격 확정 수",
    "요청1_구성품 분리 상태",
)


def column_index(reference: str) -> int:
    letters = re.match(r"[A-Z]+", reference)
    if not letters:
        raise ValueError(reference)
    number = 0
    for character in letters.group(0):
        number = number * 26 + ord(character) - 64
    return number - 1


def cell_value(cell: ET.Element) -> str | int | float:
    value_type = cell.get("t")
    if value_type == "inlineStr":
        return "".join(node.text or "" for node in cell.findall(f".//{NS}t"))
    value = cell.find(f"{NS}v")
    if value is None or value.text is None:
        return ""
    if value_type == "str":
        return value.text
    try:
        number = float(value.text)
    except ValueError:
        return value.text
    return int(number) if number.is_integer() else number


def main() -> None:
    if not OUTPUT.exists():
        raise FileNotFoundError(OUTPUT)

    headers: list[str] = []
    header_index: dict[str, int] = {}
    header_styles: dict[str, str | None] = {}
    statuses: Counter[str] = Counter()
    attachment_statuses: Counter[str] = Counter()
    comparison_rows = 0
    multi_rows = 0
    candidate_sum = 0
    alignment_errors: list[tuple[str, str, int, int]] = []
    review_style_errors: list[tuple[str, str, str | None]] = []
    attachment_status_errors: list[tuple[str, str, str]] = []
    human_review_cells = 0
    combination_labels: Counter[str] = Counter()
    component_output_labels: Counter[str] = Counter()
    combination_style_errors: list[tuple[str, str, str | None]] = []
    component_rows = 0
    component_statuses: Counter[str] = Counter()
    component_review_rows = 0
    component_style_errors: list[tuple[str, str, str | None]] = []
    component_sequence_errors: list[tuple[str, list[int]]] = []
    component_sequences: dict[str, list[int]] = {}
    component_sheet_sample: list[dict[str, object]] = []
    sheet_names: list[str] = []
    sample_ids = {"G25080006891", "G25070006116", "G25110024427"}
    samples: dict[str, dict[str, object]] = {}

    with ZipFile(OUTPUT) as archive:
        if archive.testzip() is not None:
            raise ValueError("corrupt XLSX archive")
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        sheet_names = [
            sheet.get("name", "")
            for sheet in workbook.find(f"{NS}sheets")
        ]
        if sheet_names != [
            "00_요약",
            "01_상품별_요구필드",
            "02_규격_상세",
            "03_옵션_상세",
            "04_필수값_판정",
            "05_카테고리_통계",
            "06_수집오류",
            "07_조합상품_구성품",
            "08_패턴_상태",
        ]:
            raise AssertionError(f"unexpected sheet order: {sheet_names}")
        with archive.open("xl/worksheets/sheet2.xml") as stream:
            for _, element in ET.iterparse(stream, events=("end",)):
                if element.tag != f"{NS}row":
                    continue
                cells = element.findall(f"{NS}c")
                values = {
                    column_index(cell.get("r", "")): cell_value(cell)
                    for cell in cells
                }
                styles = {
                    column_index(cell.get("r", "")): cell.get("s")
                    for cell in cells
                }
                human_review_cells += sum(style == "11" for style in styles.values())
                row_number = int(element.get("r", "0"))
                if row_number == 1:
                    headers = [
                        str(values.get(index, ""))
                        for index in range(max(values) + 1)
                    ]
                    header_index = {
                        header: index for index, header in enumerate(headers)
                    }
                    header_styles = {
                        str(values.get(column_index(cell.get("r", "")), "")):
                        cell.get("s")
                        for cell in cells
                    }
                else:
                    product_id = str(
                        values.get(header_index["상품 ID"], "")
                    )
                    attachment_status = str(
                        values.get(
                            header_index["별첨0715_필수값 판정"], ""
                        )
                    )
                    attachment_statuses[attachment_status] += 1
                    reinforcement_fields = str(
                        values.get(
                            header_index["별첨0715_보강대상 필드"], ""
                        )
                    )
                    status = str(
                        values.get(header_index["요청1_규격 상태"], "")
                    )
                    statuses[status] += 1
                    combination_label = str(
                        values.get(
                            header_index["요청1_조합상품 여부"],
                            "",
                        )
                    )
                    component_output_label = str(
                        values.get(
                            header_index["요청1_구성품 분리 상태"],
                            "",
                        )
                    )
                    combination_labels[combination_label] += 1
                    component_output_labels[component_output_label] += 1
                    combination_needs_review = (
                        combination_label == "검토필요"
                        or (
                            combination_label == "Y"
                            and component_output_label != "구성품 규격 확정"
                        )
                    )
                    if combination_needs_review:
                        for header in COMBINATION_HEADERS:
                            cell_value_for_header = str(
                                values.get(header_index[header], "")
                            )
                            expected_styles = (
                                {"9", "11"}
                                if cell_value_for_header == "미확보"
                                else {"11"}
                            )
                            if (
                                styles.get(header_index[header])
                                not in expected_styles
                            ):
                                combination_style_errors.append(
                                    (
                                        product_id,
                                        header,
                                        styles.get(header_index[header]),
                                    )
                                )
                    expected_attachment_status = {
                        "확정(DB)": "API/HTML 확정",
                        "확정(규칙)": "OCR·규칙확정",
                        "확정(수동)": "OCR·규칙확정",
                        "비교정보 제공": "비교정보제공",
                        "사람 확인 필요": "비교정보제공",
                        "OCR 보강대상": "최종무후보",
                        "규격 후보 없음": "최종무후보",
                        "규격 원천 미분류": "최종무후보",
                    }.get(status)
                    if attachment_status != expected_attachment_status:
                        attachment_status_errors.append(
                            (
                                product_id,
                                attachment_status,
                                str(expected_attachment_status),
                            )
                        )
                    if (
                        reinforcement_fields not in {"", "해당없음"}
                        and styles.get(
                            header_index["별첨0715_보강대상 필드"]
                        ) != "11"
                    ):
                        review_style_errors.append(
                            (
                                product_id,
                                "별첨0715_보강대상 필드",
                                styles.get(
                                    header_index["별첨0715_보강대상 필드"]
                                ),
                            )
                        )
                    if (
                        attachment_status in {"비교정보제공", "최종무후보"}
                        and styles.get(
                            header_index["별첨0715_필수값 판정"]
                        ) != "11"
                    ):
                        review_style_errors.append(
                            (
                                product_id,
                                "별첨0715_필수값 판정",
                                styles.get(
                                    header_index["별첨0715_필수값 판정"]
                                ),
                            )
                        )
                    if (
                        status in {"비교정보 제공", "규격 후보 없음"}
                        and styles.get(header_index["요청1_규격 상태"]) != "11"
                    ):
                        review_style_errors.append(
                            (
                                product_id,
                                "요청1_규격 상태",
                                styles.get(header_index["요청1_규격 상태"]),
                            )
                        )
                    inferred_header = "요청2_디자인 스타일 추론"
                    inferred_value = str(
                        values.get(header_index[inferred_header], "")
                    )
                    if (
                        inferred_value not in {"", "미확보", "해당없음"}
                        and styles.get(header_index[inferred_header]) != "11"
                    ):
                        review_style_errors.append(
                            (
                                product_id,
                                inferred_header,
                                styles.get(header_index[inferred_header]),
                            )
                        )
                    if status == "비교정보 제공":
                        comparison_rows += 1
                        count = int(
                            values.get(
                                header_index["요청1_규격 비교후보 수"], 0
                            )
                            or 0
                        )
                        candidate_sum += count
                        if count > 1:
                            multi_rows += 1
                        for header in COMPARISON_HEADERS:
                            parts = [
                                part.strip()
                                for part in str(
                                    values.get(header_index[header], "")
                                ).split(",")
                            ]
                            if len(parts) != count:
                                alignment_errors.append(
                                    (
                                        str(
                                            values.get(
                                                header_index["상품 ID"], ""
                                            )
                                        ),
                                        header,
                                        count,
                                        len(parts),
                                    )
                                )
                                break
                    if product_id in sample_ids:
                        samples[product_id] = {
                            header: values.get(header_index[header], "")
                            for header in (
                                "상품명",
                                *ORDERED_DIMENSION_HEADERS,
                            )
                        }
                element.clear()

        component_header_index: dict[str, int] = {}
        with archive.open("xl/worksheets/sheet8.xml") as stream:
            for _, element in ET.iterparse(stream, events=("end",)):
                if element.tag != f"{NS}row":
                    continue
                cells = element.findall(f"{NS}c")
                values = {
                    column_index(cell.get("r", "")): cell_value(cell)
                    for cell in cells
                }
                styles = {
                    column_index(cell.get("r", "")): cell.get("s")
                    for cell in cells
                }
                row_number = int(element.get("r", "0"))
                if row_number == 1:
                    component_headers = [
                        str(values.get(index, ""))
                        for index in range(max(values) + 1)
                    ]
                    component_header_index = {
                        header: index
                        for index, header in enumerate(component_headers)
                    }
                else:
                    component_rows += 1
                    product_id = str(
                        values.get(component_header_index["상품 ID"], "")
                    )
                    sequence = int(
                        values.get(component_header_index["구성품 순번"], 0)
                        or 0
                    )
                    component_sequences.setdefault(product_id, []).append(sequence)
                    resolution_status = str(
                        values.get(
                            component_header_index["구성품 규격 상태"],
                            "",
                        )
                    )
                    component_statuses[resolution_status] += 1
                    needs_review = str(
                        values.get(
                            component_header_index["사람 확인 필요"],
                            "",
                        )
                    )
                    if needs_review == "Y":
                        component_review_rows += 1
                        for header in (
                            "구성품 규격 상태",
                            "원천 유형",
                            "원천 위치",
                            "근거 원문",
                            "사람 확인 필요",
                        ):
                            if styles.get(component_header_index[header]) != "11":
                                component_style_errors.append(
                                    (
                                        product_id,
                                        header,
                                        styles.get(
                                            component_header_index[header]
                                        ),
                                    )
                                )
                    if product_id == "G25070005743":
                        component_sheet_sample.append(
                            {
                                header: values.get(index, "")
                                for header, index in component_header_index.items()
                            }
                        )
                element.clear()

    for product_id, sequences in component_sequences.items():
        expected = list(range(1, len(sequences) + 1))
        if sequences != expected:
            component_sequence_errors.append((product_id, sequences))

    expected_positions = list(
        range(
            header_index[ORDERED_DIMENSION_HEADERS[0]],
            header_index[ORDERED_DIMENSION_HEADERS[0]]
            + len(ORDERED_DIMENSION_HEADERS),
        )
    )
    actual_positions = [
        header_index[header] for header in ORDERED_DIMENSION_HEADERS
    ]
    if actual_positions != expected_positions:
        raise AssertionError(
            f"dimension comparison columns are not adjacent: {actual_positions}"
        )
    if alignment_errors:
        raise AssertionError(
            f"comma-aligned comparison columns failed: {alignment_errors[:5]}"
        )
    if review_style_errors:
        raise AssertionError(
            f"human-review styles failed: {review_style_errors[:5]}"
        )
    if attachment_status_errors:
        raise AssertionError(
            f"attachment 0715 status mapping failed: "
            f"{attachment_status_errors[:5]}"
        )
    if combination_style_errors:
        raise AssertionError(
            f"combination review styles failed: "
            f"{combination_style_errors[:5]}"
        )
    if component_style_errors:
        raise AssertionError(
            f"component review styles failed: {component_style_errors[:5]}"
        )
    if component_sequence_errors:
        raise AssertionError(
            f"component sequences failed: {component_sequence_errors[:5]}"
        )

    connection = sqlite3.connect(DB_PATH)
    progress = connection.execute(
        "SELECT * FROM vw_dimension_progress_authoritative"
    ).fetchone()
    expected_comparison_rows, expected_work_queue = connection.execute(
        """
        SELECT comparison_provided, total_remaining
        FROM vw_dimension_progress_authoritative
        """
    ).fetchone()
    database_candidate_count = connection.execute(
        "SELECT COUNT(*) FROM vw_dimension_comparison_candidates_current"
    ).fetchone()[0]
    work_queue_count = connection.execute(
        "SELECT COUNT(*) FROM vw_dimension_work_queue_current"
    ).fetchone()[0]
    integrity = connection.execute("PRAGMA integrity_check").fetchone()[0]
    database_combination_labels = {
        {
            "CONFIRMED": "Y",
            "CANDIDATE": "검토필요",
            "NOT_COMBINATION": "N",
        }[status]: count
        for status, count in connection.execute(
            """
            SELECT detection_status,COUNT(*)
            FROM vw_product_combination_current
            GROUP BY detection_status
            """
        )
    }
    database_component_output_labels = {
        {
            "ALL_COMPONENT_DIMENSIONS_CONFIRMED": "구성품 규격 확정",
            "COMPONENT_DIMENSIONS_CANDIDATE": "구성품 규격 후보",
            "PARTIAL_COMPONENT_DIMENSIONS": "구성품 규격 부분확보",
            "COMPONENT_DIMENSIONS_MISSING": "구성품 규격 미확보",
            "COMPONENT_PARSE_REQUIRED": "구성품명 분리 필요",
            "COMBINATION_REVIEW_REQUIRED": "조합 여부 검토",
            "NOT_APPLICABLE": "해당없음",
        }[status]: count
        for status, count in connection.execute(
            """
            SELECT component_output_status,COUNT(*)
            FROM vw_product_combination_current
            GROUP BY component_output_status
            """
        )
    }
    database_component_rows = connection.execute(
        "SELECT COUNT(*) FROM vw_product_component_dimensions_current"
    ).fetchone()[0]
    database_component_statuses = {
        {
            "API_CONFIRMED": "API 확정",
            "API_UNIT_INFERRED": "단위 추론 후보",
            "DIMENSION_MISSING": "규격 미확보",
            "COMPONENT_NAME_REQUIRED": "구성품명 미확보",
            "COMBINATION_REVIEW_REQUIRED": "조합 여부 검토",
        }[status]: count
        for status, count in connection.execute(
            """
            SELECT resolution_status,COUNT(*)
            FROM vw_product_component_dimensions_current
            GROUP BY resolution_status
            """
        )
    }
    connection.close()

    if comparison_rows != expected_comparison_rows:
        raise AssertionError((comparison_rows, expected_comparison_rows))
    if candidate_sum != database_candidate_count:
        raise AssertionError((candidate_sum, database_candidate_count))
    if statuses["사람 확인 필요"] != 0:
        raise AssertionError(statuses)
    if work_queue_count != expected_work_queue:
        raise AssertionError((work_queue_count, expected_work_queue))
    if dict(combination_labels) != database_combination_labels:
        raise AssertionError(
            (dict(combination_labels), database_combination_labels)
        )
    if dict(component_output_labels) != database_component_output_labels:
        raise AssertionError(
            (dict(component_output_labels), database_component_output_labels)
        )
    if component_rows != database_component_rows:
        raise AssertionError((component_rows, database_component_rows))
    if dict(component_statuses) != database_component_statuses:
        raise AssertionError(
            (dict(component_statuses), database_component_statuses)
        )
    expected_example = [
        {
            "상품 ID": "G25070005743",
            "구성품 순번": 1,
            "구성품명": "4인 소파",
            "W (mm)": 2910,
            "D (mm)": 1020,
            "H (mm)": 910,
            "구성품 규격 상태": "API 확정",
        },
        {
            "상품 ID": "G25070005743",
            "구성품 순번": 2,
            "구성품명": "스툴",
            "W (mm)": 740,
            "D (mm)": 660,
            "H (mm)": 410,
            "구성품 규격 상태": "API 확정",
        },
    ]
    actual_example = [
        {key: row.get(key) for key in expected_example[index]}
        for index, row in enumerate(component_sheet_sample)
    ]
    if actual_example != expected_example:
        raise AssertionError((actual_example, expected_example))

    print(
        json.dumps(
            {
                "output": str(OUTPUT),
                "file_bytes": OUTPUT.stat().st_size,
                "rows": sum(statuses.values()),
                "columns": len(headers),
                "dimension_column_positions": {
                    header: header_index[header] + 1
                    for header in ORDERED_DIMENSION_HEADERS
                },
                "dimension_header_styles": {
                    header: header_styles[header]
                    for header in ORDERED_DIMENSION_HEADERS
                },
                "dimension_status_counts": dict(statuses),
                "attachment_0715_status_counts": dict(attachment_statuses),
                "comparison_rows": comparison_rows,
                "multi_candidate_rows": multi_rows,
                "comparison_candidate_sum": candidate_sum,
                "alignment_errors": len(alignment_errors),
                "human_review_cells_main_sheet": human_review_cells,
                "human_review_style_errors": len(review_style_errors),
                "attachment_0715_status_errors": len(
                    attachment_status_errors
                ),
                "sheets": sheet_names,
                "combination_labels": dict(combination_labels),
                "component_output_labels": dict(component_output_labels),
                "component_rows": component_rows,
                "component_statuses": dict(component_statuses),
                "component_review_rows": component_review_rows,
                "combination_style_errors": len(combination_style_errors),
                "component_style_errors": len(component_style_errors),
                "component_sequence_errors": len(component_sequence_errors),
                "component_sample_G25070005743": component_sheet_sample,
                "database_progress": progress,
                "database_work_queue": work_queue_count,
                "database_integrity": integrity,
                "samples": samples,
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
