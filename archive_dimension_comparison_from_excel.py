from __future__ import annotations

import sqlite3
from datetime import datetime
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from build_dimension_resolution_ledger import LEDGER_VERSION, init_schema
from bulk_homestyle_collect import DB_PATH
from validate_dimension_comparison_workbook import (
    NS,
    OUTPUT,
    cell_value,
    column_index,
)


COMPARISON_COLUMNS = {
    "w": "요청1_규격 비교 W (mm)",
    "d": "요청1_규격 비교 D (mm)",
    "h": "요청1_규격 비교 H (mm)",
    "target": "요청1_규격 비교 대상/옵션",
    "raw": "요청1_규격 비교 원문",
}


def split_values(value: object) -> list[str]:
    return [part.strip() for part in str(value or "").split(",")]


def numeric_value(value: str) -> float | None:
    if value in {"", "미확보", "해당없음"}:
        return None
    return float(value)


def main() -> None:
    timestamp = datetime.now().astimezone().isoformat(timespec="seconds")
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    init_schema(connection)

    headers: list[str] = []
    header_index: dict[str, int] = {}
    archived_products = 0
    archived_candidates = 0
    with ZipFile(OUTPUT) as archive:
        with archive.open("xl/worksheets/sheet2.xml") as stream:
            for _, element in ET.iterparse(stream, events=("end",)):
                if element.tag != f"{NS}row":
                    continue
                cells = element.findall(f"{NS}c")
                values = {
                    column_index(cell.get("r", "")): cell_value(cell)
                    for cell in cells
                }
                if int(element.get("r", "0")) == 1:
                    headers = [
                        str(values.get(index, ""))
                        for index in range(max(values) + 1)
                    ]
                    header_index = {
                        header: index for index, header in enumerate(headers)
                    }
                    element.clear()
                    continue
                if (
                    values.get(header_index["요청1_규격 상태"])
                    != "비교정보 제공"
                ):
                    element.clear()
                    continue

                product_id = str(values.get(header_index["상품 ID"], ""))
                count = int(
                    values.get(header_index["요청1_규격 비교후보 수"], 0)
                    or 0
                )
                columns = {
                    key: split_values(values.get(header_index[header], ""))
                    for key, header in COMPARISON_COLUMNS.items()
                }
                if any(len(parts) != count for parts in columns.values()):
                    raise ValueError(
                        f"comparison column alignment failed: {product_id}"
                    )
                connection.execute(
                    "DELETE FROM fact_dimension_comparison_candidate "
                    "WHERE product_id=?",
                    (product_id,),
                )
                for index in range(count):
                    candidate_no = index + 1
                    connection.execute(
                        """
                        INSERT INTO fact_dimension_comparison_candidate (
                            product_id, candidate_no, is_representative,
                            comparison_target, w_mm, d_mm, h_mm,
                            shape_type, normalized_axis_mapping, source_type,
                            raw_notation, context_text, candidate_key,
                            source_snapshot_id, captured_at, ledger_version
                        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                        """,
                        (
                            product_id,
                            candidate_no,
                            1 if candidate_no == 1 else 0,
                            columns["target"][index],
                            numeric_value(columns["w"][index]),
                            numeric_value(columns["d"][index]),
                            numeric_value(columns["h"][index]),
                            "",
                            "",
                            "EXCEL_COMPARISON_ARCHIVE",
                            columns["raw"][index],
                            columns["raw"][index],
                            f"excel:{product_id}:{candidate_no}",
                            "dimension_comparison_excel_20260723",
                            timestamp,
                            LEDGER_VERSION,
                        ),
                    )
                    archived_candidates += 1
                archived_products += 1
                element.clear()

    connection.commit()
    integrity = connection.execute("PRAGMA integrity_check").fetchone()[0]
    connection.close()
    print(
        {
            "products": archived_products,
            "candidates": archived_candidates,
            "integrity": integrity,
        }
    )


if __name__ == "__main__":
    main()
