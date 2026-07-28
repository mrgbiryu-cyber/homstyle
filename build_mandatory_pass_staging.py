from __future__ import annotations

import json
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any

import build_homestyle_bulk_workbook as workbook
from bulk_homestyle_collect import DB_PATH, RUN_DIR, unpack


TABLE = "stg_mandatory_pass"
LATEST_OUTPUT = RUN_DIR / "mandatory_staging_latest.json"
STANDARD_SOURCE = "(발췌) 별첨. 3D 애셋 생성운용 자동화 필요데이터 _0715_ing_R.pptx"
STANDARD_VERSION = "0715_ing_R + set-ID-optional-2026-07-22"
PARSER_VERSION = "mandatory-v2; set-component-id-informational; dimension-confirmed-regex-v2; L-deferred"


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def status_value(value: Any) -> int | None:
    if value == "충족":
        return 1
    if value == workbook.MISSING:
        return 0
    return None


def source_value(value: Any) -> Any:
    if value in (None, "", workbook.MISSING, workbook.NOT_APPLICABLE):
        return None
    return value


def sheet_rows(sheets: list[tuple], name: str) -> list[list[Any]]:
    return next(sheet[1] for sheet in sheets if sheet[0] == name)


def create_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        f"""
        CREATE TABLE IF NOT EXISTS {TABLE} (
            snapshot_id TEXT NOT NULL,
            assessed_at TEXT NOT NULL,
            is_current INTEGER NOT NULL DEFAULT 1,
            standard_source TEXT NOT NULL,
            standard_version TEXT NOT NULL,
            parser_version TEXT NOT NULL,
            system TEXT NOT NULL,
            product_id TEXT NOT NULL,
            product_name TEXT,
            final_status TEXT NOT NULL,
            fulfilled_count INTEGER NOT NULL,
            required_count INTEGER NOT NULL,
            missing_fields TEXT,
            collection_error INTEGER NOT NULL DEFAULT 0,
            image_requirement_status TEXT,
            is_set_applicable INTEGER NOT NULL DEFAULT 0,
            brand_ok INTEGER,
            category_ok INTEGER,
            id_ok INTEGER,
            image_required_ok INTEGER,
            color_ok INTEGER,
            size_wdh_ok INTEGER,
            appliance_rear_image_ok INTEGER,
            appliance_additional_ok INTEGER,
            set_component_ids_ok INTEGER,
            brand_value TEXT,
            mid_category_value TEXT,
            small_category_value TEXT,
            representative_image_url TEXT,
            rolling_image_count INTEGER,
            rolling_image_urls_json TEXT,
            color_value TEXT,
            w_mm REAL,
            d_mm REAL,
            h_mm REAL,
            appliance_additional_value TEXT,
            set_component_ids_value TEXT,
            PRIMARY KEY (snapshot_id, product_id)
        );
        CREATE INDEX IF NOT EXISTS idx_stg_mandatory_current_status
            ON {TABLE}(is_current, final_status);
        CREATE INDEX IF NOT EXISTS idx_stg_mandatory_current_product
            ON {TABLE}(is_current, product_id);

        DROP VIEW IF EXISTS vw_mandatory_pass_current_summary;
        CREATE VIEW vw_mandatory_pass_current_summary AS
        SELECT snapshot_id, assessed_at, final_status, COUNT(*) AS product_count
        FROM {TABLE}
        WHERE is_current=1
        GROUP BY snapshot_id, assessed_at, final_status;

        DROP VIEW IF EXISTS vw_mandatory_pass_current_missing;
        CREATE VIEW vw_mandatory_pass_current_missing AS
        SELECT '브랜드' AS required_field, COUNT(*) AS product_count
          FROM {TABLE} WHERE is_current=1 AND brand_ok=0
        UNION ALL
        SELECT '카테고리(중·소)', COUNT(*)
          FROM {TABLE} WHERE is_current=1 AND category_ok=0
        UNION ALL
        SELECT 'ID', COUNT(*)
          FROM {TABLE} WHERE is_current=1 AND id_ok=0
        UNION ALL
        SELECT '대표·롤링 이미지 URL', COUNT(*)
          FROM {TABLE} WHERE is_current=1 AND image_required_ok=0
        UNION ALL
        SELECT '색상', COUNT(*)
          FROM {TABLE} WHERE is_current=1 AND color_ok=0
        UNION ALL
        SELECT '사이즈(W/D/H)', COUNT(*)
          FROM {TABLE} WHERE is_current=1 AND size_wdh_ok=0
        UNION ALL
        SELECT '가전 후면 이미지', COUNT(*)
          FROM {TABLE} WHERE is_current=1 AND appliance_rear_image_ok=0
        UNION ALL
        SELECT '가전 추가사항', COUNT(*)
          FROM {TABLE} WHERE is_current=1 AND appliance_additional_ok=0
        UNION ALL
        SELECT '상품 수집오류', COUNT(*)
          FROM {TABLE} WHERE is_current=1 AND collection_error=1;
        """
    )


def build_snapshot() -> dict[str, Any]:
    sheets, workbook_meta = workbook.build_rows()
    main_rows = sheet_rows(sheets, "01_상품별_요구필드")
    mandatory_rows = sheet_rows(sheets, "04_필수값_판정")
    main_header = {name: index for index, name in enumerate(main_rows[0])}
    mandatory_header = {name: index for index, name in enumerate(mandatory_rows[0])}
    rolling_headers = [
        name for name in main_rows[0]
        if str(name).startswith("요청1_롤링 이미지 URL ")
    ]
    main_by_id = {
        str(row[main_header["상품 ID"]]): row for row in main_rows[1:]
    }
    assessed_at = now_text()
    snapshot_id = assessed_at.replace(":", "").replace("+", "_")
    inserts = []
    for row in mandatory_rows[1:]:
        product_id = str(row[mandatory_header["상품 ID"]])
        main = main_by_id[product_id]
        set_status = status_value(row[mandatory_header["세트 구성 실제 ID"]])
        inserts.append(
            (
                snapshot_id,
                assessed_at,
                1,
                STANDARD_SOURCE,
                STANDARD_VERSION,
                PARSER_VERSION,
                row[mandatory_header["구분"]],
                product_id,
                row[mandatory_header["상품명"]],
                row[mandatory_header["필수값 판정"]],
                int(row[mandatory_header["충족수"]]),
                int(row[mandatory_header["대상수"]]),
                source_value(row[mandatory_header["보강대상 필드"]]),
                0,
                row[mandatory_header["이미지 필수 대체판정"]],
                int(set_status is not None),
                status_value(row[mandatory_header["브랜드"]]),
                status_value(row[mandatory_header["카테고리(중·소)"]]),
                status_value(row[mandatory_header["ID"]]),
                status_value(row[mandatory_header["대표·롤링 이미지 URL"]]),
                status_value(row[mandatory_header["색상"]]),
                status_value(row[mandatory_header["사이즈(W/D/H)"]]),
                status_value(row[mandatory_header["가전 후면 이미지"]]),
                status_value(row[mandatory_header["가전 추가사항"]]),
                set_status,
                source_value(main[main_header["요청1_브랜드명"]]),
                source_value(main[main_header["요청1_중카테고리"]]),
                source_value(main[main_header["요청1_소카테고리"]]),
                source_value(main[main_header["요청1_대표 이미지 URL"]]),
                int(main[main_header["요청1_롤링 이미지 수"]] or 0),
                json.dumps(
                    [
                        source_value(main[main_header[header]])
                        for header in rolling_headers
                        if source_value(main[main_header[header]]) is not None
                    ],
                    ensure_ascii=False,
                ),
                source_value(main[main_header["요청1_제품 색상"]]),
                source_value(main[main_header["요청1_W (mm)"]]),
                source_value(main[main_header["요청1_D (mm)"]]),
                source_value(main[main_header["요청1_H (mm)"]]),
                source_value(main[main_header["요청1_설치 타입 구분"]]),
                source_value(main[main_header["요청1_세트 구성 ID 리스트"]]),
            )
        )

    connection = sqlite3.connect(DB_PATH)
    create_schema(connection)
    error_rows = connection.execute(
        """
        SELECT p.product_id,p.listing_blob
        FROM products p JOIN sources s ON s.product_id=p.product_id
        WHERE s.goods_status IS NULL OR s.goods_status != 200
        ORDER BY p.product_id
        """
    ).fetchall()
    for product_id, listing_blob in error_rows:
        listing = unpack(listing_blob) or {}
        inserts.append(
            (
                snapshot_id, assessed_at, 1, STANDARD_SOURCE, STANDARD_VERSION,
                PARSER_VERSION, "홈스타일", product_id,
                listing.get("productName") or product_id,
                workbook.MANDATORY_REINFORCEMENT, 0, 6, "상품 수집오류", 1,
                workbook.MISSING, 0,
                None, None, 1, None, None, None, None, None, None,
                None, None, None, None, None, None, None, None, None, None,
                None, None,
            )
        )

    placeholders = ",".join("?" for _ in range(37))
    with connection:
        connection.execute(f"UPDATE {TABLE} SET is_current=0 WHERE is_current=1")
        connection.executemany(
            f"INSERT INTO {TABLE} VALUES ({placeholders})",
            inserts,
        )

    status_counts = {
        status: count
        for status, count in connection.execute(
            f"SELECT final_status,COUNT(*) FROM {TABLE} "
            "WHERE snapshot_id=? GROUP BY final_status",
            (snapshot_id,),
        )
    }
    missing_counts = {
        field: count
        for field, count in connection.execute(
            "SELECT required_field,product_count "
            "FROM vw_mandatory_pass_current_missing WHERE product_count>0"
        )
    }
    total = sum(status_counts.values())
    result = {
        "database": str(DB_PATH),
        "table": TABLE,
        "snapshot_id": snapshot_id,
        "assessed_at": assessed_at,
        "rows": total,
        "status_counts": status_counts,
        "pass_rate": status_counts.get(workbook.MANDATORY_PASS, 0) / total if total else 0,
        "missing_required_groups": missing_counts,
        "views": [
            "vw_mandatory_pass_current_summary",
            "vw_mandatory_pass_current_missing",
        ],
        "workbook_meta": workbook_meta,
        "excel_written": False,
    }
    connection.close()
    LATEST_OUTPUT.write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return result


if __name__ == "__main__":
    print(json.dumps(build_snapshot(), ensure_ascii=False, indent=2))
