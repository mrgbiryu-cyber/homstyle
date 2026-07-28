from __future__ import annotations

import json
import sqlite3
from collections import Counter
from datetime import datetime
from typing import Any

from bulk_homestyle_collect import DB_PATH


RUN_NAME = "dimension_remaining_reason_v1"
PASS1_RUN = "dimension_scan_pass1_all_detail_images_v1"
PASS2_RUN = "dimension_scan_pass2_layout_v1"
TARGETED_RUN = "dimension_targeted_reocr_remaining_v1"

REASONS = {
    "NO_DETAIL_IMAGES": (
        "상세 이미지 없음",
        "상품 API/HTML의 상세 이미지 원천을 추가 확인",
    ),
    "PASS1_OCR_INCOMPLETE": (
        "1차 OCR 오류·미완료",
        "이미지 재다운로드 후 OCR 재실행",
    ),
    "PASS1_NO_SIZE_SIGNAL": (
        "전체 이미지에서 크기 신호 없음",
        "추가 이미지/API/제조사 규격 원천 탐색",
    ),
    "HEADING_NOT_LOCATED": (
        "SIZE/INFO 등 대상 영역 미탐지",
        "제목 없는 도면을 찾는 시각 영역 탐지 보강",
    ),
    "TARGET_REGION_OCR_NO_VALUE": (
        "대상 영역 OCR 후 숫자 규격 미인식",
        "다른 OCR 엔진·해상도·전처리로 재시도",
    ),
    "CANDIDATES_ALL_EXCLUDED_OR_INVALID": (
        "후보가 제외 또는 유효값 판정 실패",
        "원문 이미지 수동 확인 또는 별도 규칙 검토",
    ),
}


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def create_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS stg_dimension_remaining_reason (
            snapshot_id TEXT NOT NULL,
            analyzed_at TEXT NOT NULL,
            is_current INTEGER NOT NULL,
            run_name TEXT NOT NULL,
            product_id TEXT NOT NULL,
            product_name TEXT,
            mid_category TEXT,
            small_category TEXT,
            resolution_status TEXT,
            reason_code TEXT NOT NULL,
            reason_name TEXT NOT NULL,
            next_action TEXT,
            pass1_total_image_count INTEGER,
            pass1_ocr_success_count INTEGER,
            pass1_ocr_error_count INTEGER,
            pass1_candidate_image_count INTEGER,
            pass2_status TEXT,
            targeted_crop_count INTEGER,
            targeted_candidate_count INTEGER,
            PRIMARY KEY(snapshot_id, product_id)
        );
        CREATE INDEX IF NOT EXISTS idx_dimension_remaining_reason_current
            ON stg_dimension_remaining_reason(is_current, reason_code);

        DROP VIEW IF EXISTS vw_dimension_remaining_reason_current;
        CREATE VIEW vw_dimension_remaining_reason_current AS
        SELECT *
        FROM stg_dimension_remaining_reason
        WHERE is_current = 1;

        DROP VIEW IF EXISTS vw_dimension_remaining_reason_summary;
        CREATE VIEW vw_dimension_remaining_reason_summary AS
        SELECT
            reason_code,
            reason_name,
            next_action,
            COUNT(*) AS product_count
        FROM stg_dimension_remaining_reason
        WHERE is_current = 1
        GROUP BY reason_code, reason_name, next_action
        ORDER BY product_count DESC, reason_code;
        """
    )


def main() -> None:
    analyzed_at = now_text()
    snapshot_id = analyzed_at.replace(":", "").replace("+", "_")
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA journal_mode=WAL")
    create_schema(connection)

    pass1 = {
        row["product_id"]: row
        for row in connection.execute(
            "SELECT * FROM stg_dimension_scan_pass1_product WHERE run_name = ?",
            (PASS1_RUN,),
        )
    }
    pass2 = {
        row["product_id"]: row
        for row in connection.execute(
            "SELECT * FROM stg_dimension_scan_pass2_product WHERE run_name = ?",
            (PASS2_RUN,),
        )
    }
    crop_counts = {
        row["product_id"]: int(row["n"])
        for row in connection.execute(
            """
            SELECT product_id, COUNT(*) AS n
            FROM stg_dimension_targeted_ocr_crop
            WHERE run_name = ?
            GROUP BY product_id
            """,
            (TARGETED_RUN,),
        )
    }
    candidate_counts = {
        row["product_id"]: int(row["n"])
        for row in connection.execute(
            """
            SELECT product_id, COUNT(*) AS n
            FROM stg_dimension_targeted_ocr_candidate
            WHERE run_name = ?
            GROUP BY product_id
            """,
            (TARGETED_RUN,),
        )
    }

    inserts: list[tuple[Any, ...]] = []
    counts: Counter[str] = Counter()
    for ledger in connection.execute(
        """
        SELECT *
        FROM fact_dimension_resolution_ledger
        WHERE resolution_status = 'NO_CANDIDATE'
        ORDER BY product_id
        """
    ):
        product_id = ledger["product_id"]
        p1 = pass1.get(product_id)
        p2 = pass2.get(product_id)
        total_images = int(p1["total_image_count"] or 0) if p1 else 0
        success_images = int(p1["ocr_success_image_count"] or 0) if p1 else 0
        error_images = int(p1["ocr_error_image_count"] or 0) if p1 else 0
        candidate_images = int(p1["candidate_image_count"] or 0) if p1 else 0
        crop_count = crop_counts.get(product_id, 0)
        targeted_candidate_count = candidate_counts.get(product_id, 0)

        if total_images == 0:
            reason_code = "NO_DETAIL_IMAGES"
        elif candidate_images == 0 and error_images > 0:
            reason_code = "PASS1_OCR_INCOMPLETE"
        elif candidate_images == 0:
            reason_code = "PASS1_NO_SIZE_SIGNAL"
        elif crop_count == 0:
            reason_code = "HEADING_NOT_LOCATED"
        elif targeted_candidate_count == 0:
            reason_code = "TARGET_REGION_OCR_NO_VALUE"
        else:
            reason_code = "CANDIDATES_ALL_EXCLUDED_OR_INVALID"
        reason_name, next_action = REASONS[reason_code]
        counts[reason_code] += 1
        inserts.append(
            (
                snapshot_id,
                analyzed_at,
                1,
                RUN_NAME,
                product_id,
                ledger["product_name"],
                ledger["mid_category"],
                ledger["small_category"],
                ledger["resolution_status"],
                reason_code,
                reason_name,
                next_action,
                total_images,
                success_images,
                error_images,
                candidate_images,
                p2["pass2_status"] if p2 else "NO_PASS2",
                crop_count,
                targeted_candidate_count,
            )
        )

    with connection:
        connection.execute(
            "UPDATE stg_dimension_remaining_reason SET is_current = 0 "
            "WHERE is_current = 1"
        )
        connection.executemany(
            """
            INSERT INTO stg_dimension_remaining_reason
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,
            inserts,
        )

    progress = dict(
        connection.execute(
            "SELECT * FROM vw_dimension_progress_authoritative"
        ).fetchone()
    )
    integrity = connection.execute("PRAGMA integrity_check").fetchone()[0]
    result = {
        "snapshot_id": snapshot_id,
        "rows": len(inserts),
        "reason_counts": dict(counts.most_common()),
        "authoritative_progress": progress,
        "integrity_check": integrity,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    connection.close()


if __name__ == "__main__":
    main()
