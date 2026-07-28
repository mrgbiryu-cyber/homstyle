from __future__ import annotations

import json
import sqlite3
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

from bulk_homestyle_collect import pack, unpack
from low_dimension_quality_policy import LOW_DIMENSION_THRESHOLD_MM, low_axes


ROOT = Path(__file__).resolve().parent
DB_PATH = ROOT / "homestyle_bulk_run" / "homestyle_bulk.sqlite"
RUN_NAME = "low_dimension_quality_repair_20260728"
SUMMARY_PATH = ROOT / "homestyle_bulk_run" / f"{RUN_NAME}_summary.json"


@dataclass(frozen=True)
class Option:
    label: str
    w_mm: float
    d_mm: float
    h_mm: float
    primary: bool = False


@dataclass(frozen=True)
class Repair:
    product_id: str
    w_mm: float
    d_mm: float
    h_mm: float
    raw_notation: str
    reason: str
    evidence_text: str
    evidence_url: str
    evidence_type: str = "OCR_IMAGE"
    options: tuple[Option, ...] = field(default_factory=tuple)


REPAIRS = (
    Repair(
        "G25070001578",
        3065,
        1200,
        970,
        "W3065 × D1200 × H970(mm)",
        "알로소 공통 배송 이미지의 엘리베이터 규격 W900×D1200×H2300을 제품 규격으로 잘못 확정",
        "CARPONE GRAN 4인 | INFORMATION | 사이즈 W3065 × D1200 × H970(mm)",
        "https://shop-cdn.sidiz.com/_outside/Alloso/2024ver/SOFA/CARPONE/CARPONEGRAN_4Seater_Fabric.jpg",
    ),
    Repair(
        "G25090018467",
        2770,
        1100,
        605,
        "W2770 × D1100 × H605(mm)",
        "알로소 공통 배송 이미지의 엘리베이터 규격 W900×D1200×H2300을 제품 규격으로 잘못 확정",
        "HOLIDAY 4인 오픈쇼터 | INFORMATION | 사이즈 W2770 × D1100 × H605(mm)",
        "https://shop-cdn.sidiz.com/_outside/Alloso/2024ver/SOFA/HOLIDAY/HOLIDAY_4SeaterOpenShorter.jpg",
    ),
    Repair(
        "G26020033019",
        1133,
        900,
        300,
        "W1133 × D900 × H300(mm)",
        "알로소 공통 배송 이미지의 엘리베이터 규격 W900×D1200×H2300을 제품 규격으로 잘못 확정",
        "LAYER TABLE | DIMENSIONS 도면 | TOP W1133 × D900 / FRONT H300(mm)",
        "https://shop-cdn.sidiz.com/_outside/Alloso/TABLE/LAYER%20TABLE.jpg",
    ),
    Repair(
        "G26010030486",
        330,
        475,
        660,
        "S: 330×475×660mm, M: 510×395×580mm",
        "알로소 공통 배송 이미지의 엘리베이터 규격 W900×D1200×H2300을 제품 규격으로 잘못 확정; 실제 상품은 S/M 복수 규격",
        "PEILI TABLE | Size | S 330×475×660mm / M 510×395×580mm",
        "https://shop-cdn.sidiz.com/_outside/Alloso/2024ver/TABLE/PEILI.jpg",
        options=(
            Option("S", 330, 475, 660, True),
            Option("M", 510, 395, 580),
        ),
    ),
    Repair(
        "G25120029838",
        140,
        140,
        140,
        "크기 14×14×14cm",
        "API 상품정보고시의 210×261×19mm는 상품 상세 HTML의 동일 모델 규격과 충돌하며, 상세 HTML에 노블우드휴지케이스(소) 14×14×14cm가 명시됨",
        "품명 및 모델명 노블우드휴지케이스(소) | 크기 14×14×14cm",
        "https://homestyle.lge.co.kr/item?productId=G25120029838",
        evidence_type="HTML",
    ),
)


VALID_LOW_EVIDENCE: dict[str, tuple[str, str, str]] = {
    "G25070000617": (
        "API_PRODUCT_NOTIFICATION",
        "크기 71 × 50 × D2cm",
        "액자·보드류의 두께 D=20mm로 구조상 정상",
    ),
    "G25070001390": (
        "API_PRODUCT_NOTIFICATION",
        "크기 H30 × W25 × D1.6cm",
        "액자 두께 D=16mm로 구조상 정상",
    ),
    "G25070001396": (
        "API_PRODUCT_NOTIFICATION",
        "크기 H30 × W25 × D1.6cm / Image area 25×20cm",
        "액자 두께 D=16mm로 구조상 정상",
    ),
    "G25080009387": (
        "OCR_PRODUCT_SIZE",
        "SIZE 제품 상세 사이즈 W118 × D2 × H160cm",
        "러그 두께 D=20mm로 제품 SIZE 문맥에서 확인",
    ),
    "G26020033434": (
        "OCR_PRODUCT_SIZE",
        "SIZE W1500 × D1500 × H10(mm)",
        "러그 높이·두께 H=10mm로 제품 SIZE 문맥에서 확인",
    ),
    "G26070046697": (
        "OCR_PRODUCT_SIZE",
        "SIZE W1600 × D2300 × H10(mm)",
        "러그 높이·두께 H=10mm로 제품 SIZE 문맥에서 확인",
    ),
    "G26070046698": (
        "OCR_PRODUCT_SIZE",
        "SIZE W1500 × D1500 × H10(mm)",
        "러그 높이·두께 H=10mm로 제품 SIZE 문맥에서 확인",
    ),
    "G25070002832": (
        "OCR_PRODUCT_SIZE",
        "Product Description | Size W400 × D20 × H110mm",
        "벽 선반 두께 D=20mm로 제품 Size 문맥에서 확인",
    ),
    "G25100021623": (
        "API_PRODUCT_NOTIFICATION",
        "크기 2 × 15 × 49cm",
        "넥타이걸이 폭 W=20mm로 구조상 정상",
    ),
    "G25120029829": (
        "API_PRODUCT_NOTIFICATION",
        "크기 42 × 0.4 × 21.5cm",
        "알루미늄 옷걸이 두께 D=4mm로 구조상 정상",
    ),
    "G26030035994": (
        "API_PRODUCT_NOTIFICATION",
        "크기 가로60 × 세로2 × 높이54cm",
        "거울 두께 D=20mm로 구조상 정상",
    ),
}


def now_text() -> str:
    return datetime.now(ZoneInfo("Asia/Seoul")).isoformat(timespec="seconds")


def init_schema(connection: sqlite3.Connection) -> None:
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS stg_dimension_low_value_audit (
            run_name TEXT NOT NULL,
            product_id TEXT NOT NULL,
            audited_at TEXT NOT NULL,
            is_current INTEGER NOT NULL,
            original_w_mm REAL,
            original_d_mm REAL,
            original_h_mm REAL,
            low_axes TEXT,
            threshold_mm REAL NOT NULL,
            audit_status TEXT NOT NULL,
            audit_reason TEXT,
            evidence_type TEXT,
            evidence_url TEXT,
            evidence_text TEXT,
            action_taken TEXT,
            PRIMARY KEY (run_name, product_id)
        )
        """
    )
    connection.execute(
        """
        CREATE INDEX IF NOT EXISTS idx_dimension_low_value_audit_current
        ON stg_dimension_low_value_audit(is_current, audit_status, product_id)
        """
    )
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS hist_dimension_quality_repair (
            run_name TEXT NOT NULL,
            product_id TEXT NOT NULL,
            repaired_at TEXT NOT NULL,
            old_resolution_status TEXT,
            new_resolution_status TEXT,
            old_w_mm REAL,
            old_d_mm REAL,
            old_h_mm REAL,
            new_w_mm REAL,
            new_d_mm REAL,
            new_h_mm REAL,
            raw_notation TEXT,
            repair_reason TEXT,
            evidence_url TEXT,
            evidence_text TEXT,
            PRIMARY KEY (run_name, product_id)
        )
        """
    )


def add_reinforcement(
    connection: sqlite3.Connection,
    repair: Repair,
    timestamp: str,
) -> None:
    source = connection.execute(
        "SELECT ocr_status, ocr_blob FROM sources WHERE product_id=?",
        (repair.product_id,),
    ).fetchone()
    if source is None:
        raise RuntimeError(f"sources row missing: {repair.product_id}")
    ocr_status, previous_blob = source
    connection.execute(
        """
        INSERT OR IGNORE INTO source_ocr_dimension_backup (
            run_name, product_id, backed_up_at,
            previous_ocr_status, previous_ocr_blob
        ) VALUES (?,?,?,?,?)
        """,
        (RUN_NAME, repair.product_id, timestamp, ocr_status, previous_blob),
    )
    payload = unpack(previous_blob) or {}
    reinforcements = [
        item
        for item in (payload.get("dimension_reinforcements") or [])
        if item.get("run_name") != RUN_NAME
    ]
    item: dict[str, Any] = {
        "run_name": RUN_NAME,
        "product_id": repair.product_id,
        "verification_method": (
            "stored_detail_html_product_identity_match"
            if repair.evidence_type == "HTML"
            else "stored_detail_image_product_identity_and_size_section"
        ),
        "verified_at": timestamp,
        "source_url": repair.evidence_url,
        "raw_notation": repair.raw_notation,
        "evidence_text": repair.evidence_text,
        "repair_reason": repair.reason,
        "w_mm": repair.w_mm,
        "d_mm": repair.d_mm,
        "h_mm": repair.h_mm,
        "dimensions": [[repair.w_mm, repair.d_mm, repair.h_mm]],
    }
    if repair.options:
        item["dimension_options"] = [
            {
                "label": option.label,
                "w_mm": option.w_mm,
                "d_mm": option.d_mm,
                "h_mm": option.h_mm,
                "is_primary": option.primary,
            }
            for option in repair.options
        ]
    reinforcements.append(item)
    payload["dimension_reinforcements"] = reinforcements
    if repair.evidence_type != "HTML":
        selected = dict(payload.get("selected") or {})
        selected["url"] = repair.evidence_url
        selected["tag"] = "PRODUCT_SIZE_RECHECK"
        selected["alt"] = repair.evidence_text
        payload["selected"] = selected
    connection.execute(
        "UPDATE sources SET ocr_blob=?, ocr_at=? WHERE product_id=?",
        (pack(payload), timestamp, repair.product_id),
    )


def replace_options(
    connection: sqlite3.Connection,
    repair: Repair,
    timestamp: str,
) -> None:
    options = repair.options or (
        Option("대표 규격", repair.w_mm, repair.d_mm, repair.h_mm, True),
    )
    connection.execute(
        "DELETE FROM fact_dimension_resolution_option WHERE product_id=?",
        (repair.product_id,),
    )
    for option_no, option in enumerate(options, start=1):
        connection.execute(
            """
            INSERT INTO fact_dimension_resolution_option (
                product_id, option_no, option_label, is_primary,
                w_mm, d_mm, h_mm, d_applicability, shape_type,
                diameter_mm, normalized_axis_mapping,
                resolution_rule_code, source_candidate_key,
                source_type, evidence_text, locked_at, ledger_version
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,
            (
                repair.product_id,
                option_no,
                option.label,
                int(option.primary),
                option.w_mm,
                option.d_mm,
                option.h_mm,
                "APPLICABLE",
                "RECTANGLE",
                None,
                "MANUAL_VERIFIED->W,D,H",
                "LOW_DIMENSION_RECHECK",
                "",
                (
                    "HTML_PRODUCT_SPEC"
                    if repair.evidence_type == "HTML"
                    else "OCR_IMAGE_MANUAL_VERIFICATION"
                ),
                f"{repair.evidence_text} | 원문: {repair.raw_notation}",
                timestamp,
                "dimension-resolution-ledger-v2-comparison",
            ),
        )


def repair_one(
    connection: sqlite3.Connection,
    repair: Repair,
    timestamp: str,
) -> dict[str, Any]:
    old = connection.execute(
        "SELECT * FROM fact_dimension_resolution_ledger WHERE product_id=?",
        (repair.product_id,),
    ).fetchone()
    if old is None:
        raise RuntimeError(f"ledger row missing: {repair.product_id}")

    add_reinforcement(connection, repair, timestamp)
    connection.execute(
        """
        UPDATE stg_dimension_classification_master
        SET classification_complete=1,
            classification_status='COMPLETE',
            classification_stage='LOW_DIMENSION_RECHECK',
            classification_code='B03_LOW_DIMENSION_RECHECK',
            classification_name='20mm 이하 의심값 제품별 원천 재검증',
            dimension_resolution_status='MANUAL_CONFIRMED_WDH',
            dimension_value_confirmed=1,
            confirmed_w_mm=?,
            confirmed_d_mm=?,
            confirmed_h_mm=?,
            quality_flag='LOW_DIMENSION_REPAIRED_20260728',
            requires_followup=0,
            followup_priority=NULL,
            next_action='완료',
            best_image_url=?,
            evidence_text=?
        WHERE product_id=? AND is_current=1
        """,
        (
            repair.w_mm,
            repair.d_mm,
            repair.h_mm,
            repair.evidence_url,
            repair.evidence_text,
            repair.product_id,
        ),
    )
    connection.execute(
        """
        UPDATE fact_dimension_resolution_ledger
        SET resolution_status='MANUAL_CONFIRMED',
            is_locked=1,
            pass_status='PASS',
            needs_human_review=0,
            needs_ocr=0,
            locked_w_mm=?,
            locked_d_mm=?,
            locked_h_mm=?,
            locked_d_applicability='APPLICABLE',
            locked_shape_type='RECTANGLE',
            locked_option_count=?,
            candidate_w_mm=NULL,
            candidate_d_mm=NULL,
            candidate_h_mm=NULL,
            context_status='LOW_DIMENSION_MANUAL_RECHECK',
            resolution_rule_code='LOW_DIMENSION_RECHECK',
            resolution_source=?,
            source_snapshot_id=?,
            representative_candidate_key='',
            evidence_text=?,
            last_transition_at=?,
            updated_at=?
        WHERE product_id=?
        """,
        (
            repair.w_mm,
            repair.d_mm,
            repair.h_mm,
            len(repair.options) or 1,
            (
                "detail_html_product_spec"
                if repair.evidence_type == "HTML"
                else "detail_image_product_size_section"
            ),
            RUN_NAME,
            f"{repair.evidence_text} | 원문: {repair.raw_notation}",
            timestamp,
            timestamp,
            repair.product_id,
        ),
    )
    replace_options(connection, repair, timestamp)
    connection.execute(
        """
        INSERT OR REPLACE INTO hist_dimension_quality_repair (
            run_name, product_id, repaired_at,
            old_resolution_status, new_resolution_status,
            old_w_mm, old_d_mm, old_h_mm,
            new_w_mm, new_d_mm, new_h_mm,
            raw_notation, repair_reason, evidence_url, evidence_text
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """,
        (
            RUN_NAME,
            repair.product_id,
            timestamp,
            old["resolution_status"],
            "MANUAL_CONFIRMED",
            old["locked_w_mm"],
            old["locked_d_mm"],
            old["locked_h_mm"],
            repair.w_mm,
            repair.d_mm,
            repair.h_mm,
            repair.raw_notation,
            repair.reason,
            repair.evidence_url,
            repair.evidence_text,
        ),
    )
    connection.execute(
        """
        INSERT INTO hist_dimension_resolution_event (
            event_at, product_id, old_status, new_status,
            transition_reason, source_snapshot_id,
            old_w_mm, old_d_mm, old_h_mm,
            new_w_mm, new_d_mm, new_h_mm, ledger_version
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
        """,
        (
            timestamp,
            repair.product_id,
            old["resolution_status"],
            "MANUAL_CONFIRMED",
            "LOW_DIMENSION_RECHECK_REPAIR",
            RUN_NAME,
            old["locked_w_mm"],
            old["locked_d_mm"],
            old["locked_h_mm"],
            repair.w_mm,
            repair.d_mm,
            repair.h_mm,
            "dimension-resolution-ledger-v2-comparison",
        ),
    )
    return {
        "product_id": repair.product_id,
        "old": [old["locked_w_mm"], old["locked_d_mm"], old["locked_h_mm"]],
        "new": [repair.w_mm, repair.d_mm, repair.h_mm],
        "option_count": len(repair.options) or 1,
        "reason": repair.reason,
    }


def insert_audit_row(
    connection: sqlite3.Connection,
    *,
    timestamp: str,
    product_id: str,
    original_w_mm: float | None,
    original_d_mm: float | None,
    original_h_mm: float | None,
    audit_status: str,
    audit_reason: str,
    evidence_type: str,
    evidence_url: str,
    evidence_text: str,
    action_taken: str,
) -> None:
    connection.execute(
        """
        INSERT OR REPLACE INTO stg_dimension_low_value_audit (
            run_name, product_id, audited_at, is_current,
            original_w_mm, original_d_mm, original_h_mm,
            low_axes, threshold_mm, audit_status, audit_reason,
            evidence_type, evidence_url, evidence_text, action_taken
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """,
        (
            RUN_NAME,
            product_id,
            timestamp,
            1,
            original_w_mm,
            original_d_mm,
            original_h_mm,
            ",".join(low_axes(original_w_mm, original_d_mm, original_h_mm)),
            LOW_DIMENSION_THRESHOLD_MM,
            audit_status,
            audit_reason,
            evidence_type,
            evidence_url,
            evidence_text,
            action_taken,
        ),
    )


def main() -> None:
    timestamp = now_text()
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA journal_mode=WAL")
    connection.execute("PRAGMA synchronous=NORMAL")
    init_schema(connection)

    before_rows = list(
        connection.execute(
            """
            SELECT *
            FROM fact_dimension_resolution_ledger
            WHERE is_locked=1
              AND (
                    (locked_w_mm IS NOT NULL AND locked_w_mm<=?)
                 OR (locked_d_mm IS NOT NULL AND locked_d_mm<=?)
                 OR (locked_h_mm IS NOT NULL AND locked_h_mm<=?)
              )
            ORDER BY product_id
            """,
            (LOW_DIMENSION_THRESHOLD_MM,) * 3,
        )
    )
    before_by_id = {row["product_id"]: row for row in before_rows}
    expected_ids = set(VALID_LOW_EVIDENCE) | {
        repair.product_id for repair in REPAIRS
    }
    if set(before_by_id) != expected_ids:
        raise RuntimeError(
            "20mm 이하 완료군이 사전 검증한 16건과 다릅니다: "
            f"actual={sorted(before_by_id)}"
        )

    repairs_by_id = {repair.product_id: repair for repair in REPAIRS}
    connection.commit()
    repaired: list[dict[str, Any]] = []
    try:
        connection.execute("BEGIN IMMEDIATE")
        connection.execute(
            "UPDATE stg_dimension_low_value_audit SET is_current=0 WHERE is_current=1"
        )
        for product_id, row in before_by_id.items():
            pdp_url = (
                f"https://homestyle.lge.co.kr/item?productId={product_id}"
            )
            if product_id in VALID_LOW_EVIDENCE:
                evidence_type, evidence_text, reason = VALID_LOW_EVIDENCE[
                    product_id
                ]
                insert_audit_row(
                    connection,
                    timestamp=timestamp,
                    product_id=product_id,
                    original_w_mm=row["locked_w_mm"],
                    original_d_mm=row["locked_d_mm"],
                    original_h_mm=row["locked_h_mm"],
                    audit_status="VALID_THIN_DIMENSION_EVIDENCE_CONFIRMED",
                    audit_reason=reason,
                    evidence_type=evidence_type,
                    evidence_url=pdp_url,
                    evidence_text=evidence_text,
                    action_taken="규격값 유지; 20mm 이하 예외 근거 등록",
                )
                continue

            repair = repairs_by_id[product_id]
            repaired.append(repair_one(connection, repair, timestamp))
            insert_audit_row(
                connection,
                timestamp=timestamp,
                product_id=product_id,
                original_w_mm=row["locked_w_mm"],
                original_d_mm=row["locked_d_mm"],
                original_h_mm=row["locked_h_mm"],
                audit_status="INVALID_LOW_DIMENSION_CORRECTED",
                audit_reason=repair.reason,
                evidence_type=repair.evidence_type,
                evidence_url=repair.evidence_url,
                evidence_text=repair.evidence_text,
                action_taken=(
                    f"W={repair.w_mm}, D={repair.d_mm}, H={repair.h_mm}mm로 교체"
                ),
            )
        connection.commit()
    except Exception:
        connection.rollback()
        raise

    after_rows = list(
        connection.execute(
            """
            SELECT product_id, locked_w_mm, locked_d_mm, locked_h_mm
            FROM fact_dimension_resolution_ledger
            WHERE is_locked=1
              AND (
                    (locked_w_mm IS NOT NULL AND locked_w_mm<=?)
                 OR (locked_d_mm IS NOT NULL AND locked_d_mm<=?)
                 OR (locked_h_mm IS NOT NULL AND locked_h_mm<=?)
              )
            ORDER BY product_id
            """,
            (LOW_DIMENSION_THRESHOLD_MM,) * 3,
        )
    )
    status_counts = {
        row[0]: row[1]
        for row in connection.execute(
            """
            SELECT audit_status, COUNT(*)
            FROM stg_dimension_low_value_audit
            WHERE is_current=1
            GROUP BY audit_status
            ORDER BY audit_status
            """
        )
    }
    resolution_counts = {
        row[0]: row[1]
        for row in connection.execute(
            """
            SELECT resolution_status, COUNT(*)
            FROM fact_dimension_resolution_ledger
            GROUP BY resolution_status
            ORDER BY resolution_status
            """
        )
    }
    connection.close()

    summary = {
        "run_name": RUN_NAME,
        "audited_at": timestamp,
        "threshold_mm": LOW_DIMENSION_THRESHOLD_MM,
        "before_low_value_products": len(before_rows),
        "audit_status_counts": status_counts,
        "repaired_count": len(repaired),
        "repaired": repaired,
        "after_low_value_products": len(after_rows),
        "after_low_value_product_ids": [row["product_id"] for row in after_rows],
        "resolution_status_counts": resolution_counts,
        "policy": (
            "W/D/H 중 하나라도 20mm 이하이면 자동완료 금지. "
            "액자·러그·거울·옷걸이·벽선반 등 박형 제품은 "
            "명시적 제품 규격 근거가 있을 때만 예외 유지."
        ),
    }
    SUMMARY_PATH.write_text(
        json.dumps(summary, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
