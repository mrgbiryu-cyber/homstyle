from __future__ import annotations

import json
import sqlite3
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

from bulk_homestyle_collect import pack, unpack


ROOT = Path(__file__).resolve().parent
DB_PATH = ROOT / "homestyle_bulk_run" / "homestyle_bulk.sqlite"
RUN_NAME = "locked_dimension_quality_repair_20260728"
SUMMARY_PATH = (
    ROOT
    / "homestyle_bulk_run"
    / "locked_dimension_quality_repair_20260728_summary.json"
)


@dataclass(frozen=True)
class Option:
    label: str
    w_mm: float
    d_mm: float | None
    h_mm: float
    primary: bool = False


@dataclass(frozen=True)
class Repair:
    product_id: str
    w_mm: float
    d_mm: float | None
    h_mm: float
    raw_notation: str
    reason: str
    evidence_text: str
    shape_type: str = "RECTANGLE"
    d_applicability: str = "APPLICABLE"
    options: tuple[Option, ...] = field(default_factory=tuple)


REPAIRS = (
    Repair(
        "G25080008914",
        565,
        570,
        750,
        "Open size W565 × H750 × D570 mm",
        "OCR의 5/S·7/? 혼동과 W-H-D 축 순서 오해를 원본 이미지로 교정",
        "제품명 스프링 쿠션 체어 | 제품 사용 상태(Open size) W=565 mm D=570 mm H=750 mm",
        options=(
            Option("사용 상태(Open)", 565, 570, 750, True),
            Option("접은 상태(Folding)", 565, 100, 565),
        ),
    ),
    Repair(
        "G25080010611",
        1400,
        760,
        750,
        "W1400 × D760 × H750 mm",
        "OCR이 W1400을 WI격00으로 읽어 W=0으로 저장한 값을 원본 규격 이미지로 교정",
        "제품명 노바레트로 테이블 | 제품 사이즈 W=1400 mm D=760 mm H=750 mm",
    ),
    Repair(
        "G25080010622",
        1400,
        400,
        750,
        "W1400 × D400 × H750 mm",
        "OCR이 W1400을 WI격00으로 읽어 W=0으로 저장한 값을 원본 규격 이미지로 교정",
        "제품명 노바레트로 9칸 와이드 서랍장 B | 제품 사이즈 W=1400 mm D=400 mm H=750 mm",
    ),
    Repair(
        "G25110024675",
        600,
        619.5,
        1860,
        "가로 600 × 세로 619.5 × 높이 1860 mm",
        "라인업 앞쪽 숫자 2를 W로 잘못 선택한 값을 제품명 600폭과 MODEL SIZE 문맥으로 교정",
        "제품명 프리스토 2단 600폭 렌지장 | 제품 사이즈 W=600 mm D=619.5 mm H=1860 mm",
    ),
    Repair(
        "G26010030824",
        180,
        180,
        395,
        "Ø180 × H395 mm",
        "전기 사양 숫자를 규격으로 결합한 값을 원본 INFORMATION 도면의 원형 외곽 규격으로 교정",
        "제품명 스피아노 아델리오 LED 크롬 테이블 단스탠드 | 원형 받침 지름 180 mm, 전체 높이 395 mm | W=180 mm D=180 mm H=395 mm",
        shape_type="ROUND",
    ),
    Repair(
        "G26010030838",
        220,
        220,
        776,
        "Ø220 × H776 mm",
        "전기 사양 숫자를 규격으로 결합한 값을 원본 INFORMATION 도면의 원형 외곽 규격으로 교정",
        "제품명 스피아노 아델리오 LED 크롬 플로어 장스탠드 | 원형 받침 지름 220 mm, 전체 높이 776 mm | W=220 mm D=220 mm H=776 mm",
        shape_type="ROUND",
    ),
    Repair(
        "G26040038373",
        1000,
        None,
        1500,
        "100 × 150 cm",
        "OCR이 100을 1 00/00으로 분리해 W=0으로 저장한 값을 원본 SIZE 표기로 교정",
        "제품명 잭 러그 | 제품 사이즈 100 × 150 cm | 2D 규격 W=1000 mm H=1500 mm, D=해당없음",
        shape_type="AREA_2D",
        d_applicability="NOT_APPLICABLE",
    ),
    Repair(
        "G26050042633",
        1800,
        800,
        900,
        "가로 1800 × 세로 800 × 높이 620~900 mm",
        "라인업 숫자 9를 W로 잘못 선택한 값을 상품명 1800폭과 MODEL SIZE 문맥으로 교정",
        "제품명 업모션 1800폭 타원형 테이블(LPM) | 제품 사이즈 W=1800 mm D=800 mm H=620~900 mm | 대표 배치 높이는 최대 900 mm",
        shape_type="OVAL",
        options=(
            Option("높이 최대", 1800, 800, 900, True),
            Option("높이 최소", 1800, 800, 620),
        ),
    ),
    Repair(
        "G26050042634",
        1600,
        800,
        898,
        "가로 1600 × 세로 800 × 높이 618~898 mm",
        "라인업 숫자 1을 W로 잘못 선택한 값을 상품명 1600폭과 MODEL SIZE 문맥으로 교정",
        "제품명 업모션 테이블 1600폭 타원형 (강화 사틴유리) | 제품 사이즈 W=1600 mm D=800 mm H=618~898 mm | 대표 배치 높이는 최대 898 mm",
        shape_type="OVAL",
        options=(
            Option("높이 최대", 1600, 800, 898, True),
            Option("높이 최소", 1600, 800, 618),
        ),
    ),
    Repair(
        "G26050042635",
        1800,
        800,
        898,
        "가로 1800 × 세로 800 × 높이 618~898 mm",
        "라인업 숫자 1을 W로 잘못 선택한 값을 상품명 1800폭과 MODEL SIZE 문맥으로 교정",
        "제품명 업모션 테이블 1800폭 타원형 (강화 사틴유리) | 제품 사이즈 W=1800 mm D=800 mm H=618~898 mm | 대표 배치 높이는 최대 898 mm",
        shape_type="OVAL",
        options=(
            Option("높이 최대", 1800, 800, 898, True),
            Option("높이 최소", 1800, 800, 618),
        ),
    ),
    Repair(
        "G26050042636",
        1600,
        800,
        900,
        "가로 1600 × 세로 800 × 높이 620~900 mm",
        "라인업 숫자 9를 W로 잘못 선택한 값을 상품명 1600폭과 MODEL SIZE 문맥으로 교정",
        "제품명 업모션 사각 테이블 1600폭(LPM) | 제품 사이즈 W=1600 mm D=800 mm H=620~900 mm | 대표 배치 높이는 최대 900 mm",
        options=(
            Option("높이 최대", 1600, 800, 900, True),
            Option("높이 최소", 1600, 800, 620),
        ),
    ),
    Repair(
        "G26050042637",
        1600,
        800,
        898,
        "가로 1600 × 세로 800 × 높이 618~898 mm",
        "라인업 숫자 9를 W로 잘못 선택한 값을 상품명 1600폭과 MODEL SIZE 문맥으로 교정",
        "제품명 업모션 사각 테이블 1600폭 (강화 사틴유리) | 제품 사이즈 W=1600 mm D=800 mm H=618~898 mm | 대표 배치 높이는 최대 898 mm",
        options=(
            Option("높이 최대", 1600, 800, 898, True),
            Option("높이 최소", 1600, 800, 618),
        ),
    ),
    Repair(
        "G26060045757",
        1200,
        718,
        1705,
        "1200 × 718~891 × 1705 mm",
        "중간축 범위 718~891을 포함한 3축 표기에서 W를 0으로 잘못 저장한 값을 교정",
        "제품명 로이모노 1200폭 5단 책상세트 | 제품 사이즈 W=1200 mm D=718~891 mm H=1705 mm | 단일 D 필드는 수납 상태 718 mm",
        options=(
            Option("깊이 최소(책상 수납)", 1200, 718, 1705, True),
            Option("깊이 최대(책상 인출)", 1200, 891, 1705),
        ),
    ),
    Repair(
        "G26060045758",
        1200,
        718,
        2038,
        "1200 × 718~891 × 2038 mm",
        "중간축 범위 718~891을 포함한 3축 표기에서 W를 0으로 잘못 저장한 값을 교정",
        "제품명 로이모노 1200폭 6단 책상세트 | 제품 사이즈 W=1200 mm D=718~891 mm H=2038 mm | 단일 D 필드는 수납 상태 718 mm",
        options=(
            Option("깊이 최소(책상 수납)", 1200, 718, 2038, True),
            Option("깊이 최대(책상 인출)", 1200, 891, 2038),
        ),
    ),
    Repair(
        "G26060045759",
        1447,
        718,
        1705,
        "1447 × 718~891 × 1705 mm",
        "중간축 범위 718~891을 포함한 3축 표기에서 W를 0으로 잘못 저장한 값을 교정",
        "제품명 로이모노 1400폭 5단 책상세트 | MODEL SIZE 제품 외곽 W=1447 mm D=718~891 mm H=1705 mm | 단일 D 필드는 수납 상태 718 mm",
        options=(
            Option("깊이 최소(책상 수납)", 1447, 718, 1705, True),
            Option("깊이 최대(책상 인출)", 1447, 891, 1705),
        ),
    ),
    Repair(
        "G26060045760",
        1447,
        718,
        2038,
        "1447 × 718~891 × 2038 mm",
        "중간축 범위 718~891을 포함한 3축 표기에서 W를 0으로 잘못 저장한 값을 교정",
        "제품명 로이모노 1400폭 6단 책상세트 | MODEL SIZE 제품 외곽 W=1447 mm D=718~891 mm H=2038 mm | 단일 D 필드는 수납 상태 718 mm",
        options=(
            Option("깊이 최소(책상 수납)", 1447, 718, 2038, True),
            Option("깊이 최대(책상 인출)", 1447, 891, 2038),
        ),
    ),
)


def now_text() -> str:
    return datetime.now(ZoneInfo("Asia/Seoul")).isoformat(timespec="seconds")


def init_audit_table(connection: sqlite3.Connection) -> None:
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


def source_url(ocr_payload: dict[str, Any]) -> str:
    return str((ocr_payload.get("selected") or {}).get("url") or "")


def add_reinforcement(
    connection: sqlite3.Connection,
    repair: Repair,
    timestamp: str,
) -> str:
    source = connection.execute(
        "SELECT ocr_status, ocr_blob FROM sources WHERE product_id=?",
        (repair.product_id,),
    ).fetchone()
    if source is None:
        raise RuntimeError(f"sources row missing: {repair.product_id}")
    ocr_status, previous_blob = source
    payload = unpack(previous_blob) or {}
    url = source_url(payload)
    connection.execute(
        """
        INSERT OR IGNORE INTO source_ocr_dimension_backup (
            run_name, product_id, backed_up_at,
            previous_ocr_status, previous_ocr_blob
        ) VALUES (?,?,?,?,?)
        """,
        (RUN_NAME, repair.product_id, timestamp, ocr_status, previous_blob),
    )
    reinforcements = [
        item
        for item in (payload.get("dimension_reinforcements") or [])
        if not (
            item.get("run_name") == RUN_NAME
            and item.get("product_id") == repair.product_id
        )
    ]
    item: dict[str, Any] = {
        "run_name": RUN_NAME,
        "product_id": repair.product_id,
        "verification_method": "stored_detail_image_visual_and_raw_context",
        "verified_at": timestamp,
        "source_url": url,
        "raw_notation": repair.raw_notation,
        "evidence_text": repair.evidence_text,
        "repair_reason": repair.reason,
        "w_mm": repair.w_mm,
        "d_mm": repair.d_mm,
        "h_mm": repair.h_mm,
    }
    if repair.d_mm is not None:
        item["dimensions"] = [[repair.w_mm, repair.d_mm, repair.h_mm]]
    if len(repair.options) > 1:
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
    connection.execute(
        "UPDATE sources SET ocr_blob=?, ocr_at=? WHERE product_id=?",
        (pack(payload), timestamp, repair.product_id),
    )
    return url


def update_classification_master(
    connection: sqlite3.Connection,
    repair: Repair,
) -> None:
    if repair.d_mm is None:
        return
    connection.execute(
        """
        UPDATE stg_dimension_classification_master
        SET classification_complete=1,
            classification_status='COMPLETE',
            classification_stage='IMAGE_RAW_VERIFIED_REPAIR',
            classification_code='B02_IMAGE_RAW_VERIFIED',
            classification_name='원본 이미지·OCR 문맥 수기 검증 교정',
            dimension_resolution_status='MANUAL_CONFIRMED_WDH',
            dimension_value_confirmed=1,
            confirmed_w_mm=?,
            confirmed_d_mm=?,
            confirmed_h_mm=?,
            quality_flag='LOCKED_VALUE_REPAIRED_20260728',
            requires_followup=0,
            followup_priority=NULL,
            next_action='완료',
            evidence_text=?
        WHERE product_id=? AND is_current=1
        """,
        (
            repair.w_mm,
            repair.d_mm,
            repair.h_mm,
            repair.evidence_text,
            repair.product_id,
        ),
    )


def update_area_context(
    connection: sqlite3.Connection,
    repair: Repair,
) -> None:
    if repair.shape_type != "AREA_2D":
        return
    candidate = connection.execute(
        """
        SELECT candidate_key, snapshot_id
        FROM stg_dimension_context_candidate
        WHERE product_id=? AND is_current=1
        ORDER BY candidate_score DESC
        LIMIT 1
        """,
        (repair.product_id,),
    ).fetchone()
    if candidate is None:
        return
    candidate_key, snapshot_id = candidate
    connection.execute(
        """
        UPDATE stg_dimension_context_candidate
        SET raw_notation=?,
            normalized_axis_mapping='2D_PAIR->W,H;D=N/A',
            unit_status='UNIT_PRESENT',
            unit_text='cm',
            value_1_raw=100,
            value_2_raw=150,
            w_mm=?,
            d_mm=NULL,
            h_mm=?,
            decision_status='CATEGORY_NORMALIZED',
            rejection_reason=''
        WHERE candidate_key=?
        """,
        (repair.raw_notation, repair.w_mm, repair.h_mm, candidate_key),
    )
    connection.execute(
        """
        UPDATE stg_dimension_context_product
        SET context_status='AUTO_ACCEPT_CANDIDATE',
            representative_candidate_key=?,
            representative_w_mm=?,
            representative_d_mm=NULL,
            representative_h_mm=?,
            representative_shape_type='AREA_2D',
            representative_d_applicability='NOT_APPLICABLE',
            requires_reocr=0,
            requires_human_review=0,
            next_action='완료'
        WHERE product_id=? AND snapshot_id=? AND is_current=1
        """,
        (
            candidate_key,
            repair.w_mm,
            repair.h_mm,
            repair.product_id,
            snapshot_id,
        ),
    )


def replace_options(
    connection: sqlite3.Connection,
    repair: Repair,
    timestamp: str,
) -> None:
    options = repair.options or (
        Option("대표", repair.w_mm, repair.d_mm, repair.h_mm, True),
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
                repair.d_applicability,
                repair.shape_type,
                option.w_mm if repair.shape_type == "ROUND" else None,
                (
                    "W,H->W,H;D=N/A"
                    if repair.shape_type == "AREA_2D"
                    else "MANUAL_VERIFIED->W,D,H"
                ),
                "MANUAL_IMAGE_RAW_REPAIR",
                "",
                "OCR_IMAGE_MANUAL_VERIFICATION",
                f"{repair.evidence_text} | 원문: {repair.raw_notation}",
                timestamp,
                "dimension-resolution-ledger-v1",
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
    url = add_reinforcement(connection, repair, timestamp)
    update_classification_master(connection, repair)
    update_area_context(connection, repair)
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
            locked_d_applicability=?,
            locked_shape_type=?,
            locked_option_count=?,
            candidate_w_mm=NULL,
            candidate_d_mm=NULL,
            candidate_h_mm=NULL,
            context_status='MANUAL_IMAGE_RAW_VERIFIED',
            resolution_rule_code='MANUAL_IMAGE_RAW_REPAIR',
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
            repair.d_applicability,
            repair.shape_type,
            len(repair.options) or 1,
            RUN_NAME,
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
            url,
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
            "INVALID_LOCKED_VALUE_IMAGE_RAW_REPAIR",
            RUN_NAME,
            old["locked_w_mm"],
            old["locked_d_mm"],
            old["locked_h_mm"],
            repair.w_mm,
            repair.d_mm,
            repair.h_mm,
            "dimension-resolution-ledger-v1",
        ),
    )
    return {
        "product_id": repair.product_id,
        "old_status": old["resolution_status"],
        "old_w_mm": old["locked_w_mm"],
        "old_d_mm": old["locked_d_mm"],
        "old_h_mm": old["locked_h_mm"],
        "new_status": "MANUAL_CONFIRMED",
        "new_w_mm": repair.w_mm,
        "new_d_mm": repair.d_mm,
        "new_h_mm": repair.h_mm,
        "raw_notation": repair.raw_notation,
        "evidence_url": url,
    }


def main() -> None:
    timestamp = now_text()
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA journal_mode=WAL")
    connection.execute("PRAGMA synchronous=NORMAL")
    init_audit_table(connection)
    connection.commit()
    repaired: list[dict[str, Any]] = []
    try:
        connection.execute("BEGIN IMMEDIATE")
        for repair in REPAIRS:
            repaired.append(repair_one(connection, repair, timestamp))
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()

    summary = {
        "run_name": RUN_NAME,
        "repaired_at": timestamp,
        "database": str(DB_PATH),
        "repaired_count": len(repaired),
        "repaired": repaired,
        "policy": {
            "nonpositive_dimension": "완료 금지",
            "range": "단일값은 첫 끝점, option 테이블에는 최소/최대 모두 보존",
            "manual_verification": "MANUAL_CONFIRMED / OCR·규칙확정",
        },
    }
    SUMMARY_PATH.write_text(
        json.dumps(summary, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
