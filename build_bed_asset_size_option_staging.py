from __future__ import annotations

import json
import re
import sqlite3
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any

from bulk_homestyle_collect import DB_PATH, unpack
from build_homestyle_bulk_workbook import option_groups
from build_product_component_staging import (
    SET_RE,
    is_bed_frame_mattress_single_asset,
)


RULE_VERSION = "bed-single-3d-asset-size-option-v1.0"

SIZE_PATTERNS: list[tuple[str, re.Pattern[str]]] = [
    ("KK", re.compile(r"킹오브킹|\bKK\b", re.I)),
    ("CK", re.compile(r"칼킹|\bCK\b", re.I)),
    ("RK", re.compile(r"레귤러킹|\bRK\b", re.I)),
    ("LK", re.compile(r"라지킹|\bLK\b", re.I)),
    ("SS", re.compile(r"슈퍼싱글|\bSS\b", re.I)),
    ("DS", re.compile(r"더블싱글|\bDS\b", re.I)),
    ("DD", re.compile(r"더블|\bDD\b", re.I)),
    ("Q", re.compile(r"퀸|\bQ\b", re.I)),
    ("KI", re.compile(r"\bKI\b", re.I)),
    ("K", re.compile(r"(?<![A-Z0-9])K(?![A-Z0-9])|(?<!오브)킹", re.I)),
]


def now_iso() -> str:
    return datetime.now().astimezone().replace(microsecond=0).isoformat()


def bed_size_codes(text: str) -> list[str]:
    found: list[tuple[int, str]] = []
    occupied: list[tuple[int, int]] = []
    for code, pattern in SIZE_PATTERNS:
        for match in pattern.finditer(text or ""):
            span = match.span()
            if any(span[0] < end and start < span[1] for start, end in occupied):
                continue
            occupied.append(span)
            found.append((span[0], code))
    return list(dict.fromkeys(code for _, code in sorted(found)))


def create_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS stg_bed_asset_policy (
            snapshot_id TEXT NOT NULL,
            built_at TEXT NOT NULL,
            is_current INTEGER NOT NULL,
            rule_version TEXT NOT NULL,
            product_id TEXT NOT NULL,
            product_name TEXT NOT NULL,
            asset_policy TEXT NOT NULL,
            option_group_count INTEGER NOT NULL,
            size_option_item_count INTEGER NOT NULL,
            distinct_bed_size_count INTEGER NOT NULL,
            bed_size_codes TEXT,
            requires_size_option_split INTEGER NOT NULL,
            option_evidence TEXT,
            PRIMARY KEY (snapshot_id, product_id)
        );

        CREATE TABLE IF NOT EXISTS stg_bed_asset_size_option (
            snapshot_id TEXT NOT NULL,
            built_at TEXT NOT NULL,
            is_current INTEGER NOT NULL,
            rule_version TEXT NOT NULL,
            product_id TEXT NOT NULL,
            option_group_seq INTEGER NOT NULL,
            option_seq INTEGER NOT NULL,
            option_style TEXT,
            raw_option_style TEXT,
            option_name TEXT NOT NULL,
            option_id TEXT,
            bed_size_codes TEXT,
            is_bed_size_option INTEGER NOT NULL,
            evidence_text TEXT,
            PRIMARY KEY (
                snapshot_id,
                product_id,
                option_group_seq,
                option_seq
            )
        );

        CREATE INDEX IF NOT EXISTS idx_bed_asset_policy_current
            ON stg_bed_asset_policy(is_current, requires_size_option_split);
        CREATE INDEX IF NOT EXISTS idx_bed_asset_size_option_current
            ON stg_bed_asset_size_option(
                is_current,
                product_id,
                is_bed_size_option
            );

        DROP VIEW IF EXISTS vw_bed_asset_policy_current;
        CREATE VIEW vw_bed_asset_policy_current AS
        SELECT *
        FROM stg_bed_asset_policy
        WHERE is_current = 1;

        DROP VIEW IF EXISTS vw_bed_asset_size_options_current;
        CREATE VIEW vw_bed_asset_size_options_current AS
        SELECT *
        FROM stg_bed_asset_size_option
        WHERE is_current = 1;
        """
    )


def build(db_path: Path = DB_PATH) -> dict[str, Any]:
    connection = sqlite3.connect(db_path)
    connection.row_factory = sqlite3.Row
    create_schema(connection)
    rows = connection.execute(
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

    targets = [
        row
        for row in rows
        if is_bed_frame_mattress_single_asset(
            row["product_name"] or "",
            row["mid_category"] or "",
            row["small_category"] or "",
        )
        and not SET_RE.search(row["product_name"] or "")
    ]
    if len(targets) != 121:
        raise RuntimeError(f"Expected 121 reviewed bed products, found {len(targets)}")

    built_at = now_iso()
    snapshot_id = (
        "bed_asset_policy_"
        + datetime.now().astimezone().strftime("%Y%m%d_%H%M%S")
    )
    connection.execute(
        "UPDATE stg_bed_asset_policy SET is_current = 0 WHERE is_current = 1"
    )
    connection.execute(
        "UPDATE stg_bed_asset_size_option SET is_current = 0 WHERE is_current = 1"
    )

    policy_rows: list[tuple[Any, ...]] = []
    option_rows: list[tuple[Any, ...]] = []
    split_examples: list[dict[str, Any]] = []
    split_counts: Counter[str] = Counter()

    for row in targets:
        product_id = row["product_id"]
        product_name = row["product_name"] or product_id
        payload = unpack(row["goods_blob"]) or {}
        data = payload.get("data") or {}
        groups = option_groups(data, "")

        all_size_codes: list[str] = []
        size_option_item_count = 0
        option_evidence: list[str] = []
        for group_seq, group in enumerate(groups, start=1):
            group_style = group["style"]
            raw_style = group["raw_style"]
            group_codes: list[str] = []
            parsed_items: list[tuple[int, dict[str, Any], list[str]]] = []
            for option_seq, item in enumerate(group["items"], start=1):
                codes = bed_size_codes(item["name"])
                group_codes.extend(codes)
                parsed_items.append((option_seq, item, codes))

            unique_group_codes = list(dict.fromkeys(group_codes))
            compact_style = re.sub(r"\s+", "", raw_style).casefold()
            explicit_bed_size_group = compact_style in {
                "사이즈",
                "size",
                "침대사이즈",
                "프레임사이즈",
            }
            item_mentions_bed = any(
                re.search(r"침대|프레임", item["name"], flags=re.I)
                for _, item, _ in parsed_items
            )
            size_group = (
                explicit_bed_size_group
                or (item_mentions_bed and len(unique_group_codes) >= 2)
            )

            for option_seq, item, codes in parsed_items:
                is_size_option = int(size_group and bool(codes))
                if is_size_option:
                    size_option_item_count += 1
                    all_size_codes.extend(codes)
                    option_evidence.append(
                        f"{raw_style or group_style}:{item['name']}"
                    )
                option_rows.append(
                    (
                        snapshot_id,
                        built_at,
                        1,
                        RULE_VERSION,
                        product_id,
                        group_seq,
                        option_seq,
                        group_style,
                        raw_style,
                        item["name"],
                        item["option_id"],
                        ",".join(codes),
                        is_size_option,
                        group["evidence"],
                    )
                )

        distinct_codes = list(dict.fromkeys(all_size_codes))
        requires_split = int(len(distinct_codes) >= 2)
        split_counts["SIZE_OPTION_SPLIT" if requires_split else "SINGLE_ASSET"] += 1
        if requires_split and len(split_examples) < 10:
            split_examples.append(
                {
                    "product_id": product_id,
                    "product_name": product_name,
                    "bed_size_codes": distinct_codes,
                    "option_evidence": option_evidence,
                }
            )
        policy_rows.append(
            (
                snapshot_id,
                built_at,
                1,
                RULE_VERSION,
                product_id,
                product_name,
                "SINGLE_3D_ASSET",
                len(groups),
                size_option_item_count,
                len(distinct_codes),
                ",".join(distinct_codes),
                requires_split,
                " | ".join(option_evidence),
            )
        )

    connection.executemany(
        """
        INSERT INTO stg_bed_asset_policy VALUES (
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
        )
        """,
        policy_rows,
    )
    connection.executemany(
        """
        INSERT INTO stg_bed_asset_size_option VALUES (
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
        )
        """,
        option_rows,
    )
    connection.commit()

    current_products = connection.execute(
        "SELECT COUNT(*) FROM vw_bed_asset_policy_current"
    ).fetchone()[0]
    current_options = connection.execute(
        "SELECT COUNT(*) FROM vw_bed_asset_size_options_current"
    ).fetchone()[0]
    if current_products != 121:
        raise RuntimeError(
            f"Bed asset policy current count mismatch: {current_products}"
        )
    connection.close()

    return {
        "database": str(db_path),
        "snapshot_id": snapshot_id,
        "rule_version": RULE_VERSION,
        "target_products": len(targets),
        "policy_counts": dict(sorted(split_counts.items())),
        "option_rows": current_options,
        "split_examples": split_examples,
    }


if __name__ == "__main__":
    print(json.dumps(build(), ensure_ascii=False, indent=2))
