from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path


ROOT = Path(__file__).resolve().parent
RUN_DIR = ROOT / "poc_full_run"
OUTPUT = RUN_DIR / "ocr_engine_benchmark.json"


PRODUCTS = {
    "G25070005743": "소파",
    "G25100020496": "러그",
    "G25070001871": "펜던트 조명",
    "G25070006112": "사진액자",
    "OLED48C6KNA": "TV",
    "G646GBB031": "냉장고",
    "WA2525YMZF": "워시타워",
    "SQ06GJ1WFS": "벽걸이 에어컨",
}

# Ground truth was transcribed from the dimension/detail images and cross-checked
# against the API/HTML values. Products without a reliable image-level numeric
# ground truth are intentionally excluded from numeric recall.
CRITICAL_TOKENS = {
    "G25070005743": ["2910", "1020", "730", "910", "740", "660", "410"],
    "G25100020496": ["140", "200", "160", "230", "300"],
    "OLED48C6KNA": ["1071", "620", "46.9", "120"],
    "G646GBB031": ["914", "911", "709", "698", "1860"],
    "WA2525YMZF": ["700", "1890", "830", "1410"],
    "SQ06GJ1WFS": ["837", "308", "189"],
}

CRITICAL_PHRASES = {
    "G25070005743": [
        "4인 소파+스툴",
        "카멜브라운",
        "크림 아이보리",
        "천연면피",
        "합성가죽",
    ],
    "G25100020496": ["가구 배치 추천"],
    "G25070001871": ["배송 기간"],
    "G25070006112": ["10X15CM", "LEATHER", "FLOWER"],
    "G646GBB031": ["도어쿨링"],
    "WA2525YMZF": ["문 열었을 때"],
    "SQ06GJ1WFS": ["아이스쿨파워"],
}

# Conservative list: images visually confirmed as having no useful field text.
# Decorative micro-text inside the photo-frame mockups is excluded from this list.
CONFIRMED_NO_FIELD_TEXT = {
    "G25070005743__01.jpg",
    "G25070005743__02.jpg",
    "G25100020496__01.jpg",
    "G25100020496__02.jpg",
    "G25100020496__05.jpg",
    "G25100020496__07.jpg",
    "G25070001871__01.png",
    "G25070001871__02.png",
    "G25070001871__03.png",
    "G25070001871__04.png",
    "G25070001871__05.png",
    "OLED48C6KNA__03.jpg",
    "OLED48C6KNA__04.jpg",
    "OLED48C6KNA__05.jpg",
    "OLED48C6KNA__07.jpg",
    "OLED48C6KNA__08.jpg",
    "G646GBB031__05.jpg",
    "G646GBB031__06.jpg",
    "G646GBB031__07.jpg",
    "WA2525YMZF__04.jpg",
    "WA2525YMZF__05.jpg",
    "WA2525YMZF__06.jpg",
    "WA2525YMZF__07.jpg",
    "WA2525YMZF__08.jpg",
    "SQ06GJ1WFS__01.jpg",
    "SQ06GJ1WFS__02.jpg",
    "SQ06GJ1WFS__07.jpg",
    "SQ06GJ1WFS__08.jpg",
}


def load_json(name: str):
    return json.loads((RUN_DIR / name).read_text(encoding="utf-8-sig"))


def product_id_from_file(filename: str) -> str:
    return filename.split("__", 1)[0]


def normalize_numeric_text(text: str) -> str:
    return text.replace(",", "")


def contains_exact_token(text: str, token: str) -> bool:
    text = normalize_numeric_text(text)
    escaped = re.escape(token)
    return bool(re.search(rf"(?<!\d){escaped}(?!\d)", text))


def normalize_phrase(text: str) -> str:
    return "".join(ch.lower() for ch in text if ch.isalnum())


def contains_phrase(text: str, phrase: str) -> bool:
    return normalize_phrase(phrase) in normalize_phrase(text)


def combine_windows_rows(ko_rows: list[dict], en_rows: list[dict]) -> tuple[list[dict], dict]:
    by_file: dict[str, dict[str, dict]] = defaultdict(dict)
    for language, rows in (("ko", ko_rows), ("en-US", en_rows)):
        for row in rows:
            by_file[row["file"]][language] = row

    combined = []
    for filename, language_rows in by_file.items():
        ko = language_rows.get("ko", {})
        en = language_rows.get("en-US", {})
        text = "\n".join(part for part in (ko.get("text", ""), en.get("text", "")) if part)
        statuses = [row.get("status") for row in language_rows.values()]
        combined.append(
            {
                "product_id": product_id_from_file(filename),
                "file": filename,
                "role": "",
                "status": "SUCCESS" if statuses and all(s == "SUCCESS" for s in statuses) else "ERROR",
                "line_count": len([line for line in text.splitlines() if line.strip()]),
                "character_count": len(text),
                "mean_confidence": None,
                "text": text,
                "elapsed_ms": sum(int(row.get("elapsed_ms", 0)) for row in language_rows.values()),
                "error": " | ".join(row.get("error", "") for row in language_rows.values() if row.get("error")),
            }
        )
    metadata = {
        "engine": "Windows OCR",
        "version": "Windows.Media.Ocr",
        "language": "ko + en-US (dual pass)",
        "input_count": len(combined),
        "attempt_count": len(ko_rows) + len(en_rows),
        "elapsed_ms": sum(int(row.get("elapsed_ms", 0)) for row in ko_rows + en_rows),
        "note": "Same 62 images were processed once in Korean and once in English; text was merged per image.",
    }
    return combined, metadata


def evaluate_engine(engine_name: str, rows: list[dict], metadata: dict) -> tuple[dict, list[dict], list[dict], list[dict]]:
    by_product: dict[str, list[dict]] = defaultdict(list)
    by_file = {}
    for row in rows:
        product_id = row.get("product_id") or product_id_from_file(row["file"])
        by_product[product_id].append(row)
        by_file[row["file"]] = row

    token_details = []
    for product_id, tokens in CRITICAL_TOKENS.items():
        text = "\n".join(row.get("text", "") for row in by_product[product_id])
        for token in tokens:
            found = contains_exact_token(text, token)
            token_details.append(
                {
                    "engine": engine_name,
                    "product_id": product_id,
                    "product_group": PRODUCTS[product_id],
                    "gold_type": "CRITICAL_NUMERIC_TOKEN",
                    "gold_value": token,
                    "found": found,
                }
            )

    phrase_details = []
    for product_id, phrases in CRITICAL_PHRASES.items():
        text = "\n".join(row.get("text", "") for row in by_product[product_id])
        for phrase in phrases:
            found = contains_phrase(text, phrase)
            phrase_details.append(
                {
                    "engine": engine_name,
                    "product_id": product_id,
                    "product_group": PRODUCTS[product_id],
                    "gold_type": "CRITICAL_FIELD_PHRASE",
                    "gold_value": phrase,
                    "found": found,
                }
            )

    image_details = []
    for filename in sorted(by_file):
        row = by_file[filename]
        text = row.get("text", "") or ""
        confirmed_blank = filename in CONFIRMED_NO_FIELD_TEXT
        image_details.append(
            {
                "engine": engine_name,
                "product_id": row.get("product_id") or product_id_from_file(filename),
                "product_group": PRODUCTS.get(row.get("product_id") or product_id_from_file(filename), ""),
                "file": filename,
                "role": row.get("role", ""),
                "status": row.get("status", ""),
                "line_count": int(row.get("line_count", 0)),
                "character_count": int(row.get("character_count", len(text))),
                "mean_confidence": row.get("mean_confidence"),
                "elapsed_ms": int(row.get("elapsed_ms", 0)),
                "confirmed_no_field_text": confirmed_blank,
                "false_positive_on_confirmed_blank": bool(confirmed_blank and text.strip()),
                "text_excerpt": text[:1000],
                "error": row.get("error", ""),
            }
        )

    token_hits = sum(row["found"] for row in token_details)
    phrase_hits = sum(row["found"] for row in phrase_details)
    false_positives = sum(row["false_positive_on_confirmed_blank"] for row in image_details)
    elapsed_ms = int(metadata.get("elapsed_ms") or sum(row["elapsed_ms"] for row in image_details))
    summary = {
        "engine": engine_name,
        "version": metadata.get("version", ""),
        "language": metadata.get("language", ""),
        "processed_images": len(image_details),
        "successful_images": sum(row["status"] in {"SUCCESS", "SKIPPED_TINY"} for row in image_details),
        "error_images": sum(row["status"] == "ERROR" for row in image_details),
        "nonempty_images": sum(row["character_count"] > 0 for row in image_details),
        "critical_token_hits": token_hits,
        "critical_token_total": len(token_details),
        "critical_token_recall_pct": round(token_hits / len(token_details) * 100, 1),
        "critical_phrase_hits": phrase_hits,
        "critical_phrase_total": len(phrase_details),
        "critical_phrase_recall_pct": round(phrase_hits / len(phrase_details) * 100, 1),
        "confirmed_blank_images": len(CONFIRMED_NO_FIELD_TEXT),
        "false_positive_images": false_positives,
        "false_positive_rate_pct": round(false_positives / len(CONFIRMED_NO_FIELD_TEXT) * 100, 1),
        "elapsed_ms": elapsed_ms,
        "average_ms_per_image": round(elapsed_ms / len(image_details), 1),
        "total_characters": sum(row["character_count"] for row in image_details),
        "configuration_note": metadata.get("note") or metadata.get("engine_note", ""),
    }
    return summary, token_details, phrase_details, image_details


def main() -> None:
    windows_rows, windows_meta = combine_windows_rows(load_json("ocr_ko.json"), load_json("ocr_en.json"))
    paddle = load_json("ocr_paddle_ko.json")
    tesseract = load_json("ocr_tesseract_js.json")

    engines = [
        ("Windows OCR (ko+en)", windows_rows, windows_meta),
        (
            "PaddleOCR PP-OCRv5",
            paddle["items"],
            {
                **paddle["metadata"],
                "version": f"PaddleOCR {paddle['metadata']['paddleocr_version']} / Paddle {paddle['metadata']['paddle_version']}",
                "note": "PP-OCRv5 mobile detector + Korean mobile recognizer; CPU; MKLDNN disabled for runtime compatibility.",
            },
        ),
        (
            "Tesseract.js",
            tesseract["items"],
            {
                **tesseract["metadata"],
                "version": f"Tesseract.js {tesseract['metadata']['tesseract_js_version']}",
            },
        ),
    ]

    summaries = []
    token_details = []
    phrase_details = []
    image_details = []
    for engine_name, rows, metadata in engines:
        summary, token_rows, phrase_rows, image_rows = evaluate_engine(engine_name, rows, metadata)
        summaries.append(summary)
        token_details.extend(token_rows)
        phrase_details.extend(phrase_rows)
        image_details.extend(image_rows)

    ranked = sorted(
        summaries,
        key=lambda row: (
            row["critical_token_recall_pct"],
            row["critical_phrase_recall_pct"],
            -row["false_positive_rate_pct"],
            -row["average_ms_per_image"],
        ),
        reverse=True,
    )
    for rank, row in enumerate(ranked, start=1):
        row["quality_rank"] = rank

    payload = {
        "metadata": {
            "benchmark_name": "Homestyle OCR engine comparison",
            "image_count": 62,
            "product_count": 8,
            "numeric_gold_product_count": len(CRITICAL_TOKENS),
            "critical_numeric_token_count": sum(map(len, CRITICAL_TOKENS.values())),
            "critical_field_phrase_count": sum(map(len, CRITICAL_PHRASES.values())),
            "confirmed_no_field_text_image_count": len(CONFIRMED_NO_FIELD_TEXT),
            "metric_note": "This is exact critical-token/phrase recall, not full-text CER. Full CER requires a manually transcribed corpus.",
            "native_tesseract_note": "Native Tesseract 5.5 installer, tesserocr, conda and WSL paths were attempted but blocked by UAC/compiler/repository constraints. Tesseract.js 7.0.0 (WebAssembly Tesseract port) was executed instead and is labeled separately.",
        },
        "recommendation": {
            "primary_ocr": ranked[0]["engine"],
            "reason": "Highest exact critical numeric-token recall, then field-phrase recall; false positives and speed are secondary controls.",
            "production_policy": "API/HTML first; OCR only on detail/spec images; accept numeric OCR only when it matches API/HTML or passes a dimension rule/manual review.",
        },
        "summary": summaries,
        "token_details": token_details,
        "phrase_details": phrase_details,
        "image_details": image_details,
    }
    OUTPUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"output": str(OUTPUT), "summary": summaries}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
