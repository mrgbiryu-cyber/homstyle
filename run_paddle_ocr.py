from __future__ import annotations

import argparse
import json
import os
import sys
import time
from pathlib import Path

from PIL import Image


ROOT = Path(__file__).resolve().parent
RUN_DIR = ROOT / "poc_full_run"
MANIFEST_PATH = RUN_DIR / "image_manifest.json"
IMAGE_DIR = RUN_DIR / "images_flat"
OUTPUT_PATH = RUN_DIR / "ocr_paddle_ko.json"
MODEL_ROOT = ROOT / ".tools" / "paddle_models"
DET_MODEL_DIR = MODEL_ROOT / "PP-OCRv5_mobile_det_infer"
REC_MODEL_DIR = MODEL_ROOT / "korean_PP-OCRv5_mobile_rec_infer"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run local PaddleOCR PP-OCRv5 Korean OCR on the PoC images."
    )
    parser.add_argument("--limit", type=int, default=0, help="Process only N images.")
    parser.add_argument(
        "--resume",
        action="store_true",
        help="Reuse completed items already present in the output JSON.",
    )
    return parser.parse_args()


def save_output(metadata: dict, items: list[dict]) -> None:
    payload = {"metadata": metadata, "items": items}
    temp_path = OUTPUT_PATH.with_suffix(".json.tmp")
    temp_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    temp_path.replace(OUTPUT_PATH)


def main() -> int:
    args = parse_args()
    os.environ.setdefault("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True")
    os.environ.setdefault("FLAGS_use_mkldnn", "0")

    missing = [p for p in (MANIFEST_PATH, DET_MODEL_DIR, REC_MODEL_DIR) if not p.exists()]
    if missing:
        print("Missing required paths:", *(str(p) for p in missing), sep="\n- ", file=sys.stderr)
        return 2

    manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8-sig"))
    manifest = [row for row in manifest if row.get("download_status") == "SUCCESS"]
    if args.limit:
        manifest = manifest[: args.limit]

    existing: dict[str, dict] = {}
    if args.resume and OUTPUT_PATH.exists():
        old = json.loads(OUTPUT_PATH.read_text(encoding="utf-8-sig"))
        existing = {
            row["file"]: row
            for row in old.get("items", [])
            if row.get("status") in {"SUCCESS", "SKIPPED_TINY"}
        }

    print("Importing PaddleOCR...", flush=True)
    import paddle
    import paddleocr
    from paddleocr import PaddleOCR

    init_started = time.perf_counter()
    ocr = PaddleOCR(
        text_detection_model_name="PP-OCRv5_mobile_det",
        text_detection_model_dir=str(DET_MODEL_DIR),
        text_recognition_model_name="korean_PP-OCRv5_mobile_rec",
        text_recognition_model_dir=str(REC_MODEL_DIR),
        use_doc_orientation_classify=False,
        use_doc_unwarping=False,
        use_textline_orientation=False,
        text_recognition_batch_size=16,
        enable_mkldnn=False,
        device="cpu",
    )
    init_ms = round((time.perf_counter() - init_started) * 1000)

    metadata = {
        "engine": "PaddleOCR",
        "paddle_version": paddle.__version__,
        "paddleocr_version": paddleocr.__version__,
        "ocr_version": "PP-OCRv5",
        "detection_model": "PP-OCRv5_mobile_det",
        "recognition_model": "korean_PP-OCRv5_mobile_rec",
        "language": "korean",
        "device": "cpu",
        "mkldnn": False,
        "text_det_limit_side_len": 1280,
        "text_det_limit_type": "max",
        "text_rec_score_thresh": 0.5,
        "init_ms": init_ms,
        "started_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "input_count": len(manifest),
    }

    items: list[dict] = []
    run_started = time.perf_counter()
    for index, entry in enumerate(manifest, start=1):
        filename = entry["file"]
        image_path = IMAGE_DIR / filename
        if filename in existing:
            items.append(existing[filename])
            print(f"[{index:02d}/{len(manifest):02d}] {filename}: REUSED", flush=True)
            continue

        started = time.perf_counter()
        base = {
            "product_id": entry.get("product_id", ""),
            "file": filename,
            "role": entry.get("role", ""),
            "language": "korean",
        }
        try:
            with Image.open(image_path) as image:
                width, height = image.size
            base.update({"width": width, "height": height})
            if width <= 2 or height <= 2:
                row = {
                    **base,
                    "status": "SKIPPED_TINY",
                    "line_count": 0,
                    "character_count": 0,
                    "mean_confidence": None,
                    "text": "",
                    "lines": [],
                    "scores": [],
                    "error": "Image is 2x2 pixels or smaller.",
                }
            else:
                results = list(
                    ocr.predict(
                        str(image_path),
                        text_det_limit_side_len=1280,
                        text_det_limit_type="max",
                        text_rec_score_thresh=0.5,
                    )
                )
                payload = results[0].json.get("res", {}) if results else {}
                raw_lines = payload.get("rec_texts", []) or []
                raw_scores = payload.get("rec_scores", []) or []
                kept = [
                    (str(line).strip(), float(score))
                    for line, score in zip(raw_lines, raw_scores)
                    if str(line).strip() and float(score) >= 0.5
                ]
                lines = [line for line, _ in kept]
                scores = [round(score, 6) for _, score in kept]
                text = "\n".join(lines)
                row = {
                    **base,
                    "status": "SUCCESS",
                    "line_count": len(lines),
                    "character_count": len(text),
                    "mean_confidence": round(sum(scores) / len(scores), 6) if scores else None,
                    "text": text,
                    "lines": lines,
                    "scores": scores,
                    "error": "",
                }
        except Exception as exc:  # Keep the batch running and retain the exact failure.
            row = {
                **base,
                "status": "ERROR",
                "line_count": 0,
                "character_count": 0,
                "mean_confidence": None,
                "text": "",
                "lines": [],
                "scores": [],
                "error": f"{type(exc).__name__}: {exc}",
            }

        row["elapsed_ms"] = round((time.perf_counter() - started) * 1000)
        items.append(row)
        metadata["processed_count"] = len(items)
        metadata["elapsed_ms"] = round((time.perf_counter() - run_started) * 1000)
        save_output(metadata, items)
        print(
            f"[{index:02d}/{len(manifest):02d}] {filename}: {row['status']} "
            f"lines={row['line_count']} chars={row['character_count']} "
            f"elapsed={row['elapsed_ms']}ms",
            flush=True,
        )

    metadata.update(
        {
            "completed_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "processed_count": len(items),
            "success_count": sum(r["status"] == "SUCCESS" for r in items),
            "tiny_skip_count": sum(r["status"] == "SKIPPED_TINY" for r in items),
            "error_count": sum(r["status"] == "ERROR" for r in items),
            "nonempty_count": sum(bool(r.get("text")) for r in items),
            "elapsed_ms": round((time.perf_counter() - run_started) * 1000),
        }
    )
    save_output(metadata, items)
    print(json.dumps(metadata, ensure_ascii=False, indent=2), flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
