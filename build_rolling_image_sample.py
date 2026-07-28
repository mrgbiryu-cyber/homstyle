from __future__ import annotations

import argparse
import json
from pathlib import Path
from zipfile import ZipFile

import requests
import urllib3

import build_feasibility_sheet as xlsx_writer


ROOT = Path(__file__).resolve().parent
DEFAULT_PRODUCT_ID = "G25070005743"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 Chrome/150 Safari/537.36"
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build a one-product rolling image URL sample.")
    parser.add_argument("--product-id", default=DEFAULT_PRODUCT_ID)
    return parser.parse_args()


def absolute_image_url(value: str) -> str:
    value = value.strip()
    if value.startswith("/goods/"):
        return "https://static-store.lge.co.kr" + value
    if value.startswith("//"):
        return "https:" + value
    return value


def fetch_product(session: requests.Session, product_id: str) -> tuple[str, list[dict]]:
    endpoint = (
        "https://livingapi.lge.co.kr/itemsvc/ajax/v1/pdp/goods/"
        f"{product_id}?epFlagYn=N"
    )
    response = session.get(
        endpoint,
        timeout=30,
        headers={
            "Accept": "application/json",
            "Referer": f"https://homestyle.lge.co.kr/item?productId={product_id}",
        },
    )
    response.raise_for_status()
    payload = response.json()
    data = payload.get("data") or {}
    images = sorted(
        (row for row in data.get("images") or [] if isinstance(row, dict)),
        key=lambda row: int(row.get("sortSeq") or 9999),
    )
    normalized = []
    seen: set[str] = set()
    for image in images:
        if image.get("type") not in (None, "", "IMAGE"):
            continue
        url = absolute_image_url(str(image.get("imageUrl") or ""))
        if not url.startswith(("http://", "https://")) or url in seen:
            continue
        seen.add(url)
        normalized.append(
            {
                "sequence": int(image.get("sortSeq") or len(normalized) + 1),
                "url": url,
                "alt": str(image.get("imageAlt") or ""),
            }
        )
    if not normalized:
        raise ValueError(f"No rolling images returned for {product_id}")
    return str(data.get("productName") or product_id), normalized


def verify_images(session: requests.Session, images: list[dict]) -> None:
    for image in images:
        response = session.get(
            image["url"],
            timeout=30,
            stream=True,
            headers={"Referer": "https://homestyle.lge.co.kr/"},
        )
        image["http_status"] = response.status_code
        image["content_type"] = response.headers.get("Content-Type", "")
        response.close()


def validate(path: Path, expected_count: int) -> None:
    with ZipFile(path) as archive:
        if archive.testzip() is not None:
            raise ValueError("Corrupt XLSX archive")
        sheet = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
        if sheet.count("롤링 이미지 URL") != expected_count:
            raise ValueError("Rolling image column count mismatch")
        if sheet.count("https://static-store.lge.co.kr") < expected_count + 1:
            raise ValueError("Representative/rolling image URL count mismatch")


def main() -> None:
    args = parse_args()
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    session = requests.Session()
    # Workspace Python does not have the corporate proxy CA installed.
    session.verify = False
    session.headers.update({"User-Agent": USER_AGENT})

    product_name, images = fetch_product(session, args.product_id)
    verify_images(session, images)
    output = ROOT / f"롤링이미지_대표샘플_{args.product_id}.xlsx"

    headers = [
        "상품 ID",
        "상품명",
        "대표 이미지 URL",
        "롤링 이미지 수",
        *[f"롤링 이미지 URL {index:02d}" for index in range(1, len(images) + 1)],
    ]
    row = [
        args.product_id,
        product_name,
        images[0]["url"],
        len(images),
        *[image["url"] for image in images],
    ]
    groups = ["COMMON", "COMMON"] + ["R1"] * (len(headers) - 2)
    widths = [20, 54, 58, 16] + [58] * len(images)
    rows = [headers, row]

    xlsx_writer.OUTPUT = output
    xlsx_writer.SHEETS = [
        ("롤링이미지_한상품", rows, widths, groups, ["COMMON"]),
    ]
    xlsx_writer.build_xlsx()
    validate(output, len(images))

    result = {
        "output": str(output),
        "product_id": args.product_id,
        "product_name": product_name,
        "representative_image_url": images[0]["url"],
        "rolling_image_count": len(images),
        "all_images_http_200": all(image["http_status"] == 200 for image in images),
        "images": images,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
