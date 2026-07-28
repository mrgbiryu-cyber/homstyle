#!/usr/bin/env python
"""Markdown 문서를 읽기 좋은 독립형 HTML로 변환한다."""

from __future__ import annotations

import argparse
from collections import OrderedDict
import html
import os
import re
import sys
import webbrowser
from datetime import datetime
from pathlib import Path
from urllib.parse import quote, unquote, urlsplit, urlunsplit

try:
    import markdown
    from markdown.extensions.toc import slugify_unicode
except ImportError:
    print(
        "필수 패키지 'Markdown'이 없습니다.\n"
        "설치 방법:\n"
        "  python -m pip install -r requirements-md-html.txt\n"
        "또는 이 프로젝트 환경에서는:\n"
        "  uv --system-certs pip install "
        "--python .venv-ocr\\Scripts\\python.exe -r requirements-md-html.txt",
        file=sys.stderr,
    )
    raise SystemExit(2)

try:
    from pygments.formatters import HtmlFormatter
except ImportError:  # 코드 하이라이트만 단색으로 표시하고 변환은 계속한다.
    HtmlFormatter = None


MARKDOWN_SUFFIXES = {".md", ".markdown"}
LINK_ATTRIBUTE_RE = re.compile(
    r'(?P<prefix>\b(?:href|src)\s*=\s*")(?P<url>[^"]+)(?P<suffix>")',
    re.IGNORECASE,
)
EXTERNAL_SCHEME_RE = re.compile(r"^[A-Za-z][A-Za-z0-9+.-]*:")
FIRST_H1_RE = re.compile(r"^\s*#\s+(.+?)\s*#*\s*$", re.MULTILINE)
TABLE_SEPARATOR_RE = re.compile(r"^:?-{3,}:?$")


BASE_CSS = """
:root {
  color-scheme: light;
  --bg: #f4f6f8;
  --paper: #ffffff;
  --text: #20242a;
  --muted: #68717d;
  --line: #dfe3e8;
  --accent: #a50034;
  --accent-soft: #fff0f4;
  --code-bg: #f7f8fa;
  --shadow: 0 8px 30px rgba(24, 32, 44, 0.08);
}

* { box-sizing: border-box; }
html { scroll-behavior: smooth; }
body {
  margin: 0;
  color: var(--text);
  background: var(--bg);
  font-family: Pretendard, "Noto Sans KR", "Malgun Gothic", Arial, sans-serif;
  font-size: 16px;
  line-height: 1.72;
  overflow-wrap: anywhere;
}
a { color: #075ea8; text-decoration: none; }
a:hover { text-decoration: underline; }

.page-header {
  color: white;
  background: linear-gradient(115deg, #85002a, #b8003c 65%, #d53569);
  padding: 26px max(24px, calc((100vw - 1480px) / 2));
  box-shadow: 0 2px 12px rgba(78, 0, 25, 0.2);
}
.page-header h1 {
  max-width: 1100px;
  margin: 0;
  font-size: clamp(22px, 2.5vw, 34px);
  line-height: 1.3;
  letter-spacing: -0.025em;
}
.page-header .source {
  margin-top: 8px;
  color: rgba(255, 255, 255, 0.78);
  font-size: 13px;
}

.layout {
  display: grid;
  grid-template-columns: minmax(220px, 280px) minmax(0, 1120px);
  gap: 24px;
  max-width: 1480px;
  margin: 24px auto 48px;
  padding: 0 24px;
  align-items: start;
}
.layout.no-toc {
  display: block;
  max-width: 1168px;
}
.toc-panel {
  position: sticky;
  top: 18px;
  max-height: calc(100vh - 36px);
  overflow: auto;
  padding: 18px;
  background: var(--paper);
  border: 1px solid var(--line);
  border-radius: 12px;
  box-shadow: var(--shadow);
  font-size: 13px;
}
.toc-title {
  margin: 0 0 10px;
  color: var(--accent);
  font-size: 14px;
  font-weight: 700;
}
.toc-panel ul {
  margin: 0;
  padding-left: 18px;
  list-style: none;
}
.toc-panel > .toc > ul { padding-left: 0; }
.toc-panel li { margin: 5px 0; }
.toc-panel a { color: #3d4651; }
.toc-panel a:hover { color: var(--accent); }

.document {
  min-width: 0;
  padding: clamp(24px, 4vw, 54px);
  background: var(--paper);
  border: 1px solid var(--line);
  border-radius: 14px;
  box-shadow: var(--shadow);
}
.document > :first-child { margin-top: 0; }
.document > :last-child { margin-bottom: 0; }

h1, h2, h3, h4, h5, h6 {
  scroll-margin-top: 18px;
  line-height: 1.35;
  letter-spacing: -0.02em;
}
.document h1 {
  margin: 0 0 30px;
  padding-bottom: 16px;
  border-bottom: 3px solid var(--accent);
  font-size: 32px;
}
.document h2 {
  margin-top: 48px;
  padding-bottom: 10px;
  border-bottom: 1px solid var(--line);
  font-size: 25px;
}
.document h3 { margin-top: 34px; font-size: 20px; }
.document h4 { margin-top: 28px; font-size: 17px; }
.headerlink {
  margin-left: 7px;
  color: #aeb5bd;
  font-size: 0.75em;
  opacity: 0;
}
h1:hover .headerlink, h2:hover .headerlink, h3:hover .headerlink,
h4:hover .headerlink { opacity: 1; }

p, ul, ol, blockquote, pre, table { margin-top: 0; margin-bottom: 18px; }
ul, ol { padding-left: 24px; }
li + li { margin-top: 4px; }
hr {
  margin: 36px 0;
  border: 0;
  border-top: 1px solid var(--line);
}
blockquote {
  margin-left: 0;
  padding: 12px 18px;
  color: #4d5661;
  background: #faf7f8;
  border-left: 4px solid var(--accent);
}

.document table {
  display: block;
  width: max-content;
  max-width: 100%;
  overflow-x: auto;
  border-collapse: collapse;
  border-spacing: 0;
  font-size: 14px;
}
.document th, .document td {
  min-width: 90px;
  padding: 9px 12px;
  border: 1px solid var(--line);
  text-align: left;
  vertical-align: top;
}
.document th {
  position: sticky;
  top: 0;
  color: #2d333b;
  background: #f0f2f5;
  font-weight: 700;
}
.document tbody tr:nth-child(even) { background: #fbfcfd; }

code, pre, kbd {
  font-family: "Cascadia Code", Consolas, "Noto Sans Mono CJK KR", monospace;
}
code {
  padding: 2px 5px;
  color: #8f1538;
  background: var(--code-bg);
  border-radius: 4px;
  font-size: 0.9em;
}
pre {
  max-width: 100%;
  overflow: auto;
  padding: 16px 18px;
  background: #f6f8fa;
  border: 1px solid var(--line);
  border-radius: 8px;
  line-height: 1.55;
}
pre code { padding: 0; color: inherit; background: transparent; }

.admonition, .note {
  padding: 14px 18px;
  background: #f4f8ff;
  border: 1px solid #d5e4fa;
  border-radius: 8px;
}
.footnote { margin-top: 42px; font-size: 14px; color: var(--muted); }

.page-footer {
  max-width: 1480px;
  margin: -24px auto 36px;
  padding: 0 24px;
  color: var(--muted);
  font-size: 12px;
  text-align: right;
}
.back-to-top {
  display: inline-block;
  margin-top: 12px;
  padding: 6px 10px;
  color: var(--muted);
  border: 1px solid var(--line);
  border-radius: 999px;
}

.pattern-accordion-summary {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  margin: 20px 0 12px;
  padding: 13px 16px;
  color: #4a2632;
  background: var(--accent-soft);
  border: 1px solid #f0c8d4;
  border-radius: 9px;
  font-size: 14px;
  font-weight: 650;
}
.pattern-accordion {
  margin: 10px 0;
  overflow: clip;
  background: var(--paper);
  border: 1px solid var(--line);
  border-radius: 10px;
  box-shadow: 0 2px 8px rgba(24, 32, 44, 0.04);
}
.pattern-accordion[open] {
  border-color: #d8a2b3;
  box-shadow: 0 5px 18px rgba(83, 18, 39, 0.09);
}
.pattern-accordion > summary {
  display: flex;
  align-items: center;
  gap: 12px;
  min-height: 56px;
  padding: 12px 16px;
  cursor: pointer;
  user-select: none;
  list-style: none;
  background: #fbfcfd;
  transition: background 0.15s ease;
}
.pattern-accordion > summary:hover { background: #fff6f8; }
.pattern-accordion > summary:focus-visible {
  outline: 3px solid rgba(165, 0, 52, 0.22);
  outline-offset: -3px;
}
.pattern-accordion > summary::-webkit-details-marker { display: none; }
.pattern-accordion > summary::before {
  content: "▶";
  flex: 0 0 auto;
  color: var(--accent);
  font-size: 12px;
  transition: transform 0.18s ease;
}
.pattern-accordion[open] > summary::before { transform: rotate(90deg); }
.pattern-heading {
  display: flex;
  flex: 1 1 440px;
  flex-direction: column;
  min-width: 360px;
  gap: 3px;
}
.pattern-name {
  min-width: 360px;
  color: #2e3339;
  font-weight: 700;
}
.pattern-example {
  color: var(--muted);
  font-size: 12px;
  font-weight: 500;
}
.pattern-code {
  flex: 0 1 auto;
  color: var(--muted);
  font-size: 11px;
  white-space: nowrap;
}
.pattern-code code { color: #6c5360; }
.pattern-count {
  flex: 0 0 auto;
  padding: 3px 9px;
  color: #5b2637;
  background: var(--accent-soft);
  border: 1px solid #f1d0da;
  border-radius: 999px;
  font-size: 12px;
  font-weight: 700;
}
.pattern-final {
  min-width: 0;
  color: var(--muted);
  font-size: 12px;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
.detail-label {
  flex: 0 0 auto;
  margin-left: auto;
  color: var(--accent);
  font-size: 12px;
  font-weight: 700;
}
.detail-label::after { content: "상세 보기"; }
.pattern-accordion[open] .detail-label::after { content: "접기"; }
.accordion-body {
  padding: 16px;
  border-top: 1px solid var(--line);
  background: white;
}
.accordion-body table { margin-bottom: 0; }

@media (max-width: 900px) {
  .layout { display: block; padding: 0 12px; margin-top: 12px; }
  .toc-panel {
    position: static;
    max-height: 320px;
    margin-bottom: 12px;
  }
  .document { padding: 24px 18px; border-radius: 10px; }
  .document h1 { font-size: 27px; }
  .document h2 { font-size: 22px; }
  .page-header { padding: 20px 18px; }
  .pattern-accordion > summary { flex-wrap: wrap; gap: 7px 10px; }
  .pattern-heading { min-width: 0; width: calc(100% - 28px); }
  .pattern-name { min-width: 0; }
  .pattern-code { order: 4; padding-left: 23px; }
  .pattern-final { order: 5; width: 100%; padding-left: 23px; }
  .detail-label { margin-left: auto; }
}

@media print {
  body { background: white; font-size: 11pt; }
  .page-header {
    padding: 0 0 14px;
    color: black;
    background: none;
    box-shadow: none;
    border-bottom: 2px solid #333;
  }
  .page-header .source { color: #555; }
  .layout, .layout.no-toc { display: block; max-width: none; margin: 0; padding: 0; }
  .toc-panel { display: none; }
  .document { padding: 0; border: 0; box-shadow: none; }
  .page-footer { display: none; }
  .document table { display: table; width: 100%; font-size: 8.5pt; }
  .pattern-accordion { overflow: visible; break-inside: avoid; }
  .pattern-accordion > summary::before, .detail-label { display: none; }
  .pattern-accordion > .accordion-body {
    display: block !important;
    padding: 8px 0 20px;
    border-top: 1px solid #bbb;
  }
  a { color: inherit; text-decoration: none; }
  h1, h2, h3, h4 { break-after: avoid; }
  pre, blockquote, table { break-inside: avoid; }
}
"""


def clean_title(raw_title: str) -> str:
    """첫 H1의 간단한 Markdown 표기를 제거해 문서 제목으로 사용한다."""
    title = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", raw_title)
    title = re.sub(r"[*_`~]", "", title)
    return html.unescape(title).strip()


def extract_title(markdown_text: str, fallback: str) -> str:
    match = FIRST_H1_RE.search(markdown_text)
    return clean_title(match.group(1)) if match else fallback


def split_markdown_table_row(line: str) -> list[str]:
    """이스케이프된 \|는 셀 구분자로 보지 않고 Markdown 표 한 행을 나눈다."""
    stripped = line.strip()
    if stripped.startswith("|"):
        stripped = stripped[1:]
    if stripped.endswith("|"):
        stripped = stripped[:-1]

    cells: list[str] = []
    current: list[str] = []
    backslash_count = 0
    for char in stripped:
        if char == "|" and backslash_count % 2 == 0:
            cells.append("".join(current).strip())
            current = []
            backslash_count = 0
            continue
        current.append(char)
        if char == "\\":
            backslash_count += 1
        else:
            backslash_count = 0
    cells.append("".join(current).strip())
    return cells


def plain_markdown_cell(cell: str) -> str:
    text = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", cell)
    text = re.sub(r"<br\s*/?>", " / ", text, flags=re.IGNORECASE)
    text = text.replace(r"\|", "|")
    # 패턴 코드의 밑줄은 의미가 있으므로 Markdown 기호 제거 대상에서 제외한다.
    text = re.sub(r"[*`~]", "", text)
    return html.unescape(re.sub(r"\s+", " ", text)).strip()


def render_markdown_cell(cell: str) -> str:
    """표 셀 안의 링크·코드·줄바꿈 같은 인라인 Markdown을 HTML로 바꾼다."""
    rendered = markdown.markdown(
        cell.strip(),
        extensions=["markdown.extensions.extra"],
        output_format="html5",
    ).strip()
    if rendered.startswith("<p>") and rendered.endswith("</p>"):
        rendered = rendered[3:-4]
    return rendered


def is_table_separator(cells: list[str]) -> bool:
    return bool(cells) and all(
        bool(TABLE_SEPARATOR_RE.fullmatch(cell.replace(" ", ""))) for cell in cells
    )


def build_pattern_accordion(
    headers: list[str],
    rows: list[list[str]],
    group_index: int,
    final_index: int | None,
    name_index: int | None,
    example_index: int | None,
) -> list[str]:
    groups: OrderedDict[str, list[list[str]]] = OrderedDict()
    for row in rows:
        padded = row + [""] * max(0, len(headers) - len(row))
        key = plain_markdown_cell(padded[group_index]) or "미분류"
        groups.setdefault(key, []).append(padded[: len(headers)])

    output = [
        (
            '<div class="pattern-accordion-summary">'
            f"<span>동일 패턴 기준으로 {len(groups):,}개 그룹을 묶었습니다.</span>"
            f"<span>총 {len(rows):,}개 상품</span>"
            "</div>"
        )
    ]
    header_html = "".join(
        f"<th>{render_markdown_cell(header)}</th>" for header in headers
    )

    for key, group_rows in groups.items():
        display_names: list[str] = []
        if name_index is not None:
            for row in group_rows:
                value = plain_markdown_cell(row[name_index])
                if value and value not in display_names:
                    display_names.append(value)
        display_name = display_names[0] if display_names else key

        examples: list[str] = []
        if example_index is not None:
            for row in group_rows:
                value = plain_markdown_cell(row[example_index])
                if value and value not in examples:
                    examples.append(value)
        example = examples[0] if examples else ""

        final_values: list[str] = []
        if final_index is not None:
            for row in group_rows:
                value = plain_markdown_cell(row[final_index])
                if value and value not in final_values:
                    final_values.append(value)
        final_text = " · ".join(final_values[:3])
        if len(final_values) > 3:
            final_text += f" 외 {len(final_values) - 3}종"

        body_rows = []
        for row in group_rows:
            body_rows.append(
                "<tr>"
                + "".join(
                    f"<td>{render_markdown_cell(cell)}</td>" for cell in row
                )
                + "</tr>"
            )

        final_label = (
            f'<span class="pattern-final">최종 코드: {html.escape(final_text)}</span>'
            if final_text
            else ""
        )
        output.extend(
            [
                '<details class="pattern-accordion">',
                "<summary>",
                '<span class="pattern-heading">',
                f'<span class="pattern-name">{html.escape(display_name)}</span>',
                (
                    f'<span class="pattern-example">예: '
                    f"{html.escape(example)}</span>"
                    if example
                    else ""
                ),
                "</span>",
                f'<span class="pattern-code"><code>{html.escape(key)}</code></span>',
                f'<span class="pattern-count">{len(group_rows):,}개 상품</span>',
                final_label,
                '<span class="detail-label" aria-hidden="true"></span>',
                "</summary>",
                '<div class="accordion-body">',
                "<table>",
                f"<thead><tr>{header_html}</tr></thead>",
                f"<tbody>{''.join(body_rows)}</tbody>",
                "</table>",
                "</div>",
                "</details>",
            ]
        )
    return output


def group_repeated_pattern_tables(markdown_text: str) -> str:
    """상품 ID와 패턴 코드가 있는 반복 표를 패턴별 details 아코디언으로 바꾼다."""
    lines = markdown_text.splitlines()
    output: list[str] = []
    index = 0

    while index < len(lines):
        if (
            index + 2 >= len(lines)
            or not lines[index].lstrip().startswith("|")
            or not lines[index + 1].lstrip().startswith("|")
        ):
            output.append(lines[index])
            index += 1
            continue

        headers = split_markdown_table_row(lines[index])
        separators = split_markdown_table_row(lines[index + 1])
        if len(headers) != len(separators) or not is_table_separator(separators):
            output.append(lines[index])
            index += 1
            continue

        end = index + 2
        raw_rows: list[list[str]] = []
        while end < len(lines) and lines[end].lstrip().startswith("|"):
            raw_rows.append(split_markdown_table_row(lines[end]))
            end += 1

        normalized_headers = [plain_markdown_cell(cell) for cell in headers]
        product_index = next(
            (
                i
                for i, name in enumerate(normalized_headers)
                if name in {"상품 ID", "상품ID", "제품 ID", "제품ID"}
            ),
            None,
        )
        group_index = next(
            (
                i
                for preferred in (
                    "규격 패턴 코드",
                    "기존 패턴 코드",
                    "과거 코드",
                    "패턴 코드",
                    "패턴",
                    "상품 구성 패턴 코드",
                    "최종 코드",
                )
                for i, name in enumerate(normalized_headers)
                if name == preferred
            ),
            None,
        )
        final_index = next(
            (
                i
                for i, name in enumerate(normalized_headers)
                if name in {"상품 구성 패턴 코드", "최종 코드"}
            ),
            None,
        )
        name_index = next(
            (
                i
                for i, name in enumerate(normalized_headers)
                if name in {"결정자용 패턴명", "패턴명", "상품 구성 패턴명"}
            ),
            None,
        )
        example_index = next(
            (
                i
                for i, name in enumerate(normalized_headers)
                if name in {"패턴 예시", "구성 예시", "대표 예시", "예시"}
            ),
            None,
        )

        if product_index is None or group_index is None or len(raw_rows) < 8:
            output.extend(lines[index:end])
            index = end
            continue

        keys = [
            plain_markdown_cell(
                (row + [""] * max(0, len(headers) - len(row)))[group_index]
            )
            for row in raw_rows
        ]
        if len(set(keys)) == len(keys):
            output.extend(lines[index:end])
            index = end
            continue

        output.extend(
            build_pattern_accordion(
                headers,
                raw_rows,
                group_index,
                final_index,
                name_index,
                example_index,
            )
        )
        index = end

    trailing_newline = "\n" if markdown_text.endswith("\n") else ""
    return "\n".join(output) + trailing_newline


def is_external_url(url: str) -> bool:
    return (
        not url
        or url.startswith(("#", "/", "//", "data:"))
        or bool(EXTERNAL_SCHEME_RE.match(url))
    )


def encode_relative_url(path: str, query: str, fragment: str) -> str:
    encoded_path = quote(path.replace(os.sep, "/"), safe="/:@%+~()[],'=-_.")
    return urlunsplit(("", "", encoded_path, query, fragment))


def rewrite_relative_links(
    rendered_html: str,
    source_file: Path,
    output_file: Path,
    output_map: dict[Path, Path],
) -> str:
    """출력 위치가 달라져도 Markdown 링크와 이미지 상대 경로가 유지되게 한다."""

    def replace(match: re.Match[str]) -> str:
        escaped_url = match.group("url")
        url = html.unescape(escaped_url)
        if is_external_url(url):
            return match.group(0)

        split = urlsplit(url)
        decoded_path = unquote(split.path).replace("/", os.sep)
        if not decoded_path:
            return match.group(0)

        source_target = (source_file.parent / decoded_path).resolve()
        if source_target.suffix.lower() in MARKDOWN_SUFFIXES:
            # 일괄 변환 대상이면 실제 출력 위치를 사용한다. 단일 변환에서는
            # 현재 HTML을 기준으로 원래 상대 구조를 유지한 .html 링크를 만든다.
            target = output_map.get(source_target)
            if target is None:
                linked_output = Path(decoded_path).with_suffix(".html")
                target = (output_file.parent / linked_output).resolve()
        else:
            target = source_target

        relative = os.path.relpath(target, output_file.parent)
        new_url = encode_relative_url(relative, split.query, split.fragment)
        return (
            match.group("prefix")
            + html.escape(new_url, quote=True)
            + match.group("suffix")
        )

    return LINK_ATTRIBUTE_RE.sub(replace, rendered_html)


def pygments_css() -> str:
    if HtmlFormatter is None:
        return ""
    return HtmlFormatter(style="friendly").get_style_defs(".codehilite")


def render_document(
    source_file: Path,
    output_file: Path,
    output_map: dict[Path, Path],
    *,
    forced_title: str | None = None,
    include_toc: bool = True,
    group_patterns: bool = True,
) -> tuple[str, str]:
    markdown_text = source_file.read_text(encoding="utf-8-sig")
    title = forced_title or extract_title(markdown_text, source_file.stem)
    if group_patterns:
        markdown_text = group_repeated_pattern_tables(markdown_text)

    converter = markdown.Markdown(
        extensions=[
            "markdown.extensions.extra",
            "markdown.extensions.sane_lists",
            "markdown.extensions.codehilite",
            "markdown.extensions.toc",
        ],
        extension_configs={
            "markdown.extensions.codehilite": {
                "css_class": "codehilite",
                "guess_lang": False,
                "linenums": False,
            },
            "markdown.extensions.toc": {
                "permalink": "¶",
                "slugify": slugify_unicode,
                "toc_depth": "2-4",
            },
        },
        output_format="html5",
    )
    body = converter.convert(markdown_text)
    body = rewrite_relative_links(body, source_file, output_file, output_map)
    toc = converter.toc if include_toc and converter.toc else ""

    toc_panel = ""
    layout_class = "layout no-toc"
    if toc:
        toc_panel = (
            '<aside class="toc-panel" aria-label="문서 목차">'
            '<p class="toc-title">문서 목차</p>'
            f"{toc}"
            '<a class="back-to-top" href="#top">맨 위로</a>'
            "</aside>"
        )
        layout_class = "layout"

    source_label = source_file.name
    generated_at = datetime.now().astimezone().strftime("%Y-%m-%d %H:%M %Z")
    document = f"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="generator" content="md_to_html.py">
  <title>{html.escape(title)}</title>
  <style>
{BASE_CSS}
{pygments_css()}
  </style>
</head>
<body id="top">
  <header class="page-header">
    <h1>{html.escape(title)}</h1>
    <div class="source">원본: {html.escape(source_label)}</div>
  </header>
  <div class="{layout_class}">
    {toc_panel}
    <main class="document">
{body}
    </main>
  </div>
  <footer class="page-footer">
    {html.escape(generated_at)} 변환 · md_to_html.py
  </footer>
</body>
</html>
"""
    return document, title


def render_index(entries: list[tuple[str, Path]], index_file: Path) -> str:
    items = []
    for title, output_file in entries:
        relative = os.path.relpath(output_file, index_file.parent)
        href = encode_relative_url(relative, "", "")
        items.append(
            "<li>"
            f'<a href="{html.escape(href, quote=True)}">{html.escape(title)}</a>'
            f"<small>{html.escape(output_file.stem)}</small>"
            "</li>"
        )

    generated_at = datetime.now().astimezone().strftime("%Y-%m-%d %H:%M %Z")
    return f"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="generator" content="md_to_html.py">
  <title>Markdown 문서 목록</title>
  <style>
{BASE_CSS}
  .index-list {{ margin: 0; padding: 0; list-style: none; }}
  .index-list li {{
    display: flex;
    align-items: baseline;
    justify-content: space-between;
    gap: 18px;
    padding: 12px 4px;
    border-bottom: 1px solid var(--line);
  }}
  .index-list li:last-child {{ border-bottom: 0; }}
  .index-list a {{ font-weight: 650; }}
  .index-list small {{ color: var(--muted); text-align: right; }}
  </style>
</head>
<body id="top">
  <header class="page-header">
    <h1>Markdown 문서 목록</h1>
    <div class="source">총 {len(entries):,}개 문서</div>
  </header>
  <div class="layout no-toc">
    <main class="document">
      <ul class="index-list">
        {''.join(items)}
      </ul>
    </main>
  </div>
  <footer class="page-footer">
    {html.escape(generated_at)} 변환 · md_to_html.py
  </footer>
</body>
</html>
"""


def collect_sources(input_dir: Path, recursive: bool) -> list[Path]:
    iterator = input_dir.rglob("*") if recursive else input_dir.glob("*")
    return sorted(
        (
            path.resolve()
            for path in iterator
            if path.is_file() and path.suffix.lower() in MARKDOWN_SUFFIXES
        ),
        key=lambda path: str(path).casefold(),
    )


def resolve_single_output(source: Path, output_argument: Path | None) -> Path:
    if output_argument is None:
        return source.with_suffix(".html")

    output = output_argument.expanduser().resolve()
    if output.exists() and output.is_dir():
        return output / source.with_suffix(".html").name
    if output.suffix.lower() != ".html":
        return output / source.with_suffix(".html").name
    return output


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Markdown 파일 또는 폴더를 목차와 스타일이 포함된 독립형 HTML로 변환합니다.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""사용 예:
  python md_to_html.py 문서.md
  python md_to_html.py 문서.md -o 결과.html --open
  python md_to_html.py . -o html_output
  python md_to_html.py docs -o site --recursive --no-index
""",
    )
    parser.add_argument("input", type=Path, help="입력 Markdown 파일 또는 폴더")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="단일 파일의 출력 HTML 또는 폴더 변환 시 출력 폴더",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="폴더 입력 시 하위 폴더까지 검색",
    )
    parser.add_argument("--no-toc", action="store_true", help="문서 목차를 만들지 않음")
    parser.add_argument(
        "--no-index",
        action="store_true",
        help="폴더 변환 시 _index.html을 만들지 않음",
    )
    parser.add_argument(
        "--no-accordion",
        action="store_true",
        help="같은 패턴의 상품 표를 아코디언으로 묶지 않음",
    )
    parser.add_argument(
        "--title",
        help="단일 파일의 HTML 제목 강제 지정(생략 시 첫 H1 또는 파일명)",
    )
    parser.add_argument(
        "--open",
        action="store_true",
        help="변환 후 결과를 기본 브라우저로 열기",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    input_path = args.input.expanduser().resolve()

    if not input_path.exists():
        parser.error(f"입력 경로가 없습니다: {input_path}")

    if input_path.is_file():
        if input_path.suffix.lower() not in MARKDOWN_SUFFIXES:
            parser.error("입력 파일 확장자는 .md 또는 .markdown이어야 합니다.")

        output_file = resolve_single_output(input_path, args.output)
        output_map = {input_path: output_file}
        document, _ = render_document(
            input_path,
            output_file,
            output_map,
            forced_title=args.title,
            include_toc=not args.no_toc,
            group_patterns=not args.no_accordion,
        )
        output_file.parent.mkdir(parents=True, exist_ok=True)
        output_file.write_text(document, encoding="utf-8", newline="\n")
        print(f"[완료] 1개 문서 변환: {output_file}")
        if args.open:
            webbrowser.open(output_file.as_uri())
        return 0

    if args.title:
        parser.error("--title은 단일 파일 변환에서만 사용할 수 있습니다.")

    sources = collect_sources(input_path, args.recursive)
    if not sources:
        parser.error(f"변환할 Markdown 파일이 없습니다: {input_path}")

    output_root = (
        args.output.expanduser().resolve()
        if args.output
        else (input_path / "html_output").resolve()
    )
    output_map = {
        source: output_root
        / source.relative_to(input_path).with_suffix(".html")
        for source in sources
    }

    entries: list[tuple[str, Path]] = []
    for source in sources:
        output_file = output_map[source]
        document, title = render_document(
            source,
            output_file,
            output_map,
            include_toc=not args.no_toc,
            group_patterns=not args.no_accordion,
        )
        output_file.parent.mkdir(parents=True, exist_ok=True)
        output_file.write_text(document, encoding="utf-8", newline="\n")
        entries.append((title, output_file))

    open_target = entries[0][1]
    if not args.no_index:
        index_file = output_root / "_index.html"
        index_file.parent.mkdir(parents=True, exist_ok=True)
        index_file.write_text(
            render_index(entries, index_file),
            encoding="utf-8",
            newline="\n",
        )
        open_target = index_file
        print(f"[목록] {index_file}")

    print(f"[완료] {len(entries):,}개 문서 변환: {output_root}")
    if args.open:
        webbrowser.open(open_target.as_uri())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
