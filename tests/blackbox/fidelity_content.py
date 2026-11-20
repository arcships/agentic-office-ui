from __future__ import annotations

import json
import os
import re
from pathlib import Path
import signal
import socket
import subprocess
import sys
import time
import urllib.request

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

from browser_evidence import BrowserEvidence


ROOT = Path(__file__).resolve().parents[2]
OUTPUT = Path(os.environ.get(
    "BLACKBOX_EVIDENCE_DIR",
    ROOT / "output" / "acceptance" / "fidelity-content",
)).resolve()
PORT = int(os.environ.get("CI_PREVIEW_PORT", "4173"))
BASE_URL = os.environ.get("BLACKBOX_BASE_URL", f"http://127.0.0.1:{PORT}")
MANAGE_PREVIEW = "BLACKBOX_BASE_URL" not in os.environ


def port_in_use() -> bool:
    with socket.socket() as sock:
        return sock.connect_ex(("127.0.0.1", PORT)) == 0


def wait_for_preview(process: subprocess.Popen[bytes]) -> None:
    deadline = time.monotonic() + 30
    while time.monotonic() < deadline:
        if process.poll() is not None:
            raise RuntimeError(f"preview exited early with {process.returncode}")
        try:
            with urllib.request.urlopen(BASE_URL, timeout=1) as response:
                if response.status == 200:
                    return
        except Exception:
            time.sleep(0.2)
    raise TimeoutError(f"preview did not become ready at {BASE_URL}")


def stop_process(process: subprocess.Popen[bytes] | None) -> None:
    if process is None or process.poll() is not None:
        return
    try:
        os.killpg(process.pid, signal.SIGTERM)
        process.wait(timeout=5)
    except Exception:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except Exception:
            process.kill()


def launch_browser(playwright):
    failures: list[str] = []
    try:
        return playwright.chromium.launch(headless=True), "bundled-chromium"
    except PlaywrightError as error:
        failures.append(str(error))
    for candidate in (
        Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"),
        Path("/usr/bin/google-chrome"),
        Path("/usr/bin/google-chrome-stable"),
    ):
        if not candidate.exists():
            continue
        try:
            return (
                playwright.chromium.launch(headless=True, executable_path=str(candidate)),
                str(candidate),
            )
        except PlaywrightError as error:
            failures.append(str(error))
    raise RuntimeError("no Chromium executable could launch: " + " | ".join(failures))


def expect(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


# The header table of `table-structures.docx`: a running head laid out as a table.
HEADER_TABLE_ROWS = [
    ["Register", "Period", "Owner"],
    ["Table Register", "FY2026", "O. North"],
]


def active_docx_search_highlight(page, message: str) -> dict[str, float]:
    highlight = page.locator(
        ".docx-document-surface__search-highlight--active"
    ).first
    highlight.wait_for(state="visible", timeout=10_000)
    box = highlight.bounding_box()
    expect(
        box is not None and box["width"] > 1 and box["height"] > 1,
        message,
    )
    return box


def collect_docx_pages(page) -> list[dict]:
    """Per-page header/footer/width/watermark and header-table geometry."""
    return page.evaluate(
        """() => {
          const pages = [...document.querySelectorAll('[data-docx-page-wrapper="true"]')];
          return pages.map((wrapper) => {
            const surface = wrapper.querySelector('[data-docx-page-surface]');
            const header = wrapper.querySelector('[data-docx-page-header]');
            const footer = wrapper.querySelector('[data-docx-page-footer]');
            const body = wrapper.querySelector('[data-docx-page-body]');
            const headerTables = header
              ? [...header.querySelectorAll('table')].map((table) => ({
                  rows: [...table.rows].map((row) =>
                    [...row.cells].map((cell) => (cell.innerText || '').trim())
                  ),
                }))
              : [];
            const watermarks = header
              ? [...header.querySelectorAll('img')].filter((img) =>
                  /z-index:\\s*-\\d+/.test(img.getAttribute('style') || '')
                ).length
              : 0;
            return {
              index: Number(wrapper.getAttribute('data-docx-page-index')),
              width: surface
                ? Math.round(surface.getBoundingClientRect().width)
                : 0,
              header: header ? header.innerText.trim() : '',
              footer: footer ? footer.innerText.trim() : '',
              watermarks,
              headerTables,
              bodyText: body ? body.innerText : '',
            };
          });
        }"""
    )


def scroll_docx_viewer(page, fraction: float) -> None:
    page.evaluate(
        """(fraction) => {
          const viewer = document.querySelector('[data-testid="docx-viewer"]');
          const candidates = [...(viewer ? viewer.querySelectorAll('*') : [])]
            .filter((element) => element.scrollHeight > element.clientHeight + 200);
          candidates.sort((left, right) => right.scrollHeight - left.scrollHeight);
          const target = candidates[0];
          if (target) {
            target.scrollTop =
              (target.scrollHeight - target.clientHeight) * fraction;
          }
        }""",
        fraction,
    )
    page.wait_for_timeout(400)


def collect_all_docx_pages(page, expected_pages: int) -> list[dict]:
    """Merges page snapshots while stepping through the scroll range.

    The viewer only mounts the pages inside its scroll viewport, so a single
    snapshot cannot see every page of a multi-page document.
    """
    observed: dict[int, dict] = {}

    def collect() -> None:
        for entry in collect_docx_pages(page):
            observed.setdefault(entry["index"], entry)

    collect()
    for step in range(13):
        if len(observed) >= expected_pages:
            break
        scroll_docx_viewer(page, step / 12)
        collect()
    return [observed[index] for index in sorted(observed)]


def run_docx(browser, attempt_dir: Path) -> dict[str, object]:
    context = browser.new_context(viewport={"width": 1440, "height": 900}, locale="zh-CN")
    page = context.new_page()
    evidence = BrowserEvidence(page)
    try:
        page.goto(f"{BASE_URL}/#/docx-viewer", wait_until="networkidle")
        page.locator('[data-testid="docx-sample-select"]').select_option("review-comments.docx")
        page.locator('[data-testid="docx-load-sample"]').click()
        page.locator('[data-testid="page-status"][data-state="ready"]').wait_for(timeout=30_000)
        page.locator('[data-testid="docx-footnotes"]').wait_for(timeout=10_000)
        page.locator('[data-testid="docx-endnotes"]').wait_for(timeout=10_000)
        footnote_text = page.locator('[data-testid="docx-footnotes"]').inner_text()
        endnote_text = page.locator('[data-testid="docx-endnotes"]').inner_text()
        expect("服务级别摘要" in footnote_text, "脚注正文未展示")
        expect("响应时间变更" in endnote_text, "尾注正文未展示")
        search = page.locator('[data-testid="docx-search-input"]')
        search.fill("reviewed")
        page.locator('[data-testid="docx-search-status"]').wait_for()
        status = page.locator('[data-testid="docx-search-status"]').inner_text()
        expect(status not in {"", "0/0"}, f"全文搜索没有结果: {status}")
        search_highlight = active_docx_search_highlight(page, "搜索结果没有精确高亮")
        page.screenshot(path=str(attempt_dir / "docx-notes-search.png"), full_page=True)

        page.locator('[data-testid="docx-sample-select"]').select_option("invoice-table.docx")
        page.locator('[data-testid="docx-load-sample"]').click()
        table_paragraph = page.locator(
            '[data-docx-table-cell-paragraph-host="true"]',
            has_text="OCR document page processing",
        )
        table_paragraph.wait_for(timeout=30_000)
        search.fill("OCR document page processing")
        page.wait_for_function(
            """() => document.querySelector('[data-testid="docx-search-status"]')?.textContent !== '0/0'""",
            timeout=10_000,
        )
        table_search_status = page.locator('[data-testid="docx-search-status"]').inner_text()
        table_highlight = active_docx_search_highlight(page, "表格内搜索结果没有精确高亮")
        table_box = table_paragraph.bounding_box()
        expect(
            table_box is not None
            and table_highlight["x"] < table_box["x"] + table_box["width"]
            and table_highlight["x"] + table_highlight["width"] > table_box["x"]
            and table_highlight["y"] < table_box["y"] + table_box["height"]
            and table_highlight["y"] + table_highlight["height"] > table_box["y"],
            "表格搜索高亮没有覆盖目标段落",
        )
        page.screenshot(path=str(attempt_dir / "docx-table-search.png"), full_page=True)

        # Issue #53: merged/nested tables and per-section headers were never
        # asserted. Both are document semantics a browser can check exactly.
        page.locator('[data-testid="docx-sample-select"]').select_option("table-structures.docx")
        page.wait_for_function(
            """() => document.querySelector('[data-testid="loaded-file"]')?.textContent?.includes('table-structures.docx')
              && document.querySelector('[data-testid="docx-viewer"]')?.dataset.state === 'ready'""",
            timeout=30_000,
        )
        table_structure = page.evaluate(
            """() => {
              const hosts = [...document.querySelectorAll('[data-docx-table-host="true"]')];
              const outer = hosts.find((host) => !host.closest('[data-docx-table-cell="true"]'));
              if (!outer) return null;
              const table = outer.querySelector('table');
              if (!table) return null;
              // `table.querySelectorAll('colgroup > col')` also matches the
              // nested table's own column group, so only the outer table's
              // direct colgroup may be counted.
              const ownColgroup = [...table.children].find(
                (child) => child.tagName === 'COLGROUP'
              );
              const rows = [...table.tBodies[0].rows];
              const body = rows[0] ? [...rows[0].cells] : [];
              const vertical = table.querySelector('td[rowspan="3"]');
              const spanned = table.querySelector('td[colspan="2"]');
              const nested = [...table.querySelectorAll('td table')];
              const multiParagraph = [...table.querySelectorAll('td')].some(
                (cell) => cell.querySelectorAll('[data-docx-table-cell-paragraph-host="true"]').length === 3
              );
              return {
                columnCount: ownColgroup ? ownColgroup.querySelectorAll('col').length : 0,
                headerCellCount: body.length,
                rowCount: rows.length,
                verticalMergeText: vertical ? vertical.innerText.trim() : null,
                northRegionCount: (table.innerText.match(/North region/g) || []).length,
                gridSpanText: spanned ? spanned.innerText.trim() : null,
                nestedTables: nested.length,
                nestedText: nested.map((node) => node.innerText).join(' '),
                multiParagraphCell: multiParagraph,
              };
            }"""
        )
        expect(table_structure is not None, "table-structures.docx 没有渲染表格")
        assert table_structure is not None
        expect(table_structure["columnCount"] == 4, f"合并表格列数应为 4: {table_structure}")
        expect(table_structure["headerCellCount"] == 4, f"表头单元格数应为 4: {table_structure}")
        expect(
            table_structure["northRegionCount"] == 1,
            f"纵向合并单元格文本只能绘制一次: {table_structure}",
        )
        expect(
            (table_structure["verticalMergeText"] or "").startswith("North region"),
            f"rowspan=3 单元格不是合并后的单元格: {table_structure}",
        )
        expect(
            (table_structure["gridSpanText"] or "").startswith("Q1+Q2 combined"),
            f"colspan=2 单元格内容错误: {table_structure}",
        )
        expect(table_structure["nestedTables"] == 1, f"嵌套表格缺失: {table_structure}")
        expect(
            all(marker in table_structure["nestedText"] for marker in ["N1", "N2", "N3", "N4"]),
            f"嵌套表格单元格失失: {table_structure}",
        )
        expect(
            table_structure["multiParagraphCell"],
            f"多段落单元格未渲染三段段落: {table_structure}",
        )

        # Issue #53 §2: the running head contains a table. It must render through
        # the header's own table host on every page of the section, and it must
        # not be duplicated into the page body.
        table_pages = collect_all_docx_pages(page, expected_pages=2)
        expect(
            len(table_pages) >= 2,
            f"table-structures.docx 应渲染至少 2 页: {table_pages}",
        )
        header_table_rows = [
            page_info["headerTables"][0]["rows"] if page_info["headerTables"] else None
            for page_info in table_pages
        ]
        expect(
            header_table_rows == [HEADER_TABLE_ROWS] * len(table_pages),
            f"页眉表格必须在每一页渲染且保留单元格文本: {header_table_rows}",
        )
        expect(
            all(
                "Table Structures Fixture" in page_info["header"]
                for page_info in table_pages
            ),
            f"页眉表格不得挤掉运行标题: {[page_info['header'] for page_info in table_pages]}",
        )
        expect(
            all(
                "Table Register" not in page_info["bodyText"]
                for page_info in table_pages
            ),
            f"页眉表格不能重复渲染到正文: {[page_info['bodyText'][:80] for page_info in table_pages]}",
        )
        page.screenshot(path=str(attempt_dir / "docx-header-table.png"), full_page=True)
        page.screenshot(path=str(attempt_dir / "docx-table-structures.png"), full_page=True)

        page.locator('[data-testid="docx-sample-select"]').select_option("sections-columns.docx")
        page.wait_for_function(
            """() => document.querySelector('[data-testid="loaded-file"]')?.textContent?.includes('sections-columns.docx')
              && document.querySelector('[data-testid="docx-viewer"]')?.dataset.state === 'ready'""",
            timeout=30_000,
        )
        # The viewer only mounts the pages inside its scroll viewport, so the
        # pages are collected while stepping through the scroll range and merged
        # by page index; that keeps the per-page header assertion honest for a
        # document whose pages cannot all be mounted at once.
        sections = collect_all_docx_pages(page, expected_pages=4)
        expect(len(sections) == 4, f"sections-columns.docx 应渲染 4 页: {sections}")
        expect(
            [page_info["header"] for page_info in sections]
            == [
                "Section One Header",
                "Section One Header",
                "Section Two Landscape Header",
                "Section Three Two Column Header",
            ],
            f"每页必须渲染其所属节的页眉: {sections}",
        )
        expect(
            [page_info["footer"] for page_info in sections]
            == [
                "Section One Footer",
                "Section One Footer",
                "Section Two Footer",
                "Section Three Footer",
            ],
            f"每页必须渲染其所属节的页脚: {sections}",
        )
        expect(
            sections[2]["width"] > sections[0]["width"],
            f"横向节的页面必须比纵向节宽: {sections}",
        )
        expect(
            sections[3]["width"] == sections[0]["width"],
            f"分栏节回到纵向: {sections}",
        )
        expect(
            [page_info["watermarks"] for page_info in sections]
            == [1, 1, 0, 0],
            f"水印必须只出现在第一节且每页重复: {sections}",
        )
        page.screenshot(path=str(attempt_dir / "docx-sections.png"), full_page=True)
        evidence.assert_clean()
        return {
            "status": "PASS",
            "footnoteText": footnote_text,
            "endnoteText": endnote_text,
            "searchStatus": status,
            "tableSearchStatus": table_search_status,
            "searchHighlight": search_highlight,
            "tableSearchHighlight": table_highlight,
            "tableStructure": table_structure,
            "headerTableRows": header_table_rows,
            "sections": sections,
        }
    finally:
        evidence.save(attempt_dir)
        context.close()


XLSX_CANVAS_PIXEL_PROBE = """() => {
  const grid = document.querySelector('[data-testid="xlsx-grid"]');
  const canvas = grid?.querySelector('canvas.xlsx-grid__body');
  if (!grid || !canvas) return null;
  const cssWidth = canvas.getBoundingClientRect().width;
  if (!cssWidth) return null;
  const dpr = canvas.width / cssWidth;
  const colWidth = Number(grid.dataset.firstColumnWidth) || 80;
  const rowHeight = Number(grid.dataset.firstRowHeight) || 20;
  const ctx = canvas.getContext('2d');
  const sample = (firstCol, firstRow, cols, rows) => {
    const x = Math.round(firstCol * colWidth * dpr);
    const y = Math.round(firstRow * rowHeight * dpr);
    const width = Math.max(1, Math.round(cols * colWidth * dpr));
    const height = Math.max(1, Math.round(rows * rowHeight * dpr));
    const data = ctx.getImageData(x, y, width, height).data;
    let nonWhite = 0, red = 0, green = 0, blue = 0;
    for (let index = 0; index < data.length; index += 4) {
      const r = data[index], g = data[index + 1], b = data[index + 2], a = data[index + 3];
      if (a < 16 || (r > 245 && g > 245 && b > 245)) continue;
      nonWhite += 1;
      if (r > g + 24 && r > b + 24) red += 1;
      else if (g > r + 16 && g > b + 8) green += 1;
      else if (b > r + 24 && b > g + 10) blue += 1;
    }
    return { nonWhite, red, green, blue };
  };
  return {
    colWidth,
    rowHeight,
    colorScaleMin: sample(1, 1, 1, 1),
    colorScaleMax: sample(1, 4, 1, 1),
    dataBar: sample(2, 2, 1, 1),
    sparkline: sample(3, 1, 1, 1),
    iconSet: sample(3, 2, 1, 3),
  };
}"""


def run_xlsx(browser, attempt_dir: Path) -> dict[str, object]:
    context = browser.new_context(viewport={"width": 1440, "height": 900}, locale="zh-CN")
    page = context.new_page()
    evidence = BrowserEvidence(page)
    try:
        page.goto(f"{BASE_URL}/#/xlsx-viewer", wait_until="networkidle")
        page.locator('[data-testid="xlsx-read-only"]').check()
        page.locator('[data-testid="xlsx-sample-select"]').select_option("charts-images.xlsx")
        page.locator('[data-testid="xlsx-load-sample"]').click()
        page.locator('[data-testid="page-status"][data-state="ready"]').wait_for(timeout=30_000)
        expected = {
            "xlsx-hyperlink-count": 1,
            "xlsx-comment-count": 1,
            "xlsx-shape-count": 1,
            "xlsx-form-control-count": 1,
        }
        expect(page.locator('[data-testid="xlsx-worker-actual"]').inner_text() == "Worker", "只读正式用例没有使用 Worker")
        observed: dict[str, int] = {}
        for testid, count in expected.items():
            locator = page.locator(f'[data-testid="{testid}"]')
            locator.wait_for(timeout=10_000)
            value = int(locator.inner_text().split("：")[-1].strip())
            observed[testid] = value
            expect(value == count, f"{testid} 期望 {count}，实际 {value}")
        expect(page.locator('[data-testid="xlsx-shape"]').count() == 1, "形状展示层缺失")
        expect(page.locator('[data-testid="xlsx-form-control"]').count() == 1, "表单控件展示层缺失")

        grid = page.locator('[data-testid="xlsx-grid"]')
        box = grid.bounding_box()
        expect(box is not None, "网格没有可见尺寸")
        # Fixture features that were previously unasserted: merged cells, a
        # freeze on both axes and the list validation rule.
        expect(grid.get_attribute("data-merged-region-count") == "1", "合并区域数量错误")
        expect(grid.get_attribute("data-freeze-row") == "2", "冻结行数量错误")
        expect(grid.get_attribute("data-freeze-column") == "2", "冻结列数量错误")
        expect(int(grid.get_attribute("data-data-validation-count") or 0) >= 1, "数据验证规则缺失")

        # Conditional formats and sparklines used to be asserted against the
        # model that computed them. Sample the real canvas instead: B2/B5 are
        # the colour-scale endpoints, C3 carries a data bar, D2 a sparkline and
        # D3:D5 the traffic-light icon set.
        pixels = None
        deadline = time.monotonic() + 10
        while time.monotonic() < deadline:
            pixels = page.evaluate(XLSX_CANVAS_PIXEL_PROBE)
            if (
                pixels
                and pixels["colorScaleMax"]["green"] > 40
                and pixels["dataBar"]["blue"] > 40
            ):
                break
            page.wait_for_timeout(200)
        expect(pixels is not None, "XLSX 正文画布没有可采样的像素")
        if pixels:
            expect(pixels["colWidth"] > 0 and pixels["rowHeight"] > 0, pixels)
            expect(
                pixels["colorScaleMin"]["red"] > 40,
                f"色阶最小端未绘制: {pixels['colorScaleMin']}",
            )
            expect(
                pixels["colorScaleMax"]["green"] > 40,
                f"色阶最大端未绘制: {pixels['colorScaleMax']}",
            )
            expect(
                pixels["dataBar"]["blue"] > 40,
                f"数据条未绘制: {pixels['dataBar']}",
            )
            expect(
                pixels["sparkline"]["blue"] > 3,
                f"迷你图未绘制: {pixels['sparkline']}",
            )
            expect(
                pixels["iconSet"]["nonWhite"] > 0
                and (pixels["iconSet"]["red"] + pixels["iconSet"]["green"]) > 0,
                f"图标集未绘制: {pixels['iconSet']}",
            )

        first_width = float(grid.get_attribute("data-first-column-width") or 80)
        page.mouse.move(box["x"] + 48 + first_width + 20, box["y"] + 24 + 24 + 12)
        page.locator('[data-testid="xlsx-comment-popover"]').wait_for(timeout=5_000)
        expect("财务复核" in page.locator('[data-testid="xlsx-comment-popover"]').inner_text(), "批注浮层内容错误")

        page.mouse.click(box["x"] + 48 + 20, box["y"] + 24 + 24 + 12)
        page.locator('[data-testid="xlsx-name-box"]').wait_for()
        page.wait_for_timeout(250)
        expect(page.locator('[data-testid="xlsx-name-box"]').input_value() == "A7", "内部超链接没有跳转到 A7")
        page.screenshot(path=str(attempt_dir / "xlsx-rich-content.png"), full_page=True)
        evidence.assert_clean()
        return {"status": "PASS", "counts": observed, "pixels": pixels, "hyperlinkTarget": "A7"}
    finally:
        evidence.save(attempt_dir)
        context.close()


def run_pdf(browser, attempt_dir: Path) -> dict[str, object]:
    context = browser.new_context(
        viewport={"width": 1440, "height": 900}, locale="zh-CN"
    )
    page = context.new_page()
    evidence = BrowserEvidence(page)

    def wait_pdf_ready(file_name: str) -> None:
        page.wait_for_function(
            """expected => {
              const status = document.querySelector('[data-testid="page-status"]');
              const loaded = document.querySelector('[data-testid="loaded-file"]');
              const stage = document.querySelector('[data-testid="pdf-page"]');
              return status?.dataset.state === 'ready' &&
                loaded?.textContent?.includes(expected) &&
                stage?.dataset.renderState === 'rendered';
            }""",
            arg=file_name,
            timeout=45_000,
        )

    def page_text(page_number: int, marker: str) -> str:
        page.wait_for_function(
            """expected => {
              const stage = document.querySelector('[data-testid="pdf-page"]');
              const layer = document.querySelector('[data-testid="pdf-text-layer"]');
              return stage?.dataset.pageNumber === String(expected.pageNumber) &&
                stage?.dataset.renderState === 'rendered' &&
                (layer?.textContent ?? '').includes(expected.marker);
            }""",
            arg={"pageNumber": page_number, "marker": marker},
            timeout=45_000,
        )
        return page.locator('[data-testid="pdf-text-layer"]').text_content() or ""

    def goto_page(page_number: int, marker: str) -> str:
        input_box = page.locator('[data-testid="pdf-page-input"]')
        input_box.fill(str(page_number))
        input_box.press("Enter")
        return page_text(page_number, marker)

    def search(query: str) -> dict[str, object]:
        if page.locator('[data-testid="pdf-search-input"]').count() == 0:
            page.locator('[data-testid="pdf-search-toggle"]').click()
        page.locator('[data-testid="pdf-search-input"]').fill(query)
        page.locator('[data-testid="pdf-search-submit"]').click()
        page.wait_for_function(
            """() => /第\\s*\\d+\\s*\\/\\s*\\d+\\s*项/.test(
              document.querySelector('[data-testid="pdf-search-result"]')?.textContent ?? ''
            )""",
            timeout=45_000,
        )
        result_text = page.locator('[data-testid="pdf-search-result"]').inner_text()
        status = re.search(r"第\s*(\d+)\s*/\s*(\d+)\s*项", result_text)
        expect(status is not None, f"搜索状态缺少命中计数: {result_text}")
        assert status is not None
        context_text = page.locator(".pdf-search-context").inner_text()
        expect(query in context_text, f"搜索上下文未包含 {query}: {context_text}")
        return {
            "status": result_text,
            "activeIndex": int(status.group(1)),
            "total": int(status.group(2)),
            "context": context_text,
        }

    try:
        page.goto(f"{BASE_URL}/#/pdf-viewer", wait_until="networkidle")
        wait_pdf_ready("sample.pdf")

        # Per-page extraction: every page of sample.pdf carries its own marker.
        per_page: dict[str, str] = {}
        for number in range(1, 5):
            text = goto_page(number, f"ALPHA-{number}")
            expect(f"Section {number}" in text, f"第 {number} 页正文缺失 Section {number}: {text[:120]}")
            expect(f"ALPHA-{number}" in text, f"第 {number} 页正文缺失 ALPHA-{number}: {text[:120]}")
            per_page[f"page-{number}"] = text

        # Multi-hit search must report every match, not only the active page.
        multi_hit = search("PdfViewer")
        expect(
            int(multi_hit["total"]) >= 16,
            f"PdfViewer 命中数不足: {multi_hit}",
        )
        first_index = int(multi_hit["activeIndex"])
        page.locator('[data-testid="pdf-search-next"]').click()
        page.wait_for_function(
            """previous => {
              const text = document.querySelector('[data-testid="pdf-search-result"]')?.textContent ?? '';
              const match = text.match(/第\\s*(\\d+)\\s*\\/\\s*(\\d+)\\s*项/);
              return match && Number(match[1]) !== previous;
            }""",
            arg=first_index,
            timeout=45_000,
        )
        moved = page.locator('[data-testid="pdf-search-result"]').inner_text()
        moved_status = re.search(r"第\s*(\d+)\s*/\s*(\d+)\s*项", moved)
        assert moved_status is not None
        expect(
            int(moved_status.group(2)) == int(multi_hit["total"]),
            f"翻到下一个命中后总数变化: {multi_hit} -> {moved}",
        )
        screenshot_path = attempt_dir / "pdf-text-fidelity.png"
        page.screenshot(path=str(screenshot_path), full_page=True)

        # CJK extraction and search: the fixture has a real Chinese text layer.
        page.locator('[data-testid="pdf-sample-select"]').select_option("cjk-report.pdf")
        wait_pdf_ready("cjk-report.pdf")
        cjk_page_1 = page_text(1, "CJK-ALPHA-1")
        expect("CJK-ALPHA-1" in cjk_page_1, f"中文第 1 页缺少标记: {cjk_page_1[:120]}")
        expect("合同" in cjk_page_1, f"中文第 1 页缺少合同关键字: {cjk_page_1[:120]}")
        cjk_page_2 = goto_page(2, "CJK-ALPHA-2")
        expect("CJK-ALPHA-2" in cjk_page_2, f"中文第 2 页缺少标记: {cjk_page_2[:120]}")
        cjk_search = search("合同")
        expect(
            int(cjk_search["total"]) >= 6,
            f"中文关键字命中数不足: {cjk_search}",
        )
        page.screenshot(path=str(attempt_dir / "pdf-cjk-search.png"), full_page=True)
        evidence.assert_clean()
        return {
            "status": "PASS",
            "perPageText": {key: len(value) for key, value in per_page.items()},
            "multiHit": multi_hit,
            "multiHitAfterNext": moved,
            "cjkPageLengths": [len(cjk_page_1), len(cjk_page_2)],
            "cjkSearch": cjk_search,
        }
    finally:
        evidence.save(attempt_dir)
        context.close()


def main() -> int:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    if MANAGE_PREVIEW and port_in_use():
        raise RuntimeError(f"port {PORT} is already in use")
    preview_log = (OUTPUT / "preview.log").open("wb") if MANAGE_PREVIEW else None
    preview = subprocess.Popen(
        [
            "pnpm",
            "--filter",
            "demo",
            "preview",
            "--host",
            "127.0.0.1",
            "--port",
            str(PORT),
        ],
        cwd=ROOT,
        stdout=preview_log,
        stderr=subprocess.STDOUT,
        start_new_session=True,
    ) if MANAGE_PREVIEW else None
    summary: dict[str, object] = {"baseUrl": BASE_URL, "cases": {}}
    failed = False
    try:
        if preview is not None:
            wait_for_preview(preview)
        with sync_playwright() as playwright:
            browser, browser_source = launch_browser(playwright)
            summary["browser"] = browser.version
            summary["browserSource"] = browser_source
            try:
                for case_id, runner in (
                    ("DOCX-NOTES-SEARCH", run_docx),
                    ("XLSX-RICH-CONTENT", run_xlsx),
                    ("PDF-TEXT-FIDELITY", run_pdf),
                ):
                    attempts = []
                    for attempt in (1, 2):
                        attempt_dir = OUTPUT / case_id / f"attempt-{attempt}"
                        attempt_dir.mkdir(parents=True, exist_ok=True)
                        try:
                            result = runner(browser, attempt_dir)
                            attempts.append(result)
                            break
                        except Exception as error:
                            attempts.append({"status": "FAIL", "error": str(error)})
                    status = "PASS" if attempts[0]["status"] == "PASS" else "FLAKY" if len(attempts) > 1 and attempts[-1]["status"] == "PASS" else "FAIL"
                    summary["cases"][case_id] = {"status": status, "attempts": attempts}
                    if status != "PASS":
                        failed = True
            finally:
                browser.close()
        summary["result"] = "FAIL" if failed else "PASS"
        summary["status"] = summary["result"]
        (OUTPUT / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        return 1 if failed else 0
    finally:
        stop_process(preview)
        if preview_log is not None:
            preview_log.close()


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as error:
        OUTPUT.mkdir(parents=True, exist_ok=True)
        summary = {
            "suite": "BB-FIDELITY-CONTENT",
            "result": "BLOCKED",
            "status": "BLOCKED",
            "error": str(error),
        }
        (OUTPUT / "summary.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"BLOCKED: {error}", file=sys.stderr)
        sys.exit(2)
