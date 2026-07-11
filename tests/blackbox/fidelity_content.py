from __future__ import annotations

import json
import os
from pathlib import Path
import sys

from playwright.sync_api import sync_playwright

from browser_evidence import BrowserEvidence


ROOT = Path(__file__).resolve().parents[2]
OUTPUT = Path(os.environ.get(
    "BLACKBOX_EVIDENCE_DIR",
    ROOT / "output" / "acceptance" / "fidelity-content",
)).resolve()
BASE_URL = os.environ.get("BLACKBOX_BASE_URL", "http://127.0.0.1:4173")


def expect(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


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
        expect(page.locator('[data-docx-search-match="true"]').count() > 0, "搜索结果没有高亮")
        page.screenshot(path=str(attempt_dir / "docx-notes-search.png"), full_page=True)
        evidence.assert_clean()
        return {"status": "PASS", "footnoteText": footnote_text, "endnoteText": endnote_text, "searchStatus": status}
    finally:
        evidence.save(attempt_dir)
        context.close()


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
            "xlsx-conditional-format-count": 3,
            "xlsx-sparkline-count": 1,
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
        return {"status": "PASS", "counts": observed, "hyperlinkTarget": "A7"}
    finally:
        evidence.save(attempt_dir)
        context.close()


def main() -> int:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    summary: dict[str, object] = {"baseUrl": BASE_URL, "cases": {}}
    failed = False
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        try:
            for case_id, runner in (("DOCX-NOTES-SEARCH", run_docx), ("XLSX-RICH-CONTENT", run_xlsx)):
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
    summary["status"] = "FAIL" if failed else "PASS"
    (OUTPUT / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
