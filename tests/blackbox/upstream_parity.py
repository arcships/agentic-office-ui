"""Dual-run parity harness: upstream Extend playground vs the local demo.

This is an **opt-in** blackbox check: it needs the upstream playgrounds running
(they are not vendored) and a local preview at LOCAL_DEMO_URL.

    # terminal 1 – upstream DOCX
    cd /tmp/react-docx/apps/playground && ./node_modules/.bin/vite --port 5181
    # terminal 2 – upstream XLSX (run scripts/copy-duke-wasm.mjs once)
    cd /tmp/react-xlsx/apps/playground && ./node_modules/.bin/vite --port 5182
    # terminal 3 – local demo
    cd apps/demo/dist && python3 -m http.server 4173

    UPSTREAM_DOCX_URL=http://localhost:5181 \
    UPSTREAM_XLSX_URL=http://localhost:5182 \
    LOCAL_DEMO_URL=http://127.0.0.1:4173 \
    uv run --with playwright python tests/blackbox/upstream_parity.py --strict

Without the upstream URLs the script prints SKIP and exits 0, so it is safe to
keep in the suite. `--strict` compares the local metrics against
`tests/blackbox/upstream_parity_baseline.json` (regression guard); the upstream
numbers are always reported as the parity target.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from pathlib import Path

from playwright.sync_api import Page, sync_playwright

ROOT = Path(__file__).resolve().parents[2]
CORPUS = ROOT / "test-data" / "corpus"
BASELINE = Path(__file__).with_name("upstream_parity_baseline.json")
OUTPUT = Path(os.environ.get("UPSTREAM_PARITY_OUTPUT", ROOT / "output" / "acceptance" / "upstream-parity"))

DOCX_URL = os.environ.get("UPSTREAM_DOCX_URL", "").strip()
XLSX_URL = os.environ.get("UPSTREAM_XLSX_URL", "").strip()
LOCAL_URL = os.environ.get("LOCAL_DEMO_URL", "http://127.0.0.1:4173").strip().rstrip("/")
LOCAL_DOCX_URL = os.environ.get("LOCAL_DOCX_URL", f"{LOCAL_URL}/#/docx-viewer")
LOCAL_XLSX_URL = os.environ.get("LOCAL_XLSX_URL", f"{LOCAL_URL}/#/xlsx-viewer")


def upload_docx(page: Page, path: Path, local: bool) -> None:
    """Upstream: Open button + file chooser; local: hidden file input."""
    if local:
        page.locator('input[type="file"][accept*="docx"]').first.set_input_files(str(path))
    else:
        with page.expect_file_chooser() as info:
            page.get_by_role("button", name="Open").click()
        info.value.set_files(str(path))


def upload_xlsx(page: Page, path: Path) -> None:
    page.locator('input[type="file"][accept*="xlsx"]').first.set_input_files(str(path))


def page_total_from_text(page: Page) -> int | None:
    text = page.locator("body").inner_text()
    match = re.search(r"Page\s+\d+\s+of\s+(\d+)", text)
    return int(match.group(1)) if match else None


def run_docx(browser, url: str, fixture: Path, local: bool) -> dict:
    page = browser.new_page(viewport={"width": 1440, "height": 900})
    page.goto(url, wait_until="networkidle")
    page.wait_for_timeout(1200)
    upload_docx(page, fixture, local)
    page.wait_for_timeout(5000)
    if local:
        page.wait_for_selector("[data-docx-page-count]", timeout=30000)
        result = {
            "pages": int(page.get_attribute("[data-docx-page-count]", "data-docx-page-count") or 0),
            "cards": page.locator(".docx-gutter-card").count(),
            "first_page_height": page.evaluate(
                "() => Math.round(document.querySelector('[data-docx-page-surface]')?.getBoundingClientRect().height ?? 0)"
            ),
        }
    else:
        result = {"pages": page_total_from_text(page), "cards": page.locator(".docx-gutter-card").count()}
    page.close()
    return result


def run_xlsx(browser, url: str, fixture: Path, local: bool) -> dict:
    page = browser.new_page(viewport={"width": 1440, "height": 900})
    page.goto(url, wait_until="networkidle")
    page.wait_for_timeout(1200)
    upload_xlsx(page, fixture)
    page.wait_for_timeout(6000)
    text = page.locator("body").inner_text()
    titles = [name for name in ("Bar", "Line", "Pie", "Area", "Radar", "Scatter") if name in text]
    result = {
        "chart_titles": titles,
        "charts": page.locator('[data-testid="xlsx-chart-item"]').count() if local else len(titles),
        "error": bool(re.search(r"could not load|failed to load", text, re.I)),
    }
    page.close()
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--strict", action="store_true", help="fail when local metrics drift from the baseline")
    args = parser.parse_args()

    if not DOCX_URL or not XLSX_URL:
        print("SKIP: set UPSTREAM_DOCX_URL and UPSTREAM_XLSX_URL to run the dual-run parity check")
        return 0

    baseline = json.loads(BASELINE.read_text(encoding="utf-8"))
    report: dict = {"upstream": {}, "local": {}, "baseline": baseline}
    failures: list[str] = []

    with sync_playwright() as pw:
        chromium = os.environ.get("UPSTREAM_PARITY_CHROMIUM", "/usr/bin/chromium")
        launch_args = ["--no-sandbox", "--disable-gpu"]
        browser = (
            pw.chromium.launch(executable_path=chromium, args=launch_args)
            if Path(chromium).exists()
            else pw.chromium.launch(args=launch_args)
        )

        demo = ROOT / "apps" / "demo" / "public" / "samples" / "demo.docx"
        comments = CORPUS / "docx-comments.docx"
        charts = CORPUS / "xlsx-charts.xlsx"

        report["upstream"]["docx_demo"] = run_docx(browser, DOCX_URL, demo, local=False)
        report["local"]["docx_demo"] = run_docx(browser, LOCAL_DOCX_URL, demo, local=True)
        report["upstream"]["docx_comments"] = run_docx(browser, DOCX_URL, comments, local=False)
        report["local"]["docx_comments"] = run_docx(browser, LOCAL_DOCX_URL, comments, local=True)
        report["upstream"]["xlsx_charts"] = run_xlsx(browser, XLSX_URL, charts, local=False)
        report["local"]["xlsx_charts"] = run_xlsx(browser, LOCAL_XLSX_URL, charts, local=True)

        browser.close()

    if args.strict:
        for key in ("docx_demo", "docx_comments"):
            expected = baseline.get(key, {}).get("local_pages")
            actual = report["local"][key].get("pages")
            if expected is not None and actual != expected:
                failures.append(f"{key}: local pages {actual} != baseline {expected}")
        expected_charts = baseline.get("xlsx_charts", {}).get("local_charts")
        actual_charts = report["local"]["xlsx_charts"].get("charts")
        if expected_charts is not None and actual_charts != expected_charts:
            failures.append(f"xlsx_charts: local charts {actual_charts} != baseline {expected_charts}")

    report["status"] = "FAIL" if failures else "PASS"
    report["failures"] = failures
    OUTPUT.mkdir(parents=True, exist_ok=True)
    (OUTPUT / "summary.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    print(json.dumps({k: report[k] for k in ("upstream", "local", "failures")}, ensure_ascii=False, indent=2))
    print(f"status: {report['status']} (summary: {OUTPUT / 'summary.json'})")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
