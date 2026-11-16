"""Scenario matrix (V6): viewports, dark, zoom, print, keyboard, touch, IME.

    LOCAL_DEMO_URL=http://127.0.0.1:4173 \
    uv run --with playwright python tests/blackbox/scenario_matrix.py \
        --output verification/evidence/2026-11-16/scenario/summary.json

This is observational: each scenario records the measured behaviour and the
known issue it exercises (`issues` field), rather than failing on every gap.
It fails only on crashes or unhandled page errors.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
LOCAL_URL = os.environ.get("LOCAL_DEMO_URL", "http://127.0.0.1:4173").strip().rstrip("/")
DOCX = ROOT / "apps" / "demo" / "public" / "samples" / "demo.docx"
XLSX = ROOT / "test-data" / "corpus" / "xlsx-validation.xlsx"
PDF = ROOT / "test-data" / "corpus" / "pdf-rotated-90.pdf"
PPTX = ROOT / "test-data" / "corpus" / "pptx-deck-basic.pptx"
VIEWPORTS = [(1440, 900), (1280, 720), (768, 1024), (390, 844)]


def load(page, route: str, selector: str, path: Path, wait_ms: int = 4000) -> None:
    page.goto(LOCAL_URL + route, wait_until="networkidle")
    page.wait_for_timeout(700)
    page.locator(selector).first.set_input_files(str(path))
    page.wait_for_timeout(wait_ms)


def viewport_matrix(browser) -> list[dict]:
    rows: list[dict] = []
    cases = [
        ("docx", "/#/docx-viewer", 'input[type="file"][accept*="docx"]', DOCX),
        ("xlsx", "/#/xlsx-viewer", 'input[type="file"][accept*="xlsx"]', XLSX),
        ("pdf", "/#/pdf-surface", 'input[type="file"][accept*="pdf"]', PDF),
        ("pptx", "/#/pptx-viewer", 'input[type="file"][accept*="pptx"]', PPTX),
    ]
    for width, height in VIEWPORTS:
        for name, route, selector, path in cases:
            page = browser.new_page(viewport={"width": width, "height": height})
            errors: list[str] = []
            page.on("pageerror", lambda exc: errors.append(str(exc)))
            try:
                load(page, route, selector, path)
                data = page.evaluate(
                    """() => ({
                      scrollWidth: document.documentElement.scrollWidth,
                      innerWidth: window.innerWidth,
                      mainHeight: Math.round((document.querySelector('main') ?? document.body).getBoundingClientRect().height),
                      state: document.querySelector('[data-testid="docx-viewer"],[data-testid="pptx-viewer"]')?.getAttribute('data-state') ?? null,
                    })"""
                )
            except Exception as exc:  # noqa: BLE001
                data = {"error": f"{type(exc).__name__}: {exc}"}
            data["overflow"] = data.get("scrollWidth", 0) > data.get("innerWidth", 0) + 1
            data["errors"] = errors[:3]
            rows.append({"viewport": f"{width}x{height}", "viewer": name, **data})
            page.close()
    return rows


def docx_scenarios(browser) -> dict:
    result: dict = {}
    page = browser.new_page(viewport={"width": 1440, "height": 900})
    load(page, "/#/docx-viewer", 'input[type="file"][accept*="docx"]', DOCX)
    result["baseline"] = page.evaluate(
        """() => ({
          pages: Number(document.querySelector('[data-docx-page-count]')?.getAttribute('data-docx-page-count') || 0),
          height: Math.round(document.querySelector('[data-docx-page-surface]')?.getBoundingClientRect().height || 0),
          bg: getComputedStyle(document.querySelector('[data-docx-page-surface]')).backgroundColor,
        })"""
    )
    # zoom 200 / 50
    for zoom in ("200", "50"):
        page.select_option('[data-testid="docx-zoom-select"]', zoom)
        page.wait_for_timeout(1200)
        result[f"zoom_{zoom}"] = page.evaluate(
            """() => ({
              height: Math.round(document.querySelector('[data-docx-page-surface]')?.getBoundingClientRect().height || 0),
              cardTop: Math.round(document.querySelector('.docx-gutter-card')?.getBoundingClientRect().top ?? 0),
            })"""
        )
    page.select_option('[data-testid="docx-zoom-select"]', "100")
    page.wait_for_timeout(800)
    # dark mode
    page.click('[data-testid="docx-theme-toggle"]')
    page.wait_for_timeout(600)
    result["dark"] = page.evaluate(
        """() => ({
          bg: getComputedStyle(document.querySelector('[data-docx-page-surface]')).backgroundColor,
          color: getComputedStyle(document.querySelector('[data-docx-page-surface]')).color,
        })"""
    )
    # print media: chrome should be hidden if print CSS exists
    page.emulate_media(media="print")
    result["print"] = page.evaluate(
        """() => ({
          toolbarDisplay: getComputedStyle(document.querySelector('[data-testid="docx-viewer-toolbar"]')).display,
          mountedPages: document.querySelectorAll('[data-docx-page-wrapper]').length,
        })"""
    )
    # keyboard page navigation
    page.emulate_media(media="screen")
    page.focus('[data-testid="docx-viewer"]')
    before = page.evaluate("() => document.querySelector('.docx-viewer-root')?.scrollTop ?? 0")
    page.keyboard.press("PageDown")
    page.wait_for_timeout(500)
    after = page.evaluate("() => document.querySelector('.docx-viewer-root')?.scrollTop ?? 0")
    result["keyboard"] = {"before": before, "after": after}
    page.close()
    return result


def xlsx_ime(browser) -> dict:
    page = browser.new_page(viewport={"width": 1440, "height": 900})
    load(page, "/#/xlsx-viewer", 'input[type="file"][accept*="xlsx"]', XLSX)
    canvas = page.locator("canvas.xlsx-grid__body").first
    box = canvas.bounding_box()
    page.mouse.click(box["x"] + 60, box["y"] + 60)
    page.wait_for_timeout(400)
    page.keyboard.type("abc")
    page.wait_for_timeout(400)
    editor = page.locator('[data-testid="xlsx-cell-editor"]')
    visible = editor.count() > 0
    if not visible:
        page.keyboard.press("Enter")
        page.wait_for_timeout(300)
        visible = editor.count() > 0
    state: dict = {"editor_opened": visible}
    if visible:
        try:
            editor.dispatch_event("compositionstart")
            page.keyboard.type("zhong")
            page.keyboard.press("Enter")
            page.wait_for_timeout(300)
            state.update({
                "editor_after_enter_during_composition": editor.count(),
                "editor_closed_on_composition_enter": editor.count() == 0,
            })
            page.keyboard.press("Escape")
            if editor.count():
                editor.dispatch_event("compositionend")
        except Exception as exc:  # editor may be replaced by the commit
            state["ime_error"] = f"{type(exc).__name__}: {exc}"
    else:
        state["editor_count"] = 0
    page.close()
    return {"editor_opened": visible, **state}


def touch_case(browser) -> dict:
    context = browser.new_context(viewport={"width": 390, "height": 844}, has_touch=True, is_mobile=True)
    page = context.new_page()
    errors: list[str] = []
    page.on("pageerror", lambda exc: errors.append(str(exc)))
    load(page, "/#/docx-viewer", 'input[type="file"][accept*="docx"]', DOCX)
    page.tap(".docx-page-surface")
    page.wait_for_timeout(500)
    data = page.evaluate("() => ({ pages: Number(document.querySelector('[data-docx-page-count]')?.getAttribute('data-docx-page-count') || 0) })")
    context.close()
    return {**data, "page_errors": errors[:3]}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output")
    args = parser.parse_args()
    from playwright.sync_api import sync_playwright

    report: dict = {"suite": "scenario-matrix", "date": "2026-11-16", "issues": {"keyboard": 116, "print": 92, "dark": 92, "ime": 89, "touch": 106}}
    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            executable_path="/usr/bin/chromium" if Path("/usr/bin/chromium").exists() else None,
            args=["--no-sandbox", "--disable-gpu"],
        )
        report["viewports"] = viewport_matrix(browser)
        report["docx"] = docx_scenarios(browser)
        report["xlsx_ime"] = xlsx_ime(browser)
        report["touch"] = touch_case(browser)
        browser.close()

    failures: list[str] = []
    for row in report["viewports"]:
        if row.get("error") or row.get("errors"):
            failures.append(f"{row['viewport']}/{row['viewer']}: {row.get('error') or row['errors']}")
        if row.get("overflow"):
            failures.append(f"{row['viewport']}/{row['viewer']}: horizontal overflow")
    if report["touch"].get("page_errors"):
        failures.append("touch: page errors")
    report["failures"] = failures
    report["verdict"] = "FAIL" if failures else "PASS"
    if args.output:
        output = Path(args.output)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"failures": failures, "docx": report["docx"], "ime": report["xlsx_ime"], "verdict": report["verdict"]}, ensure_ascii=False, indent=2))
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
