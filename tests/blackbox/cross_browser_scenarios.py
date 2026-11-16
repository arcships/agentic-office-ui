"""Cross-browser scenario subset for Chromium/Firefox (V3 × V6).

    LOCAL_DEMO_URL=http://127.0.0.1:4173 \
    uv run --with playwright python tests/blackbox/cross_browser_scenarios.py \
        --output verification/evidence/2026-11-16/scenario/cross-chromium-firefox.json

Covers, per engine: DOCX zoom/dark/print/PageDown/IME, XLSX zoom/print,
PPTX zoom/print, PDF print. WebKit uses `webkit_webdriver_scenarios.py`.
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
XLSX = ROOT / "test-data" / "corpus" / "xlsx-charts.xlsx"
PDF = ROOT / "test-data" / "corpus" / "pdf-rotated-90.pdf"
PPTX = ROOT / "test-data" / "corpus" / "pptx-deck-basic.pptx"


def load(page, route: str, selector: str, path: Path, wait_ms: int = 4000) -> None:
    page.goto(LOCAL_URL + route, wait_until="networkidle")
    page.wait_for_timeout(700)
    page.locator(selector).first.set_input_files(str(path))
    page.wait_for_timeout(wait_ms)


def docx(page) -> dict:
    load(page, "/#/docx-viewer", 'input[type="file"][accept*="docx"]', DOCX)
    result = {
        "baseline": page.evaluate(
            """() => ({
              pages: Number(document.querySelector('[data-docx-page-count]')?.getAttribute('data-docx-page-count') || 0),
              height: Math.round(document.querySelector('[data-docx-page-surface]')?.getBoundingClientRect().height || 0),
              bg: getComputedStyle(document.querySelector('[data-docx-page-surface]')).backgroundColor,
            })"""
        )
    }
    select = page.locator('[data-testid="docx-zoom-select"]')
    for zoom in ("200", "50"):
        select.select_option(zoom)
        page.wait_for_timeout(900)
        result[f"zoom_{zoom}"] = page.evaluate(
            "() => Math.round(document.querySelector('[data-docx-page-surface]')?.getBoundingClientRect().height || 0)"
        )
    select.select_option("100")
    page.wait_for_timeout(700)
    page.click('[data-testid="docx-theme-toggle"]')
    page.wait_for_timeout(500)
    result["dark"] = page.evaluate("() => getComputedStyle(document.querySelector('[data-docx-page-surface]')).backgroundColor")
    page.emulate_media(media="print")
    result["print"] = page.evaluate(
        """() => ({
          toolbar: getComputedStyle(document.querySelector('[data-testid="docx-viewer-toolbar"]')).display,
          mounted: document.querySelectorAll('[data-docx-page-wrapper]').length,
        })"""
    )
    page.emulate_media(media="screen")
    page.focus('[data-testid="docx-viewer"]')
    before = page.evaluate("() => document.querySelector('.docx-viewer-root')?.scrollTop ?? 0")
    page.keyboard.press("PageDown")
    page.wait_for_timeout(400)
    after = page.evaluate("() => document.querySelector('.docx-viewer-root')?.scrollTop ?? 0")
    result["keyboard"] = {"before": before, "after": after}
    page.close()
    return result


def docx_ime(browser) -> dict:
    page = browser.new_page(viewport={"width": 1440, "height": 900})
    page.goto(f"{LOCAL_URL}/#/docx-editor", wait_until="networkidle")
    page.wait_for_selector('[contenteditable="true"]', timeout=30000)
    page.wait_for_timeout(1000)
    editable = page.locator('[contenteditable="true"]').first
    editable.click()
    page.keyboard.press("End")
    before = page.locator('[contenteditable="true"]').count()
    editable.dispatch_event("compositionstart")
    page.keyboard.type("zhong")
    page.keyboard.press("Enter")
    page.wait_for_timeout(300)
    after = page.locator('[contenteditable="true"]').count()
    editable.dispatch_event("compositionend")
    page.keyboard.press("Escape")
    page.close()
    return {"before": before, "after": after, "guard_held": before == after}


def xlsx(page) -> dict:
    load(page, "/#/xlsx-viewer", 'input[type="file"][accept*="xlsx"]', XLSX)
    result = {}
    select = page.locator('select[aria-label="Zoom level"]')
    for level in ("200", "50"):
        select.select_option(level)
        page.wait_for_timeout(800)
        result[f"zoom_{level}"] = page.evaluate(
            """() => { const c = document.querySelector('canvas.xlsx-grid__body');
              return { value: document.querySelector('select[aria-label="Zoom level"]')?.value,
                       canvas: c ? { w: c.width, h: c.height } : null }; }"""
        )
    select.select_option("100")
    page.emulate_media(media="print")
    result["print"] = page.evaluate(
        """() => ({
          toolbar: getComputedStyle(document.querySelector('[data-testid="xlsx-toolbar"]')).display,
          grid: document.querySelectorAll('canvas').length,
        })"""
    )
    page.close()
    return result


def pptx(page) -> dict:
    load(page, "/#/pptx-viewer", 'input[type="file"][accept*="pptx"]', PPTX)
    page.wait_for_timeout(800)
    buttons = page.locator(".pptx-viewer__zoom button")
    before = buttons.nth(1).inner_text()
    buttons.nth(2).click()
    buttons.nth(2).click()
    page.wait_for_timeout(600)
    after = buttons.nth(1).inner_text()
    page.emulate_media(media="print")
    result = {
        "zoom": {"before": before, "after": after},
        "print": page.evaluate(
            """() => ({
              toolbar: getComputedStyle(document.querySelector('.pptx-viewer__toolbar')).display,
              slides: document.querySelectorAll('.pptx-thumbnail').length,
            })"""
        ),
    }
    page.close()
    return result


def pdf(page) -> dict:
    load(page, "/#/pdf-surface", 'input[type="file"][accept*="pdf"]', PDF)
    page.emulate_media(media="print")
    result = {
        "print": page.evaluate(
            """() => ({
              surfaces: document.querySelectorAll('.pdf-surface__page-slot').length,
              bodyOverflow: getComputedStyle(document.body).overflow,
            })"""
        )
    }
    page.close()
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output")
    parser.add_argument("--browsers", default="chromium,firefox")
    args = parser.parse_args()
    from playwright.sync_api import Error as PlaywrightError
    from playwright.sync_api import sync_playwright

    report: dict = {"suite": "cross-browser-scenarios", "date": "2026-11-16", "browsers": {}, "failures": []}
    with sync_playwright() as pw:
        for name in [b.strip() for b in args.browsers.split(",") if b.strip()]:
            launcher = getattr(pw, name)
            kwargs = {"headless": True, "args": ["--no-sandbox", "--disable-gpu"]} if name == "chromium" else {"headless": True}
            if name == "chromium" and Path("/usr/bin/chromium").exists():
                kwargs["executable_path"] = "/usr/bin/chromium"
            try:
                browser = launcher.launch(**kwargs)
            except PlaywrightError as exc:
                report["browsers"][name] = {"skipped": str(exc).splitlines()[0]}
                continue
            results: dict = {}
            for case_name, fn in {"docx": docx, "xlsx": xlsx, "pptx": pptx, "pdf": pdf}.items():
                page = browser.new_page(viewport={"width": 1440, "height": 900})
                errors: list[str] = []
                page.on("pageerror", lambda exc: errors.append(str(exc)))
                try:
                    data = fn(page)
                except Exception as exc:  # noqa: BLE001
                    data = {"error": f"{type(exc).__name__}: {exc}"}
                data["page_errors"] = errors[:3]
                results[case_name] = data
                if data.get("error") or data["page_errors"]:
                    report["failures"].append(f"{name}/{case_name}")
                page.close()
            results["docx_ime"] = docx_ime(browser)
            if not results["docx_ime"]["guard_held"]:
                report["failures"].append(f"{name}/docx_ime")
            report["browsers"][name] = results
            browser.close()
    report["verdict"] = "FAIL" if report["failures"] else "PASS"
    if args.output:
        output = Path(args.output)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"browsers": {k: list(v) for k, v in report["browsers"].items()}, "failures": report["failures"], "verdict": report["verdict"]}, ensure_ascii=False, indent=2))
    return 1 if report["failures"] else 0


if __name__ == "__main__":
    sys.exit(main())
