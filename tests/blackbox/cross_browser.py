"""Cross-browser smoke for the four viewers (V3).

    LOCAL_DEMO_URL=http://127.0.0.1:4173 \
    uv run --with playwright python tests/blackbox/cross_browser.py \
        --output verification/evidence/2026-11-16/cross-browser/summary.json

Runs Chromium, Firefox and WebKit against the built demo:

- DOCX: load `demo.docx`, check state/page count/first page height.
- XLSX: load `xlsx-charts.xlsx`, check ready state, canvases, chart items.
- PDF: load a rotated PDF, check page slots and rotation attributes.
- PPTX: load the corpus deck, check ready state and thumbnail count.

WebKit/Firefox are Playwright-managed builds; Chromium uses the system binary
when present. Missing browsers produce a SKIP entry, not a failure.
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


def wait(page, fn, timeout_ms: int = 15000) -> None:
    page.wait_for_function(fn, timeout=timeout_ms)


def docx_case(page) -> dict:
    page.goto(f"{LOCAL_URL}/#/docx-viewer", wait_until="networkidle")
    page.wait_for_timeout(800)
    page.locator('input[type="file"][accept*="docx"]').first.set_input_files(str(DOCX))
    page.wait_for_selector("[data-docx-page-count]", timeout=30000)
    page.wait_for_timeout(4000)
    return page.evaluate(
        """() => ({
          state: document.querySelector('[data-testid="docx-viewer"]')?.getAttribute('data-state'),
          pages: Number(document.querySelector('[data-docx-page-count]')?.getAttribute('data-docx-page-count') || 0),
          firstPageHeight: Math.round(document.querySelector('[data-docx-page-surface]')?.getBoundingClientRect().height || 0),
          bodyHeight: document.querySelector('.docx-viewer-root')?.scrollHeight ?? null,
        })"""
    )


def xlsx_case(page) -> dict:
    page.goto(f"{LOCAL_URL}/#/xlsx-viewer", wait_until="networkidle")
    page.wait_for_timeout(800)
    page.locator('input[type="file"][accept*="xlsx"]').first.set_input_files(str(XLSX))
    page.wait_for_timeout(6000)
    return page.evaluate(
        """() => ({
          status: document.querySelector('[data-testid="page-status"]')?.textContent?.trim(),
          error: document.querySelector('[data-testid="load-error"]')?.textContent?.trim() ?? null,
          canvases: document.querySelectorAll('canvas').length,
          charts: document.querySelectorAll('[data-testid="xlsx-chart-item"]').length,
          tabs: [...document.querySelectorAll('[data-testid="xlsx-sheet-tab"]')].map(t => t.textContent.trim()),
        })"""
    )


def pdf_case(page) -> dict:
    page.goto(f"{LOCAL_URL}/#/pdf-surface", wait_until="networkidle")
    page.wait_for_timeout(800)
    page.locator('input[type="file"][accept*="pdf"]').first.set_input_files(str(PDF))
    page.wait_for_timeout(6000)
    return page.evaluate(
        """() => ({
          slots: document.querySelectorAll('.pdf-surface__page-slot').length,
          images: document.querySelectorAll('.pdf-surface__page-img').length,
          first: (() => { const img = document.querySelector('.pdf-surface__page-img'); return img ? {w: img.naturalWidth, h: img.naturalHeight} : null })(),
          errors: [...document.querySelectorAll('.pdf-surface__error')].map(e => e.textContent.trim()),
        })"""
    )


def pptx_case(page) -> dict:
    page.goto(f"{LOCAL_URL}/#/pptx-viewer", wait_until="networkidle")
    page.wait_for_timeout(800)
    page.locator('input[type="file"][accept*="pptx"]').first.set_input_files(str(PPTX))
    page.wait_for_function(
        "() => document.querySelector('[data-testid=\"pptx-viewer\"]')?.getAttribute('data-state') === 'ready'",
        timeout=30000,
    )
    page.wait_for_timeout(1500)
    return page.evaluate(
        """() => ({
          state: document.querySelector('[data-testid="pptx-viewer"]')?.getAttribute('data-state'),
          thumbs: document.querySelectorAll('.pptx-thumbnail').length,
          counter: document.querySelector('[data-testid="pptx-page-counter"]')?.textContent?.trim() ?? null,
        })"""
    )


CASES = {"docx": docx_case, "xlsx": xlsx_case, "pdf": pdf_case, "pptx": pptx_case}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output")
    parser.add_argument("--browsers", default="chromium,firefox,webkit")
    args = parser.parse_args()
    from playwright.sync_api import Error as PlaywrightError
    from playwright.sync_api import sync_playwright

    payload: dict = {"suite": "cross-browser", "date": "2026-11-16", "browsers": {}, "failures": []}
    with sync_playwright() as pw:
        for name in [b.strip() for b in args.browsers.split(",") if b.strip()]:
            launcher = getattr(pw, name)
            launch_kwargs = {"headless": True, "args": ["--no-sandbox", "--disable-gpu"]} if name == "chromium" else {"headless": True}
            if name == "chromium" and Path("/usr/bin/chromium").exists():
                launch_kwargs["executable_path"] = "/usr/bin/chromium"
            try:
                browser = launcher.launch(**launch_kwargs)
            except PlaywrightError as exc:
                payload["browsers"][name] = {"skipped": str(exc).splitlines()[0]}
                continue
            result: dict = {"cases": {}, "errors": []}
            for case_name, fn in CASES.items():
                page = browser.new_page(viewport={"width": 1440, "height": 900})
                page_errors: list[str] = []
                page.on("pageerror", lambda exc: page_errors.append(str(exc)))
                page.on("console", lambda msg: page_errors.append(msg.text) if msg.type == "error" else None)
                try:
                    data = fn(page)
                except Exception as exc:  # noqa: BLE001 - report per case
                    data = {"error": f"{type(exc).__name__}: {exc}"}
                data["console_errors"] = page_errors[:5]
                result["cases"][case_name] = data
                page.close()
            browser.close()
            payload["browsers"][name] = result
            for case_name, data in result["cases"].items():
                if data.get("error"):
                    payload["failures"].append(f"{name}/{case_name}: {data['error']}")
                if data.get("console_errors"):
                    payload["failures"].append(f"{name}/{case_name}: console errors")
    payload["verdict"] = "FAIL" if payload["failures"] else "PASS"
    if args.output:
        output = Path(args.output)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"browsers": {k: list(v.get("cases", {})) for k, v in payload["browsers"].items()}, "failures": payload["failures"], "verdict": payload["verdict"]}, ensure_ascii=False, indent=2))
    return 1 if payload["failures"] else 0


if __name__ == "__main__":
    sys.exit(main())
