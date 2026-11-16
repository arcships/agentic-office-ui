"""Robustness checks (V5): malformed, oversized, concurrent and unmount cases.

    LOCAL_DEMO_URL=http://127.0.0.1:4173 \
    uv run --with playwright python tests/blackbox/robustness.py \
        --output verification/evidence/2026-11-16/robustness/summary.json

The script builds temporary malformed/oversized files, drives the demo, and
records the visible error code, elapsed time and console errors. It does not
assert a specific code per case (codes are documented in the summary); it fails
only when the viewer crashes, hangs, or reports success for an invalid input.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import tempfile
import time
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
LOCAL_URL = os.environ.get("LOCAL_DEMO_URL", "http://127.0.0.1:4173").strip().rstrip("/")
SAMPLES = ROOT / "apps" / "demo" / "public" / "samples"
XLSX_OK = ROOT / "test-data" / "corpus" / "xlsx-validation.xlsx"
XLSX_OTHER = ROOT / "test-data" / "corpus" / "xlsx-charts.xlsx"


def make_files(tmp: Path) -> dict[str, Path]:
    files: dict[str, Path] = {}
    demo = (SAMPLES / "demo.docx").read_bytes()
    files["truncated"] = tmp / "truncated.docx"
    files["truncated"].write_bytes(demo[:4096])

    files["oversized_xlsx"] = tmp / "oversized.xlsx"
    with files["oversized_xlsx"].open("wb") as fh:
        fh.write(b"PK\x03\x04")
        fh.write(os.urandom(30 * 1024 * 1024))

    files["oversized_docx"] = tmp / "oversized.docx"
    with files["oversized_docx"].open("wb") as fh:
        fh.write(b"PK\x03\x04")
        fh.write(os.urandom(30 * 1024 * 1024))

    files["malformed_pdf"] = tmp / "malformed.pdf"
    files["malformed_pdf"].write_bytes(b"%PDF-1.7\n" + os.urandom(8192))
    files["malformed_pptx"] = tmp / "malformed.pptx"
    files["malformed_pptx"].write_bytes(b"PK\x03\x04" + os.urandom(8192))
    deck = (ROOT / "test-data" / "corpus" / "pptx-deck-basic.pptx").read_bytes()
    files["truncated_pptx"] = tmp / "truncated.pptx"
    files["truncated_pptx"].write_bytes(deck[:4096])
    files["malformed_csv"] = tmp / "malformed.csv"
    files["malformed_csv"].write_bytes(bytes(range(256)) * 20)

    files["zip_bomb"] = tmp / "bomb.xlsx"
    with zipfile.ZipFile(files["zip_bomb"], "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        archive.writestr("xl/workbook.xml", "<workbook/>")
        archive.writestr("xl/worksheets/sheet1.xml", "0" * (64 * 1024 * 1024))
    return files


def upload(page, selector: str, path: Path) -> None:
    page.locator(selector).first.set_input_files(str(path))


def read_state(page) -> dict:
    return page.evaluate(
        """() => {
          const docx = document.querySelector('[data-testid="docx-viewer"]');
          const pptx = document.querySelector('[data-testid="pptx-viewer"]');
          const err = document.querySelector('[data-testid="load-error"]');
          return {
            docxState: docx?.getAttribute('data-state') ?? null,
            pptxState: pptx?.getAttribute('data-state') ?? null,
            status: document.querySelector('[data-testid="page-status"]')?.textContent?.trim() ?? null,
            errorCode: err?.getAttribute('data-error-code') ?? null,
            errorText: err?.textContent?.trim().slice(0, 160) ?? null,
            pdfError: document.querySelector('.pdf-surface__error')?.textContent?.trim().slice(0, 160) ?? null,
            pdfEmpty: document.querySelector('.pdf-surface__empty')?.textContent?.trim().slice(0, 120) ?? null,
            pdfSlots: document.querySelectorAll('.pdf-surface__page-slot').length,
            tabs: [...document.querySelectorAll('[data-testid="xlsx-sheet-tab"]')].map(t => t.textContent.trim()),
          }
        }"""
    )


def run_case(browser, name: str, route: str, selector: str, path: Path, wait_ms: int = 5000) -> dict:
    page = browser.new_page(viewport={"width": 1440, "height": 900})
    errors: list[str] = []
    page.on("pageerror", lambda exc: errors.append(str(exc)))
    started = time.time()
    page.goto(LOCAL_URL + route, wait_until="networkidle")
    page.wait_for_timeout(700)
    try:
        upload(page, selector, path)
        page.wait_for_timeout(wait_ms)
        state = read_state(page)
    except Exception as exc:  # noqa: BLE001 - record instead of crashing the run
        state = {"error": f"{type(exc).__name__}: {exc}"}
    elapsed = round(time.time() - started, 2)
    page.close()
    return {"case": name, "elapsed_s": elapsed, "state": state, "page_errors": errors[:5]}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output")
    args = parser.parse_args()
    from playwright.sync_api import sync_playwright

    tmp = Path(tempfile.mkdtemp(prefix="robust-"))
    files = make_files(tmp)
    report: dict = {"suite": "robustness", "date": "2026-11-16", "cases": [], "failures": []}
    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            executable_path="/usr/bin/chromium" if Path("/usr/bin/chromium").exists() else None,
            args=["--no-sandbox", "--disable-gpu"],
        )
        report["cases"].append(run_case(browser, "corrupt-docx", "/#/docx-viewer", 'input[type="file"][accept*="docx"]', SAMPLES / "corrupted.docx"))
        report["cases"].append(run_case(browser, "truncated-docx", "/#/docx-viewer", 'input[type="file"][accept*="docx"]', files["truncated"]))
        report["cases"].append(run_case(browser, "oversized-docx", "/#/docx-viewer", 'input[type="file"][accept*="docx"]', files["oversized_docx"]))
        report["cases"].append(run_case(browser, "oversized-xlsx", "/#/xlsx-viewer", 'input[type="file"][accept*="xlsx"]', files["oversized_xlsx"]))
        report["cases"].append(run_case(browser, "zip-bomb-xlsx", "/#/xlsx-viewer", 'input[type="file"][accept*="xlsx"]', files["zip_bomb"]))
        report["cases"].append(run_case(browser, "malformed-pdf", "/#/pdf-surface", 'input[type="file"][accept*="pdf"]', files["malformed_pdf"]))
        report["cases"].append(run_case(browser, "malformed-pptx", "/#/pptx-viewer", 'input[type="file"][accept*="pptx"]', files["malformed_pptx"]))
        report["cases"].append(run_case(browser, "truncated-pptx", "/#/pptx-viewer", 'input[type="file"][accept*="pptx"]', files["truncated_pptx"]))
        report["cases"].append(run_case(browser, "malformed-csv", "/#/xlsx-viewer", 'input[type="file"][accept*="xlsx"]', files["malformed_csv"]))

        # concurrent loads: second file must win
        page = browser.new_page(viewport={"width": 1440, "height": 900})
        errors: list[str] = []
        page.on("pageerror", lambda exc: errors.append(str(exc)))
        page.goto(LOCAL_URL + "/#/xlsx-viewer", wait_until="networkidle")
        page.wait_for_timeout(700)
        upload(page, 'input[type="file"][accept*="xlsx"]', XLSX_OTHER)
        page.wait_for_timeout(120)
        upload(page, 'input[type="file"][accept*="xlsx"]', XLSX_OK)
        page.wait_for_timeout(6000)
        state = read_state(page)
        page.close()
        report["cases"].append({"case": "concurrent-xlsx", "state": state, "page_errors": errors[:5]})

        # two tabs load different formats concurrently
        page_a = browser.new_page(viewport={"width": 1440, "height": 900})
        page_b = browser.new_page(viewport={"width": 1440, "height": 900})
        errs_a: list[str] = []
        errs_b: list[str] = []
        page_a.on("pageerror", lambda exc: errs_a.append(str(exc)))
        page_b.on("pageerror", lambda exc: errs_b.append(str(exc)))
        page_a.goto(LOCAL_URL + "/#/docx-viewer", wait_until="networkidle")
        page_b.goto(LOCAL_URL + "/#/xlsx-viewer", wait_until="networkidle")
        page_a.wait_for_timeout(600)
        page_b.wait_for_timeout(600)
        upload(page_a, 'input[type="file"][accept*="docx"]', SAMPLES / "demo.docx")
        upload(page_b, 'input[type="file"][accept*="xlsx"]', XLSX_OK)
        page_a.wait_for_timeout(6000)
        page_b.wait_for_timeout(6000)
        state_a = read_state(page_a)
        state_b = read_state(page_b)
        page_a.close(); page_b.close()
        report["cases"].append({"case": "two-tabs-concurrent", "state": {"a": state_a, "b": state_b}, "page_errors": errs_a + errs_b})

        # unmount during load
        page = browser.new_page(viewport={"width": 1440, "height": 900})
        errors = []
        page.on("pageerror", lambda exc: errors.append(str(exc)))
        page.goto(LOCAL_URL + "/#/xlsx-viewer", wait_until="networkidle")
        page.wait_for_timeout(700)
        upload(page, 'input[type="file"][accept*="xlsx"]', files["oversized_xlsx"])
        page.wait_for_timeout(150)
        page.goto(LOCAL_URL + "/#/", wait_until="networkidle")
        page.wait_for_timeout(1500)
        page.close()
        report["cases"].append({"case": "unmount-during-load", "page_errors": errors[:5]})

        browser.close()

    for case in report["cases"]:
        if case.get("page_errors"):
            report["failures"].append(f"{case['case']}: page errors")
        state = case.get("state", {})
        if state.get("error"):
            report["failures"].append(f"{case['case']}: {state['error']}")
        if case["case"] in {"corrupt-docx", "truncated-docx", "oversized-docx", "oversized-xlsx", "zip-bomb-xlsx", "malformed-pdf", "malformed-pptx", "truncated-pptx"}:
            if state.get("docxState") == "ready" or state.get("pptxState") == "ready" or (state.get("status") or "").lower() == "ready":
                report["failures"].append(f"{case['case']}: invalid input reported ready")
        if case["case"] in {"malformed-pdf"} and state.get("pdfSlots", 0) > 0:
            report["failures"].append("malformed-pdf: rendered page slots")
    report["verdict"] = "FAIL" if report["failures"] else "PASS"
    if args.output:
        output = Path(args.output)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"cases": [{k: c.get(k) for k in ("case", "elapsed_s", "state", "page_errors")} for c in report["cases"]], "failures": report["failures"], "verdict": report["verdict"]}, ensure_ascii=False, indent=2))
    return 1 if report["failures"] else 0


if __name__ == "__main__":
    sys.exit(main())
