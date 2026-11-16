"""Deep roundtrip: structured content and styles (V2 extension).

    LOCAL_DEMO_URL=http://127.0.0.1:4173 \
    uv run --with playwright --with python-docx --with openpyxl \
        python tests/blackbox/roundtrip_deep.py \
        --output verification/evidence/2026-11-16/roundtrip/deep.json

DOCX: insert table + image + formatted range via the demo's acceptance hooks,
export, reopen with python-docx and LibreOffice; assert table/image/bold survive.

XLSX: use a chartless workbook, toggle bold + fill on A1, export, reopen with
openpyxl; assert style survived. (Chart export is covered separately by #66.)
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
LOCAL_URL = os.environ.get("LOCAL_DEMO_URL", "http://127.0.0.1:4173").strip().rstrip("/")
XLSX = ROOT / "test-data" / "corpus" / "xlsx-validation.xlsx"


def docx_case(browser, outdir: Path) -> dict:
    page = browser.new_page(viewport={"width": 1440, "height": 900})
    page.goto(f"{LOCAL_URL}/#/docx-editor", wait_until="networkidle")
    page.wait_for_selector('[contenteditable="true"]', timeout=30000)
    page.wait_for_timeout(1200)
    page.evaluate("() => { const d = document.querySelector('[data-testid=\"editor-acceptance-panel\"]'); if (d) d.open = true; }")
    page.wait_for_timeout(300)
    page.click('[data-testid="editor-test-insert-table"]')
    page.click('[data-testid="editor-test-insert-image"]')
    page.click('[data-testid="editor-test-format-range"]')
    page.wait_for_timeout(2500)
    before = page.evaluate(
        """() => ({ images: document.querySelectorAll('img').length,
                   tables: document.querySelectorAll('table').length })"""
    )
    with page.expect_download(timeout=30000) as download:
        page.click('[data-testid="editor-export"]')
    export = outdir / "deep.docx"
    download.value.save_as(str(export))
    page.close()

    import docx
    import zipfile

    with zipfile.ZipFile(export) as archive:
        media = [name for name in archive.namelist() if name.startswith("word/media/")]
        document_xml = archive.read("word/document.xml").decode("utf8", "ignore")
    document = docx.Document(str(export))
    bold_runs = sum(1 for paragraph in document.paragraphs for run in paragraph.runs if run.bold)
    valid = None
    if shutil.which("soffice"):
        subprocess.run(
            ["soffice", f"-env:UserInstallation=file://{(outdir / 'lo').resolve()}", "--headless", "--convert-to", "pdf", "--outdir", str(outdir.resolve()), str(export)],
            capture_output=True,
            text=True,
            timeout=120,
        )
        valid = (outdir / "deep.pdf").exists()
    return {
        "before_export": before,
        "tables": len(document.tables),
        "media_parts": len(media),
        "drawings": document_xml.count("<w:drawing"),
        "bold_runs": bold_runs,
        "libreoffice_valid": valid,
        "export_bytes": export.stat().st_size,
    }


def xlsx_case(browser, outdir: Path) -> dict:
    from openpyxl import load_workbook

    page = browser.new_page(viewport={"width": 1440, "height": 900})
    page.goto(f"{LOCAL_URL}/#/xlsx-viewer", wait_until="networkidle")
    page.wait_for_timeout(800)
    page.locator('input[type="file"][accept*="xlsx"]').first.set_input_files(str(XLSX))
    page.wait_for_timeout(5000)
    canvas = page.locator("canvas.xlsx-grid__body").first
    box = canvas.bounding_box()
    page.mouse.click(box["x"] + 60, box["y"] + 60)
    page.wait_for_timeout(400)
    page.click('[aria-label="Bold"]')
    page.evaluate(
        """() => { const input = document.querySelector('input[aria-label="Fill color"]');
          input.value = '#ff0000'; input.dispatchEvent(new Event('input', {bubbles:true}));
          input.dispatchEvent(new Event('change', {bubbles:true})); }"""
    )
    page.wait_for_timeout(600)
    before = page.evaluate(
        """() => ({ boldPressed: document.querySelector('[aria-label="Bold"]')?.getAttribute('aria-pressed'),
                   fill: document.querySelector('input[aria-label="Fill color"]')?.value,
                   address: document.querySelector('[data-testid="xlsx-name-box"]')?.value })"""
    )
    with page.expect_download(timeout=30000) as download:
        page.click('[data-testid="xlsx-toolbar-export"]')
    export = outdir / "deep.xlsx"
    download.value.save_as(str(export))
    page.close()

    reopen_error = None
    bold = None
    fill = None
    try:
        workbook = load_workbook(export)
        sheet = workbook[workbook.sheetnames[0]]
        address = before.get("address") or "A1"
        cell = sheet[address]
        bold = bool(cell.font.bold)
        fill = cell.fill.fgColor.rgb if cell.fill and cell.fill.fgColor else None
    except Exception as exc:  # noqa: BLE001
        reopen_error = f"{type(exc).__name__}: {exc}"
    return {"before_export": before, "bold": bold, "fill": fill, "reopen_error": reopen_error, "export_bytes": export.stat().st_size}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output")
    args = parser.parse_args()
    from playwright.sync_api import sync_playwright

    outdir = Path(args.output).parent if args.output else Path(tempfile.mkdtemp(prefix="roundtrip-deep-"))
    outdir.mkdir(parents=True, exist_ok=True)
    report: dict = {"suite": "roundtrip-deep", "date": "2026-11-16", "cases": {}, "failures": []}
    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            executable_path="/usr/bin/chromium" if Path("/usr/bin/chromium").exists() else None,
            args=["--no-sandbox", "--disable-gpu"],
        )
        report["cases"]["docx"] = docx_case(browser, outdir)
        report["cases"]["xlsx"] = xlsx_case(browser, outdir)
        browser.close()
    docx_result = report["cases"]["docx"]
    if docx_result["tables"] < 1:
        report["failures"].append("docx: table missing after export")
    if docx_result["media_parts"] < 1 or docx_result["drawings"] < 1:
        report["failures"].append("docx: image missing after export")
    if not docx_result["bold_runs"]:
        report["failures"].append("docx: formatted range lost bold")
    if docx_result["libreoffice_valid"] is False:
        report["failures"].append("docx: exported package rejected by LibreOffice")
    xlsx_result = report["cases"]["xlsx"]
    if xlsx_result["reopen_error"]:
        report["failures"].append(f"xlsx: reopen failed ({xlsx_result['reopen_error']})")
    if not xlsx_result["bold"]:
        report["failures"].append("xlsx: bold lost after export")
    report["verdict"] = "FAIL" if report["failures"] else "PASS"
    if args.output:
        output = Path(args.output)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 1 if report["failures"] else 0


if __name__ == "__main__":
    sys.exit(main())
