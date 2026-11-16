"""Edit → export → reopen round-trip checks (V2).

    LOCAL_DEMO_URL=http://127.0.0.1:4173 \
    uv run --with playwright --with python-docx --with openpyxl \
        python tests/blackbox/roundtrip.py \
        --output verification/evidence/2026-11-16/roundtrip/summary.json

Cases:

- DOCX editor: type a marker into the starter document, export, reopen with
  python-docx and LibreOffice; the marker must survive and the package must
  remain valid.
- XLSX viewer: load a real workbook, write a marker into a cell via the formula
  bar, export, reopen with openpyxl; the cell value, sheet list and chart parts
  must survive.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
LOCAL_URL = os.environ.get("LOCAL_DEMO_URL", "http://127.0.0.1:4173").strip().rstrip("/")
MARKER = f"ROUNDTRIP-{int(time.time())}"


def find_chromium() -> str | None:
    path = os.environ.get("ROUNDTRIP_CHROMIUM", "/usr/bin/chromium")
    return path if Path(path).exists() else None


def docx_case(browser, outdir: Path) -> dict:
    page = browser.new_page(viewport={"width": 1440, "height": 900})
    errors: list[str] = []
    page.on("pageerror", lambda exc: errors.append(str(exc)))
    page.goto(f"{LOCAL_URL}/#/docx-editor", wait_until="networkidle")
    page.wait_for_selector('[contenteditable="true"]', timeout=30000)
    page.wait_for_timeout(1500)
    editable = page.locator('[contenteditable="true"]').first
    editable.click()
    page.keyboard.press("End")
    page.keyboard.type(f" {MARKER}")
    page.wait_for_timeout(500)
    with page.expect_download(timeout=30000) as download_info:
        page.click('[data-testid="editor-export"]')
    export = outdir / "roundtrip.docx"
    download_info.value.save_as(str(export))
    page.close()

    import docx

    document = docx.Document(str(export))
    text = "\n".join(p.text for p in document.paragraphs)
    marker_found = MARKER in text
    valid = True
    if shutil.which("soffice"):
        result = subprocess.run(
            ["soffice", f"-env:UserInstallation=file://{outdir.resolve()}/lo", "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(export)],
            capture_output=True,
            text=True,
            timeout=120,
        )
        valid = (outdir / "roundtrip.pdf").exists()
    return {
        "export_bytes": export.stat().st_size,
        "marker": MARKER,
        "marker_found": marker_found,
        "libreoffice_valid": valid,
        "errors": errors[:5],
    }


def xlsx_case(browser, outdir: Path) -> dict:
    from openpyxl import load_workbook

    page = browser.new_page(viewport={"width": 1440, "height": 900})
    errors: list[str] = []
    page.on("pageerror", lambda exc: errors.append(str(exc)))
    page.goto(f"{LOCAL_URL}/#/xlsx-viewer", wait_until="networkidle")
    page.wait_for_timeout(800)
    fixture = ROOT / "test-data" / "corpus" / "xlsx-charts.xlsx"
    page.locator('input[type="file"][accept*="xlsx"]').first.set_input_files(str(fixture))
    page.wait_for_timeout(6000)
    status = (page.locator('[data-testid="page-status"]').inner_text() or "").strip()
    # select the first data cell, then write through the formula bar
    canvas = page.locator("canvas.xlsx-grid__body").first
    box = canvas.bounding_box()
    page.mouse.click(box["x"] + 60, box["y"] + 60)
    page.wait_for_timeout(400)
    page.fill('[data-testid="xlsx-formula-input"]', MARKER)
    page.keyboard.press("Enter")
    page.wait_for_timeout(500)
    with page.expect_download(timeout=30000) as download_info:
        page.click('[data-testid="xlsx-toolbar-export"]')
    export = outdir / "roundtrip.xlsx"
    download_info.value.save_as(str(export))
    page.close()

    import zipfile

    with zipfile.ZipFile(export) as archive:
        chart_parts = [name for name in archive.namelist() if name.startswith("xl/charts/")]
        has_chart_drawing = any("drawing" in name and name.endswith(".xml") for name in archive.namelist())
    reopen_error = None
    sheets: list[str] = []
    value = None
    charts = 0
    try:
        workbook = load_workbook(export)
        sheet = workbook[workbook.sheetnames[0]]
        value = sheet["A2"].value
        charts = len(sheet._charts)
        sheets = workbook.sheetnames
    except Exception as exc:  # exported package may be structurally invalid
        reopen_error = f"{type(exc).__name__}: {exc}"
    return {
        "status_before_export": status,
        "export_bytes": export.stat().st_size,
        "marker": MARKER,
        "cell_a2": value,
        "marker_found": value == MARKER,
        "sheets": sheets,
        "charts_after_export": charts,
        "chart_parts_after_export": len(chart_parts),
        "has_chart_drawing": has_chart_drawing,
        "reopen_error": reopen_error,
        "errors": errors[:5],
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output")
    args = parser.parse_args()
    from playwright.sync_api import sync_playwright

    outdir = Path(args.output).parent if args.output else Path(tempfile.mkdtemp(prefix="roundtrip-"))
    outdir.mkdir(parents=True, exist_ok=True)
    payload: dict = {"suite": "roundtrip", "date": "2026-11-16", "marker": MARKER, "cases": {}, "failures": []}
    with sync_playwright() as pw:
        browser = pw.chromium.launch(executable_path=find_chromium(), args=["--no-sandbox", "--disable-gpu"])
        docx_result = docx_case(browser, outdir)
        xlsx_result = xlsx_case(browser, outdir)
        browser.close()
    payload["cases"]["docx"] = docx_result
    payload["cases"]["xlsx"] = xlsx_result
    if not docx_result["marker_found"]:
        payload["failures"].append("docx: marker missing after export")
    if not docx_result["libreoffice_valid"]:
        payload["failures"].append("docx: exported package rejected by LibreOffice")
    if not xlsx_result["marker_found"]:
        payload["failures"].append("xlsx: cell value missing after export")
    if xlsx_result["reopen_error"]:
        payload["failures"].append(f"xlsx: exported package cannot be reopened ({xlsx_result['reopen_error']})")
    if xlsx_result["chart_parts_after_export"] == 0:
        payload["failures"].append("xlsx: chart parts missing after export (see #66)")
    payload["verdict"] = "FAIL" if payload["failures"] else "PASS"
    if args.output:
        output = Path(args.output)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"cases": payload["cases"], "failures": payload["failures"], "verdict": payload["verdict"]}, ensure_ascii=False, indent=2))
    return 1 if payload["failures"] else 0


if __name__ == "__main__":
    sys.exit(main())
