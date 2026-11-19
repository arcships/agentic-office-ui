from __future__ import annotations

import hashlib
import json
import os
import signal
import socket
import subprocess
import sys
import time
import traceback
import urllib.request
from pathlib import Path
from typing import Callable

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import Page, sync_playwright

from browser_evidence import BrowserEvidence
from docx import Document
from openpyxl import __version__ as openpyxl_version
from openpyxl import load_workbook


EXPECTED_OPENPYXL = "3.1.5"


ROOT = Path(__file__).resolve().parents[2]
OUTPUT = Path(
    os.environ.get(
        "BLACKBOX_EVIDENCE_DIR",
        ROOT / "output" / "acceptance" / "p1-e2e-workflows",
    )
).resolve()
PORT = int(os.environ.get("CI_PREVIEW_PORT", "4173"))
BASE_URL = f"http://127.0.0.1:{PORT}"


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


def navigate_ready(page: Page, route: str) -> None:
    page.goto(f"{BASE_URL}/#{route}", wait_until="networkidle")
    page.locator('[data-testid="page-status"]').wait_for(state="visible", timeout=30_000)
    page.wait_for_function(
        "() => document.querySelector('[data-testid=\"page-status\"]')?.dataset.state === 'ready'",
        timeout=30_000,
    )


def save_download(download, directory: Path) -> dict[str, object]:
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / download.suggested_filename
    download.save_as(path)
    data = path.read_bytes()
    return {
        "fileName": download.suggested_filename,
        "path": str(path),
        "bytes": len(data),
        "sha256": hashlib.sha256(data).hexdigest(),
    }


def read_editor_model(page: Page) -> dict[str, object]:
    raw = page.locator('[data-testid="editor-model-snapshot"]').text_content() or "{}"
    return json.loads(raw)


def wait_for_editor_model(page: Page, predicate, message: str, timeout_ms: int = 30_000):
    deadline = time.monotonic() + timeout_ms / 1000
    last: dict[str, object] | None = None
    while time.monotonic() < deadline:
        last = read_editor_model(page)
        if predicate(last):
            return last
        page.wait_for_timeout(50)
    raise AssertionError(f"timed out waiting for {message}; last model={last}")


def docx_model_projection(snapshot: dict[str, object]) -> dict[str, object]:
    """Compare text, tables and images without depending on run boundaries.

    The editor snapshot summarizes paragraphs, tables and image runs. A DOCX
    round trip may re-segment adjacent runs, so the bold/italic/underline text
    is compared as a string instead of comparing run arrays.
    """
    nodes: list[dict[str, object]] = []
    for node in snapshot.get("nodes", []):
        if not isinstance(node, dict):
            continue
        if node.get("type") == "paragraph":
            def styled(style_key: str) -> str:
                return "".join(
                    str(run.get("text", ""))
                    for run in node.get("runs", [])
                    if isinstance(run, dict)
                    and run.get("type") == "text"
                    and isinstance(run.get("style"), dict)
                    and run["style"].get(style_key) is True
                )

            images = [
                {
                    "widthPx": run.get("widthPx"),
                    "heightPx": run.get("heightPx"),
                    "floating": run.get("floating"),
                }
                for run in node.get("runs", [])
                if isinstance(run, dict) and run.get("type") == "image"
            ]
            nodes.append(
                {
                    "type": "paragraph",
                    "text": node.get("text", ""),
                    "bold": styled("bold"),
                    "italic": styled("italic"),
                    "underline": styled("underline"),
                    "images": images,
                }
            )
        elif node.get("type") == "table":
            style = node.get("style") or {}
            nodes.append(
                {
                    "type": "table",
                    "rows": [
                        [str(cell.get("text", "")) for cell in row.get("cells", [])]
                        for row in node.get("rows", [])
                    ],
                    "layout": style.get("layout"),
                    "columnWidthsTwips": style.get("columnWidthsTwips"),
                    "widthTwips": style.get("widthTwips"),
                }
            )
    return {
        "nodes": nodes,
        "tableCount": snapshot.get("tableCount"),
        "imageCount": snapshot.get("imageCount"),
    }


def docx_switch_and_recover(page: Page, attempt_dir: Path) -> dict[str, object]:
    navigate_ready(page, "/docx-viewer")
    sample = page.locator('[data-testid="docx-sample-select"]')

    sample.select_option("invoice-table.docx")
    page.wait_for_function(
        "() => document.querySelector('[data-testid=\"loaded-file\"]')?.textContent?.includes('invoice-table.docx') && document.querySelector('[data-testid=\"page-status\"]')?.dataset.state === 'ready'",
        timeout=30_000,
    )
    page.screenshot(path=str(attempt_dir / "invoice-ready.png"), full_page=True)

    sample.select_option("corrupted.docx")
    page.wait_for_function(
        "() => document.querySelector('[data-testid=\"page-status\"]')?.dataset.state === 'error'",
        timeout=30_000,
    )
    error_code = page.locator('[data-testid="docx-page-error"]').inner_text()
    assert "PARSE_FAILED" in error_code, error_code
    page.screenshot(path=str(attempt_dir / "corrupted-error.png"), full_page=True)

    sample.select_option("report-with-image.docx")
    page.wait_for_function(
        "() => document.querySelector('[data-testid=\"loaded-file\"]')?.textContent?.includes('report-with-image.docx') && document.querySelector('[data-testid=\"page-status\"]')?.dataset.state === 'ready'",
        timeout=30_000,
    )
    runtime_details = page.locator(
        '[data-testid="docx-viewer-page"] details.runtime-details'
    )
    if runtime_details.get_attribute("open") is None:
        runtime_details.locator("summary").click()
    diagnostics = page.locator('[data-testid="docx-diagnostics"]').inner_text()
    assert "worker-success" in diagnostics, diagnostics[-2_000:]
    page.screenshot(path=str(attempt_dir / "recovered-ready.png"), full_page=True)
    return {"error": error_code, "recoveredFile": "report-with-image.docx"}


def docx_edit_undo_export(page: Page, attempt_dir: Path) -> dict[str, object]:
    navigate_ready(page, "/docx-editor")
    paragraph = page.locator('[data-testid="editor-paragraph"]').nth(2)
    original = paragraph.inner_text()
    marker = "P1 E2E edited paragraph"
    paragraph.fill(marker)
    page.locator("h2").click()
    page.locator('[data-testid="editor-undo"]').wait_for(state="visible")
    assert not page.locator('[data-testid="editor-undo"]').is_disabled()
    assert paragraph.inner_text() == marker

    page.locator('[data-testid="editor-undo"]').click()
    page.wait_for_function(
        "expected => document.querySelectorAll('[data-testid=\"editor-paragraph\"]')[2]?.textContent === expected",
        arg=original,
    )
    assert paragraph.inner_text() == original

    paragraph.fill(marker)
    page.locator("h2").click()
    edited_model = wait_for_editor_model(
        page,
        lambda model: any(
            node.get("type") == "paragraph" and node.get("text") == marker
            for node in model.get("nodes", [])
            if isinstance(node, dict)
        ),
        "the edited paragraph to reach the public editor model",
    )
    with page.expect_download(timeout=30_000) as download_info:
        page.locator('[data-testid="editor-export"]').click()
    saved = save_download(download_info.value, attempt_dir / "downloads")
    assert saved["bytes"] > 0, saved
    assert str(saved["fileName"]).endswith("-edited.docx"), saved
    page.wait_for_function(
        "() => document.querySelector('[data-testid=\"editor-status\"]')?.textContent?.includes('Exported DOCX')",
        timeout=30_000,
    )

    # Reopen the downloaded package with a real parser, not a byte-length check.
    exported_path = Path(str(saved["path"]))
    document = Document(str(exported_path))
    paragraph_texts = [item.text for item in document.paragraphs]
    assert marker in paragraph_texts, paragraph_texts[-5:]
    assert len(document.tables) == 1, len(document.tables)
    assert document.tables[0].rows[0].cells[0].text == "Launch readiness"
    assert len(document.inline_shapes) == 0, len(document.inline_shapes)

    # Re-import the exported file through the public import path and diff the
    # text, tables and images of the before/after models.
    with page.expect_file_chooser(timeout=30_000) as chooser_info:
        page.locator('[data-testid="editor-import"]').click()
    chooser_info.value.set_files(str(exported_path))
    page.wait_for_function(
        "name => document.querySelector('[data-testid=\"loaded-file\"]')?.textContent?.includes(name)",
        arg=str(saved["fileName"]),
        timeout=30_000,
    )
    reimported_model = wait_for_editor_model(
        page,
        lambda model: any(
            node.get("type") == "paragraph" and node.get("text") == marker
            for node in model.get("nodes", [])
            if isinstance(node, dict)
        ),
        "the exported file to be re-imported through the editor",
    )
    before_projection = docx_model_projection(edited_model)
    after_projection = docx_model_projection(reimported_model)
    assert before_projection == after_projection, {
        "before": before_projection,
        "after": after_projection,
    }

    page.screenshot(path=str(attempt_dir / "edited-exported.png"), full_page=True)
    return {
        "original": original,
        "edited": marker,
        "download": saved,
        "pythonDocx": {
            "paragraphs": len(document.paragraphs),
            "tables": len(document.tables),
            "inlineShapes": len(document.inline_shapes),
            "markerParagraph": marker,
        },
        "reimported": after_projection,
    }


def docx_reference_roundtrip(page: Page, attempt_dir: Path) -> dict[str, object]:
    page.goto(f"{BASE_URL}/#/docx-surface", wait_until="networkidle")
    page.locator('[data-testid="docx-reference-selection-mode"]').wait_for(
        state="visible", timeout=30_000
    )
    page.locator('[data-docx-page-index="0"]').wait_for(
        state="visible", timeout=30_000
    )
    page.locator('[data-testid="docx-reference-selection-mode"]').select_option(
        "object"
    )
    page.locator('[data-docx-page-index="0"]').click(position={"x": 120, "y": 120})
    page.wait_for_function(
        """() => document.querySelector('[data-testid="docx-reference-panel"] pre')?.textContent?.includes('"kind": "page"')""",
        timeout=10_000,
    )
    confirm_text = page.locator(
        '[data-testid="docx-reference-panel"] pre'
    ).nth(0).inner_text()
    confirm_event = json.loads(confirm_text)
    assert confirm_event["reference"]["document"]["documentId"] == "demo-docx-document"
    assert confirm_event["reference"]["source"] == "native"

    page.locator('[data-testid="docx-reference-simulate-update"]').click()
    page.wait_for_function(
        """() => document.querySelectorAll('[data-testid="docx-reference-panel"] pre')[1]?.textContent?.includes('"status": "unsupported"')""",
        timeout=30_000,
    )
    resolve_text = page.locator(
        '[data-testid="docx-reference-panel"] pre'
    ).nth(1).inner_text()
    resolve_event = json.loads(resolve_text)
    assert resolve_event["result"]["reasonCode"] == "docx.reflow-requires-reselection"
    page.screenshot(path=str(attempt_dir / "reference-reload-result.png"), full_page=True)
    return {
        "referenceKind": confirm_event["reference"]["kind"],
        "recognitionSource": confirm_event["reference"]["source"],
        "reloadStatus": resolve_event["result"]["status"],
        "reasonCode": resolve_event["result"]["reasonCode"],
    }


def xlsx_workflow(page: Page, attempt_dir: Path) -> dict[str, object]:
    navigate_ready(page, "/xlsx-viewer")
    page.wait_for_function(
        "() => document.querySelector('[data-testid=\"xlsx-grid\"]')?.getBoundingClientRect().height > 0",
        timeout=30_000,
    )
    tabs = page.locator('[data-testid="xlsx-sheet-tab"]')
    assert tabs.count() >= 2, f"expected at least two workbook tabs, got {tabs.count()}"
    first_name = tabs.nth(0).get_attribute("data-sheet-name")
    second_name = tabs.nth(1).get_attribute("data-sheet-name")
    tabs.nth(1).click()
    page.wait_for_function(
        "name => document.querySelector('[data-testid=\"xlsx-sheet-tab\"].xlsx-sheettabs__tab--active')?.dataset.sheetName === name",
        arg=second_name,
    )

    grid = page.locator('[data-testid="xlsx-grid"]')
    grid.click(position={"x": 88, "y": 36})
    formula = page.locator('[data-testid="xlsx-formula-input"]')
    page.wait_for_function(
        "() => !document.querySelector('[data-testid=\"xlsx-formula-input\"]')?.disabled",
        timeout=10_000,
    )
    original_formula = formula.input_value()
    formula.fill("=41+1")
    page.locator('[data-testid="page-status"]').click()
    page.wait_for_function(
        "() => document.querySelector('[data-testid=\"xlsx-formula-input\"]')?.value === '=41+1'",
        timeout=10_000,
    )
    undo = page.locator('[data-testid="xlsx-undo"]')
    assert not undo.is_disabled()
    undo.click()
    page.wait_for_function(
        "expected => document.querySelector('[data-testid=\"xlsx-formula-input\"]')?.value === expected",
        arg=original_formula,
        timeout=10_000,
    )

    sample = page.locator('[data-testid="xlsx-sample-select"]')
    sample.select_option("corrupted.xlsx")
    page.wait_for_function(
        "() => document.querySelector('[data-testid=\"page-status\"]')?.dataset.state === 'error'",
        timeout=30_000,
    )
    error_code = page.locator('[data-testid="load-error"]').get_attribute("data-error-code")
    assert error_code, "corrupted XLSX must expose a stable error code"
    page.screenshot(path=str(attempt_dir / "xlsx-corrupted-error.png"), full_page=True)

    sample.select_option("sales-table.xlsx")
    page.wait_for_function(
        "() => document.querySelector('[data-testid=\"loaded-file\"]')?.textContent?.includes('sales-table.xlsx') && document.querySelector('[data-testid=\"page-status\"]')?.dataset.state === 'ready'",
        timeout=30_000,
    )
    runtime_details = page.locator(
        '[data-testid="xlsx-viewer-page"] details.runtime-details'
    )
    if runtime_details.get_attribute("open") is None:
        runtime_details.locator("summary").click()
    with page.expect_download(timeout=30_000) as download_info:
        page.locator('[data-testid="xlsx-download-source"]').click()
    saved = save_download(download_info.value, attempt_dir / "downloads")
    assert saved["bytes"] > 0, saved
    assert str(saved["fileName"]).endswith(".xlsx"), saved
    # XLSX-013: the source download must be the verified fixture bytes, not a
    # re-requested URL or a re-serialized workbook.
    fixture = ROOT / "apps" / "demo" / "public" / "samples" / "sales-table.xlsx"
    assert hashlib.sha256(Path(str(saved["path"])).read_bytes()).hexdigest() == hashlib.sha256(
        fixture.read_bytes()
    ).hexdigest(), "the source download must match the fixture bytes"
    page.screenshot(path=str(attempt_dir / "xlsx-recovered-downloaded.png"), full_page=True)

    # XLSX-011/013: edit a real workbook, export it through the toolbar, and
    # reopen the downloaded file with the pinned independent parser.
    assert openpyxl_version == EXPECTED_OPENPYXL, openpyxl_version
    sample.select_option("charts-images.xlsx")
    page.wait_for_function(
        """() => document.querySelector('[data-testid="loaded-file"]')?.textContent?.includes('charts-images.xlsx')
          && document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'""",
        timeout=30_000,
    )
    grid_box = grid.bounding_box()
    assert grid_box, "the XLSX grid has no bounding box"
    page.mouse.click(grid_box["x"] + 48 + 20, grid_box["y"] + 24 + 12)
    page.wait_for_function(
        "() => !document.querySelector('[data-testid=\"xlsx-formula-input\"]')?.disabled",
        timeout=10_000,
    )
    marker = "E2E-XLSX-EXPORT"
    formula.fill(marker)
    page.locator('[data-testid="page-status"]').click()
    page.wait_for_function(
        "expected => document.querySelector('[data-testid=\"xlsx-formula-input\"]')?.value === expected",
        arg=marker,
        timeout=10_000,
    )
    with page.expect_download(timeout=30_000) as export_info:
        page.locator('[data-testid="xlsx-toolbar-export"]').click()
    exported = save_download(export_info.value, attempt_dir / "downloads")
    assert exported["bytes"] > 0, exported
    assert str(exported["fileName"]).endswith(".xlsx"), exported

    workbook = load_workbook(Path(str(exported["path"])), data_only=False)
    try:
        dashboard = workbook["Dashboard"]
        assert dashboard["A1"].value == marker, dashboard["A1"].value
        assert dashboard["A1"].font.bold is True, dashboard["A1"].font
        assert str(dashboard["A1"].fill.fgColor.rgb).endswith("1C3A5E"), dashboard["A1"].fill
        merged = {str(range_) for range_ in dashboard.merged_cells.ranges}
        assert "A7:C8" in merged, sorted(merged)
        assert dashboard.freeze_panes == "C3", dashboard.freeze_panes
        assert len(dashboard.data_validations.dataValidation) == 1
        assert workbook.sheetnames == ["Dashboard", "Revenue Chart"]
    finally:
        workbook.close()
    page.screenshot(path=str(attempt_dir / "xlsx-exported.png"), full_page=True)

    return {
        "firstSheet": first_name,
        "secondSheet": second_name,
        "originalFormula": original_formula,
        "errorCode": error_code,
        "recoveredFile": "sales-table.xlsx",
        "download": saved,
        "exported": exported,
        "exportEdit": marker,
        "merged": ["A7:C8"],
        "freezePanes": "C3",
        "openpyxl": openpyxl_version,
    }


CASES: list[tuple[str, Callable[[Page, Path], dict[str, object]]]] = [
    ("E2E-DOCX-SWITCH-RECOVERY", docx_switch_and_recover),
    ("E2E-DOCX-EDIT-UNDO-EXPORT", docx_edit_undo_export),
    ("E2E-DOCX-REFERENCE-ROUNDTRIP", docx_reference_roundtrip),
    ("E2E-XLSX-WORKFLOW", xlsx_workflow),
]


def run_attempt(browser, case_id: str, workflow, attempt: int) -> dict[str, object]:
    attempt_dir = OUTPUT / case_id / f"attempt-{attempt}"
    attempt_dir.mkdir(parents=True, exist_ok=True)
    context = browser.new_context(viewport={"width": 1440, "height": 900}, accept_downloads=True)
    context.tracing.start(screenshots=True, snapshots=True, sources=True)
    page = context.new_page()
    events = BrowserEvidence(page)
    try:
        details = workflow(page, attempt_dir)
        events.assert_clean()
        result: dict[str, object] = {
            "id": case_id,
            "attempt": attempt,
            "status": "PASS",
            "details": details,
            "events": events.events,
        }
    except Exception as error:
        try:
            page.screenshot(path=str(attempt_dir / "failure.png"), full_page=True)
        except Exception:
            pass
        result = {
            "id": case_id,
            "attempt": attempt,
            "status": "FAIL",
            "error": repr(error),
            "traceback": traceback.format_exc(),
            "events": events.events,
        }
    finally:
        events.save(attempt_dir)
        context.tracing.stop(path=attempt_dir / "trace.zip")
        context.close()
    (attempt_dir / "result.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return result


def stop_process(process: subprocess.Popen[bytes]) -> None:
    if process.poll() is not None:
        return
    try:
        os.killpg(process.pid, signal.SIGTERM)
        process.wait(timeout=5)
    except Exception:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except Exception:
            process.kill()


def main() -> int:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    if port_in_use():
        raise RuntimeError(f"port {PORT} is already in use")
    preview_log = (OUTPUT / "preview.log").open("wb")
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
    )
    try:
        wait_for_preview(preview)
        with sync_playwright() as playwright:
            browser, browser_source = launch_browser(playwright)
            browser_version = browser.version
            results = []
            for case_id, workflow in CASES:
                attempts = [run_attempt(browser, case_id, workflow, 1)]
                if attempts[0]["status"] == "FAIL":
                    attempts.append(run_attempt(browser, case_id, workflow, 2))
                if attempts[0]["status"] == "PASS":
                    status = "PASS"
                elif attempts[-1]["status"] == "PASS":
                    status = "FLAKY"
                else:
                    status = "FAIL"
                results.append({"id": case_id, "status": status, "attempts": attempts})
            browser.close()
        overall = "PASS" if all(item["status"] == "PASS" for item in results) else "FAIL"
        summary = {
            "suite": "P1-E2E-FORMAL-WORKFLOWS",
            "mode": "formal preview",
            "result": overall,
            "browser": browser_version,
            "browserSource": browser_source,
            "baseUrl": BASE_URL,
            "results": results,
        }
        (OUTPUT / "summary.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return 0 if overall == "PASS" else 1
    finally:
        stop_process(preview)
        preview_log.close()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        OUTPUT.mkdir(parents=True, exist_ok=True)
        (OUTPUT / "summary.json").write_text(
            json.dumps(
                {"suite": "P1-E2E-FORMAL-WORKFLOWS", "result": "BLOCKED", "error": str(error)},
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        print(f"BLOCKED: {error}", file=sys.stderr)
        raise SystemExit(2)
