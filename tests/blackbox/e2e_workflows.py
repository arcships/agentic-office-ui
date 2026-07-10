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
        "bytes": len(data),
        "sha256": hashlib.sha256(data).hexdigest(),
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
    with page.expect_download(timeout=30_000) as download_info:
        page.locator('[data-testid="editor-export"]').click()
    saved = save_download(download_info.value, attempt_dir / "downloads")
    assert saved["bytes"] > 0, saved
    assert str(saved["fileName"]).endswith("-edited.docx"), saved
    page.wait_for_function(
        "() => document.querySelector('[data-testid=\"editor-status\"]')?.textContent?.includes('Exported DOCX')",
        timeout=30_000,
    )
    page.screenshot(path=str(attempt_dir / "edited-exported.png"), full_page=True)
    return {"original": original, "edited": marker, "download": saved}


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
    with page.expect_download(timeout=30_000) as download_info:
        page.locator('[data-testid="xlsx-download-source"]').click()
    saved = save_download(download_info.value, attempt_dir / "downloads")
    assert saved["bytes"] > 0, saved
    assert str(saved["fileName"]).endswith(".xlsx"), saved
    page.screenshot(path=str(attempt_dir / "xlsx-recovered-downloaded.png"), full_page=True)
    return {
        "firstSheet": first_name,
        "secondSheet": second_name,
        "originalFormula": original_formula,
        "errorCode": error_code,
        "recoveredFile": "sales-table.xlsx",
        "download": saved,
    }


CASES: list[tuple[str, Callable[[Page, Path], dict[str, object]]]] = [
    ("E2E-DOCX-SWITCH-RECOVERY", docx_switch_and_recover),
    ("E2E-DOCX-EDIT-UNDO-EXPORT", docx_edit_undo_export),
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
