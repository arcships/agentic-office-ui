from __future__ import annotations

import json
import os
from pathlib import Path
import shlex
import signal
import socket
import subprocess
import time
import traceback
import urllib.request
import zipfile

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

from browser_launch import launch_browser


ROOT = Path(__file__).resolve().parents[2]
OUTPUT = Path(
    os.environ.get(
        "BLACKBOX_EVIDENCE_DIR",
        ROOT / "output" / "acceptance" / "bb-docx-export-bytes",
    )
).resolve()
PORT = int(os.environ.get("CI_DOCX_EXPORT_BYTES_PREVIEW_PORT", "4187"))
BASE_URL = f"http://127.0.0.1:{PORT}"
EDITOR_URL = f"{BASE_URL}/#/docx-editor"
DOCX_MIME = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)


def write_json(path: Path, value: object) -> None:
    path.write_text(
        json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8"
    )


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


def read_docx_parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def run_assertions(page, output: Path) -> dict[str, object]:
    page.goto(EDITOR_URL, wait_until="networkidle")
    page.wait_for_function(
        """() => document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'
          && document.querySelector('[data-testid="editor-editable-state"]')?.dataset.state === 'editable'""",
        timeout=30_000,
    )

    page.get_by_test_id("editor-test-export-bytes").click()
    state = page.get_by_test_id("editor-export-bytes-state")
    state.wait_for(state="visible", timeout=30_000)
    page.wait_for_function(
        """() => document.querySelector('[data-testid="editor-export-bytes-state"]')?.dataset.state === 'ready'""",
        timeout=30_000,
    )
    bytes_state = json.loads(state.text_content() or "{}")
    if bytes_state.get("bytes", 0) <= 0:
        raise AssertionError(f"getDocxBytes returned no bytes: {bytes_state}")
    if bytes_state.get("signature") != [0x50, 0x4B, 0x03, 0x04]:
        raise AssertionError(
            f"getDocxBytes did not return an OOXML ZIP: {bytes_state}"
        )
    if bytes_state.get("blobType") != DOCX_MIME:
        raise AssertionError(
            f"getDocxBlob returned the wrong MIME type: {bytes_state}"
        )
    if bytes_state.get("blobSize") != bytes_state.get("bytes"):
        raise AssertionError(
            f"getDocxBlob size does not match getDocxBytes: {bytes_state}"
        )

    with page.expect_download(timeout=30_000) as download_info:
        page.get_by_test_id("editor-export").click()
    download = download_info.value
    if download.suggested_filename != "quarterly-planning-brief-edited.docx":
        raise AssertionError(
            f"unexpected download name: {download.suggested_filename}"
        )
    download_path = output / download.suggested_filename
    download.save_as(download_path)
    parts = read_docx_parts(download_path)
    document_xml = parts.get("word/document.xml", b"").decode("utf-8", "replace")
    if "Quarterly Planning Brief" not in document_xml:
        raise AssertionError("the downloaded DOCX does not contain the edited document")
    if "[Content_Types].xml" not in parts:
        raise AssertionError("the downloaded DOCX has no content-types part")

    return {
        "getDocxBytes": bytes_state,
        "download": {
            "fileName": download.suggested_filename,
            "bytes": download_path.stat().st_size,
            "parts": sorted(parts),
        },
    }


def run_attempt(browser, attempt: int) -> dict[str, object]:
    attempt_dir = OUTPUT / f"attempt-{attempt}"
    attempt_dir.mkdir(parents=True, exist_ok=True)
    context = browser.new_context(
        viewport={"width": 1440, "height": 900},
        device_scale_factor=1,
        locale="en",
        timezone_id="Asia/Shanghai",
        accept_downloads=True,
    )
    page_errors: list[str] = []
    console_errors: list[str] = []
    page = context.new_page()
    page.on("pageerror", lambda error: page_errors.append(str(error)))
    page.on("console", lambda message: console_errors.append(message.text)
            if message.type == "error" else None)
    result: dict[str, object] = {"attempt": attempt}
    try:
        evidence = run_assertions(page, attempt_dir)
        if page_errors:
            raise AssertionError(f"page errors: {page_errors}")
        if console_errors:
            raise AssertionError(f"console errors: {console_errors}")
        result.update({"status": "PASS", "evidence": evidence})
    except (AssertionError, PlaywrightError) as error:
        result.update(
            {
                "status": "FAIL",
                "error": str(error),
                "pageErrors": page_errors,
                "consoleErrors": console_errors,
                "traceback": traceback.format_exc(),
            }
        )
    finally:
        context.close()
    write_json(attempt_dir / "result.json", result)
    return result


def main() -> int:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    if port_in_use():
        raise RuntimeError(f"port {PORT} is already in use")
    preview_command_text = os.environ.get("DOCX_EXPORT_BYTES_PREVIEW_COMMAND")
    if preview_command_text:
        preview_command = shlex.split(preview_command_text.format(port=PORT))
    else:
        preview_command = [
            "pnpm",
            "--filter",
            "demo",
            "preview",
            "--host",
            "127.0.0.1",
            "--port",
            str(PORT),
        ]
    preview_log = (OUTPUT / "preview.log").open("wb")
    preview = subprocess.Popen(
        preview_command,
        cwd=ROOT,
        stdout=preview_log,
        stderr=subprocess.STDOUT,
        start_new_session=True,
    )
    try:
        wait_for_preview(preview)
        with sync_playwright() as playwright:
            browser = launch_browser(playwright.chromium, headless=True)
            browser_version = browser.version
            browser_source = browser.browser_type.name
            attempts = [run_attempt(browser, 1)]
            if attempts[0]["status"] == "FAIL":
                attempts.append(run_attempt(browser, 2))
            status = (
                "PASS"
                if attempts[0]["status"] == "PASS"
                else "FLAKY"
                if attempts[-1]["status"] == "PASS"
                else "FAIL"
            )
            browser.close()
        write_json(
            OUTPUT / "summary.json",
            {
                "suite": "BB-DOCX-EXPORT-BYTES-01",
                "mode": "formal preview",
                "result": status,
                "browser": browser_version,
                "browserSource": browser_source,
                "cases": [{"id": "DOCX-EXPORT-BYTES", "status": status, "attempts": attempts}],
            },
        )
        return 0 if status == "PASS" else 1
    finally:
        stop_process(preview)
        preview_log.close()


if __name__ == "__main__":
    raise SystemExit(main())
