from __future__ import annotations

import json
import os
from pathlib import Path
import signal
import socket
import subprocess
import time
import traceback
import urllib.request

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

from browser_evidence import BrowserEvidence


ROOT = Path(__file__).resolve().parents[2]
OUTPUT = Path(
    os.environ.get(
        "BLACKBOX_EVIDENCE_DIR",
        ROOT / "output" / "acceptance" / "p2-runtime-security",
    )
).resolve()
PORT = int(os.environ.get("CI_SECURITY_PREVIEW_PORT", "4176"))
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
            return playwright.chromium.launch(headless=True, executable_path=str(candidate)), str(candidate)
        except PlaywrightError as error:
            failures.append(str(error))
    raise RuntimeError("no Chromium executable could launch: " + " | ".join(failures))


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


CASES = [
    {
        "id": "SEC-RUNTIME-PDF-JAVASCRIPT",
        "route": "pdf-viewer",
        "input": "javascript:void(window.__p2SecurityMarker='changed')",
        "inputTestId": "pdf-url-input",
        "applyTestId": "pdf-apply-url",
        "errorTestId": "pdf-page-load-error",
        "diagnosticsTestId": "pdf-diagnostics",
        "sampleTestId": "pdf-sample-select",
        "recovery": "sample.pdf",
    },
    {
        "id": "SEC-RUNTIME-XLSX-JAVASCRIPT",
        "route": "xlsx-viewer",
        "input": "javascript:void(window.__p2SecurityMarker='changed')",
        "inputTestId": "xlsx-url-input",
        "applyTestId": "xlsx-apply-url",
        "errorTestId": "load-error",
        "diagnosticsTestId": "xlsx-diagnostics",
        "sampleTestId": "xlsx-sample-select",
        "recovery": "sales-table.xlsx",
    },
    {
        "id": "SEC-RUNTIME-PDF-CROSS-ORIGIN",
        "route": "pdf-viewer",
        "input": "https://blocked.invalid/sample.pdf?token=p2-secret",
        "inputTestId": "pdf-url-input",
        "applyTestId": "pdf-apply-url",
        "errorTestId": "pdf-page-load-error",
        "diagnosticsTestId": "pdf-diagnostics",
        "sampleTestId": "pdf-sample-select",
        "recovery": "sample.pdf",
    },
    {
        "id": "SEC-RUNTIME-XLSX-CROSS-ORIGIN",
        "route": "xlsx-viewer",
        "input": "https://blocked.invalid/book.xlsx?token=p2-secret",
        "inputTestId": "xlsx-url-input",
        "applyTestId": "xlsx-apply-url",
        "errorTestId": "load-error",
        "diagnosticsTestId": "xlsx-diagnostics",
        "sampleTestId": "xlsx-sample-select",
        "recovery": "sales-table.xlsx",
    },
]


def run_attempt(browser, case: dict[str, str], attempt: int) -> dict[str, object]:
    attempt_dir = OUTPUT / case["id"] / f"attempt-{attempt}"
    attempt_dir.mkdir(parents=True, exist_ok=True)
    context = browser.new_context(viewport={"width": 1440, "height": 900})
    context.add_init_script("window.__p2SecurityMarker = 'unchanged'")
    context.tracing.start(screenshots=True, snapshots=True, sources=True)
    page = context.new_page()
    evidence = BrowserEvidence(page)
    requests: list[str] = []
    page.on("request", lambda request: requests.append(request.url))
    try:
        page.goto(f"{BASE_URL}/#/{case['route']}", wait_until="networkidle")
        page.wait_for_function(
            "() => document.querySelector('[data-testid=\"page-status\"]')?.dataset.state === 'ready'",
            timeout=30_000,
        )
        page.get_by_test_id(case["inputTestId"]).fill(case["input"])
        page.get_by_test_id(case["applyTestId"]).click()
        page.wait_for_function(
            "() => document.querySelector('[data-testid=\"page-status\"]')?.dataset.state === 'error'",
            timeout=15_000,
        )
        error = page.get_by_test_id(case["errorTestId"])
        assert error.get_attribute("data-error-code") == "SOURCE_NOT_ALLOWED", error.inner_text()
        diagnostics = page.get_by_test_id(case["diagnosticsTestId"]).inner_text()
        assert "p2-secret" not in diagnostics, diagnostics
        assert "load-start" in diagnostics, diagnostics
        assert "load-error" in diagnostics, diagnostics
        task_prefix = "pdf-controller-" if case["route"] == "pdf-viewer" else "runtime-"
        assert task_prefix in diagnostics, diagnostics
        assert page.evaluate("window.__p2SecurityMarker") == "unchanged"
        assert not any("blocked.invalid" in url for url in requests), requests

        page.get_by_test_id(case["sampleTestId"]).select_option(case["recovery"])
        page.wait_for_function(
            "name => document.querySelector('[data-testid=\"page-status\"]')?.dataset.state === 'ready' && document.querySelector('[data-testid=\"loaded-file\"]')?.textContent?.includes(name)",
            arg=case["recovery"],
            timeout=30_000,
        )
        evidence.assert_clean()
        assert not evidence.events["downloads"], evidence.events["downloads"]
        page.screenshot(path=str(attempt_dir / "recovered.png"), full_page=True)
        result: dict[str, object] = {
            "id": case["id"],
            "attempt": attempt,
            "status": "PASS",
            "errorCode": "SOURCE_NOT_ALLOWED",
            "requestCount": len(requests),
            "forbiddenRequestCount": sum("blocked.invalid" in url for url in requests),
            "marker": "unchanged",
            "diagnostics": diagnostics,
            "recovery": case["recovery"],
        }
    except Exception as error:
        try:
            page.screenshot(path=str(attempt_dir / "failure.png"), full_page=True)
        except Exception:
            pass
        result = {
            "id": case["id"],
            "attempt": attempt,
            "status": "FAIL",
            "error": repr(error),
            "traceback": traceback.format_exc(),
            "requests": requests,
        }
    finally:
        evidence.save(attempt_dir)
        context.tracing.stop(path=attempt_dir / "trace.zip")
        context.close()
    (attempt_dir / "result.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return result


def main() -> int:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    if port_in_use():
        raise RuntimeError(f"port {PORT} is already in use")
    preview_log = (OUTPUT / "preview.log").open("wb")
    preview = subprocess.Popen(
        ["pnpm", "--filter", "demo", "preview", "--host", "127.0.0.1", "--port", str(PORT)],
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
            for case in CASES:
                attempts = [run_attempt(browser, case, 1)]
                if attempts[0]["status"] == "FAIL":
                    attempts.append(run_attempt(browser, case, 2))
                status = "PASS" if attempts[0]["status"] == "PASS" else (
                    "FLAKY" if attempts[-1]["status"] == "PASS" else "FAIL"
                )
                results.append({"id": case["id"], "status": status, "attempts": attempts})
            browser.close()
        overall = "PASS" if all(item["status"] == "PASS" for item in results) else "FAIL"
        summary = {
            "suite": "BB-SEC-URL",
            "mode": "formal preview",
            "result": overall,
            "browser": browser_version,
            "browserSource": browser_source,
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
                {"suite": "BB-SEC-URL", "result": "BLOCKED", "error": repr(error), "traceback": traceback.format_exc()},
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        print(f"BLOCKED: {error}")
        raise SystemExit(2)
