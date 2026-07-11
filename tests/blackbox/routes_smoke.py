from __future__ import annotations

import json
import os
import signal
import socket
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

from browser_evidence import BrowserEvidence


ROOT = Path(__file__).resolve().parents[2]
OUTPUT = Path(
    os.environ.get(
        "BLACKBOX_EVIDENCE_DIR",
        ROOT / "output" / "acceptance" / "p1-ci-routes",
    )
).resolve()
PORT = int(os.environ.get("CI_PREVIEW_PORT", "4173"))
BASE_URL = f"http://127.0.0.1:{PORT}"
ROUTES = [
    ("CI-ROUTE-HOME", "/"),
    ("CI-ROUTE-DOCX-VIEWER", "/docx-viewer"),
    ("CI-ROUTE-DOCX-EDITOR", "/docx-editor"),
    ("CI-ROUTE-XLSX-VIEWER", "/xlsx-viewer"),
    ("CI-ROUTE-PPTX-VIEWER", "/pptx-viewer"),
    ("CI-ROUTE-RUNTIME-ISOLATION", "/runtime-isolation"),
]
READY_ROUTES = {"/docx-viewer", "/docx-editor", "/xlsx-viewer", "/runtime-isolation"}


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

    candidates = [
        Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"),
        Path("/usr/bin/google-chrome"),
        Path("/usr/bin/google-chrome-stable"),
    ]
    for candidate in candidates:
        if not candidate.exists():
            continue
        try:
            return (
                playwright.chromium.launch(
                    headless=True, executable_path=str(candidate)
                ),
                str(candidate),
            )
        except PlaywrightError as error:
            failures.append(str(error))
    raise RuntimeError("no Chromium executable could launch: " + " | ".join(failures))


def wait_ready(page, route: str) -> None:
    if route in READY_ROUTES:
        page.wait_for_function(
            """() => document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'""",
            timeout=30_000,
        )
    if route == "/xlsx-viewer":
        page.wait_for_function(
            "() => document.querySelector('.xlsx-grid')?.getBoundingClientRect().height > 0",
            timeout=10_000,
        )


def page_state(page) -> dict[str, object]:
    status = page.locator('[data-testid="page-status"]')
    grid = page.locator(".xlsx-grid")
    return {
        "bodyTextLength": len(page.locator("body").inner_text()),
        "headings": page.locator("h1, h2, h3").count(),
        "status": status.get_attribute("data-state") if status.count() else None,
        "gridHeight": (
            grid.evaluate("element => element.getBoundingClientRect().height")
            if grid.count()
            else None
        ),
    }


def run_attempt(browser, case_id: str, route: str, attempt: int) -> dict[str, object]:
    attempt_dir = OUTPUT / case_id / f"attempt-{attempt}"
    attempt_dir.mkdir(parents=True, exist_ok=True)
    context = browser.new_context(viewport={"width": 1440, "height": 900})
    page = context.new_page()
    events = BrowserEvidence(page)
    try:
        page.goto(f"{BASE_URL}/#{route}", wait_until="networkidle")
        wait_ready(page, route)
        direct = page_state(page)
        assert direct["bodyTextLength"] > 0 and direct["headings"] > 0, direct
        page.screenshot(path=str(attempt_dir / "direct.png"), full_page=True)

        page.reload(wait_until="networkidle")
        wait_ready(page, route)
        refreshed = page_state(page)
        assert refreshed["bodyTextLength"] > 0 and refreshed["headings"] > 0, refreshed
        page.screenshot(path=str(attempt_dir / "refresh.png"), full_page=True)

        if case_id == "CI-ROUTE-HOME":
            probe = os.environ.get("BLACKBOX_BROWSER_PROBE", "")
            if probe == "console":
                page.evaluate("console.error('P1 console fault probe')")
            elif probe == "vue-warning":
                page.evaluate("console.warn('[Vue warn]: P1 Vue warning fault probe')")
            elif probe == "pageerror":
                page.evaluate("setTimeout(() => { throw new Error('P1 page error fault probe') }, 0)")
                page.wait_for_timeout(100)
            elif probe == "requestfailed":
                page.evaluate("fetch('http://127.0.0.1:65534/p1-request-failure').catch(() => undefined)")
                page.wait_for_timeout(200)
            elif probe == "response404":
                page.route(
                    "**/__p1_response_404_probe__",
                    lambda intercepted: intercepted.fulfill(status=404, body="missing"),
                )
                page.evaluate("fetch('/__p1_response_404_probe__').catch(() => undefined)")
                page.wait_for_timeout(100)

        events.assert_clean()
        assert not events.events["downloads"], events.events["downloads"]
        result = {
            "id": case_id,
            "route": route,
            "attempt": attempt,
            "status": "PASS",
            "direct": direct,
            "refresh": refreshed,
            "events": events.events,
        }
    except Exception as error:
        page.screenshot(path=str(attempt_dir / "failure.png"), full_page=True)
        result = {
            "id": case_id,
            "route": route,
            "attempt": attempt,
            "status": "FAIL",
            "error": str(error),
            "events": events.events,
        }
    events.save(attempt_dir)
    (attempt_dir / "result.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    context.close()
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
            for case_id, route in ROUTES:
                attempts = [run_attempt(browser, case_id, route, 1)]
                if attempts[0]["status"] == "FAIL":
                    attempts.append(run_attempt(browser, case_id, route, 2))
                if attempts[0]["status"] == "PASS":
                    status = "PASS"
                elif attempts[-1]["status"] == "PASS":
                    status = "FLAKY"
                else:
                    status = "FAIL"
                results.append(
                    {"id": case_id, "route": route, "status": status, "attempts": attempts}
                )
            browser.close()

        overall = "PASS" if all(item["status"] == "PASS" for item in results) else "FAIL"
        summary = {
            "suite": "P1-CI-FORMAL-ROUTES",
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
                {"suite": "P1-CI-FORMAL-ROUTES", "result": "BLOCKED", "error": str(error)},
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        print(f"BLOCKED: {error}", file=sys.stderr)
        raise SystemExit(2)
