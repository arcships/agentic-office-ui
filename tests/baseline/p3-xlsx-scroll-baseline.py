from __future__ import annotations

"""Focused green regression for the large-grid scroll range.

The original failure evidence lives under output/acceptance. This source check
must pass and is also registered in the performance suite.
"""

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

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "tests" / "blackbox"))
from browser_evidence import BrowserEvidence  # noqa: E402


OUTPUT = Path(
    os.environ.get(
        "BLACKBOX_EVIDENCE_DIR",
        ROOT / "output" / "acceptance" / "p3-xlsx-scroll-regression",
    )
).resolve()
PORT = int(os.environ.get("CI_PERF_BASELINE_PORT", "4182"))
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
    failures = []
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


def wait_large_grid(page) -> None:
    page.wait_for_function(
        """() => {
          const status = document.querySelector('[data-testid="page-status"]');
          const loaded = document.querySelector('[data-testid="loaded-file"]');
          const grid = document.querySelector('[data-testid="xlsx-grid"]');
          const canvas = grid?.querySelector('.xlsx-grid__body');
          return status?.dataset.state === 'ready' &&
            loaded?.textContent?.includes('large-grid.xlsx') &&
            grid?.getBoundingClientRect().height > 0 &&
            canvas instanceof HTMLCanvasElement && canvas.width > 0;
        }""",
        timeout=45_000,
    )


def grid_snapshot(page) -> dict[str, float]:
    return page.locator('[data-testid="xlsx-grid"]').evaluate(
        """grid => ({
          clientWidth: grid.clientWidth,
          clientHeight: grid.clientHeight,
          scrollWidth: grid.scrollWidth,
          scrollHeight: grid.scrollHeight,
          maxHorizontal: Math.max(0, grid.scrollWidth - grid.clientWidth),
          maxVertical: Math.max(0, grid.scrollHeight - grid.clientHeight),
          scrollLeft: grid.scrollLeft,
          scrollTop: grid.scrollTop,
        })"""
    )


def run_attempt(browser, attempt: int) -> dict[str, object]:
    attempt_dir = OUTPUT / "PERF-004-SCROLL-RANGE" / f"attempt-{attempt}"
    attempt_dir.mkdir(parents=True, exist_ok=True)
    context = browser.new_context(
        viewport={"width": 1440, "height": 900},
        device_scale_factor=1,
        locale="zh-CN",
        timezone_id="Asia/Shanghai",
    )
    context.tracing.start(screenshots=True, snapshots=True, sources=True)
    page = context.new_page()
    evidence = BrowserEvidence(page)
    initial = None
    final = None
    try:
        page.goto(f"{BASE_URL}/#/xlsx-viewer", wait_until="networkidle")
        page.locator('[data-testid="xlsx-sample-select"]').select_option(
            "large-grid.xlsx"
        )
        wait_large_grid(page)
        initial = grid_snapshot(page)
        page.locator('[data-testid="xlsx-grid"]').evaluate(
            """grid => {
              grid.scrollTop = Math.max(0, grid.scrollHeight - grid.clientHeight);
              grid.scrollLeft = Math.max(0, grid.scrollWidth - grid.clientWidth);
              grid.dispatchEvent(new Event('scroll'));
            }"""
        )
        page.wait_for_timeout(100)
        final = grid_snapshot(page)
        page.screenshot(
            path=str(attempt_dir / "PERF-004-large-grid-scroll-range.png"),
            full_page=True,
        )
        evidence.assert_clean()
        assert initial["maxVertical"] > 0, initial
        assert initial["maxHorizontal"] > 0, initial
        assert final["scrollTop"] > 0, final
        assert final["scrollLeft"] > 0, final
        result: dict[str, object] = {
            "id": "PERF-004-SCROLL-RANGE",
            "attempt": attempt,
            "status": "PASS",
            "initial": initial,
            "final": final,
        }
    except Exception as error:
        try:
            page.screenshot(path=str(attempt_dir / "failure.png"), full_page=True)
        except Exception:
            pass
        result = {
            "id": "PERF-004-SCROLL-RANGE",
            "attempt": attempt,
            "status": "FAIL",
            "initial": initial,
            "final": final,
            "error": repr(error),
            "traceback": traceback.format_exc(),
        }
    finally:
        evidence.save(attempt_dir)
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
            attempts = [run_attempt(browser, 1)]
            if attempts[0]["status"] == "FAIL":
                attempts.append(run_attempt(browser, 2))
            browser.close()
        if attempts[0]["status"] == "PASS":
            status = "PASS"
        elif attempts[-1]["status"] == "PASS":
            status = "FLAKY"
        else:
            status = "FAIL"
        summary = {
            "suite": "BB-PERF-XLSX",
            "case": "PERF-004-SCROLL-RANGE",
            "result": status,
            "mode": "formal preview",
            "browser": browser_version,
            "browserSource": browser_source,
            "baseUrl": BASE_URL,
            "viewport": {"width": 1440, "height": 900},
            "deviceScaleFactor": 1,
            "attempts": attempts,
        }
        (OUTPUT / "summary.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return 0 if status == "PASS" else 1
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
                {
                    "suite": "BB-PERF-XLSX",
                    "case": "PERF-004-SCROLL-RANGE",
                    "result": "BLOCKED",
                    "error": repr(error),
                    "traceback": traceback.format_exc(),
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        print(f"BLOCKED: {error}", file=sys.stderr)
        raise SystemExit(2)
