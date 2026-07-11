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
        ROOT / "output" / "acceptance" / "bb-worker",
    )
).resolve()
PORT = int(os.environ.get("CI_WORKER_PREVIEW_PORT", "4183"))
BASE_URL = f"http://127.0.0.1:{PORT}"

WORKER_PROBE = r"""
(() => {
  const state = { active: 0, created: 0, terminated: 0 };
  Object.defineProperty(window, "__workerProbe", { value: state });
  const NativeWorker = window.Worker;
  function TrackedWorker(...args) {
    const worker = new NativeWorker(...args);
    state.active += 1;
    state.created += 1;
    let terminated = false;
    const terminate = worker.terminate.bind(worker);
    worker.terminate = () => {
      if (!terminated) {
        terminated = true;
        state.active -= 1;
        state.terminated += 1;
      }
      return terminate();
    };
    return worker;
  }
  TrackedWorker.prototype = NativeWorker.prototype;
  Object.setPrototypeOf(TrackedWorker, NativeWorker);
  window.Worker = TrackedWorker;
})();
"""


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
            return playwright.chromium.launch(
                headless=True, executable_path=str(candidate)
            ), str(candidate)
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


def run_attempt(browser, case_id: str, button: str, expected_state: str, attempt: int) -> dict[str, object]:
    attempt_dir = OUTPUT / case_id / f"attempt-{attempt}"
    attempt_dir.mkdir(parents=True, exist_ok=True)
    context = browser.new_context(
        viewport={"width": 1440, "height": 900},
        device_scale_factor=1,
        locale="zh-CN",
        timezone_id="Asia/Shanghai",
    )
    context.add_init_script(WORKER_PROBE)
    context.tracing.start(screenshots=True, snapshots=True, sources=True)
    page = context.new_page()
    evidence = BrowserEvidence(page)
    try:
        page.goto(f"{BASE_URL}/#/runtime-worker", wait_until="networkidle")
        page.wait_for_function(
            "() => document.querySelector('[data-testid=page-status]')?.dataset.state === 'ready'"
        )
        page.get_by_test_id(button).click()
        page.wait_for_function(
            "() => document.querySelector('[data-testid=worker-case-status]')?.dataset.state === 'running'"
        )
        page.get_by_test_id("worker-heartbeat").click()
        page.wait_for_function(
            "expected => document.querySelector('[data-testid=worker-case-status]')?.dataset.state === expected",
            arg=expected_state,
            timeout=30_000,
        )

        error_code = page.get_by_test_id("worker-error-code").inner_text().strip()
        error_phase = page.get_by_test_id("worker-error-phase").inner_text().strip()
        terminated = int(page.get_by_test_id("worker-terminated-count").inner_text())
        recovered_sheets = int(page.get_by_test_id("worker-recovered-sheets").inner_text())
        heartbeat = int(page.get_by_test_id("worker-heartbeat-count").inner_text())
        diagnostics = json.loads(page.get_by_test_id("worker-diagnostics").inner_text())
        resources = page.evaluate("() => ({ ...window.__workerProbe })")

        if case_id == "STRESS-005-TIMEOUT-RECOVERY":
            assert error_code == "TIMEOUT", error_code
            assert error_phase == "parse", error_phase
            assert recovered_sheets > 0, recovered_sheets
            assert resources["created"] >= 2, resources
            assert any(
                item.get("type") == "worker-error" and item.get("code") == "TIMEOUT"
                for item in diagnostics
            ), diagnostics
        else:
            assert error_code == "ABORTED", error_code
            assert expected_state == "cancelled"
            assert any(item.get("type") == "worker-cancelled" for item in diagnostics), diagnostics
        assert heartbeat == 1, heartbeat
        assert terminated >= 1, terminated
        assert resources["active"] == 0, resources
        assert resources["terminated"] == resources["created"], resources

        page.screenshot(path=str(attempt_dir / "final.png"), full_page=True)
        evidence.assert_clean()
        result: dict[str, object] = {
            "id": case_id,
            "attempt": attempt,
            "status": "PASS",
            "state": expected_state,
            "errorCode": error_code,
            "errorPhase": error_phase,
            "terminatedCount": terminated,
            "recoveredSheets": recovered_sheets,
            "heartbeat": heartbeat,
            "resources": resources,
            "diagnostics": diagnostics,
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
            "events": evidence.events,
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
        [
            "pnpm", "--filter", "demo", "preview",
            "--host", "127.0.0.1", "--port", str(PORT),
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
            for case_id, button, state in (
                ("STRESS-005-TIMEOUT-RECOVERY", "worker-timeout-run", "recovered"),
                ("WORKER-CANCEL", "worker-cancel-run", "cancelled"),
            ):
                attempts = [run_attempt(browser, case_id, button, state, 1)]
                if attempts[0]["status"] == "FAIL":
                    attempts.append(run_attempt(browser, case_id, button, state, 2))
                status = "PASS" if attempts[0]["status"] == "PASS" else (
                    "FLAKY" if attempts[-1]["status"] == "PASS" else "FAIL"
                )
                results.append({"id": case_id, "status": status, "attempts": attempts})
            browser.close()
        overall = "PASS" if all(item["status"] == "PASS" for item in results) else "FAIL"
        summary = {
            "suite": "BB-STRESS",
            "subset": ["STRESS-005", "WORKER-CANCEL"],
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
                {
                    "suite": "BB-STRESS",
                    "subset": ["STRESS-005", "WORKER-CANCEL"],
                    "result": "BLOCKED",
                    "error": repr(error),
                    "traceback": traceback.format_exc(),
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        raise SystemExit(2)
