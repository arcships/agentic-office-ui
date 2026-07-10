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
        ROOT / "output" / "acceptance" / "p2-config-isolation",
    )
).resolve()
PORT = int(os.environ.get("CI_CONFIG_PREVIEW_PORT", "4177"))
BASE_URL = f"http://127.0.0.1:{PORT}"

RESOURCE_PROBE = """
(() => {
  const state = { activeWorkers: 0, createdWorkers: 0, terminatedWorkers: 0 };
  Object.defineProperty(window, "__p2ConfigResources", { value: state });
  const NativeWorker = window.Worker;
  function TrackedWorker(...args) {
    const worker = new NativeWorker(...args);
    state.activeWorkers += 1;
    state.createdWorkers += 1;
    let terminated = false;
    const terminate = worker.terminate.bind(worker);
    worker.terminate = () => {
      if (!terminated) {
        terminated = true;
        state.activeWorkers -= 1;
        state.terminatedWorkers += 1;
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


def run_attempt(browser, attempt: int) -> dict[str, object]:
    attempt_dir = OUTPUT / f"attempt-{attempt}"
    attempt_dir.mkdir(parents=True, exist_ok=True)
    context = browser.new_context(viewport={"width": 1440, "height": 900})
    context.add_init_script(RESOURCE_PROBE)
    context.tracing.start(screenshots=True, snapshots=True, sources=True)
    page = context.new_page()
    evidence = BrowserEvidence(page)
    try:
        page.goto(f"{BASE_URL}/#/runtime-isolation", wait_until="networkidle")
        page.wait_for_function(
            "() => document.querySelector('[data-testid=\"page-status\"]')?.dataset.state === 'ready'",
            timeout=30_000,
        )
        diagnostics_a = json.loads(page.get_by_test_id("runtime-a-diagnostics").inner_text())
        diagnostics_b = json.loads(page.get_by_test_id("runtime-b-diagnostics").inner_text())
        runtime_ids_a = {item.get("runtimeId") for item in diagnostics_a if item.get("runtimeId")}
        runtime_ids_b = {item.get("runtimeId") for item in diagnostics_b if item.get("runtimeId")}
        assert len(runtime_ids_a) == 1, runtime_ids_a
        assert len(runtime_ids_b) == 1, runtime_ids_b
        assert runtime_ids_a.isdisjoint(runtime_ids_b), (runtime_ids_a, runtime_ids_b)
        assert all("runtime=alpha" in item.get("wasmUrl", "") for item in diagnostics_a if item.get("wasmUrl")), diagnostics_a
        assert all("runtime=beta" in item.get("wasmUrl", "") for item in diagnostics_b if item.get("wasmUrl")), diagnostics_b
        assert all(str(item.get("requestId", "")).startswith(next(iter(runtime_ids_a)) + ":") for item in diagnostics_a if item.get("requestId") and not str(item.get("requestId")).startswith("page:")), diagnostics_a
        assert all(str(item.get("requestId", "")).startswith(next(iter(runtime_ids_b)) + ":") for item in diagnostics_b if item.get("requestId") and not str(item.get("requestId")).startswith("page:")), diagnostics_b

        wasm_responses = [
            item for item in evidence.events["responses"]
            if ".wasm?runtime=" in str(item.get("url", ""))
        ]
        assert any("runtime=alpha" in str(item["url"]) and item["status"] == 200 for item in wasm_responses), wasm_responses
        assert any("runtime=beta" in str(item["url"]) and item["status"] == 200 for item in wasm_responses), wasm_responses
        assert all("application/wasm" in str(item["contentType"]) for item in wasm_responses), wasm_responses
        resources_ready = page.evaluate("() => ({ ...window.__p2ConfigResources })")
        assert resources_ready["createdWorkers"] >= 2, resources_ready
        assert resources_ready["activeWorkers"] == 0, resources_ready

        page.screenshot(path=str(attempt_dir / "two-runtimes-ready.png"), full_page=True)
        page.evaluate("location.hash = '#/'")
        page.wait_for_timeout(200)
        resources_unmounted = page.evaluate("() => ({ ...window.__p2ConfigResources })")
        assert resources_unmounted["activeWorkers"] == 0, resources_unmounted
        evidence.assert_clean()
        result: dict[str, object] = {
            "id": "CONFIG-TWO-RUNTIMES",
            "attempt": attempt,
            "status": "PASS",
            "runtimeIdsA": sorted(runtime_ids_a),
            "runtimeIdsB": sorted(runtime_ids_b),
            "wasmResponses": wasm_responses,
            "resourcesReady": resources_ready,
            "resourcesUnmounted": resources_unmounted,
        }
    except Exception as error:
        try:
            page.screenshot(path=str(attempt_dir / "failure.png"), full_page=True)
        except Exception:
            pass
        result = {
            "id": "CONFIG-TWO-RUNTIMES",
            "attempt": attempt,
            "status": "FAIL",
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
            attempts = [run_attempt(browser, 1)]
            if attempts[0]["status"] == "FAIL":
                attempts.append(run_attempt(browser, 2))
            browser.close()
        status = "PASS" if attempts[0]["status"] == "PASS" else (
            "FLAKY" if attempts[-1]["status"] == "PASS" else "FAIL"
        )
        summary = {
            "suite": "P2-CONFIG-01",
            "mode": "formal preview",
            "result": "PASS" if status == "PASS" else "FAIL",
            "status": status,
            "browser": browser_version,
            "browserSource": browser_source,
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
                {"suite": "P2-CONFIG-01", "result": "BLOCKED", "error": repr(error), "traceback": traceback.format_exc()},
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        raise SystemExit(2)
