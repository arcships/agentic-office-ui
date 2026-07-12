from __future__ import annotations

import json
import os
from pathlib import Path
import signal
import socket
import subprocess
import sys
import time
import traceback
import urllib.request
from typing import Callable

from playwright.sync_api import Download
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import Page, sync_playwright

from browser_evidence import BrowserEvidence


ROOT = Path(__file__).resolve().parents[2]
OUTPUT = Path(
    os.environ.get(
        "BLACKBOX_EVIDENCE_DIR",
        ROOT / "output" / "acceptance" / "p1-race",
    )
).resolve()
DELAY_MS = 1500

RESOURCE_PROBE = r"""
(() => {
  const state = {
    activeWorkers: 0,
    createdWorkers: 0,
    terminatedWorkers: 0,
    activeObjectUrls: 0,
    createdObjectUrls: 0,
    revokedObjectUrls: 0,
  };
  Object.defineProperty(window, "__p1ResourceProbe", { value: state });

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

  const activeUrls = new Set();
  const createObjectURL = URL.createObjectURL.bind(URL);
  const revokeObjectURL = URL.revokeObjectURL.bind(URL);
  URL.createObjectURL = (value) => {
    const url = createObjectURL(value);
    activeUrls.add(url);
    state.activeObjectUrls = activeUrls.size;
    state.createdObjectUrls += 1;
    return url;
  };
  URL.revokeObjectURL = (url) => {
    if (activeUrls.delete(url)) state.revokedObjectUrls += 1;
    state.activeObjectUrls = activeUrls.size;
    return revokeObjectURL(url);
  };
})();
"""


def available_port(preferred: int) -> int:
    with socket.socket() as sock:
        if sock.connect_ex(("127.0.0.1", preferred)) != 0:
            return preferred
    with socket.socket() as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def wait_for_server(base_url: str, process: subprocess.Popen[bytes]) -> None:
    deadline = time.monotonic() + 20
    while time.monotonic() < deadline:
        if process.poll() is not None:
            raise RuntimeError(f"fault server exited early with {process.returncode}")
        try:
            with urllib.request.urlopen(f"{base_url}/__health", timeout=1) as response:
                if response.status == 200:
                    return
        except Exception:
            time.sleep(0.1)
    raise TimeoutError(f"fault server did not become ready at {base_url}")


def server_events(base_url: str) -> list[dict[str, object]]:
    with urllib.request.urlopen(f"{base_url}/__events", timeout=2) as response:
        return json.loads(response.read().decode())


def event_count(base_url: str, event: str, file_name: str) -> int:
    return sum(
        1
        for item in server_events(base_url)
        if item.get("event") == event and item.get("file") == file_name
    )


def wait_for_event_count(
    base_url: str,
    event: str,
    file_name: str,
    previous: int,
    timeout_seconds: float = 10,
) -> None:
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        if event_count(base_url, event, file_name) > previous:
            return
        time.sleep(0.05)
    raise TimeoutError(f"did not observe {event} for {file_name}")


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


def resource_state(page: Page) -> dict[str, int]:
    return page.evaluate("() => ({ ...window.__p1ResourceProbe })")


def wait_ready(page: Page, file_name: str, timeout: int = 30_000) -> None:
    page.wait_for_function(
        "name => document.querySelector('[data-testid=\"page-status\"]')?.dataset.state === 'ready' && document.querySelector('[data-testid=\"loaded-file\"]')?.textContent?.includes(name)",
        arg=file_name,
        timeout=timeout,
    )


def open_details(page: Page, test_id: str) -> None:
    details = page.locator(f'[data-testid="{test_id}"]')
    details.wait_for(state="visible")
    if details.get_attribute("open") is None:
        details.locator("summary").click()


def assert_race_evidence(
    evidence: BrowserEvidence, delayed_file: str, case_id: str
) -> dict[str, object]:
    violations = evidence.violations()
    switch_files = {
        "RACE-DOCX-SWITCH-20": {
            "demo.docx",
            "corrupted.docx",
            "report-with-image.docx",
            "invoice-table.docx",
        },
        "RACE-XLSX-SWITCH-20": {
            "financial-model.xlsx",
            "corrupted.xlsx",
            "charts-images.xlsx",
            "sales-table.xlsx",
        },
    }.get(case_id, set())
    expected_aborts = [
        item
        for item in violations["requestFailures"]
        if "ERR_ABORTED" in str(item.get("failure", ""))
        and (
            delayed_file in str(item.get("url", ""))
            or any(
                f"/samples/{file_name}" in str(item.get("url", ""))
                for file_name in switch_files
            )
        )
    ]
    unexpected_requests = [
        item for item in violations["requestFailures"] if item not in expected_aborts
    ]
    unexpected = {
        "console": violations["console"],
        "pageErrors": violations["pageErrors"],
        "requestFailures": unexpected_requests,
        "responses": violations["responses"],
    }
    assert not any(unexpected.values()), unexpected
    return {"expectedCancelledRequests": expected_aborts, "unexpected": unexpected}


def docx_snapshot(page: Page) -> dict[str, object]:
    pages = page.locator('[data-testid="docx-page"]')
    first_text = pages.first.inner_text() if pages.count() else ""
    return {
        "status": page.locator('[data-testid="page-status"]').get_attribute("data-state"),
        "loadedFile": page.locator('[data-testid="loaded-file"]').inner_text(),
        "pageCount": pages.count(),
        "firstPage": first_text[:300],
        "diagnostics": page.locator('[data-testid="docx-diagnostics"]').inner_text(),
        "resources": resource_state(page),
    }


def xlsx_snapshot(page: Page) -> dict[str, object]:
    return {
        "status": page.locator('[data-testid="page-status"]').get_attribute("data-state"),
        "loadedFile": page.locator('[data-testid="loaded-file"]').inner_text(),
        "diagnostics": page.locator('[data-testid="xlsx-diagnostics"]').inner_text(),
        "resources": resource_state(page),
    }


def pdf_snapshot(page: Page) -> dict[str, object]:
    return {
        "status": page.locator('[data-testid="page-status"]').get_attribute("data-state"),
        "loadedFile": page.locator('[data-testid="loaded-file"]').inner_text(),
        "diagnostics": page.locator('[data-testid="pdf-diagnostics"]').inner_text(),
        "resources": resource_state(page),
    }


def public_diagnostics(page: Page, test_id: str) -> list[dict[str, str]]:
    return page.locator(
        f'[data-testid="{test_id}"] [data-diagnostic-type]'
    ).evaluate_all(
        """entries => entries.map((entry) => ({
          type: entry.getAttribute("data-diagnostic-type") || "",
          text: (entry.textContent || "").trim(),
        }))"""
    )


def diagnostic_task_id(entry: dict[str, str], prefix: str) -> str:
    for part in entry["text"].split(" · "):
        candidate = part.strip()
        if candidate.startswith(prefix):
            return candidate
    raise AssertionError(f"diagnostic does not expose a {prefix} taskId: {entry}")


def save_public_events(attempt_dir: Path, file_name: str, payload: object) -> None:
    (attempt_dir / file_name).write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def save_download(download: Download, attempt_dir: Path) -> dict[str, object]:
    download_dir = attempt_dir / "downloads"
    download_dir.mkdir(parents=True, exist_ok=True)
    file_name = Path(download.suggested_filename).name
    path = download_dir / file_name
    download.save_as(path)
    return {
        "fileName": file_name,
        "bytes": path.stat().st_size,
    }


def prepare_deferred_xlsx(page: Page, base_url: str) -> tuple[str, list[dict[str, str]]]:
    page.goto(f"{base_url}/#/xlsx-viewer", wait_until="domcontentloaded")
    wait_ready(page, "financial-model.xlsx")
    open_details(page, "xlsx-runtime-details")

    page.locator('[data-testid="xlsx-read-only"]').check()
    wait_ready(page, "financial-model.xlsx")
    page.locator('[data-testid="xlsx-sample-select"]').select_option("sales-table.xlsx")
    wait_ready(page, "sales-table.xlsx")

    page.locator('[data-testid="xlsx-defer-loading-above-bytes"]').fill("1")
    page.wait_for_function(
        """() => {
          const button = document.querySelector('[data-testid="xlsx-continue-deferred"]');
          const loaded = document.querySelector('[data-testid="loaded-file"]')?.textContent || "";
          const deferred = document.querySelectorAll(
            '[data-testid="xlsx-diagnostics"] [data-diagnostic-type="load-deferred"]'
          );
          return loaded.includes("sales-table.xlsx") && button && !button.disabled && deferred.length > 0;
        }""",
        timeout=30_000,
    )
    diagnostics = public_diagnostics(page, "xlsx-diagnostics")
    deferred = [entry for entry in diagnostics if entry["type"] == "load-deferred"]
    assert deferred, diagnostics
    return diagnostic_task_id(deferred[-1], "xlsx-runtime-"), diagnostics


def docx_slow_latest(page: Page, base_url: str, attempt_dir: Path) -> dict[str, object]:
    slow = "demo.docx"
    before = event_count(base_url, "delay-start", slow)
    page.goto(f"{base_url}/#/docx-viewer", wait_until="domcontentloaded")
    page.locator('[data-testid="docx-sample-select"]').wait_for(state="visible")
    open_details(page, "docx-runtime-details")
    wait_for_event_count(base_url, "delay-start", slow, before)
    finished = event_count(base_url, "delay-finished", slow)
    page.locator('[data-testid="docx-sample-select"]').select_option("invoice-table.docx")
    wait_ready(page, "invoice-table.docx")
    wait_for_event_count(base_url, "delay-finished", slow, finished)
    page.wait_for_timeout(250)
    final = docx_snapshot(page)
    assert "INVOICE" in str(final["firstPage"]).upper(), final
    assert "page-load-cancelled" in str(final["diagnostics"]), final
    assert final["resources"]["activeWorkers"] == 0, final
    page.screenshot(path=str(attempt_dir / "docx-slow-latest.png"), full_page=True)
    return final


def docx_switch_20(page: Page, base_url: str, attempt_dir: Path) -> dict[str, object]:
    page.goto(f"{base_url}/#/docx-viewer", wait_until="domcontentloaded")
    selector = page.locator('[data-testid="docx-sample-select"]')
    selector.wait_for(state="visible")
    sequence = ["demo.docx", "corrupted.docx", "report-with-image.docx", "invoice-table.docx"] * 5
    sequence[-1] = "invoice-table.docx"
    for file_name in sequence:
        selector.select_option(file_name)
    wait_ready(page, "invoice-table.docx")
    page.wait_for_timeout(DELAY_MS + 400)
    final = docx_snapshot(page)
    assert "INVOICE" in str(final["firstPage"]).upper(), final
    assert final["resources"]["activeWorkers"] == 0, final
    page.screenshot(path=str(attempt_dir / "docx-switch-20.png"), full_page=True)
    return {"switchCount": len(sequence), "final": final}


def docx_unmount(page: Page, base_url: str, attempt_dir: Path) -> dict[str, object]:
    slow = "demo.docx"
    before = event_count(base_url, "delay-start", slow)
    page.goto(f"{base_url}/#/docx-viewer", wait_until="domcontentloaded")
    page.locator('[data-testid="docx-sample-select"]').wait_for(state="visible")
    wait_for_event_count(base_url, "delay-start", slow, before)
    page.evaluate("location.hash = '#/'")
    page.wait_for_timeout(DELAY_MS + 400)
    resources = resource_state(page)
    assert page.locator('[data-testid="docx-viewer"]').count() == 0
    assert resources["activeWorkers"] == 0, resources
    assert resources["activeObjectUrls"] == 0, resources
    page.screenshot(path=str(attempt_dir / "docx-unmounted.png"), full_page=True)
    return {"resources": resources}


def docx_double_export(page: Page, base_url: str, attempt_dir: Path) -> dict[str, object]:
    downloads: list[Download] = []
    page.on("download", lambda download: downloads.append(download))
    page.goto(f"{base_url}/#/docx-editor", wait_until="domcontentloaded")
    wait_ready(page, "quarterly-planning-brief.docx")

    dispatch = page.locator('[data-testid="editor-export"]').evaluate(
        """button => {
          const startedAt = performance.now();
          button.click();
          button.click();
          return { startedAt, finishedAt: performance.now(), clickCount: 2 };
        }"""
    )
    page.wait_for_function(
        """() => document.querySelector('[data-testid="editor-status"]')
          ?.textContent?.includes("Exported DOCX")""",
        timeout=30_000,
    )
    deadline = time.monotonic() + 30
    while not downloads and time.monotonic() < deadline:
        page.wait_for_timeout(50)
    assert downloads, "two synchronous export calls produced no download"
    page.wait_for_timeout(250)
    assert len(downloads) == 1, [item.suggested_filename for item in downloads]

    saved = save_download(downloads[0], attempt_dir)
    assert saved["bytes"] > 0, saved
    page.wait_for_function(
        "() => window.__p1ResourceProbe.activeObjectUrls === 0",
        timeout=5_000,
    )
    resources = resource_state(page)
    assert resources["createdObjectUrls"] == 1, resources
    assert resources["revokedObjectUrls"] == 1, resources
    assert resources["activeObjectUrls"] == 0, resources
    status = page.locator('[data-testid="editor-status"]').inner_text()
    events = {
        "dispatch": dispatch,
        "downloads": [{"suggestedFilename": item.suggested_filename} for item in downloads],
        "saved": saved,
        "status": status,
        "resources": resources,
    }
    save_public_events(attempt_dir, "docx-double-export-events.json", events)
    page.screenshot(path=str(attempt_dir / "docx-double-export.png"), full_page=True)
    return events


def docx_export_then_route(page: Page, base_url: str, attempt_dir: Path) -> dict[str, object]:
    downloads: list[Download] = []
    page.on("download", lambda download: downloads.append(download))
    page.goto(f"{base_url}/#/docx-editor", wait_until="domcontentloaded")
    wait_ready(page, "quarterly-planning-brief.docx")

    dispatch = page.locator('[data-testid="editor-export"]').evaluate(
        """button => {
          const startedAt = performance.now();
          button.click();
          location.hash = "#/";
          return { startedAt, routeRequestedAt: performance.now() };
        }"""
    )
    page.wait_for_function(
        """() => location.hash === "#/" &&
          !document.querySelector('[data-testid="docx-editor-page"]')""",
        timeout=30_000,
    )
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1_200)
    page.wait_for_function(
        """() => window.__p1ResourceProbe.activeObjectUrls === 0 &&
          !document.querySelector('[data-testid="editor-status"]')""",
        timeout=5_000,
    )

    resources = resource_state(page)
    assert page.url.endswith("/#/"), page.url
    assert page.locator('[data-testid="docx-editor-page"]').count() == 0
    assert page.locator('[data-testid="editor-status"]').count() == 0
    assert len(downloads) <= 1, [item.suggested_filename for item in downloads]
    assert resources["activeObjectUrls"] == 0, resources
    assert resources["createdObjectUrls"] == resources["revokedObjectUrls"], resources
    events = {
        "dispatch": dispatch,
        "route": page.url,
        "downloads": [{"suggestedFilename": item.suggested_filename} for item in downloads],
        "resources": resources,
        "publicInvariant": (
            "the home route remains active, editor state stays absent, and every created "
            "object URL is revoked; zero or one download is accepted because serialization "
            "may finish before the public route change is observed"
        ),
    }
    save_public_events(attempt_dir, "docx-export-route-events.json", events)
    page.screenshot(path=str(attempt_dir / "docx-export-route.png"), full_page=True)
    return events


def xlsx_slow_latest(page: Page, base_url: str, attempt_dir: Path) -> dict[str, object]:
    slow = "financial-model.xlsx"
    before = event_count(base_url, "delay-start", slow)
    page.goto(f"{base_url}/#/xlsx-viewer", wait_until="domcontentloaded")
    page.locator('[data-testid="xlsx-sample-select"]').wait_for(state="visible")
    open_details(page, "xlsx-runtime-details")
    wait_for_event_count(base_url, "delay-start", slow, before)
    finished = event_count(base_url, "delay-finished", slow)
    page.locator('[data-testid="xlsx-sample-select"]').select_option("sales-table.xlsx")
    wait_ready(page, "sales-table.xlsx")
    wait_for_event_count(base_url, "delay-finished", slow, finished)
    page.wait_for_timeout(250)
    final = xlsx_snapshot(page)
    assert final["status"] == "ready", final
    assert "sales-table.xlsx" in str(final["loadedFile"]), final
    assert "xlsx-runtime-" in str(final["diagnostics"]), final
    page.screenshot(path=str(attempt_dir / "xlsx-slow-latest.png"), full_page=True)
    return final


def xlsx_switch_20(page: Page, base_url: str, attempt_dir: Path) -> dict[str, object]:
    page.goto(f"{base_url}/#/xlsx-viewer", wait_until="domcontentloaded")
    selector = page.locator('[data-testid="xlsx-sample-select"]')
    selector.wait_for(state="visible")
    sequence = ["financial-model.xlsx", "corrupted.xlsx", "charts-images.xlsx", "sales-table.xlsx"] * 5
    sequence[-1] = "sales-table.xlsx"
    for file_name in sequence:
        selector.select_option(file_name)
    wait_ready(page, "sales-table.xlsx")
    page.wait_for_timeout(DELAY_MS + 400)
    final = xlsx_snapshot(page)
    assert final["status"] == "ready", final
    assert "sales-table.xlsx" in str(final["loadedFile"]), final
    page.screenshot(path=str(attempt_dir / "xlsx-switch-20.png"), full_page=True)
    return {"switchCount": len(sequence), "final": final}


def xlsx_worker_unmount(page: Page, base_url: str, attempt_dir: Path) -> dict[str, object]:
    page.goto(f"{base_url}/#/xlsx-viewer", wait_until="domcontentloaded")
    page.locator('[data-testid="xlsx-sample-select"]').wait_for(state="visible")
    page.locator('[data-testid="xlsx-read-only"]').check()
    page.locator('[data-testid="xlsx-sample-select"]').select_option("sales-table.xlsx")
    wait_ready(page, "sales-table.xlsx")
    page.wait_for_function("() => window.__p1ResourceProbe.activeWorkers > 0")
    mounted = resource_state(page)
    page.evaluate("location.hash = '#/'")
    page.wait_for_timeout(300)
    unmounted = resource_state(page)
    assert page.locator('[data-testid="xlsx-viewer"]').count() == 0
    assert mounted["activeWorkers"] > 0, mounted
    assert unmounted["activeWorkers"] == 0, unmounted
    assert unmounted["activeObjectUrls"] == 0, unmounted
    page.screenshot(path=str(attempt_dir / "xlsx-worker-unmounted.png"), full_page=True)
    return {"mounted": mounted, "unmounted": unmounted}


def xlsx_deferred_double_continue(page: Page, base_url: str, attempt_dir: Path) -> dict[str, object]:
    task_id, deferred_diagnostics = prepare_deferred_xlsx(page, base_url)
    before = resource_state(page)
    dispatch = page.locator('[data-testid="xlsx-continue-deferred"]').evaluate(
        """button => {
          const startedAt = performance.now();
          button.click();
          button.click();
          return { startedAt, finishedAt: performance.now(), clickCount: 2 };
        }"""
    )
    wait_ready(page, "sales-table.xlsx")
    page.wait_for_function(
        """taskId => {
          const entries = Array.from(document.querySelectorAll(
            '[data-testid="xlsx-diagnostics"] [data-diagnostic-type]'
          )).filter((entry) => entry.textContent?.includes(taskId));
          const types = entries.map((entry) => entry.getAttribute("data-diagnostic-type"));
          return types.filter((type) => type === "load-resumed").length === 1 &&
            types.filter((type) => type === "load-success").length === 1;
        }""",
        arg=task_id,
        timeout=30_000,
    )

    diagnostics = public_diagnostics(page, "xlsx-diagnostics")
    task_events = [entry for entry in diagnostics if task_id in entry["text"]]
    resumed = [entry for entry in task_events if entry["type"] == "load-resumed"]
    succeeded = [entry for entry in task_events if entry["type"] == "load-success"]
    assert len(resumed) == 1, task_events
    assert len(succeeded) == 1, task_events
    resumed_task_ids = {diagnostic_task_id(entry, "xlsx-runtime-") for entry in resumed}
    success_task_ids = {diagnostic_task_id(entry, "xlsx-runtime-") for entry in succeeded}
    assert resumed_task_ids == success_task_ids == {task_id}, task_events

    final = xlsx_snapshot(page)
    events = {
        "taskId": task_id,
        "dispatch": dispatch,
        "before": before,
        "deferredDiagnostics": deferred_diagnostics,
        "taskEvents": task_events,
        "final": final,
    }
    save_public_events(attempt_dir, "xlsx-deferred-double-events.json", events)
    page.screenshot(path=str(attempt_dir / "xlsx-deferred-double.png"), full_page=True)
    return events


def xlsx_deferred_worker_unmount(page: Page, base_url: str, attempt_dir: Path) -> dict[str, object]:
    task_id, deferred_diagnostics = prepare_deferred_xlsx(page, base_url)
    before = resource_state(page)
    page.locator('[data-testid="xlsx-continue-deferred"]').click()
    page.wait_for_function(
        """taskId => window.__p1ResourceProbe.activeWorkers > 0 &&
          Array.from(document.querySelectorAll(
            '[data-testid="xlsx-diagnostics"] [data-diagnostic-type="load-resumed"]'
          )).some((entry) => entry.textContent?.includes(taskId))""",
        arg=task_id,
        timeout=30_000,
    )
    mounted = resource_state(page)
    diagnostics_before_route = public_diagnostics(page, "xlsx-diagnostics")

    page.evaluate("location.hash = '#/docx-editor'")
    wait_ready(page, "quarterly-planning-brief.docx")
    page.wait_for_function(
        """() => window.__p1ResourceProbe.activeWorkers === 0 &&
          window.__p1ResourceProbe.activeObjectUrls === 0""",
        timeout=5_000,
    )
    page.wait_for_timeout(250)
    unmounted = resource_state(page)
    assert page.locator('[data-testid="xlsx-viewer-page"]').count() == 0
    assert mounted["activeWorkers"] > 0, mounted
    assert unmounted["activeWorkers"] == 0, unmounted
    assert unmounted["createdWorkers"] == unmounted["terminatedWorkers"], unmounted
    assert unmounted["activeObjectUrls"] == 0, unmounted

    events = {
        "taskId": task_id,
        "before": before,
        "mounted": mounted,
        "unmounted": unmounted,
        "deferredDiagnostics": deferred_diagnostics,
        "diagnosticsBeforeRoute": diagnostics_before_route,
        "route": page.url,
    }
    save_public_events(attempt_dir, "xlsx-deferred-unmount-events.json", events)
    page.screenshot(path=str(attempt_dir / "xlsx-deferred-unmounted.png"), full_page=True)
    return events


def pdf_slow_latest(page: Page, base_url: str, attempt_dir: Path) -> dict[str, object]:
    slow = "sample.pdf"
    before = event_count(base_url, "delay-start", slow)
    page.goto(f"{base_url}/#/pdf-viewer", wait_until="domcontentloaded")
    page.locator('[data-testid="pdf-sample-select"]').wait_for(state="visible")
    wait_for_event_count(base_url, "delay-start", slow, before)
    finished = event_count(base_url, "delay-finished", slow)
    page.locator('[data-testid="pdf-sample-select"]').select_option("rotated-pages.pdf")
    wait_ready(page, "rotated-pages.pdf")
    wait_for_event_count(base_url, "delay-finished", slow, finished)
    page.wait_for_timeout(250)
    final = pdf_snapshot(page)
    assert final["status"] == "ready", final
    assert "rotated-pages.pdf" in str(final["loadedFile"]), final
    assert "pdf-controller-" in str(final["diagnostics"]), final
    assert final["resources"]["activeWorkers"] > 0, final
    assert final["resources"]["activeObjectUrls"] > 0, final
    page.screenshot(path=str(attempt_dir / "pdf-slow-latest.png"), full_page=True)

    page.evaluate("location.hash = '#/'")
    page.wait_for_function(
        """() => !document.querySelector('[data-testid="pdf-viewer-page"]') &&
          window.__p1ResourceProbe.activeWorkers === 0 &&
          window.__p1ResourceProbe.activeObjectUrls === 0""",
        timeout=5_000,
    )
    unmounted = resource_state(page)
    assert unmounted["createdWorkers"] == unmounted["terminatedWorkers"], unmounted
    assert unmounted["createdObjectUrls"] == unmounted["revokedObjectUrls"], unmounted
    return {"mounted": final, "unmounted": unmounted}


Case = tuple[str, Callable[[Page, str, Path], dict[str, object]]]


def run_attempt(
    browser,
    base_url: str,
    delayed_file: str,
    case_id: str,
    workflow,
    attempt: int,
) -> dict[str, object]:
    attempt_dir = OUTPUT / case_id / f"attempt-{attempt}"
    attempt_dir.mkdir(parents=True, exist_ok=True)
    context = browser.new_context(
        viewport={"width": 1440, "height": 900}, accept_downloads=True
    )
    context.add_init_script(RESOURCE_PROBE)
    context.tracing.start(screenshots=True, snapshots=True, sources=True)
    page = context.new_page()
    evidence = BrowserEvidence(page)
    try:
        details = workflow(page, base_url, attempt_dir)
        policy = assert_race_evidence(evidence, delayed_file, case_id)
        result: dict[str, object] = {
            "id": case_id,
            "attempt": attempt,
            "status": "PASS",
            "details": details,
            "racePolicy": policy,
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
        }
    finally:
        evidence.save(attempt_dir)
        context.tracing.stop(path=attempt_dir / "trace.zip")
        context.close()
    (attempt_dir / "result.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return result


def run_group(browser, name: str, port: int, delayed_file: str, cases: list[Case]) -> list[dict[str, object]]:
    group_dir = OUTPUT / name
    group_dir.mkdir(parents=True, exist_ok=True)
    base_url = f"http://127.0.0.1:{port}"
    server_log = (group_dir / "fault-server.log").open("wb")
    env = {
        **os.environ,
        "DIST_DIR": str(ROOT / "apps/demo/dist"),
        "PORT": str(port),
        "DELAY_FILE": delayed_file,
        "DELAY_MS": str(DELAY_MS),
        "DELAY_HITS": "50",
        "EVENTS_PATH": str(group_dir / "fault-events.jsonl"),
    }
    server = subprocess.Popen(
        [sys.executable, str(Path(__file__).with_name("fault_server.py"))],
        cwd=ROOT,
        env=env,
        stdout=server_log,
        stderr=subprocess.STDOUT,
        start_new_session=True,
    )
    try:
        wait_for_server(base_url, server)
        results = []
        for case_id, workflow in cases:
            attempts = [run_attempt(browser, base_url, delayed_file, case_id, workflow, 1)]
            if attempts[0]["status"] == "FAIL":
                attempts.append(run_attempt(browser, base_url, delayed_file, case_id, workflow, 2))
            if attempts[0]["status"] == "PASS":
                status = "PASS"
            elif attempts[-1]["status"] == "PASS":
                status = "FLAKY"
            else:
                status = "FAIL"
            results.append({"id": case_id, "status": status, "attempts": attempts})
        return results
    finally:
        stop_process(server)
        server_log.close()


def main() -> int:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as playwright:
        browser, browser_source = launch_browser(playwright)
        browser_version = browser.version
        results = run_group(
            browser,
            "docx-server",
            available_port(4181),
            "demo.docx",
            [
                ("RACE-DOCX-SLOW-LATEST", docx_slow_latest),
                ("RACE-DOCX-SWITCH-20", docx_switch_20),
                ("RACE-DOCX-UNMOUNT", docx_unmount),
                ("RACE-DOCX-EXPORT-DOUBLE", docx_double_export),
                ("RACE-DOCX-EXPORT-ROUTE", docx_export_then_route),
            ],
        )
        results += run_group(
            browser,
            "xlsx-server",
            available_port(4182),
            "financial-model.xlsx",
            [
                ("RACE-XLSX-SLOW-LATEST", xlsx_slow_latest),
                ("RACE-XLSX-SWITCH-20", xlsx_switch_20),
                ("RACE-XLSX-WORKER-UNMOUNT", xlsx_worker_unmount),
                ("RACE-XLSX-DEFERRED-DOUBLE", xlsx_deferred_double_continue),
                ("RACE-XLSX-DEFERRED-UNMOUNT", xlsx_deferred_worker_unmount),
            ],
        )
        results += run_group(
            browser,
            "pdf-server",
            available_port(4183),
            "sample.pdf",
            [
                ("RACE-PDF-SLOW-LATEST", pdf_slow_latest),
            ],
        )
        browser.close()
    overall = "PASS" if all(item["status"] == "PASS" for item in results) else "FAIL"
    summary = {
        "suite": "BB-RACE",
        "mode": "formal build with controlled delayed responses",
        "result": overall,
        "browser": browser_version,
        "browserSource": browser_source,
        "results": results,
    }
    (OUTPUT / "summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    if overall != "PASS":
        failed_cases = [r for r in results if r["status"] != "PASS"]
        print(f"BB-RACE FAIL: {len(failed_cases)} of {len(results)} cases failed", file=sys.stderr)
        for case in failed_cases:
            failures = [
                str(attempt.get("error", "unknown"))
                for attempt in case.get("attempts", [])
                if attempt.get("status") == "FAIL"
            ]
            print(f"  {case['id']}: {' | '.join(failures) or 'unknown'}", file=sys.stderr)
    return 0 if overall == "PASS" else 1


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        OUTPUT.mkdir(parents=True, exist_ok=True)
        (OUTPUT / "summary.json").write_text(
            json.dumps(
                {
                    "suite": "BB-RACE",
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
