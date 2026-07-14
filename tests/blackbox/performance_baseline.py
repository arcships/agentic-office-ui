from __future__ import annotations

import argparse
import gzip
import hashlib
import json
import math
import os
import platform
import signal
import socket
import statistics
import subprocess
import sys
import time
import traceback
import urllib.request
from importlib.metadata import version as package_version
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import Page, sync_playwright

from browser_evidence import BrowserEvidence


ROOT = Path(__file__).resolve().parents[2]
CONFIG_PATH = Path(
    os.environ.get(
        "PERFORMANCE_BASELINE_CONFIG",
        ROOT / "tests" / "blackbox" / "performance-baseline.json",
    )
).resolve()
OUTPUT = Path(
    os.environ.get(
        "BLACKBOX_EVIDENCE_DIR",
        ROOT / "output" / "acceptance" / "bb-perf-xlsx",
    )
).resolve()
PORT = int(os.environ.get("CI_PERF_PREVIEW_PORT", "4181"))
BASE_URL = f"http://127.0.0.1:{PORT}"
PACKAGE_NAMES = ("docx-core", "xlsx-core", "vue-docx", "vue-xlsx", "vue-pdf", "vue-ui")
FIXTURE_NAMES = (
    "financial-model.xlsx",
    "sales-table.xlsx",
    "charts-images.xlsx",
    "large-grid.xlsx",
    "corrupted.xlsx",
    "sample.pdf",
)

PERFORMANCE_PROBE = r"""
(() => {
  const activeUrls = new Set();
  const state = {
    activeWorkers: 0,
    createdWorkers: 0,
    terminatedWorkers: 0,
    createdObjectUrls: 0,
    revokedObjectUrls: 0,
    canvasDimensionWrites: 0,
    longTasks: [],
  };
  Object.defineProperty(window, "__performanceProbe", { value: state });

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

  const createObjectURL = URL.createObjectURL.bind(URL);
  const revokeObjectURL = URL.revokeObjectURL.bind(URL);
  URL.createObjectURL = (value) => {
    const url = createObjectURL(value);
    activeUrls.add(url);
    state.createdObjectUrls += 1;
    return url;
  };
  URL.revokeObjectURL = (url) => {
    if (activeUrls.delete(String(url))) state.revokedObjectUrls += 1;
    return revokeObjectURL(url);
  };
  Object.defineProperty(state, "activeObjectUrls", {
    get: () => activeUrls.size,
  });

  for (const property of ['width', 'height']) {
    const descriptor = Object.getOwnPropertyDescriptor(HTMLCanvasElement.prototype, property);
    if (!descriptor?.get || !descriptor?.set) continue;
    Object.defineProperty(HTMLCanvasElement.prototype, property, {
      configurable: descriptor.configurable,
      enumerable: descriptor.enumerable,
      get: descriptor.get,
      set(value) {
        state.canvasDimensionWrites += 1;
        return descriptor.set.call(this, value);
      },
    });
  }

  try {
    const observer = new PerformanceObserver((list) => {
      for (const entry of list.getEntries()) {
        state.longTasks.push({ startTime: entry.startTime, duration: entry.duration });
      }
    });
    observer.observe({ type: "longtask", buffered: true });
  } catch {
    // Chromium without Long Tasks support is reported by the collector.
  }
})();
"""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Capture the formal XLSX performance baseline.")
    parser.add_argument("--rounds", type=int, default=None)
    parser.add_argument("--scroll-seconds", type=float, default=None)
    parser.add_argument("--config", type=Path, default=CONFIG_PATH)
    return parser.parse_args()


def git_value(*args: str) -> str:
    result = subprocess.run(
        ["git", *args], cwd=ROOT, text=True, capture_output=True, check=False
    )
    return result.stdout.strip() if result.returncode == 0 else "unknown"


def fixture_manifest() -> dict[str, dict[str, Any]]:
    data = json.loads((ROOT / "test-data" / "manifest.json").read_text(encoding="utf-8"))
    return {entry["file"]: entry for entry in data["entries"]}


def verify_fixtures(config: dict[str, Any]) -> dict[str, dict[str, Any]]:
    manifest = fixture_manifest()
    resolved: dict[str, dict[str, Any]] = {}
    configured = config.get("fixtures", {})
    for name in FIXTURE_NAMES:
        entry = manifest.get(name)
        if not entry:
            raise AssertionError(f"fixture {name} is missing from test-data/manifest.json")
        expected = configured.get(name)
        if expected and (
            int(expected["bytes"]) != int(entry["bytes"])
            or expected["sha256"] != entry["sha256"]
        ):
            raise AssertionError(f"fixture {name} no longer matches the fixed performance baseline")
        path = ROOT / "apps" / "demo" / "public" / "samples" / name
        content = path.read_bytes()
        actual_hash = hashlib.sha256(content).hexdigest()
        if len(content) != int(entry["bytes"]) or actual_hash != entry["sha256"]:
            raise AssertionError(f"fixture {name} bytes do not match its manifest")
        resolved[name] = {
            "bytes": len(content),
            "sha256": actual_hash,
            "category": entry["category"],
            "source": entry["source"],
        }
    return resolved


def gzip_bytes(path: Path) -> int:
    return len(gzip.compress(path.read_bytes(), compresslevel=9, mtime=0))


def file_kind(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".js":
        return "worker" if "worker" in path.name else "js"
    if suffix == ".css":
        return "css"
    if suffix == ".wasm":
        return "wasm"
    if suffix in {".ts", ".map"} or path.name.endswith(".d.ts"):
        return "types"
    return "other"


def summarize_tree(root: Path) -> dict[str, Any]:
    if not root.is_dir():
        raise AssertionError(f"formal build output is missing: {root}")
    files = []
    totals: dict[str, dict[str, int]] = {}
    for path in sorted(item for item in root.rglob("*") if item.is_file()):
        relative = path.relative_to(root).as_posix()
        size = path.stat().st_size
        zipped = gzip_bytes(path)
        kind = file_kind(path)
        totals.setdefault(kind, {"rawBytes": 0, "gzipBytes": 0, "files": 0})
        totals[kind]["rawBytes"] += size
        totals[kind]["gzipBytes"] += zipped
        totals[kind]["files"] += 1
        files.append(
            {
                "path": relative,
                "kind": kind,
                "rawBytes": size,
                "gzipBytes": zipped,
                "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
            }
        )
    return {
        "root": str(root),
        "rawBytes": sum(item["rawBytes"] for item in files),
        "gzipBytes": sum(item["gzipBytes"] for item in files),
        "totals": totals,
        "files": files,
    }


def collect_build_metrics() -> tuple[dict[str, Any], dict[str, float], list[dict[str, Any]]]:
    packages = {
        name: summarize_tree(ROOT / "packages" / name / "dist") for name in PACKAGE_NAMES
    }
    demo = summarize_tree(ROOT / "apps" / "demo" / "dist")
    all_trees = {"demo": demo, **packages}
    metrics: dict[str, float] = {}
    for name, summary in all_trees.items():
        prefix = f"asset.{name}"
        metrics[f"{prefix}.totalRawBytes"] = float(summary["rawBytes"])
        metrics[f"{prefix}.totalGzipBytes"] = float(summary["gzipBytes"])
        for kind in ("js", "worker", "css", "wasm", "types"):
            values = summary["totals"].get(kind, {})
            metrics[f"{prefix}.{kind}RawBytes"] = float(values.get("rawBytes", 0))
            metrics[f"{prefix}.{kind}GzipBytes"] = float(values.get("gzipBytes", 0))

    wasm_by_hash: dict[str, list[str]] = {}
    for item in demo["files"]:
        if item["kind"] == "wasm":
            wasm_by_hash.setdefault(item["sha256"], []).append(item["path"])
    duplicate_wasm = [paths for paths in wasm_by_hash.values() if len(paths) > 1]
    findings: list[dict[str, Any]] = []
    if duplicate_wasm:
        findings.append(
            {
                "case": "PERF-001",
                "severity": "P1",
                "code": "DUPLICATE_WASM",
                "details": duplicate_wasm,
            }
        )
    metrics["asset.demo.duplicateWasmCopies"] = float(
        sum(len(paths) - 1 for paths in duplicate_wasm)
    )

    main_scripts = [
        ROOT / "apps" / "demo" / "dist" / item["path"]
        for item in demo["files"]
        if item["kind"] == "js" and item["path"].startswith("assets/index-")
    ]
    eager_xlsx = any(b"xlsx-grid" in path.read_bytes() for path in main_scripts)
    metrics["asset.demo.homeMainContainsXlsx"] = 1.0 if eager_xlsx else 0.0
    if eager_xlsx:
        findings.append(
            {
                "case": "PERF-002",
                "severity": "P1",
                "code": "XLSX_EAGER_IN_HOME_BUNDLE",
                "details": [path.name for path in main_scripts],
            }
        )
    return {"packages": packages, "demo": demo}, metrics, findings


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
    launch_args = ["--enable-precise-memory-info"]
    try:
        return playwright.chromium.launch(headless=True, args=launch_args), "bundled-chromium"
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
                playwright.chromium.launch(
                    headless=True, executable_path=str(candidate), args=launch_args
                ),
                str(candidate),
            )
        except PlaywrightError as error:
            failures.append(str(error))
    raise RuntimeError("no Chromium executable could launch: " + " | ".join(failures))


def wait_xlsx(page: Page, file_name: str, state: str = "ready") -> None:
    page.wait_for_function(
        """expected => {
          const status = document.querySelector('[data-testid="page-status"]');
          const loaded = document.querySelector('[data-testid="loaded-file"]');
          const grid = document.querySelector('[data-testid="xlsx-grid"]');
          if (status?.dataset.state !== expected.state || !loaded?.textContent?.includes(expected.fileName)) return false;
          if (expected.state !== 'ready') return true;
          const canvas = grid?.querySelector('.xlsx-grid__body');
          return grid?.getBoundingClientRect().height > 0 && canvas instanceof HTMLCanvasElement && canvas.width > 0 && canvas.height > 0;
        }""",
        arg={"fileName": file_name, "state": state},
        timeout=45_000,
    )


def click_demo_route(page: Page, href: str) -> None:
    """Use the app's native link without retaining selector-engine node handles."""
    page.evaluate(
        """href => {
          const link = [...document.querySelectorAll('a')].find((item) => item.getAttribute('href') === href);
          if (!(link instanceof HTMLAnchorElement)) throw new Error(`route link is missing: ${href}`);
          link.click();
        }""",
        href,
    )
    page.wait_for_function("href => location.hash === href", arg=href, timeout=10_000)


def wait_home(page: Page) -> None:
    page.wait_for_function(
        """() => [...document.querySelectorAll('h2')].some((heading) => heading.textContent?.trim() === 'Agentic Office UI')""",
        timeout=10_000,
    )


def wait_xlsx_terminal_state(page: Page, file_name: str) -> str:
    page.wait_for_function(
        """expected => {
          const status = document.querySelector('[data-testid="page-status"]');
          const loaded = document.querySelector('[data-testid="loaded-file"]');
          return loaded?.textContent?.includes(expected) && ['ready', 'error'].includes(status?.dataset.state ?? '');
        }""",
        arg=file_name,
        timeout=30_000,
    )
    return page.locator('[data-testid="page-status"]').get_attribute("data-state") or "unknown"


def wait_chart_renderer(page: Page) -> None:
    page.wait_for_function(
        """() => {
          const overlay = document.querySelector('[data-testid="xlsx-chart-overlay"]');
          const chartCount = Number(overlay?.getAttribute('data-chart-count') ?? '0');
          const terminal = overlay?.getAttribute('data-state') === 'ready' && chartCount > 0;
          const rendered = !!document.querySelector('[data-testid="xlsx-chart-item"] svg[role="img"]');
          const failed = !!document.querySelector('[data-testid="xlsx-optional-renderer-error"]');
          return terminal && (rendered || failed);
        }""",
        timeout=30_000,
    )
    if page.locator('[data-testid="xlsx-optional-renderer-error"]').count():
        raise AssertionError("optional XLSX chart renderer failed to load")


def resource_entries(page: Page) -> list[dict[str, Any]]:
    return page.evaluate(
        """() => performance.getEntriesByType('resource').map((entry) => ({
          name: entry.name,
          initiatorType: entry.initiatorType,
          transferSize: entry.transferSize,
          encodedBodySize: entry.encodedBodySize,
          decodedBodySize: entry.decodedBodySize,
          duration: entry.duration,
          startTime: entry.startTime,
        }))"""
    )


def transfer_total(entries: list[dict[str, Any]]) -> float:
    return float(sum(float(item.get("transferSize") or item.get("encodedBodySize") or 0) for item in entries))


def memory_snapshot(page: Page, cdp) -> dict[str, float]:
    browser_memory = page.evaluate(
        """() => ({
          usedJSHeapSize: performance.memory?.usedJSHeapSize ?? 0,
          totalJSHeapSize: performance.memory?.totalJSHeapSize ?? 0,
          liveElements: document.querySelectorAll('*').length,
          liveTreeNodes: (() => {
            const walker = document.createTreeWalker(document, NodeFilter.SHOW_ALL);
            let count = 1;
            while (walker.nextNode()) count += 1;
            return count;
          })(),
        })"""
    )
    raw = cdp.send("Performance.getMetrics")
    cdp_metrics = {entry["name"]: float(entry["value"]) for entry in raw["metrics"]}
    return {
        "usedJSHeapBytes": float(browser_memory["usedJSHeapSize"]),
        "totalJSHeapBytes": float(browser_memory["totalJSHeapSize"]),
        "liveElements": float(browser_memory["liveElements"]),
        "liveTreeNodes": float(browser_memory["liveTreeNodes"]),
        "documents": cdp_metrics.get("Documents", 0.0),
        "nodes": cdp_metrics.get("Nodes", 0.0),
        "listeners": cdp_metrics.get("JSEventListeners", 0.0),
    }


def collect_page_garbage(cdp) -> None:
    """Collect unreachable page objects before measuring retained memory."""
    cdp.send("HeapProfiler.enable")
    cdp.send("HeapProfiler.collectGarbage")


def long_tasks(page: Page, since: float) -> list[dict[str, float]]:
    return page.evaluate(
        """since => window.__performanceProbe.longTasks.filter((entry) => entry.startTime >= since)""",
        since,
    )


def percentile(values: list[float], percent: float) -> float:
    if not values:
        return 0.0
    ordered = sorted(values)
    index = max(0, min(len(ordered) - 1, math.ceil(percent * len(ordered)) - 1))
    return float(ordered[index])


def scroll_grid(page: Page, duration_ms: float) -> dict[str, Any]:
    result = page.evaluate(
        """durationMs => new Promise((resolve, reject) => {
          const grid = document.querySelector('[data-testid="xlsx-grid"]');
          if (!(grid instanceof HTMLElement)) return reject(new Error('XLSX grid is missing'));
          const canvases = [...grid.querySelectorAll('canvas')];
          const initialCanvasDimensionWrites = window.__performanceProbe.canvasDimensionWrites;
          const frameDurations = [];
          const startedAt = performance.now();
          let previous = startedAt;
          const maxTop = Math.max(0, grid.scrollHeight - grid.clientHeight);
          const maxLeft = Math.max(0, grid.scrollWidth - grid.clientWidth);
          function frame(now) {
            if (now > previous) frameDurations.push(now - previous);
            previous = now;
            const progress = Math.min(1, (now - startedAt) / durationMs);
            if (progress < 0.5) grid.scrollTop = maxTop * progress * 2;
            else grid.scrollLeft = maxLeft * (progress - 0.5) * 2;
            grid.dispatchEvent(new Event('scroll'));
            if (progress < 1) requestAnimationFrame(frame);
            else requestAnimationFrame(() => {
              resolve({
                frameDurations,
                canvasDimensionWrites: window.__performanceProbe.canvasDimensionWrites - initialCanvasDimensionWrites,
                durationMs: performance.now() - startedAt,
                finalScrollTop: grid.scrollTop,
                finalScrollLeft: grid.scrollLeft,
                maxTop,
                maxLeft,
                canvasCount: canvases.length,
              });
            });
          }
          requestAnimationFrame(frame);
        })""",
        duration_ms,
    )
    frames = [float(value) for value in result["frameDurations"] if float(value) > 0]
    result["frameP95Ms"] = percentile(frames, 0.95)
    result["frameWorstMs"] = max(frames, default=0.0)
    result["frameMedianMs"] = statistics.median(frames) if frames else 0.0
    result["frameCount"] = len(frames)
    del result["frameDurations"]
    return result


def interaction_metrics(page: Page, round_number: int) -> dict[str, Any]:
    grid = page.locator('[data-testid="xlsx-grid"]')
    name_box = page.locator('[data-testid="xlsx-name-box"]')
    formula = page.locator('[data-testid="xlsx-formula-input"]')
    before_name = name_box.input_value()
    started = page.evaluate("performance.now()")
    grid.click(position={"x": 220, "y": 180})
    page.wait_for_function(
        """before => {
          const box = document.querySelector('[data-testid="xlsx-name-box"]');
          return box instanceof HTMLInputElement && !!box.value && box.value !== before;
        }""",
        arg=before_name,
        timeout=10_000,
    )
    selected = name_box.input_value()
    select_ms = float(page.evaluate("start => performance.now() - start", started))
    if selected == "A1":
        raise AssertionError("selection after deep scroll still reports A1")
    grid.dblclick(position={"x": 220, "y": 180})
    editor = page.locator('[data-testid="xlsx-cell-editor"]')
    editor.wait_for(state="visible", timeout=10_000)
    original = editor.input_value()
    marker = f"PERF-R{round_number}"
    editor.fill(marker)
    started = page.evaluate("performance.now()")
    editor.press("Tab")
    page.wait_for_function(
        """before => document.querySelector('[data-testid="xlsx-name-box"]')?.value !== before""",
        arg=selected,
        timeout=10_000,
    )
    tab_target = name_box.input_value()
    tab_ms = float(page.evaluate("start => performance.now() - start", started))
    page.locator('[data-testid="xlsx-cell-editor"]').press("Escape")
    undo = page.locator('[data-testid="xlsx-undo"]')
    page.wait_for_function(
        """() => !document.querySelector('[data-testid="xlsx-undo"]')?.disabled""",
        timeout=10_000,
    )
    started = page.evaluate("performance.now()")
    undo.click()
    page.wait_for_function(
        """() => document.querySelector('[data-testid="xlsx-undo"]')?.disabled === true""",
        timeout=10_000,
    )
    undo_ms = float(page.evaluate("start => performance.now() - start", started))
    grid.dblclick(position={"x": 220, "y": 180})
    restored_editor = page.locator('[data-testid="xlsx-cell-editor"]')
    restored_editor.wait_for(state="visible", timeout=10_000)
    if restored_editor.input_value() != original:
        raise AssertionError("undo did not restore the edited cell")
    restored_editor.press("Escape")
    return {
        "selectedCell": selected,
        "tabTarget": tab_target,
        "originalValue": original,
        "marker": marker,
        "selectMs": select_ms,
        "tabCommitMs": tab_ms,
        "undoMs": undo_ms,
    }


def probe_state(page: Page) -> dict[str, Any]:
    return page.evaluate(
        """() => ({
          activeWorkers: window.__performanceProbe.activeWorkers,
          createdWorkers: window.__performanceProbe.createdWorkers,
          terminatedWorkers: window.__performanceProbe.terminatedWorkers,
          activeObjectUrls: window.__performanceProbe.activeObjectUrls,
          createdObjectUrls: window.__performanceProbe.createdObjectUrls,
          revokedObjectUrls: window.__performanceProbe.revokedObjectUrls,
        })"""
    )


def collect_released_home_state(page: Page, cdp, cycle: int) -> dict[str, Any]:
    page.wait_for_timeout(750)
    collect_page_garbage(cdp)
    page.wait_for_timeout(250)
    return {
        "cycle": cycle,
        "memory": memory_snapshot(page, cdp),
        "resources": probe_state(page),
    }


def run_xlsx_lifecycle_cycles(page: Page, cdp, first_sample: dict[str, Any]) -> list[dict[str, Any]]:
    samples = [first_sample]
    for cycle in range(2, 4):
        click_demo_route(page, "#/xlsx-viewer")
        wait_xlsx(page, "financial-model.xlsx")
        page.wait_for_load_state("networkidle")
        click_demo_route(page, "#/")
        wait_home(page)
        samples.append(collect_released_home_state(page, cdp, cycle))
    return samples


def run_round(browser, round_number: int, scroll_seconds: float) -> dict[str, Any]:
    round_dir = OUTPUT / f"round-{round_number}"
    round_dir.mkdir(parents=True, exist_ok=True)
    context = browser.new_context(
        viewport={"width": 1440, "height": 900},
        device_scale_factor=1,
        locale="zh-CN",
        timezone_id="Asia/Shanghai",
    )
    context.add_init_script(PERFORMANCE_PROBE)
    # DOM snapshots keep old page trees alive until tracing stops. That makes
    # CDP Nodes/JSEventListeners measure the test recorder instead of the
    # released application. We already save explicit before/after screenshots,
    # network, console and page errors, so keep the action trace without DOM
    # snapshots while collecting memory metrics.
    trace_path = round_dir / "trace.zip"
    context.tracing.start(screenshots=True, snapshots=False, sources=True)
    tracing_active = True
    page = context.new_page()
    evidence = BrowserEvidence(page)
    cdp = context.new_cdp_session(page)
    cdp.send("Performance.enable")
    partial_metrics: dict[str, float] = {}
    completed_steps: list[str] = []
    try:
        page.goto(f"{BASE_URL}/#/", wait_until="networkidle")
        wait_home(page)
        home_resources = resource_entries(page)
        home_memory = memory_snapshot(page, cdp)
        partial_metrics["browser.home.transferBytes"] = transfer_total(home_resources)
        completed_steps.append("PERF-002-home")
        page.screenshot(path=str(round_dir / "PERF-002-home.png"), full_page=True)

        click_demo_route(page, "#/xlsx-viewer")
        wait_xlsx(page, "financial-model.xlsx")
        page.wait_for_load_state("networkidle")
        xlsx_resources = resource_entries(page)
        initial_memory = memory_snapshot(page, cdp)
        partial_metrics["browser.xlsx.totalTransferBytes"] = transfer_total(xlsx_resources)
        partial_metrics["browser.xlsx.additionalTransferBytes"] = max(
            0.0, transfer_total(xlsx_resources) - transfer_total(home_resources)
        )
        completed_steps.append("PERF-002-xlsx")

        load_started = float(page.evaluate("performance.now()"))
        page.locator('[data-testid="xlsx-sample-select"]').select_option("large-grid.xlsx")
        wait_xlsx(page, "large-grid.xlsx")
        load_ms = float(page.evaluate("start => performance.now() - start", load_started))
        load_tasks = long_tasks(page, load_started)
        loaded_memory = memory_snapshot(page, cdp)
        partial_metrics.update(
            {
                "browser.largeGrid.loadToOperableMs": load_ms,
                "browser.largeGrid.loadLongTaskCount": float(len(load_tasks)),
                "browser.largeGrid.loadLongTaskWorstMs": max(
                    (float(item["duration"]) for item in load_tasks), default=0.0
                ),
            }
        )
        completed_steps.append("PERF-003")
        page.screenshot(path=str(round_dir / "PERF-003-large-grid-ready.png"), full_page=True)

        scroll_started = float(page.evaluate("performance.now()"))
        scroll = scroll_grid(page, scroll_seconds * 1000)
        scroll_tasks = long_tasks(page, scroll_started)
        partial_metrics.update(
            {
                "browser.scroll.frameP95Ms": float(scroll["frameP95Ms"]),
                "browser.scroll.frameWorstMs": float(scroll["frameWorstMs"]),
                "browser.scroll.maxVerticalPixels": float(scroll["maxTop"]),
                "browser.scroll.maxHorizontalPixels": float(scroll["maxLeft"]),
                "browser.scroll.longTaskCount": float(len(scroll_tasks)),
                "browser.scroll.longTaskWorstMs": max(
                    (float(item["duration"]) for item in scroll_tasks), default=0.0
                ),
                "browser.scroll.canvasDimensionWrites": float(scroll["canvasDimensionWrites"]),
            }
        )
        completed_steps.append("PERF-004")
        page.screenshot(path=str(round_dir / "PERF-004-large-grid-scrolled.png"), full_page=True)
        interaction = interaction_metrics(page, round_number)
        partial_metrics.update(
            {
                "browser.interaction.selectMs": float(interaction["selectMs"]),
                "browser.interaction.tabCommitMs": float(interaction["tabCommitMs"]),
                "browser.interaction.undoMs": float(interaction["undoMs"]),
            }
        )
        completed_steps.append("PERF-005")
        page.screenshot(path=str(round_dir / "PERF-005-interaction.png"), full_page=True)

        resources_before_chart = {entry["name"] for entry in resource_entries(page)}
        page.locator('[data-testid="xlsx-sample-select"]').select_option("charts-images.xlsx")
        wait_xlsx(page, "charts-images.xlsx")
        wait_chart_renderer(page)
        page.wait_for_load_state("networkidle")
        chart_resources = resource_entries(page)
        chart_new_resources = [
            entry for entry in chart_resources if entry["name"] not in resources_before_chart
        ]
        chart_js = [entry for entry in chart_new_resources if urlparse(entry["name"]).path.endswith(".js")]
        partial_metrics["browser.chart.lazyJsRequests"] = float(len(chart_js))
        completed_steps.append("PERF-006")
        page.screenshot(path=str(round_dir / "PERF-006-charts.png"), full_page=True)

        memory_samples = [home_memory, initial_memory, loaded_memory]
        sequence_findings = []
        for fixture in (
            "financial-model.xlsx",
            "sales-table.xlsx",
            "charts-images.xlsx",
            "large-grid.xlsx",
            "corrupted.xlsx",
        ):
            page.locator('[data-testid="xlsx-sample-select"]').select_option(fixture)
            actual_state = wait_xlsx_terminal_state(page, fixture)
            expected_state = "error" if fixture == "corrupted.xlsx" else "ready"
            if actual_state != expected_state:
                error = page.locator('[data-testid="load-error"]')
                sequence_findings.append(
                    {
                        "case": "PERF-007",
                        "severity": "P1",
                        "code": "WORKBOOK_SEQUENCE_LOAD_FAILED",
                        "details": {
                            "fixture": fixture,
                            "expectedState": expected_state,
                            "actualState": actual_state,
                            "error": error.inner_text() if error.count() else "",
                        },
                    }
                )
            memory_samples.append(memory_snapshot(page, cdp))
        peak_heap = max(sample["usedJSHeapBytes"] for sample in memory_samples)
        peak_nodes = max(sample["nodes"] for sample in memory_samples)
        peak_listeners = max(sample["listeners"] for sample in memory_samples)
        before_home_state = probe_state(page)

        # Stop the recorder before checking release. Even without DOM snapshots,
        # Playwright tracing retains action targets until the trace is closed.
        # Those recorder-owned nodes/listeners must not be counted as product
        # retention.
        context.tracing.stop(path=trace_path)
        tracing_active = False
        click_demo_route(page, "#/")
        wait_home(page)
        first_lifecycle_sample = collect_released_home_state(page, cdp, 1)
        lifecycle_samples = run_xlsx_lifecycle_cycles(page, cdp, first_lifecycle_sample)
        after_home_memory = lifecycle_samples[-1]["memory"]
        after_home_state = lifecycle_samples[-1]["resources"]
        lifecycle_node_growth = max(sample["memory"]["nodes"] for sample in lifecycle_samples) - min(
            sample["memory"]["nodes"] for sample in lifecycle_samples
        )
        lifecycle_listener_growth = max(sample["memory"]["listeners"] for sample in lifecycle_samples) - min(
            sample["memory"]["listeners"] for sample in lifecycle_samples
        )
        after_home_node_delta = max(0.0, after_home_memory["nodes"] - home_memory["nodes"])
        after_home_listener_delta = max(0.0, after_home_memory["listeners"] - home_memory["listeners"])
        after_home_live_tree_delta = max(
            0.0, after_home_memory["liveTreeNodes"] - home_memory["liveTreeNodes"]
        )
        after_home_live_element_delta = max(
            0.0, after_home_memory["liveElements"] - home_memory["liveElements"]
        )
        page.screenshot(path=str(round_dir / "PERF-007-home-released.png"), full_page=True)

        evidence.assert_clean()
        metrics = {
            "browser.home.transferBytes": transfer_total(home_resources),
            "browser.xlsx.totalTransferBytes": transfer_total(xlsx_resources),
            "browser.xlsx.additionalTransferBytes": max(0.0, transfer_total(xlsx_resources) - transfer_total(home_resources)),
            "browser.largeGrid.loadToOperableMs": load_ms,
            "browser.largeGrid.loadLongTaskCount": float(len(load_tasks)),
            "browser.largeGrid.loadLongTaskWorstMs": max((float(item["duration"]) for item in load_tasks), default=0.0),
            "browser.scroll.frameP95Ms": float(scroll["frameP95Ms"]),
            "browser.scroll.frameWorstMs": float(scroll["frameWorstMs"]),
            "browser.scroll.maxVerticalPixels": float(scroll["maxTop"]),
            "browser.scroll.maxHorizontalPixels": float(scroll["maxLeft"]),
            "browser.scroll.longTaskCount": float(len(scroll_tasks)),
            "browser.scroll.longTaskWorstMs": max((float(item["duration"]) for item in scroll_tasks), default=0.0),
            "browser.scroll.canvasDimensionWrites": float(scroll["canvasDimensionWrites"]),
            "browser.interaction.selectMs": float(interaction["selectMs"]),
            "browser.interaction.tabCommitMs": float(interaction["tabCommitMs"]),
            "browser.interaction.undoMs": float(interaction["undoMs"]),
            "browser.chart.lazyJsRequests": float(len(chart_js)),
            "browser.memory.peakHeapBytes": peak_heap,
            "browser.memory.afterHomeHeapBytes": after_home_memory["usedJSHeapBytes"],
            "browser.memory.retainedRatio": (after_home_memory["usedJSHeapBytes"] / peak_heap) if peak_heap else 0.0,
            "browser.memory.peakNodes": peak_nodes,
            "browser.memory.afterHomeNodes": after_home_memory["nodes"],
            "browser.memory.afterHomeNodeDelta": after_home_node_delta,
            "browser.memory.afterHomeLiveTreeDelta": after_home_live_tree_delta,
            "browser.memory.afterHomeLiveElementDelta": after_home_live_element_delta,
            "browser.memory.peakListeners": peak_listeners,
            "browser.memory.afterHomeListeners": after_home_memory["listeners"],
            "browser.memory.afterHomeListenerDelta": after_home_listener_delta,
            "browser.memory.lifecycleNodeGrowth": lifecycle_node_growth,
            "browser.memory.lifecycleListenerGrowth": lifecycle_listener_growth,
            "browser.resources.afterHomeActiveWorkers": float(after_home_state["activeWorkers"]),
            "browser.resources.afterHomeActiveObjectUrls": float(after_home_state["activeObjectUrls"]),
        }
        partial_metrics.update(metrics)
        completed_steps.append("PERF-007")
        findings = list(sequence_findings)
        if float(scroll["maxTop"]) <= 0 or float(scroll["maxLeft"]) <= 0:
            findings.append(
                {
                    "case": "PERF-004",
                    "severity": "P1",
                    "code": "LARGE_GRID_HAS_NO_SCROLL_RANGE",
                    "details": {
                        "maxTop": scroll["maxTop"],
                        "maxLeft": scroll["maxLeft"],
                        "frameMetricsAreIdleLoop": True,
                    },
                }
            )
        if int(scroll["canvasDimensionWrites"]) >= int(scroll["frameCount"]):
            findings.append(
                {
                    "case": "PERF-004",
                    "severity": "P1",
                    "code": "CANVAS_RESIZED_EVERY_FRAME",
                    "details": {
                        "dimensionWrites": scroll["canvasDimensionWrites"],
                        "frames": scroll["frameCount"],
                    },
                }
            )
        if not chart_js:
            findings.append(
                {
                    "case": "PERF-006",
                    "severity": "P1",
                    "code": "NO_LAZY_FEATURE_CHUNK",
                    "details": chart_new_resources,
                }
            )
        # CDP's total Nodes/JSEventListeners also includes objects retained by
        # Playwright's action trace and selector engine. Gate the released page
        # on its live document tree, then use the repeated route-cycle deltas
        # below to catch detached trees or listeners that keep accumulating.
        if after_home_live_tree_delta > 32 or after_home_live_element_delta > 16:
            findings.append(
                {
                    "case": "PERF-007",
                    "severity": "P1",
                    "code": "DOM_RETENTION_AFTER_ROUTE",
                    "details": {
                        "initialHome": home_memory,
                        "releasedHome": after_home_memory,
                        "nodeDelta": after_home_node_delta,
                        "listenerDelta": after_home_listener_delta,
                        "liveTreeDelta": after_home_live_tree_delta,
                        "liveElementDelta": after_home_live_element_delta,
                    },
                }
            )
        if lifecycle_node_growth > 8 or lifecycle_listener_growth > 4:
            findings.append(
                {
                    "case": "PERF-007",
                    "severity": "P1",
                    "code": "RESOURCE_GROWS_ACROSS_ROUTE_CYCLES",
                    "details": {
                        "nodeGrowth": lifecycle_node_growth,
                        "listenerGrowth": lifecycle_listener_growth,
                        "samples": lifecycle_samples,
                    },
                }
            )
        if after_home_state["activeWorkers"] or after_home_state["activeObjectUrls"]:
            findings.append(
                {
                    "case": "PERF-007",
                    "severity": "P1",
                    "code": "RESOURCE_NOT_RELEASED",
                    "details": after_home_state,
                }
            )
        result = {
            "round": round_number,
            "status": "CAPTURED",
            "metrics": metrics,
            "homeResources": home_resources,
            "xlsxResources": xlsx_resources,
            "loadLongTasks": load_tasks,
            "scroll": scroll,
            "scrollLongTasks": scroll_tasks,
            "interaction": interaction,
            "chartNewResources": chart_new_resources,
            "memorySamples": memory_samples,
            "beforeHomeResources": before_home_state,
            "afterHomeResources": after_home_state,
            "afterHomeMemory": after_home_memory,
            "lifecycleSamples": lifecycle_samples,
            "traceStoppedBeforeReleaseMetrics": True,
            "findings": findings,
            "completedSteps": completed_steps,
        }
    except Exception as error:
        try:
            page.screenshot(path=str(round_dir / "failure.png"), full_page=True)
        except Exception:
            pass
        result = {
            "round": round_number,
            "status": "FAIL",
            "error": repr(error),
            "traceback": traceback.format_exc(),
            "partialMetrics": partial_metrics,
            "completedSteps": completed_steps,
            "events": evidence.events,
        }
    finally:
        evidence.save(round_dir)
        if tracing_active:
            context.tracing.stop(path=trace_path)
        context.close()
    (round_dir / "result.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return result


def aggregate_metrics(
    build_metrics: dict[str, float], rounds: list[dict[str, Any]]
) -> dict[str, dict[str, Any]]:
    all_keys = set(build_metrics)
    for item in rounds:
        all_keys.update(item["metrics"])
    aggregated: dict[str, dict[str, Any]] = {}
    for key in sorted(all_keys):
        if key in build_metrics:
            samples = [float(build_metrics[key])] * len(rounds)
        else:
            samples = [float(item["metrics"][key]) for item in rounds]
        median = float(statistics.median(samples))
        worst = max(samples)
        minimum = min(samples)
        spread = ((worst - minimum) / median * 100) if median else (0.0 if worst == 0 else 100.0)
        aggregated[key] = {
            "samples": samples,
            "median": median,
            "worst": worst,
            "minimum": minimum,
            "spreadPercent": spread,
            "direction": "min" if key in {
                "browser.chart.lazyJsRequests",
                "browser.scroll.maxVerticalPixels",
                "browser.scroll.maxHorizontalPixels",
            } else "max",
        }
    return aggregated


def verify_budget(config: dict[str, Any], aggregate: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    if config.get("status") != "approved":
        return []
    violations = []
    regression = float(config.get("maxRegressionPercent", 10))
    absolute_ms = float(config.get("maxRegressionAbsoluteMs", 0))
    css_floor_bytes = float(config.get("cssRelativeBudgetFloorBytes", 0))
    for key, approved in config.get("metrics", {}).items():
        current = aggregate.get(key)
        if not current:
            violations.append({"metric": key, "reason": "missing"})
            continue
        baseline = float(approved["median"])
        direction = approved.get("direction", "max")
        if direction == "min":
            allowed = baseline * (1 - regression / 100)
            violated = float(current["median"]) < allowed
        else:
            allowed = 0.0 if baseline == 0 else baseline * (1 + regression / 100)
            if baseline > 0 and key.endswith("Ms"):
                allowed = max(allowed, baseline + absolute_ms)
            if key.endswith(("cssGzipBytes", "cssRawBytes")):
                allowed = max(allowed, css_floor_bytes)
            violated = float(current["median"]) > allowed
        if violated:
            violations.append(
                {
                    "metric": key,
                    "direction": direction,
                    "baselineMedian": baseline,
                    "currentMedian": current["median"],
                    "allowed": allowed,
                }
            )
    return violations


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


def write_environment(config: dict[str, Any], fixtures: dict[str, Any], extra: dict[str, Any] | None = None) -> None:
    environment = {
        "commit": git_value("rev-parse", "HEAD"),
        "branch": git_value("branch", "--show-current"),
        "cwd": str(ROOT),
        "platform": platform.platform(),
        "machine": platform.machine(),
        "python": sys.version,
        "playwright": package_version("playwright"),
        "node": subprocess.run(["node", "--version"], text=True, capture_output=True).stdout.strip(),
        "pnpm": subprocess.run(["pnpm", "--version"], text=True, capture_output=True).stdout.strip(),
        "viewport": {"width": 1440, "height": 900},
        "deviceScaleFactor": 1,
        "locale": "zh-CN",
        "timezone": "Asia/Shanghai",
        "baseUrl": BASE_URL,
        "mode": "formal preview",
        "config": str(CONFIG_PATH),
        "configStatus": config.get("status", "capture"),
        "fixtures": fixtures,
    }
    environment.update(extra or {})
    (OUTPUT / "environment.json").write_text(
        json.dumps(environment, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def main() -> int:
    args = parse_args()
    config_path = args.config.resolve()
    config = json.loads(config_path.read_text(encoding="utf-8"))
    approved_environment = config.get("approvedEnvironment")
    if approved_environment:
        actual_environment = {
            "platform": platform.system().lower(),
            "machine": platform.machine().lower(),
        }
        expected_environment = {
            "platform": str(approved_environment.get("platform", "")).lower(),
            "machine": str(approved_environment.get("machine", "")).lower(),
        }
        if actual_environment != expected_environment:
            raise RuntimeError(
                "performance baseline environment mismatch: "
                f"expected {expected_environment}, got {actual_environment}"
            )
    rounds_count = args.rounds or int(config["rounds"])
    scroll_seconds = args.scroll_seconds or float(config["scrollSeconds"])
    if rounds_count < 1:
        raise ValueError("rounds must be positive")
    OUTPUT.mkdir(parents=True, exist_ok=True)
    fixtures = verify_fixtures(config)
    write_environment(config, fixtures)
    build_report, build_metrics, build_findings = collect_build_metrics()
    (OUTPUT / "build-metrics.json").write_text(
        json.dumps(build_report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (OUTPUT / "commands.log").write_text(
        f"pnpm build\nCI_PERF_PREVIEW_PORT={PORT} python tests/blackbox/performance_baseline.py --rounds {rounds_count} --scroll-seconds {scroll_seconds}\n",
        encoding="utf-8",
    )
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
            expected_browser = config.get("browser", {}).get("version")
            if expected_browser and browser_version != expected_browser:
                raise RuntimeError(
                    f"fixed browser mismatch: expected {expected_browser}, got {browser_version}"
                )
            write_environment(
                config,
                fixtures,
                {"browser": browser_version, "browserSource": browser_source},
            )
            rounds = [
                run_round(browser, round_number, scroll_seconds)
                for round_number in range(1, rounds_count + 1)
            ]
            browser.close()
        failed_rounds = [item for item in rounds if item["status"] != "CAPTURED"]
        if failed_rounds:
            summary = {
                "suite": "BB-PERF-XLSX",
                "task": "P3-PERF-BASELINE-01",
                "result": "FAIL",
                "mode": "formal preview",
                "roundsRequested": rounds_count,
                "rounds": rounds,
            }
            (OUTPUT / "summary.json").write_text(
                json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            return 1

        aggregate = aggregate_metrics(build_metrics, rounds)
        violations = verify_budget(config, aggregate)
        findings = build_findings + [finding for item in rounds for finding in item["findings"]]
        blocking_findings = [
            finding for finding in findings
            if finding.get("severity") in {"P0", "P1"}
        ]
        status = config.get("status", "capture")
        result = "FAIL" if violations or (status == "approved" and blocking_findings) else (
            "PASS" if status == "approved" else "BASELINE_CAPTURED_PENDING_APPROVAL"
        )
        summary = {
            "suite": "BB-PERF-XLSX",
            "task": "P3-PERF-BASELINE-01",
            "result": result,
            "candidateReleasePass": result == "PASS",
            "mode": "formal preview",
            "commit": git_value("rev-parse", "HEAD"),
            "browser": browser_version,
            "browserSource": browser_source,
            "viewport": {"width": 1440, "height": 900},
            "deviceScaleFactor": 1,
            "roundsRequested": rounds_count,
            "scrollSeconds": scroll_seconds,
            "fixtures": fixtures,
            "aggregate": aggregate,
            "findings": findings,
            "blockingFindings": blocking_findings,
            "budgetStatus": status,
            "budgetViolations": violations,
            "rounds": rounds,
        }
        (OUTPUT / "summary.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        candidate = {
            **config,
            "status": "pending_approval",
            "baselineCommit": summary["commit"],
            "capturedBrowser": browser_version,
            "metrics": aggregate,
        }
        (OUTPUT / "baseline-candidate.json").write_text(
            json.dumps(candidate, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return 1 if violations else 0
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
                    "task": "P3-PERF-BASELINE-01",
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
