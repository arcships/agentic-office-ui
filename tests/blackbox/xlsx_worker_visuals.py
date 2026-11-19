#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import signal
import socket
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

from playwright.sync_api import sync_playwright

from browser_launch import launch_browser


ROOT = Path(__file__).resolve().parents[2]
OUTPUT = Path(
    os.environ.get(
        "BLACKBOX_EVIDENCE_DIR",
        ROOT / "output" / "acceptance" / "p4-fidelity-xlsx-worker",
    )
).resolve()
PORT = int(os.environ.get("CI_PREVIEW_PORT", "4173"))
BASE_URL = f"http://127.0.0.1:{PORT}"
MANAGE_PREVIEW = "BLACKBOX_BASE_URL" not in os.environ


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


def stop_process(process: subprocess.Popen[bytes] | None) -> None:
    if process is None or process.poll() is not None:
        return
    try:
        os.killpg(process.pid, signal.SIGTERM)
        process.wait(timeout=5)
    except Exception:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except Exception:
            process.kill()


def open_case(browser, base_url: str, attempt_dir: Path, mode: str) -> dict:
    case_dir = attempt_dir / mode
    case_dir.mkdir(parents=True, exist_ok=True)
    console: list[dict] = []
    page_errors: list[str] = []
    request_failures: list[dict] = []
    responses: list[dict] = []
    downloads: list[str] = []
    context = browser.new_context(
        viewport={"width": 1440, "height": 900},
        record_har_path=str(case_dir / "network.har"),
    )
    page = context.new_page()
    page.on("console", lambda message: console.append({"type": message.type, "text": message.text}))
    page.on("pageerror", lambda error: page_errors.append(str(error)))
    page.on("requestfailed", lambda request: request_failures.append({"url": request.url, "failure": request.failure}))
    page.on("response", lambda response: responses.append({"url": response.url, "status": response.status}))
    page.on("download", lambda download: downloads.append(download.suggested_filename))

    try:
        page.goto(f"{base_url}/#/xlsx-viewer", wait_until="networkidle")
        if mode == "worker":
            page.locator('[data-testid="xlsx-read-only"]').check()
            details = page.locator('[data-testid="xlsx-runtime-details"]')
            if details.get_attribute("open") is None:
                details.locator("summary").click()
            page.locator('[data-testid="xlsx-use-worker"]').check()
        page.locator('[data-testid="xlsx-sample-select"]').select_option("charts-images.xlsx")
        expected_mode = "Worker" if mode == "worker" else "主线程"
        page.wait_for_function(
            """([name, expected]) => document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'
              && document.querySelector('[data-testid="loaded-file"]')?.textContent?.includes(name)
              && document.querySelector('[data-testid="xlsx-worker-actual"]')?.textContent === expected""",
            arg=["charts-images.xlsx", expected_mode],
        )
        grid = page.locator('[data-testid="xlsx-grid"]')
        page.wait_for_function(
            """expected => document.querySelector('[data-testid="xlsx-grid"]')?.dataset.mergedRegionCount === '1'
              && document.querySelector('[data-testid="xlsx-grid"]')?.dataset.freezeRow === '2'
              && document.querySelector('[data-testid="xlsx-grid"]')?.dataset.freezeColumn === '2'
              && Number(document.querySelector('[data-testid="xlsx-grid"]')?.dataset.dataValidationCount || 0) >= 1
              && document.querySelector('[data-testid="xlsx-grid"]')?.dataset.workerBacked === expected""",
            arg="true" if mode == "worker" else "false",
        )
        if mode != "worker":
            note = page.locator('[data-testid="xlsx-worker-note"]')
            note.wait_for(timeout=10_000)
            if "只读" not in note.inner_text():
                raise AssertionError("the editable view does not explain why Worker is inactive")
        image = page.locator('[data-testid="xlsx-image-item"] img').first
        image.wait_for(state="visible")
        page.wait_for_function(
            """() => {
              const image = document.querySelector('[data-testid="xlsx-image-item"] img');
              const rect = image?.getBoundingClientRect();
              return Boolean(rect && rect.width > 0 && rect.height > 0);
            }"""
        )
        page.wait_for_function(
            """() => document.querySelector('[data-testid="xlsx-image-item"] img')?.dataset.imageState === 'ready'"""
        )
        pane_keys = sorted(
            set(
                page.locator("[data-grid-pane]").evaluate_all(
                    "elements => elements.map((element) => element.dataset.gridPane)"
                )
            )
        )
        if pane_keys != ["frozen-cols", "frozen-corner", "frozen-rows", "main"]:
            raise AssertionError(f"multi-axis freeze panes are missing: {pane_keys}")

        canvas = page.locator('[data-testid="xlsx-grid"] canvas.xlsx-grid__body')
        box = canvas.bounding_box()
        if not box:
            raise AssertionError("XLSX body canvas has no bounding box")
        page.mouse.click(box["x"] + 120, box["y"] + 7 * 24 + 12)
        page.wait_for_function("""() => document.querySelector('[data-testid="xlsx-name-box"]')?.value === 'A7'""")

        image_box = page.locator('[data-testid="xlsx-image-item"]').first.bounding_box()
        result = {
            "mode": mode,
            "actualMode": page.locator('[data-testid="xlsx-worker-actual"]').inner_text(),
            "mergedRegionCount": grid.get_attribute("data-merged-region-count"),
            "freezeRow": grid.get_attribute("data-freeze-row"),
            "freezeColumn": grid.get_attribute("data-freeze-column"),
            "dataValidationCount": grid.get_attribute("data-data-validation-count"),
            "paneKeys": pane_keys,
            "workerBacked": grid.get_attribute("data-worker-backed"),
            "workerNote": page.locator('[data-testid="xlsx-worker-note"]').inner_text()
            if page.locator('[data-testid="xlsx-worker-note"]').count()
            else None,
            "mergeSecondarySelectsAnchor": page.locator('[data-testid="xlsx-name-box"]').input_value(),
            "imageState": image.get_attribute("data-image-state"),
            "imageBox": image_box,
            "workerRequested": any("xlsx-worker" in item["url"].lower() for item in responses),
            "wasmRequested": any(".wasm" in item["url"].lower() for item in responses),
            "badResponses": [item for item in responses if item["status"] >= 400],
            "consoleErrors": [item for item in console if item["type"] == "error"],
            "pageErrors": page_errors,
            "requestFailures": request_failures,
        }
        page.screenshot(path=str(case_dir / "ready.png"), full_page=True)
        if result["mergedRegionCount"] != "1":
            raise AssertionError("exact merge range is missing")
        if result["freezeRow"] != "2" or result["freezeColumn"] != "2":
            raise AssertionError("multi-axis freeze panes are missing")
        if int(result["dataValidationCount"] or 0) < 1:
            raise AssertionError("data validation is missing from the fixture")
        if result["mergeSecondarySelectsAnchor"] != "A7":
            raise AssertionError("merged secondary cell did not select A7 anchor")
        if result["imageState"] != "ready" or not image_box:
            raise AssertionError("workbook image is not ready")
        if result["badResponses"] or result["consoleErrors"] or page_errors or request_failures:
            raise AssertionError("browser diagnostics contain failures")
        if mode == "worker" and (not result["workerRequested"] or not result["wasmRequested"]):
            raise AssertionError("Worker mode did not request Worker and WASM resources")
        return result
    finally:
        (case_dir / "console.json").write_text(json.dumps(console, ensure_ascii=False, indent=2), encoding="utf-8")
        (case_dir / "page-errors.json").write_text(json.dumps(page_errors, ensure_ascii=False, indent=2), encoding="utf-8")
        (case_dir / "request-failures.json").write_text(json.dumps(request_failures, ensure_ascii=False, indent=2), encoding="utf-8")
        (case_dir / "responses.json").write_text(json.dumps(responses, ensure_ascii=False, indent=2), encoding="utf-8")
        (case_dir / "downloads.json").write_text(json.dumps(downloads, ensure_ascii=False, indent=2), encoding="utf-8")
        context.close()


def run_attempt(browser, base_url: str, evidence_dir: Path, attempt: int) -> dict:
    attempt_dir = evidence_dir / f"attempt-{attempt}"
    attempt_dir.mkdir(parents=True, exist_ok=True)
    try:
        main = open_case(browser, base_url, attempt_dir, "main")
        worker = open_case(browser, base_url, attempt_dir, "worker")
        main_box = main["imageBox"]
        worker_box = worker["imageBox"]
        image_delta = {
            key: abs(float(main_box[key]) - float(worker_box[key]))
            for key in ["width", "height"]
        }
        if float(worker_box.get("width", 0)) > 0 and float(worker_box.get("height", 0)) > 0:
            if any(value > 1.5 for value in image_delta.values()):
                raise AssertionError(f"Worker/main image geometry differs: {image_delta}")
        elif worker.get("imageState") != "ready":
            raise AssertionError("Worker image did not decode")
        result = {"attempt": attempt, "status": "PASS", "main": main, "worker": worker, "imageDelta": image_delta}
    except Exception as cause:
        result = {"attempt": attempt, "status": "FAIL", "error": f"{type(cause).__name__}: {cause}"}
    (attempt_dir / "result.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--base-url", default=os.environ.get("BLACKBOX_BASE_URL", BASE_URL))
    parser.add_argument("--evidence-dir", default=str(OUTPUT))
    args = parser.parse_args()
    evidence_dir = Path(args.evidence_dir).resolve()
    evidence_dir.mkdir(parents=True, exist_ok=True)
    base_url = args.base_url.rstrip("/")

    preview: subprocess.Popen[bytes] | None = None
    preview_log = None
    if MANAGE_PREVIEW:
        if port_in_use():
            raise RuntimeError(f"port {PORT} is already in use")
        preview_log = (evidence_dir / "preview.log").open("wb")
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
        wait_for_preview(preview)

    try:
        with sync_playwright() as playwright:
            browser = launch_browser(playwright.chromium)
            first = run_attempt(browser, base_url, evidence_dir, 1)
            results = [first]
            if first["status"] != "PASS":
                results.append(run_attempt(browser, base_url, evidence_dir, 2))
            browser.close()
    finally:
        stop_process(preview)
        if preview_log is not None:
            preview_log.close()

    status = "PASS" if results[0]["status"] == "PASS" else "FLAKY" if results[-1]["status"] == "PASS" else "FAIL"
    summary = {"suite": "P4-FIDELITY-XLSX-WORKER-01", "status": status, "baseUrl": base_url, "attempts": results}
    (evidence_dir / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if status == "PASS" else 1


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        OUTPUT.mkdir(parents=True, exist_ok=True)
        (OUTPUT / "summary.json").write_text(
            json.dumps(
                {"suite": "P4-FIDELITY-XLSX-WORKER-01", "result": "BLOCKED", "error": str(error)},
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        print(f"BLOCKED: {error}", file=sys.stderr)
        raise SystemExit(2)
