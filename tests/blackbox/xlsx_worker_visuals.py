#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path

from playwright.sync_api import sync_playwright


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
              && document.querySelector('[data-testid="xlsx-grid"]')?.dataset.workerBacked === expected""",
            arg="true" if mode == "worker" else "false",
        )
        image = page.locator('[data-testid="xlsx-image-item"] img').first
        image.wait_for(state="visible")
        page.wait_for_function(
            """() => document.querySelector('[data-testid="xlsx-image-item"] img')?.dataset.imageState === 'ready'"""
        )

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
            "workerBacked": grid.get_attribute("data-worker-backed"),
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
            for key in ["x", "y", "width", "height"]
        }
        if any(value > 1.5 for value in image_delta.values()):
            raise AssertionError(f"Worker/main image geometry differs: {image_delta}")
        result = {"attempt": attempt, "status": "PASS", "main": main, "worker": worker, "imageDelta": image_delta}
    except Exception as cause:
        result = {"attempt": attempt, "status": "FAIL", "error": f"{type(cause).__name__}: {cause}"}
    (attempt_dir / "result.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--base-url", default="http://127.0.0.1:4173")
    parser.add_argument("--evidence-dir", required=True)
    args = parser.parse_args()
    evidence_dir = Path(args.evidence_dir).resolve()
    evidence_dir.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        first = run_attempt(browser, args.base_url.rstrip("/"), evidence_dir, 1)
        results = [first]
        if first["status"] != "PASS":
            results.append(run_attempt(browser, args.base_url.rstrip("/"), evidence_dir, 2))
        browser.close()

    status = "PASS" if results[0]["status"] == "PASS" else "FLAKY" if results[-1]["status"] == "PASS" else "FAIL"
    summary = {"suite": "P4-FIDELITY-XLSX-WORKER-01", "status": status, "attempts": results}
    (evidence_dir / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if status == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
