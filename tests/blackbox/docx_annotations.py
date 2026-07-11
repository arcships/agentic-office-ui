#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


def run_attempt(browser, base_url: str, evidence_dir: Path, attempt: int) -> dict:
    attempt_dir = evidence_dir / f"attempt-{attempt}"
    attempt_dir.mkdir(parents=True, exist_ok=True)
    console: list[dict] = []
    page_errors: list[str] = []
    request_failures: list[dict] = []
    responses: list[dict] = []
    downloads: list[str] = []
    context = browser.new_context(
        viewport={"width": 1440, "height": 900},
        record_har_path=str(attempt_dir / "network.har"),
    )
    page = context.new_page()
    page.on("console", lambda message: console.append({"type": message.type, "text": message.text}))
    page.on("pageerror", lambda error: page_errors.append(str(error)))
    page.on(
        "requestfailed",
        lambda request: request_failures.append({
            "url": request.url,
            "failure": request.failure,
        }),
    )
    page.on(
        "response",
        lambda response: responses.append({"url": response.url, "status": response.status}),
    )
    page.on("download", lambda download: downloads.append(download.suggested_filename))

    status = "PASS"
    error = None
    assertions: dict[str, object] = {}
    try:
        page.goto(f"{base_url}/#/docx-viewer", wait_until="networkidle")
        page.locator('[data-testid="docx-sample-select"]').select_option("review-comments.docx")
        page.wait_for_function(
            """() => document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'
              && document.querySelector('[data-testid="loaded-file"]')?.textContent?.includes('review-comments.docx')
              && document.querySelector('[data-testid="docx-viewer"]')?.dataset.state === 'ready'"""
        )
        page.locator('[data-testid="docx-show-tracked-changes"]').wait_for(state="visible")
        page.locator('[data-testid="docx-show-comments"]').wait_for(state="visible")

        tracked = page.locator('[data-testid="docx-show-tracked-changes"]')
        comments = page.locator('[data-testid="docx-show-comments"]')
        gutter_cards = page.locator(".docx-gutter-card")
        page.wait_for_function("() => document.querySelectorAll('.docx-gutter-card').length === 3")
        visible_text = page.locator('[data-testid="docx-viewer"]').inner_text()
        assertions["initialCardCount"] = gutter_cards.count()
        assertions["initialTrackedPressed"] = tracked.get_attribute("aria-pressed")
        assertions["initialCommentsPressed"] = comments.get_attribute("aria-pressed")
        assertions["authorsVisible"] = all(
            value in visible_text for value in ["Alex Reviewer", "Morgan Editor", "Priya Reviewer"]
        )
        assertions["commentVisible"] = "Please confirm this report includes the service-level summary." in visible_text
        page.screenshot(path=str(attempt_dir / "annotations-visible.png"), full_page=True)

        comments.click()
        page.wait_for_function(
            """() => document.querySelector('[data-testid="docx-show-comments"]')?.getAttribute('aria-pressed') === 'false'
              && document.querySelectorAll('.docx-gutter-card--comment').length === 0"""
        )
        assertions["commentToggleHidesCards"] = page.locator(".docx-gutter-card--comment").count() == 0
        assertions["trackedCardsRemain"] = page.locator(".docx-gutter-card:not(.docx-gutter-card--comment)").count() == 2
        page.screenshot(path=str(attempt_dir / "comments-hidden.png"), full_page=True)

        tracked.click()
        page.wait_for_function("() => document.querySelectorAll('.docx-gutter-card').length === 0")
        assertions["trackedToggleHidesCards"] = gutter_cards.count() == 0
        page.screenshot(path=str(attempt_dir / "annotations-hidden.png"), full_page=True)

        comments.click()
        tracked.click()
        page.wait_for_function("() => document.querySelectorAll('.docx-gutter-card').length === 3")
        assertions["togglesRecover"] = gutter_cards.count() == 3
        assertions["workerRequested"] = any("worker" in item["url"].lower() for item in responses)
        assertions["wasmRequested"] = any(".wasm" in item["url"].lower() for item in responses)
        assertions["badResponses"] = [item for item in responses if item["status"] >= 400]
        assertions["consoleErrors"] = [item for item in console if item["type"] == "error"]

        required = [
            assertions["initialCardCount"] == 3,
            assertions["initialTrackedPressed"] == "true",
            assertions["initialCommentsPressed"] == "true",
            assertions["authorsVisible"] is True,
            assertions["commentVisible"] is True,
            assertions["commentToggleHidesCards"] is True,
            assertions["trackedCardsRemain"] is True,
            assertions["trackedToggleHidesCards"] is True,
            assertions["togglesRecover"] is True,
            assertions["workerRequested"] is True,
            assertions["wasmRequested"] is True,
            not assertions["badResponses"],
            not assertions["consoleErrors"],
            not page_errors,
            not request_failures,
        ]
        if not all(required):
            raise AssertionError("DOCX annotation acceptance assertions failed")
    except (AssertionError, PlaywrightTimeoutError, Exception) as cause:
        status = "FAIL"
        error = f"{type(cause).__name__}: {cause}"
        try:
            page.screenshot(path=str(attempt_dir / "failure.png"), full_page=True)
        except Exception:
            pass
    finally:
        (attempt_dir / "console.json").write_text(json.dumps(console, ensure_ascii=False, indent=2), encoding="utf-8")
        (attempt_dir / "page-errors.json").write_text(json.dumps(page_errors, ensure_ascii=False, indent=2), encoding="utf-8")
        (attempt_dir / "request-failures.json").write_text(json.dumps(request_failures, ensure_ascii=False, indent=2), encoding="utf-8")
        (attempt_dir / "responses.json").write_text(json.dumps(responses, ensure_ascii=False, indent=2), encoding="utf-8")
        (attempt_dir / "downloads.json").write_text(json.dumps(downloads, ensure_ascii=False, indent=2), encoding="utf-8")
        context.close()

    result = {"attempt": attempt, "status": status, "error": error, "assertions": assertions}
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
    summary = {"suite": "P4-FIDELITY-DOCX-01", "status": status, "attempts": results}
    (evidence_dir / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if status == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
