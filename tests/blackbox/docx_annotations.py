#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import signal
import socket
import subprocess
import time
import urllib.request
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parents[2]
DEFAULT_PORT = int(os.environ.get("CI_ANNOTATIONS_PORT", "4187"))


def port_in_use(port: int) -> bool:
    with socket.socket() as sock:
        return sock.connect_ex(("127.0.0.1", port)) == 0


def wait_for_preview(process: subprocess.Popen[bytes], base_url: str) -> None:
    deadline = time.monotonic() + 60
    while time.monotonic() < deadline:
        if process.poll() is not None:
            raise RuntimeError(f"preview exited early with {process.returncode}")
        try:
            with urllib.request.urlopen(base_url, timeout=1) as response:
                if response.status == 200:
                    return
        except Exception:
            time.sleep(0.2)
    raise TimeoutError(f"preview did not become ready at {base_url}")


def start_preview(evidence_dir: Path, port: int) -> tuple[subprocess.Popen[bytes], object]:
    preview_log = (evidence_dir / "preview.log").open("wb")
    process = subprocess.Popen(
        [
            "pnpm",
            "--filter",
            "demo",
            "preview",
            "--host",
            "127.0.0.1",
            "--port",
            str(port),
        ],
        cwd=ROOT,
        stdout=preview_log,
        stderr=subprocess.STDOUT,
        start_new_session=True,
    )
    return process, preview_log


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


GUTTER_LINE_HEIGHT_PX = 24


def gutter_alignment(page) -> dict:
    """Measure every gutter card against the anchor marker it belongs to.

    Issue #53: the browser suite used to assert only that cards exist. It now
    asserts that each comment has exactly one card, that the card's number
    matches its inline marker, and that the card sits on its anchor (or is pushed
    down by the card above it) instead of in a fixed list at the top of the page.
    """
    return page.evaluate(
        """() => {
          const cards = [...document.querySelectorAll('.docx-gutter-card')];
          const markers = [...document.querySelectorAll('[data-docx-comment-marker]')];
          const markerById = new Map();
          for (const marker of markers) {
            markerById.set(
              marker.getAttribute('data-docx-gutter-marker-id'),
              marker
            );
          }
          return cards.map((card) => {
            const id = card.getAttribute('data-docx-gutter-card-id') || '';
            const marker = markerById.get(id) || null;
            const cardRect = card.getBoundingClientRect();
            const markerRect = marker ? marker.getBoundingClientRect() : null;
            const pageElement = card.closest('[data-docx-page-surface]');
            return {
              id,
              index: Number(card.getAttribute('data-docx-gutter-card-index')),
              markerNumber: marker
                ? Number(marker.getAttribute('data-docx-comment-marker'))
                : null,
              cardTop: cardRect.top,
              cardBottom: cardRect.bottom,
              markerTop: markerRect ? markerRect.top : null,
              pageIndex: pageElement
                ? pageElement.getAttribute('data-docx-page-index')
                : null,
            };
          });
        }"""
    )


def assert_gutter_alignment(page, report: dict, prefix: str = "") -> dict:
    cards = gutter_alignment(page)
    comment_cards = [card for card in cards if card["id"].startswith("comment:")]
    if not comment_cards:
        raise AssertionError("the gutter rendered no comment cards")
    ids = [card["id"] for card in comment_cards]
    if len(ids) != len(set(ids)):
        raise AssertionError(f"a comment is listed more than once: {ids}")
    for card in comment_cards:
        if card["markerTop"] is None:
            raise AssertionError(f"{card['id']} has no inline anchor marker")
        if card["markerNumber"] != card["index"]:
            raise AssertionError(
                f"{card['id']}: marker {card['markerNumber']} != card {card['index']}"
            )
        if card["cardTop"] < card["markerTop"] - GUTTER_LINE_HEIGHT_PX:
            raise AssertionError(
                f"{card['id']}: card top {card['cardTop']} sits above its anchor "
                f"{card['markerTop']}"
            )
    # Cards stacked under a taller card may be pushed down; every other card has
    # to start within one line of its anchor.
    by_page: dict[str, list[dict]] = {}
    for card in cards:
        by_page.setdefault(str(card["pageIndex"]), []).append(card)
    for page_cards in by_page.values():
        previous_bottom = None
        for card in page_cards:
            if card["markerTop"] is None:
                continue
            within_a_line = (
                card["cardTop"] <= card["markerTop"] + GUTTER_LINE_HEIGHT_PX
            )
            pushed_down = previous_bottom is not None and card["cardTop"] >= previous_bottom
            if not within_a_line and not pushed_down:
                raise AssertionError(
                    f"{card['id']}: card top {card['cardTop']} is not within one line "
                    f"of anchor {card['markerTop']} and was not pushed down"
                )
            previous_bottom = card["cardBottom"]
    report[f"{prefix}CommentCardIds"] = ids
    report[f"{prefix}CommentCardTops"] = [round(card["cardTop"], 1) for card in comment_cards]
    report[f"{prefix}CommentMarkerTops"] = [
        round(card["markerTop"], 1) for card in comment_cards
    ]
    return report


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

        # Issue #53: cards must sit on their anchors, not in a fixed list. One
        # comment cannot prove that, so the multi-comment fixture is loaded next.
        assert_gutter_alignment(page, assertions, "singleComment")
        page.screenshot(path=str(attempt_dir / "annotations-aligned.png"), full_page=True)

        page.locator('[data-testid="docx-sample-select"]').select_option("review-threads.docx")
        page.wait_for_function(
            """() => document.querySelector('[data-testid="loaded-file"]')?.textContent?.includes('review-threads.docx')
              && document.querySelector('[data-testid="docx-viewer"]')?.dataset.state === 'ready'
              && document.querySelectorAll('.docx-gutter-card--comment').length >= 8""",
            timeout=30_000,
        )
        thread_text = page.locator('[data-testid="docx-viewer"]').inner_text()
        assertions["threadRootVisible"] = "Root comment" in thread_text
        assertions["resolvedThreadHidden"] = "Resolved root:" not in thread_text
        assertions["headerCommentVisible"] = "Header comment:" in thread_text
        assertions["cellCommentVisible"] = "Cell comment:" in thread_text
        assert_gutter_alignment(page, assertions, "threads")
        page.screenshot(path=str(attempt_dir / "threads-aligned.png"), full_page=True)

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
            assertions["threadRootVisible"] is True,
            assertions["resolvedThreadHidden"] is True,
            assertions["headerCommentVisible"] is True,
            assertions["cellCommentVisible"] is True,
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
    parser.add_argument("--base-url", default=os.environ.get("CI_ANNOTATIONS_BASE_URL") or None)
    parser.add_argument(
        "--evidence-dir",
        default=os.environ.get("BLACKBOX_EVIDENCE_DIR") or None,
    )
    args = parser.parse_args()
    if not args.evidence_dir:
        parser.error("--evidence-dir or BLACKBOX_EVIDENCE_DIR is required")
    evidence_dir = Path(args.evidence_dir).resolve()
    evidence_dir.mkdir(parents=True, exist_ok=True)

    # The formal preview is started here so the suite can register this case
    # without a separately managed server; an explicit --base-url reuses one.
    preview = None
    preview_log = None
    base_url = args.base_url
    if base_url is None:
        if port_in_use(DEFAULT_PORT):
            raise RuntimeError(f"port {DEFAULT_PORT} is already in use")
        base_url = f"http://127.0.0.1:{DEFAULT_PORT}"
        preview, preview_log = start_preview(evidence_dir, DEFAULT_PORT)
        try:
            wait_for_preview(preview, base_url)
        except Exception:
            stop_process(preview)
            preview_log.close()
            raise

    results: list[dict] = []
    launch_error = None
    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=True)
            first = run_attempt(browser, base_url.rstrip("/"), evidence_dir, 1)
            results.append(first)
            if first["status"] != "PASS":
                results.append(run_attempt(browser, base_url.rstrip("/"), evidence_dir, 2))
            browser.close()
    except Exception as cause:
        launch_error = f"{type(cause).__name__}: {cause}"
    finally:
        if preview is not None:
            stop_process(preview)
            preview_log.close()

    if launch_error is not None:
        summary = {
            "suite": "P4-FIDELITY-DOCX-01",
            "status": "FAIL",
            "error": launch_error,
            "attempts": results,
        }
        (evidence_dir / "summary.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        return 1

    status = "PASS" if results[0]["status"] == "PASS" else "FLAKY" if results[-1]["status"] == "PASS" else "FAIL"
    summary = {"suite": "P4-FIDELITY-DOCX-01", "status": status, "attempts": results}
    (evidence_dir / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if status == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
