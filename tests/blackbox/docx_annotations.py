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


# The marker chip is centered on the anchor line, so a card that sits on its
# anchor may start slightly above or below the chip's top edge.
GUTTER_ANCHOR_TOLERANCE_PX = 28


def gutter_snapshot(page) -> dict:
    """Cards, anchor markers and the geometry of the anchors they belong to."""
    return page.evaluate(
        """() => {
          const rect = (element) =>
            element ? element.getBoundingClientRect().toJSON() : null;
          const cards = [...document.querySelectorAll('.docx-gutter-card')];
          const markerById = new Map();
          for (const marker of document.querySelectorAll('[data-docx-gutter-marker-id]')) {
            markerById.set(
              marker.getAttribute('data-docx-gutter-marker-id'),
              marker
            );
          }
          const findByText = (selector, text) =>
            [...document.querySelectorAll(selector)].find((element) =>
              (element.innerText || '').includes(text)
            ) || null;
          return {
            cards: cards.map((card) => {
              const marker =
                markerById.get(card.getAttribute('data-docx-gutter-card-id')) || null;
              return {
                id: card.getAttribute('data-docx-gutter-card-id'),
                index: Number(card.getAttribute('data-docx-gutter-card-index')),
                isComment: card.classList.contains('docx-gutter-card--comment'),
                isResolved: card.classList.contains('docx-gutter-card--resolved'),
                text: card.innerText || '',
                rect: rect(card),
                marker: marker
                  ? {
                      number: marker.getAttribute('data-docx-comment-marker'),
                      rect: rect(marker),
                    }
                  : null,
              };
            }),
            commentCardCount:
              document.querySelectorAll('.docx-gutter-card--comment').length,
            resolvedCards:
              document.querySelectorAll('.docx-gutter-card--resolved').length,
            anchors: {
              crossStart: rect(
                findByText('[data-docx-paragraph-host="true"]', 'Cross paragraph range begins')
              ),
              crossEnd: rect(
                findByText('[data-docx-paragraph-host="true"]', 'and closes at this point')
              ),
              cellAmount: rect(
                findByText('[data-docx-table-cell="true"]', '1,200.00')
              ),
              header: rect(document.querySelector('[data-docx-page-header]')),
            },
            viewerText: document.querySelector('[data-testid="docx-viewer"]')?.innerText || '',
          };
        }"""
    )


def card_by_text(snapshot: dict, text: str) -> dict:
    for card in snapshot["cards"]:
        if text in card["text"]:
            return card
    raise AssertionError(
        f"no gutter card contains {text!r}: "
        f"{[card['text'][:60] for card in snapshot['cards']]}"
    )


def band_contains(rect: dict | None, y: float, padding: float) -> bool:
    if rect is None:
        return False
    return rect["top"] - padding <= y <= rect["bottom"] + padding


def marker_top(card: dict) -> float:
    if not card["marker"] or card["marker"]["rect"] is None:
        raise AssertionError(f"{card['id']} has no anchor marker")
    return card["marker"]["rect"]["top"]


def assert_gutter_alignment(snapshot: dict, report: dict, prefix: str = "") -> dict:
    """Every comment has one card whose number matches its inline anchor marker.

    Issue #53: the suite used to assert only that a card exists. It now asserts
    one card per comment, marker/card number equality, and that the card starts
    on its anchor instead of in a fixed list at the top of the page.
    """
    comment_cards = [card for card in snapshot["cards"] if card["isComment"]]
    if not comment_cards:
        raise AssertionError("the gutter rendered no comment cards")
    ids = [card["id"] for card in comment_cards]
    if len(ids) != len(set(ids)):
        raise AssertionError(f"a comment is listed more than once: {ids}")
    if snapshot["commentCardCount"] != len(comment_cards):
        raise AssertionError(
            f"comment card count {snapshot['commentCardCount']} != {len(comment_cards)}"
        )
    for card in comment_cards:
        number = card["marker"]["number"] if card["marker"] else None
        if number is None:
            raise AssertionError(f"{card['id']} has no comment anchor marker")
        if int(number) != int(card["index"]):
            raise AssertionError(
                f"{card['id']}: marker {number} != card {card['index']}"
            )

    # Cards are laid out by anchor position; a card that collides with a taller
    # card above it may be pushed down, but never above its own anchor.
    previous_bottom: float | None = None
    for card in snapshot["cards"]:
        if card["marker"] is None or card["marker"]["rect"] is None:
            continue
        card_top = card["rect"]["top"]
        anchor_top = card["marker"]["rect"]["top"]
        if card_top < anchor_top - GUTTER_ANCHOR_TOLERANCE_PX:
            raise AssertionError(
                f"{card['id']}: card top {card_top:.1f} sits above its anchor "
                f"{anchor_top:.1f}"
            )
        within_a_line = card_top <= anchor_top + GUTTER_ANCHOR_TOLERANCE_PX
        pushed_down = previous_bottom is not None and card_top >= previous_bottom - 1
        if not within_a_line and not pushed_down:
            raise AssertionError(
                f"{card['id']}: card top {card_top:.1f} is not within one line of "
                f"anchor {anchor_top:.1f} and was not pushed down"
            )
        previous_bottom = max(previous_bottom or 0, card["rect"]["bottom"])

    report[f"{prefix}CommentCardIds"] = ids
    report[f"{prefix}CommentCardTops"] = [
        round(card["rect"]["top"], 1) for card in comment_cards
    ]
    report[f"{prefix}CommentMarkerTops"] = [
        round(marker_top(card), 1) for card in comment_cards
    ]
    return report


def assert_thread_fixture(snapshot: dict, report: dict) -> dict:
    """Threads, cross-paragraph ranges, cell/header anchors and resolved state."""
    # 11 comments minus the two in the resolved thread (root + reply).
    if snapshot["commentCardCount"] != 9:
        raise AssertionError(
            f"expected 9 visible comment cards, got {snapshot['commentCardCount']}"
        )
    if snapshot["resolvedCards"] != 0:
        raise AssertionError(
            f"the resolved thread must stay hidden, got {snapshot['resolvedCards']} cards"
        )
    viewer_text = snapshot["viewerText"]
    if "Resolved root:" in viewer_text or "Resolved reply:" in viewer_text:
        raise AssertionError("a resolved thread comment is rendered by default")

    # Thread: root and both replies are listed, and the replies share the root's
    # anchor. Ordinals stay in document order (root, reply one, reply two).
    root = card_by_text(snapshot, "Root comment")
    reply_one = card_by_text(snapshot, "Reply one")
    reply_two = card_by_text(snapshot, "Reply two")
    if [root["index"], reply_one["index"], reply_two["index"]] != [1, 2, 3]:
        raise AssertionError(
            "thread cards are not numbered in document order: "
            f"{[root['index'], reply_one['index'], reply_two['index']]}"
        )
    root_top = marker_top(root)
    for reply in (reply_one, reply_two):
        if abs(marker_top(reply) - root_top) > 3:
            raise AssertionError(
                f"{reply['id']} is not anchored to the thread root: "
                f"{marker_top(reply):.1f} vs {root_top:.1f}"
            )

    # Two comments on one paragraph share its line.
    left = card_by_text(snapshot, "Left half of the same paragraph")
    right = card_by_text(snapshot, "Right half of the same paragraph")
    if abs(marker_top(left) - marker_top(right)) > 3:
        raise AssertionError(
            "the two comments on one paragraph are not anchored to the same line"
        )

    # A range crossing three paragraphs anchors at the paragraph that starts it.
    cross = card_by_text(snapshot, "Range crosses the three paragraphs above")
    anchors = snapshot["anchors"]
    cross_top = marker_top(cross)
    if not band_contains(anchors["crossStart"], cross_top, GUTTER_ANCHOR_TOLERANCE_PX):
        raise AssertionError(
            "the cross-paragraph card is not anchored to the range start: "
            f"{cross_top:.1f} vs {anchors['crossStart']}"
        )
    if band_contains(anchors["crossEnd"], cross_top, 4):
        raise AssertionError(
            "the cross-paragraph card is anchored to the range end instead of its start"
        )

    # A cell comment anchors inside the table cell that holds its range.
    cell_card = card_by_text(snapshot, "Cell comment")
    if not band_contains(
        anchors["cellAmount"], marker_top(cell_card), GUTTER_ANCHOR_TOLERANCE_PX
    ):
        raise AssertionError(
            f"the cell comment is not anchored to its table cell: "
            f"{marker_top(cell_card):.1f} vs {anchors['cellAmount']}"
        )

    # A header comment anchors in the header band and carries its part name.
    header_card = card_by_text(snapshot, "Header comment")
    if not str(header_card["id"]).startswith("comment:header:"):
        raise AssertionError(f"header comment id is not region-tagged: {header_card['id']}")
    if not band_contains(
        anchors["header"], marker_top(header_card), GUTTER_ANCHOR_TOLERANCE_PX
    ):
        raise AssertionError(
            f"the header comment is not anchored to the header: "
            f"{marker_top(header_card):.1f} vs {anchors['header']}"
        )

    report["threadCardIds"] = [root["id"], reply_one["id"], reply_two["id"]]
    report["crossParagraphAnchorTop"] = round(cross_top, 1)
    report["cellCommentAnchorTop"] = round(marker_top(cell_card), 1)
    report["headerCommentAnchorTop"] = round(marker_top(header_card), 1)
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
        # Comments/edits toggles live in the upstream DOCX actions menu.
        def open_docx_actions() -> None:
            page.locator('[data-testid="docx-toolbar-actions"]').click()
            page.locator('[data-testid="docx-actions-menu"]').wait_for(state="visible")

        open_docx_actions()
        page.locator('[data-testid="docx-show-tracked-changes"]').wait_for(state="visible")
        page.locator('[data-testid="docx-show-comments"]').wait_for(state="visible")

        tracked = page.locator('[data-testid="docx-show-tracked-changes"]')
        comments = page.locator('[data-testid="docx-show-comments"]')
        gutter_cards = page.locator(".docx-gutter-card")
        # The insertion renders inline (Word's inline revision markup) and is
        # deliberately not ballooned, so the fixture shows one comment card and
        # one deleted-revision card.
        page.wait_for_function("() => document.querySelectorAll('.docx-gutter-card').length === 2")
        visible_text = page.locator('[data-testid="docx-viewer"]').inner_text()
        assertions["initialCardCount"] = gutter_cards.count()
        assertions["initialTrackedPressed"] = tracked.get_attribute("aria-checked")
        assertions["initialCommentsPressed"] = comments.get_attribute("aria-checked")
        assertions["authorsVisible"] = all(
            value in visible_text for value in ["Alex Reviewer", "Priya Reviewer"]
        )
        assertions["commentVisible"] = "Please confirm this report includes the service-level summary." in visible_text
        # The insertion has no card; it must still be visible as inline markup.
        assertions["insertionVisibleInline"] = "two business days" in visible_text
        page.screenshot(path=str(attempt_dir / "annotations-visible.png"), full_page=True)

        open_docx_actions()
        page.locator('[data-testid="docx-show-comments"]').click()
        page.wait_for_function(
            "() => document.querySelectorAll('.docx-gutter-card--comment').length === 0"
        )
        assertions["commentToggleHidesCards"] = page.locator(".docx-gutter-card--comment").count() == 0
        assertions["trackedCardsRemain"] = page.locator(".docx-gutter-card:not(.docx-gutter-card--comment)").count() == 1
        page.screenshot(path=str(attempt_dir / "comments-hidden.png"), full_page=True)

        open_docx_actions()
        page.locator('[data-testid="docx-show-tracked-changes"]').click()
        page.wait_for_function("() => document.querySelectorAll('.docx-gutter-card').length === 0")
        assertions["trackedToggleHidesCards"] = gutter_cards.count() == 0
        page.screenshot(path=str(attempt_dir / "annotations-hidden.png"), full_page=True)

        open_docx_actions()
        page.locator('[data-testid="docx-show-comments"]').click()
        open_docx_actions()
        page.locator('[data-testid="docx-show-tracked-changes"]').click()
        page.wait_for_function("() => document.querySelectorAll('.docx-gutter-card').length === 2")
        assertions["togglesRecover"] = gutter_cards.count() == 2
        assertions["workerRequested"] = any("worker" in item["url"].lower() for item in responses)
        assertions["wasmRequested"] = any(".wasm" in item["url"].lower() for item in responses)
        assertions["badResponses"] = [item for item in responses if item["status"] >= 400]
        assertions["consoleErrors"] = [item for item in console if item["type"] == "error"]

        # Issue #53: cards must sit on their anchors, not in a fixed list. One
        # comment cannot prove that, so the multi-comment fixture is loaded next.
        single_comment = gutter_snapshot(page)
        assert_gutter_alignment(single_comment, assertions, "singleComment")
        assertions["singleCommentCardIds"] = [
            card["id"] for card in single_comment["cards"] if card["isComment"]
        ]
        page.screenshot(path=str(attempt_dir / "annotations-aligned.png"), full_page=True)

        page.locator('[data-testid="docx-sample-select"]').select_option("review-threads.docx")
        page.wait_for_function(
            """() => document.querySelector('[data-testid="loaded-file"]')?.textContent?.includes('review-threads.docx')
              && document.querySelector('[data-testid="docx-viewer"]')?.dataset.state === 'ready'
              && document.querySelectorAll('.docx-gutter-card--comment').length === 9""",
            timeout=30_000,
        )
        threads = gutter_snapshot(page)
        assert_gutter_alignment(threads, assertions, "threads")
        assert_thread_fixture(threads, assertions)
        thread_text = threads["viewerText"]
        assertions["threadRootVisible"] = "Root comment" in thread_text
        assertions["resolvedThreadHidden"] = "Resolved root:" not in thread_text
        assertions["headerCommentVisible"] = "Header comment:" in thread_text
        assertions["cellCommentVisible"] = "Cell comment:" in thread_text
        assertions["crossParagraphVisible"] = "Range crosses the three paragraphs above." in thread_text
        page.screenshot(path=str(attempt_dir / "threads-aligned.png"), full_page=True)

        required = [
            assertions["initialCardCount"] == 2,
            assertions["initialTrackedPressed"] == "true",
            assertions["initialCommentsPressed"] == "true",
            assertions["authorsVisible"] is True,
            assertions["commentVisible"] is True,
            assertions["insertionVisibleInline"] is True,
            assertions["commentToggleHidesCards"] is True,
            assertions["trackedCardsRemain"] is True,
            assertions["trackedToggleHidesCards"] is True,
            assertions["togglesRecover"] is True,
            assertions["threadRootVisible"] is True,
            assertions["resolvedThreadHidden"] is True,
            assertions["headerCommentVisible"] is True,
            assertions["cellCommentVisible"] is True,
            assertions["crossParagraphVisible"] is True,
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
