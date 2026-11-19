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
        ROOT / "output" / "acceptance" / "p2-docx-editing",
    )
).resolve()
PORT = int(os.environ.get("CI_DOCX_EDITING_PREVIEW_PORT", "4179"))
BASE_URL = f"http://127.0.0.1:{PORT}"

CASES = (
    {"id": "DOCX-EDIT-COMPOSITION", "run": "composition"},
    {"id": "DOCX-EDIT-FORMAT-HISTORY", "run": "format_history"},
    {"id": "DOCX-EDIT-READONLY", "run": "readonly"},
    {"id": "DOCX-EDIT-TABLE-IMAGE", "run": "table_image"},
)


def write_json(path: Path, value: object) -> None:
    path.write_text(
        json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8"
    )


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


def read_snapshot(page, test_id: str) -> dict[str, object]:
    value = page.get_by_test_id(test_id).text_content() or ""
    return json.loads(value)


def wait_for_model(page, predicate, message: str, timeout_ms: int = 10_000):
    deadline = time.monotonic() + timeout_ms / 1000
    last: dict[str, object] | None = None
    while time.monotonic() < deadline:
        last = read_snapshot(page, "editor-model-snapshot")
        if predicate(last):
            return last
        page.wait_for_timeout(40)
    raise AssertionError(f"timed out waiting for {message}; last model={last}")


def navigate_ready(page) -> None:
    page.goto(f"{BASE_URL}/#/docx-editor", wait_until="networkidle")
    page.wait_for_function(
        """() => document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'
          && document.querySelector('[data-testid="editor-editable-state"]')?.dataset.state === 'editable'
          && document.querySelectorAll('[data-testid="editor-paragraph"]').length >= 3""",
        timeout=30_000,
    )
    # Acceptance drives the real toolbar. The former "自动化验收入口" buttons
    # called insertTableRow/setImageWrapMode directly, so the toolbar and the
    # pointer drag/resize paths were never exercised.
    page.get_by_test_id("docx-editor-toolbar").wait_for(state="visible", timeout=30_000)


def set_dom_text_range(locator, start: int, end: int) -> None:
    locator.evaluate(
        """(element, offsets) => {
          const textNodes = [];
          const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
          let current = walker.nextNode();
          while (current) {
            const parent = current.parentElement;
            if (!parent?.closest('[data-docx-numbering-label="true"]')) textNodes.push(current);
            current = walker.nextNode();
          }
          const resolve = (offset) => {
            let remaining = offset;
            for (const node of textNodes) {
              const length = node.textContent?.length || 0;
              if (remaining <= length) return {node, offset: remaining};
              remaining -= length;
            }
            const node = textNodes[textNodes.length - 1];
            if (!node) return {node: element, offset: 0};
            return {node, offset: node.textContent?.length || 0};
          };
          const selection = window.getSelection();
          const range = document.createRange();
          const anchor = resolve(offsets.start);
          const focus = resolve(offsets.end);
          element.focus();
          range.setStart(anchor.node, anchor.offset);
          range.setEnd(focus.node, focus.offset);
          selection.removeAllRanges();
          selection.addRange(range);
          element.dispatchEvent(new Event('selectionchange', { bubbles: true }));
        }""",
        {"start": start, "end": end},
    )


def wait_for_docx_selection_offsets(page, start: int, end: int, timeout_ms: int = 10_000) -> None:
    page.wait_for_function(
        """expected => {
          const raw = document.querySelector('[data-testid="editor-selection-snapshot"]')?.textContent;
          if (!raw) return false;
          const range = JSON.parse(raw).activeTextRange;
          return range?.start?.offset === expected.start && range?.end?.offset === expected.end;
        }""",
        arg={"start": start, "end": end},
        timeout=timeout_ms,
    )


def cell_locator(page, table_index: int, row: int = 0, cell: int = 0):
    return page.locator(
        f'[data-docx-table-cell-paragraph-host="true"]'
        f'[data-docx-table-index="{table_index}"]'
        f'[data-docx-table-row-index="{row}"]'
        f'[data-docx-table-cell-index="{cell}"]'
    ).first


def target_paragraph(model: dict[str, object]) -> tuple[int, dict[str, object]]:
    nodes = model.get("nodes", [])
    for index, node in enumerate(nodes):
        if (
            isinstance(node, dict)
            and node.get("type") == "paragraph"
            and "This paragraph is ready for editing" in str(node.get("text", ""))
        ):
            return index, node
    raise AssertionError("acceptance paragraph is missing from public model snapshot")


def dispatch_composition(page, node_index: int, marker: str, offset: int) -> None:
    paragraph = page.locator(
        f'[data-testid="editor-paragraph"][data-node-index="{node_index}"]'
    ).first
    paragraph.wait_for(state="visible")
    paragraph.evaluate(
        """(element, args) => {
          const current = element.textContent || '';
          const next = current.slice(0, args.offset) + args.marker + current.slice(args.offset);
          element.focus();
          element.dispatchEvent(new CompositionEvent('compositionstart', {
            bubbles: true,
            data: args.marker,
          }));
          element.textContent = next;
          element.dispatchEvent(new InputEvent('input', {
            bubbles: true,
            data: args.marker,
            inputType: 'insertCompositionText',
            isComposing: true,
          }));
          element.dispatchEvent(new CompositionEvent('compositionend', {
            bubbles: true,
            data: args.marker,
          }));
          element.dispatchEvent(new InputEvent('input', {
            bubbles: true,
            data: args.marker,
            inputType: 'insertText',
          }));
          element.blur();
        }""",
        {"marker": marker, "offset": offset},
    )


def run_composition(page, attempt_dir: Path) -> dict[str, object]:
    before = read_snapshot(page, "editor-model-snapshot")
    node_index, node = target_paragraph(before)
    original = str(node["text"])
    offset = original.index("editing") + len("editing")
    marker = "中英IME"
    expected = original[:offset] + marker + original[offset:]
    dispatch_composition(page, node_index, marker, offset)
    after = wait_for_model(
        page,
        lambda model: str(model["nodes"][node_index].get("text")) == expected,
        "composition text commit",
    )
    assert page.get_by_test_id("editor-history-state").get_attribute("data-can-undo") == "true"

    page.get_by_test_id("editor-undo").click()
    undone = wait_for_model(
        page,
        lambda model: str(model["nodes"][node_index].get("text")) == original,
        "composition undo",
    )
    assert page.get_by_test_id("editor-history-state").get_attribute("data-can-redo") == "true"

    page.get_by_test_id("editor-redo").click()
    redone = wait_for_model(
        page,
        lambda model: str(model["nodes"][node_index].get("text")) == expected,
        "composition redo",
    )
    write_json(
        attempt_dir / "models.json",
        {"before": before, "after": after, "undone": undone, "redone": redone},
    )
    return {
        "nodeIndex": node_index,
        "offset": offset,
        "marker": marker,
        "expected": expected,
    }


def bold_text(node: dict[str, object]) -> str:
    return "".join(
        str(run.get("text", ""))
        for run in node.get("runs", [])
        if isinstance(run, dict)
        and isinstance(run.get("style"), dict)
        and run["style"].get("bold") is True
    )


def run_format_history(page, attempt_dir: Path) -> dict[str, object]:
    before = read_snapshot(page, "editor-model-snapshot")
    node_index, node = target_paragraph(before)
    original = str(node["text"])
    start = original.index("editing")
    end = start + len("editing")
    target = page.locator(
        f'[data-testid="editor-paragraph"][data-node-index="{node_index}"]'
    ).first
    target.wait_for(state="visible")
    set_dom_text_range(target, start, end)
    wait_for_docx_selection_offsets(page, start, end)
    page.get_by_test_id("editor-bold").click()
    formatted = wait_for_model(
        page,
        lambda model: "editing" in bold_text(model["nodes"][node_index]),
        "partial bold formatting through the real toolbar",
    )
    selected = read_snapshot(page, "editor-selection-snapshot")
    active_range = selected.get("activeTextRange")
    assert isinstance(active_range, dict), selected
    assert active_range["start"]["offset"] < active_range["end"]["offset"], selected

    page.get_by_test_id("editor-undo").click()
    undone = wait_for_model(
        page,
        lambda model: "editing" not in bold_text(model["nodes"][node_index]),
        "format undo",
    )
    undo_selection = read_snapshot(page, "editor-selection-snapshot")
    assert undo_selection.get("activeTextRange") == active_range, undo_selection
    assert (
        undo_selection.get("historyRestoreRequest", {}).get("activeTextRange")
        == active_range
    ), undo_selection

    page.get_by_test_id("editor-redo").click()
    redone = wait_for_model(
        page,
        lambda model: "editing" in bold_text(model["nodes"][node_index]),
        "format redo",
    )
    redo_selection = read_snapshot(page, "editor-selection-snapshot")
    assert redo_selection.get("activeTextRange") == active_range, redo_selection
    assert (
        redo_selection.get("historyRestoreRequest", {}).get("activeTextRange")
        == active_range
    ), redo_selection
    write_json(
        attempt_dir / "format-history.json",
        {
            "before": before,
            "formatted": formatted,
            "undone": undone,
            "redone": redone,
            "selected": selected,
            "undoSelection": undo_selection,
            "redoSelection": redo_selection,
        },
    )
    return {
        "nodeIndex": node_index,
        "range": active_range,
        "boldText": "editing",
        "selection": {"start": start, "end": end},
    }


def run_readonly(page, attempt_dir: Path) -> dict[str, object]:
    before = read_snapshot(page, "editor-model-snapshot")
    node_index, node = target_paragraph(before)
    original = str(node["text"])
    checkbox = page.get_by_test_id("editor-readonly")
    checkbox.check()
    page.wait_for_function(
        "() => document.querySelector('[data-testid=editor-editable-state]')?.dataset.state === 'readonly'"
    )
    paragraph = page.locator(
        f'[data-testid="editor-paragraph"][data-node-index="{node_index}"]'
    ).first
    assert paragraph.get_attribute("contenteditable") != "true"

    dispatch_composition(page, node_index, "不得写入", 3)
    page.keyboard.press("Control+b")
    # The real toolbar is disabled by the read-only fieldset. Click through the
    # DOM so the test proves the command path refuses to mutate, not just that
    # Playwright refuses to click a disabled control.
    assert page.get_by_test_id("editor-bold").is_disabled()
    for test_id in ("editor-insert-table", "editor-insert-image"):
        page.get_by_test_id(test_id).evaluate("button => button.click()")
    page.wait_for_timeout(250)
    after = read_snapshot(page, "editor-model-snapshot")
    assert after == before, {"before": before, "after": after}
    assert str(after["nodes"][node_index].get("text")) == original
    assert page.get_by_test_id("editor-history-state").get_attribute("data-can-undo") == "false"
    write_json(attempt_dir / "readonly.json", {"before": before, "after": after})
    return {"nodeIndex": node_index, "modelUnchanged": True}


ACCEPTANCE_SVG = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="240" height="160">'
    '<rect width="240" height="160" fill="#2563eb"/></svg>'
).encode("utf-8")


def table_at(model: dict[str, object], table_index: int) -> dict[str, object] | None:
    nodes = model.get("nodes")
    if not isinstance(nodes, list) or table_index >= len(nodes):
        return None
    node = nodes[table_index]
    return node if isinstance(node, dict) and node.get("type") == "table" else None


def images_in(model: dict[str, object]) -> list[dict[str, object]]:
    return [
        run
        for node in model.get("nodes", [])
        if node.get("type") == "paragraph"
        for run in node.get("runs", [])
        if run.get("type") == "image"
    ]


def run_table_image(page, attempt_dir: Path) -> dict[str, object]:
    before = read_snapshot(page, "editor-model-snapshot")
    before_table_count = int(before.get("tableCount", 0))
    before_image_count = int(before.get("imageCount", 0))

    # Real toolbar button, not a test-only hook.
    page.get_by_test_id("editor-insert-table").click()
    with_table = wait_for_model(
        page,
        lambda model: model.get("tableCount") == before_table_count + 1,
        "table transaction",
    )
    table_index = max(
        index
        for index, node in enumerate(with_table["nodes"])
        if node.get("type") == "table"
    )
    table = table_at(with_table, table_index)
    assert table is not None
    assert len(table["rows"]) == 3, table
    assert all(len(row["cells"]) == 3 for row in table["rows"]), table
    assert table["style"]["layout"] == "fixed", table
    assert sum(table["style"]["columnWidthsTwips"]) == 7200, table

    # Select a real table cell, then extend the table through the toolbar.
    cell = cell_locator(page, table_index, 0, 0)
    cell.wait_for(state="visible", timeout=10_000)
    cell.click()
    page.get_by_test_id("editor-table-row-below").click()
    wait_for_model(
        page,
        lambda model: len((table_at(model, table_index) or {}).get("rows", [])) == 4,
        "toolbar row insertion",
    )
    cell_locator(page, table_index, 0, 0).click()
    page.get_by_test_id("editor-table-column-right").click()
    with_columns = wait_for_model(
        page,
        lambda model: all(
            len(row["cells"]) == 4
            for row in (table_at(model, table_index) or {}).get("rows", [])
        ),
        "toolbar column insertion",
    )
    extended = table_at(with_columns, table_index)
    assert extended is not None
    assert len(extended["rows"]) == 4, extended
    widths_before_drag = list(extended["style"]["columnWidthsTwips"])

    # Real pointer drag on the first column boundary.
    handle = page.locator(
        f'[data-docx-table-index="{table_index}"] .docx-table-column-handle'
    ).first
    handle.wait_for(state="visible", timeout=10_000)
    handle_box = handle.bounding_box()
    assert handle_box, "column resize handle has no box"
    handle_x = handle_box["x"] + handle_box["width"] / 2
    handle_y = handle_box["y"] + min(40, handle_box["height"] / 2)
    page.mouse.move(handle_x, handle_y)
    page.mouse.down()
    page.mouse.move(handle_x + 36, handle_y, steps=6)
    page.mouse.up()
    dragged = wait_for_model(
        page,
        lambda model: (
            (candidate := table_at(model, table_index)) is not None
            and candidate["style"]["columnWidthsTwips"][0] >= widths_before_drag[0] + 500
            and candidate["style"]["columnWidthsTwips"][1] <= widths_before_drag[1] - 500
        ),
        "pointer column resize",
    )
    resized_table = table_at(dragged, table_index)
    assert resized_table is not None
    assert sum(resized_table["style"]["columnWidthsTwips"]) == sum(widths_before_drag), resized_table

    # Real toolbar image picker (it creates the file input on demand).
    # Select a body paragraph first: an image inserted into a selected table
    # cell is inline, while the floating-image pointer path needs a paragraph.
    page.locator(
        '[data-testid="editor-paragraph"][contenteditable="true"]'
    ).filter(has_text="This paragraph is ready for editing").first.click()
    with page.expect_file_chooser(timeout=10_000) as chooser_info:
        page.get_by_test_id("editor-insert-image").click()
    chooser_info.value.set_files(
        {"name": "acceptance.svg", "mimeType": "image/svg+xml", "buffer": ACCEPTANCE_SVG}
    )
    with_image = wait_for_model(
        page,
        lambda model: model.get("imageCount") == before_image_count + 1,
        "image transaction",
    )
    images = images_in(with_image)
    assert len(images) == before_image_count + 1, images
    image = images[-1]
    assert image["widthPx"] == 240, image
    assert image["heightPx"] == 160, image
    assert image["floating"]["wrapType"] == "square", image
    assert image["floating"]["behindDocument"] is False, image

    floating = page.get_by_test_id("docx-floating-image").last
    floating.wait_for(state="visible", timeout=10_000)
    floating.click()
    page.get_by_test_id("docx-image-wrap-toolbar").wait_for(state="visible", timeout=10_000)
    page.get_by_test_id("docx-image-wrap-front").click()
    wait_for_model(
        page,
        lambda model: images_in(model)[-1]["floating"]["wrapType"] == "none",
        "image wrap mode change through the floating toolbar",
    )
    page.get_by_test_id("docx-image-wrap-square").click()
    wait_for_model(
        page,
        lambda model: images_in(model)[-1]["floating"]["wrapType"] == "square",
        "image wrap mode return to square",
    )

    # Real pointer drag on the floating image, then on its south-east handle.
    image_box = floating.bounding_box()
    assert image_box
    start_x = image_box["x"] + image_box["width"] / 2
    start_y = image_box["y"] + image_box["height"] / 2
    page.mouse.move(start_x, start_y)
    page.mouse.down()
    page.mouse.move(start_x + 24, start_y - 12, steps=4)
    page.mouse.up()
    moved = wait_for_model(
        page,
        lambda model: (
            bool(images_in(model))
            and abs(int(images_in(model)[-1]["floating"]["xPx"]) - 96) <= 2
            and abs(int(images_in(model)[-1]["floating"]["yPx"]) - 60) <= 2
        ),
        "pointer image move",
    )

    resize_handle = floating.locator(
        '.docx-image-resize-handle--se[data-testid="docx-image-resize-handle"]'
    )
    resize_box = resize_handle.bounding_box()
    assert resize_box, "image resize handle has no box"
    resize_x = resize_box["x"] + resize_box["width"] / 2
    resize_y = resize_box["y"] + resize_box["height"] / 2
    page.mouse.move(resize_x, resize_y)
    page.mouse.down()
    page.mouse.move(resize_x + 40, resize_y + 20, steps=4)
    page.mouse.up()
    resized = wait_for_model(
        page,
        lambda model: (
            bool(images_in(model))
            and abs(int(images_in(model)[-1].get("widthPx") or 0) - 280) <= 2
            and abs(int(images_in(model)[-1].get("heightPx") or 0) - 180) <= 2
        ),
        "pointer image resize",
    )
    assert not page.get_by_test_id("editor-status").inner_text().startswith("Unsupported:")
    assert page.get_by_test_id("editor-history-state").get_attribute("data-can-undo") == "true"
    write_json(
        attempt_dir / "table-image.json",
        {
            "before": before,
            "withTable": with_table,
            "withColumns": with_columns,
            "afterDrag": dragged,
            "withImage": with_image,
            "afterMove": moved,
            "afterResize": resized,
            "columnWidthsBeforeDrag": widths_before_drag,
            "columnWidthsAfterDrag": resized_table["style"]["columnWidthsTwips"],
        },
    )
    return {
        "tableRows": len(resized_table["rows"]),
        "tableColumns": len(resized_table["rows"][0]["cells"]),
        "columnWidthsBeforeDrag": widths_before_drag,
        "columnWidthsAfterDrag": resized_table["style"]["columnWidthsTwips"],
        "image": images_in(resized)[-1],
    }


RUNNERS = {
    "composition": run_composition,
    "format_history": run_format_history,
    "readonly": run_readonly,
    "table_image": run_table_image,
}


def run_attempt(browser, case: dict[str, str], attempt: int) -> dict[str, object]:
    attempt_dir = OUTPUT / case["id"] / f"attempt-{attempt}"
    attempt_dir.mkdir(parents=True, exist_ok=True)
    context = browser.new_context(viewport={"width": 1500, "height": 1000})
    context.tracing.start(screenshots=True, snapshots=True, sources=True)
    page = context.new_page()
    evidence = BrowserEvidence(page)
    try:
        navigate_ready(page)
        details = RUNNERS[case["run"]](page, attempt_dir)
        evidence.assert_clean()
        assert not evidence.events["downloads"], evidence.events["downloads"]
        page.screenshot(path=str(attempt_dir / "final.png"), full_page=True)
        result: dict[str, object] = {
            "id": case["id"],
            "attempt": attempt,
            "status": "PASS",
            "details": details,
        }
    except Exception as error:
        try:
            page.screenshot(path=str(attempt_dir / "failure.png"), full_page=True)
        except Exception:
            pass
        result = {
            "id": case["id"],
            "attempt": attempt,
            "status": "FAIL",
            "error": repr(error),
            "traceback": traceback.format_exc(),
        }
    finally:
        evidence.save(attempt_dir)
        context.tracing.stop(path=attempt_dir / "trace.zip")
        context.close()
    write_json(attempt_dir / "result.json", result)
    return result


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
            results: list[dict[str, object]] = []
            for case in CASES:
                attempts = [run_attempt(browser, case, 1)]
                if attempts[0]["status"] == "FAIL":
                    attempts.append(run_attempt(browser, case, 2))
                status = (
                    "PASS"
                    if attempts[0]["status"] == "PASS"
                    else "FLAKY"
                    if attempts[-1]["status"] == "PASS"
                    else "FAIL"
                )
                results.append(
                    {"id": case["id"], "status": status, "attempts": attempts}
                )
            browser.close()
        passed = all(item["status"] == "PASS" for item in results)
        write_json(
            OUTPUT / "summary.json",
            {
                "suite": "P2-DOCX-EDIT-01",
                "mode": "formal preview",
                "result": "PASS" if passed else "FAIL",
                "browser": browser_version,
                "browserSource": browser_source,
                "cases": results,
            },
        )
        return 0 if passed else 1
    finally:
        stop_process(preview)
        preview_log.close()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        OUTPUT.mkdir(parents=True, exist_ok=True)
        write_json(
            OUTPUT / "summary.json",
            {
                "suite": "P2-DOCX-EDIT-01",
                "result": "BLOCKED",
                "error": repr(error),
                "traceback": traceback.format_exc(),
            },
        )
        raise SystemExit(2)
