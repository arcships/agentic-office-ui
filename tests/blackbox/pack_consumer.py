from __future__ import annotations

import hashlib
import json
from pathlib import Path
import sys
from urllib.parse import unquote, urlsplit

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import Page, sync_playwright

from browser_evidence import BrowserEvidence


base_url = sys.argv[1].rstrip("/")
evidence_dir = Path(sys.argv[2]).resolve()
consumer_dir = Path(sys.argv[3]).resolve()
runtime_evidence_path = Path(sys.argv[4]).resolve()
browser_name = sys.argv[5] if len(sys.argv) > 5 else "chromium"
evidence_dir.mkdir(parents=True, exist_ok=True)


def write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


def load_runtime_evidence() -> tuple[dict[str, object], str | None]:
    try:
        value = json.loads(runtime_evidence_path.read_text(encoding="utf-8"))
        if not isinstance(value, dict):
            raise ValueError("runtime evidence must be an object")
        return value, None
    except Exception as error:  # evidence must survive malformed input
        return {}, repr(error)


runtime_evidence, runtime_evidence_error = load_runtime_evidence()
expected_resources = runtime_evidence.get("expectedResources", [])
built_runtime_files = runtime_evidence.get("builtRuntimeFiles", [])


def launch_browser(playwright):
    if browser_name in {"firefox", "webkit"}:
        browser_type = getattr(playwright, browser_name)
        return browser_type.launch(headless=True), f"bundled-{browser_name}"
    if browser_name != "chromium":
        raise RuntimeError(f"unsupported browser: {browser_name}")
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


def served_file(response: dict[str, object]) -> Path:
    relative = unquote(urlsplit(str(response["url"])).path).lstrip("/")
    candidate = (consumer_dir / "dist" / relative).resolve()
    candidate.relative_to(consumer_dir)
    return candidate


def wait_for_pdf_page(page: Page, page_number: int) -> None:
    page.wait_for_function(
        """expected => {
          const viewer = document.querySelector('[data-testid="pdf-document"]');
          const stage = document.querySelector('[data-testid="pdf-page"]');
          const image = document.querySelector('[data-testid="pdf-page-image"]');
          return Number(viewer?.dataset.activePage) === expected &&
            Number(stage?.dataset.pageNumber) === expected &&
            stage?.dataset.renderState === 'rendered' && image instanceof HTMLImageElement &&
            image.complete && image.naturalWidth > 0 && image.naturalHeight > 0;
        }""",
        arg=page_number,
        timeout=45_000,
    )


def pdf_image(page: Page) -> dict[str, object]:
    return page.locator('[data-testid="pdf-page-image"]').evaluate(
        """image => ({
          src: image.currentSrc || image.src,
          naturalWidth: image.naturalWidth,
          naturalHeight: image.naturalHeight,
          displayedWidth: image.getBoundingClientRect().width,
          displayedHeight: image.getBoundingClientRect().height,
        })"""
    )


def runtime_proof(events: dict[str, list[dict[str, object]]]) -> tuple[dict[str, object], dict[str, bool]]:
    responses = events["responses"]
    workers = [
        item
        for item in responses
        if (
            "docx-import-worker" in str(item["url"])
            or "xlsx-worker" in str(item["url"])
        )
        and int(item["status"]) == 200
    ]
    wasm = [
        item
        for item in responses
        if str(item["url"]).split("?", 1)[0].endswith(".wasm")
        and int(item["status"]) == 200
    ]
    served_by_path: dict[str, dict[str, object]] = {}
    resource_error = None
    try:
        for response in workers + wasm:
            file_path = served_file(response)
            relative = file_path.relative_to(consumer_dir).as_posix()
            served_by_path[relative] = {
                "url": response["url"],
                "path": relative,
                "exists": file_path.is_file(),
                "sha256": hashlib.sha256(file_path.read_bytes()).hexdigest()
                if file_path.is_file()
                else None,
                "contentType": response["contentType"],
            }
    except Exception as error:
        resource_error = repr(error)

    built_by_path = {
        str(item.get("path")): item.get("sha256")
        for item in built_runtime_files
        if isinstance(item, dict)
    }
    served_resources = list(served_by_path.values())
    served_worker_paths = {
        str(item["path"])
        for item in served_resources
        if "worker" in Path(str(item["path"])).name
    }
    served_wasm_paths = {
        str(item["path"])
        for item in served_resources
        if str(item["path"]).endswith(".wasm")
    }
    checks = {
        "twoWorkers": len(served_worker_paths) == 2,
        "threeWasm": len(served_wasm_paths) == 3,
        "wasmMime": bool(wasm)
        and all(
            str(item["contentType"]).split(";", 1)[0] == "application/wasm"
            for item in wasm
        ),
        "workerMime": bool(workers)
        and all(
            str(item["contentType"]).split(";", 1)[0]
            in {"application/javascript", "text/javascript"}
            for item in workers
        ),
        "runtimeSourcesFromConsumer": bool(workers and wasm)
        and all(
            urlsplit(str(item["url"])).netloc == urlsplit(base_url).netloc
            for item in workers + wasm
        ),
        "servedFilesExist": bool(served_resources)
        and all(bool(item["exists"]) for item in served_resources),
        "servedRuntimeHashesMatchBuiltConsumer": (
            resource_error is None
            and len(served_resources) == len(built_by_path) == 5
            and all(
                built_by_path.get(str(item["path"])) == item["sha256"]
                for item in served_resources
            )
        ),
        "tgzRuntimeManifestComplete": (
            runtime_evidence_error is None
            and isinstance(expected_resources, list)
            and len(expected_resources) == 5
            and len(built_runtime_files) == 5
        ),
    }
    return {
        "workers": workers,
        "wasm": wasm,
        "servedResources": served_resources,
        "resourceError": resource_error,
    }, checks


def blocked_attempt(attempt_dir: Path, error: Exception) -> dict[str, object]:
    result = {
        "status": "BLOCKED",
        "browserName": browser_name,
        "browser": None,
        "browserSource": None,
        "error": repr(error),
        "checks": {},
    }
    write_json(attempt_dir / "result.json", result)
    return result


def run_attempt(playwright, attempt_number: int) -> dict[str, object]:
    attempt_dir = evidence_dir / f"attempt-{attempt_number}"
    attempt_dir.mkdir(parents=True, exist_ok=True)
    try:
        browser, browser_source = launch_browser(playwright)
    except Exception as error:
        return blocked_attempt(attempt_dir, error)

    context = browser.new_context(
        viewport={"width": 1440, "height": 900},
        device_scale_factor=1,
        locale="zh-CN",
        timezone_id="Asia/Shanghai",
        accept_downloads=True,
    )
    context.tracing.start(screenshots=True, snapshots=True, sources=True)
    page = context.new_page()
    evidence = BrowserEvidence(page)
    error = None
    ui: dict[str, object] = {}
    feature_checks: dict[str, bool] = {}
    try:
        page.goto(base_url, wait_until="networkidle", timeout=30_000)
        page.wait_for_function(
            """() => document.querySelector('[data-testid="consumer-ready"]')?.dataset.state === 'ready' &&
              document.querySelector('[data-testid="consumer-xlsx-worker"]')?.textContent?.trim() === 'worker'""",
            timeout=60_000,
        )

        docx_viewer = page.locator('[data-testid="consumer-docx"] [data-testid="docx-viewer"]')
        docx_viewer.wait_for(state="visible", timeout=30_000)
        docx_text = page.locator('[data-testid="consumer-docx"] [data-testid="docx-page"]').all_inner_texts()
        docx_rendered_text = "\n".join(docx_text)
        assert docx_text and "INVOICE" in docx_rendered_text and "INV-2026-0705" in docx_rendered_text

        docx_toolbar = page.locator('[data-testid="consumer-docx"] [data-testid="docx-viewer-toolbar"]')
        docx_toolbar.wait_for(state="visible", timeout=30_000)
        docx_page_current = page.locator(
            '[data-testid="consumer-docx"] [data-testid="docx-page-current"]'
        )
        assert docx_page_current.input_value() == "1"
        assert f"of {len(docx_text)}" in " ".join(docx_toolbar.inner_text().split())
        docx_page = page.locator('[data-testid="consumer-docx"] [data-testid="docx-page"]').first
        docx_width_before = float(docx_page.bounding_box()["width"])
        docx_zoom = page.locator('[data-testid="consumer-docx"] [data-testid="docx-zoom-select"]')
        docx_zoom.select_option("125")
        page.wait_for_function(
            """minimum => document.querySelector('[data-testid="consumer-docx"] [data-testid="docx-page"]')?.getBoundingClientRect().width > minimum""",
            arg=docx_width_before * 1.15,
            timeout=30_000,
        )
        docx_width_after = float(docx_page.bounding_box()["width"])
        docx_sidebar = page.locator('[data-testid="consumer-docx"] [data-testid="docx-sidebar-toggle"]')
        if docx_sidebar.get_attribute("aria-pressed") != "true":
            docx_sidebar.click()
        docx_thumbnail = page.locator('[data-testid="consumer-docx"] [data-testid="docx-thumbnail"] canvas').first
        docx_thumbnail.wait_for(state="visible", timeout=30_000)
        page.wait_for_function(
            """() => document.querySelector('[data-testid="consumer-docx"] [data-testid="docx-thumbnail"] canvas')?.dataset.thumbnailState === 'ready'""",
            timeout=30_000,
        )
        docx_thumbnail_proof = docx_thumbnail.evaluate(
            """canvas => {
              const encoded = canvas.toDataURL('image/png');
              const blank = document.createElement('canvas');
              blank.width = canvas.width;
              blank.height = canvas.height;
              const context = blank.getContext('2d');
              context.fillStyle = '#ffffff';
              context.fillRect(0, 0, blank.width, blank.height);
              const blankEncoded = blank.toDataURL('image/png');
              return {
                state: canvas.dataset.thumbnailState,
                contentBlocks: Number(canvas.dataset.thumbnailContentBlocks || 0),
                encodedLength: encoded.length,
                blankLength: blankEncoded.length,
              };
            }"""
        )
        assert docx_thumbnail_proof["contentBlocks"] > 0
        assert docx_thumbnail_proof["encodedLength"] > docx_thumbnail_proof["blankLength"] + 300

        docx_editor = page.locator('[data-testid="consumer-docx-editor"] [data-testid="docx-editor"]')
        docx_editor.wait_for(state="visible", timeout=30_000)
        editor_toolbar = page.locator('[data-testid="consumer-docx-editor"] [data-testid="docx-editor-toolbar"]')
        editor_toolbar.wait_for(state="visible", timeout=30_000)
        editor_paragraph_selector = (
            '[data-testid="consumer-docx-editor"] '
            '[data-docx-paragraph-host="true"]'
            '[data-docx-paragraph-node-index="0"]'
        )
        editor_paragraph = page.locator(editor_paragraph_selector).first
        editor_paragraph.wait_for(state="visible", timeout=30_000)
        assert editor_paragraph.get_attribute("contenteditable") == "true"
        editor_original = editor_paragraph.inner_text()
        editor_marker = "〔TGZ〕"
        editor_paragraph.click()
        editor_paragraph.evaluate(
            """element => {
              element.focus();
              const range = document.createRange();
              range.selectNodeContents(element);
              range.collapse(false);
              const selection = window.getSelection();
              selection.removeAllRanges();
              selection.addRange(range);
            }"""
        )
        page.keyboard.type(editor_marker)
        page.locator('[data-testid="consumer-docx-editor"] h2').click()
        page.wait_for_function(
            """args => document.querySelector(args.selector)?.textContent === args.expected""",
            arg={
                "selector": editor_paragraph_selector,
                "expected": editor_original + editor_marker,
            },
            timeout=30_000,
        )
        editor_edit_committed = editor_paragraph.inner_text() == editor_original + editor_marker
        page.locator('[data-testid="consumer-docx-editor"] [data-testid="editor-undo"]').click()
        page.wait_for_function(
            """args => document.querySelector(args.selector)?.textContent === args.expected""",
            arg={"selector": editor_paragraph_selector, "expected": editor_original},
            timeout=30_000,
        )
        editor_undo_restored = editor_paragraph.inner_text() == editor_original
        page.locator('[data-testid="consumer-docx-readonly"]').check()
        page.wait_for_function(
            """selector => {
              const root = document.querySelector('[data-testid="consumer-docx-editor"]');
              return document.querySelector(selector)?.getAttribute('contenteditable') !== 'true'
                && root?.querySelector('[data-testid="editor-new"]')?.disabled === true
                && root?.querySelector('[data-testid="editor-import"]')?.disabled === true;
            }""",
            arg=editor_paragraph_selector,
            timeout=30_000,
        )
        editor_paragraph.click()
        page.keyboard.type("〔BLOCKED〕")
        editor_readonly_blocked = editor_paragraph.inner_text() == editor_original
        assert editor_readonly_blocked

        xlsx_canvas = page.locator('[data-testid="consumer-xlsx"] .xlsx-grid__body').evaluate(
            "canvas => ({ width: canvas.width, height: canvas.height, dataLength: canvas.toDataURL().length })"
        )
        page.locator('[data-testid="consumer-xlsx-select-a1"]').click()
        page.wait_for_function(
            """() => document.querySelector('[data-testid="consumer-xlsx-selection"]')?.textContent?.trim() === 'A1' &&
              document.querySelector('[data-testid="consumer-xlsx-value"]')?.textContent?.trim() === 'Region'""",
            timeout=30_000,
        )
        xlsx_name_box = page.locator(
            '[data-testid="consumer-xlsx"] [data-testid="xlsx-name-box"]'
        )
        xlsx_formula = page.locator(
            '[data-testid="consumer-xlsx"] [data-testid="xlsx-formula-input"]'
        )
        assert xlsx_name_box.input_value() == "A1"
        assert xlsx_formula.input_value() == "Region"
        assert not xlsx_formula.is_disabled()
        xlsx_sheet_tabs = page.locator(
            '[data-testid="consumer-xlsx"] [data-testid="xlsx-sheet-tab"]'
        )
        assert xlsx_sheet_tabs.count() >= 1
        xlsx_selected_sheet = page.locator(
            '[data-testid="consumer-xlsx"] [data-testid="xlsx-sheet-tab"][aria-selected="true"]'
        )
        assert xlsx_selected_sheet.count() == 1
        xlsx_sheet_names = [
            name
            for name in xlsx_sheet_tabs.evaluate_all(
                "tabs => tabs.map(tab => tab.dataset.sheetName || tab.textContent?.trim() || '')"
            )
            if name
        ]
        assert "Sales Data" in xlsx_sheet_names
        xlsx_grid = page.locator('[data-testid="consumer-xlsx"] [data-testid="xlsx-grid"]')
        assert int(xlsx_grid.get_attribute("data-display-row-count") or "0") >= 200
        assert int(xlsx_grid.get_attribute("data-display-column-count") or "0") >= 50
        xlsx_grid.click(position={"x": 60, "y": 36})
        page.keyboard.press("Enter")
        xlsx_cell_editor = page.locator('[data-testid="consumer-xlsx"] [data-testid="xlsx-cell-editor"]')
        xlsx_cell_editor.wait_for(state="visible", timeout=30_000)
        xlsx_cell_editor.fill("Region-TGZ")
        xlsx_cell_editor.press("Enter")
        page.wait_for_function(
            """() => document.querySelector('[data-testid="consumer-xlsx"] [data-testid="xlsx-formula-input"]')?.value === 'Region-TGZ'""",
            timeout=30_000,
        )
        xlsx_edit_committed = xlsx_formula.input_value() == "Region-TGZ"
        page.locator('[data-testid="consumer-xlsx"] [data-testid="xlsx-undo"]').click()
        page.wait_for_function(
            """() => document.querySelector('[data-testid="consumer-xlsx"] [data-testid="xlsx-formula-input"]')?.value === 'Region'""",
            timeout=30_000,
        )
        xlsx_undo_restored = xlsx_formula.input_value() == "Region"
        page.locator('[data-testid="consumer-xlsx-readonly"]').check()
        page.wait_for_function(
            """() => {
              const root = document.querySelector('[data-testid="consumer-xlsx"]');
              return root?.querySelector('[data-testid="xlsx-formula-input"]')?.disabled === true
                && root?.querySelector('[data-testid="xlsx-ribbon-read-only"]')?.checked === true
                && !root?.querySelector('[data-testid="xlsx-add-sheet"]')
                && !root?.querySelector('[data-testid="xlsx-remove-sheet"]');
            }""",
            timeout=30_000,
        )
        xlsx_grid.click(position={"x": 60, "y": 36})
        page.keyboard.type("BLOCKED")
        xlsx_readonly_blocked = (
            page.locator(
                '[data-testid="consumer-xlsx"] [data-testid="xlsx-cell-editor"]'
            ).count()
            == 0
            and xlsx_formula.input_value() == "Region"
        )
        assert xlsx_readonly_blocked

        wait_for_pdf_page(page, 1)
        initial_pdf = pdf_image(page)
        assert page.locator('[data-testid="pdf-page-count"]').inner_text().strip() == "4"
        assert page.locator('[data-testid="pdf-thumbnail"]').count() == 4
        page.locator('[data-testid="pdf-thumbnail-image"]').first.wait_for(
            state="visible", timeout=30_000
        )
        page.locator('[data-testid="pdf-next-page"]').click()
        wait_for_pdf_page(page, 2)
        page.locator('[data-testid="pdf-zoom-select"]').select_option("2")
        page.wait_for_function(
            """() => document.querySelector('[data-testid="pdf-document"]')?.dataset.zoom === '2' &&
              document.querySelector('[data-testid="pdf-page"]')?.dataset.renderState === 'rendered'""",
            timeout=45_000,
        )
        zoomed_pdf = pdf_image(page)
        assert int(zoomed_pdf["naturalWidth"]) > int(initial_pdf["naturalWidth"])
        page.locator('[data-testid="pdf-zoom-select"]').select_option("1")
        wait_for_pdf_page(page, 2)
        page.locator('[data-testid="pdf-rotate"]').click()
        page.wait_for_function(
            """() => document.querySelector('[data-testid="pdf-document"]')?.dataset.rotation === '90' &&
              document.querySelector('[data-testid="pdf-page"]')?.dataset.renderState === 'rendered'""",
            timeout=45_000,
        )
        page.locator('[data-testid="pdf-search-input"]').fill("ALPHA-4")
        page.locator('[data-testid="pdf-search-submit"]').click()
        page.wait_for_function(
            """() => document.querySelector('[data-testid="pdf-search-result"]')?.textContent?.includes('第 4 页')""",
            timeout=45_000,
        )
        wait_for_pdf_page(page, 4)
        assert "ALPHA-4" in page.locator(".pdf-search-context").inner_text()
        page.locator('[data-testid="pdf-thumbnail"][data-page-number="1"]').click()
        wait_for_pdf_page(page, 1)

        fixture = consumer_dir / "public" / "fixtures" / "sample.pdf"
        expected_pdf = fixture.read_bytes()
        downloads_dir = attempt_dir / "downloads"
        downloads_dir.mkdir(parents=True, exist_ok=True)
        with page.expect_download(timeout=30_000) as download_info:
            page.locator('[data-testid="pdf-download"]').click()
        download = download_info.value
        downloaded_path = downloads_dir / download.suggested_filename
        download.save_as(downloaded_path)
        downloaded_pdf = downloaded_path.read_bytes()
        assert download.suggested_filename == "sample.pdf"
        assert downloaded_pdf == expected_pdf and downloaded_pdf.startswith(b"%PDF-")

        native_pdf_embeds = page.locator(
            '[data-testid="pdf-viewer"] iframe, [data-testid="pdf-viewer"] embed, '
            '[data-testid="pdf-viewer"] object'
        ).count()
        feature_checks = {
            "allReady": page.locator('[data-testid="consumer-ready"]').get_attribute("data-state") == "ready",
            "docxVisibleContent": bool(docx_text) and "INV-2026-0705" in docx_rendered_text,
            "docxWorker": page.locator('[data-testid="consumer-docx-source"]').inner_text().strip() == "worker",
            "docxViewerChrome": docx_toolbar.is_visible() and docx_width_after > docx_width_before * 1.15,
            "docxPageNavigation": docx_page_current.input_value() == "1" and f"of {len(docx_text)}" in " ".join(docx_toolbar.inner_text().split()),
            "docxThumbnails": docx_thumbnail_proof["state"] == "ready" and docx_thumbnail_proof["contentBlocks"] > 0,
            "docxEditorChrome": editor_toolbar.is_visible(),
            "docxEditorInteraction": editor_edit_committed and editor_undo_restored,
            "docxEditorReadOnly": editor_readonly_blocked and page.locator('[data-testid="consumer-docx-editor"] [data-testid="editor-new"]').is_disabled(),
            "xlsxWorker": page.locator('[data-testid="consumer-xlsx-worker"]').inner_text().strip() == "worker",
            "xlsxCanvasRendered": int(xlsx_canvas["width"]) > 0 and int(xlsx_canvas["height"]) > 0 and int(xlsx_canvas["dataLength"]) > 200,
            "xlsxCellSelection": page.locator('[data-testid="consumer-xlsx-value"]').inner_text().strip() == "Region",
            "xlsxWorkspace": int(xlsx_grid.get_attribute("data-display-row-count") or "0") >= 200 and int(xlsx_grid.get_attribute("data-display-column-count") or "0") >= 50,
            "xlsxFormulaBar": xlsx_name_box.input_value() == "A1" and xlsx_formula.input_value() == "Region",
            "xlsxSheetTabs": "Sales Data" in xlsx_sheet_names and xlsx_selected_sheet.get_attribute("data-sheet-name") == "Sales Data",
            "xlsxEditorInteraction": xlsx_edit_committed and xlsx_undo_restored,
            "xlsxReadOnly": xlsx_readonly_blocked and xlsx_formula.is_disabled(),
            "pdfFourPages": page.locator('[data-testid="pdf-page-count"]').inner_text().strip() == "4",
            "pdfNavigation": page.locator('[data-testid="pdf-document"]').get_attribute("data-active-page") == "1",
            "pdfZoom": int(zoomed_pdf["naturalWidth"]) > int(initial_pdf["naturalWidth"]),
            "pdfRotation": page.locator('[data-testid="pdf-document"]').get_attribute("data-rotation") == "90",
            "pdfThumbnails": page.locator('[data-testid="pdf-thumbnail"]').count() == 4,
            "pdfSearch": "第 4 页" in page.locator('[data-testid="pdf-search-result"]').inner_text(),
            "pdfDownload": downloaded_pdf == expected_pdf,
            "pdfMaxFileSize50MiB": page.locator('[data-testid="consumer-pdf-max-file-size"]').inner_text().strip() == str(50 * 1024 * 1024),
            "pdfRenderedFromVerifiedBytes": str(initial_pdf["src"]).startswith("blob:") and native_pdf_embeds == 0,
        }
        evidence.assert_clean()
        ui = {
            "docxTextSample": docx_rendered_text[:240],
            "docxZoom": {"before": docx_width_before, "after": docx_width_after},
            "docxPages": {"current": docx_page_current.input_value(), "total": len(docx_text)},
            "docxThumbnail": docx_thumbnail_proof,
            "docxEditorOriginal": editor_original,
            "docxEditor": {
                "editCommitted": editor_edit_committed,
                "undoRestored": editor_undo_restored,
                "readOnlyBlocked": editor_readonly_blocked,
            },
            "xlsxCanvas": xlsx_canvas,
            "xlsxSelectedAddress": page.locator('[data-testid="consumer-xlsx-selection"]').inner_text(),
            "xlsxSelectedValue": page.locator('[data-testid="consumer-xlsx-value"]').inner_text(),
            "xlsxFormula": {"nameBox": xlsx_name_box.input_value(), "value": xlsx_formula.input_value()},
            "xlsxSheets": xlsx_sheet_names,
            "xlsxEditor": {
                "editCommitted": xlsx_edit_committed,
                "undoRestored": xlsx_undo_restored,
                "readOnlyBlocked": xlsx_readonly_blocked,
            },
            "pdfInitial": initial_pdf,
            "pdfZoomed": zoomed_pdf,
            "pdfSearchResult": page.locator('[data-testid="pdf-search-result"]').inner_text(),
            "downloadSha256": hashlib.sha256(downloaded_pdf).hexdigest(),
        }
    except Exception as caught:
        error = repr(caught)

    runtime_details, runtime_checks = runtime_proof(evidence.events)
    checks = {
        **feature_checks,
        **runtime_checks,
        "cleanBrowser": not any(evidence.violations().values()),
    }
    try:
        page.screenshot(path=str(attempt_dir / "page.png"), full_page=True)
    except Exception as screenshot_error:
        error = f"{error or ''}; screenshot: {screenshot_error!r}".strip("; ")
    evidence.save(attempt_dir)
    try:
        context.tracing.stop(path=attempt_dir / "trace.zip")
    except Exception as trace_error:
        error = f"{error or ''}; trace: {trace_error!r}".strip("; ")
    context.close()
    browser_version = browser.version
    browser.close()

    result = {
        "suite": "BB-PACK-CONSUMER",
        "case": "PACK-003-TGZ-FORMAL-PREVIEW",
        "attempt": attempt_number,
        "status": "PASS" if error is None and checks and all(checks.values()) else "FAIL",
        "browser": browser_version,
        "browserName": browser_name,
        "browserSource": browser_source,
        "baseUrl": base_url,
        "environment": {
            "locale": "zh-CN",
            "timezone": "Asia/Shanghai",
            "viewport": "1440x900",
            "deviceScaleFactor": 1,
        },
        "ui": ui,
        "checks": checks,
        "error": error or runtime_evidence_error,
        **runtime_details,
    }
    if runtime_evidence_error:
        result["status"] = "FAIL"
    write_json(attempt_dir / "result.json", result)
    return result


with sync_playwright() as playwright:
    attempts = [run_attempt(playwright, 1)]
    if attempts[0]["status"] == "FAIL":
        attempts.append(run_attempt(playwright, 2))

first_status = str(attempts[0]["status"])
if first_status == "PASS":
    final_status = "PASS"
    selected = attempts[0]
elif first_status == "BLOCKED":
    final_status = "BLOCKED"
    selected = attempts[0]
elif len(attempts) == 2 and attempts[1]["status"] == "PASS":
    final_status = "FLAKY"
    selected = attempts[1]
else:
    final_status = "FAIL"
    selected = attempts[-1]

result = {
    **selected,
    "status": final_status,
    "attempts": [
        {
            "attempt": item.get("attempt", index + 1),
            "status": item["status"],
            "error": item.get("error"),
            "evidenceDirectory": str(evidence_dir / f"attempt-{index + 1}"),
        }
        for index, item in enumerate(attempts)
    ],
}
write_json(evidence_dir / "result.json", result)
print(json.dumps(result, ensure_ascii=False, indent=2))
raise SystemExit(0 if final_status == "PASS" else 1)
