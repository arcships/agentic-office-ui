from __future__ import annotations

import hashlib
import json
from pathlib import Path
import sys
from urllib.parse import urlsplit

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

from browser_evidence import BrowserEvidence


base_url = sys.argv[1].rstrip("/")
evidence_dir = Path(sys.argv[2]).resolve()
consumer_dir = Path(sys.argv[3]).resolve()
expected_resources = json.loads(Path(sys.argv[4]).read_text(encoding="utf-8"))
evidence_dir.mkdir(parents=True, exist_ok=True)


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


def served_file(response: dict[str, object]) -> Path:
    relative = urlsplit(str(response["url"])).path.lstrip("/")
    return consumer_dir / "dist" / relative


with sync_playwright() as playwright:
    browser, browser_source = launch_browser(playwright)
    context = browser.new_context(viewport={"width": 1440, "height": 900})
    context.tracing.start(screenshots=True, snapshots=True, sources=True)
    page = context.new_page()
    evidence = BrowserEvidence(page)
    error = None
    ui: dict[str, object] = {}
    try:
        page.goto(base_url, wait_until="networkidle", timeout=30_000)
        page.wait_for_function(
            """() => {
              const image = document.querySelector('[data-testid="pdf-page-image"]');
              return document.querySelector('[data-testid="consumer-ready"]')?.dataset.state === 'ready' &&
                document.querySelector('[data-testid="consumer-xlsx-worker"]')?.textContent?.trim() === 'worker' &&
                image instanceof HTMLImageElement && image.complete &&
                image.naturalWidth > 0 && image.naturalHeight > 0;
            }""",
            timeout=60_000,
        )
        pdf_image = page.locator('[data-testid="pdf-page-image"]').evaluate(
            """image => ({
              src: image.currentSrc || image.src,
              naturalWidth: image.naturalWidth,
              naturalHeight: image.naturalHeight,
            })"""
        )
        ui = {
            "overall": page.locator('[data-testid="consumer-ready"]').get_attribute("data-state"),
            "docx": page.locator('[data-testid="consumer-docx-state"]').get_attribute("data-state"),
            "docxSource": page.locator('[data-testid="consumer-docx-source"]').inner_text(),
            "xlsx": page.locator('[data-testid="consumer-xlsx-state"]').get_attribute("data-state"),
            "xlsxWorker": page.locator('[data-testid="consumer-xlsx-worker"]').inner_text(),
            "pdf": page.locator('[data-testid="consumer-pdf-state"]').get_attribute("data-state"),
            "pdfRender": pdf_image,
            "pdfNativeEmbedCount": page.locator(
                '[data-testid="pdf-viewer"] iframe, '
                '[data-testid="pdf-viewer"] embed, '
                '[data-testid="pdf-viewer"] object'
            ).count(),
        }
        evidence.assert_clean()
        assert not evidence.events["downloads"], evidence.events["downloads"]
    except Exception as caught:
        error = repr(caught)
    page.screenshot(path=str(evidence_dir / "page.png"), full_page=True)
    evidence.save(evidence_dir)
    context.tracing.stop(path=evidence_dir / "trace.zip")
    context.close()
    browser_version = browser.version
    browser.close()

workers = [
    item for item in evidence.events["responses"]
    if ("docx-import-worker" in str(item["url"]) or "xlsx-worker" in str(item["url"]))
    and int(item["status"]) == 200
]
wasm = [
    item for item in evidence.events["responses"]
    if str(item["url"]).split("?", 1)[0].endswith(".wasm") and int(item["status"]) == 200
]
served_resources = []
for response in workers + wasm:
    file_path = served_file(response)
    served_resources.append(
        {
            "url": response["url"],
            "path": str(file_path.relative_to(consumer_dir)),
            "exists": file_path.is_file(),
            "sha256": hashlib.sha256(file_path.read_bytes()).hexdigest() if file_path.is_file() else None,
            "contentType": response["contentType"],
        }
    )

expected_wasm_hashes = {
    item["sha256"] for item in expected_resources if str(item["path"]).endswith(".wasm")
}
expected_pdfium = next(
    (
        item for item in expected_resources
        if item.get("package") == "@extend-ai/vue-extend"
        and item.get("path") == "dist/pdfium.wasm"
    ),
    None,
)
expected_worker_entries = [
    item for item in expected_resources if str(item["path"]).endswith("worker.js")
]
served_wasm_hashes = {
    item["sha256"] for item in served_resources
    if str(item["path"]).endswith(".wasm") and item["sha256"]
}
checks = {
    "allReady": ui.get("overall") == "ready" and ui.get("docx") == "ready" and ui.get("xlsx") == "ready" and ui.get("pdf") == "ready",
    "docxWorker": ui.get("docxSource") == "worker",
    "xlsxWorker": str(ui.get("xlsxWorker", "")).strip() == "worker",
    "twoWorkers": len({item["url"] for item in workers}) >= 2,
    "threeWasm": len({item["url"] for item in wasm}) >= 3,
    "wasmMime": bool(wasm) and all(str(item["contentType"]).split(";", 1)[0] == "application/wasm" for item in wasm),
    "runtimeSourcesFromConsumer": all(
        urlsplit(str(item["url"])).netloc == urlsplit(base_url).netloc
        for item in workers + wasm
    ),
    "servedFilesExist": bool(served_resources) and all(item["exists"] for item in served_resources),
    "servedWasmHashesFromTgz": served_wasm_hashes == expected_wasm_hashes,
    "pdfiumDeclaredByTgz": expected_pdfium is not None and expected_pdfium.get("sha256") in served_wasm_hashes,
    "pdfPageRenderedFromVerifiedBytes": (
        str(ui.get("pdfRender", {}).get("src", "")).startswith("blob:")
        and int(ui.get("pdfRender", {}).get("naturalWidth", 0)) > 0
        and int(ui.get("pdfRender", {}).get("naturalHeight", 0)) > 0
        and ui.get("pdfNativeEmbedCount") == 0
    ),
    "workerSourcesDeclaredByTgz": len(workers) >= 2 and len(expected_worker_entries) == 2,
    "cleanBrowser": not any(evidence.violations().values()),
}
result = {
    "suite": "BB-PACK-CONSUMER",
    "case": "PACK-003-TGZ-FORMAL-PREVIEW",
    "status": "PASS" if error is None and all(checks.values()) else "FAIL",
    "browser": browser_version,
    "browserSource": browser_source,
    "baseUrl": base_url,
    "ui": ui,
    "checks": checks,
    "error": error,
    "workers": workers,
    "wasm": wasm,
    "servedResources": served_resources,
}
(evidence_dir / "result.json").write_text(
    json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
)
print(json.dumps(result, ensure_ascii=False, indent=2))
raise SystemExit(0 if result["status"] == "PASS" else 1)
