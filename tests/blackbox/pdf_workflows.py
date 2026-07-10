from __future__ import annotations

import hashlib
import json
import os
import platform
import signal
import socket
import subprocess
import sys
import time
import traceback
import urllib.request
from pathlib import Path
from typing import Callable
from urllib.parse import urlparse

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import Page, sync_playwright

from browser_evidence import BrowserEvidence


ROOT = Path(__file__).resolve().parents[2]
SAMPLES = ROOT / "apps" / "demo" / "public" / "samples"
FIXTURE_MANIFEST = ROOT / "test-data" / "manifest.json"
OUTPUT = Path(
    os.environ.get(
        "BLACKBOX_EVIDENCE_DIR",
        ROOT / "output" / "acceptance" / "bb-pdf-viewer",
    )
).resolve()
PORT = int(os.environ.get("CI_PREVIEW_PORT", "4173"))
BASE_URL = f"http://127.0.0.1:{PORT}"

PDF_BLOB_PROBE = """
(() => {
  const created = [];
  const revoked = [];
  const workers = [];
  const originalCreate = URL.createObjectURL.bind(URL);
  const originalRevoke = URL.revokeObjectURL.bind(URL);
  const OriginalWorker = window.Worker;
  Object.defineProperty(window, "__pdfBlobProbe", {
    value: { created, revoked, workers },
    configurable: false,
    writable: false,
  });
  URL.createObjectURL = (blob) => {
    const url = originalCreate(blob);
    created.push({ url, type: blob?.type || "", size: blob?.size ?? -1 });
    return url;
  };
  URL.revokeObjectURL = (url) => {
    revoked.push(String(url));
    return originalRevoke(url);
  };
  function TrackedWorker(...args) {
    workers.push(String(args[0]));
    return new OriginalWorker(...args);
  }
  TrackedWorker.prototype = OriginalWorker.prototype;
  Object.setPrototypeOf(TrackedWorker, OriginalWorker);
  window.Worker = TrackedWorker;
})();
"""


def fixture_entry(file_name: str) -> dict[str, object]:
    manifest = json.loads(FIXTURE_MANIFEST.read_text(encoding="utf-8"))
    for entry in manifest["entries"]:
        if entry["file"] == file_name:
            return entry
    raise AssertionError(f"missing fixture manifest entry for {file_name}")


def git_value(*args: str) -> str:
    result = subprocess.run(
        ["git", *args], cwd=ROOT, text=True, capture_output=True, check=False
    )
    return result.stdout.strip() if result.returncode == 0 else "unknown"


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
            return (
                playwright.chromium.launch(
                    headless=True, executable_path=str(candidate)
                ),
                str(candidate),
            )
        except PlaywrightError as error:
            failures.append(str(error))
    raise RuntimeError("no Chromium executable could launch: " + " | ".join(failures))


def wait_for_rendered_page(page: Page, page_number: int, file_name: str) -> None:
    page.wait_for_function(
        """expected => {
          const status = document.querySelector('[data-testid="page-status"]');
          const loaded = document.querySelector('[data-testid="loaded-file"]');
          const viewer = document.querySelector('[data-testid="pdf-document"]');
          const stage = document.querySelector('[data-testid="pdf-page"]');
          const image = document.querySelector('[data-testid="pdf-page-image"]');
          return status?.dataset.state === 'ready' &&
            loaded?.textContent?.includes(expected.fileName) &&
            viewer?.dataset.activePage === String(expected.pageNumber) &&
            stage?.dataset.pageNumber === String(expected.pageNumber) &&
            stage?.dataset.renderState === 'rendered' &&
            image instanceof HTMLImageElement && image.complete &&
            image.naturalWidth > 0 && image.naturalHeight > 0 &&
            image.alt === `PDF 第 ${expected.pageNumber} 页`;
        }""",
        arg={"pageNumber": page_number, "fileName": file_name},
        timeout=45_000,
    )


def navigate_pdf(page: Page, file_name: str = "sample.pdf") -> None:
    page.goto(f"{BASE_URL}/#/pdf-viewer", wait_until="networkidle")
    page.locator('[data-testid="pdf-sample-select"]').wait_for(
        state="visible", timeout=30_000
    )
    if page.locator('[data-testid="pdf-sample-select"]').input_value() != file_name:
        page.locator('[data-testid="pdf-sample-select"]').select_option(file_name)
    wait_for_rendered_page(page, 1, file_name)


def select_sample(page: Page, file_name: str, page_number: int = 1) -> None:
    page.locator('[data-testid="pdf-sample-select"]').select_option(file_name)
    wait_for_rendered_page(page, page_number, file_name)


def image_snapshot(page: Page) -> dict[str, object]:
    return page.locator('[data-testid="pdf-page-image"]').evaluate(
        """image => {
          const rect = image.getBoundingClientRect();
          return {
            alt: image.alt,
            src: image.currentSrc || image.src,
            naturalWidth: image.naturalWidth,
            naturalHeight: image.naturalHeight,
            displayedWidth: rect.width,
            displayedHeight: rect.height,
          };
        }"""
    )


def document_snapshot(page: Page) -> dict[str, object]:
    return page.locator('[data-testid="pdf-document"]').evaluate(
        """element => ({
          pageCount: Number(element.dataset.pageCount),
          activePage: Number(element.dataset.activePage),
          zoom: Number(element.dataset.zoom),
          rotation: Number(element.dataset.rotation),
        })"""
    )


def assert_no_native_pdf_embed(page: Page) -> None:
    count = page.locator(
        '[data-testid="pdf-viewer"] iframe, '
        '[data-testid="pdf-viewer"] embed, '
        '[data-testid="pdf-viewer"] object'
    ).count()
    assert count == 0, f"PDF viewer contains {count} iframe/embed/object elements"


def response_count(events: BrowserEvidence, path: str) -> int:
    return sum(
        1
        for item in events.events["responses"]
        if urlparse(str(item["url"])).path == path
    )


def screenshot(page: Page, attempt_dir: Path, name: str) -> None:
    page.screenshot(path=str(attempt_dir / name), full_page=True)


def pdf_001_rendered_pages(
    page: Page, attempt_dir: Path, evidence: BrowserEvidence
) -> dict[str, object]:
    navigate_pdf(page)
    image = image_snapshot(page)
    document = document_snapshot(page)
    assert document["pageCount"] == 4, document
    assert document["activePage"] == 1, document
    assert image["naturalWidth"] > 0 and image["naturalHeight"] > 0, image
    assert str(image["src"]).startswith("blob:"), image
    page.locator('[data-testid="pdf-thumbnail-image"]').first.wait_for(
        state="visible", timeout=30_000
    )
    thumbnail = page.locator('[data-testid="pdf-thumbnail-image"]').first.evaluate(
        "image => ({ width: image.naturalWidth, height: image.naturalHeight })"
    )
    assert thumbnail["width"] > 0 and thumbnail["height"] > 0, thumbnail
    assert page.locator('[data-testid="pdf-thumbnail"]').count() == 4
    wasm = [
        item
        for item in evidence.events["responses"]
        if urlparse(str(item["url"])).path.endswith(".wasm")
    ]
    assert wasm, "real PDF rendering did not request a WASM engine"
    assert all(int(item["status"]) == 200 for item in wasm), wasm
    assert_no_native_pdf_embed(page)
    screenshot(page, attempt_dir, "PDF-001-rendered-page.png")
    return {"document": document, "image": image, "thumbnail": thumbnail, "wasm": wasm}


def pdf_002_navigation(
    page: Page, attempt_dir: Path, _evidence: BrowserEvidence
) -> dict[str, object]:
    navigate_pdf(page)
    visited: list[dict[str, object]] = []

    def capture(action: str, expected_page: int) -> None:
        wait_for_rendered_page(page, expected_page, "sample.pdf")
        current = document_snapshot(page)
        selected = page.locator(
            f'[data-testid="pdf-thumbnail"][data-page-number="{expected_page}"]'
        )
        assert selected.get_attribute("aria-current") == "page"
        visited.append({"action": action, "document": current, "image": image_snapshot(page)})

    page.locator('[data-testid="pdf-next-page"]').click()
    capture("next", 2)
    page.locator('[data-testid="pdf-page-input"]').fill("3")
    page.locator('[data-testid="pdf-page-input"]').press("Enter")
    capture("page-input", 3)
    page.locator('[data-testid="pdf-previous-page"]').click()
    capture("previous", 2)
    page.locator('[data-testid="pdf-last-page"]').click()
    capture("last", 4)
    page.locator('[data-testid="pdf-first-page"]').click()
    capture("first", 1)
    page.locator('[data-testid="pdf-thumbnail"][data-page-number="4"]').click()
    capture("thumbnail", 4)
    assert_no_native_pdf_embed(page)
    screenshot(page, attempt_dir, "PDF-002-navigation-page-4.png")
    return {"visited": visited}


def pdf_003_zoom_rotate_search(
    page: Page, attempt_dir: Path, _evidence: BrowserEvidence
) -> dict[str, object]:
    navigate_pdf(page)
    initial = image_snapshot(page)
    initial_src = initial["src"]
    page.locator('[data-testid="pdf-zoom-select"]').select_option("2")
    page.wait_for_function(
        """oldSrc => {
          const viewer = document.querySelector('[data-testid="pdf-document"]');
          const stage = document.querySelector('[data-testid="pdf-page"]');
          const image = document.querySelector('[data-testid="pdf-page-image"]');
          return viewer?.dataset.zoom === '2' && stage?.dataset.renderState === 'rendered' &&
            image instanceof HTMLImageElement && image.complete && image.src !== oldSrc &&
            image.naturalWidth > 0;
        }""",
        arg=initial_src,
        timeout=45_000,
    )
    zoomed = image_snapshot(page)
    assert zoomed["naturalWidth"] > initial["naturalWidth"], (initial, zoomed)
    assert zoomed["displayedWidth"] > initial["displayedWidth"], (initial, zoomed)
    screenshot(page, attempt_dir, "PDF-003-zoom-200.png")

    page.locator('[data-testid="pdf-zoom-select"]').select_option("1")
    wait_for_rendered_page(page, 1, "sample.pdf")
    before_rotate = image_snapshot(page)
    page.locator('[data-testid="pdf-rotate"]').click()
    page.wait_for_function(
        """oldSrc => {
          const viewer = document.querySelector('[data-testid="pdf-document"]');
          const stage = document.querySelector('[data-testid="pdf-page"]');
          const image = document.querySelector('[data-testid="pdf-page-image"]');
          return viewer?.dataset.rotation === '90' && stage?.dataset.renderState === 'rendered' &&
            image instanceof HTMLImageElement && image.complete && image.src !== oldSrc;
        }""",
        arg=before_rotate["src"],
        timeout=45_000,
    )
    rotated = image_snapshot(page)
    assert before_rotate["naturalHeight"] > before_rotate["naturalWidth"], before_rotate
    assert rotated["naturalWidth"] > rotated["naturalHeight"], rotated
    screenshot(page, attempt_dir, "PDF-003-rotated-clockwise.png")

    page.locator('[data-testid="pdf-search-input"]').fill("ALPHA-4")
    page.locator('[data-testid="pdf-search-submit"]').click()
    page.wait_for_function(
        """() => document.querySelector('[data-testid="pdf-search-result"]')?.textContent?.includes('第 4 页')""",
        timeout=45_000,
    )
    wait_for_rendered_page(page, 4, "sample.pdf")
    search_result = page.locator('[data-testid="pdf-search-result"]').inner_text()
    context_text = page.locator(".pdf-search-context").inner_text()
    assert "ALPHA-4" in context_text, context_text
    assert_no_native_pdf_embed(page)
    screenshot(page, attempt_dir, "PDF-003-search-alpha-4.png")
    return {
        "initial": initial,
        "zoomed": zoomed,
        "beforeRotate": before_rotate,
        "rotated": rotated,
        "searchResult": search_result,
        "searchContext": context_text,
    }


def pdf_004_fixture_matrix(
    page: Page, attempt_dir: Path, _evidence: BrowserEvidence
) -> dict[str, object]:
    navigate_pdf(page)
    results: dict[str, object] = {}

    select_sample(page, "scanned-invoice.pdf")
    scanned = {"document": document_snapshot(page), "image": image_snapshot(page)}
    assert scanned["document"]["pageCount"] == 3, scanned
    assert scanned["image"]["naturalWidth"] > 0, scanned
    screenshot(page, attempt_dir, "PDF-004-scanned-invoice.png")
    results["scanned-invoice.pdf"] = scanned

    select_sample(page, "rotated-pages.pdf")
    portrait = image_snapshot(page)
    page.locator('[data-testid="pdf-next-page"]').click()
    wait_for_rendered_page(page, 2, "rotated-pages.pdf")
    landscape = image_snapshot(page)
    assert portrait["naturalHeight"] > portrait["naturalWidth"], portrait
    assert landscape["naturalWidth"] > landscape["naturalHeight"], landscape
    screenshot(page, attempt_dir, "PDF-004-rotated-pages-landscape.png")
    results["rotated-pages.pdf"] = {
        "document": document_snapshot(page),
        "portrait": portrait,
        "landscape": landscape,
    }

    select_sample(page, "large-contract.pdf")
    large = document_snapshot(page)
    assert large["pageCount"] == 31, large
    assert page.locator('[data-testid="pdf-thumbnail"]').count() == 31
    page.locator('[data-testid="pdf-last-page"]').click()
    wait_for_rendered_page(page, 31, "large-contract.pdf")
    last_page = image_snapshot(page)
    assert last_page["naturalWidth"] > 0 and last_page["naturalHeight"] > 0, last_page
    screenshot(page, attempt_dir, "PDF-004-large-contract-page-31.png")
    results["large-contract.pdf"] = {
        "document": document_snapshot(page),
        "lastPage": last_page,
        "thumbnailCount": 31,
    }
    assert_no_native_pdf_embed(page)
    return results


def pdf_005_verified_download(
    page: Page, attempt_dir: Path, evidence: BrowserEvidence
) -> dict[str, object]:
    navigate_pdf(page)
    expected = fixture_entry("sample.pdf")
    source_path = "/samples/sample.pdf"
    before_requests = response_count(evidence, source_path)
    assert before_requests == 1, {
        "message": "sample PDF must be fetched exactly once before download",
        "responses": evidence.events["responses"],
    }
    downloads_dir = attempt_dir / "downloads"
    downloads_dir.mkdir(parents=True, exist_ok=True)
    with page.expect_download(timeout=30_000) as download_info:
        page.locator('[data-testid="pdf-download"]').click()
    download = download_info.value
    actual_path = downloads_dir / download.suggested_filename
    download.save_as(actual_path)
    data = actual_path.read_bytes()
    page.wait_for_timeout(250)
    after_requests = response_count(evidence, source_path)
    assert after_requests == before_requests, {
        "before": before_requests,
        "after": after_requests,
    }
    digest = hashlib.sha256(data).hexdigest()
    assert download.suggested_filename == "sample.pdf", download.suggested_filename
    assert actual_path.name == "sample.pdf", actual_path.name
    assert len(data) == expected["bytes"], (len(data), expected)
    assert digest == expected["sha256"], (digest, expected)
    assert data.startswith(b"%PDF-"), data[:16]
    blob_probe = page.evaluate("window.__pdfBlobProbe")
    matching_blobs = [
        item
        for item in blob_probe["created"]
        if item["type"] == "application/pdf" and item["size"] == len(data)
    ]
    assert matching_blobs, blob_probe
    file_result = subprocess.run(
        ["file", "--brief", "--mime-type", str(actual_path)],
        text=True,
        capture_output=True,
        check=False,
    )
    detected_mime = file_result.stdout.strip()
    assert file_result.returncode == 0 and detected_mime == "application/pdf", {
        "returnCode": file_result.returncode,
        "stdout": file_result.stdout,
        "stderr": file_result.stderr,
    }
    assert_no_native_pdf_embed(page)
    screenshot(page, attempt_dir, "PDF-005-downloaded.png")
    return {
        "suggestedFileName": download.suggested_filename,
        "actualFileName": actual_path.name,
        "bytes": len(data),
        "sha256": digest,
        "mime": detected_mime,
        "blobMime": matching_blobs[-1]["type"],
        "sourceRequestsBeforeDownload": before_requests,
        "sourceRequestsAfterDownload": after_requests,
        "downloadEvent": evidence.events["downloads"][-1],
    }


def pdf_006_corrupted_recovery(
    page: Page, attempt_dir: Path, _evidence: BrowserEvidence
) -> dict[str, object]:
    navigate_pdf(page)
    page.locator('[data-testid="pdf-search-input"]').fill("ALPHA-4")
    page.locator('[data-testid="pdf-search-submit"]').click()
    wait_for_rendered_page(page, 4, "sample.pdf")
    page.locator('[data-testid="pdf-sample-select"]').select_option("corrupted.pdf")
    page.wait_for_function(
        """() => document.querySelector('[data-testid="page-status"]')?.dataset.state === 'error'""",
        timeout=45_000,
    )
    error = page.locator('[data-testid="pdf-page-load-error"]').inner_text()
    error_code = page.locator('[data-testid="pdf-page-load-error"]').get_attribute(
        "data-error-code"
    )
    assert error_code == "INVALID_PDF", (error_code, error)
    assert page.locator('[data-testid="pdf-document"]').count() == 0
    assert page.locator('[data-testid="pdf-thumbnail"]').count() == 0
    assert page.locator('[data-testid="pdf-page-image"]').count() == 0
    assert page.locator('[data-testid="pdf-search-result"]').count() == 0
    assert page.locator('[data-testid="pdf-download"]').is_disabled()
    assert_no_native_pdf_embed(page)
    screenshot(page, attempt_dir, "PDF-006-corrupted-error.png")

    select_sample(page, "sample.pdf")
    recovered = {"document": document_snapshot(page), "image": image_snapshot(page)}
    assert recovered["document"]["pageCount"] == 4, recovered
    assert recovered["document"]["activePage"] == 1, recovered
    assert page.locator('[data-testid="pdf-thumbnail"]').count() == 4
    assert not page.locator('[data-testid="pdf-download"]').is_disabled()
    assert_no_native_pdf_embed(page)
    screenshot(page, attempt_dir, "PDF-006-recovered-sample.png")
    return {"errorCode": error_code, "error": error, "recovered": recovered}


def pdf_008_file_size_limit(
    page: Page, attempt_dir: Path, evidence: BrowserEvidence
) -> dict[str, object]:
    navigate_pdf(page)
    default_limit = 50 * 1024 * 1024
    sample_size = int(fixture_entry("sample.pdf")["bytes"])
    rejected_limit = sample_size - 1
    active_limit = page.locator('[data-testid="pdf-max-file-size-active"]')
    assert active_limit.get_attribute("data-bytes") == str(default_limit)
    baseline_wasm = sum(
        1
        for item in evidence.events["responses"]
        if urlparse(str(item["url"])).path.endswith(".wasm")
    )
    baseline_workers = len(page.evaluate("window.__pdfBlobProbe.workers"))

    page.locator('[data-testid="pdf-max-file-size"]').fill(str(rejected_limit))
    page.locator('[data-testid="pdf-apply-max-file-size"]').click()
    assert active_limit.get_attribute("data-bytes") == str(rejected_limit)
    page.locator('[data-testid="pdf-load-sample"]').click()
    page.wait_for_function(
        """() => document.querySelector('[data-testid="page-status"]')?.dataset.state === 'error' &&
          document.querySelector('[data-testid="pdf-page-load-error"]')?.dataset.errorCode === 'PDF_TOO_LARGE'""",
        timeout=30_000,
    )
    assert page.locator('[data-testid="pdf-document"]').count() == 0
    assert page.locator('[data-testid="pdf-page-image"]').count() == 0
    assert page.locator('[data-testid="pdf-status"]').get_attribute("data-state") != "ready"
    error_diagnostic = page.locator(
        '[data-testid="pdf-diagnostic"][data-diagnostic-type="load-error"]'
        '[data-error-code="PDF_TOO_LARGE"]'
    ).last
    error_diagnostic.wait_for(state="visible", timeout=10_000)
    rejected_request_id = error_diagnostic.get_attribute("data-request-id")
    assert rejected_request_id
    rejected_diagnostics = page.locator(
        f'[data-testid="pdf-diagnostic"][data-request-id="{rejected_request_id}"]'
    )
    rejected_types = [
        rejected_diagnostics.nth(index).get_attribute("data-diagnostic-type")
        for index in range(rejected_diagnostics.count())
    ]
    assert "load-start" in rejected_types and "load-error" in rejected_types, rejected_types
    assert "render-start" not in rejected_types, rejected_types
    assert "render-success" not in rejected_types, rejected_types
    assert not error_diagnostic.get_attribute("data-runtime-id")
    rejected_wasm = sum(
        1
        for item in evidence.events["responses"]
        if urlparse(str(item["url"])).path.endswith(".wasm")
    )
    rejected_workers = len(page.evaluate("window.__pdfBlobProbe.workers"))
    assert rejected_wasm == baseline_wasm, {
        "before": baseline_wasm,
        "after": rejected_wasm,
    }
    assert rejected_workers == baseline_workers, {
        "before": baseline_workers,
        "after": rejected_workers,
    }
    assert_no_native_pdf_embed(page)
    screenshot(page, attempt_dir, "PDF-008-file-too-large.png")

    page.locator('[data-testid="pdf-max-file-size"]').fill(str(default_limit))
    page.locator('[data-testid="pdf-apply-max-file-size"]').click()
    assert active_limit.get_attribute("data-bytes") == str(default_limit)
    page.locator('[data-testid="pdf-load-sample"]').click()
    wait_for_rendered_page(page, 1, "sample.pdf")
    recovered = {"document": document_snapshot(page), "image": image_snapshot(page)}
    assert recovered["document"]["pageCount"] == 4, recovered
    assert recovered["image"]["naturalWidth"] > 0, recovered
    assert_no_native_pdf_embed(page)
    screenshot(page, attempt_dir, "PDF-008-size-limit-recovered.png")
    return {
        "defaultLimit": default_limit,
        "rejectedLimit": rejected_limit,
        "sampleBytes": sample_size,
        "errorCode": "PDF_TOO_LARGE",
        "rejectedRequestId": rejected_request_id,
        "rejectedDiagnosticTypes": rejected_types,
        "wasmResponsesBefore": baseline_wasm,
        "wasmResponsesAfterRejectedLoad": rejected_wasm,
        "workersBefore": baseline_workers,
        "workersAfterRejectedLoad": rejected_workers,
        "recovered": recovered,
    }


Workflow = Callable[[Page, Path, BrowserEvidence], dict[str, object]]

CASES: list[tuple[str, Workflow]] = [
    ("PDF-001", pdf_001_rendered_pages),
    ("PDF-002", pdf_002_navigation),
    ("PDF-003", pdf_003_zoom_rotate_search),
    ("PDF-004", pdf_004_fixture_matrix),
    ("PDF-005", pdf_005_verified_download),
    ("PDF-006", pdf_006_corrupted_recovery),
    ("PDF-008", pdf_008_file_size_limit),
]


def run_attempt(
    browser, case_id: str, workflow: Workflow, attempt: int
) -> dict[str, object]:
    attempt_dir = OUTPUT / case_id / f"attempt-{attempt}"
    attempt_dir.mkdir(parents=True, exist_ok=True)
    context = browser.new_context(
        viewport={"width": 1440, "height": 900},
        accept_downloads=True,
        locale="zh-CN",
        timezone_id="Asia/Shanghai",
    )
    context.add_init_script(PDF_BLOB_PROBE)
    context.tracing.start(screenshots=True, snapshots=True, sources=True)
    page = context.new_page()
    evidence = BrowserEvidence(page)
    started_at = time.time()
    try:
        details = workflow(page, attempt_dir, evidence)
        evidence.assert_clean()
        result: dict[str, object] = {
            "id": case_id,
            "attempt": attempt,
            "status": "PASS",
            "durationMs": round((time.time() - started_at) * 1000),
            "details": details,
            "events": evidence.events,
        }
    except Exception as error:
        try:
            screenshot(page, attempt_dir, "failure.png")
        except Exception:
            pass
        result = {
            "id": case_id,
            "attempt": attempt,
            "status": "FAIL",
            "durationMs": round((time.time() - started_at) * 1000),
            "error": repr(error),
            "traceback": traceback.format_exc(),
            "events": evidence.events,
        }
    finally:
        evidence.save(attempt_dir)
        context.tracing.stop(path=attempt_dir / "trace.zip")
        context.close()
    (attempt_dir / "result.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return result


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


def write_environment(extra: dict[str, object] | None = None) -> None:
    environment: dict[str, object] = {
        "commit": git_value("rev-parse", "HEAD"),
        "branch": git_value("branch", "--show-current"),
        "cwd": str(ROOT),
        "platform": platform.platform(),
        "python": sys.version,
        "baseUrl": BASE_URL,
        "mode": "formal preview",
        "fixtureManifest": str(FIXTURE_MANIFEST),
    }
    environment.update(extra or {})
    (OUTPUT / "environment.json").write_text(
        json.dumps(environment, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def main() -> int:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    write_environment()
    preview_command = (
        f"pnpm --filter demo preview --host 127.0.0.1 --port {PORT}"
    )
    (OUTPUT / "commands.log").write_text(
        f"{preview_command}\npython tests/blackbox/pdf_workflows.py\n",
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
            write_environment(
                {"browser": browser_version, "browserSource": browser_source}
            )
            results = []
            for case_id, workflow in CASES:
                attempts = [run_attempt(browser, case_id, workflow, 1)]
                if attempts[0]["status"] == "FAIL":
                    attempts.append(run_attempt(browser, case_id, workflow, 2))
                if attempts[0]["status"] == "PASS":
                    status = "PASS"
                elif attempts[-1]["status"] == "PASS":
                    status = "FLAKY"
                else:
                    status = "FAIL"
                results.append(
                    {"id": case_id, "status": status, "attempts": attempts}
                )
            browser.close()

        if any(item["status"] == "FAIL" for item in results):
            overall = "FAIL"
        elif any(item["status"] == "FLAKY" for item in results):
            overall = "FLAKY"
        else:
            overall = "PASS"
        summary = {
            "suite": "BB-PDF-VIEWER",
            "mode": "formal preview",
            "result": overall,
            "browser": browser_version,
            "browserSource": browser_source,
            "baseUrl": BASE_URL,
            "results": results,
        }
        (OUTPUT / "summary.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return 0 if overall == "PASS" else 1
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
                    "suite": "BB-PDF-VIEWER",
                    "mode": "formal preview",
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
