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


def open_pdf_actions(page: Page) -> None:
    """Open the upstream `...` actions menu that owns the extra PDF commands."""
    menu = page.locator('[data-testid="pdf-actions-menu"]')
    if menu.count() == 0:
        page.locator('[data-testid="pdf-actions"]').click()
    menu.wait_for(state="visible", timeout=10_000)


def click_pdf_action(page: Page, test_id: str) -> None:
    open_pdf_actions(page)
    page.locator(f'[data-testid="{test_id}"]').click()


def open_pdf_search(page: Page) -> None:
    input_locator = page.locator('[data-testid="pdf-search-input"]')
    if input_locator.count() == 0:
        page.locator('[data-testid="pdf-search-toggle"]').click()
    input_locator.wait_for(state="visible", timeout=10_000)


def pdf_download_disabled(page: Page) -> bool:
    open_pdf_actions(page)
    return page.locator('[data-testid="pdf-download"]').is_disabled()


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


def text_layer_text(page: Page) -> str:
    """Read the text layer as contiguous text.

    `innerText` inserts a line break for every absolutely positioned glyph
    span even though the spans are in reading order. `textContent` is the DOM's
    actual extraction path and keeps words and sentences contiguous.
    """
    return page.locator('[data-testid="pdf-text-layer"]').evaluate(
        "element => element.textContent || ''"
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


def page_stage_snapshot(page: Page) -> dict[str, object]:
    """The stage reports the rendered page's intrinsic /Rotate separately from
    the user rotation exposed on `pdf-document[data-rotation]`."""
    return page.locator('[data-testid="pdf-page"]').evaluate(
        """element => ({
          pageNumber: Number(element.dataset.pageNumber),
          renderState: element.dataset.renderState,
          pageRotation: Number(element.dataset.pageRotation ?? 0),
        })"""
    )


# A blank white canvas satisfies `naturalWidth > 0`; sampling the decoded bitmap
# is the only way to prove that text or image pixels were actually rasterized.
MIN_TEXT_NON_WHITE_RATIO = 0.01
MIN_TEXT_DARK_RATIO = 0.002


def pixel_sample(page: Page) -> dict[str, object]:
    return page.locator('[data-testid="pdf-page-image"]').evaluate(
        """image => {
          const sampleWidth = Math.min(480, image.naturalWidth);
          const sampleHeight = Math.max(
            1,
            Math.round((image.naturalHeight * sampleWidth) / image.naturalWidth),
          );
          const canvas = document.createElement('canvas');
          canvas.width = sampleWidth;
          canvas.height = sampleHeight;
          const context = canvas.getContext('2d', { willReadFrequently: true });
          if (!context) throw new Error('2D canvas is unavailable for pixel sampling');
          context.drawImage(image, 0, 0, sampleWidth, sampleHeight);
          const pixels = context.getImageData(0, 0, sampleWidth, sampleHeight).data;
          const total = sampleWidth * sampleHeight;
          let nonWhite = 0;
          let dark = 0;
          let ink = 0;
          for (let index = 0; index < pixels.length; index += 4) {
            const luminance =
              (pixels[index] * 299 + pixels[index + 1] * 587 + pixels[index + 2] * 114) / 1000;
            if (luminance < 245) nonWhite += 1;
            if (luminance < 128) dark += 1;
            ink = (ink * 31 + Math.round(luminance)) % 4294967291;
          }
          return {
            sampleWidth,
            sampleHeight,
            nonWhiteRatio: nonWhite / total,
            darkRatio: dark / total,
            luminanceDigest: ink,
          };
        }"""
    )


def assert_rendered_pixels(page: Page, label: str) -> dict[str, object]:
    sample = pixel_sample(page)
    assert sample["nonWhiteRatio"] > MIN_TEXT_NON_WHITE_RATIO, {
        "label": label,
        "message": "rendered bitmap is effectively blank",
        "sample": sample,
    }
    assert sample["darkRatio"] > MIN_TEXT_DARK_RATIO, {
        "label": label,
        "message": "rendered bitmap has no dark ink pixels",
        "sample": sample,
    }
    return sample


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
    stage = page_stage_snapshot(page)
    assert document["pageCount"] == 4, document
    assert document["activePage"] == 1, document
    assert document["rotation"] == 0, document
    assert stage["pageRotation"] == 0, stage
    assert image["naturalWidth"] > 0 and image["naturalHeight"] > 0, image
    assert str(image["src"]).startswith("blob:"), image
    pixels = assert_rendered_pixels(page, "sample.pdf page 1")
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
    return {
        "document": document,
        "image": image,
        "pixels": pixels,
        "stage": stage,
        "thumbnail": thumbnail,
        "wasm": wasm,
    }


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

    click_pdf_action(page, "pdf-next-page")
    capture("next", 2)
    page.locator('[data-testid="pdf-page-input"]').fill("3")
    page.locator('[data-testid="pdf-page-input"]').press("Enter")
    capture("page-input", 3)
    click_pdf_action(page, "pdf-previous-page")
    capture("previous", 2)
    click_pdf_action(page, "pdf-last-page")
    capture("last", 4)
    click_pdf_action(page, "pdf-first-page")
    capture("first", 1)
    # The viewer now follows the upstream default (sidebar closed). Open it
    # before navigating through a thumbnail so the button is not behind the
    # stage's pointer interception.
    page.locator('[data-testid="pdf-sidebar-toggle"]').click()
    page.locator('[data-testid="pdf-thumbnail"][data-page-number="4"]').wait_for()
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
    zoomed_pixels = assert_rendered_pixels(page, "sample.pdf page 1 at 200%")
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
    assert document_snapshot(page)["rotation"] == 90, document_snapshot(page)
    rotated_pixels = assert_rendered_pixels(page, "sample.pdf page 1 rotated by the viewer")
    screenshot(page, attempt_dir, "PDF-003-rotated-clockwise.png")

    open_pdf_search(page)
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
        "zoomedPixels": zoomed_pixels,
        "beforeRotate": before_rotate,
        "rotated": rotated,
        "rotatedPixels": rotated_pixels,
        "searchResult": search_result,
        "searchContext": context_text,
    }


def pdf_004_fixture_matrix(
    page: Page, attempt_dir: Path, _evidence: BrowserEvidence
) -> dict[str, object]:
    navigate_pdf(page)
    results: dict[str, object] = {}

    select_sample(page, "scanned-invoice.pdf")
    scanned = {
        "document": document_snapshot(page),
        "image": image_snapshot(page),
        "page": page_stage_snapshot(page),
        "textLayers": page.locator('[data-testid="pdf-text-layer"]').count(),
        "pixels": assert_rendered_pixels(page, "scanned-invoice.pdf page 1"),
    }
    assert scanned["document"]["pageCount"] == 3, scanned
    assert scanned["image"]["naturalWidth"] > 0, scanned
    # The fixture is image-only: there is no native text layer, so selecting
    # text must be unavailable instead of fabricating selectable glyphs.
    assert scanned["textLayers"] == 0, {
        "message": "image-only fixture exposed a text layer",
        "detail": scanned,
    }
    assert page.locator('[data-testid="pdf-selection-layer"]').count() == 1, scanned
    screenshot(page, attempt_dir, "PDF-004-scanned-invoice.png")
    results["scanned-invoice.pdf"] = scanned

    select_sample(page, "rotated-pages.pdf")
    portrait = image_snapshot(page)
    portrait_stage = page_stage_snapshot(page)
    portrait_document = document_snapshot(page)
    click_pdf_action(page, "pdf-next-page")
    wait_for_rendered_page(page, 2, "rotated-pages.pdf")
    landscape = image_snapshot(page)
    landscape_stage = page_stage_snapshot(page)
    landscape_document = document_snapshot(page)
    assert portrait_document["rotation"] == 0, portrait_document
    assert portrait_stage["pageRotation"] == 0, portrait_stage
    assert portrait["naturalHeight"] > portrait["naturalWidth"], portrait
    # Page two carries real /Rotate 90 metadata. The viewer must report and
    # apply the intrinsic rotation instead of inferring orientation from the
    # MediaBox or the user rotation control.
    assert landscape_document["rotation"] == 0, landscape_document
    assert landscape_stage["pageRotation"] == 90, landscape_stage
    assert landscape["naturalWidth"] > landscape["naturalHeight"], landscape
    screenshot(page, attempt_dir, "PDF-004-rotated-pages-rotate-90.png")
    results["rotated-pages.pdf"] = {
        "document": landscape_document,
        "portrait": portrait,
        "portraitDocument": portrait_document,
        "portraitStage": portrait_stage,
        "landscape": landscape,
        "landscapeStage": landscape_stage,
    }

    select_sample(page, "large-contract.pdf")
    large = document_snapshot(page)
    assert large["pageCount"] == 240, large
    assert page.locator('[data-testid="pdf-thumbnail"]').count() == 240
    # Wait for the 240 thumbnail bitmaps before switching pages: they share the
    # PDFium worker queue, so measuring page 240 while they render would race.
    page.wait_for_function(
        """expected => {
          const images = [...document.querySelectorAll('[data-testid="pdf-thumbnail-image"]')];
          return images.length === expected &&
            images.every((image) => image.complete && image.naturalWidth > 0);
        }""",
        arg=240,
        timeout=120_000,
    )
    click_pdf_action(page, "pdf-last-page")
    wait_for_rendered_page(page, 240, "large-contract.pdf")
    last_page = image_snapshot(page)
    assert last_page["naturalWidth"] > 0 and last_page["naturalHeight"] > 0, last_page
    screenshot(page, attempt_dir, "PDF-004-large-contract-page-240.png")
    results["large-contract.pdf"] = {
        "document": document_snapshot(page),
        "lastPage": last_page,
        "thumbnailCount": 240,
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
    open_pdf_actions(page)
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
    open_pdf_search(page)
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
    assert page.locator('[data-testid="pdf-actions"]').is_disabled()
    assert_no_native_pdf_embed(page)
    screenshot(page, attempt_dir, "PDF-006-corrupted-error.png")

    select_sample(page, "sample.pdf")
    recovered = {"document": document_snapshot(page), "image": image_snapshot(page)}
    assert recovered["document"]["pageCount"] == 4, recovered
    assert recovered["document"]["activePage"] == 1, recovered
    assert page.locator('[data-testid="pdf-thumbnail"]').count() == 4
    open_pdf_actions(page)
    assert not page.locator('[data-testid="pdf-download"]').is_disabled()
    page.keyboard.press("Escape")
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
    # Changing the public limit reloads the active source automatically; the
    # rejection must happen without opening another engine or retaining a page.
    # The loader rejects from Content-Length before reading the body, so the
    # browser reports that deliberate stream cancellation as ERR_ABORTED.
    evidence.allow_request_abort(
        f"{BASE_URL}/samples/sample.pdf",
        reason="PDF_TOO_LARGE rejects the response before reading the body",
    )
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
    # Restoring the limit also reloads automatically, exercising recovery from
    # the same public configuration instead of a duplicated manual load.
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
        "requestAborts": evidence.expected_request_aborts,
        "wasmResponsesBefore": baseline_wasm,
        "wasmResponsesAfterRejectedLoad": rejected_wasm,
        "workersBefore": baseline_workers,
        "workersAfterRejectedLoad": rejected_workers,
        "recovered": recovered,
    }


def pdf_009_encrypted_document(
    page: Page, attempt_dir: Path, _evidence: BrowserEvidence
) -> dict[str, object]:
    navigate_pdf(page)
    password_input = page.locator('[data-testid="pdf-password"]')
    password_input.fill("")
    page.locator('[data-testid="pdf-sample-select"]').select_option("encrypted.pdf")
    page.wait_for_function(
        """() => document.querySelector('[data-testid="page-status"]')?.dataset.state === 'error'""",
        timeout=45_000,
    )
    error = page.locator('[data-testid="pdf-page-load-error"]')
    error.wait_for(state="visible", timeout=10_000)
    error_code = error.get_attribute("data-error-code")
    error_text = error.inner_text()
    assert error_code == "ENCRYPTED", (error_code, error_text)
    assert page.locator('[data-testid="pdf-document"]').count() == 0
    assert page.locator('[data-testid="pdf-thumbnail"]').count() == 0
    assert page.locator('[data-testid="pdf-page-image"]').count() == 0
    assert page.locator('[data-testid="pdf-actions"]').is_disabled()
    diagnostic = page.locator(
        '[data-testid="pdf-diagnostic"][data-diagnostic-type="load-error"]'
        '[data-error-code="ENCRYPTED"]'
    ).last
    diagnostic.wait_for(state="visible", timeout=10_000)
    screenshot(page, attempt_dir, "PDF-009-encrypted-locked.png")

    # The documented fixture password unlocks the same bytes through the
    # public viewer prop; recovery must not require a page reload.
    password_input.fill("arcships")
    page.locator('[data-testid="pdf-load-sample"]').click()
    wait_for_rendered_page(page, 1, "encrypted.pdf")
    unlocked = {
        "document": document_snapshot(page),
        "image": image_snapshot(page),
        "pixels": assert_rendered_pixels(page, "encrypted.pdf page 1"),
    }
    assert unlocked["document"]["pageCount"] == 2, unlocked
    assert unlocked["document"]["activePage"] == 1, unlocked
    text_layer = page.locator('[data-testid="pdf-text-layer"]')
    text_layer.wait_for(state="visible", timeout=10_000)
    unlocked_text = text_layer_text(page)
    assert "Encrypted Contract 1" in unlocked_text, unlocked_text
    assert "Encrypted page 1 row 1" in unlocked_text, unlocked_text
    assert "\n" not in unlocked_text, {
        "message": "text layer extraction must stay contiguous",
        "text": unlocked_text[:200],
    }
    assert page.locator('[data-testid="pdf-page-load-error"]').count() == 0
    assert_no_native_pdf_embed(page)
    screenshot(page, attempt_dir, "PDF-009-encrypted-unlocked.png")
    return {
        "errorCode": error_code,
        "error": error_text,
        "unlocked": unlocked,
        "text": unlocked_text,
    }


def pdf_010_surface_error_state(
    page: Page, attempt_dir: Path, _evidence: BrowserEvidence
) -> dict[str, object]:
    """The surface route owns no password input, so a locked PDF must surface
    the stable error in the page UI instead of only logging to the console."""
    page.goto(f"{BASE_URL}/#/pdf-surface", wait_until="networkidle")
    status = page.locator('[data-testid="pdf-surface-status"]')
    status.wait_for(state="visible", timeout=30_000)
    assert status.get_attribute("data-state") == "idle", status.get_attribute("data-state")
    page.locator('[data-testid="pdf-surface-empty"]').wait_for(state="visible", timeout=10_000)

    page.locator('[data-testid="pdf-surface-sample-select"]').select_option("encrypted.pdf")
    error = page.locator('[data-testid="pdf-surface-page-load-error"]')
    error.wait_for(state="visible", timeout=45_000)
    error_code = error.get_attribute("data-error-code")
    error_text = error.inner_text()
    assert error_code == "ENCRYPTED", (error_code, error_text)
    assert status.get_attribute("data-state") == "error", status.get_attribute("data-state")
    assert page.locator('.pdf-surface__pages').count() == 0
    screenshot(page, attempt_dir, "PDF-010-surface-encrypted-error.png")

    # Image-only pages must not gain a selectable text layer in the surface
    # either, and a successful load clears the page-level error.
    page.locator('[data-testid="pdf-surface-sample-select"]').select_option(
        "scanned-invoice.pdf"
    )
    page.wait_for_function(
        """() => document.querySelector('[data-testid="pdf-surface-status"]')?.dataset.state === 'ready'""",
        timeout=45_000,
    )
    page.locator('.pdf-surface__page-img').first.wait_for(state="visible", timeout=45_000)
    page.wait_for_timeout(750)
    text_layers = page.locator('.pdf-surface__text-layer').count()
    assert text_layers == 0, {
        "message": "image-only surface page exposed a text layer",
        "textLayers": text_layers,
    }
    assert page.locator('[data-testid="pdf-surface-page-load-error"]').count() == 0
    screenshot(page, attempt_dir, "PDF-010-surface-scanned.png")
    return {
        "errorCode": error_code,
        "errorText": error_text,
        "textLayers": text_layers,
        "status": status.get_attribute("data-state"),
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
    ("PDF-009", pdf_009_encrypted_document),
    ("PDF-010", pdf_010_surface_error_state),
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
