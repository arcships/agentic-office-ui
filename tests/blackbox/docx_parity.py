from __future__ import annotations

import json
from io import BytesIO
import os
from pathlib import Path
import signal
import socket
import subprocess
import time
import traceback
import urllib.parse
import urllib.request

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright
from PIL import Image, ImageChops, ImageEnhance, ImageFilter

from browser_evidence import BrowserEvidence


ROOT = Path(__file__).resolve().parents[2]
OUTPUT = Path(
    os.environ.get(
        "BLACKBOX_EVIDENCE_DIR",
        ROOT / "output" / "acceptance" / "p2-docx-render-parity",
    )
).resolve()
PORT = int(os.environ.get("CI_DOCX_PARITY_PREVIEW_PORT", "4178"))
BASE_URL = f"http://127.0.0.1:{PORT}"

CASES = (
    {"id": "DOCX-PARITY-LEGAL", "file": "legal-contract.docx"},
    {"id": "DOCX-PARITY-INVOICE", "file": "invoice-table.docx"},
    {"id": "DOCX-PARITY-IMAGE", "file": "report-with-image.docx"},
    {"id": "DOCX-PARITY-CHINESE", "file": "chinese-mixed.docx"},
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


def write_json(path: Path, value: object) -> None:
    path.write_text(
        json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def capture_visual_diff(page, attempt_dir: Path) -> dict[str, object]:
    roots = {
        side: page.locator(f'[data-testid="parity-{side}"] .docx-viewer-root')
        for side in ("viewer", "editor")
    }
    boxes = {side: locator.bounding_box() for side, locator in roots.items()}
    if not boxes["viewer"] or not boxes["editor"]:
        return {"status": "FAIL", "reason": "one or both document roots are missing"}

    common_width = int(round(min(boxes["viewer"]["width"], boxes["editor"]["width"])))
    common_height = int(
        round(min(640, boxes["viewer"]["height"], boxes["editor"]["height"]))
    )
    if common_width <= 0 or common_height <= 0:
        return {
            "status": "FAIL",
            "reason": "one or both document roots collapsed",
            "boxes": boxes,
        }

    images: dict[str, Image.Image] = {}
    for side, box in boxes.items():
        screenshot = page.screenshot(
            path=str(attempt_dir / f"{side}-visual.png"),
            clip={
                "x": round(box["x"]),
                "y": round(box["y"]),
                "width": common_width,
                "height": common_height,
            },
        )
        images[side] = Image.open(BytesIO(screenshot)).convert("RGB")

    blur_radius = 1
    difference = ImageChops.difference(
        images["viewer"].filter(ImageFilter.GaussianBlur(blur_radius)),
        images["editor"].filter(ImageFilter.GaussianBlur(blur_radius)),
    )
    ImageEnhance.Brightness(difference).enhance(4).save(attempt_dir / "visual-diff.png")
    threshold = 20
    changed_pixels = sum(
        1 for pixel in difference.getdata() if max(pixel) > threshold
    )
    total_pixels = common_width * common_height
    return {
        "status": "PASS",
        "boxes": boxes,
        "commonWidth": common_width,
        "commonHeight": common_height,
        "threshold": threshold,
        "comparisonBlurRadius": blur_radius,
        "changedPixels": changed_pixels,
        "differenceRatio": changed_pixels / total_pixels,
    }


def combined_text(structure: dict[str, object]) -> str:
    pages = structure.get("pages", [])
    return " ".join(str(page.get("text", "")) for page in pages)


def collect_legal_page_samples(page, attempt_dir: Path, page_count: int) -> list[dict[str, object]]:
    scroll_root = page.locator('[data-testid="parity-editor"] .docx-viewer-root')
    if scroll_root.count() == 0 or page_count <= 0:
        write_json(attempt_dir / "legal-page-samples.json", [])
        return []
    targets = (
        ("first", 0),
        ("middle", max(0, (page_count - 1) // 2)),
        ("last", max(0, page_count - 1)),
    )
    samples: list[dict[str, object]] = []
    for label, target_index in targets:
        scroll_root.evaluate(
            """(root, args) => {
              const max = Math.max(0, root.scrollHeight - root.clientHeight);
              root.scrollTop = args.count <= 1 ? 0 : max * args.index / (args.count - 1);
              root.dispatchEvent(new Event('scroll'));
            }""",
            {"count": page_count, "index": target_index},
        )
        page.wait_for_timeout(180)
        sample = scroll_root.evaluate(
            r"""(root, target) => {
              const wrappers = [...root.querySelectorAll('[data-docx-page-wrapper="true"]')];
              const exact = wrappers.find((item) => Number(item.dataset.docxPageIndex) === target);
              return exact ? {
                index: Number(exact.dataset.docxPageIndex),
                text: (exact.innerText || '').replace(/\s+/g, ' ').trim(),
              } : { index: null, text: '' };
            }""",
            target_index,
        )
        sample["label"] = label
        sample["targetIndex"] = target_index
        samples.append(sample)
        scroll_root.screenshot(path=str(attempt_dir / f"editor-{label}.png"))
    write_json(attempt_dir / "legal-page-samples.json", samples)
    return samples


def validate_case(
    case: dict[str, str],
    viewer: dict[str, object],
    editor: dict[str, object],
    diff: dict[str, object],
    legal_samples: list[dict[str, object]],
    visual: dict[str, object],
    evidence: BrowserEvidence,
) -> list[str]:
    failures: list[str] = []

    def expect(condition: bool, message: str) -> None:
        if not condition:
            failures.append(message)

    expect(diff.get("equal") is True, f"render structures differ: {diff.get('differences')}")
    interaction = editor.get("interaction", {})
    expect(interaction.get("contentEditableTrue") == 0, "read-only editor has contenteditable=true")
    expect(interaction.get("inputs") == 0, "read-only editor exposes input elements")
    expect(interaction.get("selects") == 0, "read-only editor exposes select elements")
    expect(interaction.get("resizeHandles") == 0, "read-only editor exposes resize/edit handles")
    expect(visual.get("status") == "PASS", f"visual capture failed: {visual}")
    expect(visual.get("commonWidth", 0) >= 200, "document surface width collapsed")
    expect(visual.get("commonHeight", 0) >= 500, "document surface height collapsed")
    expect(
        float(visual.get("differenceRatio", 1)) <= 0.01,
        f"Viewer/Editor visual difference ratio is {visual.get('differenceRatio')}",
    )

    file_name = case["file"]
    if file_name == "legal-contract.docx":
        viewer_pages = viewer.get("pages", [])
        editor_pages = editor.get("pages", [])
        expect(len(viewer_pages) >= 3, f"Viewer rendered only {len(viewer_pages)} legal pages")
        expect(len(editor_pages) == len(viewer_pages), "Editor did not expose every legal page")
        for label, pages in (("Viewer", viewer_pages), ("Editor", editor_pages)):
            if len(pages) >= 3:
                indexes = (0, (len(pages) - 1) // 2, len(pages) - 1)
                texts = [str(pages[index].get("text", "")).strip() for index in indexes]
                expect(all(texts), f"{label} legal first/middle/last page has empty text")
                expect(len(set(texts)) == 3, f"{label} legal first/middle/last page repeats content")
        expect(len(legal_samples) == 3, "legal scroll evidence is incomplete")
        if len(legal_samples) == 3:
            expect(
                [item.get("index") for item in legal_samples]
                == [item.get("targetIndex") for item in legal_samples],
                "Editor virtual pages did not reach first/middle/last targets",
            )
            sampled_texts = [str(item.get("text", "")).strip() for item in legal_samples]
            expect(all(sampled_texts), "Editor scrolled legal page sample has empty text")
            expect(len(set(sampled_texts)) == 3, "Editor scrolled legal pages repeat content")
    elif file_name == "invoice-table.docx":
        for label, structure in (("Viewer", viewer), ("Editor", editor)):
            totals = structure.get("totals", {})
            expect(totals.get("cells") == 34, f"{label} invoice cell count is {totals.get('cells')}, expected 34")
            text = combined_text(structure)
            for amount in ("$216.00", "$1,068.00", "$2,694.00", "$2,882.58"):
                expect(amount in text, f"{label} invoice is missing {amount}")
    elif file_name == "report-with-image.docx":
        for label, structure in (("Viewer", viewer), ("Editor", editor)):
            image_count = structure.get("totals", {}).get("images")
            expect(image_count == 1, f"{label} image count is {image_count}, expected exactly 1")
    elif file_name == "chinese-mixed.docx":
        required = (
            "中文、English、数字与标点混排测试",
            "合同编号：CN-EXT-2026-0705",
            "上海凌云科技有限公司",
            "人民币 1,234,567.89 元",
        )
        for label, structure in (("Viewer", viewer), ("Editor", editor)):
            text = combined_text(structure)
            for expected in required:
                expect(expected in text, f"{label} mixed-language text is missing {expected}")

    violations = evidence.violations()
    if any(violations.values()):
        failures.append(
            "browser evidence policy violations: "
            + json.dumps(violations, ensure_ascii=False, sort_keys=True)
        )
    return failures


def run_attempt(browser, case: dict[str, str], attempt: int) -> dict[str, object]:
    attempt_dir = OUTPUT / case["id"] / f"attempt-{attempt}"
    attempt_dir.mkdir(parents=True, exist_ok=True)
    context = browser.new_context(viewport={"width": 1720, "height": 1000})
    context.tracing.start(screenshots=True, snapshots=True, sources=True)
    page = context.new_page()
    evidence = BrowserEvidence(page)
    try:
        query = urllib.parse.urlencode({"sample": case["file"]})
        page.goto(f"{BASE_URL}/#/docx-parity?{query}", wait_until="networkidle")
        page.wait_for_function(
            """() => ['ready', 'error'].includes(
              document.querySelector('[data-testid="page-status"]')?.dataset.state
            )""",
            timeout=60_000,
        )
        state = page.get_by_test_id("page-status").get_attribute("data-state")
        if state != "ready":
            raise AssertionError(
                f"parity page did not become ready: {page.get_by_test_id('load-error').inner_text()}"
            )

        viewer = json.loads(page.get_by_test_id("viewer-structure").text_content() or "")
        editor = json.loads(page.get_by_test_id("editor-structure").text_content() or "")
        diff = json.loads(page.get_by_test_id("parity-diff").text_content() or "")
        write_json(attempt_dir / "viewer-structure.json", viewer)
        write_json(attempt_dir / "editor-structure.json", editor)
        write_json(attempt_dir / "diff.json", diff)
        write_json(
            attempt_dir / "structure.json",
            {"file": case["file"], "viewer": viewer, "editor": editor, "diff": diff},
        )

        page.wait_for_timeout(120)
        visual = capture_visual_diff(page, attempt_dir)
        write_json(attempt_dir / "visual-metrics.json", visual)

        legal_samples: list[dict[str, object]] = []
        if case["file"] == "legal-contract.docx":
            legal_samples = collect_legal_page_samples(
                page, attempt_dir, int(editor.get("pageCount", 0))
            )

        page.get_by_test_id("parity-viewer").screenshot(path=str(attempt_dir / "viewer.png"))
        page.get_by_test_id("parity-editor").screenshot(path=str(attempt_dir / "editor.png"))
        page.screenshot(path=str(attempt_dir / "full.png"), full_page=True)
        failures = validate_case(
            case, viewer, editor, diff, legal_samples, visual, evidence
        )
        result: dict[str, object] = {
            "id": case["id"],
            "file": case["file"],
            "attempt": attempt,
            "status": "PASS" if not failures else "FAIL",
            "failures": failures,
            "viewerTotals": viewer.get("totals"),
            "editorTotals": editor.get("totals"),
            "visual": visual,
        }
    except Exception as error:
        try:
            page.screenshot(path=str(attempt_dir / "failure.png"), full_page=True)
        except Exception:
            pass
        result = {
            "id": case["id"],
            "file": case["file"],
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
        ["pnpm", "--filter", "demo", "preview", "--host", "127.0.0.1", "--port", str(PORT)],
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
            case_results: list[dict[str, object]] = []
            for case in CASES:
                attempts = [run_attempt(browser, case, 1)]
                if attempts[0]["status"] == "FAIL":
                    attempts.append(run_attempt(browser, case, 2))
                status = "PASS" if attempts[0]["status"] == "PASS" else (
                    "FLAKY" if attempts[-1]["status"] == "PASS" else "FAIL"
                )
                case_results.append(
                    {"id": case["id"], "file": case["file"], "status": status, "attempts": attempts}
                )
            browser.close()
        passed = all(case["status"] == "PASS" for case in case_results)
        summary = {
            "suite": "P2-DOCX-RENDER-01",
            "mode": "formal preview",
            "result": "PASS" if passed else "FAIL",
            "browser": browser_version,
            "browserSource": browser_source,
            "cases": case_results,
        }
        write_json(OUTPUT / "summary.json", summary)
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
                "suite": "P2-DOCX-RENDER-01",
                "result": "BLOCKED",
                "error": repr(error),
                "traceback": traceback.format_exc(),
            },
        )
        raise SystemExit(2)
