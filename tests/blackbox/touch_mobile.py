from __future__ import annotations

import json
import os
import signal
import socket
import subprocess
import sys
import time
import traceback
import urllib.request
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable

from playwright.sync_api import Browser, Error as PlaywrightError, Page, sync_playwright

from browser_evidence import BrowserEvidence


ROOT = Path(__file__).resolve().parents[2]
OUTPUT = Path(
    os.environ.get(
        "BLACKBOX_EVIDENCE_DIR",
        ROOT / "output" / "acceptance" / "p1-touch-mobile",
    )
).resolve()
VIEWPORT = {"width": 390, "height": 844}
DEVICE_SCALE_FACTOR = 3
TOUCH_MIN_PX = 44
LONG_PRESS_MS = 700
DOUBLE_TAP_GAP_MS = 60


def free_port() -> int:
    configured = os.environ.get("CI_TOUCH_PORT") or os.environ.get("CI_PREVIEW_PORT")
    if configured:
        return int(configured)
    with socket.socket() as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


PORT = free_port()
BASE_URL = f"http://127.0.0.1:{PORT}"


def wait_for_preview(process: subprocess.Popen[bytes]) -> None:
    deadline = time.monotonic() + 45
    while time.monotonic() < deadline:
        if process.poll() is not None:
            raise RuntimeError(f"demo preview exited early with {process.returncode}")
        try:
            with urllib.request.urlopen(BASE_URL, timeout=1) as response:
                if response.status == 200:
                    return
        except Exception:
            time.sleep(0.25)
    raise TimeoutError(f"demo preview did not become ready at {BASE_URL}")


def stop_process(process: subprocess.Popen[bytes]) -> None:
    if process.poll() is not None:
        return
    try:
        os.killpg(process.pid, signal.SIGTERM)
        process.wait(timeout=8)
    except Exception:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except Exception:
            process.kill()


def launch_browser(playwright) -> tuple[Browser, str]:
    failures: list[str] = []
    try:
        return playwright.chromium.launch(headless=True, args=["--disable-dev-shm-usage"]), "bundled-chromium"
    except PlaywrightError as error:
        failures.append(str(error))
    for candidate in (
        Path("/usr/bin/google-chrome-stable"),
        Path("/usr/bin/google-chrome"),
        Path("/usr/bin/chromium"),
        Path("/usr/bin/chromium-browser"),
    ):
        if not candidate.exists():
            continue
        try:
            return (
                playwright.chromium.launch(
                    headless=True,
                    executable_path=str(candidate),
                    args=["--disable-dev-shm-usage"],
                ),
                str(candidate),
            )
        except PlaywrightError as error:
            failures.append(f"{candidate}: {error}")
    raise RuntimeError("no Chromium executable could launch: " + " | ".join(failures))


class TouchDriver:
    """Chromium CDP touch emulation that produces real pointerType=touch events."""

    def __init__(self, page: Page):
        self.page = page
        self.cdp = page.context.new_cdp_session(page)

    def _dispatch(self, event_type: str, points: list[dict[str, float]]) -> None:
        self.cdp.send(
            "Input.dispatchTouchEvent",
            {"type": event_type, "touchPoints": points},
        )

    def touch_start(self, points: list[tuple[float, float]], first_id: int = 1) -> None:
        self._dispatch(
            "touchStart",
            [
                {"x": round(x, 2), "y": round(y, 2), "id": first_id + index, "radiusX": 2, "radiusY": 2, "force": 1}
                for index, (x, y) in enumerate(points)
            ],
        )

    def touch_move(self, points: list[tuple[float, float]], first_id: int = 1) -> None:
        self._dispatch(
            "touchMove",
            [
                {"x": round(x, 2), "y": round(y, 2), "id": first_id + index, "radiusX": 2, "radiusY": 2, "force": 1}
                for index, (x, y) in enumerate(points)
            ],
        )

    def touch_end(self) -> None:
        self._dispatch("touchEnd", [])

    def touch_cancel(self) -> None:
        self._dispatch("touchCancel", [])

    def point(self, locator, fx: float = 0.5, fy: float = 0.5) -> tuple[float, float]:
        locator.scroll_into_view_if_needed()
        self.page.wait_for_timeout(250)
        box = locator.bounding_box()
        if not box:
            raise AssertionError(f"no bounding box for {locator}")
        x = box["x"] + box["width"] * fx
        y = box["y"] + box["height"] * fy
        x = min(max(x, 20), VIEWPORT["width"] - 20)
        y = min(max(y, 110), VIEWPORT["height"] - 80)
        return x, y

    def long_press(self, x: float, y: float, hold_ms: int = LONG_PRESS_MS) -> None:
        self.touch_start([(x, y)], first_id=1)
        self.page.wait_for_timeout(hold_ms)

    def double_tap(self, x: float, y: float) -> None:
        self.touch_start([(x, y)], first_id=2)
        self.touch_end()
        self.page.wait_for_timeout(DOUBLE_TAP_GAP_MS)
        self.touch_start([(x, y)], first_id=3)
        self.touch_end()

    def pinch(self, x: float, y: float, start: float = 40, end: float = 140) -> None:
        self.touch_start([(x - start, y), (x + start, y)], first_id=4)
        for step in range(1, 6):
            offset = start + (end - start) * step / 5
            self.touch_move([(x - offset, y), (x + offset, y)], first_id=4)
            self.page.wait_for_timeout(25)
        self.touch_end()


def read_zoom(page: Page, selector: str, kind: str) -> int:
    locator = page.locator(selector).first
    value = locator.input_value() if kind == "select" else locator.inner_text()
    return int(round(float(value.strip().replace("%", ""))))


def wait_for_zoom(
    page: Page,
    selector: str,
    kind: str,
    predicate: Callable[[int], bool],
    message: str,
    timeout_ms: int = 5_000,
) -> int:
    deadline = time.monotonic() + timeout_ms / 1000
    last = None
    while time.monotonic() < deadline:
        last = read_zoom(page, selector, kind)
        if predicate(last):
            return last
        page.wait_for_timeout(50)
    raise AssertionError(f"{message}; last zoom={last}")


def status_contains(page: Page, selector: str, text: str) -> bool:
    return text in page.locator(selector).first.inner_text()


@dataclass(frozen=True)
class SurfaceSpec:
    key: str
    label: str
    route: str
    target_selector: str
    ready_selector: str
    zoom_selector: str
    zoom_kind: str
    long_press_selector: str
    long_press_text: str | None
    long_press_menu: str | None
    long_press_persists: bool = True
    setup: Callable[[Page], None] | None = None

    def wait_ready(self, page: Page) -> None:
        page.wait_for_selector(self.ready_selector, timeout=45_000)
        page.wait_for_timeout(900)

    def long_press_visible(self, page: Page) -> bool:
        if self.long_press_menu:
            menu = page.locator(self.long_press_menu)
            return menu.count() > 0 and menu.first.is_visible()
        return status_contains(page, self.long_press_selector, self.long_press_text or "")


SURFACES: tuple[SurfaceSpec, ...] = (
    SurfaceSpec(
        key="docx",
        label="DOCX Surface",
        route="/docx-surface",
        target_selector=".docx-viewer-root",
        ready_selector=".docx-viewer-root",
        zoom_selector=".zoom-value",
        zoom_kind="text",
        long_press_selector=".status-grid",
        long_press_text="右键：第 1 页",
        long_press_menu=None,
    ),
    SurfaceSpec(
        key="xlsx",
        label="XLSX Viewer",
        route="/xlsx-viewer",
        target_selector=".xlsx-viewer__body",
        ready_selector='[data-testid="xlsx-grid"]',
        zoom_selector=".xlsx-toolbar__zoom-select",
        zoom_kind="select",
        long_press_selector="",
        long_press_text=None,
        long_press_menu='[data-testid="xlsx-context-menu"]',
        setup=lambda page: page.locator('[data-testid="xlsx-read-only"]').check(),
    ),
    SurfaceSpec(
        key="pptx",
        label="PPTX Surface",
        route="/pptx-surface",
        target_selector='[data-slide-index="0"]',
        ready_selector='[data-slide-index="0"]',
        zoom_selector=".zoom-value",
        zoom_kind="text",
        long_press_selector='[data-testid="pptx-surface-status"]',
        long_press_text="右键",
        long_press_menu=None,
        long_press_persists=False,
        setup=lambda page: page.set_input_files(
            '[data-testid="pptx-surface-file-input"]',
            str(ROOT / "tests/fixtures/pptx/playback/transition-cut-fade-push-wipe.pptx"),
        ),
    ),
    SurfaceSpec(
        key="pdf",
        label="PDF Surface",
        route="/pdf-surface",
        target_selector='[data-testid="pdf-surface"]',
        ready_selector='[data-testid="pdf-surface"]',
        zoom_selector=".zoom-value",
        zoom_kind="text",
        long_press_selector=".status-grid",
        long_press_text="右键：第 1 页",
        long_press_menu=None,
        setup=lambda page: page.set_input_files(
            '[data-testid="pdf-surface-file-input"]',
            str(ROOT / "apps/demo/public/samples/sample.pdf"),
        ),
    ),
)


@dataclass(frozen=True)
class TargetSpec:
    key: str
    route: str
    ready_selector: str
    selectors: tuple[str, ...]


TARGETS: tuple[TargetSpec, ...] = (
    TargetSpec(
        key="docx",
        route="/docx-viewer",
        ready_selector='[data-testid="docx-viewer"][data-state="ready"]',
        selectors=(
            ".docx-viewer-toolbar__icon-button",
            ".docx-viewer-toolbar__zoom",
            ".docx-viewer-toolbar__search button",
        ),
    ),
    TargetSpec(
        key="xlsx",
        route="/xlsx-viewer",
        ready_selector='[data-testid="xlsx-viewer"][data-state="ready"]',
        selectors=(
            ".xlsx-toolbar__icon-button",
            ".xlsx-toolbar__zoom-select",
            ".xlsx-ribbon__icon-button",
        ),
    ),
    TargetSpec(
        key="pptx",
        route="/pptx-viewer",
        ready_selector='[data-testid="pptx-viewer"]',
        selectors=(".pptx-viewer__toolbar button",),
    ),
    TargetSpec(
        key="pdf",
        route="/pdf-viewer",
        ready_selector='[data-testid="pdf-viewer"]',
        selectors=(".pdf-toolbar button", ".pdf-toolbar select", ".pdf-toolbar input"),
    ),)


def open_surface(page: Page, spec: SurfaceSpec) -> None:
    page.goto(f"{BASE_URL}/#{spec.route}", wait_until="networkidle")
    if spec.setup:
        spec.setup(page)
    spec.wait_ready(page)


def new_page(browser: Browser) -> tuple[Page, BrowserEvidence]:
    context = browser.new_context(
        viewport=VIEWPORT,
        device_scale_factor=DEVICE_SCALE_FACTOR,
        has_touch=True,
        is_mobile=True,
        color_scheme="light",
    )
    page = context.new_page()
    return page, BrowserEvidence(page)


def case_long_press(page: Page, spec: SurfaceSpec) -> None:
    open_surface(page, spec)
    driver = TouchDriver(page)
    x, y = driver.point(page.locator(spec.target_selector))
    driver.long_press(x, y)
    try:
        if not spec.long_press_visible(page):
            raise AssertionError(f"{spec.label}: long press did not open the context menu")
    finally:
        driver.touch_end()
    page.wait_for_timeout(250)
    if spec.long_press_persists and not spec.long_press_visible(page):
        raise AssertionError(f"{spec.label}: context menu closed immediately after long press")


def case_double_tap(page: Page, spec: SurfaceSpec) -> None:
    open_surface(page, spec)
    driver = TouchDriver(page)
    initial = read_zoom(page, spec.zoom_selector, spec.zoom_kind)
    x, y = driver.point(page.locator(spec.target_selector))
    driver.double_tap(x, y)
    result = wait_for_zoom(
        page,
        spec.zoom_selector,
        spec.zoom_kind,
        lambda value: value != initial,
        f"{spec.label}: double tap did not change zoom",
    )
    if result != 200:
        raise AssertionError(f"{spec.label}: double tap expected 200%, got {result}%")


def case_pinch(page: Page, spec: SurfaceSpec) -> None:
    open_surface(page, spec)
    driver = TouchDriver(page)
    initial = read_zoom(page, spec.zoom_selector, spec.zoom_kind)
    x, y = driver.point(page.locator(spec.target_selector))
    driver.pinch(x, y)
    result = wait_for_zoom(
        page,
        spec.zoom_selector,
        spec.zoom_kind,
        lambda value: value > initial,
        f"{spec.label}: pinch did not increase zoom",
    )
    if result != 200:
        raise AssertionError(f"{spec.label}: pinch expected 200%, got {result}%")


def case_pointer_cancel(page: Page, spec: SurfaceSpec) -> None:
    open_surface(page, spec)
    driver = TouchDriver(page)
    x, y = driver.point(page.locator(spec.target_selector))
    zoom_before = read_zoom(page, spec.zoom_selector, spec.zoom_kind)

    # A cancelled long press must not open the context menu.
    driver.touch_start([(x, y)], first_id=6)
    page.wait_for_timeout(200)
    driver.touch_cancel()
    page.wait_for_timeout(LONG_PRESS_MS + 100)
    if spec.long_press_visible(page):
        raise AssertionError(f"{spec.label}: pointercancel still opened the context menu")
    if read_zoom(page, spec.zoom_selector, spec.zoom_kind) != zoom_before:
        raise AssertionError(f"{spec.label}: pointercancel changed zoom")

    # A cancelled two-finger pinch must not start a zoom gesture.
    driver.touch_start([(x - 40, y), (x + 40, y)], first_id=7)
    driver.touch_cancel()
    page.wait_for_timeout(350)
    zoom_after_cancel = read_zoom(page, spec.zoom_selector, spec.zoom_kind)
    if zoom_after_cancel != zoom_before:
        raise AssertionError(
            f"{spec.label}: cancelled pinch changed zoom {zoom_before}% -> {zoom_after_cancel}%"
        )

    # Once a pinch has produced a zoom update, pointercancel must stop further
    # updates instead of leaving the gesture stuck.
    driver.touch_start([(x - 40, y), (x + 40, y)], first_id=8)
    driver.touch_move([(x - 100, y), (x + 100, y)], first_id=8)
    page.wait_for_timeout(80)
    zoom_during = read_zoom(page, spec.zoom_selector, spec.zoom_kind)
    driver.touch_cancel()
    page.wait_for_timeout(350)
    zoom_after_active_cancel = read_zoom(page, spec.zoom_selector, spec.zoom_kind)
    if zoom_after_active_cancel != zoom_during:
        raise AssertionError(
            f"{spec.label}: pointercancel did not stop pinch zoom "
            f"({zoom_during}% -> {zoom_after_active_cancel}%)"
        )


def check_targets(page: Page, spec: TargetSpec) -> None:
    page.goto(f"{BASE_URL}/#{spec.route}", wait_until="networkidle")
    page.wait_for_selector(spec.ready_selector, timeout=45_000)
    if spec.key == "pdf":
        page.wait_for_selector(".pdf-toolbar", timeout=45_000)
    page.wait_for_timeout(600)
    if not page.evaluate("matchMedia('(pointer: coarse)').matches"):
        raise AssertionError(f"{spec.key}: mobile profile did not report coarse pointer")
    checked = 0
    for selector in spec.selectors:
        sizes = page.evaluate(
            """selector => [...document.querySelectorAll(selector)]
                 .filter(element => {
                   const style = getComputedStyle(element);
                   const rect = element.getBoundingClientRect();
                   return style.display !== "none" && style.visibility !== "hidden" &&
                     rect.width > 0 && rect.height > 0;
                 })
                 .map(element => {
                   const rect = element.getBoundingClientRect();
                   return {
                     width: Math.round(rect.width),
                     height: Math.round(rect.height),
                     label: (element.getAttribute("aria-label") || element.innerText || element.tagName).trim().slice(0, 40)
                   };
                 })""",
            selector,
        )
        if not sizes:
            raise AssertionError(f"{spec.key}: no visible touch targets matched {selector}")
        for item in sizes:
            if item["width"] < TOUCH_MIN_PX or item["height"] < TOUCH_MIN_PX:
                raise AssertionError(
                    f"{spec.key}: touch target {item['label']} is {item['width']}x{item['height']}px "
                    f"(expected >= {TOUCH_MIN_PX}px) for {selector}"
                )
            checked += 1
    if checked < 3:
        raise AssertionError(f"{spec.key}: only {checked} touch targets checked")


def main() -> int:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    cases: list[dict[str, object]] = []
    failures: list[str] = []
    evidence_items: list[tuple[str, BrowserEvidence]] = []
    preview_log = (OUTPUT / "preview.log").open("wb")

    if not (ROOT / "apps/demo/dist/index.html").exists():
        raise RuntimeError("apps/demo/dist is missing; build the demo before running the touch suite")

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
            "--strictPort",
        ],
        cwd=ROOT,
        stdout=preview_log,
        stderr=subprocess.STDOUT,
        start_new_session=True,
    )

    browser: Browser | None = None
    browser_source = ""
    browser_version = ""
    try:
        wait_for_preview(preview)
        with sync_playwright() as playwright:
            browser, browser_source = launch_browser(playwright)
            browser_version = browser.version
            for spec in SURFACES:
                for case_id, gesture, runner in (
                    (f"TOUCH-{spec.key.upper()}-LONG-PRESS", "long-press", case_long_press),
                    (f"TOUCH-{spec.key.upper()}-DOUBLE-TAP", "double-tap", case_double_tap),
                    (f"TOUCH-{spec.key.upper()}-PINCH", "pinch", case_pinch),
                    (f"TOUCH-{spec.key.upper()}-POINTERCANCEL", "pointercancel", case_pointer_cancel),
                ):
                    page, evidence = new_page(browser)
                    evidence_items.append((case_id, evidence))
                    try:
                        runner(page, spec)
                        page.screenshot(path=str(OUTPUT / f"{case_id}.png"), full_page=False)
                        cases.append(
                            {
                                "id": case_id,
                                "viewer": spec.key,
                                "label": spec.label,
                                "gesture": gesture,
                                "status": "PASS",
                            }
                        )
                        print(f"PASS: {case_id} ({spec.label} {gesture})")
                    except Exception as error:
                        cases.append(
                            {
                                "id": case_id,
                                "viewer": spec.key,
                                "label": spec.label,
                                "gesture": gesture,
                                "status": "FAIL",
                                "error": str(error),
                            }
                        )
                        failures.append(f"{case_id}: {error}")
                        print(f"FAIL: {case_id} - {error}")
                        print(traceback.format_exc())
                        try:
                            page.screenshot(path=str(OUTPUT / f"{case_id}-failure.png"), full_page=False)
                        except Exception:
                            pass
                    finally:
                        page.context.close()

            for spec in TARGETS:
                case_id = f"TOUCH-{spec.key.upper()}-TARGETS-44PX"
                page, evidence = new_page(browser)
                evidence_items.append((case_id, evidence))
                try:
                    check_targets(page, spec)
                    page.screenshot(path=str(OUTPUT / f"{case_id}.png"), full_page=False)
                    cases.append(
                        {
                            "id": case_id,
                            "viewer": spec.key,
                            "label": spec.key.upper(),
                            "gesture": "targets",
                            "status": "PASS",
                        }
                    )
                    print(f"PASS: {case_id} ({spec.key} toolbar targets >= {TOUCH_MIN_PX}px)")
                except Exception as error:
                    cases.append(
                        {
                            "id": case_id,
                            "viewer": spec.key,
                            "label": spec.key.upper(),
                            "gesture": "targets",
                            "status": "FAIL",
                            "error": str(error),
                        }
                    )
                    failures.append(f"{case_id}: {error}")
                    print(f"FAIL: {case_id} - {error}")
                    print(traceback.format_exc())
                    try:
                        page.screenshot(path=str(OUTPUT / f"{case_id}-failure.png"), full_page=False)
                    except Exception:
                        pass
                finally:
                    page.context.close()
            browser.close()
    finally:
        stop_process(preview)
        preview_log.close()

    for case_id, evidence in evidence_items:
        try:
            evidence.assert_clean()
        except Exception as error:
            failures.append(f"{case_id} browser evidence: {error}")
            print(f"FAIL: {case_id} browser evidence - {error}")
        evidence.save(OUTPUT / "browser-evidence" / case_id)

    summary = {
        "suite": "P1-TOUCH-MOBILE",
        "result": "FAIL" if failures else "PASS",
        "generatedAt": datetime.now(timezone.utc).isoformat(),
        "profile": {
            "viewport": VIEWPORT,
            "deviceScaleFactor": DEVICE_SCALE_FACTOR,
            "hasTouch": True,
            "isMobile": True,
            "pointer": "coarse",
        },
        "baseUrl": BASE_URL,
        "browser": {"source": browser_source, "version": browser_version},
        "touchMinimumPx": TOUCH_MIN_PX,
        "cases": cases,
        "failures": failures,
    }
    (OUTPUT / "summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    passed = sum(1 for case in cases if case["status"] == "PASS")
    print(f"P1-TOUCH-MOBILE {summary['result']}: {passed}/{len(cases)} cases passed")
    print(f"Evidence: {OUTPUT}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
