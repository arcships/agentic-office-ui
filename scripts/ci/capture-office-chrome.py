"""Issue #158 — deterministic responsive evidence for shared Office chrome.

The harness starts from the demo gallery (`#/chrome`) and checks the geometry
contract that static CSS assertions cannot see:

* toolbar height >= 48px, no overflow, and no overlap with the document body;
* 160px sidebar width, inline at >=768px container width and overlay below;
* the 390px toolbar wraps without horizontal page overflow;
* overlay close keeps the content in place; inline close shifts it by exactly
  the sidebar width;
* search popover, zoom select, sidebar skeleton and light/dark screenshots.

Run the demo preview first (the script never starts a server):

    pnpm --filter demo build
    pnpm --filter demo preview --port 4173
    node scripts/ci/run-python.mjs scripts/ci/capture-office-chrome.py

Artifacts default to `output/office-chrome/`.
"""

from __future__ import annotations

import argparse
import functools
import json
import sys
import threading
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any

from playwright.sync_api import Locator, Page, sync_playwright

VIEWPORTS = [
    ("1440x900", 1440, 900, False),
    ("1280x720", 1280, 720, False),
    ("768x1024", 768, 1024, True),
    ("390x844", 390, 844, True),
]
THEMES = ["light", "dark"]
DEFAULT_URL = "http://127.0.0.1:4173/#/chrome"
ROOT = Path(__file__).resolve().parents[2]


class QuietStaticHandler(SimpleHTTPRequestHandler):
    """Serve the built demo without writing one request line per asset."""

    def log_message(self, format: str, *args: Any) -> None:  # noqa: A002
        pass


def start_static_server(directory: Path) -> tuple[ThreadingHTTPServer, str]:
    """Serve `apps/demo/dist` on an ephemeral loopback port.

    The server lives in this process and is shut down in `finally`; no shell
    background job or Vite child process can leak past the run.
    """
    if not (directory / "index.html").exists():
        raise SystemExit(
            f"{directory}/index.html is missing; run `pnpm --filter demo build` first"
        )
    handler = functools.partial(QuietStaticHandler, directory=str(directory))
    server = ThreadingHTTPServer(("127.0.0.1", 0), handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    port = server.server_address[1]
    return server, f"http://127.0.0.1:{port}/#/chrome"


def box(locator: Locator) -> dict[str, float]:
    value = locator.bounding_box()
    if value is None:
        raise AssertionError(f"no bounding box for {locator}")
    return {key: float(value[key]) for key in ("x", "y", "width", "height")}


def wait_idle(page: Page, ms: int = 220) -> None:
    page.wait_for_timeout(ms)


def inject_deterministic_css(page: Page) -> Any:
    return page.add_style_tag(
        content="""
        *, *::before, *::after {
          animation-duration: 0s !important;
          animation-delay: 0s !important;
          transition-duration: 0s !important;
          transition-delay: 0s !important;
          caret-color: transparent !important;
        }
        """
    )


def assert_no_page_overflow(page: Page, label: str) -> dict[str, Any]:
    metrics = page.evaluate(
        """() => ({
          innerWidth: window.innerWidth,
          scrollWidth: document.documentElement.scrollWidth,
          bodyScrollWidth: document.body.scrollWidth,
        })"""
    )
    assert metrics["scrollWidth"] <= metrics["innerWidth"] + 1, (
        f"{label}: page overflows horizontally {metrics}"
    )
    assert metrics["bodyScrollWidth"] <= metrics["innerWidth"] + 1, (
        f"{label}: body overflows horizontally {metrics}"
    )
    return metrics


def assert_toolbar(page: Page, label: str) -> dict[str, Any]:
    toolbar = page.locator('[data-testid="chrome-gallery-toolbar"]')
    toolbar.wait_for()
    toolbar_box = box(toolbar)
    content_box = box(page.locator('[data-testid="chrome-content"]'))
    overflow = page.evaluate(
        """() => {
          const toolbar = document.querySelector('[data-testid="chrome-gallery-toolbar"]');
          const rect = toolbar.getBoundingClientRect();
          const controls = [...toolbar.querySelectorAll('button, input, [role="combobox"]')];
          const outside = controls
            .map((element) => ({ element, rect: element.getBoundingClientRect() }))
            .filter(({ rect: control }) => control.width > 0 && (
              control.left < rect.left - 1 ||
              control.right > rect.right + 1 ||
              control.top < rect.top - 1 ||
              control.bottom > rect.bottom + 1
            ))
            .map(({ element }) => element.outerHTML.slice(0, 120));
          const leading = toolbar.querySelector('.office-toolbar__leading').getBoundingClientRect();
          const trailing = toolbar.querySelector('.office-toolbar__trailing').getBoundingClientRect();
          return {
            clientWidth: toolbar.clientWidth,
            scrollWidth: toolbar.scrollWidth,
            outside,
            wrapped: trailing.top > leading.top + 1,
          };
        }"""
    )
    assert toolbar_box["height"] >= 48, f"{label}: toolbar height {toolbar_box['height']} < 48"
    assert overflow["scrollWidth"] <= overflow["clientWidth"] + 1, (
        f"{label}: toolbar overflows {overflow}"
    )
    assert not overflow["outside"], f"{label}: controls outside toolbar {overflow['outside']}"
    assert toolbar_box["y"] + toolbar_box["height"] <= content_box["y"] + 1, (
        f"{label}: toolbar overlaps the content area"
    )
    if label.startswith("390x844"):
        assert overflow["wrapped"], f"{label}: toolbar must wrap at 390px"
    if label.startswith("1440x900"):
        assert not overflow["wrapped"], f"{label}: toolbar must stay on one row at 1440px"
    return {
        "height": toolbar_box["height"],
        "width": toolbar_box["width"],
        **overflow,
    }


def assert_sidebar(page: Page, label: str) -> dict[str, Any]:
    sidebar = page.locator('[data-testid="document-viewer-sidebar"]')
    sidebar.wait_for()
    sidebar_box = box(sidebar)
    expected = page.evaluate(
        """() => {
          const sidebar = document.querySelector('[data-testid="document-viewer-sidebar"]');
          const container = sidebar.parentElement;
          return {
            containerWidth: container.getBoundingClientRect().width,
            mode: sidebar.dataset.sidebarMode,
            open: sidebar.dataset.sidebarOpen,
            width: sidebar.getBoundingClientRect().width,
          };
        }"""
    )
    expected_mode = "inline" if expected["containerWidth"] >= 768 else "overlay"
    assert expected["mode"] == expected_mode, (
        f"{label}: sidebar mode {expected['mode']} != {expected_mode} "
        f"(container {expected['containerWidth']})"
    )
    assert expected["open"] == "true", f"{label}: sidebar should start open"
    assert abs(sidebar_box["width"] - 160) <= 0.5, (
        f"{label}: sidebar width {sidebar_box['width']} != 160"
    )
    return expected


def search_flow(page: Page, out: Path, prefix: str) -> dict[str, Any]:
    shell = page.locator('[data-testid="office-chrome-shell"]')
    trigger = shell.locator('[data-testid="office-search-trigger"]')
    trigger.click()
    popover = page.locator('[data-testid="office-search-popover"]')
    popover.wait_for()
    wait_idle(page, 80)
    input_locator = page.locator('[data-testid="office-search-input"]')
    assert input_locator.evaluate("(element) => element === document.activeElement"), (
        f"{prefix}: search input did not take focus"
    )
    popover_box = box(popover)
    viewport = page.viewport_size
    assert popover_box["x"] >= -1, f"{prefix}: search popover is off-screen"
    assert popover_box["x"] + popover_box["width"] <= viewport["width"] + 1, (
        f"{prefix}: search popover overflows the viewport"
    )
    input_locator.fill("page")
    wait_idle(page, 60)
    status = page.locator('[data-testid="office-search-status"]').inner_text()
    assert status == "1 of 4", f"{prefix}: search status {status!r}"
    popover.screenshot(path=str(out / f"{prefix}-search-popover.png"))
    page.keyboard.press("Enter")
    wait_idle(page, 40)
    page.keyboard.press("Escape")
    wait_idle(page, 80)
    assert trigger.get_attribute("aria-expanded") == "false", (
        f"{prefix}: search popover did not close"
    )
    return {"status": status}


def zoom_flow(page: Page, out: Path, prefix: str) -> dict[str, Any]:
    shell = page.locator('[data-testid="office-chrome-shell"]')
    trigger = shell.locator('[data-testid="office-zoom-select"] [role="combobox"]')
    trigger.click()
    wait_idle(page, 120)
    listbox = page.locator('[role="listbox"]')
    listbox.wait_for()
    listbox.screenshot(path=str(out / f"{prefix}-zoom-options.png"))
    listbox.get_by_role("option", name="Fit width", exact=True).click()
    wait_idle(page, 120)
    fit_text = trigger.inner_text().strip()
    assert fit_text == "Fit width", f"{prefix}: fit label {fit_text!r}"
    trigger.click()
    wait_idle(page, 80)
    listbox = page.locator('[role="listbox"]')
    listbox.wait_for()
    listbox.get_by_role("option", name="150%", exact=True).click()
    wait_idle(page, 120)
    zoom_text = trigger.inner_text().strip()
    assert zoom_text == "150%", f"{prefix}: zoom label {zoom_text!r}"
    return {"fit": fit_text, "zoom": zoom_text}


def assert_accessibility(page: Page, prefix: str) -> dict[str, Any]:
    """Manual axe-substitute checks for the chrome surface.

    The repository adds `axe-core` in #165; until then this asserts the
    structural contract the component tests pin: named toolbars/landmarks,
    named icon buttons, numeric page input, labelled form controls, live
    regions, and unique ids (the state components use `useId`).
    """
    toolbar = page.locator('[data-testid="chrome-gallery-toolbar"]')
    assert toolbar.get_attribute("role") == "toolbar", f"{prefix}: toolbar role"
    assert toolbar.get_attribute("aria-label"), f"{prefix}: toolbar label"
    sidebar = page.locator('[data-testid="document-viewer-sidebar"]')
    assert sidebar.get_attribute("aria-label"), f"{prefix}: sidebar label"
    status = page.locator('[data-testid="office-status-bar"]')
    assert status.get_attribute("aria-label"), f"{prefix}: status label"

    audit = page.evaluate(
        """() => {
          const unnamedButtons = [...document.querySelectorAll('button')]
            .filter((element) => {
              const text = (element.innerText || "").trim();
              return !text && !element.getAttribute("aria-label") && !element.getAttribute("title");
            })
            .map((element) => element.outerHTML.slice(0, 120));
          const unlabelledFields = [...document.querySelectorAll('input, select, textarea')]
            .filter((element) => {
              if (element.type === "hidden") return false;
              if (element.getAttribute("aria-label")) return false;
              return !(element.labels && element.labels.length);
            })
            .map((element) => element.outerHTML.slice(0, 120));
          const ids = [...document.querySelectorAll('[id]')].map((element) => element.id);
          const duplicates = ids.filter((id, index) => ids.indexOf(id) !== index);
          const liveRegions = [...document.querySelectorAll('[role="status"], [role="alert"]')]
            .map((element) => (element.innerText || "").trim());
          const pageInput = document.querySelector('[data-testid="office-page-input"]');
          return {
            unnamedButtons,
            unlabelledFields,
            duplicates,
            emptyLiveRegions: liveRegions.filter((text) => !text),
            pageInputLabel: pageInput?.getAttribute("aria-label") ?? "",
            pageInputMode: pageInput?.getAttribute("inputmode") ?? "",
          };
        }"""
    )
    assert not audit["unnamedButtons"], f"{prefix}: unnamed buttons {audit['unnamedButtons']}"
    assert not audit["unlabelledFields"], f"{prefix}: unlabelled fields {audit['unlabelledFields']}"
    assert not audit["duplicates"], f"{prefix}: duplicate ids {audit['duplicates']}"
    assert not audit["emptyLiveRegions"], f"{prefix}: empty live regions"
    assert audit["pageInputLabel"], f"{prefix}: page input has no accessible name"
    assert audit["pageInputMode"] == "numeric", f"{prefix}: page input mode"
    return audit


def page_flow(page: Page, prefix: str) -> dict[str, Any]:
    shell = page.locator('[data-testid="office-chrome-shell"]')
    input_locator = shell.locator('[data-testid="office-page-input"]')
    input_locator.fill("5")
    input_locator.press("Enter")
    wait_idle(page, 60)
    assert input_locator.input_value() == "5", f"{prefix}: typed page did not commit"
    input_locator.press("End")
    wait_idle(page, 40)
    assert input_locator.input_value() == "12", f"{prefix}: End did not jump to the last page"
    input_locator.press("Home")
    wait_idle(page, 40)
    assert input_locator.input_value() == "1", f"{prefix}: Home did not jump to page 1"
    input_locator.fill("99")
    input_locator.press("Enter")
    wait_idle(page, 40)
    assert input_locator.input_value() == "12", f"{prefix}: page input did not clamp"
    input_locator.fill("3")
    input_locator.press("Enter")
    wait_idle(page, 40)
    return {"typed": 5, "clamped": 12}


def xlsx_story_flow(page: Page, out: Path, prefix: str) -> dict[str, Any]:
    """Screenshot the XLSX toolbar/ribbon/status story and check its geometry.

    Mirrors the shared-toolbar contract from ``assert_toolbar``: a 48px row, no
    horizontal overflow, and the controls the toolbar/ribbon must keep. The
    spreadsheet body below the chrome keeps the remaining height because the
    toolbar is ``flex: 0 0 auto``.
    """
    story = page.locator('[data-testid="xlsx-chrome-story"]')
    story.wait_for()
    story.locator('[data-testid="xlsx-toolbar"]').wait_for()
    story.locator('[data-testid="xlsx-format-toolbar"]').wait_for()
    story.locator('[data-testid="xlsx-status-bar"]').wait_for()
    metrics = page.evaluate(
        """() => {
          const story = document.querySelector('[data-testid="xlsx-chrome-story"]');
          const toolbar = story.querySelector('[data-testid="xlsx-toolbar"]');
          const ribbon = story.querySelector('[data-testid="xlsx-format-toolbar"]');
          const status = story.querySelector('[data-testid="xlsx-status-bar"]');
          const rect = (element) => {
            const box = element.getBoundingClientRect();
            return { height: box.height, width: box.width, top: box.top, bottom: box.bottom };
          };
          const toolbarTop = toolbar.getBoundingClientRect().top;
          const paddingTop = parseFloat(getComputedStyle(toolbar).paddingTop) || 0;
          const cluster =
            toolbar.querySelector('.xlsx-toolbar__zoom') ||
            toolbar.querySelector('.office-toolbar__trailing');
          const unnamed = [...story.querySelectorAll('button')]
            .filter((element) => {
              const text = (element.innerText || '').trim();
              return !text && !element.getAttribute('aria-label') && !element.getAttribute('title');
            })
            .map((element) => element.outerHTML.slice(0, 120));
          return {
            toolbar: rect(toolbar),
            ribbon: rect(ribbon),
            status: rect(status),
            toolbarOverflow: toolbar.scrollWidth - toolbar.clientWidth,
            ribbonOverflow: ribbon.scrollWidth - ribbon.clientWidth,
            wrapped: Boolean(cluster) &&
              cluster.getBoundingClientRect().top > toolbarTop + paddingTop + 4,
            iconButtons: toolbar.querySelectorAll('.office-icon-button').length,
            ribbonIconButtons: ribbon.querySelectorAll('.office-icon-button').length,
            unnamedButtons: unnamed,
          };
        }"""
    )
    assert metrics["toolbar"]["height"] >= 48, f"{prefix}: xlsx toolbar height {metrics}"
    assert metrics["toolbarOverflow"] <= 1, f"{prefix}: xlsx toolbar overflows {metrics}"
    assert metrics["ribbonOverflow"] <= 1, f"{prefix}: xlsx ribbon overflows {metrics}"
    assert metrics["iconButtons"] >= 4, f"{prefix}: xlsx toolbar icon controls {metrics}"
    assert metrics["ribbonIconButtons"] >= 10, f"{prefix}: xlsx ribbon icon controls {metrics}"
    assert not metrics["unnamedButtons"], f"{prefix}: xlsx unnamed buttons {metrics['unnamedButtons']}"
    story.screenshot(path=str(out / f"{prefix}-xlsx-chrome.png"))
    story.locator('[data-testid="xlsx-format-toolbar"]').screenshot(
        path=str(out / f"{prefix}-xlsx-ribbon.png")
    )
    return metrics


def sidebar_flow(page: Page, out: Path, prefix: str, mode: str) -> dict[str, Any]:
    shell = page.locator('[data-testid="office-chrome-shell"]')
    toggle = shell.locator('[data-testid="chrome-sidebar-toggle"]')
    content = shell.locator('[data-testid="chrome-content"]')
    before = box(content)
    shell_before = box(page.locator('[data-testid="office-chrome-shell"]'))

    # Animations are still enabled here (no reduced-motion emulation yet).
    transition = page.evaluate(
        """() => {
          const sidebar = document.querySelector('[data-testid="document-viewer-sidebar"]');
          return {
            duration: getComputedStyle(sidebar).transitionDuration,
            ready: sidebar.dataset.transitionsReady,
          };
        }"""
    )
    assert transition["ready"] == "true", f"{prefix}: sidebar transitions not ready"
    assert "0.2s" in transition["duration"], (
        f"{prefix}: sidebar transition duration {transition['duration']!r}"
    )

    toggle.click()
    wait_idle(page, 300)
    sidebar = page.locator('[data-testid="document-viewer-sidebar"]')
    assert sidebar.get_attribute("data-sidebar-open") == "false", (
        f"{prefix}: sidebar did not close"
    )
    after = box(content)
    shell_after = box(page.locator('[data-testid="office-chrome-shell"]'))
    delta = after["x"] - before["x"]
    if mode == "overlay":
        assert abs(delta) <= 1.5, f"{prefix}: overlay close moved content by {delta}"
    else:
        assert abs(delta + 160) <= 1.5, f"{prefix}: inline close moved content by {delta}"
    assert abs(shell_after["height"] - shell_before["height"]) <= 1.5, (
        f"{prefix}: shell height jumped on sidebar close"
    )
    page.locator('[data-testid="office-chrome-shell"]').screenshot(
        path=str(out / f"{prefix}-sidebar-closed.png")
    )

    toggle.click()
    wait_idle(page, 300)
    assert sidebar.get_attribute("data-sidebar-open") == "true", (
        f"{prefix}: sidebar did not reopen"
    )
    return {"transition": transition, "closedContentDelta": delta}


def skeleton_flow(page: Page, out: Path, prefix: str) -> dict[str, Any]:
    checkbox = page.locator('[data-testid="chrome-gallery-sidebar-loading"]')
    checkbox.check()
    wait_idle(page, 120)
    skeleton = page.locator('[data-testid="document-viewer-sidebar-skeleton"]')
    skeleton.wait_for()
    assert page.locator('[data-slot="sidebar-skeleton-item"]').count() == 3
    page.locator('[data-testid="document-viewer-sidebar"]').screenshot(
        path=str(out / f"{prefix}-sidebar-skeleton.png")
    )
    checkbox.uncheck()
    wait_idle(page, 120)
    return {"items": 3}


def coarse_pointer(page: Page, prefix: str) -> dict[str, Any]:
    result = page.evaluate(
        """() => {
          const button = document.querySelector('[data-testid="chrome-sidebar-toggle"]');
          const after = getComputedStyle(button, "::after");
          return {
            coarse: matchMedia("(pointer: coarse)").matches,
            width: parseFloat(after.width) || 0,
            height: parseFloat(after.height) || 0,
          };
        }"""
    )
    if result["coarse"]:
        assert result["width"] >= 44 and result["height"] >= 44, (
            f"{prefix}: coarse hit area {result}"
        )
    return result


def capture_theme(
    browser: Any,
    name: str,
    width: int,
    height: int,
    touch: bool,
    theme: str,
    url: str,
    out_root: Path,
    errors: list[str],
) -> dict[str, Any]:
    """Capture one viewport/theme case in a fresh context.

    A fresh page per theme keeps `(pointer: coarse)` true for the whole case:
    Chromium can fall back to a fine primary pointer after mouse-driven
    interactions, which would otherwise shrink the touch controls halfway
    through the matrix.
    """
    context = browser.new_context(
        viewport={"width": width, "height": height},
        device_scale_factor=1,
        has_touch=touch,
        is_mobile=False,
    )
    page = context.new_page()
    prefix = f"{name}-{theme}"
    theme_out = out_root / name / theme
    theme_out.mkdir(parents=True, exist_ok=True)
    page.on(
        "console",
        lambda message: errors.append(f"{prefix}: {message.text}")
        if message.type == "error"
        else None,
    )
    page.on("pageerror", lambda error: errors.append(f"{prefix}: {error}"))

    try:
        page.goto(url, wait_until="networkidle")
        page.wait_for_selector('[data-testid="office-chrome-gallery"]')
        page.evaluate("document.fonts.ready")
        wait_idle(page, 250)
        page.locator('[data-testid="chrome-gallery-theme"]').select_option(theme)
        wait_idle(page, 180)
        assert page.locator(f'[data-office-theme="{theme}"]').count() >= 1, (
            f"{prefix}: theme attribute not applied"
        )

        overflow = assert_no_page_overflow(page, prefix)
        toolbar = assert_toolbar(page, prefix)
        sidebar = assert_sidebar(page, prefix)
        pointer = coarse_pointer(page, prefix)
        accessibility = assert_accessibility(page, prefix)
        page.locator('[data-testid="chrome-gallery-toolbar"]').screenshot(
            path=str(theme_out / f"{prefix}-toolbar.png")
        )
        page.locator('[data-testid="document-viewer-sidebar"]').screenshot(
            path=str(theme_out / f"{prefix}-sidebar-open.png")
        )
        page.locator('[data-testid="office-chrome-states"]').screenshot(
            path=str(theme_out / f"{prefix}-states.png")
        )

        # The transition contract is checked before motion is reduced; the
        # injected stylesheet then makes the screenshots deterministic.
        sidebar_interaction = sidebar_flow(page, theme_out, prefix, sidebar["mode"])
        deterministic = inject_deterministic_css(page)
        skeleton = skeleton_flow(page, theme_out, prefix)
        search = search_flow(page, theme_out, prefix)
        zoom = zoom_flow(page, theme_out, prefix)
        page_control = page_flow(page, prefix)

        page.locator('[data-testid="chrome-gallery-status"]').select_option("error")
        wait_idle(page, 80)
        shell = page.locator('[data-testid="office-chrome-shell"]')
        assert shell.locator('[data-testid="office-status-bar-error"]').count() == 1
        shell.locator('[data-testid="office-status-bar"]').screenshot(
            path=str(theme_out / f"{prefix}-status-error.png")
        )
        page.locator('[data-testid="chrome-gallery-status"]').select_option("normal")
        wait_idle(page, 60)

        # Spreadsheet chrome story (#160): toolbar + ribbon + status bar in the
        # same themed subtree, one screenshot per viewport/theme.
        story = xlsx_story_flow(page, theme_out, prefix)

        page.screenshot(path=str(theme_out / f"{prefix}-full-page.png"), full_page=True)
        deterministic.evaluate("(element) => element.remove()")
        return {
            "viewport": name,
            "width": width,
            "height": height,
            "theme": theme,
            "page": overflow,
            "toolbar": toolbar,
            "sidebar": sidebar,
            "sidebarInteraction": sidebar_interaction,
            "skeleton": skeleton,
            "search": search,
            "zoom": zoom,
            "pageControl": page_control,
            "accessibility": accessibility,
            "pointer": pointer,
            "xlsxStory": story,
        }
    finally:
        context.close()


def capture_viewport(
    browser: Any,
    name: str,
    width: int,
    height: int,
    touch: bool,
    url: str,
    out_root: Path,
    errors: list[str],
) -> list[dict[str, Any]]:
    return [
        capture_theme(
            browser,
            name,
            width,
            height,
            touch,
            theme,
            url,
            out_root,
            errors,
        )
        for theme in THEMES
    ]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--url", default=DEFAULT_URL)
    parser.add_argument(
        "--serve",
        default="",
        help="directory to serve in-process; when set, --url is replaced by an ephemeral loopback URL",
    )
    parser.add_argument(
        "--out",
        default=str(ROOT / "output" / "office-chrome"),
    )
    args = parser.parse_args()

    out_root = Path(args.out)
    out_root.mkdir(parents=True, exist_ok=True)
    errors: list[str] = []
    records: list[dict[str, Any]] = []
    server: ThreadingHTTPServer | None = None

    try:
        if args.serve:
            server, args.url = start_static_server(Path(args.serve).resolve())
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=True)
            try:
                for name, width, height, touch in VIEWPORTS:
                    records.extend(
                        capture_viewport(
                            browser,
                            name,
                            width,
                            height,
                            touch,
                            args.url,
                            out_root,
                            errors,
                        )
                    )
            finally:
                browser.close()
    finally:
        if server is not None:
            server.shutdown()
            server.server_close()

    summary = {
        "suite": "P1-OFFICE-CHROME-RESPONSIVE",
        "result": "FAIL" if errors else "PASS",
        "url": args.url,
        "viewports": [name for name, _, _, _ in VIEWPORTS],
        "themes": THEMES,
        "cases": records,
        "errors": errors,
    }
    (out_root / "summary.json").write_text(
        json.dumps(summary, indent=2) + "\n", encoding="utf-8"
    )
    if errors:
        print("console errors:", file=sys.stderr)
        for error in errors:
            print(f"  {error}", file=sys.stderr)
        return 1
    print(f"office chrome evidence: {out_root}")
    print(f"cases: {len(records)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
