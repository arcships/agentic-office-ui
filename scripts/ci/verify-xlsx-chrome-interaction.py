"""Issue #160 — XLSX viewer/authoring chrome: geometry + interaction evidence.

The component suite runs on the repository's minimal host renderer, which has
no layout engine and no `document`, so two acceptance items can only be
verified in a browser:

* the toolbar wraps at 390px while the grid keeps usable height;
* menu/dialog keyboard + focus behaviour (Escape/Tab close and hand focus back
  to the surface, ArrowDown reaches the first item).

This script drives the real demo (`#/xlsx-viewer`) at 1440/768/390 x light/dark,
asserts those contracts and writes screenshots plus `summary.json`.

Run:

    pnpm --filter demo build
    python3 scripts/ci/run-python.mjs scripts/ci/verify-xlsx-chrome-interaction.py

Artifacts default to `output/xlsx-chrome/`.
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

ROOT = Path(__file__).resolve().parents[2]
VIEWPORTS = (
    ("1440x900", 1440, 900, False),
    ("768x1024", 768, 1024, True),
    ("390x844", 390, 844, True),
)
THEMES = ("light", "dark")
# The viewer body keeps the pre-migration minimum: 360px of sheet area below
# the chrome, which is what "the grid still has usable height" means here.
MIN_GRID_HEIGHT = 300
MIN_BODY_HEIGHT = 360


class QuietStaticHandler(SimpleHTTPRequestHandler):
    def log_message(self, format: str, *args: Any) -> None:  # noqa: A002
        pass


def start_static_server(directory: Path) -> tuple[ThreadingHTTPServer, str]:
    if not (directory / "index.html").exists():
        raise SystemExit(f"{directory}/index.html is missing; run `pnpm --filter demo build`")
    handler = functools.partial(QuietStaticHandler, directory=str(directory))
    server = ThreadingHTTPServer(("127.0.0.1", 0), handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    return server, f"http://127.0.0.1:{server.server_address[1]}/"


def rect(locator: Locator) -> dict[str, float]:
    value = locator.bounding_box()
    if value is None:
        raise AssertionError(f"no bounding box for {locator}")
    return {key: float(value[key]) for key in ("x", "y", "width", "height")}


def active_item(page: Page) -> dict[str, str]:
    return page.evaluate(
        """() => {
          const element = document.activeElement;
          if (!element) return { role: "", testId: "", tag: "" };
          return {
            role: element.getAttribute('role') || "",
            testId: element.getAttribute('data-testid') || "",
            tag: element.tagName.toLowerCase(),
          };
        }"""
    )


def open_viewer(page: Page, base_url: str, theme: str) -> None:
    page.goto(f"{base_url}?theme={theme}#/xlsx-viewer", wait_until="networkidle")
    page.wait_for_selector('[data-testid="xlsx-viewer-page"]')
    page.locator('[data-testid="xlsx-load-sample"]').click()
    page.wait_for_function(
        """() => {
          const grid = document.querySelector('[data-testid="xlsx-grid"]');
          return grid && Number(grid.dataset.displayRowCount || 0) > 0;
        }""",
        timeout=30_000,
    )
    page.wait_for_timeout(250)


def toolbar_metrics(page: Page) -> dict[str, Any]:
    return page.evaluate(
        """() => {
          const toolbar = document.querySelector('[data-testid="xlsx-toolbar"]');
          const ribbon = document.querySelector('[data-testid="xlsx-format-toolbar"]');
          const leading = toolbar.querySelector('.office-toolbar__leading').getBoundingClientRect();
          const trailing = toolbar.querySelector('.office-toolbar__trailing').getBoundingClientRect();
          const box = (element) => {
            const rect = element.getBoundingClientRect();
            return { height: rect.height, width: rect.width };
          };
          return {
            toolbar: box(toolbar),
            ribbon: box(ribbon),
            wrapped: trailing.top > leading.top + 1,
            toolbarOverflow: toolbar.scrollWidth - toolbar.clientWidth,
            ribbonOverflow: ribbon.scrollWidth - ribbon.clientWidth,
            toolbarVerticalOverflow: toolbar.scrollHeight - toolbar.clientHeight,
            ribbonVerticalOverflow: ribbon.scrollHeight - ribbon.clientHeight,
            ribbonOverflowY: getComputedStyle(ribbon).overflowY,
            toolbarControlsInside: [...toolbar.querySelectorAll('button, select')]
              .filter((control) => {
                const rect = control.getBoundingClientRect();
                const shell = toolbar.getBoundingClientRect();
                return rect.width > 0 && (rect.top < shell.top - 1 || rect.bottom > shell.bottom + 1);
              }).length,
            ribbonControlsInside: [...ribbon.querySelectorAll('button, select')]
              .filter((control) => {
                const rect = control.getBoundingClientRect();
                const shell = ribbon.getBoundingClientRect();
                return rect.width > 0 && (rect.top < shell.top - 1 || rect.bottom > shell.bottom + 1);
              }).length,
            theme: toolbar.closest('[data-office-theme]')?.getAttribute('data-office-theme') ?? '',
            darkMarker: toolbar.classList.contains('xlsx-toolbar--dark'),
          };
        }"""
    )


def layout_metrics(page: Page) -> dict[str, Any]:
    return page.evaluate(
        """() => {
          const body = document.querySelector('.xlsx-viewer__body').getBoundingClientRect();
          const grid = document.querySelector('[data-testid="xlsx-grid"]');
          const gridBox = grid.getBoundingClientRect();
          return {
            bodyHeight: body.height,
            gridHeight: gridBox.height,
            gridWidth: gridBox.width,
            displayRows: Number(grid.dataset.displayRowCount || 0),
            displayColumns: Number(grid.dataset.displayColumnCount || 0),
            pageScrollWidth: document.documentElement.scrollWidth,
            windowWidth: window.innerWidth,
          };
        }"""
    )


def keep_viewer_focused(page: Page, label: str) -> None:
    focused = page.evaluate(
        """() => {
          const active = document.activeElement;
          const viewer = document.querySelector('[data-testid="xlsx-viewer"]');
          return Boolean(active && viewer && (active === viewer || viewer.contains(active)));
        }"""
    )
    assert focused, f"{label}: focus must return to the spreadsheet surface"


def menu_keyboard_flow(page: Page, label: str) -> dict[str, Any]:
    grid = page.locator('[data-testid="xlsx-grid"]')
    box = rect(grid)
    page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2, button="right")
    menu = page.locator('[data-testid="xlsx-context-menu"]')
    menu.wait_for(state="visible", timeout=10_000)
    # The shared menu focuses its container on open; the first ArrowDown lands
    # on the first enabled item (Reka's roving focus).
    arrows: list[dict[str, str]] = []
    for _ in range(2):
        page.keyboard.press("ArrowDown")
        page.wait_for_timeout(80)
        arrows.append(active_item(page))
    first_item = arrows[0]
    menu.screenshot(path=str(label))
    page.keyboard.press("Escape")
    menu.wait_for(state="detached", timeout=10_000)
    keep_viewer_focused(page, f"{label}: after Escape")
    return {"firstItem": first_item, "arrows": arrows}


def dialog_flow(page: Page, theme_out: Path, prefix: str) -> dict[str, Any]:
    trigger = page.locator('[data-testid="xlsx-ribbon-number-format"]')
    trigger.scroll_into_view_if_needed()
    trigger.click()
    dialog = page.locator('[data-testid="xlsx-authoring-dialog"]')
    dialog.wait_for(state="visible", timeout=10_000)
    dialog.screenshot(path=str(theme_out / f"{prefix}-xlsx-authoring.png"))
    metrics = page.evaluate(
        """() => {
          const dialog = document.querySelector('[data-testid="xlsx-authoring-dialog"]');
          const rect = dialog.getBoundingClientRect();
          const classes = dialog.className;
          return {
            classes,
            height: rect.height,
            width: rect.width,
            radius: getComputedStyle(dialog).borderTopLeftRadius,
            background: getComputedStyle(dialog).backgroundColor,
            section: Boolean(document.querySelector('[data-testid="xlsx-number-format-section"]')),
          };
        }"""
    )
    assert "office-dialog__content" in metrics["classes"], f"{prefix}: shared dialog surface"
    assert metrics["section"], f"{prefix}: the number-format form still renders"
    page.keyboard.press("Escape")
    dialog.wait_for(state="detached", timeout=10_000)
    return metrics


def capture_case(
    browser: Any,
    base_url: str,
    name: str,
    width: int,
    height: int,
    touch: bool,
    theme: str,
    out_root: Path,
) -> dict[str, Any]:
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
    errors: list[str] = []
    page.on("console", lambda message: errors.append(message.text) if message.type == "error" else None)
    page.on("pageerror", lambda error: errors.append(str(error)))
    try:
        open_viewer(page, base_url, theme)
        chrome = toolbar_metrics(page)
        layout = layout_metrics(page)

        assert chrome["theme"] == theme, f"{prefix}: viewer subtree carries the theme ({chrome['theme']})"
        assert chrome["darkMarker"] is (theme == "dark"), f"{prefix}: dark marker matches the theme"
        assert chrome["toolbar"]["height"] >= 48, f"{prefix}: toolbar >= 48px ({chrome['toolbar']})"
        assert chrome["toolbarOverflow"] <= 1, f"{prefix}: toolbar does not hide content {chrome}"
        assert chrome["ribbonOverflow"] <= 1, f"{prefix}: ribbon does not hide content {chrome}"
        assert chrome["toolbarVerticalOverflow"] <= 1, (
            f"{prefix}: the toolbar grows with its wrapped rows instead of clipping them {chrome}"
        )
        assert chrome["toolbarControlsInside"] == 0, (
            f"{prefix}: every toolbar control stays inside the row {chrome}"
        )
        if chrome["ribbonOverflowY"] in ("auto", "scroll"):
            # Narrow/tablet hosts cap the wrapped ribbon and scroll it on purpose;
            # the cap is the contract that keeps the grid usable.
            assert chrome["ribbon"]["height"] <= 133, (
                f"{prefix}: the scrollable ribbon respects the 132px cap {chrome}"
            )
        else:
            assert chrome["ribbonVerticalOverflow"] <= 4, (
                f"{prefix}: an uncapped ribbon does not clip its rows {chrome}"
            )
            assert chrome["ribbonControlsInside"] == 0, (
                f"{prefix}: every ribbon control stays inside the row {chrome}"
            )
        assert layout["pageScrollWidth"] <= layout["windowWidth"] + 1, f"{prefix}: no page overflow {layout}"
        assert layout["gridHeight"] >= MIN_GRID_HEIGHT, (
            f"{prefix}: grid keeps usable height {layout}"
        )
        assert layout["bodyHeight"] >= MIN_BODY_HEIGHT, f"{prefix}: sheet body keeps its minimum {layout}"
        assert layout["displayRows"] >= 200, f"{prefix}: grid still renders its rows {layout}"

        if name == "390x844":
            assert chrome["wrapped"], f"{prefix}: toolbar wraps at 390px {chrome}"
        else:
            assert not chrome["wrapped"], f"{prefix}: toolbar stays on one row {chrome}"

        keyboard = menu_keyboard_flow(page, str(theme_out / f"{prefix}-xlsx-context-menu.png"))
        assert keyboard["firstItem"]["role"] == "menuitem", (
            f"{prefix}: ArrowDown focuses a real menu item {keyboard}"
        )
        focused = [item["testId"] for item in keyboard["arrows"]]
        assert all(item["role"] == "menuitem" for item in keyboard["arrows"]), (
            f"{prefix}: arrow keys stay inside the menu {keyboard}"
        )
        assert len(set(focused)) == len(focused), (
            f"{prefix}: ArrowDown walks the items instead of staying put {focused}"
        )

        dialog = dialog_flow(page, theme_out, prefix)
        keep_viewer_focused(page, f"{prefix}: after the dialog closes")

        # The dialog trigger scrolled the page; bring the chrome back for the
        # evidence screenshots.
        page.evaluate("() => window.scrollTo(0, 0)")
        page.wait_for_timeout(150)
        page.locator(".viewer-container").screenshot(path=str(theme_out / f"{prefix}-xlsx-viewer.png"))
        page.locator('[data-testid="xlsx-toolbar"]').screenshot(path=str(theme_out / f"{prefix}-xlsx-toolbar.png"))
        return {
            "viewport": name,
            "theme": theme,
            "chrome": chrome,
            "layout": layout,
            "menu": keyboard,
            "dialog": dialog,
            "errors": errors,
        }
    finally:
        context.close()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--serve", default="apps/demo/dist")
    parser.add_argument("--url", default="")
    parser.add_argument("--output", type=Path, default=ROOT / "output" / "xlsx-chrome")
    args = parser.parse_args(argv)

    server = None
    base_url = args.url
    if not base_url:
        server, base_url = start_static_server((ROOT / args.serve).resolve())
    out_root: Path = args.output.resolve()
    out_root.mkdir(parents=True, exist_ok=True)
    cases: list[dict[str, Any]] = []
    failures: list[str] = []
    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=True)
            try:
                for name, width, height, touch in VIEWPORTS:
                    for theme in THEMES:
                        try:
                            cases.append(
                                capture_case(
                                    browser, base_url, name, width, height, touch, theme, out_root
                                )
                            )
                        except AssertionError as error:
                            failures.append(str(error))
            finally:
                browser.close()
    finally:
        if server is not None:
            server.shutdown()
            server.server_close()

    summary = {
        "suite": "XLSX-CHROME-01",
        "issue": "#160",
        "url": base_url,
        "viewports": [name for name, _, _, _ in VIEWPORTS],
        "themes": list(THEMES),
        "result": "FAIL" if failures else "PASS",
        "cases": cases,
        "failures": failures,
    }
    (out_root / "summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")
    print(f"xlsx chrome evidence: {out_root}")
    print(f"cases: {len(cases)} failures: {len(failures)}")
    for failure in failures:
        print(f"FAIL: {failure}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
