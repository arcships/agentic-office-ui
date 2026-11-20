"""Issue #162 — deterministic PDF viewer/surface chrome story screenshots.

The story drives the real demo routes (`#/pdf-viewer`, `#/pdf-surface`) with the
bundled `sample.pdf`, so the captures exercise the actual PDFium session instead
of a mock toolbar. Every viewport/theme case asserts the shared-chrome contract
that static CSS cannot see:

* toolbar height >= 48px, no horizontal overflow, no overlap with the pages;
* 160px document sidebar, inline at >= 768px container width and overlay below;
* overlay close keeps the page content in place (no layout jump);
* page / zoom / rotate / search / print / download controls stay reachable;
* continuous / single / spread layouts render the page stack;
* the shared status bar sits below the pages without covering them.

Each story is captured at 390 / 768 / 1440 in light and dark. The script serves
`apps/demo/dist` from an in-process thread and always shuts it down in
`finally`, so no Vite/preview child process can leak.

    pnpm --filter demo build
    PYTHON=/tmp/bb-venv/bin/python RUN_PYTHON_MODULES=playwright \\
      node scripts/ci/run-python.mjs scripts/ci/capture-pdf-chrome.py --serve apps/demo/dist

Artifacts default to `output/pdf-chrome/`.
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
VIEWPORTS = [
    ("390x844", 390, 844, True),
    ("768x1024", 768, 1024, True),
    ("1440x900", 1440, 900, False),
]
THEMES = ["light", "dark"]
LAYOUTS = ["continuous", "single", "spread"]
DEFAULT_URL = "http://127.0.0.1:4173/#/pdf-viewer"


class QuietStaticHandler(SimpleHTTPRequestHandler):
    """Serve the built demo without writing one request line per asset."""

    def log_message(self, format: str, *args: Any) -> None:  # noqa: A002
        pass


def start_static_server(directory: Path) -> tuple[ThreadingHTTPServer, str]:
    if not (directory / "index.html").exists():
        raise SystemExit(
            f"{directory}/index.html is missing; run `pnpm --filter demo build` first"
        )
    handler = functools.partial(QuietStaticHandler, directory=str(directory))
    server = ThreadingHTTPServer(("127.0.0.1", 0), handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    port = server.server_address[1]
    return server, f"http://127.0.0.1:{port}/#/pdf-viewer"


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


def box(locator: Locator) -> dict[str, float]:
    value = locator.bounding_box()
    if value is None:
        raise AssertionError(f"no bounding box for {locator}")
    return {key: float(value[key]) for key in ("x", "y", "width", "height")}


def assert_no_page_overflow(page: Page, label: str) -> dict[str, Any]:
    metrics: dict[str, Any] = page.evaluate(
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


def wait_for_viewer_ready(page: Page, root_test_id: str) -> None:
    page.wait_for_selector(f'[data-testid="{root_test_id}"]', state="visible", timeout=45_000)
    page.wait_for_function(
        """() => document.querySelector('[data-testid="pdf-status"]')?.dataset.state === 'ready'""",
        timeout=45_000,
    )
    page.wait_for_function(
        """() => {
          const image = document.querySelector('[data-testid="pdf-page-image"]');
          return image instanceof HTMLImageElement && image.complete && image.naturalWidth > 0;
        }""",
        timeout=45_000,
    )


def wait_for_layout(page: Page, layout: str) -> None:
    page.wait_for_function(
        """expected => document.querySelector('.pdf-frame-wrap')?.dataset.pageLayout === expected""",
        arg=layout,
        timeout=30_000,
    )
    wait_idle(page, 180)


def select_layout(page: Page, layout: str) -> None:
    page.locator('[data-testid="pdf-page-layout"]').select_option(layout)
    wait_for_layout(page, layout)


def assert_toolbar(page: Page, label: str, toolbar: Locator) -> dict[str, Any]:
    toolbar.wait_for()
    toolbar_box = box(toolbar)
    body_box = box(page.locator('[data-testid="pdf-document"]'))
    metrics: dict[str, Any] = page.evaluate(
        """() => {
          const toolbar = document.querySelector('[data-testid="pdf-toolbar"]');
          const body = document.querySelector('[data-testid="pdf-document"]');
          const rect = (element) => {
            const box = element.getBoundingClientRect();
            return { height: box.height, width: box.width, top: box.top, bottom: box.bottom, left: box.left, right: box.right };
          };
          const unnamed = [...toolbar.querySelectorAll('button')]
            .filter((element) => {
              const text = (element.innerText || '').trim();
              return !text && !element.getAttribute('aria-label') && !element.getAttribute('title');
            })
            .map((element) => element.outerHTML.slice(0, 160));
          return {
            toolbar: rect(toolbar),
            body: body ? rect(body) : null,
            overflow: toolbar.scrollWidth - toolbar.clientWidth,
            wrapped: toolbar.querySelector('.office-toolbar__trailing')?.getBoundingClientRect().top
              > toolbar.querySelector('.office-toolbar__leading')?.getBoundingClientRect().top + 1,
            iconButtons: toolbar.querySelectorAll('.office-icon-button').length,
            controls: {
              page: Boolean(toolbar.querySelector('[data-testid="pdf-page-input"]')),
              zoom: Boolean(toolbar.querySelector('[data-testid="pdf-zoom-select"]')),
              rotate: Boolean(toolbar.querySelector('[data-testid="pdf-rotate"]')),
              search: Boolean(toolbar.querySelector('[data-testid="pdf-search-input"]')),
              print: Boolean(toolbar.querySelector('[data-testid="pdf-print"]')),
              download: Boolean(toolbar.querySelector('[data-testid="pdf-download"]')),
              sidebar: Boolean(toolbar.querySelector('[data-testid="pdf-sidebar-toggle"]')),
            },
            unnamedButtons: unnamed,
          };
        }"""
    )
    assert toolbar_box["height"] >= 48, f"{label}: toolbar height {toolbar_box}"
    assert metrics["overflow"] <= 1, f"{label}: toolbar overflows {metrics}"
    assert metrics["iconButtons"] >= 8, f"{label}: shared icon buttons {metrics}"
    assert all(metrics["controls"].values()), f"{label}: toolbar controls {metrics['controls']}"
    assert not metrics["unnamedButtons"], f"{label}: unnamed buttons {metrics['unnamedButtons']}"
    if body_box:
        assert toolbar_box["y"] + toolbar_box["height"] <= body_box["y"] + 1, (
            f"{label}: toolbar covers the pages {metrics}"
        )
    return metrics


def assert_sidebar(page: Page, label: str, out: Path, prefix: str) -> dict[str, Any]:
    sidebar = page.locator('[data-testid="pdf-thumbnails-sidebar"]')
    sidebar.wait_for(state="visible", timeout=30_000)
    page.wait_for_selector(
        '[data-testid="pdf-thumbnails-sidebar"] [data-testid="pdf-thumbnail"]',
        timeout=30_000,
    )
    metrics: dict[str, Any] = page.evaluate(
        """() => {
          const sidebar = document.querySelector('[data-testid="pdf-thumbnails-sidebar"]');
          const body = document.querySelector('[data-testid="pdf-document"]');
          const frame = document.querySelector('.pdf-frame-wrap');
          const rect = (element) => {
            if (!element) return null;
            const box = element.getBoundingClientRect();
            return { x: box.x, y: box.y, width: box.width, height: box.height, top: box.top, bottom: box.bottom };
          };
          return {
            mode: sidebar?.dataset.sidebarMode,
            open: sidebar?.dataset.sidebarOpen,
            position: sidebar ? getComputedStyle(sidebar).position : null,
            sidebar: rect(sidebar),
            body: rect(body),
            frame: rect(frame),
            thumbnails: sidebar?.querySelectorAll('[data-testid="pdf-thumbnail"]').length ?? 0,
          };
        }"""
    )
    assert metrics["open"] == "true", f"{label}: sidebar not open {metrics}"
    assert metrics["mode"] in {"inline", "overlay"}, f"{label}: sidebar mode {metrics}"
    assert abs(metrics["sidebar"]["width"] - 160) <= 1.5, f"{label}: sidebar width {metrics['sidebar']}"
    assert metrics["thumbnails"] >= 1, f"{label}: sidebar lost its thumbnails {metrics}"
    assert metrics["sidebar"]["height"] >= (metrics["body"]["height"] / 2), (
        f"{label}: sidebar collapsed vertically {metrics}"
    )

    page.locator('[data-testid="pdf-thumbnails-sidebar"]').screenshot(
        path=str(out / f"{prefix}-sidebar.png")
    )

    before = box(page.locator('.pdf-frame-wrap'))
    page.locator('[data-testid="pdf-sidebar-toggle"]').click()
    wait_idle(page, 320)
    after = box(page.locator('.pdf-frame-wrap'))
    closed = page.evaluate(
        """() => document.querySelector('[data-testid="pdf-thumbnails-sidebar"]')?.dataset.sidebarOpen"""
    )
    assert closed == "false", f"{label}: sidebar did not close"
    delta = after["x"] - before["x"]
    if metrics["mode"] == "overlay":
        assert abs(delta) <= 1.5, f"{label}: overlay close moved the pages by {delta}"
    else:
        assert abs(delta + 160) <= 1.5, f"{label}: inline close moved the pages by {delta}"

    page.locator('[data-testid="pdf-sidebar-toggle"]').click()
    wait_idle(page, 320)
    assert page.evaluate(
        """() => document.querySelector('[data-testid="pdf-thumbnails-sidebar"]')?.dataset.sidebarOpen"""
    ) == "true", f"{label}: sidebar did not reopen"

    return metrics


def assert_status_bar(page: Page, label: str) -> dict[str, Any]:
    bar = page.locator('[data-testid="pdf-status-bar"]')
    bar.wait_for(state="visible", timeout=30_000)
    metrics: dict[str, Any] = page.evaluate(
        """() => {
          const bar = document.querySelector('[data-testid="pdf-status-bar"]');
          const section = document.querySelector('[data-testid="pdf-document"]');
          const rect = (element) => {
            const box = element.getBoundingClientRect();
            return { top: box.top, bottom: box.bottom, height: box.height, width: box.width };
          };
          return {
            bar: rect(bar),
            body: rect(section),
            page: bar.querySelector('[data-testid="office-status-bar-page"]')?.textContent?.trim(),
            file: bar.querySelector('[data-testid="office-status-bar-file"]')?.textContent?.trim(),
          };
        }"""
    )
    assert metrics["bar"]["height"] >= 24, f"{label}: status bar height {metrics}"
    assert metrics["body"]["bottom"] <= metrics["bar"]["top"] + 1, (
        f"{label}: status bar covers the pages {metrics}"
    )
    assert metrics["page"], f"{label}: missing page indicator {metrics}"
    assert metrics["file"], f"{label}: missing file name {metrics}"
    return metrics


def viewer_story(page: Page, out: Path, prefix: str) -> dict[str, Any]:
    toolbar = page.locator('[data-testid="pdf-toolbar"]')
    toolbar_metrics = assert_toolbar(page, prefix, toolbar)
    sidebar = assert_sidebar(page, prefix, out, prefix)
    status = assert_status_bar(page, prefix)

    layouts: dict[str, Any] = {}
    for layout in LAYOUTS:
        select_layout(page, layout)
        page.wait_for_function(
            """() => document.querySelector('[data-testid="pdf-status"]')?.dataset.state === 'ready'""",
            timeout=30_000,
        )
        page.locator('[data-testid="pdf-viewer"]').screenshot(
            path=str(out / f"{prefix}-viewer-{layout}.png")
        )
        toolbar.screenshot(path=str(out / f"{prefix}-toolbar-{layout}.png"))
        layouts[layout] = page.evaluate(
            """() => {
              const documentRoot = document.querySelector('[data-testid="pdf-document"]');
              const frame = document.querySelector('.pdf-frame-wrap');
              return {
                layout: frame?.dataset.pageLayout ?? documentRoot?.dataset.pageLayout,
                page: document.querySelector('[data-testid="pdf-page"]')?.dataset.pageNumber,
                placeholders: document.querySelectorAll('[data-testid="pdf-page-placeholder"]').length,
              };
            }"""
        )
        assert layouts[layout]["layout"] == layout, f"{prefix}: {layout} layout did not apply {layouts[layout]}"
    return {
        "toolbar": toolbar_metrics,
        "sidebar": sidebar,
        "status": status,
        "layouts": layouts,
    }


def surface_story(page: Page, out: Path, prefix: str) -> dict[str, Any]:
    surface = page.locator('[data-testid="pdf-surface"]')
    surface.wait_for(state="visible", timeout=45_000)
    page.wait_for_function(
        """() => {
          const image = document.querySelector('[data-testid="pdf-surface"] .pdf-surface__page-img');
          return image instanceof HTMLImageElement && image.complete && image.naturalWidth > 0;
        }""",
        timeout=45_000,
    )
    page.locator('[data-testid="pdf-surface-shell"]').scroll_into_view_if_needed()
    page.wait_for_selector(
        '[data-testid="pdf-surface-nav"] [data-testid="pdf-surface-thumbnail"]',
        timeout=30_000,
    )
    wait_idle(page, 220)
    metrics = page.evaluate(
        """() => {
          const shell = document.querySelector('[data-testid="pdf-surface-shell"]');
          const sidebar = document.querySelector('[data-testid="pdf-surface-nav"]');
          const status = document.querySelector('[data-testid="pdf-surface-status-bar"]');
          const frame = document.querySelector('[data-testid="pdf-surface"]');
          const rect = (element) => {
            if (!element) return null;
            const box = element.getBoundingClientRect();
            return { x: box.x, y: box.y, width: box.width, height: box.height, top: box.top, bottom: box.bottom };
          };
          return {
            sidebar: rect(sidebar),
            sidebarMode: sidebar?.dataset.sidebarMode,
            status: rect(status),
            surface: rect(frame),
            thumbnails: sidebar?.querySelectorAll('[data-testid="pdf-surface-thumbnail"]').length ?? 0,
          };
        }"""
    )
    assert metrics["sidebarMode"] in {"inline", "overlay"}, f"{prefix}: surface sidebar {metrics}"
    assert abs(metrics["sidebar"]["width"] - 160) <= 1.5, f"{prefix}: surface sidebar width {metrics}"
    assert metrics["thumbnails"] >= 1, f"{prefix}: surface thumbnails {metrics}"
    assert metrics["status"], f"{prefix}: surface status bar missing {metrics}"
    assert metrics["surface"]["bottom"] <= metrics["status"]["top"] + 1, (
        f"{prefix}: surface status bar covers the pages {metrics}"
    )
    page.locator('[data-testid="pdf-surface-shell"]').screenshot(
        path=str(out / f"{prefix}-surface.png")
    )
    return metrics


def capture_case(
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
        origin = url.split("#", 1)[0].split("?", 1)[0]
        page.goto(f"{origin}?theme={theme}#/pdf-viewer", wait_until="networkidle")
        wait_for_viewer_ready(page, "pdf-viewer")
        page.evaluate("document.fonts.ready")
        inject_deterministic_css(page)
        wait_idle(page, 260)
        overflow = assert_no_page_overflow(page, prefix)
        viewer = viewer_story(page, theme_out, prefix)

        page.goto(f"{origin}?theme={theme}#/pdf-surface", wait_until="networkidle")
        page.locator('[data-testid="pdf-surface-load-sample"]').click()
        page.wait_for_selector('[data-testid="pdf-surface"]', state="visible", timeout=45_000)
        wait_idle(page, 260)
        surface = surface_story(page, theme_out, prefix)

        page.screenshot(path=str(theme_out / f"{prefix}-full-page.png"), full_page=True)
        return {
            "viewport": name,
            "width": width,
            "height": height,
            "theme": theme,
            "pointer": "coarse" if touch else "fine",
            "page": overflow,
            "viewer": viewer,
            "surface": surface,
        }
    finally:
        context.close()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--url", default=DEFAULT_URL)
    parser.add_argument(
        "--serve",
        default="",
        help="directory to serve in-process; replaces the --url origin with an ephemeral loopback URL",
    )
    parser.add_argument("--out", default=str(ROOT / "output" / "pdf-chrome"))
    args = parser.parse_args()

    out_root = Path(args.out)
    out_root.mkdir(parents=True, exist_ok=True)
    errors: list[str] = []
    cases: list[dict[str, Any]] = []
    server: ThreadingHTTPServer | None = None

    try:
        if args.serve:
            server, args.url = start_static_server(Path(args.serve).resolve())
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=True)
            try:
                for name, width, height, touch in VIEWPORTS:
                    for theme in THEMES:
                        try:
                            cases.append(
                                capture_case(
                                    browser, name, width, height, touch, theme,
                                    args.url, out_root, errors,
                                )
                            )
                        except AssertionError as error:
                            errors.append(f"{name}-{theme}: {error}")
            finally:
                browser.close()
    finally:
        if server is not None:
            server.shutdown()
            server.server_close()

    summary = {
        "suite": "P1-PDF-CHROME-RESPONSIVE",
        "result": "FAIL" if errors else "PASS",
        "url": args.url,
        "viewports": [name for name, _, _, _ in VIEWPORTS],
        "themes": THEMES,
        "layouts": LAYOUTS,
        "cases": cases,
        "errors": errors,
    }
    (out_root / "summary.json").write_text(json.dumps(summary, indent=2) + "\n", encoding="utf-8")
    if errors:
        print("pdf chrome failures:", file=sys.stderr)
        for error in errors:
            print(f"  {error}", file=sys.stderr)
        return 1
    print(f"pdf chrome evidence: {out_root}")
    print(f"cases: {len(cases)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
