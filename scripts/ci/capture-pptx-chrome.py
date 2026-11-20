"""Issue #161 — deterministic PPTX viewer chrome story screenshots.

The story uses the real demo viewer (`#/pptx-viewer`) with a fixed fixture
uploaded through its file input, so the captures exercise the actual document
session instead of a mock toolbar:

* browse / single (toolbar, page control, zoom, notes toggle, fullscreen)
* browse / sorter (shared cards, badges, skeleton, scrollable mobile grid)
* present + presenter rail (toolbar hidden, notes/next-slide/timer aligned)

Each story is captured at 390 / 768 / 1440 in light and dark. The script serves
`apps/demo/dist` from an in-process thread and always shuts it down in
`finally`, so no Vite/preview child process can leak.

    pnpm --filter demo build
    PYTHON=/tmp/bb-venv/bin/python RUN_PYTHON_MODULES=playwright \\
      node scripts/ci/run-python.mjs scripts/ci/capture-pptx-chrome.py --serve apps/demo/dist

Artifacts default to `output/pptx-chrome/`.
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

from playwright.sync_api import Page, sync_playwright

ROOT = Path(__file__).resolve().parents[2]
FIXTURE = ROOT / "tests/fixtures/pptx/playback/hidden-and-internal-link.pptx"
VIEWPORTS = [
    ("390x844", 390, 844, True),
    ("768x1024", 768, 1024, True),
    ("1440x900", 1440, 900, False),
]
THEMES = ["light", "dark"]
DEFAULT_URL = "http://127.0.0.1:4173/#/pptx-viewer"


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
    return server, f"http://127.0.0.1:{port}/#/pptx-viewer"


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


def viewer_metrics(page: Page, label: str) -> dict[str, Any]:
    """Toolbar geometry and the no-overlap contract for the current story."""
    metrics = page.evaluate(
        """() => {
          const viewer = document.querySelector('[data-testid="pptx-viewer"]');
          const toolbar = viewer.querySelector('[data-testid="pptx-viewer-toolbar"]');
          const stageWrap = viewer.querySelector('.pptx-viewer__stage-wrap');
          const stage = viewer.querySelector('[data-testid="pptx-stage"]');
          const presenter = viewer.querySelector('[data-testid="pptx-presenter"]');
          const rect = (element) => {
            if (!element) return null;
            const box = element.getBoundingClientRect();
            return {
              bottom: box.bottom, height: box.height, left: box.left,
              right: box.right, top: box.top, width: box.width,
            };
          };
          const leading = toolbar?.querySelector('.pptx-toolbar__leading');
          const trailing = toolbar?.querySelector('.pptx-toolbar__trailing');
          const toolbarBox = toolbar?.getBoundingClientRect();
          const outside = toolbar
            ? [...toolbar.querySelectorAll('button, input, [role="combobox"]')]
                .map((element) => ({ element, box: element.getBoundingClientRect() }))
                .filter(({ box }) => box.width > 0 && (
                  box.left < toolbarBox.left - 1 ||
                  box.right > toolbarBox.right + 1 ||
                  box.top < toolbarBox.top - 1 ||
                  box.bottom > toolbarBox.bottom + 1
                ))
                .map(({ element }) => element.outerHTML.slice(0, 120))
            : [];
          const unnamed = toolbar
            ? [...toolbar.querySelectorAll('button')]
                .filter((element) => {
                  const text = (element.innerText || '').trim();
                  return !text && !element.getAttribute('aria-label') && !element.getAttribute('title');
                })
                .map((element) => element.outerHTML.slice(0, 120))
            : [];
          return {
            theme: viewer.getAttribute('data-office-theme'),
            presenter: rect(presenter),
            stage: rect(stage),
            stageWrap: rect(stageWrap),
            toolbar: rect(toolbar),
            toolbarOverflow: toolbar ? toolbar.scrollWidth - toolbar.clientWidth : 0,
            wrapped: Boolean(leading && trailing && trailing.getBoundingClientRect().top > leading.getBoundingClientRect().top + 1),
            outside,
            unnamed,
            zoomLabel: (
              toolbar?.querySelector('[data-testid="pptx-zoom-reset"] .pptx-toolbar__zoom-value')?.innerText?.trim()
              ?? ''
            ),
            pageCounter: toolbar?.querySelector('[data-testid="pptx-page-counter"]')?.innerText?.trim() ?? '',
            pageInput: toolbar?.querySelector('[data-testid="office-page-input"]')?.innerText?.trim() ?? '',
            pageTotal: toolbar?.querySelector('[data-testid="office-page-total"]')?.innerText?.trim() ?? '',
          };
        }"""
    )
    assert not metrics["outside"], f"{label}: toolbar controls outside their row {metrics['outside']}"
    assert not metrics["unnamed"], f"{label}: unnamed toolbar buttons {metrics['unnamed']}"
    return metrics


def browse_story(page: Page, out: Path, prefix: str, theme: str) -> dict[str, Any]:
    page.wait_for_selector('[data-testid="pptx-viewer"][data-state="ready"]')
    wait_idle(page)
    metrics = viewer_metrics(page, f"{prefix}-browse")
    metrics["page"] = assert_no_page_overflow(page, f"{prefix}-browse")

    assert metrics["theme"] == theme, f"{prefix}: viewer theme {metrics['theme']!r} != {theme!r}"
    assert metrics["toolbar"], f"{prefix}: browse toolbar missing"
    assert metrics["toolbar"]["height"] >= 48, (
        f"{prefix}: toolbar height {metrics['toolbar']['height']} < 48"
    )
    assert metrics["toolbarOverflow"] <= 1, f"{prefix}: toolbar overflows {metrics}"
    assert metrics["toolbar"]["bottom"] <= metrics["stageWrap"]["top"] + 1, (
        f"{prefix}: toolbar overlaps the slide area {metrics}"
    )
    assert metrics["pageInput"], f"{prefix}: page control input missing"
    assert "%" in metrics["zoomLabel"], f"{prefix}: zoom value {metrics['zoomLabel']!r}"

    viewer = page.locator('[data-testid="pptx-viewer"]')
    viewer.screenshot(path=str(out / f"{prefix}-browse-single.png"))
    page.locator('[data-testid="pptx-viewer-toolbar"]').screenshot(
        path=str(out / f"{prefix}-browse-toolbar.png")
    )
    return metrics


def open_toolbar_actions(page: Page) -> None:
    """Layout/present/notes/find live in the upstream-style actions menu."""
    menu = page.locator('[data-testid="pptx-actions-menu"]')
    if menu.count() == 0:
        page.locator('[data-testid="pptx-actions"]').click()
        menu.wait_for(state="attached")


def toolbar_menu_item(page: Page, selector: str):
    open_toolbar_actions(page)
    return page.locator(selector)


def sorter_story(page: Page, out: Path, prefix: str) -> dict[str, Any]:
    toolbar_menu_item(page, '[data-testid="pptx-view-mode-sorter"]').click()
    page.wait_for_selector('[data-testid="pptx-sorter"]')
    wait_idle(page)
    metrics = viewer_metrics(page, f"{prefix}-sorter")
    sorter = page.evaluate(
        """() => {
          const sorter = document.querySelector('[data-testid="pptx-sorter"]');
          const box = sorter.getBoundingClientRect();
          const cards = [...sorter.querySelectorAll('[data-testid="pptx-sorter-item"]')];
          const overflowing = cards
            .map((card) => ({ card, box: card.getBoundingClientRect() }))
            .filter(({ box: card }) => card.right > box.right + 1 || card.left < box.left - 1)
            .map(({ card }) => card.getAttribute('data-slide-index'));
          const grid = sorter.querySelector('[data-testid="pptx-sorter-grid"]');
          const columns = getComputedStyle(grid).gridTemplateColumns.split(' ').filter(Boolean).length;
          return {
            cards: cards.length,
            columns,
            overflowing,
            scrollHeight: sorter.scrollHeight,
            clientHeight: sorter.clientHeight,
            scrollWidth: sorter.scrollWidth,
            clientWidth: sorter.clientWidth,
            cardStyles: cards.slice(0, 2).map((card) => ({
              border: getComputedStyle(card).borderTopWidth,
              radius: getComputedStyle(card).borderTopLeftRadius,
              shadow: getComputedStyle(card).boxShadow,
            })),
          };
        }"""
    )
    assert metrics["toolbar"], f"{prefix}: sorter toolbar missing"
    assert metrics["toolbar"]["height"] >= 48, f"{prefix}: sorter toolbar height {metrics['toolbar']['height']}"
    assert metrics["toolbar"]["bottom"] <= metrics["stageWrap"]["top"] + 1, (
        f"{prefix}: sorter toolbar overlaps the content {metrics}"
    )
    assert sorter["cards"] >= 4, f"{prefix}: only {sorter['cards']} sorter cards mounted"
    assert sorter["columns"] >= 2, f"{prefix}: sorter columns {sorter['columns']}"
    assert not sorter["overflowing"], f"{prefix}: sorter cards overflow horizontally {sorter['overflowing']}"
    assert sorter["scrollWidth"] <= sorter["clientWidth"] + 1, f"{prefix}: sorter overflows {sorter}"
    assert sorter["cardStyles"], f"{prefix}: no sorter cards"
    first_card = sorter["cardStyles"][0]
    assert float(first_card["radius"].replace("px", "")) >= 4, (
        f"{prefix}: sorter cards lost the upstream rounded card radius {first_card}"
    )
    # The card frame itself carries the upstream `shadow-xs` (active: `shadow-sm`).
    frame_style = page.evaluate(
        """() => {
          const card = document.querySelector('[data-testid="pptx-sorter-item"]');
          const frame = card?.querySelector('.pptx-slide-card__frame');
          return frame ? getComputedStyle(frame).boxShadow : '';
        }"""
    )
    assert frame_style and frame_style != "none", (
        f"{prefix}: sorter card frame lost the upstream shadow {frame_style!r}"
    )

    # The mobile grid must scroll through the deck without changing the slide.
    slide_before = page.locator('[data-testid="pptx-viewer"]').get_attribute("data-slide-index")
    page.evaluate(
        """() => {
          const sorter = document.querySelector('[data-testid="pptx-sorter"]');
          sorter.scrollTop = sorter.scrollHeight;
          sorter.dispatchEvent(new Event('scroll'));
        }"""
    )
    wait_idle(page, 120)
    scrolled = page.evaluate(
        """() => {
          const sorter = document.querySelector('[data-testid="pptx-sorter"]');
          return { top: sorter.scrollTop, max: sorter.scrollHeight - sorter.clientHeight };
        }"""
    )
    if scrolled["max"] > 0:
        assert scrolled["top"] > 0, f"{prefix}: sorter did not scroll {scrolled}"
    assert page.locator('[data-testid="pptx-viewer"]').get_attribute("data-slide-index") == slide_before, (
        f"{prefix}: scrolling the sorter changed the active slide"
    )

    page.locator('[data-testid="pptx-viewer"]').screenshot(
        path=str(out / f"{prefix}-browse-sorter.png")
    )
    page.locator('[data-testid="pptx-sorter"]').screenshot(
        path=str(out / f"{prefix}-sorter-grid.png")
    )
    return {**metrics, "sorter": sorter, "scrolled": scrolled}


def present_story(page: Page, out: Path, prefix: str) -> dict[str, Any]:
    toolbar_menu_item(page, '[data-testid="pptx-view-mode-single"]').click()
    toolbar_menu_item(page, '[data-testid="pptx-notes-toggle"]').click()
    page.wait_for_selector('[data-testid="pptx-presenter"]')
    page.locator('[data-testid="pptx-mode-present"]').click()
    page.wait_for_selector('[data-testid="pptx-viewer"].pptx-viewer--present')
    page.wait_for_selector('[data-testid="pptx-viewer"][data-state="ready"]')
    wait_idle(page, 320)

    metrics = viewer_metrics(page, f"{prefix}-present")
    presenter = page.evaluate(
        """() => {
          const panel = document.querySelector('[data-testid="pptx-presenter"]');
          const timer = panel.querySelector('[data-testid="pptx-presenter-timer"]');
          const notes = panel.querySelector('[data-testid="pptx-presenter-notes"]');
          const next = panel.querySelector('.pptx-presenter__next, [data-testid="pptx-presenter-last"]');
          const box = panel.getBoundingClientRect();
          return {
            mode: panel.getAttribute('data-presenter-mode'),
            label: panel.getAttribute('aria-label'),
            timer: timer?.innerText?.trim() ?? '',
            notesVisible: Boolean(notes && notes.getBoundingClientRect().height > 0),
            nextVisible: Boolean(next && next.getBoundingClientRect().height > 0),
            width: box.width,
            height: box.height,
            card: Boolean(panel.querySelector('.office-card')),
            scrollArea: Boolean(panel.querySelector('.office-scroll-area')),
            separator: Boolean(panel.querySelector('.office-separator')),
          };
        }"""
    )
    assert metrics["toolbar"] is None, f"{prefix}: present mode must hide the browse toolbar"
    assert metrics["stage"], f"{prefix}: present stage missing"
    assert metrics["presenter"], f"{prefix}: presenter panel missing"
    assert presenter["mode"] == "present", f"{prefix}: presenter mode {presenter['mode']!r}"
    assert presenter["card"], f"{prefix}: presenter panel is not a shared card"
    assert presenter["scrollArea"], f"{prefix}: presenter notes are not in a scroll area"
    assert presenter["separator"], f"{prefix}: presenter sections have no separator"
    assert presenter["nextVisible"], f"{prefix}: next-slide preview missing"
    assert presenter["notesVisible"], f"{prefix}: notes body missing"
    assert presenter["timer"], f"{prefix}: presenter timer missing"

    page.locator('[data-testid="pptx-viewer"]').screenshot(
        path=str(out / f"{prefix}-present.png")
    )
    page.locator('[data-testid="pptx-presenter"]').screenshot(
        path=str(out / f"{prefix}-presenter.png")
    )
    return {**metrics, "presenter": presenter}


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
        page.goto(f"{origin}?theme={theme}#/pptx-viewer", wait_until="networkidle")
        page.set_input_files('[data-testid="pptx-file-input"]', str(FIXTURE))
        page.wait_for_selector('[data-testid="pptx-viewer"][data-state="ready"]')
        page.evaluate("document.fonts.ready")
        inject_deterministic_css(page)
        wait_idle(page, 260)

        browse = browse_story(page, theme_out, prefix, theme)
        sorter = sorter_story(page, theme_out, prefix)
        present = present_story(page, theme_out, prefix)
        page.screenshot(path=str(theme_out / f"{prefix}-full-page.png"), full_page=True)
        return {
            "viewport": name,
            "width": width,
            "height": height,
            "theme": theme,
            "browse": browse,
            "sorter": sorter,
            "present": present,
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
    parser.add_argument("--out", default=str(ROOT / "output" / "pptx-chrome"))
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
                        cases.append(
                            capture_case(
                                browser, name, width, height, touch, theme,
                                args.url, out_root, errors,
                            )
                        )
            finally:
                browser.close()
    finally:
        if server is not None:
            server.shutdown()
            server.server_close()

    summary = {
        "suite": "P1-PPTX-CHROME-RESPONSIVE",
        "result": "FAIL" if errors else "PASS",
        "url": args.url,
        "fixture": str(FIXTURE.relative_to(ROOT)),
        "viewports": [name for name, _, _, _ in VIEWPORTS],
        "themes": THEMES,
        "cases": cases,
        "errors": errors,
    }
    (out_root / "summary.json").write_text(json.dumps(summary, indent=2) + "\n", encoding="utf-8")
    if errors:
        print("console errors:", file=sys.stderr)
        for error in errors:
            print(f"  {error}", file=sys.stderr)
        return 1
    print(f"pptx chrome evidence: {out_root}")
    print(f"cases: {len(cases)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
