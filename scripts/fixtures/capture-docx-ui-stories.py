#!/usr/bin/env python3
"""Deterministic DOCX viewer/editor chrome stories (issue #159).

The product visual matrix (#74) captures one state per route. This script
captures the DOCX states the UI migration has to hold together:

  viewer: continuous / single / spread, light / dark, 1440 / 768 / 390
  editor: editable / read-only, light / dark, 1440 / 390

It starts the formal Vite preview, disables animation and font hinting, waits
for the document and chrome, asserts the toolbar never hides controls
horizontally, records whether the thumbnail sidebar resolved to inline or
overlay, and writes lossless WebP screenshots plus a manifest with SHA-256
hashes. The output is reproducible: every scenario is a fresh browser context,
`?theme=` selects the scheme, and no asynchronous stabiliser is left running.

Usage:
    python3 scripts/fixtures/capture-docx-ui-stories.py
    python3 scripts/fixtures/capture-docx-ui-stories.py --skip-build
    python3 scripts/fixtures/capture-docx-ui-stories.py --output /tmp/docx-stories
"""

from __future__ import annotations

import argparse
import hashlib
import json
import socket
import subprocess
import sys
import time
import urllib.request
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
BASELINE = ROOT / "tests/baseline/visual/docx-stories"
SAMPLE = "page-sizes.docx"

sys.path.insert(0, str(ROOT / "tests/blackbox"))
from browser_launch import launch_browser  # noqa: E402

from playwright.sync_api import sync_playwright  # noqa: E402
from PIL import Image  # noqa: E402

VIEWPORTS = {
    "1440x900": (1440, 900),
    "768x1024": (768, 1024),
    "390x844": (390, 844),
}

DETERMINISTIC_CSS = """
* {
  animation: none !important;
  transition: none !important;
  caret-color: transparent !important;
}
html { scrollbar-width: none; }
::-webkit-scrollbar { width: 0; height: 0; }
body {
  font-family: "Liberation Sans", "DejaVu Sans", "Noto Sans CJK SC", Arial, sans-serif !important;
}
"""


@dataclass(frozen=True)
class Story:
    id: str
    route: str
    viewport: str
    theme: str
    view_mode: str = "continuous"
    thumbnails: bool = False
    read_only: bool = False


STORIES: tuple[Story, ...] = (
    Story("viewer-continuous-light-1440", "/docx-viewer", "1440x900", "light"),
    Story("viewer-single-dark-1440", "/docx-viewer", "1440x900", "dark", view_mode="single"),
    Story("viewer-spread-light-1440", "/docx-viewer", "1440x900", "light", view_mode="spread"),
    Story("viewer-continuous-light-390-overlay", "/docx-viewer", "390x844", "light", thumbnails=True),
    Story("viewer-continuous-dark-768-overlay", "/docx-viewer", "768x1024", "dark", thumbnails=True),
    Story("editor-light-1440", "/docx-editor", "1440x900", "light", thumbnails=True),
    Story("editor-readonly-dark-1440", "/docx-editor", "1440x900", "dark", read_only=True),
    Story("editor-light-390-overlay", "/docx-editor", "390x844", "light", thumbnails=True),
)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def free_port(preferred: int = 4191) -> int:
    with socket.socket() as sock:
        if sock.connect_ex(("127.0.0.1", preferred)) != 0:
            return preferred
    with socket.socket() as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def wait_for_preview(base_url: str, process: subprocess.Popen[bytes], log_path: Path) -> None:
    deadline = time.monotonic() + 60
    while time.monotonic() < deadline:
        if process.poll() is not None:
            raise RuntimeError(
                f"preview exited early with {process.returncode}; see {log_path}"
            )
        try:
            with urllib.request.urlopen(base_url, timeout=1) as response:
                if response.status == 200:
                    return
        except Exception:
            time.sleep(0.2)
    raise TimeoutError(f"preview did not become ready at {base_url}")


def stop_preview(process: subprocess.Popen[bytes]) -> None:
    if process.poll() is not None:
        return
    try:
        import os
        import signal

        os.killpg(os.getpgid(process.pid), signal.SIGTERM)
        process.wait(timeout=10)
    except Exception:
        process.kill()


def wait_ready(page) -> None:
    page.wait_for_function(
        """() => document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'""",
        timeout=60_000,
    )
    page.wait_for_selector(
        '[data-testid="docx-page"], .docx-editor-viewer [data-docx-paragraph-host="true"]',
        state="visible",
        timeout=60_000,
    )


def prepare(page, story: Story) -> None:
    if story.route == "/docx-viewer":
        page.select_option('[data-testid="docx-sample-select"]', SAMPLE)
        page.wait_for_timeout(150)
        wait_ready(page)
        page.select_option('[data-testid="docx-prop-view-mode"]', story.view_mode)
        page.wait_for_function(
            """mode => document.querySelector('[data-docx-view-mode]')?.getAttribute('data-docx-view-mode') === mode""",
            arg=story.view_mode,
            timeout=30_000,
        )
        if story.thumbnails:
            page.set_checked('[data-testid="docx-prop-default-thumbnails-open"]', True)
            wait_ready(page)
    else:
        if story.read_only:
            page.set_checked('[data-testid="editor-readonly"]', True)
            page.wait_for_function(
                """() => document.querySelector('[data-testid="editor-editable-state"]')?.dataset.state === 'readonly'""",
                timeout=30_000,
            )
        if story.thumbnails:
            page.click('[data-testid="editor-toggle-thumbnails"]')
    page.wait_for_timeout(300)
    page.wait_for_function(
        """() => document.fonts?.status === 'loaded'""",
        timeout=30_000,
    )
    page.wait_for_timeout(150)


def toolbar_metrics(page, story: Story) -> dict[str, object]:
    selector = (
        '[data-testid="docx-viewer-toolbar"]'
        if story.route == "/docx-viewer"
        else '[data-testid="docx-editor-toolbar"]'
    )
    metrics = page.evaluate(
        """selector => {
          const toolbar = document.querySelector(selector);
          if (!toolbar) return null;
          const rect = toolbar.getBoundingClientRect();
          const controls = Array.from(toolbar.querySelectorAll('button, input, select, [role="group"]'))
            .map(element => ({
              tag: element.tagName.toLowerCase(),
              testid: element.getAttribute('data-testid') || '',
              cls: String(element.className || '').slice(0, 80),
              height: Math.round(element.getBoundingClientRect().height * 100) / 100,
              style: (() => {
                const computed = getComputedStyle(element);
                return {
                  boxSizing: computed.boxSizing,
                  height: computed.height,
                  minHeight: computed.minHeight,
                  paddingBlock: computed.paddingBlock,
                  borderBlock: computed.borderBlock,
                  heightSm: computed.getPropertyValue('--office-ui-height-sm'),
                  controlSm: computed.getPropertyValue('--office-control-height-sm'),
                  heightXs: computed.getPropertyValue('--office-ui-height-xs'),
                };
              })(),
            }));
          const sidebar = document.querySelector('[data-testid="docx-thumbnail-panel"]');
          return {
            height: Math.round(rect.height * 100) / 100,
            maxControlHeight: controls.length ? Math.max(...controls.map(item => item.height)) : 0,
            tallControls: controls.filter(item => item.height > 32),
            scrollWidth: toolbar.scrollWidth,
            clientWidth: toolbar.clientWidth,
            sidebar: sidebar ? {
              mode: sidebar.getAttribute('data-sidebar-mode'),
              open: sidebar.getAttribute('data-sidebar-open'),
              width: Math.round(sidebar.getBoundingClientRect().width * 100) / 100,
            } : null,
          };
        }""",
        selector,
    )
    if metrics is None:
        raise AssertionError(f"{story.id}: toolbar did not render")
    if metrics["scrollWidth"] > metrics["clientWidth"] + 1:
        raise AssertionError(
            f"{story.id}: toolbar hides controls horizontally "
            f"({metrics['scrollWidth']} > {metrics['clientWidth']})"
        )
    if metrics["height"] < 47:
        raise AssertionError(f"{story.id}: toolbar is shorter than 48px: {metrics}")
    return metrics


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--skip-build", action="store_true", help="reuse the current demo dist")
    parser.add_argument("--output", default=str(BASELINE), help="baseline output directory")
    parser.add_argument("--report", default="", help="optional PNG report directory")
    parser.add_argument("--only", default="", help="comma separated story ids")
    args = parser.parse_args()

    output = Path(args.output).resolve()
    output.mkdir(parents=True, exist_ok=True)
    report = Path(args.report).resolve() if args.report else ROOT / "output/docx-ui-stories"
    report.mkdir(parents=True, exist_ok=True)

    if not args.skip_build:
        subprocess.run(["pnpm", "--filter", "demo", "build"], cwd=ROOT, check=True)

    port = free_port()
    base_url = f"http://127.0.0.1:{port}"
    log_path = output / "preview.log"
    log = log_path.open("wb")
    preview = subprocess.Popen(
        ["pnpm", "--filter", "demo", "preview", "--host", "127.0.0.1", "--port", str(port)],
        cwd=ROOT,
        stdout=log,
        stderr=subprocess.STDOUT,
        start_new_session=True,
    )
    selected = {item.strip() for item in args.only.split(",") if item.strip()}
    stories = [story for story in STORIES if not selected or story.id in selected]
    if not stories:
        raise SystemExit("no stories selected")
    results: list[dict[str, object]] = []
    try:
        wait_for_preview(base_url, preview, log_path)
        with sync_playwright() as playwright:
            browser = launch_browser(playwright.chromium)
            browser_source = "chromium"
            for story in stories:
                width, height = VIEWPORTS[story.viewport]
                context = browser.new_context(
                    viewport={"width": width, "height": height},
                    device_scale_factor=1,
                    locale="en-US",
                    timezone_id="UTC",
                    color_scheme=story.theme,
                    reduced_motion="reduce",
                )
                page = context.new_page()
                page.set_default_timeout(60_000)
                page.add_init_script(
                    "document.documentElement.style.scrollbarWidth = 'none';"
                )
                page.goto(
                    f"{base_url}/?theme={story.theme}#{story.route}",
                    wait_until="networkidle",
                )
                page.add_style_tag(content=DETERMINISTIC_CSS)
                wait_ready(page)
                prepare(page, story)
                page.evaluate(
                    """() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))"""
                )
                metrics = toolbar_metrics(page, story)
                sidebar = metrics.get("sidebar") or {}
                expected_sidebar = "overlay" if width <= 768 else "inline"
                if story.thumbnails and sidebar.get("mode") != expected_sidebar:
                    raise AssertionError(
                        f"{story.id}: thumbnail sidebar must be {expected_sidebar}, got {sidebar}"
                    )
                if story.thumbnails and expected_sidebar == "inline":
                    if abs(float(sidebar.get("width", 0)) - 160) > 1:
                        raise AssertionError(
                            f"{story.id}: inline sidebar must be 160px, got {sidebar.get('width')}"
                        )
                png = report / f"{story.id}.png"
                page.screenshot(path=str(png), full_page=False)
                webp = output / f"{story.id}.webp"
                Image.open(png).save(webp, "WEBP", lossless=True, quality=100, method=6)
                results.append(
                    {
                        "id": story.id,
                        "route": story.route,
                        "viewport": story.viewport,
                        "theme": story.theme,
                        "viewMode": story.view_mode if story.route == "/docx-viewer" else None,
                        "thumbnails": story.thumbnails,
                        "readOnly": story.read_only,
                        "toolbar": metrics,
                        "screenshot": str(webp.relative_to(ROOT)),
                        "sha256": sha256(webp),
                    }
                )
                context.close()
            browser.close()
        manifest = {
            "schemaVersion": 1,
            "suite": "docx-ui-stories",
            "issue": 159,
            "capturedAt": datetime.now(timezone.utc).isoformat(),
            "sample": SAMPLE,
            "browser": browser_source,
            "stories": results,
        }
        (output / "manifest.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        (report / "manifest.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        print(f"captured {len(results)} DOCX stories into {output}")
        for result in results:
            toolbar = result["toolbar"]
            sidebar = toolbar.get("sidebar") or {}
            print(
                f"  {result['id']}: toolbar {toolbar['height']}px, "
                f"max control {toolbar.get('maxControlHeight')}px, "
                f"sidebar {sidebar.get('mode') or 'closed'}, {result['sha256'][:12]}"
            )
        return 0
    finally:
        stop_preview(preview)
        log.close()


if __name__ == "__main__":
    raise SystemExit(main())
