#!/usr/bin/env python3

import argparse
import json
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


def main() -> None:
    parser = argparse.ArgumentParser(description="验证网页 PPTX 浏览和演示组件。")
    parser.add_argument("--app-url", required=True)
    parser.add_argument("--fixture", type=Path, required=True)
    parser.add_argument("--browser", choices=("chromium", "firefox", "webkit"), required=True)
    args = parser.parse_args()

    console_errors: list[str] = []
    page_errors: list[str] = []
    with sync_playwright() as playwright:
        browser = getattr(playwright, args.browser).launch(headless=True)
        page = browser.new_page(viewport={"width": 1440, "height": 900})
        page.on("console", lambda message: console_errors.append(message.text) if message.type == "error" else None)
        page.on("pageerror", lambda error: page_errors.append(str(error)))
        page.goto(f"{args.app_url}/#/pptx-viewer", wait_until="networkidle")

        viewer = page.locator('[data-testid="pptx-viewer"]')
        page.locator('[data-testid="pptx-file-input"]').set_input_files(str(args.fixture.resolve()))
        page.wait_for_function(
            "() => document.querySelector('[data-testid=\"pptx-viewer\"]')?.dataset.state === 'ready'",
            timeout=30_000,
        )
        assert page.locator(".pptx-viewer__toolbar").count() == 1
        assert page.locator(".pptx-thumbnail").count() == 12
        assert page.locator(".pptx-thumbnail__hidden").count() == 1
        search = page.get_by_role("search")
        search.get_by_role("searchbox").fill("单击动画")
        search.get_by_role("button", name="搜索").click()
        page.wait_for_function(
            "() => document.querySelector('.pptx-viewer__search-count')?.textContent?.includes('/')",
            timeout=5_000,
        )
        page.get_by_role("button", name="放大").click()
        page.wait_for_function(
            "() => document.querySelector('[aria-label=\"恢复适合窗口\"]')?.textContent?.includes('125%')",
            timeout=5_000,
        )

        page.locator('[data-testid="pptx-mode-present"]').click()
        page.wait_for_function(
            """() => {
              const viewer = document.querySelector('[data-testid="pptx-viewer"]')
              return viewer?.dataset.state === 'ready' && viewer.classList.contains('pptx-viewer--present')
            }""",
            timeout=30_000,
        )
        assert page.locator(".pptx-viewer__playback-controls").count() == 1
        assert page.locator("[data-pptx-object-key]").count() > 0
        assert "精确" in page.locator(".pptx-viewer__capability").inner_text()

        page.get_by_role("button", name="下一步").click()
        page.wait_for_function(
            """() => {
              const viewer = document.querySelector('[data-testid="pptx-viewer"]')
              return Number(viewer?.dataset.clickBoundary ?? 0) >= 1
            }""",
            timeout=5_000,
        )
        first_boundary = int(viewer.get_attribute("data-click-boundary") or "0")
        viewer.focus()
        page.keyboard.press("ArrowRight")
        page.wait_for_function(
            """previous => {
              const viewer = document.querySelector('[data-testid="pptx-viewer"]')
              return Number(viewer?.dataset.clickBoundary ?? 0) > previous
                || Number(viewer?.dataset.slideIndex ?? 0) > 0
            }""",
            arg=first_boundary,
            timeout=5_000,
        )

        page.get_by_role("button", name="全屏").click()
        fullscreen_entered = False
        try:
            page.wait_for_function("() => document.fullscreenElement !== null", timeout=2_000)
            fullscreen_entered = True
        except PlaywrightTimeoutError:
            pass
        fullscreen_supported = page.evaluate("() => Boolean(document.documentElement.requestFullscreen)")
        fullscreen_status = page.locator('[data-testid="pptx-demo-status"]').inner_text()
        assert fullscreen_entered or not fullscreen_supported or "演示错误" in fullscreen_status
        if fullscreen_entered:
            page.keyboard.press("Escape")
            page.wait_for_function("() => document.fullscreenElement === null", timeout=3_000)

        page.locator('[data-testid="pptx-mode-browse"]').click()
        page.wait_for_function(
            """() => {
              const viewer = document.querySelector('[data-testid="pptx-viewer"]')
              return viewer?.dataset.state === 'ready' && !viewer.classList.contains('pptx-viewer--present')
            }""",
            timeout=30_000,
        )
        assert page.locator(".pptx-viewer__playback-controls").count() == 0
        assert page.locator(".pptx-viewer__toolbar").count() == 1
        assert not console_errors, console_errors
        assert not page_errors, page_errors

        print(json.dumps({
            "browser": args.browser,
            "firstBoundary": first_boundary,
            "fullscreenSupported": fullscreen_supported,
            "fullscreenEntered": fullscreen_entered,
            "status": "PASS",
        }, ensure_ascii=False))
        browser.close()


if __name__ == "__main__":
    main()
