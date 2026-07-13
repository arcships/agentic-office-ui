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
        browse_layout = page.evaluate(
            """() => {
              const stage = document.querySelector('[data-testid="pptx-stage"]')
              const scroller = document.querySelector('.pptx-viewer__stage-wrap')
              const slides = [...stage.querySelectorAll(':scope > [data-slide-index]')]
              const tops = slides.map((slide) => slide.getBoundingClientRect().top)
              return {
                count: slides.length,
                vertical: tops.every((top, index) => index === 0 || top > tops[index - 1]),
                scrollable: scroller.scrollHeight > scroller.clientHeight,
              }
            }"""
        )
        assert browse_layout == {"count": 12, "vertical": True, "scrollable": True}
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

        stage = page.locator('[data-testid="pptx-stage"]')
        stage_point = page.evaluate(
            """() => {
              const stage = document.querySelector('[data-testid="pptx-stage"]');
              if (!stage) throw new Error('PPTX stage is missing');
              const rect = stage.getBoundingClientRect();
              for (const yRatio of [0.9, 0.75, 0.5, 0.25, 0.1]) {
                for (const xRatio of [0.9, 0.75, 0.5, 0.25, 0.1]) {
                  const x = rect.left + rect.width * xRatio;
                  const y = rect.top + rect.height * yRatio;
                  const target = document.elementFromPoint(x, y);
                  if (target && stage.contains(target) && !target.closest('[data-pptx-object-key]')) {
                    return { x, y };
                  }
                }
              }
              throw new Error('PPTX stage has no clickable background point');
            }"""
        )
        page.mouse.click(stage_point["x"], stage_point["y"])
        page.wait_for_function(
            """() => {
              const viewer = document.querySelector('[data-testid="pptx-viewer"]')
              return Number(viewer?.dataset.clickBoundary ?? 0) >= 1
            }""",
            timeout=5_000,
        )
        first_boundary = int(viewer.get_attribute("data-click-boundary") or "0")
        assert int(viewer.get_attribute("data-slide-index") or "0") == 0
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

        page.goto(f"{args.app_url}/#/pptx-surface", wait_until="networkidle")
        page.locator('[data-testid="pptx-surface-file-input"]').set_input_files(str(args.fixture.resolve()))
        page.wait_for_function(
            "() => document.querySelectorAll('[data-testid=\"pptx-surface-stage\"] > [data-slide-index]').length === 12",
            timeout=30_000,
        )
        page.locator('[data-testid="pptx-surface-stage"] > [data-slide-index="1"]').dispatch_event("click")
        assert "已选中第 2 页" in page.locator('[data-testid="pptx-surface-status"]').inner_text()
        surface_object = page.locator('[data-testid="pptx-surface-stage"] [data-pptx-object-key]').first
        surface_object.click(force=True)
        assert "页对象：" in page.locator('[data-testid="pptx-surface-status"]').inner_text()
        surface_object.click(button="right", force=True)
        surface_context = page.locator('[data-testid="pptx-surface-status"]').inner_text()
        assert "对象" in surface_context and "右键" in surface_context
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
