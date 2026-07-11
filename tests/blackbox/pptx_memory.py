#!/usr/bin/env python3

import argparse
import json
from pathlib import Path
from urllib.parse import quote

from playwright.sync_api import sync_playwright


def main() -> None:
    parser = argparse.ArgumentParser(description="验证 PPTX 连续重播后的浏览器堆内存和资源释放。")
    parser.add_argument("--app-url", default="http://127.0.0.1:5197")
    parser.add_argument("--sample-url", required=True)
    args = parser.parse_args()
    workspace = Path(__file__).resolve().parents[2]
    module_url = args.app_url + "/@fs" + quote(str(workspace / "packages/pptx-core/src/browser.ts"), safe="/")

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        page = browser.new_page()
        console_errors = []
        page_errors = []
        page.on("console", lambda message: console_errors.append(message.text) if message.type == "error" else None)
        page.on("pageerror", lambda error: page_errors.append(str(error)))
        page.goto(args.app_url)
        page.wait_for_load_state("networkidle")
        page.evaluate(
            """
            async ({ moduleUrl, sampleUrl }) => {
              const api = await import(moduleUrl)
              const response = await fetch(sampleUrl)
              const root = document.createElement("div")
              root.style.width = "1280px"
              root.style.height = "720px"
              document.body.replaceChildren(root)
              const session = api.createPptxDocumentSession(root, { lazyMedia: true, lazySlides: true })
              await session.open(await response.arrayBuffer(), { initialSlide: 0 })
              const controller = session.createPlaybackController({ autoplay: false, initialSlide: 0 })
              globalThis.__pptxMemoryHarness = { controller, root, session }
            }
            """,
            {"moduleUrl": module_url, "sampleUrl": args.sample_url},
        )
        client = page.context.new_cdp_session(page)
        client.send("Performance.enable")

        def heap_size() -> int:
            metrics = client.send("Performance.getMetrics")["metrics"]
            return int(next(item["value"] for item in metrics if item["name"] == "JSHeapUsedSize"))

        page.request_gc()
        before = heap_size()
        replay_metrics = page.evaluate(
            """
            async () => {
              const { controller } = globalThis.__pptxMemoryHarness
              let maxResetMs = 0
              for (let replay = 0; replay < 10; replay += 1) {
                await controller.next()
                const resetStartedAt = performance.now()
                await controller.reset()
                maxResetMs = Math.max(maxResetMs, performance.now() - resetStartedAt)
              }
              return { maxResetMs }
            }
            """
        )
        page.request_gc()
        after = heap_size()
        remaining = page.evaluate(
            """
            () => {
              const { root, session } = globalThis.__pptxMemoryHarness
              session.dispose()
              return root.childElementCount
            }
            """
        )
        page.request_gc()
        released = heap_size()
        browser.close()

    growth = max(0, after - before)
    assert not console_errors, console_errors
    assert not page_errors, page_errors
    assert growth <= 20 * 1024 * 1024, {"before": before, "after": after, "growth": growth}
    assert replay_metrics["maxResetMs"] <= 200, replay_metrics
    assert remaining == 0, remaining
    print(json.dumps({
        "heapBefore": before,
        "heapAfterTenReplays": after,
        "growthBytes": growth,
        "heapAfterDispose": released,
        "remainingChildren": remaining,
        "maxResetMs": replay_metrics["maxResetMs"],
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
