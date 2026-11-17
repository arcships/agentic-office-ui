#!/usr/bin/env python3

import argparse
import json
from pathlib import Path
from urllib.parse import quote

from playwright.sync_api import sync_playwright

from browser_launch import launch_browser


def main() -> None:
    parser = argparse.ArgumentParser(
        description="验证 PPTX 文档会话和播放控制器的状态机。"
    )
    parser.add_argument("--app-url", default="http://127.0.0.1:5197")
    parser.add_argument("--sample-url", required=True)
    parser.add_argument(
        "--browser", choices=("chromium", "firefox", "webkit"), default="chromium"
    )
    parser.add_argument("--initial-slide", type=int, default=0)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--quiet", action="store_true")
    args = parser.parse_args()

    workspace = Path(__file__).resolve().parents[2]
    module_url = args.app_url + "/@fs" + quote(
        str(workspace / "packages/pptx-core/src/browser.ts"), safe="/"
    )

    console_errors: list[str] = []
    page_errors: list[str] = []
    with sync_playwright() as playwright:
        browser = launch_browser(getattr(playwright, args.browser), headless=True)
        page = browser.new_page()
        page.on(
            "console",
            lambda message: console_errors.append(message.text)
            if message.type == "error"
            and not (message.location or {}).get("url", "").endswith("/favicon.ico")
            else None,
        )
        page.on("pageerror", lambda error: page_errors.append(str(error)))
        page.goto(args.app_url)
        page.wait_for_load_state("networkidle")
        result = page.evaluate(
            """
            async ({ moduleUrl, sampleUrl, initialSlide }) => {
              const api = await import(moduleUrl)
              const response = await fetch(sampleUrl)
              if (!response.ok) throw new Error(`sample fetch failed: ${response.status}`)
              const bytes = await response.arrayBuffer()
              const root = document.createElement("div")
              root.style.width = "1280px"
              root.style.height = "720px"
              document.body.replaceChildren(root)
              const session = api.createPptxDocumentSession(root, {
                lazyMedia: true,
                lazySlides: true,
                approximation: "safe",
              })
              const checks = {}
              const details = {}
              try {
                const first = await session.open(bytes, { initialSlide })
                const firstPlayback = session.playbackDocument
                if (!firstPlayback) throw new Error("playback document missing")

                let failedOpenCode = null
                try {
                  await session.open(new Uint8Array([1, 2, 3]))
                } catch (error) {
                  failedOpenCode = error?.code ?? String(error)
                }
                const retainedController = session.createPlaybackController({
                  autoplay: false,
                  initialSlide,
                })
                checks.failedOpenRetainsDocument =
                  failedOpenCode === "INVALID_SOURCE"
                  && session.document === first
                  && session.playbackDocument === firstPlayback
                  && retainedController.snapshot.slideIndex === initialSlide
                details.failedOpenCode = failedOpenCode
                retainedController.dispose()

                const matches = session.searchText("切换")
                const staleMatch = matches[0]
                await session.open(bytes, { initialSlide })
                const staleHighlight = await session.highlightSearchResult(staleMatch)
                checks.searchCacheCleared =
                  matches.length > 0 && staleHighlight === null
                details.searchMatches = matches.length

                const controller = session.createPlaybackController({
                  autoplay: false,
                  initialSlide,
                })
                const events = []
                controller.subscribe((event) => events.push(event))
                const navigationResults = await Promise.allSettled([
                  controller.goToSlide(initialSlide + 1),
                  controller.goToSlide(initialSlide + 2),
                  controller.goToSlide(initialSlide + 3),
                ])
                const rejectionCodes = navigationResults
                  .filter((entry) => entry.status === "rejected")
                  .map((entry) => entry.reason?.code ?? String(entry.reason))
                const changes = events
                  .filter((event) => event.type === "slidechange")
                  .map((event) => event.to)
                checks.supersededNavigationRejects =
                  rejectionCodes.length === 1
                  && rejectionCodes[0] === "COMMAND_CANCELLED"
                  && controller.snapshot.slideIndex === initialSlide + 3
                  && JSON.stringify(changes) === JSON.stringify([
                    initialSlide + 1,
                    initialSlide + 3,
                  ])
                details.rejectionCodes = rejectionCodes
                details.navigationChanges = changes

                const beforeStatus = controller.snapshot.status
                const beforeIndex = controller.snapshot.slideIndex
                const originalRenderSlide = session.renderSlide
                session.renderSlide = async () => {
                  throw new Error("forced render failure")
                }
                let renderFailureCode = null
                try {
                  await controller.goToSlide(initialSlide + 4)
                } catch (error) {
                  renderFailureCode = error?.code ?? String(error)
                }
                session.renderSlide = originalRenderSlide
                const failureEvents = events.filter(
                  (event) => event.type === "error",
                )
                const renderError = failureEvents.at(-1)?.error
                checks.renderFailureRecovers =
                  renderFailureCode === "RENDER_FAILED"
                  && controller.snapshot.status === beforeStatus
                  && controller.snapshot.slideIndex === beforeIndex
                  && renderError?.code === "RENDER_FAILED"
                details.renderFailureCode = renderFailureCode
                details.renderFailureStatus = controller.snapshot.status
                details.renderFailureSlideIndex = controller.snapshot.slideIndex

                await controller.goToSlide(initialSlide + 4)
                checks.renderRecovery =
                  controller.snapshot.status === "ready"
                  && controller.snapshot.slideIndex === initialSlide + 4
                controller.dispose()

                const syntheticRoot = document.createElement("div")
                document.body.appendChild(syntheticRoot)
                const warningCount = 1000
                const nodes = {
                  root: {
                    acceleration: 0,
                    autoReverse: false,
                    childIds: Array.from(
                      { length: warningCount },
                      (_, index) => `effect-${index}`,
                    ),
                    conditions: [],
                    container: "root",
                    deceleration: 0,
                    delayMs: 0,
                    durationMs: 0,
                    fill: "hold",
                    id: "root",
                    kind: "group",
                    restart: "never",
                  },
                }
                for (let index = 0; index < warningCount; index += 1) {
                  nodes[`effect-${index}`] = {
                    acceleration: 0,
                    autoReverse: false,
                    childIds: [],
                    conditions: [{
                      delayMs: 0,
                      event: index === 0 ? "with-previous" : "after-previous",
                      source: "start",
                    }],
                    container: "behavior",
                    deceleration: 0,
                    delayMs: 0,
                    durationMs: 0,
                    effect: {
                      id: `animation-${index}`,
                      kind: "appear",
                      targetObjectKey: `missing-object-${index}`,
                      values: {},
                    },
                    fill: "hold",
                    id: `effect-${index}`,
                    kind: "effect",
                    restart: "never",
                  }
                }
                const capability = {
                  approximate: 0,
                  discovered: 0,
                  features: [],
                  static: 0,
                  strict: 0,
                  unparsed: 0,
                }
                const retentionController = api.createPptxPlaybackController(
                  {
                    document: {
                      capability,
                      slides: [{
                        actions: [],
                        capability,
                        hidden: false,
                        index: 0,
                        media: [],
                        morphFromPrevious: [],
                        nodes,
                        objects: [],
                        rootNodeId: "root",
                      }],
                    },
                    initialSlideIndex: 0,
                    renderSlide: async () => {},
                    root: syntheticRoot,
                  },
                  { autoplay: false },
                )
                const warnings = []
                retentionController.subscribe((event) => {
                  if (event.type === "warning") warnings.push(event.warning)
                })
                checks.retainedEventsCapped =
                  warnings.length === 50
                  && warnings[0]?.objectKey === `missing-object-${warningCount - 50}`
                  && warnings.at(-1)?.objectKey === `missing-object-${warningCount - 1}`
                details.retainedWarnings = warnings.length
                retentionController.dispose()
                syntheticRoot.remove()

                return {
                  checks,
                  details,
                  slides: first.slides.length,
                }
              } finally {
                session.dispose()
              }
            }
            """,
            {
                "moduleUrl": module_url,
                "sampleUrl": args.sample_url,
                "initialSlide": args.initial_slide,
            },
        )
        browser.close()

    assert not console_errors, f"浏览器控制台错误：{console_errors}"
    assert not page_errors, f"页面错误：{page_errors}"
    checks = result["checks"]
    assert result["slides"] > 0, result
    assert checks["failedOpenRetainsDocument"], result
    assert checks["searchCacheCleared"], result
    assert checks["supersededNavigationRejects"], result
    assert checks["renderFailureRecovers"], result
    assert checks["renderRecovery"], result
    assert checks["retainedEventsCapped"], result
    if not args.quiet:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(
            json.dumps(result, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )


if __name__ == "__main__":
    main()
