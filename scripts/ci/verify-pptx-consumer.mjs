#!/usr/bin/env node

import { execFileSync, spawn } from "node:child_process"
import { copyFileSync, mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { createServer } from "node:net"
import { tmpdir } from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..")
const coreVersion = JSON.parse(readFileSync(path.join(root, "packages/pptx-core/package.json"), "utf8")).version
const vueVersion = JSON.parse(readFileSync(path.join(root, "packages/vue-pptx/package.json"), "utf8")).version
const interactionVersion = JSON.parse(readFileSync(path.join(root, "packages/office-interaction/package.json"), "utf8")).version
const temporary = mkdtempSync(path.join(tmpdir(), "pptx-consumer-"))
const packages = path.join(temporary, "packages")
const consumer = path.join(temporary, "consumer")
const fixture = process.env.PPTX_CONSUMER_FIXTURE
  ? path.resolve(process.env.PPTX_CONSUMER_FIXTURE)
  : path.join(root, "tests/fixtures/pptx/playback-controlled.pptx")
mkdirSync(packages)
mkdirSync(path.join(consumer, "src"), { recursive: true })
mkdirSync(path.join(consumer, "public"), { recursive: true })

let previewServer = null

function run(command, args, cwd) {
  return execFileSync(command, args, {
    cwd,
    encoding: "utf8",
    maxBuffer: 32 * 1024 * 1024,
    stdio: ["ignore", "pipe", "inherit"],
  }).trim()
}

function availablePort() {
  return new Promise((resolve, reject) => {
    const server = createServer()
    server.once("error", reject)
    server.listen(0, "127.0.0.1", () => {
      const address = server.address()
      const port = typeof address === "object" && address ? address.port : 0
      server.close((error) => error ? reject(error) : resolve(port))
    })
  })
}

async function waitForServer(url) {
  for (let attempt = 0; attempt < 100; attempt += 1) {
    try {
      const response = await fetch(url)
      if (response.ok) return
    } catch { /* 服务仍在启动。 */ }
    await new Promise((resolve) => setTimeout(resolve, 100))
  }
  throw new Error("PPTX 外部消费预览服务启动超时。")
}

try {
  run("pnpm", ["pack", "--pack-destination", packages], path.join(root, "packages/pptx-core"))
  run("pnpm", ["pack", "--pack-destination", packages], path.join(root, "packages/vue-pptx"))
  run("pnpm", ["pack", "--pack-destination", packages], path.join(root, "packages/office-interaction"))
  const core = path.join(packages, `arcships-pptx-core-${coreVersion}.tgz`)
  const vuePptx = path.join(packages, `arcships-vue-pptx-${vueVersion}.tgz`)
  const interaction = path.join(packages, `arcships-office-interaction-${interactionVersion}.tgz`)
  const coreManifest = JSON.parse(run("tar", ["-xOf", core, "package/package.json"], root))
  const browserBundle = run("tar", ["-xOf", core, "package/dist/browser.js"], root)
  if (coreManifest.dependencies?.["@aiden0z/pptx-renderer"]) {
    throw new Error("PPTX 压缩包仍把补丁渲染器暴露为运行依赖。")
  }
  for (const marker of ["pptxNodeId", "pptxNodeType", "pptxPartPath", "pptxParagraphIndex"]) {
    if (!browserBundle.includes(marker)) {
      throw new Error(`PPTX 浏览器入口没有包含渲染标记：${marker}`)
    }
  }
  if (/from\s*["']@aiden0z\/pptx-renderer["']/u.test(browserBundle)) {
    throw new Error("PPTX 浏览器入口仍在外部加载未打补丁的渲染器。")
  }
  writeFileSync(path.join(consumer, "package.json"), JSON.stringify({
    private: true,
    type: "module",
    scripts: { build: "tsc --noEmit && vite build" },
    dependencies: {
      "@arcships/pptx-core": `file:${core}`,
      "@arcships/vue-pptx": `file:${vuePptx}`,
      vue: "^3.5.0",
    },
    devDependencies: { typescript: "^5.0.0", vite: "^6.0.0" },
    pnpm: {
      overrides: {
        "@arcships/pptx-core": `file:${core}`,
        "@arcships/office-interaction": `file:${interaction}`,
      },
    },
  }, null, 2))
  writeFileSync(path.join(consumer, "tsconfig.json"), JSON.stringify({
    compilerOptions: {
      target: "ES2022",
      module: "ESNext",
      moduleResolution: "Bundler",
      strict: true,
      skipLibCheck: true,
    },
    include: ["src"],
  }, null, 2))
  writeFileSync(path.join(consumer, "index.html"), "<main id=\"app\"></main><script type=\"module\" src=\"/src/main.ts\"></script>")
  writeFileSync(path.join(consumer, "src/main.ts"), `
import { computed, createApp, defineComponent, h, ref } from "vue"
import {
  PptxStage,
  PptxViewer,
  usePptxDocument,
  usePptxPlayback,
  type PptxStageExpose,
  type PptxPreviewSource,
} from "@arcships/vue-pptx"
import "@arcships/vue-pptx/style.css"
import { createPptxDocumentSession, type PptxDocumentSession } from "@arcships/pptx-core/browser"

const factory: typeof createPptxDocumentSession = createPptxDocumentSession
const sessionType: PptxDocumentSession | null = null
void factory
void sessionType
const HeadlessConsumer = defineComponent({
  props: { source: { required: true, type: Object } },
  setup(props) {
    const stage = ref<PptxStageExpose | null>(null)
    const element = computed(() => stage.value?.element ?? null)
    const document = usePptxDocument(element, {
      source: () => props.source as PptxPreviewSource,
    })
    const playback = usePptxPlayback(document)
    void playback
    return () => h(PptxStage, { ref: stage })
  },
})
void HeadlessConsumer
async function main() {
  const source = await fetch("/sample.pptx").then((response) => response.arrayBuffer())
  createApp({
    render: () => h(PptxViewer, {
      source,
      mode: "present",
      autoplay: false,
      height: "800px",
    }),
  }).mount("#app")
}
void main()
`)
  copyFileSync(
    fixture,
    path.join(consumer, "public/sample.pptx"),
  )
  run("pnpm", ["install", "--ignore-workspace"], consumer)
  run("pnpm", ["build"], consumer)
  const port = await availablePort()
  const appUrl = `http://127.0.0.1:${port}`
  previewServer = spawn("pnpm", ["exec", "vite", "preview", "--host", "127.0.0.1", "--port", String(port), "--strictPort"], {
    cwd: consumer,
    stdio: ["ignore", "ignore", "inherit"],
  })
  await waitForServer(appUrl)
  const browserCheck = path.join(consumer, "browser-check.py")
  writeFileSync(browserCheck, `
from playwright.sync_api import sync_playwright

with sync_playwright() as playwright:
    browser = playwright.chromium.launch(headless=True)
    page = browser.new_page(viewport={"width": 1440, "height": 900})
    errors = []
    page.on("pageerror", lambda error: errors.append(str(error)))
    page.goto("${appUrl}", wait_until="networkidle")
    page.wait_for_function("() => document.querySelector('[data-testid=\\"pptx-viewer\\"]')?.dataset.state === 'ready'", timeout=30000)
    assert page.locator("[data-pptx-object-key]").count() > 0
    viewer = page.locator('[data-testid="pptx-viewer"]')
    assert int(viewer.get_attribute("data-slide-index") or "0") == 0
    page.get_by_role("button", name="下一步").click()
    page.wait_for_function("() => Number(document.querySelector('[data-testid=\\"pptx-viewer\\"]')?.dataset.clickBoundary ?? 0) >= 1", timeout=5000)
    assert int(viewer.get_attribute("data-slide-index") or "0") == 0
    assert not errors, errors
    browser.close()
`)
  run("python3", [browserCheck], consumer)
  console.log("PASS: PPTX 压缩包内含补丁渲染器，并在工作区外完成安装、构建、预览和单击动画。")
} finally {
  previewServer?.kill("SIGINT")
  rmSync(temporary, { recursive: true, force: true })
}
