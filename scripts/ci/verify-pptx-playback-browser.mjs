#!/usr/bin/env node

import { spawn, spawnSync } from "node:child_process"
import path from "node:path"
import process from "node:process"

const root = path.resolve(import.meta.dirname, "../..")
const python = process.env.PYTHON ?? "python3"
const port = Number(process.env.PPTX_TEST_PORT ?? 5197)
const appUrl = `http://127.0.0.1:${port}`
const demoPort = Number(process.env.PPTX_DEMO_TEST_PORT ?? 5198)
const demoUrl = `http://127.0.0.1:${demoPort}`
const browsers = (process.env.PPTX_BROWSERS ?? "chromium,firefox,webkit")
  .split(",")
  .map((browser) => browser.trim())
  .filter(Boolean)
for (const browser of browsers) {
  if (!["chromium", "firefox", "webkit"].includes(browser)) {
    throw new Error(`Unsupported PPTX test browser: ${browser}`)
  }
}
const sampleUrl = `${appUrl}/@fs${root}/tests/fixtures/pptx/playback-controlled.pptx`
const powerPointSavedSampleUrl = `${appUrl}/@fs${root}/tests/fixtures/pptx/playback/media-audio-video.pptx`
const morphSampleUrl = `${appUrl}/@fs${root}/tests/fixtures/pptx/playback/morph-explicit-name.pptx`
const runVueChecks = process.env.PPTX_VUE !== "0"
const server = spawn("pnpm", ["--filter=pptx-playback-lab", "dev", "--host", "127.0.0.1", "--port", String(port)], {
  cwd: root,
  detached: true,
  stdio: ["ignore", "ignore", "inherit"],
  detached: true,
})
const demoServer = runVueChecks
  ? spawn("pnpm", ["--filter=demo", "dev", "--host", "127.0.0.1", "--port", String(demoPort)], {
      cwd: root,
      detached: true,
      stdio: ["ignore", "ignore", "inherit"],
    })
  : null

function stopServer(child) {
  if (!child || !child.pid || child.exitCode !== null) return
  try {
    process.kill(-child.pid, "SIGINT")
  } catch { /* 进程已退出。 */ }
}

async function waitForServer(url) {
  for (let attempt = 0; attempt < 100; attempt += 1) {
    try {
      const response = await fetch(url)
      if (response.ok) return
    } catch { /* 服务仍在启动。 */ }
    await new Promise((resolve) => setTimeout(resolve, 100))
  }
  throw new Error("PPTX 浏览器验收服务启动超时。")
}

function run(browser, extra, fileUrl = sampleUrl) {
  const result = spawnSync(python, [
    "tests/blackbox/pptx_playback_model.py",
    "--browser", browser,
    "--app-url", appUrl,
    "--sample-url", fileUrl,
    "--quiet",
    ...extra,
  ], { cwd: root, stdio: "inherit" })
  if (result.status !== 0) throw new Error(`${browser} PPTX 浏览器验收失败。`)
}

function runStateMachine(browser) {
  const result = spawnSync(python, [
    "tests/blackbox/pptx_state_machine.py",
    "--browser", browser,
    "--app-url", appUrl,
    "--sample-url", sampleUrl,
    "--quiet",
  ], { cwd: root, stdio: "inherit" })
  if (result.status !== 0) throw new Error(`${browser} PPTX 状态机浏览器验收失败。`)
}


function runVue(browser) {
  const result = spawnSync(python, [
    "tests/blackbox/pptx_vue_present.py",
    "--browser", browser,
    "--app-url", demoUrl,
    "--fixture", path.join(root, "tests/fixtures/pptx/playback-controlled.pptx"),
  ], { cwd: root, stdio: "inherit" })
  if (result.status !== 0) throw new Error(`${browser} PPTX Vue 演示组件验收失败。`)
}

try {
  await Promise.all([
    waitForServer(appUrl),
    ...(runVueChecks ? [waitForServer(demoUrl)] : []),
  ])
  for (const browser of browsers) {
    run(browser, ["--initial-slide", "0", "--steps", "2", "--transition-check", "--pause-check"])
    run(browser, ["--initial-slide", "3", "--auto-advance-check", "--approximation", "safe"])
    run(browser, ["--initial-slide", "5", "--media-check"])
    run(browser, ["--initial-slide", "6", "--repeat-check"])
    run(browser, ["--initial-slide", "7", "--indefinite-check"])
    run(browser, ["--initial-slide", "10", "--overlap-check"])
    run(browser, ["--initial-slide", "11", "--paragraph-check"])
    run(browser, ["--initial-slide", "0", "--effect-timing-check"])
    run(browser, ["--initial-slide", "0", "--steps", "1", "--previous-check"])
    if (runVueChecks) runVue(browser)
  }
  if (browsers.includes("chromium")) {
    run("chromium", ["--initial-slide", "7", "--hidden-check"])
    run("chromium", ["--initial-slide", "9", "--action-check"])
    run("chromium", ["--initial-slide", "5", "--media-check", "--resource-check", "--performance-check"])
    run("chromium", ["--initial-slide", "0", "--rapid-navigation-check", "--approximation", "safe"])
    run("chromium", ["--initial-slide", "0", "--search-cache-check"])
    runStateMachine("chromium")
    run("chromium", ["--initial-slide", "0", "--duplicate-target-check"])
    run("chromium", ["--initial-slide", "0", "--reset-check"])
    run("chromium", ["--initial-slide", "5", "--media-block-check"])
    run("chromium", ["--initial-slide", "5", "--media-check", "--resource-check"], powerPointSavedSampleUrl)
    run("chromium", ["--initial-slide", "4", "--morph-check", "--approximation", "safe"], morphSampleUrl)
    const memory = spawnSync(python, [
      "tests/blackbox/pptx_memory.py",
      "--app-url", appUrl,
      "--sample-url", sampleUrl,
    ], { cwd: root, stdio: "inherit" })
    if (memory.status !== 0) throw new Error("PPTX 连续重播内存验收失败。")
    const baselines = spawnSync(python, [
      "scripts/fixtures/capture-pptx-formal-baselines.py",
      "--check",
    ], { cwd: root, stdio: "inherit" })
    if (baselines.status !== 0) throw new Error("PPTX 浏览器事件基准校验失败。")
  }
  console.log(`PASS: PPTX ${browsers.join("/")} 网页演示、状态机、切换、重复、自动换页、Morph、媒体、书签、冲突裁决、逐段文字、正式素材、基准复核、内存和资源释放验收通过。`)
} finally {
  stopServer(server)
  stopServer(demoServer)
}
