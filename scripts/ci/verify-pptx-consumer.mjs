#!/usr/bin/env node

import { execFileSync } from "node:child_process"
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..")
const temporary = mkdtempSync(path.join(tmpdir(), "pptx-consumer-"))
const packages = path.join(temporary, "packages")
const consumer = path.join(temporary, "consumer")
mkdirSync(packages)
mkdirSync(path.join(consumer, "src"), { recursive: true })

function run(command, args, cwd) {
  return execFileSync(command, args, { cwd, encoding: "utf8", stdio: ["ignore", "pipe", "inherit"] }).trim()
}

try {
  run("pnpm", ["pack", "--pack-destination", packages], path.join(root, "packages/pptx-core"))
  run("pnpm", ["pack", "--pack-destination", packages], path.join(root, "packages/vue-pptx"))
  const core = path.join(packages, "arcships-pptx-core-0.2.0.tgz")
  const vuePptx = path.join(packages, "arcships-vue-pptx-0.2.0.tgz")
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
    pnpm: { overrides: { "@arcships/pptx-core": `file:${core}` } },
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
import { createApp, h } from "vue"
import { PptxViewer } from "@arcships/vue-pptx"
import "@arcships/vue-pptx/style.css"
import { createPptxDocumentSession, type PptxDocumentSession } from "@arcships/pptx-core/browser"

const factory: typeof createPptxDocumentSession = createPptxDocumentSession
const sessionType: PptxDocumentSession | null = null
void factory
void sessionType
createApp({ render: () => h(PptxViewer, { mode: "present", autoplay: false }) }).mount("#app")
`)
  run("pnpm", ["install", "--ignore-workspace"], consumer)
  run("pnpm", ["build"], consumer)
  console.log("PASS: PPTX 压缩包可在工作区外完成类型检查和正式构建。")
} finally {
  rmSync(temporary, { recursive: true, force: true })
}
