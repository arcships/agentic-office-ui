import { rmSync } from "node:fs"
import path from "path"
import { defineConfig } from "vite"
import vue from "@vitejs/plugin-vue"

const workspaceRoot = path.resolve(__dirname, "../..")

const workspaceSourceAliasIds = [
  "@arcships/docx-core",
  "@arcships/pptx-core",
  "@arcships/pptx-core/browser",
  "@arcships/vue-docx",
  "@arcships/vue-pptx",
] as const
const workspaceSourceAliases = [
  {
    find: /^@arcships\/docx-core$/,
    replacement: path.resolve(workspaceRoot, "packages/docx-core/src/index.ts"),
  },
  {
    find: /^@arcships\/vue-docx$/,
    replacement: path.resolve(workspaceRoot, "packages/vue-docx/src/index.ts"),
  },
  {
    find: /^@arcships\/pptx-core\/browser$/,
    replacement: path.resolve(workspaceRoot, "packages/pptx-core/src/browser.ts"),
  },
  {
    find: /^@arcships\/pptx-core$/,
    replacement: path.resolve(workspaceRoot, "packages/pptx-core/src/index.ts"),
  },
  {
    find: /^@arcships\/vue-pptx$/,
    replacement: path.resolve(workspaceRoot, "packages/vue-pptx/src/index.ts"),
  },
]

function omitLegacyPublicWasmCopies() {
  return {
    name: "omit-legacy-public-wasm-copies",
    apply: "build" as const,
    closeBundle() {
      for (const fileName of ["docx_wasm_bg.wasm", "duke_sheets_wasm_bg.wasm"]) {
        rmSync(path.resolve(__dirname, "dist", fileName), { force: true })
      }
    },
  }
}

export default defineConfig({
  plugins: [vue(), omitLegacyPublicWasmCopies()],
  server: { port: 5000 },
  worker: {
    format: "es",
  },
  optimizeDeps: {
    exclude: [
      "us-atlas",
      "world-atlas",
      ...workspaceSourceAliasIds,
    ],
  },
  resolve: {
    alias: workspaceSourceAliases,
    dedupe: ["vue"],
  },
  build: {
    rollupOptions: {
      output: {
        chunkFileNames: "assets/chunk-[name]-[hash].js",
      },
    },
  },
})
