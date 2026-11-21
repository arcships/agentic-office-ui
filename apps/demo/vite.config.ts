import { copyFileSync, existsSync, mkdirSync, rmSync, statSync } from "node:fs"
import { createRequire } from "node:module"
import path from "path"
import { defineConfig, type Plugin } from "vite"
import vue from "@vitejs/plugin-vue"

const workspaceRoot = path.resolve(__dirname, "../..")
const xlsxRequire = createRequire(path.resolve(workspaceRoot, "packages/xlsx-core/package.json"))
const pdfRequire = createRequire(path.resolve(workspaceRoot, "packages/vue-pdf/package.json"))

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

/**
 * The published packages resolve to `dist/`, so serving them from source
 * requires a complete alias table plus the assets their build steps normally
 * copy. Keeping the extra aliases dev-only preserves the packaged consumer
 * boundary exercised by `pnpm build`.
 */
const devSourceAliasIds = [
  "@arcships/office-interaction",
  "@arcships/vue-pdf",
  "@arcships/vue-ui",
  "@arcships/vue-xlsx",
  "@arcships/xlsx-core",
  "@arcships/xlsx-core/runtime",
] as const
const devSourceAliases = [
  {
    find: /^@arcships\/office-interaction$/,
    replacement: path.resolve(workspaceRoot, "packages/office-interaction/src/index.ts"),
  },
  {
    find: /^@arcships\/vue-pdf$/,
    replacement: path.resolve(workspaceRoot, "packages/vue-pdf/src/index.ts"),
  },
  {
    find: /^@arcships\/vue-ui$/,
    replacement: path.resolve(workspaceRoot, "packages/vue-ui/src/index.ts"),
  },
  {
    find: /^@arcships\/vue-xlsx$/,
    replacement: path.resolve(workspaceRoot, "packages/vue-xlsx/src/index.ts"),
  },
  {
    find: /^@arcships\/xlsx-core\/runtime$/,
    replacement: path.resolve(workspaceRoot, "packages/xlsx-core/src/runtime.ts"),
  },
  {
    find: /^@arcships\/xlsx-core$/,
    replacement: path.resolve(workspaceRoot, "packages/xlsx-core/src/index.ts"),
  },
]
// Package entry points import their own styles; SFC styles are injected by the
// Vue plugin. The explicit `style.css` imports in main.ts only matter for the
// built (dist) packages, so point them at an empty file while serving source.
const devStyleAliases = [
  {
    find: /^@arcships\/vue-ui\/theme\.css$/,
    replacement: path.resolve(workspaceRoot, "packages/vue-ui/src/styles/theme.css"),
  },
  {
    find: /^@arcships\/(?:vue-docx|vue-pdf|vue-pptx|vue-ui|vue-xlsx)\/style\.css$/,
    replacement: path.resolve(__dirname, "src/dev-empty.css"),
  },
]

const devWasmCopies = [
  {
    source: path.join(
      path.dirname(xlsxRequire.resolve("@dukelib/sheets-wasm")),
      "duke_sheets_wasm_bg.wasm",
    ),
    target: path.resolve(
      workspaceRoot,
      "packages/xlsx-core/src/assets/duke_sheets_wasm_bg.wasm",
    ),
  },
  {
    source: pdfRequire.resolve("@embedpdf/pdfium/pdfium.wasm"),
    target: path.resolve(workspaceRoot, "packages/vue-pdf/src/pdf/pdfium.wasm"),
  },
]

function prepareDevWasmAssets(): Plugin {
  return {
    name: "prepare-dev-wasm-assets",
    apply: "serve",
    buildStart() {
      for (const { source, target } of devWasmCopies) {
        if (existsSync(target) && statSync(target).size === statSync(source).size) continue
        mkdirSync(path.dirname(target), { recursive: true })
        copyFileSync(source, target)
      }
    },
  }
}

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

export default defineConfig(({ command }) => {
  const serve = command === "serve"

  return {
    plugins: [prepareDevWasmAssets(), vue(), omitLegacyPublicWasmCopies()],
    server: { port: 5173 },
    worker: {
      format: "es",
    },
    optimizeDeps: {
      exclude: [
        "us-atlas",
        "world-atlas",
        ...workspaceSourceAliasIds,
        ...(serve ? devSourceAliasIds : []),
      ],
    },
    resolve: {
      alias: [
        ...(serve ? devStyleAliases : []),
        ...workspaceSourceAliases,
        ...(serve ? devSourceAliases : []),
      ],
      dedupe: ["vue"],
    },
    build: {
      rollupOptions: {
        output: {
          chunkFileNames: "assets/chunk-[name]-[hash].js",
        },
      },
    },
  }
})
