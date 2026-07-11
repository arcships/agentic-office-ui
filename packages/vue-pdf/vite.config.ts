import { defineConfig } from "vite"
import vue from "@vitejs/plugin-vue"
import { readFileSync } from "node:fs"
import { createRequire } from "node:module"

const require = createRequire(import.meta.url)
const pdfiumWasmRuntimeExpression = 'new URL("./pdfium.wasm", import.meta.url).href'
const pdfiumWasmRuntimeMarker = "__AGENTIC_OFFICE_PDFIUM_WASM_URL__"

function packagePdfiumWasm() {
  let runtimeUrlFound = false
  let runtimeUrlRestored = false

  return {
    name: "package-pdfium-wasm",
    enforce: "pre" as const,
    buildStart() {
      this.emitFile({
        type: "asset",
        fileName: "pdfium.wasm",
        source: readFileSync(require.resolve("@embedpdf/pdfium/pdfium.wasm")),
      })
    },
    transform(code: string) {
      if (!code.includes(pdfiumWasmRuntimeExpression)) return null
      runtimeUrlFound = true
      return {
        code: code.replaceAll(
          pdfiumWasmRuntimeExpression,
          JSON.stringify(pdfiumWasmRuntimeMarker),
        ),
        map: null,
      }
    },
    renderChunk(code: string) {
      if (!code.includes(pdfiumWasmRuntimeMarker)) return null
      runtimeUrlRestored = true
      return {
        code: code.replaceAll(
          JSON.stringify(pdfiumWasmRuntimeMarker),
          pdfiumWasmRuntimeExpression,
        ),
        map: null,
      }
    },
    generateBundle() {
      if (runtimeUrlFound && !runtimeUrlRestored) {
        this.error("PDFium WASM URL was not restored as a package-relative URL")
      }
    },
  }
}

export default defineConfig({
  plugins: [packagePdfiumWasm(), vue()],
  build: {
    lib: { entry: "src/index.ts", formats: ["es"], fileName: "index", cssFileName: "style" },
    rollupOptions: {
      output: {
        assetFileNames: (assetInfo) =>
          assetInfo.names.includes("pdfium.wasm")
            ? "pdfium.wasm"
            : assetInfo.names.some((name) => name.endsWith(".css"))
              ? "[name][extname]"
            : "assets/[name]-[hash][extname]",
      },
      external: [
        "vue",
        "@embedpdf/core",
        "@embedpdf/core/vue",
        "@embedpdf/models",
        "@embedpdf/plugin-document-manager/vue",
        "@embedpdf/plugin-viewport/vue",
        "@embedpdf/plugin-scroll/vue",
        "@embedpdf/plugin-render/vue",
        "@embedpdf/plugin-tiling/vue",
        "@embedpdf/plugin-search/vue",
        "@embedpdf/plugin-selection/vue",
        "@embedpdf/plugin-thumbnail/vue",
        "@embedpdf/plugin-zoom/vue",
        "@embedpdf/plugin-rotate/vue",
        "@embedpdf/plugin-interaction-manager/vue",
      ],
    },
  },
})
