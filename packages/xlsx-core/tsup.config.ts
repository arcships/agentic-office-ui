import { defineConfig } from "tsup"
import { copyFileSync } from "node:fs"
import { createRequire } from "node:module"
import { dirname, join } from "node:path"

export default defineConfig([
  {
    entry: ["src/index.ts", "src/xlsx-worker.ts"],
    format: ["esm"],
    dts: true,
    clean: true,
    external: ["@dukelib/sheets-wasm"],
    skipNodeModulesBundle: true,
    onSuccess: () => {
      const require = createRequire(import.meta.url)
      const wasmPath = join(
        dirname(require.resolve("@dukelib/sheets-wasm")),
        "duke_sheets_wasm_bg.wasm",
      )
      copyFileSync(wasmPath, new URL("dist/duke_sheets_wasm_bg.wasm", import.meta.url))
      console.log("Copied duke_sheets_wasm_bg.wasm to dist/")
    },
  },
])
