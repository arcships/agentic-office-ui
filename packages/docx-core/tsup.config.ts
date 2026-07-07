import { defineConfig } from "tsup"
import { copyFileSync } from "node:fs"

export default defineConfig([
  {
    entry: ["src/index.ts", "src/docx-import-worker.ts"],
    format: ["esm"],
    dts: true,
    clean: true,
    external: ["@chenglou/pretext", "fast-png", "utif"],
    skipNodeModulesBundle: true,
    onSuccess: () => {
      copyFileSync(
        new URL("src/docx_wasm_bg.wasm", import.meta.url),
        new URL("dist/docx_wasm_bg.wasm", import.meta.url),
      )
      console.log("Copied docx_wasm_bg.wasm to dist/")
    },
  },
])
