import { defineConfig } from "tsup";
import { copyFileSync } from "node:fs";

export default defineConfig({
  entry: [
    "src/index.ts",
    "src/core.ts",
    "src/runtime.ts",
    "src/wasm-url.ts",
    "src/viewer/docx-import-worker.ts",
    "src/viewer/docx-export-worker.ts",
  ],
  format: ["esm"],
  dts: true,
  clean: true,
  loader: {
    ".wasm": "file",
  },
  // Bundle the generated wasm-bindgen JS glue so consumers get one file.
  // The .wasm binary is copied separately via onSuccess.
  skipNodeModulesBundle: false,
  onSuccess: () => {
    copyFileSync(
      new URL("src/engine/generated/docx_wasm_bg.wasm", import.meta.url),
      new URL("dist/docx_wasm_bg.wasm", import.meta.url),
    );
    // Worker entries are emitted to dist/viewer/, but index.js references them
    // relative to dist/ — copy to dist root to match the URL resolution.
    copyFileSync(
      new URL("dist/viewer/docx-import-worker.js", import.meta.url),
      new URL("dist/docx-import-worker.js", import.meta.url),
    );
    copyFileSync(
      new URL("dist/viewer/docx-export-worker.js", import.meta.url),
      new URL("dist/docx-export-worker.js", import.meta.url),
    );
  },
});
