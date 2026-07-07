import { defineConfig } from "tsup";
import { copyFileSync } from "node:fs";

export default defineConfig({
  entry: ["src/index.ts"],
  format: ["esm"],
  dts: true,
  clean: true,
  // Bundle the generated wasm-bindgen JS glue so consumers get one file.
  // The .wasm binary is copied separately via onSuccess.
  skipNodeModulesBundle: false,
  onSuccess: () => {
    copyFileSync(
      new URL("src/engine/generated/docx_wasm_bg.wasm", import.meta.url),
      new URL("dist/docx_wasm_bg.wasm", import.meta.url),
    );
  },
});
