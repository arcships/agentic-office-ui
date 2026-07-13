import { defineConfig } from "tsup"

export default defineConfig({
  entry: {
    index: "src/index.ts",
    browser: "src/browser.ts",
  },
  format: ["esm"],
  dts: true,
  clean: true,
  noExternal: [
    /^@arcships\/office-runtime/u,
    /^@aiden0z\/pptx-renderer$/u,
    /^fflate$/u,
  ],
  skipNodeModulesBundle: true,
})
