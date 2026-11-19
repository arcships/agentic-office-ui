import { createRequire } from "node:module";
import { readFileSync } from "node:fs";
import { fileURLToPath, URL } from "node:url";

import { defineConfig } from "vite";
import vue from "@vitejs/plugin-vue";

const require = createRequire(import.meta.url);
const here = fileURLToPath(new URL(".", import.meta.url));
const repoRoot = fileURLToPath(new URL("../..", import.meta.url));

/**
 * The comparison app imports @arcships/vue-ui and @arcships/vue-pdf sources
 * directly so the harness measures the working tree, not a stale dist. The
 * two package aliases below point at source entry points; everything else
 * resolves through the workspace links.
 */
export default defineConfig({
  plugins: [vue(), pdfiumWasm()],
  resolve: {
    alias: [
      {
        find: "@arcships/vue-ui/theme.css",
        replacement: fileURLToPath(
          new URL("../../packages/vue-ui/src/styles/theme.css", import.meta.url),
        ),
      },
      {
        find: "@arcships/vue-ui",
        replacement: fileURLToPath(new URL("./src/shims/vue-ui.ts", import.meta.url)),
      },
      {
        find: "@arcships/office-interaction",
        replacement: fileURLToPath(
          new URL("../../packages/office-interaction/src/index.ts", import.meta.url),
        ),
      },
    ],
  },
  build: {
    outDir: "dist",
    emptyOutDir: true,
    target: "chrome120",
  },
  server: {
    fs: {
      allow: [here, repoRoot],
    },
  },
});

/**
 * @arcships/vue-pdf ships pdfium.wasm through its library build. When its
 * source is consumed directly, emit the same wasm asset and point the runtime
 * URL at the app root so the PDF engine loads in the harness page.
 */
function pdfiumWasm() {
  const runtimeUrl = 'new URL("./pdfium.wasm", import.meta.url).href';
  const appUrl = 'new URL("/pdfium.wasm", import.meta.url).href';
  return {
    name: "vue-reference-pdfium-wasm",
    enforce: "pre" as const,
    buildStart() {
      this.emitFile({
        type: "asset",
        fileName: "pdfium.wasm",
        source: readFileSync(require.resolve("@embedpdf/pdfium/pdfium.wasm")),
      });
    },
    transform(code: string) {
      if (!code.includes(runtimeUrl)) return null;
      return { code: code.replaceAll(runtimeUrl, appUrl), map: null };
    },
  };
}
