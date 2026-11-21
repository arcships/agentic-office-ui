import { readdirSync } from "node:fs";
import path from "node:path";

import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";

const root = import.meta.dirname;

// L2 route entries are discovered from this app root so a new pilot route is
// one file, not a build-config edit. `index.html` is the L1 surface harness.
const routeEntries = Object.fromEntries(
  readdirSync(root)
    .filter((file) => file.endsWith(".html"))
    .map((file) => [path.basename(file, ".html"), path.join(root, file)]),
);

// The reference app only exists for the pixel-parity harness: it renders the
// minimum vendored subset of the pinned extend-hq/ui reference (see
// src/vendor/NOTICE.md) with the same fixtures as tools/vue-reference. It is
// never published.
export default defineConfig({
  plugins: [react(), tailwindcss()],
  resolve: {
    // The upstream files under src/upstream are vendored byte-for-byte, so
    // their `@/...` specifiers must keep resolving without rewrites.
    alias: {
      "@": path.join(root, "src/upstream"),
    },
  },
  build: {
    outDir: "dist",
    emptyOutDir: true,
    target: "chrome120",
    rollupOptions: {
      input: routeEntries,
    },
  },
  server: {
    fs: {
      allow: ["..", "../.."],
    },
  },
});
