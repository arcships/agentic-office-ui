import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";

// The reference app only exists for the pixel-parity harness: it renders the
// minimum vendored subset of the pinned extend-hq/ui reference (see
// src/vendor/NOTICE.md) with the same fixtures as tools/vue-reference. It is
// never published.
export default defineConfig({
  plugins: [react(), tailwindcss()],
  build: {
    outDir: "dist",
    emptyOutDir: true,
    target: "chrome120",
  },
  server: {
    fs: {
      allow: ["..", "../.."],
    },
  },
});
