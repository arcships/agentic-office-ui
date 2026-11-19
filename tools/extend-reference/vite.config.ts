import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";

// The reference app only exists for the pixel-parity harness: it renders the
// minimum vendored subset of extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
// with the same fixtures as tools/vue-reference. It is never published.
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
