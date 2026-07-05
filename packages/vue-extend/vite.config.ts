import { defineConfig } from "vite"
import vue from "@vitejs/plugin-vue"

export default defineConfig({
  plugins: [vue()],
  build: {
    lib: { entry: "src/index.ts", formats: ["es"], fileName: "index" },
    rollupOptions: {
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
        "signature_pad",
      ],
    },
  },
})
