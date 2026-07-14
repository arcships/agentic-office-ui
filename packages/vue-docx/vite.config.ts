import { defineConfig } from "vite"
import vue from "@vitejs/plugin-vue"

export default defineConfig({
  plugins: [vue()],
  build: {
    lib: { entry: "src/index.ts", formats: ["es"], fileName: "index", cssFileName: "style" },
    rollupOptions: {
      external: ["vue", "@arcships/docx-core", "@arcships/office-interaction"],
    },
  },
})
