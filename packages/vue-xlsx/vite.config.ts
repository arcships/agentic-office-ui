import { defineConfig } from "vite"
import vue from "@vitejs/plugin-vue"
import vueJsx from "@vitejs/plugin-vue-jsx"

export default defineConfig({
  plugins: [vue(), vueJsx()],
  build: {
    lib: { entry: "src/index.ts", formats: ["es"], fileName: "index" },
    rollupOptions: {
      external: [
        "vue",
        "@extend-ai/xlsx-core",
        "d3-geo",
        "d3-hierarchy",
        "d3-scale",
        "d3-shape",
        "regl",
        "topojson-client",
        "us-atlas",
        "world-atlas",
      ],
    },
  },
})
