import { defineConfig } from "vite"
import vue from "@vitejs/plugin-vue"
import vueJsx from "@vitejs/plugin-vue-jsx"

export default defineConfig({
  plugins: [vue(), vueJsx()],
  build: {
    lib: {
      entry: {
        index: "src/index.ts",
        chart: "src/optional/chart.ts",
        map: "src/optional/map.ts",
        webgl: "src/optional/webgl.ts",
      },
      formats: ["es"],
      fileName: (_format, entryName) => `${entryName}.js`,
      cssFileName: "style",
    },
    rollupOptions: {
      external: [
        "vue",
        "@arcships/xlsx-core",
        "d3-geo",
        "d3-hierarchy",
        "d3-scale",
        "d3-shape",
        "regl",
        "topojson-client",
        "us-atlas",
        "world-atlas",
      ],
      output: {
        manualChunks(id) {
          return id.includes("/src/render/") ? "optional-renderers" : undefined
        },
      },
    },
  },
})
