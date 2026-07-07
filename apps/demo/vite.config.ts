import { defineConfig } from "vite"
import vue from "@vitejs/plugin-vue"

export default defineConfig({
  plugins: [vue()],
  server: { port: 5000 },
  worker: {
    format: "es",
  },
  optimizeDeps: {
    exclude: ["us-atlas", "world-atlas"],
  },
})
