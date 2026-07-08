import path from "path"
import { defineConfig } from "vite"
import vue from "@vitejs/plugin-vue"

const workspaceRoot = path.resolve(__dirname, "../..")

const workspaceSourceAliases: Record<string, string> = {
  "@extend-ai/docx-core": path.resolve(workspaceRoot, "packages/docx-core/src/index.ts"),
  "@extend-ai/vue-docx": path.resolve(workspaceRoot, "packages/vue-docx/src/index.ts"),
}

export default defineConfig({
  plugins: [vue()],
  server: { port: 5000 },
  worker: {
    format: "es",
  },
  optimizeDeps: {
    exclude: [
      "us-atlas",
      "world-atlas",
      ...Object.keys(workspaceSourceAliases),
    ],
  },
  resolve: {
    alias: {
      ...workspaceSourceAliases,
    },
    dedupe: ["vue"],
  },
})