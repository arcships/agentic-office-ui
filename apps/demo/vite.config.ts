import path from "path"
import { defineConfig } from "vite"
import vue from "@vitejs/plugin-vue"

const workspaceRoot = path.resolve(__dirname, "../..")

const workspaceSourceAliasIds = ["@extend-ai/docx-core", "@extend-ai/vue-docx"] as const
const workspaceSourceAliases = [
  {
    find: /^@extend-ai\/docx-core$/,
    replacement: path.resolve(workspaceRoot, "packages/docx-core/src/index.ts"),
  },
  {
    find: /^@extend-ai\/vue-docx$/,
    replacement: path.resolve(workspaceRoot, "packages/vue-docx/src/index.ts"),
  },
]

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
      ...workspaceSourceAliasIds,
    ],
  },
  resolve: {
    alias: workspaceSourceAliases,
    dedupe: ["vue"],
  },
})
