import { defineConfig } from "vite"
import vue from "@vitejs/plugin-vue"

export default defineConfig({
  plugins: [vue()],
  build: {
    lib: { entry: "src/index.ts", formats: ["es"], fileName: "index", cssFileName: "style" },
    rollupOptions: {
      output: {
        assetFileNames: (assetInfo) =>
          assetInfo.names.some((name) => name.endsWith(".css"))
            ? "[name][extname]"
            : "assets/[name]-[hash][extname]",
      },
      external: ["vue", "signature_pad", "@arcships/office-interaction"],
    },
  },
})
