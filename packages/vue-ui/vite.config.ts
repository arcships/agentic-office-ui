import { readFileSync, writeFileSync } from "node:fs"
import { defineConfig, type Plugin } from "vite"
import vue from "@vitejs/plugin-vue"

const TOKENS_IMPORT = '@import "./tokens.css";'

/**
 * Emit the design-token stylesheets as standalone package assets.
 *
 * `theme.css` is the public entry (`@arcships/vue-ui/theme.css`) and is made
 * self-contained: its source `@import` is replaced by the token block so a
 * host can load a single CSS file (or let a bundler inline it). `tokens.css`
 * is still emitted for hosts that only want the token contract.
 */
function officeThemeStyles(): Plugin {
  return {
    name: "office-theme-styles",
    apply: "build",
    closeBundle() {
      const tokensUrl = new URL("./src/styles/tokens.css", import.meta.url)
      const themeUrl = new URL("./src/styles/theme.css", import.meta.url)
      const tokens = readFileSync(tokensUrl, "utf8")
      const themeSource = readFileSync(themeUrl, "utf8")
      if (!themeSource.includes(TOKENS_IMPORT)) {
        throw new Error("theme.css must import ./tokens.css")
      }
      const theme = themeSource.replace(TOKENS_IMPORT, "")
      writeFileSync(new URL("./dist/tokens.css", import.meta.url), tokens)
      writeFileSync(new URL("./dist/theme.css", import.meta.url), `${tokens}\n${theme}`)
    },
  }
}

export default defineConfig({
  plugins: [vue(), officeThemeStyles()],
  build: {
    lib: { entry: "src/index.ts", formats: ["es"], fileName: "index", cssFileName: "style" },
    rollupOptions: {
      output: {
        assetFileNames: (assetInfo) =>
          assetInfo.names.some((name) => name.endsWith(".css"))
            ? "[name][extname]"
            : "assets/[name]-[hash][extname]",
      },
      external: ["vue", "signature_pad", "@arcships/office-interaction", "@lucide/vue"],
    },
  },
})
