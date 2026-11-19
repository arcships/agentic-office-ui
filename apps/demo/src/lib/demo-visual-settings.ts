/**
 * Deterministic visual-matrix switches for the verification demo.
 *
 * The screenshot harness appends `?theme=dark` (and, in the future, other
 * stable switches) to the preview URL. The demo is a verification surface, so
 * keeping the switch in the URL makes every captured state reproducible by
 * opening the same address in a browser.
 */
export type DemoTheme = "light" | "dark"

const THEME_PARAM = "theme"

export function readDemoTheme(): DemoTheme {
  if (typeof window === "undefined") return "light"
  const value = new URLSearchParams(window.location.search).get(THEME_PARAM)
  return value === "dark" ? "dark" : "light"
}

/**
 * Applies the requested demo chrome theme. Components that own a separate
 * document theme (DOCX/XLSX/PDF/PPTX viewers) read `readDemoTheme()` directly
 * so their internal dark palette follows the same URL switch.
 */
export function applyDemoTheme(): DemoTheme {
  const theme = readDemoTheme()
  if (typeof document !== "undefined") {
    document.documentElement.dataset.theme = theme
    document.documentElement.style.colorScheme = theme
  }
  return theme
}
