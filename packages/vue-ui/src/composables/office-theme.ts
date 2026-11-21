import {
  computed,
  inject,
  provide,
  type ComputedRef,
  type InjectionKey,
} from "vue"

/** Requested theme mode. `system` follows `prefers-color-scheme`. */
export type OfficeTheme = "light" | "dark" | "system"

/** Concrete scheme a component should paint with. */
export type OfficeResolvedTheme = "light" | "dark"

/** Optional density of Office control chrome. */
export type OfficeDensity = "comfortable" | "compact"

/** Attributes written by `OfficeThemeProvider`; CSS is scoped to these. */
export const OFFICE_THEME_ATTRIBUTE = "data-office-theme"
export const OFFICE_THEME_RESOLVED_ATTRIBUTE = "data-office-theme-resolved"
export const OFFICE_DENSITY_ATTRIBUTE = "data-office-density"

/** The media query behind `theme="system"`. */
export const OFFICE_DARK_MEDIA_QUERY = "(prefers-color-scheme: dark)"

/**
 * Theme context published by `OfficeThemeProvider`.
 *
 * All members are read-only computed refs so descendants cannot mutate the
 * provider state; the provider owns the source of truth.
 */
export interface OfficeThemeContext {
  readonly theme: ComputedRef<OfficeTheme>
  readonly resolvedTheme: ComputedRef<OfficeResolvedTheme>
  readonly density: ComputedRef<OfficeDensity>
  readonly locale: ComputedRef<string | undefined>
  readonly isDark: ComputedRef<boolean>
}

export const OFFICE_THEME_INJECTION_KEY: InjectionKey<OfficeThemeContext> =
  Symbol("arcships.office-theme")

/**
 * Resolve a requested theme against the current system preference.
 *
 * Pure and framework-free so the same rule runs on the server, in tests and in
 * the browser. Unknown runtime values fall back to `system`.
 */
export function resolveOfficeTheme(
  theme: OfficeTheme,
  prefersDark: boolean,
): OfficeResolvedTheme {
  if (theme === "dark") return "dark"
  if (theme === "light") return "light"
  return prefersDark ? "dark" : "light"
}

/**
 * Read the OS colour preference without touching `window` on the server.
 * Returns `false` when `matchMedia` is unavailable, which matches the light
 * baseline in `tokens.css`.
 */
export function officeSystemPrefersDark(): boolean {
  if (typeof window === "undefined" || typeof window.matchMedia !== "function") {
    return false
  }
  return window.matchMedia(OFFICE_DARK_MEDIA_QUERY).matches
}

const fallbackContext: OfficeThemeContext = Object.freeze({
  theme: computed<OfficeTheme>(() => "system"),
  resolvedTheme: computed<OfficeResolvedTheme>(() => "light"),
  density: computed<OfficeDensity>(() => "comfortable"),
  locale: computed(() => undefined),
  isDark: computed(() => false),
})

/** Publish a theme context to descendants. */
export function provideOfficeTheme(context: OfficeThemeContext): OfficeThemeContext {
  provide(OFFICE_THEME_INJECTION_KEY, context)
  return context
}

/**
 * Read the nearest `OfficeThemeProvider` context.
 *
 * Rendering outside a provider is supported: components see the light
 * baseline (`system` / `light` / `comfortable`) instead of failing.
 */
export function useOfficeTheme(): OfficeThemeContext {
  return inject(OFFICE_THEME_INJECTION_KEY, fallbackContext)
}
