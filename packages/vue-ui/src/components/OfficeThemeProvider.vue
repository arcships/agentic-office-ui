<script setup lang="ts">
import { computed, onBeforeUnmount, onMounted, ref } from "vue"
import type { OfficeMessagesOverride } from "@arcships/office-interaction"
import {
  OFFICE_DARK_MEDIA_QUERY,
  provideOfficeTheme,
  resolveOfficeTheme,
  type OfficeDensity,
  type OfficeTheme,
  type OfficeThemeContext,
} from "../composables/office-theme"
import { provideOfficeUiMessages, useOfficeUiLocale, useOfficeUiMessages } from "../i18n"
import { provideOfficePortalTarget } from "../composables/office-portal"

export interface OfficeThemeProviderProps {
  /** Requested scheme. `system` follows `prefers-color-scheme`. */
  theme?: OfficeTheme
  /** BCP 47 locale written to `lang`; also drives the shared message catalog. */
  locale?: string
  /** Optional control density. Defaults to `comfortable`. */
  density?: OfficeDensity
  /** Optional message overrides merged on top of the locale catalog. */
  messages?: OfficeMessagesOverride
  /** Element used for the provider host. Defaults to `div`. */
  as?: string
}

const props = withDefaults(defineProps<OfficeThemeProviderProps>(), {
  theme: "system",
  locale: undefined,
  density: "comfortable",
  messages: undefined,
  as: "div",
})

// Inherit an ambient Office locale/catalog so a provider nested inside a viewer
// does not silently reset translation to English.
const inheritedLocale = useOfficeUiLocale()
const inheritedMessages = useOfficeUiMessages()

const prefersDark = ref(false)
const resolvedTheme = computed(() => resolveOfficeTheme(props.theme, prefersDark.value))

const context: OfficeThemeContext = Object.freeze({
  theme: computed(() => props.theme),
  resolvedTheme,
  density: computed(() => props.density),
  locale: computed(() => props.locale ?? inheritedLocale.value),
  isDark: computed(() => resolvedTheme.value === "dark"),
})

provideOfficeTheme(context)
provideOfficeUiMessages(
  () => props.locale ?? inheritedLocale.value,
  () => {
    if (props.messages !== undefined) return props.messages
    // The fallback catalog is English, so merging it as an override would
    // replace every zh-CN key. Only inherit an ambient catalog when the host
    // did not select a locale explicitly.
    if (props.locale !== undefined) return undefined
    return inheritedMessages.value
  },
)

// Portalled overlays must stay inside the themed subtree; see office-portal.ts.
const portalTarget = ref<HTMLElement | null>(null)
provideOfficePortalTarget(portalTarget)

let mediaQuery: MediaQueryList | undefined

function handleMediaChange(event: MediaQueryListEvent): void {
  prefersDark.value = event.matches
}

function teardownMediaQuery(): void {
  if (!mediaQuery) return
  if (typeof mediaQuery.removeEventListener === "function") {
    mediaQuery.removeEventListener("change", handleMediaChange)
  } else {
    mediaQuery.removeListener?.(handleMediaChange)
  }
  mediaQuery = undefined
}

onMounted(() => {
  // SSR and non-browser renderers never reach this hook; the guard also keeps
  // hand-rolled renderers without `matchMedia` safe.
  if (typeof window === "undefined" || typeof window.matchMedia !== "function") return
  mediaQuery = window.matchMedia(OFFICE_DARK_MEDIA_QUERY)
  prefersDark.value = mediaQuery.matches
  if (typeof mediaQuery.addEventListener === "function") {
    mediaQuery.addEventListener("change", handleMediaChange)
  } else {
    mediaQuery.addListener?.(handleMediaChange)
  }
})

onBeforeUnmount(teardownMediaQuery)
</script>

<template>
  <component
    :is="as"
    class="office-theme"
    :data-office-theme="theme"
    :data-office-theme-resolved="resolvedTheme"
    :data-office-density="density"
    :lang="locale"
  >
    <span ref="portalTarget" class="office-theme__portal" />
    <slot />
  </component>
</template>

<style scoped>
/* `contents` keeps the provider out of layout so wrapping a viewer or an app
 * shell cannot change flex/grid sizing. Custom properties still inherit. */
.office-theme {
  display: contents;
}

/* A Teleport target inherits the provider's tokens without joining layout. */
.office-theme__portal {
  display: contents;
}
</style>
