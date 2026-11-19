<script setup lang="ts">
import { computed, type Component } from "vue"
import { resolveOfficeIcon, useOfficeIcons } from "../icons"

export interface OfficeIconProps {
  /** Registry name, e.g. `chevron-down` or `ChevronDown`. */
  name?: string
  /** Explicit component; wins over `name` (host icons, one-off overrides). */
  icon?: Component
  /** Size override. Numbers are pixels; strings accept any CSS length. */
  size?: number | string
  /** Lucide stroke width. Defaults to `2`, matching the reference UI. */
  strokeWidth?: number
  /** Accessible name. Without it the icon is decorative (`aria-hidden`). */
  label?: string
}

const props = withDefaults(defineProps<OfficeIconProps>(), {
  name: undefined,
  icon: undefined,
  size: undefined,
  strokeWidth: 2,
  label: undefined,
})

const registry = useOfficeIcons()
const warnedIcons = new Set<string>()

const resolvedIcon = computed<Component | undefined>(() => {
  if (props.icon) return props.icon
  if (!props.name) return undefined
  const icon = resolveOfficeIcon(props.name, registry)
  if (!icon) warnMissingIcon(props.name)
  return icon
})

const sizeAttrs = computed<Record<string, unknown>>(() =>
  props.size === undefined
    ? {}
    : { style: { "--office-icon-size": toCssLength(props.size) } },
)

// Keep the SVG width/height attributes correct for SSR and for hosts that
// render without the package stylesheet. CSS (and the size token) still wins
// whenever the styles are present.
const attributeSize = computed(() =>
  typeof props.size === "number" ? props.size : 16,
)

function toCssLength(value: number | string): string {
  return typeof value === "number" ? `${value}px` : value
}

function warnMissingIcon(name: string): void {
  if (warnedIcons.has(name)) return
  warnedIcons.add(name)
  const processLike = (globalThis as {
    process?: { env?: Record<string, string | undefined> }
  }).process
  if (processLike?.env?.NODE_ENV === "production") return
  console.warn(
    `[OfficeIcon] Unknown icon "${name}". Use a built-in name, pass the \`icon\` prop, or register it with provideOfficeIcons().`,
  )
}
</script>

<template>
  <component
    :is="resolvedIcon"
    v-if="resolvedIcon"
    class="office-icon"
    :size="attributeSize"
    v-bind="sizeAttrs"
    :stroke-width="strokeWidth"
    :aria-hidden="label ? undefined : 'true'"
    :aria-label="label"
    :role="label ? 'img' : undefined"
    focusable="false"
  />
</template>

<style scoped>
/* Size comes from the token so a host can scale every Office icon at once.
 * No colour is set: Lucide inherits `currentColor` from the control. */
.office-icon {
  display: inline-block;
  block-size: var(--office-icon-size, 1rem);
  flex-shrink: 0;
  inline-size: var(--office-icon-size, 1rem);
  vertical-align: middle;
}
</style>
