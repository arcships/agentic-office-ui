<script setup lang="ts">
import {
  computed,
  onBeforeUnmount,
  onMounted,
  ref,
  watch,
  type CSSProperties,
} from "vue"
import { useElementWidth } from "../../composables/element-width"
import { useOfficeUiMessages } from "../../i18n"
import { officeClass, toCssLength } from "../ui/types"
import DocumentViewerSidebarSkeleton from "./DocumentViewerSidebarSkeleton.vue"
import type {
  OfficeSidebarMode,
  OfficeSidebarResolvedMode,
} from "./types"

export interface DocumentViewerSidebarProps {
  /** Controlled open state (`v-model:open`). */
  open?: boolean
  /** Initial open state for uncontrolled use. */
  defaultOpen?: boolean
  /**
   * Layout strategy. `auto` (default) switches to `overlay` when the measured
   * container is narrower than `breakpoint`; `inline` and `overlay` force a
   * layout.
   */
  mode?: OfficeSidebarMode
  /** Container width at which `auto` becomes inline. Defaults to 768px. */
  breakpoint?: number
  /** Sidebar width. Defaults to 160px. */
  width?: number | string
  /** Render the built-in thumbnail skeleton instead of the default slot. */
  loading?: boolean
  /** Accessible name. Defaults to `common.chrome.sidebar`. */
  label?: string
  /** Close an open overlay with Escape. */
  closeOnEscape?: boolean
  /**
   * Element measured in `auto` mode. Defaults to the sidebar's parent, which
   * is normally the viewer shell that also hosts the document surface.
   */
  container?: HTMLElement | null
  /** Extra classes merged on the root. */
  class?: string
}

const props = withDefaults(defineProps<DocumentViewerSidebarProps>(), {
  open: undefined,
  defaultOpen: true,
  mode: "auto",
  breakpoint: 768,
  width: 160,
  loading: false,
  label: undefined,
  closeOnEscape: true,
  container: null,
  class: "",
})

const emit = defineEmits<{
  "update:open": [open: boolean]
  "open-change": [open: boolean]
  "mode-change": [mode: OfficeSidebarResolvedMode]
}>()

const m = useOfficeUiMessages()

const root = ref<HTMLElement | null>(null)
const measureTarget = ref<HTMLElement | null>(null)
const viewportInline = ref(true)
const internalOpen = ref(props.defaultOpen)
const transitionsReady = ref(false)
let mounted = false
let firstFrame = 0
let secondFrame = 0

const measuredWidth = useElementWidth(measureTarget)

const requestedMode = computed(() => props.mode)
const inline = computed(() => {
  if (requestedMode.value === "inline") return true
  if (requestedMode.value === "overlay") return false
  if (measuredWidth.value > 0) return measuredWidth.value >= props.breakpoint
  return viewportInline.value
})
const resolvedMode = computed<OfficeSidebarResolvedMode>(() =>
  inline.value ? "inline" : "overlay",
)
const isOpen = computed(() => props.open ?? internalOpen.value)
const classes = computed(() =>
  officeClass(
    "office-ui",
    "office-document-sidebar",
    `office-document-sidebar--${resolvedMode.value}`,
    props.class,
  ),
)
const style = computed(
  () =>
    ({
      "--office-ui-sidebar-width": toCssLength(props.width),
    }) as CSSProperties,
)

watch(resolvedMode, (mode) => {
  if (mounted) emit("mode-change", mode)
})
watch(isOpen, (open) => {
  if (mounted) emit("open-change", open)
})

function measureContainer(): void {
  measureTarget.value = props.container ?? root.value?.parentElement ?? null
}

function requestClose(): void {
  if (!isOpen.value) return
  if (props.open === undefined) internalOpen.value = false
  emit("update:open", false)
}

function onEscape(event: KeyboardEvent): void {
  if (!props.closeOnEscape || !isOpen.value || event.defaultPrevented) return
  if (event.key !== "Escape") return
  event.preventDefault()
  requestClose()
}

function onDocumentKeydown(event: KeyboardEvent): void {
  if (resolvedMode.value !== "overlay") return
  onEscape(event)
}

onMounted(() => {
  mounted = true
  if (typeof window !== "undefined" && typeof window.matchMedia === "function") {
    viewportInline.value = window.matchMedia(
      `(min-width: ${props.breakpoint}px)`,
    ).matches
  }
  measureContainer()
  if (typeof document !== "undefined" && typeof document.addEventListener === "function") {
    document.addEventListener("keydown", onDocumentKeydown)
  }
  if (typeof requestAnimationFrame !== "function") {
    transitionsReady.value = true
    return
  }
  // Two frames: the first paints the initial open/closed state without a
  // transition, the second enables the animation for later toggles.
  firstFrame = requestAnimationFrame(() => {
    secondFrame = requestAnimationFrame(() => {
      transitionsReady.value = true
    })
  })
})

onBeforeUnmount(() => {
  mounted = false
  if (typeof document !== "undefined" && typeof document.removeEventListener === "function") {
    document.removeEventListener("keydown", onDocumentKeydown)
  }
  if (typeof cancelAnimationFrame === "function") {
    if (firstFrame) cancelAnimationFrame(firstFrame)
    if (secondFrame) cancelAnimationFrame(secondFrame)
  }
})

watch(
  () => props.container,
  () => measureContainer(),
  { flush: "post" },
)

defineExpose({
  inline,
  mode: resolvedMode,
  open: isOpen,
  close: requestClose,
})
</script>

<template>
  <aside
    ref="root"
    :class="classes"
    :style="style"
    :aria-label="props.label ?? m.common.chrome.sidebar"
    :aria-hidden="isOpen ? undefined : 'true'"
    :inert="isOpen ? undefined : true"
    data-document-sidebar=""
    :data-sidebar-mode="resolvedMode"
    :data-sidebar-open="isOpen ? 'true' : 'false'"
    :data-sidebar-inline="inline ? 'true' : 'false'"
    :data-transitions-ready="transitionsReady ? 'true' : 'false'"
    data-slot="document-sidebar"
    data-testid="document-viewer-sidebar"
    @keydown="onEscape"
  >
    <header v-if="$slots.header" class="office-document-sidebar__header">
      <slot name="header" />
    </header>
    <div class="office-document-sidebar__body">
      <DocumentViewerSidebarSkeleton v-if="props.loading" />
      <slot v-else />
    </div>
    <footer v-if="$slots.footer" class="office-document-sidebar__footer">
      <slot name="footer" />
    </footer>
  </aside>
</template>
