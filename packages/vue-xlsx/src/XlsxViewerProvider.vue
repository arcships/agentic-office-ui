<template>
  <slot />
</template>

<script setup lang="ts">
import { provide } from "vue"
import type { XlsxViewerController } from "./composables"
import { useXlsxViewerController, XLSX_VIEWER_KEY, XLSX_VIEWER_DARK_KEY } from "./composables"
import type { UseXlsxViewerControllerOptions } from "./composables"

export interface XlsxViewerProviderProps extends UseXlsxViewerControllerOptions {
  /** Existing controller to provide instead of creating one from options. */
  controller?: XlsxViewerController
  /** Uses the built-in dark worksheet/viewer palette. */
  isDark?: boolean
}

const props = withDefaults(defineProps<XlsxViewerProviderProps>(), {
  isDark: false,
})

const { controller: externalController, isDark, ...options } = props

provide(XLSX_VIEWER_DARK_KEY, { isDark: props.isDark })

if (externalController) {
  provide(XLSX_VIEWER_KEY, externalController)
} else {
  const internalController = useXlsxViewerController(options)
  provide(XLSX_VIEWER_KEY, internalController)
}
</script>
