/** Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. */
/**
 * Localized labels for the primitive layer.
 *
 * Generic control labels (dialog close, menu, notifications) live in the
 * shared `vueUi.primitives` section of the `@arcships/office-interaction`
 * message catalog. Primitives therefore follow the active Office locale
 * (`en` by default) and host message overrides, so a `zh-CN` provider stays
 * Chinese. Hosts can override any label through the component prop.
 */

import { computed, type ComputedRef } from "vue"
import { useOfficeUiMessages } from "../../i18n"

export interface OfficeUiPrimitiveLabels {
  close: string
  loading: string
  menu: string
  notifications: string
  dialog: string
  more: string
}

/** Resolve the primitive label set for the active Office locale. */
export function useOfficeUiLabels(): ComputedRef<OfficeUiPrimitiveLabels> {
  const messages = useOfficeUiMessages()
  return computed(() => messages.value.vueUi.primitives)
}
