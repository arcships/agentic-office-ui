/** Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. */
/**
 * Localized labels for the primitive layer.
 *
 * The shared `@arcships/office-interaction` message catalog does not own
 * generic control labels (dialog close, menu, notifications). Rather than
 * inventing new catalog keys in a UI-package issue, primitives resolve these
 * few strings from the active Office locale, so a `zh-CN` provider stays
 * Chinese. Hosts can override any label through the component prop.
 */

import { computed, type ComputedRef } from "vue"
import { useOfficeUiLocale } from "../../i18n"

export interface OfficeUiPrimitiveLabels {
  close: string
  loading: string
  menu: string
  notifications: string
  dialog: string
  more: string
}

const LABELS: Record<"en" | "zh-CN", OfficeUiPrimitiveLabels> = {
  en: {
    close: "Close",
    loading: "Loading",
    menu: "Menu",
    notifications: "Notifications",
    dialog: "Dialog",
    more: "More",
  },
  "zh-CN": {
    close: "关闭",
    loading: "加载中",
    menu: "菜单",
    notifications: "通知",
    dialog: "对话框",
    more: "更多",
  },
}

/** Resolve the primitive label set for the active Office locale. */
export function useOfficeUiLabels(): ComputedRef<OfficeUiPrimitiveLabels> {
  const locale = useOfficeUiLocale()
  return computed(() => (locale.value === "zh-CN" ? LABELS["zh-CN"] : LABELS.en))
}
