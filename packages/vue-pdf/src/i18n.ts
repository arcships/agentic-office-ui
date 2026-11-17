import { computed, inject, provide, type ComputedRef } from "vue"
import {
  OFFICE_I18N_INJECTION_KEY,
  OFFICE_LOCALE_INJECTION_KEY,
  officeMessagesEn,
  resolveOfficeLocale,
  resolveOfficeMessages,
  type OfficeMessages,
  type OfficeMessagesOverride,
} from "@arcships/office-interaction"

const fallbackMessages: ComputedRef<OfficeMessages> = computed(() => officeMessagesEn)
const fallbackLocale: ComputedRef<string> = computed(() => "en")

/** Resolve the shared message catalog provided by a parent viewer. */
export function usePdfMessages(): ComputedRef<OfficeMessages> {
  return inject(OFFICE_I18N_INJECTION_KEY, fallbackMessages)
}

/** Resolve the active built-in locale provided by a parent viewer. */
export function usePdfLocale(): ComputedRef<string> {
  return inject(OFFICE_LOCALE_INJECTION_KEY, fallbackLocale)
}

/** Provide the resolved catalog to descendant PDF components. */
export function providePdfMessages(
  locale: () => string | undefined,
  messages: () => OfficeMessagesOverride | undefined,
): ComputedRef<OfficeMessages> {
  const resolved = computed(() => resolveOfficeMessages(locale(), messages()))
  provide(OFFICE_I18N_INJECTION_KEY, resolved)
  provide(OFFICE_LOCALE_INJECTION_KEY, computed(() => resolveOfficeLocale(locale())))
  return resolved
}
