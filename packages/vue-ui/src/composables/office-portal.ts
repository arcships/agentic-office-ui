/**
 * Portal target for token-scoped overlays.
 *
 * `--office-*` tokens live under `[data-office-theme]`. Reka UI portals default
 * to `document.body`, which sits outside that subtree, so a portalled dialog
 * would lose every colour, radius and control metric. `OfficeThemeProvider`
 * renders an inert `display: contents` host inside its themed subtree and
 * provides it here; primitives pass it to their `<XPortal :to="...">`.
 *
 * Without a provider the target is `null` and Teleport keeps its default
 * `body` behaviour (correct when the host itself carries `data-office-theme`).
 */

import { inject, provide, type InjectionKey, type Ref } from "vue"

export const OFFICE_PORTAL_TARGET_INJECTION_KEY: InjectionKey<
  Ref<HTMLElement | null>
> = Symbol.for("@arcships/vue-ui.portal-target") as InjectionKey<
  Ref<HTMLElement | null>
>

/** Provide the portal host owned by an `OfficeThemeProvider`. */
export function provideOfficePortalTarget(
  target: Ref<HTMLElement | null>,
): void {
  provide(OFFICE_PORTAL_TARGET_INJECTION_KEY, target)
}

/** Resolve the nearest themed portal host, or `null` for the default body. */
export function useOfficePortalTarget(): Ref<HTMLElement | null> | null {
  return inject(OFFICE_PORTAL_TARGET_INJECTION_KEY, null)
}
