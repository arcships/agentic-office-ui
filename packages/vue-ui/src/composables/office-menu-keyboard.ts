/**
 * Shared WAI-ARIA menu keyboard semantics for the Office context menus.
 *
 * Every Office surface (DOCX, XLSX, …) used to decide independently whether a
 * context menu closes on Escape, so the same gesture behaved differently per
 * format. Both menus now delegate keyboard handling to the helpers below:
 *
 * - the menu takes focus when it opens,
 * - arrow/home/end keys cycle through its items,
 * - Escape and Tab close it and return focus to the invoking surface.
 */

import { getCurrentInstance, onBeforeUnmount } from "vue"

export const OFFICE_MENU_ITEM_SELECTOR =
  '[role="menuitem"]:not([disabled]):not([aria-disabled="true"])'

export type OfficeMenuKeyboardCommand =
  | { type: "focus"; index: number }
  | { type: "close"; restoreFocus: boolean }

/**
 * Resolve the menu command for a key press.
 *
 * `currentIndex` is the index of the focused item, or `-1` when focus is not on
 * an item yet. Returns `null` for keys the menu does not own.
 */
export function resolveOfficeMenuKeyCommand(
  key: string,
  itemCount: number,
  currentIndex: number,
): OfficeMenuKeyboardCommand | null {
  if (key === "Escape") return { type: "close", restoreFocus: true }
  if (key === "Tab") return { type: "close", restoreFocus: true }
  if (itemCount <= 0) return null
  const index = currentIndex >= 0 && currentIndex < itemCount ? currentIndex : -1
  switch (key) {
    case "ArrowDown":
      return { type: "focus", index: index < 0 ? 0 : (index + 1) % itemCount }
    case "ArrowUp":
      return { type: "focus", index: index < 0 ? itemCount - 1 : (index - 1 + itemCount) % itemCount }
    case "Home":
      return { type: "focus", index: 0 }
    case "End":
      return { type: "focus", index: itemCount - 1 }
    default:
      return null
  }
}

function activeElement(): HTMLElement | null {
  if (typeof document === "undefined") return null
  const element = document.activeElement
  if (!element || typeof (element as HTMLElement).focus !== "function") return null
  return element as HTMLElement
}

export interface OfficeMenuKeyboardOptions {
  /** The rendered menu root, or `null` while the menu is closed. */
  menu: () => HTMLElement | null | undefined
  /** Called when the menu must close. */
  close: () => void
  /**
   * Element that receives focus after the menu closes. Defaults to the element
   * that was focused when the menu opened.
   */
  restoreTarget?: () => HTMLElement | null | undefined
  itemSelector?: string
  /**
   * Also close on Escape from outside the menu. Keeps the gesture working if
   * focus never moved into the menu. Defaults to `true`.
   */
  windowFallback?: boolean
}

export interface OfficeMenuKeyboard {
  /** Remember the invoking element; call right before the menu becomes visible. */
  captureInvoker: () => void
  /** Focus the first enabled item; call after the menu is rendered. */
  focusFirst: () => void
  /** Handle a keydown on the menu root. */
  onKeydown: (event: KeyboardEvent) => void
  /** Focus the element that opened the menu. */
  restoreFocus: () => void
}

export function useOfficeMenuKeyboard(options: OfficeMenuKeyboardOptions): OfficeMenuKeyboard {
  const itemSelector = options.itemSelector ?? OFFICE_MENU_ITEM_SELECTOR
  let invoker: HTMLElement | null = null

  function menuElement(): HTMLElement | null {
    return options.menu() ?? null
  }

  function items(): HTMLElement[] {
    const menu = menuElement()
    if (!menu || typeof menu.querySelectorAll !== "function") return []
    return Array.from(menu.querySelectorAll(itemSelector)) as HTMLElement[]
  }

  function captureInvoker(): void {
    invoker = activeElement()
  }

  function focusIndex(index: number): void {
    const available = items()
    if (!available.length) return
    const target = available[Math.min(Math.max(index, 0), available.length - 1)]
    target?.focus?.()
  }

  function focusFirst(): void {
    focusIndex(0)
  }

  function restoreFocus(): void {
    const explicit = options.restoreTarget?.() ?? null
    const target = explicit ?? invoker
    invoker = null
    target?.focus?.()
  }

  function currentIndex(): number {
    const focused = activeElement()
    if (!focused) return -1
    return items().indexOf(focused)
  }

  function closeAndRestore(): void {
    options.close()
    restoreFocus()
  }

  function onKeydown(event: KeyboardEvent): void {
    if (event.defaultPrevented) return
    const command = resolveOfficeMenuKeyCommand(event.key, items().length, currentIndex())
    if (!command) return
    event.preventDefault?.()
    event.stopPropagation?.()
    if (command.type === "focus") {
      focusIndex(command.index)
      return
    }
    closeAndRestore()
  }

  function onWindowKeydown(event: KeyboardEvent): void {
    if (event.key !== "Escape" || event.defaultPrevented) return
    const menu = menuElement()
    if (!menu) return
    const focused = activeElement()
    if (focused && typeof menu.contains === "function" && menu.contains(focused)) return
    event.preventDefault?.()
    closeAndRestore()
  }

  if (options.windowFallback !== false && typeof window !== "undefined") {
    window.addEventListener("keydown", onWindowKeydown)
    if (getCurrentInstance()) {
      onBeforeUnmount(() => window.removeEventListener("keydown", onWindowKeydown))
    }
  }

  return { captureInvoker, focusFirst, onKeydown, restoreFocus }
}
