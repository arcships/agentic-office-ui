import { unref, type Ref } from "vue"

/**
 * Shared keyboard map for the DOCX, PDF and PPTX viewers.
 *
 * Every viewer attaches the returned `onKeydown` handler to its component
 * root, so a shortcut only fires while focus is inside that viewer. The
 * resolver refuses editable targets (inputs, textareas, selects and
 * contenteditable hosts) before it looks at a key, which keeps toolbar
 * fields and form controls in charge of their own keyboard behaviour.
 *
 * The map itself is intentionally small and identical across formats:
 *
 * | Keys                         | Command                                  |
 * | ---------------------------- | ---------------------------------------- |
 * | PageDown / PageUp            | page-next / page-previous                |
 * | Home / End                   | page-first / page-last                   |
 * | Ctrl/Cmd + `+` or `=`        | zoom-in                                  |
 * | Ctrl/Cmd + `-`               | zoom-out                                 |
 * | Ctrl/Cmd + `0`               | zoom-reset                               |
 * | Ctrl/Cmd + `F`               | search                                   |
 * | Escape                       | escape                                   |
 * | Present mode: Arrow keys, Space | page-previous / page-next             |
 *
 * Present mode is the only mode that claims the arrow keys and Space; browse
 * mode leaves them to the host so a scrollable slide/PDF/DOCX list keeps its
 * native scrolling. `pageNavigation` can be turned off per host when PageUp /
 * PageDown / Home / End should scroll instead of changing pages.
 */
export type ViewerKeyCommand =
  | "page-previous"
  | "page-next"
  | "page-first"
  | "page-last"
  | "zoom-in"
  | "zoom-out"
  | "zoom-reset"
  | "search"
  | "escape"

export type ViewerKeyMode = "browse" | "present"

/** The subset of `KeyboardEvent` the shared map reads, so it stays testable. */
export interface ViewerKeyEvent {
  key: string
  ctrlKey?: boolean
  metaKey?: boolean
  altKey?: boolean
  shiftKey?: boolean
  defaultPrevented?: boolean
  target?: EventTarget | null
  preventDefault?: () => void
}

export interface ViewerKeyMapOptions {
  /** PageUp / PageDown / Home / End navigate by page. Defaults to false. */
  pageNavigation?: boolean
  /** Ctrl/Cmd with `+`, `=`, `-` or `0` changes zoom. Defaults to false. */
  zoom?: boolean
  /** Ctrl/Cmd + F requests the search field. Defaults to false. */
  search?: boolean
  /** Escape requests search or selection cancellation. Defaults to false. */
  escape?: boolean
  /**
   * `present` also maps arrows and Space to page navigation. `browse` is the
   * default and leaves those keys to the browser/host.
   */
  mode?: ViewerKeyMode
}

/** Documented map, exported so guides and tests can assert the same contract. */
export const VIEWER_KEYBOARD_SHORTCUTS: ReadonlyArray<{
  keys: string
  command: ViewerKeyCommand
  scope: "all" | "present"
}> = Object.freeze([
  { keys: "PageDown", command: "page-next", scope: "all" },
  { keys: "PageUp", command: "page-previous", scope: "all" },
  { keys: "Home", command: "page-first", scope: "all" },
  { keys: "End", command: "page-last", scope: "all" },
  { keys: "Ctrl/Cmd + + or =", command: "zoom-in", scope: "all" },
  { keys: "Ctrl/Cmd + -", command: "zoom-out", scope: "all" },
  { keys: "Ctrl/Cmd + 0", command: "zoom-reset", scope: "all" },
  { keys: "Ctrl/Cmd + F", command: "search", scope: "all" },
  { keys: "Escape", command: "escape", scope: "all" },
  { keys: "ArrowDown / ArrowRight / Space", command: "page-next", scope: "present" },
  { keys: "ArrowUp / ArrowLeft", command: "page-previous", scope: "present" },
])

const INTERACTIVE_SELECTOR =
  "button, a[href], [role='button'], [role='link'], [role='checkbox'], [role='radio'], [role='switch'], [role='menuitem'], summary"

/**
 * True when the key should stay with a form field or editable content.
 * The viewer roots keep their navigation keys away from toolbar fields and
 * from embedded form controls (PDF AcroForm inputs, DOCX text editors).
 */
export function isEditableViewerTarget(target: unknown): boolean {
  if (!target || typeof target !== "object") return false
  const element = target as {
    tagName?: unknown
    isContentEditable?: unknown
    closest?: (selector: string) => unknown
  }
  if (typeof element.closest === "function") {
    try {
      if (element.closest("[contenteditable]:not([contenteditable='false'])")) return true
    } catch {
      // A detached or synthetic target may not be a real DOM node.
    }
  }
  if (element.isContentEditable === true) return true
  // Real DOM nodes expose `tagName`; lightweight renderers and synthetic
  // targets often expose only `type`.
  const name = typeof element.tagName === "string"
    ? element.tagName
    : typeof (element as { type?: unknown }).type === "string"
      ? (element as { type: string }).type
      : ""
  const tag = name.toLowerCase()
  return tag === "input" || tag === "textarea" || tag === "select"
}

/**
 * True for controls that own Space/Arrow activation. Presentation mode maps
 * Space to "next slide", but a focused playback button must still be usable.
 */
export function isInteractiveViewerTarget(target: unknown): boolean {
  if (!target || typeof target !== "object") return false
  const element = target as { closest?: (selector: string) => unknown }
  if (typeof element.closest !== "function") return false
  try {
    return Boolean(element.closest(INTERACTIVE_SELECTOR))
  } catch {
    return false
  }
}

/** Resolve one key event against the shared map. Pure; safe to unit test. */
export function resolveViewerKeyCommand(
  event: ViewerKeyEvent,
  options: ViewerKeyMapOptions = {},
): ViewerKeyCommand | null {
  if (!event || typeof event.key !== "string" || event.defaultPrevented) return null

  if (event.key === "Escape") return options.escape ? "escape" : null

  const modified = Boolean(event.ctrlKey || event.metaKey)
  if (modified) {
    const key = event.key.toLowerCase()
    if (key === "f") return options.search ? "search" : null
    if (!options.zoom) return null
    if (event.key === "+" || event.key === "=" || key === "add") return "zoom-in"
    if (event.key === "-" || event.key === "_" || key === "subtract") return "zoom-out"
    if (event.key === "0" || key === "num0") return "zoom-reset"
    return null
  }

  // Alt/Shift + navigation extends a selection or opens a menu; leave those
  // combinations to the document instead of changing pages underneath them.
  if (event.altKey || event.shiftKey) return null

  if (options.pageNavigation) {
    switch (event.key) {
      case "PageDown":
        return "page-next"
      case "PageUp":
        return "page-previous"
      case "Home":
        return "page-first"
      case "End":
        return "page-last"
      default:
        break
    }
  }

  if (options.mode === "present") {
    if (event.key === " " || event.key === "Spacebar" || event.key === "Space") {
      // A focused playback button owns Space; the viewer must not advance a slide.
      if (isInteractiveViewerTarget(event.target)) return null
      return "page-next"
    }
    switch (event.key) {
      case "ArrowDown":
      case "ArrowRight":
        return "page-next"
      case "ArrowUp":
      case "ArrowLeft":
        return "page-previous"
      default:
        break
    }
  }

  return null
}

export interface ViewerKeyOptions {
  /** Reactive switches; `true` enables the matching part of the map. */
  pageNavigation?: ViewerKeySwitch
  zoom?: ViewerKeySwitch
  search?: ViewerKeySwitch
  escape?: ViewerKeySwitch
  mode?: ViewerKeyMode | Ref<ViewerKeyMode> | (() => ViewerKeyMode)
  /**
   * Applies the command. Return `false` when the host could not handle it
   * (for example already on the first page) so the browser keeps the key.
   */
  onCommand: (command: ViewerKeyCommand, event: ViewerKeyEvent) => boolean | void
}

/** A plain value, a Vue ref or a getter, so hosts can bind props reactively. */
export type ViewerKeySwitch = boolean | Ref<boolean> | (() => boolean)

export interface ViewerKeyboard {
  /** Bind to the viewer root: `<div @keydown="keyboard.onKeydown">`. */
  onKeydown: (event: ViewerKeyEvent) => boolean
}

/**
 * Shared viewer keyboard composable. The host only supplies command handlers;
 * the scope, editable-target guard and modifier rules stay identical between
 * DOCX, PDF and PPTX.
 */
export function useViewerKeys(options: ViewerKeyOptions): ViewerKeyboard {
  function onKeydown(event: ViewerKeyEvent): boolean {
    if (!event || typeof event.key !== "string") return false
    if (isEditableViewerTarget(event.target)) return false

    const command = resolveViewerKeyCommand(event, {
      pageNavigation: resolveSwitch(options.pageNavigation),
      zoom: resolveSwitch(options.zoom),
      search: resolveSwitch(options.search),
      escape: resolveSwitch(options.escape),
      mode: resolveMode(options.mode),
    })
    if (!command) return false

    const handled = options.onCommand(command, event) !== false
    if (handled) event.preventDefault?.()
    return handled
  }

  return { onKeydown }
}

function resolveSwitch(value: ViewerKeySwitch | undefined): boolean {
  if (typeof value === "function") return Boolean(value())
  return Boolean(unref(value))
}

function resolveMode(value: ViewerKeyOptions["mode"]): ViewerKeyMode {
  if (typeof value === "function") return value()
  return unref(value) ?? "browse"
}
