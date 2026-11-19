import { sanitizeOfficeLink } from "@arcships/office-runtime/link-policy"
import type { PptxPlaybackDocument, PptxSlideAction } from "@arcships/pptx-core"
import type { PptxActionRequest } from "@arcships/pptx-core/browser"

/**
 * Link routing for the PPTX viewer.
 *
 * The playback parser already resolves `a:hlinkClick` / `a:hlinkHover` into
 * actions; this module decides what those actions mean outside the playback
 * controller. Every URL passes the shared office link policy, so `javascript:`,
 * `data:` and `vbscript:` targets can never reach `window.open`, even when a
 * host supplies its own playback document.
 */

export interface PptxLinkEntry {
  click?: PptxSlideAction
  hover?: PptxSlideAction
}

export type PptxLinkProtocol = "http" | "https" | "mailto"

export type PptxLinkTarget =
  | { kind: "slide"; slideIndex: number }
  | { kind: "url"; url: string; protocol: PptxLinkProtocol }
  /** Missing, unsupported or policy-rejected target: nothing may be opened. */
  | { kind: "unsupported" }

const OPENABLE_URL_PATTERN = /^(https?|mailto):/iu
/** Only these protocols are safe to hand to the browser as a new document. */
const OPENABLE_PROTOCOLS = new Set(["http", "https", "mailto"])
const UNSUPPORTED_TARGET: PptxLinkTarget = Object.freeze({ kind: "unsupported" })

export function pptxLinkKey(slideIndex: number, objectKey: string): string {
  return `${slideIndex}|${objectKey}`
}

/** Index the deck once so a click or a context menu is a map lookup, not a scan. */
export function createPptxLinkIndex(
  document: PptxPlaybackDocument | null | undefined,
): ReadonlyMap<string, PptxLinkEntry> {
  const index = new Map<string, PptxLinkEntry>()
  for (const slide of document?.slides ?? []) {
    for (const action of slide.actions) {
      const key = pptxLinkKey(slide.index, action.sourceObjectKey)
      const entry = index.get(key) ?? {}
      // The parser emits one action per trigger; keep the first of each so a
      // duplicated hlinkClick cannot shadow the real one.
      if (action.trigger === "click") entry.click ??= action
      else entry.hover ??= action
      index.set(key, entry)
    }
  }
  return index
}

export function resolvePptxLinkEntry(
  index: ReadonlyMap<string, PptxLinkEntry>,
  slideIndex: number,
  objectKey: string,
): PptxLinkEntry | undefined {
  return index.get(pptxLinkKey(slideIndex, objectKey))
}

function openableProtocol(value: string): PptxLinkProtocol | undefined {
  const match = /^([a-z][a-z\d+.-]*):/iu.exec(value)
  const protocol = match?.[1].toLowerCase()
  return protocol && OPENABLE_PROTOCOLS.has(protocol)
    ? protocol as PptxLinkProtocol
    : undefined
}

/**
 * Resolve one parsed action into something the viewer may act on.
 *
 * Internal jumps stay inside the deck, external targets must be http(s) or
 * mailto after `sanitizeOfficeLink()`, and everything else is refused instead
 * of being forwarded as a clickable target.
 */
export function resolvePptxLinkTarget(action: PptxSlideAction | undefined): PptxLinkTarget {
  if (!action) return UNSUPPORTED_TARGET
  if (action.kind === "go-to-slide") {
    const index = action.targetSlideIndex
    return typeof index === "number" && Number.isInteger(index) && index >= 0
      ? Object.freeze({ kind: "slide" as const, slideIndex: index })
      : UNSUPPORTED_TARGET
  }
  if (action.kind !== "open-url" && action.kind !== "mailto") return UNSUPPORTED_TARGET
  const url = sanitizeOfficeLink(action.url)
  if (!url || !OPENABLE_URL_PATTERN.test(url)) return UNSUPPORTED_TARGET
  const protocol = openableProtocol(url)
  if (!protocol) return UNSUPPORTED_TARGET
  return Object.freeze({ kind: "url" as const, url, protocol })
}

/** The `action` event a browse-mode click reports to the host. */
export function createPptxActionRequest(
  action: PptxSlideAction,
  target: PptxLinkTarget,
): PptxActionRequest {
  if (target.kind === "slide") {
    return { kind: "go-to-slide", slideIndex: target.slideIndex, sourceObjectKey: action.sourceObjectKey }
  }
  if (target.kind === "url") {
    return target.protocol === "mailto"
      ? { kind: "mailto", url: target.url, sourceObjectKey: action.sourceObjectKey }
      : { kind: "open-url", url: target.url, sourceObjectKey: action.sourceObjectKey }
  }
  return {
    kind: "unsupported",
    // Same payload as the playback controller: never echo a rejected target.
    action: action.rawAction ?? "unknown",
    sourceObjectKey: action.sourceObjectKey,
  }
}

/**
 * Re-check an action the playback controller already classified.
 *
 * The controller gates URLs while parsing, but the viewer also accepts host
 * playback documents; this keeps one policy for both modes and downgrades a
 * malformed or policy-rejected URL to `unsupported` instead of forwarding it.
 */
export function sanitizePptxActionRequest(action: PptxActionRequest): PptxActionRequest {
  if (action.kind !== "open-url" && action.kind !== "mailto") return action
  const url = sanitizeOfficeLink(action.url)
  const protocol = url && OPENABLE_URL_PATTERN.test(url) ? openableProtocol(url) : undefined
  if (!url || !protocol) {
    return { kind: "unsupported", action: "unknown", sourceObjectKey: action.sourceObjectKey }
  }
  if (action.kind === "mailto") {
    return protocol === "mailto"
      ? { kind: "mailto", url, sourceObjectKey: action.sourceObjectKey }
      : { kind: "unsupported", action: "unknown", sourceObjectKey: action.sourceObjectKey }
  }
  return protocol === "mailto"
    ? { kind: "mailto", url, sourceObjectKey: action.sourceObjectKey }
    : { kind: "open-url", url, sourceObjectKey: action.sourceObjectKey }
}

/** Open a policy-approved URL in a new tab without leaking the opener. */
export function openPptxLink(url: string): boolean {
  const opener = (globalThis as {
    open?: (target: string, name?: string, features?: string) => unknown
  }).open
  if (typeof opener !== "function") return false
  opener(url, "_blank", "noopener,noreferrer")
  return true
}

/** Text a host would copy for this target, or `undefined` when there is none. */
export function pptxLinkCopyValue(target: PptxLinkTarget): string | undefined {
  return target.kind === "url" ? target.url : undefined
}
