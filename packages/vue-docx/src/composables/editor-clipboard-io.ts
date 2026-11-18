// editor-clipboard-io.ts — clipboard write ladder for the rich DOCX payload.
//
// Mirrors the XLSX clipboard: the composite ClipboardItem write is
// all-or-nothing (Firefox rejects custom MIME types, and some hosts deny the
// write permission), so a rejected rich write still falls back to `writeText`
// and the user keeps a plain-text copy instead of losing the copy entirely.
//
// Targets are injected so the ladder is testable without a browser.

import type { DocxClipboardDataSink } from "@arcships/docx-core"
import {
  DOCX_CLIPBOARD_MIME,
  MAX_DOCX_CLIPBOARD_HTML_CHARS,
  MAX_DOCX_CLIPBOARD_PAYLOAD_CHARS,
  MAX_DOCX_CLIPBOARD_RTF_CHARS
} from "@arcships/docx-core"

export interface DocxClipboardWriteTarget {
  write?(items: unknown[]): Promise<void>
  writeText?(text: string): Promise<void>
}

export interface DocxClipboardWriteTargets {
  clipboard?: DocxClipboardWriteTarget | null
  /** `ClipboardItem` constructor; absent in browsers without the async API. */
  ClipboardItem?: (new (items: Record<string, Blob>) => unknown) | null
  /** `Blob` constructor, injected for the same reason. */
  Blob?: typeof Blob
}

export interface DocxClipboardWritePayload {
  /** Internal structured flavour (already serialized). */
  internal?: string
  /** `text/html` flavour. */
  html?: string
  /** `text/rtf` flavour. */
  rtf?: string
  /** `text/plain` flavour; always writable. */
  text: string
}

/** Resolves the browser clipboard handles; absent ones disable that flavour. */
export function resolveDocxClipboardWriteTargets(): DocxClipboardWriteTargets {
  return {
    clipboard: typeof navigator === "undefined" ? undefined : navigator.clipboard,
    ClipboardItem: typeof ClipboardItem === "function" ? ClipboardItem : undefined,
    Blob: typeof Blob === "function" ? Blob : undefined
  }
}

function flavoursWithinCaps(payload: DocxClipboardWritePayload): {
  internal?: string
  html?: string
  rtf?: string
} {
  const internal =
    payload.internal && payload.internal.length <= MAX_DOCX_CLIPBOARD_PAYLOAD_CHARS
      ? payload.internal
      : undefined
  const html =
    payload.html && payload.html.length <= MAX_DOCX_CLIPBOARD_HTML_CHARS
      ? payload.html
      : undefined
  const rtf =
    payload.rtf && payload.rtf.length <= MAX_DOCX_CLIPBOARD_RTF_CHARS
      ? payload.rtf
      : undefined
  return { internal, html, rtf }
}

/**
 * Writes the flavours to the async clipboard, returning whether anything
 * reached it.
 *
 * Browsers disagree on custom MIME types inside a `ClipboardItem` (Firefox
 * historically rejects the whole item), so a rejected rich write is retried
 * with the standard `text/html` + `text/plain` pair before falling back to
 * plain text. The copy keeps its formatting where the platform allows it.
 */
export async function writeDocxClipboardPayload(
  payload: DocxClipboardWritePayload,
  targets: DocxClipboardWriteTargets
): Promise<boolean> {
  const { clipboard } = targets
  if (!clipboard) return false
  const { internal, html, rtf } = flavoursWithinCaps(payload)

  const writeItem = async (items: Record<string, Blob>): Promise<boolean> => {
    if (!targets.ClipboardItem || !clipboard.write) return false
    try {
      await clipboard.write([new targets.ClipboardItem(items)])
      return true
    } catch {
      return false
    }
  }

  if (targets.Blob && targets.ClipboardItem) {
    const blob = targets.Blob
    if (internal) {
      const full: Record<string, Blob> = {}
      full[DOCX_CLIPBOARD_MIME] = new blob([internal], { type: DOCX_CLIPBOARD_MIME })
      if (html) full["text/html"] = new blob([html], { type: "text/html" })
      if (rtf) full["text/rtf"] = new blob([rtf], { type: "text/rtf" })
      full["text/plain"] = new blob([payload.text], { type: "text/plain" })
      if (await writeItem(full)) return true
    }

    if (html || rtf) {
      const standard: Record<string, Blob> = {}
      if (html) standard["text/html"] = new blob([html], { type: "text/html" })
      if (rtf) standard["text/rtf"] = new blob([rtf], { type: "text/rtf" })
      standard["text/plain"] = new blob([payload.text], { type: "text/plain" })
      if (await writeItem(standard)) return true
    }
  }

  if (!clipboard.writeText) return false
  try {
    await clipboard.writeText(payload.text)
    return true
  } catch {
    return false
  }
}

/**
 * Writes the flavours into a copy/cut event's clipboard sink.
 *
 * Synchronous by design: `event.clipboardData` is only writable while the event
 * is dispatched. Returns false when no flavour could be written.
 */
export function writeDocxClipboardDataSink(
  payload: DocxClipboardWritePayload,
  sink: DocxClipboardDataSink
): boolean {
  const { internal, html, rtf } = flavoursWithinCaps(payload)
  try {
    if (internal) sink.setData(DOCX_CLIPBOARD_MIME, internal)
    if (html) sink.setData("text/html", html)
    if (rtf) sink.setData("text/rtf", rtf)
    sink.setData("text/plain", payload.text)
    return true
  } catch {
    return false
  }
}
