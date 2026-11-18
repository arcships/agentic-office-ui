// editor-clipboard.ts — copy/paste operations
//
// Extracted from useDocxEditor.ts.
//
// Copy writes three flavours: the internal structured payload (exact runs +
// source `DocxTextRange`), `text/html` with inline run styles so Word and other
// editors keep bold/links/lists, and `text/plain`. Paste prefers them in that
// order and falls back to the plain-text insert, so a plain-text clipboard and
// a clipboard from an app we cannot parse still paste.

import type {
  DocxClipboardParagraph,
  DocxClipboardSources,
  DocxTextRange
} from "@arcships/docx-core"
import {
  DOCX_CLIPBOARD_MIME,
  compareTextRangeBoundaries,
  createDocxClipboardPayload,
  docxClipboardParagraphsForRange,
  paragraphText,
  parseDocxClipboardHtml,
  parseDocxClipboardPayload,
  parseDocxClipboardRtf,
  renderDocxClipboardHtml,
  renderDocxClipboardRtf,
  serializeDocxClipboardPayload
} from "@arcships/docx-core"
import { sanitizeOfficeLink } from "@arcships/office-runtime/link-policy"
import type { EditorCore } from "./editor-shared"
import {
  resolveDocxClipboardWriteTargets,
  writeDocxClipboardDataSink,
  writeDocxClipboardPayload
} from "./editor-clipboard-io"

export interface EditorClipboardOperations {
  insertTextAtSelection: (text: string, range?: DocxTextRange) => DocxTextRange | undefined
  insertClipboardParagraphsAtSelection: (
    paragraphs: DocxClipboardParagraph[],
    range?: DocxTextRange
  ) => DocxTextRange | undefined
}

interface ClipboardPayloadFlavours {
  internal: string
  html: string
  rtf: string
  text: string
}

export function createEditorClipboard(ctx: EditorCore, operations: EditorClipboardOperations) {
  // ── resolveSelectedTextForRange ──────────────────────────────────
  function resolveSelectedTextForRange(range?: DocxTextRange): string {
    if (!range) return ""
    const model = ctx.model.value
    const startLoc = range.start.location
    const endLoc = range.end.location

    // Single paragraph selection
    if (
      startLoc.kind === "paragraph" &&
      endLoc.kind === "paragraph" &&
      startLoc.nodeIndex === endLoc.nodeIndex
    ) {
      const para = model.nodes[startLoc.nodeIndex]
      if (!para || para.type !== "paragraph") return ""
      const text = paragraphText(para)
      const start = Math.max(0, Math.min(range.start.offset, text.length))
      const end = Math.max(start, Math.min(range.end.offset, text.length))
      return text.slice(start, end)
    }

    // Cross-paragraph selection — collect text from each paragraph in range
    if (startLoc.kind === "paragraph" && endLoc.kind === "paragraph") {
      const parts: string[] = []
      for (let i = startLoc.nodeIndex; i <= endLoc.nodeIndex; i++) {
        const para = model.nodes[i]
        if (!para || para.type !== "paragraph") continue
        const text = paragraphText(para)
        if (i === startLoc.nodeIndex) {
          const start = Math.max(0, Math.min(range.start.offset, text.length))
          parts.push(text.slice(start))
        } else if (i === endLoc.nodeIndex) {
          const end = Math.max(0, Math.min(range.end.offset, text.length))
          parts.push(text.slice(0, end))
        } else {
          parts.push(text)
        }
      }
      return parts.join("\n")
    }

    // Table cell range — collect text from table cell paragraphs
    if (startLoc.kind === "table-cell" && endLoc.kind === "table-cell") {
      const table = model.nodes[startLoc.tableIndex]
      if (!table || table.type !== "table") return ""
      const parts: string[] = []

      // Same cell
      if (
        startLoc.tableIndex === endLoc.tableIndex &&
        startLoc.rowIndex === endLoc.rowIndex &&
        startLoc.cellIndex === endLoc.cellIndex
      ) {
        const cell = table.rows[startLoc.rowIndex]?.cells[startLoc.cellIndex]
        if (!cell) return ""
        for (let pi = startLoc.paragraphIndex; pi <= endLoc.paragraphIndex; pi++) {
          const para = cell.nodes[pi]
          if (!para || para.type !== "paragraph") continue
          if (pi === startLoc.paragraphIndex && pi === endLoc.paragraphIndex) {
            const text = paragraphText(para)
            const start = Math.max(0, Math.min(range.start.offset, text.length))
            const end = Math.max(start, Math.min(range.end.offset, text.length))
            parts.push(text.slice(start, end))
          } else if (pi === startLoc.paragraphIndex) {
            const text = paragraphText(para)
            const start = Math.max(0, Math.min(range.start.offset, text.length))
            parts.push(text.slice(start))
          } else if (pi === endLoc.paragraphIndex) {
            const text = paragraphText(para)
            const end = Math.max(0, Math.min(range.end.offset, text.length))
            parts.push(text.slice(0, end))
          } else {
            parts.push(paragraphText(para))
          }
        }
        return parts.join("\n")
      }
    }

    return ""
  }

  // ── payload builders ─────────────────────────────────────────────
  function richFlavoursForRange(range: DocxTextRange): ClipboardPayloadFlavours | undefined {
    const paragraphs = docxClipboardParagraphsForRange(ctx.model.value, range)
    if (!paragraphs || paragraphs.length === 0) return undefined
    const payload = createDocxClipboardPayload({ paragraphs, range })
    const text = payload.text
    if (!text) return undefined
    return {
      internal: serializeDocxClipboardPayload(payload),
      html: renderDocxClipboardHtml(paragraphs, { sanitizeLink: sanitizeOfficeLink }),
      rtf: renderDocxClipboardRtf(paragraphs, { sanitizeLink: sanitizeOfficeLink }),
      text
    }
  }

  // ── copy ─────────────────────────────────────────────────────────
  const copy = async (): Promise<void> => {
    const activeRange = ctx.activeTextRange.value
    if (!activeRange) return

    const flavours = richFlavoursForRange(activeRange) ?? {
      text: resolveSelectedTextForRange(activeRange)
    }
    if (!flavours.text) return

    const wrote = await writeDocxClipboardPayload(flavours, resolveDocxClipboardWriteTargets())
    if (wrote) {
      ctx.status.value = ctx.statusMessage("copied")
    }
  }

  // ── writeSelectionToClipboardData ────────────────────────────────
  // Copy/cut events must fill `event.clipboardData` synchronously; the async
  // clipboard API cannot be awaited inside the handler.
  const writeSelectionToClipboardData = (data: {
    setData(type: string, value: string): void
  }): boolean => {
    const activeRange = ctx.activeTextRange.value
    if (!activeRange) return false
    if (compareTextRangeBoundaries(activeRange.start, activeRange.end) >= 0) return false

    const flavours = richFlavoursForRange(activeRange) ?? {
      text: resolveSelectedTextForRange(activeRange)
    }
    if (!flavours.text) return false
    return writeDocxClipboardDataSink(flavours, data)
  }

  // ── pasteClipboardData ───────────────────────────────────────────
  // Prefers the internal payload, then `text/html`, then `text/rtf`, then
  // `text/plain`. Reports success only when the model actually changed.
  const pasteClipboardData = (sources: DocxClipboardSources): boolean => {
    const internal = sources.internal
    if (internal) {
      const payload = parseDocxClipboardPayload(internal, {
        sanitizeLink: sanitizeOfficeLink
      })
      if (payload && operations.insertClipboardParagraphsAtSelection(payload.paragraphs)) {
        ctx.status.value = ctx.statusMessage("pasted")
        return true
      }
    }

    if (sources.html) {
      const paragraphs = parseDocxClipboardHtml(sources.html, {
        sanitizeLink: sanitizeOfficeLink
      })
      if (
        paragraphs.length > 0 &&
        operations.insertClipboardParagraphsAtSelection(paragraphs)
      ) {
        ctx.status.value = ctx.statusMessage("pasted")
        return true
      }
    }

    if (sources.rtf) {
      const paragraphs = parseDocxClipboardRtf(sources.rtf, {
        sanitizeLink: sanitizeOfficeLink
      })
      if (
        paragraphs.length > 0 &&
        operations.insertClipboardParagraphsAtSelection(paragraphs)
      ) {
        ctx.status.value = ctx.statusMessage("pasted")
        return true
      }
    }

    if (!sources.text) return false
    if (!operations.insertTextAtSelection(sources.text)) return false
    ctx.status.value = ctx.statusMessage("pasted")
    return true
  }

  // ── paste ────────────────────────────────────────────────────────
  const paste = async (): Promise<void> => {
    const sources = await readClipboardSources()
    if (!sources) return
    pasteClipboardData(sources)
  }

  return {
    copy,
    paste,
    pasteClipboardData,
    writeSelectionToClipboardData,
    resolveSelectedTextForRange
  }
}

interface ClipboardItemLike {
  types: readonly string[]
  getType(type: string): Promise<Blob>
}

/** Reads every flavour we understand from the async clipboard. */
async function readClipboardSources(): Promise<DocxClipboardSources | undefined> {
  if (typeof navigator === "undefined" || !navigator.clipboard) return undefined
  const clipboard = navigator.clipboard

  if (typeof clipboard.read === "function") {
    try {
      const items = (await clipboard.read()) as ClipboardItemLike[]
      const sources: DocxClipboardSources = {}
      for (const item of items) {
        const types = item.types ?? []
        if (!sources.internal && types.includes(DOCX_CLIPBOARD_MIME)) {
          sources.internal = await (await item.getType(DOCX_CLIPBOARD_MIME)).text()
        }
        if (!sources.html && types.includes("text/html")) {
          sources.html = await (await item.getType("text/html")).text()
        }
        if (!sources.rtf && types.includes("text/rtf")) {
          sources.rtf = await (await item.getType("text/rtf")).text()
        }
        if (!sources.text && types.includes("text/plain")) {
          sources.text = await (await item.getType("text/plain")).text()
        }
      }
      if (sources.internal || sources.html || sources.rtf || sources.text) return sources
    } catch {
      // Permission denied or an unreadable item — fall back to readText().
    }
  }

  if (typeof clipboard.readText !== "function") return undefined
  try {
    const text = await clipboard.readText()
    return text ? { text } : undefined
  } catch {
    return undefined
  }
}
