// editor-clipboard.ts — copy/paste operations
//
// Extracted from useDocxEditor.ts.

import type { DocxTextRange } from "@arcships/docx-core"
import { paragraphText } from "@arcships/docx-core"
import type { EditorCore } from "./editor-shared"

export function createEditorClipboard(
  ctx: EditorCore,
  insertTextAtSelection: (text: string, range?: DocxTextRange) => DocxTextRange | undefined,
) {
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

  // ── copy ─────────────────────────────────────────────────────────
  const copy = async (): Promise<void> => {
    const activeRange = ctx.activeTextRange.value
    if (!activeRange) return

    const text = resolveSelectedTextForRange(activeRange)
    if (!text) return

    if (typeof navigator !== "undefined" && navigator.clipboard?.writeText) {
      try {
        await navigator.clipboard.writeText(text)
        ctx.status.value = "Copied to clipboard"
      } catch {
        // Clipboard write failed — silently ignore
      }
    }
  }

  // ── paste ────────────────────────────────────────────────────────
  // Inserts clipboard text at the active caret/selection. Reports success
  // only when the model actually changed.
  const paste = async (): Promise<void> => {
    if (typeof navigator === "undefined" || !navigator.clipboard?.readText) return

    let text = ""
    try {
      text = await navigator.clipboard.readText()
    } catch {
      // Clipboard read failed — silently ignore
      return
    }
    if (!text) return

    const replaced = insertTextAtSelection(text)
    if (!replaced) return

    ctx.status.value = "Pasted from clipboard"
  }

  return { copy, paste, resolveSelectedTextForRange }
}
