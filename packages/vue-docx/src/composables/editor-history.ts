// editor-history.ts — undo/redo with snapshot-based history stack
//
// Extracted from useDocxEditor.ts.

import { cloneEditorSelection, cloneTextRange } from "@extend-ai/docx-core"
import type { EditorCore } from "./editor-shared"

export function createEditorHistory(ctx: EditorCore) {
  function undo(): void {
    const h = ctx.history.value
    if (h.past.length === 0) return

    const previous = h.past[h.past.length - 1]
    ctx.history.value = {
      past: h.past.slice(0, -1),
      future: [
        {
          model: ctx.modelSnapshot.value,
          selection: cloneEditorSelection(ctx.selectionSnapshot.value),
          activeTextRange: cloneTextRange(ctx.activeTextRangeSnapshot.value),
        },
        ...h.future,
      ],
    }
    ctx.model.value = previous.model
    ctx.selection.value = cloneEditorSelection(previous.selection)
    ctx.activeTextRange.value = cloneTextRange(previous.activeTextRange)

    const nextNonce = ctx.historyRestoreNonce.value + 1
    ctx.historyRestoreNonce.value = nextNonce
    ctx.historyRestoreRequest.value = {
      nonce: nextNonce,
      selection: cloneEditorSelection(ctx.selection.value),
      activeTextRange: cloneTextRange(ctx.activeTextRange.value),
    }
  }

  function redo(): void {
    const h = ctx.history.value
    if (h.future.length === 0) return

    const next = h.future[0]
    ctx.history.value = {
      past: [
        ...h.past,
        {
          model: ctx.modelSnapshot.value,
          selection: cloneEditorSelection(ctx.selectionSnapshot.value),
          activeTextRange: cloneTextRange(ctx.activeTextRangeSnapshot.value),
        },
      ],
      future: h.future.slice(1),
    }
    ctx.model.value = next.model
    ctx.selection.value = cloneEditorSelection(next.selection)
    ctx.activeTextRange.value = cloneTextRange(next.activeTextRange)
  }

  return { undo, redo }
}
