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
    const restoredSelection = cloneEditorSelection(previous.selection)
    const restoredRange = cloneTextRange(previous.activeTextRange)
    ctx.suppressNextDomSelectionRestore.value = true
    ctx.selection.value = restoredSelection
    ctx.activeTextRange.value = restoredRange
    ctx.selectionSnapshot.value = cloneEditorSelection(restoredSelection)
    ctx.activeTextRangeSnapshot.value = cloneTextRange(restoredRange)
    ctx.model.value = previous.model
    ctx.modelSnapshot.value = previous.model

    const nextNonce = ctx.historyRestoreNonce.value + 1
    ctx.historyRestoreNonce.value = nextNonce
    ctx.historyRestoreRequest.value = {
      nonce: nextNonce,
      selection: cloneEditorSelection(restoredSelection),
      activeTextRange: cloneTextRange(restoredRange),
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
    const restoredSelection = cloneEditorSelection(next.selection)
    const restoredRange = cloneTextRange(next.activeTextRange)
    ctx.suppressNextDomSelectionRestore.value = true
    ctx.selection.value = restoredSelection
    ctx.activeTextRange.value = restoredRange
    ctx.selectionSnapshot.value = cloneEditorSelection(restoredSelection)
    ctx.activeTextRangeSnapshot.value = cloneTextRange(restoredRange)
    ctx.model.value = next.model
    ctx.modelSnapshot.value = next.model

    const nextNonce = ctx.historyRestoreNonce.value + 1
    ctx.historyRestoreNonce.value = nextNonce
    ctx.historyRestoreRequest.value = {
      nonce: nextNonce,
      selection: cloneEditorSelection(restoredSelection),
      activeTextRange: cloneTextRange(restoredRange),
    }
  }

  return { undo, redo }
}
