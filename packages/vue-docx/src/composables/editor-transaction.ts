// editor-transaction.ts — dispatchEditorTransaction + applyModelChange
//
// Extracted from useDocxEditor.ts. Provides the single-entry-point transaction
// dispatcher and a convenience wrapper for model-only changes.

import type { DocModel, DocxEditorTransactionContext, DocxEditorTransactionPatch } from "@extend-ai/docx-core"
import {
  cloneDocModel,
  cloneEditorSelection,
  cloneTextRange,
  cloneTextStyle,
  normalizeEditorCursorStateForModel,
  sameEditorSelection,
  sameTextRange,
} from "@extend-ai/docx-core"
import type { EditorCore } from "./editor-shared"

function hasOwn(obj: object, prop: string): boolean {
  return Object.prototype.hasOwnProperty.call(obj, prop)
}

export function createEditorTransaction(ctx: EditorCore) {
  // ── dispatchEditorTransaction ────────────────────────────────────
  function dispatchEditorTransaction(
    resolver: (txCtx: DocxEditorTransactionContext) => DocxEditorTransactionPatch | undefined
  ): boolean {
    const currentModel = ctx.modelSnapshot.value
    const currentSelection = cloneEditorSelection(ctx.selectionSnapshot.value)
    const currentRange = cloneTextRange(ctx.activeTextRangeSnapshot.value)
    const currentPendingRunStyle = cloneTextStyle(ctx.pendingRunStyleSnapshot.value)

    const patch = resolver({
      model: currentModel,
      selection: currentSelection,
      activeTextRange: currentRange,
      pendingRunStyle: currentPendingRunStyle,
    })
    if (!patch) return false

    const nextModel = patch.model ?? currentModel
    const hasExplicitRange = hasOwn(patch, "activeTextRange")
    const hasExplicitSelection = hasOwn(patch, "selection")
    const hasPendingRunStylePatch = hasOwn(patch, "pendingRunStyle")

    const requestedSelection = patch.selection ?? currentSelection
    const requestedRange = hasExplicitRange ? patch.activeTextRange : currentRange

    const { selection: normSelection, activeTextRange: normRange } =
      normalizeEditorCursorStateForModel(nextModel, requestedSelection, requestedRange)

    const nextSelection = normSelection
    const nextRange = normRange
    const nextPendingRunStyle = hasPendingRunStylePatch
      ? cloneTextStyle(patch.pendingRunStyle)
      : currentPendingRunStyle

    const modelChanged = nextModel !== currentModel
    const selectionChanged = !sameEditorSelection(currentSelection, nextSelection)
    const rangeChanged = !sameTextRange(currentRange, nextRange)
    const pendingRunStyleChanged =
      hasPendingRunStylePatch &&
      JSON.stringify(nextPendingRunStyle ?? null) !== JSON.stringify(currentPendingRunStyle ?? null)

    if (!modelChanged && !selectionChanged && !rangeChanged && !pendingRunStyleChanged && !patch.status && !patch.clearSelectedFormField) {
      return false
    }

    // push history snapshot
    if (modelChanged && patch.pushHistory !== false) {
      ctx.history.value = {
        past: [
          ...ctx.history.value.past.slice(-99),
          {
            model: currentModel,
            selection: cloneEditorSelection(currentSelection),
            activeTextRange: cloneTextRange(currentRange),
          },
        ],
        future: [],
      }
    }

    if (modelChanged) ctx.model.value = nextModel
    if (selectionChanged) {
      ctx.suppressSelectionReset.value = true
      ctx.selection.value = cloneEditorSelection(nextSelection)
    }
    if (rangeChanged) {
      ctx.suppressSelectionReset.value = true
      ctx.activeTextRange.value = cloneTextRange(nextRange)
    }
    if (pendingRunStyleChanged) ctx.pendingRunStyle.value = nextPendingRunStyle

    // DOM selection restore
    const localSessionActive =
      ctx.selectionSessionKindInternal.value === "pointer" ||
      ctx.selectionSessionKindInternal.value === "keyboard" ||
      ctx.selectionSessionKindInternal.value === "composition"
    const shouldRequestRestore =
      selectionChanged || rangeChanged
        ? !localSessionActive || hasExplicitSelection || hasExplicitRange
        : false

    if (shouldRequestRestore) {
      const nextNonce = ctx.historyRestoreNonce.value + 1
      ctx.historyRestoreNonce.value = nextNonce
      ctx.historyRestoreRequest.value = {
        nonce: nextNonce,
        selection: cloneEditorSelection(nextSelection),
        activeTextRange: cloneTextRange(nextRange),
      }
    }

    if (patch.clearSelectedFormField) ctx.selectedFormFieldLocation.value = undefined
    if (patch.status) ctx.status.value = patch.status

    return true
  }

  // ── applyModelChange ─────────────────────────────────────────────
  function applyModelChange(updater: (current: DocModel) => DocModel, successStatus?: string): void {
    dispatchEditorTransaction((current) => {
      const nextModel = updater(current.model)
      if (nextModel === current.model) return undefined
      return { model: nextModel, status: successStatus, clearSelectedFormField: true }
    })
  }

  return { dispatchEditorTransaction, applyModelChange }
}
