// editor-transaction.ts — dispatchEditorTransaction + applyModelChange
//
// Extracted from useDocxEditor.ts. Provides the single-entry-point transaction
// dispatcher and a convenience wrapper for model-only changes.

import type {
  DocModel,
  DocxEditorTransactionContext,
  DocxEditorTransactionPatch,
  DocxTextRange,
} from "@arcships/docx-core"
import {
  cloneDocModel,
  cloneEditorSelection,
  cloneTextRange,
  cloneTextStyle,
  countParagraphsLosingRevisionMarkup,
  normalizeEditorCursorStateForModel,
  sameEditorSelection,
  sameTextRange,
} from "@arcships/docx-core"
import type { ApplyModelChangeOptions, EditorCore } from "./editor-shared"
import { trimHistoryEntriesToBudget } from "./history-budget"

function hasOwn(obj: object, prop: string): boolean {
  return Object.prototype.hasOwnProperty.call(obj, prop)
}

function revisionLossStatus(lostParagraphs: number): string {
  return lostParagraphs === 1
    ? "Tracked changes in an edited paragraph were resolved with the edit"
    : `Tracked changes in ${lostParagraphs} edited paragraphs were resolved with the edit`
}

export function createEditorTransaction(ctx: EditorCore) {
  // Text commits that share a key and have no intervening transaction share one
  // undo entry: typing a word is one undo step, not one step per keystroke.
  let openTextCommitKey: string | undefined

  function endTextCommitBurst(): void {
    openTextCommitKey = undefined
  }

  /**
   * Key of the paragraph/cell that owns a caret range, using the same format as
   * the text commands. Returns `undefined` for section (header/footer) ranges,
   * which have no body location to compare.
   */
  function textCommitKeyForRange(range?: DocxTextRange): string | undefined {
    if (!range) return undefined
    const start = range.start.location
    const end = range.end.location
    if (start.kind !== end.kind) return undefined
    if (start.kind === "paragraph" && end.kind === "paragraph") {
      if (start.nodeIndex !== end.nodeIndex) return undefined
      return `paragraph:${start.nodeIndex}`
    }
    if (start.kind === "table-cell" && end.kind === "table-cell") {
      if (
        start.tableIndex !== end.tableIndex ||
        start.rowIndex !== end.rowIndex ||
        start.cellIndex !== end.cellIndex
      ) {
        return undefined
      }
      return `table-cell:${start.tableIndex}:${start.rowIndex}:${start.cellIndex}:${
        start.paragraphIndex ?? 0
      }`
    }
    return undefined
  }

  /**
   * True when a caret/selection update stays inside the open text-commit burst's
   * own paragraph or cell and is collapsed. Typing moves the caret inside that
   * target, so those updates must not split the undo entry; jumping elsewhere or
   * selecting a range does.
   */
  function keepsTextCommitBurst(range?: DocxTextRange): boolean {
    if (openTextCommitKey === undefined || !range) return false
    if (
      range.start.location.kind !== range.end.location.kind ||
      range.start.offset !== range.end.offset
    ) {
      return false
    }
    return textCommitKeyForRange(range) === openTextCommitKey
  }

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
      const continuesTextBurst =
        patch.textCommitKey !== undefined &&
        patch.textCommitKey === openTextCommitKey &&
        ctx.history.value.past.length > 0

      if (continuesTextBurst) {
        // The entry opened by the burst already holds the pre-burst model and
        // selection, so the commit only drops the redo tail.
        ctx.history.value = {
          past: ctx.history.value.past,
          future: [],
        }
      } else {
        const nextPast = trimHistoryEntriesToBudget(
          [
            ...ctx.history.value.past,
            {
              model: currentModel,
              selection: cloneEditorSelection(currentSelection),
              activeTextRange: cloneTextRange(currentRange),
            },
          ],
          {
            ...ctx.historyBudget,
            estimateBytes: ctx.estimateHistorySnapshotBytes,
          }
        )
        ctx.history.value = {
          past: nextPast.entries,
          future: [],
        }
      }
    }

    // Any transaction without a text-commit key (selection change, formatting,
    // structural edit, undo/redo) closes the burst. A keyed commit that changed
    // nothing leaves the current burst state untouched.
    if (patch.textCommitKey === undefined) {
      openTextCommitKey = undefined
    } else if (modelChanged) {
      openTextCommitKey = patch.textCommitKey
    }

    if (modelChanged) {
      ctx.model.value = nextModel
      // Keep the synchronous snapshot in step so a batch of commits (for
      // example flushing every dirty draft before export) builds on the
      // previous commit instead of the last rendered model.
      ctx.modelSnapshot.value = nextModel
    }
    if (selectionChanged) {
      ctx.suppressSelectionReset.value = true
      ctx.selection.value = cloneEditorSelection(nextSelection)
      ctx.selectionSnapshot.value = cloneEditorSelection(nextSelection)
    }
    if (rangeChanged) {
      ctx.suppressSelectionReset.value = true
      ctx.activeTextRange.value = cloneTextRange(nextRange)
      ctx.activeTextRangeSnapshot.value = cloneTextRange(nextRange)
    }
    if (pendingRunStyleChanged) {
      ctx.pendingRunStyle.value = nextPendingRunStyle
      ctx.pendingRunStyleSnapshot.value = cloneTextStyle(nextPendingRunStyle)
    }

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

    // Revision markup lives in `sourceXml`, and every text mutation clears it.
    // Report the paragraphs that lost revisions in this transaction so the
    // gutter cards never disappear without an explanation.
    const revisionLossParagraphCount =
      modelChanged && ctx.showTrackedChanges.value
        ? countParagraphsLosingRevisionMarkup(currentModel, nextModel)
        : 0

    if (patch.status) {
      ctx.status.value = patch.status
    } else if (revisionLossParagraphCount > 0) {
      ctx.status.value = revisionLossStatus(revisionLossParagraphCount)
    }

    return true
  }

  // ── applyModelChange ─────────────────────────────────────────────
  function applyModelChange(
    updater: (current: DocModel) => DocModel,
    successStatus?: string,
    options?: ApplyModelChangeOptions
  ): boolean {
    return dispatchEditorTransaction((current) => {
      const nextModel = updater(current.model)
      if (nextModel === current.model) return undefined
      const patch: DocxEditorTransactionPatch = {
        model: nextModel,
        status: successStatus,
        clearSelectedFormField: options?.clearSelectedFormField ?? true,
        pushHistory: options?.pushHistory,
        textCommitKey: options?.textCommitKey,
      }
      if (options && hasOwn(options, "selection")) {
        patch.selection = options.selection
      }
      if (options && hasOwn(options, "activeTextRange")) {
        patch.activeTextRange = options.activeTextRange
      }
      return patch
    })
  }

  return { dispatchEditorTransaction, applyModelChange, endTextCommitBurst, keepsTextCommitBurst }
}
