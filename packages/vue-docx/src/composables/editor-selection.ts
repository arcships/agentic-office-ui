// editor-selection.ts — selection session management
//
// Extracted from useDocxEditor.ts.
// Simple selection wrappers (setActiveTextRange, selectParagraph, selectTableCell)
// are defined inline in useDocxEditor since they're thin dispatch calls.

import type { DocxSelectionSessionKind } from "@extend-ai/docx-core"
import type { EditorCore } from "./editor-shared"

export function createEditorSelection(ctx: EditorCore) {
  // ── selection session ────────────────────────────────────────────
  function clearSelectionSession(expectedKind?: DocxSelectionSessionKind): void {
    if (expectedKind && ctx.selectionSessionKindInternal.value !== expectedKind) return
    if (ctx.selectionSessionTimeout.value !== null) {
      clearTimeout(ctx.selectionSessionTimeout.value)
    }
    ctx.selectionSessionTimeout.value = null
    ctx.selectionSessionKindInternal.value = "idle"
    ctx.selectionSessionKind.value = "idle"
  }

  function beginSelectionSession(
    kind: Exclude<DocxSelectionSessionKind, "idle">,
    opts?: { settleAfterMs?: number }
  ): void {
    if (ctx.selectionSessionTimeout.value !== null) {
      clearTimeout(ctx.selectionSessionTimeout.value)
    }
    ctx.selectionSessionTimeout.value = null
    ctx.selectionSessionKindInternal.value = kind
    ctx.selectionSessionKind.value = kind

    if (Number.isFinite(opts?.settleAfterMs) && (opts?.settleAfterMs as number) > 0) {
      const expectedKind = kind
      ctx.selectionSessionTimeout.value = setTimeout(() => {
        ctx.selectionSessionTimeout.value = null
        if (ctx.selectionSessionKindInternal.value === expectedKind) {
          ctx.selectionSessionKindInternal.value = "idle"
          ctx.selectionSessionKind.value = "idle"
        }
      }, Math.max(16, Math.round(opts!.settleAfterMs as number)))
    }
  }

  function suppressNextDomSelectionRestore(): void {
    ctx.suppressNextDomSelectionRestore.value = true
  }

  return {
    beginSelectionSession,
    clearSelectionSession,
    suppressNextDomSelectionRestore,
  }
}
