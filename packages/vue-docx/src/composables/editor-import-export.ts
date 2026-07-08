// editor-import-export.ts — .docx import / export / newDocument
//
// Extracted from useDocxEditor.ts.

import type { DocModel } from "@extend-ai/docx-core"
import { cloneDocModel } from "@extend-ai/docx-core"
import type { EditorCore } from "./editor-shared"

export function createEditorImportExport(ctx: EditorCore) {
  // ── registerPendingExportModelTransformer ────────────────────────
  const registerPendingExportModelTransformer = (
    transformer?: (model: DocModel) => DocModel
  ): void => {
    ctx.pendingExportModelTransformer.value = transformer
  }

  // ── newDocument ──────────────────────────────────────────────────
  const newDocument = (): void => {
    ctx.model.value = cloneDocModel(ctx.starterTemplate)
    ctx.documentLoadNonce.value += 1
    ctx.history.value = { past: [], future: [] }
    ctx.historyRestoreRequest.value = undefined
    ctx.basePackage.value = undefined
    ctx.selection.value = { kind: "paragraph", nodeIndex: 0 }
    ctx.activeTextRange.value = undefined
    ctx.pendingRunStyle.value = undefined
    ctx.selectedFormFieldLocation.value = undefined
    ctx.importError.value = undefined
    ctx.status.value = "Ready"
  }

  // ── importDocxFile ───────────────────────────────────────────────
  const importDocxFile = async (file: File): Promise<void> => {
    ctx.activeImportAbortController.value?.abort()
    ctx.activeImportAbortController.value = undefined

    if (!/\.docx?$/i.test(file.name)) {
      ctx.status.value = `Only .docx and .doc files are supported`
      return
    }

    ctx.isImporting.value = true
    ctx.importError.value = undefined
    ctx.status.value = `Loading ${file.name}...`

    const abortController = new AbortController()
    ctx.activeImportAbortController.value = abortController

    try {
      const { importDocxBuffer } = await import("@extend-ai/docx-core")
      const buffer = await file.arrayBuffer()
      const { model: parsedModel, package: parsedPkg } = await importDocxBuffer(buffer, {
        signal: abortController.signal,
        transferBuffer: false,
      })
      ctx.model.value = parsedModel
      ctx.basePackage.value = parsedPkg
      ctx.documentLoadNonce.value += 1
      ctx.fileName.value = file.name
      ctx.history.value = { past: [], future: [] }
      ctx.historyRestoreRequest.value = undefined
      ctx.selection.value = { kind: "paragraph", nodeIndex: 0 }
      ctx.activeTextRange.value = undefined
      ctx.pendingRunStyle.value = undefined
      ctx.selectedFormFieldLocation.value = undefined
      ctx.importError.value = undefined
      ctx.status.value = "Ready"
      ctx.isImporting.value = false
    } catch (error) {
      ctx.isImporting.value = false
      const err = error instanceof Error ? error : new Error("Unknown import error")
      ctx.importError.value = err
      ctx.status.value = `Failed to load file: ${err.message}`
    }
  }

  // ── exportDocx ───────────────────────────────────────────────────
  const exportDocx = (): void => {
    ctx.status.value = "Export... (see save dialog)"
  }

  return {
    registerPendingExportModelTransformer,
    newDocument,
    importDocxFile,
    exportDocx,
  }
}
