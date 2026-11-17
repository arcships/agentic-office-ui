// editor-import-export.ts — .docx import / export / newDocument
//
// Extracted from useDocxEditor.ts.

import type { DocModel } from "@arcships/docx-core"
import {
  cloneDocModel,
  collectFormFieldsFromModel,
  createDocxRuntime,
  exportDocxBuffer,
  primeCommentAnchors,
} from "@arcships/docx-core"
import { onScopeDispose, watch } from "vue"
import type { EditorCore } from "./editor-shared"
import { isRequiredFormFieldEmpty } from "./form-field-value"
import {
  DOCX_UNSUPPORTED_FILE_MESSAGE,
  docxExportFileName,
  docxExportMimeType,
  isSupportedDocxFileName,
} from "./docx-file-types"

export function createEditorImportExport(ctx: EditorCore) {
  const runtime = createDocxRuntime()
  const loader = runtime.createLoader()
  let activeImport: symbol | undefined
  let activeExportToken: symbol | undefined

  interface ExportTask {
    token: symbol
    documentLoadNonce: number
    model: DocModel
    anchor?: HTMLAnchorElement
    objectUrl?: string
    cleanupTimer?: ReturnType<typeof setTimeout>
    cancelExport?: () => void
    delivered: boolean
    cleaned: boolean
  }

  let activeExport: ExportTask | undefined

  // New / Import / close all discard unsaved edits. Use the host's native
  // confirmation when one is available; do not block in non-browser contexts.
  const confirmDiscardChanges = (): boolean => {
    if (!ctx.isDirty.value) return true
    const confirmFn = (globalThis as { confirm?: (message?: string) => boolean }).confirm
    if (typeof confirmFn !== "function") return true
    try {
      return Boolean(confirmFn("Discard unsaved changes?"))
    } catch {
      return false
    }
  }

  const invalidateImport = (): void => {
    loader.cancel()
    activeImport = undefined
    ctx.activeImportAbortController.value = undefined
  }

  const cleanupExportTask = (task: ExportTask): void => {
    if (task.cleaned) return
    task.cleaned = true

    const cancelExport = task.cancelExport
    task.cancelExport = undefined
    if (cancelExport) {
      try {
        cancelExport()
      } catch {
        // Cancellation must not prevent the remaining resource cleanup.
      }
    }

    const cleanupTimer = task.cleanupTimer
    task.cleanupTimer = undefined
    if (cleanupTimer !== undefined) {
      try {
        clearTimeout(cleanupTimer)
      } catch {
        // Resource cleanup must continue even if a host timer shim fails.
      }
    }

    const anchor = task.anchor
    task.anchor = undefined
    if (anchor) {
      try {
        anchor.remove()
      } catch {
        // Object URL cleanup must not depend on DOM removal succeeding.
      }
    }

    const objectUrl = task.objectUrl
    task.objectUrl = undefined
    if (objectUrl) {
      try {
        URL.revokeObjectURL(objectUrl)
      } catch {
        // Cleanup is best-effort, but each owned URL is attempted only once.
      }
    }

    if (activeExport === task) {
      activeExport = undefined
    }
  }

  const invalidateExport = (reason?: "document-changed"): void => {
    const task = activeExport
    activeExportToken = undefined
    if (task) {
      const wasPending = !task.delivered
      cleanupExportTask(task)
      if (reason === "document-changed" && wasPending) {
        ctx.status.value = "Export cancelled: document changed"
      }
    }
  }

  const isCurrentExport = (task: ExportTask): boolean => (
    activeExport === task &&
    activeExportToken === task.token &&
    ctx.documentLoadNonce.value === task.documentLoadNonce &&
    ctx.model.value === task.model
  )

  // Editing commands replace the model object. Invalidate synchronously so an
  // export that is awaiting serialization cannot later download stale bytes.
  // Tell the user why the export they requested produced no file.
  watch(
    [ctx.model, ctx.documentLoadNonce],
    () => invalidateExport("document-changed"),
    { flush: "sync" },
  )

  // ── registerPendingExportModelTransformer ────────────────────────
  const registerPendingExportModelTransformer = (
    transformer?: (model: DocModel) => DocModel
  ): void => {
    ctx.pendingExportModelTransformer.value = transformer
  }

  // ── newDocument ──────────────────────────────────────────────────
  const newDocument = (): void => {
    if (!confirmDiscardChanges()) return
    invalidateImport()
    invalidateExport()
    ctx.isImporting.value = false
    const nextModel = cloneDocModel(ctx.starterTemplate)
    primeCommentAnchors(nextModel)
    ctx.model.value = nextModel
    ctx.cleanModel.value = nextModel
    ctx.documentLoadNonce.value += 1
    ctx.fileName.value = ctx.initialFileName
    ctx.history.value = { past: [], future: [] }
    ctx.historyRestoreRequest.value = undefined
    ctx.basePackage.value = undefined
    ctx.selection.value = { kind: "paragraph", nodeIndex: 0 }
    ctx.activeTextRange.value = undefined
    ctx.pendingRunStyle.value = undefined
    ctx.pendingRunStyleSnapshot.value = undefined
    ctx.selectedFormFieldLocation.value = undefined
    ctx.importError.value = undefined
    ctx.status.value = "Ready"
  }

  // ── importDocxFile ───────────────────────────────────────────────
  const importDocxFile = async (file: File): Promise<void> => {
    if (!confirmDiscardChanges()) return
    invalidateImport()
    invalidateExport()

    if (!isSupportedDocxFileName(file.name)) {
      ctx.isImporting.value = false
      ctx.status.value = DOCX_UNSUPPORTED_FILE_MESSAGE
      return
    }

    const importToken = Symbol("docx-import")
    activeImport = importToken

    ctx.isImporting.value = true
    ctx.importError.value = undefined
    ctx.status.value = `Loading ${file.name}...`

    const isCurrentImport = (): boolean => (
      activeImport === importToken
    )

    try {
      const { model: parsedModel, package: parsedPkg } = await loader.load({
        kind: "file",
        file,
        name: file.name,
      }, {
        transferBuffer: false,
      })
      if (!isCurrentImport()) return
      primeCommentAnchors(parsedModel)
      ctx.endTextCommitBurst()
      ctx.model.value = parsedModel
      ctx.cleanModel.value = parsedModel
      ctx.basePackage.value = parsedPkg
      ctx.documentLoadNonce.value += 1
      ctx.fileName.value = file.name
      ctx.history.value = { past: [], future: [] }
      ctx.historyRestoreRequest.value = undefined
      ctx.selection.value = { kind: "paragraph", nodeIndex: 0 }
      ctx.activeTextRange.value = undefined
      ctx.pendingRunStyle.value = undefined
      ctx.pendingRunStyleSnapshot.value = undefined
      ctx.selectedFormFieldLocation.value = undefined
      ctx.importError.value = undefined
      ctx.status.value = "Ready"
    } catch (error) {
      if (!isCurrentImport()) return
      const err = error instanceof Error ? error : new Error("Unknown import error")
      ctx.importError.value = err
      ctx.status.value = `Failed to load file: ${err.message}`
    } finally {
      if (!isCurrentImport()) return
      ctx.isImporting.value = false
      activeImport = undefined
      ctx.activeImportAbortController.value = undefined
    }
  }

  onScopeDispose(() => {
    invalidateImport()
    invalidateExport()
    activeImport = undefined
    loader.dispose()
    runtime.dispose()
  })

  // ── exportDocx ───────────────────────────────────────────────────
  const exportDocx = (): void => {
    // Exporting the starter model while an import is still parsing would save
    // the wrong document and race the load completion.
    if (ctx.isImporting.value || activeImport !== undefined) {
      ctx.status.value = "Wait for the document to finish loading before exporting"
      return
    }

    // contentEditable hosts buffer keystrokes in local drafts; commit them
    // before reading the model so Ctrl+S never drops the last edit.
    ctx.flushTextDrafts()
    invalidateExport()

    let exportModel: DocModel
    try {
      const sourceModel = ctx.model.value
      const transformer = ctx.pendingExportModelTransformer.value
      exportModel = transformer ? transformer(sourceModel) : sourceModel
      if (exportModel !== sourceModel) {
        // The synchronous watcher runs before the new export task is created.
        ctx.model.value = exportModel
      }
    } catch (error) {
      const err = error instanceof Error ? error : new Error("Unknown export error")
      ctx.status.value = `Export failed: ${err.message}`
      return
    }

    const invalidRequiredField = collectFormFieldsFromModel(exportModel).find(
      ({ field }) => isRequiredFormFieldEmpty(field)
    )
    if (invalidRequiredField) {
      ctx.selectedFormFieldLocation.value = invalidRequiredField.location
      const label =
        invalidRequiredField.field.widget?.name?.trim() ||
        invalidRequiredField.field.title?.trim() ||
        invalidRequiredField.field.tag?.trim()
      ctx.status.value = label
        ? `Cannot export: required form field "${label}" is empty`
        : "Cannot export: a required form field is empty"
      return
    }

    const token = Symbol("docx-export")
    const exportController = new AbortController()
    const task: ExportTask = {
      token,
      documentLoadNonce: ctx.documentLoadNonce.value,
      model: exportModel,
      cancelExport: () => exportController.abort(),
      delivered: false,
      cleaned: false,
    }
    const fileName = ctx.fileName.value
    const basePackage = ctx.basePackage.value
    const exportFileName = docxExportFileName(fileName)
    const exportMimeType = docxExportMimeType(fileName)
    activeExportToken = token
    activeExport = task

    ctx.status.value = "Exporting DOCX..."
    ;(async () => {
      try {
        const output = await exportDocxBuffer(exportModel, basePackage, {
          signal: exportController.signal,
          limits: runtime.limits,
          onProgress: (progress) => {
            if (!isCurrentExport(task)) return
            ctx.status.value = `Exporting DOCX: ${progress.phase}...`
          },
        })
        if (!isCurrentExport(task)) {
          cleanupExportTask(task)
          return
        }

        // Confirm every required DOM capability before creating an object URL.
        const ownerDocument = globalThis.document
        if (
          !ownerDocument?.body ||
          typeof ownerDocument.createElement !== "function" ||
          typeof ownerDocument.body.append !== "function"
        ) {
          throw new Error("no document available")
        }
        const anchor = ownerDocument.createElement("a")
        if (!anchor || typeof anchor.click !== "function" || typeof anchor.remove !== "function") {
          throw new Error("no download anchor available")
        }
        if (typeof URL.createObjectURL !== "function" || typeof URL.revokeObjectURL !== "function") {
          throw new Error("object URL API unavailable")
        }

        const blob = new Blob([output], {
          type: exportMimeType,
        })
        task.anchor = anchor
        task.objectUrl = URL.createObjectURL(blob)
        anchor.href = task.objectUrl
        anchor.download = exportFileName
        anchor.style.display = "none"
        ownerDocument.body.append(anchor)
        task.cleanupTimer = setTimeout(() => cleanupExportTask(task), 1000)
        anchor.click()
        task.delivered = true

        if (!isCurrentExport(task)) {
          cleanupExportTask(task)
          return
        }
        ctx.cleanModel.value = exportModel
        ctx.status.value = "Exported DOCX"
      } catch (error) {
        const shouldReport = isCurrentExport(task)
        cleanupExportTask(task)
        if (!shouldReport) return
        activeExportToken = undefined
        const err = error instanceof Error ? error : new Error("Unknown export error")
        ctx.status.value = `Export failed: ${err.message}`
      }
    })()
  }

  return {
    registerPendingExportModelTransformer,
    newDocument,
    importDocxFile,
    exportDocx,
  }
}
