// editor-import-export.ts — .docx import / export / newDocument
//
// Extracted from useDocxEditor.ts.

import type { DocModel } from "@arcships/docx-core"
import { cloneDocModel, createDocxRuntime, serializeDocx } from "@arcships/docx-core"
import { onScopeDispose, watch } from "vue"
import type { EditorCore } from "./editor-shared"

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
    cleaned: boolean
  }

  let activeExport: ExportTask | undefined

  const invalidateImport = (): void => {
    loader.cancel()
    activeImport = undefined
    ctx.activeImportAbortController.value = undefined
  }

  const cleanupExportTask = (task: ExportTask): void => {
    if (task.cleaned) return
    task.cleaned = true

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

  const invalidateExport = (): void => {
    activeExportToken = undefined
    if (activeExport) cleanupExportTask(activeExport)
  }

  const isCurrentExport = (task: ExportTask): boolean => (
    activeExport === task &&
    activeExportToken === task.token &&
    ctx.documentLoadNonce.value === task.documentLoadNonce &&
    ctx.model.value === task.model
  )

  // Editing commands replace the model object. Invalidate synchronously so an
  // export that is awaiting serialization cannot later download stale bytes.
  watch(
    [ctx.model, ctx.documentLoadNonce],
    () => invalidateExport(),
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
    invalidateImport()
    invalidateExport()
    ctx.isImporting.value = false
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
    invalidateImport()
    invalidateExport()

    if (!/\.docx?$/i.test(file.name)) {
      ctx.isImporting.value = false
      ctx.status.value = `Only .docx and .doc files are supported`
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

    const token = Symbol("docx-export")
    const task: ExportTask = {
      token,
      documentLoadNonce: ctx.documentLoadNonce.value,
      model: exportModel,
      cleaned: false,
    }
    const fileName = ctx.fileName.value
    const basePackage = ctx.basePackage.value
    activeExportToken = token
    activeExport = task

    ;(async () => {
      try {
        const output = await serializeDocx(exportModel, basePackage)
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
          type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        })
        task.anchor = anchor
        task.objectUrl = URL.createObjectURL(blob)
        anchor.href = task.objectUrl
        anchor.download = /\.docx?$/i.test(fileName)
          ? fileName.replace(/\.docx?$/i, "") + "-edited.docx"
          : "edited.docx"
        anchor.style.display = "none"
        ownerDocument.body.append(anchor)
        task.cleanupTimer = setTimeout(() => cleanupExportTask(task), 1000)
        anchor.click()

        if (!isCurrentExport(task)) {
          cleanupExportTask(task)
          return
        }
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
