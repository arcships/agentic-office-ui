// editor-import-export.ts — .docx import / export / newDocument
//
// Extracted from useDocxEditor.ts.

import type {
  DocModel,
  DocxExportProgress,
  DocxFormFieldLocation,
  OoxmlPackage,
  OoxmlPart,
} from "@arcships/docx-core"
import {
  cloneDocModel,
  collectFormFieldsFromModel,
  createDocxRuntime,
  DocxExportError,
  exportDocxBuffer,
  primeCommentAnchors,
} from "@arcships/docx-core"
import { onScopeDispose, toRaw, watch } from "vue"
import type { EditorCore } from "./editor-shared"
import { isRequiredFormFieldEmpty } from "./form-field-value"
import { prepareDocxDocumentFonts } from "./docx-document-fonts"
import {
  docxExportFileName,
  docxExportMimeType,
  isSupportedDocxFileName,
} from "./docx-file-types"

type EditorExportFailureReason = "not-ready" | "required-field"

/**
 * Controller-level refusal that still carries the reason `exportDocx()` needs
 * to choose the matching localized status message. Public callers see the
 * `DocxExportError` shape with `code: "EXPORT_FAILED"`.
 */
class EditorExportError extends DocxExportError {
  readonly reason: EditorExportFailureReason
  readonly location?: DocxFormFieldLocation
  readonly label?: string

  constructor(
    reason: EditorExportFailureReason,
    message: string,
    details: { location?: DocxFormFieldLocation; label?: string } = {},
  ) {
    super("EXPORT_FAILED", message)
    this.reason = reason
    this.location = details.location
    this.label = details.label
  }
}

/** Cancellation/progress hooks used internally by `exportDocx()`. */
interface GetDocxBytesOptions {
  signal?: AbortSignal
  onProgress?: (progress: DocxExportProgress) => void
  /** Internal: `exportDocx()` already prepared and validated this model. */
  preparedModel?: DocModel
}

interface ExportTask {
  documentLoadNonce: number
  model: DocModel
  controller: AbortController
  cancelExternal?: () => void
  anchor?: HTMLAnchorElement
  objectUrl?: string
  cleanupTimer?: ReturnType<typeof setTimeout>
  delivered: boolean
  cleaned: boolean
}

/** One `exportDocx()` invocation; invalidated when a newer export starts. */
interface ExportRequest {
  controller: AbortController
}

export function createEditorImportExport(ctx: EditorCore) {
  const runtime = createDocxRuntime()
  const loader = runtime.createLoader()
  let activeImport: symbol | undefined
  let activeExportRequest: ExportRequest | undefined
  let activeExport: ExportTask | undefined
  const pendingByteTasks = new Set<ExportTask>()
  // Applying the pending export transformer replaces the model as part of
  // preparing the export, not as a user edit. Let the export proceed instead of
  // invalidating it through the model watcher.
  let preparingExport = false

  // New / Import / close all discard unsaved edits. Use the host's native
  // confirmation when one is available; do not block in non-browser contexts.
  const confirmDiscardChanges = (): boolean => {
    if (!ctx.isDirty.value) return true
    const confirmFn = (globalThis as { confirm?: (message?: string) => boolean }).confirm
    if (typeof confirmFn !== "function") return true
    try {
      return Boolean(confirmFn(ctx.statusMessage("discardChanges")))
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
    pendingByteTasks.delete(task)
    task.controller.abort()

    const cancelExternal = task.cancelExternal
    task.cancelExternal = undefined
    if (cancelExternal) {
      try {
        cancelExternal()
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
    const tasks = new Set(pendingByteTasks)
    if (activeExport) tasks.add(activeExport)
    const wasPending = activeExportRequest !== undefined ||
      [...tasks].some((task) => !task.delivered)

    for (const task of tasks) {
      cleanupExportTask(task)
    }

    const request = activeExportRequest
    activeExportRequest = undefined
    request?.controller.abort()

    if (reason === "document-changed" && wasPending) {
      ctx.status.value = ctx.statusMessage("exportCancelledChanged")
    }
  }

  /** Cancel the previous download request without touching `getDocxBytes()` readers. */
  const invalidateDownload = (): void => {
    const request = activeExportRequest
    activeExportRequest = undefined
    request?.controller.abort()

    const task = activeExport
    if (task) cleanupExportTask(task)
  }

  const isCurrentByteTask = (task: ExportTask): boolean => (
    pendingByteTasks.has(task) &&
    !task.cleaned &&
    !task.controller.signal.aborted &&
    ctx.documentLoadNonce.value === task.documentLoadNonce &&
    ctx.model.value === task.model
  )

  const isCurrentExportRequest = (request: ExportRequest): boolean => (
    activeExportRequest === request &&
    !request.controller.signal.aborted
  )

  const isCurrentExport = (task: ExportTask): boolean => (
    activeExport === task &&
    !task.cleaned &&
    !task.controller.signal.aborted &&
    ctx.documentLoadNonce.value === task.documentLoadNonce &&
    ctx.model.value === task.model
  )

  // Editing commands replace the model object. Invalidate synchronously so an
  // export that is awaiting serialization cannot later download stale bytes.
  // Tell the user why the export they requested produced no file.
  watch(
    [ctx.model, ctx.documentLoadNonce],
    () => {
      if (preparingExport) return
      invalidateExport("document-changed")
    },
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
    ctx.commentComposer.value = undefined
    ctx.pendingRunStyle.value = undefined
    ctx.pendingRunStyleSnapshot.value = undefined
    ctx.selectedFormFieldLocation.value = undefined
    ctx.importError.value = undefined
    ctx.status.value = ctx.statusMessage("ready")
  }

  // ── importDocxFile ───────────────────────────────────────────────
  const importDocxFile = async (file: File): Promise<void> => {
    if (!confirmDiscardChanges()) return
    invalidateImport()
    invalidateExport()

    if (!isSupportedDocxFileName(file.name)) {
      ctx.isImporting.value = false
      ctx.status.value = ctx.statusMessage("unsupportedFile")
      return
    }

    const importToken = Symbol("docx-import")
    activeImport = importToken

    ctx.isImporting.value = true
    ctx.importError.value = undefined
    ctx.status.value = ctx.statusMessage("loadingFile", { name: file.name })

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
      void prepareDocxDocumentFonts(parsedModel, parsedPkg)
      ctx.documentLoadNonce.value += 1
      ctx.fileName.value = file.name
      ctx.history.value = { past: [], future: [] }
      ctx.historyRestoreRequest.value = undefined
      ctx.selection.value = { kind: "paragraph", nodeIndex: 0 }
      ctx.activeTextRange.value = undefined
      // A composer opened for the previous document anchors to a paragraph the
      // new model may not have.
      ctx.commentComposer.value = undefined
      ctx.pendingRunStyle.value = undefined
      ctx.pendingRunStyleSnapshot.value = undefined
      ctx.selectedFormFieldLocation.value = undefined
      ctx.importError.value = undefined
      ctx.status.value = ctx.statusMessage("ready")
    } catch (error) {
      if (!isCurrentImport()) return
      const err = error instanceof Error ? error : new Error(ctx.statusMessage("unknownImportError"))
      ctx.importError.value = err
      ctx.status.value = ctx.statusMessage("importFailed", { message: err.message })
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

  // ── shared export preparation / serialization ────────────────────
  const prepareExportModel = (): DocModel => {
    // Exporting the starter model while an import is still parsing would save
    // the wrong document and race the load completion.
    if (ctx.isImporting.value || activeImport !== undefined) {
      throw new EditorExportError("not-ready", ctx.statusMessage("exportNotReady"))
    }

    // contentEditable hosts buffer keystrokes in local drafts; commit them
    // before reading the model so Ctrl+S never drops the last edit.
    ctx.flushTextDrafts()

    preparingExport = true
    try {
      const sourceModel = ctx.model.value
      const transformer = ctx.pendingExportModelTransformer.value
      const exportModel = transformer ? transformer(sourceModel) : sourceModel
      if (exportModel !== sourceModel) {
        // The synchronous watcher is suppressed while preparing; the new model
        // is the one this export serializes.
        ctx.model.value = exportModel
        // Vue exposes a reactive proxy through the ref; cancellation compares
        // model identity, so serialize and track the proxy the editor exposes.
        return ctx.model.value
      }
      return exportModel
    } catch (error) {
      throw error instanceof Error ? error : new Error(ctx.statusMessage("unknownExportError"))
    } finally {
      preparingExport = false
    }
  }

  const validateExportModel = (exportModel: DocModel): void => {
    const invalidRequiredField = collectFormFieldsFromModel(exportModel).find(
      ({ field }) => isRequiredFormFieldEmpty(field)
    )
    if (!invalidRequiredField) return

    const label =
      invalidRequiredField.field.widget?.name?.trim() ||
      invalidRequiredField.field.title?.trim() ||
      invalidRequiredField.field.tag?.trim()
    throw new EditorExportError(
      "required-field",
      label
        ? ctx.statusMessage("exportRequiredField", { label })
        : ctx.statusMessage("exportRequiredFieldEmpty"),
      { location: invalidRequiredField.location, label },
    )
  }

  /** Map a preparation/serialization failure to the editor's status message. */
  const reportExportFailure = (error: unknown): void => {
    if (error instanceof EditorExportError && error.reason === "not-ready") {
      ctx.status.value = error.message
      return
    }
    if (error instanceof EditorExportError && error.reason === "required-field") {
      if (error.location) ctx.selectedFormFieldLocation.value = error.location
      ctx.status.value = error.label
        ? ctx.statusMessage("exportRequiredField", { label: error.label })
        : ctx.statusMessage("exportRequiredFieldEmpty")
      return
    }

    const err = error instanceof Error ? error : new Error(ctx.statusMessage("unknownExportError"))
    ctx.status.value = ctx.statusMessage("exportFailed", { message: err.message })
  }

  /**
   * `Worker.postMessage` structured-clones its payload; Vue refs expose
   * reactive Proxy objects that cannot cross that boundary. Copy the package
   * into plain Maps/objects before posting. Binary assets are cloneable typed
   * arrays and are shared rather than copied.
   */
  const cloneExportPackage = (pkg: OoxmlPackage | undefined): OoxmlPackage | undefined => {
    if (!pkg) return undefined
    const parts = new Map<string, OoxmlPart>()
    for (const [name, part] of pkg.parts) {
      parts.set(name, { name: part.name, content: part.content })
    }
    const binaryAssets = new Map<string, Uint8Array>()
    for (const [name, bytes] of pkg.binaryAssets) {
      binaryAssets.set(name, toRaw(bytes))
    }
    return { parts, binaryAssets }
  }

  const createByteTask = (
    exportModel: DocModel,
    options: GetDocxBytesOptions,
  ): ExportTask => {
    const task: ExportTask = {
      documentLoadNonce: ctx.documentLoadNonce.value,
      model: exportModel,
      controller: new AbortController(),
      delivered: false,
      cleaned: false,
    }

    const externalSignal = options.signal
    if (externalSignal) {
      const abort = (): void => task.controller.abort()
      if (externalSignal.aborted) {
        task.controller.abort()
      } else {
        externalSignal.addEventListener("abort", abort, { once: true })
        task.cancelExternal = () => externalSignal.removeEventListener("abort", abort)
      }
    }

    pendingByteTasks.add(task)
    return task
  }

  // ── getDocxBytes / getDocxBlob ───────────────────────────────────
  // The public controller methods take no options; `exportDocx()` passes an
  // internal request signal and progress callback so both share one path.
  const getDocxBytes = async (
    options: GetDocxBytesOptions = {},
  ): Promise<ArrayBuffer> => {
    const exportModel = options.preparedModel ?? prepareExportModel()
    validateExportModel(exportModel)

    const basePackage = ctx.basePackage.value
    const usesWorker = typeof Worker !== "undefined"
    // `cloneDocModel` also deep-copies the model graph, so nested reactive
    // values cannot surface as Proxy objects inside the Worker message.
    const messageModel = usesWorker ? cloneDocModel(toRaw(exportModel)) : exportModel
    const messageBasePackage = usesWorker ? cloneExportPackage(basePackage) : basePackage
    const task = createByteTask(exportModel, options)
    try {
      const output = await exportDocxBuffer(messageModel, messageBasePackage, {
        signal: task.controller.signal,
        limits: runtime.limits,
        onProgress: (progress) => {
          if (!isCurrentByteTask(task)) return
          options.onProgress?.(progress)
        },
      })
      if (!isCurrentByteTask(task)) {
        throw new DocxExportError("ABORTED", "DOCX export was aborted")
      }
      return output
    } finally {
      cleanupExportTask(task)
    }
  }

  const getDocxBlob = async (
    options: GetDocxBytesOptions = {},
  ): Promise<Blob> => {
    const fileName = ctx.fileName.value
    const bytes = await getDocxBytes(options)
    return new Blob([bytes], { type: docxExportMimeType(fileName) })
  }

  // ── exportDocx ───────────────────────────────────────────────────
  const exportDocx = (): void => {
    // Exporting the starter model while an import is still parsing would save
    // the wrong document and race the load completion.
    if (ctx.isImporting.value || activeImport !== undefined) {
      ctx.status.value = ctx.statusMessage("exportNotReady")
      return
    }

    invalidateDownload()

    let exportModel: DocModel
    try {
      exportModel = prepareExportModel()
      validateExportModel(exportModel)
    } catch (error) {
      reportExportFailure(error)
      return
    }

    const request: ExportRequest = {
      controller: new AbortController(),
    }
    activeExportRequest = request
    const exportFileName = docxExportFileName(ctx.fileName.value)
    const exportMimeType = docxExportMimeType(ctx.fileName.value)
    let downloadTask: ExportTask | undefined

    ctx.status.value = ctx.statusMessage("exporting")
    ;(async () => {
      try {
        const output = await getDocxBytes({
          preparedModel: exportModel,
          signal: request.controller.signal,
          onProgress: (progress) => {
            if (!isCurrentExportRequest(request)) return
            ctx.status.value = ctx.statusMessage("exportingPhase", { phase: progress.phase })
          },
        })
        if (!isCurrentExportRequest(request)) return

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

        downloadTask = {
          documentLoadNonce: ctx.documentLoadNonce.value,
          model: ctx.model.value,
          controller: new AbortController(),
          delivered: false,
          cleaned: false,
        }
        activeExport = downloadTask
        const task = downloadTask

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
          if (activeExportRequest === request) activeExportRequest = undefined
          return
        }
        activeExportRequest = undefined
        ctx.cleanModel.value = task.model
        ctx.status.value = ctx.statusMessage("exported")
      } catch (error) {
        const shouldReport = isCurrentExportRequest(request)
        if (shouldReport) activeExportRequest = undefined
        if (downloadTask) cleanupExportTask(downloadTask)
        if (!shouldReport) return

        reportExportFailure(error)
      }
    })()
  }

  return {
    registerPendingExportModelTransformer,
    newDocument,
    importDocxFile,
    exportDocx,
    getDocxBytes,
    getDocxBlob,
  }
}
