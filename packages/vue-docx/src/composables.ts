/**
 * @extend-ai/vue-docx 鈥?Vue 3 composables & components for DOCX
 *
 * Vue API aligned with the public @extend-ai/react-docx hooks.
 * Built on @extend-ai/docx-core (framework-agnostic engine).
 */

import {
  ref, shallowRef, computed, watch, readonly,
  type Ref, type ComputedRef, type ShallowRef,
} from "vue"
import {
  insertParagraph,
  removeParagraph,
  duplicateParagraph,
  setParagraphHeading,
  setParagraphAlignment,
  updateParagraphText,
  toggleRunStyleFlag,
  buildDocModelFromBytes,
  type DocModel, type DocxEditorController, type DocxEditorStatus,
  type DocxSelection, type DocxDocumentTheme, type DocxLineSpacingInfo,
  type DocxBorderContext, type DocxBorderPreset, type DocxTrackedChange,
  type DocxComment, type ParagraphStyleDefinition, type DocxPageThumbnailItem,
  type DocxTextRange, type UseDocxPageThumbnailsOptions,
} from "@extend-ai/docx-core"

export type * from "@extend-ai/docx-core"

// ============================================================
// useDocxModel 鈥?Lightweight hook for loading .docx into model
// ============================================================

export interface UseDocxModelState {
  model: Ref<DocModel | null>
  isLoading: Ref<boolean>
  error: Ref<Error | null>
}

export function useDocxModel(file?: Ref<ArrayBuffer | null | undefined> | ArrayBuffer | null): UseDocxModelState {
  const model = ref<DocModel | null>(null) as Ref<DocModel | null>
  const isLoading = ref(false)
  const error = ref<Error | null>(null)

  const fileRef = ref(file)

  watch(fileRef, async (f) => {
    if (!f) {
      model.value = null
      return
    }
    isLoading.value = true
    error.value = null
    try {
      const { buildDocModelFromBytes } = await import("@extend-ai/docx-core")
      const result = await buildDocModelFromBytes(
        f instanceof ArrayBuffer ? f : new Uint8Array(f as unknown as ArrayBuffer).buffer
      )
      model.value = result.model
    } catch (e) {
      error.value = e instanceof Error ? e : new Error(String(e))
    } finally {
      isLoading.value = false
    }
  }, { immediate: true })

  return { model, isLoading, error }
}

// ============================================================
// useDocxEditor — Core editor controller aligned with react-docx useDocxEditor
// ============================================================

export interface UseDocxEditorOptions {
  starterModel?: DocModel
  initialFileName?: string
  initialStatus?: DocxEditorStatus
  initialDocumentTheme?: DocxDocumentTheme
  initialShowTrackedChanges?: boolean
  initialShowComments?: boolean
}

export function useDocxEditor(options: UseDocxEditorOptions = {}): DocxEditorController {
  const model = ref<DocModel>(
    options.starterModel ?? { nodes: [], metadata: { headerSections: [], footerSections: [] } }
  )
  const fileName = ref(options.initialFileName ?? "Untitled.docx")
  const status = ref<DocxEditorStatus>(options.initialStatus ?? "ready")
  const isDirty = ref(false)
  const documentTheme = ref<DocxDocumentTheme>(options.initialDocumentTheme ?? { mode: "light" })
  const showTrackedChanges = ref(options.initialShowTrackedChanges ?? false)
  const showComments = ref(options.initialShowComments ?? false)

  // Selection
  const selection: { text: string; nodeIndex: number | null; runIndex: number | null; textRange: DocxTextRange | null; isCollapsed: boolean } = {
    text: "",
    nodeIndex: null,
    runIndex: 0,
    textRange: null,
    isCollapsed: true,
  }

  // Pagination
  const currentPage = ref(1)
  const totalPages = ref(1)

  // History stacks for undo/redo
  const history = ref<DocModel[]>([])
  const future = ref<DocModel[]>([])
  const canUndo = computed(() => history.value.length > 0)
  const canRedo = computed(() => future.value.length > 0)

  function pushHistory(newModel: DocModel) {
    history.value.push(model.value)
    future.value = []
    model.value = newModel
    isDirty.value = true
    if (history.value.length > 100) history.value.shift()
  }

  // Build the controller matching the DocxEditorController interface
  const controller: DocxEditorController = {
    get model() { return model.value },
    get fileName() { return fileName.value },
    get selection() { return selection },
    get status() { return status.value },
    get isDirty() { return isDirty.value },

    undo() {
      if (history.value.length === 0) return
      future.value.push(model.value)
      model.value = history.value.pop()!
      isDirty.value = history.value.length > 0
    },
    redo() {
      if (future.value.length === 0) return
      history.value.push(model.value)
      model.value = future.value.pop()!
      isDirty.value = true
    },
    get canUndo() { return canUndo.value },
    get canRedo() { return canRedo.value },

    async importDocxFile(file: File | Blob) {
      status.value = "loading"
      try {
        const buf = await file.arrayBuffer()
        const { buildDocModelFromBytes } = await import("@extend-ai/docx-core")
        const result = await buildDocModelFromBytes(buf)
        pushHistory(result.model)
        fileName.value = (file as File).name ?? "Imported.docx"
        status.value = "ready"
      } catch (e) {
        status.value = "error"
        throw e
      }
    },

    async importDocxBuffer(buffer: ArrayBuffer, name?: string) {
      status.value = "loading"
      try {
        const { buildDocModelFromBytes } = await import("@extend-ai/docx-core")
        const result = await buildDocModelFromBytes(buffer)
        pushHistory(result.model)
        if (name) fileName.value = name
        status.value = "ready"
      } catch (e) {
        status.value = "error"
        throw e
      }
    },

    async exportDocx(): Promise<Blob> {
      const { serializeDocx } = await import("@extend-ai/docx-core")
      const buf = await serializeDocx(model.value)
      return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" })
    },

    // Text formatting
    toggleBold() {
      if (selection.nodeIndex != null) pushHistory(toggleRunStyleFlag(model.value, selection.nodeIndex, selection.runIndex ?? 0, "bold"))
    },
    toggleItalic() {
      if (selection.nodeIndex != null) pushHistory(toggleRunStyleFlag(model.value, selection.nodeIndex, selection.runIndex ?? 0, "italic"))
    },
    toggleUnderline() {
      if (selection.nodeIndex != null) pushHistory(toggleRunStyleFlag(model.value, selection.nodeIndex, selection.runIndex ?? 0, "underline"))
    },
    toggleStrike() {
      if (selection.nodeIndex != null) pushHistory(toggleRunStyleFlag(model.value, selection.nodeIndex, selection.runIndex ?? 0, "strike"))
    },
    setFontSize(_size: number) {},
    setFontFamily(_family: string) {},
    setTextColor(_color: string | null) {},
    setHighlight(_color: string | null) {},
    setSuperscript() {},
    setSubscript() {},


    updateParagraphText(nodeIndex: number, text: string) {
      selection.nodeIndex = nodeIndex
      selection.runIndex = 0
      selection.text = text
      pushHistory(updateParagraphText(model.value, nodeIndex, text))
    },

    selectParagraph(nodeIndex: number, runIndex = 0) {
      selection.nodeIndex = nodeIndex
      selection.runIndex = runIndex
      const node = model.value.nodes[nodeIndex]
      selection.text = node?.type === "paragraph" ? node.children.map((child) => child.type === "text" ? child.text : "").join("") : ""
    },

    insertParagraph() {
      pushHistory(insertParagraph(model.value, ""))
    },
    removeParagraph() {
      if (selection.nodeIndex != null) {
        pushHistory(removeParagraph(model.value, selection.nodeIndex))
      }
    },
    duplicateParagraph() {
      if (selection.nodeIndex != null) {
        pushHistory(duplicateParagraph(model.value, selection.nodeIndex))
      }
    },
    setParagraphHeading(level: number | null) {
      if (selection.nodeIndex != null) {
        pushHistory(setParagraphHeading(model.value, selection.nodeIndex, level as 1|2|3|4|5|6|undefined))
      }
    },
    setParagraphAlignment(align: "left" | "center" | "right" | "justify" | null) {
      if (selection.nodeIndex != null) {
        pushHistory(setParagraphAlignment(model.value, selection.nodeIndex, align ?? undefined))
      }
    },
    setLineSpacing(_info: DocxLineSpacingInfo) {},
    setParagraphStyle(_styleId: string) {},

    insertTable(_rows: number, _cols: number) {},
    removeTable() {},
    insertTableRow() {},
    insertTableColumn() {},

    async insertImage(_file: File) {},

    applyBorderPreset(_preset: DocxBorderPreset) {},
    get selectedBorderContext() { return null },
    get activeBorderPresets() { return [] },

    setFormFieldValue(_fieldId: string, _value: string) {},
    toggleFormCheckbox(_fieldId: string) {},
    selectFormField(_fieldId: string | null) {},
    get selectedFormField() { return null },

    get trackedChanges() { return [] },
    get showTrackedChanges() { return showTrackedChanges.value },
    setShowTrackedChanges(show: boolean) { showTrackedChanges.value = show },
    toggleShowTrackedChanges() { showTrackedChanges.value = !showTrackedChanges.value },

    get comments() { return [] },
    get showComments() { return showComments.value },
    setShowComments(show: boolean) { showComments.value = show },
    toggleShowComments() { showComments.value = !showComments.value },

    get documentTheme() { return documentTheme.value },
    setDocumentTheme(theme: DocxDocumentTheme) { documentTheme.value = theme },
    toggleDocumentTheme() {
      documentTheme.value = {
        ...documentTheme.value,
        mode: documentTheme.value.mode === "light" ? "dark" : "light",
      }
    },

    get currentPage() { return currentPage.value },
    get totalPages() { return totalPages.value },

    get availableParagraphStyles() { return [] },
    get selectedParagraphStyleId() { return null },
    get selectedLineSpacing() { return null },
  }

  return controller
}

// ============================================================
// useDocxDocumentTheme
// ============================================================

export interface UseDocxDocumentThemeResult {
  documentTheme: ComputedRef<DocxDocumentTheme>
  setDocumentTheme: (theme: DocxDocumentTheme) => void
  toggleDocumentTheme: () => void
}

export function useDocxDocumentTheme(
  editor: Pick<DocxEditorController, "documentTheme" | "setDocumentTheme" | "toggleDocumentTheme">
): UseDocxDocumentThemeResult {
  const theme = ref<DocxDocumentTheme>(editor.documentTheme)

  watch(() => editor.documentTheme, (t) => { theme.value = t })

  return {
    documentTheme: computed(() => theme.value),
    setDocumentTheme: editor.setDocumentTheme,
    toggleDocumentTheme: editor.toggleDocumentTheme,
  }
}

// ============================================================
// useDocxTrackChanges
// ============================================================

export interface UseDocxTrackChangesResult {
  trackedChanges: ComputedRef<DocxTrackedChange[]>
  showTrackedChanges: ComputedRef<boolean>
  setShowTrackedChanges: (show: boolean) => void
  toggleShowTrackedChanges: () => void
}

export function useDocxTrackChanges(
  editor: Pick<DocxEditorController, "trackedChanges" | "showTrackedChanges" | "setShowTrackedChanges" | "toggleShowTrackedChanges">
): UseDocxTrackChangesResult {
  return {
    trackedChanges: computed(() => editor.trackedChanges),
    showTrackedChanges: computed(() => editor.showTrackedChanges),
    setShowTrackedChanges: editor.setShowTrackedChanges,
    toggleShowTrackedChanges: editor.toggleShowTrackedChanges,
  }
}

// ============================================================
// useDocxComments
// ============================================================

export interface UseDocxCommentsResult {
  comments: ComputedRef<DocxComment[]>
  showComments: ComputedRef<boolean>
  setShowComments: (show: boolean) => void
  toggleShowComments: () => void
}

export function useDocxComments(
  editor: Pick<DocxEditorController, "comments" | "showComments" | "setShowComments" | "toggleShowComments">
): UseDocxCommentsResult {
  return {
    comments: computed(() => editor.comments),
    showComments: computed(() => editor.showComments),
    setShowComments: editor.setShowComments,
    toggleShowComments: editor.toggleShowComments,
  }
}

// ============================================================
// useDocxPageLayout
// ============================================================

export interface UseDocxPageLayoutResult {
  pageWidth: ComputedRef<number>
  pageHeight: ComputedRef<number>
  marginTop: ComputedRef<number>
  marginBottom: ComputedRef<number>
  marginLeft: ComputedRef<number>
  marginRight: ComputedRef<number>
}

export function useDocxPageLayout(
  editor: Pick<DocxEditorController, "model">
): UseDocxPageLayoutResult {
  return {
    pageWidth: computed(() => 816),
    pageHeight: computed(() => 1056),
    marginTop: computed(() => 72),
    marginBottom: computed(() => 72),
    marginLeft: computed(() => 72),
    marginRight: computed(() => 72),
  }
}

// ============================================================
// useDocxPagination
// ============================================================

export interface UseDocxPaginationResult {
  currentPage: ComputedRef<number>
  totalPages: ComputedRef<number>
}

export function useDocxPagination(
  editor: Pick<DocxEditorController, "currentPage" | "totalPages">
): UseDocxPaginationResult {
  return {
    currentPage: computed(() => editor.currentPage),
    totalPages: computed(() => editor.totalPages),
  }
}

// ============================================================
// useDocxParagraphStyles
// ============================================================

export interface UseDocxParagraphStylesResult {
  availableStyles: ComputedRef<ParagraphStyleDefinition[]>
  selectedStyleId: ComputedRef<string | null>
  setStyle: (styleId: string) => void
}

export function useDocxParagraphStyles(
  editor: Pick<DocxEditorController, "availableParagraphStyles" | "selectedParagraphStyleId" | "setParagraphStyle">
): UseDocxParagraphStylesResult {
  return {
    availableStyles: computed(() => editor.availableParagraphStyles),
    selectedStyleId: computed(() => editor.selectedParagraphStyleId),
    setStyle: editor.setParagraphStyle,
  }
}

// ============================================================
// useDocxLineSpacing
// ============================================================

export interface UseDocxLineSpacingResult {
  lineSpacing: ComputedRef<DocxLineSpacingInfo | null>
  setLineSpacing: (info: DocxLineSpacingInfo) => void
}

export function useDocxLineSpacing(
  editor: Pick<DocxEditorController, "selectedLineSpacing" | "setLineSpacing">
): UseDocxLineSpacingResult {
  return {
    lineSpacing: computed(() => editor.selectedLineSpacing),
    setLineSpacing: editor.setLineSpacing,
  }
}

// ============================================================
// useDocxBorders
// ============================================================

export interface UseDocxBordersResult {
  borderContext: ComputedRef<DocxBorderContext | null>
  activePresets: ComputedRef<DocxBorderPreset[]>
  applyPreset: (preset: DocxBorderPreset) => void
}

export function useDocxBorders(
  editor: Pick<DocxEditorController, "selectedBorderContext" | "activeBorderPresets" | "applyBorderPreset">
): UseDocxBordersResult {
  return {
    borderContext: computed(() => editor.selectedBorderContext),
    activePresets: computed(() => editor.activeBorderPresets),
    applyPreset: editor.applyBorderPreset,
  }
}

// ============================================================
// useDocxViewerThumbnails / useDocxPageThumbnails
// ============================================================

export interface UseDocxPageThumbnailsResult {
  items: ComputedRef<DocxPageThumbnailItem[]>
  isLoading: Ref<boolean>
}

export function useDocxViewerThumbnails(
  editor: DocxEditorController,
  _options?: UseDocxPageThumbnailsOptions,
): UseDocxPageThumbnailsResult {
  const items = ref<DocxPageThumbnailItem[]>([])
  const isLoading = ref(false)

  // Thumbnail generation via canvas - to be implemented with actual renderer
  watch(() => editor.model, () => {
    isLoading.value = true
    const pages = editor.totalPages
    items.value = Array.from({ length: pages }, (_, i) => ({
      pageNumber: i + 1,
      width: 120,
      height: 160,
      paint(_canvas: HTMLCanvasElement) {},
      canvasRef: null,
      async renderToCanvas(_canvas: HTMLCanvasElement) {},
    }))
    isLoading.value = false
  })

  return { items: computed(() => items.value), isLoading }
}

// Alias
export const useDocxPageThumbnails = useDocxViewerThumbnails

