// useDocxEditor Vue composable — orchestrator
// Migrated from upstream @extend-ai/react-docx, editor.tsx lines 24954-29580
//
// Creates shared state, delegates to sub-modules, and returns the
// DocxEditorController. Each sub-module lives in its own file.

import { ref, computed, watch, shallowRef } from "vue"
import {
  cloneDocModel,
  cloneEditorSelection,
  cloneTextRange,
  cloneTextStyle,
  sameEditorSelection,
  sameTextRange,
  normalizeEditorCursorStateForModel,
  defaultStarterModel,
  assertValidDocxModel,
  DEFAULT_PARAGRAPH_LINE_MULTIPLE,
  resolveAutoLineSpacingMultiple,
  tableBorderPresetState,
  paragraphBorderPresetState,
} from "@extend-ai/docx-core"
import {
  firstRunStyle,
  firstTextStyleAtOffset,
  linkAtOffset,
  uniformLinkInRange,
  paragraphListType,
  normalizeParagraphStyleDefinitionsForUi,
  resolveSelectedParagraphLocation,
  getParagraphAtLocation,
  collectTrackedChangesFromModel,
  collectCommentsFromModel,
} from "@extend-ai/docx-core"

import type {
  DocModel,
  ParagraphNode,
  TextRunNode,
  OoxmlPackage,
  ParagraphStyleDefinition,
} from "@extend-ai/docx-core"
import type {
  DocxEditorController,
  DocxEditorSelection,
  DocxTextRange,
  DocxHistorySnapshot,
  DocxHistoryRestoreRequest,
  DocxEditorTransactionContext,
  DocxEditorTransactionPatch,
  DocxSelectedFormField,
  DocxFormFieldLocation,
  DocxDocumentTheme,
  DocxSelectionSessionKind,
  DocxLineSpacingInfo,
  DocxLineSpacingRule,
  DocxBorderContext,
  DocxBorderPreset,
  DocxBorderPresetState,
  DocxListType,
  DocxPaginationInfo,
  DocxImageLocation,
  DocxImageDropTarget,
  DocxImageWrapMode,
  DocxSectionParagraphLocation,
  DocxSectionImageLocation,
  UseDocxEditorOptions,
  ParagraphLocation,
  DocxTextRangeLocation,
} from "@extend-ai/docx-core"

import type { EditorCore } from "./editor-shared"
import { createEditorTransaction } from "./editor-transaction"
import { createEditorHistory } from "./editor-history"
import { createEditorSelection } from "./editor-selection"
import { createEditorFormat } from "./editor-format"
import { createEditorTextInput } from "./editor-text-input"
import { createEditorTable } from "./editor-table"
import { createEditorImage } from "./editor-image"
import { createEditorFormField } from "./editor-form-field"
import { createEditorList } from "./editor-list"
import { createEditorClipboard } from "./editor-clipboard"
import { createEditorImportExport } from "./editor-import-export"

// helpers
function typedKeys<T extends object>(obj: T): (keyof T)[] {
  return Object.keys(obj) as (keyof T)[]
}

// ── useDocxEditor ──────────────────────────────────────────────────

export function useDocxEditor(
  options: UseDocxEditorOptions = {}
): DocxEditorController {
  // --- starter template ---
  const starterModel = options.starterModel ?? defaultStarterModel
  assertValidDocxModel(starterModel)
  const starterTemplate = cloneDocModel(starterModel)

  // --- core state ---
  const model = ref<DocModel>(cloneDocModel(starterTemplate))
  const basePackage = ref<OoxmlPackage | undefined>(undefined)
  const documentLoadNonce = ref(0)
  const fileName = ref(options.initialFileName ?? "(new document)")
  const status = ref(options.initialStatus ?? "Ready")
  const importError = ref<Error | undefined>(undefined)
  const isImporting = ref(false)

  // --- selection ---
  const selection = ref<DocxEditorSelection>({ kind: "paragraph", nodeIndex: 0 })
  const activeTextRange = ref<DocxTextRange | undefined>(undefined)
  const selectedFormFieldLocation = ref<DocxFormFieldLocation | undefined>(undefined)
  const pendingRunStyle = ref<TextRunNode["style"] | undefined>(undefined)

  // --- theme / UI toggles ---
  const documentTheme = ref<DocxDocumentTheme>(options.initialDocumentTheme ?? "light")
  const showTrackedChanges = ref(options.initialShowTrackedChanges ?? false)
  const showComments = ref(options.initialShowComments ?? false)

  // --- pagination ---
  const paginationInfo = ref<DocxPaginationInfo>({ currentPage: 1, totalPages: 1 })

  // --- history ---
  const history = ref<{ past: DocxHistorySnapshot[]; future: DocxHistorySnapshot[] }>({
    past: [],
    future: [],
  })

  // --- selection session ---
  const selectionSessionKind = ref<DocxSelectionSessionKind>("idle")
  let selectionSessionTimeout: ReturnType<typeof setTimeout> | null = null
  const selectionSessionKindInternal = shallowRef<DocxSelectionSessionKind>("idle")

  // --- refs (mutable snapshots for synchronous access in callbacks) ---
  let modelSnapshot = model.value
  let selectionSnapshot = cloneEditorSelection(selection.value)
  let activeTextRangeSnapshot = cloneTextRange(activeTextRange.value)
  let pendingRunStyleSnapshot = cloneTextStyle(pendingRunStyle.value)

  let suppressNextDomSelectionRestore = false
  let domSelectionRestoreModel = model.value
  let suppressSelectionReset = false
  let historyRestoreNonceValue = 0
  const embeddedFontLoadNonce = ref(0)
  let loadedEmbeddedFontFaces: FontFace[] = []

  // --- sync snapshots ---
  watch(model, (val) => { modelSnapshot = val })
  watch(selection, (val) => { selectionSnapshot = cloneEditorSelection(val) })
  watch(activeTextRange, (val) => { activeTextRangeSnapshot = cloneTextRange(val) })
  watch(pendingRunStyle, (val) => { pendingRunStyleSnapshot = cloneTextStyle(val) })

  // --- derived: selected paragraph ---
  const selectedParagraphLocation = computed(() =>
    resolveSelectedParagraphLocation(selection.value, activeTextRange.value)
  )

  const selectedParagraph = computed(() => {
    const loc = selectedParagraphLocation.value
    return getParagraphAtLocation(model.value, loc).paragraph
  })

  // --- derived: available paragraph styles ---
  const availableParagraphStyles = computed(() => {
    const embeddedStyles = model.value.metadata.paragraphStyles ?? []
    if (embeddedStyles.length > 0) {
      return normalizeParagraphStyleDefinitionsForUi(embeddedStyles)
    }
    return normalizeParagraphStyleDefinitionsForUi(
      defaultStarterModel.metadata.paragraphStyles
    )
  })

  const paragraphStyleById = computed(() =>
    new Map(availableParagraphStyles.value.map((s) => [s.id, s]))
  )

  // --- derived: selected paragraph style id ---
  const selectedParagraphStyleId = computed(() => {
    const sp = selectedParagraph.value
    if (!sp) {
      return (
        model.value.metadata.defaultParagraphStyleId ??
        availableParagraphStyles.value.find((s) => s.isDefault)?.id
      )
    }
    if (sp.style?.styleId) return sp.style.styleId
    if (sp.style?.headingLevel) {
      const headingStyleId = `Heading${sp.style.headingLevel}`
      if (paragraphStyleById.value.has(headingStyleId)) return headingStyleId
    }
    return (
      model.value.metadata.defaultParagraphStyleId ??
      availableParagraphStyles.value.find((s) => s.isDefault)?.id
    )
  })

  // --- derived: selected line spacing ---
  const selectedLineSpacing = computed<DocxLineSpacingInfo>(() => {
    const style = selectedParagraph.value?.style?.spacing
    const lineTwips = Number.isFinite(style?.lineTwips)
      ? Math.max(1, Math.round(style!.lineTwips as number))
      : undefined
    const lineRule: DocxLineSpacingRule = style?.lineRule ?? "auto"
    const multiple =
      lineTwips !== undefined
        ? resolveAutoLineSpacingMultiple(lineTwips, DEFAULT_PARAGRAPH_LINE_MULTIPLE)
        : DEFAULT_PARAGRAPH_LINE_MULTIPLE
    return { lineRule, lineTwips, multiple: Number(multiple.toFixed(3)) }
  })

  // --- derived: border context ---
  const selectedBorderContext = computed<DocxBorderContext>(() =>
    selectedParagraphLocation.value.kind === "table-cell" ? "table" : "paragraph"
  )

  const activeBorderPresets = computed<DocxBorderPresetState>(() => {
    const loc = selectedParagraphLocation.value
    if (loc.kind === "table-cell") {
      const node = model.value.nodes[loc.tableIndex]
      if (!node || node.type !== "table") return tableBorderPresetState(undefined)
      const cell = node.rows[loc.rowIndex]?.cells[loc.cellIndex]
      return tableBorderPresetState(node.style?.borders, cell?.style?.borders)
    }
    return paragraphBorderPresetState(selectedParagraph.value?.style?.borders)
  })

  // --- derived: selected run style ---
  const selectedRunStyleFromSelection = computed(() => {
    const sp = selectedParagraph.value
    if (!sp) return undefined
    const range = activeTextRange.value
    if (!range) return firstRunStyle(sp)

    const normalized = { start: range.start, end: range.end }
    const cmp = (a: typeof normalized.start, b: typeof normalized.end): number => {
      if (a.location.kind === "paragraph" && b.location.kind === "paragraph") {
        if (a.location.nodeIndex !== b.location.nodeIndex)
          return Math.sign(a.location.nodeIndex - b.location.nodeIndex)
        return Math.sign(a.offset - b.offset)
      }
      if (a.location.kind === "table-cell" && b.location.kind === "table-cell") {
        if (a.location.tableIndex !== b.location.tableIndex)
          return Math.sign(a.location.tableIndex - b.location.tableIndex)
        if (a.location.rowIndex !== b.location.rowIndex)
          return Math.sign(a.location.rowIndex - b.location.rowIndex)
        if (a.location.cellIndex !== b.location.cellIndex)
          return Math.sign(a.location.cellIndex - b.location.cellIndex)
        if (a.location.paragraphIndex !== b.location.paragraphIndex)
          return Math.sign(a.location.paragraphIndex - b.location.paragraphIndex)
        return Math.sign(a.offset - b.offset)
      }
      if (a.location.kind === "paragraph") return -1
      return 1
    }

    const isExpanded = cmp(normalized.start, normalized.end) < 0
    const selLoc = selectedParagraphLocation.value
    const startLoc = normalized.start.location
    const endLoc = normalized.end.location
    let sameKind = false
    if (startLoc.kind === "paragraph" && endLoc.kind === "paragraph") {
      sameKind = startLoc.nodeIndex === endLoc.nodeIndex
    } else if (startLoc.kind === "table-cell" && endLoc.kind === "table-cell") {
      sameKind = true
    }
    const offset =
      isExpanded && !sameKind
        ? 0
        : startLoc.kind !== selLoc.kind
        ? 0
        : normalized.start.offset

    return firstTextStyleAtOffset(sp, offset, false) ?? firstRunStyle(sp)
  })

  const selectedRunStyle = computed(() => {
    const range = activeTextRange.value
    let expanded = false
    if (range) {
      const s = range.start.location
      const e = range.end.location
      if (s.kind === e.kind) {
        if (s.kind === "paragraph" && e.kind === "paragraph") {
          expanded = s.nodeIndex !== e.nodeIndex || range.start.offset < range.end.offset
        }
      }
    }
    if (expanded || !pendingRunStyle.value) {
      return selectedRunStyleFromSelection.value
    }
    return { ...(selectedRunStyleFromSelection.value ?? {}), ...pendingRunStyle.value }
  })

  // --- derived: selected link ---
  const selectedLink = computed(() => {
    const sp = selectedParagraph.value
    const range = activeTextRange.value
    if (!sp || !range) return undefined
    const loc = selectedParagraphLocation.value
    if (range.start.location.kind !== range.end.location.kind) return undefined
    if (range.start.location.kind === "paragraph") {
      const startLoc = range.start.location as { kind: "paragraph"; nodeIndex: number }
      const endLoc = range.end.location as { kind: "paragraph"; nodeIndex: number }
      if (startLoc.nodeIndex !== endLoc.nodeIndex) {
        if (loc.kind === "paragraph" && startLoc.nodeIndex !== loc.nodeIndex) return undefined
        return linkAtOffset(sp, range.start.offset, false)
      }
      if (range.start.offset < range.end.offset) {
        return uniformLinkInRange(sp, range.start.offset, range.end.offset)
      }
      return linkAtOffset(sp, range.start.offset, true)
    }
    return undefined
  })

  // --- derived: list type ---
  const selectedListType = computed(() =>
    selectedParagraph.value
      ? paragraphListType(selectedParagraph.value, model.value.metadata.numberingDefinitions)
      : undefined
  )

  // --- derived: tracked changes / comments ---
  const trackedChanges = computed(() => collectTrackedChangesFromModel(model.value))
  const comments = computed(() => collectCommentsFromModel(model.value))

  // --- derived: booleans ---
  const hasUnorderedList = computed(() => selectedListType.value === "unordered")
  const hasOrderedList = computed(() => selectedListType.value === "ordered")
  const canUndo = computed(() => history.value.past.length > 0)
  const canRedo = computed(() => history.value.future.length > 0)

  // --- derived: selected form field ---
  const selectedFormField = computed<DocxSelectedFormField | undefined>(() => {
    const loc = selectedFormFieldLocation.value
    if (!loc) return undefined
    const { paragraph } = getParagraphAtLocation(model.value, loc)
    if (!paragraph) return undefined
    const child = paragraph.children[loc.childIndex]
    if (!child || child.type !== "form-field") return undefined
    return { location: { ...loc }, field: child }
  })

  // --- derived: history restore request ---
  const historyRestoreRequest = ref<DocxHistoryRestoreRequest | undefined>(undefined)

  // --- cursor normalization effect ---
  watch(model, (currentModel) => {
    const { selection: nextSelection, activeTextRange: nextRange } =
      normalizeEditorCursorStateForModel(
        currentModel,
        selectionSnapshot,
        activeTextRangeSnapshot
      )

    const modelChanged = domSelectionRestoreModel !== currentModel
    domSelectionRestoreModel = currentModel

    const selectionChanged = !sameEditorSelection(selectionSnapshot, nextSelection)
    const rangeChanged = !sameTextRange(activeTextRangeSnapshot, nextRange)

    if (!selectionChanged && !rangeChanged) {
      if (suppressNextDomSelectionRestore) {
        suppressNextDomSelectionRestore = false
        return
      }
      if (modelChanged) {
        const nextNonce = historyRestoreNonceValue + 1
        historyRestoreNonceValue = nextNonce
        historyRestoreRequest.value = {
          nonce: nextNonce,
          selection: cloneEditorSelection(selectionSnapshot),
          activeTextRange: cloneTextRange(activeTextRangeSnapshot),
        }
      }
      return
    }

    suppressSelectionReset = true
    if (selectionChanged) {
      selection.value = cloneEditorSelection(nextSelection)
    }
    if (rangeChanged) {
      activeTextRange.value = cloneTextRange(nextRange)
    }

    const nextNonce = historyRestoreNonceValue + 1
    historyRestoreNonceValue = nextNonce
    historyRestoreRequest.value = {
      nonce: nextNonce,
      selection: cloneEditorSelection(nextSelection),
      activeTextRange: cloneTextRange(nextRange),
    }
  })

  // --- clear invalid form field location ---
  watch([model, selectedFormFieldLocation], ([currentModel, loc]) => {
    if (!loc) return
    const { paragraph } = getParagraphAtLocation(currentModel, loc)
    const ff = paragraph?.children[loc.childIndex]
    if (!ff || ff.type !== "form-field") {
      selectedFormFieldLocation.value = undefined
    }
  })

  // --- pagination reset on document load ---
  watch(documentLoadNonce, () => {
    paginationInfo.value = { currentPage: 1, totalPages: 1 }
  })

  // ═══════════════════════════════════════════════════════════════════
  // Build EditorCore context
  // ═══════════════════════════════════════════════════════════════════

  const ctx: EditorCore = {
    // refs
    model, selection, activeTextRange, pendingRunStyle, history, status,
    selectedFormFieldLocation, documentLoadNonce, fileName, importError,
    isImporting, documentTheme, showTrackedChanges, showComments,
    paginationInfo, historyRestoreRequest, basePackage,
    selectionSessionKind, selectionSessionKindInternal, embeddedFontLoadNonce,

    // snapshots
    modelSnapshot: { get value() { return modelSnapshot }, set value(v) { modelSnapshot = v } },
    selectionSnapshot: { get value() { return selectionSnapshot }, set value(v) { selectionSnapshot = v } },
    activeTextRangeSnapshot: { get value() { return activeTextRangeSnapshot }, set value(v) { activeTextRangeSnapshot = v } },
    pendingRunStyleSnapshot: { get value() { return pendingRunStyleSnapshot }, set value(v) { pendingRunStyleSnapshot = v } },

    // flags
    historyRestoreNonce: { get value() { return historyRestoreNonceValue }, set value(v) { historyRestoreNonceValue = v } },
    suppressNextDomSelectionRestore: { get value() { return suppressNextDomSelectionRestore }, set value(v) { suppressNextDomSelectionRestore = v } },
    domSelectionRestoreModel: { get value() { return domSelectionRestoreModel }, set value(v) { domSelectionRestoreModel = v } },
    suppressSelectionReset: { get value() { return suppressSelectionReset }, set value(v) { suppressSelectionReset = v } },
    selectionSessionTimeout: { get value() { return selectionSessionTimeout }, set value(v) { selectionSessionTimeout = v } },
    loadedEmbeddedFontFaces: { get value() { return loadedEmbeddedFontFaces }, set value(v) { loadedEmbeddedFontFaces = v } },
    activeImportAbortController: { get value() { return activeImportAbortController }, set value(v) { activeImportAbortController = v } },
    pendingExportModelTransformer: { get value() { return pendingExportModelTransformer }, set value(v) { pendingExportModelTransformer = v } },
    starterTemplate,

    // computed
    selectedParagraph,
    selectedParagraphLocation,
    selectedRunStyle,
    selectedLink,
    selectedParagraphStyleId,
    selectedLineSpacing,
    selectedBorderContext,
    activeBorderPresets,
    availableParagraphStyles,
    trackedChanges,
    comments,
    selectedListType,
    selectedFormField,
    hasUnorderedList,
    hasOrderedList,
    canUndo,
    canRedo,
  }

  // ═══════════════════════════════════════════════════════════════════
  // Initialize sub-modules
  // ═══════════════════════════════════════════════════════════════════

  // 1. editor-transaction (must be first — creates dispatch + applyChange)
  const { dispatchEditorTransaction, applyModelChange } = createEditorTransaction(ctx)

  // 2. editor-history
  const { undo, redo } = createEditorHistory(ctx)

  // 3. editor-selection
  const selModule = createEditorSelection(ctx)

  // 4. editor-format
  const fmt = createEditorFormat(ctx, dispatchEditorTransaction, applyModelChange)

  // 5. editor-text-input
  const txt = createEditorTextInput(ctx, applyModelChange)

  // 6. editor-table
  const tbl = createEditorTable(ctx, applyModelChange)

  // 7. editor-image
  const img = createEditorImage(ctx, applyModelChange)

  // 8. editor-form-field
  const ff = createEditorFormField(ctx, applyModelChange)

  // 9. editor-list
  const lst = createEditorList(ctx, applyModelChange)

  // 10. editor-clipboard
  const clp = createEditorClipboard(ctx)

  // 11. editor-import-export
  const ie = createEditorImportExport(ctx)

  // ═══════════════════════════════════════════════════════════════════
  // Thin wrappers (state accessors + trivial dispatch calls)
  // ═══════════════════════════════════════════════════════════════════

  let activeImportAbortController: AbortController | undefined
  let pendingExportModelTransformer: ((model: DocModel) => DocModel) | undefined

  const setStatus = (v: string | ((prev: string) => string)): void => {
    status.value = typeof v === "function" ? v(status.value) : v
  }
  const setDocumentThemeFn = (theme: DocxDocumentTheme): void => { documentTheme.value = theme }
  const setShowTrackedChangesFn = (v: boolean): void => { showTrackedChanges.value = v }
  const toggleShowTrackedChangesFn = (): void => { showTrackedChanges.value = !showTrackedChanges.value }
  const setShowCommentsFn = (v: boolean): void => { showComments.value = v }
  const toggleShowCommentsFn = (): void => { showComments.value = !showComments.value }
  const syncPaginationInfo = (pagination: DocxPaginationInfo): void => {
    paginationInfo.value = pagination
  }

  // Selection wrappers (thin dispatch calls)
  const setActiveTextRange = (range?: DocxTextRange): void => {
    dispatchEditorTransaction(() => ({ activeTextRange: range }))
  }
  const selectParagraph = (nodeIndex: number): void => {
    dispatchEditorTransaction(() => ({
      selection: { kind: "paragraph", nodeIndex },
      activeTextRange: undefined,
    }))
  }
  const selectTableCell = (tableIndex: number, rowIndex: number, cellIndex: number): void => {
    dispatchEditorTransaction(() => ({
      selection: { kind: "table-cell", tableIndex, rowIndex, cellIndex },
      activeTextRange: undefined,
    }))
  }

  // ═══════════════════════════════════════════════════════════════════
  // Controller assembly
  // ═══════════════════════════════════════════════════════════════════

  return {
    get model() { return model.value },
    get documentLoadNonce() { return documentLoadNonce.value },
    get fileName() { return fileName.value },
    get status() { return status.value },
    get importError() { return importError.value },
    get isImporting() { return isImporting.value },
    get documentTheme() { return documentTheme.value },
    get selection() { return selection.value },
    get activeTextRange() { return activeTextRange.value },
    get historyRestoreRequest() { return historyRestoreRequest.value },
    get selectedFormField() { return selectedFormField.value },
    get selectedParagraph() { return selectedParagraph.value },
    get selectedRunStyle() { return selectedRunStyle.value },
    get selectedLink() { return selectedLink.value },
    get pendingRunStyle() { return pendingRunStyle.value },
    get selectionSessionKind() { return selectionSessionKind.value },
    suppressNextDomSelectionRestore: selModule.suppressNextDomSelectionRestore,
    beginSelectionSession: selModule.beginSelectionSession,
    clearSelectionSession: selModule.clearSelectionSession,
    get selectedParagraphStyleId() { return selectedParagraphStyleId.value },
    get selectedLineSpacing() { return selectedLineSpacing.value },
    get selectedBorderContext() { return selectedBorderContext.value },
    get activeBorderPresets() { return activeBorderPresets.value },
    get availableParagraphStyles() { return availableParagraphStyles.value },
    get trackedChanges() { return trackedChanges.value },
    get showTrackedChanges() { return showTrackedChanges.value },
    get comments() { return comments.value },
    get showComments() { return showComments.value },
    get currentPage() { return paginationInfo.value.currentPage },
    get totalPages() { return paginationInfo.value.totalPages },
    get hasUnorderedList() { return hasUnorderedList.value },
    get hasOrderedList() { return hasOrderedList.value },
    get canUndo() { return canUndo.value },
    get canRedo() { return canRedo.value },
    registerPendingExportModelTransformer: ie.registerPendingExportModelTransformer,
    setStatus,
    setDocumentTheme: setDocumentThemeFn,
    setShowTrackedChanges: setShowTrackedChangesFn,
    setShowComments: setShowCommentsFn,
    syncPaginationInfo,
    toggleShowTrackedChanges: toggleShowTrackedChangesFn,
    toggleShowComments: toggleShowCommentsFn,
    importDocxFile: ie.importDocxFile,
    newDocument: ie.newDocument,
    exportDocx: ie.exportDocx,
    undo,
    redo,
    setHeading: fmt.setHeading,
    setParagraphStyle: fmt.setParagraphStyle,
    setLineSpacing: fmt.setLineSpacing,
    setFontFamily: fmt.setFontFamily,
    setFontSize: fmt.setFontSize,
    toggleBold: fmt.toggleBold,
    toggleItalic: fmt.toggleItalic,
    toggleUnderline: fmt.toggleUnderline,
    toggleStrike: fmt.toggleStrike,
    toggleSuperscript: fmt.toggleSuperscript,
    toggleSubscript: fmt.toggleSubscript,
    setHighlight: fmt.setHighlight,
    setTextColor: fmt.setTextColor,
    setLink: fmt.setLink,
    selectFormField: ff.selectFormField,
    toggleFormCheckbox: ff.toggleFormCheckbox,
    setFormFieldValue: ff.setFormFieldValue,
    updateFormFieldWidget: ff.updateFormFieldWidget,
    applyBorderPreset: fmt.applyBorderPreset,
    setAlignment: fmt.setAlignment,
    toggleList: lst.toggleList,
    adjustSelectedListDepth: lst.adjustSelectedListDepth,
    insertListItemAfterSelection: txt.insertListItemAfterSelection,
    splitParagraphAtSelection: txt.splitParagraphAtSelection,
    insertTable: tbl.insertTable,
    insertImageFile: img.insertImageFile,
    appendParagraph: txt.appendParagraph,
    resizeImage: img.resizeImage,
    setSyntheticTextBoxText: img.setSyntheticTextBoxText,
    setImageWrapMode: img.setImageWrapMode,
    moveFloatingImage: img.moveFloatingImage,
    moveSectionFloatingImage: img.moveSectionFloatingImage,
    moveParagraphDropCap: img.moveParagraphDropCap,
    setParagraphDropCapFontSizePt: img.setParagraphDropCapFontSizePt,
    setParagraphDropCapText: img.setParagraphDropCapText,
    moveImage: img.moveImage,
    setActiveTextRange,
    selectParagraph,
    selectTableCell,
    clearTableCellContents: tbl.clearTableCellContents,
    setTableColumnWidths: tbl.setTableColumnWidths,
    insertTableRow: tbl.insertTableRow,
    insertTableColumn: tbl.insertTableColumn,
    deleteTableRow: tbl.deleteTableRow,
    deleteTableColumn: tbl.deleteTableColumn,
    deleteTable: tbl.deleteTable,
    moveTable: tbl.moveTable,
    moveEmbeddedTableToBody: tbl.moveEmbeddedTableToBody,
    replaceExpandedSelection: txt.replaceExpandedSelection,
    deleteExpandedSelection: txt.deleteExpandedSelection,
    commitParagraphText: txt.commitParagraphText,
    commitTableCellText: txt.commitTableCellText,
    commitTableCellParagraphTextRecursive: txt.commitTableCellParagraphTextRecursive,
    commitSectionParagraphText: txt.commitSectionParagraphText,
    copy: clp.copy,
    paste: clp.paste,
  }
}
