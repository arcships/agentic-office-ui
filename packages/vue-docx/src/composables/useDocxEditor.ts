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
  shouldReissueDomSelectionRestore,
  defaultStarterModel,
  assertValidDocxModel,
  DEFAULT_PARAGRAPH_LINE_MULTIPLE,
  resolveAutoLineSpacingMultiple,
  tableBorderPresetState,
  paragraphBorderPresetState,
  parseEmbeddedTableRuntimeKey,
  cloneTableCellSelectionRange,
  normalizeTableCellSelectionRange,
  tableCellSelectionRangeHasMultipleCells,
  isOpenDocumentModel,
} from "@arcships/docx-core"
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
  primeCommentAnchors,
  applyTrackedChangeDecision,
  addHeaderFooter as addSectionHeaderFooter,
  setPageMargins as applySectionPageMargins,
  setPageSize as applySectionPageSize,
} from "@arcships/docx-core"

import type {
  DocModel,
  DocxTrackedChange,
  DocxTrackedChangeDecision,
  DocxTrackedChangeDecisionScope,
  ParagraphNode,
  TextRunNode,
  OoxmlPackage,
  ParagraphStyleDefinition,
} from "@arcships/docx-core"
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
  DocxCommentComposer,
  DocxSelectionSessionKind,
  DocxTableCellSelectionRange,
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
  DocxSectionRegion,
  DocxPageMarginsUpdate,
  DocxPageSizeUpdate,
  DocxPageSetupCommandOptions,
  DocxRevisionDisplay,
  AddHeaderFooterOptions,
  UseDocxEditorOptions,
  DocxEditorStatusMessages,
  ParagraphLocation,
} from "@arcships/docx-core"
import { formatOfficeMessage, officeMessagesEn } from "@arcships/office-interaction"

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
import { createEditorComments } from "./editor-comments"
import { createEditorSearchReplace } from "./editor-search-replace"
import {
  DEFAULT_DOCX_HISTORY_MAX_BYTES,
  estimateHistoryEntryBytes,
} from "./history-budget"

// helpers
function typedKeys<T extends object>(obj: T): (keyof T)[] {
  return Object.keys(obj) as (keyof T)[]
}

// Shared empty result so a hidden revision list keeps a stable identity and
// never allocates on keystrokes.
const EMPTY_TRACKED_CHANGES: DocxTrackedChange[] = []

const DEFAULT_STATUS_MESSAGES: DocxEditorStatusMessages =
  officeMessagesEn.docx.editor.status

function resolveStatusMessages(
  option: UseDocxEditorOptions["statusMessages"],
): DocxEditorStatusMessages {
  const provided = typeof option === "function" ? option() : option
  return { ...DEFAULT_STATUS_MESSAGES, ...provided }
}

// ── useDocxEditor ──────────────────────────────────────────────────

type RuntimeKeyTableController = {
  selectTableCellByRuntimeKey: (
    tableRuntimeKey: string,
    rowIndex: number,
    cellIndex: number
  ) => void
  commitTableCellTextByRuntimeKey: (
    tableRuntimeKey: string,
    rowIndex: number,
    cellIndex: number,
    text: string
  ) => void
  commitTableCellParagraphTextByRuntimeKey: (
    tableRuntimeKey: string,
    rowIndex: number,
    cellIndex: number,
    paragraphIndex: number,
    text: string
  ) => void
}

export function useDocxEditor(
  options: UseDocxEditorOptions = {}
): DocxEditorController & RuntimeKeyTableController {
  // --- starter template ---
  const starterModel = options.starterModel ?? defaultStarterModel
  assertValidDocxModel(starterModel)
  const starterTemplate = cloneDocModel(starterModel)

  // --- core state ---
  const model = ref<DocModel>(cloneDocModel(starterTemplate))
  primeCommentAnchors(model.value)
  const cleanModel = ref<DocModel>(model.value)
  const isDirty = computed(() => model.value !== cleanModel.value)
  const basePackage = ref<OoxmlPackage | undefined>(undefined)
  const documentLoadNonce = ref(0)
  const initialFileName = options.initialFileName ?? officeMessagesEn.docx.editor.newDocument
  const fileName = ref(initialFileName)
  const statusMessage = (
    key: keyof DocxEditorStatusMessages,
    values: Record<string, string | number> = {},
  ): string => formatOfficeMessage(
    resolveStatusMessages(options.statusMessages)[key] ?? DEFAULT_STATUS_MESSAGES[key],
    values,
  )
  const status = ref(options.initialStatus ?? statusMessage("ready"))
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
  const revisionDisplay = ref<DocxRevisionDisplay>(
    options.initialRevisionDisplay ?? "balloons"
  )
  const setRevisionDisplay = (display: DocxRevisionDisplay): void => {
    revisionDisplay.value = display
  }
  const showComments = ref(options.initialShowComments ?? false)
  const commentComposer = ref<DocxCommentComposer | undefined>(undefined)

  // --- pagination ---
  const paginationInfo = ref<DocxPaginationInfo>({ currentPage: 1, totalPages: 1 })

  // --- history ---
  const history = ref<{ past: DocxHistorySnapshot[]; future: DocxHistorySnapshot[] }>({
    past: [],
    future: [],
  })
  const requestedHistoryMaxEntries = options.historyMaxEntries ?? 100
  const requestedHistoryMaxBytes = options.historyMaxBytes ?? DEFAULT_DOCX_HISTORY_MAX_BYTES
  const historyBudget = Object.freeze({
    maxEntries:
      Number.isFinite(requestedHistoryMaxEntries) && requestedHistoryMaxEntries > 0
        ? Math.max(1, Math.floor(requestedHistoryMaxEntries))
        : 100,
    maxBytes:
      Number.isFinite(requestedHistoryMaxBytes) && requestedHistoryMaxBytes > 0
        ? Math.max(1, Math.floor(requestedHistoryMaxBytes))
      : DEFAULT_DOCX_HISTORY_MAX_BYTES,
  })
  const historyEntryByteEstimates = new WeakMap<DocxHistorySnapshot, number>()
  const estimateHistorySnapshotBytes = (snapshot: DocxHistorySnapshot): number => {
    const cached = historyEntryByteEstimates.get(snapshot)
    if (cached !== undefined) return cached
    const estimatedBytes = estimateHistoryEntryBytes(snapshot)
    historyEntryByteEstimates.set(snapshot, estimatedBytes)
    return estimatedBytes
  }

  // --- selection session ---
  const selectionSessionKind = ref<DocxSelectionSessionKind>("idle")
  let selectionSessionTimeout: ReturnType<typeof setTimeout> | null = null
  const selectionSessionKindInternal = shallowRef<DocxSelectionSessionKind>("idle")
  const textDraftFlushers = new Set<() => void>()
  const flushTextDrafts = (): void => {
    for (const flusher of [...textDraftFlushers]) flusher()
  }

  // The transaction module owns the burst state; this holder lets the editor
  // context (built first) expose a stable closer to every sub-module.
  let endTextCommitBurstImpl: () => void = () => {}
  const endTextCommitBurst = (): void => endTextCommitBurstImpl()

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
  // Collecting revisions walks every paragraph and table cell, so it only runs
  // while the markup is actually rendered. Hidden revisions cost nothing on
  // keystrokes; `trackedChangeCount` stays `null` until they are shown.
  const trackedChanges = computed<DocxTrackedChange[]>(() =>
    showTrackedChanges.value ? collectTrackedChangesFromModel(model.value) : EMPTY_TRACKED_CHANGES
  )
  const trackedChangeCount = computed<number | null>(() =>
    showTrackedChanges.value ? trackedChanges.value.length : null
  )
  const comments = computed(() => collectCommentsFromModel(model.value))

  // --- edits that silently drop revision markup ---
  // Every text edit clears `sourceXml`, which is where revisions live. The
  // transaction dispatcher reports exactly which paragraphs lost their
  // revision markup (see createEditorTransaction) instead of letting the
  // gutter cards disappear without explanation.

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
      if (!shouldReissueDomSelectionRestore({
        modelChanged,
        selectionChanged,
        rangeChanged,
        activeTextRange: activeTextRangeSnapshot,
        suppressNext: false,
        selectionSessionKind: selectionSessionKindInternal.value,
      })) return
      const nextNonce = historyRestoreNonceValue + 1
      historyRestoreNonceValue = nextNonce
      historyRestoreRequest.value = {
        nonce: nextNonce,
        selection: cloneEditorSelection(selectionSnapshot),
        activeTextRange: cloneTextRange(activeTextRangeSnapshot),
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
    activeFormFieldEditKey = undefined
  })

  // ═══════════════════════════════════════════════════════════════════
  // Build EditorCore context
  // ═══════════════════════════════════════════════════════════════════

  const ctx: EditorCore = {
    // refs
    model, cleanModel, isDirty, selection, activeTextRange, pendingRunStyle, history, historyBudget,
    estimateHistorySnapshotBytes, status,
    statusMessage,
    selectedFormFieldLocation, documentLoadNonce, fileName, initialFileName, importError,
    isImporting, documentTheme, showTrackedChanges, showComments,
    commentComposer,
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
    flushTextDrafts,
    endTextCommitBurst,
    activeFormFieldEditKey: { get value() { return activeFormFieldEditKey }, set value(v) { activeFormFieldEditKey = v } },
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
  const {
    dispatchEditorTransaction,
    applyModelChange,
    endTextCommitBurst: endBurstInTransaction,
    keepsTextCommitBurst,
  } = createEditorTransaction(ctx)
  endTextCommitBurstImpl = endBurstInTransaction

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
  const clp = createEditorClipboard(ctx, {
    insertTextAtSelection: txt.insertTextAtSelection,
    insertClipboardParagraphsAtSelection: txt.insertClipboardParagraphsAtSelection,
  })

  // 11. editor-import-export
  const ie = createEditorImportExport(ctx)

  // 12. editor-comments (authoring: add/reply/resolve/delete)
  const cmt = createEditorComments(ctx, applyModelChange, {
    author: options.commentAuthor ?? officeMessagesEn.docx.editor.commentAuthor,
    initials: options.commentInitials,
    clock: options.commentClock ?? (() => new Date()),
  })

  // 13. editor-search-replace (find & replace over search results)
  const repl = createEditorSearchReplace(ctx, applyModelChange)

  // ═══════════════════════════════════════════════════════════════════
  // Thin wrappers (state accessors + trivial dispatch calls)
  // ═══════════════════════════════════════════════════════════════════

  let activeImportAbortController: AbortController | undefined
  let pendingExportModelTransformer: ((model: DocModel) => DocModel) | undefined
  let activeFormFieldEditKey: string | undefined

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
  const registerTextDraftFlusher = (flusher: () => void): (() => void) => {
    textDraftFlushers.add(flusher)
    return () => { textDraftFlushers.delete(flusher) }
  }
  const clearPendingRunStyle = (): void => {
    if (pendingRunStyleSnapshot === undefined && pendingRunStyle.value === undefined) return
    pendingRunStyleSnapshot = undefined
    pendingRunStyle.value = undefined
  }

  // ── tracked-change decisions ─────────────────────────────────────
  // Every decision goes through the transaction dispatcher, so accepting or
  // rejecting a revision is a single undoable step and shares the history
  // budget with ordinary edits.
  const decideTrackedChanges = (
    decision: DocxTrackedChangeDecision,
    scope?: DocxTrackedChangeDecisionScope
  ): boolean => {
    let applied = false
    dispatchEditorTransaction((current) => {
      const result = applyTrackedChangeDecision(current.model, decision, scope)
      if (result.model === current.model) return undefined
      applied = true
      const count = result.resolvedChangeIds.length
      const status = decision === "accept"
        ? (count === 1
          ? ctx.statusMessage("acceptedChange")
          : ctx.statusMessage("acceptedChanges", { count }))
        : (count === 1
          ? ctx.statusMessage("rejectedChange")
          : ctx.statusMessage("rejectedChanges", { count }))
      const keptFormatting =
        result.clearedOnlyChangeIds.length > 0
          ? (result.clearedOnlyChangeIds.length === 1
            ? ctx.statusMessage("formattingClearedOne")
            : ctx.statusMessage("formattingCleared", { count: result.clearedOnlyChangeIds.length }))
          : ""
      return {
        model: result.model,
        status: `${status}${keptFormatting}`,
        clearSelectedFormField: true,
      }
    })
    return applied
  }

  const acceptTrackedChange = (changeId: string): boolean =>
    decideTrackedChanges("accept", { changeIds: [changeId] })
  const rejectTrackedChange = (changeId: string): boolean =>
    decideTrackedChanges("reject", { changeIds: [changeId] })
  const acceptTrackedChangesInTableRow = (tableIndex: number, rowIndex: number): boolean =>
    decideTrackedChanges("accept", {
      nodeIndex: tableIndex,
      tableRowRange: { startRowIndex: rowIndex, endRowIndex: rowIndex + 1 },
    })
  const rejectTrackedChangesInTableRow = (tableIndex: number, rowIndex: number): boolean =>
    decideTrackedChanges("reject", {
      nodeIndex: tableIndex,
      tableRowRange: { startRowIndex: rowIndex, endRowIndex: rowIndex + 1 },
    })
  const acceptAllTrackedChanges = (): boolean => decideTrackedChanges("accept")
  const rejectAllTrackedChanges = (): boolean => decideTrackedChanges("reject")

  // ── page setup / header + footer authoring ────────────────────────
  // Every command is a single undoable step: the edited section lives in
  // `metadata.sections`, which the history snapshot already covers.
  const setPageMargins = (
    sectionIndex: number,
    margins: DocxPageMarginsUpdate,
    options?: DocxPageSetupCommandOptions
  ): boolean =>
    applyModelChange(
      (current) => applySectionPageMargins(current, sectionIndex, margins),
      ctx.statusMessage("pageMarginsUpdated"),
      { pushHistory: options?.pushHistory }
    )

  const setPageSize = (
    sectionIndex: number,
    pageSize: DocxPageSizeUpdate,
    options?: DocxPageSetupCommandOptions
  ): boolean =>
    applyModelChange(
      (current) => applySectionPageSize(current, sectionIndex, pageSize),
      ctx.statusMessage("pageSizeUpdated"),
      { pushHistory: options?.pushHistory }
    )

  const addHeaderFooter = (
    sectionIndex: number,
    kind: DocxSectionRegion,
    options?: AddHeaderFooterOptions
  ): boolean => {
    const status = ctx.statusMessage(kind === "header" ? "headerAdded" : "footerAdded")
    return applyModelChange(
      (current) => addSectionHeaderFooter(current, sectionIndex, kind, options),
      status
    )
  }

  // Browser-originated ranges are already reflected in the DOM. Keep the
  // synchronous snapshot current without issuing a DOM restore request that
  // would fight the user's live selection.
  const setActiveTextRange = (
    range?: DocxTextRange,
    options?: { preservePendingRunStyle?: boolean },
  ): void => {
    const nextRange = normalizeEditorCursorStateForModel(
      modelSnapshot,
      selectionSnapshot,
      range
    ).activeTextRange
    if (sameTextRange(activeTextRangeSnapshot, nextRange)) return
    // A real selection/caret move discards a pending caret style. Typing
    // passes preservePendingRunStyle because the input event extends the range
    // before the draft is committed on blur/export.
    if (!options?.preservePendingRunStyle) clearPendingRunStyle()
    // A selection change ends the typing burst, except for the caret following
    // the text that the burst just committed (which stays in the same
    // paragraph/cell and is collapsed).
    if (!keepsTextCommitBurst(nextRange)) endTextCommitBurst()
    activeTextRangeSnapshot = cloneTextRange(nextRange)
    activeTextRange.value = nextRange
  }
  const selectParagraph = (nodeIndex: number): void => {
    clearPendingRunStyle()
    dispatchEditorTransaction(() => ({
      selection: { kind: "paragraph", nodeIndex },
      activeTextRange: undefined,
    }))
  }
  const selectTableCell = (tableIndex: number, rowIndex: number, cellIndex: number): void => {
    clearPendingRunStyle()
    dispatchEditorTransaction(() => ({
      selection: { kind: "table-cell", tableIndex, rowIndex, cellIndex },
      activeTextRange: undefined,
    }))
  }
  const selectTableCellRange = (
    tableIndex: number,
    range: DocxTableCellSelectionRange,
    tableRuntimeKey?: string
  ): void => {
    const table = model.value.nodes[tableIndex]
    const normalized =
      !tableRuntimeKey && table?.type === "table"
        ? normalizeTableCellSelectionRange(table, range)
        : cloneTableCellSelectionRange(range)
    if (!normalized) return
    normalized.tableIndex = tableIndex
    clearPendingRunStyle()
    dispatchEditorTransaction(() => ({
      selection: {
        kind: "table-cell",
        tableIndex,
        rowIndex: normalized.focusRowIndex,
        cellIndex: normalized.focusCellIndex,
        ...(tableRuntimeKey ? { tableRuntimeKey } : undefined),
        cellRange: normalized,
      },
      activeTextRange: undefined,
    }))
  }
  /**
   * Selects every cell of the table as one rectangular range. The anchor is
   * the top-left cell and the focus is the bottom-right cell; normalization
   * expands the grid span over merged cells.
   */
  const selectWholeTable = (tableIndex?: number): void => {
    const selection = selectionSnapshot
    const targetTableIndex =
      tableIndex ??
      (selection.kind === "table-cell" ? selection.tableIndex : undefined)
    if (targetTableIndex === undefined) return
    const table = model.value.nodes[targetTableIndex]
    if (!table || table.type !== "table" || table.rows.length === 0) return
    let focusRowIndex = table.rows.length - 1
    while (focusRowIndex >= 0 && (table.rows[focusRowIndex]?.cells.length ?? 0) === 0) {
      focusRowIndex -= 1
    }
    if (focusRowIndex < 0) return
    let anchorRowIndex = 0
    while (anchorRowIndex <= focusRowIndex && (table.rows[anchorRowIndex]?.cells.length ?? 0) === 0) {
      anchorRowIndex += 1
    }
    if (anchorRowIndex > focusRowIndex) return
    const columnCount = table.rows.reduce(
      (max, row) =>
        Math.max(
          max,
          row.cells.reduce(
            (count, cell) => count + (cell.style?.gridSpan ?? 1),
            0
          )
        ),
      0
    )
    selectTableCellRange(targetTableIndex, {
      tableIndex: targetTableIndex,
      anchorRowIndex,
      anchorCellIndex: 0,
      focusRowIndex,
      focusCellIndex: (table.rows[focusRowIndex]?.cells.length ?? 1) - 1,
      gridRange: {
        startColumnIndex: 0,
        endColumnIndex: Math.max(0, columnCount - 1),
      },
    })
  }
  /** Collapses a multi-cell range to its anchor cell. Returns true if it did. */
  const collapseTableCellSelection = (): boolean => {
    const selection = selectionSnapshot
    if (
      selection.kind !== "table-cell" ||
      !selection.cellRange ||
      !tableCellSelectionRangeHasMultipleCells(selection.cellRange)
    ) {
      return false
    }
    const range = selection.cellRange
    clearPendingRunStyle()
    dispatchEditorTransaction(() => ({
      selection: {
        kind: "table-cell",
        tableIndex: selection.tableIndex,
        rowIndex: range.anchorRowIndex,
        cellIndex: range.anchorCellIndex,
        ...(selection.tableRuntimeKey
          ? { tableRuntimeKey: selection.tableRuntimeKey }
          : undefined),
      },
      activeTextRange: undefined,
    }))
    return true
  }
  const selectTableCellByRuntimeKey = (
    tableRuntimeKey: string,
    rowIndex: number,
    cellIndex: number
  ): void => {
    const location = parseEmbeddedTableRuntimeKey(tableRuntimeKey)
    if (!location) return
    dispatchEditorTransaction(() => ({
      selection: {
        kind: "table-cell",
        tableIndex: location.hostTableIndex,
        rowIndex,
        cellIndex,
        tableRuntimeKey,
      },
      activeTextRange: undefined,
    }))
  }

  // ═══════════════════════════════════════════════════════════════════
  // Controller assembly
  // ═══════════════════════════════════════════════════════════════════

  return {
    get model() { return model.value },
    get isDirty() { return isDirty.value },
    get viewOnlyReason() {
      return isOpenDocumentModel(model.value) ? statusMessage("viewOnly") : undefined
    },
    get documentLoadNonce() { return documentLoadNonce.value },
    get fileName() { return fileName.value },
    get status() { return status.value },
    get importError() { return importError.value },
    get isImporting() { return isImporting.value },
    get documentTheme() { return documentTheme.value },
    get selection() { return selection.value },
    get selectedTableCellRange() {
      return selection.value.kind === "table-cell"
        ? selection.value.cellRange
        : undefined
    },
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
    get trackedChangeCount() { return trackedChangeCount.value },
    get showTrackedChanges() { return showTrackedChanges.value },
    get revisionDisplay() { return revisionDisplay.value },
    setRevisionDisplay,
    get comments() { return comments.value },
    get showComments() { return showComments.value },
    get commentComposer() { return commentComposer.value },
    get currentPage() { return paginationInfo.value.currentPage },
    get totalPages() { return paginationInfo.value.totalPages },
    get hasUnorderedList() { return hasUnorderedList.value },
    get hasOrderedList() { return hasOrderedList.value },
    get canUndo() { return canUndo.value },
    get canRedo() { return canRedo.value },
    registerPendingExportModelTransformer: ie.registerPendingExportModelTransformer,
    registerTextDraftFlusher,
    flushTextDrafts,
    endTextCommitBurst,
    setStatus,
    setDocumentTheme: setDocumentThemeFn,
    setShowTrackedChanges: setShowTrackedChangesFn,
    setShowComments: setShowCommentsFn,
    beginComment: cmt.beginComment,
    beginCommentReply: cmt.beginCommentReply,
    cancelCommentComposer: cmt.cancelCommentComposer,
    submitCommentComposer: cmt.submitCommentComposer,
    addComment: cmt.addComment,
    replyToComment: cmt.replyToComment,
    resolveComment: cmt.resolveComment,
    deleteComment: cmt.deleteComment,
    syncPaginationInfo,
    toggleShowTrackedChanges: toggleShowTrackedChangesFn,
    toggleShowComments: toggleShowCommentsFn,
    acceptTrackedChange,
    rejectTrackedChange,
    acceptTrackedChangesInTableRow,
    rejectTrackedChangesInTableRow,
    acceptAllTrackedChanges,
    rejectAllTrackedChanges,
    importDocxFile: ie.importDocxFile,
    newDocument: ie.newDocument,
    exportDocx: ie.exportDocx,
    getDocxBytes: ie.getDocxBytes,
    getDocxBlob: ie.getDocxBlob,
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
    endFormFieldEdit: ff.endFormFieldEdit,
    updateFormFieldWidget: ff.updateFormFieldWidget,
    updateFormField: ff.updateFormField,
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
    bringToFront: img.bringToFront,
    bringForward: img.bringForward,
    sendBackward: img.sendBackward,
    sendToBack: img.sendToBack,
    moveFloatingImage: img.moveFloatingImage,
    moveSectionFloatingImage: img.moveSectionFloatingImage,
    moveParagraphDropCap: img.moveParagraphDropCap,
    setParagraphDropCapFontSizePt: img.setParagraphDropCapFontSizePt,
    setParagraphDropCapText: img.setParagraphDropCapText,
    moveImage: img.moveImage,
    setActiveTextRange,
    selectParagraph,
    selectTableCell,
    selectTableCellRange,
    selectWholeTable,
    collapseTableCellSelection,
    selectTableCellByRuntimeKey,
    clearTableCellContents: tbl.clearTableCellContents,
    clearTableSelection: tbl.clearTableSelection,
    setTableColumnWidths: tbl.setTableColumnWidths,
    setTableRowHeight: tbl.setTableRowHeight,
    setTableFloating: tbl.setTableFloating,
    insertTableRow: tbl.insertTableRow,
    insertTableColumn: tbl.insertTableColumn,
    deleteTableRow: tbl.deleteTableRow,
    deleteTableColumn: tbl.deleteTableColumn,
    deleteTable: tbl.deleteTable,
    mergeTableCells: tbl.mergeTableCells,
    splitTableCell: tbl.splitTableCell,
    moveTable: tbl.moveTable,
    moveBodyNode: tbl.moveBodyNode,
    moveEmbeddedTableToBody: tbl.moveEmbeddedTableToBody,
    replaceExpandedSelection: txt.replaceExpandedSelection,
    deleteExpandedSelection: txt.deleteExpandedSelection,
    replaceMatch: repl.replaceMatch,
    replaceAll: repl.replaceAll,
    commitParagraphText: txt.commitParagraphText,
    commitTableCellText: txt.commitTableCellText,
    commitTableCellTextByRuntimeKey: txt.commitTableCellTextByRuntimeKey,
    commitTableCellParagraphTextByRuntimeKey: txt.commitTableCellParagraphTextByRuntimeKey,
    commitTableCellParagraphTextRecursive: txt.commitTableCellParagraphTextRecursive,
    commitSectionParagraphText: txt.commitSectionParagraphText,
    setPageMargins,
    setPageSize,
    addHeaderFooter,
    writeSelectionToClipboardData: clp.writeSelectionToClipboardData,
    pasteClipboardData: async (sources) => clp.pasteClipboardData(sources),
    copy: clp.copy,
    paste: clp.paste,
  }
}
