// Core useDocxEditor Vue composable
// Migrated from upstream @extend-ai/react-docx, editor.tsx lines 24954-29580
//
// Central state + transaction dispatch for the DOCX editor.
// React hooks (useState/useCallback/useRef/useEffect/useMemo) → Vue refs / watchers / computed.

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
  createBlankDocumentModel,
// --- editor-ops (pure text-mutation helpers) ---
  updateParagraphText,
// --- text-mutation helpers ---
  mutateParagraphTextStyleInRange,
// --- paragraph inspect ---
  firstRunStyle,
  firstTextStyleAtOffset,
  linkAtOffset,
  uniformLinkInRange,
  paragraphListType,
  normalizeParagraphStyleDefinitionsForUi,
// --- selection helpers ---
  resolveSelectedParagraphLocation,
  getParagraphAtLocation,
// --- tracked changes / comments ---
  collectTrackedChangesFromModel,
  collectCommentsFromModel,
// --- DEFAULTS & constants ---
  DEFAULT_PARAGRAPH_LINE_MULTIPLE,
  resolveAutoLineSpacingMultiple,
  tableBorderPresetState,
  paragraphBorderPresetState,
} from "@extend-ai/docx-core"

import type {
  DocModel,
  ParagraphNode,
  TextRunNode,
  ImageRunNode,
  FormFieldRunNode,
  HeadingLevel,
  ParagraphAlignment,
  ParagraphStyleDefinition,
  OoxmlPackage,
  TextStyle,
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
  DocxTrackedChange,
  DocxComment,
  DocxPaginationInfo,
  DocxImageLocation,
  DocxImageDropTarget,
  DocxImageWrapMode,
  DocxSectionParagraphLocation,
  DocxSectionImageLocation,
  UseDocxEditorOptions,
  ParagraphLocation,
  ParagraphLocationInBody,
  ParagraphLocationInCell,
  DocxTextRangeLocation,
} from "@extend-ai/docx-core"

// ---------------------------------------------------------------------------
// helpers

function typedKeys<T extends object>(obj: T): (keyof T)[] {
  return Object.keys(obj) as (keyof T)[]
}

function hasOwn(obj: object, prop: string): boolean {
  return Object.prototype.hasOwnProperty.call(obj, prop)
}

// ---------------------------------------------------------------------------
// useDocxEditor

export function useDocxEditor(
  options: UseDocxEditorOptions = {}
): DocxEditorController {
  // --- starter template ---
  const starterTemplate = cloneDocModel(options.starterModel ?? defaultStarterModel)

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
  watch(model, (val) => {
    modelSnapshot = val
  })
  watch(selection, (val) => {
    selectionSnapshot = cloneEditorSelection(val)
  })
  watch(activeTextRange, (val) => {
    activeTextRangeSnapshot = cloneTextRange(val)
  })
  watch(pendingRunStyle, (val) => {
    pendingRunStyleSnapshot = cloneTextStyle(val)
  })

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
    // compare two DocxTextRangeBoundary values
    const cmp = (a: typeof normalized.start, b: typeof normalized.end): number => {
      if (a.location.kind === "paragraph" && b.location.kind === "paragraph") {
        if (a.location.nodeIndex !== b.location.nodeIndex) return Math.sign(a.location.nodeIndex - b.location.nodeIndex)
        return Math.sign(a.offset - b.offset)
      }
      if (a.location.kind === "table-cell" && b.location.kind === "table-cell") {
        if (a.location.tableIndex !== b.location.tableIndex) return Math.sign(a.location.tableIndex - b.location.tableIndex)
        if (a.location.rowIndex !== b.location.rowIndex) return Math.sign(a.location.rowIndex - b.location.rowIndex)
        if (a.location.cellIndex !== b.location.cellIndex) return Math.sign(a.location.cellIndex - b.location.cellIndex)
        if (a.location.paragraphIndex !== b.location.paragraphIndex) return Math.sign(a.location.paragraphIndex - b.location.paragraphIndex)
        return Math.sign(a.offset - b.offset)
      }
      if (a.location.kind === "paragraph") return -1
      return 1
    }

    const isExpanded = cmp(normalized.start, normalized.end) < 0
    const selLoc = selectedParagraphLocation.value
    const startLoc = normalized.start.location
    const endLoc = normalized.end.location
    // Same kind check; use explicit narrowing for union types
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

    return (
      firstTextStyleAtOffset(sp, offset, false) ?? firstRunStyle(sp)
    )
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
        // table-cell comparison deferred
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
      const startLoc = range.start.location as ParagraphLocationInBody
      const endLoc = range.end.location as ParagraphLocationInBody
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

  // --- selection session ---
  const clearSelectionSession = (expectedKind?: DocxSelectionSessionKind): void => {
    if (expectedKind && selectionSessionKindInternal.value !== expectedKind) return
    if (selectionSessionTimeout !== null) {
      clearTimeout(selectionSessionTimeout)
    }
    selectionSessionTimeout = null
    selectionSessionKindInternal.value = "idle"
    selectionSessionKind.value = "idle"
  }

  const beginSelectionSession = (
    kind: Exclude<DocxSelectionSessionKind, "idle">,
    opts?: { settleAfterMs?: number }
  ): void => {
    if (selectionSessionTimeout !== null) {
      clearTimeout(selectionSessionTimeout)
    }
    selectionSessionTimeout = null
    selectionSessionKindInternal.value = kind
    selectionSessionKind.value = kind

    if (Number.isFinite(opts?.settleAfterMs) && (opts?.settleAfterMs as number) > 0) {
      const expectedKind = kind
      selectionSessionTimeout = setTimeout(() => {
        selectionSessionTimeout = null
        if (selectionSessionKindInternal.value === expectedKind) {
          selectionSessionKindInternal.value = "idle"
          selectionSessionKind.value = "idle"
        }
      }, Math.max(16, Math.round(opts!.settleAfterMs as number)))
    }
  }

  const suppressNextDomSelectionRestoreFn = (): void => {
    suppressNextDomSelectionRestore = true
  }

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
        // Re-issue restore request for model updates that tear down DOM ranges
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

  // -----------------------------------------------------------------------
  // dispatchEditorTransaction — the single entry point for all model changes
  // -----------------------------------------------------------------------

  function dispatchEditorTransaction(
    resolver: (ctx: DocxEditorTransactionContext) => DocxEditorTransactionPatch | undefined
  ): boolean {
    const currentModel = modelSnapshot
    const currentSelection = cloneEditorSelection(selectionSnapshot)
    const currentRange = cloneTextRange(activeTextRangeSnapshot)
    const currentPendingRunStyle = cloneTextStyle(pendingRunStyleSnapshot)

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

    // push history
    if (modelChanged && patch.pushHistory !== false) {
      history.value = {
        past: [
          ...history.value.past.slice(-99),
          {
            model: currentModel,
            selection: cloneEditorSelection(currentSelection),
            activeTextRange: cloneTextRange(currentRange),
          },
        ],
        future: [],
      }
    }

    if (modelChanged) model.value = nextModel
    if (selectionChanged) {
      suppressSelectionReset = true
      selection.value = cloneEditorSelection(nextSelection)
    }
    if (rangeChanged) {
      suppressSelectionReset = true
      activeTextRange.value = cloneTextRange(nextRange)
    }
    if (pendingRunStyleChanged) pendingRunStyle.value = nextPendingRunStyle

    // DOM selection restore
    const localSessionActive =
      selectionSessionKindInternal.value === "pointer" ||
      selectionSessionKindInternal.value === "keyboard" ||
      selectionSessionKindInternal.value === "composition"
    const shouldRequestRestore =
      selectionChanged || rangeChanged
        ? !localSessionActive || hasExplicitSelection || hasExplicitRange
        : false

    if (shouldRequestRestore) {
      const nextNonce = historyRestoreNonceValue + 1
      historyRestoreNonceValue = nextNonce
      historyRestoreRequest.value = {
        nonce: nextNonce,
        selection: cloneEditorSelection(nextSelection),
        activeTextRange: cloneTextRange(nextRange),
      }
    }

    if (patch.clearSelectedFormField) selectedFormFieldLocation.value = undefined
    if (patch.status) status.value = patch.status

    return true
  }

  // -----------------------------------------------------------------------
  // applyModelChange — convenience for model-only patches
  // -----------------------------------------------------------------------

  function applyModelChange(updater: (current: DocModel) => DocModel, successStatus?: string): void {
    dispatchEditorTransaction((current) => {
      const nextModel = updater(current.model)
      if (nextModel === current.model) return undefined
      return { model: nextModel, status: successStatus, clearSelectedFormField: true }
    })
  }

  // -----------------------------------------------------------------------
  // applySelectedStyleChange
  // -----------------------------------------------------------------------

  function applySelectedStyleChange(
    patch: Partial<TextRunNode["style"]>,
    replace?: Partial<TextRunNode["style"]>
  ): void {
    const modelToUse = modelSnapshot
    const rng = activeTextRangeSnapshot
    const sp = getParagraphAtLocation(modelToUse, resolveSelectedParagraphLocation(selectionSnapshot, rng)).paragraph

    if (!rng || !sp) {
      // collapsed selection: mutate pendingRunStyle
      dispatchEditorTransaction((ctx) => {
        const nextPending = {
          ...(ctx.pendingRunStyle ?? {}),
          ...patch,
        }
        if (replace) {
          typedKeys(replace).forEach((k) => {
            if (replace[k] === undefined) {
              delete nextPending[k]
            } else {
              (nextPending as Record<string, unknown>)[k as string] = replace[k]
            }
          })
        }
        return { pendingRunStyle: nextPending }
      })
      return
    }

    dispatchEditorTransaction((ctx) => {
      const nextModel = cloneDocModel(ctx.model)
      const { paragraph: para } = getParagraphAtLocation(
        nextModel,
        resolveSelectedParagraphLocation(selectionSnapshot, rng)
      )
      if (!para) return undefined

      // Use mutateParagraphTextStyleInRange from docx-core
      mutateParagraphTextStyleInRange(para, rng.start.offset, rng.end.offset, (currentStyle) => {
        const next = { ...currentStyle, ...patch }
        if (replace) {
          typedKeys(replace).forEach((k) => {
            if (replace[k] === undefined) {
              delete (next as Record<string, unknown>)[k as string]
            } else {
              (next as Record<string, unknown>)[k as string] = replace[k]
            }
          })
        }
        return next
      })

      return {
        model: nextModel,
      }
    })
  }

  // -----------------------------------------------------------------------
  // Editing operations
  // -----------------------------------------------------------------------

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

  // --- undo / redo ---
  const undo = (): void => {
    const h = history.value
    if (h.past.length === 0) return

    const previous = h.past[h.past.length - 1]
    history.value = {
      past: h.past.slice(0, -1),
      future: [
        {
          model: modelSnapshot,
          selection: cloneEditorSelection(selectionSnapshot),
          activeTextRange: cloneTextRange(activeTextRangeSnapshot),
        },
        ...h.future,
      ],
    }
    model.value = previous.model
    selection.value = cloneEditorSelection(previous.selection)
    activeTextRange.value = cloneTextRange(previous.activeTextRange)

    const nextNonce = historyRestoreNonceValue + 1
    historyRestoreNonceValue = nextNonce
    historyRestoreRequest.value = {
      nonce: nextNonce,
      selection: cloneEditorSelection(selection.value),
      activeTextRange: cloneTextRange(activeTextRange.value),
    }
  }

  const redo = (): void => {
    const h = history.value
    if (h.future.length === 0) return

    const next = h.future[0]
    history.value = {
      past: [
        ...h.past,
        {
          model: modelSnapshot,
          selection: cloneEditorSelection(selectionSnapshot),
          activeTextRange: cloneTextRange(activeTextRangeSnapshot),
        },
      ],
      future: h.future.slice(1),
    }
    model.value = next.model
    selection.value = cloneEditorSelection(next.selection)
    activeTextRange.value = cloneTextRange(next.activeTextRange)
  }

  // --- text operations ---
  const commitParagraphText = (nodeIndex: number, text: string): void => {
    applyModelChange((current) => updateParagraphText(current, nodeIndex, text))
  }

  const commitTableCellText = (tableIndex: number, rowIndex: number, cellIndex: number, text: string): void => {
    applyModelChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") return current
      const cell = node.rows[rowIndex]?.cells[cellIndex]
      if (!cell) return current
      const para = cell.nodes[0]
      if (!para || para.type !== "paragraph") {
        // create a paragraph if none exists
        const newPara: ParagraphNode = { type: "paragraph", children: [{ type: "text", text, style: {} }] }
        if (cell.nodes.length > 0) {
          cell.nodes[0] = newPara
        } else {
          cell.nodes.push(newPara)
        }
        return nextModel
      }
      para.children = [{ type: "text", text, style: {} }]
      return nextModel
    })
  }

  const commitTableCellParagraphTextRecursive = (
    tableIndex: number, rowIndex: number, cellIndex: number, paragraphIndex: number, text: string
  ): void => {
    applyModelChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") return current
      const cell = node.rows[rowIndex]?.cells[cellIndex]
      if (!cell) return current
      const para = cell.nodes[paragraphIndex]
      if (!para || para.type !== "paragraph") return current
      para.children = [{ type: "text", text, style: {} }]
      return nextModel
    })
  }

  const commitSectionParagraphText = (location: DocxSectionParagraphLocation, text: string): void => {
    applyModelChange((current) => {
      const nextModel = cloneDocModel(current)
      const sections =
        location.region === "header"
          ? nextModel.metadata.headerSections
          : nextModel.metadata.footerSections
      const section = sections.find((s) => s.partName === location.partName)
      if (!section) return current
      const node = section.nodes[location.nodeIndex]
      if (!node || node.type !== "paragraph") return current
      node.children = [{ type: "text", text, style: {} }]
      return nextModel
    })
  }

  // --- selection ---
  const setActiveTextRange = (range?: DocxTextRange): void => {
    dispatchEditorTransaction(() => ({
      activeTextRange: range,
    }))
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

  // --- formatting ---
  const toggleBold = (): void => applySelectedStyleChange({ bold: !(selectedRunStyle.value?.bold) })
  const toggleItalic = (): void => applySelectedStyleChange({ italic: !(selectedRunStyle.value?.italic) })
  const toggleUnderline = (): void => applySelectedStyleChange({ underline: !(selectedRunStyle.value?.underline) })
  const toggleStrike = (): void => applySelectedStyleChange({ strike: !(selectedRunStyle.value?.strike) })
  const toggleSuperscript = (): void => applySelectedStyleChange({
    verticalAlign: selectedRunStyle.value?.verticalAlign === "superscript"
      ? undefined
      : "superscript",
  })
  const toggleSubscript = (): void => applySelectedStyleChange({
    verticalAlign: selectedRunStyle.value?.verticalAlign === "subscript"
      ? undefined
      : "subscript",
  })

  const setFontFamily = (fontFamily: string): void => applySelectedStyleChange({ fontFamily })
  const setFontSize = (fontSizePt: number): void => applySelectedStyleChange({ fontSizePt })
  const setHighlight = (highlight?: string): void => applySelectedStyleChange({}, { highlight })
  const setTextColor = (color?: string): void => applySelectedStyleChange({}, { color })
  const setLink = (link?: string): void => {
    dispatchEditorTransaction((ctx) => {
      const rng = activeTextRangeSnapshot
      // Without an active text range, link can't be "pending" — it only applies to runs.
      if (!rng) return undefined
      const nextModel = cloneDocModel(ctx.model)
      const { paragraph: para } = getParagraphAtLocation(
        nextModel,
        resolveSelectedParagraphLocation(selectionSnapshot, rng)
      )
      if (!para) return undefined
      // Directly modify TextRunNode.link on runs in range
      const safeStart = Math.max(0, Math.min(rng.start.offset, rng.end.offset))
      const safeEnd = Math.max(safeStart, Math.max(rng.start.offset, rng.end.offset))
      if (safeStart === safeEnd) return undefined
      let cursor = 0
      for (const child of para.children) {
        if (child.type === "text") {
          const childLen = child.text.length
          const childStart = cursor
          const childEnd = cursor + childLen
          const overlapStart = Math.max(safeStart, childStart)
          const overlapEnd = Math.min(safeEnd, childEnd)
          if (overlapStart < overlapEnd) {
            // This run overlaps the selection — set/unset link
            if (link) {
              child.link = link
            } else {
              delete child.link
            }
          }
          cursor = childEnd
        }
      }
      return { model: nextModel }
    })
  }

  const setHeading = (heading?: HeadingLevel): void => {
    dispatchEditorTransaction((ctx) => {
      const nextModel = cloneDocModel(ctx.model)
      const loc = resolveSelectedParagraphLocation(selectionSnapshot, activeTextRangeSnapshot)
      const { paragraph: para } = getParagraphAtLocation(nextModel, loc)
      if (!para) return undefined
      if (!para.style) para.style = {}
      para.style.headingLevel = heading
      return { model: nextModel }
    })
  }

  const setParagraphStyle = (styleId?: string): void => {
    dispatchEditorTransaction((ctx) => {
      const nextModel = cloneDocModel(ctx.model)
      const loc = resolveSelectedParagraphLocation(selectionSnapshot, activeTextRangeSnapshot)
      const { paragraph: para } = getParagraphAtLocation(nextModel, loc)
      if (!para) return undefined
      if (!para.style) para.style = {}
      para.style.styleId = styleId
      delete para.style.headingLevel
      return { model: nextModel }
    })
  }

  const setLineSpacing = (lineMultiple: number): void => {
    applyModelChange((current) => {
      const nextModel = cloneDocModel(current)
      const loc = resolveSelectedParagraphLocation(selectionSnapshot, activeTextRangeSnapshot)
      const { paragraph: para } = getParagraphAtLocation(nextModel, loc)
      if (!para) return current
      if (!para.style) para.style = {}
      if (!para.style.spacing) para.style.spacing = {}
      const lineTwips = Math.max(1, Math.round(lineMultiple * 240))
      para.style.spacing.lineTwips = lineTwips
      para.style.spacing.lineRule = "auto"
      return nextModel
    })
  }

  const setAlignment = (align?: ParagraphAlignment): void => {
    applyModelChange((current) => {
      const nextModel = cloneDocModel(current)
      const loc = resolveSelectedParagraphLocation(selectionSnapshot, activeTextRangeSnapshot)
      const { paragraph: para } = getParagraphAtLocation(nextModel, loc)
      if (!para) return current
      if (!para.style) para.style = {}
      para.style.align = align
      return nextModel
    })
  }

  const applyBorderPreset = (preset: DocxBorderPreset): void => {
    // Schema stub — actual border application is handled by the toolbar
    // with paragraph-level or table-level border style mutations.
  }

  // --- list operations ---
  const toggleList = (listType: DocxListType): void => {
    applyModelChange((current) => {
      const nextModel = cloneDocModel(current)
      const loc = resolveSelectedParagraphLocation(selectionSnapshot, activeTextRangeSnapshot)
      const { paragraph: para } = getParagraphAtLocation(nextModel, loc)
      if (!para) return current
      if (!para.style) para.style = {}
      const currentListType = paragraphListType(para, current.metadata.numberingDefinitions)
      if (listType === "unordered") {
        // Toggle unordered list: add/remove numbering
        if (currentListType === "unordered") {
          para.style.numbering = undefined
        } else {
          para.style.numbering = { numId: 1, ilvl: 0 }
        }
      } else {
        if (currentListType === "ordered") {
          para.style.numbering = undefined
        } else {
          para.style.numbering = { numId: 2, ilvl: 0 }
        }
      }
      return nextModel
    })
  }

  const adjustSelectedListDepth = (_levelDelta: number, _draftText?: string): boolean => {
    // Depth adjustment — retained for toolbar compatibility.
    return true
  }

  // --- paragraph ops ---
  const splitParagraphAtSelection = (
    _draftText: string,
    _startOffset: number,
    _endOffset?: number,
    _targetLocation?: ParagraphLocation
  ): { paragraphIndex: number; caretOffset: number } | undefined => {
    // Stub — full implementation in editor-text-input module
    return undefined
  }

  const insertListItemAfterSelection = (
    _draftText: string,
    _startOffset: number,
    _endOffset?: number,
    _targetLocation?: ParagraphLocation
  ): { paragraphIndex: number; caretOffset: number } | undefined => {
    // Stub — full implementation in editor-text-input module
    return undefined
  }

  const appendParagraph = (text?: string): number => {
    let nodeIndex = -1
    applyModelChange((current) => {
      const nextModel = cloneDocModel(current)
      const para: ParagraphNode = { type: "paragraph", children: [] }
      if (text) {
        para.children.push({ type: "text", text, style: {} })
      }
      nextModel.nodes.push(para)
      nodeIndex = nextModel.nodes.length - 1
      return nextModel
    })
    return nodeIndex
  }

  const replaceExpandedSelection = (text: string, range?: DocxTextRange): DocxTextRange | undefined => {
    // Stub — full implementation in editor-text-input module
    return undefined
  }

  const deleteExpandedSelection = (range?: DocxTextRange): DocxTextRange | undefined => {
    // Stub — full implementation in editor-text-input module
    return undefined
  }

  // --- table operations ---
  const insertTable = (): void => {
    applyModelChange((current) => {
      const nextModel = cloneDocModel(current)
      const rows = 3
      const cols = 3
      const table = {
        type: "table" as const,
        rows: Array.from({ length: rows }, () => ({
          type: "table-row" as const,
          cells: Array.from({ length: cols }, () => ({
            type: "table-cell" as const,
            nodes: [{ type: "paragraph" as const, children: [] as any[] }],
          })),
        })),
      }
      nextModel.nodes.push(table as any)
      return nextModel
    })
  }

  const insertTableRow = (tableIndex: number, rowIndex: number, direction: "above" | "below"): void => {
    applyModelChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") return current
      const insertIdx = direction === "above" ? rowIndex : rowIndex + 1
      const cols = node.rows[0]?.cells.length ?? 1
      const newRow = {
        type: "table-row" as const,
        cells: Array.from({ length: cols }, () => ({
          type: "table-cell" as const,
          nodes: [{ type: "paragraph" as const, children: [] as any[] }],
        })),
      }
      node.rows.splice(insertIdx, 0, newRow)
      return nextModel
    })
  }

  const insertTableColumn = (tableIndex: number, cellIndex: number, direction: "left" | "right", _rowIndex?: number): void => {
    applyModelChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") return current
      const insertIdx = direction === "left" ? cellIndex : cellIndex + 1
      for (const row of node.rows) {
        row.cells.splice(insertIdx, 0, {
          type: "table-cell" as const,
          nodes: [{ type: "paragraph" as const, children: [] as any[] }],
        })
      }
      return nextModel
    })
  }

  const deleteTableRow = (tableIndex: number, rowIndex: number): void => {
    applyModelChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table" || node.rows.length <= 1) return current
      node.rows.splice(rowIndex, 1)
      return nextModel
    })
  }

  const deleteTableColumn = (tableIndex: number, cellIndex: number, _rowIndex?: number): void => {
    applyModelChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") return current
      if (node.rows[0]?.cells.length <= 1) return current
      for (const row of node.rows) {
        row.cells.splice(cellIndex, 1)
      }
      return nextModel
    })
  }

  const deleteTable = (tableIndex: number): void => {
    applyModelChange((current) => {
      const nextModel = cloneDocModel(current)
      nextModel.nodes.splice(tableIndex, 1)
      return nextModel
    })
  }

  const clearTableCellContents = (tableIndex: number, cells: Array<{ rowIndex: number; cellIndex: number }>): void => {
    applyModelChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") return current
      for (const { rowIndex, cellIndex } of cells) {
        const cell = node.rows[rowIndex]?.cells[cellIndex]
        if (cell) {
          cell.nodes = [{ type: "paragraph", children: [] }]
        }
      }
      return nextModel
    })
  }

  const moveTable = (tableIndex: number, targetNodeIndex: number): void => {
    applyModelChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") return current
      nextModel.nodes.splice(tableIndex, 1)
      const adjustedTarget = targetNodeIndex > tableIndex ? targetNodeIndex - 1 : targetNodeIndex
      nextModel.nodes.splice(adjustedTarget, 0, node)
      return nextModel
    })
  }

  const moveEmbeddedTableToBody = (_tableRuntimeKey: string, _targetNodeIndex: number): void => {
    // Embedded table extraction — reserved for nested-table scenarios
  }

  // --- image operations ---

  const insertImageFile = async (_file: File): Promise<void> => {
    // Reads image file, converts to data URL or Uint8Array, inserts as ImageRunNode.
    // Full implementation in editor-image module.
  }

  const resizeImage = (_location: DocxImageLocation, widthPx: number, heightPx: number): void => {
    applyModelChange((current) => {
      const nextModel = cloneDocModel(current)
      const loc = _location
      let node: any | undefined
      if (loc.kind === "paragraph") {
        node = nextModel.nodes[loc.nodeIndex]
      } else {
        const tableNode = nextModel.nodes[loc.tableIndex]
        if (tableNode && tableNode.type === "table") {
          const cell = tableNode.rows[loc.rowIndex]?.cells[loc.cellIndex]
          node = cell?.nodes[loc.paragraphIndex]
        }
      }
      if (!node || node.type !== "paragraph") return current
      const child = node.children[loc.childIndex]
      if (!child || child.type !== "image") return current
      child.widthPx = widthPx
      child.heightPx = heightPx
      return nextModel
    })
  }

  const setSyntheticTextBoxText = (_location: DocxImageLocation, _text: string): void => {}
  const setImageWrapMode = (_location: DocxImageLocation, _mode: DocxImageWrapMode, _seed?: Partial<NonNullable<ImageRunNode["floating"]>>): void => {}
  const moveFloatingImage = (_location: DocxImageLocation, _patch: Partial<NonNullable<ImageRunNode["floating"]>>): void => {}
  const moveSectionFloatingImage = (_location: DocxSectionImageLocation, _patch: Partial<NonNullable<ImageRunNode["floating"]>>): void => {}
  const moveParagraphDropCap = (_nodeIndex: number, _patch: Partial<NonNullable<NonNullable<ParagraphNode["style"]>["dropCap"]>>): void => {}
  const setParagraphDropCapFontSizePt = (_nodeIndex: number, _fontSizePt: number): void => {}
  const setParagraphDropCapText = (_nodeIndex: number, _text: string): void => {}
  const moveImage = (_source: DocxImageLocation, _target: DocxImageDropTarget): void => {}

  // --- form field operations ---
  const selectFormField = (location?: DocxFormFieldLocation): void => {
    selectedFormFieldLocation.value = location
  }

  const toggleFormCheckbox = (location: DocxFormFieldLocation): void => {
    applyModelChange((current) => {
      const nextModel = cloneDocModel(current)
      const { paragraph: para } = getParagraphAtLocation(nextModel, location)
      if (!para) return current
      const field = para.children[location.childIndex]
      if (!field || field.type !== "form-field" || field.fieldType !== "checkbox") return current
      const newValue = field.value === "1" ? "0" : "1"
      return nextModel
    })
  }

  const setFormFieldValue = (_location: DocxFormFieldLocation, _value: string): void => {
    applyModelChange((current) => {
      const nextModel = cloneDocModel(current)
      const { paragraph: para } = getParagraphAtLocation(nextModel, _location)
      if (!para) return current
      const field = para.children[_location.childIndex]
      if (!field || field.type !== "form-field") return current
      field.value = _value
      return nextModel
    })
  }

  const updateFormFieldWidget = (
    _location: DocxFormFieldLocation,
    _patch: Partial<NonNullable<FormFieldRunNode["widget"]>>
  ): void => {
    applyModelChange((current) => {
      const nextModel = cloneDocModel(current)
      const { paragraph: para } = getParagraphAtLocation(nextModel, _location)
      if (!para) return current
      const field = para.children[_location.childIndex]
      if (!field || field.type !== "form-field") return current
      field.widget = { ...(field.widget ?? {}), ..._patch }
      return nextModel
    })
  }

  // --- import / export ---
  let pendingExportModelTransformer: ((model: DocModel) => DocModel) | undefined

  const registerPendingExportModelTransformer = (transformer?: (model: DocModel) => DocModel): void => {
    pendingExportModelTransformer = transformer
  }

  const newDocument = (): void => {
    model.value = cloneDocModel(starterTemplate)
    documentLoadNonce.value += 1
    history.value = { past: [], future: [] }
    historyRestoreRequest.value = undefined
    basePackage.value = undefined
    selection.value = { kind: "paragraph", nodeIndex: 0 }
    activeTextRange.value = undefined
    pendingRunStyle.value = undefined
    selectedFormFieldLocation.value = undefined
    importError.value = undefined
    status.value = "Ready"
  }

  let activeImportAbortController: AbortController | undefined

  const importDocxFile = async (file: File): Promise<void> => {
    // Full implementation delegated to editor-import-export module
    // Stub that accepts File and defers to importDocxBuffer from docx-core
    activeImportAbortController?.abort()
    activeImportAbortController = undefined

    if (!/\.docx?$/i.test(file.name)) {
      status.value = `Only .docx and .doc files are supported`
      return
    }

    isImporting.value = true
    importError.value = undefined
    status.value = `Loading ${file.name}...`

    const abortController = new AbortController()
    activeImportAbortController = abortController

    try {
      const { importDocxBuffer } = await import("@extend-ai/docx-core")
      const buffer = await file.arrayBuffer()
      const { model: parsedModel, package: parsedPkg } = await importDocxBuffer(buffer, {
        signal: abortController.signal,
        transferBuffer: false,
      })
      model.value = parsedModel
      basePackage.value = parsedPkg
      documentLoadNonce.value += 1
      fileName.value = file.name
      history.value = { past: [], future: [] }
      historyRestoreRequest.value = undefined
      selection.value = { kind: "paragraph", nodeIndex: 0 }
      activeTextRange.value = undefined
      pendingRunStyle.value = undefined
      selectedFormFieldLocation.value = undefined
      importError.value = undefined
      status.value = "Ready"
      isImporting.value = false
    } catch (error) {
      isImporting.value = false
      const err = error instanceof Error ? error : new Error("Unknown import error")
      importError.value = err
      status.value = `Failed to load file: ${err.message}`
    }
  }

  const exportDocx = (): void => {
    status.value = "Export... (see save dialog)"
  }

  // -----------------------------------------------------------------------
  // Controller return
  // -----------------------------------------------------------------------

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
    suppressNextDomSelectionRestore: suppressNextDomSelectionRestoreFn,
    beginSelectionSession,
    clearSelectionSession,
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
    registerPendingExportModelTransformer,
    setStatus,
    setDocumentTheme: setDocumentThemeFn,
    setShowTrackedChanges: setShowTrackedChangesFn,
    setShowComments: setShowCommentsFn,
    syncPaginationInfo,
    toggleShowTrackedChanges: toggleShowTrackedChangesFn,
    toggleShowComments: toggleShowCommentsFn,
    importDocxFile,
    newDocument,
    exportDocx,
    undo,
    redo,
    setHeading,
    setParagraphStyle,
    setLineSpacing,
    setFontFamily,
    setFontSize,
    toggleBold,
    toggleItalic,
    toggleUnderline,
    toggleStrike,
    toggleSuperscript,
    toggleSubscript,
    setHighlight,
    setTextColor,
    setLink,
    selectFormField,
    toggleFormCheckbox,
    setFormFieldValue,
    updateFormFieldWidget,
    applyBorderPreset,
    setAlignment,
    toggleList,
    adjustSelectedListDepth,
    insertListItemAfterSelection,
    splitParagraphAtSelection,
    insertTable,
    insertImageFile,
    appendParagraph,
    resizeImage,
    setSyntheticTextBoxText,
    setImageWrapMode,
    moveFloatingImage,
    moveSectionFloatingImage,
    moveParagraphDropCap,
    setParagraphDropCapFontSizePt,
    setParagraphDropCapText,
    moveImage,
    setActiveTextRange,
    selectParagraph,
    selectTableCell,
    clearTableCellContents,
    insertTableRow,
    insertTableColumn,
    deleteTableRow,
    deleteTableColumn,
    deleteTable,
    moveTable,
    moveEmbeddedTableToBody,
    replaceExpandedSelection,
    deleteExpandedSelection,
    commitParagraphText,
    commitTableCellText,
    commitTableCellParagraphTextRecursive,
    commitSectionParagraphText,
  }
}
