// editor-selection.ts — selection session management
//
// Extracted from useDocxEditor.ts.
// Simple selection wrappers (setActiveTextRange, selectParagraph, selectTableCell)
// are defined inline in useDocxEditor since they're thin dispatch calls.

import type {
  DocModel,
  DocxEditorController,
  DocxImageLocation,
  DocxRevisionDisplay,
  DocxSelectionSessionKind,
  DocxTableContextMenuContext,
  DocxTextRange,
  DocxTextRangeLocation,
  ParagraphNode,
} from "@arcships/docx-core"
import {
  compareTextRangeBoundaries,
  displayTextOffsetForSourceOffset,
  getParagraphAtLocation,
  hiddenParagraphTextRanges,
  paragraphText,
  reviewDisplayParagraph,
  sameTextRange,
  sourceTextOffsetForDisplayOffset,
} from "@arcships/docx-core"
import type { EditorCore } from "./editor-shared"

const PARAGRAPH_HOST_SELECTOR = "[data-docx-paragraph-host='true']"
const TABLE_CELL_PARAGRAPH_HOST_SELECTOR =
  "[data-docx-table-cell-paragraph-host='true']"
const TEXT_HOST_SELECTOR =
  `${PARAGRAPH_HOST_SELECTOR},${TABLE_CELL_PARAGRAPH_HOST_SELECTOR}`

type SelectionHost = HTMLElement & {
  dataset: DOMStringMap
}

type SelectionSyncOptions = {
  preserveExpandedSelection?: boolean
  /** Typing extends the active range; keep the caret's pending run style. */
  preservePendingRunStyle?: boolean
  /**
   * Review mode the DOM was rendered with. Defaults to the controller's
   * effective mode (final text while revisions are hidden).
   */
  revisionDisplay?: DocxRevisionDisplay
}

function elementForNode(node: Node | null): HTMLElement | undefined {
  if (!node) return undefined
  const candidate = node as Node & {
    closest?: (selector: string) => Element | null
    parentElement?: HTMLElement | null
  }
  if (typeof candidate.closest === "function") {
    return candidate as unknown as HTMLElement
  }
  return candidate.parentElement ?? undefined
}

function closestTextHost(node: Node | null): SelectionHost | undefined {
  const element = elementForNode(node)
  const host = element?.closest?.(TEXT_HOST_SELECTOR)
  return host ? host as SelectionHost : undefined
}

/** Resolve the DOCX paragraph/table-cell host that contains a DOM node. */
export function docxTextHostFromNode(
  node: Node | null | undefined
): HTMLElement | undefined {
  return closestTextHost(node ?? null) as HTMLElement | undefined
}

/**
 * Text location (body paragraph or table cell) of the host that owns a DOM
 * node. Returns undefined for hosts without a model index (for example the
 * header/footer stories, which use negative display indexes).
 */
export function docxTextLocationForNode(
  node: Node | null | undefined
): DocxTextRangeLocation | undefined {
  const host = closestTextHost(node ?? null)
  return host ? locationForHost(host) : undefined
}

/** Read the visible text of a DOCX text host, preserving line breaks. */
export function readDocxTextHostText(host: HTMLElement): string {
  function readNode(node: Node): string {
    const element = node as Element
    if (element.matches?.("[data-docx-numbering-label='true']")) return ""
    // Footnote/endnote markers are display-only: their label is not paragraph text.
    if (element.matches?.("[data-docx-note-marker='true']")) return ""
    if (element.matches?.("[data-docx-image-child-index]")) return ""
    // Form fields are atomic model runs edited through the field layer, not
    // text characters. Skipping their display value keeps draft text/offsets
    // aligned with paragraphText() and lets an unrelated edit preserve them.
    if (element.matches?.("[data-docx-form-field='true']")) return ""
    // Deleted/moved-from runs stay on screen while revisions are shown, but they
    // are not part of the text the model keeps: skipping them here keeps
    // selection offsets and committed text in agreement.
    const trackedChange = element.getAttribute?.("data-docx-tracked-change")
    if (trackedChange === "deletion" || trackedChange === "move-from") return ""
    if (element.tagName?.toLowerCase() === "br") return "\n"
    if (node.nodeType === 3) return node.textContent ?? ""
    // Test hosts and other lightweight stand-ins only expose `textContent`.
    if (!node.childNodes) return node.textContent ?? ""
    return Array.from(node.childNodes).map(readNode).join("")
  }

  return readNode(host)
}

/** Resolve the table cell under a DOM node for context-menu actions. */
export function docxTableContextFromNode(
  node: Node | null | undefined
): DocxTableContextMenuContext | undefined {
  const element = elementForNode(node ?? null)
  const cell = element?.closest?.("[data-docx-table-cell='true']") as HTMLElement | null | undefined
  if (!cell) return undefined
  const tableIndex = nonNegativeInteger(cell.getAttribute?.("data-docx-table-index"))
  const rowIndex = nonNegativeInteger(cell.getAttribute?.("data-docx-table-row-index"))
  const cellIndex = nonNegativeInteger(cell.getAttribute?.("data-docx-table-cell-index"))
  if (tableIndex === undefined || rowIndex === undefined || cellIndex === undefined) {
    return undefined
  }
  return { tableIndex, rowIndex, cellIndex }
}

/**
 * Resolve the image under a DOM node for context-menu actions.
 *
 * Images rendered inside a paragraph host carry `data-docx-image-child-index`;
 * floating images rendered by the interaction layer carry
 * `data-image-node-index` + `data-image-child-index` instead of a paragraph
 * host. Table-cell images resolve through their paragraph host so the same
 * model location drives the layer commands as every other image edit.
 */
export function docxImageLocationFromNode(
  node: Node | null | undefined
): DocxImageLocation | undefined {
  const element = elementForNode(node ?? null)
  const image = element?.closest?.(
    "[data-docx-image-child-index],[data-image-child-index]"
  ) as HTMLElement | null | undefined
  if (!image) return undefined

  const childIndex = nonNegativeInteger(
    image.getAttribute?.("data-docx-image-child-index") ??
      image.getAttribute?.("data-image-child-index")
  )
  if (childIndex === undefined) return undefined

  const paragraphHost = image.closest?.(
    TABLE_CELL_PARAGRAPH_HOST_SELECTOR
  ) as SelectionHost | null | undefined
  if (paragraphHost) {
    const location = locationForHost(paragraphHost)
    return location ? { ...location, childIndex } as DocxImageLocation : undefined
  }

  const paragraph = image.closest?.("[data-docx-paragraph-node-index]") as
    | HTMLElement
    | null
    | undefined
  const paragraphIndex = nonNegativeInteger(
    paragraph?.getAttribute?.("data-docx-paragraph-node-index")
  )
  if (paragraphIndex !== undefined) {
    return { kind: "paragraph", nodeIndex: paragraphIndex, childIndex }
  }

  // The image layer container is not inside a paragraph host; its own data
  // attributes identify the owning body paragraph.
  const floatingContainer = image.closest?.("[data-image-node-index]") as
    | HTMLElement
    | null
    | undefined
  const nodeIndex = nonNegativeInteger(
    floatingContainer?.getAttribute?.("data-image-node-index")
  )
  return nodeIndex === undefined
    ? undefined
    : { kind: "paragraph", nodeIndex, childIndex }
}

function nonNegativeInteger(value: string | null | undefined): number | undefined {
  if (value === undefined || value === null || value.trim() === "") return undefined
  const parsed = Number(value)
  return Number.isInteger(parsed) && parsed >= 0 ? parsed : undefined
}

function hostAttribute(host: SelectionHost, name: string): string | undefined {
  const value = host.getAttribute?.(name)
  return value === null || value === undefined ? undefined : String(value)
}

function locationForHost(host: SelectionHost): DocxTextRangeLocation | undefined {
  if (host.matches?.(TABLE_CELL_PARAGRAPH_HOST_SELECTOR)) {
    const tableIndex = nonNegativeInteger(hostAttribute(host, "data-docx-table-index"))
    const rowIndex = nonNegativeInteger(hostAttribute(host, "data-docx-table-row-index"))
    const cellIndex = nonNegativeInteger(hostAttribute(host, "data-docx-table-cell-index"))
    const paragraphIndex = nonNegativeInteger(hostAttribute(host, "data-docx-paragraph-index"))
    if (
      tableIndex === undefined ||
      rowIndex === undefined ||
      cellIndex === undefined ||
      paragraphIndex === undefined
    ) {
      return undefined
    }
    return { kind: "table-cell", tableIndex, rowIndex, cellIndex, paragraphIndex }
  }

  const nodeIndex = nonNegativeInteger(
    hostAttribute(host, "data-docx-paragraph-node-index")
  )
  return nodeIndex === undefined ? undefined : { kind: "paragraph", nodeIndex }
}

function clampOffset(offset: number, length: number): number {
  return Math.max(0, Math.min(Math.round(offset), Math.max(0, length)))
}

/**
 * Review mode the page was actually rendered with. Hidden revisions render
 * final text even when the preferred mode is `inline`.
 */
export function effectiveControllerRevisionDisplay(
  editor: Pick<DocxEditorController, "showTrackedChanges" | "revisionDisplay">
): DocxRevisionDisplay {
  return editor.showTrackedChanges ? editor.revisionDisplay : "none"
}

/**
 * Maps a paragraph's display-text offset (the DOM/caret space) to the source
 * model offset. Works for body, table, header, footer and note paragraphs.
 */
export function paragraphDisplayOffsetToSourceOffset(
  paragraph: ParagraphNode,
  display: DocxRevisionDisplay,
  offset: number
): number {
  if (display === "inline") return clampOffset(offset, paragraphText(paragraph).length)
  const displayText = paragraphText(reviewDisplayParagraph(paragraph, display))
  const ranges = hiddenParagraphTextRanges(paragraph, display)
  return clampOffset(
    sourceTextOffsetForDisplayOffset(ranges, clampOffset(offset, displayText.length)),
    paragraphText(paragraph).length
  )
}

/**
 * Maps a paragraph's source-model offset back to the display text the DOM and
 * caret use. Works for every story; callers that already hold the paragraph do
 * not need a range location.
 */
export function paragraphSourceOffsetToDisplayOffset(
  paragraph: ParagraphNode,
  display: DocxRevisionDisplay,
  offset: number
): number {
  if (display === "inline") return clampOffset(offset, paragraphText(paragraph).length)
  const displayText = paragraphText(reviewDisplayParagraph(paragraph, display))
  const ranges = hiddenParagraphTextRanges(paragraph, display)
  return clampOffset(
    displayTextOffsetForSourceOffset(ranges, clampOffset(offset, paragraphText(paragraph).length)),
    displayText.length
  )
}

function paragraphAtRangeLocation(
  model: DocModel,
  location: DocxTextRangeLocation
): ParagraphNode | undefined {
  return getParagraphAtLocation(model, location).paragraph
}

/** Translate DOM display offsets into source-model offsets before writing. */
export function displayTextRangeToSourceTextRange(
  model: DocModel,
  display: DocxRevisionDisplay,
  range: DocxTextRange
): DocxTextRange {
  const mapBoundary = (boundary: DocxTextRange["start"]): DocxTextRange["start"] => {
    const paragraph = paragraphAtRangeLocation(model, boundary.location)
    return {
      location: { ...boundary.location },
      offset: paragraph
        ? paragraphDisplayOffsetToSourceOffset(paragraph, display, boundary.offset)
        : boundary.offset,
    }
  }
  return { start: mapBoundary(range.start), end: mapBoundary(range.end) }
}

/** Translate source-model offsets into the DOM display offsets on restore. */
export function sourceTextRangeToDisplayTextRange(
  model: DocModel,
  display: DocxRevisionDisplay,
  range: DocxTextRange
): DocxTextRange {
  const mapBoundary = (boundary: DocxTextRange["start"]): DocxTextRange["start"] => {
    const paragraph = paragraphAtRangeLocation(model, boundary.location)
    return {
      location: { ...boundary.location },
      offset: paragraph
        ? paragraphSourceOffsetToDisplayOffset(paragraph, display, boundary.offset)
        : boundary.offset,
    }
  }
  return { start: mapBoundary(range.start), end: mapBoundary(range.end) }
}

// Elements that are rendered in a text host but are not part of its text model
// (`paragraphText`): list numbering labels, note markers and tracked change
// deletions. Counting them shifts every caret/selection offset.
export const PARAGRAPH_NON_TEXT_CONTENT_SELECTOR =
  "[data-docx-numbering-label='true'],[data-docx-note-marker='true'],[data-docx-tracked-change='deletion']"

function fragmentTextLength(fragment: DocumentFragment): number {
  fragment.querySelectorAll?.(PARAGRAPH_NON_TEXT_CONTENT_SELECTOR).forEach((element) => element.remove())
  fragment.querySelectorAll?.("br").forEach((lineBreak) => {
    lineBreak.replaceWith("\n")
  })
  return fragment.textContent?.length ?? 0
}

function boundaryOffsetWithinHost(
  host: SelectionHost,
  node: Node,
  offset: number
): number | undefined {
  if (node !== host && !host.contains(node)) return undefined
  try {
    const prefix = host.ownerDocument.createRange()
    prefix.setStart(host, 0)
    prefix.setEnd(node, offset)
    return fragmentTextLength(prefix.cloneContents())
  } catch {
    return undefined
  }
}

/** Map a viewport point inside a text host to a plain-text caret offset. */
export function textOffsetAtDomPoint(
  host: HTMLElement,
  clientX: number,
  clientY: number
): number | undefined {
  const ownerDocument = host.ownerDocument as Document & {
    caretRangeFromPoint?: (x: number, y: number) => Range | null
    caretPositionFromPoint?: (x: number, y: number) => { offsetNode: Node; offset: number } | null
  }
  if (!ownerDocument) return undefined

  let node: Node | undefined
  let offset: number | undefined

  if (typeof ownerDocument.caretRangeFromPoint === "function") {
    const range = ownerDocument.caretRangeFromPoint(clientX, clientY)
    if (range) {
      node = range.startContainer
      offset = range.startOffset
    }
  } else if (typeof ownerDocument.caretPositionFromPoint === "function") {
    const position = ownerDocument.caretPositionFromPoint(clientX, clientY)
    if (position) {
      node = position.offsetNode
      offset = position.offset
    }
  }

  if (!node || offset === undefined) return undefined
  return boundaryOffsetWithinHost(host as SelectionHost, node, offset)
}

function hostForLocation(
  root: HTMLElement,
  location: DocxTextRangeLocation
): SelectionHost | undefined {
  const selector = location.kind === "paragraph"
    ? `${PARAGRAPH_HOST_SELECTOR}[data-docx-paragraph-node-index="${location.nodeIndex}"]`
    : `${TABLE_CELL_PARAGRAPH_HOST_SELECTOR}` +
      `[data-docx-table-index="${location.tableIndex}"]` +
      `[data-docx-table-row-index="${location.rowIndex}"]` +
      `[data-docx-table-cell-index="${location.cellIndex}"]` +
      `[data-docx-paragraph-index="${location.paragraphIndex}"]`
  return (root.querySelector?.(selector) ?? undefined) as SelectionHost | undefined
}

function domPositionForTextOffset(
  host: SelectionHost,
  targetOffset: number
): { node: Node; offset: number } {
  const safeTarget = Math.max(0, Math.round(targetOffset))
  let traversed = 0
  let resolved: { node: Node; offset: number } | undefined

  function visit(parent: Node): void {
    if (resolved) return
    const children = Array.from(parent.childNodes)
    for (let index = 0; index < children.length && !resolved; index += 1) {
      const child = children[index]
      const element = child as Element
      if (element.matches?.(PARAGRAPH_NON_TEXT_CONTENT_SELECTOR)) continue
      if (element.tagName?.toLowerCase() === "br") {
        if (safeTarget <= traversed) {
          resolved = { node: parent, offset: index }
          break
        }
        traversed += 1
        if (safeTarget <= traversed) {
          resolved = { node: parent, offset: index + 1 }
          break
        }
        continue
      }
      if (child.nodeType === 3) {
        const length = child.textContent?.length ?? 0
        if (traversed + length >= safeTarget) {
          resolved = { node: child, offset: Math.max(0, safeTarget - traversed) }
          break
        }
        traversed += length
        continue
      }
      visit(child)
    }
  }

  visit(host)
  return resolved ?? { node: host, offset: host.childNodes.length }
}

function expandedRange(range?: DocxTextRange): boolean {
  return Boolean(
    range && compareTextRangeBoundaries(range.start, range.end) < 0
  )
}

function toolbarOwnsFocus(root: HTMLElement): boolean {
  const activeElement = root.ownerDocument?.activeElement
  return Boolean(activeElement && "closest" in activeElement &&
    activeElement.closest(".docx-toolbar, .docx-toolbar-shell"))
}

/** Map the browser's current selection to the public DOCX text-range shape. */
export function docxTextRangeFromDomSelection(
  selection: Selection,
  root: HTMLElement
): DocxTextRange | undefined {
  if (selection.rangeCount === 0) return undefined
  const nativeRange = selection.getRangeAt(0)
  const startHost = closestTextHost(nativeRange.startContainer)
  const endHost = closestTextHost(nativeRange.endContainer)
  if (!startHost || !endHost || !root.contains(startHost) || !root.contains(endHost)) {
    return undefined
  }

  const startLocation = locationForHost(startHost)
  const endLocation = locationForHost(endHost)
  const startOffset = boundaryOffsetWithinHost(
    startHost,
    nativeRange.startContainer,
    nativeRange.startOffset
  )
  const endOffset = boundaryOffsetWithinHost(
    endHost,
    nativeRange.endContainer,
    nativeRange.endOffset
  )
  if (
    !startLocation ||
    !endLocation ||
    startOffset === undefined ||
    endOffset === undefined
  ) {
    return undefined
  }

  return {
    start: { location: startLocation, offset: startOffset },
    end: { location: endLocation, offset: endOffset },
  }
}

/**
 * Synchronize a real browser selection into one editor instance. A toolbar
 * focus transition is intentionally not allowed to collapse an expanded model
 * selection before the toolbar command runs.
 */
export function syncDomSelectionToEditor(
  editor: DocxEditorController,
  root: HTMLElement,
  selection?: Selection | null,
  options: SelectionSyncOptions = {}
): DocxTextRange | undefined {
  const currentSelection = selection ?? root.ownerDocument?.getSelection?.() ??
    (typeof window !== "undefined" ? window.getSelection() : null)
  if (!currentSelection) return undefined
  const displayRange = docxTextRangeFromDomSelection(currentSelection, root)
  if (!displayRange) return undefined
  const revisionDisplay = options.revisionDisplay ?? effectiveControllerRevisionDisplay(editor)
  const range = displayTextRangeToSourceTextRange(
    editor.model,
    revisionDisplay,
    displayRange
  )
  if (
    currentSelection.isCollapsed &&
    expandedRange(editor.activeTextRange) &&
    (options.preserveExpandedSelection || toolbarOwnsFocus(root))
  ) {
    return editor.activeTextRange
  }
  if (!sameTextRange(editor.activeTextRange, range)) {
    editor.setActiveTextRange(range, {
      preservePendingRunStyle: options.preservePendingRunStyle,
    })
  }
  return range
}

/** Restore a model range after Vue has replaced a contenteditable subtree. */
export function restoreDomSelectionFromTextRange(
  root: HTMLElement,
  range: DocxTextRange,
  options: {
    focus?: boolean
    owner?: HTMLElement
    /** Source model used to translate source offsets back to display text. */
    sourceModel?: DocModel
    /** Review mode the DOM was rendered with. */
    revisionDisplay?: DocxRevisionDisplay
  } = {}
): boolean {
  const selection = root.ownerDocument?.getSelection?.() ??
    (typeof window !== "undefined" ? window.getSelection() : null)
  if (!selection) return false
  const displayRange = options.sourceModel
    ? sourceTextRangeToDisplayTextRange(
        options.sourceModel,
        options.revisionDisplay ?? "inline",
        range
      )
    : range
  const startHost = hostForLocation(root, displayRange.start.location)
  const endHost = hostForLocation(root, displayRange.end.location)
  if (!startHost || !endHost) return false
  if (options.owner && !options.owner.contains(startHost)) return false

  try {
    const start = domPositionForTextOffset(startHost, displayRange.start.offset)
    const end = domPositionForTextOffset(endHost, displayRange.end.offset)
    const nativeRange = root.ownerDocument.createRange()
    nativeRange.setStart(start.node, start.offset)
    nativeRange.setEnd(end.node, end.offset)
    if (options.focus) startHost.focus()
    selection.removeAllRanges()
    selection.addRange(nativeRange)
    return true
  } catch {
    return false
  }
}

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
