// editor-selection.ts — selection session management
//
// Extracted from useDocxEditor.ts.
// Simple selection wrappers (setActiveTextRange, selectParagraph, selectTableCell)
// are defined inline in useDocxEditor since they're thin dispatch calls.

import type {
  DocxEditorController,
  DocxSelectionSessionKind,
  DocxTextRange,
  DocxTextRangeLocation,
} from "@arcships/docx-core"
import { compareTextRangeBoundaries, sameTextRange } from "@arcships/docx-core"
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

function fragmentTextLength(fragment: DocumentFragment): number {
  fragment
    .querySelectorAll?.("[data-docx-numbering-label='true']")
    .forEach((label) => label.remove())
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
      if (element.matches?.("[data-docx-numbering-label='true']")) continue
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
  const range = docxTextRangeFromDomSelection(currentSelection, root)
  if (!range) return undefined
  if (
    currentSelection.isCollapsed &&
    expandedRange(editor.activeTextRange) &&
    (options.preserveExpandedSelection || toolbarOwnsFocus(root))
  ) {
    return editor.activeTextRange
  }
  if (!sameTextRange(editor.activeTextRange, range)) {
    editor.setActiveTextRange(range)
  }
  return range
}

/** Restore a model range after Vue has replaced a contenteditable subtree. */
export function restoreDomSelectionFromTextRange(
  root: HTMLElement,
  range: DocxTextRange,
  options: { focus?: boolean; owner?: HTMLElement } = {}
): boolean {
  const selection = root.ownerDocument?.getSelection?.() ??
    (typeof window !== "undefined" ? window.getSelection() : null)
  if (!selection) return false
  const startHost = hostForLocation(root, range.start.location)
  const endHost = hostForLocation(root, range.end.location)
  if (!startHost || !endHost) return false
  if (options.owner && !options.owner.contains(startHost)) return false

  try {
    const start = domPositionForTextOffset(startHost, range.start.offset)
    const end = domPositionForTextOffset(endHost, range.end.offset)
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
