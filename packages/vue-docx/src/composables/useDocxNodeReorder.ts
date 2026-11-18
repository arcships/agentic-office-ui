// useDocxNodeReorder.ts — drag-to-reorder for body paragraphs and tables.
//
// The editor shows a handle at the hovered body node. Dragging the handle
// resolves an insertion slot from the node frames under the pointer and commits
// one `moveBodyNode()` call, so a reordered paragraph/table is one undo entry.
//
// The DOM lookups are injectable so the decisions (hovered node, insertion slot,
// no-op drops) can be tested without a browser.

import { computed, ref, type ComputedRef, type Ref } from "vue"
import type { DocxEditorController } from "@arcships/docx-core"

/** A body node host: a paragraph or a top-level table. */
export const DOCX_REORDER_NODE_SELECTOR =
  "[data-docx-paragraph-host='true'],[data-docx-table-host='true']"

/** Stories that are not part of the body node list. */
export const DOCX_REORDER_IGNORED_SELECTOR =
  "[data-docx-page-header='true'],[data-docx-page-footer='true'],[data-docx-notes-section],[data-docx-table-runtime-key]"

export interface DocxReorderFrame {
  /** Body node index from the host's data attribute. */
  nodeIndex: number
  top: number
  bottom: number
}

export interface DocxReorderHover extends DocxReorderFrame {
  left: number
  width: number
}

export interface DocxReorderRect {
  top: number
  bottom: number
  left: number
  width: number
}

export interface UseDocxNodeReorderOptions {
  /** Handle/drag are off for read-only surfaces and while review locks editing. */
  enabled?: () => boolean
  /** Element under a viewport point. */
  elementFromPoint?: (x: number, y: number) => Element | null
  /** Viewport rect of an element. */
  rectOf?: (element: Element) => DocxReorderRect | undefined
  /** Every reorderable body node frame, in document order. */
  collectFrames?: () => DocxReorderFrame[]
}

function readIndex(element: Element | null | undefined, attribute: string): number | undefined {
  const raw = element?.getAttribute?.(attribute)
  if (raw === null || raw === undefined || raw === "") return undefined
  const value = Number(raw)
  return Number.isInteger(value) && value >= 0 ? value : undefined
}

/**
 * Body node index for a hovered element, or undefined when the element is not
 * part of the body (header/footer, notes, nested table, non-node chrome).
 */
export function docxReorderNodeIndexFromElement(
  element: Element | null | undefined,
): number | undefined {
  if (!element) return undefined
  const host = (element as Element & {
    closest?: (selector: string) => Element | null
  }).closest?.(DOCX_REORDER_NODE_SELECTOR)
  if (!host) return undefined
  const insideIgnoredStory = (
    host as Element & { closest?: (selector: string) => Element | null }
  ).closest?.(DOCX_REORDER_IGNORED_SELECTOR)
  if (insideIgnoredStory) return undefined
  return (
    readIndex(host, "data-docx-paragraph-node-index") ??
    readIndex(host, "data-docx-table-index")
  )
}

/**
 * Insertion slot for a pointer position: before the first node whose midpoint
 * is below the pointer, otherwise after the last node.
 */
export function resolveDocxReorderInsertionIndex(
  pointerY: number,
  frames: readonly DocxReorderFrame[],
): number | undefined {
  if (frames.length === 0) return undefined
  for (const frame of frames) {
    if (pointerY < (frame.top + frame.bottom) / 2) return frame.nodeIndex
  }
  return frames[frames.length - 1].nodeIndex + 1
}

/**
 * Node index after the move, or undefined when the drop is a no-op (dropping
 * onto either side of the node's own slot).
 */
export function resolveDocxReorderTarget(
  sourceIndex: number,
  insertionIndex: number,
): number | undefined {
  if (!Number.isInteger(sourceIndex) || sourceIndex < 0) return undefined
  if (!Number.isInteger(insertionIndex) || insertionIndex < 0) return undefined
  const target = insertionIndex > sourceIndex ? insertionIndex - 1 : insertionIndex
  return target === sourceIndex ? undefined : target
}

export interface UseDocxNodeReorderResult {
  hover: Ref<DocxReorderHover | undefined>
  dragging: Ref<boolean>
  dropIndex: Ref<number | undefined>
  /** Viewport Y of the drop indicator line. */
  dropTop: Ref<number | undefined>
  enabled: ComputedRef<boolean>
  onHover: (event: { clientX: number; clientY: number; target?: unknown }) => void
  clearHover: () => void
  startDrag: (nodeIndex: number) => void
  updateDrag: (event: { clientY: number }) => void
  finishDrag: () => boolean
  cancelDrag: () => void
  /** Keyboard reorder: -1 moves up one slot, +1 moves down one slot. */
  moveBy: (nodeIndex: number, delta: -1 | 1) => boolean
}

export function useDocxNodeReorder(
  controller: Pick<DocxEditorController, "moveBodyNode"> | undefined,
  options: UseDocxNodeReorderOptions = {},
): UseDocxNodeReorderResult {
  const hover = ref<DocxReorderHover | undefined>()
  const dragging = ref(false)
  const dropIndex = ref<number | undefined>()
  const dropTop = ref<number | undefined>()
  let dragSourceIndex: number | undefined

  const enabled = computed(() => options.enabled ? options.enabled() : Boolean(controller))

  function elementFromPoint(x: number, y: number): Element | null {
    if (options.elementFromPoint) return options.elementFromPoint(x, y)
    const ownerDocument = (globalThis as { document?: Document }).document
    return ownerDocument?.elementFromPoint?.(x, y) ?? null
  }

  function rectOf(element: Element): DocxReorderRect | undefined {
    if (options.rectOf) return options.rectOf(element)
    const rect = (element as Element & {
      getBoundingClientRect?: () => DocxReorderRect
    }).getBoundingClientRect?.()
    return rect ? { top: rect.top, bottom: rect.bottom, left: rect.left, width: rect.width } : undefined
  }

  function collectFrames(): DocxReorderFrame[] {
    if (options.collectFrames) return options.collectFrames()
    const ownerDocument = (globalThis as { document?: Document }).document
    if (!ownerDocument?.querySelectorAll) return []
    const frames: DocxReorderFrame[] = []
    ownerDocument
      .querySelectorAll<HTMLElement>(DOCX_REORDER_NODE_SELECTOR)
      .forEach((element) => {
        const nodeIndex = docxReorderNodeIndexFromElement(element)
        if (nodeIndex === undefined) return
        const rect = rectOf(element)
        if (!rect) return
        frames.push({ nodeIndex, top: rect.top, bottom: rect.bottom })
      })
    return frames.sort((left, right) => left.top - right.top)
  }

  function onHover(event: {
    clientX: number
    clientY: number
    target?: unknown
  }): void {
    if (!enabled.value || dragging.value) return
    const target = event.target as Element | null | undefined
    const nodeIndex = docxReorderNodeIndexFromElement(target)
    if (nodeIndex === undefined) {
      hover.value = undefined
      return
    }
    const host = (target as Element & {
      closest?: (selector: string) => Element | null
    }).closest?.(DOCX_REORDER_NODE_SELECTOR)
    if (!host) {
      hover.value = undefined
      return
    }
    const rect = rectOf(host)
    if (!rect) {
      hover.value = undefined
      return
    }
    hover.value = {
      nodeIndex,
      top: rect.top,
      bottom: rect.bottom,
      left: rect.left,
      width: rect.width,
    }
  }

  function clearHover(): void {
    if (dragging.value) return
    hover.value = undefined
  }

  function startDrag(nodeIndex: number): void {
    if (!enabled.value || !controller) return
    dragging.value = true
    dragSourceIndex = nodeIndex
    dropIndex.value = undefined
    dropTop.value = hover.value?.top
  }

  function updateDrag(event: { clientY: number }): void {
    if (!dragging.value) return
    const frames = collectFrames()
    if (frames.length === 0) {
      dropIndex.value = undefined
      dropTop.value = undefined
      return
    }
    const insertionIndex = resolveDocxReorderInsertionIndex(event.clientY, frames)
    dropIndex.value = insertionIndex
    const boundaryFrame = frames.find((frame) => frame.nodeIndex >= (insertionIndex ?? 0))
    dropTop.value = boundaryFrame
      ? boundaryFrame.top
      : frames[frames.length - 1].bottom
  }

  function reset(): void {
    dragging.value = false
    dragSourceIndex = undefined
    dropIndex.value = undefined
    dropTop.value = undefined
    hover.value = undefined
  }

  function finishDrag(): boolean {
    if (!dragging.value) return false
    const sourceIndex = dragSourceIndex
    const insertionIndex = dropIndex.value
    reset()
    if (sourceIndex === undefined || insertionIndex === undefined) return false
    if (!controller) return false
    if (resolveDocxReorderTarget(sourceIndex, insertionIndex) === undefined) return false
    return controller.moveBodyNode(sourceIndex, insertionIndex)
  }

  function cancelDrag(): void {
    if (!dragging.value) return
    reset()
  }

  function moveBy(nodeIndex: number, delta: -1 | 1): boolean {
    if (!controller) return false
    const insertionIndex = delta < 0 ? nodeIndex - 1 : nodeIndex + 2
    if (insertionIndex < 0) return false
    return controller.moveBodyNode(nodeIndex, insertionIndex)
  }

  return {
    hover,
    dragging,
    dropIndex,
    dropTop,
    enabled,
    onHover,
    clearHover,
    startDrag,
    updateDrag,
    finishDrag,
    cancelDrag,
    moveBy,
  }
}
