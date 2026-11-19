// Page surface registry — shared registry that connects the viewer's mounted
// page DOM elements to thumbnail consumers (useDocxPageThumbnails).
//
// Migrated from upstream @extend-ai/react-docx, editor.tsx lines 29582-29668

import type {
  DocxEditorController,
  DocxSectionPageNumberTracker,
  LayoutSnapshot
} from "@arcships/docx-core"

/** A page surface was published or unpublished by the viewer. */
export interface DocxViewerPageSurfaceChange {
  /** Zero-based page index that changed. */
  pageIndex: number
  kind: "publish" | "unpublish"
  /** Content signature attached to a publish notification. */
  contentKey?: string
}

export interface DocxViewerPageSurfaceSize {
  widthPx: number
  heightPx: number
}

export type DocxViewerPageSurfaceListener = (
  change?: DocxViewerPageSurfaceChange
) => void

// ---------------------------------------------------------------------------
// Registry

export interface DocxViewerPageSurfaceRegistry {
  /** DOM elements for mounted page surfaces, keyed by zero-based page index. */
  pageElements: Map<number, HTMLDivElement>

  /** Content signatures per page, synced when pagination or content changes. */
  pageContentKeys: Map<number, string>

  /** Page sizes reported by mounted page surfaces. */
  pageSizes: Map<number, DocxViewerPageSurfaceSize>

  /**
   * Page count from the viewer's pagination pass. Thumbnail consumers use it
   * instead of re-paginating the document themselves.
   */
  layoutPageCount: number

  /**
   * Layout-derived page sizes for every page, including unmounted ones. Kept
   * separate from `pageSizes` so unmounting an offscreen surface cannot delete
   * the size the viewer resolved from section layout.
   */
  layoutPageSizes: Map<number, DocxViewerPageSurfaceSize>

  /** Viewer-owned layout snapshot shared with thumbnail consumers. */
  layoutSnapshot?: LayoutSnapshot

  /**
   * Viewer-owned per-section page numbering. Thumbnail consumers label pages
   * from it, so the rail cannot disagree with the numbers a page's own
   * header/footer fields resolve.
   */
  pageNumberTracker?: DocxSectionPageNumberTracker

  /** Change listeners, notified when a page surface changes. */
  listeners: Set<DocxViewerPageSurfaceListener>
}

// ---------------------------------------------------------------------------
// Editor-scoped singleton

const registryByEditor = new WeakMap<object, DocxViewerPageSurfaceRegistry>()

function registryOwner(editor: Pick<DocxEditorController, "syncPaginationInfo">): object {
  return editor.syncPaginationInfo as object
}

export function createDocxViewerPageSurfaceRegistry(): DocxViewerPageSurfaceRegistry {
  return {
    pageElements: new Map(),
    pageContentKeys: new Map(),
    pageSizes: new Map(),
    layoutPageCount: 0,
    layoutPageSizes: new Map(),
    listeners: new Set(),
  }
}

export function ensureDocxViewerPageSurfaceRegistry(
  editor: Pick<DocxEditorController, "syncPaginationInfo">
): DocxViewerPageSurfaceRegistry {
  const owner = registryOwner(editor)
  let registry = registryByEditor.get(owner)
  if (!registry) {
    registry = createDocxViewerPageSurfaceRegistry()
    registryByEditor.set(owner, registry)
  }
  return registry
}

export function subscribeDocxViewerPageSurfaces(
  editor: Pick<DocxEditorController, "syncPaginationInfo">,
  listener: DocxViewerPageSurfaceListener
): () => void {
  const registry = ensureDocxViewerPageSurfaceRegistry(editor)
  registry.listeners.add(listener)
  return () => {
    registry.listeners.delete(listener)
  }
}

/**
 * Notifies listeners about one changed page. Passing no change keeps the
 * legacy "something changed" semantics for callers that cannot identify the
 * page; consumers should prefer the per-page form.
 */
export function notifyDocxViewerPageSurfaceListeners(
  editor: Pick<DocxEditorController, "syncPaginationInfo">,
  change?: DocxViewerPageSurfaceChange
): void {
  const registry = ensureDocxViewerPageSurfaceRegistry(editor)
  // Copy so listeners may unsubscribe while being called.
  for (const listener of [...registry.listeners]) {
    listener(change)
  }
}

// ---------------------------------------------------------------------------
// Registry owner for thumbnail composable

export function docxViewerPageSurfaceRegistryOwner(
  editor: Pick<DocxEditorController, "syncPaginationInfo">
): object {
  return registryOwner(editor)
}
