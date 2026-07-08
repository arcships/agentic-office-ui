// Page surface registry — shared registry that connects the viewer's mounted
// page DOM elements to thumbnail consumers (useDocxPageThumbnails).
//
// Migrated from upstream @extend-ai/react-docx, editor.tsx lines 29582-29668

import type { DocxEditorController, DocxPaginationInfo } from "@extend-ai/docx-core"

// ---------------------------------------------------------------------------
// Snapshots used by the fast (direct-draw) thumbnail renderer

export interface DocxPageThumbnailTextRunSnapshot {
  text: string
  bold?: boolean
  italic?: boolean
  color?: string
  backgroundColor?: string
  fontSizePx: number
  fontFamily?: string
}

export interface DocxPageThumbnailSnapshotElement {
  kind: "paragraph" | "image" | "image-placeholder" | "table"
  xPx: number
  yPx: number
  widthPx: number
  heightPx: number
  align?: string
  backgroundColor?: string
  lineHeightPx?: number
  startLineIndex?: number
  runs?: DocxPageThumbnailTextRunSnapshot[]
  src?: string
  cells?: DocxPageThumbnailTableCellSnapshot[]
}

export interface DocxPageThumbnailTableCellSnapshot {
  xPx: number
  yPx: number
  widthPx: number
  heightPx: number
  backgroundColor?: string
  runs?: DocxPageThumbnailTextRunSnapshot[]
}

export interface DocxPageThumbnailRenderSnapshot {
  key: string
  sourceWidthPx: number
  sourceHeightPx: number
  pageBackgroundColor: string
  elements: DocxPageThumbnailSnapshotElement[]
}

export interface DocxViewerPageThumbnailSnapshotEntry {
  key: string
  getSnapshot: () => DocxPageThumbnailRenderSnapshot
}

export interface DocxViewerPageSurfaceSize {
  widthPx: number
  heightPx: number
}

// ---------------------------------------------------------------------------
// Registry

export interface DocxViewerPageSurfaceRegistry {
  /** DOM elements for mounted page surfaces, keyed by zero-based page index. */
  pageElements: Map<number, HTMLDivElement>

  /** Content signatures per page, synced when pagination or content changes. */
  pageContentKeys: Map<number, string>

  /** Layout-derived page sizes, available for offscreen pages. */
  pageSizes: Map<number, DocxViewerPageSurfaceSize>

  /** Fast-paint draw snapshots for offscreen pages. */
  pageThumbnailSnapshots: Map<number, DocxViewerPageThumbnailSnapshotEntry>

  /** Change listeners, notified when surface state changes. */
  listeners: Set<() => void>
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
    pageThumbnailSnapshots: new Map(),
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
  listener: () => void
): () => void {
  const registry = ensureDocxViewerPageSurfaceRegistry(editor)
  registry.listeners.add(listener)
  return () => {
    registry.listeners.delete(listener)
  }
}

export function notifyDocxViewerPageSurfaceListeners(
  editor: Pick<DocxEditorController, "syncPaginationInfo">
): void {
  const registry = ensureDocxViewerPageSurfaceRegistry(editor)
  for (const listener of registry.listeners) {
    listener()
  }
}

// ---------------------------------------------------------------------------
// Registry owner for thumbnail composable

export function docxViewerPageSurfaceRegistryOwner(
  editor: Pick<DocxEditorController, "syncPaginationInfo">
): object {
  return registryOwner(editor)
}
