// Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn.
// Tree index for the Vue FileSystem block (#179).
//
// The upstream block keeps a flat `items` manifest keyed by path and infers
// folders from file paths. The Vue block takes an explicit nested tree instead:
// hosts already own the folder structure (an S3 listing, an upload manifest, a
// drive folder), and inferring it from strings loses empty folders and makes
// duplicate names ambiguous. Everything the component needs — ancestors,
// breadth-first children, breadcrumb labels, sizes — is derived here so the
// behaviour is unit-testable without a DOM.
import type { FileSystemItem } from "./types"

export interface FileSystemNode {
  item: FileSystemItem
  /** Parent folder id; `null` for root entries. */
  parentId: string | null
  /** Display path from the root: `contracts/2026/january.pdf`. */
  path: string
  /** 0 for root entries. */
  depth: number
  hasChildren: boolean
}

export interface FileSystemIndex {
  nodes: Map<string, FileSystemNode>
  /** Child ids by folder id. Root entries live under the empty string key. */
  children: Map<string, string[]>
  /** Ids that appeared more than once; the first occurrence wins. */
  duplicateIds: string[]
  /** Non-fatal manifest problems (missing names/ids), for host diagnostics. */
  warnings: string[]
}

const ROOT_KEY = ""

function compareNames(a: FileSystemItem, b: FileSystemItem): number {
  const byKind = a.kind === b.kind ? 0 : a.kind === "folder" ? -1 : 1
  if (byKind !== 0) return byKind
  return a.name.localeCompare(b.name, "en", { numeric: true, sensitivity: "base" })
}

/**
 * Sort children folders-first then by natural name order. Exported because the
 * component and hosts that render a custom list should agree on the order.
 */
export function sortFileSystemItems(items: readonly FileSystemItem[]): FileSystemItem[] {
  return [...items].sort(compareNames)
}

/**
 * Build the lookup index for a nested manifest. Duplicate ids keep the first
 * node (breadth-first) and are reported in `duplicateIds` instead of silently
 * overwriting a subtree.
 */
export function buildFileSystemIndex(items: readonly FileSystemItem[]): FileSystemIndex {
  const nodes = new Map<string, FileSystemNode>()
  const children = new Map<string, string[]>()
  const duplicateIds: string[] = []
  const warnings: string[] = []

  const queue: Array<{ item: FileSystemItem; parentId: string | null; parentPath: string; depth: number }> = []
  for (const item of items) queue.push({ item, parentId: null, parentPath: ROOT_KEY, depth: 0 })

  while (queue.length > 0) {
    const next = queue.shift()!
    const { item, parentId, parentPath, depth } = next
    if (typeof item.id !== "string" || item.id.length === 0) {
      warnings.push(`FileSystem item "${item.name ?? "?"}" has no id and was skipped.`)
      continue
    }
    if (nodes.has(item.id)) {
      duplicateIds.push(item.id)
      continue
    }
    if (typeof item.name !== "string" || item.name.trim().length === 0) {
      warnings.push(`FileSystem item ${item.id} has no name.`)
    }
    const path = parentPath ? `${parentPath}/${item.name}` : item.name
    const hasChildren = item.kind === "folder" && (item.children?.length ?? 0) > 0
    nodes.set(item.id, { item, parentId, path, depth, hasChildren })

    const parentKey = parentId ?? ROOT_KEY
    const siblings = children.get(parentKey) ?? []
    siblings.push(item.id)
    children.set(parentKey, siblings)

    if (item.kind === "folder") {
      for (const child of item.children ?? []) {
        queue.push({ item: child, parentId: item.id, parentPath: path, depth: depth + 1 })
      }
    }
  }

  for (const [parentKey, ids] of children) {
    const sorted = ids
      .map((id) => nodes.get(id))
      .filter((node): node is FileSystemNode => Boolean(node))
      .sort((a, b) => compareNames(a.item, b.item))
      .map((node) => node.item.id)
    children.set(parentKey, sorted)
  }

  return { nodes, children, duplicateIds, warnings }
}

/** Folder ids from the root to `folderId` (exclusive of the root, inclusive of the folder). */
export function fileSystemAncestors(
  index: FileSystemIndex,
  folderId: string | null,
): FileSystemNode[] {
  if (!folderId) return []
  const chain: FileSystemNode[] = []
  let current = index.nodes.get(folderId)
  const guard = new Set<string>()
  while (current && !guard.has(current.item.id)) {
    guard.add(current.item.id)
    chain.unshift(current)
    current = current.parentId ? index.nodes.get(current.parentId) : undefined
  }
  return chain
}

/** Entries directly inside `folderId`, already sorted folders-first. */
export function fileSystemEntries(
  index: FileSystemIndex,
  folderId: string | null,
): FileSystemNode[] {
  const ids = index.children.get(folderId ?? ROOT_KEY) ?? []
  return ids
    .map((id) => index.nodes.get(id))
    .filter((node): node is FileSystemNode => Boolean(node))
}

/**
 * Resolve a folder id that may be stale (deleted folder, controlled prop from a
 * previous manifest) to the nearest existing folder, falling back to the root.
 */
export function resolveFileSystemFolder(
  index: FileSystemIndex,
  folderId: string | null,
): string | null {
  if (!folderId) return null
  const node = index.nodes.get(folderId)
  if (!node || node.item.kind !== "folder") return null
  return node.item.id
}

const SIZE_UNITS = ["B", "KB", "MB", "GB", "TB"] as const

/** Human readable byte size; unit suffixes stay untranslated (SI symbols). */
export function formatFileSystemSize(size: number | undefined): string | null {
  if (typeof size !== "number" || !Number.isFinite(size) || size < 0) return null
  if (size < 1024) return `${Math.round(size)} ${SIZE_UNITS[0]}`
  let value = size
  let unit = 0
  while (value >= 1024 && unit < SIZE_UNITS.length - 1) {
    value /= 1024
    unit += 1
  }
  const rounded = value >= 10 ? Math.round(value) : Math.round(value * 10) / 10
  return `${rounded} ${SIZE_UNITS[unit]}`
}

/** Lowercase extension without the dot; empty when the name has none. */
export function fileSystemExtension(name: string): string {
  const dot = name.lastIndexOf(".")
  return dot <= 0 ? "" : name.slice(dot + 1).toLowerCase()
}

/**
 * Whether a file can show a preview image: an explicit thumb URL wins, else an
 * image MIME with a resolved URL.
 */
export function fileSystemPreviewUrl(item: FileSystemItem): string | null {
  if (item.kind !== "file") return null
  if (item.thumbnailUrl) return item.thumbnailUrl
  if (item.url && item.mime?.startsWith("image/")) return item.url
  return null
}

/** Flatten a nested manifest into `id -> item`, useful for host side updates. */
export function flattenFileSystemItems(items: readonly FileSystemItem[]): FileSystemItem[] {
  const result: FileSystemItem[] = []
  const guard = new Set<string>()
  const visit = (list: readonly FileSystemItem[]) => {
    for (const item of list) {
      if (guard.has(item.id)) continue
      guard.add(item.id)
      result.push(item)
      if (item.kind === "folder" && item.children?.length) visit(item.children)
    }
  }
  visit(items)
  return result
}
