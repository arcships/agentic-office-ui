// editor-clone-helpers.ts — copy-on-write helpers shared by the edit commands.
//
// Every interactive command edits a bounded number of top-level nodes. These
// helpers turn a model plus a location into a shallow model whose touched nodes
// are safe to mutate, so the command never clones the whole document graph.

import type { DocModel } from "@arcships/docx-core"
import type { ParagraphLocation } from "@arcships/docx-core"
import {
  cloneDocModelNodes,
  parseEmbeddedTableRuntimeKey,
} from "@arcships/docx-core"

/**
 * Index of the top-level node that owns a paragraph location: the paragraph
 * itself in the body, or the table when the paragraph lives in a cell.
 */
export function topLevelNodeIndex(location: ParagraphLocation): number {
  return location.kind === "table-cell" ? location.tableIndex : location.nodeIndex
}

/**
 * Copy-on-write model whose host table owns the given runtime key, or
 * `undefined` when the key does not resolve to a table.
 *
 * Nested tables live inside their host table node, so cloning the host table
 * gives the runtime-key helpers a private graph to mutate.
 */
export function cloneHostTableForRuntimeKey(
  model: DocModel,
  tableRuntimeKey: string
): DocModel | undefined {
  const location = parseEmbeddedTableRuntimeKey(tableRuntimeKey)
  if (!location) return undefined

  const hostTable = model.nodes[location.hostTableIndex]
  if (!hostTable || hostTable.type !== "table") return undefined

  return cloneDocModelNodes(model, [location.hostTableIndex])
}
