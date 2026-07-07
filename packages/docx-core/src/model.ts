/**
 * DOCX Document Model — Immutable editing functions
 *
 * Pure functions for manipulating the DocModel tree.
 * All operations are immutable (return new model, don't mutate).
 */

import type {
  DocModel, DocNode, ParagraphNode, TableNode, TableCellNode,
  ParagraphChildNode, TextRunNode, ImageRunNode,
  ParagraphStyle, TextStyle,
} from "./types"

// ============================================================
// Clone
// ============================================================

export function cloneDocModel(model: DocModel): DocModel {
  return {
    nodes: model.nodes.map(cloneDocNode),
    metadata: {
      ...model.metadata,
      headerSections: model.metadata.headerSections.map(s => ({
        ...s,
        nodes: s.nodes.map(cloneDocNode),
      })),
      footerSections: model.metadata.footerSections.map(s => ({
        ...s,
        nodes: s.nodes.map(cloneDocNode),
      })),
      sections: model.metadata.sections?.map(s => ({
        ...s,
        headerSections: s.headerSections.map(h => ({
          ...h,
          nodes: h.nodes.map(cloneDocNode),
        })),
        footerSections: s.footerSections.map(f => ({
          ...f,
          nodes: f.nodes.map(cloneDocNode),
        })),
      })),
      footnotes: model.metadata.footnotes?.map(n => ({
        ...n,
        nodes: n.nodes?.map(cloneDocNode),
      })),
      endnotes: model.metadata.endnotes?.map(n => ({
        ...n,
        nodes: n.nodes?.map(cloneDocNode),
      })),
    },
  }
}

function cloneDocNode(node: DocNode): DocNode {
  if (node.type === "paragraph") {
    return {
      ...node,
      children: node.children.map(cloneParagraphChild),
    }
  }
  return {
    ...node,
    rows: node.rows.map(row => ({
      ...row,
      cells: row.cells.map(cell => ({
        ...cell,
        nodes: cell.nodes.map(n => {
          if (n.type === "table") return cloneDocNode(n)
          return {
            ...n,
            children: n.children.map(cloneParagraphChild),
          }
        }),
      })),
    })),
  }
}

function sameTextStyle(a?: TextStyle, b?: TextStyle): boolean {
  return JSON.stringify(a ?? {}) === JSON.stringify(b ?? {})
}

function mergeAdjacentTextRuns(children: ParagraphChildNode[]): ParagraphChildNode[] {
  const merged: ParagraphChildNode[] = []
  for (const child of children) {
    const previous = merged[merged.length - 1]
    if (child.type === "text" && previous?.type === "text" && sameTextStyle(previous.style, child.style)) {
      previous.text += child.text
    } else {
      merged.push(cloneParagraphChild(child))
    }
  }
  return merged
}

function cloneParagraphChild(child: ParagraphChildNode): ParagraphChildNode {
  if (child.type === "image") {
    const data = child.data ? (child.data instanceof Uint8Array ? new Uint8Array(child.data) : child.data) : undefined
    return { ...child, data }
  }
  return { ...child }
}

// ============================================================
// Normalize (legacy array data → Uint8Array)
// ============================================================

export function normalizeDocModel(model: DocModel): DocModel {
  return cloneDocModel(model)
}

// ============================================================
// Paragraph Insertion / Removal / Duplication
// ============================================================

export function insertParagraph(
  model: DocModel,
  text: string,
  index?: number,
  options?: { style?: ParagraphStyle },
): DocModel {
  const pos = index ?? model.nodes.length
  const newParagraph: ParagraphNode = {
    type: "paragraph",
    children: [{ type: "text", text }],
    style: options?.style,
  }
  const nodes = [...model.nodes]
  nodes.splice(pos, 0, newParagraph)
  return { ...model, nodes }
}

export function removeParagraph(model: DocModel, index: number): DocModel {
  if (index < 0 || index >= model.nodes.length) return model
  const nodes = [...model.nodes]
  nodes.splice(index, 1)
  return { ...model, nodes }
}

export function duplicateParagraph(model: DocModel, index: number): DocModel {
  if (index < 0 || index >= model.nodes.length) return model
  const original = model.nodes[index]
  const copy = cloneDocNode(original)
  const nodes = [...model.nodes]
  nodes.splice(index + 1, 0, copy)
  return { ...model, nodes }
}

export function copyParagraphs(model: DocModel, startIndex: number, endIndex?: number): ParagraphNode[] {
  const end = endIndex ?? startIndex + 1
  return model.nodes.slice(startIndex, end)
    .filter(n => n.type === "paragraph")
    .map(n => cloneDocNode(n) as ParagraphNode)
}

export function pasteParagraphs(
  model: DocModel,
  index: number,
  paragraphs: ParagraphNode[],
): DocModel {
  const nodes = [...model.nodes]
  nodes.splice(index, 0, ...paragraphs.map(p => cloneDocNode(p)))
  return { ...model, nodes }
}

// ============================================================
// Paragraph Text Updates
// ============================================================

export function updateParagraphText(
  model: DocModel,
  nodeIndex: number,
  text: string,
  options?: { style?: TextStyle },
): DocModel {
  const node = model.nodes[nodeIndex]
  if (!node || node.type !== "paragraph") return model

  const textChildren = node.children.filter((child): child is TextRunNode => child.type === "text")
  const oldText = textChildren.map(child => child.text).join("")
  if (oldText === text) return model

  let prefix = 0
  const maxPrefix = Math.min(oldText.length, text.length)
  while (prefix < maxPrefix && oldText[prefix] === text[prefix]) prefix++

  let suffix = 0
  const maxSuffix = Math.min(oldText.length - prefix, text.length - prefix)
  while (suffix < maxSuffix && oldText[oldText.length - 1 - suffix] === text[text.length - 1 - suffix]) suffix++

  const replaceStart = prefix
  const replaceEnd = oldText.length - suffix
  const insertedText = text.slice(prefix, text.length - suffix)

  let textOffset = 0
  let inserted = false
  const nextChildren: ParagraphChildNode[] = []

  const insertStyle = options?.style ?? textChildren.find((child, index) => {
    const start = textChildren.slice(0, index).reduce((sum, run) => sum + run.text.length, 0)
    return replaceStart <= start + child.text.length
  })?.style ?? textChildren[textChildren.length - 1]?.style

  for (const child of node.children) {
    if (child.type !== "text") {
      nextChildren.push(cloneParagraphChild(child))
      continue
    }

    const runStart = textOffset
    const runEnd = textOffset + child.text.length
    textOffset = runEnd

    if (runEnd <= replaceStart || runStart >= replaceEnd) {
      nextChildren.push(cloneParagraphChild(child))
      continue
    }

    const keepBefore = child.text.slice(0, Math.max(0, replaceStart - runStart))
    const keepAfter = child.text.slice(Math.max(0, replaceEnd - runStart))

    if (keepBefore) nextChildren.push({ ...child, text: keepBefore })
    if (!inserted) {
      if (insertedText) nextChildren.push({ type: "text", text: insertedText, style: insertStyle })
      inserted = true
    }
    if (keepAfter) nextChildren.push({ ...child, text: keepAfter })
  }

  if (!inserted && insertedText) {
    nextChildren.push({ type: "text", text: insertedText, style: insertStyle })
  }

  const updatedParagraph: ParagraphNode = {
    ...node,
    children: mergeAdjacentTextRuns(nextChildren),
  }

  const nodes = [...model.nodes]
  nodes[nodeIndex] = updatedParagraph
  return { ...model, nodes }
}

export function updateTableCellText(
  model: DocModel,
  tableIndex: number,
  rowIndex: number,
  cellIndex: number,
  text: string,
  options?: { style?: TextStyle },
): DocModel {
  const node = model.nodes[tableIndex]
  if (!node || node.type !== "table") return model

  const rows = [...node.rows]
  const row = { ...rows[rowIndex], cells: [...rows[rowIndex].cells] }
  const cell = { ...row.cells[cellIndex], nodes: [...row.cells[cellIndex].nodes] }
  rows[rowIndex] = row

  if (cell.nodes[0]?.type === "paragraph") {
    cell.nodes[0] = {
      ...cell.nodes[0],
      children: [{ type: "text", text, style: options?.style }],
    }
  }

  const updatedTable: TableNode = { ...node, rows }
  const nodes = [...model.nodes]
  nodes[tableIndex] = updatedTable
  return { ...model, nodes }
}

// ============================================================
// Text Replacement (find & replace)
// ============================================================

export function replaceText(
  model: DocModel,
  searchValue: string | RegExp,
  replacement: string,
): DocModel {
  const nodes = model.nodes.map(node => {
    if (node.type === "paragraph") {
      return {
        ...node,
        children: node.children.map(child => {
          if (child.type === "text") {
            return {
              ...child,
              text: child.text.replaceAll(searchValue, replacement),
            }
          }
          return child
        }),
      }
    }
    return node
  })
  return { ...model, nodes }
}

// ============================================================
// Paragraph Formatting
// ============================================================

export function setParagraphHeading(
  model: DocModel,
  nodeIndex: number,
  headingLevel?: 1 | 2 | 3 | 4 | 5 | 6,
): DocModel {
  const node = model.nodes[nodeIndex]
  if (!node || node.type !== "paragraph") return model

  const updatedParagraph: ParagraphNode = {
    ...node,
    style: {
      ...node.style,
      align: node.style?.align,
      headingLevel: headingLevel || undefined,
    },
  }

  const nodes = [...model.nodes]
  nodes[nodeIndex] = updatedParagraph
  return { ...model, nodes }
}

export function setParagraphAlignment(
  model: DocModel,
  nodeIndex: number,
  align?: "left" | "center" | "right" | "justify",
): DocModel {
  const node = model.nodes[nodeIndex]
  if (!node || node.type !== "paragraph") return model

  const updatedParagraph: ParagraphNode = {
    ...node,
    style: { ...node.style, align },
  }

  const nodes = [...model.nodes]
  nodes[nodeIndex] = updatedParagraph
  return { ...model, nodes }
}

// ============================================================
// Run-level Text Formatting
// ============================================================

export function applyRunStyle(
  model: DocModel,
  nodeIndex: number,
  runIndex: number,
  style: Partial<TextStyle>,
): DocModel {
  const node = model.nodes[nodeIndex]
  if (!node || node.type !== "paragraph" || !node.children[runIndex] || node.children[runIndex].type !== "text")
    return model

  const children = [...node.children]
  children[runIndex] = {
    ...children[runIndex],
    style: { ...(children[runIndex] as TextRunNode).style, ...style },
  } as ParagraphChildNode

  const nodes = [...model.nodes]
  nodes[nodeIndex] = { ...node, children }
  return { ...model, nodes }
}

export function toggleRunStyleFlag(
  model: DocModel,
  nodeIndex: number,
  runIndex: number,
  key: "bold" | "italic" | "underline" | "strike",
): DocModel {
  const node = model.nodes[nodeIndex]
  if (!node || node.type !== "paragraph" || !node.children[runIndex] || node.children[runIndex].type !== "text")
    return model

  const run = node.children[runIndex] as TextRunNode
  const current = run.style?.[key] ?? false

  return applyRunStyle(model, nodeIndex, runIndex, { [key]: !current })
}

export function setRunHighlight(
  model: DocModel,
  nodeIndex: number,
  runIndex: number,
  highlight?: string,
): DocModel {
  return applyRunStyle(model, nodeIndex, runIndex, { highlight })
}

export function setRunColor(
  model: DocModel,
  nodeIndex: number,
  runIndex: number,
  color?: string,
): DocModel {
  return applyRunStyle(model, nodeIndex, runIndex, { color })
}

// ============================================================
// Table Cell Text Update (recursive for nested tables)
// ============================================================

export function updateTableCellParagraphText(
  model: DocModel,
  tableIndex: number,
  rowIndex: number,
  cellIndex: number,
  paragraphIndex: number,
  text: string,
  options?: { style?: TextStyle },
): DocModel {
  const node = model.nodes[tableIndex]
  if (!node || node.type !== "table") return model

  const rows = [...node.rows]
  const row = { ...rows[rowIndex], cells: [...rows[rowIndex].cells] }
  const cell = { ...row.cells[cellIndex], nodes: [...row.cells[cellIndex].nodes] }
  rows[rowIndex] = row

  if (cell.nodes[paragraphIndex]?.type === "paragraph") {
    cell.nodes[paragraphIndex] = {
      ...cell.nodes[paragraphIndex],
      children: [{ type: "text", text, style: options?.style }],
    }
  }

  const updatedTable: TableNode = { ...node, rows }
  const nodes = [...model.nodes]
  nodes[tableIndex] = updatedTable
  return { ...model, nodes }
}

// ============================================================
// Clipboard Serialization
// ============================================================

export function serializeParagraphsForClipboard(paragraphs: ParagraphNode[]): string {
  return JSON.stringify(paragraphs)
}

export function parseParagraphsFromClipboard(input: string): ParagraphNode[] | undefined {
  try {
    const parsed = JSON.parse(input)
    if (Array.isArray(parsed) && parsed.every(p => p?.type === "paragraph")) {
      return parsed
    }
  } catch {
    return undefined
  }
}

// ============================================================
// Split Paragraph Children at Text Offsets
// ============================================================

export function splitParagraphChildrenAtTextOffsets(
  paragraph: ParagraphNode,
  text: string,
  startOffset: number,
  endOffset: number,
): { beforeChildren: ParagraphChildNode[]; afterChildren: ParagraphChildNode[] } {
  const beforeChildren: ParagraphChildNode[] = []
  const afterChildren: ParagraphChildNode[] = []
  let charOffset = 0

  for (const child of paragraph.children) {
    const childText = child.type === "text" ? child.text : ""
    const childStart = charOffset
    const childEnd = charOffset + childText.length

    if (childEnd <= startOffset) {
      beforeChildren.push(child)
    } else if (childStart >= endOffset) {
      afterChildren.push(child)
    } else {
      // Child is split by the selection
      const beforeText = childText.slice(0, startOffset - childStart)
      const afterText = childText.slice(endOffset - childStart)
      if (beforeText) {
        beforeChildren.push({ ...child, text: beforeText } as ParagraphChildNode)
      }
      if (afterText) {
        afterChildren.push({ ...child, text: afterText } as ParagraphChildNode)
      }
    }

    charOffset = childEnd
  }

  return { beforeChildren, afterChildren }
}

// ============================================================
// Default Starter Model
// ============================================================

export function createBlankDocumentModel(): DocModel {
  return {
    nodes: [{ type: "paragraph", children: [{ type: "text", text: "" }] }],
    metadata: {
      headerSections: [],
      footerSections: [],
    },
  }
}

export const defaultStarterModel: DocModel = createBlankDocumentModel()
