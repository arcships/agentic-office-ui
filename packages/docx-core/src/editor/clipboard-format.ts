// clipboard-format.ts — structured (internal MIME) DOCX clipboard payload plus
// the model operations that read a selection into it and paste it back.
//
// The payload carries the source `DocxTextRange` and the runs themselves, so a
// copy inside the editor round-trips exact styles and links instead of relying
// on the lossy `text/html` mapping. `text/html` and `text/plain` stay the
// cross-application flavours (see clipboard-html.ts).

import type {
  DocModel,
  ParagraphChildNode,
  ParagraphNode,
  ParagraphStyle,
  TextRunNode
} from "../engine/types"
import type { DocxTextRange, ParagraphLocation } from "./helpers/editor-types"
import { cloneDocModelShallow, cloneDocNode, cloneNumberingDefinitions } from "../engine/clone"
import { remapParagraphCommentRanges } from "../engine/comment-ranges"
import { invalidateParagraphSourceXml } from "../engine/section-break"
import {
  cloneParagraphStyle,
  firstRunStyle,
  firstTextStyleAtOffset,
  paragraphText
} from "./helpers"
import {
  cloneTextRangeLocation,
  getParagraphAtLocation,
  sameParagraphLocation
} from "./helpers/selection-helpers"
import { ensureListNumberingId } from "./helpers/list-numbering"
import { compareTextRangeBoundaries, normalizeTextRange } from "./helpers/editor-types"
import { splitParagraphChildrenAtTextOffsets } from "./paragraph-ops"
import { pasteParagraphs } from "./run-style-ops"
import type {
  DocxClipboardParagraph,
  DocxClipboardParseOptions,
  DocxClipboardRun
} from "./clipboard-html"
import {
  MAX_DOCX_CLIPBOARD_PARAGRAPHS,
  MAX_DOCX_CLIPBOARD_RUNS,
  clipboardTextFromParagraphs,
  sanitizeClipboardParagraphs
} from "./clipboard-html"

/** Internal flavour written next to `text/html` and `text/plain`. */
export const DOCX_CLIPBOARD_MIME = "application/x-arcships-docx-fragment+json"
/** Payload schema version; bump when the shape changes incompatibly. */
export const DOCX_CLIPBOARD_VERSION = 1
/** Cap for the serialized internal payload (also caps `text/plain` reads). */
export const MAX_DOCX_CLIPBOARD_PAYLOAD_CHARS = 512 * 1024

export interface DocxClipboardPayload {
  version: typeof DOCX_CLIPBOARD_VERSION
  kind: "docx-fragment"
  text: string
  /** Source range; informational for consumers, not required to paste. */
  range?: DocxTextRange
  paragraphs: DocxClipboardParagraph[]
}

export interface CreateDocxClipboardPayloadInput {
  paragraphs: readonly DocxClipboardParagraph[]
  range?: DocxTextRange
}

/** Builds the internal payload for a copied selection. */
export function createDocxClipboardPayload(
  input: CreateDocxClipboardPayloadInput
): DocxClipboardPayload {
  const paragraphs = input.paragraphs.map((paragraph) => {
    const copy: DocxClipboardParagraph = {
      runs: paragraph.runs.map((run) => ({
        text: run.text,
        ...(run.style ? { style: { ...run.style } } : undefined),
        ...(run.link ? { link: run.link } : undefined)
      }))
    }
    if (paragraph.style) copy.style = { ...paragraph.style }
    if (paragraph.list) copy.list = { ...paragraph.list }
    return copy
  })
  const payload: DocxClipboardPayload = {
    version: DOCX_CLIPBOARD_VERSION,
    kind: "docx-fragment",
    text: clipboardTextFromParagraphs(paragraphs),
    paragraphs
  }
  if (input.range) {
    payload.range = {
      start: {
        location: cloneTextRangeLocation(input.range.start.location),
        offset: input.range.start.offset
      },
      end: {
        location: cloneTextRangeLocation(input.range.end.location),
        offset: input.range.end.offset
      }
    }
  }
  return payload
}

/** Serializes the internal payload; callers cap the size before writing it. */
export function serializeDocxClipboardPayload(payload: DocxClipboardPayload): string {
  return JSON.stringify(payload)
}

function sanitizeClipboardRangeBoundary(value: unknown): DocxTextRange["start"] | undefined {
  if (!value || typeof value !== "object") return undefined
  const boundary = value as Record<string, unknown>
  if (typeof boundary.offset !== "number" || !Number.isFinite(boundary.offset)) {
    return undefined
  }
  const offset = Math.max(0, Math.round(boundary.offset))
  const location = boundary.location
  if (!location || typeof location !== "object") return undefined
  const source = location as Record<string, unknown>
  const index = (key: string): number | undefined => {
    const candidate = source[key]
    return typeof candidate === "number" && Number.isInteger(candidate) && candidate >= 0
      ? candidate
      : undefined
  }
  if (source.kind === "paragraph") {
    const nodeIndex = index("nodeIndex")
    if (nodeIndex === undefined) return undefined
    return { location: { kind: "paragraph", nodeIndex }, offset }
  }
  if (source.kind === "table-cell") {
    const tableIndex = index("tableIndex")
    const rowIndex = index("rowIndex")
    const cellIndex = index("cellIndex")
    const paragraphIndex = index("paragraphIndex")
    if (
      tableIndex === undefined ||
      rowIndex === undefined ||
      cellIndex === undefined ||
      paragraphIndex === undefined
    ) {
      return undefined
    }
    return {
      location: { kind: "table-cell", tableIndex, rowIndex, cellIndex, paragraphIndex },
      offset
    }
  }
  return undefined
}

function sanitizeClipboardRange(value: unknown): DocxTextRange | undefined {
  if (!value || typeof value !== "object") return undefined
  const source = value as Record<string, unknown>
  const start = sanitizeClipboardRangeBoundary(source.start)
  const end = sanitizeClipboardRangeBoundary(source.end)
  if (!start || !end) return undefined
  return { start, end }
}

/**
 * Parses the internal clipboard flavour.
 *
 * Everything is re-validated: the clipboard is external input, so the payload,
 * paragraph and run caps, the style whitelist and the caller's link resolver
 * apply exactly as they do for `text/html`.
 */
export function parseDocxClipboardPayload(
  raw: unknown,
  options?: DocxClipboardParseOptions & { maxPayloadChars?: number }
): DocxClipboardPayload | undefined {
  if (typeof raw !== "string" || !raw) return undefined
  const maxChars = Math.max(1, options?.maxPayloadChars ?? MAX_DOCX_CLIPBOARD_PAYLOAD_CHARS)
  if (raw.length > maxChars) return undefined

  let parsed: unknown
  try {
    parsed = JSON.parse(raw)
  } catch {
    return undefined
  }
  if (!parsed || typeof parsed !== "object") return undefined
  const source = parsed as Record<string, unknown>
  if (source.version !== DOCX_CLIPBOARD_VERSION || source.kind !== "docx-fragment") {
    return undefined
  }

  const paragraphs = sanitizeClipboardParagraphs(source.paragraphs, {
    ...options,
    maxParagraphs: options?.maxParagraphs ?? MAX_DOCX_CLIPBOARD_PARAGRAPHS,
    maxRuns: options?.maxRuns ?? MAX_DOCX_CLIPBOARD_RUNS
  })
  if (paragraphs.length === 0) return undefined

  const payload: DocxClipboardPayload = {
    version: DOCX_CLIPBOARD_VERSION,
    kind: "docx-fragment",
    text:
      typeof source.text === "string" ? source.text : clipboardTextFromParagraphs(paragraphs),
    paragraphs
  }
  const range = sanitizeClipboardRange(source.range)
  if (range) payload.range = range
  return payload
}

// ── selection → clipboard paragraphs ───────────────────────────────

function clampOffset(value: number, length: number): number {
  if (!Number.isFinite(value)) return 0
  return Math.max(0, Math.min(length, Math.round(value)))
}

/** Copies a run style without the editor clone's undefined placeholders. */
function copyRunStyle(style: TextRunNode["style"]): TextRunNode["style"] | undefined {
  return style ? { ...style } : undefined
}

/** Slices text runs by paragraph character offsets, keeping style and link. */
function sliceRunsAtOffsets(
  children: readonly ParagraphChildNode[],
  from: number,
  to: number
): DocxClipboardRun[] {
  const runs: DocxClipboardRun[] = []
  if (to <= from) return runs
  let offset = 0
  for (const child of children) {
    if (child.type !== "text") continue
    const runStart = offset
    const runEnd = offset + child.text.length
    offset = runEnd
    if (runEnd <= from) continue
    if (runStart >= to) break
    const start = Math.max(from, runStart) - runStart
    const end = Math.min(to, runEnd) - runStart
    if (end <= start) continue
    const run: DocxClipboardRun = { text: child.text.slice(start, end) }
    const style = copyRunStyle(child.style)
    if (style) run.style = style
    if (child.link) run.link = child.link
    runs.push(run)
  }
  return runs
}

function clipboardParagraphForNode(
  node: ParagraphNode,
  from: number,
  to: number
): DocxClipboardParagraph {
  const paragraph: DocxClipboardParagraph = { runs: sliceRunsAtOffsets(node.children, from, to) }
  const style: ParagraphStyle = {}
  if (node.style?.headingLevel) style.headingLevel = node.style.headingLevel
  if (node.style?.align) style.align = node.style.align
  if (Object.keys(style).length > 0) paragraph.style = style
  // Only real numbering becomes a list: text-prefix bullets ("• item") are
  // already part of the text and would otherwise get a second marker on paste.
  const numbering = node.style?.numbering
  if (numbering && numbering.numId > 0) {
    paragraph.list = {
      type: "unordered",
      level: Math.max(0, Math.min(8, numbering.ilvl ?? 0))
    }
  }
  return paragraph
}

/**
 * Reads the selected range into clipboard paragraphs.
 *
 * Mirrors `resolveSelectedTextForRange`: a body selection may span paragraphs
 * (tables in between are skipped), and a table-cell selection must stay in one
 * cell. Anything else returns `undefined`, so `copy()` can fall back to the
 * plain-text flavour.
 */
export function docxClipboardParagraphsForRange(
  model: DocModel,
  range: DocxTextRange
): DocxClipboardParagraph[] | undefined {
  const normalized = normalizeTextRange(range)
  const start = normalized.start
  const end = normalized.end
  if (compareTextRangeBoundaries(start, end) >= 0) return undefined

  if (start.location.kind === "paragraph" && end.location.kind === "paragraph") {
    const paragraphs: DocxClipboardParagraph[] = []
    for (let index = start.location.nodeIndex; index <= end.location.nodeIndex; index += 1) {
      const node = model.nodes[index]
      if (!node || node.type !== "paragraph") continue
      const text = paragraphText(node)
      const from = index === start.location.nodeIndex ? clampOffset(start.offset, text.length) : 0
      const to = index === end.location.nodeIndex ? clampOffset(end.offset, text.length) : text.length
      paragraphs.push(clipboardParagraphForNode(node, from, to))
    }
    return paragraphs.length > 0 ? paragraphs : undefined
  }

  if (start.location.kind === "table-cell" && end.location.kind === "table-cell") {
    const sameCell =
      start.location.tableIndex === end.location.tableIndex &&
      start.location.rowIndex === end.location.rowIndex &&
      start.location.cellIndex === end.location.cellIndex
    if (!sameCell) return undefined

    const table = model.nodes[start.location.tableIndex]
    if (!table || table.type !== "table") return undefined
    const cell = table.rows[start.location.rowIndex]?.cells[start.location.cellIndex]
    if (!cell) return undefined

    const paragraphs: DocxClipboardParagraph[] = []
    const cellParagraphs = cell.nodes.filter(
      (node): node is ParagraphNode => node.type === "paragraph"
    )
    for (
      let index = start.location.paragraphIndex;
      index <= end.location.paragraphIndex;
      index += 1
    ) {
      const node = cellParagraphs[index]
      if (!node) continue
      const text = paragraphText(node)
      const from = index === start.location.paragraphIndex ? clampOffset(start.offset, text.length) : 0
      const to = index === end.location.paragraphIndex ? clampOffset(end.offset, text.length) : text.length
      paragraphs.push(clipboardParagraphForNode(node, from, to))
    }
    return paragraphs.length > 0 ? paragraphs : undefined
  }

  return undefined
}

// ── paste insertion ────────────────────────────────────────────────

export interface InsertDocxClipboardResult {
  model: DocModel
  /** Collapsed caret after the inserted content. */
  caret: DocxTextRange
}

function runsForBlock(block: DocxClipboardParagraph): TextRunNode[] {
  const runs: TextRunNode[] = []
  for (const run of block.runs) {
    if (!run.text) continue
    const textRun: TextRunNode = { type: "text", text: run.text }
    if (run.style) textRun.style = { ...run.style }
    if (run.link) textRun.link = run.link
    runs.push(textRun)
  }
  if (runs.length === 0) {
    // An empty pasted paragraph is still a paragraph.
    runs.push({ type: "text", text: "" })
  }
  return runs
}

function blockTextLength(block: DocxClipboardParagraph): number {
  return block.runs.reduce((total, run) => total + run.text.length, 0)
}

function blockParagraphStyle(block: DocxClipboardParagraph): ParagraphStyle | undefined {
  const style: ParagraphStyle = {}
  if (block.style?.headingLevel) style.headingLevel = block.style.headingLevel
  if (block.style?.align) style.align = block.style.align
  return Object.keys(style).length > 0 ? style : undefined
}

/** Applies the pasted paragraph's heading/alignment when it owns the paragraph. */
function applyBlockParagraphStyle(
  paragraph: ParagraphNode,
  block: DocxClipboardParagraph
): void {
  const style = blockParagraphStyle(block)
  if (!style) return
  paragraph.style = { ...(paragraph.style ?? {}), ...style }
  paragraph.sourceXml = undefined
}

/** Drops the empty placeholder runs a split leaves at either edge. */
function trimEmptyEdgeRuns(children: ParagraphChildNode[]): ParagraphChildNode[] {
  while (children.length > 0) {
    const first = children[0]
    if (first.type === "text" && first.text === "") {
      children.shift()
      continue
    }
    break
  }
  while (children.length > 0) {
    const last = children[children.length - 1]
    if (last.type === "text" && last.text === "") {
      children.pop()
      continue
    }
    break
  }
  return children
}

function blockText(block: DocxClipboardParagraph): string {
  return block.runs.map((run) => run.text).join("")
}

function hasBlockList(paragraphs: readonly DocxClipboardParagraph[]): boolean {
  return paragraphs.some((block) => Boolean(block.list))
}

function applyBlockList(
  model: DocModel,
  paragraph: ParagraphNode,
  block: DocxClipboardParagraph
): void {
  if (!block.list) return
  const numId = ensureListNumberingId(model, block.list.type)
  paragraph.style = {
    ...(paragraph.style ?? {}),
    numbering: { numId, ilvl: Math.max(0, Math.min(8, block.list.level ?? 0)) }
  }
  paragraph.sourceXml = undefined
}

function emptyCaret(location: ParagraphLocation, offset: number): DocxTextRange {
  return {
    start: { location: cloneTextRangeLocation(location), offset },
    end: { location: cloneTextRangeLocation(location), offset }
  }
}

/**
 * Inserts clipboard paragraphs at a range inside one paragraph or table cell.
 *
 * The caller collapses expanded, cross-paragraph selections first, so this
 * function only has to deal with one target paragraph: its text before the
 * range, the pasted runs and its text after the range. Body pastes with several
 * paragraphs split the target paragraph; table-cell pastes stay in one cell
 * paragraph with `\n` between the pasted blocks, which is what a plain-text
 * paste already did.
 */
export function insertDocxClipboardParagraphs(
  model: DocModel,
  range: DocxTextRange,
  paragraphs: readonly DocxClipboardParagraph[]
): InsertDocxClipboardResult | undefined {
  if (paragraphs.length === 0) return undefined

  const normalized = normalizeTextRange(range)
  const startLocation = normalized.start.location
  const endLocation = normalized.end.location
  if (!sameParagraphLocation(startLocation, endLocation)) return undefined

  const location = cloneTextRangeLocation(startLocation)
  const lookup = getParagraphAtLocation(model, location)
  const paragraph = lookup.paragraph
  if (!paragraph) return undefined

  const text = paragraphText(paragraph)
  const safeStart = clampOffset(normalized.start.offset, text.length)
  const safeEnd = clampOffset(normalized.end.offset, text.length)
  const beforeInsertedStyle =
    firstTextStyleAtOffset(paragraph, safeStart, true) ?? firstRunStyle(paragraph)
  const afterInsertedStyle =
    firstTextStyleAtOffset(paragraph, safeEnd, false) ?? firstRunStyle(paragraph)
  const { beforeChildren, afterChildren } = splitParagraphChildrenAtTextOffsets(
    paragraph,
    text,
    safeStart,
    safeEnd,
    { beforeInsertedStyle, afterInsertedStyle }
  )

  const nextModel = cloneDocModelShallow(model)
  let target: ParagraphNode | undefined
  if (location.kind === "paragraph") {
    target = cloneDocNode(paragraph)
    nextModel.nodes[location.nodeIndex] = target
  } else {
    const tableNode = lookup.tableNode
    if (!tableNode) return undefined
    const tableClone = cloneDocNode(tableNode)
    nextModel.nodes[location.tableIndex] = tableClone
    const cell = tableClone.rows[location.rowIndex]?.cells[location.cellIndex]
    const cellParagraph = cell?.nodes.filter(
      (node): node is ParagraphNode => node.type === "paragraph"
    )[location.paragraphIndex]
    if (!cellParagraph) return undefined
    tableClone.sourceXml = undefined
    target = cellParagraph
  }
  if (!target) return undefined

  const applyChildren = (children: ParagraphChildNode[]): void => {
    const nextChildren = trimEmptyEdgeRuns(children)
    const nextText = nextChildren
      .filter((child): child is TextRunNode => child.type === "text")
      .map((child) => child.text)
      .join("")
    target.children = nextChildren
    target.commentRanges = remapParagraphCommentRanges(text, nextText, target.commentRanges)
    invalidateParagraphSourceXml(target)
  }

  // A table cell owns a list of blocks; adding paragraphs there would change
  // the cell structure, so the blocks join into the one target paragraph.
  const singleParagraph = location.kind !== "paragraph" || paragraphs.length === 1
  if (singleParagraph) {
    const inserted: ParagraphChildNode[] = []
    paragraphs.forEach((block, index) => {
      if (location.kind !== "paragraph" && index > 0) {
        inserted.push({ type: "text", text: "\n" })
      }
      inserted.push(...runsForBlock(block))
    })
    applyChildren([...beforeChildren, ...inserted, ...afterChildren])

    if (location.kind === "paragraph") {
      const survivingText = `${text.slice(0, safeStart)}${text.slice(safeEnd)}`
      if (!survivingText.trim()) {
        // The paste owns the whole paragraph, so its heading/alignment applies.
        applyBlockParagraphStyle(target, paragraphs[0])
        if (paragraphs[0].list) {
          nextModel.metadata.numberingDefinitions = cloneNumberingDefinitions(
            model.metadata.numberingDefinitions
          )
          applyBlockList(nextModel, target, paragraphs[0])
        }
      }
    }

    const insertedTextLength =
      paragraphs.reduce((total, block) => total + blockTextLength(block), 0) +
      (location.kind === "paragraph" ? 0 : Math.max(0, paragraphs.length - 1))
    return {
      model: nextModel,
      caret: emptyCaret(location, safeStart + insertedTextLength)
    }
  }

  applyChildren([...beforeChildren, ...runsForBlock(paragraphs[0])])

  const inserted: ParagraphNode[] = []
  for (let index = 1; index < paragraphs.length; index += 1) {
    const block = paragraphs[index]
    const isLast = index === paragraphs.length - 1
    inserted.push({
      type: "paragraph",
      // The trailing text after the caret belongs to the target paragraph, so
      // the last pasted block keeps the target's style; earlier blocks keep
      // their own heading/alignment.
      style: isLast ? cloneParagraphStyle(paragraph.style) : blockParagraphStyle(block),
      children: trimEmptyEdgeRuns(
        isLast ? [...runsForBlock(block), ...afterChildren] : runsForBlock(block)
      )
    })
  }

  // The model owns numbering definitions before list pastes mutate them, and
  // `pasteParagraphs` copies `metadata` by reference.
  if (hasBlockList(paragraphs)) {
    nextModel.metadata.numberingDefinitions = cloneNumberingDefinitions(
      model.metadata.numberingDefinitions
    )
  }
  inserted.forEach((node, index) => {
    applyBlockList(nextModel, node, paragraphs[index + 1])
  })
  const survivingText = `${text.slice(0, safeStart)}${text.slice(safeEnd)}`
  if (!survivingText.trim()) {
    applyBlockParagraphStyle(target, paragraphs[0])
    applyBlockList(nextModel, target, paragraphs[0])
  }

  const pasted = pasteParagraphs(nextModel, location.nodeIndex + 1, inserted)
  const caretLocation: ParagraphLocation = {
    kind: "paragraph",
    nodeIndex: location.nodeIndex + inserted.length
  }
  const caretOffset = blockTextLength(paragraphs[paragraphs.length - 1])
  return {
    model: pasted,
    caret: emptyCaret(caretLocation, caretOffset)
  }
}
