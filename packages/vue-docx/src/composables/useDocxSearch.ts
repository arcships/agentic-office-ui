import {
  paragraphText,
  type DocModel,
  type DocNode,
  type DocxTextRangeLocation,
  type FooterSection,
  type HeaderSection,
  type ParagraphNode,
  type TableCellContentNode,
  type TableNode,
} from "@arcships/docx-core"
import {
  createSurfaceSearchExpression,
} from "@arcships/office-runtime"

/** Query options accepted by the DOCX search session. */
export interface DocxSearchOptions {
  caseSensitive?: boolean
  wholeWord?: boolean
  regex?: boolean
}

/** Story a search match belongs to. */
export type DocxSearchRegion = "body" | "header" | "footer" | "footnote" | "endnote"

/** One step from a table into a table nested inside one of its cells. */
export interface DocxNestedTableStep {
  rowIndex: number
  cellIndex: number
  /** Content-node index of the nested table inside the cell's `nodes`. */
  contentIndex: number
}

/**
 * Search-specific location for a paragraph inside a nested table.
 *
 * `DocxTextRangeLocation` addresses body paragraphs and direct cell paragraphs
 * only, so nested-table matches carry their nesting path. The path is rendered
 * as `data-docx-table-nesting-path` by `DocxTableHost`, which lets the surface
 * project a match back into the DOM without a core model change or a
 * page-specific runtime key.
 */
export interface ParagraphLocationInNestedTableCell {
  kind: "nested-table-cell"
  /** Top-level table node index within the story (body/header/footer/note). */
  tableIndex: number
  /** Nesting path from the top-level table into the innermost table. */
  path: readonly DocxNestedTableStep[]
  rowIndex: number
  cellIndex: number
  paragraphIndex: number
  /** Slash-separated `rowIndex-cellIndex-contentIndex` DOM nesting path. */
  nestingPath: string
}

/** Note body that only carries plain `text` (no paragraph nodes). */
export interface ParagraphLocationInNoteText {
  kind: "note-text"
  noteId: number
}

export type DocxSearchLocation =
  | DocxTextRangeLocation
  | ParagraphLocationInNestedTableCell
  | ParagraphLocationInNoteText

export interface DocxSearchBoundary {
  location: DocxSearchLocation
  offset: number
}

export interface DocxSearchRange {
  start: DocxSearchBoundary
  end: DocxSearchBoundary
}

export interface DocxSearchMatch {
  kind: "docx-text"
  /** Story that owns the match: body, running head/foot, or a note. */
  region: DocxSearchRegion
  /** Top-level body/story node index (table index for table content). */
  nodeIndex: number
  /** Header/footer part name; empty for body and notes. */
  partName?: string
  /** Footnote/endnote id; empty for body and header/footer matches. */
  noteId?: number
  range: DocxSearchRange
  before: string
  match: string
  after: string
}

export type DocxSearchStatus = "idle" | "searching" | "ready" | "error"

export interface DocxSearchState {
  status: DocxSearchStatus
  query: string
  matches: readonly DocxSearchMatch[]
  activeIndex: number
  error?: {
    code: "SEARCH_FAILED" | "ACTIVATION_FAILED"
    message: string
  }
}

interface StoryContext {
  region: DocxSearchRegion
  partName?: string
  noteId?: number
}

interface SearchableParagraph extends StoryContext {
  nodeIndex: number
  location: DocxSearchLocation
  paragraph: ParagraphNode
}

function abortError(): Error {
  if (typeof DOMException !== "undefined") return new DOMException("DOCX search aborted.", "AbortError")
  const error = new Error("DOCX search aborted.")
  error.name = "AbortError"
  return error
}

function nestingPathForSteps(path: readonly DocxNestedTableStep[]): string {
  return path
    .map((step) => `${step.rowIndex}-${step.cellIndex}-${step.contentIndex}`)
    .join("/")
}

function appendTableParagraphs(
  table: TableNode,
  topLevelTableIndex: number,
  path: readonly DocxNestedTableStep[],
  story: StoryContext,
  paragraphs: SearchableParagraph[],
): void {
  table.rows.forEach((row, rowIndex) => {
    row.cells.forEach((cell, cellIndex) => {
      // `paragraphIndex` counts paragraphs only; a nested table occupies a
      // content slot but must not shift a following paragraph's index. This
      // mirrors `cellParagraphIndex()` in DocxTableHost.
      let paragraphIndex = 0
      cell.nodes.forEach((cellNode: TableCellContentNode, contentIndex) => {
        if (cellNode.type === "paragraph") {
          const location: DocxSearchLocation = path.length > 0
            ? {
                kind: "nested-table-cell",
                tableIndex: topLevelTableIndex,
                path,
                rowIndex,
                cellIndex,
                paragraphIndex,
                nestingPath: nestingPathForSteps(path),
              }
            : {
                kind: "table-cell",
                tableIndex: topLevelTableIndex,
                rowIndex,
                cellIndex,
                paragraphIndex,
              }
          paragraphs.push({
            ...story,
            nodeIndex: topLevelTableIndex,
            location,
            paragraph: cellNode,
          })
          paragraphIndex += 1
          return
        }

        appendTableParagraphs(
          cellNode,
          topLevelTableIndex,
          [...path, { rowIndex, cellIndex, contentIndex }],
          story,
          paragraphs,
        )
      })
    })
  })
}

function appendStoryParagraphs(
  nodes: readonly DocNode[],
  story: StoryContext,
  paragraphs: SearchableParagraph[],
): void {
  nodes.forEach((node, nodeIndex) => {
    if (node.type === "paragraph") {
      paragraphs.push({
        ...story,
        nodeIndex,
        location: { kind: "paragraph", nodeIndex },
        paragraph: node,
      })
      return
    }
    appendTableParagraphs(node, nodeIndex, [], story, paragraphs)
  })
}

/**
 * Header/footer parts can be referenced by more than one section (for example
 * a shared default header). Search each part once.
 */
function appendSectionParagraphs(
  sections: readonly (HeaderSection | FooterSection)[],
  region: "header" | "footer",
  paragraphs: SearchableParagraph[],
): void {
  const seen = new Set<string>()
  sections.forEach((section, sectionIndex) => {
    const key = section.partName || `${region}:${sectionIndex}`
    if (seen.has(key)) return
    seen.add(key)
    appendStoryParagraphs(section.nodes, { region, partName: section.partName }, paragraphs)
  })
}

/**
 * Parts that are actually rendered come from `metadata.sections` when the
 * document has section metadata; the top-level lists are the editing fallback
 * for single-section documents. Prefer the rendered section lists so search
 * never reports stale text that no page shows.
 */
function renderedHeaderFooterSections(
  model: DocModel,
  region: "header" | "footer",
): (HeaderSection | FooterSection)[] {
  const sectionKey = region === "header" ? "headerSections" : "footerSections"
  const topLevel = region === "header"
    ? model.metadata.headerSections ?? []
    : model.metadata.footerSections ?? []
  const perSection = (model.metadata.sections ?? []).flatMap(
    (section) => section[sectionKey] ?? [],
  )
  return [...perSection, ...topLevel]
}

function appendNoteParagraphs(
  notes: readonly { id: number; text: string; nodes?: DocNode[] }[],
  region: "footnote" | "endnote",
  paragraphs: SearchableParagraph[],
): void {
  notes.forEach((note) => {
    if (note.nodes?.length) {
      appendStoryParagraphs(note.nodes, { region, noteId: note.id }, paragraphs)
      return
    }
    if (!note.text) return
    paragraphs.push({
      region,
      noteId: note.id,
      nodeIndex: 0,
      location: { kind: "note-text", noteId: note.id },
      // Note bodies without nodes render `note.text` as a single span.
      paragraph: {
        type: "paragraph",
        children: [{ type: "text", text: note.text }],
      },
    })
  })
}

/**
 * Paragraphs searchable by "find in document": body content, header/footer
 * stories, and footnote/endnote bodies. Tables are walked recursively so text
 * inside nested tables is searchable too.
 */
function searchableParagraphs(model: DocModel): SearchableParagraph[] {
  const paragraphs: SearchableParagraph[] = []
  appendStoryParagraphs(model.nodes, { region: "body" }, paragraphs)
  appendSectionParagraphs(renderedHeaderFooterSections(model, "header"), "header", paragraphs)
  appendSectionParagraphs(renderedHeaderFooterSections(model, "footer"), "footer", paragraphs)
  appendNoteParagraphs(model.metadata.footnotes ?? [], "footnote", paragraphs)
  appendNoteParagraphs(model.metadata.endnotes ?? [], "endnote", paragraphs)
  return paragraphs
}

/**
 * Search the DOCX model with Word-compatible defaults: literal,
 * case-insensitive matching. Options add case sensitivity, whole-word bounds
 * and regular-expression syntax. Match offsets are UTF-16 offsets in the
 * paragraph source text, matching DOM Range's offset model.
 */
export function findDocxSearchMatches(
  model: DocModel,
  query: string,
  signal?: AbortSignal,
  options: DocxSearchOptions = {},
): DocxSearchMatch[] {
  const normalizedQuery = query.trim()
  if (!normalizedQuery) return []
  const expression = createSurfaceSearchExpression(normalizedQuery, options)
  const matches: DocxSearchMatch[] = []

  for (const item of searchableParagraphs(model)) {
    if (signal?.aborted) throw abortError()
    const text = paragraphText(item.paragraph)
    expression.lastIndex = 0
    let result: RegExpExecArray | null
    while ((result = expression.exec(text))) {
      // A user regex can match the empty string. Advancing lastIndex keeps the
      // loop finite; empty matches are not useful highlights.
      if (!result[0].length) {
        expression.lastIndex += 1
        continue
      }
      const start = result.index
      const end = start + result[0].length
      const match: DocxSearchMatch = {
        kind: "docx-text",
        region: item.region,
        nodeIndex: item.nodeIndex,
        range: {
          start: { location: item.location, offset: start },
          end: { location: item.location, offset: end },
        },
        before: text.slice(Math.max(0, start - 24), start),
        match: result[0],
        after: text.slice(end, end + 24),
      }
      if (item.partName !== undefined) match.partName = item.partName
      if (item.noteId !== undefined) match.noteId = item.noteId
      matches.push(match)
    }
  }

  return matches
}
