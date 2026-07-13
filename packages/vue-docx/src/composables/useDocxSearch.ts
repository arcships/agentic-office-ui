import {
  paragraphText,
  type DocModel,
  type DocxTextRange,
  type DocxTextRangeLocation,
  type ParagraphNode,
} from "@arcships/docx-core"

export interface DocxSearchMatch {
  kind: "docx-text"
  nodeIndex: number
  range: DocxTextRange
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

interface SearchableParagraph {
  nodeIndex: number
  location: DocxTextRangeLocation
  paragraph: ParagraphNode
}

function abortError(): Error {
  if (typeof DOMException !== "undefined") return new DOMException("DOCX search aborted.", "AbortError")
  const error = new Error("DOCX search aborted.")
  error.name = "AbortError"
  return error
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^\${}()|[\]\\]/g, "\\$&")
}

function searchableParagraphs(model: DocModel): SearchableParagraph[] {
  const paragraphs: SearchableParagraph[] = []
  model.nodes.forEach((node, nodeIndex) => {
    if (node.type === "paragraph") {
      paragraphs.push({
        nodeIndex,
        location: { kind: "paragraph", nodeIndex },
        paragraph: node,
      })
      return
    }

    node.rows.forEach((row, rowIndex) => {
      row.cells.forEach((cell, cellIndex) => {
        cell.nodes.forEach((cellNode, paragraphIndex) => {
          // DocxTextRange currently addresses direct cell paragraphs only.
          if (cellNode.type !== "paragraph") return
          paragraphs.push({
            nodeIndex,
            location: {
              kind: "table-cell",
              tableIndex: nodeIndex,
              rowIndex,
              cellIndex,
              paragraphIndex,
            },
            paragraph: cellNode,
          })
        })
      })
    })
  })
  return paragraphs
}

/**
 * Literal, case-insensitive DOCX search. Match offsets are UTF-16 offsets in
 * the paragraph source text, matching DOM Range's offset model.
 */
export function findDocxSearchMatches(
  model: DocModel,
  query: string,
  signal?: AbortSignal,
): DocxSearchMatch[] {
  const normalizedQuery = query.trim()
  if (!normalizedQuery) return []
  const expression = new RegExp(escapeRegExp(normalizedQuery), "giu")
  const matches: DocxSearchMatch[] = []

  for (const item of searchableParagraphs(model)) {
    if (signal?.aborted) throw abortError()
    const text = paragraphText(item.paragraph)
    expression.lastIndex = 0
    let result: RegExpExecArray | null
    while ((result = expression.exec(text))) {
      const start = result.index
      const end = start + result[0].length
      matches.push({
        kind: "docx-text",
        nodeIndex: item.nodeIndex,
        range: {
          start: { location: item.location, offset: start },
          end: { location: item.location, offset: end },
        },
        before: text.slice(Math.max(0, start - 24), start),
        match: result[0],
        after: text.slice(end, end + 24),
      })
      if (!result[0].length) expression.lastIndex += 1
    }
  }

  return matches
}
