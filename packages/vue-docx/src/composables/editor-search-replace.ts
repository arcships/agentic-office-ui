// editor-search-replace.ts — replace commands over search results.
//
// `replaceText` in run-style-ops rewrites the whole document with a
// per-run regex, which cannot address "the match the user is looking at" and
// loses run boundaries. Find & replace therefore works from search results:
// every result carries its paragraph location plus UTF-16 offsets, so one
// command edits exactly those offsets, and the whole batch is one model
// revision (one undo entry).

import type {
  DocModel,
  DocxEditorController,
  DocxReplaceAllOptions,
  DocxReplaceSummary,
  DocxReplaceTarget,
  DocxParagraphTextReplacement,
} from "@arcships/docx-core"
import {
  replaceTextInDocxParagraphs,
  replaceTextInParagraphLocation,
} from "@arcships/docx-core"
import type { ApplyModelChangeOptions, EditorCore } from "./editor-shared"
import {
  findDocxSearchMatches,
  type DocxSearchMatch,
  type DocxSearchOptions,
} from "./useDocxSearch"

type ApplyChange = (
  updater: (current: DocModel) => DocModel,
  successStatus?: string,
  options?: ApplyModelChangeOptions,
) => boolean

/**
 * Convert one search result into a replace target.
 *
 * Returns `undefined` for results the editor cannot rewrite: matches in
 * headers, footers and note bodies belong to stories that render read-only, so
 * replace keeps them intact instead of rewriting text the user cannot edit.
 */
export function docxReplaceTargetFromMatch(
  match: DocxSearchMatch,
): DocxReplaceTarget | undefined {
  if (match.region !== "body") return undefined
  const location = match.range.start.location
  if (location.kind === "note-text") return undefined
  if (location.kind === "paragraph") {
    return {
      region: "body",
      location: { kind: "paragraph", nodeIndex: location.nodeIndex },
      start: match.range.start.offset,
      end: match.range.end.offset,
    }
  }
  if (location.kind === "table-cell") {
    return {
      region: "body",
      location: {
        kind: "table-cell",
        tableIndex: location.tableIndex,
        rowIndex: location.rowIndex,
        cellIndex: location.cellIndex,
        paragraphIndex: location.paragraphIndex,
      },
      start: match.range.start.offset,
      end: match.range.end.offset,
    }
  }
  return {
    region: "body",
    location: {
      kind: "nested-table-cell",
      tableIndex: location.tableIndex,
      path: location.path.map((step) => ({
        rowIndex: step.rowIndex,
        cellIndex: step.cellIndex,
        contentIndex: step.contentIndex,
      })),
      rowIndex: location.rowIndex,
      cellIndex: location.cellIndex,
      paragraphIndex: location.paragraphIndex,
    },
    start: match.range.start.offset,
    end: match.range.end.offset,
  }
}

/** True when a search result can be rewritten by this editor. */
export function canReplaceDocxSearchMatch(match: DocxSearchMatch): boolean {
  return docxReplaceTargetFromMatch(match) !== undefined
}

export function createEditorSearchReplace(ctx: EditorCore, applyChange: ApplyChange) {
  const refuseUnsupportedTarget = (): void => {
    ctx.status.value = ctx.statusMessage("replaceTargetUnavailable")
  }

  const replaceMatch = (target: DocxReplaceTarget, text: string): boolean => {
    if (target.region !== "body") {
      refuseUnsupportedTarget()
      return false
    }
    let changed = false
    applyChange((current) => {
      const result = replaceTextInParagraphLocation(current, {
        location: target.location,
        start: target.start,
        end: target.end,
        replacement: text,
      })
      if (!result) return current
      changed = true
      return result.model
    })
    if (changed) {
      ctx.status.value = ctx.statusMessage("replacedMatches", { count: 1 })
    }
    return changed
  }

  const replaceAll = (
    query: string,
    text: string,
    options: DocxReplaceAllOptions = {},
  ): DocxReplaceSummary => {
    let replaced = 0
    let skipped = 0
    const changed = applyChange((current) => {
      const matches = findDocxSearchMatches(
        current,
        query,
        undefined,
        options as DocxSearchOptions,
      )
      const edits: DocxParagraphTextReplacement[] = []
      replaced = 0
      skipped = 0
      for (const match of matches) {
        const target = docxReplaceTargetFromMatch(match)
        if (!target) {
          skipped += 1
          continue
        }
        edits.push({
          location: target.location,
          start: target.start,
          end: target.end,
          replacement: text,
        })
      }
      if (edits.length === 0) return current
      const result = replaceTextInDocxParagraphs(current, edits)
      replaced = result.applied
      return result.applied > 0 ? result.model : current
    })

    if (changed) {
      ctx.status.value = ctx.statusMessage("replacedMatches", { count: replaced })
    } else if (skipped > 0) {
      refuseUnsupportedTarget()
    }
    return { replaced, skipped }
  }

  return { replaceMatch, replaceAll }
}

/** Subset of `DocxDocumentSurface`'s exposed search API the find bar needs. */
export interface DocxFindReplaceSurface {
  search(query: string, options?: DocxSearchOptions): Promise<import("./useDocxSearch").DocxSearchState>
  refreshSearch(): Promise<import("./useDocxSearch").DocxSearchState>
  searchNext(): Promise<void>
  searchPrevious(): Promise<void>
  clearSearch(): void
  getSearchState(): import("./useDocxSearch").DocxSearchState
}

/** Convenience type for hosts that keep the controller around. */
export type DocxReplaceController = Pick<DocxEditorController, "replaceMatch" | "replaceAll">
