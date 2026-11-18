// Model-backed resolvers for DOCX field results that depend on other content:
// STYLEREF (nearest heading of a style), REF/PAGEREF (bookmark text/page), and
// the page index map used by PAGEREF.
//
// Kept in vue-docx because the document model is already the renderer's input;
// docx-core field parsing stays platform-neutral.

import {
  createDocxBookmarkIndex,
  paragraphText,
  type DocModel,
  type ParagraphNode,
} from "@arcships/docx-core"

export interface DocxFieldResolvers {
  resolveStyleRefFieldValue: (
    target: string,
    beforeNodeIndex?: number,
  ) => string | undefined
  resolveBookmarkFieldValue: (target: string) => string | undefined
  resolveBookmarkPageNumber: (target: string) => number | undefined
}

function normalizeStyleRefTarget(value: string): string {
  return value.trim().toLowerCase().replace(/[\s_-]+/g, "")
}

function paragraphMatchesStyleRefTarget(
  paragraph: ParagraphNode,
  normalizedTarget: string,
): boolean {
  const candidates = [paragraph.style?.styleId, paragraph.style?.styleName]
  return candidates.some(
    (candidate) =>
      Boolean(candidate) &&
      normalizeStyleRefTarget(candidate as string) === normalizedTarget,
  )
}

export function createDocxFieldResolvers(
  model: DocModel,
  options?: {
    pageNumberByNodeIndex?: ReadonlyMap<number, number>
  },
): DocxFieldResolvers {
  const bookmarkIndex = createDocxBookmarkIndex(model)

  function resolveStyleRefFieldValue(
    target: string,
    beforeNodeIndex?: number,
  ): string | undefined {
    const normalizedTarget = normalizeStyleRefTarget(target)
    if (!normalizedTarget) return undefined

    const before = Number.isFinite(beforeNodeIndex)
      ? Math.max(0, Math.min(model.nodes.length, Math.round(beforeNodeIndex as number)))
      : model.nodes.length

    // Word resolves STYLEREF against the closest preceding paragraph with the
    // requested style. For a first-page header with no preceding heading, fall
    // forward to the first matching paragraph so the field is still useful.
    for (let nodeIndex = before - 1; nodeIndex >= 0; nodeIndex -= 1) {
      const node = model.nodes[nodeIndex]
      if (
        node?.type === "paragraph" &&
        paragraphMatchesStyleRefTarget(node, normalizedTarget)
      ) {
        return paragraphText(node).trim()
      }
    }
    for (let nodeIndex = 0; nodeIndex < model.nodes.length; nodeIndex += 1) {
      const node = model.nodes[nodeIndex]
      if (
        node?.type === "paragraph" &&
        paragraphMatchesStyleRefTarget(node, normalizedTarget)
      ) {
        return paragraphText(node).trim()
      }
    }
    return undefined
  }

  function resolveBookmarkFieldValue(target: string): string | undefined {
    return bookmarkIndex.get(target)?.text
  }

  function resolveBookmarkPageNumber(target: string): number | undefined {
    const bookmark = bookmarkIndex.get(target)
    if (!bookmark) return undefined
    return options?.pageNumberByNodeIndex?.get(bookmark.nodeIndex)
  }

  return {
    resolveStyleRefFieldValue,
    resolveBookmarkFieldValue,
    resolveBookmarkPageNumber,
  }
}
