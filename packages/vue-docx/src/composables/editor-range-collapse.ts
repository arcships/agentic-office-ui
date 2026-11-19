// editor-range-collapse.ts — collapse a text range to its replacement caret.
//
// Extracted from `replaceExpandedSelection` so a rich clipboard paste can
// collapse an expanded selection and insert styled runs inside one transaction
// (one undo step) instead of replacing the selection with plain text first.
//
// The function is pure: it clones the touched nodes, applies the same
// whole-document / same-paragraph / cross-paragraph / table-cell handling the
// delete path always had, and reports the collapsed caret.

import type {
  DocModel,
  DocxTextRange,
  ParagraphLocation,
  TextRunNode,
  TextStyle,
} from "@arcships/docx-core"
import {
  adjustLocationAfterRemovedNodeIndexes,
  cloneDocModelNodes,
  cloneParagraphStyle,
  cloneTextRangeLocation,
  cloneTextRunWithMetadata,
  cloneTextStyle,
  firstParagraphLocationInDocument,
  firstRunStyle,
  firstTextStyleAtOffset,
  fullyCoveredTableNodeIndexesForRange,
  getParagraphAtLocation,
  mergeParagraphCommentRanges,
  normalizeRangeBoundaryParagraphOffset,
  paragraphHasOnlyTextRuns,
  paragraphIsEffectivelyEmpty,
  paragraphRangeForMutate,
  paragraphText,
  rangeCoversEntireDocument,
  sameParagraphLocation,
  shiftParagraphCommentRanges,
  splitParagraphCommentRanges,
  splitTextRunsAtOffset,
  updateParagraphTextAtLocation,
} from "@arcships/docx-core"
import { topLevelNodeIndex } from "./editor-clone-helpers"

export interface CollapsedRangeResult {
  model: DocModel
  location: ParagraphLocation
  offset: number
}

/**
 * Deletes the text inside `normalizedRange` and inserts `replacementText` at
 * the collapse point.
 *
 * `pendingRunStyle` (Ctrl+B before typing) overrides the inherited run style of
 * the inserted text, exactly like the previous inline implementation.
 */
export function collapseSelectionRange(
  current: DocModel,
  normalizedRange: DocxTextRange,
  replacementText: string,
  pendingRunStyle?: TextStyle,
): CollapsedRangeResult | undefined {
  let collapsedLocation = cloneTextRangeLocation(normalizedRange.start.location)
  let collapsedOffset = Math.max(0, Math.round(normalizedRange.start.offset))

  let next = cloneDocModelNodes(current, [
    topLevelNodeIndex(normalizedRange.start.location),
    topLevelNodeIndex(normalizedRange.end.location),
  ])

  const clearNumberingIfParagraphEmpty = (location: ParagraphLocation): void => {
    const { paragraph, tableNode } = getParagraphAtLocation(next, location)
    if (!paragraph || !paragraphIsEffectivelyEmpty(paragraph) || !paragraph.style?.numbering) {
      return
    }
    paragraph.style = { ...(cloneParagraphStyle(paragraph.style) ?? {}), numbering: undefined }
    paragraph.sourceXml = undefined
    if (tableNode) {
      tableNode.sourceXml = undefined
    }
  }

  // Entire document replacement — keep a single paragraph with the new text
  if (rangeCoversEntireDocument(current, normalizedRange)) {
    const firstLocation = firstParagraphLocationInDocument(next)
    const firstPara = firstLocation
      ? getParagraphAtLocation(next, firstLocation).paragraph
      : undefined
    const replacementRunStyle = cloneTextStyle(
      pendingRunStyle ??
        (firstPara ? firstTextStyleAtOffset(firstPara, 0, true) ?? firstRunStyle(firstPara) : undefined)
    )
    next.nodes = [
      {
        type: "paragraph",
        children: [
          {
            type: "text",
            text: replacementText,
            ...(replacementRunStyle ? { style: replacementRunStyle } : undefined),
          },
        ],
      },
    ]
    collapsedLocation = { kind: "paragraph", nodeIndex: 0 }
    collapsedOffset = replacementText.length
    return { model: next, location: collapsedLocation, offset: collapsedOffset }
  }

  // Same paragraph case
  const startLookup = getParagraphAtLocation(next, normalizedRange.start.location)
  const endLookup = getParagraphAtLocation(next, normalizedRange.end.location)
  const startParagraph = startLookup.paragraph
  const endParagraph = endLookup.paragraph
  if (!startParagraph || !endParagraph) return undefined

  const safeStart = normalizeRangeBoundaryParagraphOffset(startParagraph, normalizedRange.start.offset)
  const safeEnd = normalizeRangeBoundaryParagraphOffset(endParagraph, normalizedRange.end.offset)
  collapsedOffset = safeStart + replacementText.length
  const endParagraphStyleForMerge = cloneParagraphStyle(endParagraph.style)

  const insertedStyle = cloneTextStyle(
    pendingRunStyle ??
      firstTextStyleAtOffset(startParagraph, safeStart, true) ??
      firstRunStyle(startParagraph) ??
      firstTextStyleAtOffset(endParagraph, safeEnd, false) ??
      firstRunStyle(endParagraph)
  )

  if (sameParagraphLocation(normalizedRange.start.location, normalizedRange.end.location)) {
    const existingText = paragraphText(startParagraph)
    const nextText = `${existingText.slice(0, safeStart)}${replacementText}${existingText.slice(safeEnd)}`
    if (nextText === existingText) return undefined
    next = updateParagraphTextAtLocation(next, normalizedRange.start.location, nextText, { insertedStyle })
    clearNumberingIfParagraphEmpty(normalizedRange.start.location)
    return { model: next, location: collapsedLocation, offset: collapsedOffset }
  }

  // Cross-paragraph case
  const startText = paragraphText(startParagraph)
  const endText = paragraphText(endParagraph)
  const carriesStartPrefix = safeStart > 0
  const carriesEndSuffix = safeEnd < endText.length
  const shouldAdoptEndParagraphStyleOnDelete =
    replacementText.length === 0 && !carriesStartPrefix && carriesEndSuffix
  const canPreserveCrossParagraphRunStyles =
    paragraphHasOnlyTextRuns(startParagraph) && paragraphHasOnlyTextRuns(endParagraph)

  if (canPreserveCrossParagraphRunStyles) {
    const startRuns = startParagraph.children
      .filter((child): child is TextRunNode => child.type === "text")
      .map(cloneTextRunWithMetadata)
    const endRuns = endParagraph.children
      .filter((child): child is TextRunNode => child.type === "text")
      .map(cloneTextRunWithMetadata)
    const splitStartRuns = splitTextRunsAtOffset(startRuns, safeStart)
    const splitEndRuns = splitTextRunsAtOffset(endRuns, safeEnd)
    const mergedRuns: TextRunNode[] = [...splitStartRuns.left]
    if (replacementText.length > 0) {
      const previousRun = mergedRuns[mergedRuns.length - 1]
      const nextRun = splitEndRuns.right[0]
      const inferredLink =
        previousRun?.link && nextRun?.link && previousRun.link === nextRun.link
          ? previousRun.link
          : previousRun?.link ?? nextRun?.link
      mergedRuns.push({
        type: "text",
        text: replacementText,
        style: cloneTextStyle(insertedStyle ?? previousRun?.style ?? nextRun?.style),
        link: inferredLink,
      })
    }
    mergedRuns.push(...splitEndRuns.right)
    startParagraph.children = mergedRuns.length > 0
      ? mergedRuns
      : [{ type: "text", text: "", style: cloneTextStyle(insertedStyle) }]
    startParagraph.sourceXml = undefined
    if (startLookup.tableNode) startLookup.tableNode.sourceXml = undefined
  } else {
    const mergedText = `${startText.slice(0, safeStart)}${replacementText}${endText.slice(safeEnd)}`
    next = updateParagraphTextAtLocation(next, normalizedRange.start.location, mergedText, { insertedStyle })
  }

  // Comment ranges span character offsets, so a join must take the
  // surviving left prefix and the shifted right suffix. Fragments that
  // become adjacent once the deleted text is gone are coalesced again.
  const mergedCommentRanges = mergeParagraphCommentRanges(
    [
      ...splitParagraphCommentRanges(startParagraph.commentRanges, safeStart).before,
      ...shiftParagraphCommentRanges(
        splitParagraphCommentRanges(endParagraph.commentRanges, safeEnd).after,
        safeStart + replacementText.length
      ),
    ],
    `${startText.slice(0, safeStart)}${replacementText}${endText.slice(safeEnd)}`
  )
  const mergedParagraphAfterReplace = getParagraphAtLocation(
    next,
    normalizedRange.start.location
  ).paragraph
  if (mergedParagraphAfterReplace) {
    mergedParagraphAfterReplace.commentRanges =
      mergedCommentRanges.length > 0 ? mergedCommentRanges : undefined
  }

  if (shouldAdoptEndParagraphStyleOnDelete) {
    const mergedLookup = getParagraphAtLocation(next, normalizedRange.start.location)
    if (mergedLookup.paragraph) {
      mergedLookup.paragraph.style = endParagraphStyleForMerge
        ? cloneParagraphStyle(endParagraphStyleForMerge)
        : undefined
      mergedLookup.paragraph.sourceXml = undefined
      if (mergedLookup.tableNode) mergedLookup.tableNode.sourceXml = undefined
    }
  }
  clearNumberingIfParagraphEmpty(normalizedRange.start.location)

  // Compact removed paragraphs between start and end
  let compactedRemovedParagraphs = false
  if (
    normalizedRange.start.location.kind === "paragraph" &&
    normalizedRange.end.location.kind === "paragraph" &&
    normalizedRange.end.location.nodeIndex > normalizedRange.start.location.nodeIndex
  ) {
    const removeStartIndex = normalizedRange.start.location.nodeIndex + 1
    const removeEndIndex = normalizedRange.end.location.nodeIndex
    const removeCount = removeEndIndex - removeStartIndex + 1
    if (removeCount > 0) {
      next.nodes.splice(removeStartIndex, removeCount)
      compactedRemovedParagraphs = true
    }
  } else if (
    normalizedRange.start.location.kind === "table-cell" &&
    normalizedRange.end.location.kind === "table-cell" &&
    normalizedRange.start.location.tableIndex === normalizedRange.end.location.tableIndex &&
    normalizedRange.start.location.rowIndex === normalizedRange.end.location.rowIndex &&
    normalizedRange.start.location.cellIndex === normalizedRange.end.location.cellIndex &&
    normalizedRange.end.location.paragraphIndex > normalizedRange.start.location.paragraphIndex
  ) {
    const tableNode = next.nodes[normalizedRange.start.location.tableIndex]
    if (tableNode && tableNode.type === "table") {
      const cell = tableNode.rows[normalizedRange.start.location.rowIndex]?.cells[normalizedRange.start.location.cellIndex]
      if (cell) {
        const paragraphNodeIndexes = cell.nodes.reduce<number[]>((indexes, node, idx) => {
          if (node.type === "paragraph") indexes.push(idx)
          return indexes
        }, [])
        for (let pi = normalizedRange.end.location.paragraphIndex; pi > normalizedRange.start.location.paragraphIndex; pi -= 1) {
          const nodeIndex = paragraphNodeIndexes[pi]
          if (nodeIndex !== undefined) {
            cell.nodes.splice(nodeIndex, 1)
            compactedRemovedParagraphs = true
          }
        }
        if (compactedRemovedParagraphs) tableNode.sourceXml = undefined
      }
    }
  }

  // Remove fully covered tables
  const fullyCoveredTableIndexes =
    replacementText.length === 0
      ? fullyCoveredTableNodeIndexesForRange(next, normalizedRange)
      : []
  if (fullyCoveredTableIndexes.length > 0) {
    const uniqueDescendingIndexes = Array.from(new Set(fullyCoveredTableIndexes)).sort((a, b) => b - a)
    let removedAnyTable = false
    uniqueDescendingIndexes.forEach((tableIndex) => {
      if (next.nodes[tableIndex]?.type !== "table") return
      next.nodes.splice(tableIndex, 1)
      removedAnyTable = true
    })
    if (removedAnyTable) {
      compactedRemovedParagraphs = true
      const adjustedLocation = adjustLocationAfterRemovedNodeIndexes(
        collapsedLocation,
        [...uniqueDescendingIndexes].sort((a, b) => a - b)
      )
      if (adjustedLocation) {
        collapsedLocation = adjustedLocation
      } else {
        const firstLocation = firstParagraphLocationInDocument(next)
        if (firstLocation) {
          collapsedLocation = cloneTextRangeLocation(firstLocation)
        } else {
          next.nodes = [{ type: "paragraph", children: [{ type: "text", text: "" }] }]
          collapsedLocation = { kind: "paragraph", nodeIndex: 0 }
        }
        collapsedOffset = 0
      }
    }
  }

  // Clear remaining intermediate paragraphs if not already compacted
  if (!compactedRemovedParagraphs) {
    const rangeParagraphs = paragraphRangeForMutate(next, normalizedRange.start.location, normalizedRange.end.location)
    for (const { location } of rangeParagraphs) {
      if (sameParagraphLocation(location, normalizedRange.start.location)) continue
      next = updateParagraphTextAtLocation(next, location, "")
      clearNumberingIfParagraphEmpty(location)
    }
  }

  // Ensure collapsed paragraph exists
  const collapsedParagraph = getParagraphAtLocation(next, collapsedLocation).paragraph
  if (!collapsedParagraph) {
    const firstLocation = firstParagraphLocationInDocument(next)
    if (firstLocation) {
      collapsedLocation = cloneTextRangeLocation(firstLocation)
      collapsedOffset = 0
    } else {
      next.nodes = [{ type: "paragraph", children: [{ type: "text", text: "" }] }]
      collapsedLocation = { kind: "paragraph", nodeIndex: 0 }
      collapsedOffset = 0
    }
  } else {
    collapsedOffset = normalizeRangeBoundaryParagraphOffset(collapsedParagraph, collapsedOffset)
  }

  return { model: next, location: collapsedLocation, offset: collapsedOffset }
}
