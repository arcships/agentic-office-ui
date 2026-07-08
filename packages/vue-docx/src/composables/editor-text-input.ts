// editor-text-input.ts — text input operations (contentEditable + draft)
//
// Extracted from useDocxEditor.ts.

import type { DocModel, ParagraphNode, TextRunNode } from "@extend-ai/docx-core"
import type { DocxTextRange, ParagraphLocation } from "@extend-ai/docx-core"
import {
  cloneDocModel,
  updateParagraphText,
  normalizeTextRange,
  compareTextRangeBoundaries,
  sameParagraphLocation,
  paragraphText,
  paragraphIsEffectivelyEmpty,
  rangeCoversEntireDocument,
  firstParagraphLocationInDocument,
  normalizeRangeBoundaryParagraphOffset,
  updateParagraphTextAtLocation,
  paragraphRangeForMutate,
  fullyCoveredTableNodeIndexesForRange,
  adjustLocationAfterRemovedNodeIndexes,
  cloneParagraphStyle,
  cloneTextStyle,
  firstTextStyleAtOffset,
  firstRunStyle,
  paragraphHasOnlyTextRuns,
  cloneTextRunWithMetadata,
  splitTextRunsAtOffset,
  cloneTextRangeLocation,
  getParagraphAtLocation,
  resolveSelectedParagraphLocation,
  splitParagraphChildrenAtTextOffsets,
  splitParagraphStyleWithDefaultSpacing,
  paragraphIsList,
  isUnorderedListText,
  isOrderedListText,
  paragraphHasNumbering,
  nextOrderedListItemText,
  textWithListType,
  listPrefixLength,
  stripListPrefix,
  ORDERED_LIST_PREFIX_CAPTURE_PATTERN,
} from "@extend-ai/docx-core"
import type { DocxListType } from "@extend-ai/docx-core"
import type { EditorCore } from "./editor-shared"

export function createEditorTextInput(
  ctx: EditorCore,
  applyChange: (updater: (current: DocModel) => DocModel, successStatus?: string) => void,
) {
  // ── commitParagraphText ──────────────────────────────────────────
  const commitParagraphText = (nodeIndex: number, text: string): void => {
    applyChange((current) => updateParagraphText(current, nodeIndex, text))
  }

  // ── commitTableCellText ──────────────────────────────────────────
  const commitTableCellText = (
    tableIndex: number, rowIndex: number, cellIndex: number, text: string
  ): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") return current
      const cell = node.rows[rowIndex]?.cells[cellIndex]
      if (!cell) return current
      const para = cell.nodes[0]
      if (!para || para.type !== "paragraph") {
        const newPara: ParagraphNode = {
          type: "paragraph",
          children: [{ type: "text", text, style: {} }],
        }
        if (cell.nodes.length > 0) {
          cell.nodes[0] = newPara
        } else {
          cell.nodes.push(newPara)
        }
        return nextModel
      }
      para.children = [{ type: "text", text, style: {} }]
      return nextModel
    })
  }

  // ── commitTableCellParagraphTextRecursive ────────────────────────
  const commitTableCellParagraphTextRecursive = (
    tableIndex: number, rowIndex: number, cellIndex: number, paragraphIndex: number, text: string
  ): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") return current
      const cell = node.rows[rowIndex]?.cells[cellIndex]
      if (!cell) return current
      const para = cell.nodes[paragraphIndex]
      if (!para || para.type !== "paragraph") return current
      para.children = [{ type: "text", text, style: {} }]
      return nextModel
    })
  }

  // ── commitSectionParagraphText ───────────────────────────────────
  const commitSectionParagraphText = (
    location: { region: "header" | "footer"; partName: string; nodeIndex: number },
    text: string
  ): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const sections =
        location.region === "header"
          ? nextModel.metadata.headerSections
          : nextModel.metadata.footerSections
      const section = sections.find((s) => s.partName === location.partName)
      if (!section) return current
      const node = section.nodes[location.nodeIndex]
      if (!node || node.type !== "paragraph") return current
      node.children = [{ type: "text", text, style: {} }]
      return nextModel
    })
  }

  // ── appendParagraph ──────────────────────────────────────────────
  const appendParagraph = (text?: string): number => {
    let nodeIndex = -1
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const para: ParagraphNode = { type: "paragraph", children: [] }
      if (text) {
        para.children.push({ type: "text", text, style: {} })
      }
      nextModel.nodes.push(para)
      nodeIndex = nextModel.nodes.length - 1
      return nextModel
    })
    return nodeIndex
  }

  // ── replaceExpandedSelection ─────────────────────────────────────
  // Replaces the text within the expanded selection range with new text.
  // Handles same-paragraph, cross-paragraph, table-cell, and entire-document cases.
  const replaceExpandedSelection = (
    text: string,
    range?: DocxTextRange
  ): DocxTextRange | undefined => {
    const sourceRange = range ?? ctx.activeTextRangeSnapshot.value
    if (!sourceRange) return undefined

    const normalizedRange = normalizeTextRange({
      start: {
        location: cloneTextRangeLocation(sourceRange.start.location),
        offset: sourceRange.start.offset,
      },
      end: {
        location: cloneTextRangeLocation(sourceRange.end.location),
        offset: sourceRange.end.offset,
      },
    })

    if (compareTextRangeBoundaries(normalizedRange.start, normalizedRange.end) >= 0) {
      return undefined
    }

    const replacementText = text.replace(/\r\n?/g, "\n")
    let collapsedLocation = cloneTextRangeLocation(normalizedRange.start.location)
    let collapsedOffset = Math.max(0, Math.round(normalizedRange.start.offset))
    let replaced = false

    applyChange(
      (current) => {
        let next = cloneDocModel(current)

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
            ctx.pendingRunStyleSnapshot.value ??
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
          replaced = true
          return next
        }

        // Same paragraph case
        const startLookup = getParagraphAtLocation(next, normalizedRange.start.location)
        const endLookup = getParagraphAtLocation(next, normalizedRange.end.location)
        const startParagraph = startLookup.paragraph
        const endParagraph = endLookup.paragraph
        if (!startParagraph || !endParagraph) return current

        const safeStart = normalizeRangeBoundaryParagraphOffset(startParagraph, normalizedRange.start.offset)
        const safeEnd = normalizeRangeBoundaryParagraphOffset(endParagraph, normalizedRange.end.offset)
        collapsedOffset = safeStart + replacementText.length
        const endParagraphStyleForMerge = cloneParagraphStyle(endParagraph.style)

        const insertedStyle = cloneTextStyle(
          ctx.pendingRunStyleSnapshot.value ??
            firstTextStyleAtOffset(startParagraph, safeStart, true) ??
            firstRunStyle(startParagraph) ??
            firstTextStyleAtOffset(endParagraph, safeEnd, false) ??
            firstRunStyle(endParagraph)
        )

        if (sameParagraphLocation(normalizedRange.start.location, normalizedRange.end.location)) {
          const existingText = paragraphText(startParagraph)
          const nextText = `${existingText.slice(0, safeStart)}${replacementText}${existingText.slice(safeEnd)}`
          if (nextText === existingText) return current
          next = updateParagraphTextAtLocation(next, normalizedRange.start.location, nextText, { insertedStyle })
          clearNumberingIfParagraphEmpty(normalizedRange.start.location)
          replaced = true
          return next
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

        replaced = true
        return next
      },
      replacementText.length > 0 ? "Replaced selection" : "Deleted selection"
    )

    if (!replaced) return undefined

    const collapsedRange: DocxTextRange = {
      start: { location: cloneTextRangeLocation(collapsedLocation), offset: collapsedOffset },
      end: { location: cloneTextRangeLocation(collapsedLocation), offset: collapsedOffset },
    }

    ctx.suppressSelectionReset.value = true
    if (collapsedLocation.kind === "paragraph") {
      ctx.selection.value = { kind: "paragraph", nodeIndex: collapsedLocation.nodeIndex }
    } else {
      ctx.selection.value = {
        kind: "table-cell",
        tableIndex: collapsedLocation.tableIndex,
        rowIndex: collapsedLocation.rowIndex,
        cellIndex: collapsedLocation.cellIndex,
      }
    }
    ctx.activeTextRange.value = collapsedRange
    ctx.pendingRunStyle.value = undefined

    return collapsedRange
  }

  // ── deleteExpandedSelection ──────────────────────────────────────
  const deleteExpandedSelection = (
    range?: DocxTextRange
  ): DocxTextRange | undefined => {
    return replaceExpandedSelection("", range)
  }

  // ── splitParagraphAtSelection ─────────────────────────────────────
  const splitParagraphAtSelection = (
    draftText: string,
    startOffset: number,
    endOffset?: number,
    targetLocationOverride?: ParagraphLocation
  ): { paragraphIndex: number; caretOffset: number } | undefined => {
    const targetLocation =
      targetLocationOverride ??
      resolveSelectedParagraphLocation(ctx.selection.value, ctx.activeTextRange.value)
    if (targetLocation.kind !== "paragraph") return undefined

    const normalizedDraftText = draftText ?? ""
    const normalizedStartOffset = Math.max(0, Math.round(startOffset))
    const normalizedEndOffset = Math.max(normalizedStartOffset, Math.round(endOffset ?? normalizedStartOffset))
    const insertionIndex = targetLocation.nodeIndex + 1

    let splitResult: { paragraphIndex: number; caretOffset: number } | undefined

    applyChange((current) => {
      const next = cloneDocModel(current)
      const paragraphNode = next.nodes[targetLocation.nodeIndex]
      if (!paragraphNode || paragraphNode.type !== "paragraph") return current

      const safeStart = Math.max(0, Math.min(normalizedStartOffset, normalizedDraftText.length))
      const safeEnd = Math.max(safeStart, Math.min(normalizedEndOffset, normalizedDraftText.length))
      const splitParagraphStyle = splitParagraphStyleWithDefaultSpacing(paragraphNode.style, paragraphNode.sourceXml)
      const beforeInsertedStyle = cloneTextStyle(
        ctx.pendingRunStyleSnapshot.value ??
          firstTextStyleAtOffset(paragraphNode, safeStart, true) ??
          firstRunStyle(paragraphNode)
      )
      const inheritedRunStyle = cloneTextStyle(
        ctx.pendingRunStyleSnapshot.value ??
          firstTextStyleAtOffset(paragraphNode, safeEnd, false) ??
          firstRunStyle(paragraphNode)
      ) ?? {}

      const splitChildren = splitParagraphChildrenAtTextOffsets(
        paragraphNode,
        normalizedDraftText,
        safeStart,
        safeEnd,
        { beforeInsertedStyle, afterInsertedStyle: inheritedRunStyle }
      )

      paragraphNode.style = cloneParagraphStyle(splitParagraphStyle)
      paragraphNode.children = splitChildren.beforeChildren
      paragraphNode.sourceXml = undefined

      next.nodes.splice(insertionIndex, 0, {
        type: "paragraph",
        style: cloneParagraphStyle(splitParagraphStyle),
        children: splitChildren.afterChildren,
      })

      splitResult = { paragraphIndex: insertionIndex, caretOffset: 0 }
      return next
    }, "Split paragraph")

    if (!splitResult) return undefined

    ctx.suppressSelectionReset.value = true
    ctx.selection.value = { kind: "paragraph", nodeIndex: splitResult.paragraphIndex }
    ctx.activeTextRange.value = {
      start: { location: { kind: "paragraph", nodeIndex: splitResult.paragraphIndex }, offset: splitResult.caretOffset },
      end: { location: { kind: "paragraph", nodeIndex: splitResult.paragraphIndex }, offset: splitResult.caretOffset },
    }
    return splitResult
  }

  // ── insertListItemAfterSelection ──────────────────────────────────
  const insertListItemAfterSelection = (
    draftText: string,
    startOffset: number,
    endOffset?: number,
    targetLocationOverride?: ParagraphLocation
  ): { paragraphIndex: number; caretOffset: number } | undefined => {
    const targetLocation =
      targetLocationOverride ??
      resolveSelectedParagraphLocation(ctx.selection.value, ctx.activeTextRange.value)
    if (targetLocation.kind !== "paragraph") return undefined

    const insertionIndex = targetLocation.nodeIndex + 1
    let insertedParagraphResult: { paragraphIndex: number; caretOffset: number } | undefined
    const normalizedDraftText = draftText ?? ""
    const normalizedStartOffset = Math.max(0, Math.round(startOffset))
    const normalizedEndOffset = Math.max(normalizedStartOffset, Math.round(endOffset ?? normalizedStartOffset))

    applyChange((current) => {
      let next = cloneDocModel(current)
      let paragraphNode = next.nodes[targetLocation.nodeIndex]
      if (!paragraphNode || paragraphNode.type !== "paragraph") return current

      if (!paragraphIsList(paragraphNode, normalizedDraftText)) return current

      const safeStart = Math.max(0, Math.min(normalizedStartOffset, normalizedDraftText.length))
      const safeEnd = Math.max(safeStart, Math.min(normalizedEndOffset, normalizedDraftText.length))
      const beforeText = normalizedDraftText.slice(0, safeStart)
      const afterText = normalizedDraftText.slice(safeEnd)
      const insertRunStyle = cloneTextStyle(
        ctx.pendingRunStyleSnapshot.value ??
          firstTextStyleAtOffset(paragraphNode, safeStart, true) ??
          firstRunStyle(paragraphNode)
      ) ?? {}

      const paragraphListType: DocxListType | undefined = isUnorderedListText(normalizedDraftText)
        ? "unordered"
        : isOrderedListText(normalizedDraftText)
        ? "ordered"
        : undefined
      const hasDocxNumberingVal = paragraphHasNumbering(paragraphNode)
      const textForInsertedParagraph =
        hasDocxNumberingVal || !paragraphListType
          ? afterText
          : paragraphListType === "ordered"
          ? nextOrderedListItemText(normalizedDraftText, afterText)
          : textWithListType(afterText, paragraphListType)

      const insertedCaretOffset = hasDocxNumberingVal ? 0 : listPrefixLength(textForInsertedParagraph)

      const textForCurrentParagraph =
        hasDocxNumberingVal || !paragraphListType
          ? beforeText
          : paragraphListType === "ordered"
          ? (() => {
              const currentMatch = normalizedDraftText.match(ORDERED_LIST_PREFIX_CAPTURE_PATTERN)
              if (!currentMatch) return beforeText
              return `${currentMatch[1] ?? ""}${currentMatch[2] ?? "1"}. ${stripListPrefix(beforeText)}`
            })()
          : textWithListType(beforeText, paragraphListType)

      next = updateParagraphText(next, targetLocation.nodeIndex, textForCurrentParagraph, {
        insertedStyle: cloneTextStyle(ctx.pendingRunStyleSnapshot.value),
      })
      paragraphNode = next.nodes[targetLocation.nodeIndex]
      if (!paragraphNode || paragraphNode.type !== "paragraph") return current
      paragraphNode.sourceXml = undefined

      next.nodes.splice(insertionIndex, 0, {
        type: "paragraph",
        style: cloneParagraphStyle(paragraphNode.style),
        children: [{ type: "text", text: textForInsertedParagraph, style: insertRunStyle }],
      })
      insertedParagraphResult = { paragraphIndex: insertionIndex, caretOffset: insertedCaretOffset }
      return next
    }, "Inserted list item")

    if (insertedParagraphResult === undefined) return undefined

    ctx.suppressSelectionReset.value = true
    ctx.selection.value = { kind: "paragraph", nodeIndex: insertedParagraphResult.paragraphIndex }
    ctx.activeTextRange.value = {
      start: { location: { kind: "paragraph", nodeIndex: insertedParagraphResult.paragraphIndex }, offset: insertedParagraphResult.caretOffset },
      end: { location: { kind: "paragraph", nodeIndex: insertedParagraphResult.paragraphIndex }, offset: insertedParagraphResult.caretOffset },
    }
    return insertedParagraphResult
  }

  return {
    commitParagraphText,
    commitTableCellText,
    commitTableCellParagraphTextRecursive,
    commitSectionParagraphText,
    appendParagraph,
    splitParagraphAtSelection,
    insertListItemAfterSelection,
    replaceExpandedSelection,
    deleteExpandedSelection,
  }
}
