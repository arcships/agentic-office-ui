// editor-text-input.ts — text input operations (contentEditable + draft)
//
// Extracted from useDocxEditor.ts.

import type { DocModel, ParagraphNode, TableNode, TextRunNode } from "@arcships/docx-core"
import type { DocxTextRange, ParagraphLocation } from "@arcships/docx-core"
import {
  cloneDocModelNodes,
  cloneDocModelShallow,
  parseEmbeddedTableRuntimeKey,
  updateParagraphText,
  normalizeTextRange,
  compareTextRangeBoundaries,
  sameParagraphLocation,
  paragraphText,
  distributeTextAcrossParagraphChildren,
  paragraphIsEffectivelyEmpty,
  rangeCoversEntireDocument,
  firstParagraphLocationInDocument,
  normalizeRangeBoundaryParagraphOffset,
  updateParagraphTextAtLocation,
  updateTableCellText,
  updateTableCellParagraphText,
  updateSectionParagraphTextAtLocation,
  paragraphRangeForMutate,
  fullyCoveredTableNodeIndexesForRange,
  adjustLocationAfterRemovedNodeIndexes,
  cloneParagraphStyle,
  cloneTextStyle,
  firstTextStyleAtOffset,
  firstRunStyle,
  paragraphHasOnlyTextRuns,
  remapParagraphCommentRanges,
  splitParagraphCommentRanges,
  shiftParagraphCommentRanges,
  mergeParagraphCommentRanges,
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
} from "@arcships/docx-core"
import type { DocxListType } from "@arcships/docx-core"
import type { ApplyModelChangeOptions, EditorCore } from "./editor-shared"
import { cloneHostTableForRuntimeKey, topLevelNodeIndex } from "./editor-clone-helpers"

function commonPrefixLength(left: string, right: string): number {
  const limit = Math.min(left.length, right.length)
  let index = 0
  while (index < limit && left[index] === right[index]) index += 1
  return index
}

export function createEditorTextInput(
  ctx: EditorCore,
  applyChange: (
    updater: (current: DocModel) => DocModel,
    successStatus?: string,
    options?: ApplyModelChangeOptions,
  ) => void,
) {
  // Resolve the style of text inserted by a commit: inherit the run style at
  // the insertion point and overlay a pending caret style (Ctrl+B before
  // typing), instead of dropping run formatting or discarding the pending run.
  const insertedStyleForTextChange = (
    paragraph: ParagraphNode, previousText: string, nextText: string
  ): TextRunNode["style"] | undefined => {
    const pending = cloneTextStyle(ctx.pendingRunStyleSnapshot.value)
    if (!pending || Object.keys(pending).length === 0) return undefined
    const baseStyle =
      firstTextStyleAtOffset(paragraph, commonPrefixLength(previousText, nextText), true) ??
      firstRunStyle(paragraph)
    return cloneTextStyle({ ...(baseStyle ?? {}), ...pending })
  }

  function resolveRuntimeKeyTable(
    model: DocModel,
    tableRuntimeKey: string,
    invalidateSources = false
  ): TableNode | undefined {
    const location = parseEmbeddedTableRuntimeKey(tableRuntimeKey)
    if (!location) return undefined

    const hostTable = model.nodes[location.hostTableIndex]
    if (!hostTable || hostTable.type !== "table") return undefined
    if (invalidateSources) hostTable.sourceXml = undefined
    const hostCell = hostTable.rows[location.hostRowIndex]?.cells[location.hostCellIndex]
    const rootContent = hostCell?.nodes[location.rootContentIndex]
    if (!rootContent || rootContent.type !== "table") return undefined

    let table = rootContent
    if (invalidateSources) table.sourceXml = undefined
    for (const segment of location.descendants) {
      const cell = table.rows[segment.rowIndex]?.cells[segment.cellIndex]
      const content = cell?.nodes[segment.contentIndex]
      if (!content || content.type !== "table") return undefined
      table = content
      if (invalidateSources) table.sourceXml = undefined
    }
    return table
  }

  function writeFirstTableCellText(
    table: TableNode,
    rowIndex: number,
    cellIndex: number,
    text: string
  ): boolean {
    const cell = table.rows[rowIndex]?.cells[cellIndex]
    if (!cell) return false
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
      table.sourceXml = undefined
      return true
    }
    // Diff instead of replacing the run list: a cell paragraph can carry a
    // footnote/endnote reference (or an image/form field) that must survive the
    // text edit, exactly like body paragraphs and index-based cell commits.
    // Comment offsets are character offsets, so they follow the same diff.
    const previousText = paragraphText(para)
    para.commentRanges = remapParagraphCommentRanges(
      previousText,
      text,
      para.commentRanges
    )
    para.children = distributeTextAcrossParagraphChildren(para, text)
    para.sourceXml = undefined
    table.sourceXml = undefined
    return true
  }

  function updateRuntimeKeyParagraph(
    model: DocModel,
    tableRuntimeKey: string,
    rowIndex: number,
    cellIndex: number,
    paragraphIndex: number,
    text: string
  ): boolean {
    const table = resolveRuntimeKeyTable(model, tableRuntimeKey, true)
    if (!table) return false
    const cell = table.rows[rowIndex]?.cells[cellIndex]
    const paragraph = cell?.nodes
      .filter((child): child is ParagraphNode => child.type === "paragraph")[paragraphIndex]
    if (!paragraph) return false
    const previousText = paragraphText(paragraph)
    if (previousText === text) return false
    const insertedStyle = insertedStyleForTextChange(paragraph, previousText, text)
    paragraph.children = distributeTextAcrossParagraphChildren(paragraph, text, {
      insertedStyle,
    })
    paragraph.commentRanges = remapParagraphCommentRanges(
      previousText,
      text,
      paragraph.commentRanges
    )
    paragraph.sourceXml = undefined
    table.sourceXml = undefined
    return true
  }

  // ── commitParagraphText ──────────────────────────────────────────
  const commitParagraphText = (nodeIndex: number, text: string): void => {
    applyChange(
      (current) => {
        const paragraph = current.nodes[nodeIndex]
        if (!paragraph || paragraph.type !== "paragraph") return current
        const previousText = paragraphText(paragraph)
        if (previousText === text) return current
        const insertedStyle = insertedStyleForTextChange(paragraph, previousText, text)
        return updateParagraphText(current, nodeIndex, text, { insertedStyle })
      },
      undefined,
      { textCommitKey: `paragraph:${nodeIndex}` },
    )
  }

  // ── commitTableCellText ──────────────────────────────────────────
  const commitTableCellText = (
    tableIndex: number, rowIndex: number, cellIndex: number, text: string
  ): void => {
    applyChange(
      (current) => {
        const node = current.nodes[tableIndex]
        if (!node || node.type !== "table") return current
        const cell = node.rows[rowIndex]?.cells[cellIndex]
        if (!cell) return current
        const currentText = cell.nodes
          .filter((child): child is ParagraphNode => child.type === "paragraph")
          .map((paragraph) => paragraphText(paragraph))
          .join("\n")
        if (currentText === text) return current
        // Diff per paragraph so run styles, images and form fields survive, and
        // the paragraph/table sourceXml caches are invalidated for export.
        return updateTableCellText(current, tableIndex, rowIndex, cellIndex, text)
      },
      undefined,
      { textCommitKey: `table-cell:${tableIndex}:${rowIndex}:${cellIndex}:0` },
    )
  }

  // ── commitTableCellTextByRuntimeKey ──────────────────────────────
  const commitTableCellTextByRuntimeKey = (
    tableRuntimeKey: string, rowIndex: number, cellIndex: number, text: string
  ): void => {
    applyChange(
      (current) => {
        const nextModel = cloneHostTableForRuntimeKey(current, tableRuntimeKey)
        if (!nextModel) return current
        const node = resolveRuntimeKeyTable(nextModel, tableRuntimeKey, true)
        if (!node) return current
        return writeFirstTableCellText(node, rowIndex, cellIndex, text)
          ? nextModel
          : current
      },
      undefined,
      { textCommitKey: `runtime-table:${tableRuntimeKey}:${rowIndex}:${cellIndex}:0` },
    )
  }

  // ── commitTableCellParagraphTextByRuntimeKey ─────────────────────
  const commitTableCellParagraphTextByRuntimeKey = (
    tableRuntimeKey: string,
    rowIndex: number,
    cellIndex: number,
    paragraphIndex: number,
    text: string
  ): void => {
    applyChange((current) => {
      const nextModel = cloneHostTableForRuntimeKey(current, tableRuntimeKey)
      if (!nextModel) return current
      return updateRuntimeKeyParagraph(
        nextModel, tableRuntimeKey, rowIndex, cellIndex, paragraphIndex, text
      )
        ? nextModel
        : current
    },
    undefined,
    { textCommitKey: `runtime-table:${tableRuntimeKey}:${rowIndex}:${cellIndex}:${paragraphIndex}` },
  )
  }

  // ── commitTableCellParagraphTextRecursive ────────────────────────
  const commitTableCellParagraphTextRecursive = (
    tableIndex: number, rowIndex: number, cellIndex: number, paragraphIndex: number, text: string
  ): void => {
    applyChange(
      (current) => {
        const node = current.nodes[tableIndex]
        if (!node || node.type !== "table") return current
        const cell = node.rows[rowIndex]?.cells[cellIndex]
        const paragraph = cell?.nodes
          .filter((child): child is ParagraphNode => child.type === "paragraph")[paragraphIndex]
        if (!cell || !paragraph) return current
        const previousText = paragraphText(paragraph)
        if (previousText === text) return current
        const insertedStyle = insertedStyleForTextChange(paragraph, previousText, text)
        return updateTableCellParagraphText(
          current, tableIndex, rowIndex, cellIndex, paragraphIndex, text, { insertedStyle }
        )
      },
      undefined,
      { textCommitKey: `table-cell:${tableIndex}:${rowIndex}:${cellIndex}:${paragraphIndex}` },
    )
  }

  // ── commitSectionParagraphText ───────────────────────────────────
  const commitSectionParagraphText = (
    location: { region: "header" | "footer"; partName: string; nodeIndex: number },
    text: string
  ): void => {
    applyChange(
      (current) => updateSectionParagraphTextAtLocation(current, location, text),
      undefined,
      {
        textCommitKey:
          `section:${location.region}:${location.partName}:${location.nodeIndex}`,
      },
    )
  }

  // ── appendParagraph ──────────────────────────────────────────────
  const appendParagraph = (text?: string): number => {
    let nodeIndex = -1
    applyChange((current) => {
      const nextModel = cloneDocModelShallow(current)
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

  // ── insertTextAtSelection ────────────────────────────────────────
  // Inserts text at a collapsed caret, or replaces an expanded selection.
  // Unlike replaceExpandedSelection it also handles the collapsed case, which
  // is what clipboard paste needs.
  const insertTextAtSelection = (
    text: string,
    range?: DocxTextRange
  ): DocxTextRange | undefined => {
    const sourceRange = range ?? ctx.activeTextRangeSnapshot.value
    if (!sourceRange) return undefined

    const replacementText = text.replace(/\r\n?/g, "\n")
    if (!replacementText) return undefined

    if (compareTextRangeBoundaries(sourceRange.start, sourceRange.end) < 0) {
      return replaceExpandedSelection(replacementText, sourceRange)
    }

    const location = sourceRange.start.location
    const { paragraph } = getParagraphAtLocation(ctx.model.value, location)
    if (!paragraph) return undefined

    const currentText = paragraphText(paragraph)
    const offset = normalizeRangeBoundaryParagraphOffset(paragraph, sourceRange.start.offset)
    const nextText = `${currentText.slice(0, offset)}${replacementText}${currentText.slice(offset)}`
    if (nextText === currentText) return undefined

    const insertedStyle = cloneTextStyle(
      ctx.pendingRunStyleSnapshot.value ??
        firstTextStyleAtOffset(paragraph, offset, true) ??
        firstRunStyle(paragraph)
    )

    applyChange(
      (current) => updateParagraphTextAtLocation(current, location, nextText, { insertedStyle }),
      "Inserted text"
    )

    const caretOffset = offset + replacementText.length
    const collapsedRange: DocxTextRange = {
      start: { location: cloneTextRangeLocation(location), offset: caretOffset },
      end: { location: cloneTextRangeLocation(location), offset: caretOffset },
    }

    ctx.suppressSelectionReset.value = true
    ctx.selection.value = location.kind === "paragraph"
      ? { kind: "paragraph", nodeIndex: location.nodeIndex }
      : {
          kind: "table-cell",
          tableIndex: location.tableIndex,
          rowIndex: location.rowIndex,
          cellIndex: location.cellIndex,
        }
    ctx.activeTextRange.value = collapsedRange
    ctx.pendingRunStyle.value = undefined

    return collapsedRange
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
      const next = cloneDocModelNodes(current, [targetLocation.nodeIndex])
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

      // Split comment ranges with the text: each side keeps the part of a range
      // that survives on it, so highlights stay exact across the break.
      const draftCommentRanges = remapParagraphCommentRanges(
        paragraphText(paragraphNode),
        normalizedDraftText,
        paragraphNode.commentRanges
      )
      const beforeCommentRanges = splitParagraphCommentRanges(
        draftCommentRanges,
        safeStart
      ).before
      const afterCommentRanges = splitParagraphCommentRanges(
        draftCommentRanges,
        safeEnd
      ).after

      paragraphNode.style = cloneParagraphStyle(splitParagraphStyle)
      paragraphNode.children = splitChildren.beforeChildren
      paragraphNode.commentRanges =
        beforeCommentRanges.length > 0 ? beforeCommentRanges : undefined
      paragraphNode.sourceXml = undefined

      next.nodes.splice(insertionIndex, 0, {
        type: "paragraph",
        style: cloneParagraphStyle(splitParagraphStyle),
        children: splitChildren.afterChildren,
        commentRanges: afterCommentRanges.length > 0 ? afterCommentRanges : undefined,
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
      let next = cloneDocModelShallow(current)
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

      const draftCommentRanges = remapParagraphCommentRanges(
        paragraphText(paragraphNode),
        normalizedDraftText,
        paragraphNode.commentRanges
      )
      const beforeCommentRanges = splitParagraphCommentRanges(
        draftCommentRanges,
        safeStart
      ).before
      const afterCommentRanges = splitParagraphCommentRanges(
        draftCommentRanges,
        safeEnd
      ).after

      next = updateParagraphText(next, targetLocation.nodeIndex, textForCurrentParagraph, {
        insertedStyle: cloneTextStyle(ctx.pendingRunStyleSnapshot.value),
      })
      paragraphNode = next.nodes[targetLocation.nodeIndex]
      if (!paragraphNode || paragraphNode.type !== "paragraph") return current
      paragraphNode.commentRanges =
        beforeCommentRanges.length > 0 ? beforeCommentRanges : undefined
      paragraphNode.sourceXml = undefined

      next.nodes.splice(insertionIndex, 0, {
        type: "paragraph",
        style: cloneParagraphStyle(paragraphNode.style),
        children: [{ type: "text", text: textForInsertedParagraph, style: insertRunStyle }],
        commentRanges: afterCommentRanges.length > 0 ? afterCommentRanges : undefined,
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
    commitTableCellTextByRuntimeKey,
    commitTableCellParagraphTextByRuntimeKey,
    commitTableCellParagraphTextRecursive,
    commitSectionParagraphText,
    appendParagraph,
    splitParagraphAtSelection,
    insertListItemAfterSelection,
    replaceExpandedSelection,
    deleteExpandedSelection,
    insertTextAtSelection,
  }
}
