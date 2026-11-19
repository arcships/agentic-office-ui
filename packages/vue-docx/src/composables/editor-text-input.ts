// editor-text-input.ts — text input operations (contentEditable + draft)
//
// Extracted from useDocxEditor.ts.

import type { DocModel, ParagraphNode, TableNode, TextRunNode } from "@arcships/docx-core"
import type { DocxClipboardParagraph } from "@arcships/docx-core"
import type { DocxTextRange, ParagraphLocation } from "@arcships/docx-core"
import {
  cloneDocModelNodes,
  cloneDocModelShallow,
  parseEmbeddedTableRuntimeKey,
  updateParagraphText,
  normalizeTextRange,
  compareTextRangeBoundaries,
  paragraphText,
  distributeTextAcrossParagraphChildren,
  normalizeRangeBoundaryParagraphOffset,
  updateParagraphTextAtLocation,
  updateTableCellText,
  updateTableCellParagraphText,
  updateSectionParagraphTextAtLocation,
  cloneParagraphStyle,
  cloneTextStyle,
  firstTextStyleAtOffset,
  firstRunStyle,
  remapParagraphCommentRanges,
  splitParagraphCommentRanges,
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
  insertDocxClipboardParagraphs,
} from "@arcships/docx-core"
import type { DocxListType } from "@arcships/docx-core"
import type { ApplyModelChangeOptions, EditorCore } from "./editor-shared"
import { cloneHostTableForRuntimeKey } from "./editor-clone-helpers"
import { collapseSelectionRange, type CollapsedRangeResult } from "./editor-range-collapse"

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
  // Handles same-paragraph, cross-paragraph, table-cell, and entire-document
  // cases through the shared `collapseSelectionRange` helper.
  const applyCollapsedCaret = (collapsed: CollapsedRangeResult): DocxTextRange => {
    const collapsedRange: DocxTextRange = {
      start: {
        location: cloneTextRangeLocation(collapsed.location),
        offset: collapsed.offset,
      },
      end: {
        location: cloneTextRangeLocation(collapsed.location),
        offset: collapsed.offset,
      },
    }

    ctx.suppressSelectionReset.value = true
    if (collapsed.location.kind === "paragraph") {
      ctx.selection.value = { kind: "paragraph", nodeIndex: collapsed.location.nodeIndex }
    } else {
      ctx.selection.value = {
        kind: "table-cell",
        tableIndex: collapsed.location.tableIndex,
        rowIndex: collapsed.location.rowIndex,
        cellIndex: collapsed.location.cellIndex,
      }
    }
    ctx.activeTextRange.value = collapsedRange
    ctx.pendingRunStyle.value = undefined

    return collapsedRange
  }

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
    let collapsed: CollapsedRangeResult | undefined

    applyChange(
      (current) => {
        collapsed = collapseSelectionRange(
          current,
          normalizedRange,
          replacementText,
          ctx.pendingRunStyleSnapshot.value
        )
        return collapsed ? collapsed.model : current
      },
      replacementText.length > 0 ? "Replaced selection" : "Deleted selection"
    )

    return collapsed ? applyCollapsedCaret(collapsed) : undefined
  }

  // ── insertClipboardParagraphsAtSelection ─────────────────────────
  // Rich paste: replace the selection (if any) and insert the clipboard runs in
  // the same transaction. An expanded, cross-paragraph selection is collapsed
  // first, so the paste stays one undo step and keeps its formatting.
  const insertClipboardParagraphsAtSelection = (
    paragraphs: DocxClipboardParagraph[],
    range?: DocxTextRange
  ): DocxTextRange | undefined => {
    const sourceRange = range ?? ctx.activeTextRangeSnapshot.value
    if (!sourceRange || paragraphs.length === 0) return undefined

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

    let inserted: { model: DocModel; caret: DocxTextRange } | undefined

    applyChange((current) => {
      let base = current
      let targetRange = normalizedRange
      if (compareTextRangeBoundaries(normalizedRange.start, normalizedRange.end) < 0) {
        const collapsed = collapseSelectionRange(
          current,
          normalizedRange,
          "",
          ctx.pendingRunStyleSnapshot.value
        )
        if (!collapsed) return current
        base = collapsed.model
        const collapsedRange: DocxTextRange = {
          start: {
            location: cloneTextRangeLocation(collapsed.location),
            offset: collapsed.offset,
          },
          end: {
            location: cloneTextRangeLocation(collapsed.location),
            offset: collapsed.offset,
          },
        }
        targetRange = collapsedRange
      }

      inserted = insertDocxClipboardParagraphs(base, targetRange, paragraphs)
      return inserted ? inserted.model : current
    })

    if (!inserted) return undefined

    return applyCollapsedCaret({
      model: inserted.model,
      location: inserted.caret.start.location,
      offset: inserted.caret.start.offset,
    })
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
    insertClipboardParagraphsAtSelection,
    insertTextAtSelection,
  }
}
