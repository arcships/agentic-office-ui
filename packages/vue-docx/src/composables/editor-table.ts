// editor-table.ts — table insert/delete/move/resize operations
//
// Extracted from useDocxEditor.ts.

import type {
  DocModel,
  DocxEditorStatusMessages,
  DocxTableCellSelectionRange,
  ParagraphNode,
} from "@arcships/docx-core"
import {
  cloneDocModelShallow,
  cloneDocNode,
  cloneParagraphStyle,
  createDefaultEditorTableBorders,
  createEmptyParagraphFromTemplate,
  deleteTableColumn as deleteTableColumnOnGrid,
  deleteTableRow as deleteTableRowOnGrid,
  insertTableColumn as insertTableColumnOnGrid,
  insertTableRow as insertTableRowOnGrid,
  mergeTableCells as mergeTableCellsOnGrid,
  mergeTableFloatingStyle,
  moveBodyNodeInModel,
  paragraphLocationFromTextRangeLocation,
  paragraphText,
  parseSectionLayout,
  shiftSectionStartIndexesForInsertion,
  shiftSectionStartIndexesForRemoval,
  splitParagraphChildrenAtTextOffsets,
  splitParagraphStyleWithDefaultSpacing,
  splitTableCell as splitTableCellOnGrid,
  tableColumnCount,
  type DocxTableFloatingUpdate,
} from "@arcships/docx-core"
import type { ApplyModelChangeOptions, EditorCore } from "./editor-shared"

const INSERTED_TABLE_ROWS = 3
const INSERTED_TABLE_COLUMNS = 3
const TWIPS_PER_PIXEL = 15
const MIN_INSERTED_TABLE_WIDTH_TWIPS = 1440

type TableCellPosition = { rowIndex: number; cellIndex: number }

function insertedTableContentWidthTwips(model: DocModel): number {
  const sectionXml =
    model.metadata.sections?.[0]?.sectionPropertiesXml ??
    model.metadata.sectionPropertiesXml
  const layout = parseSectionLayout(sectionXml)
  const contentWidthPx =
    layout.pageWidthPx - layout.marginsPx.left - layout.marginsPx.right
  return Math.max(
    MIN_INSERTED_TABLE_WIDTH_TWIPS,
    Math.round(contentWidthPx * TWIPS_PER_PIXEL)
  )
}

function insertedTableColumnWidths(widthTwips: number): number[] {
  const baseWidth = Math.floor(widthTwips / INSERTED_TABLE_COLUMNS)
  const widths = Array.from(
    { length: INSERTED_TABLE_COLUMNS },
    () => Math.max(1, baseWidth)
  )
  widths[widths.length - 1] += widthTwips - baseWidth * INSERTED_TABLE_COLUMNS
  return widths
}

function insertedTableNode(model: DocModel) {
  const widthTwips = insertedTableContentWidthTwips(model)
  const columnWidthsTwips = insertedTableColumnWidths(widthTwips)
  return {
    type: "table" as const,
    style: {
      layout: "fixed" as const,
      widthTwips,
      columnWidthsTwips,
      borders: createDefaultEditorTableBorders(),
    },
    rows: Array.from({ length: INSERTED_TABLE_ROWS }, () => ({
      type: "table-row" as const,
      cells: Array.from(
        { length: INSERTED_TABLE_COLUMNS },
        (_, columnIndex) => ({
          type: "table-cell" as const,
          style: { widthTwips: columnWidthsTwips[columnIndex] },
          nodes: [createEmptyParagraphFromTemplate()],
        })
      ),
    })),
  }
}

/** Copy-on-write shell plus the one top-level table node the command edits. */
function cloneModelForTableEdit(
  current: DocModel,
  tableIndex?: number
): DocModel {
  const next = cloneDocModelShallow(current)
  if (tableIndex !== undefined) {
    const node = current.nodes[tableIndex]
    if (node) {
      next.nodes[tableIndex] = cloneDocNode(node)
    }
  }
  return next
}

/** Shallow shell with copied section wrappers for startNodeIndex edits. */
function cloneModelForSectionEdit(current: DocModel): DocModel {
  const next = cloneDocModelShallow(current)
  next.metadata.sections = current.metadata.sections?.map((section) => ({
    ...section,
  }))
  return next
}

export function createEditorTable(
  ctx: EditorCore,
  applyChange: (
    updater: (current: DocModel) => DocModel,
    successStatus?: string,
    options?: ApplyModelChangeOptions
  ) => void,
) {
  const unsupported = (key: keyof DocxEditorStatusMessages): void => {
    ctx.status.value = ctx.statusMessage(key)
  }

  const refuse = (key: keyof DocxEditorStatusMessages): void => {
    ctx.status.value = ctx.statusMessage(key)
  }

  const selectMergedCell = (
    tableIndex: number,
    position: TableCellPosition
  ): void => {
    const selection = {
      kind: "table-cell" as const,
      tableIndex,
      rowIndex: position.rowIndex,
      cellIndex: position.cellIndex,
    }
    ctx.selection.value = selection
    ctx.selectionSnapshot.value = selection
    ctx.suppressSelectionReset.value = true
    ctx.activeTextRange.value = undefined
    ctx.activeTextRangeSnapshot.value = undefined
  }

  // Inserts a table at the caret instead of appending it to the document. The
  // caret paragraph is split when the caret sits in the middle of its text, and
  // a caret inside a table cell nests the new table in that cell.
  const insertTable = (): void => {
    const range = ctx.activeTextRangeSnapshot.value
    const location = range
      ? paragraphLocationFromTextRangeLocation(range.start.location)
      : undefined
    const caretOffset = range ? Math.max(0, Math.round(range.start.offset)) : undefined

    applyChange((current) => {
      const nextModel = cloneModelForSectionEdit(current)
      if (location?.kind === "paragraph") {
        const node = current.nodes[location.nodeIndex]
        if (node) nextModel.nodes[location.nodeIndex] = cloneDocNode(node)
      } else if (location?.kind === "table-cell") {
        const node = current.nodes[location.tableIndex]
        if (node) nextModel.nodes[location.tableIndex] = cloneDocNode(node)
      }
      const table = insertedTableNode(nextModel)

      if (!location || caretOffset === undefined) {
        nextModel.nodes.push(table)
        return nextModel
      }

      if (location.kind === "table-cell") {
        const tableNode = nextModel.nodes[location.tableIndex]
        if (!tableNode || tableNode.type !== "table") {
          nextModel.nodes.push(table)
          return nextModel
        }
        const cell = tableNode.rows[location.rowIndex]?.cells[location.cellIndex]
        if (!cell) {
          nextModel.nodes.push(table)
          return nextModel
        }
        const insertIdx = Math.max(0, Math.min(cell.nodes.length, location.paragraphIndex + 1))
        cell.nodes.splice(insertIdx, 0, table)
        tableNode.sourceXml = undefined
        return nextModel
      }

      const paragraphNode = nextModel.nodes[location.nodeIndex]
      if (!paragraphNode || paragraphNode.type !== "paragraph") {
        nextModel.nodes.push(table)
        return nextModel
      }

      const text = paragraphText(paragraphNode)
      const splitOffset = Math.max(0, Math.min(text.length, caretOffset))
      if (splitOffset <= 0) {
        nextModel.nodes.splice(location.nodeIndex, 0, table)
        shiftSectionStartIndexesForInsertion(nextModel, location.nodeIndex)
        return nextModel
      }
      if (splitOffset >= text.length) {
        nextModel.nodes.splice(location.nodeIndex + 1, 0, table)
        shiftSectionStartIndexesForInsertion(nextModel, location.nodeIndex + 1)
        return nextModel
      }

      // Split the caret paragraph so the table lands exactly at the caret.
      const splitStyle = cloneParagraphStyle(
        splitParagraphStyleWithDefaultSpacing(paragraphNode.style, paragraphNode.sourceXml)
      )
      const splitChildren = splitParagraphChildrenAtTextOffsets(
        paragraphNode,
        text,
        splitOffset,
        splitOffset
      )
      paragraphNode.style = cloneParagraphStyle(splitStyle)
      paragraphNode.children = splitChildren.beforeChildren
      paragraphNode.sourceXml = undefined
      const afterParagraph: ParagraphNode = {
        type: "paragraph",
        style: cloneParagraphStyle(splitStyle),
        children: splitChildren.afterChildren,
      }
      nextModel.nodes.splice(location.nodeIndex + 1, 0, table, afterParagraph)
      shiftSectionStartIndexesForInsertion(nextModel, location.nodeIndex + 1)
      shiftSectionStartIndexesForInsertion(nextModel, location.nodeIndex + 1)
      return nextModel
    }, ctx.statusMessage("insertedTable"))
  }

  const insertTableRow = (
    tableIndex: number, rowIndex: number, direction: "above" | "below"
  ): void => {
    applyChange((current) => {
      const nextModel = cloneModelForTableEdit(current, tableIndex)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") {
        refuse("cannotInsertRowTableNotFound")
        return current
      }
      // `rowIndex` is the reference row; `direction` decides which side of it
      // the new row lands on.
      if (!insertTableRowOnGrid(node, rowIndex, direction)) {
        refuse("cannotInsertTableRow")
        return current
      }
      return nextModel
    }, ctx.statusMessage("insertedTableRow", { direction }))
  }

  const insertTableColumn = (
    tableIndex: number, cellIndex: number, direction: "left" | "right", rowIndex?: number
  ): void => {
    applyChange((current) => {
      const nextModel = cloneModelForTableEdit(current, tableIndex)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") {
        refuse("cannotInsertColumnTableNotFound")
        return current
      }
      // `cellIndex` is the reference cell; `direction` decides the side.
      if (!insertTableColumnOnGrid(node, cellIndex, direction, rowIndex)) {
        refuse("cannotInsertTableColumn")
        return current
      }
      return nextModel
    }, ctx.statusMessage("insertedTableColumn", { direction }))
  }

  const deleteTableRow = (tableIndex: number, rowIndex: number): void => {
    applyChange((current) => {
      const nextModel = cloneModelForTableEdit(current, tableIndex)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") {
        refuse("cannotDeleteRowTableNotFound")
        return current
      }
      if (!deleteTableRowOnGrid(node, rowIndex)) {
        refuse("cannotDeleteOnlyTableRow")
        return current
      }
      return nextModel
    }, ctx.statusMessage("deletedTableRow"))
  }

  const deleteTableColumn = (
    tableIndex: number, cellIndex: number, rowIndex?: number
  ): void => {
    applyChange((current) => {
      const nextModel = cloneModelForTableEdit(current, tableIndex)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") {
        refuse("cannotDeleteColumnTableNotFound")
        return current
      }
      if (tableColumnCount(node) <= 1) {
        refuse("cannotDeleteOnlyTableColumn")
        return current
      }
      if (!deleteTableColumnOnGrid(node, cellIndex, rowIndex)) {
        refuse("cannotDeleteTableColumn")
        return current
      }
      return nextModel
    }, ctx.statusMessage("deletedTableColumn"))
  }

  const deleteTable = (tableIndex: number): void => {
    const currentNodes = ctx.modelSnapshot.value.nodes
    const node = currentNodes[tableIndex]
    if (!node || node.type !== "table") {
      refuse("cannotDeleteTableNotFound")
      return
    }

    const fallbackNodeIndex = Math.min(
      Math.max(0, tableIndex),
      Math.max(0, currentNodes.length - 2)
    )

    applyChange((current) => {
      const nextModel = cloneModelForSectionEdit(current)
      const tableNode = nextModel.nodes[tableIndex]
      if (!tableNode || tableNode.type !== "table") return current
      nextModel.nodes.splice(tableIndex, 1)
      if (!nextModel.nodes.some((candidate) => candidate.type === "paragraph")) {
        nextModel.nodes.push(createEmptyParagraphFromTemplate())
      }
      shiftSectionStartIndexesForRemoval(nextModel, tableIndex)
      return nextModel
    }, ctx.statusMessage("deletedTable"), {
      selection: { kind: "paragraph", nodeIndex: fallbackNodeIndex },
      activeTextRange: undefined,
    })
  }

  const clearTableCellContents = (
    tableIndex: number, cells: Array<{ rowIndex: number; cellIndex: number }>
  ): void => {
    applyChange((current) => {
      const nextModel = cloneModelForTableEdit(current, tableIndex)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") {
        refuse("cannotClearCellsTableNotFound")
        return current
      }
      for (const { rowIndex, cellIndex } of cells) {
        const cell = node.rows[rowIndex]?.cells[cellIndex]
        if (cell) {
          cell.nodes = [createEmptyParagraphFromTemplate()]
        }
      }
      node.sourceXml = undefined
      return nextModel
    }, ctx.statusMessage("clearedTableCells"))
  }

  const setTableColumnWidths = (tableIndex: number, widthsPx: number[]): void => {
    if (
      widthsPx.length === 0 ||
      widthsPx.some((width) => !Number.isFinite(width) || width <= 0)
    ) {
      unsupported("tableColumnWidthPositive")
      return
    }

    const table = ctx.modelSnapshot.value.nodes[tableIndex]
    const columnCount = table?.type === "table" ? tableColumnCount(table) : 0
    if (table?.type !== "table" || columnCount !== widthsPx.length) {
      unsupported("tableColumnWidthTargetUnavailable")
      return
    }

    const normalizedPx = widthsPx.map((width) => Math.max(1, Math.round(width)))
    const columnWidthsTwips = normalizedPx.map((width) => width * 15)
    applyChange((current) => {
      const nextModel = cloneModelForTableEdit(current, tableIndex)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") return current
      node.style = {
        ...(node.style ?? {}),
        layout: "fixed",
        widthTwips: columnWidthsTwips.reduce((sum, width) => sum + width, 0),
        columnWidthsTwips,
      }
      // A manual column resize materializes the physical grid: the logical
      // zero-width columns no longer describe what is drawn, so the collapsed
      // marker must not survive into export.
      if (node.style) delete node.style.collapsedGridColumns
      node.sourceXml = undefined
      return nextModel
    }, ctx.statusMessage("resizedTableColumns"))
  }

  const setTableRowHeight = (
    tableIndex: number,
    rowIndex: number,
    heightPx: number
  ): void => {
    if (!Number.isFinite(heightPx) || heightPx <= 0) {
      unsupported("tableRowHeightPositive")
      return
    }

    const table = ctx.modelSnapshot.value.nodes[tableIndex]
    if (
      table?.type !== "table" ||
      !Number.isInteger(rowIndex) ||
      !table.rows[rowIndex]
    ) {
      unsupported("tableRowHeightTargetUnavailable")
      return
    }

    const heightTwips = Math.max(1, Math.round(Math.round(heightPx) * TWIPS_PER_PIXEL))
    applyChange((current) => {
      const nextModel = cloneModelForTableEdit(current, tableIndex)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") return current
      const row = node.rows[rowIndex]
      if (!row) return current
      if (row.style?.heightTwips === heightTwips && row.style?.heightRule === "atLeast") {
        return current
      }
      // A dragged height is a minimum: content that needs more room still grows
      // the row instead of being clipped like an `exact` height would.
      row.style = { ...(row.style ?? {}), heightTwips, heightRule: "atLeast" }
      node.sourceXml = undefined
      return nextModel
    }, ctx.statusMessage("resizedTableRow"))
  }

  /**
   * Floats a body table at a drop position, or (`update === null`) puts it back
   * into the text flow. The model carries the position in `TableStyle.floating`,
   * which the wrapper style renders and the serializer writes as `w:tblpPr`.
   */
  const setTableFloating = (
    tableIndex: number,
    update: DocxTableFloatingUpdate | null
  ): void => {
    const table = ctx.modelSnapshot.value.nodes[tableIndex]
    if (table?.type !== "table") {
      refuse("cannotFloatTableNotFound")
      return
    }

    if (update === null) {
      // An inline table has nothing to clear, and a no-op must not open an undo
      // entry the user cannot see.
      if (!table.style?.floating) return
      applyChange((current) => {
        const nextModel = cloneModelForTableEdit(current, tableIndex)
        const node = nextModel.nodes[tableIndex]
        if (!node || node.type !== "table") return current
        if (node.style) {
          const { floating: _floating, ...rest } = node.style
          node.style = rest
        }
        // The rebuilt table must not carry the old `w:tblpPr`.
        node.sourceXml = undefined
        return nextModel
      }, ctx.statusMessage("inlinedTable"))
      return
    }

    const floating = mergeTableFloatingStyle(table.style?.floating, update)
    if (!floating) return
    if (
      JSON.stringify(floating) ===
      JSON.stringify(table.style?.floating ?? null)
    ) {
      return
    }

    applyChange((current) => {
      const nextModel = cloneModelForTableEdit(current, tableIndex)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") return current
      node.style = { ...(node.style ?? {}), floating }
      node.sourceXml = undefined
      return nextModel
    }, ctx.statusMessage("floatedTable"))
  }

  const mergeTableCells = (
    tableIndex: number, range: DocxTableCellSelectionRange
  ): void => {
    let merged: TableCellPosition | undefined
    applyChange((current) => {
      const nextModel = cloneModelForTableEdit(current, tableIndex)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") {
        refuse("cannotMergeCellsTableNotFound")
        return current
      }
      merged = mergeTableCellsOnGrid(node, range)
      if (!merged) {
        refuse("cannotMergeSingleCell")
        return current
      }
      return nextModel
    }, ctx.statusMessage("mergedTableCells"))

    if (merged) {
      selectMergedCell(tableIndex, merged)
    }
  }

  const splitTableCell = (
    tableIndex: number, rowIndex: number, cellIndex: number
  ): void => {
    let split: TableCellPosition | undefined
    applyChange((current) => {
      const nextModel = cloneModelForTableEdit(current, tableIndex)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") {
        refuse("cannotSplitCellTableNotFound")
        return current
      }
      split = splitTableCellOnGrid(node, rowIndex, cellIndex)
      if (!split) {
        refuse("cannotSplitCellNotMerged")
        return current
      }
      return nextModel
    }, ctx.statusMessage("splitTableCell"))

    if (split) {
      selectMergedCell(tableIndex, split)
    }
  }

  const moveTable = (tableIndex: number, targetNodeIndex: number): void => {
    applyChange((current) => {
      const nextModel = cloneModelForSectionEdit(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") {
        refuse("cannotMoveTableNotFound")
        return current
      }
      if (!moveBodyNodeInModel(nextModel, tableIndex, targetNodeIndex)) {
        return current
      }
      return nextModel
    }, ctx.statusMessage("movedTable"))
  }

  // Paragraph and table reordering share one body-node move: section starts are
  // re-anchored by `moveBodyNodeInModel`, so a node that crosses a section break
  // joins the section at its new position instead of redefining the break.
  const moveBodyNode = (
    fromNodeIndex: number,
    targetNodeIndex: number
  ): boolean => {
    let moved = false
    applyChange((current) => {
      const node = current.nodes[fromNodeIndex]
      if (!node) {
        refuse("cannotMoveBodyNodeNotFound")
        return current
      }
      const nextModel = cloneModelForSectionEdit(current)
      if (!moveBodyNodeInModel(nextModel, fromNodeIndex, targetNodeIndex)) {
        return current
      }
      moved = true
      return nextModel
    }, ctx.statusMessage("movedBodyNode"))
    return moved
  }

  const moveEmbeddedTableToBody = (
    tableRuntimeKey: string, targetNodeIndex: number
  ): void => {
    // Extracts a nested table from a table cell into the document body.
    // tableRuntimeKey identifies the nested table via its parent location.
    // Reserved for nested-table scenarios — the runtime key format will be
    // determined during component integration.
    const parts = tableRuntimeKey.split(":")
    if (parts.length < 3) return

    applyChange((current) => {
      const nextModel = cloneModelForSectionEdit(current)
      const parentTableIndex = Number(parts[0])
      const parentRowIndex = Number(parts[1])
      const parentCellIndex = Number(parts[2])
      const nestedTableIndex = parts[3] !== undefined ? Number(parts[3]) : undefined

      const sourceParentTable = current.nodes[parentTableIndex]
      if (sourceParentTable?.type === "table") {
        nextModel.nodes[parentTableIndex] = cloneDocNode(sourceParentTable)
      }

      const parentTable = nextModel.nodes[parentTableIndex]
      if (!parentTable || parentTable.type !== "table") {
        refuse("cannotMoveEmbeddedHostTableNotFound")
        return current
      }

      const parentCell = parentTable.rows[parentRowIndex]?.cells[parentCellIndex]
      if (!parentCell) {
        refuse("cannotMoveEmbeddedHostCellNotFound")
        return current
      }

      if (nestedTableIndex !== undefined && !isNaN(nestedTableIndex)) {
        const cellNodes = parentCell.nodes
        const nestedTable = cellNodes[nestedTableIndex]
        if (!nestedTable || nestedTable.type !== "table") {
          refuse("cannotMoveEmbeddedTableNotFound")
          return current
        }

        // Extract nested table from cell
        cellNodes.splice(nestedTableIndex, 1)

        // Insert into body at target position
        const adjustedTarget = Math.max(
          0,
          Math.min(targetNodeIndex, nextModel.nodes.length)
        )
        shiftSectionStartIndexesForInsertion(nextModel, adjustedTarget)
        nextModel.nodes.splice(adjustedTarget, 0, nestedTable)
        parentTable.sourceXml = undefined
      } else {
        return current
      }

      return nextModel
    }, ctx.statusMessage("movedEmbeddedTable"))
  }

  return {
    insertTable,
    insertTableRow,
    insertTableColumn,
    deleteTableRow,
    deleteTableColumn,
    deleteTable,
    clearTableCellContents,
    setTableColumnWidths,
    setTableRowHeight,
    setTableFloating,
    mergeTableCells,
    splitTableCell,
    moveTable,
    moveBodyNode,
    moveEmbeddedTableToBody,
  }
}
