// editor-table.ts — table insert/delete/move/resize operations
//
// Extracted from useDocxEditor.ts.

import type {
  DocModel,
  DocxTableCellSelectionRange,
  ParagraphNode,
} from "@arcships/docx-core"
import {
  cloneDocModel,
  cloneParagraphStyle,
  createDefaultEditorTableBorders,
  createEmptyParagraphFromTemplate,
  deleteTableColumn as deleteTableColumnOnGrid,
  deleteTableRow as deleteTableRowOnGrid,
  insertTableColumn as insertTableColumnOnGrid,
  insertTableRow as insertTableRowOnGrid,
  mergeTableCells as mergeTableCellsOnGrid,
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

export function createEditorTable(
  ctx: EditorCore,
  applyChange: (
    updater: (current: DocModel) => DocModel,
    successStatus?: string,
    options?: ApplyModelChangeOptions
  ) => void,
) {
  const unsupported = (feature: string): void => {
    ctx.status.value = `Unsupported: ${feature}`
  }

  const refuse = (reason: string): void => {
    ctx.status.value = reason
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
      const nextModel = cloneDocModel(current)
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
    }, "Inserted table")
  }

  const insertTableRow = (
    tableIndex: number, rowIndex: number, direction: "above" | "below"
  ): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") {
        refuse("Cannot insert a row: table not found")
        return current
      }
      // `rowIndex` is the reference row; `direction` decides which side of it
      // the new row lands on.
      if (!insertTableRowOnGrid(node, rowIndex, direction)) {
        refuse("Cannot insert a table row")
        return current
      }
      return nextModel
    }, `Inserted table row ${direction}`)
  }

  const insertTableColumn = (
    tableIndex: number, cellIndex: number, direction: "left" | "right", rowIndex?: number
  ): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") {
        refuse("Cannot insert a column: table not found")
        return current
      }
      // `cellIndex` is the reference cell; `direction` decides the side.
      if (!insertTableColumnOnGrid(node, cellIndex, direction, rowIndex)) {
        refuse("Cannot insert a table column")
        return current
      }
      return nextModel
    }, `Inserted table column ${direction}`)
  }

  const deleteTableRow = (tableIndex: number, rowIndex: number): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") {
        refuse("Cannot delete a row: table not found")
        return current
      }
      if (!deleteTableRowOnGrid(node, rowIndex)) {
        refuse("Cannot delete the only table row")
        return current
      }
      return nextModel
    }, "Deleted table row")
  }

  const deleteTableColumn = (
    tableIndex: number, cellIndex: number, rowIndex?: number
  ): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") {
        refuse("Cannot delete a column: table not found")
        return current
      }
      if (tableColumnCount(node) <= 1) {
        refuse("Cannot delete the only table column")
        return current
      }
      if (!deleteTableColumnOnGrid(node, cellIndex, rowIndex)) {
        refuse("Cannot delete the table column")
        return current
      }
      return nextModel
    }, "Deleted table column")
  }

  const deleteTable = (tableIndex: number): void => {
    const currentNodes = ctx.modelSnapshot.value.nodes
    const node = currentNodes[tableIndex]
    if (!node || node.type !== "table") {
      refuse("Cannot delete the table: table not found")
      return
    }

    const fallbackNodeIndex = Math.min(
      Math.max(0, tableIndex),
      Math.max(0, currentNodes.length - 2)
    )

    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const tableNode = nextModel.nodes[tableIndex]
      if (!tableNode || tableNode.type !== "table") return current
      nextModel.nodes.splice(tableIndex, 1)
      if (!nextModel.nodes.some((candidate) => candidate.type === "paragraph")) {
        nextModel.nodes.push(createEmptyParagraphFromTemplate())
      }
      shiftSectionStartIndexesForRemoval(nextModel, tableIndex)
      return nextModel
    }, "Deleted table", {
      selection: { kind: "paragraph", nodeIndex: fallbackNodeIndex },
      activeTextRange: undefined,
    })
  }

  const clearTableCellContents = (
    tableIndex: number, cells: Array<{ rowIndex: number; cellIndex: number }>
  ): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") {
        refuse("Cannot clear cells: table not found")
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
    }, "Cleared table cells")
  }

  const setTableColumnWidths = (tableIndex: number, widthsPx: number[]): void => {
    if (
      widthsPx.length === 0 ||
      widthsPx.some((width) => !Number.isFinite(width) || width <= 0)
    ) {
      unsupported("table column widths must be positive")
      return
    }

    const table = ctx.modelSnapshot.value.nodes[tableIndex]
    const columnCount = table?.type === "table" ? tableColumnCount(table) : 0
    if (table?.type !== "table" || columnCount !== widthsPx.length) {
      unsupported("table column width target is unavailable")
      return
    }

    const normalizedPx = widthsPx.map((width) => Math.max(1, Math.round(width)))
    const columnWidthsTwips = normalizedPx.map((width) => width * 15)
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") return current
      node.style = {
        ...(node.style ?? {}),
        layout: "fixed",
        widthTwips: columnWidthsTwips.reduce((sum, width) => sum + width, 0),
        columnWidthsTwips,
      }
      node.sourceXml = undefined
      return nextModel
    }, "Resized table columns")
  }

  const mergeTableCells = (
    tableIndex: number, range: DocxTableCellSelectionRange
  ): void => {
    let merged: TableCellPosition | undefined
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") {
        refuse("Cannot merge cells: table not found")
        return current
      }
      merged = mergeTableCellsOnGrid(node, range)
      if (!merged) {
        refuse("Cannot merge a single cell")
        return current
      }
      return nextModel
    }, "Merged table cells")

    if (merged) {
      selectMergedCell(tableIndex, merged)
    }
  }

  const splitTableCell = (
    tableIndex: number, rowIndex: number, cellIndex: number
  ): void => {
    let split: TableCellPosition | undefined
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") {
        refuse("Cannot split a cell: table not found")
        return current
      }
      split = splitTableCellOnGrid(node, rowIndex, cellIndex)
      if (!split) {
        refuse("Cannot split a cell that is not merged")
        return current
      }
      return nextModel
    }, "Split table cell")

    if (split) {
      selectMergedCell(tableIndex, split)
    }
  }

  const moveTable = (tableIndex: number, targetNodeIndex: number): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") {
        refuse("Cannot move the table: table not found")
        return current
      }
      if (!moveBodyNodeInModel(nextModel, tableIndex, targetNodeIndex)) {
        return current
      }
      return nextModel
    }, "Moved table")
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
      const nextModel = cloneDocModel(current)
      const parentTableIndex = Number(parts[0])
      const parentRowIndex = Number(parts[1])
      const parentCellIndex = Number(parts[2])
      const nestedTableIndex = parts[3] !== undefined ? Number(parts[3]) : undefined

      const parentTable = nextModel.nodes[parentTableIndex]
      if (!parentTable || parentTable.type !== "table") {
        refuse("Cannot move the embedded table: host table not found")
        return current
      }

      const parentCell = parentTable.rows[parentRowIndex]?.cells[parentCellIndex]
      if (!parentCell) {
        refuse("Cannot move the embedded table: host cell not found")
        return current
      }

      if (nestedTableIndex !== undefined && !isNaN(nestedTableIndex)) {
        const cellNodes = parentCell.nodes
        const nestedTable = cellNodes[nestedTableIndex]
        if (!nestedTable || nestedTable.type !== "table") {
          refuse("Cannot move the embedded table: table not found")
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
    }, "Moved embedded table")
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
    mergeTableCells,
    splitTableCell,
    moveTable,
    moveEmbeddedTableToBody,
  }
}
