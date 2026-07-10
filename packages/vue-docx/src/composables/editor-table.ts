// editor-table.ts — table insert/delete/move/resize operations
//
// Extracted from useDocxEditor.ts.

import type { DocModel } from "@extend-ai/docx-core"
import { cloneDocModel } from "@extend-ai/docx-core"
import type { EditorCore } from "./editor-shared"

export function createEditorTable(
  ctx: EditorCore,
  applyChange: (updater: (current: DocModel) => DocModel, successStatus?: string) => void,
) {
  const unsupported = (feature: string): void => {
    ctx.status.value = `Unsupported: ${feature}`
  }

  const insertTable = (): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const rows = 3
      const cols = 3
      const table = {
        type: "table" as const,
        rows: Array.from({ length: rows }, () => ({
          type: "table-row" as const,
          cells: Array.from({ length: cols }, () => ({
            type: "table-cell" as const,
            nodes: [{ type: "paragraph" as const, children: [] as any[] }],
          })),
        })),
      }
      nextModel.nodes.push(table as any)
      return nextModel
    }, "Inserted table")
  }

  const insertTableRow = (
    tableIndex: number, rowIndex: number, direction: "above" | "below"
  ): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") return current
      const insertIdx = direction === "above" ? rowIndex : rowIndex + 1
      const cols = node.rows[0]?.cells.length ?? 1
      const newRow = {
        type: "table-row" as const,
        cells: Array.from({ length: cols }, () => ({
          type: "table-cell" as const,
          nodes: [{ type: "paragraph" as const, children: [] as any[] }],
        })),
      }
      node.rows.splice(insertIdx, 0, newRow)
      node.sourceXml = undefined
      return nextModel
    }, `Inserted table row ${direction}`)
  }

  const insertTableColumn = (
    tableIndex: number, cellIndex: number, direction: "left" | "right", _rowIndex?: number
  ): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") return current
      const insertIdx = direction === "left" ? cellIndex : cellIndex + 1
      for (const row of node.rows) {
        row.cells.splice(insertIdx, 0, {
          type: "table-cell" as const,
          nodes: [{ type: "paragraph" as const, children: [] as any[] }],
        })
      }
      node.sourceXml = undefined
      return nextModel
    }, `Inserted table column ${direction}`)
  }

  const deleteTableRow = (tableIndex: number, rowIndex: number): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table" || node.rows.length <= 1) return current
      node.rows.splice(rowIndex, 1)
      node.sourceXml = undefined
      return nextModel
    }, "Deleted table row")
  }

  const deleteTableColumn = (
    tableIndex: number, cellIndex: number, _rowIndex?: number
  ): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") return current
      if (node.rows[0]?.cells.length <= 1) return current
      for (const row of node.rows) {
        row.cells.splice(cellIndex, 1)
      }
      node.sourceXml = undefined
      return nextModel
    }, "Deleted table column")
  }

  const deleteTable = (tableIndex: number): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") return current
      nextModel.nodes.splice(tableIndex, 1)
      return nextModel
    }, "Deleted table")
  }

  const clearTableCellContents = (
    tableIndex: number, cells: Array<{ rowIndex: number; cellIndex: number }>
  ): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") return current
      for (const { rowIndex, cellIndex } of cells) {
        const cell = node.rows[rowIndex]?.cells[cellIndex]
        if (cell) {
          cell.nodes = [{ type: "paragraph", children: [] }]
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
    const columnCount = table?.type === "table" ? table.rows[0]?.cells.length ?? 0 : 0
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

  const moveTable = (tableIndex: number, targetNodeIndex: number): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") return current
      nextModel.nodes.splice(tableIndex, 1)
      const adjustedTarget =
        targetNodeIndex > tableIndex ? targetNodeIndex - 1 : targetNodeIndex
      nextModel.nodes.splice(adjustedTarget, 0, node)
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
      if (!parentTable || parentTable.type !== "table") return current

      const parentCell = parentTable.rows[parentRowIndex]?.cells[parentCellIndex]
      if (!parentCell) return current

      if (nestedTableIndex !== undefined && !isNaN(nestedTableIndex)) {
        const cellNodes = parentCell.nodes
        const nestedTable = cellNodes[nestedTableIndex]
        if (!nestedTable || nestedTable.type !== "table") return current

        // Extract nested table from cell
        cellNodes.splice(nestedTableIndex, 1)

        // Insert into body at target position
        const adjustedTarget = Math.min(targetNodeIndex, nextModel.nodes.length)
        nextModel.nodes.splice(adjustedTarget, 0, nestedTable)
        parentTable.sourceXml = undefined
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
    moveTable,
    moveEmbeddedTableToBody,
  }
}
