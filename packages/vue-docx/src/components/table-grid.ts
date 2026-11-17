// Grid materialization for DOCX tables.
//
// The document model keeps one cell per row, including `vMerge` continuation
// cells. HTML tables cannot render a continuation cell: the cell above must
// carry a `rowspan` that covers the continuation rows instead. Page slices
// make this harder — a render window can start in the middle of a vertical
// merge, in which case there is no visible restart cell to carry the span.
//
// This module converts the model rows into the exact cells that should be
// emitted for a render window. Continuation cells covered by a visible
// restart are dropped; continuation cells whose restart is outside the window
// degrade to ordinary cells so column alignment is preserved.

import type {
  TableCellNode,
  TableNode,
  TableRowNode,
} from "@arcships/docx-core"

export interface TableGridCell {
  cell: TableCellNode
  /** Row index in the source table. */
  rowIndex: number
  /** Cell index in the source row. */
  cellIndex: number
  /** First logical grid column covered by the cell. */
  columnIndex: number
  colSpan: number
  rowSpan: number
  /** True when a vertical merge continuation has no visible restart cell. */
  continuationFallback: boolean
}

export interface TableGridRow {
  row: TableRowNode
  /** Row index in the source table. */
  rowIndex: number
  cells: TableGridCell[]
}

export interface TableGridWindow {
  /** First row in the render window (inclusive). */
  startRowIndex?: number
  /** Last row in the render window (exclusive). */
  endRowIndex?: number
}

interface CellGridPosition {
  cellIndex: number
  columnIndex: number
  colSpan: number
}

function cellGridSpan(cell: TableCellNode): number {
  return cell.style?.gridSpan && cell.style.gridSpan > 1
    ? cell.style.gridSpan
    : 1
}

function gridPositions(row: TableRowNode): CellGridPosition[] {
  let columnCursor = 0
  return row.cells.map((cell, cellIndex) => {
    const colSpan = cellGridSpan(cell)
    const position = { cellIndex, columnIndex: columnCursor, colSpan }
    columnCursor += colSpan
    return position
  })
}

function clampRowIndex(value: number, fallback: number, rowCount: number): number {
  const normalized = Number.isFinite(value) ? Math.trunc(value) : fallback
  return Math.min(rowCount, Math.max(0, normalized))
}

function isContinuation(
  rows: TableRowNode[],
  rowIndex: number,
  cellIndex: number
): boolean {
  return Boolean(rows[rowIndex]?.cells[cellIndex]?.style?.vMergeContinuation)
}

function verticalMergeSpans(
  rows: TableRowNode[],
  positionsByRow: CellGridPosition[][]
): Map<string, number> {
  const spans = new Map<string, number>()

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    for (const position of positionsByRow[rowIndex]) {
      const cell = rows[rowIndex].cells[position.cellIndex]
      if (cell.style?.vMergeContinuation) continue

      let rowSpan = 1
      let nextRowIndex = rowIndex + 1
      while (nextRowIndex < rows.length) {
        const continuation = positionsByRow[nextRowIndex].find(
          (candidate) =>
            candidate.columnIndex === position.columnIndex &&
            isContinuation(
              rows,
              nextRowIndex,
              candidate.cellIndex
            )
        )
        if (!continuation) break
        rowSpan += 1
        nextRowIndex += 1
      }

      spans.set(`${rowIndex}:${position.cellIndex}`, rowSpan)
    }
  }

  return spans
}

function continuationHasVisibleRestart(
  rows: TableRowNode[],
  positionsByRow: CellGridPosition[][],
  spans: Map<string, number>,
  rowIndex: number,
  columnIndex: number,
  windowStart: number
): boolean {
  for (let candidateRow = rowIndex - 1; candidateRow >= windowStart; candidateRow -= 1) {
    const position = positionsByRow[candidateRow].find(
      (candidate) => candidate.columnIndex === columnIndex
    )
    if (!position) return false
    if (isContinuation(rows, candidateRow, position.cellIndex)) {
      continue
    }

    const rowSpan = spans.get(`${candidateRow}:${position.cellIndex}`) ?? 1
    return candidateRow + rowSpan > rowIndex
  }
  return false
}

export function resolveTableGrid(
  table: TableNode,
  window: TableGridWindow = {}
): TableGridRow[] {
  const rows = table.rows
  const rowCount = rows.length
  const startRowIndex = clampRowIndex(window.startRowIndex ?? 0, 0, rowCount)
  const endRowIndex = Math.max(
    startRowIndex,
    clampRowIndex(window.endRowIndex ?? rowCount, rowCount, rowCount)
  )
  const positionsByRow = rows.map(gridPositions)
  const spans = verticalMergeSpans(rows, positionsByRow)

  const gridRows: TableGridRow[] = []
  for (let rowIndex = startRowIndex; rowIndex < endRowIndex; rowIndex += 1) {
    const row = rows[rowIndex]
    const cells: TableGridCell[] = []

    for (const position of positionsByRow[rowIndex]) {
      const cell = row.cells[position.cellIndex]

      if (cell.style?.vMergeContinuation) {
        if (
          continuationHasVisibleRestart(
            rows,
            positionsByRow,
            spans,
            rowIndex,
            position.columnIndex,
            startRowIndex
          )
        ) {
          continue
        }
        cells.push({
          cell,
          rowIndex,
          cellIndex: position.cellIndex,
          columnIndex: position.columnIndex,
          colSpan: position.colSpan,
          rowSpan: 1,
          continuationFallback: true,
        })
        continue
      }

      cells.push({
        cell,
        rowIndex,
        cellIndex: position.cellIndex,
        columnIndex: position.columnIndex,
        colSpan: position.colSpan,
        rowSpan: Math.min(
          spans.get(`${rowIndex}:${position.cellIndex}`) ?? 1,
          endRowIndex - rowIndex
        ),
        continuationFallback: false,
      })
    }

    gridRows.push({ row, rowIndex, cells })
  }

  return gridRows
}
