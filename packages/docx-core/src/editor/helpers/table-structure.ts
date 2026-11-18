// Table grid structure operations.
//
// The editor's table commands are driven by a logical grid: a cell may cover
// several grid columns (`gridSpan`) and several grid rows (`rowSpan`, followed
// by `vMergeContinuation` cells in the rows below). These helpers keep row and
// column edits on that grid instead of treating every row as a flat array.

import type {
  DocModel,
  DocNode,
  DocumentSection,
  TableCellContentNode,
  TableCellNode,
  TableNode,
  TableRowNode,
} from "../../engine/types";
import type { DocxTableCellSelectionRange } from "./editor-types";
import {
  cloneTableRowStyle,
  createEmptyParagraphFromTemplate,
  createEmptyTableCellFromTemplate,
} from "./numbering";
import {
  computeColumnWidthsFromTableDefinition,
  tableColumnCount,
} from "./table-utils";

/** A cell together with the logical grid position it starts at. */
export interface TableGridCellRef {
  cell: TableCellNode;
  cellIndex: number;
  startColumn: number;
  span: number;
}

export interface TableCellPosition {
  rowIndex: number;
  cellIndex: number;
}

const DEFAULT_COLUMN_WIDTH_TWIPS = 1440;

export function tableRowGridCells(row: TableRowNode): TableGridCellRef[] {
  let column = 0;
  return row.cells.map((cell, cellIndex) => {
    const span =
      cell.style?.gridSpan && cell.style.gridSpan > 1
        ? Math.max(1, Math.round(cell.style.gridSpan))
        : 1;
    const ref = { cell, cellIndex, startColumn: column, span };
    column += span;
    return ref;
  });
}

export function tableCellGridRef(
  table: TableNode,
  rowIndex: number,
  cellIndex: number
): TableGridCellRef | undefined {
  const row = table.rows[rowIndex];
  if (!row) {
    return undefined;
  }
  return tableRowGridCells(row)[cellIndex];
}

function cellRefAtColumn(
  table: TableNode,
  rowIndex: number,
  column: number
): TableGridCellRef | undefined {
  const row = table.rows[rowIndex];
  if (!row) {
    return undefined;
  }
  return tableRowGridCells(row).find(
    (ref) => column >= ref.startColumn && column < ref.startColumn + ref.span
  );
}

function clamp(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, value));
}

function resolvedColumnWidths(
  table: TableNode,
  columnCount: number
): number[] {
  // A collapsed grid keeps its zero-width logical columns: rewriting them to a
  // default width would silently materialize a grid whose marker still names
  // them, and export would then reject the table as contradictory.
  const collapsed = new Set(
    (table.style?.collapsedGridColumns ?? []).filter(
      (index) => Number.isInteger(index) && index >= 0 && index < columnCount
    )
  );
  const normalize = (width: number, index: number, fallback: number): number => {
    if (collapsed.has(index)) return 0;
    return Number.isFinite(width) && width > 0 ? Math.round(width) : fallback;
  };

  const explicit = table.style?.columnWidthsTwips;
  if (explicit && explicit.length === columnCount) {
    return explicit.map((width, index) =>
      normalize(width, index, DEFAULT_COLUMN_WIDTH_TWIPS)
    );
  }

  const derived = computeColumnWidthsFromTableDefinition(table, columnCount);
  if (derived && derived.length === columnCount) {
    return derived.map((width, index) =>
      normalize(
        width,
        index,
        Math.max(
          1,
          Math.round(
            (table.style?.widthTwips ?? DEFAULT_COLUMN_WIDTH_TWIPS * columnCount) /
              columnCount
          )
        )
      )
    );
  }

  const total = Math.max(1, table.style?.widthTwips ?? DEFAULT_COLUMN_WIDTH_TWIPS * columnCount);
  const fallback = Math.max(1, Math.round(total / Math.max(1, columnCount)));
  return Array.from({ length: columnCount }, (_, index) =>
    collapsed.has(index) ? 0 : fallback
  );
}

function applyColumnWidths(table: TableNode, widths: number[]): void {
  table.style = {
    ...(table.style ?? {}),
    columnWidthsTwips: widths,
    widthTwips: widths.reduce((sum, width) => sum + width, 0),
  };
}

/**
 * Keep `collapsedGridColumns` pointing at the same zero-width columns after a
 * physical grid edit. A stale marker makes export reject the table, because the
 * marked index would name a column with a positive width.
 */
function remapCollapsedGridColumns(
  table: TableNode,
  remap: (index: number) => number | undefined
): void {
  const marker = table.style?.collapsedGridColumns;
  if (!marker || marker.length === 0) return;
  const next = marker
    .map((index) => remap(index))
    .filter((index): index is number => index !== undefined);
  table.style = { ...(table.style ?? {}), collapsedGridColumns: next };
}

function averageInsertedColumnWidth(widths: number[], boundary: number): number {
  const left = widths[boundary - 1];
  const right = widths[boundary];
  if (Number.isFinite(left) && Number.isFinite(right)) {
    return Math.max(1, Math.round((left + right) / 2));
  }
  if (Number.isFinite(left)) {
    return Math.max(1, Math.round(left));
  }
  if (Number.isFinite(right)) {
    return Math.max(1, Math.round(right));
  }
  return DEFAULT_COLUMN_WIDTH_TWIPS;
}

function findVerticalMergeRestart(
  table: TableNode,
  rowIndex: number,
  column: number
): { ref: TableGridCellRef; rowIndex: number } | undefined {
  for (let index = rowIndex; index >= 0; index -= 1) {
    const ref = cellRefAtColumn(table, index, column);
    if (!ref) {
      continue;
    }
    if (!ref.cell.style?.vMergeContinuation) {
      return { ref, rowIndex: index };
    }
  }
  return undefined;
}

function verticalMergeHeightAt(
  table: TableNode,
  restartRowIndex: number,
  column: number,
  cell: TableCellNode
): number {
  const declaredSpan = cell.style?.rowSpan;
  if (Number.isFinite(declaredSpan) && (declaredSpan as number) > 1) {
    return Math.max(1, Math.round(declaredSpan as number));
  }

  let height = 1;
  for (let index = restartRowIndex + 1; index < table.rows.length; index += 1) {
    const ref = cellRefAtColumn(table, index, column);
    if (!ref?.cell.style?.vMergeContinuation) {
      break;
    }
    height += 1;
  }
  return height;
}

/**
 * Returns the restart cell of a vertical merge that crosses the horizontal
 * boundary between `boundary - 1` and `boundary` at `column`.
 */
function verticalMergeAtBoundary(
  table: TableNode,
  boundary: number,
  column: number
): { ref: TableGridCellRef; rowIndex: number } | undefined {
  const below = cellRefAtColumn(table, boundary, column);
  if (below?.cell.style?.vMergeContinuation) {
    const restart = findVerticalMergeRestart(table, boundary - 1, column);
    if (restart) {
      return restart;
    }
  }

  const above = cellRefAtColumn(table, boundary - 1, column);
  if (!above) {
    return undefined;
  }
  if (above.cell.style?.vMergeContinuation) {
    const restart = findVerticalMergeRestart(table, boundary - 1, column);
    if (!restart) {
      return undefined;
    }
    const height = verticalMergeHeightAt(
      table,
      restart.rowIndex,
      column,
      restart.ref.cell
    );
    return restart.rowIndex + height > boundary ? restart : undefined;
  }

  const height = verticalMergeHeightAt(
    table,
    boundary - 1,
    column,
    above.cell
  );
  return boundary - 1 + height > boundary ? { ref: above, rowIndex: boundary - 1 } : undefined;
}

function extendVerticalMerge(restart: { ref: TableGridCellRef; rowIndex: number }, column: number, table: TableNode): void {
  const rowSpan = restart.ref.cell.style?.rowSpan;
  const nextSpan =
    Number.isFinite(rowSpan) && (rowSpan as number) > 1
      ? Math.max(1, Math.round(rowSpan as number)) + 1
      : verticalMergeHeightAt(table, restart.rowIndex, column, restart.ref.cell) + 1;
  restart.ref.cell.style = {
    ...(restart.ref.cell.style ?? {}),
    rowSpan: nextSpan,
  };
}

export function insertTableRow(
  table: TableNode,
  rowIndex: number,
  direction: "above" | "below"
): boolean {
  if (table.rows.length === 0) {
    return false;
  }

  const sourceRowIndex = clamp(rowIndex, 0, table.rows.length - 1);
  const insertIndex = clamp(
    direction === "above" ? rowIndex : rowIndex + 1,
    0,
    table.rows.length
  );
  const columnCount = Math.max(1, tableColumnCount(table));
  const widths = resolvedColumnWidths(table, columnCount);
  const templateRow = table.rows[sourceRowIndex];
  const templateCells = tableRowGridCells(templateRow);

  const row: TableRowNode = {
    type: "table-row",
    style: cloneTableRowStyle(templateRow.style),
    cells: [],
  };
  if (row.style) {
    row.style.isHeader = undefined;
  }

  const extendedRestarts = new Set<TableCellNode>();
  const buildCell = (
    templateCell: TableCellNode | undefined,
    span: number,
    startColumn: number,
    merge: { ref: TableGridCellRef; rowIndex: number } | undefined
  ): TableCellNode => {
    const cell = createEmptyTableCellFromTemplate(templateCell);
    cell.style = {
      ...(cell.style ?? {}),
      widthTwips: widths
        .slice(startColumn, startColumn + span)
        .reduce((sum, value) => sum + value, 0),
      gridSpan: span > 1 ? span : undefined,
      rowSpan: undefined,
      vMergeContinuation: merge ? true : undefined,
    };
    if (merge && !extendedRestarts.has(merge.ref.cell)) {
      extendVerticalMerge(merge, startColumn, table);
      extendedRestarts.add(merge.ref.cell);
    }
    return cell;
  };

  let column = 0;
  let lastTemplate: TableCellNode | undefined;
  for (const template of templateCells) {
    if (column >= columnCount) {
      break;
    }
    const span = Math.min(template.span, columnCount - column);
    lastTemplate = template.cell;
    const merges = Array.from({ length: span }, (_, offset) =>
      verticalMergeAtBoundary(table, insertIndex, column + offset)
    );
    const firstMerge = merges[0];
    const uniformMerge = merges.every(
      (merge) =>
        merge === firstMerge ||
        (merge !== undefined &&
          firstMerge !== undefined &&
          merge.ref.cell === firstMerge.ref.cell)
    );

    if (uniformMerge) {
      row.cells.push(buildCell(template.cell, span, column, firstMerge));
    } else {
      merges.forEach((merge, offset) => {
        row.cells.push(buildCell(template.cell, 1, column + offset, merge));
      });
    }
    column += span;
  }

  while (column < columnCount) {
    row.cells.push(
      buildCell(
        lastTemplate,
        1,
        column,
        verticalMergeAtBoundary(table, insertIndex, column)
      )
    );
    column += 1;
  }

  table.rows.splice(insertIndex, 0, row);
  table.sourceXml = undefined;
  return true;
}

export function insertTableColumn(
  table: TableNode,
  cellIndex: number,
  direction: "left" | "right",
  rowIndex?: number
): boolean {
  const columnCount = Math.max(1, tableColumnCount(table));
  const safeRowIndex = clamp(
    rowIndex ?? 0,
    0,
    Math.max(0, table.rows.length - 1)
  );
  const templateRef = tableCellGridRef(table, safeRowIndex, cellIndex);
  if (!templateRef) {
    return false;
  }

  const boundary = clamp(
    direction === "left"
      ? templateRef.startColumn
      : templateRef.startColumn + templateRef.span,
    0,
    columnCount
  );
  const widths = resolvedColumnWidths(table, columnCount);
  const insertedWidth = averageInsertedColumnWidth(widths, boundary);
  const nextWidths = [...widths];
  nextWidths.splice(boundary, 0, insertedWidth);

  for (const row of table.rows) {
    const refs = tableRowGridCells(row);
    const covering = refs.find(
      (ref) => boundary > ref.startColumn && boundary < ref.startColumn + ref.span
    );
    if (covering) {
      const nextSpan = covering.span + 1;
      covering.cell.style = {
        ...(covering.cell.style ?? {}),
        gridSpan: nextSpan > 1 ? nextSpan : undefined,
        widthTwips: (covering.cell.style?.widthTwips ?? 0) + insertedWidth,
      };
      continue;
    }

    const before = refs.find((ref) => ref.startColumn >= boundary);
    const insertAt = before ? before.cellIndex : row.cells.length;
    const template = (before
      ? refs[before.cellIndex - 1]
      : refs[refs.length - 1])?.cell;
    const cell = createEmptyTableCellFromTemplate(template);
    cell.style = {
      ...(cell.style ?? {}),
      widthTwips: insertedWidth,
      gridSpan: undefined,
      rowSpan: undefined,
      vMergeContinuation: undefined,
    };
    row.cells.splice(insertAt, 0, cell);
  }

  applyColumnWidths(table, nextWidths);
  remapCollapsedGridColumns(table, (index) =>
    index >= boundary ? index + 1 : index
  );
  table.sourceXml = undefined;
  return true;
}

function clearVerticalMergeForDeletedRow(
  table: TableNode,
  deletedRowIndex: number
): void {
  const deletedRow = table.rows[deletedRowIndex];
  if (!deletedRow) {
    return;
  }

  const refs = tableRowGridCells(deletedRow);
  const clearedContinuations = new Set<TableCellNode>();
  for (const ref of refs) {
    const spanRows = Math.max(0, Math.round(ref.cell.style?.rowSpan ?? 0));
    if (spanRows <= 1) {
      continue;
    }

    for (
      let rowIndex = deletedRowIndex + 1;
      rowIndex < deletedRowIndex + spanRows && rowIndex < table.rows.length;
      rowIndex += 1
    ) {
      for (
        let column = ref.startColumn;
        column < ref.startColumn + ref.span;
        column += 1
      ) {
        const continuation = cellRefAtColumn(table, rowIndex, column);
        if (
          continuation?.cell.style?.vMergeContinuation &&
          !clearedContinuations.has(continuation.cell)
        ) {
          continuation.cell.style = {
            ...(continuation.cell.style ?? {}),
            vMergeContinuation: undefined,
            rowSpan: undefined,
          };
          clearedContinuations.add(continuation.cell);
        }
      }
    }
  }

  const decrementedRestarts = new Set<TableCellNode>();
  for (const ref of refs) {
    for (let column = ref.startColumn; column < ref.startColumn + ref.span; column += 1) {
      const restart = findVerticalMergeRestart(
        table,
        deletedRowIndex - 1,
        column
      );
      if (!restart || decrementedRestarts.has(restart.ref.cell)) {
        continue;
      }
      const span = restart.ref.cell.style?.rowSpan;
      if (!Number.isFinite(span) || (span as number) <= 1) {
        continue;
      }
      if (restart.rowIndex + (span as number) <= deletedRowIndex) {
        continue;
      }
      const nextSpan = Math.round(span as number) - 1;
      restart.ref.cell.style = {
        ...(restart.ref.cell.style ?? {}),
        rowSpan: nextSpan > 1 ? nextSpan : undefined,
      };
      decrementedRestarts.add(restart.ref.cell);
    }
  }
}

export function deleteTableRow(table: TableNode, rowIndex: number): boolean {
  if (table.rows.length <= 1) {
    return false;
  }

  const safeRowIndex = clamp(rowIndex, 0, table.rows.length - 1);
  clearVerticalMergeForDeletedRow(table, safeRowIndex);
  table.rows.splice(safeRowIndex, 1);
  table.sourceXml = undefined;
  return true;
}

export function deleteTableColumn(
  table: TableNode,
  cellIndex: number,
  rowIndex?: number
): boolean {
  const columnCount = Math.max(1, tableColumnCount(table));
  if (columnCount <= 1) {
    return false;
  }

  const safeRowIndex = clamp(
    rowIndex ?? 0,
    0,
    Math.max(0, table.rows.length - 1)
  );
  const targetRef = tableCellGridRef(table, safeRowIndex, cellIndex);
  if (!targetRef) {
    return false;
  }

  const targetColumn = targetRef.startColumn;
  const widths = resolvedColumnWidths(table, columnCount);
  const removedWidth = widths[targetColumn] ?? DEFAULT_COLUMN_WIDTH_TWIPS;

  for (const row of table.rows) {
    const refs = tableRowGridCells(row);
    const covering = refs.find(
      (ref) =>
        targetColumn >= ref.startColumn &&
        targetColumn < ref.startColumn + ref.span
    );
    if (!covering) {
      continue;
    }

    if (covering.span > 1) {
      const nextSpan = covering.span - 1;
      covering.cell.style = {
        ...(covering.cell.style ?? {}),
        gridSpan: nextSpan > 1 ? nextSpan : undefined,
        widthTwips: Math.max(
          1,
          (covering.cell.style?.widthTwips ?? removedWidth) - removedWidth
        ),
      };
      continue;
    }

    row.cells.splice(covering.cellIndex, 1);
    if (row.cells.length === 0) {
      const cell = createEmptyTableCellFromTemplate();
      cell.style = {
        ...(cell.style ?? {}),
        widthTwips: Math.max(1, widths[targetColumn - 1] ?? DEFAULT_COLUMN_WIDTH_TWIPS),
      };
      row.cells.push(cell);
    }
  }

  const nextWidths = [...widths];
  nextWidths.splice(targetColumn, 1);
  applyColumnWidths(table, nextWidths);
  remapCollapsedGridColumns(table, (index) =>
    index === targetColumn ? undefined : index > targetColumn ? index - 1 : index
  );
  table.sourceXml = undefined;
  return true;
}

interface GridRectangle {
  minRow: number;
  maxRow: number;
  minColumn: number;
  maxColumn: number;
}

function expandRectangleForMerges(
  table: TableNode,
  rectangle: GridRectangle
): GridRectangle {
  let next = { ...rectangle };
  let changed = true;
  let guard = 0;

  while (changed && guard < 64) {
    changed = false;
    guard += 1;
    for (let rowIndex = next.minRow; rowIndex <= next.maxRow; rowIndex += 1) {
      const row = table.rows[rowIndex];
      if (!row) {
        continue;
      }
      for (const ref of tableRowGridCells(row)) {
        const refEnd = ref.startColumn + ref.span - 1;
        if (refEnd < next.minColumn || ref.startColumn > next.maxColumn) {
          continue;
        }

        if (ref.startColumn < next.minColumn) {
          next.minColumn = ref.startColumn;
          changed = true;
        }
        if (refEnd > next.maxColumn) {
          next.maxColumn = refEnd;
          changed = true;
        }

        const restart = ref.cell.style?.vMergeContinuation
          ? findVerticalMergeRestart(table, rowIndex - 1, ref.startColumn)
          : { ref, rowIndex };
        if (!restart) {
          continue;
        }
        const height = verticalMergeHeightAt(
          table,
          restart.rowIndex,
          ref.startColumn,
          restart.ref.cell
        );
        const mergeEnd = restart.rowIndex + height - 1;
        if (restart.rowIndex < next.minRow) {
          next.minRow = restart.rowIndex;
          changed = true;
        }
        if (mergeEnd > next.maxRow) {
          next.maxRow = mergeEnd;
          changed = true;
        }
      }
    }
  }

  return {
    minRow: clamp(next.minRow, 0, Math.max(0, table.rows.length - 1)),
    maxRow: clamp(next.maxRow, 0, Math.max(0, table.rows.length - 1)),
    minColumn: Math.max(0, next.minColumn),
    maxColumn: Math.max(0, next.maxColumn),
  };
}

function cellNodesForMerge(nodes: TableCellContentNode[]): TableCellContentNode[] {
  return nodes.length > 0 ? nodes : [createEmptyParagraphFromTemplate()];
}

export function mergeTableCells(
  table: TableNode,
  range: DocxTableCellSelectionRange
): TableCellPosition | undefined {
  const anchor = tableCellGridRef(
    table,
    range.anchorRowIndex,
    range.anchorCellIndex
  );
  const focus = tableCellGridRef(
    table,
    range.focusRowIndex,
    range.focusCellIndex
  );
  if (!anchor || !focus) {
    return undefined;
  }

  const rectangle = expandRectangleForMerges(table, {
    minRow: Math.min(range.anchorRowIndex, range.focusRowIndex),
    maxRow: Math.max(range.anchorRowIndex, range.focusRowIndex),
    minColumn: Math.min(anchor.startColumn, focus.startColumn),
    maxColumn:
      Math.max(
        anchor.startColumn + anchor.span,
        focus.startColumn + focus.span
      ) - 1,
  });

  const width = rectangle.maxColumn - rectangle.minColumn + 1;
  const height = rectangle.maxRow - rectangle.minRow + 1;
  if (width <= 1 && height <= 1) {
    return undefined;
  }

  const columnCount = Math.max(1, tableColumnCount(table));
  const widths = resolvedColumnWidths(table, columnCount);
  const anchorCell =
    cellRefAtColumn(table, rectangle.minRow, rectangle.minColumn)?.cell;
  if (!anchorCell) {
    return undefined;
  }

  let mergedPosition: TableCellPosition | undefined;
  let mergedContent: TableCellContentNode[] = [];

  for (
    let rowIndex = rectangle.minRow;
    rowIndex <= rectangle.maxRow;
    rowIndex += 1
  ) {
    const row = table.rows[rowIndex];
    if (!row) {
      continue;
    }
    const affected = tableRowGridCells(row).filter((ref) => {
      const refEnd = ref.startColumn + ref.span - 1;
      return refEnd >= rectangle.minColumn && ref.startColumn <= rectangle.maxColumn;
    });
    if (affected.length === 0) {
      continue;
    }
    if (rowIndex === rectangle.minRow) {
      mergedPosition = { rowIndex, cellIndex: affected[0].cellIndex };
    }
    mergedContent = mergedContent.concat(
      affected.flatMap((ref) => cloneCellContentForMerge(ref.cell.nodes))
    );
  }

  for (
    let rowIndex = rectangle.minRow;
    rowIndex <= rectangle.maxRow;
    rowIndex += 1
  ) {
    const row = table.rows[rowIndex];
    if (!row) {
      continue;
    }
    const refs = tableRowGridCells(row);
    const affected = refs.filter((ref) => {
      const refEnd = ref.startColumn + ref.span - 1;
      return refEnd >= rectangle.minColumn && ref.startColumn <= rectangle.maxColumn;
    });
    if (affected.length === 0) {
      continue;
    }

    for (let index = affected.length - 1; index >= 0; index -= 1) {
      row.cells.splice(affected[index].cellIndex, 1);
    }

    const cell = createEmptyTableCellFromTemplate(anchorCell);
    cell.style = {
      ...(cell.style ?? {}),
      gridSpan: width > 1 ? width : undefined,
      rowSpan: height > 1 ? height : undefined,
      vMergeContinuation: rowIndex === rectangle.minRow ? undefined : true,
      widthTwips: widths
        .slice(rectangle.minColumn, rectangle.maxColumn + 1)
        .reduce((sum, value) => sum + value, 0),
    };
    cell.nodes =
      rowIndex === rectangle.minRow
        ? cellNodesForMerge(mergedContent)
        : [createEmptyParagraphFromTemplate()];
    row.cells.splice(affected[0].cellIndex, 0, cell);
  }

  table.sourceXml = undefined;
  return mergedPosition;
}

function cloneCellContentForMerge(
  nodes: TableCellContentNode[]
): TableCellContentNode[] {
  return nodes.map((node) =>
    node.type === "paragraph"
      ? {
          type: "paragraph",
          style: node.style ? { ...node.style } : undefined,
          sourceXml: node.sourceXml,
          children: node.children.map((child) => ({ ...child })),
        }
      : node
  );
}

export function splitTableCell(
  table: TableNode,
  rowIndex: number,
  cellIndex: number
): TableCellPosition | undefined {
  let ref = tableCellGridRef(table, rowIndex, cellIndex);
  if (!ref) {
    return undefined;
  }

  let restartRowIndex = rowIndex;
  if (ref.cell.style?.vMergeContinuation) {
    const restart = findVerticalMergeRestart(table, rowIndex - 1, ref.startColumn);
    if (!restart) {
      return undefined;
    }
    ref = restart.ref;
    restartRowIndex = restart.rowIndex;
  }

  const width = ref.span;
  const height = verticalMergeHeightAt(
    table,
    restartRowIndex,
    ref.startColumn,
    ref.cell
  );
  if (width <= 1 && height <= 1) {
    return undefined;
  }

  const columnCount = Math.max(1, tableColumnCount(table));
  const widths = resolvedColumnWidths(table, columnCount);
  const startColumn = ref.startColumn;
  const endColumn = startColumn + width - 1;
  const anchorCell = ref.cell;
  const anchorContent = cloneCellContentForMerge(anchorCell.nodes);
  let position: TableCellPosition | undefined;

  for (
    let currentRowIndex = restartRowIndex;
    currentRowIndex < restartRowIndex + height &&
    currentRowIndex < table.rows.length;
    currentRowIndex += 1
  ) {
    const row = table.rows[currentRowIndex];
    if (!row) {
      continue;
    }
    const refs = tableRowGridCells(row);
    const affected = refs.filter((cellRef) => {
      const cellEnd = cellRef.startColumn + cellRef.span - 1;
      return cellEnd >= startColumn && cellRef.startColumn <= endColumn;
    });
    if (affected.length === 0) {
      continue;
    }

    const insertAt = affected[0].cellIndex;
    if (currentRowIndex === restartRowIndex) {
      position = { rowIndex: currentRowIndex, cellIndex: insertAt };
    }
    for (let index = affected.length - 1; index >= 0; index -= 1) {
      row.cells.splice(affected[index].cellIndex, 1);
    }

    const newCells: TableCellNode[] = [];
    for (let column = startColumn; column <= endColumn; column += 1) {
      const cell = createEmptyTableCellFromTemplate(anchorCell);
      cell.style = {
        ...(cell.style ?? {}),
        gridSpan: undefined,
        rowSpan: undefined,
        vMergeContinuation: undefined,
        widthTwips: widths[column] ?? DEFAULT_COLUMN_WIDTH_TWIPS,
      };
      cell.nodes =
        currentRowIndex === restartRowIndex && column === startColumn
          ? anchorContent
          : [createEmptyParagraphFromTemplate()];
      newCells.push(cell);
    }
    row.cells.splice(insertAt, 0, ...newCells);
  }

  table.sourceXml = undefined;
  return position;
}

// ── Body-level section bookkeeping ────────────────────────────────

function sectionStartIndexes(model: DocModel): DocumentSection[] {
  return model.metadata.sections ?? [];
}

export function shiftSectionStartIndexesForInsertion(
  model: DocModel,
  insertIndex: number
): void {
  for (const section of sectionStartIndexes(model)) {
    const start = section.startNodeIndex;
    if (start > 0 && start >= insertIndex) {
      section.startNodeIndex = start + 1;
    }
  }
}

export function shiftSectionStartIndexesForRemoval(
  model: DocModel,
  removeIndex: number
): void {
  const maxIndex = Math.max(0, model.nodes.length - 1);
  for (const section of sectionStartIndexes(model)) {
    const start = section.startNodeIndex;
    if (start > removeIndex) {
      section.startNodeIndex = Math.min(maxIndex, start - 1);
      continue;
    }
    if (start < 0 || start > maxIndex) {
      section.startNodeIndex = Math.min(maxIndex, Math.max(0, start));
    }
  }
}

/**
 * Reorders a body node while keeping every existing section boundary attached
 * to the same node. A node that crosses a section break therefore joins the
 * section at its new position instead of silently redefining the break.
 */
export function moveBodyNodeInModel(
  model: DocModel,
  fromIndex: number,
  targetNodeIndex: number
): boolean {
  const nodes = model.nodes;
  if (
    fromIndex < 0 ||
    fromIndex >= nodes.length ||
    !Number.isFinite(targetNodeIndex)
  ) {
    return false;
  }

  const adjustedTarget = clamp(
    targetNodeIndex > fromIndex ? targetNodeIndex - 1 : targetNodeIndex,
    0,
    nodes.length - 1
  );
  if (adjustedTarget === fromIndex) {
    return false;
  }

  const movedNode = nodes[fromIndex];
  const anchors = sectionStartIndexes(model).map((section) => {
    const start = clamp(section.startNodeIndex, 0, nodes.length - 1);
    if (start === 0) {
      return { section, anchor: undefined as DocNode | undefined, keepZero: true };
    }
    if (start === fromIndex) {
      const successor = nodes.slice(fromIndex + 1).find((node) => node !== movedNode);
      return { section, anchor: successor, keepZero: false };
    }
    return { section, anchor: nodes[start], keepZero: false };
  });

  nodes.splice(fromIndex, 1);
  nodes.splice(adjustedTarget, 0, movedNode);

  for (const { section, anchor, keepZero } of anchors) {
    if (keepZero) {
      section.startNodeIndex = 0;
      continue;
    }
    if (anchor) {
      const index = nodes.indexOf(anchor);
      section.startNodeIndex =
        index >= 0 ? index : Math.min(section.startNodeIndex, nodes.length - 1);
    } else {
      section.startNodeIndex = Math.max(
        0,
        Math.min(section.startNodeIndex, nodes.length - 1)
      );
    }
  }

  return true;
}
