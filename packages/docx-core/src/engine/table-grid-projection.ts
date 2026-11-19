// Collapsed table grid projection.
//
// A table can carry a *logical* grid that is wider than the grid Word lays out:
// when a merge collapses a grid column, the column keeps its zero width in
// `w:tblGrid` (so the span arithmetic still addresses the original columns) but
// no physical column is emitted for it. Rendering and serialization both need
// the *physical* view:
//
// - every cell span is expressed in physical columns, and
// - `columnWidthsTwips` only lists the columns that are actually drawn.
//
// `projectCollapsedTableGrid` derives that view without mutating the imported
// logical model, and rejects grids whose marker, widths or spans disagree
// instead of producing a package Word cannot open.
//
// Ported from upstream doc-model/src/table-grid-projection.ts (commit 6f70b92)
// against the local `TableNode`/`TableStyle` shape.

export interface GridCell {
  style?: object;
}

export interface GridRow {
  cells: GridCell[];
  style?: object;
}

export interface GridTable {
  rows: GridRow[];
  sourceXml?: string;
  style?: object;
}

interface GridCellStyle {
  gridSpan?: number;
}

interface GridRowStyle {
  gridBefore?: number;
  gridAfter?: number;
}

interface GridTableStyle {
  layout?: string;
  columnWidthsTwips?: number[];
  collapsedGridColumns?: number[];
  cellSpacingTwips?: number;
}

export interface GridProjectionOptions {
  /**
   * Effective border spacing after the consumer resolved inherited styles.
   * A collapsed grid only holds when that spacing is zero, so the default is
   * "not confirmed" and therefore rejected.
   */
  resolvedBorderSpacingPx?: number;
}

export interface LogicalBoundaryRange {
  start: number;
  end: number;
}

export type GridProjectionResult<T extends GridTable> =
  | { kind: "identity"; sourceTable: T; table: T }
  /**
   * A marked grid rejected here must not be handed to a physical consumer: its
   * spans and widths cannot be reconciled.
   */
  | { kind: "rejected"; sourceTable: T; reason: string }
  | {
      kind: "projected";
      sourceTable: T;
      table: T;
      /** Logical column boundary → physical column boundary. */
      logicalToPhysicalBoundary: number[];
      /** Physical column boundary → the logical columns that collapsed onto it. */
      physicalToLogicalBoundaryRanges: LogicalBoundaryRange[];
      /** Physical column → its logical column index. */
      physicalToLogicalColumn: number[];
      collapsedGridColumns: number[];
    };

/**
 * The physical view of a table whose grid collapsed zero-width columns.
 *
 * Returns the source table unchanged (`identity`) when it has no marker, and a
 * `rejected` result when the marker and the grid contradict each other.
 */
export function projectCollapsedTableGrid<T extends GridTable>(
  sourceTable: T,
  options: GridProjectionOptions = {}
): GridProjectionResult<T> {
  const tableStyle = sourceTable.style as GridTableStyle | undefined;
  const marker = tableStyle?.collapsedGridColumns;
  if (marker === undefined || (Array.isArray(marker) && marker.length === 0)) {
    return { kind: "identity", sourceTable, table: sourceTable };
  }
  const reject = (reason: string): GridProjectionResult<T> => ({
    kind: "rejected",
    sourceTable,
    reason
  });
  if (!Array.isArray(marker) || !tableStyle) {
    return reject("invalid-collapsed-marker");
  }
  if (tableStyle.layout !== "fixed") return reject("fixed-layout-required");
  if (options.resolvedBorderSpacingPx !== 0) {
    return reject("zero-resolved-spacing-not-confirmed");
  }
  if (
    tableStyle.cellSpacingTwips !== undefined &&
    tableStyle.cellSpacingTwips !== 0
  ) {
    return reject("nonzero-source-cell-spacing");
  }
  const widths = tableStyle.columnWidthsTwips;
  if (
    !Array.isArray(widths) ||
    widths.length === 0 ||
    widths.some((value) => !Number.isFinite(value) || value < 0)
  ) {
    return reject("invalid-logical-grid");
  }
  const collapsed = new Set<number>();
  for (const index of marker) {
    if (!Number.isInteger(index) || index < 0 || index >= widths.length) {
      return reject("invalid-collapsed-index");
    }
    if (collapsed.has(index)) return reject("duplicate-collapsed-index");
    if (widths[index] !== 0) return reject("marked-column-has-positive-width");
    collapsed.add(index);
  }
  const logicalToPhysicalBoundary = [0];
  const physicalToLogicalColumn: number[] = [];
  const physicalWidths: number[] = [];
  for (let index = 0; index < widths.length; index += 1) {
    if (!collapsed.has(index)) {
      if (widths[index] === 0) return reject("unmarked-zero-column");
      physicalToLogicalColumn.push(index);
      physicalWidths.push(widths[index]);
    }
    logicalToPhysicalBoundary.push(physicalWidths.length);
  }
  if (physicalWidths.length === 0) return reject("empty-physical-grid");
  const physicalToLogicalBoundaryRanges: LogicalBoundaryRange[] = [];
  for (let logical = 0; logical < logicalToPhysicalBoundary.length; logical += 1) {
    const physical = logicalToPhysicalBoundary[logical];
    const range = physicalToLogicalBoundaryRanges[physical];
    if (range) range.end = logical;
    else physicalToLogicalBoundaryRanges[physical] = { start: logical, end: logical };
  }
  const physicalSums = [0];
  for (const value of physicalWidths) {
    const total = physicalSums[physicalSums.length - 1] + value;
    if (!Number.isFinite(total)) return reject("invalid-physical-grid-width");
    physicalSums.push(total);
  }
  if (!(physicalSums[physicalSums.length - 1] > 0)) {
    return reject("invalid-physical-grid-width");
  }
  const rows: GridRow[] = [];
  for (const row of sourceTable.rows) {
    const rowStyle = row.style as GridRowStyle | undefined;
    if ((rowStyle?.gridBefore ?? 0) !== 0 || (rowStyle?.gridAfter ?? 0) !== 0) {
      return reject("row-grid-gaps-not-supported");
    }
    let logicalStart = 0;
    let changed = false;
    const cells: GridCell[] = [];
    for (const cell of row.cells) {
      const cellStyle = cell.style as GridCellStyle | undefined;
      const sourceSpan = cellStyle?.gridSpan ?? 1;
      if (!Number.isInteger(sourceSpan) || sourceSpan <= 0) {
        return reject("invalid-cell-span");
      }
      const logicalEnd = logicalStart + sourceSpan;
      if (logicalEnd > widths.length) return reject("cell-exceeds-logical-grid");
      const physicalStart = logicalToPhysicalBoundary[logicalStart];
      const physicalEnd = logicalToPhysicalBoundary[logicalEnd];
      const span = physicalEnd - physicalStart;
      if (span <= 0) return reject("cell-collapses-to-empty-span");
      if (!(physicalSums[physicalEnd] - physicalSums[physicalStart] > 0)) {
        return reject("cell-has-nonpositive-physical-width");
      }
      if (span === sourceSpan) {
        cells.push(cell);
      } else {
        changed = true;
        const style = { ...cellStyle } as GridCellStyle;
        if (span === 1) delete style.gridSpan;
        else style.gridSpan = span;
        cells.push({ ...cell, style });
      }
      logicalStart = logicalEnd;
    }
    rows.push(changed ? { ...row, cells } : row);
  }
  const { collapsedGridColumns: _marker, ...style } = tableStyle;
  const table = {
    ...sourceTable,
    style: { ...style, columnWidthsTwips: physicalWidths },
    rows
  } as unknown as T;
  return {
    kind: "projected",
    sourceTable,
    table,
    logicalToPhysicalBoundary,
    physicalToLogicalBoundaryRanges,
    physicalToLogicalColumn,
    collapsedGridColumns: [...collapsed].sort((left, right) => left - right)
  };
}

/** True when the table carries a non-empty collapsed grid marker. */
export function hasCollapsedTableGrid(table: GridTable): boolean {
  const style = table.style as GridTableStyle | undefined;
  const marker = style?.collapsedGridColumns;
  return Array.isArray(marker) && marker.length > 0;
}
