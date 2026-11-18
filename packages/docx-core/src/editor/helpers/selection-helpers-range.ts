// Selection / caret range normalization, cloning, equality, and resolution
// helpers. Range-to-model normalization keeps stale selection/range state
// valid after model mutations. Split from selection-helpers.ts to keep each
// file under the 1000-line hard constraint.
// Upstream editor.tsx: lines 23949-24145, 24261-24639.

import type {
  DocModel,
  ParagraphNode,
  TableCellNode,
  TableNode,
} from "../../engine/types";
import type {
  DocxEditorSelection,
  DocxTableCellSelectionRange,
  DocxTextRange,
  DocxTextRangeBoundary,
  DocxTextRangeLocation,
  DocxImageLocation,
  ParagraphLocation
} from "./editor-types";
import {
  compareTextRangeBoundaries,
  normalizeTextRange
} from "./editor-types";
import { paragraphText, tableCellParagraphs } from "./paragraph-inspect";
import {
  tableRowGridCells,
  tableSelectionGridRectangle,
} from "./table-structure";
import { clampNumber } from "./zoom-utils";
import {
  cloneTextRangeLocation,
  firstParagraphLocationFromNode,
  firstParagraphLocationInDocument,
  firstParagraphLocationInTable,
  getParagraphAtLocation,
  lastParagraphLocationInNode,
  normalizeRangeBoundaryParagraphOffset,
  sameParagraphLocation,
} from "./selection-helpers";

// ── Paragraph location (last-in-document) ───────────────────────────

export function lastParagraphLocationInDocument(
  model: DocModel
): ParagraphLocation | undefined {
  for (let nodeIndex = model.nodes.length - 1; nodeIndex >= 0; nodeIndex -= 1) {
    const node = model.nodes[nodeIndex];
    if (!node) {
      continue;
    }

    if (node.type === "paragraph") {
      return {
        kind: "paragraph",
        nodeIndex,
      };
    }

    for (let rowIndex = node.rows.length - 1; rowIndex >= 0; rowIndex -= 1) {
      const row = node.rows[rowIndex];
      if (!row) {
        continue;
      }

      for (
        let cellIndex = row.cells.length - 1;
        cellIndex >= 0;
        cellIndex -= 1
      ) {
        const cell = row.cells[cellIndex];
        if (!cell || cell.style?.vMergeContinuation) {
          continue;
        }

        const paragraphs = tableCellParagraphs(cell.nodes);
        if (paragraphs.length === 0) {
          continue;
        }

        return {
          kind: "table-cell",
          tableIndex: nodeIndex,
          rowIndex,
          cellIndex,
          paragraphIndex: paragraphs.length - 1,
        };
      }
    }
  }

  return undefined;
}

// ── Range coverage / boundary resolution ────────────────────────────

export function rangeCoversEntireDocument(
  model: DocModel,
  range: DocxTextRange
): boolean {
  const normalizedRange = normalizeTextRange(range);
  const firstLocation = firstParagraphLocationInDocument(model);
  const lastLocation = lastParagraphLocationInDocument(model);
  if (!firstLocation || !lastLocation) {
    return false;
  }

  if (!sameParagraphLocation(normalizedRange.start.location, firstLocation)) {
    return false;
  }

  if (!sameParagraphLocation(normalizedRange.end.location, lastLocation)) {
    return false;
  }

  const firstParagraph = getParagraphAtLocation(model, firstLocation).paragraph;
  const lastParagraph = getParagraphAtLocation(model, lastLocation).paragraph;
  if (!firstParagraph || !lastParagraph) {
    return false;
  }

  const safeStart = normalizeRangeBoundaryParagraphOffset(
    firstParagraph,
    normalizedRange.start.offset
  );
  const safeEnd = normalizeRangeBoundaryParagraphOffset(
    lastParagraph,
    normalizedRange.end.offset
  );
  return safeStart <= 0 && safeEnd >= paragraphText(lastParagraph).length;
}

export function resolveRangeBoundaryOffsetsForParagraph(
  currentLocation: ParagraphLocation,
  rangeStart: DocxTextRangeBoundary,
  rangeEnd: DocxTextRangeBoundary,
  paragraph: ParagraphNode
): [number, number] {
  const ordered =
    compareTextRangeBoundaries(rangeStart, rangeEnd) <= 0
      ? [rangeStart, rangeEnd]
      : [rangeEnd, rangeStart];
  const startBoundary = ordered[0];
  const endBoundary = ordered[1];
  const isStartBoundaryHere = sameParagraphLocation(
    currentLocation,
    startBoundary.location
  );
  const isEndBoundaryHere = sameParagraphLocation(
    currentLocation,
    endBoundary.location
  );

  const safeStart = isStartBoundaryHere
    ? normalizeRangeBoundaryParagraphOffset(paragraph, startBoundary.offset)
    : 0;
  const safeEnd = isEndBoundaryHere
    ? normalizeRangeBoundaryParagraphOffset(paragraph, endBoundary.offset)
    : paragraphText(paragraph).length;

  return [safeStart, safeEnd];
}

// ── Cloning helpers ──────────────────────────────────────────────────

export function cloneTextRange(range?: DocxTextRange): DocxTextRange | undefined {
  if (!range) {
    return undefined;
  }

  return {
    start: {
      location: cloneTextRangeLocation(range.start.location),
      offset: range.start.offset,
    },
    end: {
      location: cloneTextRangeLocation(range.end.location),
      offset: range.end.offset,
    },
  };
}

export function cloneEditorSelection(
  selection: DocxEditorSelection
): DocxEditorSelection {
  if (selection.kind === "paragraph") {
    return {
      kind: "paragraph",
      nodeIndex: selection.nodeIndex,
    };
  }

  const cloned: Extract<DocxEditorSelection, { kind: "table-cell" }> = {
    kind: "table-cell",
    tableIndex: selection.tableIndex,
    rowIndex: selection.rowIndex,
    cellIndex: selection.cellIndex,
  };
  if (selection.tableRuntimeKey !== undefined) {
    cloned.tableRuntimeKey = selection.tableRuntimeKey;
  }
  if (selection.cellRange) {
    cloned.cellRange = cloneTableCellSelectionRange(selection.cellRange);
  }
  return cloned;
}

export function cloneTableCellSelectionRange(
  range: DocxTableCellSelectionRange
): DocxTableCellSelectionRange {
  const cloned: DocxTableCellSelectionRange = {
    tableIndex: range.tableIndex,
    anchorRowIndex: range.anchorRowIndex,
    anchorCellIndex: range.anchorCellIndex,
    focusRowIndex: range.focusRowIndex,
    focusCellIndex: range.focusCellIndex,
  };
  if (range.gridRange) {
    cloned.gridRange = {
      startColumnIndex: range.gridRange.startColumnIndex,
      endColumnIndex: range.gridRange.endColumnIndex,
    };
  }
  return cloned;
}

function sameTableGridRange(
  a: DocxTableCellSelectionRange["gridRange"],
  b: DocxTableCellSelectionRange["gridRange"]
): boolean {
  if (!a || !b) {
    return a === b;
  }
  return (
    a.startColumnIndex === b.startColumnIndex &&
    a.endColumnIndex === b.endColumnIndex
  );
}

/** True when the range covers more than one model cell. */
export function tableCellSelectionRangeHasMultipleCells(
  range: DocxTableCellSelectionRange
): boolean {
  return (
    range.anchorRowIndex !== range.focusRowIndex ||
    range.anchorCellIndex !== range.focusCellIndex
  );
}

/** Anchor cell of a table range, which Escape and Shift+click extend from. */
export function tableCellSelectionAnchor(range: DocxTableCellSelectionRange): {
  rowIndex: number;
  cellIndex: number;
} {
  return {
    rowIndex: range.anchorRowIndex,
    cellIndex: range.anchorCellIndex,
  };
}

/**
 * Normalizes a table-cell range against the model table.
 *
 * Anchor and focus are clamped to existing cells, then the physical grid span
 * is resolved through `tableSelectionGridRectangle` and stored on the range.
 * Returns undefined when either endpoint no longer exists.
 */
export function normalizeTableCellSelectionRange(
  table: TableNode,
  range: DocxTableCellSelectionRange
): DocxTableCellSelectionRange | undefined {
  const anchorCell =
    table.rows[range.anchorRowIndex]?.cells[range.anchorCellIndex];
  const focusCell = table.rows[range.focusRowIndex]?.cells[range.focusCellIndex];
  if (!anchorCell || !focusCell) {
    return undefined;
  }

  const normalized = cloneTableCellSelectionRange({
    ...range,
    gridRange: undefined,
  });
  const rectangle = tableSelectionGridRectangle(table, normalized);
  if (!rectangle) {
    return undefined;
  }
  normalized.gridRange = {
    startColumnIndex: rectangle.minColumn,
    endColumnIndex: rectangle.maxColumn,
  };
  return normalized;
}

/**
 * Cells whose physical grid rectangle intersects the selection rectangle.
 * Includes cells reached only through a grid span or a vertical merge, so
 * formatting, copying and clearing cover everything the highlight covers.
 */
export function tableCellSelectionRangeCells(
  table: TableNode,
  range: DocxTableCellSelectionRange
): Array<{ rowIndex: number; cellIndex: number }> {
  const rectangle = tableSelectionGridRectangle(table, range);
  if (!rectangle) {
    return [];
  }

  const cells: Array<{ rowIndex: number; cellIndex: number }> = [];
  for (
    let rowIndex = rectangle.minRow;
    rowIndex <= rectangle.maxRow;
    rowIndex += 1
  ) {
    const row = table.rows[rowIndex];
    if (!row) {
      continue;
    }
    for (const ref of tableRowGridCells(row)) {
      const refEnd = ref.startColumn + ref.span - 1;
      if (
        refEnd >= rectangle.minColumn &&
        ref.startColumn <= rectangle.maxColumn
      ) {
        cells.push({ rowIndex, cellIndex: ref.cellIndex });
      }
    }
  }
  return cells;
}

/**
 * True when a model cell is inside the selection rectangle. A cell that only
 * partially overlaps through `gridSpan` still counts: Word highlights the
 * whole cell rather than half of it.
 */
export function isCellWithinTableSelectionRange(
  table: TableNode,
  range: DocxTableCellSelectionRange,
  rowIndex: number,
  cellIndex: number
): boolean {
  const rectangle = tableSelectionGridRectangle(table, range);
  if (!rectangle) {
    return false;
  }
  if (rowIndex < rectangle.minRow || rowIndex > rectangle.maxRow) {
    return false;
  }
  const row = table.rows[rowIndex];
  if (!row) {
    return false;
  }
  const ref = tableRowGridCells(row).find(
    (candidate) => candidate.cellIndex === cellIndex
  );
  if (!ref) {
    return false;
  }
  const refEnd = ref.startColumn + ref.span - 1;
  return refEnd >= rectangle.minColumn && ref.startColumn <= rectangle.maxColumn;
}

/**
 * True when the selection points at this cell. Single-cell selections match
 * exactly; range selections are tested against the physical grid so merged and
 * spanning cells highlight as a whole. Pass the table to make the grid test
 * exact; without it the model row/cell bounds are used as a fallback.
 */
export function isCellSelected(
  selection: DocxEditorSelection,
  tableIndex: number,
  rowIndex: number,
  cellIndex: number,
  table?: TableNode
): boolean {
  if (selection.kind !== "table-cell" || selection.tableIndex !== tableIndex) {
    return false;
  }
  if (!selection.cellRange) {
    return selection.rowIndex === rowIndex && selection.cellIndex === cellIndex;
  }
  if (table) {
    return isCellWithinTableSelectionRange(
      table,
      selection.cellRange,
      rowIndex,
      cellIndex
    );
  }
  return cellIsInsideSelectionRange({ rowIndex, cellIndex }, selection.cellRange);
}

/** Cell paragraphs covered by a multi-cell table selection. */
export function tableCellSelectionRangeParagraphEntries(
  table: TableNode,
  range: DocxTableCellSelectionRange
): Array<{
  rowIndex: number;
  cellIndex: number;
  cell: TableCellNode;
  paragraph: ParagraphNode;
  paragraphIndex: number;
}> {
  const entries: Array<{
    rowIndex: number;
    cellIndex: number;
    cell: TableCellNode;
    paragraph: ParagraphNode;
    paragraphIndex: number;
  }> = [];
  for (const { rowIndex, cellIndex } of tableCellSelectionRangeCells(
    table,
    range
  )) {
    const cell = table.rows[rowIndex]?.cells[cellIndex];
    if (!cell) {
      continue;
    }
    tableCellParagraphs(cell.nodes).forEach((paragraph, paragraphIndex) => {
      entries.push({ rowIndex, cellIndex, cell, paragraph, paragraphIndex });
    });
  }
  return entries;
}

function sameTableCellSelectionRange(
  a: DocxTableCellSelectionRange | undefined,
  b: DocxTableCellSelectionRange | undefined
): boolean {
  if (!a || !b) {
    return a === b;
  }
  return (
    a.tableIndex === b.tableIndex &&
    a.anchorRowIndex === b.anchorRowIndex &&
    a.anchorCellIndex === b.anchorCellIndex &&
    a.focusRowIndex === b.focusRowIndex &&
    a.focusCellIndex === b.focusCellIndex &&
    sameTableGridRange(a.gridRange, b.gridRange)
  );
}

export function sameEditorSelection(
  a: DocxEditorSelection,
  b: DocxEditorSelection
): boolean {
  if (a.kind !== b.kind) {
    return false;
  }

  if (a.kind === "paragraph") {
    return (
      a.nodeIndex ===
      (b as Extract<DocxEditorSelection, { kind: "paragraph" }>).nodeIndex
    );
  }

  const tableSelection = b as Extract<
    DocxEditorSelection,
    { kind: "table-cell" }
  >;
  return (
    a.tableIndex === tableSelection.tableIndex &&
    a.rowIndex === tableSelection.rowIndex &&
    a.cellIndex === tableSelection.cellIndex &&
    a.tableRuntimeKey === tableSelection.tableRuntimeKey &&
    sameTableCellSelectionRange(a.cellRange, tableSelection.cellRange)
  );
}

export function sameTextRangeBoundary(
  a: DocxTextRangeBoundary,
  b: DocxTextRangeBoundary
): boolean {
  return compareTextRangeBoundaries(a, b) === 0;
}

export function sameTextRange(a?: DocxTextRange, b?: DocxTextRange): boolean {
  if (!a && !b) {
    return true;
  }
  if (!a || !b) {
    return false;
  }

  const normalizedA = normalizeTextRange(a);
  const normalizedB = normalizeTextRange(b);
  return (
    sameTextRangeBoundary(normalizedA.start, normalizedB.start) &&
    sameTextRangeBoundary(normalizedA.end, normalizedB.end)
  );
}

// ── Location conversion helpers ──────────────────────────────────────

export function imageLocationToParagraphLocation(
  location: DocxImageLocation
): ParagraphLocation {
  if (location.kind === "paragraph") {
    return {
      kind: "paragraph",
      nodeIndex: location.nodeIndex,
    };
  }

  return {
    kind: "table-cell",
    tableIndex: location.tableIndex,
    rowIndex: location.rowIndex,
    cellIndex: location.cellIndex,
    paragraphIndex: location.paragraphIndex,
  };
}

export function selectionFallbackParagraphLocation(
  selection: DocxEditorSelection
): ParagraphLocation {
  if (selection.kind === "paragraph") {
    return {
      kind: "paragraph",
      nodeIndex: selection.nodeIndex,
    };
  }

  return {
    kind: "table-cell",
    tableIndex: selection.tableIndex,
    rowIndex: selection.rowIndex,
    cellIndex: selection.cellIndex,
    paragraphIndex: 0,
  };
}

export function selectionFromTextRangeLocation(
  location: DocxTextRangeLocation
): DocxEditorSelection {
  if (location.kind === "paragraph") {
    return {
      kind: "paragraph",
      nodeIndex: location.nodeIndex,
    };
  }

  return {
    kind: "table-cell",
    tableIndex: location.tableIndex,
    rowIndex: location.rowIndex,
    cellIndex: location.cellIndex,
  };
}

export function paragraphLocationFromTextRangeLocation(
  location: DocxTextRangeLocation
): ParagraphLocation {
  if (location.kind === "paragraph") {
    return {
      kind: "paragraph",
      nodeIndex: location.nodeIndex,
    };
  }

  return {
    kind: "table-cell",
    tableIndex: location.tableIndex,
    rowIndex: location.rowIndex,
    cellIndex: location.cellIndex,
    paragraphIndex: Math.max(0, Math.round(location.paragraphIndex)),
  };
}

export function resolveSelectedParagraphLocation(
  selection: DocxEditorSelection,
  activeTextRange?: DocxTextRange
): ParagraphLocation {
  const activeLocation = activeTextRange?.start.location;
  if (activeLocation) {
    return paragraphLocationFromTextRangeLocation(activeLocation);
  }

  return selectionFallbackParagraphLocation(selection);
}

// ── Model normalization ─────────────────────────────────────────────

export function normalizeParagraphLocationForModel(
  model: DocModel,
  location: ParagraphLocation
): ParagraphLocation | undefined {
  const exact = getParagraphAtLocation(model, location).paragraph;
  if (exact) {
    return cloneTextRangeLocation(location);
  }

  if (location.kind === "table-cell") {
    const tableNode = model.nodes[location.tableIndex];
    if (tableNode && tableNode.type === "table") {
      const safeRowIndex = clampNumber(
        location.rowIndex,
        0,
        Math.max(0, tableNode.rows.length - 1)
      );
      const row = tableNode.rows[safeRowIndex];
      if (row) {
        const safeCellIndex = clampNumber(
          location.cellIndex,
          0,
          Math.max(0, row.cells.length - 1)
        );
        const candidateCell = row.cells[safeCellIndex];
        if (candidateCell && !candidateCell.style?.vMergeContinuation) {
          const cellParagraphs = tableCellParagraphs(candidateCell.nodes);
          if (cellParagraphs.length > 0) {
            return {
              kind: "table-cell",
              tableIndex: location.tableIndex,
              rowIndex: safeRowIndex,
              cellIndex: safeCellIndex,
              paragraphIndex: clampNumber(
                location.paragraphIndex,
                0,
                Math.max(0, cellParagraphs.length - 1)
              ),
            };
          }
        }
      }

      const firstInTable = firstParagraphLocationInTable(
        model,
        location.tableIndex
      );
      if (firstInTable) {
        return firstInTable;
      }
    }
  }

  if (model.nodes.length === 0) {
    return undefined;
  }

  const anchorNodeIndex =
    location.kind === "paragraph" ? location.nodeIndex : location.tableIndex;
  const clampedAnchorNodeIndex = clampNumber(
    anchorNodeIndex,
    0,
    Math.max(0, model.nodes.length - 1)
  );

  const sameNodeFallback = firstParagraphLocationFromNode(
    model,
    clampedAnchorNodeIndex
  );
  if (sameNodeFallback) {
    return sameNodeFallback;
  }

  for (
    let nodeIndex = clampedAnchorNodeIndex + 1;
    nodeIndex < model.nodes.length;
    nodeIndex += 1
  ) {
    const forward = firstParagraphLocationFromNode(model, nodeIndex);
    if (forward) {
      return forward;
    }
  }

  for (
    let nodeIndex = clampedAnchorNodeIndex - 1;
    nodeIndex >= 0;
    nodeIndex -= 1
  ) {
    const backward = firstParagraphLocationFromNode(model, nodeIndex);
    if (backward) {
      return backward;
    }
  }

  return firstParagraphLocationInDocument(model);
}

export function normalizeTextRangeForModel(
  model: DocModel,
  range?: DocxTextRange
): DocxTextRange | undefined {
  if (!range) {
    return undefined;
  }

  const normalized = normalizeTextRange(range);
  const startLocation = normalizeParagraphLocationForModel(
    model,
    paragraphLocationFromTextRangeLocation(normalized.start.location)
  );
  const endLocation = normalizeParagraphLocationForModel(
    model,
    paragraphLocationFromTextRangeLocation(normalized.end.location)
  );
  if (!startLocation || !endLocation) {
    return undefined;
  }

  const startParagraph = getParagraphAtLocation(model, startLocation).paragraph;
  const endParagraph = getParagraphAtLocation(model, endLocation).paragraph;
  if (!startParagraph || !endParagraph) {
    return undefined;
  }

  return normalizeTextRange({
    start: {
      location: cloneTextRangeLocation(startLocation),
      offset: normalizeRangeBoundaryParagraphOffset(
        startParagraph,
        normalized.start.offset
      ),
    },
    end: {
      location: cloneTextRangeLocation(endLocation),
      offset: normalizeRangeBoundaryParagraphOffset(
        endParagraph,
        normalized.end.offset
      ),
    },
  });
}

export function normalizeSelectionForModel(
  model: DocModel,
  selection: DocxEditorSelection,
  preferredLocation?: DocxTextRangeLocation
): DocxEditorSelection {
  // Nested table cells live inside a cell of the host table. The host table
  // cannot normalize inner row/cell coordinates, so preserve the runtime-key
  // selection verbatim.
  if (selection.kind === "table-cell" && selection.tableRuntimeKey) {
    const preserved: Extract<DocxEditorSelection, { kind: "table-cell" }> = {
      kind: "table-cell",
      tableIndex: selection.tableIndex,
      rowIndex: selection.rowIndex,
      cellIndex: selection.cellIndex,
      tableRuntimeKey: selection.tableRuntimeKey,
    };
    if (selection.cellRange) {
      preserved.cellRange = cloneTableCellSelectionRange(selection.cellRange);
    }
    return preserved;
  }

  const normalizedParagraphLocation = normalizeParagraphLocationForModel(
    model,
    preferredLocation
      ? paragraphLocationFromTextRangeLocation(preferredLocation)
      : selectionFallbackParagraphLocation(selection)
  );
  if (!normalizedParagraphLocation) {
    return {
      kind: "paragraph",
      nodeIndex: 0,
    };
  }

  const normalizedSelection = selectionFromTextRangeLocation(
    normalizedParagraphLocation
  );
  if (
    selection.kind === "table-cell" &&
    selection.cellRange &&
    normalizedSelection.kind === "table-cell" &&
    normalizedSelection.tableIndex === selection.tableIndex
  ) {
    const table = model.nodes[selection.tableIndex];
    const normalizedRange =
      table?.type === "table"
        ? normalizeTableCellSelectionRange(table, selection.cellRange)
        : cloneTableCellSelectionRange(selection.cellRange);
    const insideRange =
      normalizedRange &&
      (table?.type === "table"
        ? isCellWithinTableSelectionRange(
            table,
            normalizedRange,
            normalizedSelection.rowIndex,
            normalizedSelection.cellIndex
          )
        : cellIsInsideSelectionRange(normalizedSelection, normalizedRange));
    if (normalizedRange && insideRange) {
      return {
        ...normalizedSelection,
        cellRange: normalizedRange,
      };
    }
  }

  return normalizedSelection;
}

function cellIsInsideSelectionRange(
  cell: { rowIndex: number; cellIndex: number },
  range: DocxTableCellSelectionRange
): boolean {
  const minRow = Math.min(range.anchorRowIndex, range.focusRowIndex);
  const maxRow = Math.max(range.anchorRowIndex, range.focusRowIndex);
  const minCell = Math.min(range.anchorCellIndex, range.focusCellIndex);
  const maxCell = Math.max(range.anchorCellIndex, range.focusCellIndex);
  return (
    cell.rowIndex >= minRow &&
    cell.rowIndex <= maxRow &&
    cell.cellIndex >= minCell &&
    cell.cellIndex <= maxCell
  );
}

export function normalizeEditorCursorStateForModel(
  model: DocModel,
  selection: DocxEditorSelection,
  activeTextRange?: DocxTextRange
): {
  selection: DocxEditorSelection;
  activeTextRange?: DocxTextRange;
} {
  const normalizedRange = normalizeTextRangeForModel(model, activeTextRange);
  if (normalizedRange) {
    return {
      selection: normalizeSelectionForModel(
        model,
        selection,
        normalizedRange.start.location
      ),
      activeTextRange: normalizedRange,
    };
  }

  return {
    selection: normalizeSelectionForModel(model, selection),
    activeTextRange: undefined,
  };
}
