// Selection / caret range normalization, cloning, equality, and resolution
// helpers. Range-to-model normalization keeps stale selection/range state
// valid after model mutations. Split from selection-helpers.ts to keep each
// file under the 1000-line hard constraint.
// Upstream editor.tsx: lines 23949-24145, 24261-24639.

import type {
  DocModel,
  ParagraphNode,
} from "../../engine/types";
import type {
  DocxEditorSelection,
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

  return {
    kind: "table-cell",
    tableIndex: selection.tableIndex,
    rowIndex: selection.rowIndex,
    cellIndex: selection.cellIndex,
  };
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
    a.cellIndex === tableSelection.cellIndex
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
  selection: DocxEditorSelection
): DocxEditorSelection {
  const normalizedParagraphLocation = normalizeParagraphLocationForModel(
    model,
    selectionFallbackParagraphLocation(selection)
  );
  if (!normalizedParagraphLocation) {
    return {
      kind: "paragraph",
      nodeIndex: 0,
    };
  }

  return selectionFromTextRangeLocation(normalizedParagraphLocation);
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
        selectionFromTextRangeLocation(normalizedRange.start.location)
      ),
      activeTextRange: normalizedRange,
    };
  }

  return {
    selection: normalizeSelectionForModel(model, selection),
    activeTextRange: undefined,
  };
}
