// Selection / caret location helpers: paragraph-location navigation, text-range
// boundary comparison/cloning, editor-selection cloning/equality, location
// resolution, form-field collection, and model normalization for stale
// selection/range state. DOM selection restore heuristics live in
// selection-restore.ts.
// Upstream editor.tsx: lines 23538-24145, 24261-24639.

import type {
  DocModel,
  FormFieldRunNode,
  ParagraphNode,
  TableNode
} from "../../engine/types";
import type {
  DocxEditorSelection,
  DocxImageLocation,
  DocxSelectedFormField,
  DocxTextRange,
  DocxTextRangeBoundary,
  DocxTextRangeLocation,
  ParagraphLocation
} from "./editor-types";
import {
  compareParagraphLocations,
  compareTextRangeBoundaries,
  normalizeTextRange
} from "./editor-types";
import { paragraphText, tableCellParagraphs } from "./paragraph-inspect";
import { clampNumber } from "./zoom-utils";

export function sameParagraphLocation(
  a: ParagraphLocation,
  b: ParagraphLocation
): boolean {
  if (a.kind === "paragraph") {
    return b.kind === "paragraph" && a.nodeIndex === b.nodeIndex;
  }

  if (b.kind === "paragraph") {
    return false;
  }

  return (
    a.tableIndex === b.tableIndex &&
    a.rowIndex === b.rowIndex &&
    a.cellIndex === b.cellIndex &&
    a.paragraphIndex === b.paragraphIndex
  );
}

export function firstParagraphLocationInTable(
  model: DocModel,
  tableIndex: number
): ParagraphLocation | undefined {
  const tableNode = model.nodes[tableIndex];
  if (!tableNode || tableNode.type !== "table") {
    return undefined;
  }

  for (let rowIndex = 0; rowIndex < tableNode.rows.length; rowIndex += 1) {
    const row = tableNode.rows[rowIndex];
    if (!row) {
      continue;
    }

    for (let cellIndex = 0; cellIndex < row.cells.length; cellIndex += 1) {
      const cell = row.cells[cellIndex];
      if (!cell || cell.style?.vMergeContinuation) {
        continue;
      }

      const cellParagraphs = tableCellParagraphs(cell.nodes);
      if (cellParagraphs.length === 0) {
        continue;
      }

      return {
        kind: "table-cell",
        tableIndex,
        rowIndex,
        cellIndex,
        paragraphIndex: 0,
      };
    }
  }

  return undefined;
}

export function lastParagraphLocationInTable(
  model: DocModel,
  tableIndex: number
): ParagraphLocation | undefined {
  const tableNode = model.nodes[tableIndex];
  if (!tableNode || tableNode.type !== "table") {
    return undefined;
  }

  for (let rowIndex = tableNode.rows.length - 1; rowIndex >= 0; rowIndex -= 1) {
    const row = tableNode.rows[rowIndex];
    if (!row) {
      continue;
    }

    for (let cellIndex = row.cells.length - 1; cellIndex >= 0; cellIndex -= 1) {
      const cell = row.cells[cellIndex];
      if (!cell || cell.style?.vMergeContinuation) {
        continue;
      }

      const cellParagraphs = tableCellParagraphs(cell.nodes);
      if (cellParagraphs.length === 0) {
        continue;
      }

      return {
        kind: "table-cell",
        tableIndex,
        rowIndex,
        cellIndex,
        paragraphIndex: cellParagraphs.length - 1,
      };
    }
  }

  return undefined;
}

export function nodeIndexFromParagraphLocation(
  location: ParagraphLocation
): number {
  return location.kind === "paragraph"
    ? location.nodeIndex
    : location.tableIndex;
}

export function adjustLocationAfterRemovedNodeIndexes(
  location: DocxTextRangeLocation,
  removedNodeIndexes: number[]
): DocxTextRangeLocation | undefined {
  if (removedNodeIndexes.length === 0) {
    return cloneTextRangeLocation(location);
  }

  const normalizedRemoved = [...removedNodeIndexes]
    .filter((value) => Number.isFinite(value) && value >= 0)
    .map((value) => Math.round(value))
    .sort((left, right) => left - right);

  const sourceNodeIndex =
    location.kind === "paragraph" ? location.nodeIndex : location.tableIndex;
  let adjustedNodeIndex = sourceNodeIndex;
  for (const removedIndex of normalizedRemoved) {
    if (removedIndex === adjustedNodeIndex) {
      return undefined;
    }
    if (removedIndex < adjustedNodeIndex) {
      adjustedNodeIndex -= 1;
    }
  }

  if (location.kind === "paragraph") {
    return {
      kind: "paragraph",
      nodeIndex: Math.max(0, adjustedNodeIndex),
    };
  }

  return {
    kind: "table-cell",
    tableIndex: Math.max(0, adjustedNodeIndex),
    rowIndex: location.rowIndex,
    cellIndex: location.cellIndex,
    paragraphIndex: location.paragraphIndex,
  };
}

export function tableCoverageBoundaries(
  model: DocModel,
  tableIndex: number
):
  | {
      start: DocxTextRangeBoundary;
      end: DocxTextRangeBoundary;
    }
  | undefined {
  const firstLocation = firstParagraphLocationInTable(model, tableIndex);
  const lastLocation = lastParagraphLocationInTable(model, tableIndex);
  if (!firstLocation || !lastLocation) {
    return undefined;
  }

  const firstParagraph = getParagraphAtLocation(model, firstLocation).paragraph;
  const lastParagraph = getParagraphAtLocation(model, lastLocation).paragraph;
  if (!firstParagraph || !lastParagraph) {
    return undefined;
  }

  return {
    start: {
      location: cloneTextRangeLocation(firstLocation),
      offset: 0,
    },
    end: {
      location: cloneTextRangeLocation(lastLocation),
      offset: paragraphText(lastParagraph).length,
    },
  };
}

export function fullyCoveredTableNodeIndexesForRange(
  model: DocModel,
  normalizedRange: DocxTextRange
): number[] {
  const startNodeIndex = nodeIndexFromParagraphLocation(
    normalizedRange.start.location
  );
  const endNodeIndex = nodeIndexFromParagraphLocation(
    normalizedRange.end.location
  );
  const firstIndex = Math.min(startNodeIndex, endNodeIndex);
  const lastIndex = Math.max(startNodeIndex, endNodeIndex);
  const coveredTableIndexes: number[] = [];

  for (let nodeIndex = firstIndex; nodeIndex <= lastIndex; nodeIndex += 1) {
    const node = model.nodes[nodeIndex];
    if (!node || node.type !== "table") {
      continue;
    }

    const boundaries = tableCoverageBoundaries(model, nodeIndex);
    if (!boundaries) {
      continue;
    }

    const coversFromStart =
      compareTextRangeBoundaries(normalizedRange.start, boundaries.start) <= 0;
    const coversToEnd =
      compareTextRangeBoundaries(normalizedRange.end, boundaries.end) >= 0;
    if (!coversFromStart || !coversToEnd) {
      continue;
    }

    coveredTableIndexes.push(nodeIndex);
  }

  return coveredTableIndexes;
}

export function firstParagraphLocationFromNode(
  model: DocModel,
  nodeIndex: number
): ParagraphLocation | undefined {
  const node = model.nodes[nodeIndex];
  if (!node) {
    return undefined;
  }

  if (node.type === "paragraph") {
    return {
      kind: "paragraph",
      nodeIndex,
    };
  }

  return firstParagraphLocationInTable(model, nodeIndex);
}

export function lastParagraphLocationInNode(
  model: DocModel,
  nodeIndex: number
): ParagraphLocation | undefined {
  const node = model.nodes[nodeIndex];
  if (!node) {
    return undefined;
  }

  if (node.type === "paragraph") {
    return {
      kind: "paragraph",
      nodeIndex,
    };
  }

  return lastParagraphLocationInTable(model, nodeIndex);
}

export function nextParagraphLocation(
  model: DocModel,
  location: ParagraphLocation
): ParagraphLocation | undefined {
  if (location.kind === "paragraph") {
    for (
      let nodeIndex = location.nodeIndex + 1;
      nodeIndex < model.nodes.length;
      nodeIndex += 1
    ) {
      const next = firstParagraphLocationFromNode(model, nodeIndex);
      if (next) {
        return next;
      }
    }
    return undefined;
  }

  const tableNode = model.nodes[location.tableIndex];
  if (!tableNode || tableNode.type !== "table") {
    return undefined;
  }

  const currentCell =
    tableNode.rows[location.rowIndex]?.cells[location.cellIndex];
  if (!currentCell) {
    return undefined;
  }

  const cellParagraphs = tableCellParagraphs(currentCell.nodes);
  if (location.paragraphIndex < cellParagraphs.length - 1) {
    return {
      kind: "table-cell",
      tableIndex: location.tableIndex,
      rowIndex: location.rowIndex,
      cellIndex: location.cellIndex,
      paragraphIndex: location.paragraphIndex + 1,
    };
  }

  for (
    let rowIndex = location.rowIndex;
    rowIndex < tableNode.rows.length;
    rowIndex += 1
  ) {
    const row = tableNode.rows[rowIndex];
    if (!row) {
      continue;
    }

    for (
      let cellIndex =
        rowIndex === location.rowIndex ? location.cellIndex + 1 : 0;
      cellIndex < row.cells.length;
      cellIndex += 1
    ) {
      const cell = row.cells[cellIndex];
      if (!cell || cell.style?.vMergeContinuation) {
        continue;
      }

      if (tableCellParagraphs(cell.nodes).length === 0) {
        continue;
      }

      return {
        kind: "table-cell",
        tableIndex: location.tableIndex,
        rowIndex,
        cellIndex,
        paragraphIndex: 0,
      };
    }
  }

  for (
    let nodeIndex = location.tableIndex + 1;
    nodeIndex < model.nodes.length;
    nodeIndex += 1
  ) {
    const next = firstParagraphLocationFromNode(model, nodeIndex);
    if (next) {
      return next;
    }
  }

  return undefined;
}

export function firstParagraphLocationInDocument(
  model: DocModel
): ParagraphLocation | undefined {
  for (let nodeIndex = 0; nodeIndex < model.nodes.length; nodeIndex += 1) {
    const first = firstParagraphLocationFromNode(model, nodeIndex);
    if (first) {
      return first;
    }
  }

  return undefined;
}

export function paragraphRangeForMutate(
  model: DocModel,
  start: ParagraphLocation,
  end: ParagraphLocation
): {
  location: ParagraphLocation;
}[] {
  const items: { location: ParagraphLocation }[] = [];

  const ordered =
    compareParagraphLocations(start, end) <= 0 ? [start, end] : [end, start];
  let current: ParagraphLocation | undefined = ordered[0];
  const limit = ordered[1];
  while (current) {
    items.push({ location: current });
    if (compareParagraphLocations(current, limit) >= 0) {
      break;
    }
    current = nextParagraphLocation(model, current);
  }

  return items;
}

export function normalizeRangeBoundaryParagraphOffset(
  paragraph: ParagraphNode,
  offset: number
): number {
  const length = paragraphText(paragraph).length;
  return Math.max(0, Math.min(Math.max(0, length), Math.round(offset)));
}

export function cloneTextRangeLocation(
  location: DocxTextRangeLocation
): DocxTextRangeLocation {
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

export function getParagraphAtLocation(
  model: DocModel,
  location: ParagraphLocation
): {
  paragraph?: ParagraphNode;
  tableNode?: TableNode;
} {
  if (location.kind === "paragraph") {
    const node = model.nodes[location.nodeIndex];
    if (!node || node.type !== "paragraph") {
      return {};
    }

    return { paragraph: node };
  }

  const tableNode = model.nodes[location.tableIndex];
  if (!tableNode || tableNode.type !== "table") {
    return {};
  }

  const cell = tableNode.rows[location.rowIndex]?.cells[location.cellIndex];
  if (!cell) {
    return {};
  }

  const paragraph = tableCellParagraphs(cell.nodes)[location.paragraphIndex];
  if (!paragraph) {
    return {};
  }

  return { paragraph, tableNode };
}

export function collectFormFieldsFromModel(
  model: DocModel
): DocxSelectedFormField[] {
  const collected: DocxSelectedFormField[] = [];

  model.nodes.forEach((node, nodeIndex) => {
    if (node.type === "paragraph") {
      node.children.forEach((child, childIndex) => {
        if (child.type !== "form-field") {
          return;
        }

        collected.push({
          location: {
            kind: "paragraph",
            nodeIndex,
            childIndex,
          },
          field: child as FormFieldRunNode,
        });
      });
      return;
    }

    node.rows.forEach((row, rowIndex) => {
      row.cells.forEach((cell, cellIndex) => {
        tableCellParagraphs(cell.nodes).forEach((paragraph, paragraphIndex) => {
          paragraph.children.forEach((child, childIndex) => {
            if (child.type !== "form-field") {
              return;
            }

            collected.push({
              location: {
                kind: "table-cell",
                tableIndex: nodeIndex,
                rowIndex,
                cellIndex,
                paragraphIndex,
                childIndex,
              },
              field: child as FormFieldRunNode,
            });
          });
        });
      });
    });
  });

  return collected;
}
