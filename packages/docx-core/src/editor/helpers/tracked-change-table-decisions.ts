// Table revision decisions: accept/reject rows, cells and table properties.
//
// Table revisions live in three element families (see tracked-changes-table.ts):
//   * `<w:tblPrChange>`, `<w:trPrChange>`, `<w:tcPrChange>` — property snapshots
//   * `<w:trPr><w:ins|w:del/></w:trPr>` — inserted/deleted rows
//   * `<w:tcPr><w:cellIns|w:cellDel/></w:tcPr>` — inserted/deleted cells
//
// The model only keeps current properties, so a decision has to edit the table's
// `sourceXml` and the row/cell arrays together:
//
//   * accepting an inserted row (or rejecting a deleted one) keeps the row and
//     drops the revision marker;
//   * rejecting an inserted row (or accepting a deleted one) removes both the
//     model row and its `<w:tr>` element;
//   * property changes only have their markup cleared — the model stores no
//     previous table formatting, so there is nothing to revert.

import type { DocModel, TableNode } from "../../engine/types";
import type { DocxTrackedChange } from "./editor-types";
import {
  extractBalancedTagRanges,
  extractTopLevelTagRanges,
  type XmlBalancedTagRange
} from "./xml-parsing";
import { xmlAttribute } from "./ooxml-helpers";
import {
  resolveTableTrackedMarkup,
  type TableTrackedRevision
} from "./tracked-changes-table";
import {
  applyXmlEdits,
  collectPropertyChangeRanges,
  type XmlEdit
} from "./tracked-change-xml";

export type TableDecisionKind = "accept" | "reject";

const TABLE_PROPERTY_TAG_BY_KIND: Record<string, string> = {
  "table-format-change": "w:tblPrChange",
  "row-format-change": "w:trPrChange",
  "cell-format-change": "w:tcPrChange"
};

/** `t:0:1:0:0:row-insertion-4:2` → `row-insertion-4`. */
export function tableRevisionStableId(
  change: DocxTrackedChange
): string | undefined {
  if (change.inlineAnchorId) {
    return undefined;
  }
  const parts = change.id.split(":");
  return parts.length >= 2 ? parts[parts.length - 2] : undefined;
}

export function findTableRevision(
  table: TableNode,
  stableId: string | undefined
): TableTrackedRevision | undefined {
  if (!stableId) {
    return undefined;
  }
  return resolveTableTrackedMarkup(table)?.revisions.find(
    (revision) => revision.stableId === stableId
  );
}

interface TableRowCellXml {
  rows: XmlBalancedTagRange[];
  cellsByRow: Array<Map<number, XmlBalancedTagRange>>;
}

function tableContainerRanges(sourceXml: string): TableRowCellXml {
  const rows = extractTopLevelTagRanges(sourceXml, "w:tr").sort(
    (left, right) => left.start - right.start
  );
  const cellsByRow = rows.map((row) => {
    const cells = new Map<number, XmlBalancedTagRange>();
    extractTopLevelTagRanges(sourceXml.slice(row.start, row.end), "w:tc")
      .sort((left, right) => left.start - right.start)
      .forEach((cell, index) => {
        cells.set(index, {
          ...cell,
          start: cell.start + row.start,
          end: cell.end + row.start
        });
      });
    return cells;
  });
  return { rows, cellsByRow };
}

/** Locates `<w:ins/>` / `<w:del/>` / `<w:cellIns/>` / `<w:cellDel/>` markup. */
function findRowCellMarkupRange(
  sourceXml: string,
  container: XmlBalancedTagRange,
  tagName: string,
  revisionId: string | undefined
): XmlBalancedTagRange | undefined {
  const slice = sourceXml.slice(container.start, container.end);
  const ranges = extractBalancedTagRanges(slice, tagName);
  const match = revisionId
    ? ranges.find(
        (range) => xmlAttribute(range.openTag, "w:id")?.trim() === revisionId
      )
    : ranges[0];
  if (!match) {
    return undefined;
  }
  return {
    ...match,
    start: match.start + container.start,
    end: match.end + container.start
  };
}

export interface TableDecisionPlan {
  edits: XmlEdit[];
  removeRowIndex?: number;
  removeCell?: { rowIndex: number; cellIndex: number };
  clearedOnly: boolean;
}

/**
 * Plans one table revision. Returns `undefined` when the revision cannot be
 * located in the table's `sourceXml`, which marks it unresolved for the caller.
 */
export function planTableDecision(
  table: TableNode,
  change: DocxTrackedChange,
  decision: TableDecisionKind
): TableDecisionPlan | undefined {
  const sourceXml = table.sourceXml ?? "";
  if (!sourceXml) {
    return undefined;
  }

  const stableId = tableRevisionStableId(change);
  const revision = findTableRevision(table, stableId);
  if (!stableId || !revision) {
    return undefined;
  }

  const containers = tableContainerRanges(sourceXml);
  const plan: TableDecisionPlan = { edits: [], clearedOnly: false };

  const propertyTag = TABLE_PROPERTY_TAG_BY_KIND[revision.kind];
  if (propertyTag) {
    const candidates = collectPropertyChangeRanges(sourceXml, propertyTag).filter(
      (range) => {
        if (revision.revisionId && range.revisionId !== revision.revisionId) {
          return false;
        }
        if (revision.scope === "table") {
          return true;
        }
        const rowRange =
          revision.rowIndex !== undefined
            ? containers.rows[revision.rowIndex]
            : undefined;
        if (!rowRange || range.start < rowRange.start || range.end > rowRange.end) {
          return false;
        }
        if (revision.scope === "row") {
          return true;
        }
        const cellRange =
          revision.rowIndex !== undefined && revision.cellIndex !== undefined
            ? containers.cellsByRow[revision.rowIndex]?.get(revision.cellIndex)
            : undefined;
        return Boolean(
          cellRange && range.start >= cellRange.start && range.end <= cellRange.end
        );
      }
    );
    const range = candidates[0];
    if (!range) {
      return undefined;
    }
    plan.edits.push({ start: range.start, end: range.end, text: "" });
    // The model stores no previous table formatting, so rejecting and accepting
    // both drop the markup; report the difference to callers.
    plan.clearedOnly = decision === "reject";
    return plan;
  }

  if (revision.kind !== "insertion" && revision.kind !== "deletion") {
    return undefined;
  }

  const isRowRevision = revision.scope === "row";
  const isCellRevision = revision.scope === "cell";
  if (!isRowRevision && !isCellRevision) {
    return undefined;
  }

  const rowIndex = revision.rowIndex;
  const rowRange = rowIndex !== undefined ? containers.rows[rowIndex] : undefined;
  if (rowIndex === undefined || !rowRange) {
    return undefined;
  }

  if (isRowRevision) {
    const markupRange = findRowCellMarkupRange(
      sourceXml,
      rowRange,
      revision.kind === "insertion" ? "w:ins" : "w:del",
      revision.revisionId
    );
    if (!markupRange) {
      return undefined;
    }
    const keepRow =
      decision === "accept"
        ? revision.kind === "insertion"
        : revision.kind === "deletion";
    if (keepRow) {
      plan.edits.push({ start: markupRange.start, end: markupRange.end, text: "" });
      return plan;
    }
    plan.edits.push({ start: rowRange.start, end: rowRange.end, text: "" });
    plan.removeRowIndex = rowIndex;
    return plan;
  }

  const cellIndex = revision.cellIndex;
  if (cellIndex === undefined) {
    return undefined;
  }
  const cellRange = containers.cellsByRow[rowIndex]?.get(cellIndex);
  if (!cellRange) {
    return undefined;
  }
  const markupRange = findRowCellMarkupRange(
    sourceXml,
    cellRange,
    revision.kind === "insertion" ? "w:cellIns" : "w:cellDel",
    revision.revisionId
  );
  if (!markupRange) {
    return undefined;
  }
  const keepCell =
    decision === "accept"
      ? revision.kind === "insertion"
      : revision.kind === "deletion";
  if (keepCell) {
    plan.edits.push({ start: markupRange.start, end: markupRange.end, text: "" });
    return plan;
  }
  plan.edits.push({ start: cellRange.start, end: cellRange.end, text: "" });
  plan.removeCell = { rowIndex, cellIndex };
  return plan;
}

export interface TableDecisionOutcome {
  resolvedChangeIds: string[];
  unresolvedChangeIds: string[];
  clearedOnlyChangeIds: string[];
  changed: boolean;
}

/**
 * Applies every selected table revision of one table to a cloned model, editing
 * the table's `sourceXml` and its row/cell arrays together.
 */
export function applyTableRevisionDecisions(
  model: DocModel,
  nodeIndex: number,
  changes: readonly DocxTrackedChange[],
  decision: TableDecisionKind
): TableDecisionOutcome {
  const outcome: TableDecisionOutcome = {
    resolvedChangeIds: [],
    unresolvedChangeIds: [],
    clearedOnlyChangeIds: [],
    changed: false
  };

  const table = model.nodes[nodeIndex];
  if (!table || table.type !== "table") {
    outcome.unresolvedChangeIds.push(...changes.map((change) => change.id));
    return outcome;
  }

  // Rows and cells are removed from the highest index down so earlier removals
  // cannot shift the indices of later ones.
  const removeRows: number[] = [];
  const removeCells: Array<{ rowIndex: number; cellIndex: number }> = [];
  const edits: XmlEdit[] = [];
  const tableSourceXml = table.sourceXml ?? "";

  for (const change of changes) {
    const plan = planTableDecision(table, change, decision);
    if (!plan) {
      outcome.unresolvedChangeIds.push(change.id);
      continue;
    }
    outcome.resolvedChangeIds.push(change.id);
    if (plan.clearedOnly) {
      outcome.clearedOnlyChangeIds.push(change.id);
    }
    edits.push(...plan.edits);
    if (plan.removeRowIndex !== undefined) {
      removeRows.push(plan.removeRowIndex);
    }
    if (plan.removeCell) {
      removeCells.push(plan.removeCell);
    }
  }

  if (edits.length > 0) {
    table.sourceXml = applyXmlEdits(tableSourceXml, edits);
    outcome.changed = true;
  }
  for (const cell of [...removeCells].sort(
    (left, right) => right.cellIndex - left.cellIndex
  )) {
    table.rows[cell.rowIndex]?.cells.splice(cell.cellIndex, 1);
    outcome.changed = true;
  }
  for (const rowIndex of [...new Set(removeRows)].sort((left, right) => right - left)) {
    table.rows.splice(rowIndex, 1);
    outcome.changed = true;
  }

  return outcome;
}
