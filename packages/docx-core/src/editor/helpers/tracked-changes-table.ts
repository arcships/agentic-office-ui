// Tracked-change parsing for DOCX tables.
//
// Table revisions are spread across several element families:
//   * <w:tblPrChange>, <w:trPrChange>, <w:tcPrChange> — property there-and-back
//     snapshots
//   * <w:trPr><w:ins|w:del/></w:trPr> — inserted/deleted rows
//   * <w:tcPr><w:cellIns|w:cellDel/></w:tcPr> — inserted/deleted cells
//
// The model only keeps the current properties, so this module reads the table's
// source XML and resolves each revision back to the top-level row/cell that
// actually contains it. Both the gutter collector and the table renderer use
// the resulting markup.

import type {
  ParagraphNode,
  TableCellNode,
  TableNode,
  TableRowNode
} from "../../engine/types";
import type {
  DocxTextRangeLocation,
  DocxTrackedChangeKind
} from "./editor-types";
import { setCacheEntry } from "./cache-utils";
import { xmlAttribute } from "./ooxml-helpers";
import { tableCellParagraphs } from "./paragraph-inspect";
import { firstTableCellAnchorLocation } from "./text-mutation";
import {
  decodeXmlText,
  extractBalancedTagRanges,
  extractTopLevelTagRanges,
  normalizeTrackedChangeSnippet,
  type XmlBalancedTagRange
} from "./xml-parsing";

export type TableTrackedRevisionScope = "table" | "row" | "cell";

export interface TableTrackedRevision {
  stableId: string;
  kind: DocxTrackedChangeKind;
  scope: TableTrackedRevisionScope;
  author?: string;
  date?: string;
  text?: string;
  revisionId?: string;
  /** Top-level row this revision belongs to (undefined for table properties). */
  rowIndex?: number;
  /** Top-level cell this revision belongs to (undefined for row properties). */
  cellIndex?: number;
}

export interface TableTrackedMarkup {
  revisions: TableTrackedRevision[];
  /** Inserted/deleted rows keyed by top-level row index. */
  rowChangeByIndex: Map<number, TableTrackedRevision>;
  /** Inserted/deleted cells keyed by `${rowIndex}:${cellIndex}`. */
  cellChangeByIndex: Map<string, TableTrackedRevision>;
}

const tableTrackedMarkupBySourceXml = new Map<
  string,
  TableTrackedMarkup | null
>();

interface TableContainerRanges {
  rows: XmlBalancedTagRange[];
  cellsByRow: XmlBalancedTagRange[][];
}

function containsRange(
  outer: { start: number; end: number },
  inner: { start: number; end: number }
): boolean {
  return (
    outer.start <= inner.start &&
    outer.end >= inner.end &&
    (outer.start < inner.start || outer.end > inner.end)
  );
}

function tableContainerRanges(sourceXml: string): TableContainerRanges {
  const rows = extractTopLevelTagRanges(sourceXml, "w:tr").sort(
    (left, right) => left.start - right.start
  );
  const cellsByRow = rows.map((row) =>
    extractTopLevelTagRanges(sourceXml.slice(row.start, row.end), "w:tc")
      .map((cell) => ({
        ...cell,
        start: cell.start + row.start,
        end: cell.end + row.start
      }))
      .filter(
        (cell) =>
          !rows.some(
            (otherRow) =>
              otherRow !== row && containsRange(otherRow, cell)
          )
      )
      .sort((left, right) => left.start - right.start)
  );
  return { rows, cellsByRow };
}

function rangeIndexContaining(
  ranges: XmlBalancedTagRange[],
  offset: number
): number {
  return ranges.findIndex(
    (range) => offset >= range.start && offset < range.end
  );
}

function revisionAttributes(openTag: string): {
  revisionId?: string;
  author?: string;
  date?: string;
} {
  return {
    revisionId: xmlAttribute(openTag, "w:id")?.trim() || undefined,
    author: decodeXmlText(xmlAttribute(openTag, "w:author") ?? "") || undefined,
    date: xmlAttribute(openTag, "w:date")?.trim() || undefined
  };
}

function paragraphTextSnippet(paragraph: ParagraphNode): string | undefined {
  const text = paragraph.children
    .filter((child) => child.type === "text")
    .map((child) => child.text)
    .join("");
  return normalizeTrackedChangeSnippet(text);
}

function cellTextSnippet(cell: TableCellNode): string | undefined {
  for (const paragraph of tableCellParagraphs(cell.nodes)) {
    const snippet = paragraphTextSnippet(paragraph);
    if (snippet) {
      return snippet;
    }
  }
  return undefined;
}

function rowTextSnippet(row: TableRowNode): string | undefined {
  for (const cell of row.cells) {
    const snippet = cellTextSnippet(cell);
    if (snippet) {
      return snippet;
    }
  }
  return undefined;
}

function summarizeTableFormattingChange(
  scope: TableTrackedRevisionScope,
  changeXml: string
): string {
  const features: string[] = [];
  if (/<w:tblW\b|<w:tcW\b|<w:gridSpan\b/i.test(changeXml)) {
    features.push("width");
  }
  if (/<w:tblLayout\b/i.test(changeXml)) {
    features.push("layout");
  }
  if (/<w:tblInd\b|<w:ind\b/i.test(changeXml)) {
    features.push("indent");
  }
  if (/<w:tblCellMar\b|<w:tcMar\b/i.test(changeXml)) {
    features.push("margins");
  }
  if (/<w:(?:tblBorders|tcBorders|trBorders|pBdr)\b/i.test(changeXml)) {
    features.push("borders");
  }
  if (/<w:trHeight\b/i.test(changeXml)) {
    features.push("row height");
  }
  if (/<w:vAlign\b/i.test(changeXml)) {
    features.push("vertical align");
  }
  if (/<w:jc\b/i.test(changeXml)) {
    features.push("alignment");
  }
  if (/<w:shd\b/i.test(changeXml)) {
    features.push("shading");
  }

  const prefix =
    scope === "table"
      ? "Table formatting"
      : scope === "row"
      ? "Row formatting"
      : "Cell formatting";
  if (features.length === 0) {
    return prefix;
  }
  const unique = Array.from(new Set(features));
  return `${prefix}: ${unique.join(", ")}`;
}

function directPropertyChangeContent(
  containerXml: string,
  changeTagName: "w:trPrChange" | "w:tcPrChange"
): string | undefined {
  const propertyMatch =
    changeTagName === "w:trPrChange"
      ? /<w:trPr\b[^>]*>([\s\S]*?)<\/w:trPr>/i.exec(containerXml)
      : /<w:tcPr\b[^>]*>([\s\S]*?)<\/w:tcPr>/i.exec(containerXml);
  if (!propertyMatch) {
    return undefined;
  }
  return propertyMatch[1]?.replace(
    new RegExp(`<${changeTagName}\\b[\\s\\S]*?</${changeTagName}>`, "gi"),
    ""
  );
}

function findRowOrCellRevisionTag(
  containerXml: string,
  tagName: "w:ins" | "w:del" | "w:cellIns" | "w:cellDel"
): string | undefined {
  const escapedTag = tagName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return containerXml.match(new RegExp(`<${escapedTag}\\b[^>]*/?>`, "i"))?.[0];
}

/**
 * Parses every revision kind a table can carry and resolves it to the
 * top-level row/cell that contains it. Cached by table source XML.
 */
export function resolveTableTrackedMarkup(
  table: TableNode
): TableTrackedMarkup | undefined {
  const sourceXml = table.sourceXml ?? "";
  if (!sourceXml) {
    return undefined;
  }

  const cached = tableTrackedMarkupBySourceXml.get(sourceXml);
  if (cached !== undefined) {
    return cached ?? undefined;
  }

  const revisions: TableTrackedRevision[] = [];
  const revisionByKey = new Map<string, TableTrackedRevision>();
  const rowChangeByIndex = new Map<number, TableTrackedRevision>();
  const cellChangeByIndex = new Map<string, TableTrackedRevision>();
  let anonymousCounter = 0;

  const appendRevision = (
    scope: TableTrackedRevisionScope,
    kind: DocxTrackedChangeKind,
    openTag: string,
    text: string | undefined,
    rowIndex?: number,
    cellIndex?: number
  ): TableTrackedRevision => {
    const { revisionId, author, date } = revisionAttributes(openTag);
    const key = revisionId
      ? `${scope}:${kind}:id:${revisionId}`
      : `${scope}:${kind}:${author ?? ""}:${date ?? ""}:${text ?? ""}`;
    const existing = revisionByKey.get(key);
    if (existing) {
      if (!existing.text && text) {
        existing.text = text;
      }
      if (!existing.author && author) {
        existing.author = author;
      }
      if (!existing.date && date) {
        existing.date = date;
      }
      return existing;
    }

    const stableId = revisionId
      ? `${scope}-${kind}-${revisionId}`
      : `${scope}-${kind}-${anonymousCounter}`;
    anonymousCounter += 1;
    const revision: TableTrackedRevision = {
      stableId,
      kind,
      scope,
      author,
      date,
      text,
      revisionId,
      rowIndex,
      cellIndex
    };
    revisions.push(revision);
    revisionByKey.set(key, revision);

    if (rowIndex !== undefined && (kind === "insertion" || kind === "deletion")) {
      rowChangeByIndex.set(rowIndex, revision);
    }
    if (
      rowIndex !== undefined &&
      cellIndex !== undefined &&
      (kind === "insertion" || kind === "deletion")
    ) {
      cellChangeByIndex.set(`${rowIndex}:${cellIndex}`, revision);
    }
    return revision;
  };

  const { rows, cellsByRow } = tableContainerRanges(sourceXml);

  // Inserted / deleted rows.
  rows.forEach((rowRange, rowIndex) => {
    const firstCellStart = cellsByRow[rowIndex]?.[0]?.start ?? rowRange.end;
    const rowPrefix = sourceXml.slice(rowRange.start, firstCellStart);
    const propertyContent = directPropertyChangeContent(
      rowPrefix,
      "w:trPrChange"
    );
    if (!propertyContent) {
      return;
    }
    const insertedTag = findRowOrCellRevisionTag(propertyContent, "w:ins");
    const deletedTag = findRowOrCellRevisionTag(propertyContent, "w:del");
    const tag = insertedTag ?? deletedTag;
    if (!tag) {
      return;
    }
    const row = table.rows[rowIndex];
    appendRevision(
      "row",
      insertedTag ? "insertion" : "deletion",
      tag,
      (row && rowTextSnippet(row)) ?? "Table row",
      rowIndex
    );
  });

  // Inserted / deleted cells.
  cellsByRow.forEach((cellRanges, rowIndex) => {
    cellRanges.forEach((cellRange, cellIndex) => {
      const segment = sourceXml.slice(cellRange.start, cellRange.end);
      const propertyContent = directPropertyChangeContent(
        segment,
        "w:tcPrChange"
      );
      if (!propertyContent) {
        return;
      }
      const insertedTag = findRowOrCellRevisionTag(
        propertyContent,
        "w:cellIns"
      );
      const deletedTag = findRowOrCellRevisionTag(
        propertyContent,
        "w:cellDel"
      );
      const tag = insertedTag ?? deletedTag;
      if (!tag) {
        return;
      }
      const cell = table.rows[rowIndex]?.cells[cellIndex];
      appendRevision(
        "cell",
        insertedTag ? "insertion" : "deletion",
        tag,
        (cell && cellTextSnippet(cell)) ?? "Table cell",
        rowIndex,
        cellIndex
      );
    });
  });

  // Table / row / cell property changes.
  const propertyChangeKinds: Array<{
    tagName: string;
    scope: TableTrackedRevisionScope;
    kind: DocxTrackedChangeKind;
  }> = [
    { tagName: "w:tblPrChange", scope: "table", kind: "table-format-change" },
    { tagName: "w:trPrChange", scope: "row", kind: "row-format-change" },
    { tagName: "w:tcPrChange", scope: "cell", kind: "cell-format-change" }
  ];

  propertyChangeKinds.forEach(({ tagName, scope, kind }) => {
    extractBalancedTagRanges(sourceXml, tagName).forEach((range) => {
      const rowIndex =
        scope === "table" ? undefined : rangeIndexContaining(rows, range.start);
      const cellIndex =
        rowIndex !== undefined
          ? rangeIndexContaining(cellsByRow[rowIndex] ?? [], range.start)
          : undefined;
      const changeXml = sourceXml.slice(range.start, range.end);
      appendRevision(
        scope,
        kind,
        range.openTag,
        summarizeTableFormattingChange(scope, changeXml),
        rowIndex === -1 ? undefined : rowIndex,
        cellIndex === -1 ? undefined : cellIndex
      );
    });
  });

  const resolved: TableTrackedMarkup | undefined =
    revisions.length > 0
      ? { revisions, rowChangeByIndex, cellChangeByIndex }
      : undefined;
  setCacheEntry(tableTrackedMarkupBySourceXml, sourceXml, resolved ?? null);
  return resolved;
}

/**
 * Resolves a table revision's gutter anchor. Row/cell revisions anchor at the
 * first paragraph of the real row/cell that contains the revision, falling
 * back to the table's first cell when the target cell has no paragraphs.
 */
export function tableRevisionAnchorLocation(
  table: TableNode,
  tableIndex: number,
  revision: TableTrackedRevision
): Extract<DocxTextRangeLocation, { kind: "table-cell" }> | undefined {
  const rowIndex = revision.rowIndex;
  if (rowIndex === undefined || !table.rows[rowIndex]) {
    return firstTableCellAnchorLocation(table, tableIndex);
  }

  const row = table.rows[rowIndex];
  const preferredCellIndex = Math.max(0, revision.cellIndex ?? 0);
  const preferredCell = row.cells[preferredCellIndex];
  if (preferredCell && tableCellParagraphs(preferredCell.nodes).length > 0) {
    return {
      kind: "table-cell",
      tableIndex,
      rowIndex,
      cellIndex: preferredCellIndex,
      paragraphIndex: 0
    };
  }

  for (let cellIndex = 0; cellIndex < row.cells.length; cellIndex += 1) {
    const cell = row.cells[cellIndex];
    if (cell && tableCellParagraphs(cell.nodes).length > 0) {
      return {
        kind: "table-cell",
        tableIndex,
        rowIndex,
        cellIndex,
        paragraphIndex: 0
      };
    }
  }

  return {
    kind: "table-cell",
    tableIndex,
    rowIndex,
    cellIndex: preferredCellIndex,
    paragraphIndex: 0
  };
}
