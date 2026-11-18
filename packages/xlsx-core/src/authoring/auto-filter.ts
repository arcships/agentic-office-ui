import { strToU8, unzipSync, zipSync } from "fflate";
import {
  getChildElements,
  getLocalElements,
  getRelationshipId,
  parseRelationships,
  parseWorkbookSheets,
  parseXml,
  readArchiveText,
  relsPathForDocument,
  serializeXml,
  SPREADSHEET_NS,
  type ArchiveEntries
} from "../images";
import type { XlsxAutoFilterColumnState } from "../types";

/**
 * Value filters the viewer applied or cleared, keyed by workbook sheet index
 * and AutoFilter range reference.
 */
export type WorkbookAutoFilterOverrides = Readonly<
  Record<number, Readonly<Record<string, readonly XlsxAutoFilterColumnState[]>>>
>;

function createSpreadsheetElement(document: XMLDocument, localName: string) {
  return document.createElementNS(SPREADSHEET_NS, localName);
}

/**
 * Replaces an `<autoFilter>` element's criteria.
 *
 * Existing `filterColumn` children are removed first, so clearing a filter
 * writes a bare `<autoFilter ref="…"/>` instead of leaving the file's original
 * values behind.
 */
function writeAutoFilterColumns(
  document: XMLDocument,
  autoFilterNode: Element,
  columns: readonly XlsxAutoFilterColumnState[]
) {
  for (const existing of getChildElements(autoFilterNode, "filterColumn")) {
    autoFilterNode.removeChild(existing);
  }
  for (const column of columns) {
    if (column.values.length === 0 && !column.includeBlank) {
      continue;
    }
    const columnNode = createSpreadsheetElement(document, "filterColumn");
    columnNode.setAttribute("colId", String(column.columnIndex));
    const filtersNode = createSpreadsheetElement(document, "filters");
    if (column.includeBlank) {
      filtersNode.setAttribute("blank", "1");
    }
    for (const value of column.values) {
      const filterNode = createSpreadsheetElement(document, "filter");
      filterNode.setAttribute("val", value);
      filtersNode.appendChild(filterNode);
    }
    columnNode.appendChild(filtersNode);
    autoFilterNode.appendChild(columnNode);
  }
}

function findOverrideColumns(
  reference: string,
  overrides: Readonly<Record<string, readonly XlsxAutoFilterColumnState[]>>
) {
  if (Object.prototype.hasOwnProperty.call(overrides, reference)) {
    return overrides[reference] ?? [];
  }
  // A worksheet may declare several ranges in one `ref` sqref; the override is
  // keyed by the individual range the header menu belongs to.
  const tokens = reference.split(/\s+/);
  for (const token of tokens) {
    if (Object.prototype.hasOwnProperty.call(overrides, token)) {
      return overrides[token] ?? [];
    }
  }
  return null;
}

function patchAutoFilterElement(
  document: XMLDocument,
  overrides: Readonly<Record<string, readonly XlsxAutoFilterColumnState[]>>
) {
  const autoFilterNode = getLocalElements(document, "autoFilter")[0];
  if (!autoFilterNode) {
    return false;
  }
  const columns = findOverrideColumns(autoFilterNode.getAttribute("ref") ?? "", overrides);
  if (columns === null) {
    return false;
  }
  writeAutoFilterColumns(document, autoFilterNode, columns);
  return true;
}

function applyToSheetArchive(
  archive: ArchiveEntries,
  sheetPath: string,
  overrides: Readonly<Record<string, readonly XlsxAutoFilterColumnState[]>>
) {
  const sheetXml = readArchiveText(archive, sheetPath);
  const sheetDocument = sheetXml ? parseXml(sheetXml) : null;
  if (!sheetDocument) {
    return false;
  }

  let changed = patchAutoFilterElement(sheetDocument, overrides);

  // Table filters live in `xl/tables/tableN.xml`, next to the table itself.
  const relationships = parseRelationships(archive, relsPathForDocument(sheetPath), sheetPath);
  for (const tablePart of getLocalElements(sheetDocument, "tablePart")) {
    const relationshipId = getRelationshipId(tablePart);
    const relationship = relationshipId ? relationships.get(relationshipId) : undefined;
    const tableXml = relationship ? readArchiveText(archive, relationship.target) : null;
    const tableDocument = tableXml ? parseXml(tableXml) : null;
    if (!relationship || !tableDocument) {
      continue;
    }
    if (patchAutoFilterElement(tableDocument, overrides)) {
      archive[relationship.target] = strToU8(serializeXml(tableDocument));
      changed = true;
    }
  }

  if (changed) {
    archive[sheetPath] = strToU8(serializeXml(sheetDocument));
  }
  return changed;
}

/**
 * Writes the viewer's applied and cleared AutoFilter criteria into saved bytes.
 *
 * The filter itself is view state, but Excel stores the criteria in the
 * worksheet and table parts. Exporting without this patch would resurrect the
 * file's original filter (or lose a cleared one) on reload.
 *
 * A workbook without overrides, or one whose parts cannot be parsed, is
 * returned unchanged.
 */
export function applyWorkbookAutoFilterOverrides(
  bytes: Uint8Array,
  overridesByWorkbookSheetIndex: WorkbookAutoFilterOverrides
): Uint8Array {
  const sheetIndexes = Object.keys(overridesByWorkbookSheetIndex);
  if (sheetIndexes.length === 0) {
    return bytes;
  }

  const archive = unzipSync(bytes);
  const worksheetSheets = parseWorkbookSheets(archive).filter(
    (sheet) => sheet.kind === "worksheet"
  );
  let changed = false;
  for (const sheetIndex of sheetIndexes) {
    const worksheetIndex = Number.parseInt(sheetIndex, 10);
    const overrides = overridesByWorkbookSheetIndex[worksheetIndex];
    if (!overrides || Object.keys(overrides).length === 0) {
      continue;
    }
    const sheet = worksheetSheets[worksheetIndex];
    if (!sheet) {
      continue;
    }
    if (applyToSheetArchive(archive, sheet.path, overrides)) {
      changed = true;
    }
  }
  return changed ? zipSync(archive) : bytes;
}
