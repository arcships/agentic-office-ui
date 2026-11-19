import type { XlsxDefinedName } from "./types";
import {
  type ArchiveEntries,
  getLocalElements,
  parseXml,
  readArchiveText,
} from "./images";

/** Defined names of a workbook, split by their scope. */
export interface WorkbookDefinedNameMetadata {
  /** Names without a `localSheetId`, visible on every worksheet. */
  workbook: XlsxDefinedName[];
  /**
   * Names with a `localSheetId`, keyed by the workbook sheet slot
   * (chartsheets occupy slots too, matching Excel's `localSheetId`).
   */
  bySheetIndex: Map<number, XlsxDefinedName[]>;
}

const BUILTIN_NAME_PREFIX = "_xlnm.";

/**
 * Reads `xl/workbook.xml`'s `<definedNames>` list.
 *
 * Excel's built-in names (`_xlnm.Print_Area`, `_xlnm.Print_Titles`,
 * `_xlnm._FilterDatabase`, …) are omitted: they are worksheet plumbing rather
 * than user-visible names, and Name Manager does not list them either. The
 * print ones are parsed separately by `parseWorkbookPrintMetadata`.
 */
export function parseWorkbookDefinedNames(
  archive: ArchiveEntries,
  sheets: ReadonlyArray<{ name: string }>,
): WorkbookDefinedNameMetadata {
  const metadata: WorkbookDefinedNameMetadata = {
    bySheetIndex: new Map<number, XlsxDefinedName[]>(),
    workbook: [],
  };
  const xml = readArchiveText(archive, "xl/workbook.xml");
  const document = xml ? parseXml(xml) : null;
  if (!document) {
    return metadata;
  }

  for (const node of getLocalElements(document, "definedName")) {
    const name = (node.getAttribute("name") ?? "").trim();
    const range = (node.textContent ?? "").trim();
    if (!name || !range || name.toLowerCase().startsWith(BUILTIN_NAME_PREFIX)) {
      continue;
    }
    const hidden = node.getAttribute("hidden") === "1" || node.getAttribute("hidden")?.toLowerCase() === "true";
    const comment = node.getAttribute("comment")?.trim() || undefined;
    const rawLocalSheetId = node.getAttribute("localSheetId");
    if (rawLocalSheetId === null) {
      metadata.workbook.push({
        comment,
        hidden,
        name,
        range,
        scope: "workbook",
      });
      continue;
    }

    const localSheetId = Number(rawLocalSheetId);
    const sheetName = sheets[localSheetId]?.name;
    if (!Number.isInteger(localSheetId) || localSheetId < 0 || !sheetName) {
      continue;
    }
    const scoped = metadata.bySheetIndex.get(localSheetId) ?? [];
    scoped.push({
      comment,
      hidden,
      name,
      range,
      scope: "sheet",
      sheetName,
    });
    metadata.bySheetIndex.set(localSheetId, scoped);
  }

  return metadata;
}

/** Workbook-scoped names plus the names scoped to one worksheet. */
export function definedNamesForWorksheet(
  metadata: WorkbookDefinedNameMetadata,
  workbookSheetIndex: number,
): XlsxDefinedName[] {
  return [
    ...metadata.workbook,
    ...(metadata.bySheetIndex.get(workbookSheetIndex) ?? []),
  ];
}

/**
 * Converts the Name Manager's `XlsxNamedRange[]` view (engine names plus
 * in-memory edits) into the sheet model's defined-name shape.
 */
export function toDefinedNameSummary(
  entry: { comment?: string; hidden?: boolean; name: string; refersTo: string; sheetIndex?: number },
  sheets: ReadonlyArray<{ name: string; workbookSheetIndex: number }>,
): XlsxDefinedName {
  const sheet = entry.sheetIndex === undefined
    ? undefined
    : sheets.find((candidate) => candidate.workbookSheetIndex === entry.sheetIndex);
  return {
    comment: entry.comment,
    hidden: entry.hidden,
    name: entry.name,
    range: entry.refersTo,
    scope: entry.sheetIndex === undefined ? "workbook" : "sheet",
    ...(sheet ? { sheetName: sheet.name } : {}),
  };
}
