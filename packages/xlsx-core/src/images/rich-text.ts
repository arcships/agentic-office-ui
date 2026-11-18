/**
 * Rich-text cells.
 *
 * A cell can store several inline runs with their own font, colour and
 * vertical alignment. Excel keeps them in `xl/sharedStrings.xml` (`<si><r>`)
 * or inline in the worksheet (`<c t="inlineStr"><is><r>`). The workbook engine
 * only exposes the concatenated string, so the runs are parsed here and travel
 * with the sheet state for renderers that paint per run.
 */
import type { XlsxCellTextRun, XlsxCellTextRunFont } from "../types";
import {
  type ArchiveEntries,
  getChildElements,
  getFirstChild,
  getLocalElements,
  parseXml,
  readArchiveText,
} from "./image-parser";
import { parseSpreadsheetFont } from "./theme-palette";

/** Path of the shared-string table inside an XLSX package. */
export const SHARED_STRINGS_PART = "xl/sharedStrings.xml";

/** An `<r>`/`<r …>` run element, optionally namespace-prefixed. */
const RICH_TEXT_RUN_ELEMENT = /<(?:[\w.-]+:)?r[\s/>]/;

/**
 * Runs of one `<si>`/`<is>` node.
 *
 * Returns `null` when the node is a plain `<t>` string or a single run without
 * formatting: both render exactly like the concatenated cell value, so keeping
 * them as runs would only cost memory.
 */
export function parseRichTextRuns(node: Element | null): XlsxCellTextRun[] | null {
  if (!node) {
    return null;
  }

  const runs: XlsxCellTextRun[] = [];
  for (const runNode of getChildElements(node, "r")) {
    const text = getFirstChild(runNode, "t")?.textContent ?? "";
    const font = toRunFont(parseSpreadsheetFont(getFirstChild(runNode, "rPr")));
    runs.push(font ? { text, font } : { text });
  }

  if (runs.length === 0) {
    return null;
  }

  const hasFormatting = runs.some((run) => run.font !== undefined);
  return hasFormatting ? runs : null;
}

/**
 * Narrows a resolved `rPr` font to the public run-font shape.
 *
 * `parseSpreadsheetFont` also returns unrelated fields (`family`, `scheme`,
 * shadow); only the properties a renderer can express per run survive.
 */
function toRunFont(raw: Record<string, unknown> | undefined): XlsxCellTextRunFont | undefined {
  if (!raw) {
    return undefined;
  }

  const font: XlsxCellTextRunFont = {};
  if (raw.bold === true) font.bold = true;
  if (raw.italic === true) font.italic = true;
  if (raw.strikethrough === true) font.strikethrough = true;
  if (typeof raw.underline === "string" && raw.underline) font.underline = raw.underline;
  if (typeof raw.size === "number" && Number.isFinite(raw.size)) font.sizePt = raw.size;
  if (typeof raw.name === "string" && raw.name) font.name = raw.name;
  if (typeof raw.verticalAlign === "string" && raw.verticalAlign) {
    font.verticalAlign = raw.verticalAlign;
  }
  if (raw.color && typeof raw.color === "object") {
    font.color = raw.color as Record<string, unknown>;
  }

  return Object.keys(font).length > 0 ? font : undefined;
}

/**
 * Parses the shared-string table into per-index runs.
 *
 * Plain strings stay `null` so callers can keep using the concatenated value
 * without allocating runs for every shared string in the workbook.
 */
export function parseSharedStringRichText(
  archive: ArchiveEntries
): Array<XlsxCellTextRun[] | null> {
  const xml = readArchiveText(archive, SHARED_STRINGS_PART);
  if (!xml) {
    return [];
  }

  if (!RICH_TEXT_RUN_ELEMENT.test(xml)) {
    // Workbooks without rich text never pay for a second DOM parse of the
    // shared-string table.
    return [];
  }

  const document = parseXml(xml);
  if (!document) {
    return [];
  }

  return getLocalElements(document, "si").map((item) => parseRichTextRuns(item));
}

/**
 * Runs of a `t="inlineStr"` cell node (`<c><is>…</is></c>`).
 */
export function parseInlineRichTextCell(cellNode: Element): XlsxCellTextRun[] | null {
  return parseRichTextRuns(getFirstChild(cellNode, "is"));
}

/** Concatenated text of a run list, matching the engine's display value. */
export function richTextRunsToPlainText(runs: readonly XlsxCellTextRun[] | null | undefined): string {
  if (!runs || runs.length === 0) {
    return "";
  }
  let text = "";
  for (const run of runs) {
    text += run.text;
  }
  return text;
}

/**
 * Runs of a cell, but only while they still describe `value`.
 *
 * A cell edit, a recalculation or a paste replaces the rich string with plain
 * text while the parsed runs stay behind in the sheet state. Comparing the
 * concatenation keeps stale runs from painting over the new value.
 */
export function resolveRichTextRuns(
  runs: readonly XlsxCellTextRun[] | null | undefined,
  value: string
): XlsxCellTextRun[] | undefined {
  if (!runs || runs.length < 2) {
    return undefined;
  }
  if (richTextRunsToPlainText(runs) !== value) {
    return undefined;
  }
  return runs as XlsxCellTextRun[];
}
