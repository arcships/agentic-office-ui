import { strToU8, unzipSync, zipSync } from "fflate";
import type { XlsxSheetVisibility } from "../types";
import {
  getChildElements,
  getFirstDescendant,
  normalizeArchivePath,
  parseXml,
  readArchiveText,
  serializeXml,
  type ArchiveEntries,
} from "../images";

function normalizeVisibility(value: string | null): XlsxSheetVisibility {
  return value === "hidden" || value === "veryHidden" ? value : "visible";
}

/**
 * Hides, unhides or very-hides a worksheet inside saved XLSX bytes.
 *
 * `@dukelib/sheets-wasm` 0.1.17 exposes `Worksheet.visibility` as read-only, so
 * the viewer changes visibility by rewriting the `state` attribute of the
 * sheet's `<sheet>` entry in `xl/workbook.xml` and parsing the result back into
 * a fresh `Workbook`. That keeps the workbook model, a later save and the
 * exported package in agreement.
 *
 * The sheet is addressed by name rather than by `<sheets>` position: the
 * workbook part interleaves worksheets and chartsheets, while the controller
 * works with worksheet indices, and a name is the only identifier shared by
 * both views. Excel keeps names unique case-insensitively, so an exact match
 * identifies one worksheet.
 *
 * @param bytes Saved XLSX bytes, as produced by `Workbook.saveXlsxBytes()`.
 * @param sheetName Name of the worksheet whose visibility changes.
 * @param visibility Visibility to write. `"visible"` removes the `state`
 *   attribute, matching how Excel writes a visible sheet.
 * @returns The rewritten bytes, the input bytes when the sheet already has the
 *   requested visibility, or `null` when the name is unknown, the visibility
 *   is not an Excel state, or the archive cannot be read.
 */
export function applyWorkbookSheetVisibility(
  bytes: Uint8Array,
  sheetName: string,
  visibility: XlsxSheetVisibility,
): Uint8Array | null {
  if (typeof sheetName !== "string" || !sheetName) {
    return null;
  }
  if (visibility !== "visible" && visibility !== "hidden" && visibility !== "veryHidden") {
    return null;
  }
  let archive: ArchiveEntries;
  try {
    archive = unzipSync(bytes);
  } catch {
    return null;
  }
  const workbookPath = normalizeArchivePath("xl/workbook.xml");
  const workbookXml = readArchiveText(archive, workbookPath);
  const document = workbookXml ? parseXml(workbookXml) : null;
  if (!document || !document.documentElement) {
    return null;
  }
  const sheetsElement = getFirstDescendant(document, "sheets");
  const sheetNodes = sheetsElement ? getChildElements(sheetsElement, "sheet") : [];
  const sheetNode = sheetNodes.find((node) => node.getAttribute("name") === sheetName);
  if (!sheetNode) {
    return null;
  }
  if (normalizeVisibility(sheetNode.getAttribute("state")) === visibility) {
    return bytes;
  }

  if (visibility === "visible") {
    sheetNode.removeAttribute("state");
  } else {
    sheetNode.setAttribute("state", visibility);
  }
  archive[workbookPath] = strToU8(serializeXml(document));
  return zipSync(archive, { level: 6 });
}
