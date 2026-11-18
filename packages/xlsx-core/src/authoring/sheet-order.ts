import { strToU8, unzipSync, zipSync } from "fflate";
import {
  getChildElements,
  getFirstDescendant,
  isElementNode,
  normalizeArchivePath,
  parseXml,
  readArchiveText,
  serializeXml,
  type ArchiveEntries,
} from "../images";

// ── Saved-bytes worksheet reorder ────────────────────────────────────────────

function isPermutation(order: readonly number[], length: number) {
  if (order.length !== length) {
    return false;
  }
  const seen = new Set<number>();
  for (const value of order) {
    if (!Number.isInteger(value) || value < 0 || value >= length || seen.has(value)) {
      return false;
    }
    seen.add(value);
  }
  return true;
}

/**
 * Remaps every sheet index the workbook part keeps beside the `<sheets>` list.
 *
 * `localSheetId` on defined names, and `activeTab`/`firstSheet` on workbook
 * views, are positions in the sheet list, not `sheetId`s. Reordering the
 * `<sheet>` elements without moving them would silently retarget sheet-scoped
 * names or make the file open on a different tab.
 */
function remapWorkbookSheetIndexes(document: Document, order: readonly number[]) {
  const newPositionByOldPosition = new Map<number, number>();
  order.forEach((oldPosition, newPosition) => {
    newPositionByOldPosition.set(oldPosition, newPosition);
  });

  const remapAttribute = (element: Element, attribute: string) => {
    if (!element.hasAttribute(attribute)) {
      return;
    }
    const raw = element.getAttribute(attribute);
    const value = raw === null ? Number.NaN : Number.parseInt(raw, 10);
    const mapped = Number.isInteger(value) ? newPositionByOldPosition.get(value) : undefined;
    if (mapped !== undefined) {
      element.setAttribute(attribute, String(mapped));
    }
  };

  for (const node of Array.from(document.getElementsByTagName("*"))) {
    if (!isElementNode(node)) {
      continue;
    }
    remapAttribute(node, "localSheetId");
    remapAttribute(node, "activeTab");
    remapAttribute(node, "firstSheet");
  }
}

/**
 * Reorders the worksheets and chartsheets inside saved XLSX bytes.
 *
 * `@dukelib/sheets-wasm` 0.1.17 exposes no reorder setter (`Workbook` has
 * `addSheet`/`removeSheet`/`sheetNames`/`sheetOrder` but no order write), so
 * the viewer reorders by rewriting the workbook part of its own saved bytes
 * and parsing the result back into a fresh `Workbook`. This keeps the workbook
 * model, the exported package and a later reload in agreement, and every
 * per-sheet asset (worksheet part path, images, charts, sheet state) follows
 * its sheet because it is re-read from the reordered package.
 *
 * The order is a permutation of the current `<sheet>` slots; chartsheets are
 * interleaved with worksheets exactly as in `Workbook.sheetOrder`. Sheet-scoped
 * defined names (`localSheetId`) and workbook-view tabs (`activeTab`,
 * `firstSheet`) are remapped so they keep pointing at the same sheet.
 *
 * @param bytes Saved XLSX bytes, as produced by `Workbook.saveXlsxBytes()`.
 * @param order New slot order: `order[newPosition] = currentPosition`. Every
 *   current position must appear exactly once.
 * @returns The rewritten bytes, the input bytes for an identity order, or
 *   `null` when the order is not a permutation of the workbook's sheets or the
 *   archive cannot be read.
 */
export function applyWorkbookSheetOrder(
  bytes: Uint8Array,
  order: readonly number[],
): Uint8Array | null {
  if (!Array.isArray(order) || order.length === 0) {
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
  if (!sheetsElement || !isPermutation(order, sheetNodes.length)) {
    return null;
  }
  if (order.every((value, index) => value === index)) {
    return bytes;
  }

  // Appending an existing child moves it, so appending in the requested order
  // rewrites the list without cloning any node and keeps its attributes.
  for (const currentPosition of order) {
    const sheetNode = sheetNodes[currentPosition];
    if (sheetNode) {
      sheetsElement.appendChild(sheetNode);
    }
  }
  remapWorkbookSheetIndexes(document, order);
  archive[workbookPath] = strToU8(serializeXml(document));
  return zipSync(archive, { level: 6 });
}
