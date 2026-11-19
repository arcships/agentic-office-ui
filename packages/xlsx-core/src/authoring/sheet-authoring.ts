import { strToU8 } from "fflate";
import type {
  XlsxCellAddress,
  XlsxCellComment,
  XlsxCellCommentInput,
  XlsxCellHyperlink,
  XlsxCellHyperlinkInput,
  XlsxCellRange,
  XlsxConditionalFormatDraft,
  XlsxConditionalFormatRule,
  XlsxDataValidation,
  XlsxFreezePanes,
  XlsxTable,
  XlsxThemePalette,
} from "../types";
import { columnLabel } from "../core";
import {
  CONTENT_TYPES_NS,
  HYPERLINK_REL_TYPE,
  PKG_REL_NS,
  REL_NS,
  SPREADSHEET_NS,
  VML_DRAWING_REL_TYPE,
  type ArchiveEntries,
  type XlsxImageAttachment,
  type WorkbookImageSheetOrigin,
  type WorkbookTableMetadata,
  parseSheetState,
  parseWorksheetFreezePanesDocument,
  parseWorksheetTableMetadata,
  getChildElements,
  getFirstChild,
  getLocalElements,
  getRelationshipId,
  normalizeArchivePath,
  parseA1RangeReference,
  parseXml,
  readArchiveText,
  relativeArchivePath,
  relsPathForDocument,
  serializeXml,
} from "../images";

// ── Types ────────────────────────────────────────────────────────────────────

/**
 * Minimal workbook state the sheet-authoring mutations need: the source
 * archive and the worksheet part paths. `WorkbookImageAssets` satisfies it.
 */
export type WorkbookAuthoringAssets = {
  archive: ArchiveEntries;
  sheetOrigins?: Array<WorkbookImageSheetOrigin | null>;
  sheetPathsByWorkbookSheetIndex: string[];
  tableMetadataByWorkbookSheetIndex?: WorkbookTableMetadata[][];
  themePalette?: XlsxThemePalette;
};

export type { XlsxCellCommentInput, XlsxCellHyperlinkInput } from "../types";

export type XlsxTableInsertInput = {
  /** Header labels, one per column of `range`. Missing names fall back to `Column N`. */
  columnNames: string[];
  displayName?: string;
  /** First row of `range` holds the column headers. Defaults to `true`. */
  hasHeaders?: boolean;
  name?: string;
  range: XlsxCellRange;
  showColumnStripes?: boolean;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  showRowStripes?: boolean;
  styleName?: string;
};

export type WorkbookSheetAuthoringState = {
  comments: XlsxCellComment[];
  conditionalFormatRules: XlsxConditionalFormatRule[];
  dataValidations: XlsxDataValidation[];
  freezePanes: XlsxFreezePanes | null;
  hyperlinks: XlsxCellHyperlink[];
  tables: XlsxTable[];
};

const REL_NS_DECLARATION = "http://www.w3.org/2000/xmlns/";
const COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const TABLE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
const COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
const TABLE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
const VML_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.vmlDrawing";

/**
 * OOXML worksheet child order. Mutations insert new elements in schema order so
 * Excel and LibreOffice accept the patched part without a repair pass.
 */
const WORKSHEET_CHILD_ORDER = [
  "sheetPr",
  "dimension",
  "sheetViews",
  "sheetFormatPr",
  "cols",
  "sheetData",
  "sheetCalcPr",
  "sheetProtection",
  "protectedRanges",
  "scenarios",
  "autoFilter",
  "sortState",
  "dataConsolidate",
  "customSheetViews",
  "mergeCells",
  "phoneticPr",
  "conditionalFormatting",
  "dataValidations",
  "hyperlinks",
  "printOptions",
  "pageMargins",
  "pageSetup",
  "headerFooter",
  "rowBreaks",
  "colBreaks",
  "customProperties",
  "cellWatches",
  "ignoredErrors",
  "smartTags",
  "drawing",
  "legacyDrawing",
  "legacyDrawingHF",
  "picture",
  "oleObjects",
  "controls",
  "webPublishItems",
  "tableParts",
  "pivotTableParts",
  "extLst",
];

// ── Small helpers ────────────────────────────────────────────────────────────

export function cellAddressToA1(cell: XlsxCellAddress) {
  return `${columnLabel(cell.col)}${cell.row + 1}`;
}

function normalizeCell(cell: XlsxCellAddress): XlsxCellAddress {
  return {
    col: Math.max(0, Math.floor(Number(cell.col) || 0)),
    row: Math.max(0, Math.floor(Number(cell.row) || 0)),
  };
}

function normalizeRange(range: XlsxCellRange): XlsxCellRange {
  const start = normalizeCell(range.start);
  const end = normalizeCell(range.end);
  return {
    start: { col: Math.min(start.col, end.col), row: Math.min(start.row, end.row) },
    end: { col: Math.max(start.col, end.col), row: Math.max(start.row, end.row) }
  };
}

function rangeReference(range: XlsxCellRange) {
  const normalized = normalizeRange(range);
  const start = cellAddressToA1(normalized.start);
  const end = cellAddressToA1(normalized.end);
  return start === end ? start : `${start}:${end}`;
}

export function escapeXmlText(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function resolveSheetPath(assets: WorkbookAuthoringAssets, workbookSheetIndex: number) {
  if (!Number.isInteger(workbookSheetIndex) || workbookSheetIndex < 0) {
    return null;
  }
  return assets.sheetPathsByWorkbookSheetIndex[workbookSheetIndex] ?? null;
}

function readSheetDocument(assets: WorkbookAuthoringAssets, workbookSheetIndex: number) {
  const sheetPath = resolveSheetPath(assets, workbookSheetIndex);
  if (!sheetPath) {
    return null;
  }
  const xml = readArchiveText(assets.archive, sheetPath);
  const document = xml ? parseXml(xml) : null;
  if (!document?.documentElement) {
    return null;
  }
  return { document, sheetPath };
}

function writeSheetDocument(
  assets: WorkbookAuthoringAssets,
  sheetPath: string,
  document: XMLDocument
) {
  assets.archive[normalizeArchivePath(sheetPath)] = strToU8(serializeXml(document));
}

function createSheetElement(document: XMLDocument, localName: string) {
  return document.createElementNS(SPREADSHEET_NS, localName);
}

function removeSheetChildren(document: XMLDocument, localName: string) {
  for (const node of Array.from(document.getElementsByTagName("*"))) {
    if (node.localName === localName && node.parentNode) {
      node.parentNode.removeChild(node);
    }
  }
}

function insertSheetChild(document: XMLDocument, node: Element) {
  const root = document.documentElement;
  if (!root) {
    return;
  }
  const rank = WORKSHEET_CHILD_ORDER.indexOf(node.localName);
  if (rank < 0) {
    root.appendChild(node);
    return;
  }
  for (const child of Array.from(root.childNodes)) {
    if (child.nodeType !== 1) {
      continue;
    }
    const childRank = WORKSHEET_CHILD_ORDER.indexOf((child as Element).localName);
    if (childRank > rank) {
      root.insertBefore(node, child);
      return;
    }
  }
  root.appendChild(node);
}

function ensureSheetChild(document: XMLDocument, localName: string) {
  const existing = getLocalElements(document, localName)[0];
  if (existing) {
    return existing;
  }
  const node = createSheetElement(document, localName);
  insertSheetChild(document, node);
  return node;
}

function ensureRelationshipNamespace(document: XMLDocument) {
  const root = document.documentElement;
  if (!root || root.getAttribute("xmlns:r")) {
    return;
  }
  root.setAttributeNS(REL_NS_DECLARATION, "xmlns:r", REL_NS);
}

function setRelationshipId(element: Element, relationshipId: string) {
  ensureRelationshipNamespace(element.ownerDocument);
  element.setAttributeNS(REL_NS, "r:id", relationshipId);
}

function ensureRelationshipsDocument(archive: ArchiveEntries, relsPath: string) {
  const existing = readArchiveText(archive, relsPath);
  if (existing) {
    const document = parseXml(existing);
    if (document?.documentElement) {
      return document;
    }
  }
  return parseXml(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}"></Relationships>`
  );
}

function writeRelationshipsDocument(archive: ArchiveEntries, relsPath: string, document: XMLDocument) {
  archive[normalizeArchivePath(relsPath)] = strToU8(serializeXml(document));
}

function nextRelationshipId(document: XMLDocument, prefix: string) {
  const existingIds = new Set(
    getLocalElements(document, "Relationship")
      .map((node) => node.getAttribute("Id"))
      .filter((value): value is string => Boolean(value))
  );
  let index = 1;
  while (existingIds.has(`rId${prefix}${index}`)) {
    index += 1;
  }
  return `rId${prefix}${index}`;
}

function createRelationship(
  document: XMLDocument,
  type: string,
  target: string,
  targetMode?: "External"
) {
  const root = document.documentElement;
  if (!root) {
    return null;
  }
  const id = nextRelationshipId(document, "XlsxAuthoring");
  const node = document.createElementNS(PKG_REL_NS, "Relationship");
  node.setAttribute("Id", id);
  node.setAttribute("Type", type);
  node.setAttribute("Target", target);
  if (targetMode) {
    node.setAttribute("TargetMode", targetMode);
  }
  root.appendChild(node);
  return id;
}

function removeRelationshipById(document: XMLDocument, relationshipId: string) {
  const node = getLocalElements(document, "Relationship").find(
    (candidate) => candidate.getAttribute("Id") === relationshipId
  );
  node?.parentNode?.removeChild(node);
}

type RawRelationship = { target: string; targetMode: string | null; type: string };

function readRelationships(
  archive: ArchiveEntries,
  relsPath: string,
  sheetPath: string
): Map<string, RawRelationship> {
  const xml = readArchiveText(archive, relsPath);
  const document = xml ? parseXml(xml) : null;
  if (!document) {
    return new Map();
  }
  const relationships = new Map<string, RawRelationship>();
  for (const node of getLocalElements(document, "Relationship")) {
    const id = node.getAttribute("Id");
    const target = node.getAttribute("Target");
    const type = node.getAttribute("Type");
    if (!id || !target || !type) {
      continue;
    }
    relationships.set(id, {
      target: target.startsWith("/")
        ? normalizeArchivePath(target)
        : target.startsWith("#")
          ? target
          : resolveRelativeTarget(sheetPath, target),
      targetMode: node.getAttribute("TargetMode"),
      type
    });
  }
  return relationships;
}

/** Resolve a relative package target against the owning part path. */
function resolveRelativeTarget(ownerPath: string, target: string) {
  if (target.startsWith("/")) {
    return normalizeArchivePath(target);
  }
  const ownerSegments = normalizeArchivePath(ownerPath).split("/").slice(0, -1);
  for (const segment of target.replace(/\\/g, "/").split("/")) {
    if (!segment || segment === ".") {
      continue;
    }
    if (segment === "..") {
      ownerSegments.pop();
      continue;
    }
    ownerSegments.push(segment);
  }
  return normalizeArchivePath(ownerSegments.join("/"));
}

function ensureContentTypesDocument(archive: ArchiveEntries) {
  const existing = readArchiveText(archive, "[Content_Types].xml");
  if (existing) {
    const document = parseXml(existing);
    if (document?.documentElement) {
      return document;
    }
  }
  return parseXml(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${CONTENT_TYPES_NS}"></Types>`
  );
}

function writeContentTypesDocument(archive: ArchiveEntries, document: XMLDocument) {
  archive["[Content_Types].xml"] = strToU8(serializeXml(document));
}

function ensureContentTypeOverride(archive: ArchiveEntries, partPath: string, contentType: string) {
  const document = ensureContentTypesDocument(archive);
  const root = document?.documentElement;
  if (!document || !root) {
    return false;
  }
  const partName = `/${normalizeArchivePath(partPath)}`;
  const existing = getLocalElements(document, "Override").find(
    (node) => node.getAttribute("PartName") === partName
  );
  if (existing) {
    existing.setAttribute("ContentType", contentType);
  } else {
    const node = document.createElementNS(CONTENT_TYPES_NS, "Override");
    node.setAttribute("PartName", partName);
    node.setAttribute("ContentType", contentType);
    root.appendChild(node);
  }
  writeContentTypesDocument(archive, document);
  return true;
}

function ensureContentTypeDefault(archive: ArchiveEntries, extension: string, contentType: string) {
  const document = ensureContentTypesDocument(archive);
  const root = document?.documentElement;
  if (!document || !root) {
    return false;
  }
  const existing = getLocalElements(document, "Default").find(
    (node) => (node.getAttribute("Extension") ?? "").toLowerCase() === extension.toLowerCase()
  );
  if (!existing) {
    const node = document.createElementNS(CONTENT_TYPES_NS, "Default");
    node.setAttribute("Extension", extension);
    node.setAttribute("ContentType", contentType);
    root.appendChild(node);
  }
  writeContentTypesDocument(archive, document);
  return true;
}

function nextPartPath(archive: ArchiveEntries, directory: string, baseName: string, extension: string) {
  let index = 1;
  while (archive[`${directory}/${baseName}${index}.${extension}`]) {
    index += 1;
  }
  return `${directory}/${baseName}${index}.${extension}`;
}

function parseWorksheetState(assets: WorkbookAuthoringAssets, sheetPath: string) {
  return parseSheetState(assets.archive, sheetPath, {
    includeCachedFormulaValues: false,
    themePalette: assets.themePalette ?? null
  });
}

// ── Freeze panes ─────────────────────────────────────────────────────────────

/**
 * Freezes rows/columns above and left of `freeze` on a worksheet, or clears the
 * frozen pane when both split counts are zero or `null` is passed.
 *
 * @returns `false` when the worksheet part is unavailable (non-XML load modes).
 */
export function setWorkbookFreezePanes(
  assets: WorkbookAuthoringAssets,
  workbookSheetIndex: number,
  freeze: XlsxFreezePanes | null
): boolean {
  const sheet = readSheetDocument(assets, workbookSheetIndex);
  if (!sheet) {
    return false;
  }
  const { document, sheetPath } = sheet;
  const target = freeze
    ? {
        col: Math.max(0, Math.floor(Number(freeze.col) || 0)),
        row: Math.max(0, Math.floor(Number(freeze.row) || 0))
      }
    : null;
  const sheetViews = ensureSheetChild(document, "sheetViews");
  let sheetView = getFirstChild(sheetViews, "sheetView");
  if (!sheetView) {
    sheetView = createSheetElement(document, "sheetView");
    sheetView.setAttribute("workbookViewId", "0");
    sheetViews.appendChild(sheetView);
  }
  for (const existing of getChildElements(sheetView, "pane")) {
    sheetView.removeChild(existing);
  }
  for (const selection of getChildElements(sheetView, "selection")) {
    selection.removeAttribute("pane");
  }
  if (target && (target.row > 0 || target.col > 0)) {
    const activePane = target.row > 0 && target.col > 0
      ? "bottomRight"
      : target.row > 0
        ? "bottomLeft"
        : "topRight";
    const pane = createSheetElement(document, "pane");
    if (target.col > 0) {
      pane.setAttribute("xSplit", String(target.col));
    }
    if (target.row > 0) {
      pane.setAttribute("ySplit", String(target.row));
    }
    pane.setAttribute("topLeftCell", cellAddressToA1(target));
    pane.setAttribute("activePane", activePane);
    pane.setAttribute("state", "frozen");
    sheetView.insertBefore(pane, sheetView.firstChild);
    const selection = getFirstChild(sheetView, "selection");
    if (selection) {
      selection.setAttribute("pane", activePane);
    }
  }
  writeSheetDocument(assets, sheetPath, document);
  return true;
}

// ── Hyperlinks ───────────────────────────────────────────────────────────────

function splitInternalTarget(target: string) {
  const trimmed = target.trim();
  return trimmed.startsWith("#") ? trimmed.slice(1) : null;
}

function removeHyperlinkNodesForCell(document: XMLDocument, cell: XlsxCellAddress) {
  const reference = cellAddressToA1(cell).toUpperCase();
  for (const node of getLocalElements(document, "hyperlink")) {
    if ((node.getAttribute("ref") ?? "").toUpperCase() === reference) {
      node.parentNode?.removeChild(node);
    }
  }
}

/**
 * Creates, updates or removes the hyperlink on a single cell.
 *
 * `#Sheet!A1` targets become an internal `location`; every other target becomes
 * an external relationship so Excel keeps it as a link.
 *
 * @returns The updated hyperlink list, or `null` when the sheet part is unavailable.
 */
export function setWorkbookCellHyperlink(
  assets: WorkbookAuthoringAssets,
  workbookSheetIndex: number,
  cell: XlsxCellAddress,
  input: XlsxCellHyperlinkInput | null
): XlsxCellHyperlink[] | null {
  const sheet = readSheetDocument(assets, workbookSheetIndex);
  if (!sheet) {
    return null;
  }
  const { document, sheetPath } = sheet;
  const address = normalizeCell(cell);
  const target = input?.target.trim() ?? "";
  if (input && !target) {
    return null;
  }
  const relsPath = relsPathForDocument(sheetPath);
  const relationshipsDocument = ensureRelationshipsDocument(assets.archive, relsPath);
  if (!relationshipsDocument?.documentElement) {
    return null;
  }

  removeHyperlinkNodesForCell(document, address);
  if (input) {
    const hyperlinks = ensureSheetChild(document, "hyperlinks");
    const node = createSheetElement(document, "hyperlink");
    node.setAttribute("ref", cellAddressToA1(address));
    const internalTarget = splitInternalTarget(target);
    if (internalTarget) {
      node.setAttribute("location", internalTarget);
    } else {
      const relationshipId = createRelationship(relationshipsDocument, HYPERLINK_REL_TYPE, target, "External");
      if (!relationshipId) {
        return null;
      }
      setRelationshipId(node, relationshipId);
    }
    if (input.tooltip?.trim()) {
      node.setAttribute("tooltip", input.tooltip.trim());
    }
    hyperlinks.appendChild(node);
  }

  const remaining = getLocalElements(document, "hyperlink");
  if (remaining.length === 0) {
    removeSheetChildren(document, "hyperlinks");
  } else {
    // CT_Hyperlinks does not declare `count`; keep the part readable by strict
    // OOXML consumers such as openpyxl.
    getLocalElements(document, "hyperlinks")[0]?.removeAttribute("count");
  }

  // Drop relationships no hyperlink references any more so export stays clean.
  const referencedIds = new Set(
    getLocalElements(document, "hyperlink")
      .map((node) => getRelationshipId(node))
      .filter((value): value is string => Boolean(value))
  );
  for (const node of getLocalElements(relationshipsDocument, "Relationship")) {
    if (node.getAttribute("Type") !== HYPERLINK_REL_TYPE) {
      continue;
    }
    const id = node.getAttribute("Id");
    if (id && !referencedIds.has(id)) {
      node.parentNode?.removeChild(node);
    }
  }

  writeSheetDocument(assets, sheetPath, document);
  writeRelationshipsDocument(assets.archive, relsPath, relationshipsDocument);
  return parseWorksheetState(assets, sheetPath)?.hyperlinks ?? [];
}

// ── Comments ─────────────────────────────────────────────────────────────────

function commentKey(cell: XlsxCellAddress) {
  return `${cell.row}:${cell.col}`;
}

function readCommentEntries(assets: WorkbookAuthoringAssets, sheetPath: string) {
  const comments = parseWorksheetState(assets, sheetPath)?.comments ?? [];
  const entries = new Map<string, XlsxCellComment>();
  for (const comment of comments) {
    entries.set(commentKey(comment.cell), {
      author: comment.author,
      cell: comment.cell,
      text: comment.text
    });
  }
  return entries;
}

function buildCommentsXml(entries: XlsxCellComment[]) {
  const authors: string[] = [];
  const authorIndex = new Map<string, number>();
  const rows = entries
    .slice()
    .sort((left, right) => left.cell.row - right.cell.row || left.cell.col - right.cell.col)
    .map((comment) => {
      const author = comment.author?.trim() || "";
      let index = authorIndex.get(author);
      if (index === undefined) {
        index = authors.length;
        authors.push(author);
        authorIndex.set(author, index);
      }
      const authorAttribute = author ? ` authorId="${index}"` : "";
      return `<comment ref="${cellAddressToA1(comment.cell)}"${authorAttribute}><text><r><t xml:space="preserve">${escapeXmlText(comment.text)}</t></r></text></comment>`;
    });
  const authorNodes = authors.map((author) => `<author>${escapeXmlText(author)}</author>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="${SPREADSHEET_NS}"><authors>${authorNodes}</authors><commentList>${rows.join("")}</commentList></comments>`;
}

function buildCommentsVmlXml(entries: XlsxCellComment[]) {
  const shapes = entries
    .slice()
    .sort((left, right) => left.cell.row - right.cell.row || left.cell.col - right.cell.col)
    .map((comment, index) => {
      const shapeId = 1025 + index;
      const column = comment.cell.col;
      const row = comment.cell.row;
      const anchor = `${column}, 15, ${Math.max(0, row - 1)}, 2, ${column + 2}, 15, ${row + 3}, 4`;
      return `<v:shape id="_x0000_s${shapeId}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${anchor}</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${row}</x:Row><x:Column>${column}</x:Column></x:ClientData></v:shape>`;
    });
  return `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>${shapes.join("")}</xml>`;
}

function removeCommentParts(
  assets: WorkbookAuthoringAssets,
  sheetPath: string,
  document: XMLDocument,
  relationshipsDocument: XMLDocument
) {
  const relsPath = relsPathForDocument(sheetPath);
  const relationships = readRelationships(assets.archive, relsPath, sheetPath);
  for (const [id, relationship] of relationships) {
    if (relationship.type !== COMMENTS_REL_TYPE && relationship.type !== VML_DRAWING_REL_TYPE) {
      continue;
    }
    delete assets.archive[relationship.target];
    removeRelationshipById(relationshipsDocument, id);
  }
  removeSheetChildren(document, "legacyDrawing");
}

/**
 * Creates, edits or deletes the note attached to a single cell.
 *
 * The comments part and the legacy VML drawing are written together so Excel
 * shows the note indicator and the comment popup after export.
 *
 * @returns The updated comment list, or `null` when the sheet part is unavailable.
 */
export function setWorkbookCellComment(
  assets: WorkbookAuthoringAssets,
  workbookSheetIndex: number,
  cell: XlsxCellAddress,
  input: XlsxCellCommentInput | null
): XlsxCellComment[] | null {
  const sheet = readSheetDocument(assets, workbookSheetIndex);
  if (!sheet) {
    return null;
  }
  const { document, sheetPath } = sheet;
  const address = normalizeCell(cell);
  const text = input?.text?.trim() ?? "";
  if (input && !text) {
    return null;
  }
  const relsPath = relsPathForDocument(sheetPath);
  const relationshipsDocument = ensureRelationshipsDocument(assets.archive, relsPath);
  if (!relationshipsDocument?.documentElement) {
    return null;
  }

  const entries = readCommentEntries(assets, sheetPath);
  if (input) {
    entries.set(commentKey(address), {
      author: input.author?.trim() || undefined,
      cell: address,
      text
    });
  } else {
    entries.delete(commentKey(address));
  }
  const nextComments = [...entries.values()];
  removeCommentParts(assets, sheetPath, document, relationshipsDocument);

  if (nextComments.length > 0) {
    const commentsPath = nextPartPath(assets.archive, "xl/comments", "commentAuthoring", "xml");
    const vmlPath = nextPartPath(assets.archive, "xl/drawings", "vmlDrawingCommentAuthoring", "vml");
    assets.archive[commentsPath] = strToU8(buildCommentsXml(nextComments));
    assets.archive[vmlPath] = strToU8(buildCommentsVmlXml(nextComments));
    ensureContentTypeOverride(assets.archive, commentsPath, COMMENTS_CONTENT_TYPE);
    ensureContentTypeDefault(assets.archive, "vml", VML_CONTENT_TYPE);

    createRelationship(relationshipsDocument, COMMENTS_REL_TYPE, relativeArchivePath(sheetPath, commentsPath));
    const vmlRelationshipId = createRelationship(
      relationshipsDocument,
      VML_DRAWING_REL_TYPE,
      relativeArchivePath(sheetPath, vmlPath)
    );
    if (vmlRelationshipId) {
      removeSheetChildren(document, "legacyDrawing");
      const legacyDrawing = createSheetElement(document, "legacyDrawing");
      setRelationshipId(legacyDrawing, vmlRelationshipId);
      insertSheetChild(document, legacyDrawing);
    }
  }

  writeSheetDocument(assets, sheetPath, document);
  writeRelationshipsDocument(assets.archive, relsPath, relationshipsDocument);
  return parseWorksheetState(assets, sheetPath)?.comments ?? [];
}

// ── Conditional formatting ───────────────────────────────────────────────────

const DEFAULT_COLOR_SCALE = ["F8696B", "FFEB84", "63BE7B"];
const DEFAULT_DATA_BAR_COLOR = "5B9BD5";

export function normalizeConditionalFormatColor(value: string) {
  const hex = value.trim().replace(/^#/, "");
  const six = hex.length === 8 ? hex.slice(2) : hex;
  return /^[0-9a-fA-F]{6}$/.test(six) ? `#${six.toLowerCase()}` : "#5b9bd5";
}

function createCfvoDraft(
  threshold: { type: string; value?: number | string } | undefined,
  fallbackType: string
): XlsxConditionalFormatRule["cfvos"][number] {
  const type = threshold?.type?.trim() || fallbackType;
  const rawValue = threshold?.value;
  if (rawValue === undefined || rawValue === null || `${rawValue}`.length === 0) {
    return { type };
  }
  const numeric = typeof rawValue === "number" ? rawValue : Number(rawValue);
  if (type === "formula") {
    return { type, value: String(rawValue) };
  }
  return { type, value: Number.isFinite(numeric) ? numeric : undefined };
}

/**
 * Build a persisted conditional-format rule from a host draft.
 *
 * @param draft Rule description coming from an editor.
 * @param ranges Target ranges; defaults to `draft.ranges`.
 * @param priority Rule priority; defaults to `1`.
 */
export function createConditionalFormatRuleFromDraft(
  draft: XlsxConditionalFormatDraft,
  ranges?: readonly XlsxCellRange[],
  priority = 1
): XlsxConditionalFormatRule | null {
  const targetRanges = (draft.ranges?.length ? draft.ranges : ranges ?? []).map((range) =>
    normalizeRange(range)
  );
  if (targetRanges.length === 0) {
    return null;
  }
  const thresholds = draft.thresholds ?? [];
  if (draft.kind === "colorScale") {
    const colors = draft.colors?.length ? draft.colors : DEFAULT_COLOR_SCALE;
    const stopCount = colors.length >= 3 ? 3 : 2;
    const cfvos = stopCount === 3
      ? [
          createCfvoDraft(thresholds[0], "min"),
          createCfvoDraft(thresholds[1], "percentile"),
          createCfvoDraft(thresholds[2], "max")
        ]
      : [createCfvoDraft(thresholds[0], "min"), createCfvoDraft(thresholds[1], "max")];
    if (cfvos.length < 2) {
      return null;
    }
    return {
      cfvos,
      colors: colors.slice(0, stopCount).map((color) => ({
        rgb: normalizeConditionalFormatColor(color)
      })),
      kind: "colorScale",
      priority,
      ranges: targetRanges
    };
  }
  if (draft.kind === "dataBar") {
    return {
      cfvos: [
        createCfvoDraft(thresholds[0], "min"),
        createCfvoDraft(thresholds[1], "max")
      ],
      color: {
        rgb: normalizeConditionalFormatColor(draft.colors?.[0] ?? DEFAULT_DATA_BAR_COLOR)
      },
      kind: "dataBar",
      priority,
      ranges: targetRanges,
      showValue: draft.showValue ?? true
    };
  }
  const iconSet = draft.iconSet?.trim() || "3TrafficLights1";
  const iconCount = iconSetIconCount(iconSet);
  const cfvos = Array.from({ length: iconCount }, (_, index) => {
    const threshold = thresholds[index];
    if (threshold && (threshold.type || threshold.value !== undefined)) {
      return createCfvoDraft(threshold, "percent");
    }
    return {
      type: "percent",
      value: index === 0 ? 0 : Math.round((index * 100) / iconCount)
    };
  });
  return {
    cfvos,
    icons: cfvos.map((_, index) => ({ iconId: index, iconSet })),
    kind: "iconSet",
    priority,
    ranges: targetRanges,
    reverse: draft.reverse ?? false,
    showValue: draft.showValue ?? true
  };
}

function iconSetIconCount(iconSet: string) {
  const match = /^(\d)/.exec(iconSet);
  const count = match ? Number(match[1]) : 3;
  return count >= 3 && count <= 5 ? count : 3;
}

function readConditionalFormatRules(document: XMLDocument) {
  const scratch: ArchiveEntries = { "xl/worksheets/sheet.xml": strToU8(serializeXml(document)) };
  return parseSheetState(scratch, "xl/worksheets/sheet.xml", {
    includeCachedFormulaValues: false
  })?.conditionalFormatRules ?? [];
}

const MODELED_CONDITIONAL_FORMAT_TYPES = new Set(["colorScale", "dataBar", "iconSet"]);

/**
 * Removes the rules this editor owns while leaving rules it cannot model
 * (`cellIs`, `top10`, `expression`, …) untouched, so a user editing one rule
 * never silently deletes the others.
 */
function removeModeledConditionalFormatNodes(document: XMLDocument) {
  for (const container of getLocalElements(document, "conditionalFormatting")) {
    for (const rule of getChildElements(container, "cfRule")) {
      const type = rule.getAttribute("type") ?? "";
      if (!MODELED_CONDITIONAL_FORMAT_TYPES.has(type)) {
        continue;
      }
      rule.parentNode?.removeChild(rule);
    }
    if (getChildElements(container, "cfRule").length === 0) {
      container.parentNode?.removeChild(container);
    }
  }
}

function writeConditionalFormatNodes(document: XMLDocument, rules: XlsxConditionalFormatRule[]) {
  removeModeledConditionalFormatNodes(document);
  rules.forEach((rule, index) => {
    const container = createSheetElement(document, "conditionalFormatting");
    container.setAttribute("sqref", rule.ranges.map(rangeReference).join(" "));
    const cfRule = createRuleElement(document, rule, index + 1);
    if (!cfRule) {
      return;
    }
    container.appendChild(cfRule);
    insertSheetChild(document, container);
  });
}

function createRuleElement(document: XMLDocument, rule: XlsxConditionalFormatRule, priority: number) {
  const cfRule = createSheetElement(document, "cfRule");
  cfRule.setAttribute("type", rule.kind);
  cfRule.setAttribute("priority", String(priority));
  if (rule.kind === "colorScale") {
    const colorScale = createSheetElement(document, "colorScale");
    for (const cfvo of rule.cfvos) {
      colorScale.appendChild(createCfvoElement(document, cfvo));
    }
    for (const color of rule.colors) {
      const colorNode = createSheetElement(document, "color");
      applySpreadsheetColor(colorNode, color);
      colorScale.appendChild(colorNode);
    }
    cfRule.appendChild(colorScale);
    return cfRule;
  }
  if (rule.kind === "dataBar") {
    const dataBar = createSheetElement(document, "dataBar");
    if (rule.showValue !== undefined) {
      dataBar.setAttribute("showValue", rule.showValue ? "1" : "0");
    }
    for (const cfvo of rule.cfvos) {
      dataBar.appendChild(createCfvoElement(document, cfvo));
    }
    const colorNode = createSheetElement(document, "color");
    applySpreadsheetColor(colorNode, rule.color);
    dataBar.appendChild(colorNode);
    cfRule.appendChild(dataBar);
    return cfRule;
  }
  if (rule.kind === "iconSet") {
    const iconSet = createSheetElement(document, "iconSet");
    iconSet.setAttribute("iconSet", rule.icons[0]?.iconSet ?? "3TrafficLights1");
    if (rule.showValue !== undefined) {
      iconSet.setAttribute("showValue", rule.showValue ? "1" : "0");
    }
    if (rule.reverse) {
      iconSet.setAttribute("reverse", "1");
    }
    for (const cfvo of rule.cfvos) {
      iconSet.appendChild(createCfvoElement(document, cfvo));
    }
    cfRule.appendChild(iconSet);
    return cfRule;
  }
  return null;
}

function createCfvoElement(
  document: XMLDocument,
  cfvo: XlsxConditionalFormatRule["cfvos"][number]
) {
  const node = createSheetElement(document, "cfvo");
  node.setAttribute("type", cfvo.type);
  if (cfvo.value !== undefined && cfvo.value !== null && `${cfvo.value}`.length > 0) {
    if (cfvo.type === "formula") {
      const formula = createSheetElement(document, "f");
      formula.textContent = String(cfvo.value);
      node.appendChild(formula);
    } else {
      node.setAttribute("val", String(cfvo.value));
    }
  }
  return node;
}

/**
 * Normalise a colour value to the eight-digit `AARRGGBB` form ST_HexColorRGB
 * requires.
 *
 * The in-memory model keeps CSS-style `#rrggbb` (that is what the XML reader
 * returns and what the renderer resolves), but OOXML `rgb` attributes are
 * ARGB: writing `#rrggbb` verbatim makes Excel and LibreOffice read the colour
 * as black.
 *
 * @param value `#rrggbb`, `rrggbb`, `#aarrggbb` or `aarrggbb`.
 * @returns Uppercase `AARRGGBB`, or `null` when the value is not a hex colour.
 */
function toOoxmlArgbHex(value: unknown): string | null {
  if (typeof value !== "string") {
    return null;
  }
  const hex = value.trim().replace(/^#/, "");
  if (/^[0-9a-fA-F]{8}$/.test(hex)) {
    return hex.toUpperCase();
  }
  if (/^[0-9a-fA-F]{6}$/.test(hex)) {
    return `FF${hex.toUpperCase()}`;
  }
  return null;
}

function readSpreadsheetNumber(value: unknown): number | null {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }
  if (typeof value === "string" && value.trim().length > 0) {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}

/**
 * Write an OOXML spreadsheet colour element (`<color/>`).
 *
 * Handles every representation the workbook model can carry:
 * `rgb`/`argb`/`hex` (converted to ARGB), `theme` with an optional `tint`,
 * `indexed`, and `auto`. `theme` and `indexed` win over a hex value, matching
 * how Excel resolves a colour element that carries several attributes.
 */
function applySpreadsheetColor(node: Element, color: Record<string, unknown> | undefined) {
  if (!color) {
    return;
  }
  const auto = color.auto === true || color.colorType === "auto";
  const tint = readSpreadsheetNumber(color.tint);
  const theme = readSpreadsheetNumber(color.theme);
  const indexed = readSpreadsheetNumber(color.indexed);
  if (auto) {
    node.setAttribute("auto", "1");
    return;
  }
  if (theme !== null) {
    node.setAttribute("theme", String(theme));
    if (tint !== null) {
      node.setAttribute("tint", String(tint));
    }
    return;
  }
  if (indexed !== null) {
    node.setAttribute("indexed", String(indexed));
    if (tint !== null) {
      node.setAttribute("tint", String(tint));
    }
    return;
  }
  const rgb = toOoxmlArgbHex(color.rgb ?? color.argb ?? color.hex);
  if (!rgb) {
    return;
  }
  node.setAttribute("rgb", rgb);
  if (tint !== null) {
    node.setAttribute("tint", String(tint));
  }
}

/**
 * Applies a conditional-format rule to the worksheet.
 *
 * Priorities are renumbered in list order, which is how the renderer resolves
 * overlapping rules (lowest number wins).
 *
 * @returns The updated rule list, or `null` when the sheet part is unavailable.
 */
export function addWorkbookConditionalFormatRule(
  assets: WorkbookAuthoringAssets,
  workbookSheetIndex: number,
  rule: XlsxConditionalFormatRule
): XlsxConditionalFormatRule[] | null {
  return mutateConditionalFormatRules(assets, workbookSheetIndex, (rules) => [...rules, rule]);
}

export function updateWorkbookConditionalFormatRule(
  assets: WorkbookAuthoringAssets,
  workbookSheetIndex: number,
  index: number,
  rule: XlsxConditionalFormatRule
): XlsxConditionalFormatRule[] | null {
  return mutateConditionalFormatRules(assets, workbookSheetIndex, (rules) => {
    if (index < 0 || index >= rules.length) {
      return null;
    }
    const nextRules = rules.slice();
    nextRules[index] = rule;
    return nextRules;
  });
}

export function deleteWorkbookConditionalFormatRule(
  assets: WorkbookAuthoringAssets,
  workbookSheetIndex: number,
  index: number
): XlsxConditionalFormatRule[] | null {
  return mutateConditionalFormatRules(assets, workbookSheetIndex, (rules) => {
    if (index < 0 || index >= rules.length) {
      return null;
    }
    return rules.filter((_, ruleIndex) => ruleIndex !== index);
  });
}

/**
 * Replaces every conditional-format rule on a worksheet in one write.
 *
 * Structure edits shift each rule's ranges and must rewrite the whole list so
 * priorities stay in list order without one XML pass per rule.
 *
 * @returns The updated rule list, or `null` when the sheet part is unavailable.
 */
export function replaceWorkbookConditionalFormatRules(
  assets: WorkbookAuthoringAssets,
  workbookSheetIndex: number,
  rules: XlsxConditionalFormatRule[]
): XlsxConditionalFormatRule[] | null {
  return mutateConditionalFormatRules(assets, workbookSheetIndex, () => rules);
}

/**
 * Shifts every `conditionalFormatting@sqref` and unmodeled rule formula.
 *
 * Structure edits rewrite the rules this editor models from the in-memory
 * state, but `expression`, `cellIs`, `top10` and similar rules are preserved
 * as raw XML. Without this pass their `sqref` keeps pointing at the old rows
 * and columns after an insert or delete.
 *
 * @param remapRanges Maps the container's parsed ranges; an empty result
 *   removes the whole container.
 * @param remapFormula Optional formula rewrite for unmodeled rule formulas.
 * @returns `true` when a container was rewritten.
 */
export function updateWorkbookConditionalFormatSqrefs(
  assets: WorkbookAuthoringAssets,
  workbookSheetIndex: number,
  remapRanges: (ranges: XlsxCellRange[]) => XlsxCellRange[],
  remapFormula?: (formula: string) => string
): boolean {
  const sheet = readSheetDocument(assets, workbookSheetIndex);
  if (!sheet) {
    return false;
  }
  const { document, sheetPath } = sheet;
  let changed = false;
  for (const container of [...getLocalElements(document, "conditionalFormatting")]) {
    const sqref = container.getAttribute("sqref") ?? "";
    const ranges = sqref.split(/\s+/).flatMap((token) => {
      const range = token ? parseA1RangeReference(token) : null;
      return range ? [range] : [];
    });
    if (ranges.length === 0) {
      continue;
    }
    const nextRanges = remapRanges(ranges);
    if (nextRanges.length === 0) {
      container.parentNode?.removeChild(container);
      changed = true;
      continue;
    }
    const nextSqref = nextRanges.map(rangeReference).join(" ");
    if (nextSqref !== sqref) {
      container.setAttribute("sqref", nextSqref);
      changed = true;
    }
    if (remapFormula) {
      for (const rule of getChildElements(container, "cfRule")) {
        for (const formula of getChildElements(rule, "formula")) {
          const current = formula.textContent ?? "";
          const next = remapFormula(current);
          if (next !== current) {
            formula.textContent = next;
            changed = true;
          }
        }
      }
    }
  }
  if (changed) {
    writeSheetDocument(assets, sheetPath, document);
  }
  return changed;
}

function mutateConditionalFormatRules(
  assets: WorkbookAuthoringAssets,
  workbookSheetIndex: number,
  mutate: (rules: XlsxConditionalFormatRule[]) => XlsxConditionalFormatRule[] | null
) {
  const sheet = readSheetDocument(assets, workbookSheetIndex);
  if (!sheet) {
    return null;
  }
  const { document, sheetPath } = sheet;
  const nextRules = mutate(readConditionalFormatRules(document));
  if (!nextRules) {
    return null;
  }
  writeConditionalFormatNodes(document, nextRules);
  writeSheetDocument(assets, sheetPath, document);
  return parseWorksheetState(assets, sheetPath)?.conditionalFormatRules ?? [];
}

// ── Data validation ──────────────────────────────────────────────────────────

function expandRangeCells(ranges: readonly XlsxCellRange[]) {
  const cells = new Set<string>();
  for (const range of ranges) {
    const normalized = normalizeRange(range);
    for (let row = normalized.start.row; row <= normalized.end.row; row += 1) {
      for (let col = normalized.start.col; col <= normalized.end.col; col += 1) {
        cells.add(`${row}:${col}`);
      }
    }
  }
  return cells;
}

/**
 * Applies a data-validation rule to one or more ranges, replacing any rule that
 * targets the same cells. Passing `null` removes validations for the ranges.
 *
 * @returns The updated validation list, or `null` when the sheet part is unavailable.
 */
export function setWorkbookDataValidation(
  assets: WorkbookAuthoringAssets,
  workbookSheetIndex: number,
  ranges: readonly XlsxCellRange[],
  validation: XlsxDataValidation | null
): XlsxDataValidation[] | null {
  const sheet = readSheetDocument(assets, workbookSheetIndex);
  if (!sheet) {
    return null;
  }
  const { document, sheetPath } = sheet;
  const targetCells = expandRangeCells(ranges);
  if (targetCells.size === 0) {
    return null;
  }
  const current = parseWorksheetState(assets, sheetPath)?.dataValidations ?? [];
  const kept = current.filter((entry) => {
    const ruleCells = expandRangeCells(entry.ranges);
    for (const cell of ruleCells) {
      if (targetCells.has(cell)) {
        return false;
      }
    }
    return true;
  });
  const next = validation
    ? [...kept, { ...validation, ranges: ranges.map((range) => normalizeRange(range)) }]
    : kept;
  writeDataValidationNodes(document, next);
  writeSheetDocument(assets, sheetPath, document);
  return parseWorksheetState(assets, sheetPath)?.dataValidations ?? [];
}

function writeDataValidationNodes(document: XMLDocument, validations: XlsxDataValidation[]) {
  removeSheetChildren(document, "dataValidations");
  if (validations.length === 0) {
    return;
  }
  const container = createSheetElement(document, "dataValidations");
  container.setAttribute("count", String(validations.length));
  for (const validation of validations) {
    const node = createSheetElement(document, "dataValidation");
    if (validation.validationType) {
      node.setAttribute("type", validation.validationType);
    }
    if (validation.operator) {
      node.setAttribute("operator", validation.operator);
    }
    if (validation.allowBlank !== undefined) {
      node.setAttribute("allowBlank", validation.allowBlank ? "1" : "0");
    }
    if (validation.showDropdown !== undefined) {
      // OOXML inverts the flag: showDropDown="1" hides the in-cell dropdown.
      node.setAttribute("showDropDown", validation.showDropdown ? "0" : "1");
    }
    if (validation.showErrorAlert !== undefined) {
      node.setAttribute("showErrorMessage", validation.showErrorAlert ? "1" : "0");
    }
    if (validation.showInputMessage !== undefined) {
      node.setAttribute("showInputMessage", validation.showInputMessage ? "1" : "0");
    }
    if (validation.errorStyle) {
      node.setAttribute("errorStyle", validation.errorStyle);
    }
    if (validation.errorTitle) {
      node.setAttribute("errorTitle", validation.errorTitle);
    }
    if (validation.errorMessage) {
      node.setAttribute("error", validation.errorMessage);
    }
    if (validation.inputTitle) {
      node.setAttribute("promptTitle", validation.inputTitle);
    }
    if (validation.inputMessage) {
      node.setAttribute("prompt", validation.inputMessage);
    }
    node.setAttribute("sqref", validation.ranges.map(rangeReference).join(" "));
    const formula1 = validation.value1 ?? validation.listSource;
    if (formula1 !== undefined && formula1 !== null && `${formula1}`.length > 0) {
      const formulaNode = createSheetElement(document, "formula1");
      formulaNode.textContent = String(formula1);
      node.appendChild(formulaNode);
    }
    if (validation.value2 !== undefined && validation.value2 !== null && `${validation.value2}`.length > 0) {
      const formulaNode = createSheetElement(document, "formula2");
      formulaNode.textContent = String(validation.value2);
      node.appendChild(formulaNode);
    }
    container.appendChild(node);
  }
  insertSheetChild(document, container);
}

/**
 * Replaces every data-validation rule on a worksheet in one write.
 *
 * Structure edits shift each rule's ranges; rewriting the whole
 * `dataValidations` block keeps the OOXML `count` and rule order consistent
 * without clearing rules cell by cell.
 *
 * @returns The updated validation list, or `null` when the sheet part is unavailable.
 */
export function replaceWorkbookDataValidations(
  assets: WorkbookAuthoringAssets,
  workbookSheetIndex: number,
  validations: XlsxDataValidation[]
): XlsxDataValidation[] | null {
  const sheet = readSheetDocument(assets, workbookSheetIndex);
  if (!sheet) {
    return null;
  }
  writeDataValidationNodes(sheet.document, validations);
  writeSheetDocument(assets, sheet.sheetPath, sheet.document);
  return parseWorksheetState(assets, sheet.sheetPath)?.dataValidations ?? [];
}

// ── Tables ───────────────────────────────────────────────────────────────────

function buildTableXml(options: {
  columnNames: string[];
  hasHeaders: boolean;
  name: string;
  range: XlsxCellRange;
  styleInfo: XlsxTable["styleInfo"];
  tableId: number;
}) {
  const columns = options.columnNames
    .map((name, index) => `<tableColumn id="${index + 1}" name="${escapeXmlText(name)}"/>`)
    .join("");
  const reference = rangeReference(options.range);
  const filterReference = options.hasHeaders
    ? rangeReference({
        end: options.range.end,
        start: { col: options.range.start.col, row: options.range.start.row + 1 }
      })
    : reference;
  const style = options.styleInfo
    ? `<tableStyleInfo name="${escapeXmlText(options.styleInfo.name ?? "TableStyleMedium2")}" showFirstColumn="${options.styleInfo.showFirstColumn ? 1 : 0}" showLastColumn="${options.styleInfo.showLastColumn ? 1 : 0}" showRowStripes="${options.styleInfo.showRowStripes ? 1 : 0}" showColumnStripes="${options.styleInfo.showColumnStripes ? 1 : 0}"/>`
    : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="${SPREADSHEET_NS}" id="${options.tableId}" name="${escapeXmlText(options.name)}" displayName="${escapeXmlText(options.name)}" ref="${reference}" headerRowCount="${options.hasHeaders ? 1 : 0}" totalsRowShown="0"><autoFilter ref="${filterReference}"/><tableColumns count="${options.columnNames.length}">${columns}</tableColumns>${style}</table>`;
}

/**
 * Excel requires table names and ids to be unique across the workbook, so the
 * authoring helpers look at every worksheet's table parts, not just the active
 * sheet.
 */
function readWorkbookTableUsage(assets: WorkbookAuthoringAssets) {
  const names = new Set<string>();
  const ids: number[] = [];
  for (const sheetPath of assets.sheetPathsByWorkbookSheetIndex) {
    if (!sheetPath) {
      continue;
    }
    for (const table of parseWorksheetTableMetadata(assets.archive, sheetPath)) {
      if (table.name) {
        names.add(table.name.toLowerCase());
      }
      if (table.displayName) {
        names.add(table.displayName.toLowerCase());
      }
    }
    const relationships = readRelationships(assets.archive, relsPathForDocument(sheetPath), sheetPath);
    for (const relationship of relationships.values()) {
      if (relationship.type !== TABLE_REL_TYPE) {
        continue;
      }
      const tableXml = readArchiveText(assets.archive, relationship.target);
      const tableDocument = tableXml ? parseXml(tableXml) : null;
      const tableNode = tableDocument ? getLocalElements(tableDocument, "table")[0] : null;
      const id = Number(tableNode?.getAttribute("id") ?? Number.NaN);
      if (Number.isFinite(id)) {
        ids.push(id);
      }
    }
  }
  return { ids, names };
}

function readTableMetadata(assets: WorkbookAuthoringAssets, workbookSheetIndex: number) {
  const sheetPath = resolveSheetPath(assets, workbookSheetIndex);
  if (!sheetPath) {
    return [];
  }
  const metadata = parseWorksheetTableMetadata(assets.archive, sheetPath);
  if (assets.tableMetadataByWorkbookSheetIndex) {
    assets.tableMetadataByWorkbookSheetIndex[workbookSheetIndex] = metadata;
  }
  return normalizeTableMetadata(metadata);
}

function normalizeTableMetadata(metadata: WorkbookTableMetadata[]) {
  return metadata.flatMap((table: WorkbookTableMetadata) => {
    const reference = table.reference ?? "";
    const range = table.start && table.end
      ? { start: table.start, end: table.end }
      : parseA1RangeReference(reference);
    const name = table.name ?? "";
    if (!range || !name) {
      return [];
    }
    return [{
      columns: table.columns ?? [],
      displayName: table.displayName ?? name,
      end: range.end,
      headerRowCount: table.headerRowCount ?? 1,
      headerRowCellStyle: table.headerRowCellStyle,
      name,
      reference: reference || rangeReference(range),
      start: range.start,
      styleInfo: table.styleInfo,
      totalsRowCount: table.totalsRowCount ?? 0,
      totalsRowShown: table.totalsRowShown ?? false
    } satisfies XlsxTable];
  });
}

/**
 * Creates a ListObject table from a cell range.
 *
 * Column names default to the header row cells when `hasHeaders` is true and to
 * `Column N` otherwise. The table part is written into the source archive so it
 * is included in every later save and export.
 *
 * @returns The created table plus the refreshed table list, or `null` when the
 * range is invalid, the name is taken, or the sheet part is unavailable.
 */
export function addWorkbookTable(
  assets: WorkbookAuthoringAssets,
  workbookSheetIndex: number,
  input: XlsxTableInsertInput
): { table: XlsxTable; tables: XlsxTable[] } | null {
  const sheet = readSheetDocument(assets, workbookSheetIndex);
  if (!sheet) {
    return null;
  }
  const { document, sheetPath } = sheet;
  const range = normalizeRange(input.range);
  const columnCount = range.end.col - range.start.col + 1;
  const rowCount = range.end.row - range.start.row + 1;
  const hasHeaders = input.hasHeaders !== false;
  if (columnCount < 1 || rowCount < (hasHeaders ? 2 : 1)) {
    return null;
  }

  const workbookTables = readWorkbookTableUsage(assets);
  const usedNames = workbookTables.names;
  const requestedName = (input.name ?? input.displayName ?? "").trim()
    || `Table${usedNames.size + 1}`;
  if (!/^[A-Za-z_\\][A-Za-z0-9_.]*$/.test(requestedName) || usedNames.has(requestedName.toLowerCase())) {
    return null;
  }
  const columnNames = Array.from({ length: columnCount }, (_, index) =>
    (input.columnNames[index] ?? "").trim() || `Column${index + 1}`
  );
  if (new Set(columnNames.map((name) => name.toLowerCase())).size !== columnNames.length) {
    return null;
  }

  const relsPath = relsPathForDocument(sheetPath);
  const relationshipsDocument = ensureRelationshipsDocument(assets.archive, relsPath);
  if (!relationshipsDocument?.documentElement) {
    return null;
  }
  const tablePath = nextPartPath(assets.archive, "xl/tables", "tableAuthoring", "xml");
  assets.archive[tablePath] = strToU8(buildTableXml({
    columnNames,
    hasHeaders,
    name: requestedName,
    range,
    styleInfo: {
      name: input.styleName ?? "TableStyleMedium2",
      showColumnStripes: input.showColumnStripes ?? false,
      showFirstColumn: input.showFirstColumn ?? false,
      showLastColumn: input.showLastColumn ?? false,
      showRowStripes: input.showRowStripes ?? true
    },
    tableId: Math.max(0, ...workbookTables.ids) + 1
  }));
  ensureContentTypeOverride(assets.archive, tablePath, TABLE_CONTENT_TYPE);

  const relationshipId = createRelationship(
    relationshipsDocument,
    TABLE_REL_TYPE,
    relativeArchivePath(sheetPath, tablePath)
  );
  if (!relationshipId) {
    return null;
  }
  let tableParts = getLocalElements(document, "tableParts")[0];
  if (!tableParts) {
    tableParts = createSheetElement(document, "tableParts");
    insertSheetChild(document, tableParts);
  }
  const tablePart = createSheetElement(document, "tablePart");
  setRelationshipId(tablePart, relationshipId);
  tableParts.appendChild(tablePart);
  tableParts.setAttribute("count", String(getChildElements(tableParts, "tablePart").length));

  writeSheetDocument(assets, sheetPath, document);
  writeRelationshipsDocument(assets.archive, relsPath, relationshipsDocument);

  const tables = readTableMetadata(assets, workbookSheetIndex);
  const created = tables.find((table) => table.name === requestedName) ?? null;
  return created ? { table: created, tables } : null;
}

/**
 * Removes a table definition from the worksheet.
 *
 * @returns The remaining table list, or `null` when the sheet or table is unavailable.
 */
export function deleteWorkbookTable(
  assets: WorkbookAuthoringAssets,
  workbookSheetIndex: number,
  tableName: string
): XlsxTable[] | null {
  const sheet = readSheetDocument(assets, workbookSheetIndex);
  if (!sheet) {
    return null;
  }
  const { document, sheetPath } = sheet;
  const relsPath = relsPathForDocument(sheetPath);
  const relationshipsDocument = ensureRelationshipsDocument(assets.archive, relsPath);
  if (!relationshipsDocument?.documentElement) {
    return null;
  }
  const relationships = readRelationships(assets.archive, relsPath, sheetPath);
  let removed = false;
  for (const tablePart of getLocalElements(document, "tablePart")) {
    const relationshipId = getRelationshipId(tablePart);
    const relationship = relationshipId ? relationships.get(relationshipId) : undefined;
    if (!relationship || relationship.type !== TABLE_REL_TYPE) {
      continue;
    }
    const tableXml = readArchiveText(assets.archive, relationship.target);
    const tableDocument = tableXml ? parseXml(tableXml) : null;
    const tableNode = tableDocument ? getLocalElements(tableDocument, "table")[0] : null;
    const name = tableNode?.getAttribute("name") ?? tableNode?.getAttribute("displayName");
    if (name !== tableName) {
      continue;
    }
    delete assets.archive[relationship.target];
    if (relationshipId) {
      removeRelationshipById(relationshipsDocument, relationshipId);
    }
    tablePart.parentNode?.removeChild(tablePart);
    removed = true;
  }
  if (!removed) {
    return null;
  }
  const tableParts = getLocalElements(document, "tableParts")[0];
  if (tableParts) {
    const remaining = getChildElements(tableParts, "tablePart").length;
    if (remaining === 0) {
      tableParts.parentNode?.removeChild(tableParts);
    } else {
      tableParts.setAttribute("count", String(remaining));
    }
  }
  writeSheetDocument(assets, sheetPath, document);
  writeRelationshipsDocument(assets.archive, relsPath, relationshipsDocument);
  return readTableMetadata(assets, workbookSheetIndex);
}

/**
 * Moves a table to a new range after a structural edit.
 *
 * Rewrites the table part's `ref` (and its auto-filter range) without touching
 * the column names, style or id, so Excel keeps the ListObject identity.
 *
 * @returns The updated table list, or `null` when the table was not found.
 */
export function updateWorkbookTableRange(
  assets: WorkbookAuthoringAssets,
  workbookSheetIndex: number,
  tableName: string,
  range: XlsxCellRange
): XlsxTable[] | null {
  const sheet = readSheetDocument(assets, workbookSheetIndex);
  if (!sheet) {
    return null;
  }
  const { sheetPath } = sheet;
  const relationships = readRelationships(assets.archive, relsPathForDocument(sheetPath), sheetPath);
  let updated = false;
  for (const relationship of relationships.values()) {
    if (relationship.type !== TABLE_REL_TYPE) {
      continue;
    }
    const tableXml = readArchiveText(assets.archive, relationship.target);
    const tableDocument = tableXml ? parseXml(tableXml) : null;
    if (!tableDocument) {
      continue;
    }
    const tableNode = getLocalElements(tableDocument, "table")[0];
    if (!tableNode) {
      continue;
    }
    const name = tableNode.getAttribute("name") ?? tableNode.getAttribute("displayName");
    if (name !== tableName) {
      continue;
    }
    const normalized = normalizeRange(range);
    const reference = rangeReference(normalized);
    tableNode.setAttribute("ref", reference);
    const hasHeaders = Number(tableNode.getAttribute("headerRowCount") ?? "1") > 0;
    const autoFilter = getLocalElements(tableDocument, "autoFilter")[0];
    if (autoFilter) {
      autoFilter.setAttribute(
        "ref",
        hasHeaders && normalized.end.row > normalized.start.row
          ? rangeReference({
              end: normalized.end,
              start: { col: normalized.start.col, row: normalized.start.row + 1 }
            })
          : reference
      );
    }
    const sortState = getLocalElements(tableDocument, "sortState")[0];
    if (sortState) {
      sortState.setAttribute("ref", reference);
    }
    assets.archive[normalizeArchivePath(relationship.target)] = strToU8(serializeXml(tableDocument));
    updated = true;
    break;
  }
  if (!updated) {
    return null;
  }
  return readTableMetadata(assets, workbookSheetIndex);
}

// ── Read-back ────────────────────────────────────────────────────────────────

/**
 * Reads the authoring-visible objects a worksheet currently carries.
 *
 * Useful for hydrating controller state after a mutation or a history restore
 * without rebuilding every sheet.
 */
export function readWorkbookSheetAuthoringState(
  assets: WorkbookAuthoringAssets,
  workbookSheetIndex: number
): WorkbookSheetAuthoringState | null {
  const sheet = readSheetDocument(assets, workbookSheetIndex);
  if (!sheet) {
    return null;
  }
  const state = parseWorksheetState(assets, sheet.sheetPath);
  if (!state) {
    return null;
  }
  return {
    comments: state.comments,
    conditionalFormatRules: state.conditionalFormatRules,
    dataValidations: state.dataValidations ?? [],
    freezePanes: parseWorksheetFreezePanesDocument(sheet.document),
    hyperlinks: state.hyperlinks,
    tables: readTableMetadata(assets, workbookSheetIndex)
  };
}

// ── Chart attachment helpers (shared with charts/chart-insert) ──────────────

/** Appends a drawing attachment so a newly inserted drawing is exported. */
export function appendWorkbookSheetDrawingAttachment(
  assets: WorkbookAuthoringAssets,
  workbookSheetIndex: number,
  attachment: XlsxImageAttachment
) {
  const origins = assets.sheetOrigins;
  if (!origins) {
    return false;
  }
  const current = origins[workbookSheetIndex];
  if (!current) {
    origins[workbookSheetIndex] = {
      attachments: [attachment],
      workbookSheetIndex
    };
    return true;
  }
  if (current.attachments.some((entry) => entry.drawingPath === attachment.drawingPath)) {
    return false;
  }
  current.attachments = [...current.attachments, attachment];
  return true;
}

export {
  COMMENTS_CONTENT_TYPE,
  COMMENTS_REL_TYPE,
  TABLE_CONTENT_TYPE,
  TABLE_REL_TYPE,
  VML_CONTENT_TYPE,
  ensureContentTypeDefault,
  ensureContentTypeOverride,
  insertSheetChild
};
