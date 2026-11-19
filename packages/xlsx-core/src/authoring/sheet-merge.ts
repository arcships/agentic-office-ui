import { strToU8, unzipSync, zipSync } from "fflate";
import {
  CONTENT_TYPES_NS,
  HYPERLINK_REL_TYPE,
  PKG_REL_NS,
  REL_NS,
  VML_DRAWING_REL_TYPE,
  type ArchiveEntries,
  cloneBytes,
  getChildElements,
  getFirstChild,
  getFirstDescendant,
  getLocalElements,
  getRelationshipId,
  normalizeArchivePath,
  parseWorkbookSheets,
  parseXml,
  readArchiveText,
  relativeArchivePath,
  relsPathForDocument,
  serializeXml,
} from "../images";
import {
  COMMENTS_CONTENT_TYPE,
  COMMENTS_REL_TYPE,
  TABLE_CONTENT_TYPE,
  TABLE_REL_TYPE,
  VML_CONTENT_TYPE,
  ensureContentTypeDefault,
  ensureContentTypeOverride,
  insertSheetChild,
} from "./sheet-authoring";

const AUTHORING_FRAGMENT_PATTERN =
  /<(?:[A-Za-z0-9]+:)?(pane|hyperlinks|conditionalFormatting|dataValidations|tableParts|pivotTableParts|slicerList|legacyDrawing)[\s>/]/;

const PIVOT_TABLE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
const PIVOT_CACHE_DEFINITION_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
const SLICER_REL_TYPE = "http://schemas.microsoft.com/office/2007/relationships/slicer";
const SLICER_CACHE_REL_TYPE = "http://schemas.microsoft.com/office/2007/relationships/slicerCache";
const PIVOT_TABLE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
const PIVOT_CACHE_DEFINITION_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
const PIVOT_CACHE_RECORDS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";
const SLICER_CONTENT_TYPE = "application/vnd.ms-excel.slicer+xml";
const SLICER_CACHE_CONTENT_TYPE = "application/vnd.ms-excel.slicerCache+xml";

const PAGE_LAYOUT_ELEMENTS = [
  "printOptions",
  "pageMargins",
  "pageSetup",
  "rowBreaks",
  "colBreaks",
] as const;

type RelationshipSource = {
  target: string;
  targetMode: string | null;
  type: string;
};

function readRawRelationships(archive: ArchiveEntries, relsPath: string) {
  const xml = readArchiveText(archive, relsPath);
  const document = xml ? parseXml(xml) : null;
  const relationships = new Map<string, RelationshipSource>();
  if (!document) {
    return relationships;
  }
  for (const node of getLocalElements(document, "Relationship")) {
    const id = node.getAttribute("Id");
    const target = node.getAttribute("Target");
    const type = node.getAttribute("Type");
    if (!id || !target || !type) {
      continue;
    }
    relationships.set(id, {
      target,
      targetMode: node.getAttribute("TargetMode"),
      type
    });
  }
  return relationships;
}

function ensureRelationshipsXml(archive: ArchiveEntries, relsPath: string) {
  const existing = readArchiveText(archive, relsPath);
  if (existing) {
    const document = parseXml(existing);
    if (document?.documentElement) {
      return document;
    }
  }
  const document = parseXml(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}"></Relationships>`
  );
  return document;
}

function nextRelationshipId(document: XMLDocument) {
  const existingIds = new Set(
    getLocalElements(document, "Relationship")
      .map((node) => node.getAttribute("Id"))
      .filter((value): value is string => Boolean(value))
  );
  let index = 1;
  while (existingIds.has(`rIdXlsxAuthored${index}`)) {
    index += 1;
  }
  return `rIdXlsxAuthored${index}`;
}

function addRelationship(
  document: XMLDocument,
  type: string,
  target: string,
  targetMode: string | null
) {
  const root = document.documentElement;
  if (!root) {
    return null;
  }
  const id = nextRelationshipId(document);
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

function removeRelationshipsByType(document: XMLDocument, types: ReadonlySet<string>) {
  for (const node of getLocalElements(document, "Relationship")) {
    if (types.has(node.getAttribute("Type") ?? "")) {
      node.parentNode?.removeChild(node);
    }
  }
}

function removeSheetElement(document: XMLDocument, localName: string) {
  for (const node of Array.from(document.getElementsByTagName("*"))) {
    if (node.localName === localName && node.parentNode) {
      node.parentNode.removeChild(node);
    }
  }
}

function ensureSheetView(document: XMLDocument) {
  const existing = getLocalElements(document, "sheetViews")[0];
  if (existing) {
    const sheetView = getFirstChild(existing, "sheetView");
    if (sheetView) {
      return sheetView;
    }
    const created = document.createElementNS(existing.namespaceURI ?? "", "sheetView");
    created.setAttribute("workbookViewId", "0");
    existing.appendChild(created);
    return created;
  }
  const namespace = document.documentElement.namespaceURI ?? "";
  const sheetViews = document.createElementNS(namespace, "sheetViews");
  const sheetView = document.createElementNS(namespace, "sheetView");
  sheetView.setAttribute("workbookViewId", "0");
  sheetViews.appendChild(sheetView);
  insertSheetChild(document, sheetViews);
  return sheetView;
}

/** Resolve a package-relative relationship target against its owning part. */
function resolvePartPath(ownerPath: string, target: string) {
  if (target.startsWith("/")) {
    return normalizeArchivePath(target);
  }
  const segments = normalizeArchivePath(ownerPath).split("/").slice(0, -1);
  for (const segment of target.replace(/\\/g, "/").split("/")) {
    if (!segment || segment === ".") {
      continue;
    }
    if (segment === "..") {
      segments.pop();
      continue;
    }
    segments.push(segment);
  }
  return normalizeArchivePath(segments.join("/"));
}

function nextArchivePartPath(
  archive: ArchiveEntries,
  directory: string,
  baseName: string,
  extension: string
) {
  let index = 1;
  while (archive[`${directory}/${baseName}${index}.${extension}`]) {
    index += 1;
  }
  return `${directory}/${baseName}${index}.${extension}`;
}

function ensureRelationshipNamespace(document: XMLDocument) {
  const root = document.documentElement;
  if (!root || root.getAttribute("xmlns:r")) {
    return;
  }
  root.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:r", REL_NS);
}

function setRelationshipId(element: Element, relationshipId: string) {
  ensureRelationshipNamespace(element.ownerDocument);
  element.setAttributeNS(REL_NS, "r:id", relationshipId);
}

function clearRelationshipId(element: Element) {
  element.removeAttributeNS(REL_NS, "id");
  element.removeAttribute("r:id");
}

function applyFreezePane(sourceDocument: XMLDocument, savedDocument: XMLDocument) {
  const sourcePane = getLocalElements(sourceDocument, "pane")[0] ?? null;
  for (const sheetView of getLocalElements(savedDocument, "sheetView")) {
    for (const pane of getChildElements(sheetView, "pane")) {
      sheetView.removeChild(pane);
    }
  }
  if (!sourcePane) {
    return;
  }
  const sheetView = ensureSheetView(savedDocument);
  const imported = savedDocument.importNode(sourcePane, true) as Element;
  sheetView.insertBefore(imported, sheetView.firstChild);
  const activePane = imported.getAttribute("activePane");
  for (const selection of getChildElements(sheetView, "selection")) {
    if (activePane) {
      selection.setAttribute("pane", activePane);
    } else {
      selection.removeAttribute("pane");
    }
  }
}

function applyHyperlinks(
  sourceDocument: XMLDocument,
  sourceRelationships: Map<string, RelationshipSource>,
  savedDocument: XMLDocument,
  savedRelationships: XMLDocument
) {
  const sourceHyperlinks = getLocalElements(sourceDocument, "hyperlinks")[0] ?? null;
  removeSheetElement(savedDocument, "hyperlinks");
  removeRelationshipsByType(savedRelationships, new Set([HYPERLINK_REL_TYPE]));
  if (!sourceHyperlinks) {
    return;
  }
  const imported = savedDocument.importNode(sourceHyperlinks, true) as Element;
  for (const hyperlink of getChildElements(imported, "hyperlink")) {
    const relationshipId = getRelationshipId(hyperlink);
    if (!relationshipId) {
      continue;
    }
    const relationship = sourceRelationships.get(relationshipId);
    if (!relationship) {
      clearRelationshipId(hyperlink);
      continue;
    }
    const nextId = addRelationship(
      savedRelationships,
      relationship.type,
      relationship.target,
      relationship.targetMode
    );
    if (!nextId) {
      clearRelationshipId(hyperlink);
      continue;
    }
    setRelationshipId(hyperlink, nextId);
  }
  const hyperlinkCount = getChildElements(imported, "hyperlink").length;
  if (hyperlinkCount === 0) {
    return;
  }
  imported.setAttribute("count", String(hyperlinkCount));
  insertSheetChild(savedDocument, imported);
}

function applyConditionalFormats(sourceDocument: XMLDocument, savedDocument: XMLDocument) {
  const sourceNodes = getLocalElements(sourceDocument, "conditionalFormatting");
  removeSheetElement(savedDocument, "conditionalFormatting");
  for (const node of sourceNodes) {
    if (getChildElements(node, "cfRule").length === 0) {
      continue;
    }
    insertSheetChild(savedDocument, savedDocument.importNode(node, true) as Element);
  }
}

function applyDataValidations(sourceDocument: XMLDocument, savedDocument: XMLDocument) {
  const sourceNode = getLocalElements(sourceDocument, "dataValidations")[0] ?? null;
  removeSheetElement(savedDocument, "dataValidations");
  if (!sourceNode) {
    return;
  }
  const validationCount = getChildElements(sourceNode, "dataValidation").length;
  if (validationCount === 0) {
    return;
  }
  const imported = savedDocument.importNode(sourceNode, true) as Element;
  imported.setAttribute("count", String(validationCount));
  insertSheetChild(savedDocument, imported);
}

function copySourcePart(
  savedArchive: ArchiveEntries,
  sourceArchive: ArchiveEntries,
  sourcePartPath: string,
  directory: string,
  baseName: string,
  extension: string
) {
  const bytes = sourceArchive[normalizeArchivePath(sourcePartPath)];
  if (!bytes) {
    return null;
  }
  const targetPath = nextArchivePartPath(savedArchive, directory, baseName, extension);
  savedArchive[targetPath] = cloneBytes(bytes);
  return targetPath;
}

function applyTables(options: {
  sourceArchive: ArchiveEntries;
  sourceDocument: XMLDocument;
  sourceRelationships: Map<string, RelationshipSource>;
  sourceSheetPath: string;
  savedArchive: ArchiveEntries;
  savedDocument: XMLDocument;
  savedRelationships: XMLDocument;
  savedSheetPath: string;
}) {
  const sourceTableParts = getLocalElements(options.sourceDocument, "tableParts")[0] ?? null;
  removeSheetElement(options.savedDocument, "tableParts");
  removeRelationshipsByType(options.savedRelationships, new Set([TABLE_REL_TYPE]));
  if (!sourceTableParts) {
    return;
  }
  const importedParts = options.savedDocument.importNode(sourceTableParts, true) as Element;
  let copied = 0;
  for (const tablePart of getChildElements(importedParts, "tablePart")) {
    const relationshipId = getRelationshipId(tablePart);
    const relationship = relationshipId
      ? options.sourceRelationships.get(relationshipId)
      : undefined;
    if (!relationship || relationship.type !== TABLE_REL_TYPE) {
      tablePart.parentNode?.removeChild(tablePart);
      continue;
    }
    const sourceTablePath = resolvePartPath(options.sourceSheetPath, relationship.target);
    const copiedPath = copySourcePart(
      options.savedArchive,
      options.sourceArchive,
      sourceTablePath,
      "xl/tables",
      "tableAuthored",
      "xml"
    );
    if (!copiedPath) {
      tablePart.parentNode?.removeChild(tablePart);
      continue;
    }
    ensureContentTypeOverride(options.savedArchive, copiedPath, TABLE_CONTENT_TYPE);
    const nextId = addRelationship(
      options.savedRelationships,
      TABLE_REL_TYPE,
      relativeArchivePath(options.savedSheetPath, copiedPath),
      null
    );
    if (!nextId) {
      tablePart.parentNode?.removeChild(tablePart);
      continue;
    }
    setRelationshipId(tablePart, nextId);
    copied += 1;
  }
  if (copied === 0) {
    return;
  }
  importedParts.setAttribute("count", String(copied));
  insertSheetChild(options.savedDocument, importedParts);
}

function copyPartPreservingPath(
  savedArchive: ArchiveEntries,
  sourceArchive: ArchiveEntries,
  partPath: string,
  contentType: string,
) {
  const normalized = normalizeArchivePath(partPath);
  const bytes = sourceArchive[normalized];
  if (!bytes) {
    return null;
  }
  // The source part keeps its own relationships (a pivot cache definition
  // points at its records part), so preserving the package path is simpler and
  // safer than renumbering a dependency tree.
  savedArchive[normalized] = cloneBytes(bytes);
  ensureContentTypeOverride(savedArchive, normalized, contentType);
  return normalized;
}

/**
 * Copies the pivot table parts a worksheet references into the saved package.
 *
 * The WASM engine rewrites cell data but drops `<pivotTableParts>` and the
 * related parts, so an exported workbook would lose its pivots without this
 * merge.
 */
function applyPivotTableParts(options: {
  sourceArchive: ArchiveEntries;
  sourceDocument: XMLDocument;
  sourceRelationships: Map<string, RelationshipSource>;
  sourceSheetPath: string;
  savedArchive: ArchiveEntries;
  savedDocument: XMLDocument;
  savedRelationships: XMLDocument;
  savedSheetPath: string;
}) {
  const sourceParts = getLocalElements(options.sourceDocument, "pivotTableParts")[0] ?? null;
  removeSheetElement(options.savedDocument, "pivotTableParts");
  removeRelationshipsByType(options.savedRelationships, new Set([PIVOT_TABLE_REL_TYPE]));
  if (!sourceParts) {
    return;
  }

  const importedParts = options.savedDocument.importNode(sourceParts, true) as Element;
  let copied = 0;
  for (const pivotPart of getChildElements(importedParts, "pivotTablePart")) {
    const relationshipId = getRelationshipId(pivotPart);
    const relationship = relationshipId
      ? options.sourceRelationships.get(relationshipId)
      : undefined;
    if (!relationship || relationship.type !== PIVOT_TABLE_REL_TYPE) {
      pivotPart.parentNode?.removeChild(pivotPart);
      continue;
    }
    const sourcePath = resolvePartPath(options.sourceSheetPath, relationship.target);
    const copiedPath = copyPartPreservingPath(
      options.savedArchive,
      options.sourceArchive,
      sourcePath,
      PIVOT_TABLE_CONTENT_TYPE,
    );
    if (!copiedPath) {
      pivotPart.parentNode?.removeChild(pivotPart);
      continue;
    }
    const nextId = addRelationship(
      options.savedRelationships,
      PIVOT_TABLE_REL_TYPE,
      relativeArchivePath(options.savedSheetPath, copiedPath),
      null,
    );
    if (!nextId) {
      pivotPart.parentNode?.removeChild(pivotPart);
      continue;
    }
    setRelationshipId(pivotPart, nextId);
    copied += 1;
  }
  if (copied === 0) {
    return;
  }
  importedParts.setAttribute("count", String(copied));
  insertSheetChild(options.savedDocument, importedParts);
}

/** The `<extLst>` of a worksheet or workbook, created in schema order. */
function ensureExtLst(document: XMLDocument, insert: (node: Element) => void) {
  const existing = getLocalElements(document, "extLst")[0];
  if (existing) {
    return existing;
  }
  const extLst = document.createElementNS(document.documentElement.namespaceURI ?? "", "extLst");
  insert(extLst);
  return extLst;
}

function removeExtByUri(extLst: Element, uri: string) {
  for (const extNode of getChildElements(extLst, "ext")) {
    if (extNode.getAttribute("uri") === uri) {
      extLst.removeChild(extNode);
    }
  }
}

/** Copies the slicer parts referenced from a worksheet's `<x14:slicerList>`. */
function applySlicerParts(options: {
  sourceArchive: ArchiveEntries;
  sourceDocument: XMLDocument;
  sourceRelationships: Map<string, RelationshipSource>;
  sourceSheetPath: string;
  savedArchive: ArchiveEntries;
  savedDocument: XMLDocument;
  savedRelationships: XMLDocument;
  savedSheetPath: string;
}) {
  const sourceSlicerList = getLocalElements(options.sourceDocument, "slicerList")[0] ?? null;
  if (!sourceSlicerList) {
    return;
  }
  const sourceExt = sourceSlicerList.parentNode;
  if (!sourceExt || !("localName" in sourceExt) || (sourceExt as Element).localName !== "ext") {
    return;
  }
  const importedExt = options.savedDocument.importNode(sourceExt as Element, true) as Element;
  const uri = importedExt.getAttribute("uri") ?? "";
  let copied = 0;
  for (const slicerNode of getLocalElements(importedExt, "slicer")) {
    const relationshipId = getRelationshipId(slicerNode);
    const relationship = relationshipId
      ? options.sourceRelationships.get(relationshipId)
      : undefined;
    if (!relationship || relationship.type !== SLICER_REL_TYPE) {
      slicerNode.parentNode?.removeChild(slicerNode);
      continue;
    }
    const sourcePath = resolvePartPath(options.sourceSheetPath, relationship.target);
    const copiedPath = copyPartPreservingPath(
      options.savedArchive,
      options.sourceArchive,
      sourcePath,
      SLICER_CONTENT_TYPE,
    );
    if (!copiedPath) {
      slicerNode.parentNode?.removeChild(slicerNode);
      continue;
    }
    const nextId = addRelationship(
      options.savedRelationships,
      SLICER_REL_TYPE,
      relativeArchivePath(options.savedSheetPath, copiedPath),
      null,
    );
    if (!nextId) {
      slicerNode.parentNode?.removeChild(slicerNode);
      continue;
    }
    setRelationshipId(slicerNode, nextId);
    copied += 1;
  }
  if (copied === 0) {
    return;
  }
  const extLst = ensureExtLst(options.savedDocument, (node) => insertSheetChild(options.savedDocument, node));
  if (uri) {
    removeExtByUri(extLst, uri);
  }
  extLst.appendChild(importedExt);
}

function applyCommentParts(options: {
  sourceArchive: ArchiveEntries;
  sourceDocument: XMLDocument;
  sourceRelationships: Map<string, RelationshipSource>;
  sourceSheetPath: string;
  savedArchive: ArchiveEntries;
  savedDocument: XMLDocument;
  savedRelationships: XMLDocument;
  savedSheetPath: string;
}) {
  const sourceLegacyDrawing = getLocalElements(options.sourceDocument, "legacyDrawing")[0] ?? null;
  removeSheetElement(options.savedDocument, "legacyDrawing");
  removeRelationshipsByType(
    options.savedRelationships,
    new Set([COMMENTS_REL_TYPE, VML_DRAWING_REL_TYPE])
  );

  const commentsRelationship = [...options.sourceRelationships.values()].find(
    (relationship) => relationship.type === COMMENTS_REL_TYPE
  );
  if (commentsRelationship) {
    const commentsPath = resolvePartPath(options.sourceSheetPath, commentsRelationship.target);
    const copied = copySourcePart(
      options.savedArchive,
      options.sourceArchive,
      commentsPath,
      "xl/comments",
      "commentAuthored",
      "xml"
    );
    if (copied) {
      ensureContentTypeOverride(options.savedArchive, copied, COMMENTS_CONTENT_TYPE);
      addRelationship(
        options.savedRelationships,
        COMMENTS_REL_TYPE,
        relativeArchivePath(options.savedSheetPath, copied),
        null
      );
    }
  }

  if (!sourceLegacyDrawing) {
    return;
  }
  const relationshipId = getRelationshipId(sourceLegacyDrawing);
  const relationship = relationshipId
    ? options.sourceRelationships.get(relationshipId)
    : undefined;
  if (!relationship) {
    return;
  }
  const vmlPath = resolvePartPath(options.sourceSheetPath, relationship.target);
  const copiedVml = copySourcePart(
    options.savedArchive,
    options.sourceArchive,
    vmlPath,
    "xl/drawings",
    "vmlDrawingAuthored",
    "vml"
  );
  if (!copiedVml) {
    return;
  }
  ensureContentTypeDefault(options.savedArchive, "vml", VML_CONTENT_TYPE);
  const nextId = addRelationship(
    options.savedRelationships,
    relationship.type,
    relativeArchivePath(options.savedSheetPath, copiedVml),
    relationship.targetMode
  );
  if (!nextId) {
    return;
  }
  const imported = options.savedDocument.importNode(sourceLegacyDrawing, true) as Element;
  setRelationshipId(imported, nextId);
  insertSheetChild(options.savedDocument, imported);
}

function ensureContentTypesDocument(archive: ArchiveEntries) {
  const existing = readArchiveText(archive, "[Content_Types].xml");
  if (existing) {
    const document = parseXml(existing);
    if (document?.documentElement) {
      return document;
    }
  }
  const document = parseXml(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${CONTENT_TYPES_NS}"></Types>`
  );
  if (document) {
    archive["[Content_Types].xml"] = strToU8(serializeXml(document));
  }
  return document;
}

function createSheetElement(document: XMLDocument, localName: string) {
  const root = document.documentElement;
  return document.createElementNS(root?.namespaceURI ?? "", localName);
}

const OOXML_DEFAULT_PAGE_MARGINS: Record<string, number> = {
  bottom: 1,
  footer: 0.5,
  header: 0.5,
  left: 0.75,
  right: 0.75,
  top: 1,
};

const DEFAULT_PAGE_SETUP_ATTRIBUTES: Record<string, string> = {
  blackAndWhite: "0",
  cellComments: "none",
  draft: "0",
  errors: "displayed",
  firstPageNumber: "1",
  fitToHeight: "1",
  fitToWidth: "1",
  orientation: "portrait",
  pageOrder: "downThenOver",
  paperSize: "1",
  scale: "100",
  useFirstPageNumber: "0",
};

function hasTruthyAttribute(node: Element | null, name: string) {
  const value = node?.getAttribute(name);
  return value === "1" || value?.toLowerCase() === "true";
}

/** Whether the source worksheet carries print settings the engine drops. */
function hasAuthoredPrintState(document: XMLDocument) {
  if (hasTruthyAttribute(getFirstDescendant(document, "pageSetUpPr"), "fitToPage")) {
    return true;
  }
  const options = getLocalElements(document, "printOptions")[0] ?? null;
  if (options && ["gridLines", "headings", "horizontalCentered", "verticalCentered"]
    .some((name) => hasTruthyAttribute(options, name))) {
    return true;
  }
  const margins = getLocalElements(document, "pageMargins")[0] ?? null;
  if (margins) {
    for (const [name, fallback] of Object.entries(OOXML_DEFAULT_PAGE_MARGINS)) {
      const value = Number(margins.getAttribute(name) ?? fallback);
      if (Number.isFinite(value) && Math.abs(value - fallback) > 1e-6) return true;
    }
  }
  const setup = getLocalElements(document, "pageSetup")[0] ?? null;
  if (setup) {
    for (const [name, fallback] of Object.entries(DEFAULT_PAGE_SETUP_ATTRIBUTES)) {
      const value = setup.getAttribute(name);
      if (value != null && value !== fallback) return true;
    }
    if (setup.hasAttribute("horizontalDpi") || setup.hasAttribute("verticalDpi") || setup.hasAttribute("copies")) {
      return true;
    }
  }
  for (const kind of ["rowBreaks", "colBreaks"] as const) {
    const container = getLocalElements(document, kind)[0] ?? null;
    if (container && getChildElements(container, "brk").length > 0) return true;
  }
  return false;
}

/**
 * Copies the worksheet's print set-up into engine-saved bytes.
 *
 * The engine keeps `pageSetup` and the manual breaks but drops `pageMargins`,
 * `printOptions` and `sheetPr/pageSetUpPr`, so an exported file would print
 * with different margins and without fit-to-page. The source archive owns
 * those records, exactly like the other authored fragments.
 */
function applyPageLayout(sourceDocument: XMLDocument, savedDocument: XMLDocument) {
  for (const localName of PAGE_LAYOUT_ELEMENTS) {
    removeSheetElement(savedDocument, localName);
    const sourceNode = getLocalElements(sourceDocument, localName)[0] ?? null;
    if (!sourceNode) continue;
    insertSheetChild(savedDocument, savedDocument.importNode(sourceNode, true) as Element);
  }

  removeSheetElement(savedDocument, "pageSetUpPr");
  const sourceFitToPage = getFirstDescendant(sourceDocument, "pageSetUpPr");
  if (!sourceFitToPage) return;
  const sheetPr = getLocalElements(savedDocument, "sheetPr")[0] ?? null;
  if (sheetPr) {
    sheetPr.appendChild(savedDocument.importNode(sourceFitToPage, true));
    return;
  }
  const created = createSheetElement(savedDocument, "sheetPr");
  created.appendChild(savedDocument.importNode(sourceFitToPage, true));
  insertSheetChild(savedDocument, created);
}

const PRINT_DEFINED_NAME_PATTERN = /^_xlnm\.(?:print_area|print_titles)$/i;

/**
 * Copies `_xlnm.Print_Area` / `_xlnm.Print_Titles` definitions into the saved
 * workbook part. The engine round-trips them today, but copying the source nodes
 * keeps multi-area print ranges and `localSheetId` assignments exact after a
 * sheet reorder.
 */
function applyWorkbookPrintDefinedNames(
  sourceArchive: ArchiveEntries,
  savedArchive: ArchiveEntries
) {
  const sourceXml = readArchiveText(sourceArchive, "xl/workbook.xml");
  const savedXml = readArchiveText(savedArchive, "xl/workbook.xml");
  if (!sourceXml || !savedXml) return false;
  const sourceDocument = parseXml(sourceXml);
  const savedDocument = parseXml(savedXml);
  const sourceRoot = sourceDocument?.documentElement;
  const savedRoot = savedDocument?.documentElement;
  if (!sourceDocument || !savedDocument || !sourceRoot || !savedRoot) return false;

  const sourceDefinitions = getLocalElements(sourceDocument, "definedName")
    .filter((node) => PRINT_DEFINED_NAME_PATTERN.test(node.getAttribute("name") ?? ""));
  if (sourceDefinitions.length === 0) return false;

  for (const node of getLocalElements(savedDocument, "definedName")) {
    if (PRINT_DEFINED_NAME_PATTERN.test(node.getAttribute("name") ?? "")) {
      node.parentNode?.removeChild(node);
    }
  }

  let container = getLocalElements(savedDocument, "definedNames")[0] ?? null;
  if (!container) {
    container = savedDocument.createElementNS(savedRoot.namespaceURI ?? "", "definedNames");
    const sheetsNode = getLocalElements(savedDocument, "sheets")[0] ?? null;
    if (sheetsNode?.nextSibling) savedRoot.insertBefore(container, sheetsNode.nextSibling);
    else savedRoot.appendChild(container);
  }
  for (const definition of sourceDefinitions) {
    container.appendChild(savedDocument.importNode(definition, true));
  }
  savedArchive["xl/workbook.xml"] = strToU8(serializeXml(savedDocument));
  return true;
}

/**
 * Copies authored print set-up from the source archive into saved bytes.
 *
 * Workbooks whose page margins and page set-up are still the OOXML defaults
 * are left untouched, so exporting an unmodified workbook stays byte-for-byte
 * identical to the engine's output.
 */
const WORKBOOK_CHILD_ORDER = [
  "fileVersion",
  "fileSharing",
  "workbookPr",
  "workbookProtection",
  "bookViews",
  "sheets",
  "functionGroups",
  "externalReferences",
  "definedNames",
  "calcPr",
  "oleSize",
  "customWorkbookViews",
  "pivotCaches",
  "smartTagPr",
  "smartTagTypes",
  "webPublishing",
  "fileRecoveryPr",
  "webPublishObjects",
  "extLst",
];

function insertWorkbookChild(document: XMLDocument, node: Element) {
  const root = document.documentElement;
  if (!root) {
    return;
  }
  const rank = WORKBOOK_CHILD_ORDER.indexOf(node.localName);
  if (rank < 0) {
    root.appendChild(node);
    return;
  }
  for (const child of Array.from(root.childNodes)) {
    if (child.nodeType !== 1) {
      continue;
    }
    const childRank = WORKBOOK_CHILD_ORDER.indexOf((child as Element).localName);
    if (childRank > rank) {
      root.insertBefore(node, child);
      return;
    }
  }
  root.appendChild(node);
}

function pivotPartContentType(path: string) {
  if (/^xl\/pivotTables\/[^/]+\.xml$/i.test(path)) return PIVOT_TABLE_CONTENT_TYPE;
  if (/^xl\/pivotCache\/pivotCacheRecords[^/]*\.xml$/i.test(path)) return PIVOT_CACHE_RECORDS_CONTENT_TYPE;
  if (/^xl\/pivotCache\/[^/]+\.xml$/i.test(path)) return PIVOT_CACHE_DEFINITION_CONTENT_TYPE;
  if (/^xl\/slicers\/[^/]+\.xml$/i.test(path)) return SLICER_CONTENT_TYPE;
  if (/^xl\/slicerCaches\/[^/]+\.xml$/i.test(path)) return SLICER_CACHE_CONTENT_TYPE;
  return null;
}

function copyPivotPackageParts(savedArchive: ArchiveEntries, sourceArchive: ArchiveEntries) {
  for (const path of Object.keys(sourceArchive)) {
    if (!/^xl\/(pivotCache|pivotTables|slicers|slicerCaches)\//i.test(path)) {
      continue;
    }
    savedArchive[path] = cloneBytes(sourceArchive[path]);
    const contentType = pivotPartContentType(path);
    if (contentType) {
      ensureContentTypeOverride(savedArchive, path, contentType);
    }
  }
}

/**
 * Merges workbook-level pivot caches and slicer caches into saved bytes.
 *
 * The workbook part references cache definitions and slicer caches through
 * relationships, and Excel's slicer list lives in the workbook extLst. Both are
 * dropped by the engine's save, so they are re-imported here after the sheet
 * loop.
 */
function applyWorkbookPivotParts(sourceArchive: ArchiveEntries, savedArchive: ArchiveEntries) {
  const sourceXml = readArchiveText(sourceArchive, "xl/workbook.xml");
  if (!sourceXml) {
    return false;
  }
  const sourceDocument = parseXml(sourceXml);
  if (!sourceDocument?.documentElement) {
    return false;
  }
  const sourcePivotCaches = getLocalElements(sourceDocument, "pivotCaches")[0] ?? null;
  const sourceSlicerCaches = getLocalElements(sourceDocument, "slicerCaches")[0] ?? null;
  if (!sourcePivotCaches && !sourceSlicerCaches) {
    return false;
  }

  const savedXml = readArchiveText(savedArchive, "xl/workbook.xml");
  const savedDocument = savedXml ? parseXml(savedXml) : null;
  if (!savedDocument?.documentElement) {
    return false;
  }
  const savedRelsPath = "xl/_rels/workbook.xml.rels";
  const savedRelationships = ensureRelationshipsXml(savedArchive, savedRelsPath);
  if (!savedRelationships) {
    return false;
  }
  const sourceRelationships = readRawRelationships(sourceArchive, savedRelsPath);
  copyPivotPackageParts(savedArchive, sourceArchive);
  ensureRelationshipNamespace(savedDocument);
  let changed = false;

  if (sourcePivotCaches) {
    removeSheetElement(savedDocument, "pivotCaches");
    removeRelationshipsByType(savedRelationships, new Set([PIVOT_CACHE_DEFINITION_REL_TYPE]));
    const importedCaches = savedDocument.importNode(sourcePivotCaches, true) as Element;
    let copied = 0;
    for (const cacheNode of getChildElements(importedCaches, "pivotCache")) {
      const relationshipId = getRelationshipId(cacheNode);
      const relationship = relationshipId ? sourceRelationships.get(relationshipId) : undefined;
      if (!relationship || relationship.type !== PIVOT_CACHE_DEFINITION_REL_TYPE) {
        cacheNode.parentNode?.removeChild(cacheNode);
        continue;
      }
      const targetPath = resolvePartPath("xl/workbook.xml", relationship.target);
      const nextId = addRelationship(
        savedRelationships,
        PIVOT_CACHE_DEFINITION_REL_TYPE,
        relativeArchivePath("xl/workbook.xml", targetPath),
        null,
      );
      if (!nextId) {
        cacheNode.parentNode?.removeChild(cacheNode);
        continue;
      }
      setRelationshipId(cacheNode, nextId);
      copied += 1;
    }
    if (copied > 0) {
      insertWorkbookChild(savedDocument, importedCaches);
      changed = true;
    }
  }

  if (sourceSlicerCaches) {
    const sourceExt = sourceSlicerCaches.parentNode;
    if (sourceExt && "localName" in sourceExt && (sourceExt as Element).localName === "ext") {
      const importedExt = savedDocument.importNode(sourceExt as Element, true) as Element;
      const uri = importedExt.getAttribute("uri") ?? "";
      let copied = 0;
      for (const cacheNode of getLocalElements(importedExt, "slicerCache")) {
        const relationshipId = getRelationshipId(cacheNode);
        const relationship = relationshipId ? sourceRelationships.get(relationshipId) : undefined;
        if (!relationship || relationship.type !== SLICER_CACHE_REL_TYPE) {
          cacheNode.parentNode?.removeChild(cacheNode);
          continue;
        }
        const targetPath = resolvePartPath("xl/workbook.xml", relationship.target);
        const nextId = addRelationship(
          savedRelationships,
          SLICER_CACHE_REL_TYPE,
          relativeArchivePath("xl/workbook.xml", targetPath),
          null,
        );
        if (!nextId) {
          cacheNode.parentNode?.removeChild(cacheNode);
          continue;
        }
        setRelationshipId(cacheNode, nextId);
        copied += 1;
      }
      if (copied > 0) {
        const extLst = ensureExtLst(savedDocument, (node) => insertWorkbookChild(savedDocument, node));
        if (uri) {
          removeExtByUri(extLst, uri);
        }
        extLst.appendChild(importedExt);
        changed = true;
      }
    }
  }

  if (!changed) {
    return false;
  }
  savedArchive["xl/workbook.xml"] = strToU8(serializeXml(savedDocument));
  savedArchive[savedRelsPath] = strToU8(serializeXml(savedRelationships));
  return true;
}

function applyWorksheetPrintLayouts(
  sourceArchive: ArchiveEntries,
  savedArchive: ArchiveEntries
) {
  const sourceSheets = parseWorkbookSheets(sourceArchive);
  const savedSheets = parseWorkbookSheets(savedArchive);
  let changed = false;
  for (let index = 0; index < Math.min(sourceSheets.length, savedSheets.length); index += 1) {
    const sourceSheet = sourceSheets[index];
    const savedSheet = savedSheets[index];
    const sourceXml = readArchiveText(sourceArchive, sourceSheet.path);
    const savedXml = readArchiveText(savedArchive, savedSheet.path);
    if (!sourceXml || !savedXml) continue;
    const sourceDocument = parseXml(sourceXml);
    const savedDocument = parseXml(savedXml);
    if (!sourceDocument?.documentElement || !savedDocument?.documentElement) continue;
    if (!hasAuthoredPrintState(sourceDocument)) continue;
    applyPageLayout(sourceDocument, savedDocument);
    savedArchive[savedSheet.path] = strToU8(serializeXml(savedDocument));
    changed = true;
  }
  return changed;
}

const STYLE_CHILD_ORDER = [
  "numFmts",
  "fonts",
  "fills",
  "borders",
  "cellStyleXfs",
  "cellXfs",
  "cellStyles",
  "dxfs",
  "tableStyles",
  "colors",
  "extLst",
];

function replaceStyleElement(
  document: XMLDocument,
  localName: string,
  sourceNode: Element,
) {
  const imported = document.importNode(sourceNode, true) as Element;
  const existing = getLocalElements(document, localName)[0] ?? null;
  if (existing?.parentNode) {
    existing.parentNode.replaceChild(imported, existing);
    return;
  }
  const root = document.documentElement;
  if (!root) {
    return;
  }
  const rank = STYLE_CHILD_ORDER.indexOf(localName);
  if (rank >= 0) {
    for (const child of Array.from(root.childNodes)) {
      if (child.nodeType !== 1) {
        continue;
      }
      const childRank = STYLE_CHILD_ORDER.indexOf((child as Element).localName);
      if (childRank > rank) {
        root.insertBefore(imported, child);
        return;
      }
    }
  }
  root.appendChild(imported);
}

/**
 * Carries workbook-defined table styles and their differential styles into the
 * engine's saved `xl/styles.xml`.
 *
 * The engine rewrites styles.xml from the cell styles it knows about and emits
 * `<dxfs count="0"/>` / `<tableStyles count="0"/>`. Custom table styles and
 * conditional-format `dxf`s are authored OOXML that the engine does not model,
 * so without this merge an exported workbook loses them and the table falls
 * back to a built-in style (or no conditional formatting). Replacing the dxfs
 * list keeps the source `dxfId` indexes valid for both tables and conditional
 * formats copied elsewhere in this merge.
 */
function applyWorkbookTableStyles(
  sourceArchive: ArchiveEntries,
  savedArchive: ArchiveEntries,
) {
  const sourceXml = readArchiveText(sourceArchive, "xl/styles.xml");
  const savedXml = readArchiveText(savedArchive, "xl/styles.xml");
  if (!sourceXml || !savedXml) {
    return false;
  }
  const sourceDocument = parseXml(sourceXml);
  const savedDocument = parseXml(savedXml);
  if (!sourceDocument?.documentElement || !savedDocument?.documentElement) {
    return false;
  }
  const sourceDxfs = getLocalElements(sourceDocument, "dxfs")[0] ?? null;
  const sourceTableStyles = getLocalElements(sourceDocument, "tableStyles")[0] ?? null;
  const hasDxfs = Boolean(sourceDxfs && getChildElements(sourceDxfs, "dxf").length > 0);
  const hasTableStyles = Boolean(
    sourceTableStyles && getChildElements(sourceTableStyles, "tableStyle").length > 0,
  );
  if (!hasDxfs && !hasTableStyles) {
    return false;
  }
  if (hasDxfs && sourceDxfs) {
    replaceStyleElement(savedDocument, "dxfs", sourceDxfs);
  }
  if (hasTableStyles && sourceTableStyles) {
    replaceStyleElement(savedDocument, "tableStyles", sourceTableStyles);
  }
  savedArchive["xl/styles.xml"] = strToU8(serializeXml(savedDocument));
  return true;
}

/**
 * Writes the authored worksheet objects — freeze pane, hyperlinks, conditional
 * formats, data validations, comments, tables and print set-up — into freshly
 * saved workbook bytes.
 *
 * `savedBytes` carry the WASM engine's authoritative cell data, which drops all
 * of those OOXML constructs. This merge copies the authored fragments from the
 * source archive so an exported file, an undo snapshot or a reload keeps them.
 *
 * @returns Patched workbook bytes, or a copy of `savedBytes` when nothing is
 * authored or the patch cannot be applied safely.
 */
export function mergeWorkbookSheetAuthoringAssets(
  savedBytes: Uint8Array,
  sourceArchive: ArchiveEntries | null | undefined
) {
  if (!sourceArchive || !sourceArchive["xl/workbook.xml"]) {
    return cloneBytes(savedBytes);
  }

  try {
    const sourceSheets = parseWorkbookSheets(sourceArchive);
    if (sourceSheets.length === 0) {
      return cloneBytes(savedBytes);
    }
    const savedArchive = unzipSync(savedBytes);
    const savedSheets = parseWorkbookSheets(savedArchive);
    if (savedSheets.length === 0) {
      return cloneBytes(savedBytes);
    }
    ensureContentTypesDocument(savedArchive);

    let changed = false;
    const sheetCount = Math.min(sourceSheets.length, savedSheets.length);
    for (let index = 0; index < sheetCount; index += 1) {
      const sourceSheet = sourceSheets[index];
      const savedSheet = savedSheets[index];
      const sourceXml = readArchiveText(sourceArchive, sourceSheet.path);
      const savedXml = readArchiveText(savedArchive, savedSheet.path);
      if (!sourceXml || !savedXml || !AUTHORING_FRAGMENT_PATTERN.test(sourceXml)) {
        continue;
      }
      const sourceDocument = parseXml(sourceXml);
      const savedDocument = parseXml(savedXml);
      if (!sourceDocument?.documentElement || !savedDocument?.documentElement) {
        continue;
      }
      const sourceRelationships = readRawRelationships(
        sourceArchive,
        relsPathForDocument(sourceSheet.path)
      );
      const savedRelsPath = relsPathForDocument(savedSheet.path);
      const savedRelationships = ensureRelationshipsXml(savedArchive, savedRelsPath);
      if (!savedRelationships) {
        continue;
      }

      applyFreezePane(sourceDocument, savedDocument);
      applyHyperlinks(sourceDocument, sourceRelationships, savedDocument, savedRelationships);
      applyConditionalFormats(sourceDocument, savedDocument);
      applyDataValidations(sourceDocument, savedDocument);
      applyTables({
        savedArchive,
        savedDocument,
        savedRelationships,
        savedSheetPath: savedSheet.path,
        sourceArchive,
        sourceDocument,
        sourceRelationships,
        sourceSheetPath: sourceSheet.path
      });
      applyPivotTableParts({
        savedArchive,
        savedDocument,
        savedRelationships,
        savedSheetPath: savedSheet.path,
        sourceArchive,
        sourceDocument,
        sourceRelationships,
        sourceSheetPath: sourceSheet.path
      });
      applySlicerParts({
        savedArchive,
        savedDocument,
        savedRelationships,
        savedSheetPath: savedSheet.path,
        sourceArchive,
        sourceDocument,
        sourceRelationships,
        sourceSheetPath: sourceSheet.path
      });
      applyCommentParts({
        savedArchive,
        savedDocument,
        savedRelationships,
        savedSheetPath: savedSheet.path,
        sourceArchive,
        sourceDocument,
        sourceRelationships,
        sourceSheetPath: sourceSheet.path
      });

      savedArchive[savedSheet.path] = strToU8(serializeXml(savedDocument));
      savedArchive[savedRelsPath] = strToU8(serializeXml(savedRelationships));
      changed = true;
    }

    const tableStylesChanged = applyWorkbookTableStyles(sourceArchive, savedArchive);
    if (tableStylesChanged) changed = true;
    const printLayoutsChanged = applyWorksheetPrintLayouts(sourceArchive, savedArchive);
    if (printLayoutsChanged) changed = true;
    const pivotPartsChanged = applyWorkbookPivotParts(sourceArchive, savedArchive);
    if (pivotPartsChanged) changed = true;
    const workbookChanged = applyWorkbookPrintDefinedNames(sourceArchive, savedArchive);
    if (workbookChanged) changed = true;
    return changed ? zipSync(savedArchive, { level: 6 }) : cloneBytes(savedBytes);
  } catch {
    return cloneBytes(savedBytes);
  }
}
