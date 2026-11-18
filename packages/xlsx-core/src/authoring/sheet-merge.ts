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
  /<(?:[A-Za-z0-9]+:)?(pane|hyperlinks|conditionalFormatting|dataValidations|tableParts|legacyDrawing)[\s>/]/;

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

/**
 * Writes the authored worksheet objects — freeze pane, hyperlinks, conditional
 * formats, data validations, comments and tables — into freshly saved workbook
 * bytes.
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

    return changed ? zipSync(savedArchive, { level: 6 }) : cloneBytes(savedBytes);
  } catch {
    return cloneBytes(savedBytes);
  }
}
