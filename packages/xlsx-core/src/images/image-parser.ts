import type { Workbook } from "@dukelib/sheets-wasm";
import { strFromU8 } from "fflate";
import type {
  XlsxImage,
  XlsxShape,
} from "../types";

// ── Constants ───────────────────────────────────────────────────────────────

export const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
export const SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
export const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
export const DRAWING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
export const VML_DRAWING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
export const CTRL_PROP_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp";
export const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
export const HYPERLINK_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
export const DRAWING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawing+xml";
export const EMU_PER_PIXEL = 9525;

// ── Types ───────────────────────────────────────────────────────────────────

export type DukeWorksheet = ReturnType<Workbook["getSheet"]>;

export type ArchiveEntries = Record<string, Uint8Array>;

export type ContentTypesState = {
  defaultEntries: Map<string, string>;
  overrideEntries: Map<string, string>;
};

export type RelationshipRecord = {
  id: string;
  target: string;
  targetMode: string | null;
  type: string;
};

export type DrawingRectEmu = {
  cx: number;
  cy: number;
  x: number;
  y: number;
};

export type GroupTransform = {
  chCx: number;
  chCy: number;
  chX: number;
  chY: number;
  cx: number;
  cy: number;
  scaleX: number;
  scaleY: number;
  x: number;
  y: number;
};

export type XlsxImageAttachment = {
  drawingPath: string;
  drawingRelsPath: string | null;
  mediaPaths: string[];
};

export type WorkbookImageOrigin = {
  anchorIndex: number;
  workbookSheetIndex: number;
};

export type WorkbookImageSheetOrigin = {
  attachments: XlsxImageAttachment[];
  workbookSheetIndex: number;
};

// ── Archive path utilities ──────────────────────────────────────────────────

export function normalizeArchivePath(path: string) {
  return path.replace(/\\/g, "/").replace(/^\/+/, "");
}

export function joinArchivePath(...parts: string[]) {
  return normalizeArchivePath(parts.join("/"));
}

export function dirname(path: string) {
  const normalized = normalizeArchivePath(path);
  const lastSlash = normalized.lastIndexOf("/");
  return lastSlash >= 0 ? normalized.slice(0, lastSlash) : "";
}

export function resolveArchiveTarget(baseDocumentPath: string, target: string) {
  if (!target) {
    return normalizeArchivePath(baseDocumentPath);
  }

  if (target.startsWith("#")) {
    return target;
  }

  if (target.startsWith("/")) {
    return normalizeArchivePath(target);
  }

  const baseParts = dirname(baseDocumentPath).split("/").filter(Boolean);
  for (const segment of target.split("/")) {
    if (!segment || segment === ".") {
      continue;
    }
    if (segment === "..") {
      baseParts.pop();
      continue;
    }
    baseParts.push(segment);
  }

  return normalizeArchivePath(baseParts.join("/"));
}

export function relativeArchivePath(fromDocumentPath: string, toPath: string) {
  const fromParts = dirname(fromDocumentPath).split("/").filter(Boolean);
  const toParts = normalizeArchivePath(toPath).split("/").filter(Boolean);
  let shared = 0;
  while (shared < fromParts.length && shared < toParts.length && fromParts[shared] === toParts[shared]) {
    shared += 1;
  }

  const upSegments = fromParts.slice(shared).map(() => "..");
  const downSegments = toParts.slice(shared);
  return [...upSegments, ...downSegments].join("/") || ".";
}

export function relsPathForDocument(documentPath: string) {
  const baseName = documentPath.split("/").pop();
  const parentDir = dirname(documentPath);
  return joinArchivePath(parentDir, "_rels", `${baseName}.rels`);
}

export function cloneBytes(bytes: Uint8Array) {
  const nextBytes = new Uint8Array(bytes.byteLength);
  nextBytes.set(bytes);
  return nextBytes;
}

// ── XML utilities ────────────────────────────────────────────────────────────

export function parseXml(xml: string) {
  const parser = new DOMParser();
  const document = parser.parseFromString(xml, "application/xml");
  if (document.querySelector("parsererror")) {
    return null;
  }
  return document;
}

export function serializeXml(document: XMLDocument) {
  return new XMLSerializer().serializeToString(document);
}

export function readArchiveText(archive: ArchiveEntries, path: string) {
  const entry = archive[path];
  return entry ? strFromU8(entry) : null;
}

export function isElementNode(node: Node | null | undefined): node is Element {
  return Boolean(node && node.nodeType === 1);
}

export function getLocalElements(parent: Document | Element, localName: string) {
  return Array.from(parent.getElementsByTagName("*")).filter((node): node is Element => isElementNode(node) && node.localName === localName);
}

export function getChildElements(parent: Element, localName: string) {
  return Array.from(parent.childNodes).filter((node): node is Element => isElementNode(node) && node.localName === localName);
}

export function getFirstChild(parent: Element, localName: string) {
  return getChildElements(parent, localName)[0] ?? null;
}

export function getFirstDescendant(parent: Document | Element, localName: string) {
  return getLocalElements(parent, localName)[0] ?? null;
}

export function getRelationshipId(element: Element) {
  return element.getAttributeNS(REL_NS, "id") ?? element.getAttribute("r:id") ?? element.getAttribute("id");
}

export function getEmbeddedRelationshipId(element: Element) {
  return element.getAttributeNS(REL_NS, "embed") ?? element.getAttribute("r:embed") ?? element.getAttribute("embed");
}

export function setChildText(parent: Element, localName: string, value: string) {
  const child = getFirstChild(parent, localName);
  if (child) {
    child.textContent = value;
  }
}

export function updateMarkerElement(element: Element | null, marker: { col: number; colOffsetEmu: number; row: number; rowOffsetEmu: number }) {
  if (!element) {
    return;
  }

  setChildText(element, "col", String(Math.max(0, marker.col)));
  setChildText(element, "colOff", String(Math.max(0, Math.round(marker.colOffsetEmu))));
  setChildText(element, "row", String(Math.max(0, marker.row)));
  setChildText(element, "rowOff", String(Math.max(0, Math.round(marker.rowOffsetEmu))));
}

// ── Content types ────────────────────────────────────────────────────────────

export function parseContentTypes(archive: ArchiveEntries): ContentTypesState {
  const xml = readArchiveText(archive, "[Content_Types].xml");
  const defaultEntries = new Map<string, string>();
  const overrideEntries = new Map<string, string>();
  if (!xml) {
    return { defaultEntries, overrideEntries };
  }

  const document = parseXml(xml);
  if (!document) {
    return { defaultEntries, overrideEntries };
  }

  for (const defaultNode of getLocalElements(document, "Default")) {
    const extension = defaultNode.getAttribute("Extension");
    const contentType = defaultNode.getAttribute("ContentType");
    if (extension && contentType) {
      defaultEntries.set(extension.toLowerCase(), contentType);
    }
  }

  for (const overrideNode of getLocalElements(document, "Override")) {
    const partName = overrideNode.getAttribute("PartName");
    const contentType = overrideNode.getAttribute("ContentType");
    if (partName && contentType) {
      overrideEntries.set(normalizeArchivePath(partName), contentType);
    }
  }

  return { defaultEntries, overrideEntries };
}

export function resolveContentType(contentTypes: ContentTypesState, path: string) {
  const normalized = normalizeArchivePath(path);
  const override = contentTypes.overrideEntries.get(normalized);
  if (override) {
    return override;
  }

  const extension = normalized.split(".").pop()?.toLowerCase();
  if (!extension) {
    return "application/octet-stream";
  }

  return contentTypes.defaultEntries.get(extension) ?? "application/octet-stream";
}

// ── Relationships ────────────────────────────────────────────────────────────

export function parseRelationships(archive: ArchiveEntries, relsPath: string, baseDocumentPath: string) {
  const xml = readArchiveText(archive, relsPath);
  const relationships = new Map<string, RelationshipRecord>();
  if (!xml) {
    return relationships;
  }

  const document = parseXml(xml);
  if (!document) {
    return relationships;
  }

  for (const relationshipNode of getLocalElements(document, "Relationship")) {
    const id = relationshipNode.getAttribute("Id");
    const target = relationshipNode.getAttribute("Target");
    const type = relationshipNode.getAttribute("Type");
    if (!id || !target || !type) {
      continue;
    }

    relationships.set(id, {
      id,
      target: resolveArchiveTarget(baseDocumentPath, target),
      targetMode: relationshipNode.getAttribute("TargetMode"),
      type
    });
  }

  return relationships;
}

// ── A1 reference parsing ────────────────────────────────────────────────────

export function parseColumnReference(reference: string) {
  let value = 0;
  for (const character of reference.toUpperCase()) {
    if (character < "A" || character > "Z") {
      return null;
    }
    value = value * 26 + (character.charCodeAt(0) - 64);
  }
  return value > 0 ? value - 1 : null;
}

export function parseA1CellReference(reference: string) {
  const match = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(reference.trim());
  if (!match) {
    return null;
  }

  const col = parseColumnReference(match[1] ?? "");
  const row = Number(match[2] ?? Number.NaN) - 1;
  if (col === null || !Number.isFinite(row) || row < 0) {
    return null;
  }

  return { col, row };
}

export function parseA1RangeReference(reference: string) {
  const [startRef, endRef] = reference.split(":");
  const start = parseA1CellReference(startRef ?? "");
  const end = parseA1CellReference(endRef ?? startRef ?? "");
  return start && end ? { end, start } : null;
}

// ── Image parsing ───────────────────────────────────────────────────────────

export function createImageAnchorNodes(document: XMLDocument) {
  return document.documentElement
    ? Array.from(document.documentElement.childNodes).filter(
        (node): node is Element =>
          isElementNode(node) &&
          (node.localName === "twoCellAnchor" || node.localName === "oneCellAnchor" || node.localName === "absoluteAnchor")
      )
    : [];
}

export function createImageSource(bytes: Uint8Array, mimeType: string, objectUrls: string[]) {
  if (typeof URL !== "undefined" && typeof Blob !== "undefined") {
    const objectUrl = URL.createObjectURL(new Blob([cloneBytes(bytes)], { type: mimeType }));
    objectUrls.push(objectUrl);
    return objectUrl;
  }

  let binary = "";
  for (const value of bytes) {
    binary += String.fromCharCode(value);
  }
  const base64 = typeof btoa === "function" ? btoa(binary) : "";
  return `data:${mimeType};base64,${base64}`;
}

export function parseMarker(node: Element | null) {
  if (!node) {
    return null;
  }

  const col = Number(getFirstChild(node, "col")?.textContent ?? 0);
  const row = Number(getFirstChild(node, "row")?.textContent ?? 0);
  const colOffsetEmu = Number(getFirstChild(node, "colOff")?.textContent ?? 0);
  const rowOffsetEmu = Number(getFirstChild(node, "rowOff")?.textContent ?? 0);

  return {
    col,
    colOffsetEmu,
    row,
    rowOffsetEmu
  };
}

export function parseAnchor(anchorNode: Element) {
  if (anchorNode.localName === "anchor") {
    const from = parseMarker(getFirstChild(anchorNode, "from"));
    const to = parseMarker(getFirstChild(anchorNode, "to"));
    return from && to ? { from, kind: "two-cell" as const, to } : null;
  }

  if (anchorNode.localName === "twoCellAnchor") {
    const from = parseMarker(getFirstChild(anchorNode, "from"));
    const to = parseMarker(getFirstChild(anchorNode, "to"));
    return from && to ? { from, kind: "two-cell" as const, to } : null;
  }

  if (anchorNode.localName === "oneCellAnchor") {
    const from = parseMarker(getFirstChild(anchorNode, "from"));
    const extNode = getFirstChild(anchorNode, "ext");
    return from && extNode
      ? {
          from,
          kind: "one-cell" as const,
          sizeEmu: {
            cx: Number(extNode.getAttribute("cx") ?? 0),
            cy: Number(extNode.getAttribute("cy") ?? 0)
          }
        }
      : null;
  }

  const positionNode = getFirstChild(anchorNode, "pos");
  const extNode = getFirstChild(anchorNode, "ext");
  return positionNode && extNode
    ? {
        kind: "absolute" as const,
        positionEmu: {
          x: Number(positionNode.getAttribute("x") ?? 0),
          y: Number(positionNode.getAttribute("y") ?? 0)
        },
        sizeEmu: {
          cx: Number(extNode.getAttribute("cx") ?? 0),
          cy: Number(extNode.getAttribute("cy") ?? 0)
        }
      }
    : null;
}

// ── Coordinate conversion ───────────────────────────────────────────────────

const DEFAULT_COL_WIDTH_EMU = 64 * EMU_PER_PIXEL;
const DEFAULT_ROW_HEIGHT_EMU = 20 * EMU_PER_PIXEL;

export function anchorToRect(anchor: XlsxImage["anchor"]): DrawingRectEmu {
  if (anchor.kind === "absolute") {
    return {
      cx: anchor.sizeEmu.cx,
      cy: anchor.sizeEmu.cy,
      x: anchor.positionEmu.x,
      y: anchor.positionEmu.y
    };
  }

  if (anchor.kind === "one-cell") {
    return {
      cx: anchor.sizeEmu.cx,
      cy: anchor.sizeEmu.cy,
      x: anchor.from.colOffsetEmu,
      y: anchor.from.rowOffsetEmu
    };
  }

  return {
    cx: Math.max(0, (anchor.to.col - anchor.from.col) * DEFAULT_COL_WIDTH_EMU + anchor.to.colOffsetEmu - anchor.from.colOffsetEmu),
    cy: Math.max(0, (anchor.to.row - anchor.from.row) * DEFAULT_ROW_HEIGHT_EMU + anchor.to.rowOffsetEmu - anchor.from.rowOffsetEmu),
    x: anchor.from.colOffsetEmu,
    y: anchor.from.rowOffsetEmu
  };
}

export function parseTransformRect(xfrmNode: Element | null) {
  if (!xfrmNode) {
    return null;
  }

  const offNode = getFirstChild(xfrmNode, "off");
  const extNode = getFirstChild(xfrmNode, "ext");
  if (!offNode || !extNode) {
    return null;
  }

  return {
    cx: Number(extNode.getAttribute("cx") ?? 0),
    cy: Number(extNode.getAttribute("cy") ?? 0),
    flipH: xfrmNode.getAttribute("flipH") === "1",
    flipV: xfrmNode.getAttribute("flipV") === "1",
    rot: Number(xfrmNode.getAttribute("rot") ?? 0) / 60000,
    x: Number(offNode.getAttribute("x") ?? 0),
    y: Number(offNode.getAttribute("y") ?? 0)
  };
}

export function applyGroupTransform(rect: DrawingRectEmu, group: GroupTransform): DrawingRectEmu {
  return {
    cx: rect.cx * group.scaleX,
    cy: rect.cy * group.scaleY,
    x: group.x + (rect.x - group.chX) * group.scaleX,
    y: group.y + (rect.y - group.chY) * group.scaleY
  };
}

export function emuToPixels(value: number) {
  return value / EMU_PER_PIXEL;
}

export function pixelsToEmu(value: number) {
  return value * EMU_PER_PIXEL;
}

export function rectToAbsoluteAnchor(rect: DrawingRectEmu): XlsxImage["anchor"] {
  return {
    kind: "absolute",
    positionEmu: {
      x: rect.x,
      y: rect.y
    },
    sizeEmu: {
      cx: rect.cx,
      cy: rect.cy
    }
  };
}

// ── Color utilities ──────────────────────────────────────────────────────────

export function normalizeHexColor(value: string) {
  const hex = value.replace(/^#/, "");
  if (hex.length === 8) {
    return `#${hex.slice(2).toLowerCase()}`;
  }
  if (hex.length === 6) {
    return `#${hex.toLowerCase()}`;
  }
  return "#000000";
}

export function anchorFromNodeOrFallback(
  node: Element,
  fallbackAnchor: XlsxImage["anchor"],
  parentGroup: GroupTransform | null
) {
  const xfrmNode = getFirstDescendant(node, "xfrm");
  const rect = parseTransformRect(xfrmNode);
  if (rect) {
    const scaleX = parentGroup?.scaleX ?? 1;
    const scaleY = parentGroup?.scaleY ?? 1;
    return {
      anchor: parentGroup ? rectToAbsoluteAnchor(applyGroupTransform(rect, parentGroup)) : fallbackAnchor,
      flipH: rect.flipH,
      flipV: rect.flipV,
      rotationDeg: rect.rot,
      scaleX,
      scaleY
    };
  }

  return {
    anchor: fallbackAnchor,
    flipH: false,
    flipV: false,
    rotationDeg: 0,
    scaleX: parentGroup?.scaleX ?? 1,
    scaleY: parentGroup?.scaleY ?? 1
  };
}
