import type {
  XlsxChartDrawing,
  XlsxChartUserShape,
  XlsxChartUserShapeFill,
  XlsxChartUserShapeStroke,
  XlsxImageRect,
  XlsxShapeParagraph,
  XlsxShapeTextBox,
  XlsxShapeTextRun,
  XlsxThemePalette,
} from "../types";

import {
  EMU_PER_PIXEL,
  findFirstChartColorElement,
  isElementNode,
  resolveChartColorNode,
  resolveChartFillColor,
  resolveChartLineStyle,
  resolveThemeTypeface,
} from "./chart-colors";

import {
  dirname,
  getFirstLocalChild,
  getFirstLocalDescendant,
  getLocalChildren,
  getLocalDescendants,
  normalizeArchivePath,
  parseXml,
  readArchiveText,
  resolveRelationshipPath,
} from "./chart-xml-utils";

/** Relationship namespace used by `r:embed` / `r:id` references in DrawingML. */
const DRAWINGML_RELATIONSHIP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CHART_USER_SHAPES_REL_SUFFIX = "/chartUserShapes";

/**
 * Per-part byte cap for a picture embedded in a chart part or in one of its
 * `chartUserShapes` parts. Mirrors the upstream reader so a hostile package
 * cannot make the viewer materialise an arbitrarily large data URL.
 */
export const CHART_EMBEDDED_IMAGE_MAX_BYTES = 72 * 1024 * 1024;
/** Total byte cap for every image a single chart part references. */
export const CHART_EMBEDDED_IMAGE_TOTAL_MAX_BYTES = 100 * 1024 * 1024;
/** Byte cap for one `chartUserShapes` DrawingML part (8 MB, as upstream). */
export const CHART_USER_SHAPES_MAX_BYTES = 8 * 1024 * 1024;
/** Fallback chart drawing space (Excel's default 480 × 300 px chart size). */
const DEFAULT_CHART_EXTENT_EMU = {
  cx: 480 * EMU_PER_PIXEL,
  cy: 300 * EMU_PER_PIXEL,
};
const DEFAULT_TEXT_INSET_EMU = {
  bIns: 45720,
  lIns: 91440,
  rIns: 91440,
  tIns: 45720,
};
const DEFAULT_CHART_SHAPE_FONT_SIZE_PT = 10;
const BASE64_ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

export interface ChartEmbeddedMediaOptions {
  /** Override `CHART_EMBEDDED_IMAGE_TOTAL_MAX_BYTES`. */
  maxTotalImageBytes?: number;
  /** Override `CHART_EMBEDDED_IMAGE_MAX_BYTES`. */
  maxImageBytes?: number;
  /** Override `CHART_USER_SHAPES_MAX_BYTES`. */
  maxUserShapesBytes?: number;
  /** Theme palette used to resolve `schemeClr` fills, lines and text. */
  themePalette?: XlsxThemePalette | null;
  /**
   * Chart drawing-space size in EMU, used to normalise user-shape anchors to
   * chart fractions. Callers pass the chart graphic frame ext; when omitted the
   * parser falls back to the shapes' own bounding box.
   */
  extentEmu?: { cx: number; cy: number } | null;
}

export interface ChartEmbeddedMedia {
  drawings?: Record<string, XlsxChartDrawing>;
  images?: Record<string, string>;
}

const mediaCache = new WeakMap<Record<string, Uint8Array>, Map<string, ChartEmbeddedMedia>>();
const partImagesCache = new WeakMap<Record<string, Uint8Array>, Map<string, Record<string, string> | undefined>>();
const paletteIdentities = new WeakMap<object, number>();
let nextPaletteIdentity = 1;

function paletteIdentity(palette: XlsxThemePalette | null | undefined) {
  if (!palette) {
    return 0;
  }
  let identity = paletteIdentities.get(palette);
  if (identity === undefined) {
    identity = nextPaletteIdentity;
    nextPaletteIdentity += 1;
    paletteIdentities.set(palette, identity);
  }
  return identity;
}

function encodeBase64(bytes: Uint8Array) {
  let output = "";
  for (let index = 0; index < bytes.length; index += 3) {
    const first = bytes[index] ?? 0;
    const second = bytes[index + 1];
    const third = bytes[index + 2];
    const triple = (first << 16) | ((second ?? 0) << 8) | (third ?? 0);
    output += BASE64_ALPHABET[(triple >> 18) & 63];
    output += BASE64_ALPHABET[(triple >> 12) & 63];
    output += second === undefined ? "=" : BASE64_ALPHABET[(triple >> 6) & 63];
    output += third === undefined ? "=" : BASE64_ALPHABET[triple & 63];
  }
  return output;
}

function asciiSlice(bytes: Uint8Array, start: number, end: number) {
  let value = "";
  for (let index = start; index < Math.min(end, bytes.length); index += 1) {
    value += String.fromCharCode(bytes[index] ?? 0);
  }
  return value;
}

/**
 * Detects a supported raster image from its signature.
 *
 * The part name and `[Content_Types].xml` are not trusted: a `.png` entry that
 * holds HTML (or an SVG with script) must never become an `<image>` source. SVG
 * is deliberately unsupported so a chart part cannot inject active content.
 */
export function sniffChartImageMime(bytes: Uint8Array | null | undefined) {
  if (!bytes || bytes.length < 4) {
    return null;
  }
  if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) {
    return "image/png";
  }
  if (bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) {
    return "image/jpeg";
  }
  const header = asciiSlice(bytes, 0, 12);
  if (header.startsWith("GIF87a") || header.startsWith("GIF89a")) {
    return "image/gif";
  }
  if (header.startsWith("RIFF") && asciiSlice(bytes, 8, 12) === "WEBP") {
    return "image/webp";
  }
  if (bytes[0] === 0x42 && bytes[1] === 0x4d) {
    return "image/bmp";
  }
  if (header.startsWith("II*\0") || header.startsWith("MM\0*")) {
    return "image/tiff";
  }
  return null;
}

/**
 * @internal Exposed for the chart parser's unit tests.
 */
export function chartImageDataUrl(bytes: Uint8Array) {
  const mime = sniffChartImageMime(bytes);
  return mime ? `data:${mime};base64,${encodeBase64(bytes)}` : null;
}

function relsPathForPart(partPath: string) {
  const normalized = normalizeArchivePath(partPath);
  const fileName = normalized.split("/").pop() ?? "";
  return normalizeArchivePath(`${dirname(normalized)}/_rels/${fileName}.rels`);
}

/**
 * Resolves a relationship target to an archive key, refusing anything that
 * escapes the package (`../`, absolute OS paths, backslashes) or is empty.
 */
function resolveArchiveTarget(relsPath: string, target: string | null | undefined) {
  if (!target) {
    return null;
  }
  if (target.includes("\\") || target.trim() !== target) {
    return null;
  }
  if (/^[A-Za-z][A-Za-z\d+.-]*:/.test(target)) {
    // `http:`, `file:`, `javascript:` … — a package part never needs a scheme.
    return null;
  }
  const resolved = resolveRelationshipPath(relsPath, target);
  if (!resolved || resolved.startsWith("..") || resolved.startsWith("/")) {
    return null;
  }
  return resolved;
}

function readRelationshipId(node: Element | null | undefined) {
  if (!node) {
    return null;
  }
  return node.getAttributeNS(DRAWINGML_RELATIONSHIP_NS, "embed")
    ?? node.getAttribute("r:embed")
    ?? node.getAttribute("embed")
    ?? node.getAttributeNS(DRAWINGML_RELATIONSHIP_NS, "link")
    ?? node.getAttribute("r:link");
}

interface ChartRelationship {
  id: string;
  path: string;
  type: string;
}

function readChartRelationshipEntries(
  archive: Record<string, Uint8Array>,
  partPath: string
): ChartRelationship[] {
  const relsPath = relsPathForPart(partPath);
  const relsXml = readArchiveText(archive, relsPath);
  const relsDocument = relsXml ? parseXml(relsXml) : null;
  if (!relsDocument) {
    return [];
  }
  const relationships: ChartRelationship[] = [];
  for (const node of getLocalDescendants(relsDocument, "Relationship")) {
    const id = node.getAttribute("Id");
    const type = node.getAttribute("Type") ?? "";
    if (!id || node.getAttribute("TargetMode") === "External") {
      continue;
    }
    const path = resolveArchiveTarget(relsPath, node.getAttribute("Target"));
    if (!path) {
      continue;
    }
    relationships.push({ id, path, type });
  }
  return relationships;
}

function readPartImages(
  archive: Record<string, Uint8Array>,
  partPath: string,
  relationships: ChartRelationship[],
  options: ChartEmbeddedMediaOptions | undefined
): Record<string, string> | undefined {
  const maxImageBytes = options?.maxImageBytes ?? CHART_EMBEDDED_IMAGE_MAX_BYTES;
  const maxTotalImageBytes = options?.maxTotalImageBytes ?? CHART_EMBEDDED_IMAGE_TOTAL_MAX_BYTES;
  let archiveCache = partImagesCache.get(archive);
  if (!archiveCache) {
    archiveCache = new Map();
    partImagesCache.set(archive, archiveCache);
  }
  // Image bytes do not depend on the chart extent or theme, so a chart styled
  // twice (before and after its anchor is known) encodes each picture once.
  const cacheKey = `${partPath}|${maxImageBytes}:${maxTotalImageBytes}`;
  if (archiveCache.has(cacheKey)) {
    return archiveCache.get(cacheKey);
  }
  const images: Record<string, string> = {};
  let totalBytes = 0;
  for (const relationship of relationships) {
    if (!relationship.type.endsWith("/image")) {
      continue;
    }
    const bytes = archive[relationship.path];
    if (!bytes || bytes.length === 0 || bytes.length > maxImageBytes) {
      continue;
    }
    if (totalBytes + bytes.length > maxTotalImageBytes) {
      continue;
    }
    const dataUrl = chartImageDataUrl(bytes);
    if (!dataUrl) {
      continue;
    }
    totalBytes += bytes.length;
    images[relationship.id] = dataUrl;
  }
  const result = Object.keys(images).length > 0 ? images : undefined;
  archiveCache.set(cacheKey, result);
  return result;
}

/**
 * Reads the images and `chartUserShapes` parts linked from a chart part.
 *
 * The chart part itself keeps only relationship ids (`r:embed`), so the parsed
 * media is handed to the model as relationship-id maps plus fully parsed user
 * shapes. External relationships are ignored, image bytes are mime-sniffed and
 * every read is bounded so a malformed package degrades to "no image" instead
 * of allocating without limit.
 *
 * Results are memoised per archive object: chart styling runs on every workbook
 * refresh, and re-encoding a picture's base64 on each refresh would be pure
 * waste. The chart path, drawing extent, caps and theme palette are part of the
 * cache key.
 */
export function readChartEmbeddedMedia(
  archive: Record<string, Uint8Array>,
  chartPath: string | null | undefined,
  options?: ChartEmbeddedMediaOptions
): ChartEmbeddedMedia {
  if (!chartPath) {
    return {};
  }
  const cacheKey = `${chartPath}|${options?.extentEmu?.cx ?? 0}x${options?.extentEmu?.cy ?? 0}|`
    + `${options?.maxImageBytes ?? 0}:${options?.maxTotalImageBytes ?? 0}:${options?.maxUserShapesBytes ?? 0}|`
    + `${paletteIdentity(options?.themePalette)}`;
  let cache = mediaCache.get(archive);
  if (!cache) {
    cache = new Map();
    mediaCache.set(archive, cache);
  }
  const cached = cache.get(cacheKey);
  if (cached) {
    return cached;
  }
  const media = readChartEmbeddedMediaUncached(archive, chartPath, options);
  cache.set(cacheKey, media);
  return media;
}

function readChartEmbeddedMediaUncached(
  archive: Record<string, Uint8Array>,
  chartPath: string,
  options: ChartEmbeddedMediaOptions | undefined
): ChartEmbeddedMedia {
  const relationships = readChartRelationshipEntries(archive, chartPath);
  if (relationships.length === 0) {
    return {};
  }
  const images = readPartImages(archive, chartPath, relationships, options);
  const maxUserShapesBytes = options?.maxUserShapesBytes ?? CHART_USER_SHAPES_MAX_BYTES;
  const drawings: Record<string, XlsxChartDrawing> = {};
  for (const relationship of relationships) {
    if (!relationship.type.endsWith(CHART_USER_SHAPES_REL_SUFFIX)) {
      continue;
    }
    const bytes = archive[relationship.path];
    if (!bytes || bytes.length === 0 || bytes.length > maxUserShapesBytes) {
      continue;
    }
    const xml = readArchiveText(archive, relationship.path);
    if (!xml) {
      continue;
    }
    const partRelationships = readChartRelationshipEntries(archive, relationship.path);
    const partImages = readPartImages(archive, relationship.path, partRelationships, options);
    drawings[relationship.id] = {
      images: partImages,
      shapes: parseChartUserShapes(xml, {
        extentEmu: options?.extentEmu ?? null,
        images: partImages,
        themePalette: options?.themePalette ?? null,
      }),
      xml,
    };
  }
  return {
    drawings: Object.keys(drawings).length > 0 ? drawings : undefined,
    images,
  };
}

/**
 * Resolves the `a:blipFill` picture fill of a chart node (`spPr`, `spPr` under a
 * series/data point, or the chart-space shape properties) to a data URL.
 */
export function readChartBlipFillImage(
  node: Element | null | undefined,
  images: Record<string, string> | undefined
) {
  if (!node || !images) {
    return undefined;
  }
  const blipFill = node.localName === "blipFill" ? node : getFirstLocalChild(node, "blipFill");
  const blip = blipFill ? getFirstLocalDescendant(blipFill, "blip") : null;
  const relationshipId = readRelationshipId(blip);
  return relationshipId ? images[relationshipId] : undefined;
}

// ── chartUserShapes parsing ──────────────────────────────────────────────

interface ShapeRectEmu {
  cx: number;
  cy: number;
  x: number;
  y: number;
}

export interface ChartUserShapeParseOptions {
  /** Chart drawing-space size in EMU; falls back to the shapes' bounding box. */
  extentEmu?: { cx: number; cy: number } | null;
  /** Images of the `chartUserShapes` part, keyed by relationship id. */
  images?: Record<string, string>;
  themePalette?: XlsxThemePalette | null;
}

function readEmuValue(parent: Element | null, localName: string) {
  if (!parent) {
    return 0;
  }
  const child = getFirstLocalChild(parent, localName);
  const raw = child
    ? child.getAttribute(localName) ?? child.getAttribute("val") ?? child.textContent?.trim() ?? ""
    : parent.getAttribute(localName) ?? parent.textContent?.trim() ?? "";
  const value = Number(raw);
  return Number.isFinite(value) ? value : 0;
}

function readExtRect(anchorNode: Element | null, positionNode: Element | null, fromNode: Element | null) {
  if (!anchorNode) {
    return null;
  }
  const extNode = getFirstLocalChild(anchorNode, "ext");
  const posNode = getFirstLocalChild(anchorNode, "pos");
  const x = posNode ? readEmuValue(posNode, "x") : fromNode ? readEmuValue(fromNode, "x") : 0;
  const y = posNode ? readEmuValue(posNode, "y") : fromNode ? readEmuValue(fromNode, "y") : 0;
  const cx = readEmuValue(extNode, "cx");
  const cy = readEmuValue(extNode, "cy");
  return { cx, cy, x, y } satisfies ShapeRectEmu;
}

/**
 * Reads a chart-drawing anchor.
 *
 * `chartUserShapes` anchors use their own `cdr:from`/`cdr:to`/`cdr:pos`/`cdr:ext`
 * elements whose `x`/`y`/`cx`/`cy` values are absolute EMU inside the chart's
 * drawing space (unlike worksheet anchors, which address cells).
 */
export function readChartUserShapeAnchorRect(anchorNode: Element): ShapeRectEmu | null {
  const fromNode = getFirstLocalChild(anchorNode, "from");
  const toNode = getFirstLocalChild(anchorNode, "to");
  if (fromNode && toNode) {
    const x = readEmuValue(fromNode, "x");
    const y = readEmuValue(fromNode, "y");
    const toX = readEmuValue(toNode, "x");
    const toY = readEmuValue(toNode, "y");
    return {
      cx: Math.max(0, toX - x),
      cy: Math.max(0, toY - y),
      x,
      y,
    };
  }
  return readExtRect(anchorNode, null, fromNode);
}

function applyGroupTransform(rect: ShapeRectEmu, group: {
  child?: ShapeRectEmu | null;
  cx: number;
  cy: number;
  x: number;
  y: number;
}): ShapeRectEmu {
  const childRect = group.child;
  const scaleX = childRect && childRect.cx !== 0 ? group.cx / childRect.cx : 1;
  const scaleY = childRect && childRect.cy !== 0 ? group.cy / childRect.cy : 1;
  return {
    cx: rect.cx * scaleX,
    cy: rect.cy * scaleY,
    x: group.x + (childRect ? rect.x - childRect.x : rect.x) * scaleX,
    y: group.y + (childRect ? rect.y - childRect.y : rect.y) * scaleY,
  };
}

type ShapeRectTransform = (rect: ShapeRectEmu) => ShapeRectEmu;

function readGroupTransform(groupNode: Element): ShapeRectTransform | null {
  const shapeProperties = getFirstLocalChild(groupNode, "grpSpPr");
  const transform = shapeProperties ? getFirstLocalChild(shapeProperties, "xfrm") : null;
  if (!transform) {
    return null;
  }
  const offset = getFirstLocalChild(transform, "off");
  const extent = getFirstLocalChild(transform, "ext");
  const childOffset = getFirstLocalChild(transform, "chOff");
  const childExtent = getFirstLocalChild(transform, "chExt");
  const group = {
    child: childOffset && childExtent
      ? {
          cx: readEmuValue(childExtent, "cx"),
          cy: readEmuValue(childExtent, "cy"),
          x: readEmuValue(childOffset, "x"),
          y: readEmuValue(childOffset, "y"),
        }
      : null,
    cx: readEmuValue(extent, "cx"),
    cy: readEmuValue(extent, "cy"),
    x: readEmuValue(offset, "x"),
    y: readEmuValue(offset, "y"),
  };
  return (rect) => applyGroupTransform(rect, group);
}

function toFractionRect(rect: ShapeRectEmu, extent: { cx: number; cy: number }): XlsxImageRect {
  const clamp = (value: number, min: number, max: number) => Math.max(min, Math.min(max, value));
  // Shapes may hang off the chart area; keep a bounded range so a hostile part
  // cannot ask a renderer for an enormous box while still honouring moderate
  // overflow the way Excel paints it (clipped by the chart frame).
  return {
    height: clamp(extent.cy > 0 ? rect.cy / extent.cy : 0, 0, 2),
    left: clamp(extent.cx > 0 ? rect.x / extent.cx : 0, -1, 2),
    top: clamp(extent.cy > 0 ? rect.y / extent.cy : 0, -1, 2),
    width: clamp(extent.cx > 0 ? rect.cx / extent.cx : 0, 0, 2),
  };
}

function readGeometryAdjustments(presetGeometry: Element | null) {
  const adjustmentsNode = presetGeometry ? getFirstLocalChild(presetGeometry, "avLst") : null;
  if (!adjustmentsNode) {
    return undefined;
  }
  const adjustments: Record<string, number> = {};
  for (const adjustmentNode of getLocalChildren(adjustmentsNode, "gd")) {
    const name = adjustmentNode.getAttribute("name");
    const formula = adjustmentNode.getAttribute("fmla") ?? "";
    const match = /val\s+(-?\d+(?:\.\d+)?)/i.exec(formula);
    const value = match?.[1] === undefined ? Number.NaN : Number(match[1]);
    if (!name || !Number.isFinite(value)) {
      continue;
    }
    adjustments[name] = value;
  }
  return Object.keys(adjustments).length > 0 ? adjustments : undefined;
}

function readShapeFill(
  shapeNode: Element,
  styleNode: Element | null,
  images: Record<string, string> | undefined,
  themePalette?: XlsxThemePalette | null
): XlsxChartUserShapeFill | undefined {
  const shapeProperties = getFirstLocalChild(shapeNode, "spPr");
  // A picture (`xdr:pic` / `cdr:pic`) stores its `blipFill` next to `spPr`.
  const pictureFill = readChartBlipFillImage(getFirstLocalChild(shapeNode, "blipFill"), images)
    ?? readChartBlipFillImage(shapeProperties, images);
  if (pictureFill) {
    return { image: pictureFill };
  }
  if (shapeProperties && getFirstLocalChild(shapeProperties, "noFill")) {
    return { none: true };
  }
  const solidColor = resolveChartFillColor(shapeProperties, themePalette);
  if (solidColor) {
    return { color: solidColor };
  }
  const fillRefNode = styleNode ? getFirstLocalChild(styleNode, "fillRef") : null;
  const fillRefColor = fillRefNode
    ? resolveChartColorNode(
        findFirstChartColorElement(Array.from(fillRefNode.childNodes).find(isElementNode) ?? null),
        themePalette
      )
    : null;
  if (fillRefColor) {
    return { color: fillRefColor };
  }
  return undefined;
}

function readShapeStroke(
  shapeNode: Element,
  themePalette?: XlsxThemePalette | null
): XlsxChartUserShapeStroke | undefined {
  const shapeProperties = getFirstLocalChild(shapeNode, "spPr");
  if (!shapeProperties) {
    return undefined;
  }
  const lineStyle = resolveChartLineStyle(shapeProperties, themePalette);
  const lineNode = getFirstLocalChild(shapeProperties, "ln");
  const dash = lineNode ? getFirstLocalChild(lineNode, "prstDash")?.getAttribute("val") ?? undefined : undefined;
  if (!lineStyle.color && !lineStyle.hidden && !dash) {
    return undefined;
  }
  return {
    color: lineStyle.color ?? undefined,
    dash,
    hidden: lineStyle.hidden || undefined,
    widthPx: lineStyle.widthPx,
  };
}

interface ShapeTextStyle {
  bold?: boolean;
  color?: string;
  fontFamily?: string;
  fontSizePt?: number;
  italic?: boolean;
  underline?: boolean;
}

function hasShapeTextStyle(style: ShapeTextStyle) {
  return style.bold !== undefined
    || style.color !== undefined
    || style.fontFamily !== undefined
    || style.fontSizePt !== undefined
    || style.italic !== undefined
    || style.underline !== undefined;
}

function readShapeTextStyle(
  node: Element | null,
  themePalette?: XlsxThemePalette | null
): ShapeTextStyle | undefined {
  if (!node) {
    return undefined;
  }
  const style: ShapeTextStyle = {};
  const size = Number(node.getAttribute("sz") ?? Number.NaN);
  if (Number.isFinite(size) && size > 0) {
    style.fontSizePt = size / 100;
  }
  const bold = node.getAttribute("b");
  if (bold === "1" || bold === "true") {
    style.bold = true;
  } else if (bold === "0" || bold === "false") {
    style.bold = false;
  }
  const italic = node.getAttribute("i");
  if (italic === "1" || italic === "true") {
    style.italic = true;
  } else if (italic === "0" || italic === "false") {
    style.italic = false;
  }
  const underline = node.getAttribute("u");
  if (underline) {
    style.underline = underline !== "none";
  }
  const solidFill = getFirstLocalChild(node, "solidFill");
  const color = solidFill
    ? resolveChartColorNode(
        findFirstChartColorElement(Array.from(solidFill.childNodes).find(isElementNode) ?? null),
        themePalette
      )
    : null;
  if (color) {
    style.color = color;
  }
  const latinTypeface = getFirstLocalChild(node, "latin")?.getAttribute("typeface") ?? null;
  const typeface = resolveThemeTypeface(latinTypeface, themePalette);
  if (typeface) {
    style.fontFamily = typeface;
  }
  return hasShapeTextStyle(style) ? style : undefined;
}

function mergeShapeTextStyle(base: ShapeTextStyle, patch: ShapeTextStyle | undefined): ShapeTextStyle {
  if (!patch) {
    return base;
  }
  return {
    bold: patch.bold ?? base.bold,
    color: patch.color ?? base.color,
    fontFamily: patch.fontFamily ?? base.fontFamily,
    fontSizePt: patch.fontSizePt ?? base.fontSizePt,
    italic: patch.italic ?? base.italic,
    underline: patch.underline ?? base.underline,
  };
}

function mapParagraphAlignment(value: string | null | undefined): XlsxShapeParagraph["align"] {
  switch (value) {
    case "ctr":
      return "center";
    case "r":
      return "right";
    case "just":
    case "justLow":
    case "dist":
    case "thaiDist":
      return "justify";
    case "l":
      return "left";
    default:
      return undefined;
  }
}

function mapVerticalAlignment(value: string | null | undefined): XlsxShapeTextBox["verticalAlign"] {
  switch (value) {
    case "b":
      return "bottom";
    case "ctr":
      return "middle";
    case "t":
      return "top";
    default:
      return undefined;
  }
}

function parseShapeParagraphs(
  shapeNode: Element,
  themePalette?: XlsxThemePalette | null
): XlsxShapeParagraph[] {
  const textBody = getFirstLocalChild(shapeNode, "txBody");
  if (!textBody) {
    return [];
  }
  const styleNode = getFirstLocalChild(shapeNode, "style");
  const fontRefNode = styleNode ? getFirstLocalChild(styleNode, "fontRef") : null;
  const fontRefColor = fontRefNode
    ? resolveChartColorNode(
        findFirstChartColorElement(Array.from(fontRefNode.childNodes).find(isElementNode) ?? null),
        themePalette
      )
    : null;
  const defaultStyle: ShapeTextStyle = {
    color: fontRefColor ?? themePalette?.colorsByIndex[1] ?? "#000000",
    fontSizePt: DEFAULT_CHART_SHAPE_FONT_SIZE_PT,
  };
  const listStyleNode = getFirstLocalChild(textBody, "lstStyle");
  const paragraphs: XlsxShapeParagraph[] = [];
  for (const paragraphNode of getLocalChildren(textBody, "p")) {
    const paragraphProperties = getFirstLocalChild(paragraphNode, "pPr");
    const level = Number(paragraphProperties?.getAttribute("lvl") ?? 0);
    const levelKey = `lvl${Number.isFinite(level) ? level + 1 : 1}pPr`;
    const listLevelProperties = listStyleNode ? getFirstLocalChild(listStyleNode, levelKey) : null;
    const inheritedStyle = mergeShapeTextStyle(
      mergeShapeTextStyle(defaultStyle, readShapeTextStyle(
        listLevelProperties ? getFirstLocalChild(listLevelProperties, "defRPr") : null,
        themePalette
      )),
      readShapeTextStyle(
        paragraphProperties ? getFirstLocalChild(paragraphProperties, "defRPr") : null,
        themePalette
      )
    );
    const runs: XlsxShapeTextRun[] = [];
    for (const child of Array.from(paragraphNode.childNodes).filter(isElementNode)) {
      if (child.localName === "br") {
        runs.push({ ...inheritedStyle, text: "\n" });
        continue;
      }
      if (child.localName !== "r") {
        continue;
      }
      const text = getFirstLocalChild(child, "t")?.textContent ?? "";
      const runStyle = mergeShapeTextStyle(
        inheritedStyle,
        readShapeTextStyle(getFirstLocalChild(child, "rPr"), themePalette)
      );
      runs.push({
        bold: runStyle.bold,
        color: runStyle.color,
        fontFamily: runStyle.fontFamily,
        fontSizePt: runStyle.fontSizePt,
        italic: runStyle.italic,
        text,
        underline: runStyle.underline,
      });
    }
    if (runs.length === 0 && !getFirstLocalChild(paragraphNode, "endParaRPr")) {
      continue;
    }
    if (runs.length === 0) {
      runs.push({
        bold: inheritedStyle.bold,
        color: inheritedStyle.color,
        fontFamily: inheritedStyle.fontFamily,
        fontSizePt: inheritedStyle.fontSizePt,
        italic: inheritedStyle.italic,
        text: " ",
        underline: inheritedStyle.underline,
      });
    }
    paragraphs.push({
      align: mapParagraphAlignment(paragraphProperties?.getAttribute("algn")),
      runs,
    });
  }
  return paragraphs;
}

function parseShapeTextBox(shapeNode: Element): XlsxShapeTextBox | undefined {
  const textBody = getFirstLocalChild(shapeNode, "txBody");
  if (!textBody) {
    return undefined;
  }
  const bodyProperties = getFirstLocalChild(textBody, "bodyPr");
  const inset = (name: "bIns" | "lIns" | "rIns" | "tIns") => {
    const value = Number(bodyProperties?.getAttribute(name) ?? Number.NaN);
    return (Number.isFinite(value) ? value : DEFAULT_TEXT_INSET_EMU[name]) / EMU_PER_PIXEL;
  };
  return {
    horizontalAlign: bodyProperties?.getAttribute("anchorCtr") === "1" ? "center" : "left",
    insetPx: {
      bottom: inset("bIns"),
      left: inset("lIns"),
      right: inset("rIns"),
      top: inset("tIns"),
    },
    verticalAlign: mapVerticalAlignment(bodyProperties?.getAttribute("anchor")),
  };
}

function readShapeNodeRect(shapeNode: Element, anchorRect: ShapeRectEmu): ShapeRectEmu {
  const shapeProperties = getFirstLocalChild(shapeNode, "spPr");
  const transform = shapeProperties ? getFirstLocalChild(shapeProperties, "xfrm") : null;
  const offset = transform ? getFirstLocalChild(transform, "off") : null;
  const extent = transform ? getFirstLocalChild(transform, "ext") : null;
  return offset && extent
    ? {
        cx: readEmuValue(extent, "cx"),
        cy: readEmuValue(extent, "cy"),
        x: readEmuValue(offset, "x"),
        y: readEmuValue(offset, "y"),
      }
    : anchorRect;
}

function collectShapeNodes(
  parent: Element,
  transform: ShapeRectTransform,
  output: Array<{ node: Element; transform: ShapeRectTransform }>
) {
  for (const localName of ["sp", "pic", "cxnSp"]) {
    for (const child of getLocalChildren(parent, localName)) {
      output.push({ node: child, transform });
    }
  }
  for (const child of getLocalChildren(parent, "grpSp")) {
    const nested = readGroupTransform(child);
    collectShapeNodes(child, nested ? (rect) => transform(nested(rect)) : transform, output);
  }
}

/**
 * Parses a `chartUserShapes` DrawingML part into renderable shapes.
 *
 * Anchors are converted to chart-relative fractions so an SVG renderer can
 * multiply them by its own layout. Shapes without a parsable anchor are
 * dropped rather than painted at an arbitrary position.
 */
export function parseChartUserShapes(
  xml: string,
  options: ChartUserShapeParseOptions = {}
): XlsxChartUserShape[] {
  const document = parseXml(xml);
  const root = document?.documentElement ?? null;
  if (!root) {
    return [];
  }
  const anchorNodes = Array.from(root.childNodes).filter(
    (node): node is Element => (
      isElementNode(node)
      && (node.localName === "twoCellAnchor"
        || node.localName === "oneCellAnchor"
        || node.localName === "absoluteAnchor")
    )
  );
  const pending: Array<{
    anchorRect: ShapeRectEmu;
    node: Element;
    transform: ShapeRectTransform;
  }> = [];
  let maxRight = 0;
  let maxBottom = 0;
  for (const anchorNode of anchorNodes) {
    const anchorRect = readChartUserShapeAnchorRect(anchorNode);
    if (!anchorRect || anchorRect.cx <= 0 || anchorRect.cy <= 0) {
      continue;
    }
    maxRight = Math.max(maxRight, anchorRect.x + anchorRect.cx);
    maxBottom = Math.max(maxBottom, anchorRect.y + anchorRect.cy);
    const nodes: Array<{ node: Element; transform: ShapeRectTransform }> = [];
    collectShapeNodes(anchorNode, (rect) => rect, nodes);
    for (const entry of nodes) {
      pending.push({ ...entry, anchorRect });
    }
  }
  const declaredExtent = options.extentEmu;
  const extent = declaredExtent && declaredExtent.cx > 0 && declaredExtent.cy > 0
    ? declaredExtent
    : maxRight > 0 && maxBottom > 0
      ? { cx: maxRight, cy: maxBottom }
      : DEFAULT_CHART_EXTENT_EMU;

  const shapes: XlsxChartUserShape[] = [];
  for (const entry of pending) {
    const nonVisualProperties = getFirstLocalDescendant(entry.node, "cNvPr");
    const presetGeometry = getFirstLocalDescendant(entry.node, "prstGeom");
    const shapeRect = entry.transform(readShapeNodeRect(entry.node, entry.anchorRect));
    const fill = readShapeFill(entry.node, getFirstLocalChild(entry.node, "style"), options.images, options.themePalette);
    const stroke = readShapeStroke(entry.node, options.themePalette);
    const paragraphs = parseShapeParagraphs(entry.node, options.themePalette);
    if (!fill && !stroke && paragraphs.length === 0) {
      continue;
    }
    shapes.push({
      description: nonVisualProperties?.getAttribute("descr") ?? undefined,
      fill,
      geometry: presetGeometry?.getAttribute("prst") ?? (getFirstLocalDescendant(entry.node, "custGeom") ? "custom" : "rect"),
      geometryAdjustments: readGeometryAdjustments(presetGeometry),
      name: nonVisualProperties?.getAttribute("name") ?? undefined,
      paragraphs,
      rect: toFractionRect(shapeRect, extent),
      stroke,
      textBox: parseShapeTextBox(entry.node),
    });
  }
  return shapes;
}
