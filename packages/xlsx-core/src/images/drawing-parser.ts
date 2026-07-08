import type {
  XlsxImage,
  XlsxShape,
} from "../types";
import {
  type ArchiveEntries,
  type ContentTypesState,
  type DrawingRectEmu,
  EMU_PER_PIXEL,
  type GroupTransform,
  IMAGE_REL_TYPE,
  type RelationshipRecord,
  type WorkbookImageOrigin,
  anchorFromNodeOrFallback,
  anchorToRect,
  applyGroupTransform,
  cloneBytes,
  createImageAnchorNodes,
  createImageSource,
  emuToPixels,
  getChildElements,
  getEmbeddedRelationshipId,
  getFirstChild,
  getFirstDescendant,
  getLocalElements,
  getRelationshipId,
  isElementNode,
  normalizeArchivePath,
  parseAnchor,
  parseRelationships,
  parseTransformRect,
  parseXml,
  readArchiveText,
  rectToAbsoluteAnchor,
  relsPathForDocument,
  resolveContentType,
  serializeXml,
} from "./image-parser";
import { normalizeHexColor } from "./image-parser";
import { anchorToAbsoluteRect } from "./column-width";
import {
  type ThemeState,
} from "./theme-palette";
import {
  type WorkbookSheetState,
} from "./grid-render";

// ── Hyperlink ───────────────────────────────────────────────────────────────

export function getHyperlinkTarget(
  node: Element | null,
  drawingRelationships: Map<string, RelationshipRecord>
) {
  const hyperlinkNode = node ? getFirstDescendant(node, "hlinkClick") : null;
  const hyperlinkTargetNode = hyperlinkNode ?? node;
  const hyperlinkId = hyperlinkTargetNode ? getRelationshipId(hyperlinkTargetNode) : null;
  return hyperlinkId ? drawingRelationships.get(hyperlinkId)?.target ?? undefined : undefined;
}

// ── Drawing color helpers ───────────────────────────────────────────────────

function clampChannel(value: number) {
  return Math.max(0, Math.min(255, Math.round(value)));
}

function hexToRgb(hex: string) {
  const normalized = normalizeHexColor(hex);
  return {
    b: Number.parseInt(normalized.slice(5, 7), 16),
    g: Number.parseInt(normalized.slice(3, 5), 16),
    r: Number.parseInt(normalized.slice(1, 3), 16)
  };
}

function rgbToHex(red: number, green: number, blue: number) {
  return `#${[red, green, blue]
    .map((value) => clampChannel(value).toString(16).padStart(2, "0"))
    .join("")}`;
}

type DrawingColor = {
  color: string;
  opacity: number;
};

function applyDrawingColorTransforms(
  baseColor: string,
  transformNodes: Element[]
): DrawingColor {
  let { r, g, b } = hexToRgb(baseColor);
  let opacity = 1;
  let lumMod = 1;
  let lumOff = 0;

  transformNodes.forEach((node) => {
    const rawValue = Number(node.getAttribute("val") ?? 0);
    const value = Number.isFinite(rawValue) ? rawValue / 100000 : 0;
    switch (node.localName) {
      case "alpha":
        opacity *= value;
        break;
      case "lumMod":
        lumMod *= value;
        break;
      case "lumOff":
        lumOff += value;
        break;
      case "shade":
        r *= value;
        g *= value;
        b *= value;
        break;
      case "tint":
        r = r + (255 - r) * value;
        g = g + (255 - g) * value;
        b = b + (255 - b) * value;
        break;
      default:
        break;
    }
  });

  if (lumMod !== 1 || lumOff !== 0) {
    r = r * lumMod + 255 * lumOff;
    g = g * lumMod + 255 * lumOff;
    b = b * lumMod + 255 * lumOff;
  }

  return {
    color: rgbToHex(r, g, b),
    opacity: Math.max(0, Math.min(1, opacity))
  };
}

function resolveThemeColorName(theme: ThemeState, name: string | null) {
  if (!name) {
    return null;
  }

  const aliases: Record<string, string> = {
    bg1: "bg1",
    bg2: "bg2",
    tx1: "tx1",
    tx2: "tx2"
  };
  const key = aliases[name] ?? name;
  return theme.colors.get(key) ?? null;
}

function resolveColorValue(colorNode: Element | null, theme: ThemeState): DrawingColor | null {
  if (!colorNode) {
    return null;
  }

  let baseColor: string | null = null;
  if (colorNode.localName === "srgbClr") {
    baseColor = normalizeHexColor(colorNode.getAttribute("val") ?? "");
  } else if (colorNode.localName === "schemeClr") {
    baseColor = resolveThemeColorName(theme, colorNode.getAttribute("val"));
  } else if (colorNode.localName === "scrgbClr") {
    const r = Number(colorNode.getAttribute("r") ?? 0) * 255 / 100000;
    const g = Number(colorNode.getAttribute("g") ?? 0) * 255 / 100000;
    const b = Number(colorNode.getAttribute("b") ?? 0) * 255 / 100000;
    baseColor = rgbToHex(r, g, b);
  } else if (colorNode.localName === "sysClr") {
    baseColor = normalizeHexColor(colorNode.getAttribute("lastClr") ?? "");
  }

  if (!baseColor) {
    return null;
  }

  return applyDrawingColorTransforms(baseColor, Array.from(colorNode.childNodes).filter(isElementNode));
}

export function resolveFillColor(fillParent: Element | null, theme: ThemeState): DrawingColor | null {
  if (!fillParent) {
    return null;
  }

  const solidFillNode = getFirstChild(fillParent, "solidFill");
  if (!solidFillNode) {
    return null;
  }

  return resolveColorValue(Array.from(solidFillNode.childNodes).filter(isElementNode)[0] ?? null, theme);
}

// ── Text typeface ───────────────────────────────────────────────────────────

function resolveTextTypeface(typeface: string | null, theme: ThemeState) {
  if (!typeface) {
    return undefined;
  }

  if (typeface === "+mn-lt" || typeface === "+mn-ea" || typeface === "+mn-cs") {
    return theme.minorLatinFont ?? undefined;
  }
  if (typeface === "+mj-lt" || typeface === "+mj-ea" || typeface === "+mj-cs") {
    return theme.majorLatinFont ?? undefined;
  }

  return typeface;
}

function isThemeTypeface(typeface: string | null | undefined) {
  return Boolean(typeface && typeface.startsWith("+"));
}

export function resolvePreferredTextTypeface(node: Element | null, theme: ThemeState) {
  if (!node) {
    return undefined;
  }

  const latin = getFirstChild(node, "latin")?.getAttribute("typeface") ?? null;
  const eastAsian = getFirstChild(node, "ea")?.getAttribute("typeface") ?? null;
  const complexScript = getFirstChild(node, "cs")?.getAttribute("typeface") ?? null;
  const candidates = [latin, eastAsian, complexScript];
  const explicit = candidates.find((candidate) => candidate && !isThemeTypeface(candidate));
  if (explicit) {
    return explicit;
  }

  return resolveTextTypeface(candidates.find(Boolean) ?? null, theme);
}

// ── Shape text style ────────────────────────────────────────────────────────

type ShapeTextStyle = {
  bold?: boolean;
  color?: string;
  fontFamily?: string;
  fontSizePt?: number;
  italic?: boolean;
  underline?: boolean;
};

function parseShapeTextStyle(node: Element | null, theme: ThemeState, fallbackColor?: DrawingColor | null): ShapeTextStyle {
  if (!node) {
    return {
      color: fallbackColor?.color
    };
  }

  const fillColor = resolveFillColor(node, theme) ?? fallbackColor ?? null;
  const underlineValue = node.getAttribute("u");
  return {
    bold: node.getAttribute("b") === "1" || undefined,
    color: fillColor?.color,
    fontFamily: resolvePreferredTextTypeface(node, theme),
    fontSizePt: node.getAttribute("sz") ? Number(node.getAttribute("sz")) / 100 : undefined,
    italic: node.getAttribute("i") === "1" || undefined,
    underline: underlineValue && underlineValue !== "none" ? true : undefined
  };
}

function mergeShapeTextStyles(...styles: Array<ShapeTextStyle | undefined>): ShapeTextStyle {
  return styles.reduce<ShapeTextStyle>((acc, style) => {
    if (!style) {
      return acc;
    }
    return {
      bold: style.bold ?? acc.bold,
      color: style.color ?? acc.color,
      fontFamily: style.fontFamily ?? acc.fontFamily,
      fontSizePt: style.fontSizePt ?? acc.fontSizePt,
      italic: style.italic ?? acc.italic,
      underline: style.underline ?? acc.underline
    };
  }, {});
}

// ── Paragraph alignment ─────────────────────────────────────────────────────

function mapParagraphAlign(value: string | null | undefined): XlsxShape["paragraphs"][number]["align"] | undefined {
  switch (value) {
    case "ctr":
      return "center";
    case "just":
      return "justify";
    case "r":
      return "right";
    case "l":
      return "left";
    default:
      return undefined;
  }
}

function mapVerticalAnchor(value: string | null | undefined): "bottom" | "middle" | "top" {
  switch (value) {
    case "b":
      return "bottom";
    case "ctr":
      return "middle";
    default:
      return "top";
  }
}

// ── Shape geometry ──────────────────────────────────────────────────────────

function isStrokeOnlyGeometry(geometry: string) {
  return geometry === "line" || geometry === "arc" || geometry === "leftBrace";
}

function parsePointNode(pointNode: Element | null) {
  if (!pointNode) {
    return null;
  }

  const x = Number(pointNode.getAttribute("x") ?? Number.NaN);
  const y = Number(pointNode.getAttribute("y") ?? Number.NaN);
  return Number.isFinite(x) && Number.isFinite(y) ? { x, y } : null;
}

export function parseCustomGeometryPath(shapeNode: Element) {
  const pathNode = getFirstDescendant(shapeNode, "path");
  if (!pathNode) {
    return undefined;
  }

  const width = Number(pathNode.getAttribute("w") ?? Number.NaN);
  const height = Number(pathNode.getAttribute("h") ?? Number.NaN);
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    return undefined;
  }

  const commands: string[] = [];
  let lastPoint: { x: number; y: number } | null = null;
  for (const child of Array.from(pathNode.childNodes).filter(isElementNode)) {
    if (child.localName === "moveTo") {
      const point = parsePointNode(getFirstChild(child, "pt"));
      if (point) {
        commands.push(`M ${point.x} ${point.y}`);
        lastPoint = point;
      }
      continue;
    }

    if (child.localName === "lnTo") {
      const point = parsePointNode(getFirstChild(child, "pt"));
      if (point && (!lastPoint || point.x !== lastPoint.x || point.y !== lastPoint.y)) {
        commands.push(`L ${point.x} ${point.y}`);
        lastPoint = point;
      }
      continue;
    }

    if (child.localName === "cubicBezTo") {
      const points = getChildElements(child, "pt").map(parsePointNode).filter((point): point is { x: number; y: number } => point !== null);
      if (points.length === 3) {
        commands.push(`C ${points[0].x} ${points[0].y} ${points[1].x} ${points[1].y} ${points[2].x} ${points[2].y}`);
        lastPoint = points[2];
      }
      continue;
    }

    if (child.localName === "close") {
      commands.push("Z");
      lastPoint = null;
    }
  }

  if (commands.length === 0) {
    return undefined;
  }

  return {
    path: commands.join(" "),
    viewBox: {
      height,
      width
    }
  };
}

export function parseShapeGeometryAdjustments(shapeNode: Element) {
  const presetGeometry = getFirstDescendant(shapeNode, "prstGeom");
  if (!presetGeometry) {
    return undefined;
  }

  const adjustments: Record<string, number> = {};
  getChildElements(getFirstChild(presetGeometry, "avLst"), "gd").forEach((adjustmentNode) => {
    const name = adjustmentNode.getAttribute("name");
    const formula = adjustmentNode.getAttribute("fmla") ?? "";
    const match = formula.match(/^val\s+(-?\d+(?:\.\d+)?)$/);
    if (!name || !match) {
      return;
    }

    const value = Number.parseFloat(match[1] ?? "");
    if (Number.isFinite(value)) {
      adjustments[name] = value;
    }
  });

  return Object.keys(adjustments).length > 0 ? adjustments : undefined;
}

// ── Shape parsing ───────────────────────────────────────────────────────────

function parseShapeStroke(node: Element, styleNode: Element | null, theme: ThemeState): XlsxShape["stroke"] | undefined {
  const lineNode = getFirstDescendant(node, "ln");
  const lineRefNode = styleNode ? getFirstChild(styleNode, "lnRef") : null;
  if (!lineNode && !lineRefNode) {
    return undefined;
  }

  if (lineNode && getFirstChild(lineNode, "noFill")) {
    return { none: true };
  }

  const color = resolveFillColor(lineNode, theme)
    ?? resolveColorValue(Array.from(lineRefNode?.childNodes ?? []).filter(isElementNode)[0] ?? null, theme);
  const widthEmu = Number(lineNode?.getAttribute("w") ?? 0);
  return {
    color: color?.color,
    dash: getFirstChild(lineNode ?? node, "prstDash")?.getAttribute("val") ?? undefined,
    headEndType: getFirstChild(lineNode ?? node, "headEnd")?.getAttribute("type") ?? undefined,
    none: false,
    opacity: color?.opacity,
    tailEndType: getFirstChild(lineNode ?? node, "tailEnd")?.getAttribute("type") ?? undefined,
    widthPx: widthEmu > 0 ? emuToPixels(widthEmu) : undefined
  };
}

function parseShapeFill(
  node: Element,
  styleNode: Element | null,
  theme: ThemeState,
  geometry: string
): XlsxShape["fill"] | undefined {
  const shapePropsNode = getFirstChild(node, "spPr") ?? node;
  const noFillNode = getFirstChild(shapePropsNode, "noFill");
  if (noFillNode) {
    return { none: true };
  }

  if (isStrokeOnlyGeometry(geometry)) {
    return { none: true };
  }

  const fillRefNode = styleNode ? getFirstChild(styleNode, "fillRef") : null;
  const fillColor = resolveFillColor(shapePropsNode, theme)
    ?? resolveColorValue(Array.from(fillRefNode?.childNodes ?? []).filter(isElementNode)[0] ?? null, theme);
  if (!fillColor) {
    return undefined;
  }

  return {
    color: fillColor.color,
    none: false,
    opacity: fillColor.opacity
  };
}

function parseTextBox(shapeNode: Element): XlsxShape["textBox"] | undefined {
  const txBodyNode = getFirstChild(shapeNode, "txBody");
  if (!txBodyNode) {
    return undefined;
  }

  const bodyProps = getFirstChild(txBodyNode, "bodyPr");
  const leftInset = emuToPixels(Number(bodyProps?.getAttribute("lIns") ?? 91440));
  const rightInset = emuToPixels(Number(bodyProps?.getAttribute("rIns") ?? 91440));
  const topInset = emuToPixels(Number(bodyProps?.getAttribute("tIns") ?? 45720));
  const bottomInset = emuToPixels(Number(bodyProps?.getAttribute("bIns") ?? 45720));

  return {
    horizontalAlign: bodyProps?.getAttribute("anchorCtr") === "1" ? "center" : "left",
    insetPx: {
      bottom: bottomInset,
      left: leftInset,
      right: rightInset,
      top: topInset
    },
    verticalAlign: mapVerticalAnchor(bodyProps?.getAttribute("anchor"))
  };
}

function parseShapeParagraphs(
  shapeNode: Element,
  styleNode: Element | null,
  theme: ThemeState
): XlsxShape["paragraphs"] {
  const txBodyNode = getFirstChild(shapeNode, "txBody");
  if (!txBodyNode) {
    return [];
  }

  const defaultFontRef = styleNode ? getFirstChild(styleNode, "fontRef") : null;
  const defaultFontColor = resolveColorValue(Array.from(defaultFontRef?.childNodes ?? []).filter(isElementNode)[0] ?? null, theme);
  const listStyleNode = getFirstChild(txBodyNode, "lstStyle");
  const paragraphs: XlsxShape["paragraphs"] = [];

  getChildElements(txBodyNode, "p").forEach((paragraphNode) => {
    const paragraphProps = getFirstChild(paragraphNode, "pPr");
    const paragraphLevel = Number(paragraphProps?.getAttribute("lvl") ?? 0);
    const listLevelProps = getFirstChild(listStyleNode ?? txBodyNode, `lvl${paragraphLevel + 1}pPr`);
    const inheritedStyle = mergeShapeTextStyles(
      parseShapeTextStyle(getFirstChild(listLevelProps ?? txBodyNode, "defRPr"), theme, defaultFontColor),
      parseShapeTextStyle(getFirstChild(paragraphProps ?? paragraphNode, "defRPr"), theme, defaultFontColor)
    );
    const runs: XlsxShape["paragraphs"][number]["runs"] = [];
    let sawRenderableChild = false;

    Array.from(paragraphNode.childNodes).filter(isElementNode).forEach((child) => {
      if (child.localName === "br") {
        sawRenderableChild = true;
        runs.push({ text: "\n" });
        return;
      }
      if (child.localName !== "r") {
        return;
      }

      sawRenderableChild = true;
      const text = getFirstChild(child, "t")?.textContent ?? "";
      const runProps = getFirstChild(child, "rPr");
      const runStyle = mergeShapeTextStyles(
        inheritedStyle,
        parseShapeTextStyle(runProps, theme, defaultFontColor)
      );

      runs.push({
        bold: runStyle.bold,
        color: runStyle.color,
        fontFamily: runStyle.fontFamily,
        fontSizePt: runStyle.fontSizePt,
        italic: runStyle.italic,
        text,
        underline: runStyle.underline
      });
    });

    if (runs.length === 0) {
      if (!sawRenderableChild && !getFirstChild(paragraphNode, "endParaRPr")) {
        return;
      }

      runs.push({
        bold: inheritedStyle.bold,
        color: inheritedStyle.color,
        fontFamily: inheritedStyle.fontFamily,
        fontSizePt: inheritedStyle.fontSizePt,
        italic: inheritedStyle.italic,
        text: " ",
        underline: inheritedStyle.underline
      });
    }

    paragraphs.push({
      align: mapParagraphAlign(paragraphProps?.getAttribute("algn")),
      runs
    });
  });

  return paragraphs;
}

function parseShapeNode(
  shapeNode: Element,
  fallbackAnchor: XlsxImage["anchor"],
  drawingRelationships: Map<string, RelationshipRecord>,
  theme: ThemeState,
  workbookSheetIndex: number,
  shapeId: string,
  zIndex: number,
  parentGroup: GroupTransform | null
) {
  const nonVisualProps = getFirstDescendant(shapeNode, "cNvPr");
  const styleNode = getFirstChild(shapeNode, "style");
  const transform = anchorFromNodeOrFallback(shapeNode, fallbackAnchor, parentGroup);
  const geometry = getFirstDescendant(shapeNode, "prstGeom")?.getAttribute("prst")
    ?? (getFirstDescendant(shapeNode, "custGeom") ? "custom" : "rect");
  const customPath = geometry === "custom" ? parseCustomGeometryPath(shapeNode) : undefined;
  const geometryAdjustments = parseShapeGeometryAdjustments(shapeNode);
  return {
    anchor: transform.anchor,
    description: nonVisualProps?.getAttribute("descr") ?? undefined,
    fill: parseShapeFill(shapeNode, styleNode, theme, geometry),
    flipH: transform.flipH,
    flipV: transform.flipV,
    geometry,
    geometryAdjustments,
    hidden: nonVisualProps?.getAttribute("hidden") === "1" || undefined,
    hyperlink: getHyperlinkTarget(nonVisualProps ?? shapeNode, drawingRelationships),
    id: shapeId,
    name: nonVisualProps?.getAttribute("name") ?? undefined,
    paragraphs: parseShapeParagraphs(shapeNode, styleNode, theme),
    rotationDeg: transform.rotationDeg,
    scaleX: transform.scaleX,
    scaleY: transform.scaleY,
    sheetIndex: workbookSheetIndex,
    svgPath: customPath?.path,
    svgViewBox: customPath?.viewBox,
    stroke: parseShapeStroke(shapeNode, styleNode, theme),
    textBox: parseTextBox(shapeNode),
    workbookSheetIndex,
    zIndex
  } satisfies XlsxShape;
}

// ── Picture parsing ─────────────────────────────────────────────────────────

function parsePictureNode(
  pictureNode: Element,
  fallbackAnchor: XlsxImage["anchor"],
  drawingRelationships: Map<string, RelationshipRecord>,
  archive: ArchiveEntries,
  contentTypes: ContentTypesState,
  objectUrls: string[],
  workbookSheetIndex: number,
  imageId: string,
  zIndex: number,
  parentGroup: GroupTransform | null
) {
  const blipNode = getFirstDescendant(pictureNode, "blip");
  const svgBlipNode = blipNode ? getFirstDescendant(blipNode, "svgBlip") : null;
  const embedId = (svgBlipNode ? getEmbeddedRelationshipId(svgBlipNode) : null) ?? (blipNode ? getEmbeddedRelationshipId(blipNode) : null);
  if (!embedId) {
    return null;
  }

  const mediaRelationship = drawingRelationships.get(embedId);
  if (!mediaRelationship || mediaRelationship.type !== IMAGE_REL_TYPE) {
    return null;
  }

  const mediaBytes = archive[mediaRelationship.target];
  if (!mediaBytes) {
    return null;
  }

  const nonVisualProps = getFirstDescendant(pictureNode, "cNvPr");
  const transform = anchorFromNodeOrFallback(pictureNode, fallbackAnchor, parentGroup);
  return {
    image: {
      anchor: transform.anchor,
      description: nonVisualProps?.getAttribute("descr") ?? undefined,
      hyperlink: getHyperlinkTarget(nonVisualProps ?? pictureNode, drawingRelationships),
      id: imageId,
      mediaPath: mediaRelationship.target,
      mimeType: resolveContentType(contentTypes, mediaRelationship.target),
      name: nonVisualProps?.getAttribute("name") ?? undefined,
      sheetIndex: workbookSheetIndex,
      src: createImageSource(mediaBytes, resolveContentType(contentTypes, mediaRelationship.target), objectUrls),
      workbookSheetIndex,
      zIndex
    } satisfies XlsxImage,
    mediaPath: mediaRelationship.target
  };
}

// ── Group transform ─────────────────────────────────────────────────────────

export function parseGroupTransform(
  groupNode: Element,
  parentGroup: GroupTransform | null,
  fallbackAnchor: XlsxImage["anchor"],
  sheetState: WorkbookSheetState | null
): GroupTransform {
  const xfrmNode = getFirstDescendant(getFirstChild(groupNode, "grpSpPr") ?? groupNode, "xfrm");
  const anchorRect = anchorToAbsoluteRect(fallbackAnchor, sheetState);
  const rect = parseTransformRect(xfrmNode) ?? anchorRect;
  const rootRectMatchesAnchorOrigin = Math.abs(rect.x - anchorRect.x) <= EMU_PER_PIXEL
    && Math.abs(rect.y - anchorRect.y) <= EMU_PER_PIXEL;
  const chOffNode = getFirstChild(xfrmNode ?? groupNode, "chOff");
  const chExtNode = getFirstChild(xfrmNode ?? groupNode, "chExt");
  const rootScaleX = rect.cx !== 0 ? anchorRect.cx / rect.cx : 1;
  const rootScaleY = rect.cy !== 0 ? anchorRect.cy / rect.cy : 1;
  const useRectFrameForRoot = !parentGroup && (
    rootScaleX < 0.85
    || rootScaleX > 1.15
    || rootScaleY < 0.85
    || rootScaleY > 1.15
  );
  const absoluteRect = parentGroup
    ? applyGroupTransform(rect, parentGroup)
    : rootRectMatchesAnchorOrigin
      ? rect
      : anchorRect;
  const childRectX = parentGroup
    ? Number(chOffNode?.getAttribute("x") ?? 0)
    : useRectFrameForRoot
      ? rect.x
      : Number(chOffNode?.getAttribute("x") ?? 0);
  const childRectY = parentGroup
    ? Number(chOffNode?.getAttribute("y") ?? 0)
    : useRectFrameForRoot
      ? rect.y
      : Number(chOffNode?.getAttribute("y") ?? 0);
  const childRectCx = parentGroup
    ? Number(chExtNode?.getAttribute("cx") ?? rect.cx)
    : useRectFrameForRoot
      ? rect.cx
      : Number(chExtNode?.getAttribute("cx") ?? rect.cx);
  const childRectCy = parentGroup
    ? Number(chExtNode?.getAttribute("cy") ?? rect.cy)
    : useRectFrameForRoot
      ? rect.cy
      : Number(chExtNode?.getAttribute("cy") ?? rect.cy);
  return {
    chCx: childRectCx,
    chCy: childRectCy,
    chX: childRectX,
    chY: childRectY,
    cx: absoluteRect.cx,
    cy: absoluteRect.cy,
    scaleX: childRectCx !== 0 ? absoluteRect.cx / childRectCx : (parentGroup?.scaleX ?? 1),
    scaleY: childRectCy !== 0 ? absoluteRect.cy / childRectCy : (parentGroup?.scaleY ?? 1),
    x: absoluteRect.x,
    y: absoluteRect.y
  };
}

// ── Anchor contents parsing ─────────────────────────────────────────────────

export function parseAnchorContents(
  anchorNode: Element,
  fallbackAnchor: XlsxImage["anchor"],
  drawingRelationships: Map<string, RelationshipRecord>,
  archive: ArchiveEntries,
  contentTypes: ContentTypesState,
  objectUrls: string[],
  workbookSheetIndex: number,
  theme: ThemeState,
  ids: { image: number; shape: number; z: number },
  imageOriginsById: Map<string, WorkbookImageOrigin>,
  anchorIndex: number,
  sheetState: WorkbookSheetState | null,
  parentGroup: GroupTransform | null = null
) {
  const images: XlsxImage[] = [];
  const shapes: XlsxShape[] = [];
  const mediaPaths = new Set<string>();

  Array.from(anchorNode.childNodes).filter(isElementNode).forEach((child) => {
    if (child.localName === "pic") {
      const imageId = `sheet-${workbookSheetIndex}-${ids.image}`;
      const parsed = parsePictureNode(
        child,
        fallbackAnchor,
        drawingRelationships,
        archive,
        contentTypes,
        objectUrls,
        workbookSheetIndex,
        imageId,
        ids.z++,
        parentGroup
      );
      ids.image += 1;
      if (parsed) {
        images.push(parsed.image);
        mediaPaths.add(parsed.mediaPath);
        imageOriginsById.set(imageId, {
          anchorIndex,
          workbookSheetIndex
        });
      }
      return;
    }

    if (child.localName === "sp") {
      shapes.push(parseShapeNode(
        child,
        fallbackAnchor,
        drawingRelationships,
        theme,
        workbookSheetIndex,
        `shape-${workbookSheetIndex}-${ids.shape++}`,
        ids.z++,
        parentGroup
      ));
      return;
    }

    if (child.localName === "cxnSp") {
      shapes.push(parseShapeNode(
        child,
        fallbackAnchor,
        drawingRelationships,
        theme,
        workbookSheetIndex,
        `shape-${workbookSheetIndex}-${ids.shape++}`,
        ids.z++,
        parentGroup
      ));
      return;
    }

    if (child.localName !== "grpSp") {
      return;
    }

    const nextGroup = parseGroupTransform(
      child,
      parentGroup,
      fallbackAnchor,
      sheetState
    );
    const groupFallbackAnchor = rectToAbsoluteAnchor({
      cx: nextGroup.cx,
      cy: nextGroup.cy,
      x: nextGroup.x,
      y: nextGroup.y
    });
    const parsedGroup = parseAnchorContents(
      child,
      groupFallbackAnchor,
      drawingRelationships,
      archive,
      contentTypes,
      objectUrls,
      workbookSheetIndex,
      theme,
      ids,
      imageOriginsById,
      anchorIndex,
      sheetState,
      nextGroup
    );
    parsedGroup.images.forEach((image) => images.push(image));
    parsedGroup.shapes.forEach((shape) => shapes.push(shape));
    parsedGroup.mediaPaths.forEach((path) => mediaPaths.add(path));
  });

  return {
    images,
    mediaPaths,
    shapes
  };
}

// ── Drawing objects parsing ─────────────────────────────────────────────────

export function parseDrawingObjects(
  archive: ArchiveEntries,
  contentTypes: ContentTypesState,
  drawingPath: string,
  objectUrls: string[],
  workbookSheetIndex: number,
  zIndexBase: number,
  theme: ThemeState,
  sheetState: WorkbookSheetState | null,
  imageOriginsById: Map<string, WorkbookImageOrigin>
) {
  const drawingXml = readArchiveText(archive, drawingPath);
  if (!drawingXml) {
    return {
      images: [] as XlsxImage[],
      mediaPaths: [] as string[],
      shapes: [] as XlsxShape[]
    };
  }

  const drawingDocument = parseXml(drawingXml);
  if (!drawingDocument) {
    return {
      images: [] as XlsxImage[],
      mediaPaths: [] as string[],
      shapes: [] as XlsxShape[]
    };
  }

  const drawingRelationships = parseRelationships(archive, relsPathForDocument(drawingPath), drawingPath);
  const images: XlsxImage[] = [];
  const shapes: XlsxShape[] = [];
  const mediaPaths = new Set<string>();
  const anchorNodes = createImageAnchorNodes(drawingDocument);
  const ids = {
    image: zIndexBase,
    shape: zIndexBase,
    z: zIndexBase
  };

  anchorNodes.forEach((anchorNode, anchorIndex) => {
    const anchor = parseAnchor(anchorNode);
    if (!anchor) {
      return;
    }

    const parsed = parseAnchorContents(
      anchorNode,
      anchor,
      drawingRelationships,
      archive,
      contentTypes,
      objectUrls,
      workbookSheetIndex,
      theme,
      ids,
      imageOriginsById,
      anchorIndex,
      sheetState
    );
    parsed.images.forEach((image) => images.push(image));
    parsed.shapes.forEach((shape) => shapes.push(shape));
    parsed.mediaPaths.forEach((path) => mediaPaths.add(path));
  });

  return {
    images,
    mediaPaths: [...mediaPaths],
    shapes
  };
}
// Re-export form control parsing from separate module
export {
  normalizeControlLabel,
  flattenShapeText,
  parseSheetFormControls,
  enrichFormControlsWithHiddenShapes,
} from "./form-control-parser";
