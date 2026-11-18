// Drawing objects: parse the non-picture content of a paragraph.
//
// The WASM model builder turns text boxes into synthetic images and drops
// everything else (shapes, WordArt, groups, SmartArt, charts, OLE objects, 3D
// models, equations). This module scans a paragraph's own markup and reports the
// objects it carries, in document order, with the text offset they sit at, so
// the importer can attach `DrawingRunNode` / `MathRunNode` children and the
// serializer can write the stored markup back after an edit rebuilt the
// paragraph.

import type {
  ChartData,
  DrawingAnchor,
  DrawingObjectKind,
  DrawingObjectShapeStyle,
  ParagraphAlignment,
  TextBoxContent,
  TextBoxParagraph,
  TextBoxRunSegment,
  TextStyle,
} from "./types";
import { linearizeOmml } from "./omml";
import { findXmlElements, parseXmlElements } from "./xml-elements";

const EMU_PER_PIXEL = 9525;
const OBJECT_ROOT_NAMES = new Set([
  "w:drawing",
  "w:object",
  "w:pict",
  "m:oMath",
  "m:oMathPara",
  "mc:AlternateContent",
]);
/** Cheap pre-check so ordinary paragraphs cost one regex test. */
const OBJECT_MARKER_PATTERN =
  /<(?:w:drawing|w:object|w:pict|m:oMath|m:oMathPara|mc:AlternateContent)\b/i;
const TAG_PATTERN =
  /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[([\s\S]*?)\]\]>|<(\/?)([A-Za-z_][\w.:-]*)\b([^>]*?)(\/?)>/g;

/** One object found in a paragraph, positioned by its preceding text length. */
export interface ParagraphDrawingObject {
  kind: DrawingObjectKind | "picture" | "math";
  sourceXml: string;
  /** Characters of paragraph text before this object. */
  textOffset: number;
}

export function decodeXmlEntities(value: string): string {
  return value
    .replace(/&#x([0-9a-fA-F]+);/g, (_match, hex: string) =>
      String.fromCodePoint(Number.parseInt(hex, 16)),
    )
    .replace(/&#(\d+);/g, (_match, decimal: string) =>
      String.fromCodePoint(Number.parseInt(decimal, 10)),
    )
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function emuToPixels(value: string | undefined): number | undefined {
  if (!value) return undefined;
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed <= 0) return undefined;
  return Number((parsed / EMU_PER_PIXEL).toFixed(3));
}

function readAttribute(attributes: string | undefined, name: string): string | undefined {
  if (!attributes) return undefined;
  const match = attributes.match(
    new RegExp(`(?:^|\\s)${name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`),
  );
  return match?.[1] ?? match?.[2];
}

function onOff(value: string | undefined): boolean | undefined {
  if (value === undefined) return undefined;
  const normalized = value.trim().toLowerCase();
  if (normalized === "1" || normalized === "true" || normalized === "on") return true;
  if (normalized === "0" || normalized === "false" || normalized === "off") return false;
  return undefined;
}

/** `<wp:extent>` / `<a:ext>` of a drawing, in CSS pixels. */
export function drawingExtentPx(
  sourceXml: string,
): { widthPx?: number; heightPx?: number } {
  const extent =
    /<wp:extent\b[^>]*>/i.exec(sourceXml)?.[0] ??
    /<a:ext\b[^>]*\bcx="[^"]*"[^>]*>/i.exec(sourceXml)?.[0];
  if (!extent) return {};
  const widthPx = emuToPixels(readAttribute(extent, "cx"));
  const heightPx = emuToPixels(readAttribute(extent, "cy"));
  return {
    ...(widthPx ? { widthPx } : {}),
    ...(heightPx ? { heightPx } : {}),
  };
}

/** The `wp:anchor` of a drawing as the model's floating anchor. */
export function parseDrawingAnchor(sourceXml: string): DrawingAnchor | undefined {
  const anchorMatch = /<(?:[A-Za-z_][\w.-]*:)?anchor\b([^>]*)>/i.exec(sourceXml);
  if (!anchorMatch) return undefined;
  const attributes = anchorMatch[1] ?? "";
  const anchor: DrawingAnchor = {};

  const distL = emuToPixels(readAttribute(attributes, "distL"));
  const distR = emuToPixels(readAttribute(attributes, "distR"));
  const distT = emuToPixels(readAttribute(attributes, "distT"));
  const distB = emuToPixels(readAttribute(attributes, "distB"));
  if (distL !== undefined) anchor.distLPx = distL;
  if (distR !== undefined) anchor.distRPx = distR;
  if (distT !== undefined) anchor.distTPx = distT;
  if (distB !== undefined) anchor.distBPx = distB;

  const behindDocument = onOff(readAttribute(attributes, "behindDoc"));
  if (behindDocument !== undefined) anchor.behindDocument = behindDocument;
  const layoutInCell = onOff(readAttribute(attributes, "layoutInCell"));
  if (layoutInCell !== undefined) anchor.layoutInCell = layoutInCell;
  const allowOverlap = onOff(readAttribute(attributes, "allowOverlap"));
  if (allowOverlap !== undefined) anchor.allowOverlap = allowOverlap;
  const lockedAnchor = onOff(readAttribute(attributes, "locked"));
  if (lockedAnchor !== undefined) anchor.lockedAnchor = lockedAnchor;
  const simplePos = onOff(readAttribute(attributes, "simplePos"));
  if (simplePos !== undefined) anchor.simplePos = simplePos;

  const relativeHeight = Number(readAttribute(attributes, "relativeHeight"));
  if (Number.isFinite(relativeHeight)) anchor.zIndex = relativeHeight;

  const horizontal = /<wp:positionH\b([^>]*)>([\s\S]*?)<\/wp:positionH>/i.exec(sourceXml);
  if (horizontal) {
    const relativeTo = readAttribute(horizontal[1], "relativeFrom");
    if (relativeTo) anchor.horizontalRelativeTo = relativeTo;
    const offset = /<wp:posOffset>(-?\d+)<\/wp:posOffset>/i.exec(horizontal[2]);
    const align = /<wp:align>([^<]*)<\/wp:align>/i.exec(horizontal[2]);
    if (offset) {
      const pixels = emuToPixels(offset[1]);
      if (pixels !== undefined) anchor.xPx = pixels;
    } else if (align) {
      const value = align[1].trim().toLowerCase();
      if (["left", "center", "right", "inside", "outside"].includes(value)) {
        anchor.horizontalAlign = value as NonNullable<DrawingAnchor["horizontalAlign"]>;
      }
    }
  }

  const vertical = /<wp:positionV\b([^>]*)>([\s\S]*?)<\/wp:positionV>/i.exec(sourceXml);
  if (vertical) {
    const relativeTo = readAttribute(vertical[1], "relativeFrom");
    if (relativeTo) anchor.verticalRelativeTo = relativeTo;
    const offset = /<wp:posOffset>(-?\d+)<\/wp:posOffset>/i.exec(vertical[2]);
    const align = /<wp:align>([^<]*)<\/wp:align>/i.exec(vertical[2]);
    if (offset) {
      const pixels = emuToPixels(offset[1]);
      if (pixels !== undefined) anchor.yPx = pixels;
    } else if (align) {
      const value = align[1].trim().toLowerCase();
      if (["top", "center", "bottom", "inside", "outside"].includes(value)) {
        anchor.verticalAlign = value as NonNullable<DrawingAnchor["verticalAlign"]>;
      }
    }
  }

  const wrap =
    /<wp:wrapNone\b/i.test(sourceXml)
      ? { wrapType: "none" as const }
      : /<wp:wrapSquare\b([^>]*)>/i.exec(sourceXml)
        ? {
            wrapType: "square" as const,
            wrapText: parseWrapText(/<wp:wrapSquare\b([^>]*)>/i.exec(sourceXml)?.[1]),
          }
        : /<wp:wrapTight\b([^>]*)>/i.test(sourceXml)
          ? {
              wrapType: "tight" as const,
              wrapText: parseWrapText(/<wp:wrapTight\b([^>]*)>/i.exec(sourceXml)?.[1]),
            }
          : /<wp:wrapThrough\b([^>]*)>/i.test(sourceXml)
            ? {
                wrapType: "through" as const,
                wrapText: parseWrapText(/<wp:wrapThrough\b([^>]*)>/i.exec(sourceXml)?.[1]),
              }
            : /<wp:wrapTopAndBottom\b/i.test(sourceXml)
              ? { wrapType: "topAndBottom" as const }
              : undefined;
  if (wrap) {
    anchor.wrapType = wrap.wrapType;
    if (wrap.wrapText) anchor.wrapText = wrap.wrapText;
  }

  return Object.keys(anchor).length > 0 ? anchor : undefined;
}

function parseWrapText(attributes: string | undefined): DrawingAnchor["wrapText"] {
  const value = readAttribute(attributes, "wrapText")?.toLowerCase();
  if (value === "bothsides" || value === "left" || value === "right" || value === "largest") {
    return value.charAt(0).toUpperCase() + value.slice(1) as NonNullable<
      DrawingAnchor["wrapText"]
    >;
  }
  return undefined;
}

/** Fill/stroke of a shape, used to paint its placeholder. */
export function parseDrawingShapeStyle(
  sourceXml: string,
): DrawingObjectShapeStyle | undefined {
  const shapeProperties = /<wps:spPr\b[\s\S]*?<\/wps:spPr>|<a:spPr\b[\s\S]*?<\/a:spPr>/i.exec(
    sourceXml,
  )?.[0];
  const style: DrawingObjectShapeStyle = {};

  const preset = /<a:prstGeom\b[^>]*\bprst="([^"]+)"/i.exec(
    shapeProperties ?? sourceXml,
  )?.[1];
  if (preset) style.preset = preset;

  const fill = shapeProperties?.match(
    /<a:solidFill>\s*<a:(?:srgbClr|schemeClr)\b[^>]*val="([0-9A-Fa-f]{6})"/i,
  )?.[1];
  if (fill && !/<a:noFill\b/i.test(shapeProperties ?? "")) {
    style.fillColor = `#${fill.toUpperCase()}`;
  }

  const line = /<a:ln\b([^>]*)>([\s\S]*?)<\/a:ln>/i.exec(shapeProperties ?? "");
  if (line && !/<a:noFill\b/i.test(line[2] ?? "")) {
    const lineColor = /<a:(?:srgbClr|schemeClr)\b[^>]*val="([0-9A-Fa-f]{6})"/i.exec(
      line[2] ?? "",
    )?.[1];
    if (lineColor) style.strokeColor = `#${lineColor.toUpperCase()}`;
    const widthEmu = Number(readAttribute(line[1], "w"));
    if (Number.isFinite(widthEmu) && widthEmu > 0) {
      style.strokeWidthPx = Math.max(1, Math.round(widthEmu / EMU_PER_PIXEL));
    }
  }

  return Object.keys(style).length > 0 ? style : undefined;
}

/** `wp:docPr` / VML `v:shape` identity and description of a drawing. */
export function drawingObjectIdentity(sourceXml: string): {
  name?: string;
  alt?: string;
} {
  const docPr = /<wp:docPr\b[^>]*>/i.exec(sourceXml)?.[0];
  const vml = /<v:shape\b[^>]*>/i.exec(sourceXml)?.[0];
  const name =
    readAttribute(docPr, "name") ??
    readAttribute(vml, "id") ??
    /<v:shape\b[^>]*\balt="([^"]*)"/i.exec(sourceXml)?.[1];
  const alt = readAttribute(docPr, "descr") ?? readAttribute(vml, "alt");
  return {
    ...(name ? { name } : {}),
    ...(alt ? { alt } : {}),
  };
}

/**
 * `a:graphicData` URI of a DrawingML object. Used to tell pictures, charts,
 * diagrams, groups and shapes apart.
 */
export function drawingGraphicUri(sourceXml: string): string | undefined {
  return /<a:graphicData\b[^>]*\buri="([^"]*)"/i.exec(sourceXml)?.[1];
}

function drawingContainsTextBox(sourceXml: string): boolean {
  return /<w:txbxContent\b/i.test(sourceXml) || /<v:textbox\b/i.test(sourceXml);
}

/** Classifies the content of one `w:drawing` / `w:pict` / `w:object` block. */
export function classifyDrawingXml(
  sourceXml: string,
): DrawingObjectKind | "picture" | undefined {
  if (/<w:object\b/i.test(sourceXml) && /<o:OLEObject\b/i.test(sourceXml)) {
    return "oleObject";
  }
  if (/<o:OLEObject\b/i.test(sourceXml)) {
    return "oleObject";
  }

  const uri = drawingGraphicUri(sourceXml) ?? "";
  if (/\/chart$/i.test(uri) || /<c:chart\b/i.test(sourceXml)) {
    return "chart";
  }
  if (/\/diagram$/i.test(uri) || /<dgm:relIds\b/i.test(sourceXml)) {
    return "smartArt";
  }
  if (/<a:model3d\b/i.test(sourceXml) || /model3d/i.test(uri)) {
    return "model3d";
  }
  if (/\/wordprocessingGroup$/i.test(uri) || /<wpg:wgp\b/i.test(sourceXml)) {
    return "group";
  }
  if (/\/wordprocessingShape$/i.test(uri) || /<wps:wsp\b/i.test(sourceXml)) {
    if (drawingContainsTextBox(sourceXml)) return "textBox";
    if (/<a:prstTxWarp\b/i.test(sourceXml) || /\bprst="text[A-Za-z]*"/i.test(sourceXml)) {
      return "wordArt";
    }
    return "shape";
  }
  if (/\/picture$/i.test(uri)) {
    return "picture";
  }
  if (/<w:txbxContent\b/i.test(sourceXml)) {
    return "textBox";
  }
  if (/<pic:pic\b/i.test(sourceXml)) {
    return "picture";
  }
  if (/<v:textbox\b/i.test(sourceXml)) {
    return "textBox";
  }
  if (/<v:shape\b/i.test(sourceXml) || /<v:rect\b/i.test(sourceXml)) {
    return /<v:imagedata\b/i.test(sourceXml) ? "picture" : "shape";
  }
  if (/<w:pict\b/i.test(sourceXml)) {
    // A `w:pict` without any recognized content is still a shape Word paints.
    return /<v:imagedata\b/i.test(sourceXml) ? "picture" : "shape";
  }
  return undefined;
}

// ── Text box content ────────────────────────────────────────────────

function parseRunStyle(runPropertiesXml: string): TextStyle | undefined {
  if (!runPropertiesXml) return undefined;
  const style: TextStyle = {};
  const onOffTag = (tag: string): boolean | undefined => {
    const match = new RegExp(`<w:${tag}\\b([^>]*)/?>`).exec(runPropertiesXml);
    if (!match) return undefined;
    const value = readAttribute(match[1], "w:val");
    return value === undefined ? true : onOff(value) ?? true;
  };

  const bold = onOffTag("b");
  if (bold !== undefined) style.bold = bold;
  const italic = onOffTag("i");
  if (italic !== undefined) style.italic = italic;
  const underline = onOffTag("u");
  if (underline !== undefined) style.underline = underline;
  const strike = onOffTag("strike");
  if (strike !== undefined) style.strike = strike;

  const size = /<w:sz\b[^>]*w:val="(\d+)"/.exec(runPropertiesXml)?.[1];
  if (size) style.fontSizePt = Number(size) / 2;
  const color = /<w:color\b[^>]*w:val="([0-9A-Fa-f]{6})"/.exec(runPropertiesXml)?.[1];
  if (color) style.color = `#${color}`;
  const font = /<w:rFonts\b[^>]*w:ascii="([^"]+)"/.exec(runPropertiesXml)?.[1]
    ?? /<w:rFonts\b[^>]*w:eastAsia="([^"]+)"/.exec(runPropertiesXml)?.[1];
  if (font) style.fontFamily = font;

  return Object.keys(style).length > 0 ? style : undefined;
}

function parseTextBoxParagraph(paragraphXml: string): TextBoxParagraph | undefined {
  const paragraphProperties =
    /<w:pPr\b[\s\S]*?<\/w:pPr>|<w:pPr\b[^>]*\/>/i.exec(paragraphXml)?.[0] ?? "";
  const runProperties =
    /<w:rPr\b[\s\S]*?<\/w:rPr>|<w:rPr\b[^>]*\/>/i.exec(paragraphProperties)?.[0] ?? "";
  const paragraphStyle = parseRunStyle(runProperties);

  const alignment = /<w:jc\b[^>]*w:val="([^"]+)"/.exec(paragraphProperties)?.[1];
  const spacing = /<w:spacing\b[^>]*\/?>/.exec(paragraphProperties)?.[0] ?? "";
  const line = Number(readAttribute(spacing, "w:line"));
  const lineRule = readAttribute(spacing, "w:lineRule")?.toLowerCase();
  const fontSizePt = paragraphStyle?.fontSizePt ?? 12;
  const lineHeightPx =
    Number.isFinite(line) && line > 0 && lineRule === "exact"
      ? Math.max(1, Math.round((line / 20) * (96 / 72)))
      : Math.max(14, Math.round((fontSizePt * 96) / 72 * 1.24));

  const segments: TextBoxRunSegment[] = [];
  for (const runMatch of paragraphXml.matchAll(/<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g)) {
    const runXml = runMatch[1] ?? "";
    const styleProperties =
      /<w:rPr\b[\s\S]*?<\/w:rPr>|<w:rPr\b[^>]*\/>/i.exec(runXml)?.[0] ?? "";
    const style = { ...(paragraphStyle ?? {}), ...(parseRunStyle(styleProperties) ?? {}) };
    const text = [...runXml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)]
      .map((match) => decodeXmlEntities(match[1] ?? ""))
      .join("");
    if (!text) continue;
    const previous = segments[segments.length - 1];
    if (previous && JSON.stringify(previous.style ?? {}) === JSON.stringify(style)) {
      previous.text += text;
      continue;
    }
    segments.push({
      text,
      ...(Object.keys(style).length > 0 ? { style } : {}),
    });
  }

  if (segments.length === 0) return undefined;
  return {
    ...(alignment && isParagraphAlignment(alignment) ? { align: alignment } : {}),
    lineHeightPx,
    segments,
  };
}

function isParagraphAlignment(value: string): value is ParagraphAlignment {
  return [
    "left",
    "center",
    "right",
    "justify",
    "both",
    "distribute",
    "thaiDistribute",
    "low",
    "start",
    "end",
  ].includes(value);
}

/** Span of one element inside a larger markup fragment. */
export interface XmlElementSpan {
  /** Index of the open tag. */
  start: number;
  /** Index just past the close tag. */
  end: number;
  /** Index just past the open tag. */
  innerStart: number;
  /** Index of the close tag. */
  innerEnd: number;
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Balanced spans of every element with exactly `name`, in document order. Used
 * for `w:txbxContent`, which appears once in the `mc:Choice` and once in the
 * `mc:Fallback` of a text box.
 */
export function findElementSpans(xml: string, name: string): XmlElementSpan[] {
  const pattern = new RegExp(
    `<(/?)${escapeRegExp(name)}\\b[^>]*?(/?)>`,
    "g",
  );
  const spans: XmlElementSpan[] = [];
  const open: Array<{ start: number; innerStart: number }> = [];
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(xml)) !== null) {
    if (match[1] === "/") {
      const current = open.pop();
      if (!current) continue;
      spans.push({
        start: current.start,
        end: match.index + match[0].length,
        innerStart: current.innerStart,
        innerEnd: match.index,
      });
      continue;
    }
    if (match[2] === "/") {
      spans.push({
        start: match.index,
        end: match.index + match[0].length,
        innerStart: match.index + match[0].length,
        innerEnd: match.index + match[0].length,
      });
      continue;
    }
    open.push({
      start: match.index,
      innerStart: match.index + match[0].length,
    });
  }
  return spans.sort((left, right) => left.start - right.start);
}

/** Balanced span of the first element with `name`. */
export function findElementSpan(
  xml: string,
  name: string,
): XmlElementSpan | undefined {
  const span = findElementSpans(xml, name)[0];
  if (!span) return undefined;
  // An outer element close would have popped an inner one first: keep the
  // outermost span that starts first.
  return span;
}

/** Parses the `w:txbxContent` of a drawing into model paragraphs. */
export function parseTextBoxContent(xml: string | undefined): TextBoxContent | undefined {
  if (!xml) return undefined;
  const span = findElementSpan(xml, "w:txbxContent");
  if (!span) return undefined;
  const paragraphs: TextBoxParagraph[] = [];
  for (const match of xml
    .slice(span.innerStart, span.innerEnd)
    .matchAll(/<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g)) {
    const paragraph = parseTextBoxParagraph(`<w:p>${match[1] ?? ""}</w:p>`);
    if (paragraph) paragraphs.push(paragraph);
  }
  if (paragraphs.length === 0) return undefined;
  return { paragraphs };
}

/** Plain text of a text box, one line per paragraph. */
export function textBoxContentText(content: TextBoxContent | undefined): string | undefined {
  if (!content || content.paragraphs.length === 0) return undefined;
  return content.paragraphs
    .map((paragraph) => paragraph.segments.map((segment) => segment.text).join(""))
    .join("\n");
}

// ── Paragraph scan ──────────────────────────────────────────────────

/**
 * Every drawing object and equation of a paragraph, in document order.
 *
 * Text inside an object (a text box's `w:txbxContent`, for instance) is not
 * counted towards `textOffset`; the offset is the paragraph-flow text position
 * the object sits at.
 */
export function paragraphDrawingObjects(paragraphXml: string): ParagraphDrawingObject[] {
  if (!OBJECT_MARKER_PATTERN.test(paragraphXml)) {
    return [];
  }
  const objects: ParagraphDrawingObject[] = [];
  const stack: string[] = [];
  let objectStart = -1;
  let objectName = "";
  let textLength = 0;
  let lastIndex = 0;

  const flushText = (value: string): void => {
    if (value) textLength += decodeXmlEntities(value).length;
  };

  for (const match of paragraphXml.matchAll(TAG_PATTERN)) {
    const index = match.index ?? 0;
    const name = match[3];

    if (objectStart < 0) {
      flushText(paragraphXml.slice(lastIndex, index));
    }
    lastIndex = index + match[0].length;

    if (match[1] !== undefined) {
      // CDATA counts as text unless it belongs to the captured object.
      if (objectStart < 0) flushText(match[1]);
      continue;
    }
    if (!name) continue;

    const closing = match[2] === "/";
    const selfClosing = match[5] === "/";

    if (objectStart >= 0) {
      if (closing) {
        const open = stack.pop();
        if (!open) continue;
        if (stack.length === 0) {
          objects.push({
            kind: classifyObjectBlock(objectName, paragraphXml.slice(objectStart, lastIndex)),
            sourceXml: paragraphXml.slice(objectStart, lastIndex),
            textOffset: textLength,
          });
          objectStart = -1;
          objectName = "";
        }
        continue;
      }
      if (!selfClosing) stack.push(name);
      else {
        // A self-closing tag never closes an object.
        if (name === objectName && stack.length === 0) {
          objects.push({
            kind: classifyObjectBlock(objectName, paragraphXml.slice(objectStart, lastIndex)),
            sourceXml: paragraphXml.slice(objectStart, lastIndex),
            textOffset: textLength,
          });
          objectStart = -1;
          objectName = "";
        }
      }
      continue;
    }

    if (OBJECT_ROOT_NAMES.has(name)) {
      if (selfClosing) {
        continue;
      }
      objectStart = index;
      objectName = name;
      stack.length = 0;
      stack.push(name);
      continue;
    }

    if (name === "w:t" && !selfClosing) {
      // Handled by the generic text flush below.
      continue;
    }
  }

  return objects;
}

function classifyObjectBlock(
  rootName: string,
  sourceXml: string,
): DrawingObjectKind | "picture" | "math" {
  if (rootName === "m:oMath" || rootName === "m:oMathPara") {
    return "math";
  }
  return classifyDrawingXml(sourceXml) ?? "shape";
}

/** Linearized equation text of an `m:oMath` / `m:oMathPara` block. */
export function mathTextFromXml(sourceXml: string): string {
  return linearizeOmml(sourceXml);
}

export function mathIsDisplay(sourceXml: string): boolean {
  return /<m:oMathPara\b/i.test(sourceXml);
}

// ── Chart parts ─────────────────────────────────────────────────────

function chartSeriesNames(series: ReturnType<typeof findXmlElements>): string | undefined {
  const text = series
    .filter((element) => element.name === "c:v")
    .map((element) => element.text.trim())
    .filter(Boolean)
    .join(" ");
  return text || undefined;
}

/**
 * Reads the cached data of a chart part: plot type, title, categories and one
 * entry per series. Chart parts without cached values still yield a title and
 * the plot type.
 */
export function parseChartData(chartXml: string | undefined): ChartData | undefined {
  if (!chartXml) return undefined;
  const elements = parseXmlElements(chartXml);
  if (elements.length === 0) return undefined;

  const chartType =
    /<c:(barChart|bar3DChart|lineChart|line3DChart|pieChart|pie3DChart|doughnutChart|areaChart|area3DChart|scatterChart|radarChart|bubbleChart|stockChart|surfaceChart|surface3DChart|ofPieChart)\b/.exec(
      chartXml,
    )?.[1];

  const titleElement = findXmlElements(elements, "c:title")[0];
  const title = titleElement
    ? findXmlElements([titleElement], "a:t")
        .map((element) => element.text)
        .join("")
        .trim()
    : "";

  const series: NonNullable<ChartData["series"]> = [];
  let categories: string[] | undefined;
  for (const ser of findXmlElements(elements, "c:ser")) {
    const nameElement = findXmlElements([ser], "c:tx")[0];
    const name = nameElement ? chartSeriesNames(findXmlElements([nameElement], "c:v")) : undefined;

    const categoryElement = findXmlElements([ser], "c:cat")[0];
    const seriesCategories = categoryElement
      ? findXmlElements([categoryElement], "c:pt")
          .map((point) => {
            const value = findXmlElements([point], "c:v")[0];
            return value ? value.text.trim() : "";
          })
          .filter(Boolean)
      : [];
    if (!categories && seriesCategories.length > 0) categories = seriesCategories;

    const valueElement = findXmlElements([ser], "c:val")[0];
    const values = valueElement
      ? findXmlElements([valueElement], "c:pt").map((point) => {
          const value = findXmlElements([point], "c:v")[0];
          const parsed = Number(value?.text);
          return Number.isFinite(parsed) ? parsed : null;
        })
      : [];

    series.push({
      ...(name ? { name } : {}),
      ...(seriesCategories.length > 0 ? { categories: seriesCategories } : {}),
      ...(values.length > 0 ? { values } : {}),
    });
  }

  if (!chartType && !title && series.length === 0) return undefined;
  return {
    ...(chartType ? { chartType } : {}),
    ...(title ? { title } : {}),
    ...(categories ? { categories } : {}),
    ...(series.length > 0 ? { series } : {}),
  };
}

/** `r:id` / `r:embed` references of a drawing's chart / model poster. */
export function drawingRelationshipId(
  sourceXml: string,
  tagName: string,
): string | undefined {
  const tag = new RegExp(`<${tagName}\\b[^>]*>`, "i").exec(sourceXml)?.[0];
  if (!tag) return undefined;
  return readAttribute(tag, "r:id") ?? readAttribute(tag, "r:embed");
}

/** Poster image relationship of a `a:model3d` block. */
export function model3dPosterRelationshipId(sourceXml: string): string | undefined {
  const model = /<a:model3d\b[\s\S]*?<\/a:model3d>/i.exec(sourceXml)?.[0];
  if (!model) return undefined;
  const blip = /<a:blip\b[^>]*>/i.exec(model)?.[0];
  return readAttribute(blip, "r:embed") ?? readAttribute(blip, "r:link");
}

/** OLE program id (`o:OLEObject/@ProgID`) of an embedded object. */
export function oleObjectInfo(
  sourceXml: string,
): { progId?: string; drawAspect?: string } | undefined {
  const tag = /<o:OLEObject\b[^>]*>/i.exec(sourceXml)?.[0];
  if (!tag) return undefined;
  const progId = readAttribute(tag, "ProgID");
  const drawAspect = readAttribute(tag, "DrawAspect");
  return {
    ...(progId ? { progId } : {}),
    ...(drawAspect ? { drawAspect } : {}),
  };
}
