import type { DocumentBackgroundPattern } from "./types";

export interface ParsedDocumentBackground {
  color?: string;
  pattern?: DocumentBackgroundPattern;
}

function normalizeHexColor(value?: string): string | undefined {
  if (!value) return undefined;
  const normalized = value.trim();
  if (!normalized || normalized.toLowerCase() === "auto") return undefined;
  if (/^[0-9a-fA-F]{6}$/.test(normalized)) return `#${normalized}`;
  if (/^#[0-9a-fA-F]{6}$/.test(normalized)) return normalized;
  return undefined;
}

function readAttribute(tagXml: string | undefined, attribute: string): string | undefined {
  if (!tagXml) return undefined;
  const escaped = attribute.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return tagXml.match(new RegExp(`\\b${escaped}\\s*=\\s*(?:"([^"]+)"|'([^']+)')`, "i"))
    ?.slice(1)
    .find((value) => value !== undefined);
}

function drawingMlColor(
  patternXml: string | undefined,
  container: "fgClr" | "bgClr"
): string | undefined {
  if (!patternXml) return undefined;
  const block = patternXml.match(
    new RegExp(`<a:${container}\\b[^>]*>[\\s\\S]*?</a:${container}>`, "i")
  )?.[0];
  if (!block) return undefined;
  const srgb = block.match(/<a:srgbClr\b[^>]*\bval\s*=\s*(?:"([^"]+)"|'([^']+)')/i);
  return normalizeHexColor(srgb?.[1] ?? srgb?.[2]);
}

function parseDrawingMlPattern(backgroundXml: string): DocumentBackgroundPattern | undefined {
  const fillTag = backgroundXml.match(/<a:pattFill\b[^>]*>/i)?.[0];
  if (!fillTag) return undefined;
  const fillBlock = backgroundXml.match(/<a:pattFill\b[\s\S]*?<\/a:pattFill>/i)?.[0];
  const preset = readAttribute(fillTag, "prst");
  if (!preset) return undefined;
  return {
    preset,
    foregroundColor: drawingMlColor(fillBlock, "fgClr"),
    backgroundColor: drawingMlColor(fillBlock, "bgClr"),
  };
}

function parseVmlPattern(backgroundXml: string): DocumentBackgroundPattern | undefined {
  const fillTag = backgroundXml.match(/<v:fill\b[^>]*>/i)?.[0];
  if (!fillTag) return undefined;
  const fillType = readAttribute(fillTag, "type")?.trim().toLowerCase();
  const preset = readAttribute(fillTag, "pattern");
  if (!preset || (fillType && fillType !== "pattern" && fillType !== "tile")) {
    return undefined;
  }
  return {
    preset,
    foregroundColor: normalizeHexColor(readAttribute(fillTag, "color")),
    backgroundColor: normalizeHexColor(readAttribute(fillTag, "color2")),
  };
}

function parseShadingPattern(backgroundXml: string): DocumentBackgroundPattern | undefined {
  const shadingTag = backgroundXml.match(/<w:shd\b[^>]*>/i)?.[0];
  if (!shadingTag) return undefined;
  const preset = readAttribute(shadingTag, "w:val")?.trim();
  if (!preset || /^(?:clear|nil|solid)$/i.test(preset)) return undefined;
  return {
    preset,
    foregroundColor: normalizeHexColor(readAttribute(shadingTag, "w:color")),
    backgroundColor: normalizeHexColor(readAttribute(shadingTag, "w:fill")),
  };
}

/**
 * Reads a page background from `word/document.xml`. Word stores a page pattern
 * in the optional `w:background/w:drawing` shape fill; some producers use VML
 * or a `w:shd` shading element instead. Solid `w:background/@w:color` is also
 * returned so callers keep a single background parsing path.
 */
export function parseDocumentBackground(
  documentXml?: string
): ParsedDocumentBackground | undefined {
  if (!documentXml) return undefined;

  const block = documentXml.match(
    /<w:background\b[^>]*>[\s\S]*?<\/w:background>/i
  )?.[0];
  const tag =
    block?.match(/<w:background\b[^>]*>/i)?.[0] ??
    documentXml.match(/<w:background\b[^>]*\/?>/i)?.[0];
  if (!tag) return undefined;

  const color = normalizeHexColor(readAttribute(tag, "w:color"));
  const pattern = block
    ? parseDrawingMlPattern(block) ??
      parseVmlPattern(block) ??
      parseShadingPattern(block)
    : undefined;

  return color || pattern ? { color, pattern } : undefined;
}
