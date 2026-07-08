import type { XlsxThemePalette } from "../types";

import {
  getLocalChildren,
  getFirstLocalChild,
  getFirstLocalDescendant,
  getLocalDescendants,
  readChartNumericAttribute,
} from "./chart-parser";

export function normalizeWorksheetVisibility(value: unknown): "hidden" | "veryHidden" | "visible" {
  return value === "hidden" || value === "veryHidden" ? value : "visible";
}
export const EMU_PER_PIXEL = 9525;
const THEME_COLOR_INDEX_BY_NAME: Record<string, number> = {
  accent1: 4,
  accent2: 5,
  accent3: 6,
  accent4: 7,
  accent5: 8,
  accent6: 9,
  dk1: 1,
  dk2: 3,
  folHlink: 11,
  hlink: 10,
  lt1: 0,
  lt2: 2,
  tx1: 1,
  tx2: 3,
  bg1: 0,
  bg2: 2
};

export function clampUnitInterval(value: number) {
  return Math.max(0, Math.min(1, value));
}

export function isElementNode(node: Node | ChildNode | null | undefined): node is Element {
  return node != null && node.nodeType === 1;
}

export function normalizeHexColor(value: string) {
  const hex = value.replace(/^#/, "");
  if (hex.length === 8) {
    return `#${hex.slice(2).toLowerCase()}`;
  }
  if (hex.length === 6) {
    return `#${hex.toLowerCase()}`;
  }
  return null;
}

export function resolveColorFromXmlFragment(fragment: string, themePalette?: XlsxThemePalette | null) {
  if (!fragment) {
    return undefined;
  }

  const srgbMatch = fragment.match(/<a:srgbClr\b[^>]*\bval="([0-9a-fA-F]{6,8})"/i);
  if (srgbMatch?.[1]) {
    return normalizeHexColor(srgbMatch[1]) ?? undefined;
  }

  const schemeMatch = fragment.match(/<a:schemeClr\b[^>]*\bval="([^"]+)"[^>]*>([\s\S]*?)<\/a:schemeClr>/i)
    ?? fragment.match(/<a:schemeClr\b[^>]*\bval="([^"]+)"[^>]*/i);
  if (!schemeMatch?.[1]) {
    return undefined;
  }

  const baseColor = resolveThemeColor(schemeMatch[1], themePalette);
  if (!baseColor) {
    return undefined;
  }

  const transforms = schemeMatch[2] ?? "";
  let lightnessModifier = 1;
  let lightnessOffset = 0;
  for (const match of transforms.matchAll(/<a:(lumMod|lumOff|tint|shade)\b[^>]*\bval="(-?\d+(?:\.\d+)?)"/gi)) {
    const transform = match[1]?.toLowerCase();
    const rawValue = Number(match[2] ?? Number.NaN);
    if (!transform || !Number.isFinite(rawValue)) {
      continue;
    }
    if (transform === "lummod") {
      lightnessModifier *= rawValue / 100000;
    } else if (transform === "lumoff") {
      lightnessOffset += rawValue / 100000;
    } else if (transform === "tint") {
      lightnessOffset += (1 - lightnessOffset) * (rawValue / 100000);
    } else if (transform === "shade") {
      lightnessModifier *= rawValue / 100000;
    }
  }

  return applyLightnessTransform(baseColor, lightnessModifier, lightnessOffset) ?? undefined;
}

export function readHexColorFromXmlFragment(
  fragment: string,
  preferLine = false,
  themePalette?: XlsxThemePalette | null
) {
  const source = preferLine
    ? fragment.match(/<a:ln\b[\s\S]*?<\/a:ln>/i)?.[0] ?? ""
    : fragment.match(/<a:solidFill\b[\s\S]*?<\/a:solidFill>/i)?.[0] ?? "";
  return resolveColorFromXmlFragment(source, themePalette);
}

export function parseHexColor(color: string): [number, number, number] | null {
  const normalized = normalizeHexColor(color);
  if (!normalized) {
    return null;
  }
  const match = /^#([0-9a-f]{6})$/.exec(normalized);
  if (!match) {
    return null;
  }
  return [
    Number.parseInt(match[1].slice(0, 2), 16),
    Number.parseInt(match[1].slice(2, 4), 16),
    Number.parseInt(match[1].slice(4, 6), 16)
  ];
}

export function rgbToHsl(red: number, green: number, blue: number): [number, number, number] {
  const normalizedRed = red / 255;
  const normalizedGreen = green / 255;
  const normalizedBlue = blue / 255;
  const max = Math.max(normalizedRed, normalizedGreen, normalizedBlue);
  const min = Math.min(normalizedRed, normalizedGreen, normalizedBlue);
  const lightness = (max + min) / 2;

  if (max === min) {
    return [0, 0, lightness];
  }

  const delta = max - min;
  const saturation = lightness > 0.5 ? delta / (2 - max - min) : delta / (max + min);
  let hue = 0;

  switch (max) {
    case normalizedRed:
      hue = (normalizedGreen - normalizedBlue) / delta + (normalizedGreen < normalizedBlue ? 6 : 0);
      break;
    case normalizedGreen:
      hue = (normalizedBlue - normalizedRed) / delta + 2;
      break;
    default:
      hue = (normalizedRed - normalizedGreen) / delta + 4;
      break;
  }

  return [hue / 6, saturation, lightness];
}

export function hueToRgb(p: number, q: number, t: number) {
  let nextT = t;
  if (nextT < 0) {
    nextT += 1;
  }
  if (nextT > 1) {
    nextT -= 1;
  }
  if (nextT < 1 / 6) {
    return p + (q - p) * 6 * nextT;
  }
  if (nextT < 1 / 2) {
    return q;
  }
  if (nextT < 2 / 3) {
    return p + (q - p) * (2 / 3 - nextT) * 6;
  }
  return p;
}

export function hslToRgb(hue: number, saturation: number, lightness: number): [number, number, number] {
  if (saturation === 0) {
    const gray = Math.round(lightness * 255);
    return [gray, gray, gray];
  }

  const q = lightness < 0.5 ? lightness * (1 + saturation) : lightness + saturation - lightness * saturation;
  const p = 2 * lightness - q;

  return [
    Math.round(hueToRgb(p, q, hue + 1 / 3) * 255),
    Math.round(hueToRgb(p, q, hue) * 255),
    Math.round(hueToRgb(p, q, hue - 1 / 3) * 255)
  ];
}

export function rgbToHex(red: number, green: number, blue: number) {
  return `#${[red, green, blue]
    .map((channel) => Math.max(0, Math.min(255, Math.round(channel))).toString(16).padStart(2, "0"))
    .join("")}`;
}

export function applyLightnessTransform(baseColor: string, modifier = 1, offset = 0) {
  const rgb = parseHexColor(baseColor);
  if (!rgb) {
    return normalizeHexColor(baseColor);
  }

  const [hue, saturation, lightness] = rgbToHsl(rgb[0], rgb[1], rgb[2]);
  const nextLightness = clampUnitInterval(lightness * modifier + offset);
  const [nextRed, nextGreen, nextBlue] = hslToRgb(hue, saturation, nextLightness);
  return rgbToHex(nextRed, nextGreen, nextBlue);
}

export function resolveThemeColor(name: string | null, themePalette?: XlsxThemePalette | null) {
  if (!name) {
    return null;
  }
  const index = THEME_COLOR_INDEX_BY_NAME[name];
  return index === undefined ? null : themePalette?.colorsByIndex[index] ?? null;
}

export function resolveThemeTypeface(typeface: string | null, themePalette?: XlsxThemePalette | null) {
  if (!typeface) {
    return null;
  }
  if (typeface === "+mn-lt" || typeface === "+mn-ea" || typeface === "+mn-cs") {
    return themePalette?.minorLatinFont ?? null;
  }
  if (typeface === "+mj-lt" || typeface === "+mj-ea" || typeface === "+mj-cs") {
    return themePalette?.majorLatinFont ?? null;
  }
  return typeface;
}

export function readChartTextTypeface(textPropertiesNode: Element | null, themePalette?: XlsxThemePalette | null) {
  if (!textPropertiesNode) {
    return null;
  }
  const defaultRunProperties = getFirstLocalDescendant(textPropertiesNode, "defRPr")
    ?? getFirstLocalDescendant(textPropertiesNode, "rPr");
  if (!defaultRunProperties) {
    return null;
  }
  const typeface = getFirstLocalChild(defaultRunProperties, "latin")?.getAttribute("typeface")
    ?? getFirstLocalChild(defaultRunProperties, "ea")?.getAttribute("typeface")
    ?? getFirstLocalChild(defaultRunProperties, "cs")?.getAttribute("typeface")
    ?? null;
  const resolved = resolveThemeTypeface(typeface, themePalette)?.trim() ?? "";
  return resolved.length > 0 ? resolved : null;
}

export function resolveChartColorNode(node: Element | null, themePalette?: XlsxThemePalette | null): string | null {
  if (!node) {
    return null;
  }

  let baseColor: string | null = null;
  if (node.localName === "srgbClr") {
    baseColor = normalizeHexColor(`#${node.getAttribute("val") ?? ""}`);
  } else if (node.localName === "schemeClr") {
    baseColor = resolveThemeColor(node.getAttribute("val"), themePalette);
  } else if (node.localName === "sysClr") {
    baseColor = normalizeHexColor(`#${node.getAttribute("lastClr") ?? ""}`);
  }

  if (!baseColor) {
    return null;
  }

  let lightnessModifier = 1;
  let lightnessOffset = 0;
  for (const transformNode of Array.from(node.childNodes).filter(isElementNode)) {
    const rawValue = Number(transformNode.getAttribute("val") ?? Number.NaN);
    if (!Number.isFinite(rawValue)) {
      continue;
    }
    if (transformNode.localName === "lumMod") {
      lightnessModifier *= rawValue / 100000;
    } else if (transformNode.localName === "lumOff") {
      lightnessOffset += rawValue / 100000;
    } else if (transformNode.localName === "tint") {
      lightnessOffset += (1 - lightnessOffset) * (rawValue / 100000);
    } else if (transformNode.localName === "shade") {
      lightnessModifier *= rawValue / 100000;
    }
  }

  return applyLightnessTransform(baseColor, lightnessModifier, lightnessOffset);
}

export function isChartColorElement(node: Element | null | undefined): node is Element {
  return Boolean(node && (node.localName === "schemeClr" || node.localName === "srgbClr" || node.localName === "sysClr"));
}

export function findFirstChartColorElement(node: Element | null) {
  if (!node) {
    return null;
  }
  if (isChartColorElement(node)) {
    return node;
  }

  for (const localName of ["srgbClr", "schemeClr", "sysClr"]) {
    for (const candidate of getLocalDescendants(node, localName)) {
      if (isChartColorElement(candidate)) {
        return candidate;
      }
    }
  }

  return null;
}

export function resolveChartFillColor(shapeNode: Element | null, themePalette?: XlsxThemePalette | null) {
  if (!shapeNode || getFirstLocalChild(shapeNode, "noFill")) {
    return null;
  }
  const solidFill = getFirstLocalChild(shapeNode, "solidFill");
  if (solidFill) {
    const colorNode = findFirstChartColorElement(Array.from(solidFill.childNodes).find(isElementNode) ?? null);
    return resolveChartColorNode(colorNode, themePalette);
  }

  const gradientFill = getFirstLocalChild(shapeNode, "gradFill");
  const gradientStops = gradientFill
    ? getLocalDescendants(gradientFill, "gs")
        .map((stopNode) => ({
          colorNode: Array.from(stopNode.childNodes).find(isElementNode) ?? null,
          position: Number(stopNode.getAttribute("pos") ?? Number.NaN)
        }))
        .filter((stop) => Boolean(stop.colorNode))
    : [];
  if (gradientStops.length === 0) {
    return null;
  }

  gradientStops.sort((left, right) => {
    const leftPos = Number.isFinite(left.position) ? left.position : 0;
    const rightPos = Number.isFinite(right.position) ? right.position : 0;
    return leftPos - rightPos;
  });
  const midpointStop = gradientStops.find((stop) => Number.isFinite(stop.position) && stop.position >= 50000)
    ?? gradientStops[Math.floor(gradientStops.length / 2)]
    ?? gradientStops[0];
  return resolveChartColorNode(midpointStop.colorNode, themePalette);
}

export function resolveChartLineStyle(shapeNode: Element | null, themePalette?: XlsxThemePalette | null) {
  const lineNode = shapeNode?.localName === "ln" ? shapeNode : (shapeNode ? getFirstLocalChild(shapeNode, "ln") : null);
  if (!lineNode) {
    return { color: null, hidden: false, widthPx: undefined };
  }
  if (getFirstLocalChild(lineNode, "noFill")) {
    return { color: null, hidden: true, widthPx: undefined };
  }

  const solidFill = getFirstLocalChild(lineNode, "solidFill");
  const colorNode = solidFill ? findFirstChartColorElement(Array.from(solidFill.childNodes).find(isElementNode) ?? null) : null;
  const widthValue = Number(lineNode.getAttribute("w") ?? Number.NaN);
  return {
    color: resolveChartColorNode(colorNode, themePalette),
    hidden: false,
    widthPx: Number.isFinite(widthValue) ? Math.max(1, widthValue / EMU_PER_PIXEL) : undefined
  };
}


