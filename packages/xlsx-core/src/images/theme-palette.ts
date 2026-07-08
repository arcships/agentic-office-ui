import { resolveWorkbookColor } from "../colors";
import type {
  XlsxResolvedCellStyle,
  XlsxTableStyleDefinition,
  XlsxThemePalette,
} from "../types";
import {
  type ArchiveEntries,
  getChildElements,
  getFirstChild,
  getFirstDescendant,
  getLocalElements,
  isElementNode,
  normalizeArchivePath,
  normalizeHexColor,
  readArchiveText,
  parseXml,
} from "./image-parser";

// ── Theme types ────────────────────────────────────────────────────────────

export type ThemeState = {
  colors: Map<string, string>;
  majorLatinFont: string | null;
  minorLatinFont: string | null;
};

// ── Theme palette ──────────────────────────────────────────────────────────

export function buildThemePalette(theme: ThemeState): XlsxThemePalette {
  const themeOrder = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
  const colorsByIndex: Record<number, string> = {};

  themeOrder.forEach((key, index) => {
    const color = theme.colors.get(key);
    if (color) {
      colorsByIndex[index] = color;
    }
  });

  return {
    colorsByIndex,
    majorLatinFont: theme.majorLatinFont ?? undefined,
    minorLatinFont: theme.minorLatinFont ?? undefined
  };
}

// ── Theme parsing ──────────────────────────────────────────────────────────

export function parseWorkbookTheme(archive: ArchiveEntries): ThemeState {
  const defaultTheme: ThemeState = {
    colors: new Map([
      ["accent1", "#5b9bd5"],
      ["accent2", "#ed7d31"],
      ["accent3", "#a5a5a5"],
      ["accent4", "#ffc000"],
      ["accent5", "#4472c4"],
      ["accent6", "#70ad47"],
      ["bg1", "#ffffff"],
      ["bg2", "#e7e6e6"],
      ["dk1", "#000000"],
      ["dk2", "#6e747a"],
      ["folHlink", "#993366"],
      ["hlink", "#085296"],
      ["lt1", "#ffffff"],
      ["lt2", "#e7e6e6"],
      ["tx1", "#000000"],
      ["tx2", "#6e747a"]
    ]),
    majorLatinFont: null,
    minorLatinFont: null
  };

  const themeXml = readArchiveText(archive, "xl/theme/theme1.xml");
  if (!themeXml) {
    return defaultTheme;
  }

  const themeDocument = parseXml(themeXml);
  if (!themeDocument) {
    return defaultTheme;
  }

  const colors = new Map(defaultTheme.colors);
  const colorSchemeNode = getLocalElements(themeDocument, "clrScheme")[0] ?? null;
  if (colorSchemeNode) {
    for (const colorNode of Array.from(colorSchemeNode.childNodes).filter(isElementNode)) {
      const key = colorNode.localName;
      const srgbNode = getFirstChild(colorNode, "srgbClr");
      const sysNode = getFirstChild(colorNode, "sysClr");
      const hex = srgbNode?.getAttribute("val") ?? sysNode?.getAttribute("lastClr");
      if (hex) {
        colors.set(key, normalizeHexColor(hex));
      }
    }
  }

  const fontSchemeNode = getLocalElements(themeDocument, "fontScheme")[0] ?? null;
  const majorLatinFont = getFirstChild(getFirstChild(fontSchemeNode, "majorFont"), "latin")?.getAttribute("typeface") ?? null;
  const minorLatinFont = getFirstChild(getFirstChild(fontSchemeNode, "minorFont"), "latin")?.getAttribute("typeface") ?? null;

  colors.set("bg1", colors.get("lt1") ?? defaultTheme.colors.get("bg1") ?? "#ffffff");
  colors.set("tx1", colors.get("dk1") ?? defaultTheme.colors.get("tx1") ?? "#000000");
  colors.set("bg2", colors.get("lt2") ?? defaultTheme.colors.get("bg2") ?? "#e7e6e6");
  colors.set("tx2", colors.get("dk2") ?? defaultTheme.colors.get("tx2") ?? "#6e747a");

  return {
    colors,
    majorLatinFont,
    minorLatinFont
  };
}

// ── Spreadsheet color parsing ──────────────────────────────────────────────

export function parseSpreadsheetColor(node: Element | null) {
  if (!node) {
    return undefined;
  }

  const color: Record<string, unknown> = {};
  const rgb = node.getAttribute("rgb");
  const theme = node.getAttribute("theme");
  const tint = node.getAttribute("tint");
  const indexed = node.getAttribute("indexed");
  if (rgb) {
    color.rgb = normalizeHexColor(rgb);
  }
  if (theme !== null) {
    color.theme = Number(theme);
  }
  if (tint !== null) {
    color.tint = Number(tint);
  }
  if (indexed !== null) {
    color.indexed = Number(indexed);
  }

  return Object.keys(color).length > 0 ? color : undefined;
}

// ── Spreadsheet boolean flag ───────────────────────────────────────────────

export function hasEnabledSpreadsheetFlag(node: Element | null) {
  if (!node) {
    return false;
  }

  const value = node.getAttribute("val");
  return value === null || (value !== "0" && value !== "false");
}

// ── Style parsing ──────────────────────────────────────────────────────────

export function parseSpreadsheetFont(node: Element | null): XlsxResolvedCellStyle["font"] {
  if (!node) {
    return undefined;
  }

  const font: Record<string, unknown> = {};
  const size = getFirstChild(node, "sz")?.getAttribute("val");
  const name = getFirstChild(node, "name")?.getAttribute("val");
  const family = getFirstChild(node, "family")?.getAttribute("val");
  const scheme = getFirstChild(node, "scheme")?.getAttribute("val");
  const charset = getFirstChild(node, "charset")?.getAttribute("val");
  const verticalAlign = getFirstChild(node, "vertAlign")?.getAttribute("val");
  const color = parseSpreadsheetColor(getFirstChild(node, "color"));
  if (hasEnabledSpreadsheetFlag(getFirstChild(node, "b"))) {
    font.bold = true;
  }
  if (hasEnabledSpreadsheetFlag(getFirstChild(node, "i"))) {
    font.italic = true;
  }
  if (hasEnabledSpreadsheetFlag(getFirstChild(node, "strike"))) {
    font.strikethrough = true;
  }
  if (getFirstChild(node, "u")) {
    font.underline = getFirstChild(node, "u")?.getAttribute("val") ?? "single";
  }
  if (size !== null && size !== undefined) {
    font.size = Number(size);
  }
  if (name) {
    font.name = name;
  }
  if (family !== null && family !== undefined) {
    font.family = Number(family);
  }
  if (scheme) {
    font.scheme = scheme;
  }
  if (charset !== null && charset !== undefined) {
    font.charset = Number(charset);
  }
  if (verticalAlign) {
    font.verticalAlign = verticalAlign;
  }
  if (hasEnabledSpreadsheetFlag(getFirstChild(node, "shadow"))) {
    font.shadow = true;
  }
  if (hasEnabledSpreadsheetFlag(getFirstChild(node, "outline"))) {
    font.outline = true;
  }
  if (hasEnabledSpreadsheetFlag(getFirstChild(node, "condense"))) {
    font.condense = true;
  }
  if (hasEnabledSpreadsheetFlag(getFirstChild(node, "extend"))) {
    font.extend = true;
  }
  if (color) {
    font.color = color;
  }

  return Object.keys(font).length > 0 ? font : undefined;
}

export function parseSpreadsheetFill(node: Element | null): XlsxResolvedCellStyle["fill"] {
  if (!node) {
    return undefined;
  }

  const gradientFill = getFirstChild(node, "gradientFill");
  if (gradientFill) {
    const stops = Array.from(gradientFill.childNodes)
      .filter(isElementNode)
      .filter((child) => child.localName === "stop")
      .map((stopNode) => ({
        color: parseSpreadsheetColor(Array.from(stopNode.childNodes).find(isElementNode) ?? null),
        position: Number(stopNode.getAttribute("position") ?? Number.NaN)
      }))
      .filter((stop) => stop.color && Number.isFinite(stop.position));
    if (stops.length > 0) {
      return {
        degree: Number(gradientFill.getAttribute("degree") ?? 0),
        fillType: "gradient",
        gradientType: gradientFill.getAttribute("type") ?? "linear",
        stops
      };
    }
  }

  const patternFill = getFirstChild(node, "patternFill");
  if (!patternFill) {
    return undefined;
  }

  const patternType = patternFill.getAttribute("patternType") ?? "none";
  const foreground = parseSpreadsheetColor(getFirstChild(patternFill, "fgColor"));
  const background = parseSpreadsheetColor(getFirstChild(patternFill, "bgColor"));
  const solidColor = foreground ?? background;
  if (patternType === "solid" && solidColor) {
    return {
      color: solidColor,
      fillType: "solid"
    };
  }
  // Differential table styles sometimes omit patternType and only set bgColor.
  // Preserve those fills so header/table overrides are not dropped during import.
  if ((patternType === "none" || patternType === "gray125") && (foreground || background)) {
    return {
      background,
      fillType: "pattern",
      foreground,
      patternType
    };
  }
  if (patternType !== "none" && patternType !== "gray125" && (foreground || background)) {
    return {
      background,
      fillType: "pattern",
      foreground,
      patternType
    };
  }

  return undefined;
}

function parseSpreadsheetBorderEdge(node: Element | null) {
  if (!node) {
    return undefined;
  }

  const style = node.getAttribute("style");
  const color = parseSpreadsheetColor(getFirstChild(node, "color"));
  if (!style || style === "none") {
    return undefined;
  }

  return {
    color,
    style
  };
}

export function parseSpreadsheetBorder(node: Element | null): XlsxResolvedCellStyle["border"] {
  if (!node) {
    return undefined;
  }

  const border: Record<string, Record<string, unknown>> = {};
  (["top", "right", "bottom", "left", "horizontal", "vertical"] as const).forEach((edge) => {
    const parsedEdge = parseSpreadsheetBorderEdge(getFirstChild(node, edge));
    if (parsedEdge) {
      border[edge] = parsedEdge;
    }
  });

  return Object.keys(border).length > 0 ? border : undefined;
}

export function parseSpreadsheetAlignment(node: Element | null): XlsxResolvedCellStyle["alignment"] {
  if (!node) {
    return undefined;
  }

  const alignment: Record<string, unknown> = {};
  const horizontal = node.getAttribute("horizontal");
  const vertical = node.getAttribute("vertical");
  const wrapText = node.getAttribute("wrapText");
  const indent = node.getAttribute("indent");
  const shrinkToFit = node.getAttribute("shrinkToFit");
  const textRotation = node.getAttribute("textRotation");
  if (horizontal) {
    alignment.horizontal = horizontal;
  }
  if (vertical) {
    alignment.vertical = vertical;
  }
  if (wrapText !== null) {
    alignment.wrapText = wrapText === "1";
  }
  if (shrinkToFit !== null) {
    alignment.shrinkToFit = shrinkToFit === "1";
  }
  if (indent !== null) {
    alignment.indent = Number(indent);
  }
  if (textRotation !== null) {
    const parsedRotation = Number(textRotation);
    if (Number.isFinite(parsedRotation)) {
      alignment.textRotation = parsedRotation;
    }
  }

  return Object.keys(alignment).length > 0 ? alignment : undefined;
}

function parseDifferentialStyle(node: Element | null): XlsxResolvedCellStyle {
  if (!node) {
    return {};
  }

  const style: XlsxResolvedCellStyle = {};
  const font = parseSpreadsheetFont(getFirstChild(node, "font"));
  const fill = parseSpreadsheetFill(getFirstChild(node, "fill"));
  const border = parseSpreadsheetBorder(getFirstChild(node, "border"));
  const alignment = parseSpreadsheetAlignment(getFirstChild(node, "alignment"));

  if (font) {
    style.font = font;
  }
  if (fill) {
    style.fill = fill;
  }
  if (border) {
    style.border = border;
  }
  if (alignment) {
    style.alignment = alignment;
  }

  return style;
}

function parseResolvedXfStyle(
  xfNode: Element,
  fonts: Array<XlsxResolvedCellStyle["font"]>,
  fills: Array<XlsxResolvedCellStyle["fill"]>,
  borders: Array<XlsxResolvedCellStyle["border"]>,
  checkboxComplementIndices?: Set<number>
) {
  const style: XlsxResolvedCellStyle = {};
  const fontId = Number(xfNode.getAttribute("fontId") ?? Number.NaN);
  const fillId = Number(xfNode.getAttribute("fillId") ?? Number.NaN);
  const borderId = Number(xfNode.getAttribute("borderId") ?? Number.NaN);
  const alignment = parseSpreadsheetAlignment(getFirstChild(xfNode, "alignment"));

  if (Number.isFinite(fontId) && fonts[fontId]) {
    style.font = fonts[fontId];
  }
  if (Number.isFinite(fillId) && fills[fillId]) {
    style.fill = fills[fillId];
  }
  if (Number.isFinite(borderId) && borders[borderId]) {
    style.border = borders[borderId];
  }
  if (alignment) {
    style.alignment = alignment;
  }
  const xfComplementNode = getFirstDescendant(xfNode, "xfComplement");
  const xfComplementIndex = Number(xfComplementNode?.getAttribute("i") ?? Number.NaN);
  if (Number.isFinite(xfComplementIndex) && checkboxComplementIndices?.has(xfComplementIndex)) {
    style.cellControl = { kind: "checkbox" };
  }

  return style;
}

// ── Feature property bag ───────────────────────────────────────────────────

function readFeaturePropertyBagCheckboxComplements(archive: ArchiveEntries) {
  const xml = readArchiveText(archive, "xl/featurePropertyBag/featurePropertyBag.xml");
  if (!xml) {
    return new Set<number>();
  }

  const document = parseXml(xml);
  if (!document?.documentElement) {
    return new Set<number>();
  }

  const bagNodes = getChildElements(document.documentElement, "bag");
  const bagTypeById = bagNodes.map((node) => node.getAttribute("type") ?? "");
  const checkboxComplementIndices = new Set<number>();

  const xfComplementsBag = bagNodes.find((node) => node.getAttribute("type") === "XFComplements") ?? null;
  const mappedBagIds = xfComplementsBag
    ? getLocalElements(xfComplementsBag, "bagId")
      .map((node) => Number(node.textContent ?? Number.NaN))
      .filter((value) => Number.isFinite(value))
    : [];

  mappedBagIds.forEach((bagId, complementIndex) => {
    const xfComplementBag = bagNodes[bagId];
    if (!xfComplementBag || bagTypeById[bagId] !== "XFComplement") {
      return;
    }

    const xfControlsBagId = getLocalElements(xfComplementBag, "bagId")
      .map((node) => Number(node.textContent ?? Number.NaN))
      .find((value) => Number.isFinite(value));
    if (xfControlsBagId === undefined) {
      return;
    }

    const xfControlsBag = bagNodes[xfControlsBagId];
    if (!xfControlsBag || bagTypeById[xfControlsBagId] !== "XFControls") {
      return;
    }

    const cellControlBagId = getLocalElements(xfControlsBag, "bagId")
      .map((node) => Number(node.textContent ?? Number.NaN))
      .find((value) => Number.isFinite(value));
    if (cellControlBagId === undefined) {
      return;
    }

    if (bagTypeById[cellControlBagId] === "Checkbox") {
      checkboxComplementIndices.add(complementIndex);
    }
  });

  return checkboxComplementIndices;
}

// ── Workbook styles parsing ────────────────────────────────────────────────

export function parseWorkbookStyles(archive: ArchiveEntries) {
  const xml = readArchiveText(archive, "xl/styles.xml");
  if (!xml) {
    return {
      defaultFont: null,
      namedCellStyleByName: {},
      styleById: {},
      tableStyleByName: {}
    };
  }

  const document = parseXml(xml);
  if (!document) {
    return {
      defaultFont: null,
      namedCellStyleByName: {},
      styleById: {},
      tableStyleByName: {}
    };
  }

  const fontsNode = getFirstDescendant(document, "fonts");
  const fillsNode = getFirstDescendant(document, "fills");
  const bordersNode = getFirstDescendant(document, "borders");
  const cellStyleXfsNode = getFirstDescendant(document, "cellStyleXfs");
  const cellStylesNode = getFirstDescendant(document, "cellStyles");
  const cellXfsNode = getFirstDescendant(document, "cellXfs");
  const dxfsNode = getFirstDescendant(document, "dxfs");
  const tableStylesNode = getFirstDescendant(document, "tableStyles");
  if (!cellXfsNode) {
    return {
      defaultFont: null,
      namedCellStyleByName: {},
      styleById: {},
      tableStyleByName: {}
    };
  }

  const checkboxComplementIndices = readFeaturePropertyBagCheckboxComplements(archive);
  const fonts = getChildElements(fontsNode ?? document.documentElement, "font").map((node) => parseSpreadsheetFont(node));
  const fills = getChildElements(fillsNode ?? document.documentElement, "fill").map((node) => parseSpreadsheetFill(node));
  const borders = getChildElements(bordersNode ?? document.documentElement, "border").map((node) => parseSpreadsheetBorder(node));
  const differentialStyles = getChildElements(dxfsNode ?? document.documentElement, "dxf").map((node) => parseDifferentialStyle(node));
  const cellStyleXfs = getChildElements(cellStyleXfsNode ?? document.documentElement, "xf").map(
    (node) => parseResolvedXfStyle(node, fonts, fills, borders, checkboxComplementIndices)
  );
  const namedCellStyleByName: Record<string, XlsxResolvedCellStyle> = {};
  const styleById: Record<number, XlsxResolvedCellStyle> = {};
  const tableStyleByName: Record<string, XlsxTableStyleDefinition> = {};

  getChildElements(cellXfsNode, "xf").forEach((xfNode, index) => {
    styleById[index] = parseResolvedXfStyle(xfNode, fonts, fills, borders, checkboxComplementIndices);
  });

  getChildElements(cellStylesNode ?? document.documentElement, "cellStyle").forEach((cellStyleNode) => {
    const name = cellStyleNode.getAttribute("name");
    const xfId = Number(cellStyleNode.getAttribute("xfId") ?? Number.NaN);
    if (!name || !Number.isFinite(xfId)) {
      return;
    }

    const resolvedStyle = cellStyleXfs[xfId];
    if (resolvedStyle) {
      namedCellStyleByName[name] = resolvedStyle;
    }
  });

  getChildElements(tableStylesNode ?? document.documentElement, "tableStyle").forEach((tableStyleNode) => {
    const name = tableStyleNode.getAttribute("name");
    if (!name) {
      return;
    }

    const elements: XlsxTableStyleDefinition = {};
    getChildElements(tableStyleNode, "tableStyleElement").forEach((elementNode) => {
      const type = elementNode.getAttribute("type");
      const dxfId = Number(elementNode.getAttribute("dxfId") ?? Number.NaN);
      if (!type || !Number.isFinite(dxfId)) {
        return;
      }

      const differentialStyle = differentialStyles[dxfId];
      if (differentialStyle) {
        elements[type] = differentialStyle;
      }
    });
    tableStyleByName[name] = elements;
  });

  const normalFont = (namedCellStyleByName.Normal?.font ?? styleById[0]?.font ?? fonts[0]) as Record<string, unknown> | undefined;
  const defaultFont = normalFont ? {
    family: typeof normalFont.name === "string" ? normalFont.name : undefined,
    sizePt: typeof normalFont.size === "number" ? normalFont.size : undefined
  } : null;

  return {
    defaultFont,
    namedCellStyleByName,
    styleById,
    tableStyleByName
  };
}
