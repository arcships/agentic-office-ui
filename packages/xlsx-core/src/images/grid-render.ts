import { resolveWorkbookColor } from "../colors";
import type {
  XlsxCellAddress,
  XlsxConditionalDataBarRule,
  XlsxConditionalFormatIcon,
  XlsxConditionalFormatRule,
  XlsxConditionalFormatValueObject,
  XlsxConditionalIconSetRule,
  XlsxCellRange,
  XlsxFormControl,
  XlsxImage,
  XlsxResolvedCellStyle,
  XlsxShape,
  XlsxSparkline,
  XlsxTableStyleDefinition,
  XlsxThemePalette,
} from "../types";
import {
  type ArchiveEntries,
  type ContentTypesState,
  DRAWING_REL_TYPE,
  REL_NS,
  type RelationshipRecord,
  type WorkbookImageOrigin,
  type WorkbookImageSheetOrigin,
  getChildElements,
  getFirstChild,
  getFirstDescendant,
  getLocalElements,
  getRelationshipId,
  parseA1CellReference,
  parseA1RangeReference,
  parseColumnReference,
  parseContentTypes,
  parseRelationships,
  parseXml,
  readArchiveText,
  relsPathForDocument,
} from "./image-parser";
import {
  measureColumnCharacterWidthPx,
  MIN_COL_WIDTH_PX,
  MIN_ROW_HEIGHT_PX,
  sheetColumnWidthToPixels,
  resolveWorksheetDefaultColumnWidthPixels,
  resolveWorksheetDefaultRowHeightPixels,
} from "./column-width";
import {
  type ThemeState,
  buildThemePalette,
  hasEnabledSpreadsheetFlag,
  parseSpreadsheetAlignment,
  parseSpreadsheetColor,
  parseSpreadsheetFont,
  parseSpreadsheetFill,
  parseSpreadsheetBorder,
  parseWorkbookStyles,
  parseWorkbookTheme,
} from "./theme-palette";

// ── Types ───────────────────────────────────────────────────────────────────

type WorkbookSheetInfo = {
  name: string;
  path: string;
};

export type WorkbookSheetState = {
  cachedFormulaValues: Record<string, string>;
  columnWidthCharacterWidthPx?: number;
  colWidthOverridesPx: Record<number, number>;
  colStyleIds: Record<number, number>;
  conditionalFormatRules: XlsxConditionalFormatRule[];
  defaultColWidthPx: number;
  defaultRowHeightPx: number;
  hasHorizontalMerges: boolean;
  hasVerticalMerges: boolean;
  maxHorizontalMergeEndCol: number;
  maxVerticalMergeEndRow: number;
  maxContentCol: number;
  maxContentRow: number;
  minContentCol: number;
  minContentRow: number;
  hiddenCols: number[];
  hiddenRows: number[];
  rowHeightOverridesPx: Record<number, number>;
  rowStyleIds: Record<number, number>;
  showGridLines: boolean;
  sparklines: XlsxSparkline[];
  zoomScale: number;
};

type ParseWorkbookStructureOptions = {
  includeCachedFormulaValues?: boolean;
  themePalette?: XlsxThemePalette | null;
};

export type WorkbookTableMetadata = {
  displayName?: string;
  headerRowCount?: number;
  headerRowCellStyle?: string;
  name?: string;
  reference?: string;
  totalsRowCount?: number;
  totalsRowShown?: boolean;
};

export type WorkbookImageAssets = {
  archive: ArchiveEntries;
  formControlsByWorkbookSheetIndex: XlsxFormControl[][];
  imageOriginsById: Map<string, WorkbookImageOrigin>;
  imagesByWorkbookSheetIndex: XlsxImage[][];
  namedCellStyleByName: Record<string, XlsxResolvedCellStyle>;
  objectUrls: string[];
  shapesByWorkbookSheetIndex: XlsxShape[][];
  sheetStatesByWorkbookSheetIndex: Array<WorkbookSheetState | null>;
  sheetOrigins: Array<WorkbookImageSheetOrigin | null>;
  styleById: Record<number, XlsxResolvedCellStyle>;
  tableMetadataByWorkbookSheetIndex: WorkbookTableMetadata[][];
  tableStyleByName: Record<string, XlsxTableStyleDefinition>;
  themePalette: XlsxThemePalette;
};

export type WorkbookStructureAssets = Pick<
  WorkbookImageAssets,
  | "namedCellStyleByName"
  | "sheetStatesByWorkbookSheetIndex"
  | "styleById"
  | "tableMetadataByWorkbookSheetIndex"
  | "tableStyleByName"
  | "themePalette"
>;

export type WorkbookChartStyleAssets = Pick<
  WorkbookImageAssets,
  | "archive"
  | "sheetOrigins"
  | "themePalette"
>;

// ── Sparklines ─────────────────────────────────────────────────────────────

function parseSheetSparklines(
  document: XMLDocument,
  themePalette?: XlsxThemePalette | null
) {
  const sparklines: XlsxSparkline[] = [];

  for (const groupNode of getLocalElements(document, "sparklineGroup")) {
    const rawType = groupNode.getAttribute("type");
    const sparklineType: XlsxSparkline["type"] = rawType === "column"
      ? "column"
      : rawType === "stacked"
        ? "winLoss"
        : "line";

    const markersNode = getFirstChild(groupNode, "markers");
    const negativeNode = getFirstChild(groupNode, "negative");
    const colorSeries = resolveWorkbookColor(parseSpreadsheetColor(getFirstChild(groupNode, "colorSeries")), themePalette);
    const colorNegative = resolveWorkbookColor(parseSpreadsheetColor(getFirstChild(groupNode, "colorNegative")), themePalette);
    const colorMarkers = resolveWorkbookColor(parseSpreadsheetColor(getFirstChild(groupNode, "colorMarkers")), themePalette);
    const colorFirst = resolveWorkbookColor(parseSpreadsheetColor(getFirstChild(groupNode, "colorFirst")), themePalette);
    const colorLast = resolveWorkbookColor(parseSpreadsheetColor(getFirstChild(groupNode, "colorLast")), themePalette);
    const colorHigh = resolveWorkbookColor(parseSpreadsheetColor(getFirstChild(groupNode, "colorHigh")), themePalette);
    const colorLow = resolveWorkbookColor(parseSpreadsheetColor(getFirstChild(groupNode, "colorLow")), themePalette);
    const sparklineCollectionNode = getFirstChild(groupNode, "sparklines");
    if (!sparklineCollectionNode) {
      continue;
    }

    for (const sparklineNode of getChildElements(sparklineCollectionNode, "sparkline")) {
      const formula = getFirstChild(sparklineNode, "f")?.textContent ?? "";
      const targetReference = getFirstChild(sparklineNode, "sqref")?.textContent ?? "";
      const range = parseFormulaRangeReference(formula);
      const target = parseFormulaCellReference(targetReference);
      if (!range || !target) {
        continue;
      }

      sparklines.push({
        color: colorSeries ?? undefined,
        firstColor: colorFirst ?? undefined,
        highColor: colorHigh ?? undefined,
        lastColor: colorLast ?? undefined,
        lowColor: colorLow ?? undefined,
        markerColor: colorMarkers ?? undefined,
        markers: hasEnabledSpreadsheetFlag(markersNode),
        negative: hasEnabledSpreadsheetFlag(negativeNode),
        negativeColor: colorNegative ?? undefined,
        range,
        target,
        type: sparklineType
      });
    }
  }

  return sparklines;
}

// ── Formula reference helpers ──────────────────────────────────────────────

function stripSheetNameFromFormulaReference(reference: string) {
  const trimmed = reference.trim();
  const bangIndex = trimmed.lastIndexOf("!");
  return bangIndex >= 0 ? trimmed.slice(bangIndex + 1) : trimmed;
}

function parseFormulaCellReference(reference: string) {
  const normalized = stripSheetNameFromFormulaReference(reference).split(/\s+/)[0] ?? "";
  return parseA1CellReference(normalized);
}

function parseFormulaRangeReference(reference: string) {
  return parseA1RangeReference(stripSheetNameFromFormulaReference(reference));
}

// ── Conditional formatting ─────────────────────────────────────────────────

function parseSqrefRanges(sqref: string | null | undefined): XlsxCellRange[] {
  if (!sqref) {
    return [];
  }

  return sqref
    .trim()
    .split(/\s+/)
    .flatMap((reference) => {
      const range = parseA1RangeReference(reference);
      return range ? [range] : [];
    });
}

function parseConditionalFormatValueObject(node: Element | null): XlsxConditionalFormatValueObject | null {
  if (!node) {
    return null;
  }

  const type = node.getAttribute("type");
  if (!type) {
    return null;
  }

  const rawValue = node.getAttribute("val") ?? getFirstChild(node, "f")?.textContent ?? undefined;
  const numericValue = rawValue !== undefined ? Number(rawValue) : Number.NaN;
  return {
    type,
    value: Number.isFinite(numericValue) ? numericValue : undefined
  };
}

function parseSpreadsheetBooleanAttribute(node: Element | null, name: string) {
  if (!node) {
    return undefined;
  }

  const value = node.getAttribute(name);
  if (value === null) {
    return undefined;
  }

  return value !== "0" && value !== "false";
}

function parseStandardConditionalFormatRule(
  cfRuleNode: Element,
  ranges: XlsxCellRange[]
): (XlsxConditionalFormatRule & { id?: string }) | null {
  const type = cfRuleNode.getAttribute("type");
  const rawPriority = Number(cfRuleNode.getAttribute("priority") ?? Number.NaN);
  const priority = Number.isFinite(rawPriority) ? rawPriority : Number.MAX_SAFE_INTEGER;

  if (type === "colorScale") {
    const colorScaleNode = getFirstChild(cfRuleNode, "colorScale");
    if (!colorScaleNode) {
      return null;
    }

    const cfvos = getChildElements(colorScaleNode, "cfvo")
      .map((node) => parseConditionalFormatValueObject(node))
      .filter((value): value is XlsxConditionalFormatValueObject => Boolean(value));
    const colors = getChildElements(colorScaleNode, "color")
      .map((node) => parseSpreadsheetColor(node))
      .filter((value): value is Record<string, unknown> => Boolean(value));
    if (cfvos.length === 0 || colors.length === 0) {
      return null;
    }

    return {
      cfvos,
      colors,
      kind: "colorScale",
      priority,
      ranges
    };
  }

  if (type === "dataBar") {
    const dataBarNode = getFirstChild(cfRuleNode, "dataBar");
    if (!dataBarNode) {
      return null;
    }

    const cfvos = getChildElements(dataBarNode, "cfvo")
      .map((node) => parseConditionalFormatValueObject(node))
      .filter((value): value is XlsxConditionalFormatValueObject => Boolean(value));
    if (cfvos.length === 0) {
      return null;
    }

    const extId = getFirstDescendant(cfRuleNode, "id")?.textContent?.trim() || undefined;
    return {
      cfvos,
      color: parseSpreadsheetColor(getFirstChild(dataBarNode, "color")),
      kind: "dataBar",
      priority,
      ranges,
      id: extId
    };
  }

  if (type === "iconSet") {
    const iconSetNode = getFirstChild(cfRuleNode, "iconSet");
    if (!iconSetNode) {
      return null;
    }

    const iconSetName = iconSetNode.getAttribute("iconSet");
    const cfvos = getChildElements(iconSetNode, "cfvo")
      .map((node) => parseConditionalFormatValueObject(node))
      .filter((value): value is XlsxConditionalFormatValueObject => Boolean(value));
    if (!iconSetName || cfvos.length === 0) {
      return null;
    }

    return {
      cfvos,
      icons: cfvos.map((_, index) => ({
        iconId: index,
        iconSet: iconSetName
      })),
      kind: "iconSet",
      priority,
      ranges,
      reverse: parseSpreadsheetBooleanAttribute(iconSetNode, "reverse"),
      showValue: parseSpreadsheetBooleanAttribute(iconSetNode, "showValue")
    };
  }

  return null;
}

function parseExtendedConditionalFormatRule(
  cfRuleNode: Element,
  ranges: XlsxCellRange[]
): (XlsxConditionalFormatRule & { id?: string }) | null {
  const type = cfRuleNode.getAttribute("type");
  const ruleId = cfRuleNode.getAttribute("id") ?? undefined;
  const rawPriority = Number(cfRuleNode.getAttribute("priority") ?? Number.NaN);
  const priority = Number.isFinite(rawPriority) ? rawPriority : Number.MAX_SAFE_INTEGER;

  if (type === "dataBar") {
    const dataBarNode = getFirstChild(cfRuleNode, "dataBar");
    if (!dataBarNode) {
      return null;
    }

    const cfvos = getChildElements(dataBarNode, "cfvo")
      .map((node) => parseConditionalFormatValueObject(node))
      .filter((value): value is XlsxConditionalFormatValueObject => Boolean(value));
    if (cfvos.length === 0) {
      return null;
    }

    return {
      axisColor: parseSpreadsheetColor(getFirstChild(dataBarNode, "axisColor")),
      border: parseSpreadsheetBooleanAttribute(dataBarNode, "border"),
      borderColor: parseSpreadsheetColor(getFirstChild(dataBarNode, "borderColor")),
      cfvos,
      color: parseSpreadsheetColor(getFirstChild(dataBarNode, "fillColor")),
      gradient: parseSpreadsheetBooleanAttribute(dataBarNode, "gradient"),
      kind: "dataBar",
      maxLength: Number(dataBarNode.getAttribute("maxLength") ?? Number.NaN),
      minLength: Number(dataBarNode.getAttribute("minLength") ?? Number.NaN),
      negativeBarBorderColorSameAsPositive: parseSpreadsheetBooleanAttribute(dataBarNode, "negativeBarBorderColorSameAsPositive"),
      negativeBorderColor: parseSpreadsheetColor(getFirstChild(dataBarNode, "negativeBorderColor")),
      negativeFillColor: parseSpreadsheetColor(getFirstChild(dataBarNode, "negativeFillColor")),
      priority,
      ranges,
      showValue: parseSpreadsheetBooleanAttribute(dataBarNode, "showValue"),
      id: ruleId
    };
  }

  if (type === "iconSet") {
    const iconSetNode = getFirstChild(cfRuleNode, "iconSet");
    if (!iconSetNode) {
      return null;
    }

    const cfvos = getChildElements(iconSetNode, "cfvo")
      .map((node) => parseConditionalFormatValueObject(node))
      .filter((value): value is XlsxConditionalFormatValueObject => Boolean(value));
    const icons = getChildElements(iconSetNode, "cfIcon")
      .map((iconNode) => {
        const iconSet = iconNode.getAttribute("iconSet");
        const rawIconId = Number(iconNode.getAttribute("iconId") ?? Number.NaN);
        if (!iconSet || !Number.isFinite(rawIconId)) {
          return null;
        }

        return {
          iconId: rawIconId,
          iconSet
        } satisfies XlsxConditionalFormatIcon;
      })
      .filter((icon): icon is XlsxConditionalFormatIcon => Boolean(icon));

    if (cfvos.length === 0 || icons.length === 0) {
      return null;
    }

    return {
      cfvos,
      icons,
      kind: "iconSet",
      priority,
      ranges,
      reverse: parseSpreadsheetBooleanAttribute(iconSetNode, "reverse"),
      showValue: parseSpreadsheetBooleanAttribute(iconSetNode, "showValue"),
      id: ruleId
    };
  }

  return null;
}

function mergeConditionalFormatRule(
  baseRule: XlsxConditionalFormatRule & { id?: string },
  extendedRule: XlsxConditionalFormatRule & { id?: string }
) {
  if (baseRule.kind !== extendedRule.kind) {
    return baseRule;
  }

  if (baseRule.kind === "colorScale" && extendedRule.kind === "colorScale") {
    return {
      ...baseRule,
      ...extendedRule,
      cfvos: extendedRule.cfvos.length > 0 ? extendedRule.cfvos : baseRule.cfvos,
      colors: extendedRule.colors.length > 0 ? extendedRule.colors : baseRule.colors,
      priority: Number.isFinite(extendedRule.priority) ? extendedRule.priority : baseRule.priority,
      ranges: extendedRule.ranges.length > 0 ? extendedRule.ranges : baseRule.ranges
    };
  }

  if (baseRule.kind === "dataBar" && extendedRule.kind === "dataBar") {
    const merged: XlsxConditionalDataBarRule & { id?: string } = {
      ...baseRule,
      ...extendedRule,
      axisColor: extendedRule.axisColor ?? baseRule.axisColor,
      border: extendedRule.border ?? baseRule.border,
      cfvos: extendedRule.cfvos.length > 0 ? extendedRule.cfvos : baseRule.cfvos,
      color: extendedRule.color ?? baseRule.color,
      negativeBarBorderColorSameAsPositive: extendedRule.negativeBarBorderColorSameAsPositive ?? baseRule.negativeBarBorderColorSameAsPositive,
      negativeBorderColor: extendedRule.negativeBorderColor ?? baseRule.negativeBorderColor,
      negativeFillColor: extendedRule.negativeFillColor ?? baseRule.negativeFillColor,
      priority: Number.isFinite(extendedRule.priority) ? extendedRule.priority : baseRule.priority,
      ranges: extendedRule.ranges.length > 0 ? extendedRule.ranges : baseRule.ranges
    };
    return merged;
  }

  if (baseRule.kind === "iconSet" && extendedRule.kind === "iconSet") {
    const merged: XlsxConditionalIconSetRule & { id?: string } = {
      ...baseRule,
      ...extendedRule,
      cfvos: extendedRule.cfvos.length > 0 ? extendedRule.cfvos : baseRule.cfvos,
      icons: extendedRule.icons.length > 0 ? extendedRule.icons : baseRule.icons,
      priority: Number.isFinite(extendedRule.priority) ? extendedRule.priority : baseRule.priority,
      ranges: extendedRule.ranges.length > 0 ? extendedRule.ranges : baseRule.ranges
    };
    return merged;
  }

  return baseRule;
}

import { SPREADSHEET_NS } from "./image-parser";

function parseConditionalFormatRules(document: Document) {
  const standardRules: Array<XlsxConditionalFormatRule & { id?: string }> = [];
  const extendedRules: Array<XlsxConditionalFormatRule & { id?: string }> = [];

  getLocalElements(document, "conditionalFormatting").forEach((conditionalFormattingNode) => {
    const isExtended = conditionalFormattingNode.namespaceURI !== SPREADSHEET_NS;
    const ranges = isExtended
      ? parseSqrefRanges(getFirstChild(conditionalFormattingNode, "sqref")?.textContent ?? "")
      : parseSqrefRanges(conditionalFormattingNode.getAttribute("sqref"));

    getChildElements(conditionalFormattingNode, "cfRule").forEach((cfRuleNode) => {
      const parsedRule = isExtended
        ? parseExtendedConditionalFormatRule(cfRuleNode, ranges)
        : parseStandardConditionalFormatRule(cfRuleNode, ranges);
      if (parsedRule) {
        if (isExtended) {
          extendedRules.push(parsedRule);
        } else {
          standardRules.push(parsedRule);
        }
      }
    });
  });

  const mergedRules: XlsxConditionalFormatRule[] = [];
  const usedExtendedRuleIds = new Set<string>();
  const extendedRulesById = new Map(
    extendedRules
      .filter((rule) => typeof rule.id === "string" && rule.id.length > 0)
      .map((rule) => [rule.id as string, rule])
  );

  standardRules.forEach((rule) => {
    const matchingExtendedRule = rule.id ? extendedRulesById.get(rule.id) : undefined;
    if (matchingExtendedRule) {
      usedExtendedRuleIds.add(rule.id as string);
      mergedRules.push(mergeConditionalFormatRule(rule, matchingExtendedRule));
      return;
    }

    mergedRules.push(rule);
  });

  extendedRules.forEach((rule) => {
    if (rule.id && usedExtendedRuleIds.has(rule.id)) {
      return;
    }

    mergedRules.push(rule);
  });

  return mergedRules
    .map((rule) => {
      const nextRule = { ...rule } as XlsxConditionalFormatRule & { id?: string };
      delete nextRule.id;
      return nextRule;
    })
    .filter((rule) => rule.ranges.length > 0)
    .sort((left, right) => left.priority - right.priority);
}

// ── Sheet state parsing ────────────────────────────────────────────────────

export function parseSheetState(
  archive: ArchiveEntries,
  path: string,
  options?: ParseWorkbookStructureOptions & {
    defaultFont?: {
      family?: string;
      sizePt?: number;
    } | null;
  }
): WorkbookSheetState | null {
  const xml = readArchiveText(archive, path);
  if (!xml) {
    return null;
  }

  const document = parseXml(xml);
  if (!document) {
    return null;
  }

  const includeCachedFormulaValues = options?.includeCachedFormulaValues ?? true;
  const cachedFormulaValues: Record<string, string> = {};
  const conditionalFormatRules = parseConditionalFormatRules(document);
  const sparklines = parseSheetSparklines(document, options?.themePalette);
  const sheetFormatNode = getLocalElements(document, "sheetFormatPr")[0] ?? null;
  const sheetViewNode = getLocalElements(document, "sheetView")[0] ?? null;
  const rowHeightOverridesPx: Record<number, number> = {};
  const colWidthOverridesPx: Record<number, number> = {};
  const rowStyleIds: Record<number, number> = {};
  const colStyleIds: Record<number, number> = {};
  let minContentCol = Number.POSITIVE_INFINITY;
  let minContentRow = Number.POSITIVE_INFINITY;
  let maxContentCol = -1;
  let maxContentRow = -1;
  const columnWidthCharacterWidthPx = measureColumnCharacterWidthPx(
    options?.defaultFont?.family,
    options?.defaultFont?.sizePt
  );

  const defaultRowHeight = Number(sheetFormatNode?.getAttribute("defaultRowHeight") ?? 15);
  const defaultColWidth = Number(
    sheetFormatNode?.getAttribute("defaultColWidth")
    ?? sheetFormatNode?.getAttribute("baseColWidth")
    ?? 8.43
  );
  const rawZoomScale = Number(
    sheetViewNode?.getAttribute("zoomScale")
    ?? sheetViewNode?.getAttribute("zoomScaleNormal")
    ?? Number.NaN
  );
  const zoomScale = Number.isFinite(rawZoomScale) && rawZoomScale > 0
    ? rawZoomScale
    : 100;
  const trackContentCell = (cellRef: string | null) => {
    if (!cellRef) {
      return;
    }

    const cell = parseA1CellReference(cellRef);
    if (!cell) {
      return;
    }

    minContentCol = Math.min(minContentCol, cell.col);
    minContentRow = Math.min(minContentRow, cell.row);
    maxContentCol = Math.max(maxContentCol, cell.col);
    maxContentRow = Math.max(maxContentRow, cell.row);
  };
  const isMeaningfulCellNode = (cellNode: Element) => {
    if (getFirstChild(cellNode, "f") || getFirstChild(cellNode, "is")) {
      return true;
    }

    const valueNode = getFirstChild(cellNode, "v");
    return Boolean(valueNode && (valueNode.textContent ?? "").length > 0);
  };

  getLocalElements(document, "row").forEach((rowNode) => {
    const rowIndex = Number(rowNode.getAttribute("r") ?? 0) - 1;
    const height = Number(rowNode.getAttribute("ht") ?? Number.NaN);
    const styleId = Number(rowNode.getAttribute("s") ?? Number.NaN);
    if (rowIndex >= 0 && Number.isFinite(height)) {
      rowHeightOverridesPx[rowIndex] = Math.max(MIN_ROW_HEIGHT_PX, Math.round(height * 1.33));
    }
    if (rowIndex >= 0 && Number.isFinite(styleId)) {
      rowStyleIds[rowIndex] = styleId;
    }

    getChildElements(rowNode, "c").forEach((cellNode) => {
      const cellRef = cellNode.getAttribute("r");
      if (isMeaningfulCellNode(cellNode)) {
        trackContentCell(cellRef);
      }

      if (includeCachedFormulaValues) {
        const formulaNode = getFirstChild(cellNode, "f");
        const valueNode = getFirstChild(cellNode, "v");
        if (formulaNode && valueNode && cellRef) {
          cachedFormulaValues[cellRef] = valueNode.textContent ?? "";
        }
      }
    });
  });

  const maxMetadataCol = Math.max(maxContentCol, 0) + 256;
  getLocalElements(document, "col").forEach((colNode) => {
    const min = Number(colNode.getAttribute("min") ?? 0) - 1;
    const max = Number(colNode.getAttribute("max") ?? 0) - 1;
    const width = Number(colNode.getAttribute("width") ?? Number.NaN);
    const styleId = Number(colNode.getAttribute("style") ?? Number.NaN);
    if (!Number.isFinite(width)) {
      if (!Number.isFinite(styleId)) {
        return;
      }
    }

    for (let col = min; col <= Math.min(max, maxMetadataCol); col += 1) {
      if (col >= 0) {
        if (Number.isFinite(width)) {
          const widthPx = sheetColumnWidthToPixels(width, columnWidthCharacterWidthPx);
          colWidthOverridesPx[col] = widthPx;
        }
        if (Number.isFinite(styleId)) {
          colStyleIds[col] = styleId;
        }
      }
    }
  });

  return {
    cachedFormulaValues,
    columnWidthCharacterWidthPx,
    colWidthOverridesPx,
    colStyleIds,
    conditionalFormatRules,
    defaultColWidthPx: sheetColumnWidthToPixels(defaultColWidth, columnWidthCharacterWidthPx),
    defaultRowHeightPx: Math.max(MIN_ROW_HEIGHT_PX, Math.round(defaultRowHeight * 1.33)),
    hasHorizontalMerges: false,
    hasVerticalMerges: false,
    maxHorizontalMergeEndCol: -1,
    maxVerticalMergeEndRow: -1,
    maxContentCol,
    maxContentRow,
    minContentCol: Number.isFinite(minContentCol) ? minContentCol : -1,
    minContentRow: Number.isFinite(minContentRow) ? minContentRow : -1,
    hiddenCols: [],
    hiddenRows: [],
    rowHeightOverridesPx,
    rowStyleIds,
    showGridLines: (sheetViewNode?.getAttribute("showGridLines") ?? "1") !== "0",
    sparklines,
    zoomScale
  };
}

// ── Workbook sheets ────────────────────────────────────────────────────────

export function parseWorkbookSheets(archive: ArchiveEntries) {
  const workbookXml = readArchiveText(archive, "xl/workbook.xml");
  if (!workbookXml) {
    return [];
  }

  const workbookDocument = parseXml(workbookXml);
  if (!workbookDocument) {
    return [];
  }

  const workbookRelationships = parseRelationships(archive, "xl/_rels/workbook.xml.rels", "xl/workbook.xml");
  const sheets: WorkbookSheetInfo[] = [];

  for (const sheetNode of getLocalElements(workbookDocument, "sheet")) {
    const relationshipId = getRelationshipId(sheetNode);
    if (!relationshipId) {
      continue;
    }

    const relationship = workbookRelationships.get(relationshipId);
    if (!relationship) {
      continue;
    }

    sheets.push({
      name: sheetNode.getAttribute("name") ?? `Sheet ${sheets.length + 1}`,
      path: relationship.target
    });
  }

  return sheets;
}

// ── Workbook structure assets ──────────────────────────────────────────────

export function parseWorkbookStructureAssetsFromArchive(
  archive: ArchiveEntries,
  options?: ParseWorkbookStructureOptions
): WorkbookStructureAssets & {
  contentTypes: ContentTypesState;
  theme: ThemeState;
  workbookSheets: WorkbookSheetInfo[];
} {
  const contentTypes = parseContentTypes(archive);
  const workbookSheets = parseWorkbookSheets(archive);
  const theme = parseWorkbookTheme(archive);
  const themePalette = buildThemePalette(theme);
  const { defaultFont, namedCellStyleByName, styleById, tableStyleByName } = parseWorkbookStyles(archive);
  return {
    contentTypes,
    namedCellStyleByName,
    sheetStatesByWorkbookSheetIndex: workbookSheets.map((sheet) => parseSheetState(archive, sheet.path, {
      ...options,
      defaultFont,
      themePalette
    })),
    styleById,
    tableMetadataByWorkbookSheetIndex: workbookSheets.map(() => [] as WorkbookTableMetadata[]),
    tableStyleByName,
    theme,
    themePalette,
    workbookSheets
  };
}
