import type {
  XlsxPivotAggregation,
  XlsxPivotDataField,
  XlsxPivotField,
  XlsxPivotFieldAxis,
  XlsxPivotFieldItem,
  XlsxPivotShowAs,
  XlsxPivotSource,
  XlsxPivotTable,
  XlsxSlicer,
  XlsxSlicerItem,
  XlsxTableStyleInfo,
} from "../types";
import {
  type ArchiveEntries,
  getChildElements,
  getFirstChild,
  getLocalElements,
  getRelationshipId,
  isElementNode,
  parseA1RangeReference,
  parseAnchor,
  parseRelationships,
  parseXml,
  readArchiveText,
  relsPathForDocument,
} from "../images";

const PIVOT_TABLE_REL_SUFFIX = "/pivotTable";
const SLICER_REL_SUFFIX = "/slicer";
const DRAWING_REL_SUFFIX = "/drawing";

/**
 * Pivot-cache data needed to label pivot fields, slicers and filters.
 *
 * The workbook references cache definitions from `xl/workbook.xml`; a pivot
 * table only stores the `cacheId`, so field names and member captions have to
 * be resolved through that list.
 */
interface PivotCacheDefinition {
  cacheId: number;
  fields: Array<{
    name: string;
    sharedItems: string[];
  }>;
  refreshOnLoad: boolean;
  source?: XlsxPivotSource;
}

function readPivotCacheDefinitions(archive: ArchiveEntries): Map<number, PivotCacheDefinition> {
  const definitions = new Map<number, PivotCacheDefinition>();
  const workbookXml = readArchiveText(archive, "xl/workbook.xml");
  const workbookDocument = workbookXml ? parseXml(workbookXml) : null;
  if (!workbookDocument) {
    return definitions;
  }
  const relationships = parseRelationships(
    archive,
    "xl/_rels/workbook.xml.rels",
    "xl/workbook.xml",
  );
  for (const cacheNode of getLocalElements(workbookDocument, "pivotCache")) {
    const cacheId = Number(cacheNode.getAttribute("cacheId") ?? Number.NaN);
    if (!Number.isInteger(cacheId) || cacheId < 0) {
      continue;
    }
    const relationshipId = getRelationshipId(cacheNode);
    const relationship = relationshipId ? relationships.get(relationshipId) : undefined;
    const cacheXml = relationship ? readArchiveText(archive, relationship.target) : null;
    const cacheDocument = cacheXml ? parseXml(cacheXml) : null;
    const cacheRoot = cacheDocument?.documentElement;
    if (!cacheRoot) {
      continue;
    }

    const fields = getLocalElements(cacheRoot, "cacheField").map((fieldNode) => ({
      name: fieldNode.getAttribute("name") ?? "",
      sharedItems: parseSharedItems(fieldNode),
    }));
    const worksheetSource = getLocalElements(cacheRoot, "worksheetSource")[0];
    const source: XlsxPivotSource | undefined = worksheetSource
      ? {
          sheet: worksheetSource.getAttribute("sheet") ?? undefined,
          range: worksheetSource.getAttribute("ref") ?? undefined,
          table: worksheetSource.getAttribute("name") ?? null,
        }
      : undefined;

    definitions.set(cacheId, {
      cacheId,
      fields,
      refreshOnLoad: isEnabledFlag(cacheRoot.getAttribute("refreshOnLoad")),
      source,
    });
  }
  return definitions;
}

function parseSharedItems(fieldNode: Element): string[] {
  const sharedItemsNode = getFirstChild(fieldNode, "sharedItems");
  if (!sharedItemsNode) {
    return [];
  }
  const values: string[] = [];
  for (const itemNode of Array.from(sharedItemsNode.childNodes).filter(isElementNode)) {
    const value = itemNode.getAttribute("v") ?? itemNode.textContent ?? "";
    if (itemNode.localName === "m") {
      values.push("(blank)");
    } else {
      values.push(value);
    }
  }
  return values;
}

function isEnabledFlag(value: string | null | undefined) {
  return value === "1" || value?.toLowerCase() === "true";
}

function parsePivotAggregation(value: string | null): XlsxPivotAggregation {
  switch (value) {
    case "count":
      return "count";
    case "countNums":
      return "countNums";
    case "average":
    case "avg":
      return "average";
    case "max":
      return "max";
    case "min":
      return "min";
    case "product":
      return "product";
    case "stdDev":
      return "stdDev";
    case "stdDevp":
      return "stdDevp";
    case "var":
      return "var";
    case "varp":
      return "varp";
    case "sum":
    default:
      return "sum";
  }
}

function parsePivotShowAs(value: string | null): XlsxPivotShowAs {
  switch (value) {
    case "difference":
      return "difference";
    case "index":
      return "index";
    case "percent":
      return "percent";
    case "percentDiff":
      return "percentDifference";
    case "percentOfCol":
      return "percentOfColumnTotal";
    case "percentOfParentCol":
      return "percentOfParentColumnTotal";
    case "percentOfParentRow":
      return "percentOfParentRowTotal";
    case "percentOfRow":
      return "percentOfRowTotal";
    case "percentOfTotal":
      return "percentOfGrandTotal";
    case "percentRunningTotal":
      return "percentRunningTotal";
    case "rankAscending":
      return "rankAscending";
    case "rankDescending":
      return "rankDescending";
    case "runTotal":
      return "runningTotal";
    case "normal":
    default:
      return "normal";
  }
}

function parsePivotFieldItems(
  fieldNode: Element,
  sharedItems: readonly string[],
): XlsxPivotFieldItem[] {
  const itemsNode = getFirstChild(fieldNode, "items");
  if (!itemsNode) {
    return [];
  }
  const items: XlsxPivotFieldItem[] = [];
  for (const itemNode of getChildElements(itemsNode, "item")) {
    const type = itemNode.getAttribute("t");
    if (type === "default" || type === "grand" || type === "blank") {
      continue;
    }
    const rawIndex = itemNode.getAttribute("x");
    const key = rawIndex ?? type ?? "";
    const index = rawIndex === null ? Number.NaN : Number(rawIndex);
    items.push({
      caption: Number.isInteger(index) && index >= 0
        ? sharedItems[index] ?? `Item ${index + 1}`
        : key,
      hidden: isEnabledFlag(itemNode.getAttribute("h")) || undefined,
      key,
    });
  }
  return items;
}

function parsePivotFields(
  pivotRoot: Element,
  cache: PivotCacheDefinition | undefined,
): XlsxPivotField[] {
  return getLocalElements(pivotRoot, "pivotField").map((fieldNode, index) => {
    const axisAttribute = fieldNode.getAttribute("axis");
    let axis: XlsxPivotFieldAxis = "hidden";
    if (axisAttribute === "axisRow") {
      axis = "row";
    } else if (axisAttribute === "axisCol") {
      axis = "column";
    } else if (axisAttribute === "axisPage") {
      axis = "page";
    } else if (fieldNode.getAttribute("dataField") === "1") {
      axis = "data";
    }
    const cacheField = cache?.fields[index];
    return {
      axis,
      index,
      items: parsePivotFieldItems(fieldNode, cacheField?.sharedItems ?? []),
      name: cacheField?.name || `Field ${index + 1}`,
    };
  });
}

function parsePivotAxisFields(pivotRoot: Element, containerName: string): number[] {
  const container = getLocalElements(pivotRoot, containerName)[0];
  if (!container) {
    return [];
  }
  return getChildElements(container, "field")
    .map((fieldNode) => Number(fieldNode.getAttribute("x") ?? Number.NaN))
    .filter((index): index is number => Number.isInteger(index) && index >= 0);
}

function parsePivotDataFields(pivotRoot: Element): XlsxPivotDataField[] {
  return getLocalElements(pivotRoot, "dataField").map((fieldNode, index) => {
    const rawField = Number(fieldNode.getAttribute("fld") ?? Number.NaN);
    const fieldIndex = Number.isInteger(rawField) && rawField >= 0 ? rawField : index;
    const count = Number(fieldNode.getAttribute("numFmtId") ?? Number.NaN);
    return {
      aggregation: parsePivotAggregation(fieldNode.getAttribute("subtotal")),
      fieldIndex,
      name: fieldNode.getAttribute("name") ?? `Value ${index + 1}`,
      ...(Number.isInteger(count) && count >= 0 ? { numberFormatId: count } : {}),
      showAs: parsePivotShowAs(fieldNode.getAttribute("showDataAs")),
    };
  });
}

function parsePivotStyleInfo(pivotRoot: Element): XlsxTableStyleInfo | undefined {
  const styleNode = getLocalElements(pivotRoot, "pivotTableStyleInfo")[0];
  if (!styleNode) {
    return undefined;
  }
  return {
    name: styleNode.getAttribute("name") ?? undefined,
    showColumnStripes: isEnabledFlag(styleNode.getAttribute("showColStripes")),
    showLastColumn: isEnabledFlag(styleNode.getAttribute("showLastColumn")),
    showRowStripes: isEnabledFlag(styleNode.getAttribute("showRowStripes")),
  };
}

function parsePivotLayout(pivotRoot: Element): XlsxPivotTable["layout"] {
  const compact = pivotRoot.getAttribute("compact") !== "0";
  if (compact) {
    return "compact";
  }
  const outline = isEnabledFlag(pivotRoot.getAttribute("outline"))
    || isEnabledFlag(pivotRoot.getAttribute("outlineData"));
  const gridDropZones = isEnabledFlag(pivotRoot.getAttribute("gridDropZones"));
  return outline || !gridDropZones ? "outline" : "tabular";
}

/** Reads every pivot table part a worksheet references. */
export function parseWorksheetPivotTables(
  archive: ArchiveEntries,
  sheetPath: string,
): XlsxPivotTable[] {
  const relationships = parseRelationships(
    archive,
    relsPathForDocument(sheetPath),
    sheetPath,
  );
  const caches = readPivotCacheDefinitions(archive);
  const pivotTables: XlsxPivotTable[] = [];
  const seenParts = new Set<string>();

  for (const relationship of relationships.values()) {
    if (!relationship.type.endsWith(PIVOT_TABLE_REL_SUFFIX) || seenParts.has(relationship.target)) {
      continue;
    }
    seenParts.add(relationship.target);
    const xml = readArchiveText(archive, relationship.target);
    const document = xml ? parseXml(xml) : null;
    const pivotRoot = document?.documentElement;
    if (!pivotRoot) {
      continue;
    }
    const locationNode = getLocalElements(pivotRoot, "location")[0];
    const range = locationNode
      ? parseA1RangeReference(locationNode.getAttribute("ref") ?? "")
      : null;
    if (!range) {
      continue;
    }
    const rawCacheId = Number(pivotRoot.getAttribute("cacheId") ?? Number.NaN);
    const cacheId = Number.isInteger(rawCacheId) && rawCacheId >= 0 ? rawCacheId : 0;
    const cache = caches.get(cacheId);
    const styleInfo = parsePivotStyleInfo(pivotRoot);
    pivotTables.push({
      cacheId,
      columnFields: parsePivotAxisFields(pivotRoot, "colFields"),
      dataCaption: pivotRoot.getAttribute("dataCaption") ?? undefined,
      dataFields: parsePivotDataFields(pivotRoot),
      fields: parsePivotFields(pivotRoot, cache),
      layout: parsePivotLayout(pivotRoot),
      name: pivotRoot.getAttribute("name") ?? `PivotTable${pivotTables.length + 1}`,
      pageFields: parsePivotAxisFields(pivotRoot, "pageFields"),
      range,
      refreshOnLoad: cache?.refreshOnLoad ?? false,
      rowFields: parsePivotAxisFields(pivotRoot, "rowFields"),
      ...(cache?.source ? { source: cache.source } : {}),
      ...(styleInfo ? { styleInfo } : {}),
    });
  }
  return pivotTables;
}

interface SlicerCacheDefinition {
  items: Array<{ key: string; selected: boolean }>;
  name: string;
  pivotCacheId?: number;
  pivotTables: string[];
  sourceName?: string;
}

function readSlicerCacheDefinitions(archive: ArchiveEntries): Map<string, SlicerCacheDefinition> {
  const definitions = new Map<string, SlicerCacheDefinition>();
  for (const path of Object.keys(archive)) {
    if (!/^xl\/slicerCaches\/[^/]+\.xml$/i.test(path)) {
      continue;
    }
    const xml = readArchiveText(archive, path);
    const document = xml ? parseXml(xml) : null;
    const root = document?.documentElement;
    if (!root) {
      continue;
    }
    const name = root.getAttribute("name") ?? "";
    if (!name) {
      continue;
    }
    const items: Array<{ key: string; selected: boolean }> = [];
    const dataNode = getLocalElements(root, "data")[0];
    const containerNode = dataNode
      ? getLocalElements(dataNode, "tabular")[0] ?? getLocalElements(dataNode, "olap")[0]
      : undefined;
    if (containerNode) {
      for (const itemNode of getLocalElements(containerNode, "i")) {
        const key = itemNode.getAttribute("x") ?? "";
        items.push({ key, selected: isEnabledFlag(itemNode.getAttribute("s")) });
      }
    }
    const pivotCacheId = containerNode
      ? Number(containerNode.getAttribute("pivotCacheId") ?? Number.NaN)
      : Number.NaN;
    const pivotTables = getLocalElements(root, "pivotTable")
      .map((node) => node.getAttribute("name") ?? "")
      .filter(Boolean);
    definitions.set(name.toLowerCase(), {
      items,
      name,
      ...(Number.isInteger(pivotCacheId) && pivotCacheId >= 0 ? { pivotCacheId } : {}),
      pivotTables,
      sourceName: root.getAttribute("sourceName") ?? undefined,
    });
  }
  return definitions;
}

function slicerItemsFromCache(
  cache: SlicerCacheDefinition | undefined,
  caches: Map<number, PivotCacheDefinition>,
): XlsxSlicerItem[] {
  if (!cache) {
    return [];
  }
  const pivotCache = cache.pivotCacheId === undefined ? undefined : caches.get(cache.pivotCacheId);
  const sourceField = pivotCache?.fields.find((field) =>
    cache.sourceName ? field.name === cache.sourceName : false
  );
  return cache.items.map((item, index) => {
    const numericIndex = Number(item.key);
    const caption = Number.isInteger(numericIndex) && numericIndex >= 0
      ? sourceField?.sharedItems[numericIndex] ?? `Item ${numericIndex + 1}`
      : item.key || `Item ${index + 1}`;
    return {
      caption,
      key: item.key || String(index),
      selected: item.selected,
    };
  });
}

function findSlicerAnchors(archive: ArchiveEntries, sheetPath: string) {
  const anchors = new Map<string, ReturnType<typeof parseAnchor>>();
  const relationships = parseRelationships(
    archive,
    relsPathForDocument(sheetPath),
    sheetPath,
  );
  for (const relationship of relationships.values()) {
    if (!relationship.type.endsWith(DRAWING_REL_SUFFIX)) {
      continue;
    }
    const xml = readArchiveText(archive, relationship.target);
    const document = xml ? parseXml(xml) : null;
    if (!document) {
      continue;
    }
    for (const localName of ["twoCellAnchor", "oneCellAnchor", "absoluteAnchor"] as const) {
      for (const anchorNode of getLocalElements(document, localName)) {
        const name = getLocalElements(anchorNode, "cNvPr")
          .map((node) => node.getAttribute("name") ?? "")
          .find((value) => value.toLowerCase().startsWith("slicer_"))
          ?? getLocalElements(anchorNode, "slicer")
            .map((node) => node.getAttribute("name") ?? "")
            .find(Boolean);
        if (!name) {
          continue;
        }
        const slicerName = name.replace(/^slicer_/i, "");
        const parsed = parseAnchor(anchorNode);
        if (parsed && !anchors.has(slicerName.toLowerCase())) {
          anchors.set(slicerName.toLowerCase(), parsed);
        }
      }
    }
  }
  return anchors;
}

/** Reads every slicer part a worksheet references, with cache items and anchors. */
export function parseWorksheetSlicers(
  archive: ArchiveEntries,
  sheetPath: string,
  workbookSheetIndex: number,
): XlsxSlicer[] {
  const relationships = parseRelationships(
    archive,
    relsPathForDocument(sheetPath),
    sheetPath,
  );
  const slicerCaches = readSlicerCacheDefinitions(archive);
  const pivotCaches = readPivotCacheDefinitions(archive);
  const anchors = findSlicerAnchors(archive, sheetPath);
  const slicers: XlsxSlicer[] = [];
  const seenParts = new Set<string>();

  for (const relationship of relationships.values()) {
    if (!relationship.type.endsWith(SLICER_REL_SUFFIX) || seenParts.has(relationship.target)) {
      continue;
    }
    seenParts.add(relationship.target);
    const xml = readArchiveText(archive, relationship.target);
    const document = xml ? parseXml(xml) : null;
    if (!document) {
      continue;
    }
    for (const slicerNode of getLocalElements(document, "slicer")) {
      const name = slicerNode.getAttribute("name") ?? "";
      if (!name) {
        continue;
      }
      const cacheName = slicerNode.getAttribute("cache") ?? undefined;
      const cache = cacheName ? slicerCaches.get(cacheName.toLowerCase()) : undefined;
      const anchor = anchors.get(name.toLowerCase())
        ?? (cacheName ? anchors.get(cacheName.replace(/^slicer_/i, "").toLowerCase()) : undefined);
      const rawColumnCount = Number(slicerNode.getAttribute("columnCount") ?? Number.NaN);
      slicers.push({
        ...(anchor ? { anchor } : {}),
        caption: slicerNode.getAttribute("caption") ?? undefined,
        cacheName,
        ...(Number.isInteger(rawColumnCount) && rawColumnCount > 0 ? { columnCount: rawColumnCount } : {}),
        items: slicerItemsFromCache(cache, pivotCaches),
        name,
        pivotTables: cache?.pivotTables ?? [],
        sourceName: cache?.sourceName,
        style: slicerNode.getAttribute("style") ?? undefined,
        workbookSheetIndex,
      });
    }
  }
  return slicers;
}

/** Slicer parts that are not referenced by the sheet but live in the archive. */
export function listPivotParts(archive: ArchiveEntries) {
  return Object.keys(archive).filter((path) =>
    /^xl\/(pivotCache|pivotTables|slicers|slicerCaches)\//i.test(path)
  );
}
