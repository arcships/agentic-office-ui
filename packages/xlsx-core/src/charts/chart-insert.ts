import { strToU8 } from "fflate";
import type { XlsxCellRange, XlsxChartFamily, XlsxImageAnchor } from "../types";
import { columnLabel } from "../core";
import {
  CONTENT_TYPES_NS,
  DRAWING_CONTENT_TYPE,
  DRAWING_REL_TYPE,
  PKG_REL_NS,
  SPREADSHEET_NS,
  type ArchiveEntries,
  getLocalElements,
  parseXml,
  readArchiveText,
  relativeArchivePath,
  relsPathForDocument,
  serializeXml,
} from "../images";
import {
  appendWorkbookSheetDrawingAttachment,
  ensureContentTypeOverride,
  escapeXmlText,
  insertSheetChild,
  type WorkbookAuthoringAssets,
} from "../authoring/sheet-authoring";

const CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const SPREADSHEET_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const CHART_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

const DEFAULT_CHART_WIDTH_EMU = 5_486_400;
const DEFAULT_CHART_HEIGHT_EMU = 3_200_400;
const CHART_CONTENT_TYPE_PREFIX = "application/vnd.openxmlformats-officedocument.drawingml.";

export type { XlsxChartFamily } from "../types";

export type XlsxChartInsertSpec = {
  /** Row/column range that supplies categories and values. */
  range: XlsxCellRange;
  /** First row of `range` holds series names. Defaults to `true`. */
  hasHeaderRow?: boolean;
  /** First column of `range` holds category labels. Defaults to `true`. */
  hasHeaderColumn?: boolean;
  /** Chart family to generate. Bar/column, line and pie are supported. */
  family?: XlsxChartFamily;
  /** Optional anchor; defaults to one cell below the source range. */
  anchor?: XlsxImageAnchor;
  sheetName: string;
  title?: string;
};

export type XlsxChartInsertResult = {
  chartPath: string;
  /** OOXML chart type written into the new part, for example `ColumnClustered`. */
  chartType: string;
  drawingPath: string;
  drawingRelsPath: string;
};

function normalizeRange(range: XlsxCellRange): XlsxCellRange {
  const startRow = Math.max(0, Math.min(range.start.row, range.end.row));
  const startCol = Math.max(0, Math.min(range.start.col, range.end.col));
  const endRow = Math.max(0, Math.max(range.start.row, range.end.row));
  const endCol = Math.max(0, Math.max(range.start.col, range.end.col));
  return {
    start: { col: startCol, row: startRow },
    end: { col: endCol, row: endRow }
  };
}

function cellRef(row: number, col: number, absolute = true) {
  const column = columnLabel(col);
  return absolute ? `$${column}$${row + 1}` : `${column}${row + 1}`;
}

function quoteSheetName(name: string) {
  const trimmed = name.trim() || "Sheet1";
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(trimmed) && !/^\d/.test(trimmed)
    ? trimmed
    : `'${trimmed.replace(/'/g, "''")}'`;
}

function rangeFormula(sheetName: string, startRow: number, startCol: number, endRow: number, endCol: number) {
  const start = cellRef(startRow, startCol);
  const end = cellRef(endRow, endCol);
  return `${quoteSheetName(sheetName)}!${start === end ? start : `${start}:${end}`}`;
}

export function resolveChartTypeForFamily(family: XlsxChartFamily) {
  return family === "line" ? "Line" : family === "pie" ? "Pie" : "ColumnClustered";
}

type ChartSeriesInput = {
  categoriesFormula: string | null;
  nameFormula: string | null;
  name: string;
  valuesFormula: string;
};

function chartTitleXml(title: string | undefined) {
  if (!title?.trim()) {
    return "";
  }
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeXmlText(title.trim())}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
}

function seriesXml(series: ChartSeriesInput, index: number, family: XlsxChartFamily) {
  const nameXml = series.nameFormula
    ? `<c:tx><c:strRef><c:f>${escapeXmlText(series.nameFormula)}</c:f></c:strRef></c:tx>`
    : `<c:tx><c:v>${escapeXmlText(series.name)}</c:v></c:tx>`;
  const categoriesXml = series.categoriesFormula
    ? `<c:cat><c:strRef><c:f>${escapeXmlText(series.categoriesFormula)}</c:f></c:strRef></c:cat>`
    : "";
  const markerXml = family === "line" ? "<c:marker><c:symbol val=\"none\"/></c:marker>" : "";
  const smoothXml = family === "line" ? "<c:smooth val=\"0\"/>" : "";
  return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>${nameXml}${markerXml}${categoriesXml}<c:val><c:numRef><c:f>${escapeXmlText(series.valuesFormula)}</c:f></c:numRef></c:val>${smoothXml}</c:ser>`;
}

function buildChartXml(options: {
  family: XlsxChartFamily;
  series: ChartSeriesInput[];
  title?: string;
}) {
  const axis1 = 50000001;
  const axis2 = 50000002;
  const seriesNodes = options.series
    .map((series, index) => seriesXml(series, index, options.family))
    .join("");
  const titleNode = chartTitleXml(options.title);
  const legendNode = options.series.length > 0
    ? `<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>`
    : "";
  const plotArea = options.family === "pie"
    ? `<c:plotArea><c:layout/><c:pieChart><c:varyColors val="1"/>${seriesNodes}<c:firstSliceAngle val="0"/></c:pieChart></c:plotArea>`
    : options.family === "line"
      ? `<c:plotArea><c:layout/><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${seriesNodes}<c:marker val="1"/><c:axId val="${axis1}"/><c:axId val="${axis2}"/></c:lineChart>${categoryAxisXml(axis1, axis2)}${valueAxisXml(axis2, axis1)}</c:plotArea>`
      : `<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${seriesNodes}<c:gapWidth val="150"/><c:axId val="${axis1}"/><c:axId val="${axis2}"/></c:barChart>${categoryAxisXml(axis1, axis2)}${valueAxisXml(axis2, axis1)}</c:plotArea>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${DRAWINGML_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0"/><c:chart>${titleNode}<c:autoTitleDeleted val="${options.title?.trim() ? 0 : 1}"/>${plotArea}${legendNode}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
}

function categoryAxisXml(axisId: number, crossAxisId: number) {
  return `<c:catAx><c:axId val="${axisId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="${crossAxisId}"/></c:catAx>`;
}

function valueAxisXml(axisId: number, crossAxisId: number) {
  return `<c:valAx><c:axId val="${axisId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="${crossAxisId}"/></c:valAx>`;
}

function buildDrawingXml(options: {
  chartRelationshipId: string;
  chartTitle: string;
  fromCol: number;
  fromRow: number;
  shapeId: number;
}) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="${SPREADSHEET_DRAWING_NS}" xmlns:a="${DRAWINGML_NS}" xmlns:c="${CHART_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:oneCellAnchor><xdr:from><xdr:col>${options.fromCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${options.fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${DEFAULT_CHART_WIDTH_EMU}" cy="${DEFAULT_CHART_HEIGHT_EMU}"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${options.shapeId}" name="${escapeXmlText(options.chartTitle)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${CHART_NS}"><c:chart r:id="${options.chartRelationshipId}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`;
}

function buildDrawingRelsXml(chartPath: string, drawingPath: string) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId1" Type="${CHART_REL_TYPE}" Target="${escapeXmlText(relativeArchivePath(drawingPath, chartPath))}"/></Relationships>`;
}

function ensureDrawingNamespace(document: XMLDocument) {
  const root = document.documentElement;
  if (!root || root.getAttribute("xmlns:r")) {
    return;
  }
  root.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
}

function nextPartPath(archive: ArchiveEntries, directory: string, baseName: string, extension: string) {
  let index = 1;
  while (archive[`${directory}/${baseName}${index}.${extension}`]) {
    index += 1;
  }
  return `${directory}/${baseName}${index}.${extension}`;
}

/** Resolve a relationship target against the part that owns it. */
function resolvePartPath(ownerPath: string, target: string) {
  if (target.startsWith("/")) {
    return target.replace(/^\/+/, "");
  }
  const segments = ownerPath.split("/").slice(0, -1);
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
  return segments.join("/");
}

function nextShapeId(archive: ArchiveEntries, sheetPath: string) {
  const relationshipsXml = readArchiveText(archive, relsPathForDocument(sheetPath));
  const relationshipsDocument = relationshipsXml ? parseXml(relationshipsXml) : null;
  const usedIds: number[] = [];
  if (relationshipsDocument) {
    for (const node of getLocalElements(relationshipsDocument, "Relationship")) {
      if (node.getAttribute("Type") !== DRAWING_REL_TYPE) {
        continue;
      }
      const target = node.getAttribute("Target") ?? "";
      const drawingPath = resolvePartPath(sheetPath, target);
      const drawingXml = readArchiveText(archive, drawingPath);
      const drawingDocument = drawingXml ? parseXml(drawingXml) : null;
      if (!drawingDocument) {
        continue;
      }
      for (const node of getLocalElements(drawingDocument, "cNvPr")) {
        const id = Number(node.getAttribute("id") ?? Number.NaN);
        if (Number.isFinite(id)) {
          usedIds.push(id);
        }
      }
    }
  }
  return Math.max(1, ...usedIds, 1) + 1;
}

/**
 * Inserts a chart built from a cell range into a worksheet.
 *
 * The chart part, its drawing part and the drawing relationships are written to
 * the source archive and registered as a sheet drawing attachment, so the next
 * save/export carries the chart and the viewer can render it immediately after
 * re-parsing chart assets.
 *
 * @returns The created part paths, or `null` when the range is empty, the
 * source archive has no worksheet part, or the sheet has no addressable name.
 */
export function insertWorkbookChart(
  assets: WorkbookAuthoringAssets,
  workbookSheetIndex: number,
  spec: XlsxChartInsertSpec
): XlsxChartInsertResult | null {
  const sheetPath = assets.sheetPathsByWorkbookSheetIndex[workbookSheetIndex];
  if (!sheetPath || !assets.archive[sheetPath]) {
    return null;
  }
  const sheetXml = readArchiveText(assets.archive, sheetPath);
  const sheetDocument = sheetXml ? parseXml(sheetXml) : null;
  if (!sheetDocument?.documentElement) {
    return null;
  }
  const family: XlsxChartFamily = spec.family === "line" || spec.family === "pie" ? spec.family : "bar";
  const range = normalizeRange(spec.range);
  const hasHeaderRow = spec.hasHeaderRow !== false;
  const hasHeaderColumn = spec.hasHeaderColumn !== false;
  const valueStartRow = range.start.row + (hasHeaderRow ? 1 : 0);
  const valueStartCol = range.start.col + (hasHeaderColumn ? 1 : 0);
  if (valueStartRow > range.end.row || valueStartCol > range.end.col) {
    return null;
  }
  const categoriesFormula = hasHeaderColumn
    ? rangeFormula(spec.sheetName, valueStartRow, range.start.col, range.end.row, range.start.col)
    : null;
  const series: ChartSeriesInput[] = [];
  for (let col = valueStartCol; col <= range.end.col; col += 1) {
    series.push({
      categoriesFormula,
      name: `${columnLabel(col)} — ${spec.sheetName}`,
      nameFormula: hasHeaderRow
        ? rangeFormula(spec.sheetName, range.start.row, col, range.start.row, col)
        : null,
      valuesFormula: rangeFormula(spec.sheetName, valueStartRow, col, range.end.row, col)
    });
  }
  if (series.length === 0) {
    return null;
  }

  const chartTitle = spec.title?.trim() || "";
  const chartPath = nextPartPath(assets.archive, "xl/charts", "chartAuthoring", "xml");
  const drawingPath = nextPartPath(assets.archive, "xl/drawings", "drawingAuthoring", "xml");
  const drawingRelsPath = relsPathForDocument(drawingPath);
  const shapeId = nextShapeId(assets.archive, sheetPath);
  const displayName = chartTitle || `Chart ${shapeId}`;

  assets.archive[chartPath] = strToU8(buildChartXml({
    family,
    series,
    title: chartTitle || undefined
  }));
  assets.archive[drawingPath] = strToU8(buildDrawingXml({
    chartRelationshipId: "rId1",
    chartTitle: displayName,
    fromCol: Math.min(range.start.col, 16_000),
    fromRow: Math.min(range.end.row + 2, 1_000_000),
    shapeId
  }));
  assets.archive[drawingRelsPath] = strToU8(buildDrawingRelsXml(chartPath, drawingPath));
  ensureContentTypeOverride(assets.archive, chartPath, `${CHART_CONTENT_TYPE_PREFIX}chart+xml`);
  ensureContentTypeOverride(assets.archive, drawingPath, DRAWING_CONTENT_TYPE);

  appendWorkbookSheetDrawingAttachment(assets, workbookSheetIndex, {
    drawingPath,
    drawingRelsPath,
    mediaPaths: []
  });

  // Keep the overlay worksheet pointing at its drawing so the archive stays
  // self-consistent for debugging and for hosts that reload it directly.
  ensureDrawingNamespace(sheetDocument);
  const existingDrawing = getLocalElements(sheetDocument, "drawing")[0];
  if (!existingDrawing) {
    const relationshipsPath = relsPathForDocument(sheetPath);
    const relationshipsXml = readArchiveText(assets.archive, relationshipsPath);
    const relationshipsDocument = relationshipsXml
      ? parseXml(relationshipsXml)
      : parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}"></Relationships>`);
    if (relationshipsDocument?.documentElement) {
      const relationshipsRoot = relationshipsDocument.documentElement;
      const existingIds = new Set(
        getLocalElements(relationshipsDocument, "Relationship")
          .map((node) => node.getAttribute("Id"))
          .filter((value): value is string => Boolean(value))
      );
      let drawingId = 1;
      while (existingIds.has(`rIdXlsxChartDrawing${drawingId}`)) {
        drawingId += 1;
      }
      const relationshipId = `rIdXlsxChartDrawing${drawingId}`;
      const relationshipNode = relationshipsDocument.createElementNS(PKG_REL_NS, "Relationship");
      relationshipNode.setAttribute("Id", relationshipId);
      relationshipNode.setAttribute("Type", DRAWING_REL_TYPE);
      relationshipNode.setAttribute("Target", relativeArchivePath(sheetPath, drawingPath));
      relationshipsRoot.appendChild(relationshipNode);
      assets.archive[relationshipsPath] = strToU8(serializeXml(relationshipsDocument));

      const drawingNode = sheetDocument.createElementNS(SPREADSHEET_NS, "drawing");
      drawingNode.setAttributeNS(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "r:id",
        relationshipId
      );
      insertSheetChild(sheetDocument, drawingNode);
      assets.archive[sheetPath] = strToU8(serializeXml(sheetDocument));
    }
  }

  return {
    chartPath,
    chartType: resolveChartTypeForFamily(family),
    drawingPath,
    drawingRelsPath
  };
}

export { CHART_CONTENT_TYPE_PREFIX, CHART_NS, CHART_REL_TYPE, CONTENT_TYPES_NS };