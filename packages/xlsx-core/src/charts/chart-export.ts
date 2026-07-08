import type { Workbook } from "@dukelib/sheets-wasm";
import { strToU8 } from "fflate";
import type { WorkbookImageAssets, WorkbookImageSheetOrigin } from "../images";
import type {
  XlsxChart,
  XlsxChartAxis,
  XlsxChartDataLabels,
  XlsxChartLegend,
  XlsxChartPointDataLabel,
  XlsxChartPointStyle,
  XlsxChartReference,
  XlsxChartSeries,
  XlsxChartTypeGroup,
  XlsxChartWall,
  XlsxChartsheet,
  XlsxImageAnchor,
  XlsxThemePalette,
  XlsxWorkbookTab,
} from "../types";

import {
  CHART_NS,
  CHART_REL_TYPE,
  CHART_EX_REL_TYPE,
  DRAWINGML_NS,
  DRAWING_SPREADSHEET_NS,
  PKG_REL_NS,
  normalizeChartReference,
  normalizeChartSeries,
  buildChartSeriesFormula,
  resolveReferenceSheet,
} from "./chart-series";

import {
  resolveChartFillColor,
  resolveChartLineStyle,
  normalizeHexColor,
  normalizeWorksheetVisibility,
} from "./chart-colors";

import {
  parseXml,
  serializeXml,
  getLocalChildren,
  getLocalDescendants,
  getFirstLocalChild,
  getFirstLocalDescendant,
  ensureChild,
  setLeafValue,
  setBooleanValue,
  setNumericValue,
  readArchiveText,
  readChartNumericAttribute,
  dirname,
  resolveRelationshipPath,
  removeLocalChildren,
  normalizeArchivePath,
} from "./chart-parser";

import {
  findPrimaryChartTypeNode,
} from "./chart-types";

import type { WorkbookChartOrigin, WorkbookChartAssets } from "./chart-types";

export function normalizeChartAxis(raw: unknown): XlsxChartAxis | null {
  if (!raw || typeof raw !== "object") {
    return null;
  }

  const rawAxis = raw as Record<string, unknown>;
  const axis = rawAxis.axis && typeof rawAxis.axis === "object"
    ? rawAxis.axis as Record<string, unknown>
    : rawAxis;
  const numberFormat = axis.numberFormat && typeof axis.numberFormat === "object"
    ? axis.numberFormat as Record<string, unknown>
    : null;

  return {
    crossId: typeof rawAxis.crossId === "number" && Number.isFinite(rawAxis.crossId) ? rawAxis.crossId : undefined,
    crosses: typeof axis.crosses === "string" ? axis.crosses : undefined,
    crossBetween: typeof axis.crossBetween === "string" ? axis.crossBetween : undefined,
    delete: typeof axis.delete === "boolean" ? axis.delete : undefined,
    id: typeof rawAxis.id === "number" && Number.isFinite(rawAxis.id) ? rawAxis.id : undefined,
    labelPosition: typeof axis.labelPosition === "string" ? axis.labelPosition : undefined,
    logBase: typeof axis.logBase === "number" ? axis.logBase : undefined,
    orientation: typeof axis.orientation === "string" ? axis.orientation : undefined,
    majorUnit: typeof axis.majorUnit === "number" ? axis.majorUnit : undefined,
    max: typeof axis.max === "number" ? axis.max : undefined,
    min: typeof axis.min === "number" ? axis.min : undefined,
    majorGridlines: typeof axis.majorGridlines === "boolean" ? axis.majorGridlines : undefined,
    majorTickMark: typeof axis.majorTickMark === "string" ? axis.majorTickMark : undefined,
    minorUnit: typeof axis.minorUnit === "number" ? axis.minorUnit : undefined,
    minorGridlines: typeof axis.minorGridlines === "boolean" ? axis.minorGridlines : undefined,
    minorTickMark: typeof axis.minorTickMark === "string" ? axis.minorTickMark : undefined,
    numberFormat: numberFormat ? {
      formatCode: typeof numberFormat.formatCode === "string" ? numberFormat.formatCode : undefined,
      sourceLinked: typeof numberFormat.sourceLinked === "boolean" ? numberFormat.sourceLinked : undefined
    } : undefined,
    position: typeof axis.position === "string" ? axis.position : undefined,
    raw: axis,
    shapeProperties: axis.shapeProperties && typeof axis.shapeProperties === "object"
      ? axis.shapeProperties as Record<string, unknown>
      : undefined,
    tickLabelSkip: typeof axis.tickLabelSkip === "number" && Number.isFinite(axis.tickLabelSkip) ? axis.tickLabelSkip : undefined,
    tickMarkSkip: typeof axis.tickMarkSkip === "number" && Number.isFinite(axis.tickMarkSkip) ? axis.tickMarkSkip : undefined
  };
}

export function mergeChartAxis(target: XlsxChartAxis | null | undefined, patch: Partial<XlsxChartAxis> | null | undefined) {
  if (!patch) {
    return target ?? null;
  }
  return {
    ...(target ?? {}),
    ...patch
  };
}

export function readChartAxisFromXml(axisNode: Element | null): Partial<XlsxChartAxis> | null {
  if (!axisNode) {
    return null;
  }

  const numFmt = getFirstLocalChild(axisNode, "numFmt");
  const scalingNode = getFirstLocalChild(axisNode, "scaling");
  return {
    crossId: readChartNumericAttribute(axisNode, "crossAx"),
    crosses: getFirstLocalChild(axisNode, "crosses")?.getAttribute("val") ?? undefined,
    crossBetween: getFirstLocalChild(axisNode, "crossBetween")?.getAttribute("val") ?? undefined,
    delete: getFirstLocalChild(axisNode, "delete")?.getAttribute("val") === "1"
      ? true
      : getFirstLocalChild(axisNode, "delete")?.getAttribute("val") === "0"
        ? false
        : undefined,
    id: readChartNumericAttribute(axisNode, "axId"),
    labelPosition: getFirstLocalChild(axisNode, "tickLblPos")?.getAttribute("val") ?? undefined,
    logBase: readChartNumericAttribute(getFirstLocalChild(axisNode, "scaling"), "logBase"),
    orientation: getFirstLocalChild(scalingNode ?? axisNode, "orientation")?.getAttribute("val") ?? undefined,
    majorGridlines: Boolean(getFirstLocalChild(axisNode, "majorGridlines")),
    majorTickMark: getFirstLocalChild(axisNode, "majorTickMark")?.getAttribute("val") ?? undefined,
    majorUnit: readChartNumericAttribute(axisNode, "majorUnit"),
    max: readChartNumericAttribute(scalingNode, "max"),
    min: readChartNumericAttribute(scalingNode, "min"),
    minorGridlines: Boolean(getFirstLocalChild(axisNode, "minorGridlines")),
    minorTickMark: getFirstLocalChild(axisNode, "minorTickMark")?.getAttribute("val") ?? undefined,
    minorUnit: readChartNumericAttribute(axisNode, "minorUnit"),
    numberFormat: numFmt
      ? {
          formatCode: numFmt.getAttribute("formatCode") ?? undefined,
          sourceLinked: numFmt.getAttribute("sourceLinked") === "1"
            ? true
            : numFmt.getAttribute("sourceLinked") === "0"
              ? false
              : undefined
        }
      : undefined,
    position: getFirstLocalChild(axisNode, "axPos")?.getAttribute("val") ?? undefined,
    tickLabelSkip: readChartNumericAttribute(axisNode, "tickLblSkip"),
    tickMarkSkip: readChartNumericAttribute(axisNode, "tickMarkSkip")
  };
}

export function readChartWallFromXml(wallNode: Element | null, themePalette?: XlsxThemePalette | null): XlsxChartWall | null {
  if (!wallNode) {
    return null;
  }
  const shapeProperties = getFirstLocalChild(wallNode, "spPr");
  const lineStyle = resolveChartLineStyle(shapeProperties, themePalette);
  return {
    fillColor: resolveChartFillColor(shapeProperties, themePalette) ?? undefined,
    hidden: shapeProperties ? getFirstLocalChild(shapeProperties, "noFill") != null : undefined,
    lineColor: lineStyle.color ?? undefined,
    thickness: readChartNumericAttribute(wallNode, "thickness")
  };
}

export function normalizeChartDataLabels(raw: unknown): XlsxChartDataLabels | null {
  if (!raw || typeof raw !== "object") {
    return null;
  }

  const labels = raw as Record<string, unknown>;
  const pointLabels = Array.isArray(labels.pointLabels)
    ? (() => {
        const normalized: XlsxChartPointDataLabel[] = [];
        for (const entry of labels.pointLabels) {
          if (!entry || typeof entry !== "object") {
            continue;
          }
          const point = entry as Record<string, unknown>;
          const index = typeof point.index === "number" && Number.isFinite(point.index)
            ? point.index
            : null;
          if (index == null) {
            continue;
          }

          const nextPoint: XlsxChartPointDataLabel = { index };
          if (typeof point.deleted === "boolean") {
            nextPoint.deleted = point.deleted;
          }
          if (typeof point.fontSizePt === "number" && Number.isFinite(point.fontSizePt)) {
            nextPoint.fontSizePt = point.fontSizePt;
          }
          if (typeof point.showBubbleSize === "boolean") {
            nextPoint.showBubbleSize = point.showBubbleSize;
          }
          if (typeof point.showCategoryName === "boolean") {
            nextPoint.showCategoryName = point.showCategoryName;
          }
          if (typeof point.showPercent === "boolean") {
            nextPoint.showPercent = point.showPercent;
          }
          if (typeof point.showSeriesName === "boolean") {
            nextPoint.showSeriesName = point.showSeriesName;
          }
          if (typeof point.showValue === "boolean") {
            nextPoint.showValue = point.showValue;
          }
          if (typeof point.x === "number" && Number.isFinite(point.x)) {
            nextPoint.x = point.x;
          }
          if (typeof point.y === "number" && Number.isFinite(point.y)) {
            nextPoint.y = point.y;
          }
          normalized.push(nextPoint);
        }
        return normalized;
      })()
    : undefined;
  return {
    pointLabels: pointLabels && pointLabels.length > 0 ? pointLabels : undefined,
    raw: labels,
    showBubbleSize: typeof labels.showBubbleSize === "boolean" ? labels.showBubbleSize : undefined,
    showCategoryName: typeof labels.showCategoryName === "boolean" ? labels.showCategoryName : undefined,
    showLegendKey: typeof labels.showLegendKey === "boolean" ? labels.showLegendKey : undefined,
    showPercent: typeof labels.showPercent === "boolean" ? labels.showPercent : undefined,
    showSeriesName: typeof labels.showSeriesName === "boolean" ? labels.showSeriesName : undefined,
    showValue: typeof labels.showValue === "boolean" ? labels.showValue : undefined
  };
}

export function normalizeChartAnchor(raw: unknown): XlsxImageAnchor {
  if (!raw || typeof raw !== "object") {
    return {
      kind: "two-cell",
      from: { col: 0, colOffsetEmu: 0, row: 0, rowOffsetEmu: 0 },
      to: { col: 8, colOffsetEmu: 0, row: 15, rowOffsetEmu: 0 }
    };
  }

  const anchor = raw as Record<string, unknown>;
  const fromCol = typeof anchor.fromCol === "number" ? anchor.fromCol : 0;
  const fromColOffsetEmu = typeof anchor.fromColOffset === "number" ? anchor.fromColOffset : 0;
  const fromRow = typeof anchor.fromRow === "number" ? anchor.fromRow : 0;
  const fromRowOffsetEmu = typeof anchor.fromRowOffset === "number" ? anchor.fromRowOffset : 0;
  const rawToCol = typeof anchor.toCol === "number" ? anchor.toCol : null;
  const rawToColOffsetEmu = typeof anchor.toColOffset === "number" ? anchor.toColOffset : 0;
  const rawToRow = typeof anchor.toRow === "number" ? anchor.toRow : null;
  const rawToRowOffsetEmu = typeof anchor.toRowOffset === "number" ? anchor.toRowOffset : 0;
  const hasExplicitTo = rawToCol !== null && rawToRow !== null;
  const collapsedWidth = hasExplicitTo && (
    rawToCol < fromCol ||
    (rawToCol === fromCol && rawToColOffsetEmu <= fromColOffsetEmu)
  );
  const collapsedHeight = hasExplicitTo && (
    rawToRow < fromRow ||
    (rawToRow === fromRow && rawToRowOffsetEmu <= fromRowOffsetEmu)
  );
  const fallbackToCol = Math.max(fromCol + 8, 8);
  const fallbackToRow = Math.max(fromRow + 15, 15);

  return {
    kind: "two-cell",
    from: {
      col: fromCol,
      colOffsetEmu: fromColOffsetEmu,
      row: fromRow,
      rowOffsetEmu: fromRowOffsetEmu
    },
    to: {
      col: !hasExplicitTo || collapsedWidth ? fallbackToCol : rawToCol,
      colOffsetEmu: !hasExplicitTo || collapsedWidth ? 0 : rawToColOffsetEmu,
      row: !hasExplicitTo || collapsedHeight ? fallbackToRow : rawToRow,
      rowOffsetEmu: !hasExplicitTo || collapsedHeight ? 0 : rawToRowOffsetEmu
    }
  };
}

export function parseMarkerNode(node: Element | null) {
  if (!node) {
    return null;
  }

  const col = Number(getFirstLocalChild(node, "col")?.textContent ?? Number.NaN);
  const row = Number(getFirstLocalChild(node, "row")?.textContent ?? Number.NaN);
  const colOffsetEmu = Number(getFirstLocalChild(node, "colOff")?.textContent ?? 0);
  const rowOffsetEmu = Number(getFirstLocalChild(node, "rowOff")?.textContent ?? 0);

  if (!Number.isFinite(col) || !Number.isFinite(row)) {
    return null;
  }

  return {
    col: Math.max(0, Math.round(col)),
    colOffsetEmu: Number.isFinite(colOffsetEmu) ? Math.max(0, Math.round(colOffsetEmu)) : 0,
    row: Math.max(0, Math.round(row)),
    rowOffsetEmu: Number.isFinite(rowOffsetEmu) ? Math.max(0, Math.round(rowOffsetEmu)) : 0
  };
}

export function parseChartAnchorNode(anchorNode: Element): XlsxImageAnchor | null {
  if (anchorNode.localName === "twoCellAnchor") {
    const from = parseMarkerNode(getFirstLocalChild(anchorNode, "from"));
    const to = parseMarkerNode(getFirstLocalChild(anchorNode, "to"));
    return from && to ? { from, kind: "two-cell", to } : null;
  }

  if (anchorNode.localName === "oneCellAnchor") {
    const from = parseMarkerNode(getFirstLocalChild(anchorNode, "from"));
    const ext = getFirstLocalChild(anchorNode, "ext");
    const cx = Number(ext?.getAttribute("cx") ?? Number.NaN);
    const cy = Number(ext?.getAttribute("cy") ?? Number.NaN);
    return from && Number.isFinite(cx) && Number.isFinite(cy)
      ? {
          from,
          kind: "one-cell",
          sizeEmu: {
            cx: Math.max(0, Math.round(cx)),
            cy: Math.max(0, Math.round(cy))
          }
        }
      : null;
  }

  const pos = getFirstLocalChild(anchorNode, "pos");
  const ext = getFirstLocalChild(anchorNode, "ext");
  const x = Number(pos?.getAttribute("x") ?? Number.NaN);
  const y = Number(pos?.getAttribute("y") ?? Number.NaN);
  const cx = Number(ext?.getAttribute("cx") ?? Number.NaN);
  const cy = Number(ext?.getAttribute("cy") ?? Number.NaN);
  return Number.isFinite(x) && Number.isFinite(y) && Number.isFinite(cx) && Number.isFinite(cy)
    ? {
        kind: "absolute",
        positionEmu: {
          x: Math.round(x),
          y: Math.round(y)
        },
        sizeEmu: {
          cx: Math.max(0, Math.round(cx)),
          cy: Math.max(0, Math.round(cy))
        }
      }
    : null;
}

export function isCollapsedChartAnchor(anchor: XlsxImageAnchor) {
  if (anchor.kind !== "two-cell") {
    return false;
  }

  const collapsedWidth = anchor.to.col < anchor.from.col
    || (anchor.to.col === anchor.from.col && anchor.to.colOffsetEmu <= anchor.from.colOffsetEmu);
  const collapsedHeight = anchor.to.row < anchor.from.row
    || (anchor.to.row === anchor.from.row && anchor.to.rowOffsetEmu <= anchor.from.rowOffsetEmu);
  return collapsedWidth || collapsedHeight;
}

export function normalizeChartTypeGroup(
  workbook: Workbook,
  workbookSheetIndex: number,
  chartId: string,
  raw: unknown,
  index: number
): XlsxChartTypeGroup | null {
  if (!raw || typeof raw !== "object") {
    return null;
  }

  const group = raw as Record<string, unknown>;
  const rawSeries = Array.isArray(group.series) ? group.series : [];
  return {
    axisIds: Array.isArray(group.axisIds)
      ? group.axisIds.filter((value): value is number => typeof value === "number" && Number.isFinite(value))
      : undefined,
    chartType: typeof group.chartType === "string" ? group.chartType : "ColumnClustered",
    dataLabels: normalizeChartDataLabels(group.dataLabels),
    gapWidth: typeof group.gapWidth === "number" && Number.isFinite(group.gapWidth) ? group.gapWidth : undefined,
    is3d: typeof group.is3d === "boolean" ? group.is3d : undefined,
    overlap: typeof group.overlap === "number" && Number.isFinite(group.overlap) ? group.overlap : undefined,
    raw: group,
    series: rawSeries.map((entry, seriesIndex) => (
      normalizeChartSeries(workbook, workbookSheetIndex, `${chartId}-group-${index}`, entry, seriesIndex)
    )),
    varyColors: typeof group.varyColors === "boolean" ? group.varyColors : undefined
  };
}


export function normalizeChartsheet(raw: unknown, index: number): XlsxChartsheet {
  const chartsheet = raw && typeof raw === "object" ? raw as Record<string, unknown> : {};
  return {
    chartIds: Array.isArray(chartsheet.chartIds) ? chartsheet.chartIds.filter((value): value is string => typeof value === "string") : [],
    chartPath: typeof chartsheet.chartPath === "string" ? chartsheet.chartPath : undefined,
    id: `chartsheet-${index}`,
    index,
    name: typeof chartsheet.name === "string" ? chartsheet.name : `Chart ${index + 1}`,
    raw: chartsheet,
    workbookSheetIndex: typeof chartsheet.workbookSheetIndex === "number" ? chartsheet.workbookSheetIndex : undefined
  };
}

export function buildTabs(
  workbook: Workbook,
  chartsheets: XlsxChartsheet[],
  visibleSheetIndexByWorkbookSheetIndex: Map<number, number>,
  showHiddenSheets = false
): XlsxWorkbookTab[] {
  const rawOrder = Array.isArray(workbook.sheetOrder) ? workbook.sheetOrder as Array<Record<string, unknown>> : [];
  if (rawOrder.length === 0) {
    return workbook.sheetNames.flatMap((name, index) => {
      const worksheet = workbook.getSheet(index);
      const visibility = normalizeWorksheetVisibility(worksheet.visibility);
      if (!showHiddenSheets && visibility !== "visible") {
        return [];
      }

      return [{
        id: `sheet-${index}`,
        index,
        kind: "sheet" as const,
        name,
        sheetIndex: visibleSheetIndexByWorkbookSheetIndex.get(index) ?? index,
        visibility,
        workbookSheetIndex: index
      }];
    });
  }

  return rawOrder.flatMap<XlsxWorkbookTab>((entry, index) => {
    const slotType = typeof entry.slotType === "string" ? entry.slotType : "worksheet";
    const slotIndex = typeof entry.index === "number" ? entry.index : index;
    if (slotType === "chartsheet") {
      const chartsheet = chartsheets[slotIndex];
      return chartsheet ? [{
        chartsheetIndex: slotIndex,
        id: `chartsheet-${slotIndex}`,
        index,
        kind: "chartsheet" as const,
        name: chartsheet.name
      }] : [];
    }

    const worksheet = workbook.getSheet(slotIndex);
    const visibility = normalizeWorksheetVisibility(worksheet.visibility);
    if (!showHiddenSheets && visibility !== "visible") {
      return [];
    }

    return [{
      id: `sheet-${slotIndex}`,
      index,
      kind: "sheet" as const,
      name: worksheet.name,
      sheetIndex: visibleSheetIndexByWorkbookSheetIndex.get(slotIndex) ?? slotIndex,
      visibility,
      workbookSheetIndex: slotIndex
    }];
  });
}

export function collectChartOriginsForSheet(
  archive: Record<string, Uint8Array>,
  origin: WorkbookImageSheetOrigin | null
) {
  if (!origin) {
    return [] as WorkbookChartOrigin[];
  }

  const chartOrigins: WorkbookChartOrigin[] = [];

  for (const attachment of origin.attachments) {
    const drawingXml = readArchiveText(archive, attachment.drawingPath);
    const relsXml = readArchiveText(archive, attachment.drawingRelsPath);
    if (!drawingXml || !relsXml) {
      continue;
    }

    const drawingDocument = parseXml(drawingXml);
    const relsDocument = parseXml(relsXml);
    if (!drawingDocument || !relsDocument) {
      continue;
    }

    const relationships = new Map<string, { target: string; type: string | null }>();
    for (const node of getLocalDescendants(relsDocument, "Relationship")) {
      const id = node.getAttribute("Id");
      const target = node.getAttribute("Target");
      const type = node.getAttribute("Type");
      if (id && target) {
        relationships.set(id, {
          target: resolveRelationshipPath(attachment.drawingRelsPath ?? attachment.drawingPath, target),
          type
        });
      }
    }

    const anchorNodes = Array.from(drawingDocument.documentElement.childNodes).filter(
      (node): node is Element => (
        node.nodeType === Node.ELEMENT_NODE
        && (
          (node as Element).localName === "twoCellAnchor"
          || (node as Element).localName === "oneCellAnchor"
          || (node as Element).localName === "absoluteAnchor"
        )
      )
    );

    let chartAnchorIndex = 0;
    for (const anchorNode of anchorNodes) {
      const graphicFrame = getFirstLocalDescendant(anchorNode, "graphicFrame");
      const chartNode = graphicFrame ? getFirstLocalDescendant(graphicFrame, "chart") : null;
      const relationshipId = chartNode?.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id")
        ?? chartNode?.getAttribute("r:id")
        ?? chartNode?.getAttribute("id");
      if (!relationshipId) {
        continue;
      }
      const relationship = relationships.get(relationshipId);
      if (!relationship || (relationship.type !== CHART_REL_TYPE && relationship.type !== CHART_EX_REL_TYPE)) {
        continue;
      }

      chartOrigins.push({
        anchorIndex: chartAnchorIndex,
        anchor: parseChartAnchorNode(anchorNode),
        chartKind: relationship.type === CHART_EX_REL_TYPE ? "modern" : "classic",
        chartPath: relationship.target,
        drawingPath: attachment.drawingPath,
        workbookSheetIndex: origin.workbookSheetIndex
      });
      chartAnchorIndex += 1;
    }
  }

  return chartOrigins;
}

export function applyChartOrigins(
  chartsByWorkbookSheetIndex: XlsxChart[][],
  chartOriginsById: Map<string, WorkbookChartOrigin>,
  archive: Record<string, Uint8Array>,
  sheetOrigins: Array<WorkbookImageSheetOrigin | null>
) {
  for (let workbookSheetIndex = 0; workbookSheetIndex < chartsByWorkbookSheetIndex.length; workbookSheetIndex += 1) {
    const charts = chartsByWorkbookSheetIndex[workbookSheetIndex] ?? [];
    const origins = collectChartOriginsForSheet(archive, sheetOrigins[workbookSheetIndex] ?? null);
    const originsByKind = {
      classic: origins.filter((origin) => origin.chartKind === "classic"),
      modern: origins.filter((origin) => origin.chartKind === "modern")
    };
    const chartIndexByKind = {
      classic: 0,
      modern: 0
    };

    charts.forEach((chart) => {
      const chartKind = chart.id.startsWith("chart-ex-") ? "modern" : "classic";
      const origin = originsByKind[chartKind][chartIndexByKind[chartKind]];
      chartIndexByKind[chartKind] += 1;
      if (!origin) {
        return;
      }
      if (origin.anchor && isCollapsedChartAnchor(chart.anchor)) {
        chart.anchor = origin.anchor;
      } else if (origin.anchor && chart.anchor.kind === "two-cell" && chart.anchor.from.col === 0 && chart.anchor.from.row === 0) {
        chart.anchor = origin.anchor;
      }
      chart.chartPath = origin.chartPath ?? undefined;
      chartOriginsById.set(chart.id, origin);
    });
  }
}


export function getChartAnchorNodes(drawingDocument: XMLDocument) {
  return Array.from(drawingDocument.documentElement.childNodes).filter(
    (node): node is Element => (
      node.nodeType === Node.ELEMENT_NODE
      && (
        (node as Element).localName === "twoCellAnchor"
        || (node as Element).localName === "oneCellAnchor"
        || (node as Element).localName === "absoluteAnchor"
      )
    )
  )
    .filter((anchorNode) => {
      const graphicFrame = getFirstLocalChild(anchorNode, "graphicFrame");
      return Boolean(graphicFrame && getFirstLocalDescendant(graphicFrame, "chart"));
    });
}

export function updateMarkerNode(markerNode: Element | null, marker: { col: number; colOffsetEmu: number; row: number; rowOffsetEmu: number }) {
  if (!markerNode) {
    return;
  }

  setLeafValue(markerNode, "col", String(Math.max(0, Math.round(marker.col))));
  setLeafValue(markerNode, "colOff", String(Math.max(0, Math.round(marker.colOffsetEmu))));
  setLeafValue(markerNode, "row", String(Math.max(0, Math.round(marker.row))));
  setLeafValue(markerNode, "rowOff", String(Math.max(0, Math.round(marker.rowOffsetEmu))));
}

export function updateAnchorNode(anchorNode: Element, anchor: XlsxImageAnchor) {
  if (anchor.kind === "two-cell") {
    updateMarkerNode(getFirstLocalChild(anchorNode, "from"), anchor.from);
    updateMarkerNode(getFirstLocalChild(anchorNode, "to"), anchor.to);
    return;
  }

  if (anchor.kind === "one-cell") {
    updateMarkerNode(getFirstLocalChild(anchorNode, "from"), anchor.from);
    const ext = getFirstLocalChild(anchorNode, "ext");
    if (ext) {
      ext.setAttribute("cx", String(Math.max(0, Math.round(anchor.sizeEmu.cx))));
      ext.setAttribute("cy", String(Math.max(0, Math.round(anchor.sizeEmu.cy))));
    }
    return;
  }

  const pos = getFirstLocalChild(anchorNode, "pos");
  if (pos) {
    pos.setAttribute("x", String(Math.max(0, Math.round(anchor.positionEmu.x))));
    pos.setAttribute("y", String(Math.max(0, Math.round(anchor.positionEmu.y))));
  }
  const ext = getFirstLocalChild(anchorNode, "ext");
  if (ext) {
    ext.setAttribute("cx", String(Math.max(0, Math.round(anchor.sizeEmu.cx))));
    ext.setAttribute("cy", String(Math.max(0, Math.round(anchor.sizeEmu.cy))));
  }
}

export function setChartTitle(chartNode: Element, value: string | undefined) {
  const existing = getFirstLocalChild(chartNode, "title");
  if (!value) {
    existing?.remove();
    return;
  }

  const titleNode = existing ?? chartNode.insertBefore(
    chartNode.ownerDocument.createElementNS(CHART_NS, "c:title"),
    chartNode.firstChild
  );
  while (titleNode.firstChild) {
    titleNode.removeChild(titleNode.firstChild);
  }
  const tx = titleNode.ownerDocument.createElementNS(CHART_NS, "c:tx");
  const rich = titleNode.ownerDocument.createElementNS(CHART_NS, "c:rich");
  const bodyPr = titleNode.ownerDocument.createElementNS(DRAWINGML_NS, "a:bodyPr");
  const lstStyle = titleNode.ownerDocument.createElementNS(DRAWINGML_NS, "a:lstStyle");
  const p = titleNode.ownerDocument.createElementNS(DRAWINGML_NS, "a:p");
  const r = titleNode.ownerDocument.createElementNS(DRAWINGML_NS, "a:r");
  const t = titleNode.ownerDocument.createElementNS(DRAWINGML_NS, "a:t");
  t.textContent = value;
  r.appendChild(t);
  p.appendChild(r);
  rich.append(bodyPr, lstStyle, p);
  tx.appendChild(rich);
  titleNode.appendChild(tx);
}

export function setRefFormula(parent: Element, refNodeName: string, formula: string | undefined) {
  if (!formula) {
    removeLocalChildren(parent, refNodeName);
    return;
  }

  const refNode = ensureChild(parent, refNodeName);
  setLeafValue(refNode, "f", formula);
}

export function setSeriesText(seriesNode: Element, series: XlsxChartSeries) {
  const raw = series.raw && typeof series.raw === "object" ? series.raw as Record<string, unknown> : null;
  const nameFormula = typeof raw?.name === "string" && raw.name.length > 0 ? raw.name : undefined;
  if (!nameFormula && series.name === undefined) {
    return;
  }

  const tx = ensureChild(seriesNode, "tx");
  removeLocalChildren(tx, "strRef");
  removeLocalChildren(tx, "v");
  if (nameFormula) {
    const strRef = ensureChild(tx, "strRef");
    setLeafValue(strRef, "f", nameFormula);
    return;
  }

  setLeafValue(tx, "v", series.name ?? "");
}

export function updateSeriesNodes(plotAreaNode: Element, chart: Partial<XlsxChart>) {
  if (!chart.series) {
    return;
  }

  const seriesNodes = getLocalDescendants(plotAreaNode, "ser");
  chart.series.forEach((series, index) => {
    const seriesNode = seriesNodes[index];
    if (!seriesNode) {
      return;
    }

    setSeriesText(seriesNode, series);
    if (series.categoriesRef?.formula) {
      const target = getFirstLocalChild(seriesNode, "xVal")
        ?? getFirstLocalChild(seriesNode, "cat")
        ?? (
          chart.chartType === "Scatter" || chart.chartType === "ScatterLines" || chart.chartType === "ScatterSmooth" || chart.chartType === "Bubble"
            ? ensureChild(seriesNode, "xVal")
            : ensureChild(seriesNode, "cat")
        );
      const categoryRefName = target.localName === "xVal" || getFirstLocalChild(target, "numRef") ? "numRef" : "strRef";
      setRefFormula(target, categoryRefName, series.categoriesRef.formula);
    }
    if (series.valuesRef?.formula) {
      const target = getFirstLocalChild(seriesNode, "yVal")
        ?? getFirstLocalChild(seriesNode, "val")
        ?? (
          chart.chartType === "Scatter" || chart.chartType === "ScatterLines" || chart.chartType === "ScatterSmooth" || chart.chartType === "Bubble"
            ? ensureChild(seriesNode, "yVal")
            : ensureChild(seriesNode, "val")
        );
      setRefFormula(target, "numRef", series.valuesRef.formula);
    }
    if (series.bubbleSizeRef) {
      const target = getFirstLocalChild(seriesNode, "bubbleSize") ?? ensureChild(seriesNode, "bubbleSize");
      setRefFormula(target, "numRef", series.bubbleSizeRef.formula);
    }
    if (series.invertIfNegative !== undefined) {
      setBooleanValue(seriesNode, "invertIfNegative", series.invertIfNegative);
    }
    if (series.smooth !== undefined) {
      setBooleanValue(seriesNode, "smooth", series.smooth);
    }
  });
}

export function updateAxisNode(axisNode: Element | null, axis: XlsxChartAxis | null | undefined) {
  if (!axisNode || !axis) {
    return;
  }

  if (axis.position) {
    setLeafValue(ensureChild(axisNode, "axPos"), "val", axis.position);
    getFirstLocalChild(axisNode, "axPos")?.setAttribute("val", axis.position);
  }
  if (axis.majorGridlines !== undefined) {
    const gridlines = getFirstLocalChild(axisNode, "majorGridlines");
    if (axis.majorGridlines && !gridlines) {
      axisNode.appendChild(axisNode.ownerDocument.createElementNS(CHART_NS, "c:majorGridlines"));
    } else if (!axis.majorGridlines) {
      gridlines?.remove();
    }
  }
  if (axis.minorGridlines !== undefined) {
    const gridlines = getFirstLocalChild(axisNode, "minorGridlines");
    if (axis.minorGridlines && !gridlines) {
      axisNode.appendChild(axisNode.ownerDocument.createElementNS(CHART_NS, "c:minorGridlines"));
    } else if (!axis.minorGridlines) {
      gridlines?.remove();
    }
  }
  if (axis.majorTickMark) {
    getFirstLocalChild(axisNode, "majorTickMark")?.setAttribute("val", axis.majorTickMark)
      ?? setBooleanValue(axisNode, "majorTickMark", false).setAttribute("val", axis.majorTickMark);
  }
  if (axis.minorTickMark) {
    getFirstLocalChild(axisNode, "minorTickMark")?.setAttribute("val", axis.minorTickMark)
      ?? setBooleanValue(axisNode, "minorTickMark", false).setAttribute("val", axis.minorTickMark);
  }
  if (axis.labelPosition) {
    getFirstLocalChild(axisNode, "tickLblPos")?.setAttribute("val", axis.labelPosition)
      ?? setBooleanValue(axisNode, "tickLblPos", false).setAttribute("val", axis.labelPosition);
  }
  if (axis.crosses) {
    getFirstLocalChild(axisNode, "crosses")?.setAttribute("val", axis.crosses)
      ?? setBooleanValue(axisNode, "crosses", false).setAttribute("val", axis.crosses);
  }
  if (axis.crossBetween) {
    getFirstLocalChild(axisNode, "crossBetween")?.setAttribute("val", axis.crossBetween)
      ?? setBooleanValue(axisNode, "crossBetween", false).setAttribute("val", axis.crossBetween);
  }
  if (axis.delete !== undefined) {
    setBooleanValue(axisNode, "delete", axis.delete);
  }
  if (axis.numberFormat?.formatCode) {
    const numFmt = ensureChild(axisNode, "numFmt");
    numFmt.setAttribute("formatCode", axis.numberFormat.formatCode);
    if (axis.numberFormat.sourceLinked !== undefined) {
      numFmt.setAttribute("sourceLinked", axis.numberFormat.sourceLinked ? "1" : "0");
    }
  }
}

export function updateDataLabels(chartTypeNode: Element, labels: XlsxChartDataLabels | null | undefined) {
  if (!labels) {
    return;
  }

  const labelsNode = ensureChild(chartTypeNode, "dLbls");
  if (labels.showLegendKey !== undefined) {
    setBooleanValue(labelsNode, "showLegendKey", labels.showLegendKey);
  }
  if (labels.showValue !== undefined) {
    setBooleanValue(labelsNode, "showVal", labels.showValue);
  }
  if (labels.showCategoryName !== undefined) {
    setBooleanValue(labelsNode, "showCatName", labels.showCategoryName);
  }
  if (labels.showSeriesName !== undefined) {
    setBooleanValue(labelsNode, "showSerName", labels.showSeriesName);
  }
  if (labels.showPercent !== undefined) {
    setBooleanValue(labelsNode, "showPercent", labels.showPercent);
  }
  if (labels.showBubbleSize !== undefined) {
    setBooleanValue(labelsNode, "showBubbleSize", labels.showBubbleSize);
  }
}

export function updateWorkbookChartAnchor(
  imageAssets: Pick<WorkbookImageAssets, "archive">,
  chartAssets: WorkbookChartAssets,
  chartId: string,
  anchor: XlsxImageAnchor
) {
  const origin = chartAssets.chartOriginsById.get(chartId);
  if (!origin) {
    return false;
  }

  const drawingXml = readArchiveText(imageAssets.archive, origin.drawingPath);
  if (!drawingXml) {
    return false;
  }

  const drawingDocument = parseXml(drawingXml);
  if (!drawingDocument) {
    return false;
  }

  const anchorNode = getChartAnchorNodes(drawingDocument)[origin.anchorIndex];
  if (!anchorNode) {
    return false;
  }

  updateAnchorNode(anchorNode, anchor);
  imageAssets.archive[normalizeArchivePath(origin.drawingPath)] = strToU8(serializeXml(drawingDocument));
  return true;
}

export function updateWorkbookChartDefinition(
  imageAssets: Pick<WorkbookImageAssets, "archive">,
  chartAssets: WorkbookChartAssets,
  chartId: string,
  patch: Partial<XlsxChart>
) {
  const origin = chartAssets.chartOriginsById.get(chartId);
  if (!origin?.chartPath) {
    return false;
  }

  const chartXml = readArchiveText(imageAssets.archive, origin.chartPath);
  if (!chartXml) {
    return false;
  }

  const chartDocument = parseXml(chartXml);
  if (!chartDocument) {
    return false;
  }

  const chartNode = getFirstLocalDescendant(chartDocument, "chart");
  const plotAreaNode = chartNode ? getFirstLocalChild(chartNode, "plotArea") : null;
  const chartTypeNode = findPrimaryChartTypeNode(plotAreaNode);
  if (!chartNode || !plotAreaNode || !chartTypeNode) {
    return false;
  }

  if (patch.title !== undefined) {
    setChartTitle(chartNode, patch.title);
  }
  if (patch.displayBlanksAs) {
    const node = ensureChild(chartNode, "dispBlanksAs");
    node.setAttribute("val", patch.displayBlanksAs);
  }
  if (patch.roundedCorners !== undefined) {
    setBooleanValue(chartNode, "roundedCorners", patch.roundedCorners);
  }
  if (patch.showDlblsOverMax !== undefined) {
    setBooleanValue(chartNode, "showDLblsOverMax", patch.showDlblsOverMax);
  }
  if (patch.varyColors !== undefined) {
    setBooleanValue(chartTypeNode, "varyColors", patch.varyColors);
  }
  if (patch.gapWidth !== undefined) {
    setNumericValue(chartTypeNode, "gapWidth", patch.gapWidth);
  }
  if (patch.overlap !== undefined) {
    const overlapNode = ensureChild(chartTypeNode, "overlap");
    overlapNode.setAttribute("val", String(Math.round(patch.overlap)));
  }
  if (patch.dataLabels) {
    updateDataLabels(chartTypeNode, patch.dataLabels);
  }
  updateSeriesNodes(plotAreaNode, patch);
  updateAxisNode(
    getLocalChildren(plotAreaNode, "catAx")[0]
      ?? getLocalChildren(plotAreaNode, "dateAx")[0]
      ?? getLocalChildren(plotAreaNode, "serAx")[0]
      ?? null,
    patch.categoryAxis
  );
  updateAxisNode(getLocalChildren(plotAreaNode, "valAx")[0] ?? null, patch.valueAxis);

  imageAssets.archive[normalizeArchivePath(origin.chartPath)] = strToU8(serializeXml(chartDocument));
  return true;
}

