import type { XlsxImageAnchor, XlsxImageRect } from "./image-types";

// ── Chart reference / data labels ────────────────────────────────────────

export interface XlsxChartReference {
  formula?: string;
  refType?: string;
  values?: Array<number | string | null>;
}

export interface XlsxChartDataLabels {
  pointLabels?: XlsxChartPointDataLabel[];
  raw?: Record<string, unknown>;
  showBubbleSize?: boolean;
  showCategoryName?: boolean;
  showLegendKey?: boolean;
  showPercent?: boolean;
  showSeriesName?: boolean;
  showValue?: boolean;
}

export interface XlsxChartPointDataLabel {
  deleted?: boolean;
  fontSizePt?: number;
  index: number;
  showBubbleSize?: boolean;
  showCategoryName?: boolean;
  showPercent?: boolean;
  showSeriesName?: boolean;
  showValue?: boolean;
  x?: number;
  y?: number;
}

// ── Legend ────────────────────────────────────────────────────────────────

export interface XlsxChartLegend {
  overlay?: boolean;
  position?: string;
  raw?: Record<string, unknown>;
}

// ── Axis ──────────────────────────────────────────────────────────────────

export interface XlsxChartAxis {
  crossId?: number;
  crosses?: string;
  crossBetween?: string;
  delete?: boolean;
  id?: number;
  labelPosition?: string;
  logBase?: number;
  orientation?: string;
  majorUnit?: number;
  max?: number;
  min?: number;
  majorGridlines?: boolean;
  majorTickMark?: string;
  minorUnit?: number;
  minorGridlines?: boolean;
  minorTickMark?: string;
  numberFormat?: {
    formatCode?: string;
    sourceLinked?: boolean;
  };
  position?: string;
  raw?: Record<string, unknown>;
  shapeProperties?: Record<string, unknown>;
  tickLabelSkip?: number;
  tickMarkSkip?: number;
}

// ── Point style ──────────────────────────────────────────────────────────

export interface XlsxChartPointStyle {
  color?: string;
  explosion?: number;
  index: number;
  lineColor?: string;
}

// ── Series ────────────────────────────────────────────────────────────────

export interface XlsxChartSeries {
  bubbleSizeRef?: XlsxChartReference | null;
  bubbleSizes?: Array<number | null>;
  categories: Array<number | string | null>;
  categoriesRef?: XlsxChartReference | null;
  color?: string;
  dataPoints: unknown[];
  dataPointStyles?: XlsxChartPointStyle[];
  formatIdx?: number;
  hidden?: boolean;
  id: string;
  invertIfNegative?: boolean;
  lineColor?: string;
  lineWidthPx?: number;
  marker?: Record<string, unknown>;
  markerColor?: string;
  markerLineColor?: string;
  markerSize?: number;
  markerSymbol?: string;
  name?: string;
  negativeColor?: string;
  negativeLineColor?: string;
  raw?: Record<string, unknown>;
  shapeProperties?: Record<string, unknown>;
  smooth?: boolean;
  values: Array<number | null>;
  valuesRef?: XlsxChartReference | null;
}

// ── Element selection / type group / wall ────────────────────────────────

export type XlsxChartElementSelection =
  | { kind: "chart"; chartId: string }
  | { kind: "series"; chartId: string; seriesId: string; seriesIndex: number }
  | { kind: "point"; chartId: string; seriesId: string; seriesIndex: number; pointIndex: number }
  | { kind: "legendEntry"; chartId: string; seriesId: string; seriesIndex: number };

export interface XlsxChartTypeGroup {
  axisIds?: number[];
  chartType: string;
  dataLabels?: XlsxChartDataLabels | null;
  gapWidth?: number;
  is3d?: boolean;
  overlap?: number;
  raw?: Record<string, unknown>;
  series: XlsxChartSeries[];
  varyColors?: boolean;
}

export interface XlsxChartWall {
  fillColor?: string;
  hidden?: boolean;
  lineColor?: string;
  thickness?: number;
}

// ── Chart ─────────────────────────────────────────────────────────────────

export interface XlsxChart {
  anchor: XlsxImageAnchor;
  autoTitleDeleted?: boolean;
  axes: XlsxChartAxis[];
  axisLabelColor?: string;
  axisLineColor?: string;
  categoryAxis?: XlsxChartAxis | null;
  chartExLayout?: string;
  chartAreaBorderColor?: string;
  chartAreaFillColor?: string;
  chartColorPalette?: string[];
  chartColorPaletteOffset?: number;
  chartPath?: string;
  chartStyleId?: number;
  chartType: string;
  dataLabels?: XlsxChartDataLabels | null;
  displayBlanksAs?: string;
  editable?: boolean;
  firstSliceAngle?: number;
  fontFamily?: string;
  gapWidth?: number;
  holeSize?: number;
  id: string;
  is3d?: boolean;
  legend?: XlsxChartLegend | null;
  name?: string;
  overlap?: number;
  plotVisibleOnly?: boolean;
  raw?: Record<string, unknown>;
  radarStyle?: string;
  scatterStyle?: string;
  roundedCorners?: boolean;
  shape3d?: string;
  seriesAxis?: XlsxChartAxis | null;
  series: XlsxChartSeries[];
  sheetIndex: number;
  showDlblsOverMax?: boolean;
  sideWall?: XlsxChartWall | null;
  backWall?: XlsxChartWall | null;
  bubbleScale?: number;
  bubble3d?: boolean;
  floor?: XlsxChartWall | null;
  surfaceMaterial?: string;
  textColor?: string;
  title?: string;
  titleColor?: string;
  titleFontFamily?: string;
  typeGroups?: XlsxChartTypeGroup[];
  valueAxis?: XlsxChartAxis | null;
  varyColors?: boolean;
  view3d?: {
    depthPercent?: number;
    perspective?: number;
    rAngAx?: boolean;
    rotX?: number;
    rotY?: number;
  };
  wireframe?: boolean;
  workbookSheetIndex: number;
  zIndex: number;
}

// ── Chart sheet ──────────────────────────────────────────────────────────

export interface XlsxChartsheet {
  chartIds: string[];
  chartPath?: string;
  id: string;
  index: number;
  name: string;
  raw?: Record<string, unknown>;
  workbookSheetIndex?: number;
}

// ── Render props ─────────────────────────────────────────────────────────

export interface XlsxChartLoadingRenderProps {
  /** The chart that is waiting for its renderer or data. */
  chart: XlsxChart;
  /** The built-in chart loading placeholder. */
  defaultNode: unknown;
  /** The chart rectangle in viewer pixels, including the current zoom level. */
  rect: XlsxImageRect;
}
