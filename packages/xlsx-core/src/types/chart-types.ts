import type {
  XlsxImageAnchor,
  XlsxImageRect,
  XlsxShapeParagraph,
  XlsxShapeTextBox,
} from "./image-types";

// ── Chart reference / data labels ────────────────────────────────────────

/** Chart family a host can insert from a cell range. */
export type XlsxChartFamily = "bar" | "line" | "pie";

export interface XlsxChartReference {
  formula?: string;
  refType?: string;
  values?: Array<number | string | null>;
}

export interface XlsxChartDataLabels {
  numberFormat?: {
    formatCode?: string;
    sourceLinked?: boolean;
  };
  pointLabels?: XlsxChartPointDataLabel[];
  position?: string;
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
  axisType?: "category" | "date" | "value" | "series";
  crossId?: number;
  crosses?: string;
  crossBetween?: string;
  delete?: boolean;
  id?: number;
  labelPosition?: string;
  logBase?: number;
  majorTimeUnit?: string;
  minorTimeUnit?: string;
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
  /** Point picture fill (`c:dPt/c:spPr/a:blipFill`) as a data URL. */
  fillImage?: string;
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
  /** Series picture fill (`c:ser/c:spPr/a:blipFill`) as a data URL. */
  fillImage?: string;
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

// ── Chart-embedded media / user shapes ───────────────────────────────────

export interface XlsxChartUserShapeFill {
  /** Solid fill resolved to a CSS colour. */
  color?: string;
  /** `a:blipFill` picture fill resolved to a data URL by `readChartEmbeddedMedia`. */
  image?: string;
  /** `a:noFill`, or a stroke-only preset such as `line`. */
  none?: boolean;
}

export interface XlsxChartUserShapeStroke {
  color?: string;
  /** Preset dash name from `a:ln/a:prstDash@val`. */
  dash?: string;
  hidden?: boolean;
  widthPx?: number;
}

/**
 * One shape from a chart's `chartUserShapes` part.
 *
 * Excel anchors chart user shapes in the chart's own drawing space, so the rect
 * is normalised to fractions (0–1) of the chart area: `left`/`top` are the
 * offset from the chart's top-left corner and `width`/`height` its size. A
 * renderer multiplies those fractions by its own chart layout, which keeps the
 * overlay aligned at any zoom level and chart size.
 */
export interface XlsxChartUserShape {
  description?: string;
  fill?: XlsxChartUserShapeFill;
  /** `a:prstGeom@prst` (`rect` when the part omits it). */
  geometry: string;
  /** `a:avLst/a:gd` adjustment values as raw percentages (`adj="16667"` → 16667). */
  geometryAdjustments?: Record<string, number>;
  name?: string;
  paragraphs: XlsxShapeParagraph[];
  rect: XlsxImageRect;
  stroke?: XlsxChartUserShapeStroke;
  textBox?: XlsxShapeTextBox;
}

/** One `chartUserShapes` part linked from a chart part. */
export interface XlsxChartDrawing {
  /** Images owned by the shapes part, keyed by relationship id → data URL. */
  images?: Record<string, string>;
  /** Shapes parsed from the part, in document order. */
  shapes: XlsxChartUserShape[];
  /** Raw DrawingML XML, kept so hosts can inspect or re-export the part. */
  xml: string;
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
  /** Chart-area `a:blipFill` picture fill resolved to a data URL. */
  chartAreaFillImage?: string;
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
  /**
   * Chart-embedded images keyed by relationship id → data URL.
   *
   * Picture fills (`a:blipFill/a:blip@r:embed`) in the chart part resolve
   * through this map, which is why it is kept on the model rather than baked
   * into a single series colour.
   */
  images?: Record<string, string>;
  id: string;
  is3d?: boolean;
  /** `chartUserShapes` parts keyed by relationship id. */
  drawings?: Record<string, XlsxChartDrawing>;
  legend?: XlsxChartLegend | null;
  name?: string;
  overlap?: number;
  /** Plot-area `a:blipFill` picture fill resolved to a data URL. */
  plotAreaFillImage?: string;
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
  /** Read-only charts resolved from the chartsheet drawing relationship. */
  charts?: XlsxChart[];
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
