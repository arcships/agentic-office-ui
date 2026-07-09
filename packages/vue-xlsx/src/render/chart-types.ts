/** @jsxImportSource vue */
import type { HierarchyNode, HierarchyRectangularNode } from "d3-hierarchy";
import type { Feature, Geometry } from "geojson";
import type { XlsxChart, XlsxChartAxis, XlsxChartElementSelection, XlsxChartSeries, XlsxImageRect } from "@extend-ai/xlsx-core";

export type ChartRendererPalette = {
  border: string;
  mutedText: string;
  surface: string;
  text: string;
};

export type ChartSvgProps = {
  chart: XlsxChart;
  onChartElementDoubleClick?: (selection: XlsxChartElementSelection, event: MouseEvent) => void;
  onChartElementPointerDown?: (selection: XlsxChartElementSelection, event: PointerEvent) => void;
  palette: ChartRendererPalette;
  rect: XlsxImageRect;
  selectedChartElement?: XlsxChartElementSelection | null;
};

export type LegendItem = {
  color: string;
  label: string;
};

export type PlotRect = {
  height: number;
  left: number;
  top: number;
  width: number;
};

export type ChartLayout = {
  height: number;
  legendItems: LegendItem[];
  legendPosition: string | undefined;
  plot: PlotRect;
  titleHeight: number;
  width: number;
};

export type BarRect = {
  capEnd?: boolean;
  capStart?: boolean;
  categoryIndex: number;
  bottomScale?: number;
  color: string;
  depthOrder?: number;
  depthOffsetX?: number;
  depthOffsetY?: number;
  depthX?: number;
  depthY?: number;
  gradientId?: string;
  height: number;
  invertedNegative?: boolean;
  isHorizontal: boolean;
  key: string;
  left: number;
  shape3d?: string;
  seriesIndex: number;
  stroke: string;
  strokeWidth: number;
  topScale?: number;
  value: number;
  width: number;
  top: number;
};

export type ChartElementDataOptions = {
  selectionMode?: "seriesFirst";
};

export type ChartHierarchyDatum = {
  children?: ChartHierarchyDatum[];
  colorIndex?: number;
  name: string;
  value?: number;
};

export type ComboRenderableGroup = {
  axisIds: number[];
  categoryAxis: XlsxChartAxis | null;
  chartType: string;
  gapWidth?: number;
  is3d?: boolean;
  raw?: Record<string, unknown>;
  series: XlsxChartSeries[];
  valueAxis: XlsxChartAxis | null;
};

export type ChartStage = {
  color: string;
  isSubtotal: boolean;
  label: string;
  value: number;
};

export type BoxWhiskerStats = {
  lowerFence: number;
  lowerWhisker: number;
  max: number;
  mean: number;
  median: number;
  min: number;
  outliers: number[];
  q1: number;
  q3: number;
  upperFence: number;
  upperWhisker: number;
  visiblePoints: number[];
};

export type SurfaceDomain = {
  maxValue: number;
  minValue: number;
  safeMax: number;
  ticks: number[];
};

export type RegionMapFeature = Feature<Geometry, { name?: string; regionSet?: "country" | "us-state"; stateCode?: string }>;

export type SurfacePathPoint = { x: number; y: number };

export type SurfaceContourSegment = {
  end: SurfacePathPoint;
  start: SurfacePathPoint;
};
