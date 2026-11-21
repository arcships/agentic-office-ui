import type { OfficeFormat } from "./types"

/**
 * Framework-agnostic chart-internal geometry contract.
 *
 * A chart renderer (or a host adapter around one) publishes the rectangles and
 * point anchors it actually paints, so a Surface can hit-test series, points,
 * legend entries and axis ticks without reading private DOM or duplicating the
 * renderer's layout engine. The contract deliberately carries no drawing
 * commands and no style: it is about *where* things are, not how they look.
 *
 * Coordinate space: `chart-local-px` means CSS pixels with the origin at the
 * chart area's top-left corner. A Surface adds the chart's own page/slide/sheet
 * offset to map the snapshot onto the stage. Providers return the frame they
 * measured, so a snapshot stays valid when a chart is resized or zoomed by
 * asking for a new snapshot at the new frame.
 */
export type ChartGeometrySpace = "chart-local-px"

/**
 * How much a field can be trusted. This is a statement about the provider, not
 * about Microsoft Office parity.
 *
 * - `exact`: computed with the same formulas that painted the chart at the
 *   returned frame. The number is the painted coordinate, not an estimate.
 * - `approximate`: derived with estimated text metrics or collision heuristics
 *   (for example a label's glyph box, or a rotated tick label's bounding box).
 * - `provider`: the engine has no rule for this part; only the host renderer
 *   can supply it (for example 3-D projection hit regions or ECharts internals).
 */
export type ChartGeometryConfidence = "exact" | "approximate" | "provider"

export interface ChartGeometryPoint {
  x: number
  y: number
}

export interface ChartGeometryRect extends ChartGeometryPoint {
  width: number
  height: number
}

export interface ChartGeometryFrame {
  width: number
  height: number
}

/** A painted text label: its text, anchor and (optionally) its glyph box. */
export interface ChartGeometryLabel {
  text: string
  /** SVG-style text anchor point in chart-local pixels. */
  anchor: ChartGeometryPoint
  /** Estimated glyph box. Text metrics depend on the loaded font. */
  rect?: ChartGeometryRect
  confidence: ChartGeometryConfidence
}

export type ChartLegendPosition =
  | "left"
  | "right"
  | "top"
  | "bottom"
  | "topRight"
  | "unknown"

export interface ChartLegendItemGeometry {
  /** Position in the provider's legend item list (paint order). */
  index: number
  /** Index into `ChartGeometrySnapshot.series` when the entry maps to a series. */
  seriesIndex?: number
  /** Category/point index for pie-like legends that list data points. */
  pointIndex?: number
  label: string
  /** Hit box for the swatch plus label. Approximate when it includes text. */
  rect: ChartGeometryRect
  /** Exact swatch box when the renderer paints one. */
  swatchRect?: ChartGeometryRect
  /** SVG-style anchor of the first label line. */
  labelAnchor?: ChartGeometryPoint
  /** Wrapped label lines, when the renderer wraps. */
  lines?: readonly string[]
  labelRect?: ChartGeometryRect
  confidence: ChartGeometryConfidence
}

export interface ChartLegendGeometry {
  position: ChartLegendPosition
  /** Union of the item boxes, in chart-local pixels. */
  rect: ChartGeometryRect
  items: readonly ChartLegendItemGeometry[]
  confidence: ChartGeometryConfidence
}

export interface ChartAxisTickGeometry {
  /** Position in the provider's tick list. */
  index: number
  label: string
  /** Category index for category/date axes. */
  categoryIndex?: number
  /** Numeric value for value axes. */
  value?: number
  /** Point where the tick crosses the axis line. Exact for painted axes. */
  position: ChartGeometryPoint
  /** SVG-style anchor of the painted label, omitted when the label is hidden. */
  labelAnchor?: ChartGeometryPoint
  /** Estimated glyph box; font metrics make this approximate. */
  labelRect?: ChartGeometryRect
  /** False when the renderer skips this label (tick label interval). */
  labelVisible?: boolean
  labelConfidence: ChartGeometryConfidence
}

export interface ChartAxisGeometry {
  axisId?: number
  kind: "category" | "date" | "value" | "series"
  title?: string
  /** Axis line endpoints. Omitted when the axis is hidden. */
  line?: { start: ChartGeometryPoint; end: ChartGeometryPoint }
  /** Zero-crossing line for value axes, when the renderer paints one. */
  zeroLine?: { start: ChartGeometryPoint; end: ChartGeometryPoint }
  ticks: readonly ChartAxisTickGeometry[]
  confidence: ChartGeometryConfidence
}

export type ChartSeriesShape =
  | "bar"
  | "line"
  | "area"
  | "scatter"
  | "pie"
  | "doughnut"
  | "other"

export interface ChartSeriesPointGeometry {
  /** Category/point index in the source data. */
  index: number
  /** Axis category index; equals `index` for ordinary category charts. */
  categoryIndex: number
  value: number | null
  /** Mark anchor: bar centre, line vertex, marker centre. */
  position: ChartGeometryPoint
  /**
   * Mark/hit box. Bars and other area marks are exact; line marker boxes and
   * rotated/shaped marks are approximate.
   */
  rect?: ChartGeometryRect
  /** Confidence for `position`. */
  confidence: ChartGeometryConfidence
  /** Confidence for `rect`, when present. */
  rectConfidence?: ChartGeometryConfidence
}

export interface ChartSeriesGeometry {
  /** Index into the chart model's series array. */
  index: number
  /** Stable series id from the model, when the renderer has one. */
  seriesId?: string
  name: string
  shape: ChartSeriesShape
  /** Index into `ChartGeometrySnapshot.axes`. */
  categoryAxisIndex?: number
  /** Index into `ChartGeometrySnapshot.axes`. */
  valueAxisIndex?: number
  points: readonly ChartSeriesPointGeometry[]
  confidence: ChartGeometryConfidence
}

/**
 * A part the provider cannot describe. Presence of a limitation never changes
 * the confidence of unrelated parts; consumers should treat `exact` as
 * "every requested part I expose is exact", not "the chart matches Office".
 */
export interface ChartGeometryLimitation {
  /** Stable machine code, for example `xlsx.manual-layout-not-parsed`. */
  code: string
  /** Human-readable explanation, safe to surface in diagnostics. */
  detail: string
  /** Confidence a consumer should assign to the affected part. */
  confidence: Extract<ChartGeometryConfidence, "approximate" | "provider">
}

export interface ChartGeometrySnapshot {
  chartId: string
  /** Provider id, for diagnostics and caching. */
  provider: string
  format: OfficeFormat
  space: ChartGeometrySpace
  /** Frame the provider laid out. Chart-local pixels. */
  frame: ChartGeometryFrame
  layout: {
    /** Whole chart area. */
    chartArea: ChartGeometryRect
    /** Cartesian plot area. For non-cartesian charts this is the drawing area. */
    plotArea: ChartGeometryRect
    /** Reserved title band (not the glyph box). */
    title?: ChartGeometryRect
    titleLabel?: ChartGeometryLabel
    legend?: ChartLegendGeometry
    confidence: ChartGeometryConfidence
  }
  axes: readonly ChartAxisGeometry[]
  series: readonly ChartSeriesGeometry[]
  limitations?: readonly ChartGeometryLimitation[]
  /**
   * True only for a full-detail snapshot whose structural parts (layout, axes,
   * series) are all `exact`. Text boxes can still be `approximate`; inspect the
   * per-field confidences. Layout-only requests return false because axes and
   * series were intentionally omitted.
   */
  exact: boolean
}

export interface ChartGeometryRequest {
  /** Reject a mismatched chart instead of returning geometry for the wrong one. */
  chartId?: string
  frame: ChartGeometryFrame
  /** Detail level; `layout` omits axes and series. Defaults to `full`. */
  detail?: "layout" | "full"
  /** Optional document revision the snapshot was measured against. */
  documentId?: string
  revision?: string
}

/**
 * A chart-internal geometry provider. The generic parameter is the format's
 * native chart model (for XLSX, `XlsxChart`), so adapters stay framework-free.
 */
export interface ChartGeometryProvider<TChart = unknown> {
  readonly id: string
  readonly format: OfficeFormat
  getChartGeometry(
    chart: TChart,
    request: ChartGeometryRequest,
  ): ChartGeometrySnapshot | undefined
}

export type ChartGeometryHit =
  | { kind: "chart"; chartId: string; confidence: ChartGeometryConfidence }
  | { kind: "plotArea"; chartId: string; confidence: ChartGeometryConfidence }
  | {
      kind: "series"
      chartId: string
      seriesIndex: number
      pointIndex?: number
      confidence: ChartGeometryConfidence
    }
  | {
      kind: "legendEntry"
      chartId: string
      legendIndex: number
      seriesIndex?: number
      confidence: ChartGeometryConfidence
    }
  | {
      kind: "axisTick"
      chartId: string
      axisIndex: number
      tickIndex: number
      confidence: ChartGeometryConfidence
    }

export function chartGeometryRectContains(
  rect: ChartGeometryRect,
  point: ChartGeometryPoint,
): boolean {
  return point.x >= rect.x
    && point.x <= rect.x + rect.width
    && point.y >= rect.y
    && point.y <= rect.y + rect.height
}

export function chartGeometryRectIntersects(
  left: ChartGeometryRect,
  right: ChartGeometryRect,
): boolean {
  const x = Math.max(left.x, right.x)
  const y = Math.max(left.y, right.y)
  const width = Math.min(left.x + left.width, right.x + right.width) - x
  const height = Math.min(left.y + left.height, right.y + right.height) - y
  return width > 0 && height > 0
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value)
}

function isRect(value: unknown): value is ChartGeometryRect {
  if (!value || typeof value !== "object") return false
  const rect = value as ChartGeometryRect
  return isFiniteNumber(rect.x)
    && isFiniteNumber(rect.y)
    && isFiniteNumber(rect.width)
    && isFiniteNumber(rect.height)
    && rect.width >= 0
    && rect.height >= 0
}

function isPoint(value: unknown): value is ChartGeometryPoint {
  if (!value || typeof value !== "object") return false
  const point = value as ChartGeometryPoint
  return isFiniteNumber(point.x) && isFiniteNumber(point.y)
}

/**
 * Validate the structural invariants every provider must satisfy. Returns an
 * empty array when the snapshot is usable. Strings intentionally address the
 * field path so a failing provider can be fixed from the test output.
 */
export function validateChartGeometrySnapshot(
  snapshot: ChartGeometrySnapshot,
): readonly string[] {
  const issues: string[] = []
  if (!snapshot || typeof snapshot !== "object") {
    return ["snapshot: must be an object"]
  }
  if (!snapshot.chartId?.trim()) issues.push("chartId: must not be empty")
  if (!snapshot.provider?.trim()) issues.push("provider: must not be empty")
  if (!Number.isInteger(snapshot.frame?.width) || snapshot.frame.width <= 0) {
    issues.push("frame.width: must be a positive integer")
  }
  if (!Number.isInteger(snapshot.frame?.height) || snapshot.frame.height <= 0) {
    issues.push("frame.height: must be a positive integer")
  }
  if (!isRect(snapshot.layout?.chartArea)) issues.push("layout.chartArea: invalid rect")
  if (!isRect(snapshot.layout?.plotArea)) issues.push("layout.plotArea: invalid rect")
  if (snapshot.layout?.title !== undefined && !isRect(snapshot.layout.title)) {
    issues.push("layout.title: invalid rect")
  }
  if (snapshot.layout?.titleLabel !== undefined) {
    const label = snapshot.layout.titleLabel
    if (!label.text?.trim()) issues.push("layout.titleLabel.text: must not be empty")
    if (!isPoint(label.anchor)) issues.push("layout.titleLabel.anchor: invalid point")
    if (label.rect !== undefined && !isRect(label.rect)) {
      issues.push("layout.titleLabel.rect: invalid rect")
    }
  }
  const legend = snapshot.layout?.legend
  if (legend) {
    if (!isRect(legend.rect)) issues.push("layout.legend.rect: invalid rect")
    if (!Array.isArray(legend.items)) {
      issues.push("layout.legend.items: must be an array")
    }
    const legendItems: readonly ChartLegendItemGeometry[] = Array.isArray(legend.items) ? legend.items : []
    legendItems.forEach((item, index) => {
      if (!Number.isInteger(item.index) || item.index < 0) {
        issues.push(`layout.legend.items[${index}].index: must be a non-negative integer`)
      }
      if (!item.label?.trim()) {
        issues.push(`layout.legend.items[${index}].label: must not be empty`)
      }
      if (!isRect(item.rect)) {
        issues.push(`layout.legend.items[${index}].rect: invalid rect`)
      }
      if (item.swatchRect !== undefined && !isRect(item.swatchRect)) {
        issues.push(`layout.legend.items[${index}].swatchRect: invalid rect`)
      }
      if (item.labelAnchor !== undefined && !isPoint(item.labelAnchor)) {
        issues.push(`layout.legend.items[${index}].labelAnchor: invalid point`)
      }
      if (item.labelRect !== undefined && !isRect(item.labelRect)) {
        issues.push(`layout.legend.items[${index}].labelRect: invalid rect`)
      }
      if (item.seriesIndex !== undefined
        && (!Number.isInteger(item.seriesIndex) || item.seriesIndex < 0)) {
        issues.push(`layout.legend.items[${index}].seriesIndex: must be a non-negative integer`)
      }
      if (item.pointIndex !== undefined
        && (!Number.isInteger(item.pointIndex) || item.pointIndex < 0)) {
        issues.push(`layout.legend.items[${index}].pointIndex: must be a non-negative integer`)
      }
    })
  }
  if (!Array.isArray(snapshot.axes)) {
    issues.push("axes: must be an array")
  }
  const axes: readonly ChartAxisGeometry[] = Array.isArray(snapshot.axes) ? snapshot.axes : []
  axes.forEach((axis, axisIndex) => {
    if (!["category", "date", "value", "series"].includes(axis.kind)) {
      issues.push(`axes[${axisIndex}].kind: unknown axis kind ${String(axis.kind)}`)
    }
    if (axis.line && (!isPoint(axis.line.start) || !isPoint(axis.line.end))) {
      issues.push(`axes[${axisIndex}].line: invalid endpoints`)
    }
    if (axis.zeroLine && (!isPoint(axis.zeroLine.start) || !isPoint(axis.zeroLine.end))) {
      issues.push(`axes[${axisIndex}].zeroLine: invalid endpoints`)
    }
    if (!Array.isArray(axis.ticks)) {
      issues.push(`axes[${axisIndex}].ticks: must be an array`)
    }
    const ticks: readonly ChartAxisTickGeometry[] = Array.isArray(axis.ticks) ? axis.ticks : []
    ticks.forEach((tick, tickIndex) => {
      if (!Number.isInteger(tick.index) || tick.index < 0) {
        issues.push(`axes[${axisIndex}].ticks[${tickIndex}].index: must be a non-negative integer`)
      }
      if (!isPoint(tick.position)) {
        issues.push(`axes[${axisIndex}].ticks[${tickIndex}].position: invalid point`)
      }
      if (tick.labelAnchor !== undefined && !isPoint(tick.labelAnchor)) {
        issues.push(`axes[${axisIndex}].ticks[${tickIndex}].labelAnchor: invalid point`)
      }
      if (tick.labelRect !== undefined && !isRect(tick.labelRect)) {
        issues.push(`axes[${axisIndex}].ticks[${tickIndex}].labelRect: invalid rect`)
      }
      if (tick.categoryIndex !== undefined
        && (!Number.isInteger(tick.categoryIndex) || tick.categoryIndex < 0)) {
        issues.push(`axes[${axisIndex}].ticks[${tickIndex}].categoryIndex: must be a non-negative integer`)
      }
      if (tick.value !== undefined && !isFiniteNumber(tick.value)) {
        issues.push(`axes[${axisIndex}].ticks[${tickIndex}].value: must be finite`)
      }
    })
  })
  if (!Array.isArray(snapshot.series)) {
    issues.push("series: must be an array")
  }
  const seriesEntries: readonly ChartSeriesGeometry[] = Array.isArray(snapshot.series) ? snapshot.series : []
  seriesEntries.forEach((series, seriesIndex) => {
    if (!Number.isInteger(series.index) || series.index < 0) {
      issues.push(`series[${seriesIndex}].index: must be a non-negative integer`)
    }
    if (!series.name?.trim()) {
      issues.push(`series[${seriesIndex}].name: must not be empty`)
    }
    if (series.categoryAxisIndex !== undefined
      && (!Number.isInteger(series.categoryAxisIndex)
        || series.categoryAxisIndex < 0
        || series.categoryAxisIndex >= axes.length)) {
      issues.push(`series[${seriesIndex}].categoryAxisIndex: outside axes`)
    }
    if (series.valueAxisIndex !== undefined
      && (!Number.isInteger(series.valueAxisIndex)
        || series.valueAxisIndex < 0
        || series.valueAxisIndex >= axes.length)) {
      issues.push(`series[${seriesIndex}].valueAxisIndex: outside axes`)
    }
    if (!Array.isArray(series.points)) {
      issues.push(`series[${seriesIndex}].points: must be an array`)
    }
    const points: readonly ChartSeriesPointGeometry[] = Array.isArray(series.points) ? series.points : []
    points.forEach((point, pointIndex) => {
      if (!Number.isInteger(point.index) || point.index < 0) {
        issues.push(`series[${seriesIndex}].points[${pointIndex}].index: must be a non-negative integer`)
      }
      if (!Number.isInteger(point.categoryIndex) || point.categoryIndex < 0) {
        issues.push(`series[${seriesIndex}].points[${pointIndex}].categoryIndex: must be a non-negative integer`)
      }
      if (!isPoint(point.position)) {
        issues.push(`series[${seriesIndex}].points[${pointIndex}].position: invalid point`)
      }
      if (point.rect !== undefined && !isRect(point.rect)) {
        issues.push(`series[${seriesIndex}].points[${pointIndex}].rect: invalid rect`)
      }
      if (point.value !== null && !isFiniteNumber(point.value)) {
        issues.push(`series[${seriesIndex}].points[${pointIndex}].value: must be finite or null`)
      }
    })
  })
  return issues
}

export function assertChartGeometrySnapshot(
  snapshot: ChartGeometrySnapshot,
): ChartGeometrySnapshot {
  const issues = validateChartGeometrySnapshot(snapshot)
  if (issues.length > 0) {
    throw new TypeError(`invalid chart geometry snapshot: ${issues.join("; ")}`)
  }
  return snapshot
}

function rectDistance(rect: ChartGeometryRect, point: ChartGeometryPoint): number {
  const dx = Math.max(rect.x - point.x, 0, point.x - (rect.x + rect.width))
  const dy = Math.max(rect.y - point.y, 0, point.y - (rect.y + rect.height))
  return Math.hypot(dx, dy)
}

/**
 * Resolve the topmost chart part under a chart-local point.
 *
 * Precedence follows paint order: data marks, then legend entries, then axis
 * labels, then the plot area, then the chart area. Marks without a rect use a
 * small tolerance around their anchor so line/scatter points remain hittable.
 */
export function hitTestChartGeometry(
  snapshot: ChartGeometrySnapshot,
  point: ChartGeometryPoint,
  options: { pointTolerancePx?: number; axisLabels?: boolean } = {},
): ChartGeometryHit | null {
  const tolerance = options.pointTolerancePx ?? 4
  const includeAxisLabels = options.axisLabels !== false
  for (let seriesIndex = snapshot.series.length - 1; seriesIndex >= 0; seriesIndex -= 1) {
    const series = snapshot.series[seriesIndex]
    for (let pointIndex = series.points.length - 1; pointIndex >= 0; pointIndex -= 1) {
      const mark = series.points[pointIndex]
      const hit = mark.rect
        ? chartGeometryRectContains(mark.rect, point)
        : Math.hypot(mark.position.x - point.x, mark.position.y - point.y) <= tolerance
      if (hit) {
        return {
          kind: "series",
          chartId: snapshot.chartId,
          seriesIndex: series.index,
          pointIndex: mark.index,
          confidence: mark.rect ? (mark.rectConfidence ?? mark.confidence) : mark.confidence,
        }
      }
    }
  }
  if (includeAxisLabels) {
    for (let axisIndex = snapshot.axes.length - 1; axisIndex >= 0; axisIndex -= 1) {
      const axis = snapshot.axes[axisIndex]
      for (let tickIndex = axis.ticks.length - 1; tickIndex >= 0; tickIndex -= 1) {
        const tick = axis.ticks[tickIndex]
        if (tick.labelRect && chartGeometryRectContains(tick.labelRect, point)) {
          return {
            kind: "axisTick",
            chartId: snapshot.chartId,
            axisIndex,
            tickIndex,
            confidence: tick.labelConfidence,
          }
        }
      }
    }
  }
  const legend = snapshot.layout.legend
  if (legend) {
    for (let legendIndex = legend.items.length - 1; legendIndex >= 0; legendIndex -= 1) {
      const item = legend.items[legendIndex]
      if (rectDistance(item.rect, point) <= 0) {
        return {
          kind: "legendEntry",
          chartId: snapshot.chartId,
          legendIndex: item.index,
          ...(item.seriesIndex === undefined ? {} : { seriesIndex: item.seriesIndex }),
          confidence: item.confidence,
        }
      }
    }
  }
  if (chartGeometryRectContains(snapshot.layout.plotArea, point)) {
    return {
      kind: "plotArea",
      chartId: snapshot.chartId,
      confidence: snapshot.layout.confidence,
    }
  }
  if (chartGeometryRectContains(snapshot.layout.chartArea, point)) {
    return { kind: "chart", chartId: snapshot.chartId, confidence: "exact" }
  }
  return null
}
