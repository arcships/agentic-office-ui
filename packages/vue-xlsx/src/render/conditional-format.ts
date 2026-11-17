// Conditional-format resolution for the XLSX grid renderer.
//
// Excel evaluates conditional formats per cell: the rule with the lowest
// `priority` number wins, its `cfvos` define the bounds of the scale (a fixed
// number, a percentage of the range, a percentile rank or a formula), and the
// rule geometry (data-bar length, icon bucket) derives from those bounds.
// The helpers here are pure: the grid only has to turn their results into
// canvas drawing commands, and tests can exercise them without a canvas.
import type {
  XlsxCellAddress,
  XlsxConditionalDataBarRule,
  XlsxConditionalFormatRule,
  XlsxConditionalFormatValueObject,
  XlsxConditionalIconSetRule,
} from "@arcships/xlsx-core";

export type ConditionalColorResolver = (
  color: Record<string, unknown> | undefined,
) => string | undefined;

export interface ConditionalRuleStats {
  max: number;
  min: number;
  values: number[];
}

export interface ConditionalBounds {
  high: number;
  /** True when the high bound comes from the data/host, not a fixed number. */
  highAutomatic: boolean;
  low: number;
  /** True when the low bound comes from the data/host, not a fixed number. */
  lowAutomatic: boolean;
}

export interface ConditionalDataBarStyle {
  axisColor: string;
  border: boolean;
  borderColor: string;
  color: string;
  gradient: boolean;
  maxLength: number;
  minLength: number;
  negativeBorderColor: string;
  negativeFillColor: string;
}

export interface ConditionalIconVisual {
  color: string;
  glyph: string;
}

export interface ConditionalEnvironment {
  bounds: ConditionalBounds;
  colorScale: {
    colors: string[];
    /** position in the resolved bounds (0..1) for each colour stop. */
    stops: number[];
  };
  dataBar: ConditionalDataBarStyle | null;
  icons: ConditionalIconVisual[];
  /** Absolute values of the icon-set cfvos, in document order. */
  iconThresholds: number[];
  rule: XlsxConditionalFormatRule;
  stats: ConditionalRuleStats;
}

export interface ConditionalDataBarGeometry {
  /** Offset of the zero axis inside the bar area, in pixels. */
  axisX: number;
  /** Offset of the bar's left edge inside the bar area, in pixels. */
  barX: number;
  barWidth: number;
  negative: boolean;
  /** Excel draws the axis only when the scale spans zero. */
  showAxis: boolean;
}

export interface ConditionalIconGeometry extends ConditionalIconVisual {
  bucket: number;
  iconX: number;
  size: number;
  /** Width left for the cell text so glyphs and values never overlap. */
  textWidth: number;
}

const DEFAULT_BAR_COLOR = "#5b9bd5";
const DEFAULT_NEGATIVE_BAR_COLOR = "#dc2626";
const DEFAULT_AXIS_COLOR = "#3f3f46";
const DEFAULT_SCALE_COLORS = ["#f8696b", "#ffeb84", "#63be7b"];
const DEFAULT_MIN_LENGTH = 10;
const DEFAULT_MAX_LENGTH = 90;

const AUTOMATIC_BOUND_TYPES = new Set([
  "formula",
  "max",
  "min",
  "percent",
  "percentile",
]);

function clamp(value: number, min: number, max: number) {
  return Math.min(max, Math.max(min, value));
}

function clampPercent(value: number | undefined, fallback: number) {
  return Number.isFinite(value) ? clamp(Math.round(value as number), 0, 100) : fallback;
}

const BOUND_TYPE_ALIASES: Record<string, string> = {
  automaticmax: "max",
  automaticmin: "min",
  automax: "max",
  automin: "min",
};

function normalizeType(type: string) {
  const normalized = type.trim().toLowerCase();
  return BOUND_TYPE_ALIASES[normalized] ?? normalized;
}

function isAutomaticValueObject(cfvo: XlsxConditionalFormatValueObject) {
  const type = normalizeType(cfvo.type);
  if (type === "num") return false;
  // A formula whose literal could be read (e.g. "42" or "=42") is a fixed bound.
  if (type === "formula") return formulaNumber(cfvo) === undefined;
  return AUTOMATIC_BOUND_TYPES.has(type);
}

function formulaNumber(cfvo: XlsxConditionalFormatValueObject): number | undefined {
  if (typeof cfvo.value === "number" && Number.isFinite(cfvo.value)) return cfvo.value;
  if (typeof cfvo.value !== "string") return undefined;
  const literal = Number(cfvo.value.trim().replace(/^=/, "").replace(/^\+/, ""));
  return Number.isFinite(literal) ? literal : undefined;
}

function numericValue(cfvo: XlsxConditionalFormatValueObject): number | undefined {
  if (typeof cfvo.value === "number" && Number.isFinite(cfvo.value)) return cfvo.value;
  if (typeof cfvo.value === "string") {
    const parsed = Number(cfvo.value.trim());
    if (Number.isFinite(parsed)) return parsed;
  }
  return undefined;
}

export function withAlpha(color: string, alpha: string) {
  return /^#[0-9a-f]{6}$/i.test(color) ? `${color}${alpha}` : color;
}

/** Excel PERCENTILE.INC over the values collected for the rule's ranges. */
export function resolvePercentile(values: readonly number[], percent: number): number {
  const sorted = values.filter((value) => Number.isFinite(value)).sort((left, right) => left - right);
  if (!sorted.length) return 0;
  const rank = clamp(percent, 0, 100) / 100 * (sorted.length - 1);
  const lowIndex = Math.floor(rank);
  const highIndex = Math.min(sorted.length - 1, lowIndex + 1);
  const weight = rank - lowIndex;
  return sorted[lowIndex] + (sorted[highIndex] - sorted[lowIndex]) * weight;
}

export function createConditionalRuleStats(values: readonly number[]): ConditionalRuleStats {
  const finite = values.filter((value) => Number.isFinite(value));
  if (!finite.length) return { max: 0, min: 0, values: [] };
  let min = finite[0];
  let max = finite[0];
  for (const value of finite) {
    if (value < min) min = value;
    if (value > max) max = value;
  }
  return { max, min, values: finite };
}

export function sortConditionalRules(rules: readonly XlsxConditionalFormatRule[]) {
  // Array.prototype.sort is stable (ES2019), so equal priorities keep document
  // order, which is what Excel falls back to as well.
  return [...rules].sort((left, right) => left.priority - right.priority);
}

export function conditionalRulesForCell(
  rules: readonly XlsxConditionalFormatRule[],
  cell: XlsxCellAddress,
) {
  const matches = rules.filter((rule) => rule.ranges.some((range) => (
    cell.row >= Math.min(range.start.row, range.end.row)
    && cell.row <= Math.max(range.start.row, range.end.row)
    && cell.col >= Math.min(range.start.col, range.end.col)
    && cell.col <= Math.max(range.start.col, range.end.col)
  )));
  // Excel applies the rule with the lowest priority number first; the grid draws
  // a single effect, so the caller picks index 0 of the ordered result.
  return sortConditionalRules(matches);
}

export function resolveConditionalValue(
  cfvo: XlsxConditionalFormatValueObject,
  stats: ConditionalRuleStats,
): number | undefined {
  const type = normalizeType(cfvo.type);
  switch (type) {
    case "num":
      return numericValue(cfvo) ?? 0;
    case "min":
      return stats.min;
    case "max":
      return stats.max;
    case "percent": {
      const percent = numericValue(cfvo) ?? 0;
      return stats.min + (stats.max - stats.min) * clamp(percent, 0, 100) / 100;
    }
    case "percentile":
      return resolvePercentile(stats.values, numericValue(cfvo) ?? 0);
    case "formula":
      // Only literal formulas are resolvable here; references fall through to
      // the automatic bound so the rule still renders a sensible scale.
      return formulaNumber(cfvo);
    default:
      return numericValue(cfvo);
  }
}

export function resolveConditionalBounds(
  rule: XlsxConditionalFormatRule,
  stats: ConditionalRuleStats,
): ConditionalBounds {
  const cfvos = rule.cfvos ?? [];
  const lowCfvo = cfvos[0];
  const highCfvo = cfvos[cfvos.length - 1];
  let low = lowCfvo ? resolveConditionalValue(lowCfvo, stats) ?? stats.min : stats.min;
  let high = highCfvo ? resolveConditionalValue(highCfvo, stats) ?? stats.max : stats.max;
  if (high < low) [low, high] = [high, low];
  return {
    high,
    highAutomatic: highCfvo ? isAutomaticValueObject(highCfvo) : true,
    low,
    lowAutomatic: lowCfvo ? isAutomaticValueObject(lowCfvo) : true,
  };
}

export function resolveConditionalRatio(value: number, bounds: ConditionalBounds) {
  const span = bounds.high - bounds.low;
  if (!Number.isFinite(span) || span <= 0) return value >= bounds.high ? 1 : 0;
  return clamp((value - bounds.low) / span, 0, 1);
}

export interface ConditionalColorStop {
  colorIndex: number;
  position: number;
}

export function resolveColorScaleStops(
  rule: XlsxConditionalFormatRule,
  stats: ConditionalRuleStats,
): ConditionalColorStop[] {
  const cfvos = rule.cfvos ?? [];
  const colorCount = rule.kind === "colorScale" ? rule.colors?.length ?? 0 : 0;
  if (!cfvos.length || !colorCount) return [];
  const bounds = resolveConditionalBounds(rule, stats);
  return cfvos
    .map((cfvo, index) => ({
      colorIndex: Math.round(index * (colorCount - 1) / Math.max(1, cfvos.length - 1)),
      position: resolveConditionalRatio(resolveConditionalValue(cfvo, stats) ?? bounds.low, bounds),
    }))
    .sort((left, right) => left.position - right.position);
}

function mixHex(start: string, end: string, ratio: number) {
  if (ratio <= 0) return start;
  if (ratio >= 1) return end;
  const startHex = start.match(/^#?([0-9a-f]{6})$/i)?.[1];
  const endHex = end.match(/^#?([0-9a-f]{6})$/i)?.[1];
  if (!startHex || !endHex) return ratio < .5 ? start : end;
  const channel = (offset: number) => {
    const from = Number.parseInt(startHex.slice(offset, offset + 2), 16);
    const to = Number.parseInt(endHex.slice(offset, offset + 2), 16);
    return Math.round(from + (to - from) * clamp(ratio, 0, 1));
  };
  return `rgb(${channel(0)}, ${channel(2)}, ${channel(4)})`;
}

export function resolveColorScaleFill(
  environment: ConditionalEnvironment,
  value: number,
): string | undefined {
  const { colors, stops } = environment.colorScale;
  if (!colors.length) return undefined;
  if (!stops.length) return colors[0];
  const ratio = resolveConditionalRatio(value, environment.bounds);
  if (ratio <= stops[0]) return colors[0];
  const last = stops[stops.length - 1];
  if (ratio >= last) return colors[colors.length - 1];
  for (let index = 0; index < stops.length - 1; index += 1) {
    const from = stops[index];
    const to = stops[index + 1];
    if (ratio < from || ratio > to) continue;
    const span = to - from;
    return mixHex(colors[index], colors[index + 1], span > 0 ? (ratio - from) / span : 0);
  }
  return colors[colors.length - 1];
}

export function resolveDataBarGeometry(
  environment: ConditionalEnvironment,
  value: number,
  width: number,
): ConditionalDataBarGeometry {
  const style = environment.dataBar;
  const usable = Math.max(0, width);
  const maxLength = style?.maxLength ?? DEFAULT_MAX_LENGTH;
  const minLength = style?.minLength ?? DEFAULT_MIN_LENGTH;
  const maxBar = usable * maxLength / 100;
  const minBar = Math.min(usable * minLength / 100, maxBar);
  const { high, highAutomatic, low, lowAutomatic } = environment.bounds;

  let axisFraction = 0;
  if (low < 0 && high > 0) axisFraction = (0 - low) / (high - low);
  else if (high <= 0) axisFraction = 1;
  const axisX = usable * axisFraction;

  const negative = value < 0;
  let ratio = 0;
  let automatic = highAutomatic;
  let sideWidth = usable;
  if (negative) {
    const span = high <= 0 ? high - low : -low;
    ratio = span > 0 ? clamp((high <= 0 ? high - value : -value) / span, 0, 1) : 0;
    automatic = lowAutomatic;
    sideWidth = axisX;
  } else {
    const span = low >= 0 ? high - low : high;
    ratio = span > 0 ? clamp((low >= 0 ? value - low : value) / span, 0, 1) : (value >= high ? 1 : 0);
    automatic = low >= 0 ? lowAutomatic || highAutomatic : highAutomatic;
    sideWidth = usable - axisX;
  }

  // A fixed bound scales straight to the maximum bar length (with minLength as a
  // floor for any non-zero value); an automatic bound stretches the data range
  // across minLength..maxLength like Excel does.
  let barWidth = automatic
    ? minBar + ratio * (maxBar - minBar)
    : ratio <= 0 ? 0 : Math.min(maxBar, Math.max(minBar, ratio * maxBar));
  barWidth = Math.min(barWidth, sideWidth);

  return {
    axisX,
    barWidth,
    barX: negative ? axisX - barWidth : axisX,
    negative,
    showAxis: low < 0 && high > 0,
  };
}

// Excel icon sets, lowest value bucket first. Rating/quarter sets repeat a shape
// so the bucket is readable from the glyph itself.
const ICON_SETS: Record<string, { colors: string[]; glyphs: string[] }> = {
  "3Arrows": { colors: ["#dc2626", "#d97706", "#16a34a"], glyphs: ["▼", "▶", "▲"] },
  "3ArrowsGray": { colors: ["#a1a1aa", "#71717a", "#3f3f46"], glyphs: ["▼", "▶", "▲"] },
  "3Flags": { colors: ["#dc2626", "#d97706", "#16a34a"], glyphs: ["⚑", "⚑", "⚑"] },
  "3Signs": { colors: ["#dc2626", "#d97706", "#16a34a"], glyphs: ["✖", "▲", "✔"] },
  "3Symbols": { colors: ["#dc2626", "#d97706", "#16a34a"], glyphs: ["✖", "!", "✔"] },
  "3Symbols2": { colors: ["#dc2626", "#d97706", "#16a34a"], glyphs: ["✖", "!", "✔"] },
  "3TrafficLights1": { colors: ["#dc2626", "#d97706", "#16a34a"], glyphs: ["●", "●", "●"] },
  "3TrafficLights2": { colors: ["#dc2626", "#d97706", "#16a34a"], glyphs: ["◯", "◉", "●"] },
  "4Arrows": { colors: ["#dc2626", "#ea580c", "#d97706", "#16a34a"], glyphs: ["▼", "◢", "◥", "▲"] },
  "4ArrowsGray": { colors: ["#a1a1aa", "#8b8b93", "#71717a", "#3f3f46"], glyphs: ["▼", "◢", "◥", "▲"] },
  "4Rating": { colors: ["#3f3f46", "#3f3f46", "#3f3f46", "#3f3f46"], glyphs: ["●", "●●", "●●●", "●●●●"] },
  "4RedToBlack": { colors: ["#dc2626", "#991b1b", "#52525b", "#18181b"], glyphs: ["●", "●", "●", "●"] },
  "4TrafficLights": { colors: ["#dc2626", "#d97706", "#84cc16", "#16a34a"], glyphs: ["●", "●", "●", "●"] },
  "5Arrows": { colors: ["#dc2626", "#ea580c", "#d97706", "#65a30d", "#16a34a"], glyphs: ["▼", "◢", "▶", "◥", "▲"] },
  "5ArrowsGray": { colors: ["#a1a1aa", "#8b8b93", "#71717a", "#52525b", "#3f3f46"], glyphs: ["▼", "◢", "▶", "◥", "▲"] },
  "5Boxes": { colors: ["#dbeafe", "#93c5fd", "#60a5fa", "#3b82f6", "#1d4ed8"], glyphs: ["▪", "▪", "▪", "▪", "▪"] },
  "5Quarters": { colors: ["#3f3f46", "#3f3f46", "#3f3f46", "#3f3f46", "#3f3f46"], glyphs: ["○", "◔", "◑", "◕", "●"] },
  "5Rating": { colors: ["#3f3f46", "#3f3f46", "#3f3f46", "#3f3f46", "#3f3f46"], glyphs: ["●", "●●", "●●●", "●●●●", "●●●●●"] },
};

const FALLBACK_ICONS = { colors: ["#dc2626", "#d97706", "#16a34a"], glyphs: ["▼", "●", "▲"] };

function clampIndex(index: number, length: number) {
  if (!Number.isFinite(index) || length <= 0) return 0;
  return clamp(Math.round(index), 0, length - 1);
}

function resolveIconVisual(iconSet: string, iconId: number): ConditionalIconVisual {
  const table = ICON_SETS[iconSet] ?? FALLBACK_ICONS;
  const index = clampIndex(iconId, table.glyphs.length);
  return { color: table.colors[index], glyph: table.glyphs[index] };
}

export function resolveConditionalIcons(rule: XlsxConditionalFormatRule): ConditionalIconVisual[] {
  if (rule.kind !== "iconSet" || !rule.icons?.length) return [];
  return rule.icons.map((icon) => resolveIconVisual(icon.iconSet, icon.iconId));
}

export function resolveIconBucket(environment: ConditionalEnvironment, value: number) {
  const count = Math.max(1, environment.icons.length);
  const thresholds = environment.iconThresholds;
  let bucket = 0;
  for (let index = 1; index < Math.min(thresholds.length, count); index += 1) {
    if (value >= thresholds[index]) bucket = index;
  }
  if (environment.rule.kind === "iconSet" && environment.rule.reverse) {
    bucket = count - 1 - bucket;
  }
  return clamp(bucket, 0, count - 1);
}

export function resolveIconGeometry(
  environment: ConditionalEnvironment,
  value: number,
  cellWidth: number,
  rowHeight: number,
): ConditionalIconGeometry | null {
  if (!environment.icons.length) return null;
  const bucket = resolveIconBucket(environment, value);
  const icon = environment.icons[bucket] ?? environment.icons[environment.icons.length - 1];
  const size = Math.max(10, Math.min(18, rowHeight - 6));
  const iconX = Math.max(0, cellWidth - size - 2);
  return {
    ...icon,
    bucket,
    iconX,
    size,
    // Reserve the glyph column plus a gutter so right-aligned values stop short
    // of the icon instead of drawing underneath it.
    textWidth: Math.max(0, iconX - 2),
  };
}

export function createConditionalEnvironment(
  rule: XlsxConditionalFormatRule,
  values: readonly number[],
  resolveColor: ConditionalColorResolver,
): ConditionalEnvironment {
  const stats = createConditionalRuleStats(values);
  const bounds = resolveConditionalBounds(rule, stats);
  const cfvos = rule.cfvos ?? [];

  let dataBar: ConditionalDataBarStyle | null = null;
  if (rule.kind === "dataBar") {
    const barRule: XlsxConditionalDataBarRule = rule;
    const color = resolveColor(barRule.color) ?? DEFAULT_BAR_COLOR;
    const borderColor = resolveColor(barRule.borderColor) ?? color;
    dataBar = {
      axisColor: resolveColor(barRule.axisColor) ?? DEFAULT_AXIS_COLOR,
      border: barRule.border === true,
      borderColor,
      color,
      gradient: barRule.gradient !== false,
      maxLength: clampPercent(barRule.maxLength, DEFAULT_MAX_LENGTH),
      minLength: clampPercent(barRule.minLength, DEFAULT_MIN_LENGTH),
      negativeBorderColor: barRule.negativeBarBorderColorSameAsPositive === true
        ? borderColor
        : resolveColor(barRule.negativeBorderColor) ?? DEFAULT_NEGATIVE_BAR_COLOR,
      negativeFillColor: resolveColor(barRule.negativeFillColor) ?? DEFAULT_NEGATIVE_BAR_COLOR,
    };
  }

  const colorScaleColors = rule.kind === "colorScale"
    ? (rule.colors ?? []).map((color, index) => (
      resolveColor(color) ?? DEFAULT_SCALE_COLORS[clampIndex(index, DEFAULT_SCALE_COLORS.length)]
    ))
    : [];
  const colorStops = resolveColorScaleStops(rule, stats)
    .filter((stop) => Number.isFinite(stop.position));

  const iconRule: XlsxConditionalIconSetRule | null = rule.kind === "iconSet" ? rule : null;

  return {
    bounds,
    colorScale: {
      colors: colorStops.map((stop) => colorScaleColors[stop.colorIndex] ?? colorScaleColors[0]),
      stops: colorStops.map((stop) => stop.position),
    },
    dataBar,
    iconThresholds: iconRule
      ? cfvos.map((cfvo) => resolveConditionalValue(cfvo, stats) ?? bounds.low)
      : [],
    icons: resolveConditionalIcons(rule),
    rule,
    stats,
  };
}
