/** @jsxImportSource vue */
import type { XlsxChart } from "@arcships/xlsx-core";

// ---- Color utilities ----

export function parseRgbColor(color: string) {
  const match = /^#?([0-9a-f]{6})$/i.exec(color);
  if (!match) {
    return null;
  }
  return {
    blue: Number.parseInt(match[1].slice(4, 6), 16),
    green: Number.parseInt(match[1].slice(2, 4), 16),
    red: Number.parseInt(match[1].slice(0, 2), 16)
  };
}

export function mixRgbColor(color: string, mixWith: string, ratio: number) {
  const base = parseRgbColor(color);
  const target = parseRgbColor(mixWith);
  if (!base || !target) {
    return color;
  }
  const clamped = Math.max(0, Math.min(1, ratio));
  const mixChannel = (left: number, right: number) => Math.round(left + (right - left) * clamped);
  return `#${[
    mixChannel(base.red, target.red),
    mixChannel(base.green, target.green),
    mixChannel(base.blue, target.blue)
  ].map((value) => value.toString(16).padStart(2, "0")).join("")}`;
}

export function lightenColor(color: string, ratio: number) {
  return mixRgbColor(color, "#ffffff", ratio);
}

export function darkenColor(color: string, ratio: number) {
  return mixRgbColor(color, "#000000", ratio);
}

export function clamp(value: number, min: number, max: number) {
  return Math.max(min, Math.min(max, value));
}

export function normalizeRendererHexColor(value: unknown) {
  if (typeof value !== "string") {
    return null;
  }
  const trimmed = value.trim();
  if (!trimmed) {
    return null;
  }
  const normalized = trimmed.startsWith("#") ? trimmed.slice(1) : trimmed;
  return /^[0-9a-f]{6}$/i.test(normalized) ? `#${normalized.toLowerCase()}` : null;
}

// ---- Font utilities ----

export const DEFAULT_CHART_FONT_STACK = [
  "\"Aptos\"",
  "Calibri",
  "Carlito",
  "\"Segoe UI\"",
  "Tahoma",
  "Arial",
  "sans-serif"
].join(", ");

export const DEFAULT_CHART_TEXT_COLOR = "#000000";

export const DEFAULT_CHART_MUTED_TEXT_COLOR = "#7f7f7f";

export function escapeCssFontFamilyToken(value: string) {
  const trimmed = value.trim();
  if (!trimmed) {
    return "";
  }
  if (
    trimmed.startsWith("\"")
    || trimmed.startsWith("'")
    || /^(serif|sans-serif|monospace|cursive|fantasy|system-ui|math|emoji|fangsong)$/i.test(trimmed)
  ) {
    return trimmed;
  }
  return /\s/.test(trimmed) ? `"${trimmed.replace(/"/g, "\\\"")}"` : trimmed;
}

export function buildChartFontFamily(fontFamily: string | undefined) {
  if (!fontFamily || fontFamily.trim().length === 0) {
    return DEFAULT_CHART_FONT_STACK;
  }
  const tokens = fontFamily
    .split(",")
    .map(escapeCssFontFamilyToken)
    .filter((token) => token.length > 0);
  if (tokens.length === 0) {
    return DEFAULT_CHART_FONT_STACK;
  }
  return [...tokens, "\"Segoe UI\"", "Tahoma", "Arial", "sans-serif"].join(", ");
}

export function resolveChartTextColor(chart: XlsxChart) {
  return chart.textColor ?? chart.titleColor ?? DEFAULT_CHART_TEXT_COLOR;
}

export function resolveChartAxisTextColor(chart: XlsxChart) {
  return chart.axisLabelColor ?? chart.textColor ?? chart.titleColor ?? DEFAULT_CHART_TEXT_COLOR;
}

export function resolveChartMutedTextColor(chart: XlsxChart) {
  return chart.textColor ?? chart.axisLabelColor ?? DEFAULT_CHART_MUTED_TEXT_COLOR;
}

export function safeNumber(value: unknown) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (typeof value === "string") {
    const normalized = value.trim().replace(/,/g, "");
    if (!normalized) {
      return null;
    }
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}

export function normalizeLegendPosition(position: string | undefined) {
  switch (position) {
    case "bottom":
      return "bottom";
    case "left":
      return "left";
    case "right":
      return "right";
    case "top":
      return "top";
    case "b":
      return "bottom";
    case "l":
      return "left";
    case "r":
      return "right";
    case "t":
      return "top";
    default:
      return position;
  }
}

export function normalizeCategoryLabel(value: unknown) {
  if (value == null) {
    return "";
  }
  return String(value)
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .trim();
}

export function truncateSvgText(value: string, maxWidth: number, fontSize = 10) {
  if (!value || maxWidth <= 0) {
    return "";
  }
  const charWidth = Math.max(4.2, fontSize * 0.56);
  const maxChars = Math.max(1, Math.floor(maxWidth / charWidth));
  if (value.length <= maxChars) {
    return value;
  }
  if (maxChars <= 1) {
    return "…";
  }
  return `${value.slice(0, Math.max(0, maxChars - 1)).trimEnd()}…`;
}
