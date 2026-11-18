import type { XlsxThemePalette } from "./types/theme-types";

/**
 * ECMA-376 18.8.27 default indexed colour palette, indices 0-63, plus the two
 * system colours Excel reserves at 64 (foreground) and 65 (background).
 *
 * Workbooks may override any entry with `<colors><indexedColors>` in
 * `xl/styles.xml`; missing entries keep these values.
 */
export const XLSX_INDEXED_COLOR_PALETTE: readonly string[] = Object.freeze([
  "#000000", "#ffffff", "#ff0000", "#00ff00", "#0000ff", "#ffff00", "#ff00ff", "#00ffff",
  "#000000", "#ffffff", "#ff0000", "#00ff00", "#0000ff", "#ffff00", "#ff00ff", "#00ffff",
  "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#c0c0c0", "#808080",
  "#9999ff", "#993366", "#ffffcc", "#ccffff", "#660066", "#ff8080", "#0066cc", "#ccccff",
  "#000080", "#ff00ff", "#ffff00", "#00ffff", "#800080", "#800000", "#008080", "#0000ff",
  "#00ccff", "#ccffff", "#ccffcc", "#ffff99", "#99ccff", "#ff99cc", "#cc99ff", "#ffcc99",
  "#3366ff", "#33cccc", "#99cc00", "#ffcc00", "#ff9900", "#ff6600", "#666699", "#969696",
  "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333",
  "#000000", "#ffffff"
]);

/** Build the effective indexed palette: standard entries plus workbook overrides. */
export function buildIndexedColorPalette(
  overrides?: Record<number, string> | null
): Record<number, string> {
  const palette: Record<number, string> = {};
  XLSX_INDEXED_COLOR_PALETTE.forEach((color, index) => {
    palette[index] = color;
  });

  for (const [key, value] of Object.entries(overrides ?? {})) {
    const index = Number(key);
    const normalized = typeof value === "string" ? normalizeHexColor(value) : null;
    if (Number.isInteger(index) && index >= 0 && normalized) {
      palette[index] = normalized;
    }
  }

  return palette;
}

function normalizeHexColor(value: string) {
  const hex = value.replace(/^#/, "");
  if (hex.length === 8) {
    return `#${hex.slice(2).toLowerCase()}`;
  }
  if (hex.length === 6) {
    return `#${hex.toLowerCase()}`;
  }
  return null;
}

function parseHexColor(color: string): [number, number, number] | null {
  const normalized = normalizeHexColor(color);
  const match = normalized ? /^#([0-9a-f]{6})$/.exec(normalized) : null;
  if (!match) {
    return null;
  }

  const hex = match[1];
  return [
    Number.parseInt(hex.slice(0, 2), 16),
    Number.parseInt(hex.slice(2, 4), 16),
    Number.parseInt(hex.slice(4, 6), 16)
  ];
}

function rgbToHsl(red: number, green: number, blue: number): [number, number, number] {
  const normalizedRed = red / 255;
  const normalizedGreen = green / 255;
  const normalizedBlue = blue / 255;
  const max = Math.max(normalizedRed, normalizedGreen, normalizedBlue);
  const min = Math.min(normalizedRed, normalizedGreen, normalizedBlue);
  const lightness = (max + min) / 2;

  if (max === min) {
    return [0, 0, lightness];
  }

  const delta = max - min;
  const saturation = lightness > 0.5 ? delta / (2 - max - min) : delta / (max + min);
  let hue = 0;

  switch (max) {
    case normalizedRed:
      hue = (normalizedGreen - normalizedBlue) / delta + (normalizedGreen < normalizedBlue ? 6 : 0);
      break;
    case normalizedGreen:
      hue = (normalizedBlue - normalizedRed) / delta + 2;
      break;
    default:
      hue = (normalizedRed - normalizedGreen) / delta + 4;
      break;
  }

  return [hue / 6, saturation, lightness];
}

function hueToRgb(p: number, q: number, t: number) {
  let nextT = t;
  if (nextT < 0) {
    nextT += 1;
  }
  if (nextT > 1) {
    nextT -= 1;
  }
  if (nextT < 1 / 6) {
    return p + (q - p) * 6 * nextT;
  }
  if (nextT < 1 / 2) {
    return q;
  }
  if (nextT < 2 / 3) {
    return p + (q - p) * (2 / 3 - nextT) * 6;
  }
  return p;
}

function hslToRgb(hue: number, saturation: number, lightness: number): [number, number, number] {
  if (saturation === 0) {
    const gray = Math.round(lightness * 255);
    return [gray, gray, gray];
  }

  const q = lightness < 0.5 ? lightness * (1 + saturation) : lightness + saturation - lightness * saturation;
  const p = 2 * lightness - q;

  return [
    Math.round(hueToRgb(p, q, hue + 1 / 3) * 255),
    Math.round(hueToRgb(p, q, hue) * 255),
    Math.round(hueToRgb(p, q, hue - 1 / 3) * 255)
  ];
}

function rgbToHex(red: number, green: number, blue: number) {
  return `#${[red, green, blue]
    .map((channel) => Math.max(0, Math.min(255, channel)).toString(16).padStart(2, "0"))
    .join("")}`;
}

function applyExcelTint(baseColor: string, tint: number) {
  const rgb = parseHexColor(baseColor);
  if (!rgb || !Number.isFinite(tint) || tint === 0) {
    return normalizeHexColor(baseColor);
  }

  const [hue, saturation, lightness] = rgbToHsl(rgb[0], rgb[1], rgb[2]);
  const nextLightness = tint < 0
    ? lightness * (1 + tint)
    : lightness * (1 - tint) + tint;
  const [nextRed, nextGreen, nextBlue] = hslToRgb(hue, saturation, Math.max(0, Math.min(1, nextLightness)));
  return rgbToHex(nextRed, nextGreen, nextBlue);
}

function readColorNumber(value: unknown) {
  const numeric = typeof value === "number"
    ? value
    : typeof value === "string" && value.trim().length > 0
      ? Number(value)
      : Number.NaN;
  return Number.isFinite(numeric) ? numeric : null;
}

function applyOptionalTint(color: string, tint: number | null) {
  return tint === null ? color : applyExcelTint(color, tint);
}

function readIndexedColorIndex(color: Record<string, unknown>) {
  // Raw parsed styles use `indexed`; engine-shaped colours carry
  // `paletteIndex` (usually with `colorType: "indexed"`) and a default-palette
  // `hex`. The palette index must win over that `hex`, because the engine does
  // not know about a workbook's custom `<indexedColors>` override.
  const value = color.paletteIndex ?? color.indexed;
  const index = readColorNumber(value);
  return index !== null && Number.isInteger(index) && index >= 0 ? index : null;
}

export function resolveWorkbookColor(
  color: Record<string, unknown> | undefined,
  themePalette?: XlsxThemePalette | null
): string | null {
  if (!color) {
    return null;
  }

  const tint = readColorNumber(color.tint);
  const indexed = readIndexedColorIndex(color);
  if (indexed !== null) {
    const indexedColor = themePalette?.indexedColors?.[indexed]
      ?? XLSX_INDEXED_COLOR_PALETTE[indexed]
      ?? null;
    if (indexedColor) {
      return applyOptionalTint(indexedColor, tint);
    }
  }

  const directHex = ["hex", "rgb", "argb"]
    .map((key) => color[key])
    .find((value): value is string => typeof value === "string" && value.trim().length > 0);
  if (directHex) {
    return normalizeHexColor(directHex);
  }

  const numericTheme = readColorNumber(color.theme);
  const themeColor = numericTheme === null
    ? null
    : themePalette?.colorsByIndex[numericTheme] ?? null;
  if (themeColor) {
    return applyOptionalTint(themeColor, tint);
  }

  return null;
}

export function resolveWorkbookFillColor(
  fill: Record<string, unknown> | undefined,
  themePalette?: XlsxThemePalette | null
): string | null {
  if (!fill) {
    return null;
  }

  if (fill.fillType === "solid") {
    return resolveWorkbookColor(
      (fill.color as Record<string, unknown> | undefined)
      ?? (fill.foreground as Record<string, unknown> | undefined)
      ?? (fill.background as Record<string, unknown> | undefined),
      themePalette
    );
  }

  if (fill.fillType === "pattern") {
    return resolveWorkbookColor(
      (fill.foreground as Record<string, unknown> | undefined)
      ?? (fill.color as Record<string, unknown> | undefined)
      ?? (fill.background as Record<string, unknown> | undefined),
      themePalette
    );
  }

  return null;
}

export function resolveWorkbookFillStyle(
  fill: Record<string, unknown> | undefined,
  themePalette?: XlsxThemePalette | null
) {
  if (!fill) {
    return {
      backgroundColor: null,
      backgroundImage: null
    };
  }

  if (fill.fillType === "gradient") {
    const rawStops = Array.isArray(fill.stops) ? fill.stops as Array<Record<string, unknown>> : [];
    const stops = rawStops
      .map((stop) => ({
        color: resolveWorkbookColor(stop.color as Record<string, unknown> | undefined, themePalette),
        position: typeof stop.position === "number" ? stop.position : Number(stop.position)
      }))
      .filter((stop): stop is { color: string; position: number } => Boolean(stop.color) && Number.isFinite(stop.position))
      .sort((left, right) => left.position - right.position);
    if (stops.length > 0) {
      const gradientType = typeof fill.gradientType === "string" ? fill.gradientType : "linear";
      if (gradientType === "path") {
        const backgroundColor = stops[stops.length - 1]?.color ?? null;
        return {
          backgroundColor,
          backgroundImage: `radial-gradient(circle at center, ${stops.map((stop) => `${stop.color} ${Math.max(0, Math.min(100, stop.position * 100))}%`).join(", ")})`
        };
      }

      const degree = typeof fill.degree === "number" ? fill.degree : Number(fill.degree ?? 0);
      const cssAngle = Number.isFinite(degree) ? 90 - degree : 90;
      return {
        backgroundColor: stops[stops.length - 1]?.color ?? null,
        backgroundImage: `linear-gradient(${cssAngle}deg, ${stops.map((stop) => `${stop.color} ${Math.max(0, Math.min(100, stop.position * 100))}%`).join(", ")})`
      };
    }
  }

  return {
    backgroundColor: resolveWorkbookFillColor(fill, themePalette),
    backgroundImage: null
  };
}
