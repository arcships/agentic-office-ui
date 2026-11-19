/**
 * Bridges the sheet model's parsed style shape to the WASM engine's style input.
 *
 * `styleById` mirrors the source OOXML — colours arrive as `{ theme: 1 }`,
 * `{ rgb: "FFFF0000" }`, `{ indexed: 64 }` or `{ auto: true }` — while
 * `Worksheet.setCellStyleAt` only accepts colours that carry `colorType`,
 * `hex`, `rgb`, `themeIndex` or `paletteIndex`. Undo restores a parsed default
 * style through that engine call, so an untranslated theme colour fails at the
 * WASM boundary (`color requires colorType, hex, rgb, themeIndex, or
 * paletteIndex`) instead of clearing the border.
 */

const BORDER_EDGE_KEYS = [
  "left",
  "right",
  "top",
  "bottom",
  "diagonal",
  "vertical",
  "horizontal",
  "start",
  "end"
] as const;

function asRecord(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" ? value as Record<string, unknown> : null;
}

function finiteNumber(value: unknown): number | null {
  const numeric = typeof value === "number" ? value : typeof value === "string" && value.trim() ? Number(value) : Number.NaN;
  return Number.isFinite(numeric) ? numeric : null;
}

/**
 * Translate one parsed colour into an engine colour input.
 *
 * @returns The engine-shaped colour, or `undefined` when the source shape has
 *   nothing the engine understands, so the caller can drop the key and let the
 *   engine keep its own value instead of throwing.
 */
export function toEngineStyleColor(value: unknown): unknown {
  const record = asRecord(value);
  if (!record) {
    return value;
  }

  // Engine-resolved colours always carry one of these keys and round-trip as-is.
  if (
    typeof record.hex === "string" ||
    typeof record.themeIndex === "number" ||
    typeof record.paletteIndex === "number"
  ) {
    return value;
  }

  const hex = [record.rgb, record.argb]
    .find((candidate): candidate is string => typeof candidate === "string" && candidate.trim().length > 0);
  if (hex) {
    return { colorType: "rgb", hex };
  }

  const themeIndex = finiteNumber(record.theme);
  if (themeIndex !== null) {
    // The engine types the tint as an integer, so a parsed fractional tint is
    // dropped rather than rejected.
    return { colorType: "theme", themeIndex };
  }

  const paletteIndex = finiteNumber(record.indexed);
  if (paletteIndex !== null) {
    return { colorType: "indexed", paletteIndex };
  }

  if (record.auto === true || record.colorType === "auto") {
    return { colorType: "auto", hex: "000000" };
  }

  return undefined;
}

function convertColorKeys(record: Record<string, unknown>, keys: readonly string[]) {
  for (const key of keys) {
    if (!(key in record)) {
      continue;
    }
    const color = toEngineStyleColor(record[key]);
    if (color === undefined) {
      delete record[key];
    } else {
      record[key] = color;
    }
  }
}

function convertFill(fill: Record<string, unknown>) {
  const converted: Record<string, unknown> = { ...fill };
  convertColorKeys(converted, ["color", "foreground", "background"]);
  if (converted.fillType === "pattern") {
    // The parser names the pattern `patternType`; the engine calls it `pattern`.
    if (converted.pattern === undefined && converted.patternType !== undefined) {
      converted.pattern = converted.patternType;
    }
    delete converted.patternType;
    return converted;
  }
  if (converted.fillType === "solid" && converted.color === undefined) {
    // A parsed solid fill keeps its colour in `foreground` when the source
    // workbook wrote `fgColor`; the engine expects it in `color`.
    if (converted.foreground !== undefined) {
      converted.color = converted.foreground;
    }
    delete converted.foreground;
    delete converted.background;
  }
  return converted;
}

function convertBorder(border: Record<string, unknown>) {
  const converted: Record<string, unknown> = { ...border };
  for (const edge of BORDER_EDGE_KEYS) {
    const edgeStyle = asRecord(converted[edge]);
    if (!edgeStyle) {
      continue;
    }
    const nextEdge = { ...edgeStyle };
    convertColorKeys(nextEdge, ["color"]);
    converted[edge] = nextEdge;
  }
  return converted;
}

/**
 * Rewrite a parsed cell style so the engine accepts it.
 *
 * Only the colour-bearing groups are touched: font, fill, border edges and the
 * gradient stops inside a fill. Every other group already matches the engine's
 * input shape, or is a render-only addition the engine ignores.
 */
export function toEngineCellStyle(style: unknown): unknown {
  const record = asRecord(style);
  if (!record) {
    return style;
  }

  const converted: Record<string, unknown> = { ...record };
  // Renderer-only metadata recorded by the table-style parser; the engine's
  // `setCellStyleAt` accepts only font/fill/border/alignment/... groups.
  delete converted.tableStyleOverrides;

  const font = asRecord(record.font);
  if (font) {
    const nextFont = { ...font };
    convertColorKeys(nextFont, ["color"]);
    converted.font = nextFont;
  }

  const fill = asRecord(record.fill);
  if (fill) {
    const nextFill = convertFill(fill);
    const stops = Array.isArray(nextFill.stops) ? nextFill.stops : null;
    if (stops) {
      nextFill.stops = stops.map((stop) => {
        const stopRecord = asRecord(stop);
        if (!stopRecord) {
          return stop;
        }
        const nextStop = { ...stopRecord };
        convertColorKeys(nextStop, ["color"]);
        return nextStop;
      });
    }
    converted.fill = nextFill;
  }

  const border = asRecord(record.border);
  if (border) {
    converted.border = convertBorder(border);
  }

  return converted;
}
