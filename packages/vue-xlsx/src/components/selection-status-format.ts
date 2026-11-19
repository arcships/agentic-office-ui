import type { XlsxCellNumberFormatInput } from "@arcships/xlsx-core";
import {
  CURRENCY_SYMBOLS,
  classifyNumberFormatCode,
  resolveFormatLocale,
  stripFormatLiterals,
} from "../composables/number-locale";
import { resolveNumberFormatCode } from "../composables/cell-input";

/** Excel pads numbers to the decimal places a format spells out. */
function countFormatDecimals(code: string): number {
  const match = /\.([0#]+)/.exec(stripFormatLiterals(code.split(";")[0] ?? ""));
  return match?.[1]?.length ?? 0;
}

function usesGrouping(code: string): boolean {
  return /[0#?],[0#?]/.test(stripFormatLiterals(code.split(";")[0] ?? ""));
}

/** Currency symbol written as a bracket locale section, a literal or raw text. */
function resolveCurrencySymbol(code: string): string | null {
  const bracket = /\[\$([^\]-]*)-[0-9A-Fa-f]+\]/.exec(code);
  if (bracket?.[1]) return bracket[1];
  const literal = /"([^"]*[^\s"][^"]*)"/.exec(code);
  const literalSymbol = literal?.[1]
    ? [...literal[1]].find((character) => CURRENCY_SYMBOLS.includes(character))
    : undefined;
  if (literalSymbol) return literalSymbol;
  const withoutSections = code.replace(/\[[^\]]*\]/g, "").replace(/"[^"]*"/g, "");
  return [...withoutSections].find((character) => CURRENCY_SYMBOLS.includes(character)) ?? null;
}

function currencyIsSuffix(code: string, symbol: string): boolean {
  const section = code.split(";")[0] ?? "";
  const symbolIndex = section.indexOf(symbol);
  if (symbolIndex < 0) return false;
  // Locale ids inside `[$€-407]` are digits, not digit placeholders; mask
  // quoted literals and bracket sections before looking past the symbol.
  const masked = section
    .replace(/"[^"]*"/g, (match) => " ".repeat(match.length))
    .replace(/\[[^\]]*\]/g, (match) => " ".repeat(match.length));
  for (let index = symbolIndex + 1; index < section.length; index += 1) {
    if (/[0#?]/.test(masked[index] ?? "")) return false;
  }
  return true;
}

function resolveLocaleTag(code: string, fallbackLocale: string | undefined): string {
  const resolved = resolveFormatLocale(code, fallbackLocale)?.tag ?? fallbackLocale ?? "en-US";
  try {
    // Validate before Intl sees it; hosts can pass any locale string.
    new Intl.NumberFormat(resolved);
    return resolved;
  } catch {
    return "en-US";
  }
}

function formatFinite(
  value: number,
  localeTag: string,
  options: Intl.NumberFormatOptions,
): string {
  try {
    return new Intl.NumberFormat(localeTag, options).format(value);
  } catch {
    return String(value);
  }
}

function decimalSeparator(localeTag: string): string {
  try {
    return new Intl.NumberFormat(localeTag)
      .formatToParts(1.1)
      .find((part) => part.type === "decimal")?.value ?? ".";
  } catch {
    return ".";
  }
}

/** Excel's General look: no grouping, no forced decimals, no float noise. */
function formatGeneral(value: number, localeTag: string): string {
  const magnitude = Math.abs(value);
  if (magnitude !== 0 && (magnitude >= 1e15 || magnitude < 1e-10)) {
    return formatFinite(value, localeTag, {
      maximumSignificantDigits: 10,
      useGrouping: false,
    });
  }
  const rounded = Number(value.toPrecision(12));
  return formatFinite(rounded, localeTag, {
    maximumFractionDigits: 10,
    useGrouping: false,
  });
}

/**
 * Render one status-bar total with the number format of the selection.
 *
 * Number, currency, percent and scientific formats are honoured; date, time,
 * text and fraction formats fall back to General, because a sum of serial
 * numbers or text has no meaningful date/fraction rendering. Locale
 * separators come from the format's `[$-…]` section, then the viewer locale.
 */
export function formatSelectionStatValue(
  value: number | null | undefined,
  numberFormat: XlsxCellNumberFormatInput | null | undefined,
  fallbackLocale?: string,
): string {
  if (value === null || value === undefined || !Number.isFinite(value)) return "";
  const code = resolveNumberFormatCode(numberFormat);
  const category = classifyNumberFormatCode(code);
  const localeTag = resolveLocaleTag(code, fallbackLocale);
  // `[$$-409]` puts the currency in a bracket section, which the category
  // classifier ignores so that `[$-409]` stays a plain number format. An
  // explicit symbol makes it currency here.
  const bracketCurrency = resolveCurrencySymbol(code);

  if (category === "percent") {
    return `${formatFinite(value * 100, localeTag, {
      maximumFractionDigits: countFormatDecimals(code),
      minimumFractionDigits: countFormatDecimals(code),
      useGrouping: usesGrouping(code),
    })}%`;
  }

  if (category === "currency" || (bracketCurrency && (category === "number" || category === "general"))) {
    const symbol = resolveCurrencySymbol(code);
    const decimals = countFormatDecimals(code);
    const grouping = usesGrouping(code);
    if (!symbol) {
      return formatFinite(value, localeTag, {
        maximumFractionDigits: decimals,
        minimumFractionDigits: decimals,
        useGrouping: grouping,
      });
    }
    const magnitude = formatFinite(Math.abs(value), localeTag, {
      maximumFractionDigits: decimals,
      minimumFractionDigits: decimals,
      useGrouping: grouping,
    });
    const signed = value < 0 ? `-${magnitude}` : magnitude;
    if (currencyIsSuffix(code, symbol)) return `${signed}${symbol}`;
    return value < 0 ? `-${symbol}${magnitude}` : `${symbol}${magnitude}`;
  }

  if (category === "number") {
    return formatFinite(value, localeTag, {
      maximumFractionDigits: countFormatDecimals(code),
      minimumFractionDigits: countFormatDecimals(code),
      useGrouping: usesGrouping(code),
    });
  }

  if (category === "scientific") {
    const decimals = countFormatDecimals(code);
    const decimalMark = decimalSeparator(localeTag);
    const formattedCode = stripFormatLiterals(code);
    const exponentPlaceholders = /[eE][+-](0+)/.exec(formattedCode)?.[1] ?? "0";
    const marker = /[eE]/.exec(code.split(";")[0] ?? "")?.[0] ?? "E";
    const exponentDigits = Math.max(1, exponentPlaceholders.length);
    const [mantissa = "", rawExponent = "0"] = value.toExponential(decimals).split("e");
    const negative = rawExponent.startsWith("-");
    const digits = rawExponent.replace(/^[+-]/, "").padStart(exponentDigits, "0");
    return `${mantissa.replace(".", decimalMark)}${marker}${negative ? "-" : "+"}${digits}`;
  }

  return formatGeneral(value, localeTag);
}
