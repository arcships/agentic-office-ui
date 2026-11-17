import type { XlsxCellNumberFormatInput } from "@arcships/xlsx-core";

/**
 * Value typing and number-format helpers for XLSX cell input.
 *
 * The module is intentionally free of runtime imports so the rules stay a
 * pure, unit-testable description of how typed/pasted text becomes a cell
 * value and how a number format is reported to the UI.
 */

export type NumberFormatCategory =
  | "general"
  | "text"
  | "number"
  | "currency"
  | "percent"
  | "date"
  | "time"
  | "scientific"
  | "fraction";

export type CellInputOptions = {
  /** Number format of the target cell (explicit style or inherited row/column style). */
  numberFormat?: unknown;
  /** Workbook uses the 1904 date system. */
  date1904?: boolean;
};

export type CellInputResolution =
  | {
      kind: "formula";
      formula: string;
      /** Format to apply when the target cell is General. */
      numberFormat?: XlsxCellNumberFormatInput;
    }
  | {
      kind: "value";
      value: unknown;
      /** Format to apply when the target cell is General. */
      numberFormat?: XlsxCellNumberFormatInput;
    };

const CURRENCY_SYMBOLS = "$€£¥₹₽₩₪¢₫₴₦₱₸₺฿";
const CURRENCY_PREFIX = new RegExp(`^([${CURRENCY_SYMBOLS}])\\s*`);
const CURRENCY_SUFFIX = new RegExp(`\\s*([${CURRENCY_SYMBOLS}])$`);
const CURRENCY_SYMBOL_TEST = new RegExp(`[${CURRENCY_SYMBOLS}]`);

const MS_PER_DAY = 86_400_000;
const EXCEL_1900_EPOCH_UTC = Date.UTC(1899, 11, 30);
const EXCEL_1904_EPOCH_UTC = Date.UTC(1904, 0, 1);

const DATE_FORMULA_FORMATS: Readonly<Record<string, string>> = Object.freeze({
  DATE: "m/d/yyyy",
  DATEVALUE: "m/d/yyyy",
  EDATE: "m/d/yyyy",
  EOMONTH: "m/d/yyyy",
  TODAY: "m/d/yyyy",
  NOW: "m/d/yyyy h:mm",
});

/**
 * Built-in Excel number format ids (ECMA-376 §18.8.30) resolved to their
 * format code. Ids 27–36 and 50–58 are locale variants of 14–22; only the
 * category matters for those, so they reuse the matching date/time code.
 */
export const BUILTIN_NUMBER_FORMAT_CODES: Readonly<Record<number, string>> = Object.freeze({
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  5: "$#,##0_);($#,##0)",
  6: "$#,##0_);[Red]($#,##0)",
  7: "$#,##0.00_);($#,##0.00)",
  8: "$#,##0.00_);[Red]($#,##0.00)",
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "mm-dd-yy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yy h:mm",
  27: "mm-dd-yy",
  28: "d-mmm-yy",
  29: "d-mmm",
  30: "mmm-yy",
  31: "h:mm AM/PM",
  32: "h:mm:ss AM/PM",
  33: "h:mm",
  34: "h:mm:ss",
  35: "m/d/yy h:mm",
  36: "mm-dd-yy",
  37: "#,##0_);(#,##0)",
  38: "#,##0_);[Red](#,##0)",
  39: "#,##0.00_);(#,##0.00)",
  40: "#,##0.00_);[Red](#,##0.00)",
  41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
  42: '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
  44: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mmss.0",
  48: "##0.0E+0",
  49: "@",
  50: "mm-dd-yy",
  51: "d-mmm-yy",
  52: "d-mmm",
  53: "mmm-yy",
  54: "h:mm AM/PM",
  55: "h:mm:ss AM/PM",
  56: "h:mm",
  57: "h:mm:ss",
  58: "m/d/yy h:mm",
});

/**
 * Style patch that restores Excel's default cell formatting, i.e. "the cell has
 * no explicit style". The wasm workbook exposes no way to unset a cell style,
 * so undo and paste reset every style group to its default value instead, and
 * the workbook normalizes a cell that matches its default back to "no style".
 * The values are Excel's default cell style (Calibri 11) and are only used when
 * the captured state has no style at all.
 */
export const DEFAULT_CELL_STYLE_PATCH: Readonly<Record<string, unknown>> = Object.freeze({
  alignment: {
    horizontal: "general",
    indent: 0,
    readingOrder: "contextDependent",
    rotation: 0,
    shrinkToFit: false,
    vertical: "bottom",
    wrapText: false,
  },
  border: { diagonalDirection: "none" },
  fill: { fillType: "none" },
  font: {
    bold: false,
    color: { colorType: "auto", hex: "000000" },
    italic: false,
    name: "Calibri",
    size: 11,
    strikethrough: false,
    underline: "none",
    verticalAlign: "baseline",
  },
  numberFormat: { formatType: "general" },
  protection: { hidden: false, locked: true },
});

function asRecord(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" ? value as Record<string, unknown> : null;
}

const BRACKET_CURRENCY = /\[[^\]]*[€£¥₹₽₩₪¢₫₴₦₱₸₺฿][^\]]*\]/;

function hasCurrencySymbol(code: string): boolean {
  if (BRACKET_CURRENCY.test(code)) {
    return true;
  }
  return CURRENCY_SYMBOL_TEST.test(code.replace(/\[[^\]]*\]/g, ""));
}

/** Resolve a number-format descriptor (custom string or built-in id) to its code. */
export function resolveNumberFormatCode(format: unknown): string {
  const record = asRecord(format);
  if (!record) {
    return "";
  }

  const formatString = typeof record.formatString === "string" ? record.formatString.trim() : "";
  if (formatString && !/^general$/i.test(formatString)) {
    return formatString;
  }

  const id = typeof record.id === "number" && Number.isFinite(record.id) ? record.id : null;
  if (id !== null && id !== 0) {
    return BUILTIN_NUMBER_FORMAT_CODES[id] ?? formatString;
  }

  if (formatString || record.formatType === "general") {
    return "General";
  }

  return "";
}

/**
 * Remove quoted literals, escaped characters, colour/condition/locale
 * sections and padding directives so only format codes remain. This is what
 * keeps `#,##0.00_);[Red](#,##0.00)` from looking like a date.
 */
export function stripFormatLiterals(code: string): string {
  return code
    .replace(/"[^"]*"/g, "")
    .replace(/\\./g, "")
    .replace(/\[[^\]]*\]/g, "")
    .replace(/[_*]./g, "")
    .toLowerCase();
}

/** Classify a format code into the category the ribbon and coercion rules need. */
export function classifyNumberFormatCode(code: string): NumberFormatCategory {
  const trimmed = code.trim();
  if (!trimmed || /^general$/i.test(trimmed)) {
    return "general";
  }

  // Currency symbols stay visible inside `[$€-407]` sections, so bracket
  // sections are checked directly. `[$-409]`-style sections only carry a
  // locale id and must not be read as a dollar format.
  if (hasCurrencySymbol(trimmed)) {
    return "currency";
  }

  const stripped = stripFormatLiterals(trimmed);
  if (stripped.includes("%")) {
    return "percent";
  }
  if (stripped.includes("e+") || stripped.includes("e-")) {
    return "scientific";
  }
  if (/[yd]/.test(stripped)) {
    return "date";
  }
  if (/[hs]/.test(stripped)) {
    return "time";
  }
  if (stripped.includes("m")) {
    return "date";
  }
  if (/[0#?]\s*\/\s*[0#?]/.test(stripped)) {
    return "fraction";
  }
  if (!stripped) {
    return "general";
  }
  if (!/[0#?]/.test(stripped)) {
    return stripped.includes("@") ? "text" : "general";
  }

  return "number";
}

/** Classify a resolved style's number format. */
export function classifyNumberFormat(format: unknown): NumberFormatCategory {
  return classifyNumberFormatCode(resolveNumberFormatCode(format));
}

function prefersDayFirst(code: string): boolean {
  const stripped = stripFormatLiterals(code);
  const firstDay = stripped.search(/d/);
  const firstMonth = stripped.search(/m/);
  if (firstDay < 0 || firstMonth < 0) {
    return false;
  }
  return firstDay < firstMonth;
}

/** Turn a calendar date into an Excel serial number (1900 or 1904 date system). */
function dateToSerial(
  year: number,
  month: number,
  day: number,
  time: { hour?: number; millisecond?: number; minute?: number; second?: number } = {},
  date1904 = false
): number | null {
  if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) {
    return null;
  }
  if (month < 1 || month > 12 || day < 1 || day > 31 || year < 1900) {
    return null;
  }

  const date = new Date(0);
  date.setUTCHours(0, 0, 0, 0);
  date.setUTCFullYear(year, month - 1, day);
  if (
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() !== month - 1 ||
    date.getUTCDate() !== day
  ) {
    return null;
  }

  const hour = time.hour ?? 0;
  const minute = time.minute ?? 0;
  const second = time.second ?? 0;
  const millisecond = time.millisecond ?? 0;
  const dayFraction = (hour * 3600 + minute * 60 + second + millisecond / 1000) / 86_400;
  const epoch = date1904 ? EXCEL_1904_EPOCH_UTC : EXCEL_1900_EPOCH_UTC;
  return (date.getTime() - epoch) / MS_PER_DAY + dayFraction;
}

function normalizeYear(year: number): number {
  if (year >= 100) {
    return year;
  }
  return year <= 29 ? 2000 + year : 1900 + year;
}

function parseIsoDateTime(text: string) {
  const match = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ](\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?)?\s*(?:Z|[+-]\d{2}:?\d{2})?$/.exec(text);
  if (!match) {
    return null;
  }
  return {
    day: Number(match[3]),
    hour: match[4] === undefined ? undefined : Number(match[4]),
    hasTime: match[4] !== undefined,
    millisecond: match[7] === undefined ? 0 : Number(match[7].padEnd(3, "0")),
    minute: match[5] === undefined ? 0 : Number(match[5]),
    month: Number(match[2]),
    second: match[6] === undefined ? 0 : Number(match[6]),
    year: Number(match[1]),
  };
}

function parseSlashDate(text: string, dayFirst: boolean) {
  const usMatch = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(text);
  if (usMatch) {
    const first = Number(usMatch[1]);
    const second = Number(usMatch[2]);
    let monthFirst = !dayFirst;
    if (first > 12 && second <= 12) {
      monthFirst = false;
    } else if (second > 12 && first <= 12) {
      monthFirst = true;
    }
    return {
      day: monthFirst ? second : first,
      month: monthFirst ? first : second,
      year: normalizeYear(Number(usMatch[3])),
    };
  }

  const isoMatch = /^(\d{4})\/(\d{1,2})\/(\d{1,2})$/.exec(text);
  if (isoMatch) {
    return { day: Number(isoMatch[3]), month: Number(isoMatch[2]), year: Number(isoMatch[1]) };
  }

  return null;
}

function parseTimeOnly(text: string) {
  const match = /^(\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?\s*(am|pm)?$/i.exec(text);
  if (!match) {
    return null;
  }
  let hour = Number(match[1]);
  const minute = Number(match[2]);
  const second = match[3] === undefined ? 0 : Number(match[3]);
  const millisecond = match[4] === undefined ? 0 : Number(match[4].padEnd(3, "0"));
  const meridiem = match[5]?.toLowerCase();
  if (meridiem) {
    if (hour < 1 || hour > 12) {
      return null;
    }
    hour = (hour % 12) + (meridiem === "pm" ? 12 : 0);
  }
  if (hour > 23 || minute > 59 || second > 59) {
    return null;
  }
  return { hour, millisecond, minute, second };
}

type ParsedDateInput = {
  hasDate: boolean;
  hasTime: boolean;
  serial: number;
};

function parseDateInput(text: string, options: Required<Pick<CellInputOptions, "date1904">> & { code: string }): ParsedDateInput | null {
  const iso = parseIsoDateTime(text);
  if (iso) {
    const serial = dateToSerial(
      iso.year,
      iso.month,
      iso.day,
      { hour: iso.hour, millisecond: iso.millisecond, minute: iso.minute, second: iso.second },
      options.date1904
    );
    return serial === null ? null : { hasDate: true, hasTime: iso.hasTime, serial };
  }

  const slash = parseSlashDate(text, prefersDayFirst(options.code));
  if (slash) {
    const serial = dateToSerial(slash.year, slash.month, slash.day, {}, options.date1904);
    return serial === null ? null : { hasDate: true, hasTime: false, serial };
  }

  const time = parseTimeOnly(text);
  if (time) {
    const fraction = (time.hour * 3600 + time.minute * 60 + time.second + time.millisecond / 1000) / 86_400;
    return { hasDate: false, hasTime: true, serial: fraction };
  }

  return null;
}

type ParsedNumber = {
  grouped: boolean;
  percent: boolean;
  scientific: boolean;
  symbol: string;
  value: number;
};

const PLAIN_NUMBER = /^[+-]?(?:\d+\.?\d*|\.\d+)$/;
const GROUPED_NUMBER = /^[+-]?\d{1,3}(?:,\d{3})+(?:\.\d+)?$/;
const SCIENTIFIC_NUMBER = /^[+-]?(?:\d+\.?\d*|\.\d+)[eE][+-]?\d+$/;

function parseNumberInput(raw: string): ParsedNumber | null {
  let text = raw.trim();
  let negative = false;
  let percent = false;
  let symbol = "";

  const parenthesised = /^\((.*)\)$/.exec(text);
  if (parenthesised) {
    negative = true;
    text = parenthesised[1].trim();
  }

  const sign = /^([+-])\s*/.exec(text);
  if (sign) {
    if (sign[1] === "-") {
      negative = !negative;
    }
    text = text.slice(sign[0].length);
  }

  if (/\s*%$/.test(text)) {
    percent = true;
    text = text.replace(/\s*%$/, "");
  }

  const prefixed = CURRENCY_PREFIX.exec(text);
  if (prefixed) {
    symbol = prefixed[1];
    text = text.slice(prefixed[0].length);
  }
  const suffixed = CURRENCY_SUFFIX.exec(text);
  if (!symbol && suffixed) {
    symbol = suffixed[1];
    text = text.slice(0, suffixed.index);
  }

  text = text.trim();
  if (!text) {
    return null;
  }

  const grouped = GROUPED_NUMBER.test(text);
  const normalized = grouped ? text.replace(/,/g, "") : text;
  const scientific = SCIENTIFIC_NUMBER.test(normalized);
  if (!scientific && !PLAIN_NUMBER.test(normalized)) {
    return null;
  }

  const parsed = Number(normalized);
  if (!Number.isFinite(parsed)) {
    return null;
  }

  return {
    grouped,
    percent,
    scientific,
    symbol,
    value: negative ? -parsed : parsed,
  };
}

function buildFormat(formatString: string): XlsxCellNumberFormatInput {
  return { formatString, formatType: "custom" };
}

function suggestNumberFormat(parsed: ParsedNumber): XlsxCellNumberFormatInput | undefined {
  if (parsed.percent) {
    return buildFormat("0.00%");
  }
  if (parsed.symbol) {
    return buildFormat(`${parsed.symbol}#,##0.00`);
  }
  if (parsed.scientific) {
    return buildFormat("0.00E+00");
  }
  if (parsed.grouped) {
    return buildFormat(Number.isInteger(parsed.value) ? "#,##0" : "#,##0.00");
  }
  return undefined;
}

function suggestDateFormat(parsed: ParsedDateInput): XlsxCellNumberFormatInput {
  if (!parsed.hasDate) {
    return buildFormat("h:mm:ss");
  }
  return buildFormat(parsed.hasTime ? "m/d/yyyy h:mm" : "m/d/yyyy");
}

function isCellReferenceBoundary(text: string, start: number, end: number): boolean {
  const before = start > 0 ? text[start - 1] : "";
  const after = end < text.length ? text[end] : "";
  if (before && /[A-Za-z0-9_.$]/.test(before)) {
    return false;
  }
  if (after && /[A-Za-z0-9_(!]/.test(after)) {
    return false;
  }
  return true;
}

function canonicalizeBareSegment(segment: string): string {
  return segment
    .replace(/\s*:\s*/g, ":")
    .replace(/(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/g, (match, prefix: string, letters: string, suffix: string, digits: string, offset: number, whole: string) => {
      if (!isCellReferenceBoundary(whole, offset, offset + match.length)) {
        return match;
      }
      return `${prefix}${letters.toUpperCase()}${suffix}${digits}`;
    });
}

/**
 * Normalize a formula the way Excel does on commit: cell references become
 * uppercase and spaces around range operators are dropped, while quoted
 * strings and quoted sheet names are left untouched (`=sum(a1 : b2)` →
 * `=sum(A1:B2)`).
 */
export function canonicalizeFormula(formula: string): string {
  const trimmed = formula.trim();
  if (!trimmed) {
    return trimmed;
  }

  let result = "";
  let index = 0;
  while (index < trimmed.length) {
    const char = trimmed[index];
    if (char === '"' || char === "'") {
      let end = index + 1;
      while (end < trimmed.length) {
        if (trimmed[end] === char) {
          if (trimmed[end + 1] === char) {
            end += 2;
            continue;
          }
          break;
        }
        end += 1;
      }
      end = Math.min(end + 1, trimmed.length);
      result += trimmed.slice(index, end);
      index = end;
      continue;
    }

    let end = index;
    while (end < trimmed.length && trimmed[end] !== '"' && trimmed[end] !== "'") {
      end += 1;
    }
    result += canonicalizeBareSegment(trimmed.slice(index, end));
    index = end;
  }

  return result;
}

/** Format to apply when a date function is typed into a General cell. */
export function resolveFormulaNumberFormat(
  formula: string,
  targetFormat?: unknown
): XlsxCellNumberFormatInput | undefined {
  if (classifyNumberFormat(targetFormat) !== "general") {
    return undefined;
  }

  const body = formula.trim().replace(/^=/, "").trim();
  const match = /^([A-Za-z][A-Za-z0-9.]*)\s*\(/.exec(body);
  if (!match) {
    return undefined;
  }

  const formatString = DATE_FORMULA_FORMATS[match[1].toUpperCase()];
  return formatString ? buildFormat(formatString) : undefined;
}

/**
 * Decide what a typed or pasted string should become in the target cell.
 *
 * `numberFormat` is the format of the *target* cell. A `@` (text) cell keeps
 * the string as typed, and a General cell receives the format Excel would
 * pick for the recognized value (date, percent, currency, thousands,
 * scientific). Any other format keeps its formatting and only the value is
 * converted.
 */
export function resolveCellInput(text: string, options: CellInputOptions = {}): CellInputResolution {
  const trimmed = text.trim();
  if (!trimmed) {
    return { kind: "value", value: "" };
  }

  if (trimmed.startsWith("'")) {
    return { kind: "value", value: trimmed.slice(1) };
  }

  if (trimmed.startsWith("=") && trimmed.length > 1) {
    const formula = canonicalizeFormula(trimmed);
    return {
      kind: "formula",
      formula,
      numberFormat: resolveFormulaNumberFormat(formula, options.numberFormat),
    };
  }

  const code = resolveNumberFormatCode(options.numberFormat);
  const category = classifyNumberFormatCode(code);
  if (category === "text") {
    return { kind: "value", value: text };
  }

  const dateOptions = { code, date1904: options.date1904 === true };
  const number = parseNumberInput(trimmed);
  if (number) {
    return {
      kind: "value",
      numberFormat: category === "general" ? suggestNumberFormat(number) : undefined,
      value: number.percent ? number.value / 100 : number.value,
    };
  }

  const date = parseDateInput(trimmed, dateOptions);
  if (date) {
    return {
      kind: "value",
      numberFormat: category === "general" ? suggestDateFormat(date) : undefined,
      value: date.serial,
    };
  }

  if (/^(true|false)$/i.test(trimmed)) {
    return { kind: "value", value: trimmed.toLowerCase() === "true" };
  }

  // Lotus 1-2-3 compatibility: a leading sign followed by an expression is a
  // formula, while `-3`/`+1,234` stay numbers (handled above).
  if (/^[+-]/.test(trimmed) && trimmed.length > 1) {
    const formula = canonicalizeFormula(`=${trimmed}`);
    return {
      kind: "formula",
      formula,
      numberFormat: resolveFormulaNumberFormat(formula, options.numberFormat),
    };
  }

  return { kind: "value", value: text };
}

/** Value-only variant of {@link resolveCellInput}; formula text is returned as-is. */
export function coerceEnteredValue(text: string, options: CellInputOptions = {}): unknown {
  const resolution = resolveCellInput(text, options);
  return resolution.kind === "formula" ? text : resolution.value;
}
