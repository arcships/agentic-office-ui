/**
 * Locale-aware number and date format support for XLSX cells.
 *
 * Excel format codes can pin a locale inside a section such as `[$-407]` or
 * `[$€-40C]`. This module resolves those sections to a BCP-47 tag and derives
 * the separators, date order and month names that value parsing and display
 * need, using `Intl` as the single source of truth so no per-locale tables can
 * drift out of date.
 */

export type NumberFormatLocale = {
  /** Order of the numeric date parts, e.g. `"dmy"` for `31.12.2024`. */
  dateOrder: "dmy" | "mdy" | "ymd";
  /** Decimal mark used when the value is rendered, e.g. `","`. */
  decimalSeparator: string;
  /** Thousands mark used when the value is rendered, e.g. `"."`. */
  groupSeparator: string;
  /** Excel locale id (LCID) the format asked for, when it named one. */
  localeId: number | null;
  /** Lowercased long/short month names mapped to a 0-based month. */
  monthNames: ReadonlyMap<string, number>;
  /** Resolved BCP-47 tag. */
  tag: string;
};

export type NumberSeparators = {
  decimalSeparator: string;
  groupSeparator: string;
};

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

/** Currency symbols Excel treats as part of a numeric format or typed value. */
export const CURRENCY_SYMBOLS = "$€£¥₹₽₩₪¢₫₴₦₱₸₺฿";

const BRACKET_CURRENCY = new RegExp(`\\[[^\\]]*[${CURRENCY_SYMBOLS.replace("$", "")}][^\\]]*\\]`);
const CURRENCY_SYMBOL_TEST = new RegExp(`[${CURRENCY_SYMBOLS}]`);

function hasCurrencySymbol(code: string): boolean {
  if (BRACKET_CURRENCY.test(code)) {
    return true;
  }
  return CURRENCY_SYMBOL_TEST.test(code.replace(/\[[^\]]*\]/g, ""));
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

/** Classify a format code into the category display and coercion rules need. */
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

/**
 * Excel locale ids (LCIDs) mapped to BCP-47 tags. Formats normally reference a
 * locale as `[$-409]` (US English) or `[$€-407]` (German with the euro sign).
 */
export const BUILTIN_LOCALE_TAGS: Readonly<Record<number, string>> = Object.freeze({
  0x0401: "ar-SA",
  0x0402: "bg-BG",
  0x0403: "ca-ES",
  0x0405: "cs-CZ",
  0x0406: "da-DK",
  0x0407: "de-DE",
  0x0408: "el-GR",
  0x0409: "en-US",
  0x040a: "es-ES",
  0x040b: "fi-FI",
  0x040c: "fr-FR",
  0x040d: "he-IL",
  0x040e: "hu-HU",
  0x040f: "is-IS",
  0x0410: "it-IT",
  0x0411: "ja-JP",
  0x0412: "ko-KR",
  0x0413: "nl-NL",
  0x0414: "nb-NO",
  0x0415: "pl-PL",
  0x0416: "pt-BR",
  0x0418: "ro-RO",
  0x0419: "ru-RU",
  0x041a: "hr-HR",
  0x041b: "sk-SK",
  0x041c: "sq-AL",
  0x041d: "sv-SE",
  0x041e: "th-TH",
  0x041f: "tr-TR",
  0x0421: "id-ID",
  0x0422: "uk-UA",
  0x0424: "sl-SI",
  0x0425: "et-EE",
  0x0426: "lv-LV",
  0x0427: "lt-LT",
  0x042a: "vi-VN",
  0x042d: "eu-ES",
  0x042f: "mk-MK",
  0x0434: "xh-ZA",
  0x0436: "af-ZA",
  0x0438: "fo-FO",
  0x0439: "hi-IN",
  0x043e: "ms-MY",
  0x0441: "sw-KE",
  0x0443: "uz-UZ",
  0x0444: "tt-RU",
  0x0445: "bn-IN",
  0x0446: "pa-IN",
  0x0447: "gu-IN",
  0x0449: "ta-IN",
  0x044a: "te-IN",
  0x044b: "kn-IN",
  0x044c: "ml-IN",
  0x044e: "mr-IN",
  0x0450: "mn-MN",
  0x0456: "gl-ES",
  0x0457: "kok-IN",
  0x045a: "syr-SY",
  0x045b: "si-LK",
  0x0461: "ne-NP",
  0x0462: "fy-NL",
  0x0463: "ps-AF",
  0x0464: "fil-PH",
  0x0465: "dv-MV",
  0x0468: "ha-NG",
  0x046a: "yo-NG",
  0x046b: "quz-BO",
  0x046c: "nso-ZA",
  0x046d: "ba-RU",
  0x046e: "lb-LU",
  0x046f: "kl-GL",
  0x0470: "ig-NG",
  0x0478: "ii-CN",
  0x047a: "arn-CL",
  0x047c: "moh-CA",
  0x047e: "br-FR",
  0x0480: "ug-CN",
  0x0481: "mi-NZ",
  0x0482: "oc-FR",
  0x0483: "co-FR",
  0x0484: "gsw-FR",
  0x0485: "sah-RU",
  0x0486: "qut-GT",
  0x0487: "rw-RW",
  0x0488: "wo-SN",
  0x048c: "prs-AF",
  0x0801: "ur-PK",
  0x0804: "zh-CN",
  0x0807: "de-CH",
  0x0809: "en-GB",
  0x080a: "es-MX",
  0x080c: "fr-BE",
  0x0810: "it-CH",
  0x0813: "nl-BE",
  0x0814: "nn-NO",
  0x0816: "pt-PT",
  0x081a: "sr-Latn-CS",
  0x081d: "sv-FI",
  0x0820: "ur-IN",
  0x082c: "az-Cyrl-AZ",
  0x082e: "dsb-DE",
  0x083b: "se-SE",
  0x083c: "ga-IE",
  0x083e: "ms-BN",
  0x0843: "uz-Cyrl-UZ",
  0x0845: "bn-BD",
  0x0846: "pa-Arab-PK",
  0x0849: "ta-LK",
  0x0850: "mn-Mong-CN",
  0x0851: "bo-CN",
  0x0859: "sd-Arab-PK",
  0x085d: "iu-Latn-CA",
  0x085f: "tzm-Latn-DZ",
  0x0861: "ne-IN",
  0x086b: "quz-EC",
  0x0873: "ti-ER",
  0x0c01: "ar-EG",
  0x0c04: "zh-HK",
  0x0c07: "de-AT",
  0x0c09: "en-AU",
  0x0c0a: "es-ES",
  0x0c0c: "fr-CA",
  0x0c1a: "sr-Cyrl-CS",
  0x0c3b: "se-FI",
  0x0c6b: "quz-PE",
  0x1004: "zh-SG",
  0x1007: "de-LU",
  0x1009: "en-CA",
  0x100a: "es-GT",
  0x100c: "fr-CH",
  0x101a: "hr-BA",
  0x103b: "smj-NO",
  0x1404: "zh-MO",
  0x1407: "de-LI",
  0x1409: "en-NZ",
  0x140a: "es-CR",
  0x140c: "fr-LU",
  0x141a: "bs-Latn-BA",
  0x143b: "smj-SE",
  0x1809: "en-IE",
  0x180a: "es-PA",
  0x180c: "fr-MC",
  0x181a: "sr-Latn-BA",
  0x183b: "sma-NO",
  0x1c09: "en-ZA",
  0x1c0a: "es-DO",
  0x1c1a: "sr-Cyrl-BA",
  0x1c3b: "sma-SE",
  0x2009: "en-JM",
  0x200a: "es-VE",
  0x201a: "bs-Cyrl-BA",
  0x203b: "sms-FI",
  0x2409: "en-029",
  0x240a: "es-CO",
  0x241a: "sr-Latn-RS",
  0x243b: "smn-FI",
  0x2809: "en-BZ",
  0x280a: "es-PE",
  0x281a: "sr-Cyrl-RS",
  0x2c09: "en-TT",
  0x2c0a: "es-AR",
  0x2c1a: "sr-Latn-ME",
  0x3009: "en-ZW",
  0x300a: "es-EC",
  0x301a: "sr-Cyrl-ME",
  0x3409: "en-PH",
  0x340a: "es-CL",
  0x380a: "es-UY",
  0x3c0a: "es-PY",
  0x400a: "es-BO",
  0x440a: "es-SV",
  0x480a: "es-HN",
  0x4c0a: "es-NI",
  0x500a: "es-PR",
});

/** English month names are accepted in every locale, like Excel does. */
const ENGLISH_MONTH_NAMES = [
  "january", "february", "march", "april", "may", "june",
  "july", "august", "september", "october", "november", "december",
  "jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec",
] as const;

const DEFAULT_SEPARATORS: NumberSeparators = Object.freeze({
  decimalSeparator: ".",
  groupSeparator: ",",
});

const localeCache = new Map<string, NumberFormatLocale>();

const ENGLISH_MONTHS: ReadonlyMap<string, number> = new Map(
  ENGLISH_MONTH_NAMES.map((name, index) => [name, index % 12] as const)
);

function stripDiacritics(value: string) {
  return value.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function normalizeMonthName(value: string) {
  return stripDiacritics(value.trim().toLowerCase().replace(/\.+$/, "").replace(/\s+/g, " "));
}

function buildMonthNames(tag: string) {
  const names = new Map<string, number>();
  const long = new Intl.DateTimeFormat(tag, { month: "long", timeZone: "UTC" });
  const short = new Intl.DateTimeFormat(tag, { month: "short", timeZone: "UTC" });
  for (let month = 0; month < 12; month += 1) {
    const date = new Date(Date.UTC(2001, month, 15));
    for (const formatter of [long, short]) {
      const raw = formatter.format(date);
      const normalized = normalizeMonthName(raw);
      if (normalized) {
        names.set(normalized, month);
      }
      const folded = stripDiacritics(raw.trim().toLowerCase().replace(/\.+$/, ""));
      if (folded && !names.has(folded)) {
        names.set(folded, month);
      }
    }
  }
  ENGLISH_MONTH_NAMES.forEach((name, index) => {
    const month = index % 12;
    if (!names.has(name)) {
      names.set(name, month);
    }
  });
  return names;
}

function buildDateOrder(tag: string): NumberFormatLocale["dateOrder"] {
  const parts = new Intl.DateTimeFormat(tag, {
    day: "2-digit",
    month: "2-digit",
    timeZone: "UTC",
    year: "numeric",
  }).formatToParts(new Date(Date.UTC(2001, 10, 22)));
  const order = parts
    .filter((part) => part.type === "day" || part.type === "month" || part.type === "year")
    .map((part) => part.type[0])
    .join("");
  return order === "dmy" || order === "ymd" ? order : "mdy";
}

function buildLocale(tag: string, localeId: number | null): NumberFormatLocale {
  const parts = new Intl.NumberFormat(tag).formatToParts(-1234.5);
  return {
    dateOrder: buildDateOrder(tag),
    decimalSeparator: parts.find((part) => part.type === "decimal")?.value ?? ".",
    groupSeparator: parts.find((part) => part.type === "group")?.value ?? ",",
    localeId,
    monthNames: buildMonthNames(tag),
    tag,
  };
}

function localeForTag(tag: string, localeId: number | null) {
  const cached = localeCache.get(tag);
  if (cached) {
    return cached;
  }
  let locale: NumberFormatLocale;
  try {
    locale = buildLocale(tag, localeId);
  } catch {
    return null;
  }
  localeCache.set(tag, locale);
  return locale;
}

/**
 * Normalize a host-provided locale for number/date formats.
 *
 * Invalid tags fall back to `en-US` instead of throwing, so a bad host value
 * can never break rendering; an empty value selects the default locale.
 */
export function normalizeDocumentLocale(locale: string | null | undefined): string {
  const trimmed = typeof locale === "string" ? locale.trim() : "";
  if (!trimmed) {
    return "en-US";
  }
  try {
    const [canonical] = Intl.getCanonicalLocales(trimmed);
    return canonical ?? "en-US";
  } catch {
    return "en-US";
  }
}

/**
 * Resolve the locale named by a format code's `[$…]` section.
 *
 * `[$-409]` uses an LCID, `[$€-407]` additionally names the currency symbol.
 * System placeholders such as `[$-F800]` fall back to `fallbackLocale`, and a
 * code without any locale section only uses the fallback when one is given.
 */
export function resolveFormatLocale(
  formatCode: string | null | undefined,
  fallbackLocale?: string
): NumberFormatLocale | null {
  // Quoted literals are display text, not format directives.
  const code = String(formatCode ?? "").replace(/"[^"]*"/g, "");
  const sections = /\[\$([^\]]*)\]/.exec(code);
  if (sections) {
    const inner = sections[1] ?? "";
    const parts = inner.split("-");
    const idText = parts.length > 1 ? parts[parts.length - 1] : "";
    const localeId = /^[0-9A-Fa-f]{1,4}$/.test(idText) ? Number.parseInt(idText, 16) : null;
    const tag = localeId === null ? undefined : BUILTIN_LOCALE_TAGS[localeId];
    if (tag) {
      return localeForTag(tag, localeId);
    }
  }

  return fallbackLocale ? localeForTag(fallbackLocale, null) : null;
}

/** Locale separators with US defaults for cells that name no locale. */
export function resolveNumberSeparators(locale: NumberFormatLocale | null | undefined): NumberSeparators {
  if (!locale) {
    return DEFAULT_SEPARATORS;
  }
  return {
    decimalSeparator: locale.decimalSeparator,
    groupSeparator: locale.groupSeparator,
  };
}

type LocaleSeparatorSource = string | NumberFormatLocale | null | undefined;

function separatorsForLocale(locale: LocaleSeparatorSource): NumberSeparators {
  if (typeof locale === "string") {
    return resolveNumberSeparators(resolveFormatLocale(null, locale));
  }
  return resolveNumberSeparators(locale);
}

const LITERAL_TOKEN_START = 0xe000;
const LITERAL_TOKEN_END = 0xf8ff;
const NUMERIC_RUN = /[+-]?\d(?:[\d.,\u00a0\u202f' ]*\d)?/g;

/** Literal text a numeric format prints verbatim (quotes, escapes, currencies). */
function formatLiterals(code: string): string[] {
  const literals: string[] = [];
  for (let index = 0; index < code.length; index += 1) {
    const char = code[index];
    if (char === '"') {
      let value = "";
      let cursor = index + 1;
      while (cursor < code.length) {
        if (code[cursor] === "\\" && cursor + 1 < code.length) {
          value += code[cursor + 1];
          cursor += 2;
          continue;
        }
        if (code[cursor] === '"') {
          if (code[cursor + 1] === '"') {
            value += '"';
            cursor += 2;
            continue;
          }
          break;
        }
        value += code[cursor];
        cursor += 1;
      }
      if (value) {
        literals.push(value);
      }
      index = cursor;
    } else if (char === "\\" && index + 1 < code.length) {
      literals.push(code[index + 1] ?? "");
      index += 1;
    } else if (char === "[") {
      const end = code.indexOf("]", index + 1);
      if (end > index) {
        const inner = code.slice(index + 1, end);
        if (inner.startsWith("$")) {
          const symbol = inner.slice(1).split("-")[0] ?? "";
          if (symbol) {
            literals.push(symbol);
          }
        }
        index = end;
      }
    }
  }
  return [...new Set(literals)].sort((left, right) => right.length - left.length);
}

/** Replace edge literals with private-use tokens so they are never rewritten. */
function maskFormatLiterals(formatted: string, literals: readonly string[]) {
  let masked = formatted;
  const tokens: string[] = [];
  const tokenFor = (literal: string) => {
    if (tokens.length > LITERAL_TOKEN_END - LITERAL_TOKEN_START) {
      return null;
    }
    tokens.push(literal);
    return String.fromCharCode(LITERAL_TOKEN_START + tokens.length - 1);
  };

  for (const literal of literals) {
    if (!literal) continue;
    if (masked.startsWith(literal)) {
      const token = tokenFor(literal);
      if (token) masked = token + masked.slice(literal.length);
    }
    if (masked.endsWith(literal)) {
      const token = tokenFor(literal);
      if (token) masked = masked.slice(0, -literal.length) + token;
    }
  }
  return { masked, tokens };
}

function unmaskFormatLiterals(text: string, tokens: readonly string[]) {
  return text.replace(/[\uE000-\uF8FF]/g, (token) => {
    const index = token.charCodeAt(0) - LITERAL_TOKEN_START;
    return tokens[index] ?? token;
  });
}

/** Rewrite one en-US engine number run with the target locale's separators. */
function localizeNumberRun(run: string, locale: NumberFormatLocale): string {
  let sign = "";
  let body = run;
  if (body.startsWith("+") || body.startsWith("-")) {
    sign = body[0] ?? "";
    body = body.slice(1);
  }
  body = body.replace(/[\s\u00a0\u202f']/g, "");
  if (!/^(?:\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?)$/.test(body)) {
    return run;
  }
  const [integerPart = "", fractionPart] = body.split(".");
  const grouped = integerPart.includes(",");
  const digits = integerPart.replace(/,/g, "");
  const localizedInteger = grouped
    ? digits.replace(/\B(?=(\d{3})+(?!\d))/g, locale.groupSeparator)
    : digits;
  return `${sign}${localizedInteger}${fractionPart === undefined ? "" : `${locale.decimalSeparator}${fractionPart}`}`;
}

/**
 * Localize a formatted cell value for the viewer locale.
 *
 * The workbook engine formats with en-US separators, so numeric output is
 * rewritten to the locale's decimal and group marks. Quoted/escaped format
 * literals (including currency symbols) are masked first, and text, date and
 * time formats are returned untouched so their own punctuation survives.
 */
export function localizeCellNumber(
  formatted: string,
  formatCode: string | null | undefined,
  fallbackLocale?: string
): string {
  if (!formatted) {
    return formatted;
  }
  const code = typeof formatCode === "string" ? formatCode : String(formatCode ?? "");
  const category = classifyNumberFormatCode(code);
  if (
    category !== "number"
    && category !== "currency"
    && category !== "percent"
    && category !== "scientific"
  ) {
    return formatted;
  }
  // A `[$-…]` section wins; otherwise the host locale applies.
  const locale = resolveFormatLocale(code, fallbackLocale);
  if (!locale || (locale.decimalSeparator === "." && locale.groupSeparator === ",")) {
    return formatted;
  }

  const { masked, tokens } = maskFormatLiterals(formatted, formatLiterals(code));
  const rewritten = masked.replace(NUMERIC_RUN, (run) => localizeNumberRun(run, locale));
  return unmaskFormatLiterals(rewritten, tokens);
}

/**
 * Parse a number that uses the separators of `locale`.
 *
 * `parseLocaleNumber("1.234,5")` is 1234.5 in every locale because both marks
 * are present; with a single mark the tag decides (`1.234` is 1.234 in en-US
 * and 1234 in de-DE). Currency and percent markings are not part of the
 * number itself and stay with the caller.
 */
export function parseLocaleNumber(
  text: string,
  locale?: LocaleSeparatorSource
): number | null {
  let source = String(text ?? "").trim();
  if (!source) {
    return null;
  }

  let negative = false;
  const parenthesised = /^\((.*)\)$/.exec(source);
  if (parenthesised) {
    negative = true;
    source = (parenthesised[1] ?? "").trim();
  }

  const sign = /^([+-])\s*/.exec(source);
  if (sign) {
    if (sign[1] === "-") {
      negative = !negative;
    }
    source = source.slice(sign[0].length).trim();
  }

  const normalized = normalizeNumberDigits(source, separatorsForLocale(locale));
  if (!normalized) {
    return null;
  }
  const value = Number(normalized.normalized);
  if (!Number.isFinite(value)) {
    return null;
  }
  return negative ? -value : value;
}

function groupingPattern(separator: string) {
  return new RegExp(`^\\d{1,3}(?:\\${separator}\\d{3})+$`);
}

/**
 * Normalize a numeric string that may use locale separators.
 *
 * Both separators present means the rightmost one is the decimal mark, which
 * is unambiguous for `1.234,56` and `1,234.56`. With a single separator the
 * cell's locale decides; `1.234` is 1.234 in en-US and 1234 in de-DE.
 */
export function normalizeNumberDigits(
  text: string,
  separators: NumberSeparators = DEFAULT_SEPARATORS
): { grouped: boolean; normalized: string } | null {
  const compact = text.replace(/[\s\u00a0\u202f']/g, "");
  if (!compact || !/^\d[\d.,]*$/.test(compact)) {
    return null;
  }

  const marks: Array<{ index: number; mark: string }> = [];
  for (let index = 0; index < compact.length; index += 1) {
    const mark = compact[index];
    if (mark === "." || mark === ",") {
      marks.push({ index, mark });
    }
  }

  if (marks.length === 0) {
    return { grouped: false, normalized: compact };
  }

  const usedMarks = new Set(marks.map((entry) => entry.mark));
  if (usedMarks.size === 2) {
    const decimal = marks[marks.length - 1];
    const integerPart = compact.slice(0, decimal.index).replace(/[.,]/g, "") || "0";
    const fractionPart = compact.slice(decimal.index + 1).replace(/[.,]/g, "");
    if (!/^\d+$/.test(integerPart)) return null;
    if (!fractionPart || !/^\d+$/.test(fractionPart)) return null;
    return { grouped: true, normalized: `${integerPart}.${fractionPart}` };
  }

  const mark = marks[0].mark;
  const isGrouped = groupingPattern(mark).test(compact);
  if (mark === separators.decimalSeparator) {
    if (marks.length > 1) {
      return isGrouped
        ? { grouped: true, normalized: compact.split(mark).join("") }
        : null;
    }
    const [integerPart = "", fractionPart = ""] = compact.split(mark);
    if (!fractionPart || !/^\d+$/.test(fractionPart)) return null;
    return { grouped: false, normalized: `${integerPart || "0"}.${fractionPart}` };
  }

  if (isGrouped) {
    return { grouped: true, normalized: compact.split(mark).join("") };
  }

  if (mark === separators.groupSeparator && marks.length === 1) {
    const [integerPart = "", fractionPart = ""] = compact.split(mark);
    if (!fractionPart || !/^\d+$/.test(fractionPart)) return null;
    return { grouped: false, normalized: `${integerPart || "0"}.${fractionPart}` };
  }

  return null;
}

/** Parse a number that is already rendered as text, e.g. in a grid cell. */
export function parseDisplayedNumber(
  text: string,
  locale?: LocaleSeparatorSource
): number | null {
  const source = String(text ?? "").trim();
  if (!source) {
    return null;
  }
  const cleaned = source
    .replace(/[%％]/g, "")
    .replace(new RegExp(`[${CURRENCY_SYMBOLS}]`, "g"), "")
    .trim();
  return parseLocaleNumber(cleaned, locale);
}

const NUMERIC_DATE_PARTS = /^(\d{1,4})[./-](\d{1,2})[./-](\d{1,4})$/;
const CJK_DATE = /^(\d{4})\s*[年년]\s*(\d{1,2})\s*[月월](?:\s*(\d{1,2})\s*[日일])?$/u;
const DAY_MONTH_NAME = /^(\d{1,2})[.\s/-]*([\p{L}.]+)[\s,./-]+(\d{2,4})$/u;
const MONTH_NAME_DAY = /^([\p{L}.]+)[\s,/-]+(\d{1,2})[\s,./-]+(\d{2,4})$/u;

function normalizeYear(year: number) {
  if (year >= 100) {
    return year;
  }
  return year <= 29 ? 2000 + year : 1900 + year;
}

function resolveMonthIndex(name: string, locale: NumberFormatLocale | null) {
  const normalized = normalizeMonthName(name);
  if (!normalized) {
    return null;
  }
  const fromLocale = locale?.monthNames.get(normalized);
  if (fromLocale !== undefined) {
    return fromLocale;
  }
  return ENGLISH_MONTHS.get(normalized) ?? null;
}

/**
 * Parse a date that uses locale date order or month names, e.g. `31.12.2024`,
 * `15. Januar 2024`, `Jan 15, 2024` or `2024年1月15日`-style numeric order.
 */
export function parseLocaleDate(
  text: string,
  locale: NumberFormatLocale | null,
  dateOrder?: NumberFormatLocale["dateOrder"]
): { day: number; month: number; year: number } | null {
  const normalized = String(text ?? "").trim().replace(/\u00a0|\u202f/g, " ");
  if (!normalized) {
    return null;
  }

  const numeric = NUMERIC_DATE_PARTS.exec(normalized);
  if (numeric) {
    const [, firstText, secondText, thirdText] = numeric;
    const first = Number(firstText);
    const second = Number(secondText);
    const third = Number(thirdText);
    if (firstText.length === 4) {
      return { day: third, month: second, year: first };
    }
    if (thirdText.length !== 4 && thirdText.length !== 2) {
      return null;
    }
    const year = normalizeYear(third);
    const order = dateOrder ?? locale?.dateOrder ?? (first > 12 ? "dmy" : "mdy");
    // Prefer the locale's order, but fall back to the other reading when it
    // cannot be a calendar date (for example `1/15/2024` in a day-first cell).
    const candidates = order === "dmy"
      ? [{ day: first, month: second }, { day: second, month: first }]
      : [{ day: second, month: first }, { day: first, month: second }];
    const match = candidates.find(({ day, month }) => month >= 1 && month <= 12 && day >= 1 && day <= 31);
    return match ? { ...match, year } : null;
  }

  const cjk = CJK_DATE.exec(normalized);
  if (cjk) {
    return { day: cjk[3] === undefined ? 1 : Number(cjk[3]), month: Number(cjk[2]), year: Number(cjk[1]) };
  }

  const dayFirst = DAY_MONTH_NAME.exec(normalized);
  if (dayFirst) {
    const month = resolveMonthIndex(dayFirst[2] ?? "", locale);
    if (month !== null) {
      return { day: Number(dayFirst[1]), month: month + 1, year: normalizeYear(Number(dayFirst[3])) };
    }
  }

  const monthFirst = MONTH_NAME_DAY.exec(normalized);
  if (monthFirst) {
    const month = resolveMonthIndex(monthFirst[1] ?? "", locale);
    if (month !== null) {
      return { day: Number(monthFirst[2]), month: month + 1, year: normalizeYear(Number(monthFirst[3])) };
    }
  }

  return null;
}
