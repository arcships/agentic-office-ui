const MILLISECONDS_PER_DAY = 86_400_000;

export interface ExcelSerialFormatOptions {
  /** Workbook date system. `true` selects the 1904 epoch. */
  date1904?: boolean;
  /** Backwards-compatible alias for `date1904`. */
  use1904?: boolean;
  /** Locale used for date/time labels. Defaults to the viewer locale. */
  locale?: string | string[];
}

type NormalizedFormatCode = {
  code: string;
  /** Format code with quoted text, escapes and non-elapsed brackets removed. */
  tokens: string;
  hasDateTokens: boolean;
  hasTimeTokens: boolean;
  hasElapsedToken: boolean;
};

function normalizeFormatCode(formatCode: string | undefined): NormalizedFormatCode {
  const code = (formatCode ?? "").toLowerCase();
  const tokens = code
    // Keep elapsed-time tokens such as [h], [mm] and [ss]; drop colors,
    // locales and conditions such as [Red] or [$-409].
    .replace(/\[(h+|m+|s+)]/g, " $1 ")
    .replace(/\[[^\]]*]/g, " ")
    .replace(/"[^"]*"/g, " ")
    .replace(/\\[\s\S]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  const hasElapsedToken = /\[(?:h+|m+|s+)]/.test(code);
  const hasHourOrSecond = /[hs]/.test(tokens);
  const hasMinute = /m/.test(tokens);
  const hasDateToken = /[dy]/.test(tokens) || (hasMinute && !hasHourOrSecond);
  const hasTimeToken = hasHourOrSecond || /\[(?:h+|m+|s+)]/.test(tokens);
  return {
    code,
    tokens,
    hasDateTokens: hasDateToken,
    hasTimeTokens: hasTimeToken,
    hasElapsedToken
  };
}

function resolveDate1904(options: ExcelSerialFormatOptions | boolean | undefined) {
  if (typeof options === "boolean") {
    return options;
  }
  return options?.date1904 === true || options?.use1904 === true;
}

function resolveLocale(options: ExcelSerialFormatOptions | boolean | undefined) {
  return typeof options === "object" && options ? options.locale : undefined;
}

export function excelSerialToDate(serial: number, use1904 = false) {
  const wholeDays = Math.trunc(serial);
  const milliseconds = Math.round((serial - wholeDays) * MILLISECONDS_PER_DAY);
  // Excel's 1900 leap-year bug reserves serial 60 for the nonexistent
  // 1900-02-29. Dates before it use the real 1899-12-31 epoch; later dates
  // use the conventional 1899-12-30 epoch so modern serials stay aligned.
  const baseUtc = use1904
    ? Date.UTC(1904, 0, 1)
    : wholeDays < 60
      ? Date.UTC(1899, 11, 31)
      : Date.UTC(1899, 11, 30);
  return new Date(baseUtc + wholeDays * MILLISECONDS_PER_DAY + milliseconds);
}

/**
 * Whether a number format code should be rendered as a date/time value.
 *
 * Quoted literals and formatting directives are ignored so `0.00 "m"` is
 * numeric while `h:mm` and `[h]:mm` are temporal.
 */
export function isDateLikeFormat(formatCode: string | undefined): boolean {
  if (!formatCode) {
    return false;
  }
  const { tokens } = normalizeFormatCode(formatCode);
  return /(?:[dmyhs]|\[h+|\[m+|\[s+)/.test(tokens);
}

/** Alias matching the historical chart-data helper. */
export function isLikelyDateFormatCode(formatCode: string | undefined): boolean {
  return isDateLikeFormat(formatCode);
}

function hasAmPm(code: string) {
  return /am\/pm|a\/p/.test(code);
}

function padNumber(value: number, width: number) {
  return String(Math.trunc(Math.abs(value))).padStart(width, "0");
}

function formatElapsedTime(serial: number, code: string) {
  const totalMilliseconds = Math.round(serial * MILLISECONDS_PER_DAY);
  const sign = totalMilliseconds < 0 ? "-" : "";
  const absolute = Math.abs(totalMilliseconds);
  const totalSeconds = Math.floor(absolute / 1_000);
  const milliseconds = absolute % 1_000;
  const hourToken = /\[(h+)]/.exec(code)?.[1] ?? "";
  const minuteToken = /\[(m+)]/.exec(code)?.[1] ?? "";
  const secondToken = /\[(s+)]/.exec(code)?.[1] ?? "";

  if (hourToken) {
    const hours = Math.floor(totalSeconds / 3_600);
    const minutes = Math.floor(totalSeconds / 60) % 60;
    const seconds = totalSeconds % 60;
    const pieces = [`${sign}${padNumber(hours, hourToken.length)}`];
    if (/m/.test(code)) {
      pieces.push(padNumber(minutes, 2));
    }
    if (/s/.test(code)) {
      pieces.push(padNumber(seconds, 2));
    }
    if (/\.0+/.test(code)) {
      pieces[pieces.length - 1] += `.${padNumber(milliseconds, 3)}`;
    }
    return pieces.join(":");
  }

  if (minuteToken) {
    const minutes = Math.floor(totalSeconds / 60);
    const seconds = totalSeconds % 60;
    const pieces = [`${sign}${padNumber(minutes, minuteToken.length)}`];
    if (/s/.test(code)) {
      pieces.push(padNumber(seconds, 2));
    }
    if (/\.0+/.test(code)) {
      pieces[pieces.length - 1] += `.${padNumber(milliseconds, 3)}`;
    }
    return pieces.join(":");
  }

  if (secondToken) {
    const seconds = totalSeconds;
    const suffix = /\.0+/.test(code) ? `.${padNumber(milliseconds, 3)}` : "";
    return `${sign}${padNumber(seconds, secondToken.length)}${suffix}`;
  }

  return String(serial);
}

function buildDateTimeOptions(normalized: NormalizedFormatCode): Intl.DateTimeFormatOptions {
  const { code, tokens } = normalized;
  const options: Intl.DateTimeFormatOptions = {
    timeZone: "UTC"
  };

  if (normalized.hasDateTokens) {
    if (/yyyy/.test(tokens)) {
      options.year = "numeric";
    } else if (/yy/.test(tokens)) {
      options.year = "2-digit";
    }

    if (/mmmm/.test(tokens)) {
      options.month = "long";
    } else if (/mmm/.test(tokens)) {
      options.month = "short";
    } else if (/m/.test(tokens)) {
      options.month = "numeric";
    }

    if (/dddd/.test(tokens)) {
      options.weekday = "long";
    } else if (/ddd/.test(tokens)) {
      options.weekday = "short";
    }

    if (/dd/.test(tokens)) {
      options.day = "2-digit";
    } else if (/d/.test(tokens)) {
      options.day = "numeric";
    }
  }

  if (normalized.hasTimeTokens) {
    if (/hh/.test(tokens)) {
      options.hour = "2-digit";
    } else if (/h/.test(tokens)) {
      options.hour = "numeric";
    }
    if (/mm/.test(tokens)) {
      options.minute = "2-digit";
    } else if (/m/.test(tokens)) {
      options.minute = "numeric";
    }
    if (/ss/.test(tokens)) {
      options.second = "2-digit";
    } else if (/s/.test(tokens)) {
      options.second = "numeric";
    }
    const fractionalDigits = /\.(0+)/.exec(tokens)?.[1]?.length;
    if (fractionalDigits) {
      options.fractionalSecondDigits = Math.min(3, fractionalDigits) as 1 | 2 | 3;
    }
    if (hasAmPm(code)) {
      options.hour12 = true;
    } else {
      options.hourCycle = "h23";
    }
  }

  if (Object.keys(options).length === 1) {
    options.month = "short";
    options.day = "numeric";
  }

  return options;
}

/**
 * Render an Excel serial date/time using an Excel number format code.
 *
 * The serial is interpreted in UTC and formatted in UTC, so the same workbook
 * produces the same calendar day in every viewer timezone. Time-only formats
 * (`h:mm`, `mm:ss`, `[h]:mm`, ...) do not fall back to an 1899 date label.
 */
export function formatExcelSerial(
  value: number,
  formatCode: string | undefined,
  options: ExcelSerialFormatOptions | boolean = {}
) {
  const use1904 = resolveDate1904(options);
  const normalized = normalizeFormatCode(formatCode);
  const date = excelSerialToDate(value, use1904);

  if (normalized.hasElapsedToken) {
    return formatElapsedTime(value, normalized.code);
  }

  return new Intl.DateTimeFormat(
    resolveLocale(options),
    buildDateTimeOptions(normalized)
  ).format(date);
}

/** Compatibility wrapper kept for chart-data callers. */
export function formatExcelDateSerial(
  value: number,
  formatCode: string | undefined,
  use1904: boolean
) {
  return formatExcelSerial(value, formatCode, { date1904: use1904 });
}
