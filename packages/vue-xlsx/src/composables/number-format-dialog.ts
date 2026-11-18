import type { XlsxCellNumberFormatInput } from "@arcships/xlsx-core";

export type XlsxNumberFormatCategory =
  | "number"
  | "currency"
  | "percent"
  | "date"
  | "text"
  | "custom";

export type XlsxNumberFormatOptions = {
  category: XlsxNumberFormatCategory;
  /** Decimal places for numeric/currency/percent formats; clamped to 0–10. */
  decimals?: number;
  /** Custom Excel format code, used by the `custom` category. */
  customPattern?: string;
  /** Thousands separator for numeric/currency formats. */
  thousands?: boolean;
};

const MAX_DECIMALS = 10;

export function clampNumberFormatDecimals(value: unknown) {
  const parsed = Math.floor(Number(value));
  if (!Number.isFinite(parsed)) {
    return 2;
  }
  return Math.max(0, Math.min(MAX_DECIMALS, parsed));
}

/**
 * Builds the Excel format code for the Format Cells dialog.
 *
 * The number/currency categories always spell the decimal places out, so the
 * exported XLSX shows exactly the precision the user picked instead of Excel's
 * "General" fallback.
 */
export function buildNumberFormatPattern(options: XlsxNumberFormatOptions) {
  const decimals = clampNumberFormatDecimals(options.decimals);
  const digits = decimals > 0 ? `0.${"0".repeat(decimals)}` : "0";
  const grouped = options.thousands === false ? digits : digits.replace(/^0/, "#,##0");
  switch (options.category) {
    case "currency":
      return `$${grouped}`;
    case "percent":
      return `0${decimals > 0 ? `.${"0".repeat(decimals)}` : ""}%`;
    case "date":
      return "m/d/yyyy";
    case "text":
      return "@";
    case "custom": {
      const pattern = options.customPattern?.trim() ?? "";
      return pattern || "General";
    }
    case "number":
    default:
      return grouped;
  }
}

/** Maps the dialog's choices onto the controller's number-format input. */
export function resolveNumberFormatInput(
  options: XlsxNumberFormatOptions
): XlsxCellNumberFormatInput {
  if (options.category === "custom" && !options.customPattern?.trim()) {
    return { formatType: "general" };
  }
  return { formatString: buildNumberFormatPattern(options), formatType: "custom" };
}

/**
 * Renders a representative value with the chosen format, for the dialog's
 * preview line. Custom codes fall back to the raw code so the user still sees
 * what will be written.
 */
export function buildNumberFormatPreview(options: XlsxNumberFormatOptions) {
  const decimals = clampNumberFormatDecimals(options.decimals);
  const useGrouping = options.thousands !== false;
  switch (options.category) {
    case "currency":
      return new Intl.NumberFormat("en-US", {
        currency: "USD",
        maximumFractionDigits: decimals,
        minimumFractionDigits: decimals,
        style: "currency",
        useGrouping
      }).format(1234.567);
    case "percent":
      return new Intl.NumberFormat("en-US", {
        maximumFractionDigits: decimals,
        minimumFractionDigits: decimals,
        style: "percent"
      }).format(0.12345);
    case "date":
      return "3/14/2026";
    case "text":
      return "Text";
    case "custom":
      return options.customPattern?.trim() || "General";
    case "number":
    default:
      return new Intl.NumberFormat("en-US", {
        maximumFractionDigits: decimals,
        minimumFractionDigits: decimals,
        useGrouping
      }).format(1234.567);
  }
}

/** A1 range reference for the current selection, e.g. `Sheet1!$A$1:$C$3`. */
export function formatDefinedNameRange(sheetName: string, address: string | null | undefined) {
  const range = (address ?? "").trim();
  if (!range) {
    return "";
  }
  const escapedSheet = sheetName.replace(/'/g, "''");
  const [start, end] = range.split(":");
  const absolute = (part: string) => {
    const match = /^(\$?)([A-Za-z]+)(\$?)(\d+)$/.exec((part ?? "").trim());
    return match ? `$${match[2].toUpperCase()}$${match[4]}` : (part ?? "").trim();
  };
  const absoluteRange = end ? `${absolute(start)}:${absolute(end)}` : absolute(start);
  return `'${escapedSheet}'!${absoluteRange}`;
}
