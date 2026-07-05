// ---------------------------------------------------------------------------
// XLSX Core Utilities — pure functions with zero React/Vue dependencies
// Utilities aligned with spreadsheet viewer/editor behavior
// ---------------------------------------------------------------------------

import type { XlsxCellAddress, XlsxCellRange, XlsxCellStyleColorInput, XlsxCellBorderEdgeInput } from "./types"

// ── Column Label ───────────────────────────────────────────────────────────

/**
 * Convert a 0-based column index to an Excel column label (A, B, …, Z, AA, …).
 *
 * @example columnLabel(0)  → "A"
 * @example columnLabel(25) → "Z"
 * @example columnLabel(26) → "AA"
 */
export function columnLabel(col: number): string {
  let label = ""
  let value = col

  do {
    label = String.fromCharCode(65 + (value % 26)) + label
    value = Math.floor(value / 26) - 1
  } while (value >= 0)

  return label
}

/**
 * Alias for `columnLabel`.
 */
export const getCellColumnLabel = columnLabel

// ── A1 Notation ────────────────────────────────────────────────────────────

/**
 * Convert a zero-based cell address to A1 notation.
 *
 * @example cellAddressToA1({ col: 0, row: 0 }) → "A1"
 */
export function cellAddressToA1({ col, row }: XlsxCellAddress): string {
  return `${columnLabel(col)}${row + 1}`
}

/**
 * Convert a cell range to A1 notation.
 *
 * @example rangeToA1({ start: {col:0,row:0}, end: {col:2,row:5} }) → "A1:C6"
 */
export function rangeToA1(range: XlsxCellRange): string {
  const n = normalizeCellRange(range)
  return `${cellAddressToA1(n.start)}:${cellAddressToA1(n.end)}`
}

// ── Range Utilities ────────────────────────────────────────────────────────

/**
 * Normalize a cell range so start ≤ end in both row and column dimensions.
 */
export function normalizeCellRange(range: XlsxCellRange): XlsxCellRange {
  return {
    start: {
      row: Math.min(range.start.row, range.end.row),
      col: Math.min(range.start.col, range.end.col),
    },
    end: {
      row: Math.max(range.start.row, range.end.row),
      col: Math.max(range.start.col, range.end.col),
    },
  }
}

/**
 * Return true when two ranges overlap in any cell.
 */
export function rangesIntersect(
  firstRange: XlsxCellRange,
  secondRange: XlsxCellRange
): boolean {
  const first = normalizeCellRange(firstRange)
  const second = normalizeCellRange(secondRange)

  return (
    first.start.row <= second.end.row &&
    first.end.row >= second.start.row &&
    first.start.col <= second.end.col &&
    first.end.col >= second.start.col
  )
}

/**
 * Return true when a range covers exactly one cell.
 */
export function isSingleCellRange(range: XlsxCellRange): boolean {
  const n = normalizeCellRange(range)
  return n.start.row === n.end.row && n.start.col === n.end.col
}

// ── Range Key ──────────────────────────────────────────────────────────────

/**
 * Produce a stable, sortable string key for a cell range.
 */
export function getRangeKey(range: XlsxCellRange): string {
  const n = normalizeCellRange(range)
  return `${n.start.row}:${n.start.col}:${n.end.row}:${n.end.col}`
}

// ── Color Utilities ────────────────────────────────────────────────────────

/**
 * Normalize a hex color string to 6-digit lowercase `#rrggbb` form.
 */
export function normalizeHexColor(value: string, fallback = "#111827"): string {
  const trimmed = value.trim()
  const threeDigit = /^#?([0-9a-f]{3})$/i.exec(trimmed)
  if (threeDigit?.[1]) {
    const [r, g, b] = threeDigit[1].split("") as [string, string, string]
    return `#${r}${r}${g}${g}${b}${b}`.toLowerCase()
  }
  const sixDigit = /^#?([0-9a-f]{6})$/i.exec(trimmed)
  if (sixDigit?.[1]) {
    return `#${sixDigit[1].toLowerCase()}`
  }
  return fallback
}

/**
 * Convert a CSS color string to an XlsxCellStyleColorInput.
 */
export function toXlsxRgbColor(value: string): XlsxCellStyleColorInput {
  return {
    colorType: "rgb" as const,
    hex: normalizeHexColor(value).slice(1).toUpperCase(),
  }
}

// ── String Helpers ─────────────────────────────────────────────────────────

/**
 * Ensure a file name ends with `.xlsx`.
 */
export function ensureWorkbookExtension(fileName: string): string {
  const lower = fileName.toLowerCase()
  return lower.endsWith(".xlsx") || lower.endsWith(".xls")
    ? fileName
    : `${fileName}.xlsx`
}

/**
 * Extract a human-readable workbook name from an optional display name and URL.
 */
export function formatWorkbookName(fileName: string | undefined, url: string): string {
  if (fileName?.trim()) return fileName
  const pathname = url.split("?")[0] ?? ""
  const rawName = pathname.split("/").pop() ?? "workbook.xlsx"
  try {
    return decodeURIComponent(rawName)
  } catch {
    return rawName
  }
}

// ── Search Helpers ─────────────────────────────────────────────────────────

/**
 * Coerce a value to a string for search indexing.
 */
export function normalizeSearchText(value: unknown): string {
  if (typeof value === "string") return value
  if (value === null || value === undefined) return ""
  return String(value)
}

/**
 * Extract searchable text from a @dukelib/sheets-wasm cell value object.
 */
export function cellValueToSearchText(value: unknown): string {
  if (!value || typeof value !== "object") return normalizeSearchText(value)
  const record = value as {
    asBoolean?: () => boolean | null
    asError?: () => string | null
    asNumber?: () => number | null
    asText?: () => string | null
    is_boolean?: boolean
    is_empty?: boolean
    is_error?: boolean
    is_number?: boolean
    is_text?: boolean
  }
  if (record.is_empty) return ""
  if (record.is_error) return record.asError?.() ?? ""
  if (record.is_text) return record.asText?.() ?? ""
  if (record.is_number) return normalizeSearchText(record.asNumber?.())
  if (record.is_boolean) return record.asBoolean?.() ? "TRUE" : "FALSE"
  return normalizeSearchText(value)
}

/**
 * Case-insensitive cell match against a search query.
 */
export function cellMatchesQuery(
  displayValue: string,
  formula: string,
  query: string
): boolean {
  const q = query.toLowerCase()
  return displayValue.toLowerCase().includes(q) || formula.toLowerCase().includes(q)
}

// ── Border Utilities ───────────────────────────────────────────────────────

/**
 * Create a thin border edge styled with the given color.
 */
export function createBorderEdge(color = "#111827"): XlsxCellBorderEdgeInput {
  return { color: toXlsxRgbColor(color), style: "thin" }
}

/**
 * Compute the sub-range corresponding to a single border edge.
 */
export function getBorderEdgeRange(
  range: XlsxCellRange,
  edgeKey: "top" | "bottom" | "left" | "right"
): XlsxCellRange {
  const n = normalizeCellRange(range)
  switch (edgeKey) {
    case "top":
      return { start: n.start, end: { row: n.start.row, col: n.end.col } }
    case "bottom":
      return { start: { row: n.end.row, col: n.start.col }, end: n.end }
    case "left":
      return { start: n.start, end: { row: n.end.row, col: n.start.col } }
    case "right":
      return { start: { row: n.start.row, col: n.end.col }, end: n.end }
  }
}
