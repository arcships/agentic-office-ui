/**
 * CSV parsing, serialization, search and viewport helpers for `CsvViewer`.
 *
 * The upstream Extend block hands parsing to `papaparse` and rendering to
 * `@glideapps/glide-data-grid`. This module keeps the same observable
 * behaviour (header row, quoted cells, greedy blank-line skipping, substring
 * search) without another canvas grid runtime or parser dependency:
 *
 * - RFC 4180 quoting, doubled quotes and newlines inside quoted cells.
 * - UTF-8 (with or without BOM), UTF-16LE and UTF-16BE buffers.
 * - Delimiter auto-detection for CSV, TSV, semicolon and pipe files.
 * - Row/column padding so uneven records stay renderable.
 *
 * The helpers are pure so hosts and tests can reuse them without mounting the
 * component.
 */

/** A single-character field separator or the auto-detection sentinel. */
export type CsvDelimiter = "," | "\t" | ";" | "|" | "auto"

/** Parsing failures that a caller can turn into a localized message. */
export type CsvParseErrorCode = "INVALID_DELIMITER" | "UNTERMINATED_QUOTE"

export interface CsvParseError {
  code: CsvParseErrorCode
  /** Optional diagnostic detail (character offset, offending delimiter). */
  detail?: string
}

export interface CsvParseResult {
  /** Header titles, trimmed. Empty titles stay empty for the renderer to label. */
  headers: string[]
  /** Data rows, padded to `headers.length`. */
  rows: string[][]
  /** Resolved field separator (always a single character). */
  delimiter: string
  /** Number of cells in the physical header record. */
  sourceHeaderCount: number
  /** Number of columns after padding uneven records. */
  columnCount: number
  error: CsvParseError | null
}

export interface ParseCsvOptions {
  /** Explicit separator, or `auto` (default). `tab` and `\\t` are aliases. */
  delimiter?: CsvDelimiter | string
}

export interface SerializeCsvOptions {
  delimiter?: string
  /** Record separator. Defaults to `\r\n`, matching common CSV exporters. */
  newline?: "\n" | "\r\n"
}

export interface CsvVisibleRangeOptions {
  scrollTop: number
  viewportHeight: number
  rowCount: number
  rowHeight: number
  /** Extra rows rendered above and below the viewport. Defaults to 6. */
  overscan?: number
}

export interface CsvVisibleRange {
  start: number
  end: number
  /** Height of the spacer row above the rendered window. */
  padTop: number
  /** Height of the spacer row below the rendered window. */
  padBottom: number
}

export interface CsvMatch {
  row: number
  col: number
  value: string
  /** Spreadsheet address, for example `C4`. */
  address: string
}

export interface FindCsvMatchesOptions {
  /** Cells scanned between event-loop yields. Defaults to 500 rows. */
  batchSize?: number
  /** Return early when this reports true (new query, unmounted host). */
  shouldCancel?: () => boolean
  /** Disable event-loop yields for small inputs and synchronous callers. */
  yieldBetweenBatches?: boolean
}

const DEFAULT_DELIMITERS: readonly string[] = [",", "\t", ";", "|"]
const UTF16LE_BOM = [0xff, 0xfe] as const
const UTF16BE_BOM = [0xfe, 0xff] as const

/** Decode a CSV/TSV buffer using its BOM, defaulting to UTF-8. */
export function decodeCsvBuffer(input: ArrayBuffer | Uint8Array): string {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input)
  if (startsWith(bytes, UTF16LE_BOM)) {
    return new TextDecoder("utf-16le").decode(bytes.subarray(2))
  }
  if (startsWith(bytes, UTF16BE_BOM)) {
    return new TextDecoder("utf-16be").decode(bytes.subarray(2))
  }
  return new TextDecoder("utf-8").decode(bytes)
}

function startsWith(bytes: Uint8Array, prefix: readonly number[]): boolean {
  if (bytes.byteLength < prefix.length) return false
  return prefix.every((value, index) => bytes[index] === value)
}

/** Normalize a caller-supplied delimiter; returns `undefined` for `auto`. */
export function normalizeCsvDelimiter(
  delimiter: string | undefined,
): string | undefined {
  if (!delimiter || delimiter === "auto") return undefined
  if (delimiter === "tab" || delimiter === "\\t") return "\t"
  if (delimiter.length !== 1) return delimiter
  return delimiter
}

/**
 * Pick the separator that appears most often in the first records, ignoring
 * delimiters inside quoted cells. Ties prefer comma, then tab, semicolon and
 * pipe so an extensionless single-column file stays a CSV.
 */
export function detectCsvDelimiter(text: string): string {
  const counts = new Map<string, number>()
  for (const delimiter of DEFAULT_DELIMITERS) counts.set(delimiter, 0)

  let inQuotes = false
  let recordCount = 0
  let index = 0
  while (index < text.length && recordCount < 8) {
    const character = text[index]
    if (character === '"') {
      if (inQuotes && text[index + 1] === '"') {
        index += 2
        continue
      }
      inQuotes = !inQuotes
      index += 1
      continue
    }
    if (!inQuotes) {
      const count = counts.get(character)
      if (count !== undefined) counts.set(character, count + 1)
      if (character === "\n") recordCount += 1
      else if (character === "\r") {
        if (text[index + 1] === "\n") index += 1
        recordCount += 1
      }
    }
    index += 1
  }

  let best = ","
  let bestCount = counts.get(best) ?? 0
  for (const delimiter of DEFAULT_DELIMITERS) {
    const count = counts.get(delimiter) ?? 0
    if (count > bestCount) {
      best = delimiter
      bestCount = count
    }
  }
  return best
}

/**
 * Parse delimited text. The first non-blank record is the header row; blank
 * records are skipped greedily, matching the reference viewer.
 */
export function parseCsvText(
  text: string,
  options: ParseCsvOptions = {},
): CsvParseResult {
  const normalizedText = stripBom(text)
  const requestedDelimiter = normalizeCsvDelimiter(options.delimiter)
  if (requestedDelimiter !== undefined && requestedDelimiter.length !== 1) {
    return {
      headers: [],
      rows: [],
      delimiter: ",",
      sourceHeaderCount: 0,
      columnCount: 0,
      error: {
        code: "INVALID_DELIMITER",
        detail: requestedDelimiter,
      },
    }
  }

  const delimiter = requestedDelimiter ?? detectCsvDelimiter(normalizedText)
  const records = splitRecords(normalizedText, delimiter)
  if ("error" in records) {
    return {
      headers: [],
      rows: [],
      delimiter,
      sourceHeaderCount: 0,
      columnCount: 0,
      error: records.error,
    }
  }

  const nonBlank = records.rows.filter(
    (row) => !(row.length === 1 && row[0].trim() === ""),
  )
  if (nonBlank.length === 0) {
    return {
      headers: [],
      rows: [],
      delimiter,
      sourceHeaderCount: 0,
      columnCount: 0,
      error: null,
    }
  }

  const headerRecord = nonBlank[0]
  const sourceHeaderCount = headerRecord.length
  const body = nonBlank.slice(1)
  let columnCount = sourceHeaderCount
  for (const row of body) {
    if (row.length > columnCount) columnCount = row.length
  }
  const headers = Array.from({ length: columnCount }, (_, index) =>
    (headerRecord[index] ?? "").trim(),
  )
  const rows = body.map((row) =>
    Array.from({ length: columnCount }, (_, index) => row[index] ?? ""),
  )

  return {
    headers,
    rows,
    delimiter,
    sourceHeaderCount,
    columnCount,
    error: null,
  }
}

function stripBom(text: string): string {
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text
}

function splitRecords(
  text: string,
  delimiter: string,
): { rows: string[][] } | { error: CsvParseError } {
  const rows: string[][] = []
  let row: string[] = []
  let field = ""
  let inQuotes = false
  let quoteOffset = -1
  let index = 0

  const commitField = () => {
    row.push(field)
    field = ""
  }
  const commitRow = () => {
    commitField()
    rows.push(row)
    row = []
  }

  while (index < text.length) {
    const character = text[index]

    if (inQuotes) {
      if (character === '"') {
        if (text[index + 1] === '"') {
          field += '"'
          index += 2
          continue
        }
        inQuotes = false
        index += 1
        continue
      }
      field += character
      index += 1
      continue
    }

    if (character === '"' && field.length === 0) {
      inQuotes = true
      quoteOffset = index
      index += 1
      continue
    }
    if (character === delimiter) {
      commitField()
      index += 1
      continue
    }
    if (character === "\n") {
      commitRow()
      index += 1
      continue
    }
    if (character === "\r") {
      commitRow()
      index += text[index + 1] === "\n" ? 2 : 1
      continue
    }

    field += character
    index += 1
  }

  if (inQuotes) {
    return {
      error: {
        code: "UNTERMINATED_QUOTE",
        detail: String(quoteOffset),
      },
    }
  }
  if (field.length > 0 || row.length > 0) commitRow()
  return { rows }
}

/** Serialize a parsed table back to CSV/TSV with RFC 4180 quoting. */
export function serializeCsv(
  headers: readonly string[],
  rows: readonly (readonly string[])[],
  options: SerializeCsvOptions = {},
): string {
  const delimiter = options.delimiter ?? ","
  if (delimiter.length !== 1) {
    throw new Error("CSV delimiter must be a single character")
  }
  const newline = options.newline ?? "\r\n"
  const lines = [headers, ...rows].map((record) =>
    record.map((value) => escapeCsvCell(value, delimiter)).join(delimiter),
  )
  return lines.join(newline)
}

function escapeCsvCell(value: string, delimiter: string): string {
  const text = value ?? ""
  if (
    text.includes(delimiter) ||
    text.includes('"') ||
    text.includes("\n") ||
    text.includes("\r")
  ) {
    return `"${text.replace(/"/g, '""')}"`
  }
  return text
}

/** Convert a zero-based column index to spreadsheet letters (`A`, `AA`). */
export function csvColumnLetters(column: number): string {
  let remaining = Math.max(0, Math.floor(column)) + 1
  let letters = ""
  while (remaining > 0) {
    const remainder = (remaining - 1) % 26
    letters = String.fromCharCode(65 + remainder) + letters
    remaining = Math.floor((remaining - 1) / 26)
  }
  return letters
}

/** Convert a zero-based row/column pair to an A1 address. */
export function csvCellAddress(row: number, column: number): string {
  return `${csvColumnLetters(column)}${Math.max(0, Math.floor(row)) + 1}`
}

/**
 * Find case-insensitive substring matches in the data rows. Large tables are
 * scanned in batches with an event-loop yield so a 100k-row file does not
 * freeze typing in the search field.
 */
export async function findCsvMatches(
  headers: readonly string[],
  rows: readonly (readonly string[])[],
  rawQuery: string,
  options: FindCsvMatchesOptions = {},
): Promise<CsvMatch[]> {
  const query = rawQuery.trim().toLowerCase()
  if (!query) return []

  const batchSize = Math.max(1, Math.floor(options.batchSize ?? 500))
  const shouldYield = options.yieldBetweenBatches ?? true
  let columnCount = headers.length
  for (const row of rows) {
    if (row.length > columnCount) columnCount = row.length
  }
  const matches: CsvMatch[] = []

  for (let batchStart = 0; batchStart < rows.length; batchStart += batchSize) {
    if (options.shouldCancel?.()) return matches
    const batchEnd = Math.min(batchStart + batchSize, rows.length)
    for (let row = batchStart; row < batchEnd; row += 1) {
      const values = rows[row] ?? []
      for (let col = 0; col < columnCount; col += 1) {
        const value = values[col] ?? ""
        if (!value.toLowerCase().includes(query)) continue
        matches.push({
          row,
          col,
          value,
          address: csvCellAddress(row, col),
        })
      }
    }
    if (shouldYield && batchEnd < rows.length) {
      await new Promise<void>((resolve) => {
        setTimeout(resolve, 0)
      })
    }
  }

  return matches
}

/**
 * Resolve the rendered row window for a fixed-height virtual table.
 * Spacer heights keep the scrollbar proportional to the whole document.
 */
export function resolveCsvVisibleRange(
  options: CsvVisibleRangeOptions,
): CsvVisibleRange {
  const rowCount = Math.max(0, Math.floor(options.rowCount))
  const rowHeight = Math.max(1, options.rowHeight)
  const viewportHeight = Math.max(0, options.viewportHeight)
  const overscan = Math.max(0, Math.floor(options.overscan ?? 6))
  const scrollTop = Math.max(0, options.scrollTop)

  if (rowCount === 0) {
    return { start: 0, end: 0, padTop: 0, padBottom: 0 }
  }

  const firstVisible = Math.floor(scrollTop / rowHeight)
  const visibleCount = Math.ceil(viewportHeight / rowHeight)
  const start = Math.min(
    rowCount,
    Math.max(0, firstVisible - overscan),
  )
  const end = Math.min(
    rowCount,
    Math.max(start, firstVisible + visibleCount + overscan),
  )

  return {
    start,
    end,
    padTop: start * rowHeight,
    padBottom: (rowCount - end) * rowHeight,
  }
}
