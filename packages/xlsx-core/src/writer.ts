import { strToU8, zipSync } from "fflate"
import type { XlsxSheetData } from "./types"

export interface WritableXlsxSheet extends XlsxSheetData {
  cellText?: Record<string, string>
  cellStyles?: Record<string, Record<string, unknown>>
  mergedRanges?: Array<{ start: { row: number; col: number }; end: { row: number; col: number } }>
}

const XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"

function xmlEscape(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;")
}

function sheetNameEscape(value: string): string {
  return xmlEscape(value.slice(0, 31) || "Sheet")
}

function columnName(col: number): string {
  let n = Math.max(0, Math.trunc(col)) + 1
  let name = ""
  while (n > 0) {
    const rem = (n - 1) % 26
    name = String.fromCharCode(65 + rem) + name
    n = Math.floor((n - 1) / 26)
  }
  return name
}

function cellRef(row: number, col: number): string {
  return `${columnName(col)}${Math.max(0, Math.trunc(row)) + 1}`
}

function parseCellKey(key: string): { row: number; col: number } | null {
  const match = /^(\d+):(\d+)$/.exec(key)
  if (!match) return null
  const row = Number(match[1])
  const col = Number(match[2])
  return Number.isInteger(row) && Number.isInteger(col) && row >= 0 && col >= 0
    ? { row, col }
    : null
}

function styleKey(style: Record<string, unknown> | undefined): string {
  return JSON.stringify(style ?? {})
}

function collectStyles(sheets: WritableXlsxSheet[]): Record<string, number> {
  const ids: Record<string, number> = { "{}": 0 }
  for (const sheet of sheets) {
    for (const style of Object.values(sheet.cellStyles ?? {})) {
      const key = styleKey(style)
      if (ids[key] == null) ids[key] = Object.keys(ids).length
    }
  }
  return ids
}

function styleXml(styleIds: Record<string, number>): string {
  const count = Math.max(1, Object.keys(styleIds).length)
  const xfs = Object.entries(styleIds)
    .sort((a, b) => a[1] - b[1])
    .map(([key, id]) => id === 0 ? `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` : `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment wrapText="${key.includes('wrapText') ? 1 : 0}"/></xf>`)
    .join("")
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="${count}">${xfs}</cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`
}

function worksheetXml(sheet: WritableXlsxSheet, styleIds: Record<string, number>): string {
  const cells = Object.entries(sheet.cellText ?? {})
    .map(([key, value]) => ({ ...parseCellKey(key), value }))
    .filter((entry): entry is { row: number; col: number; value: string } =>
      entry.row != null && entry.col != null && entry.value !== ""
    )
    .sort((left, right) => left.row - right.row || left.col - right.col)

  const rows = new Map<number, Array<{ col: number; value: string }>>()
  for (const cell of cells) {
    const rowCells = rows.get(cell.row) ?? []
    rowCells.push({ col: cell.col, value: cell.value })
    rows.set(cell.row, rowCells)
  }

  const maxRow = Math.max(sheet.maxUsedRow ?? 0, ...cells.map((cell) => cell.row), 0)
  const maxCol = Math.max(sheet.maxUsedCol ?? 0, ...cells.map((cell) => cell.col), 0)
  const dimension = `A1:${cellRef(maxRow, maxCol)}`

  const colXml = Object.entries(sheet.colWidthOverridesPx ?? {})
    .map(([col, width]) => ({ col: Number(col), width }))
    .filter(({ col, width }) => Number.isInteger(col) && col >= 0 && Number.isFinite(width) && width > 0)
    .sort((left, right) => left.col - right.col)
    .map(({ col, width }) => `<col min="${col + 1}" max="${col + 1}" width="${Math.max(1, Number(width) / 7)}" customWidth="1"/>`)
    .join("")

  const mergeXml = (sheet.mergedRanges ?? [])
    .map((range) => `${cellRef(range.start.row, range.start.col)}:${cellRef(range.end.row, range.end.col)}`)
    .map((ref) => `<mergeCell ref="${ref}"/>`)
    .join("")

  const rowXml = Array.from(rows.entries())
    .sort(([left], [right]) => left - right)
    .map(([row, rowCells]) => {
      const cellXml = rowCells
        .sort((left, right) => left.col - right.col)
        .map(({ col, value }) => {
          const ref = cellRef(row, col)
          const styleId = styleIds[styleKey(sheet.cellStyles?.[`${row}:${col}`])] ?? 0
          const styleAttr = styleId ? ` s="${styleId}"` : ""
          if (value.startsWith("=") && value.length > 1) {
            const cached = sheet.cachedFormulaValues?.[`${row}:${col}`]
            const cachedXml = cached ? `<v>${xmlEscape(cached)}</v>` : ""
            return `<c r="${ref}"${styleAttr}><f>${xmlEscape(value.slice(1))}</f>${cachedXml}</c>`
          }
          const numeric = value.trim() !== "" && Number.isFinite(Number(value))
          if (numeric) return `<c r="${ref}"${styleAttr}><v>${xmlEscape(value.trim())}</v></c>`
          return `<c r="${ref}"${styleAttr} t="inlineStr"><is><t>${xmlEscape(value)}</t></is></c>`
        })
        .join("")
      const height = sheet.rowHeightOverridesPx?.[row]
      const rowAttrs = height ? ` r="${row + 1}" ht="${Math.max(1, Number(height) * 0.75)}" customHeight="1"` : ` r="${row + 1}"`
      return `<row${rowAttrs}>${cellXml}</row>`
    })
    .join("")

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="${dimension}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="${sheet.defaultRowHeightPx || 24}"/>
  ${colXml ? `<cols>${colXml}</cols>` : ""}
  <sheetData>${rowXml}</sheetData>
  ${mergeXml ? `<mergeCells count="${(sheet.mergedRanges ?? []).length}">${mergeXml}</mergeCells>` : ""}
</worksheet>`
}

export function writeXlsxWorkbook(sheets: WritableXlsxSheet[]): Uint8Array {
  const normalizedSheets = sheets.length > 0 ? sheets : []
  if (normalizedSheets.length === 0) {
    throw new Error("Cannot export XLSX: workbook has no sheets")
  }

  const styleIds = collectStyles(normalizedSheets)

  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="${XLSX_CONTENT_TYPE}"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
${normalizedSheets.map((_, index) => `  <Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("\n")}
</Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    "xl/workbook.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
${normalizedSheets.map((sheet, index) => `    <sheet name="${sheetNameEscape(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join("\n")}
  </sheets>
</workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${normalizedSheets.map((_, index) => `  <Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`).join("\n")}
  <Relationship Id="rId${normalizedSheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`),
    "xl/styles.xml": strToU8(styleXml(styleIds)),
  }

  normalizedSheets.forEach((sheet, index) => {
    files[`xl/worksheets/sheet${index + 1}.xml`] = strToU8(worksheetXml(sheet, styleIds))
  })

  return zipSync(files, { level: 6 })
}
