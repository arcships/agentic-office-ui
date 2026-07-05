import { strFromU8, unzipSync } from "fflate"
import type { XlsxSheetData, XlsxWorkbookTab } from "./types"

export interface ParsedXlsxSheet extends XlsxSheetData {
  cellText: Record<string, string>
}

export interface ParsedXlsxWorkbook {
  sheets: ParsedXlsxSheet[]
  tabs: XlsxWorkbookTab[]
  images: unknown[]
  charts: unknown[]
  tables: unknown[]
}

function parseXml(xml: string): Document {
  const doc = new DOMParser().parseFromString(xml, "application/xml")
  const parserError = doc.querySelector("parsererror")
  if (parserError) throw new Error(parserError.textContent || "Invalid XML")
  return doc
}

function readText(files: Record<string, Uint8Array>, path: string): string | null {
  const data = files[path]
  return data ? strFromU8(data) : null
}

function attr(el: Element, name: string): string {
  return el.getAttribute(name) ?? el.getAttribute(`r:${name}`) ?? ""
}

function columnIndexFromA1(ref: string): number {
  const letters = ref.match(/^[A-Z]+/i)?.[0].toUpperCase() ?? "A"
  let value = 0
  for (const ch of letters) value = value * 26 + (ch.charCodeAt(0) - 64)
  return value - 1
}

function rowIndexFromA1(ref: string): number {
  const n = Number(ref.match(/\d+/)?.[0] ?? "1")
  return Math.max(0, n - 1)
}

function makeEmptySheet(name: string, workbookSheetIndex: number, rowCount = 50, colCount = 26): ParsedXlsxSheet {
  return {
    cachedFormulaValues: {},
    cellText: {},
    colCount,
    colStyleIds: {},
    colWidthOverridesPx: {},
    colWidths: new Array(colCount).fill(80),
    conditionalFormatRules: [],
    dataValidations: [],
    defaultColWidthPx: 80,
    defaultRowHeightPx: 24,
    freezePanes: null,
    hasHorizontalMerges: false,
    hasVerticalMerges: false,
    maxHorizontalMergeEndCol: 0,
    maxUsedCol: Math.max(0, colCount - 1),
    maxUsedRow: Math.max(0, rowCount - 1),
    maxVerticalMergeEndRow: 0,
    minUsedCol: 0,
    minUsedRow: 0,
    name,
    namedCellStyleByName: {},
    rowCount,
    rowHeightOverridesPx: {},
    rowHeights: new Array(rowCount).fill(24),
    rowStyleIds: {},
    showGridLines: true,
    sparklines: [],
    styleById: {},
    tableStyleByName: {},
    themePalette: { colorsByIndex: {} },
    visibility: "visible",
    visibleCols: [],
    visibleRows: [],
    workbookSheetIndex,
    zoomScale: 1,
  }
}

function parseSharedStrings(files: Record<string, Uint8Array>): string[] {
  const xml = readText(files, "xl/sharedStrings.xml")
  if (!xml) return []
  const doc = parseXml(xml)
  return Array.from(doc.getElementsByTagName("si")).map((si) =>
    Array.from(si.getElementsByTagName("t")).map((t) => t.textContent ?? "").join("")
  )
}

function parseSheetRels(files: Record<string, Uint8Array>): Record<string, string> {
  const xml = readText(files, "xl/_rels/workbook.xml.rels")
  if (!xml) return {}
  const doc = parseXml(xml)
  const rels: Record<string, string> = {}
  for (const rel of Array.from(doc.getElementsByTagName("Relationship"))) {
    const id = rel.getAttribute("Id") ?? ""
    const target = rel.getAttribute("Target") ?? ""
    if (id && target) rels[id] = target.startsWith("/") ? target.slice(1) : `xl/${target.replace(/^\.\.\//, "")}`
  }
  return rels
}

function parseWorkbook(files: Record<string, Uint8Array>): Array<{ name: string; path: string }> {
  const xml = readText(files, "xl/workbook.xml")
  if (!xml) throw new Error("Missing xl/workbook.xml")
  const rels = parseSheetRels(files)
  const doc = parseXml(xml)
  return Array.from(doc.getElementsByTagName("sheet")).map((sheet, index) => {
    const id = attr(sheet, "id") || `rId${index + 1}`
    return {
      name: sheet.getAttribute("name") || `Sheet${index + 1}`,
      path: rels[id] || `xl/worksheets/sheet${index + 1}.xml`,
    }
  })
}

function parseWorksheet(files: Record<string, Uint8Array>, path: string, name: string, workbookSheetIndex: number, sharedStrings: string[]): ParsedXlsxSheet {
  const xml = readText(files, path)
  if (!xml) return makeEmptySheet(name, workbookSheetIndex)
  const doc = parseXml(xml)
  const sheet = makeEmptySheet(name, workbookSheetIndex)
  let maxRow = 0
  let maxCol = 0

  for (const c of Array.from(doc.getElementsByTagName("c"))) {
    const ref = c.getAttribute("r") || "A1"
    const row = rowIndexFromA1(ref)
    const col = columnIndexFromA1(ref)
    maxRow = Math.max(maxRow, row)
    maxCol = Math.max(maxCol, col)
    const type = c.getAttribute("t")
    const v = c.getElementsByTagName("v")[0]?.textContent ?? ""
    const inline = c.getElementsByTagName("t")[0]?.textContent ?? ""
    const formula = c.getElementsByTagName("f")[0]?.textContent ?? ""
    let text = inline || v
    if (type === "s") text = sharedStrings[Number(v)] ?? v
    if (!text && formula) text = `=${formula}`
    if (text) sheet.cellText[`${row}:${col}`] = text
  }

  const dimension = doc.getElementsByTagName("dimension")[0]?.getAttribute("ref") ?? ""
  const endRef = dimension.includes(":") ? dimension.split(":").pop()! : dimension
  if (endRef) {
    maxRow = Math.max(maxRow, rowIndexFromA1(endRef))
    maxCol = Math.max(maxCol, columnIndexFromA1(endRef))
  }

  sheet.rowCount = Math.max(20, maxRow + 1)
  sheet.colCount = Math.max(10, maxCol + 1)
  sheet.maxUsedRow = Math.max(0, maxRow)
  sheet.maxUsedCol = Math.max(0, maxCol)
  sheet.colWidths = new Array(sheet.colCount).fill(80)
  sheet.rowHeights = new Array(sheet.rowCount).fill(24)
  return sheet
}

export function parseXlsxBuffer(buffer: ArrayBuffer): ParsedXlsxWorkbook {
  const files = unzipSync(new Uint8Array(buffer))
  const sharedStrings = parseSharedStrings(files)
  const workbookSheets = parseWorkbook(files)
  const sheets = workbookSheets.map((sheet, index) => parseWorksheet(files, sheet.path, sheet.name, index, sharedStrings))
  const tabs = sheets.map((sheet, index): XlsxWorkbookTab => ({
    id: `sheet-${index}`,
    index,
    kind: "sheet",
    name: sheet.name,
    sheetIndex: index,
    visibility: sheet.visibility,
    workbookSheetIndex: index,
  }))
  return { sheets, tabs, images: [], charts: [], tables: [] }
}
