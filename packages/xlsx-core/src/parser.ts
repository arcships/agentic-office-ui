import { strFromU8, unzipSync } from "fflate"
import type { XlsxSheetData, XlsxTable, XlsxTableStyleInfo, XlsxWorkbookTab } from "./types"

export interface ParsedXlsxSheet extends XlsxSheetData {
  cellText: Record<string, string>
  mergedRanges?: Array<{ start: { row: number; col: number }; end: { row: number; col: number } }>
}

export interface ParsedXlsxWorkbook {
  sheets: ParsedXlsxSheet[]
  tabs: XlsxWorkbookTab[]
  images: unknown[]
  charts: unknown[]
  tables: XlsxTable[]
}

type XmlNode = { attrs: Record<string, string>; body: string; raw: string }

function readText(files: Record<string, Uint8Array>, path: string): string | null {
  const data = files[path]
  return data ? strFromU8(data) : null
}

function decodeXml(value: string): string {
  return value
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
}

function parseAttrs(source: string): Record<string, string> {
  const attrs: Record<string, string> = {}
  const attrPattern = /([\w:.-]+)\s*=\s*("[^"]*"|'[^']*')/g
  let match: RegExpExecArray | null
  while ((match = attrPattern.exec(source))) {
    attrs[match[1]] = decodeXml(match[2].slice(1, -1))
  }
  return attrs
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
}

function nodes(xml: string, tag: string): XmlNode[] {
  const escaped = escapeRegExp(tag)
  const pattern = new RegExp(`<${escaped}\\b([^>]*)>([\\s\\S]*?)<\\/${escaped}>|<${escaped}\\b([^>]*)\\/>`, "g")
  const out: XmlNode[] = []
  let match: RegExpExecArray | null
  while ((match = pattern.exec(xml))) {
    const attrText = match[1] ?? match[3] ?? ""
    out.push({ attrs: parseAttrs(attrText), body: match[2] ?? "", raw: match[0] })
  }
  return out
}

function firstNode(xml: string, tag: string): XmlNode | null {
  return nodes(xml, tag)[0] ?? null
}

function textOf(xml: string, tag: string): string {
  return decodeXml(nodes(xml, tag).map((node) => node.body).join(""))
}

function attr(attrs: Record<string, string>, name: string): string {
  return attrs[name] ?? attrs[`r:${name}`] ?? ""
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

function parseCellRef(ref: string): { row: number; col: number } {
  return { row: rowIndexFromA1(ref), col: columnIndexFromA1(ref) }
}

function parseRangeRef(ref: string): { start: { row: number; col: number }; end: { row: number; col: number } } {
  const [startRef, endRef = startRef] = ref.replace(/\$/g, "").split(":")
  return { start: parseCellRef(startRef), end: parseCellRef(endRef) }
}

function normalizeRelationshipTarget(basePath: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1)
  const baseParts = basePath.split("/")
  baseParts.pop()
  for (const part of target.split("/")) {
    if (!part || part === ".") continue
    if (part === "..") baseParts.pop()
    else baseParts.push(part)
  }
  return baseParts.join("/")
}

function relsPathFor(partPath: string): string {
  const parts = partPath.split("/")
  const file = parts.pop() ?? partPath
  return `${parts.join("/")}/_rels/${file}.rels`
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
  return nodes(xml, "si").map((si) => nodes(si.body, "t").map((t) => decodeXml(t.body)).join(""))
}

function parseRelationships(files: Record<string, Uint8Array>, relsPath: string, basePath: string): Record<string, { target: string; type: string }> {
  const xml = readText(files, relsPath)
  if (!xml) return {}
  const rels: Record<string, { target: string; type: string }> = {}
  for (const rel of nodes(xml, "Relationship")) {
    const id = rel.attrs.Id ?? ""
    const target = rel.attrs.Target ?? ""
    if (!id || !target) continue
    rels[id] = { target: normalizeRelationshipTarget(basePath, target), type: rel.attrs.Type ?? "" }
  }
  return rels
}

function parseWorkbook(files: Record<string, Uint8Array>): Array<{ name: string; path: string; visibility: "hidden" | "veryHidden" | "visible" }> {
  const xml = readText(files, "xl/workbook.xml")
  if (!xml) throw new Error("Missing xl/workbook.xml")
  const rels = parseRelationships(files, "xl/_rels/workbook.xml.rels", "xl/workbook.xml")
  return nodes(xml, "sheet").map((sheet, index) => {
    const id = attr(sheet.attrs, "id") || `rId${index + 1}`
    const visibility = sheet.attrs.state === "hidden" || sheet.attrs.state === "veryHidden" ? sheet.attrs.state : "visible"
    return {
      name: sheet.attrs.name || `Sheet${index + 1}`,
      path: rels[id]?.target || `xl/worksheets/sheet${index + 1}.xml`,
      visibility,
    }
  })
}

function parseTables(files: Record<string, Uint8Array>, worksheetPath: string, worksheetXml: string, sheetIndex: number): XlsxTable[] {
  const rels = parseRelationships(files, relsPathFor(worksheetPath), worksheetPath)
  const tables: XlsxTable[] = []
  for (const tablePart of nodes(worksheetXml, "tablePart")) {
    const id = attr(tablePart.attrs, "id")
    const target = rels[id]?.target
    const xml = target ? readText(files, target) : null
    if (!xml) continue
    const tableNode = firstNode(xml, "table")
    if (!tableNode) continue
    const reference = tableNode.attrs.ref || "A1:A1"
    const range = parseRangeRef(reference)
    const columns = nodes(xml, "tableColumn").map((column, index) => ({
      id: Number(column.attrs.id ?? index + 1),
      index,
      name: column.attrs.name || `Column${index + 1}`,
    }))
    const styleInfoNode = firstNode(xml, "tableStyleInfo")
    const styleInfo: XlsxTableStyleInfo | undefined = styleInfoNode ? {
      name: styleInfoNode.attrs.name,
      showColumnStripes: styleInfoNode.attrs.showColumnStripes === "1",
      showFirstColumn: styleInfoNode.attrs.showFirstColumn === "1",
      showLastColumn: styleInfoNode.attrs.showLastColumn === "1",
      showRowStripes: styleInfoNode.attrs.showRowStripes === "1",
    } : undefined
    tables.push({
      columns,
      displayName: tableNode.attrs.displayName || tableNode.attrs.name || `Table${tables.length + 1}`,
      end: range.end,
      headerRowCount: Number(tableNode.attrs.headerRowCount ?? 1),
      name: tableNode.attrs.name || tableNode.attrs.displayName || `Table${tables.length + 1}`,
      reference,
      start: range.start,
      styleInfo,
      totalsRowCount: Number(tableNode.attrs.totalsRowCount ?? 0),
      totalsRowShown: tableNode.attrs.totalsRowShown === "1",
    })
  }
  return tables.map((table) => ({ ...table, name: table.name || `Sheet${sheetIndex + 1}Table` }))
}

function parseWorksheet(
  files: Record<string, Uint8Array>,
  path: string,
  name: string,
  workbookSheetIndex: number,
  visibility: "hidden" | "veryHidden" | "visible",
  sharedStrings: string[],
): { sheet: ParsedXlsxSheet; tables: XlsxTable[] } {
  const xml = readText(files, path)
  if (!xml) return { sheet: makeEmptySheet(name, workbookSheetIndex), tables: [] }
  const sheet = makeEmptySheet(name, workbookSheetIndex)
  sheet.visibility = visibility
  let maxRow = 0
  let maxCol = 0

  const pane = firstNode(xml, "pane")
  if (pane) {
    const xSplit = Number(pane.attrs.xSplit ?? 0)
    const ySplit = Number(pane.attrs.ySplit ?? 0)
    if (xSplit > 0 || ySplit > 0) sheet.freezePanes = { col: xSplit, row: ySplit }
  }

  for (const colEl of nodes(xml, "col")) {
    const min = Math.max(1, Number(colEl.attrs.min ?? 1))
    const max = Math.max(min, Number(colEl.attrs.max ?? min))
    const width = Number(colEl.attrs.width ?? 0)
    const hidden = colEl.attrs.hidden === "1"
    if (width > 0) for (let c = min - 1; c <= max - 1; c++) sheet.colWidthOverridesPx[c] = Math.round(width * 7)
    if (hidden) sheet.hiddenCols = [...(sheet.hiddenCols ?? []), ...Array.from({ length: max - min + 1 }, (_, i) => min - 1 + i)]
  }

  for (const rowEl of nodes(xml, "row")) {
    const row = Math.max(0, Number(rowEl.attrs.r ?? 1) - 1)
    const height = Number(rowEl.attrs.ht ?? 0)
    if (height > 0) sheet.rowHeightOverridesPx[row] = Math.round(height / 0.75)
    if (rowEl.attrs.hidden === "1") sheet.hiddenRows = [...(sheet.hiddenRows ?? []), row]
  }

  for (const mergeEl of nodes(xml, "mergeCell")) {
    const ref = mergeEl.attrs.ref ?? ""
    if (!ref.includes(":")) continue
    sheet.mergedRanges ??= []
    const range = parseRangeRef(ref)
    sheet.mergedRanges.push(range)
    sheet.hasHorizontalMerges = sheet.hasHorizontalMerges || range.start.col !== range.end.col
    sheet.hasVerticalMerges = sheet.hasVerticalMerges || range.start.row !== range.end.row
    sheet.maxHorizontalMergeEndCol = Math.max(sheet.maxHorizontalMergeEndCol, range.end.col)
    sheet.maxVerticalMergeEndRow = Math.max(sheet.maxVerticalMergeEndRow, range.end.row)
  }

  for (const c of nodes(xml, "c")) {
    const ref = c.attrs.r || "A1"
    const row = rowIndexFromA1(ref)
    const col = columnIndexFromA1(ref)
    maxRow = Math.max(maxRow, row)
    maxCol = Math.max(maxCol, col)
    const type = c.attrs.t
    const v = textOf(c.body, "v")
    const inline = textOf(c.body, "t")
    const formula = textOf(c.body, "f")
    let text = inline || v
    if (type === "s") text = sharedStrings[Number(v)] ?? v
    if (formula) {
      sheet.cellText[`${row}:${col}`] = `=${formula}`
      if (v) sheet.cachedFormulaValues[`${row}:${col}`] = v
    } else if (text) {
      sheet.cellText[`${row}:${col}`] = text
    }
  }

  for (const cf of nodes(xml, "conditionalFormatting")) {
    const ranges = (cf.attrs.sqref ?? "").split(/\s+/).filter(Boolean).map(parseRangeRef)
    for (const rule of nodes(cf.body, "cfRule")) {
      const priority = Number(rule.attrs.priority ?? 0)
      if (rule.attrs.type === "dataBar") {
        const dataBar = firstNode(rule.body, "dataBar")
        const color = firstNode(dataBar?.body ?? "", "color")?.attrs.rgb
        sheet.conditionalFormatRules.push({
          cfvos: nodes(dataBar?.body ?? "", "cfvo").map((node) => ({ type: node.attrs.type ?? "min", value: node.attrs.val == null ? undefined : Number(node.attrs.val) })),
          color: color ? { rgb: color } : undefined,
          kind: "dataBar",
          priority,
          ranges,
          showValue: dataBar?.attrs.showValue !== "0",
        })
      } else if (rule.attrs.type === "colorScale") {
        const colorScale = firstNode(rule.body, "colorScale")
        sheet.conditionalFormatRules.push({
          cfvos: nodes(colorScale?.body ?? "", "cfvo").map((node) => ({ type: node.attrs.type ?? "min", value: node.attrs.val == null ? undefined : Number(node.attrs.val) })),
          colors: nodes(colorScale?.body ?? "", "color").map((node) => ({ rgb: node.attrs.rgb })),
          kind: "colorScale",
          priority,
          ranges,
        })
      } else if (rule.attrs.type === "iconSet") {
        const iconSet = firstNode(rule.body, "iconSet")
        sheet.conditionalFormatRules.push({
          cfvos: nodes(iconSet?.body ?? "", "cfvo").map((node) => ({ type: node.attrs.type ?? "min", value: node.attrs.val == null ? undefined : Number(node.attrs.val) })),
          icons: [],
          iconSet: iconSet?.attrs.iconSet ?? "3TrafficLights1",
          kind: "iconSet",
          priority,
          ranges,
          reverse: iconSet?.attrs.reverse === "1",
          showValue: iconSet?.attrs.showValue !== "0",
        } as never)
      }
    }
  }

  for (const validations of nodes(xml, "dataValidations")) {
    for (const validation of nodes(validations.body, "dataValidation")) {
      const ranges = (validation.attrs.sqref ?? "").split(/\s+/).filter(Boolean).map(parseRangeRef)
      sheet.dataValidations.push({
        allowBlank: validation.attrs.allowBlank === "1",
        errorMessage: validation.attrs.error,
        errorStyle: validation.attrs.errorStyle,
        inputMessage: validation.attrs.prompt,
        listSource: textOf(validation.body, "formula1"),
        ranges,
        showDropdown: validation.attrs.showDropDown !== "1",
        showErrorAlert: validation.attrs.showErrorMessage !== "0",
        showInputMessage: validation.attrs.showInputMessage === "1",
        validationType: validation.attrs.type ?? "custom",
      })
    }
  }

  for (const sparklineGroup of nodes(xml, "x14:sparklineGroup")) {
    const type = (sparklineGroup.attrs.type === "column" || sparklineGroup.attrs.type === "stacked") ? "column" : sparklineGroup.attrs.type === "winLoss" ? "winLoss" : "line"
    for (const sparkline of nodes(sparklineGroup.body, "x14:sparkline")) {
      const f = textOf(sparkline.body, "xm:f")
      const sqref = textOf(sparkline.body, "xm:sqref")
      if (!sqref) continue
      sheet.sparklines.push({
        range: parseRangeRef(f.includes("!") ? f.split("!").pop() ?? "A1:A1" : f || "A1:A1"),
        target: parseCellRef(sqref),
        type,
      })
    }
  }

  const dimension = firstNode(xml, "dimension")?.attrs.ref ?? ""
  const endRef = dimension.includes(":") ? dimension.split(":").pop()! : dimension
  if (endRef) {
    maxRow = Math.max(maxRow, rowIndexFromA1(endRef))
    maxCol = Math.max(maxCol, columnIndexFromA1(endRef))
  }
  for (const range of sheet.mergedRanges ?? []) {
    maxRow = Math.max(maxRow, range.end.row)
    maxCol = Math.max(maxCol, range.end.col)
  }

  sheet.rowCount = Math.max(20, maxRow + 1)
  sheet.colCount = Math.max(10, maxCol + 1)
  sheet.maxUsedRow = Math.max(0, maxRow)
  sheet.maxUsedCol = Math.max(0, maxCol)
  sheet.colWidths = new Array(sheet.colCount).fill(80).map((fallback, index) => sheet.colWidthOverridesPx[index] ?? fallback)
  sheet.rowHeights = new Array(sheet.rowCount).fill(24).map((fallback, index) => sheet.rowHeightOverridesPx[index] ?? fallback)
  return { sheet, tables: parseTables(files, path, xml, workbookSheetIndex) }
}

export function parseXlsxBuffer(buffer: ArrayBuffer): ParsedXlsxWorkbook {
  const files = unzipSync(new Uint8Array(buffer))
  const sharedStrings = parseSharedStrings(files)
  const workbookSheets = parseWorkbook(files)
  const parsed = workbookSheets.map((sheet, index) => parseWorksheet(files, sheet.path, sheet.name, index, sheet.visibility, sharedStrings))
  const sheets = parsed.map((entry) => entry.sheet)
  const tables = parsed.flatMap((entry, sheetIndex) => entry.tables.map((table) => ({ ...table, workbookSheetIndex: sheetIndex }) as XlsxTable & { workbookSheetIndex: number }))
  const tabs = sheets.map((sheet, index): XlsxWorkbookTab => ({
    id: `sheet-${index}`,
    index,
    kind: "sheet",
    name: sheet.name,
    sheetIndex: index,
    visibility: sheet.visibility,
    workbookSheetIndex: index,
  }))
  return { sheets, tabs, images: [], charts: [], tables }
}
