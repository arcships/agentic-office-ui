import assert from "node:assert/strict"
import { readFileSync } from "node:fs"
import test from "node:test"

const fixtures = JSON.parse(readFileSync(
  new URL("../fixtures/office-references/xlsx-object-revisions.json", import.meta.url),
  "utf8",
))
const xlsx = await import(new URL("../../packages/xlsx-core/dist/index.js", import.meta.url).href)
const interaction = await import(new URL("../../packages/office-interaction/dist/index.js", import.meta.url).href)

function chart(value) {
  return {
    ...value,
    anchor: { kind: "absolute", positionEmu: { x: 0, y: 0 }, sizeEmu: { cx: 1, cy: 1 } },
    axes: [],
    series: [],
  }
}

function context(revision, fixture, values = {}) {
  return {
    revision: { format: "xlsx", documentId: "forecast.xlsx", revision },
    sheets: fixture.sheets,
    charts: fixture.charts.map(chart),
    getCellSnapshot(sheet, cell) {
      return values[`${sheet.id ?? sheet.name}:${cell.row}:${cell.col}`]
    },
  }
}

function reference(draft, referenceId = "ref-xlsx-basic") {
  return interaction.parseOfficeObjectReference({ ...draft, referenceId })
}

const original = context("xlsx-v1", fixtures.original, {
  "sheet-data:0:0": { displayValue: "120", formula: "=SUM(B1:B4)" },
})
const dataSheet = original.sheets[0]

test("XLSX adapter creates valid worksheet, cell, range, row, column, chart, and sheet-region references", () => {
  const worksheet = reference(xlsx.createXlsxWorksheetReferenceDraft(original, dataSheet), "ref-sheet")
  assert.equal(worksheet.locator.value.sheet.sheetId, "sheet-data")

  const cellDraft = xlsx.createXlsxCellReferenceDraft(original, dataSheet, { row: 0, col: 0 })
  assert.equal(cellDraft.kind, "cell")
  assert.equal(cellDraft.locator.value.a1, "A1")
  assert.equal(cellDraft.fingerprint.exactText, "120")
  assert.equal(reference(cellDraft, "ref-cell").kind, "cell")
  cellDraft.reliability.semantic.reasonCodes.push("caller-mutation")
  assert.deepEqual(
    xlsx.createXlsxCellReferenceDraft(original, dataSheet, { row: 0, col: 0 }).reliability.semantic.reasonCodes,
    ["xlsx.native-cell-range"],
  )

  const range = reference(xlsx.createXlsxRangeReferenceDraft(original, dataSheet, {
    start: { row: 4, col: 2 },
    end: { row: 1, col: 0 },
  }), "ref-range")
  assert.equal(range.locator.value.a1, "A2:C5")

  const rows = reference(xlsx.createXlsxRowReferenceDraft(original, dataSheet, 4, 1), "ref-rows")
  assert.equal(rows.kind, "row")
  assert.deepEqual({ start: rows.locator.value.start, end: rows.locator.value.end }, { start: 1, end: 4 })
  assert.equal(xlsx.resolveXlsxReference(original, rows).descriptor.label, "Data!2:5")

  const columns = reference(xlsx.createXlsxColumnReferenceDraft(original, dataSheet, 2, 0), "ref-columns")
  assert.equal(columns.kind, "column")
  assert.deepEqual({ start: columns.locator.value.start, end: columns.locator.value.end }, { start: 0, end: 2 })
  assert.equal(xlsx.resolveXlsxReference(original, columns).descriptor.label, "Data!A:C")

  const region = {
    space: "sheet",
    sheetId: "sheet-data",
    start: { row: 0, col: 0, xOffset: 0.1, yOffset: 0.1 },
    end: { row: 2, col: 3, xOffset: 0.8, yOffset: 0.9 },
  }
  assert.equal(reference(xlsx.createXlsxRegionReferenceDraft(original, dataSheet, region), "ref-region").kind, "region")

  const chartReference = reference(xlsx.createXlsxChartReferenceDraft(
    original,
    dataSheet,
    original.charts[0],
    { fallbackRegion: region },
  ), "ref-chart")
  assert.equal(chartReference.locator.value.id, "chart-revenue")
  assert.equal(chartReference.fallbackRegion.sheetId, "sheet-data")
})

test("XLSX row and column references follow stable worksheet identity", () => {
  const rows = reference(xlsx.createXlsxRowReferenceDraft(original, dataSheet, 2, 4), "ref-rows")
  const columns = reference(xlsx.createXlsxColumnReferenceDraft(original, dataSheet, 1, 3), "ref-columns")
  const moved = context("xlsx-v2", fixtures.reordered)

  const relocatedRows = xlsx.resolveXlsxReference(moved, rows)
  assert.equal(relocatedRows.status, "relocated")
  assert.equal(relocatedRows.descriptor.label, "Revenue Data!3:5")
  assert.deepEqual(relocatedRows.reasonCodes, ["xlsx.revision-changed", "xlsx.sheet-id-match", "xlsx.row-stable"])

  const relocatedColumns = xlsx.resolveXlsxReference(moved, columns)
  assert.equal(relocatedColumns.status, "relocated")
  assert.equal(relocatedColumns.descriptor.label, "Revenue Data!B:D")
  assert.deepEqual(relocatedColumns.reasonCodes, ["xlsx.revision-changed", "xlsx.sheet-id-match", "xlsx.column-stable"])
})

test("XLSX sheet ids keep cells and ranges locatable through rename and tab reorder", () => {
  const selected = reference(xlsx.createXlsxCellReferenceDraft(original, dataSheet, { row: 0, col: 0 }), "ref-cell")
  const exact = xlsx.resolveXlsxReference(original, selected)
  assert.equal(exact.status, "exact")
  assert.deepEqual(exact.descriptor.content, { value: "120", formula: "=SUM(B1:B4)" })

  const moved = context("xlsx-v2", fixtures.reordered, {
    "sheet-data:0:0": { displayValue: "135", formula: "=SUM(B1:B4)" },
  })
  const relocated = xlsx.resolveXlsxReference(moved, selected)
  assert.equal(relocated.status, "relocated")
  assert.equal(relocated.reference.document.revision, "xlsx-v2")
  assert.equal(relocated.reference.locator.value.sheet.name, "Revenue Data")
  assert.equal(relocated.reference.locator.value.sheet.index, 1)
  assert.equal(relocated.descriptor.content.value, "135")
  assert.deepEqual(relocated.reasonCodes, ["xlsx.revision-changed", "xlsx.sheet-id-match", "xlsx.a1-stable"])
})

test("XLSX chart references prefer native ids and use unique name/type evidence after id churn", () => {
  const selected = reference(xlsx.createXlsxChartReferenceDraft(original, dataSheet, original.charts[0]), "ref-chart")
  assert.equal(xlsx.resolveXlsxReference(original, selected).status, "exact")

  const renumbered = xlsx.resolveXlsxReference(context("xlsx-v2", fixtures.renumberedChart), selected)
  assert.equal(renumbered.status, "relocated")
  assert.equal(renumbered.reference.locator.value.id, "chart-revenue-v2")
  assert.deepEqual(renumbered.reasonCodes, [
    "xlsx.revision-changed",
    "xlsx.sheet-id-match",
    "xlsx.chart-name-type-match",
  ])

  const ambiguous = xlsx.resolveXlsxReference(context("xlsx-v2", fixtures.ambiguousChart), selected)
  assert.equal(ambiguous.status, "ambiguous")
  assert.equal(ambiguous.candidates.length, 2)
  assert.deepEqual(ambiguous.reasonCodes, ["xlsx.multiple-chart-matches"])

  assert.deepEqual(
    xlsx.resolveXlsxReference(context("xlsx-v2", { sheets: fixtures.original.sheets, charts: [] }), selected),
    { status: "not-found", reasonCode: "xlsx.chart-not-found" },
  )
})

test("XLSX manual regions follow stable sheet ids and reject missing documents", () => {
  const region = {
    space: "sheet",
    sheetId: "sheet-data",
    start: { row: 0, col: 0, xOffset: 0, yOffset: 0 },
    end: { row: 1, col: 1, xOffset: 1, yOffset: 1 },
  }
  const selected = reference(xlsx.createXlsxRegionReferenceDraft(original, dataSheet, region), "ref-region")
  const relocated = xlsx.resolveXlsxReference(context("xlsx-v2", fixtures.reordered), selected)
  assert.equal(relocated.status, "relocated")
  assert.equal(relocated.descriptor.label, "Region on Revenue Data")

  const missing = context("xlsx-v2", { sheets: [], charts: [] })
  assert.deepEqual(xlsx.resolveXlsxReference(missing, selected), {
    status: "not-found",
    reasonCode: "xlsx.region-sheet-not-found",
  })
})
