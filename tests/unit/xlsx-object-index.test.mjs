import assert from "node:assert/strict"
import test from "node:test"

const xlsx = await import(
  new URL("../../packages/xlsx-core/dist/index.js", import.meta.url).href
)
const interaction = await import(
  new URL("../../packages/office-interaction/dist/index.js", import.meta.url).href
)

const sheets = [
  { id: "sheet-data", index: 0, name: "Data", workbookSheetIndex: 0 },
  { id: "sheet-notes", index: 1, name: "Notes", workbookSheetIndex: 1 },
]
const dataSheet = sheets[0]

function chart(series) {
  return {
    id: "chart-revenue",
    name: "Revenue",
    title: "Revenue over time",
    chartType: "bar",
    sheetIndex: 0,
    workbookSheetIndex: 0,
    zIndex: 3,
    anchor: { kind: "absolute", positionEmu: { x: 0, y: 0 }, sizeEmu: { cx: 1, cy: 1 } },
    axes: [],
    series,
  }
}

function series(name, values) {
  return {
    id: `series-${name}`,
    name,
    values,
    categories: values.map((_value, index) => `Q${index + 1}`),
    dataPoints: [],
  }
}

function context(revision = "xlsx-rev-1", { includeBehaviors = true } = {}) {
  const behaviors = includeBehaviors
    ? xlsx.createXlsxBehaviorRecords({
      sheet: { id: "sheet-data", index: 0, name: "Data" },
      formulas: [{ a1: "C2", formula: "=SUM(B2:B5)", displayValue: "42" }],
      conditionalFormats: [{ appliesToA1: "B2:B5", ruleType: "cellIs", index: 0 }],
      dataValidations: [{ appliesToA1: "B2:B5", validationType: "list", formula1: "$D$1:$D$3", index: 0 }],
      definedNames: [
        { name: "Revenue", refersTo: "Data!$B$2:$B$5", sheet: { id: "sheet-data", index: 0, name: "Data" } },
        { name: "GlobalTotal", refersTo: "Data!$C$2" },
      ],
      protection: { enabled: true, options: { formatCells: false } },
    })
    : []
  return {
    revision: { format: "xlsx", documentId: "forecast.xlsx", revision },
    sheets,
    charts: [chart([series("EMEA", [10, 12, 14, 16]), series("APAC", [4, 6, 8, 10])])],
    tables: [{ name: "SalesTable", displayName: "Sales", sheet: dataSheet, a1: "A1:B5", id: "table-1" }],
    drawings: [
      { id: "image-logo", objectType: "image", sheet: dataSheet, name: "Logo", alt: "Company logo" },
      { id: "shape-note", objectType: "shape", sheet: dataSheet, name: "Callout" },
      { id: "control-check", objectType: "form-control", sheet: dataSheet, name: "Approved" },
    ],
    comments: [{ sheet: dataSheet, cellA1: "B3", text: "Check forecast", author: "Ada" }],
    behaviors,
    getCellSnapshot(_sheet, cell) {
      return cell.row === 1 && cell.col === 1
        ? { displayValue: "42", formula: "=SUM(B2:B5)" }
        : undefined
    },
  }
}

function confirm(draft, referenceId) {
  return interaction.confirmOfficeReferenceDraft(draft, { referenceId, trigger: "pointer" })
}

test("XLSX chart series exemplar confirms with descriptor, fingerprint and reliability", () => {
  const ctx = context()
  const sheet = dataSheet
  const target = { element: "series", seriesIndex: 0 }
  const draft = xlsx.createXlsxChartElementReferenceDraft(ctx, sheet, "chart-revenue", target)
  assert.equal(draft.kind, "chart-series")
  assert.equal(draft.source, "native")
  assert.equal(draft.locator.value.kind, "chart-element")
  assert.equal(draft.locator.value.element, "series")
  assert.equal(draft.locator.value.seriesIndex, 0)
  assert.equal(draft.fingerprint.objectName, "EMEA")
  assert.ok(draft.fingerprint.contentHash.startsWith("fnv1a32:"))
  assert.ok(draft.fingerprint.ancestorKeys.includes("xlsx.chart-id:chart-revenue"))
  assert.equal(draft.reliability.semantic.level, "exact")
  assert.equal(draft.reliability.semantic.reasonCodes[0], "xlsx.native-chart-element")

  const confirmed = confirm(draft, "ref-xlsx-series")
  assert.equal(confirmed.reference.kind, "chart-series")
  assert.deepEqual(confirmed.reference.fingerprint, draft.fingerprint)

  const resolved = xlsx.resolveXlsxReference(ctx, confirmed.reference)
  assert.equal(resolved.status, "exact")
  assert.equal(resolved.descriptor.label, "Series: EMEA")
  assert.equal(resolved.descriptor.draft.kind, "chart-series")
  assert.equal(resolved.descriptor.content, undefined)
  assert.deepEqual(resolved.descriptor.path.map((entry) => entry.kind), [
    "workbook", "worksheet", "chart", "chart-series",
  ])

  const moved = xlsx.resolveXlsxReference(context("xlsx-rev-2"), confirmed.reference)
  assert.equal(moved.status, "relocated")
  assert.equal(moved.reference.document.revision, "xlsx-rev-2")
  assert.equal(moved.descriptor.label, "Series: EMEA")

  const point = xlsx.createXlsxChartElementReferenceDraft(ctx, sheet, "chart-revenue", {
    element: "point",
    seriesIndex: 0,
    pointIndex: 2,
  })
  assert.equal(point.kind, "chart-point")
  const pointDescriptor = xlsx.describeXlsxChartElement(ctx, sheet, ctx.charts[0], {
    element: "point",
    seriesIndex: 0,
    pointIndex: 2,
  })
  assert.equal(pointDescriptor.label, "Point 3")
  assert.equal(pointDescriptor.content.value, "14")

  const legend = xlsx.createXlsxChartElementReferenceDraft(ctx, sheet, "chart-revenue", {
    element: "legend-entry",
    legendIndex: 1,
  })
  assert.equal(legend.kind, "chart-legend-entry")
  assert.throws(
    () => xlsx.createXlsxChartElementReferenceDraft(ctx, sheet, "chart-revenue", { element: "series", seriesIndex: 9 }),
    /series is outside/u,
  )
})

test("XLSX chart hierarchy nests series and points under the chart object", () => {
  const ctx = context()
  const registry = xlsx.createXlsxObjectRegistry()
  const descriptors = registry.list(ctx, { kinds: ["chart", "chart-series", "chart-point"] })
  const chartObject = descriptors.find((descriptor) => descriptor.draft.kind === "chart")
  const chartSeries = descriptors.find((descriptor) =>
    descriptor.draft.kind === "chart-series" && descriptor.label === "Series: EMEA")
  const chartPoint = descriptors.find((descriptor) =>
    descriptor.draft.kind === "chart-point" && descriptor.parentObjectId === chartSeries?.objectId)
  assert.ok(chartObject)
  assert.ok(chartSeries)
  assert.ok(chartPoint)
  assert.equal(chartSeries.parentObjectId, chartObject.objectId)
  assert.equal(chartObject.objectId, "xlsx:chart:sheet-data:chart-revenue")
  assert.deepEqual(
    interaction.officeObjectAncestors(descriptors, chartPoint.objectId).map((entry) => entry.draft.kind),
    ["chart", "chart-series", "chart-point"],
  )
  assert.ok(chartObject.childrenState === "available")
})

test("XLSX tables, drawings and comments are stage 2 objects with resolvable references", () => {
  const ctx = context()
  const sheet = dataSheet

  const tableDraft = xlsx.createXlsxTableReferenceDraft(ctx, ctx.tables[0])
  assert.equal(tableDraft.kind, "table")
  assert.equal(tableDraft.locator.value.tableName, "SalesTable")
  assert.equal(tableDraft.fingerprint.objectName, "SalesTable")
  const tableResolved = xlsx.resolveXlsxReference(ctx, confirm(tableDraft, "ref-table").reference)
  assert.equal(tableResolved.status, "exact")
  assert.equal(tableResolved.descriptor.label, "Sales")

  const imageDraft = xlsx.createXlsxDrawingReferenceDraft(ctx, ctx.drawings[0])
  assert.equal(imageDraft.kind, "image")
  assert.equal(imageDraft.fingerprint.objectName, "Logo")
  const imageResolved = xlsx.resolveXlsxReference(ctx, confirm(imageDraft, "ref-image").reference)
  assert.equal(imageResolved.status, "exact")
  assert.equal(imageResolved.descriptor.label, "Logo")

  assert.equal(xlsx.createXlsxDrawingReferenceDraft(ctx, ctx.drawings[1]).kind, "shape")
  assert.equal(xlsx.createXlsxDrawingReferenceDraft(ctx, ctx.drawings[2]).kind, "form-control")

  const commentDraft = xlsx.createXlsxCommentReferenceDraft(ctx, ctx.comments[0])
  assert.equal(commentDraft.kind, "comment")
  assert.equal(commentDraft.locator.value.cellA1, "B3")
  const commentResolved = xlsx.resolveXlsxReference(ctx, confirm(commentDraft, "ref-comment").reference)
  assert.equal(commentResolved.status, "exact")
  assert.match(commentResolved.descriptor.label, /^Comment: Check forecast$/u)

  assert.throws(
    () => xlsx.createXlsxTableReferenceDraft(context(), { ...ctx.tables[0], name: "Missing" }),
    /not present in the reference context/u,
  )
})

test("XLSX stage 4 behaviors stay invisible references with deterministic ids", () => {
  const ctx = context()
  const records = ctx.behaviors
  assert.deepEqual(records.map((record) => record.behavior), [
    "formula",
    "conditional-format",
    "data-validation",
    "named-range",
    "named-range",
    "protection",
  ])
  const again = xlsx.createXlsxBehaviorRecords({
    sheet: { id: "sheet-data", index: 0, name: "Data" },
    formulas: [{ a1: "C2", formula: "=SUM(B2:B5)", displayValue: "42" }],
    conditionalFormats: [{ appliesToA1: "B2:B5", ruleType: "cellIs", index: 0 }],
    dataValidations: [{ appliesToA1: "B2:B5", validationType: "list", formula1: "$D$1:$D$3", index: 0 }],
    definedNames: [
      { name: "Revenue", refersTo: "Data!$B$2:$B$5", sheet: { id: "sheet-data", index: 0, name: "Data" } },
      { name: "GlobalTotal", refersTo: "Data!$C$2" },
    ],
    protection: { enabled: true, options: { formatCells: false } },
  })
  assert.deepEqual(again.map((record) => record.instanceId), records.map((record) => record.instanceId))

  const formulaRecord = records.find((record) => record.behavior === "formula")
  const draft = xlsx.createXlsxBehaviorReferenceDraft(ctx, formulaRecord)
  assert.equal(draft.kind, "formula")
  assert.equal(draft.locator.value.scope.kind, "sheet")
  assert.equal(draft.locator.value.scope.appliesToA1, "C2")
  const descriptor = xlsx.describeXlsxBehavior(ctx, formulaRecord)
  assert.equal(interaction.isOfficeObjectVisible(descriptor), false)
  assert.equal(descriptor.label, "Formula: =SUM(B2:B5)")

  const confirmed = confirm(draft, "ref-xlsx-formula")
  const resolved = xlsx.resolveXlsxReference(ctx, confirmed.reference)
  assert.equal(resolved.status, "exact")
  assert.equal(resolved.descriptor.draft.kind, "formula")

  const namedRange = records.find((record) => record.behavior === "named-range" && record.facets.name === "GlobalTotal")
  const namedDraft = xlsx.createXlsxBehaviorReferenceDraft(ctx, namedRange)
  assert.equal(namedDraft.locator.value.scope.kind, "workbook")
  const namedResolved = xlsx.resolveXlsxReference(ctx, confirm(namedDraft, "ref-xlsx-name").reference)
  assert.equal(namedResolved.status, "exact")
  assert.equal(namedResolved.descriptor.label, "Named range: GlobalTotal")

  const localized = interaction.localizeOfficeDescriptorLabel(
    resolved.descriptor.label,
    interaction.resolveOfficeMessages("zh-CN").common,
  )
  assert.equal(localized, "公式：=SUM(B2:B5)")

  // A behavior is a reference with scope, never a canvas outline.
  assert.equal("visual" in descriptor, false)
})

test("XLSX object registry registers and lists the built-in families", () => {
  const registry = xlsx.createXlsxObjectRegistry()
  assert.deepEqual(
    registry.providers().map((provider) => provider.id),
    ["xlsx.tables", "xlsx.drawings", "xlsx.comments", "xlsx.charts", "xlsx.behaviors"],
  )
  const ctx = context()
  const all = registry.list(ctx)
  const kinds = new Set(all.map((descriptor) => descriptor.draft.kind))
  for (const kind of ["table", "image", "shape", "form-control", "comment", "chart", "chart-series", "formula"]) {
    assert.ok(kinds.has(kind), `registry is missing ${kind}`)
  }
  const visible = registry.list(ctx, { visibleOnly: true })
  assert.ok(visible.every((descriptor) => descriptor.draft.locator.type !== "format"
    || descriptor.draft.locator.value.kind !== "behavior"))
})

test("XLSX table rows and cells are selectable objects with table/row/cell hierarchy", () => {
  const ctx = context()
  const table = ctx.tables[0]
  const rowDraft = xlsx.createXlsxTableRowReferenceDraft(ctx, table, 1)
  assert.equal(rowDraft.kind, "table-row")
  assert.equal(rowDraft.source, "native")
  assert.equal(rowDraft.locator.value.kind, "table-row")
  assert.equal(rowDraft.locator.value.rowIndex, 1)
  assert.ok(rowDraft.fingerprint.contentHash.startsWith("fnv1a32:"))
  assert.ok(rowDraft.fingerprint.ancestorKeys.includes("xlsx.table:SalesTable"))
  assert.equal(rowDraft.reliability.semantic.level, "exact")
  assert.equal(rowDraft.reliability.relocation.level, "likely")

  const rowConfirmed = confirm(rowDraft, "ref-xlsx-row")
  assert.equal(rowConfirmed.reference.kind, "table-row")
  const rowResolved = xlsx.resolveXlsxReference(ctx, rowConfirmed.reference)
  assert.equal(rowResolved.status, "exact")
  assert.equal(rowResolved.descriptor.label, "Row 2")
  assert.deepEqual(rowResolved.descriptor.path.map((entry) => entry.kind), [
    "workbook", "worksheet", "table", "table-row",
  ])
  assert.ok(rowResolved.descriptor.content.text.includes("42"))
  assert.equal(rowResolved.descriptor.childrenState, "available")

  const cellDraft = xlsx.createXlsxTableCellReferenceDraft(ctx, table, 1, 1)
  assert.equal(cellDraft.kind, "table-cell")
  assert.equal(cellDraft.locator.value.kind, "table-cell")
  assert.equal(cellDraft.locator.value.a1, "B2")
  assert.equal(cellDraft.locator.value.rowIndex, 1)
  assert.equal(cellDraft.locator.value.cellIndex, 1)
  const cellConfirmed = confirm(cellDraft, "ref-xlsx-cell")
  const cellResolved = xlsx.resolveXlsxReference(ctx, cellConfirmed.reference)
  assert.equal(cellResolved.status, "exact")
  assert.equal(cellResolved.descriptor.label, "Cell 2, 2")
  assert.deepEqual(cellResolved.descriptor.path.map((entry) => entry.kind), [
    "workbook", "worksheet", "table", "table-row", "table-cell",
  ])
  assert.deepEqual(cellResolved.descriptor.content, {
    value: "42",
    formula: "=SUM(B2:B5)",
  })

  const moved = xlsx.resolveXlsxReference(context("xlsx-rev-2"), cellConfirmed.reference)
  assert.equal(moved.status, "relocated")
  assert.equal(moved.reference.document.revision, "xlsx-rev-2")
  assert.equal(moved.reference.locator.value.a1, "B2")

  const registry = xlsx.createXlsxObjectRegistry()
  const descriptors = registry.list(ctx, { kinds: ["table", "table-row", "table-cell"] })
  const tableObject = descriptors.find((descriptor) => descriptor.draft.kind === "table")
  const rowObject = descriptors.find((descriptor) =>
    descriptor.draft.kind === "table-row" && descriptor.facets.rowIndex === 1)
  const cellObject = descriptors.find((descriptor) =>
    descriptor.draft.kind === "table-cell" && descriptor.facets.rowIndex === 1 && descriptor.facets.cellIndex === 1)
  assert.ok(tableObject)
  assert.ok(rowObject)
  assert.ok(cellObject)
  assert.equal(rowObject.parentObjectId, tableObject.objectId)
  assert.equal(cellObject.parentObjectId, rowObject.objectId)
  assert.deepEqual(
    interaction.officeObjectAncestors(descriptors, cellObject.objectId).map((entry) => entry.draft.kind),
    ["table", "table-row", "table-cell"],
  )

  assert.throws(
    () => xlsx.createXlsxTableRowReferenceDraft(ctx, table, 99),
    /outside the current table range/u,
  )
  assert.throws(
    () => xlsx.createXlsxTableCellReferenceDraft(ctx, table, 0, 99),
    /outside the current table range/u,
  )
})

test("XLSX overlapping objects rank by series paint order and drawing z-index", () => {
  const ctx = context()
  const registry = xlsx.createXlsxObjectRegistry()
  const seriesDescriptors = registry.list(ctx, { kinds: ["chart-series"] })
  const emea = seriesDescriptors.find((descriptor) => descriptor.label === "Series: EMEA")
  const apac = seriesDescriptors.find((descriptor) => descriptor.label === "Series: APAC")
  assert.ok(emea)
  assert.ok(apac)
  assert.ok(xlsx.xlsxObjectCandidateZIndex(apac) > xlsx.xlsxObjectCandidateZIndex(emea))

  const seriesCandidates = [emea, apac].map((descriptor) => ({
    candidateId: descriptor.objectId,
    descriptor,
    hit: "direct",
    depth: 3,
  }))
  assert.deepEqual(
    xlsx.rankXlsxObjectCandidates(seriesCandidates).map((candidate) => candidate.descriptor.label),
    ["Series: APAC", "Series: EMEA"],
    "later chart series sit above earlier series in Excel paint order",
  )

  const sheet = dataSheet
  const lowSpec = {
    id: "shape-low", objectType: "shape", sheet, name: "Low", zIndex: 2,
  }
  const highSpec = {
    id: "shape-high", objectType: "shape", sheet, name: "High", zIndex: 9,
  }
  const drawingContext = { ...ctx, drawings: [...ctx.drawings, lowSpec, highSpec] }
  const low = xlsx.describeXlsxDrawing(drawingContext, lowSpec)
  const high = xlsx.describeXlsxDrawing(drawingContext, highSpec)
  assert.deepEqual(
    xlsx.rankXlsxObjectCandidates([
      { candidateId: low.objectId, descriptor: low, hit: "direct", depth: 3 },
      { candidateId: high.objectId, descriptor: high, hit: "direct", depth: 3 },
    ]).map((candidate) => candidate.descriptor.label),
    ["High", "Low"],
  )
})
