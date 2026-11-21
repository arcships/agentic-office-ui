import assert from "node:assert/strict"
import test from "node:test"

const pptx = await import(
  new URL("../../packages/pptx-core/dist/index.js", import.meta.url).href
)
const interaction = await import(
  new URL("../../packages/office-interaction/dist/index.js", import.meta.url).href
)

const slidePath = "ppt/slides/slide1.xml"

function objectKey(shapeId, groupPath = []) {
  return pptx.createPptxObjectKey({
    slidePath,
    source: "slide",
    shapeId,
    groupPath,
  })
}

const groupKey = objectKey("10")
const groupChildKey = objectKey("11", ["10"])
const tableKey = objectKey("20")
const chartKey = objectKey("30")
const shapeKey = objectKey("2")

function slideFixture() {
  return {
    index: 0,
    hidden: false,
    objects: [
      { key: shapeKey, slidePath, source: "slide", shapeId: "2", groupPath: [], nodeType: "shape", name: "Summary" },
      { key: groupKey, slidePath, source: "slide", shapeId: "10", groupPath: [], nodeType: "group", name: "Group 1" },
      { key: groupChildKey, slidePath, source: "slide", shapeId: "11", groupPath: ["10"], nodeType: "shape", name: "Child" },
      { key: tableKey, slidePath, source: "slide", shapeId: "20", groupPath: [], nodeType: "table", name: "Forecast table" },
      { key: chartKey, slidePath, source: "slide", shapeId: "30", groupPath: [], nodeType: "chart", name: "Revenue chart" },
    ],
    morphFromPrevious: [],
    nodes: {
      "anim-1": {
        id: "anim-1",
        container: "main",
        kind: "effect",
        delayMs: 0,
        durationMs: 500,
        autoReverse: false,
        fill: "hold",
        restart: "never",
        acceleration: 0,
        deceleration: 0,
        conditions: [],
        childIds: [],
        effect: {
          id: "effect-1",
          kind: "entrance",
          targetObjectKey: shapeKey,
          presetId: "fade",
          values: {},
        },
      },
    },
    media: [],
    transition: { kind: "fade", durationMs: 700, advanceOnClick: true },
    actions: [{
      id: "action-1",
      sourceObjectKey: shapeKey,
      trigger: "click",
      kind: "go-to-slide",
      targetSlideIndex: 1,
    }],
    capability: { discovered: 0, strict: 0, approximate: 0, static: 0, unparsed: 0, features: [] },
  }
}

const tableSnapshot = {
  rows: [
    ["Region", "Revenue"],
    ["EMEA", "42"],
  ],
}

const chartSnapshot = {
  series: [
    { name: "EMEA", points: ["10", "12", "14"] },
    { name: "APAC", points: ["4", "6", "8"] },
  ],
}

function context(revision = "pptx-rev-1", { withProviders = true, chart = chartSnapshot } = {}) {
  const slide = slideFixture()
  return {
    revision: { format: "pptx", documentId: "stage.pptx", revision },
    document: {
      width: 1280,
      height: 720,
      slides: [{ index: 0, number: 1, hidden: false, notes: "Talk about the forecast" }],
      warnings: [],
    },
    playbackDocument: {
      slides: [slide],
      capability: { discovered: 0, strict: 0, approximate: 0, static: 0, unparsed: 0, features: [] },
    },
    slideIds: [slidePath],
    ...(withProviders
      ? {
          getObjectTable: (_slideIndex, key) => key === tableKey ? tableSnapshot : undefined,
          getObjectChart: (_slideIndex, key) => key === chartKey ? chart : undefined,
        }
      : {}),
  }
}

function confirm(draft, referenceId) {
  return interaction.confirmOfficeReferenceDraft(draft, { referenceId, trigger: "pointer" })
}

test("PPTX table cell exemplar confirms with descriptor, fingerprint and reliability", () => {
  const ctx = context()
  const target = { kind: "table-cell", rowIndex: 1, cellIndex: 1 }
  const draft = pptx.createPptxSubObjectReferenceDraft(ctx, 0, tableKey, target)
  assert.equal(draft.kind, "table-cell")
  assert.equal(draft.source, "native")
  assert.equal(draft.locator.value.kind, "sub-object")
  assert.deepEqual(draft.locator.value.path[0], target)
  assert.ok(draft.fingerprint.contentHash.startsWith("fnv1a32:"))
  assert.ok(draft.fingerprint.ancestorKeys.includes("pptx.sub-object:table-cell"))
  assert.equal(draft.reliability.semantic.level, "exact")
  assert.equal(draft.reliability.relocation.level, "likely")

  const confirmed = confirm(draft, "ref-pptx-cell")
  assert.equal(confirmed.reference.kind, "table-cell")
  assert.deepEqual(confirmed.reference.fingerprint, draft.fingerprint)

  const resolved = pptx.resolvePptxReference(ctx, confirmed.reference)
  assert.equal(resolved.status, "exact")
  assert.equal(resolved.descriptor.label, "Cell 2, 2: 42")
  assert.equal(resolved.descriptor.content.text, "42")
  assert.deepEqual(resolved.descriptor.path.map((entry) => entry.kind), [
    "document", "slide", "table", "table-cell",
  ])

  const moved = pptx.resolvePptxReference(context("pptx-rev-2"), confirmed.reference)
  assert.equal(moved.status, "relocated")
  assert.equal(moved.reference.document.revision, "pptx-rev-2")

  const noEvidence = context("pptx-rev-1", { withProviders: false })
  assert.throws(
    () => pptx.createPptxSubObjectReferenceDraft(noEvidence, 0, tableKey, target),
    /getObjectTable evidence/u,
  )
  assert.throws(
    () => pptx.createPptxSubObjectReferenceDraft(ctx, 0, tableKey, { kind: "table-cell", rowIndex: 9, cellIndex: 0 }),
    /outside the rendered table grid/u,
  )
})

test("PPTX table row exemplar confirms with descriptor, fingerprint and row/cell hierarchy", () => {
  const ctx = context()
  const target = { kind: "table-row", rowIndex: 1 }
  const draft = pptx.createPptxSubObjectReferenceDraft(ctx, 0, tableKey, target)
  assert.equal(draft.kind, "table-row")
  assert.equal(draft.source, "native")
  assert.equal(draft.locator.value.kind, "sub-object")
  assert.deepEqual(draft.locator.value.path[0], target)
  assert.ok(draft.fingerprint.contentHash.startsWith("fnv1a32:"))
  assert.ok(draft.fingerprint.ancestorKeys.includes("pptx.sub-object:table-row"))

  const confirmed = confirm(draft, "ref-pptx-row")
  assert.equal(confirmed.reference.kind, "table-row")
  const resolved = pptx.resolvePptxReference(ctx, confirmed.reference)
  assert.equal(resolved.status, "exact")
  assert.equal(resolved.descriptor.label, "Row 2: EMEA 42")
  assert.equal(resolved.descriptor.content.text, "EMEA\t42")
  assert.equal(resolved.descriptor.childrenState, "available")
  assert.deepEqual(resolved.descriptor.path.map((entry) => entry.kind), [
    "document", "slide", "table", "table-row",
  ])

  const moved = pptx.resolvePptxReference(context("pptx-rev-2"), confirmed.reference)
  assert.equal(moved.status, "relocated")
  assert.equal(moved.reference.document.revision, "pptx-rev-2")

  assert.throws(
    () => pptx.createPptxSubObjectReferenceDraft(ctx, 0, tableKey, { kind: "table-row", rowIndex: 9 }),
    /outside the rendered table grid/u,
  )
})

test("PPTX chart internals consume the #184 host geometry + viewport seam", () => {
  const frame = { width: 500, height: 400 };
  const geometryViewport = { x: 0.2, y: 0.1, width: 0.5, height: 0.4 };
  const geometry = {
    chartId: "chart-revenue",
    provider: "host-echarts",
    format: "pptx",
    space: "chart-local-px",
    frame,
    layout: {
      chartArea: { x: 0, y: 0, width: frame.width, height: frame.height },
      plotArea: { x: 40, y: 80, width: 420, height: 240 },
      confidence: "exact",
      legend: {
        position: "bottom",
        rect: { x: 350, y: 10, width: 100, height: 20 },
        confidence: "exact",
        items: [{
          index: 0,
          seriesIndex: 0,
          label: "EMEA",
          rect: { x: 350, y: 10, width: 100, height: 20 },
          confidence: "exact",
        }],
      },
    },
    axes: [],
    series: [
      {
        index: 0,
        name: "EMEA",
        shape: "bar",
        confidence: "exact",
        points: [{
          index: 0,
          categoryIndex: 0,
          value: 10,
          position: { x: 50, y: 120 },
          rect: { x: 40, y: 100, width: 20, height: 100 },
          confidence: "exact",
          rectConfidence: "exact",
        }],
      },
      { index: 1, name: "APAC", shape: "bar", confidence: "exact", points: [] },
    ],
    exact: true,
  };
  const slideRect = (rect) => ({
    x: geometryViewport.x + (rect.x / frame.width) * geometryViewport.width,
    y: geometryViewport.y + (rect.y / frame.height) * geometryViewport.height,
    width: (rect.width / frame.width) * geometryViewport.width,
    height: (rect.height / frame.height) * geometryViewport.height,
  });
  const assertRectClose = (actual, expected) => {
    for (const key of ["x", "y", "width", "height"]) {
      assert.ok(Math.abs(actual[key] - expected[key]) < 1e-12, `${key}: ${actual[key]} != ${expected[key]}`);
    }
  };
  const ctx = context("pptx-rev-1", {
    chart: { ...chartSnapshot, geometry, geometryViewport },
  });
  const descriptors = pptx.createPptxObjectIndex(ctx).descriptors
  const series = descriptors.find((descriptor) =>
    descriptor.draft.kind === "chart-series" && descriptor.label === "Series: EMEA")
  const point = descriptors.find((descriptor) =>
    descriptor.draft.kind === "chart-point" && descriptor.parentObjectId === series?.objectId)
  const legend = descriptors.find((descriptor) => descriptor.draft.kind === "chart-legend-entry")
  assert.ok(series)
  assert.ok(point)
  assert.ok(legend)
  assert.ok(series.draft.fingerprint.ancestorKeys.includes("pptx.sub-object:chart-series"))
  assert.ok(point.draft.fingerprint.ancestorKeys.includes("pptx.sub-object:chart-point"))
  assert.ok(legend.draft.fingerprint.ancestorKeys.includes("pptx.sub-object:chart-legend-entry"))
  assertRectClose(series.visual.fragments[0].rect, slideRect({ x: 40, y: 100, width: 20, height: 100 }));
  assert.equal(series.visual.fragments[0].zIndex, 1_000);
  assertRectClose(point.visual.fragments[0].rect, slideRect({ x: 40, y: 100, width: 20, height: 100 }));
  assertRectClose(legend.visual.fragments[0].rect, slideRect({ x: 350, y: 10, width: 100, height: 20 }));

  const confirmed = confirm(series.draft, "ref-pptx-series-geometry")
  const resolved = pptx.resolvePptxReference(ctx, confirmed.reference)
  assert.equal(resolved.status, "exact")
  assertRectClose(resolved.descriptor.visual.fragments[0].rect, slideRect({ x: 40, y: 100, width: 20, height: 100 }))

  const withoutGeometry = context()
  const bare = pptx.createPptxObjectIndex(withoutGeometry).descriptors.find((descriptor) =>
    descriptor.draft.kind === "chart-series" && descriptor.label === "Series: EMEA")
  assert.ok(bare)
  assert.equal(bare.visual, undefined, "without the host seam chart internals stay references")
  assert.equal(
    pptx.resolvePptxReference(withoutGeometry, confirm(bare.draft, "ref-pptx-bare").reference).status,
    "exact",
  )
})

test("PPTX stage 3 hierarchy nests table cells and chart internals under their objects", () => {
  const ctx = context()
  const descriptors = pptx.createPptxObjectIndex(ctx).descriptors
  const table = descriptors.find((descriptor) => descriptor.objectId === `pptx:object:${tableKey}`)
  const row = descriptors.find((descriptor) =>
    descriptor.draft.kind === "table-row" && descriptor.content?.text === "Region\tRevenue")
  const cell = descriptors.find((descriptor) => descriptor.draft.kind === "table-cell")
  const chart = descriptors.find((descriptor) => descriptor.objectId === `pptx:object:${chartKey}`)
  const series = descriptors.find((descriptor) =>
    descriptor.draft.kind === "chart-series" && descriptor.label === "Series: EMEA")
  const point = descriptors.find((descriptor) =>
    descriptor.draft.kind === "chart-point" && descriptor.parentObjectId === series?.objectId)
  assert.ok(table)
  assert.ok(row)
  assert.ok(cell)
  assert.ok(chart)
  assert.ok(series)
  assert.ok(point)
  assert.equal(row.parentObjectId, table.objectId)
  assert.equal(cell.parentObjectId, row.objectId)
  assert.equal(series.parentObjectId, chart.objectId)
  const group = descriptors.find((descriptor) =>
    descriptor.draft.kind === "group")
  const groupChild = descriptors.find((descriptor) =>
    descriptor.objectId === `pptx:object:${groupChildKey}`)
  assert.ok(group)
  assert.ok(groupChild)
  assert.equal(groupChild.parentObjectId, group.objectId)
  assert.deepEqual(
    interaction.officeObjectAncestors(descriptors, point.objectId).map((entry) => entry.draft.kind),
    ["slide", "chart", "chart-series", "chart-point"],
  )
  assert.deepEqual(
    interaction.officeObjectAncestors(descriptors, cell.objectId).map((entry) => entry.draft.kind),
    ["slide", "table", "table-row", "table-cell"],
  )
})

test("PPTX stage 4 behaviors stay invisible references with native ids", () => {
  const ctx = context()
  const records = pptx.enumeratePptxBehaviors(ctx, 0)
  assert.deepEqual(records.map((record) => record.behavior).sort(), [
    "action",
    "animation",
    "note",
    "transition",
  ])
  const animation = records.find((record) => record.behavior === "animation")
  const transition = records.find((record) => record.behavior === "transition")
  const action = records.find((record) => record.behavior === "action")
  const note = records.find((record) => record.behavior === "note")
  assert.equal(animation.ownerObjectKey, shapeKey)
  assert.equal(action.ownerObjectKey, shapeKey)
  assert.equal(note.content.text, "Talk about the forecast")

  for (const record of records) {
    const descriptor = interaction.createOfficeObjectDescriptor({
      objectId: `test:${record.instanceId}`,
      draft: pptx.createPptxBehaviorDraftFromRecord(ctx, 0, record),
      label: record.label,
      path: record.path,
    })
    assert.equal(interaction.isOfficeObjectVisible(descriptor), false)
    assert.equal(descriptor.draft.locator.value.kind, "behavior")
  }

  const draft = pptx.createPptxBehaviorReferenceDraft(ctx, 0, {
    behavior: "transition",
    instanceId: transition.instanceId,
  })
  assert.equal(draft.kind, "transition")
  assert.equal(draft.locator.value.nativeId, transition.instanceId)
  const resolved = pptx.resolvePptxReference(ctx, confirm(draft, "ref-pptx-transition").reference)
  assert.equal(resolved.status, "exact")
  assert.equal(resolved.descriptor.label, "Slide transition")
  assert.equal("visual" in resolved.descriptor, false)

  const localized = interaction.localizeOfficeDescriptorLabel(
    resolved.descriptor.label,
    interaction.resolveOfficeMessages("zh-CN").common,
  )
  assert.equal(localized, "幻灯片切换")

  const animationDraft = pptx.createPptxBehaviorReferenceDraft(ctx, 0, {
    behavior: "animation",
    instanceId: animation.instanceId,
  })
  assert.equal(animationDraft.kind, "animation")
  assert.equal(animationDraft.locator.value.ownerObjectKey, shapeKey)
  assert.equal(pptx.resolvePptxReference(ctx, confirm(animationDraft, "ref-pptx-anim").reference).status, "exact")
})

test("PPTX object registry exposes stage 2 and stage 4 families", () => {
  const registry = pptx.createPptxObjectRegistry()
  assert.deepEqual(registry.providers().map((provider) => provider.id), ["pptx.objects", "pptx.behaviors"])
  const ctx = context()
  const kinds = new Set(registry.list(ctx).map((descriptor) => descriptor.draft.kind))
  for (const kind of ["shape", "table", "chart", "group", "table-cell", "chart-series", "animation", "note"]) {
    assert.ok(kinds.has(kind), `registry is missing ${kind}`)
  }
  const behaviors = registry.list(ctx, { kinds: ["animation", "transition", "action", "master", "note"] })
  assert.equal(behaviors.length, 4)
  assert.ok(behaviors.every((descriptor) => interaction.isOfficeObjectVisible(descriptor) === false))
})
