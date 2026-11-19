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

function context(revision = "pptx-rev-1", { withProviders = true } = {}) {
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
          getObjectChart: (_slideIndex, key) => key === chartKey ? chartSnapshot : undefined,
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

test("PPTX stage 3 hierarchy nests table cells and chart internals under their objects", () => {
  const ctx = context()
  const descriptors = pptx.createPptxObjectIndex(ctx).descriptors
  const table = descriptors.find((descriptor) => descriptor.objectId === `pptx:object:${tableKey}`)
  const cell = descriptors.find((descriptor) => descriptor.draft.kind === "table-cell")
  const chart = descriptors.find((descriptor) => descriptor.objectId === `pptx:object:${chartKey}`)
  const series = descriptors.find((descriptor) =>
    descriptor.draft.kind === "chart-series" && descriptor.label === "Series: EMEA")
  const point = descriptors.find((descriptor) =>
    descriptor.draft.kind === "chart-point" && descriptor.parentObjectId === series?.objectId)
  assert.ok(table)
  assert.ok(cell)
  assert.ok(chart)
  assert.ok(series)
  assert.ok(point)
  assert.equal(cell.parentObjectId, table.objectId)
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
