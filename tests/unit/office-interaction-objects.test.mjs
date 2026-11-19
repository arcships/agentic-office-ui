import assert from "node:assert/strict"
import test from "node:test"

const interaction = await import(
  new URL("../../packages/office-interaction/dist/index.js", import.meta.url).href
)

function reliability() {
  return {
    semantic: { level: "exact", score: 1, reasonCodes: ["test.semantic"] },
    boundary: { level: "exact", score: 1, reasonCodes: ["test.boundary"] },
    hierarchy: { level: "likely", reasonCodes: ["test.hierarchy"] },
    relocation: { level: "likely", reasonCodes: ["test.relocation"] },
  }
}

function draft(kind, objectId = "p0") {
  return {
    schemaVersion: 1,
    document: { format: "docx", documentId: "objects.docx", revision: "rev-1" },
    kind,
    source: "native",
    locator: {
      type: "format",
      format: "docx",
      value: { kind: "structure", part: { kind: "body" }, path: [{ kind: "paragraph", index: 0 }] },
    },
    fingerprint: { ancestorKeys: [`test.object:${objectId}`] },
    reliability: reliability(),
  }
}

function descriptor(objectId, { kind = "paragraph", parentObjectId, visual, label } = {}) {
  return interaction.createOfficeObjectDescriptor({
    objectId,
    draft: draft(kind, objectId),
    label: label ?? objectId,
    path: [{ kind: "document", label: "Document" }, { kind, label: objectId }],
    ...(parentObjectId ? { parentObjectId } : {}),
    ...(visual
      ? {
          visual: {
            fragments: [{ container: { space: "page", pageIndex: 0 }, rect: { x: 0.1, y: 0.1, width: 0.2, height: 0.2 } }],
            layoutVersion: "rev-1",
          },
        }
      : {}),
    facets: { test: true },
  })
}

function provider(id, format, kinds, list) {
  return { id, format, kinds, list }
}

test("object descriptors copy caller data and reject unusable identities", () => {
  const visual = { x: 0.1, y: 0.1, width: 0.2, height: 0.2 }
  const path = [{ kind: "document", label: "Document" }]
  const built = interaction.createOfficeObjectDescriptor({
    objectId: "  table-1  ",
    draft: draft("table"),
    label: "Table 1",
    path,
    visual: { fragments: [{ container: { space: "page", pageIndex: 0 }, rect: visual }], layoutVersion: "rev-1" },
    facets: { rows: 2 },
  })
  assert.equal(built.objectId, "table-1")
  assert.equal(built.childrenState, "none")
  path[0].label = "mutated"
  visual.x = 0.9
  assert.equal(built.path[0].label, "Document")
  assert.equal(built.visual.fragments[0].rect.x, 0.1)

  assert.throws(() => interaction.createOfficeObjectDescriptor({
    objectId: " ",
    draft: draft("table"),
    label: "x",
    path: [],
  }), /objectId/u)
  assert.throws(() => interaction.createOfficeObjectDescriptor({
    objectId: "a",
    draft: draft("table"),
    label: "x",
    path: [],
    parentObjectId: "a",
  }), /own parent/u)

  const fingerprint = interaction.createOfficeObjectFingerprint({
    exactText: "Revenue",
    ancestorKeys: ["a", "b"],
  })
  assert.throws(() => {
    fingerprint.ancestorKeys.push("c")
  }, /not extensible|read only/u)
  assert.deepEqual(fingerprint.ancestorKeys, ["a", "b"])
})

test("object registry registers, replaces and unregisters providers without leaking state", () => {
  const registry = interaction.createOfficeObjectRegistry()
  const first = provider("docx.paragraphs", "docx", ["paragraph"], () => [descriptor("p1", { visual: true })])
  const unregister = registry.register(first)
  assert.equal(registry.providers().length, 1)
  assert.equal(registry.get("docx.paragraphs"), first)

  const replacement = provider("docx.paragraphs", "docx", ["paragraph"], () => [descriptor("p2", { visual: true })])
  registry.register(replacement)
  assert.deepEqual(registry.list({}).map((entry) => entry.objectId), ["p2"])
  unregister()
  // The first unregister closure must not remove the replacement.
  assert.equal(registry.get("docx.paragraphs"), replacement)
  assert.equal(registry.unregister("docx.paragraphs"), true)

  const tableProvider = provider("docx.tables", "docx", ["table"], () => [
    descriptor("t1", { kind: "table", visual: true }),
    descriptor("t2", { kind: "table" }),
  ])
  registry.register(tableProvider)
  assert.deepEqual(registry.list({}, { kinds: ["table"] }).map((entry) => entry.objectId), ["t1", "t2"])
  assert.deepEqual(registry.list({}, { visibleOnly: true }).map((entry) => entry.objectId), ["t1"])
  assert.deepEqual(registry.list({}, { maxObjects: 1 }).map((entry) => entry.objectId), ["t1"])
  assert.throws(() => registry.list({}, { maxObjects: 0 }), /maxObjects/u)
  assert.throws(() => registry.list({}, { maxObjects: interaction.MAX_OFFICE_OBJECT_CANDIDATES + 1 }), /maxObjects/u)

  const wrongFormat = provider("docx.broken", "docx", ["paragraph"], () => [{
    ...descriptor("wrong"),
    draft: { ...draft("paragraph"), document: { format: "pdf", documentId: "d", revision: "r" } },
  }])
  registry.register(wrongFormat)
  assert.throws(() => registry.list({}), /pdf descriptor|docx descriptor/u)
})

test("object hierarchy exposes parents, children and deterministic ancestor chains", () => {
  const descriptors = [
    descriptor("document", { kind: "document" }),
    descriptor("table", { kind: "table", parentObjectId: "document", visual: true }),
    descriptor("row", { kind: "table-row", parentObjectId: "table", visual: true }),
    descriptor("cell", { kind: "table-cell", parentObjectId: "row", visual: true }),
  ]
  assert.deepEqual(
    interaction.officeObjectAncestors(descriptors, "cell").map((entry) => entry.objectId),
    ["document", "table", "row", "cell"],
  )
  assert.deepEqual(
    interaction.officeObjectChildren(descriptors, "row").map((entry) => entry.objectId),
    ["cell"],
  )
  assert.deepEqual(interaction.officeObjectChildren(descriptors, "cell"), [])

  const cyclic = [
    descriptor("a", { parentObjectId: "b" }),
    descriptor("b", { parentObjectId: "a" }),
  ]
  assert.throws(() => interaction.officeObjectAncestors(cyclic, "a"), /cycle/u)
})

test("object candidates rank deterministically and convert into Surface previews", () => {
  const candidates = [
    { candidateId: "page", descriptor: descriptor("page", { kind: "page", visual: true }), hit: "ancestor", depth: 1 },
    { candidateId: "cell", descriptor: descriptor("cell", { kind: "table-cell", visual: true }), hit: "direct", depth: 3 },
    { candidateId: "table", descriptor: descriptor("table", { kind: "table", visual: true }), hit: "direct", depth: 2 },
    { candidateId: "inferred", descriptor: descriptor("inferred", { visual: true }), hit: "inferred", depth: 4 },
  ]
  const ranked = interaction.rankOfficeObjectCandidates(candidates)
  assert.deepEqual(ranked.map((entry) => entry.candidateId), ["cell", "table", "inferred", "page"])
  assert.deepEqual(
    interaction.rankOfficeObjectCandidates([...candidates].reverse()).map((entry) => entry.candidateId),
    ["cell", "table", "inferred", "page"],
  )
  const preview = interaction.officeObjectCandidatePreview(ranked[0])
  assert.equal(preview.candidateId, "cell")
  assert.equal(preview.draft.kind, "table-cell")
  assert.ok(!("runtimeTarget" in preview))
  assert.ok(preview.preview.visual)
})

test("invisible behavior descriptors are references without an outline", () => {
  const visible = descriptor("table", { kind: "table", visual: true })
  const behavior = descriptor("bookmark", { kind: "bookmark" })
  assert.equal(interaction.isOfficeObjectVisible(visible), true)
  assert.equal(interaction.isOfficeObjectVisible(behavior), false)
  assert.equal(behavior.visual, undefined)
})

test("descriptor localization covers stage 2-4 object labels in English and Chinese", () => {
  const localized = (label, locale) => interaction.localizeOfficeDescriptorLabel(
    label,
    interaction.resolveOfficeMessages(locale).common,
  )
  assert.equal(localized("Series: Revenue", "en"), "Series: Revenue")
  assert.equal(localized("Series: Revenue", "zh-CN"), "系列：Revenue")
  assert.equal(localized("Bookmark: Intro", "zh-CN"), "书签：Intro")
  assert.equal(localized("Formula: =SUM(A1:A2)", "zh-CN"), "公式：=SUM(A1:A2)")
  assert.equal(localized("Table Sales", "en"), "Table Sales")
  assert.equal(localized("Slide transition", "zh-CN"), "幻灯片切换")
  assert.equal(localized("Link", "zh-CN"), "链接")
  assert.equal(localized("Table 3", "en"), "Table 3")
  assert.equal(localized("Unmapped object label", "en"), "Unmapped object label")
})
