import assert from "node:assert/strict"
import test from "node:test"

const docx = await import(
  new URL("../../packages/docx-core/dist/index.js", import.meta.url).href
)
const interaction = await import(
  new URL("../../packages/office-interaction/dist/index.js", import.meta.url).href
)

function paragraph(text, extra = {}) {
  return { type: "paragraph", children: [{ type: "text", text }], ...extra }
}

function tableNode(rows) {
  return {
    type: "table",
    rows: rows.map((cells) => ({
      cells: cells.map((text) => ({ nodes: [paragraph(text)] })),
    })),
  }
}

function model({ nodes, comments, styles = [], sections } = {}) {
  return {
    nodes: nodes ?? [paragraph("Quarterly report")],
    metadata: {
      sourceParts: 1,
      warnings: [],
      headerSections: [],
      footerSections: [],
      paragraphStyles: styles,
      ...(comments ? { comments } : {}),
      ...(sections ? { sections } : {}),
    },
  }
}

function context(docModel, revision = "docx-rev-1") {
  return {
    revision: { format: "docx", documentId: "stage-two.docx", revision },
    model: docModel,
    pageCount: 3,
  }
}

function descriptorFor(index, kind, predicate = () => true) {
  return index.descriptors.find((descriptor) =>
    descriptor.draft.kind === kind && predicate(descriptor))
}

test("DOCX stage 2 blocks carry descriptors, fingerprints and a resolvable table exemplar", () => {
  const docModel = model({
    nodes: [
      paragraph("Introduction"),
      paragraph("Overview", { style: { headingLevel: 2 } }),
      paragraph("First item", { style: { numbering: { numId: 1, ilvl: 0 } } }),
      {
        type: "paragraph",
        children: [{ type: "image", alt: "Revenue chart" }],
      },
      tableNode([["Region", "Revenue"], ["EMEA", "42"]]),
    ],
  })
  const index = docx.createDocxObjectIndex(context(docModel))

  const paragraphObject = descriptorFor(index, "paragraph")
  assert.ok(paragraphObject, "paragraph descriptors are registered")
  assert.equal(paragraphObject.draft.source, "native")
  assert.ok(paragraphObject.draft.fingerprint.contentHash.startsWith("fnv1a32:"))
  assert.ok(paragraphObject.draft.fingerprint.ancestorKeys.includes("docx.part:body"))

  const heading = descriptorFor(index, "heading")
  assert.ok(heading, "a paragraph style with headingLevel becomes a heading object")
  assert.match(heading.label, /^Heading:/u)

  const listItem = descriptorFor(index, "list-item")
  assert.ok(listItem, "numbering becomes a list item object")
  assert.match(listItem.label, /^List item:/u)

  const image = descriptorFor(index, "image")
  assert.ok(image, "image children become image objects")
  assert.equal(image.draft.fingerprint.objectName, "Revenue chart")

  const table = descriptorFor(index, "table")
  assert.ok(table, "tables are stage 2 objects")
  assert.equal(table.draft.reliability.semantic.level, "exact")
  assert.equal(table.draft.reliability.relocation.level, "likely")
  assert.ok(table.draft.fingerprint.contentHash.startsWith("fnv1a32:"))
  assert.deepEqual(
    table.draft.fingerprint.ancestorKeys.slice(0, 2),
    ["docx.part:body", "docx.table:4"],
  )

  const confirmed = interaction.confirmOfficeReferenceDraft(table.draft, {
    referenceId: "ref-docx-table",
    trigger: "pointer",
  })
  assert.equal(confirmed.reference.kind, "table")
  assert.equal(confirmed.reference.fingerprint.contentHash, table.draft.fingerprint.contentHash)
  assert.equal(confirmed.reference.locator.value.kind, "structure")
  const resolved = docx.resolveDocxReference(context(docModel), confirmed.reference)
  assert.equal(resolved.status, "exact")
  assert.match(resolved.descriptor.label, /^Table/u)
  assert.equal(resolved.descriptor.facets, undefined)
})

test("DOCX stage 3 table hierarchy links cells to rows and tables", () => {
  const docModel = model({
    nodes: [tableNode([["A1", "B1"], ["A2", "B2"]])],
  })
  const index = docx.createDocxObjectIndex(context(docModel))
  const table = descriptorFor(index, "table")
  const cell = index.descriptors.find((descriptor) =>
    descriptor.draft.kind === "table-cell"
    && descriptor.label.includes("Cell 2, 1"))
  assert.ok(table)
  assert.ok(cell)
  const chain = interaction.officeObjectAncestors(index.descriptors, cell.objectId)
  assert.deepEqual(chain.map((entry) => entry.draft.kind), [
    "table",
    "table-row",
    "table-cell",
  ])
  const children = interaction.officeObjectChildren(index.descriptors, table.objectId)
  assert.deepEqual(children.map((entry) => entry.draft.kind), ["table-row", "table-row"])
  assert.equal(cell.childrenState, "available")

  const confirmed = interaction.confirmOfficeReferenceDraft(cell.draft, {
    referenceId: "ref-docx-cell",
    trigger: "pointer",
  })
  const resolved = docx.resolveDocxReference(context(docModel), confirmed.reference)
  assert.equal(resolved.status, "exact")
  assert.equal(resolved.descriptor.draft.kind, "table-cell")
})

test("DOCX stage 4 behaviors are reference kinds without outlines", () => {
  const docModel = model({
    nodes: [
      paragraph("Intro", {
        sourceXml: '<w:p><w:bookmarkStart w:id="1" w:name="Intro"/><w:instrText> PAGE </w:instrText><w:instrText> REF Intro </w:instrText><w:ins w:id="2" w:author="Ada"><w:t>x</w:t></w:ins></w:p>',
      }),
    ],
    comments: [{ id: 7, author: "Ada", text: "Check this", resolved: true }],
    styles: [{ id: "Heading1", name: "Heading 1", headingLevel: 1 }],
    sections: [{ startNodeIndex: 0, headerSections: [], footerSections: [] }],
  })
  const contextValue = context(docModel)
  const index = docx.createDocxObjectIndex(contextValue)

  const behaviors = index.behaviors
  assert.ok(behaviors.length >= 5)
  for (const descriptor of behaviors) {
    assert.equal(interaction.isOfficeObjectVisible(descriptor), false)
    assert.equal(descriptor.visual, undefined)
  }
  const byKind = (kind) => behaviors.find((descriptor) => descriptor.draft.kind === kind)
  const style = byKind("style")
  const bookmark = byKind("bookmark")
  const field = byKind("field")
  const crossReference = byKind("cross-reference")
  const trackedChange = byKind("tracked-change")
  const commentState = byKind("comment")
  for (const descriptor of [style, bookmark, field, crossReference, trackedChange, commentState]) {
    assert.ok(descriptor, `missing behavior descriptor: ${descriptor?.draft.kind}`)
    assert.equal(descriptor.draft.locator.value.kind, "behavior")
    assert.ok(descriptor.draft.locator.value.instanceId.startsWith("docx."))
    assert.equal(descriptor.draft.source, ["field", "cross-reference", "tracked-change"].includes(descriptor.draft.kind)
      ? "structural"
      : "native")
  }
  assert.equal(style.draft.reliability.semantic.level, "exact")
  assert.equal(commentState.facets.resolved, true)
  assert.equal(commentState.label, "Comment state: resolved")

  const confirmed = interaction.confirmOfficeReferenceDraft(bookmark.draft, {
    referenceId: "ref-docx-bookmark",
    trigger: "programmatic",
  })
  assert.equal(confirmed.reference.kind, "bookmark")
  const resolved = docx.resolveDocxReference(contextValue, confirmed.reference)
  assert.equal(resolved.status, "exact")
  assert.equal(resolved.descriptor.draft.kind, "bookmark")
  assert.match(resolved.descriptor.label, /^Bookmark: Intro$/u)

  const localized = interaction.localizeOfficeDescriptorLabel(
    resolved.descriptor.label,
    interaction.resolveOfficeMessages("zh-CN").common,
  )
  assert.equal(localized, "书签：Intro")

  const updated = docx.resolveDocxReference(context(docModel, "docx-rev-2"), confirmed.reference)
  assert.equal(updated.status, "relocated")
  assert.equal(updated.reference.document.revision, "docx-rev-2")
  assert.equal(updated.descriptor.label, "Bookmark: Intro")
})

test("DOCX object registration is extensible and captured per model", () => {
  const registry = docx.createDocxObjectRegistry()
  assert.deepEqual(registry.providers().map((provider) => provider.id), [
    "docx.blocks",
    "docx.comments",
    "docx.behaviors",
  ])
  const docModel = model({ nodes: [paragraph("Custom target")] })
  const custom = {
    id: "host.custom",
    format: "docx",
    kinds: ["unknown"],
    list: () => [],
  }
  const unregister = registry.register(custom)
  assert.equal(registry.providers().length, 4)
  assert.deepEqual(registry.list(context(docModel), { kinds: ["paragraph"] })
    .some((descriptor) => descriptor.draft.kind === "paragraph"), true)
  unregister()
  assert.equal(registry.providers().length, 3)
  assert.equal(docx.createDocxObjectIndex(context(docModel)), docx.createDocxObjectIndex(context(docModel)))
})
