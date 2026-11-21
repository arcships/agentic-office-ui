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

function model({ nodes, comments, styles = [], sections, headerSections = [], footerSections = [], footnotes, endnotes } = {}) {
  return {
    nodes: nodes ?? [paragraph("Quarterly report")],
    metadata: {
      sourceParts: 1,
      warnings: [],
      headerSections,
      footerSections,
      paragraphStyles: styles,
      ...(comments ? { comments } : {}),
      ...(sections ? { sections } : {}),
      ...(footnotes ? { footnotes } : {}),
      ...(endnotes ? { endnotes } : {}),
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

test("DOCX fields and revisions inside headers, footers and notes become stage 4 behaviors", () => {
  const withField = (text, instruction) => paragraph(text, {
    sourceXml: `<w:p><w:r><w:instrText> ${instruction} </w:instrText></w:r></w:p>`,
  })
  const docModel = model({
    nodes: [withField("Body", "DATE")],
    headerSections: [{
      partName: "word/header1.xml",
      nodes: [withField("Header", "PAGE")],
    }],
    footerSections: [{
      partName: "word/footer1.xml",
      nodes: [withField("Footer", "REF Intro")],
    }],
    footnotes: [{ id: 1, text: "note", nodes: [withField("Note", "NUMPAGES")] }],
    endnotes: [{ id: 2, text: "endnote", nodes: [withField("Endnote", "PAGEREF Target")] }],
  })
  const contextValue = context(docModel)
  const index = docx.createDocxObjectIndex(contextValue)
  const fields = index.behaviors.filter((descriptor) =>
    descriptor.draft.kind === "field" || descriptor.draft.kind === "cross-reference")
  const byPart = (part) => fields.filter((descriptor) => descriptor.facets.part === part)

  assert.equal(byPart("body").length, 1)
  assert.equal(byPart("header:word/header1.xml").length, 1)
  assert.equal(byPart("footer:word/footer1.xml").length, 1)
  assert.equal(byPart("footnote:1").length, 1)
  assert.equal(byPart("endnote:2").length, 1)

  const headerField = byPart("header:word/header1.xml")[0]
  assert.equal(headerField.label, "Field: PAGE")
  assert.deepEqual(headerField.draft.locator.value.owner.part, {
    kind: "header",
    partName: "word/header1.xml",
  })
  assert.equal(headerField.draft.locator.value.owner.path[0].kind, "paragraph")
  assert.deepEqual(headerField.path.map((entry) => entry.kind), [
    "document", "section", "field",
  ])
  assert.equal(interaction.isOfficeObjectVisible(headerField), false)
  assert.ok(headerField.draft.fingerprint.ancestorKeys.includes(
    `docx.behavior-instance:${headerField.draft.locator.value.instanceId}`,
  ))

  const footerField = byPart("footer:word/footer1.xml")[0]
  assert.equal(footerField.draft.kind, "cross-reference")
  assert.equal(footerField.label, "Cross-reference: Intro")
  const confirmed = interaction.confirmOfficeReferenceDraft(headerField.draft, {
    referenceId: "ref-docx-header-field",
    trigger: "programmatic",
  })
  const resolved = docx.resolveDocxReference(contextValue, confirmed.reference)
  assert.equal(resolved.status, "exact")
  assert.equal(resolved.descriptor.draft.kind, "field")
  assert.match(resolved.descriptor.label, /^Field: PAGE$/u)
  assert.equal("visual" in resolved.descriptor, false)

  const relocated = docx.resolveDocxReference(context(docModel, "docx-rev-2"), confirmed.reference)
  assert.equal(relocated.status, "relocated")
  assert.equal(relocated.reference.document.revision, "docx-rev-2")
  assert.equal(relocated.descriptor.label, "Field: PAGE")

  const ids = fields.map((descriptor) => descriptor.draft.locator.value.instanceId)
  assert.equal(new Set(ids).size, ids.length, "field instances stay unique across story parts")
})

test("DOCX nested tables keep their own table, row and cell hierarchy", () => {
  const nestedTable = tableNode([["inner"]])
  const outerTable = {
    type: "table",
    rows: [{
      cells: [{ nodes: [paragraph("outer"), nestedTable] }],
    }],
  }
  const docModel = model({ nodes: [outerTable] })
  const contextValue = context(docModel)
  const index = docx.createDocxObjectIndex(contextValue)
  const tables = index.descriptors.filter((descriptor) => descriptor.draft.kind === "table")
  assert.equal(tables.length, 2, "both the outer and nested table are objects")
  const nestedTableObject = tables.find((descriptor) =>
    descriptor.path.filter((entry) => entry.kind === "table").length === 2)
  const outerCell = index.descriptors.find((descriptor) =>
    descriptor.draft.kind === "table-cell" && descriptor.facets.path === "table:0/row:0/cell:0")
  assert.ok(nestedTableObject)
  assert.ok(outerCell)
  assert.equal(nestedTableObject.parentObjectId, outerCell.objectId)

  const nestedRow = index.descriptors.find((descriptor) =>
    descriptor.draft.kind === "table-row"
    && descriptor.facets.path === "table:0/row:0/cell:0/table:1/row:0")
  const nestedCell = index.descriptors.find((descriptor) =>
    descriptor.draft.kind === "table-cell"
    && descriptor.facets.path === "table:0/row:0/cell:0/table:1/row:0/cell:0")
  assert.ok(nestedRow)
  assert.ok(nestedCell)
  assert.equal(nestedRow.parentObjectId, nestedTableObject.objectId)
  assert.equal(nestedCell.parentObjectId, nestedRow.objectId)
  assert.deepEqual(
    interaction.officeObjectAncestors(index.descriptors, nestedCell.objectId).map((entry) => entry.draft.kind),
    ["table", "table-row", "table-cell", "table", "table-row", "table-cell"],
  )

  const confirmed = interaction.confirmOfficeReferenceDraft(nestedCell.draft, {
    referenceId: "ref-docx-nested-cell",
    trigger: "pointer",
  })
  const resolved = docx.resolveDocxReference(contextValue, confirmed.reference)
  assert.equal(resolved.status, "exact")
  assert.equal(resolved.descriptor.draft.kind, "table-cell")
  assert.equal(resolved.descriptor.label, "Cell 1, 1")
  assert.deepEqual(resolved.descriptor.path.map((entry) => entry.kind), [
    "document", "table", "table-row", "table-cell", "table", "table-row", "table-cell",
  ])
  assert.ok(nestedCell.draft.fingerprint.ancestorKeys.includes("docx.table:1"))
  assert.equal(
    index.descriptors.filter((descriptor) => descriptor.objectId === nestedCell.objectId).length,
    1,
    "nested row indices must not collide with top-level row object ids",
  )
})
