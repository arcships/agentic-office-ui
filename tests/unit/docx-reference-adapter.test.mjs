import assert from "node:assert/strict"
import { readFileSync } from "node:fs"
import test from "node:test"

const fixtures = JSON.parse(readFileSync(
  new URL("../fixtures/office-references/docx-text-revisions.json", import.meta.url),
  "utf8",
))
const docx = await import(new URL("../../packages/docx-core/dist/index.js", import.meta.url).href)
const interaction = await import(new URL("../../packages/office-interaction/dist/index.js", import.meta.url).href)

function modelFromParagraphs(paragraphs) {
  return {
    nodes: paragraphs.map((text) => ({ type: "paragraph", children: [{ type: "text", text }] })),
    metadata: {
      sourceParts: 1,
      warnings: [],
      headerSections: [],
      footerSections: [],
      paragraphStyles: [],
    },
  }
}

function context(revision, paragraphs, pageCount = 2) {
  return {
    revision: { format: "docx", documentId: "report.docx", revision },
    model: modelFromParagraphs(paragraphs),
    pageCount,
  }
}

function reference(draft, referenceId = "ref-docx-basic") {
  return interaction.parseOfficeObjectReference({ ...draft, referenceId })
}

const originalContext = context("docx-v1", fixtures.original)
const originalRange = {
  start: { location: { kind: "paragraph", nodeIndex: 1 }, offset: 4 },
  end: { location: { kind: "paragraph", nodeIndex: 1 }, offset: 21 },
}

test("DOCX adapter converts editor ranges, pages, and normalized regions into valid references", () => {
  const draft = docx.createDocxTextReferenceDraft(originalContext, originalRange, {
    fallbackRegion: { space: "page", pageIndex: 0, rect: { x: 0.1, y: 0.2, width: 0.4, height: 0.08 } },
  })
  assert.equal(draft.fingerprint.exactText, "quarterly revenue")
  assert.equal(draft.locator.value.start.path[0].kind, "paragraph")
  assert.equal(docx.docxTextRangeText(originalContext.model, originalRange), "quarterly revenue")
  assert.equal(reference(draft).kind, "text-range")
  draft.reliability.semantic.reasonCodes.push("caller-mutation")
  assert.deepEqual(
    docx.createDocxTextReferenceDraft(originalContext, originalRange).reliability.semantic.reasonCodes,
    ["docx.native-text"],
  )

  const page = reference(docx.createDocxPageReferenceDraft(originalContext, 1), "ref-page")
  assert.equal(page.locator.value.pageIndex, 1)

  const regionValue = { space: "page", pageIndex: 0, rect: { x: 0.2, y: 0.3, width: 0.5, height: 0.2 } }
  const region = reference(docx.createDocxRegionReferenceDraft(originalContext, regionValue), "ref-region")
  assert.deepEqual(region.locator.value, regionValue)
  assert.throws(
    () => docx.createDocxRegionReferenceDraft(originalContext, {
      space: "page",
      pageIndex: 0,
      rect: { x: 0.8, y: 0.3, width: 0.5, height: 0.2 },
    }),
    /normalized page space/u,
  )
})

test("DOCX text references resolve exactly and relocate after paragraph insertion", () => {
  const selected = reference(docx.createDocxTextReferenceDraft(originalContext, originalRange))
  const exact = docx.resolveDocxReference(originalContext, selected)
  assert.equal(exact.status, "exact")
  assert.equal(exact.descriptor.content.text, "quarterly revenue")

  const movedContext = context("docx-v2", fixtures.relocated)
  const moved = docx.resolveDocxReference(movedContext, selected)
  assert.equal(moved.status, "relocated")
  assert.equal(moved.reference.referenceId, selected.referenceId)
  assert.equal(moved.reference.document.revision, "docx-v2")
  assert.equal(moved.reference.locator.value.start.path[0].index, 2)
  assert.deepEqual(moved.reasonCodes, ["docx.revision-changed", "docx.text-quote-match"])
})

test("DOCX same-revision native locators remain usable when a caller omits optional fingerprints", () => {
  const selected = structuredClone(reference(docx.createDocxTextReferenceDraft(originalContext, originalRange)))
  delete selected.fingerprint
  const exact = docx.resolveDocxReference(originalContext, interaction.parseOfficeObjectReference(selected))
  assert.equal(exact.status, "exact")
  assert.equal(exact.descriptor.content.text, "quarterly revenue")
})

test("DOCX resolver reports deletion and equally credible duplicate text without guessing", () => {
  const selected = reference(docx.createDocxTextReferenceDraft(originalContext, originalRange))
  const deleted = docx.resolveDocxReference(context("docx-v2", fixtures.deleted), selected)
  assert.deepEqual(deleted, { status: "not-found", reasonCode: "docx.text-not-found" })

  const ambiguous = docx.resolveDocxReference(context("docx-v2", fixtures.ambiguous), selected)
  assert.equal(ambiguous.status, "ambiguous")
  assert.equal(ambiguous.candidates.length, 2)
  assert.deepEqual(ambiguous.reasonCodes, ["docx.multiple-text-matches"])
})

test("DOCX table-cell paths round-trip through the current model", () => {
  const tableModel = {
    nodes: [{
      type: "table",
      rows: [{ cells: [{
        type: "table-cell",
        nodes: [{ type: "paragraph", children: [{ type: "text", text: "Cell value" }] }],
      }] }],
    }],
    metadata: originalContext.model.metadata,
  }
  const tableContext = { ...originalContext, model: tableModel }
  const range = {
    start: {
      location: { kind: "table-cell", tableIndex: 0, rowIndex: 0, cellIndex: 0, paragraphIndex: 0 },
      offset: 0,
    },
    end: {
      location: { kind: "table-cell", tableIndex: 0, rowIndex: 0, cellIndex: 0, paragraphIndex: 0 },
      offset: 4,
    },
  }
  const selected = reference(docx.createDocxTextReferenceDraft(tableContext, range), "ref-cell-text")
  assert.deepEqual(selected.locator.value.start.path.map((segment) => segment.kind), [
    "table", "row", "cell", "paragraph",
  ])
  assert.equal(docx.resolveDocxReference(tableContext, selected).status, "exact")
})

test("DOCX page and manual region stay exact only within the same pagination revision", () => {
  const page = reference(docx.createDocxPageReferenceDraft(originalContext, 0), "ref-page")
  const region = reference(docx.createDocxRegionReferenceDraft(originalContext, {
    space: "page",
    pageIndex: 0,
    rect: { x: 0.1, y: 0.1, width: 0.3, height: 0.2 },
  }), "ref-region")
  assert.equal(docx.resolveDocxReference(originalContext, page).status, "exact")
  assert.equal(docx.resolveDocxReference(originalContext, region).status, "exact")
  assert.deepEqual(docx.resolveDocxReference(context("docx-v2", fixtures.original), page), {
    status: "unsupported",
    reasonCode: "docx.reflow-requires-reselection",
  })
  assert.equal(docx.resolveDocxReference(context("docx-v2", fixtures.original), region).status, "unsupported")
})
