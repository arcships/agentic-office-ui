import assert from "node:assert/strict"
import { readFileSync } from "node:fs"
import test from "node:test"

const fixtures = JSON.parse(readFileSync(
  new URL("../fixtures/office-references/pdf-text-revisions.json", import.meta.url),
  "utf8",
))
const pdf = await import(new URL("../../packages/vue-pdf/dist/index.js", import.meta.url).href)
const interaction = await import(new URL("../../packages/office-interaction/dist/index.js", import.meta.url).href)

function createContext(revision, pageTexts, pageOverrides = {}) {
  const pages = pageTexts.map((_text, index) => ({
    index,
    width: pageOverrides.width ?? 100,
    height: pageOverrides.height ?? 200,
    rotation: pageOverrides.rotation ?? 0,
  }))
  const document = { id: `runtime-${revision}`, pageCount: pages.length, pages }
  return {
    revision: { format: "pdf", documentId: "paper.pdf", revision },
    document,
    runtime: {
      async getTextSlices(_document, slices) {
        return slices.map((slice) => pageTexts[slice.pageIndex]?.slice(
          slice.charIndex,
          slice.charIndex + slice.charCount,
        ) ?? "")
      },
      async search(_document, query) {
        const hits = []
        pageTexts.forEach((text, pageIndex) => {
          let from = 0
          while (from <= text.length - query.length) {
            const charIndex = text.indexOf(query, from)
            if (charIndex < 0) break
            hits.push({
              kind: "pdf-text",
              pageIndex,
              charIndex,
              charCount: query.length,
              before: text.slice(0, charIndex),
              match: query,
              after: text.slice(charIndex + query.length),
              rects: [{ x: charIndex, y: 20, width: query.length, height: 10 }],
            })
            from = charIndex + Math.max(1, query.length)
          }
        })
        return hits
      },
    },
  }
}

function reference(draft, referenceId = "ref-pdf-basic") {
  return interaction.parseOfficeObjectReference({ ...draft, referenceId })
}

const originalContext = createContext("pdf-v1", fixtures.original)
const selection = {
  range: { pageIndex: 0, charIndex: 6, charCount: 8 },
  text: "selected",
  rects: [{ x: 6, y: 20, width: 8, height: 10 }],
}

test("PDF adapter converts PageTextSlice geometry, pages, and regions into valid references", () => {
  const draft = pdf.createPdfTextReferenceDraft(originalContext, selection)
  assert.equal(draft.locator.value.charIndex, 6)
  assert.equal(draft.fingerprint.exactText, "selected")
  assert.equal(draft.fallbackRegion.space, "page")
  assert.equal(draft.fallbackRegion.pageIndex, 0)
  assert.equal(draft.fallbackRegion.rect.x, 0.06)
  assert.equal(draft.fallbackRegion.rect.y, 0.1)
  assert.ok(Math.abs(draft.fallbackRegion.rect.width - 0.08) < 1e-12)
  assert.ok(Math.abs(draft.fallbackRegion.rect.height - 0.05) < 1e-12)
  assert.equal(reference(draft).kind, "text-range")
  draft.reliability.semantic.reasonCodes.push("caller-mutation")
  assert.deepEqual(
    pdf.createPdfTextReferenceDraft(originalContext, selection).reliability.semantic.reasonCodes,
    ["pdf.native-text"],
  )

  assert.deepEqual(
    pdf.normalizePdfReferenceRect(originalContext.document.pages[0], { x: -10, y: 190, width: 30, height: 30 }),
    { x: 0, y: 0.95, width: 0.2, height: 0.05 },
  )
  assert.equal(reference(pdf.createPdfPageReferenceDraft(originalContext, 0), "ref-page").kind, "page")
  assert.equal(reference(pdf.createPdfRegionReferenceDraft(originalContext, {
    space: "page",
    pageIndex: 0,
    rect: { x: 0.1, y: 0.2, width: 0.4, height: 0.3 },
  }), "ref-region").kind, "region")
})

test("PDF text references resolve exactly and relocate through the runtime search boundary", async () => {
  const selected = reference(pdf.createPdfTextReferenceDraft(originalContext, selection))
  const exact = await pdf.resolvePdfReference(originalContext, selected)
  assert.equal(exact.status, "exact")
  assert.equal(exact.descriptor.content.text, "selected")

  const relocatedContext = createContext("pdf-v2", fixtures.relocated)
  const relocated = await pdf.resolvePdfReference(relocatedContext, selected)
  assert.equal(relocated.status, "relocated")
  assert.equal(relocated.reference.referenceId, selected.referenceId)
  assert.equal(relocated.reference.document.revision, "pdf-v2")
  assert.equal(relocated.reference.locator.value.charIndex, 11)
  assert.deepEqual(relocated.reasonCodes, ["pdf.revision-changed", "pdf.text-quote-match"])
})

test("PDF same-revision native locators remain usable without an optional fingerprint", async () => {
  const selected = structuredClone(reference(pdf.createPdfTextReferenceDraft(originalContext, selection)))
  delete selected.fingerprint
  const exact = await pdf.resolvePdfReference(originalContext, interaction.parseOfficeObjectReference(selected))
  assert.equal(exact.status, "exact")
  assert.equal(exact.descriptor.content.text, "selected")
})

test("PDF resolver reports deletion and duplicate text as distinct outcomes", async () => {
  const selected = reference(pdf.createPdfTextReferenceDraft(originalContext, selection))
  assert.deepEqual(await pdf.resolvePdfReference(createContext("pdf-v2", fixtures.deleted), selected), {
    status: "not-found",
    reasonCode: "pdf.text-not-found",
  })
  const ambiguous = await pdf.resolvePdfReference(createContext("pdf-v2", fixtures.ambiguous), selected)
  assert.equal(ambiguous.status, "ambiguous")
  assert.equal(ambiguous.candidates.length, 2)
  assert.deepEqual(ambiguous.reasonCodes, ["pdf.multiple-text-matches"])
})

test("PDF page and region relocation requires an unchanged page signature", async () => {
  const page = reference(pdf.createPdfPageReferenceDraft(originalContext, 0), "ref-page")
  const region = reference(pdf.createPdfRegionReferenceDraft(originalContext, {
    space: "page",
    pageIndex: 0,
    rect: { x: 0.1, y: 0.2, width: 0.4, height: 0.3 },
  }), "ref-region")
  assert.equal((await pdf.resolvePdfReference(originalContext, page)).status, "exact")
  assert.equal((await pdf.resolvePdfReference(createContext("pdf-v2", fixtures.original), page)).status, "relocated")
  assert.equal((await pdf.resolvePdfReference(createContext("pdf-v2", fixtures.original), region)).status, "relocated")
  assert.deepEqual(
    await pdf.resolvePdfReference(createContext("pdf-v2", fixtures.original, { width: 120 }), page),
    { status: "not-found", reasonCode: "pdf.page-signature-changed" },
  )

  const unsignedPage = structuredClone(page)
  delete unsignedPage.fingerprint
  assert.deepEqual(
    await pdf.resolvePdfReference(
      createContext("pdf-v2", fixtures.original),
      interaction.parseOfficeObjectReference(unsignedPage),
    ),
    { status: "unsupported", reasonCode: "pdf.page-signature-unavailable" },
  )
})

test("PDF resolver forwards cancellation instead of turning it into document invalidation", async () => {
  const selected = reference(pdf.createPdfTextReferenceDraft(originalContext, selection))
  const controller = new AbortController()
  controller.abort()
  await assert.rejects(
    () => pdf.resolvePdfReference(originalContext, selected, controller.signal),
    (error) => error?.name === "AbortError",
  )
})

test("PDF resolver does not misreport runtime failures as deleted content", async () => {
  const selected = reference(pdf.createPdfTextReferenceDraft(originalContext, selection))
  const failedContext = {
    ...originalContext,
    runtime: {
      async getTextSlices() { throw new Error("PDF engine unavailable") },
      async search() { throw new Error("PDF engine unavailable") },
    },
  }
  await assert.rejects(
    () => pdf.resolvePdfReference(failedContext, selected),
    /PDF engine unavailable/u,
  )
})
