import assert from "node:assert/strict"
import { readFileSync } from "node:fs"
import test from "node:test"

const fixtures = JSON.parse(readFileSync(
  new URL("../fixtures/office-references/pptx-object-revisions.json", import.meta.url),
  "utf8",
))
const pptx = await import(new URL("../../packages/pptx-core/dist/index.js", import.meta.url).href)
const interaction = await import(new URL("../../packages/office-interaction/dist/index.js", import.meta.url).href)

function context(revision, slideFixtures, dimensions = {}, textByShapeId = {}) {
  const slideIds = slideFixtures.map((slide) => slide.slideId)
  const playbackSlides = slideFixtures.map((slide, index) => ({
    index,
    hidden: false,
    objects: slide.objects.map((object) => ({
      ...object,
      key: pptx.createPptxObjectKey({
        slidePath: slide.slideId,
        source: "slide",
        shapeId: object.shapeId,
        groupPath: object.groupPath,
      }),
      slidePath: slide.slideId,
      source: "slide",
    })),
    morphFromPrevious: [],
    nodes: {},
    media: [],
    actions: [],
    capability: { discovered: 0, strict: 0, approximate: 0, static: 0, unparsed: 0, features: [] },
  }))
  const result = {
    revision: { format: "pptx", documentId: "pitch.pptx", revision },
    document: {
      width: dimensions.width ?? 1280,
      height: dimensions.height ?? 720,
      slides: playbackSlides.map((slide, index) => ({ index, number: index + 1, hidden: slide.hidden })),
      warnings: [],
    },
    playbackDocument: {
      slides: playbackSlides,
      capability: { discovered: 0, strict: 0, approximate: 0, static: 0, unparsed: 0, features: [] },
    },
    slideIds,
  }
  result.getObjectParagraphs = (slideIndex, objectKey) => {
    const identity = result.playbackDocument.slides[slideIndex]?.objects.find((object) => object.key === objectKey)
    return identity ? textByShapeId[identity.shapeId] : undefined
  }
  return result
}

function reference(draft, referenceId = "ref-pptx-basic") {
  return interaction.parseOfficeObjectReference({ ...draft, referenceId })
}

const original = context("pptx-v1", fixtures.original)
const summaryObject = original.playbackDocument.slides[0].objects[0]

test("PPTX adapter creates valid slide, object, nested group, and region references", () => {
  const slide = reference(pptx.createPptxSlideReferenceDraft(original, 0), "ref-slide")
  assert.equal(slide.locator.value.slide.slideId, "ppt/slides/slide1.xml")

  const objectDraft = pptx.createPptxObjectReferenceDraft(original, 0, summaryObject.key, {
    fallbackRegion: { space: "slide", slideIndex: 0, rect: { x: 0.1, y: 0.2, width: 0.4, height: 0.2 } },
  })
  assert.equal(objectDraft.kind, "shape")
  assert.equal(objectDraft.locator.value.object.shapeId, "2")
  assert.equal(reference(objectDraft, "ref-object").kind, "shape")
  objectDraft.reliability.semantic.reasonCodes.push("caller-mutation")
  assert.deepEqual(
    pptx.createPptxObjectReferenceDraft(original, 0, summaryObject.key).reliability.semantic.reasonCodes,
    ["pptx.native-object-type"],
  )

  const group = original.playbackDocument.slides[0].objects[1]
  const child = original.playbackDocument.slides[0].objects[2]
  assert.equal(pptx.createPptxObjectReferenceDraft(original, 0, group.key).kind, "group")
  const childReference = reference(pptx.createPptxObjectReferenceDraft(original, 0, child.key), "ref-child")
  assert.deepEqual(childReference.locator.value.object.groupPath, ["10"])
  assert.equal(pptx.resolvePptxReference(original, childReference).descriptor.path[2].kind, "group")

  const region = reference(pptx.createPptxRegionReferenceDraft(original, {
    space: "slide",
    slideIndex: 0,
    rect: { x: 0.2, y: 0.3, width: 0.3, height: 0.25 },
  }), "ref-region")
  assert.equal(region.kind, "region")
})

test("PPTX adapter creates, validates, relocates, and disambiguates precise text ranges", () => {
  const withText = context("pptx-v1", fixtures.original, {}, { "2": ["Quarterly outlook", "Revenue grows steadily"] })
  const object = withText.playbackDocument.slides[0].objects[0]
  const selected = reference(pptx.createPptxTextReferenceDraft(
    withText,
    0,
    object.key,
    { paragraphIndex: 1, offset: 0 },
    { paragraphIndex: 1, offset: 7 },
  ), "ref-text")
  assert.equal(selected.kind, "text-range")
  assert.equal(selected.fingerprint.exactText, "Revenue")
  const exact = pptx.resolvePptxReference(withText, selected)
  assert.equal(exact.status, "exact")
  assert.equal(exact.descriptor.content.text, "Revenue")

  const moved = context("pptx-v2", fixtures.original, {}, { "2": ["Revenue", "Quarterly outlook", "grows steadily"] })
  const relocated = pptx.resolvePptxReference(moved, selected)
  assert.equal(relocated.status, "relocated")
  assert.deepEqual(relocated.reference.locator.value.start, { paragraphIndex: 0, offset: 0 })
  assert.deepEqual(relocated.reference.locator.value.end, { paragraphIndex: 0, offset: 7 })

  const duplicate = context("pptx-v2", fixtures.original, {}, { "2": ["Revenue", "Revenue"] })
  const ambiguous = pptx.resolvePptxReference(duplicate, selected)
  assert.equal(ambiguous.status, "ambiguous")
  assert.equal(ambiguous.candidates.length, 2)

  const deleted = context("pptx-v2", fixtures.original, {}, { "2": ["Quarterly outlook"] })
  assert.deepEqual(pptx.resolvePptxReference(deleted, selected), {
    status: "not-found",
    reasonCode: "pptx.text-not-found",
  })
})

test("PPTX slide paths keep slides and objects locatable through slide reorder", () => {
  const slide = reference(pptx.createPptxSlideReferenceDraft(original, 0), "ref-slide")
  const object = reference(pptx.createPptxObjectReferenceDraft(original, 0, summaryObject.key), "ref-object")
  assert.equal(pptx.resolvePptxReference(original, object).status, "exact")

  const reordered = context("pptx-v2", fixtures.reordered)
  const relocatedSlide = pptx.resolvePptxReference(reordered, slide)
  assert.equal(relocatedSlide.status, "relocated")
  assert.equal(relocatedSlide.reference.locator.value.slide.index, 1)

  const relocatedObject = pptx.resolvePptxReference(reordered, object)
  assert.equal(relocatedObject.status, "relocated")
  assert.equal(relocatedObject.reference.locator.value.slide.index, 1)
  assert.equal(relocatedObject.reference.locator.value.object.objectKey, summaryObject.key)
  assert.deepEqual(relocatedObject.reasonCodes, [
    "pptx.revision-changed",
    "pptx.slide-id-match",
    "pptx.object-key-match",
  ])
})

test("PPTX object resolver falls back to unique name/type and reports duplicate evidence", () => {
  const selected = reference(pptx.createPptxObjectReferenceDraft(original, 0, summaryObject.key), "ref-object")
  const renumbered = pptx.resolvePptxReference(context("pptx-v2", fixtures.renumberedObject), selected)
  assert.equal(renumbered.status, "relocated")
  assert.equal(renumbered.reference.locator.value.object.shapeId, "22")
  assert.deepEqual(renumbered.reasonCodes, [
    "pptx.revision-changed",
    "pptx.slide-id-match",
    "pptx.object-name-type-match",
  ])

  const ambiguous = pptx.resolvePptxReference(context("pptx-v2", fixtures.ambiguousObject), selected)
  assert.equal(ambiguous.status, "ambiguous")
  assert.equal(ambiguous.candidates.length, 2)
  assert.deepEqual(ambiguous.reasonCodes, ["pptx.multiple-object-matches"])

  assert.deepEqual(
    pptx.resolvePptxReference(context("pptx-v2", [{
      slideId: "ppt/slides/slide1.xml",
      objects: [],
    }]), selected),
    { status: "not-found", reasonCode: "pptx.object-not-found" },
  )
})

test("PPTX regions follow reordered slides only while normalized slide space is stable", () => {
  const selected = reference(pptx.createPptxRegionReferenceDraft(original, {
    space: "slide",
    slideIndex: 0,
    rect: { x: 0.2, y: 0.3, width: 0.3, height: 0.25 },
  }), "ref-region")
  const relocated = pptx.resolvePptxReference(context("pptx-v2", fixtures.reordered), selected)
  assert.equal(relocated.status, "relocated")
  assert.equal(relocated.reference.locator.value.slideIndex, 1)

  assert.deepEqual(
    pptx.resolvePptxReference(context("pptx-v2", fixtures.reordered, { width: 1440 }), selected),
    { status: "not-found", reasonCode: "pptx.slide-size-changed" },
  )
})
