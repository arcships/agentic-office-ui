import assert from "node:assert/strict"
import { readFileSync } from "node:fs"
import test from "node:test"

const packageRoot = new URL("../../packages/office-interaction/", import.meta.url)
const interaction = await import(new URL("dist/index.js", packageRoot).href)

const reliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["native-id"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["native-geometry"] },
  hierarchy: { level: "likely", reasonCodes: ["model-path"] },
  relocation: { level: "likely", reasonCodes: ["text-quote"] },
}

const docxRevision = {
  format: "docx",
  documentId: "contract.docx",
  revision: "sha256:docx-v1",
  contentDigest: `sha256:${"a".repeat(64)}`,
}

const docxTextReference = {
  schemaVersion: 1,
  referenceId: "ref-docx-text",
  document: docxRevision,
  kind: "text-range",
  source: "native",
  locator: {
    type: "format",
    format: "docx",
    value: {
      kind: "text-range",
      part: { kind: "body" },
      start: { path: [{ kind: "node", index: 2 }], offset: 4 },
      end: { path: [{ kind: "node", index: 2 }], offset: 12 },
    },
  },
  fingerprint: { exactText: "selected", prefixText: "before", suffixText: "after" },
  fallbackRegion: { space: "page", pageIndex: 0, rect: { x: 0.1, y: 0.2, width: 0.3, height: 0.1 } },
  reliability,
}

function expectInvalid(value, code) {
  const result = interaction.safeParseOfficeObjectReference(value)
  assert.equal(result.success, false)
  assert.equal(result.error.issues[0].code, code)
}

test("office-interaction is publishable, dependency-free, and does not leak Vue or office-runtime", () => {
  const manifest = JSON.parse(readFileSync(new URL("package.json", packageRoot), "utf8"))
  const declaration = readFileSync(new URL("dist/index.d.ts", packageRoot), "utf8")
  assert.equal(manifest.private, undefined)
  assert.equal(manifest.publishConfig.access, "public")
  assert.deepEqual(manifest.dependencies, undefined)
  assert.equal(manifest.peerDependencies, undefined)
  assert.doesNotMatch(declaration, /(?:@arcships\/office-runtime|\bvue\b|\.vue)/u)
})

test("reference parsing enforces format, kind, source, locator, and confidence invariants", () => {
  assert.deepEqual(interaction.parseOfficeObjectReference(docxTextReference), docxTextReference)

  expectInvalid({
    ...docxTextReference,
    locator: { ...docxTextReference.locator, format: "xlsx" },
  }, "FORMAT_MISMATCH")

  expectInvalid({ ...docxTextReference, kind: "chart" }, "KIND_LOCATOR_MISMATCH")

  expectInvalid({
    ...docxTextReference,
    kind: "region",
    source: "native",
    locator: {
      type: "manual-region",
      format: "docx",
      value: { space: "page", pageIndex: 0, rect: { x: 0, y: 0, width: 1, height: 1 } },
    },
  }, "KIND_LOCATOR_MISMATCH")

  expectInvalid({
    ...docxTextReference,
    reliability: { ...reliability, semantic: { level: "likely", score: 1.2, reasonCodes: [] } },
  }, "INVALID_VALUE")
})

test("format behavior locators require deterministic instance identity and valid scope", () => {
  const xlsxRevision = { format: "xlsx", documentId: "book", revision: "book-v1" }
  const conditionalFormat = {
    schemaVersion: 1,
    referenceId: "conditional-1",
    document: xlsxRevision,
    kind: "conditional-format",
    source: "native",
    locator: {
      type: "format",
      format: "xlsx",
      value: {
        kind: "behavior",
        behavior: "conditional-format",
        scope: {
          kind: "sheet",
          sheet: { sheetId: "sheet-1", name: "Sales", index: 0 },
          appliesToA1: "A2:D20",
        },
        instanceId: "rule:0:hash",
      },
    },
    reliability,
  }
  assert.equal(interaction.parseOfficeObjectReference(conditionalFormat).kind, "conditional-format")

  const missingId = structuredClone(conditionalFormat)
  delete missingId.locator.value.instanceId
  expectInvalid(missingId, "INVALID_TYPE")

  const wrongScope = structuredClone(conditionalFormat)
  wrongScope.locator.value.scope = { kind: "workbook" }
  expectInvalid(wrongScope, "INVALID_VALUE")

  const namedRange = structuredClone(conditionalFormat)
  namedRange.referenceId = "name-1"
  namedRange.kind = "named-range"
  namedRange.locator.value.behavior = "named-range"
  namedRange.locator.value.scope = { kind: "workbook" }
  namedRange.locator.value.instanceId = "Revenue"
  assert.equal(interaction.parseOfficeObjectReference(namedRange).kind, "named-range")

  const zeroHeightRegion = {
    ...conditionalFormat,
    referenceId: "region-1",
    kind: "region",
    source: "manual",
    locator: {
      type: "manual-region",
      format: "xlsx",
      value: {
        space: "sheet",
        sheetId: "sheet-1",
        start: { row: 2, col: 1, xOffset: 0.1, yOffset: 0.2 },
        end: { row: 2, col: 4, xOffset: 0.7, yOffset: 0.2 },
      },
    },
  }
  expectInvalid(zeroHeightRegion, "INVALID_VALUE")

  const docxField = structuredClone(docxTextReference)
  docxField.referenceId = "field-1"
  docxField.kind = "field"
  docxField.locator.value = {
    kind: "behavior",
    behavior: "field",
    owner: { scope: "part", part: { kind: "header", partName: "word/header1.xml" }, path: [{ kind: "paragraph", index: 0 }] },
    instanceId: "field:path:0",
  }
  assert.equal(interaction.parseOfficeObjectReference(docxField).kind, "field")
})

test("PPTX sub-objects and PDF visual objects keep their semantic source contracts", () => {
  const pptx = {
    schemaVersion: 1,
    referenceId: "pptx-cell",
    document: { format: "pptx", documentId: "deck", revision: "deck-v1" },
    kind: "table-cell",
    source: "native",
    locator: {
      type: "format",
      format: "pptx",
      value: {
        kind: "sub-object",
        slide: { slideId: "256", index: 0 },
        object: { objectKey: "slide1:shape4", shapeId: "4", source: "slide", groupPath: [] },
        path: [{ kind: "table-cell", rowIndex: 1, cellIndex: 2 }],
      },
    },
    reliability,
  }
  assert.equal(interaction.parseOfficeObjectReference(pptx).kind, "table-cell")

  const pdfVisual = {
    schemaVersion: 1,
    referenceId: "pdf-table",
    document: { format: "pdf", documentId: "scan", revision: "scan-v1" },
    kind: "table",
    source: "visual",
    locator: {
      type: "format",
      format: "pdf",
      value: {
        kind: "visual-object",
        pageIndex: 1,
        providerId: "local-layout",
        objectId: "table-2",
        region: { x: 0.1, y: 0.1, width: 0.8, height: 0.5 },
      },
    },
    reliability,
  }
  assert.equal(interaction.parseOfficeObjectReference(pdfVisual).source, "visual")
  expectInvalid({ ...pdfVisual, source: "native" }, "KIND_LOCATOR_MISMATCH")
})

test("canonical serialization is stable and rejects Blob URLs and cycles", () => {
  const first = interaction.serializeOfficeObjectReference(docxTextReference)
  const second = interaction.serializeOfficeObjectReference(structuredClone(docxTextReference))
  assert.equal(first, second)
  assert.deepEqual(interaction.deserializeOfficeObjectReference(first), docxTextReference)
  assert.ok(first.indexOf('"document"') < first.indexOf('"fingerprint"'))

  assert.throws(
    () => interaction.canonicalStringifyOfficeJson({ previewUrl: "blob:https://example.test/secret" }),
    (error) => error?.code === "INVALID_OFFICE_INTERACTION",
  )
  const circular = {}
  circular.self = circular
  assert.throws(
    () => interaction.canonicalStringifyOfficeJson(circular),
    (error) => error?.code === "INVALID_OFFICE_INTERACTION",
  )
})

test("geometry helpers normalize drag direction and sheet cell offsets", () => {
  assert.deepEqual(
    interaction.normalizedRectFromPoints({ x: 0.8, y: 0.9 }, { x: 0.2, y: 0.3 }),
    { x: 0.2, y: 0.3, width: 0.6000000000000001, height: 0.6000000000000001 },
  )
  assert.equal(
    interaction.normalizedRectContains({ x: 0.1, y: 0.1, width: 0.4, height: 0.4 }, { x: 0.2, y: 0.3 }),
    true,
  )
  assert.deepEqual(
    interaction.intersectNormalizedRects(
      { x: 0, y: 0, width: 0.5, height: 0.5 },
      { x: 0.25, y: 0.25, width: 0.5, height: 0.5 },
    ),
    { x: 0.25, y: 0.25, width: 0.25, height: 0.25 },
  )
  assert.deepEqual(
    interaction.normalizeSheetRegion(
      "sheet-1",
      { row: 8, col: 4, xOffset: 0.7, yOffset: 0.9 },
      { row: 2, col: 1, xOffset: 0.1, yOffset: 0.2 },
    ),
    {
      space: "sheet",
      sheetId: "sheet-1",
      start: { row: 2, col: 1, xOffset: 0.1, yOffset: 0.2 },
      end: { row: 8, col: 4, xOffset: 0.7, yOffset: 0.9 },
    },
  )
  assert.throws(
    () => interaction.normalizeSheetRegion(
      "sheet-1",
      { row: 2, col: 1, xOffset: 0.1, yOffset: 0.2 },
      { row: 2, col: 4, xOffset: 0.7, yOffset: 0.2 },
    ),
    /positive area/u,
  )
})

test("candidate ranking is deterministic and never exposes runtime targets in previews", () => {
  const { referenceId: _referenceId, ...draft } = docxTextReference
  const base = {
    draft,
    depth: 1,
    preview: {
      label: "paragraph",
      path: [{ kind: "paragraph", label: "Paragraph 1" }],
      visual: {
        layoutVersion: "1",
        fragments: [{ container: { space: "page", pageIndex: 0 }, rect: { x: 0, y: 0, width: 0.5, height: 0.5 } }],
      },
    },
  }
  const candidates = [
    { ...base, candidateId: "inferred", hit: "inferred", runtimeTarget: { secret: true } },
    { ...base, candidateId: "direct", hit: "direct", runtimeTarget: { secret: true } },
  ]
  const ranked = interaction.rankOfficeHitCandidates(candidates)
  assert.deepEqual(ranked.map((candidate) => candidate.candidateId), ["direct", "inferred"])
  assert.equal("runtimeTarget" in interaction.toOfficeReferenceCandidatePreview(ranked[0]), false)
})

test("selection session navigation follows keyboard priority without committing implicitly", () => {
  const { referenceId: _referenceId, ...draft } = docxTextReference
  const candidates = ["paragraph", "page"].map((candidateId, index) => ({
    candidateId,
    draft: {
      ...draft,
      kind: index === 0 ? "paragraph" : "page",
      locator: {
        type: "format",
        format: "docx",
        value: index === 0
          ? { kind: "structure", part: { kind: "body" }, path: [{ kind: "paragraph", index: 2 }] }
          : { kind: "page", pageIndex: 0 },
      },
    },
    preview: { label: candidateId, path: [] },
    hit: index === 0 ? "direct" : "ancestor",
    depth: 2 - index,
  }))
  assert.throws(
    () => interaction.createOfficeCandidateNavigationState([candidates[0], candidates[0]]),
    /duplicate candidateId/u,
  )
  assert.throws(
    () => interaction.createOfficeCandidateNavigationState(Array.from(
      { length: interaction.MAX_OFFICE_VISIBLE_CANDIDATES + 1 },
      (_, index) => ({ ...candidates[0], candidateId: `candidate-${index}` }),
    )),
    /at most 20/u,
  )
  assert.throws(
    () => interaction.createOfficeSelectionSessionState({ mode: "layers" }),
    /unsupported Office selection mode/u,
  )

  let state = interaction.createOfficeSelectionSessionState({ mode: "object", candidates })
  assert.equal(state.activeCandidateId, "paragraph")
  assert.equal(state.phase, "choosing")

  let result = interaction.applyOfficeSelectionKeyboard(state, { key: "Tab" })
  assert.equal(result.command, "next-candidate")
  assert.equal(result.state.activeCandidateId, "page")
  state = result.state

  result = interaction.applyOfficeSelectionKeyboard(state, { key: "Tab", shiftKey: true })
  assert.equal(result.command, "previous-candidate")
  assert.equal(result.state.activeCandidateId, "paragraph")
  state = result.state

  result = interaction.applyOfficeSelectionKeyboard(state, { key: "Enter" })
  assert.equal(result.command, "confirm-candidate")
  assert.equal(result.activeCandidate.candidateId, "paragraph")
  assert.equal(result.state, state)

  result = interaction.applyOfficeSelectionKeyboard(state, { key: "Escape" })
  assert.equal(result.command, "dismiss-candidates")
  assert.equal(result.state.mode, "object")
  assert.equal(result.state.phase, "pointing")
  assert.deepEqual(result.state.candidates, [])

  result = interaction.applyOfficeSelectionKeyboard(result.state, { key: "Escape" })
  assert.equal(result.command, undefined)
  assert.equal(result.handled, false)
  assert.equal(result.state.mode, "object")

  const hierarchy = interaction.applyOfficeSelectionKeyboard(state, { key: "ArrowRight" }, { canEnterChild: true })
  assert.equal(hierarchy.command, "enter-child")
  assert.equal(hierarchy.handled, true)
})

test("candidate confirmation emits one serializable reference without owning host intent", () => {
  const { referenceId: _referenceId, ...draft } = docxTextReference
  const candidate = {
    candidateId: "paragraph",
    draft,
    preview: { label: "Selected text", path: [{ kind: "paragraph", label: "Paragraph 3" }] },
    hit: "direct",
    depth: 2,
  }
  const state = interaction.createOfficeSelectionSessionState({ mode: "object", candidates: [candidate] })
  const result = interaction.confirmOfficeCandidate(state, {
    referenceId: "selection-1",
    trigger: "keyboard",
    additiveRequested: true,
    snapshot: {
      label: "Selected text",
      path: [{ kind: "paragraph", label: "Paragraph 3" }],
      content: { text: "selected" },
    },
  })

  assert.equal(result.event.reference.referenceId, "selection-1")
  assert.equal(result.event.trigger, "keyboard")
  assert.equal(result.event.additiveRequested, true)
  assert.equal("role" in result.event, false)
  assert.equal("action" in result.event, false)
  assert.equal(result.state.mode, "object")
  assert.equal(result.state.phase, "pointing")
  assert.deepEqual(result.state.candidates, [])

  const serialized = interaction.serializeOfficeReferenceConfirmEvent(result.event)
  assert.deepEqual(interaction.deserializeOfficeReferenceConfirmEvent(serialized), result.event)
  assert.equal(interaction.safeParseOfficeReferenceConfirmEvent({ ...result.event, trigger: "agent" }).success, false)
  assert.equal(interaction.safeParseOfficeReferenceConfirmEvent({ ...result.event, role: "target" }).success, false)
  assert.throws(
    () => interaction.confirmOfficeCandidate(state, {
      candidateId: "missing",
      referenceId: "selection-2",
      trigger: "pointer",
    }),
    /unknown candidateId/u,
  )

  assert.match(interaction.createOfficeReferenceId(), /^ref_[0-9a-f-]{36}$/u)
})

test("region session cancellation clears only transient state and keeps host-controlled mode", () => {
  const first = { space: "page", pageIndex: 0, rect: { x: 0.1, y: 0.2, width: 0.2, height: 0.3 } }
  const second = { space: "page", pageIndex: 0, rect: { x: 0.1, y: 0.2, width: 0.4, height: 0.5 } }
  let state = interaction.createOfficeSelectionSessionState({ mode: "region" })
  state = interaction.reduceOfficeSelectionSession(state, { type: "begin-region", region: first })
  state = interaction.reduceOfficeSelectionSession(state, { type: "update-region", region: second })
  assert.equal(state.phase, "drawing")
  assert.deepEqual(state.regionDraft, second)

  const result = interaction.applyOfficeSelectionKeyboard(state, { key: "Escape" })
  assert.equal(result.command, "cancel-selection")
  assert.equal(result.state.mode, "region")
  assert.equal(result.state.phase, "idle")
  assert.equal(result.state.regionDraft, undefined)
})
