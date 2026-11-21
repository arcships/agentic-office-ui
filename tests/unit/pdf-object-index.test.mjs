import assert from "node:assert/strict"
import test from "node:test"

const pdf = await import(
  new URL("../../packages/vue-pdf/dist/index.js", import.meta.url).href
)
const interaction = await import(
  new URL("../../packages/office-interaction/dist/index.js", import.meta.url).href
)

function page(index) {
  return { index, width: 600, height: 800, rotation: 0 }
}

function context(revision = "pdf-rev-1", overrides = {}) {
  const document = {
    id: `pdf-runtime-${revision}`,
    pageCount: 2,
    pages: [page(0), page(1)],
  }
  return {
    revision: { format: "pdf", documentId: "brief.pdf", revision },
    document,
    runtime: {
      async getTextSlices() {
        return []
      },
      async search() {
        return []
      },
    },
    ...overrides,
  }
}

const annotations = [
  {
    id: "link-7",
    pageIndex: 0,
    kind: "link",
    rect: { x: 30, y: 60, width: 180, height: 20 },
    url: "https://example.com/report",
  },
  {
    id: "widget-3",
    pageIndex: 0,
    kind: "widget",
    rect: { x: 40, y: 200, width: 120, height: 24 },
    fieldName: "Customer",
    fieldType: "text",
    fieldValue: "Ada",
  },
  {
    id: "annot-9",
    pageIndex: 1,
    kind: "annotation",
    rect: { x: 20, y: 40, width: 60, height: 30 },
    contents: "Please review",
    flags: ["printable"],
  },
]

function confirm(draft, referenceId) {
  return interaction.confirmOfficeReferenceDraft(draft, { referenceId, trigger: "pointer" })
}

test("PDF link exemplar confirms with descriptor, fingerprint and documented reliability", async () => {
  const ctx = context("pdf-rev-1", { annotations })
  const link = annotations[0]
  const draft = pdf.createPdfNativeObjectReferenceDraft(ctx, link)
  assert.equal(draft.kind, "link")
  assert.equal(draft.source, "native")
  assert.deepEqual(draft.locator.value, {
    kind: "native-object",
    pageIndex: 0,
    objectType: "link",
    nativeId: "link-7",
  })
  assert.ok(draft.fingerprint.contentHash.startsWith("fnv1a32:"))
  assert.ok(draft.fingerprint.ancestorKeys.includes("pdf.native-object:link:link-7"))
  assert.ok(draft.fallbackRegion)
  assert.deepEqual(draft.fallbackRegion.rect, { x: 0.05, y: 0.075, width: 0.3, height: 0.025 })
  assert.equal(draft.reliability.semantic.level, "exact")
  assert.equal(draft.reliability.semantic.reasonCodes[0], "pdf.native-link")
  assert.equal(draft.reliability.relocation.level, "likely")

  const confirmed = confirm(draft, "ref-pdf-link")
  assert.equal(confirmed.reference.kind, "link")
  assert.deepEqual(confirmed.reference.fingerprint, draft.fingerprint)

  const resolved = await pdf.resolvePdfReference(ctx, confirmed.reference)
  assert.equal(resolved.status, "exact")
  assert.equal(resolved.descriptor.label, "Link")
  assert.equal(resolved.descriptor.draft.kind, "link")
  assert.equal(resolved.descriptor.visual.fragments[0].rect.height, 0.025)
  assert.equal(resolved.descriptor.facets.url, "https://example.com/report")

  const moved = await pdf.resolvePdfReference(context("pdf-rev-2", { annotations }), confirmed.reference)
  assert.equal(moved.status, "relocated")
  assert.equal(moved.reference.document.revision, "pdf-rev-2")

  const withoutEvidence = context("pdf-rev-1")
  assert.equal((await pdf.resolvePdfReference(withoutEvidence, confirmed.reference)).status, "unsupported")

  const form = pdf.createPdfNativeObjectReferenceDraft(ctx, annotations[1])
  assert.equal(form.kind, "form-control")
  assert.equal(form.locator.value.objectType, "form")
  const formResolved = await pdf.resolvePdfReference(ctx, confirm(form, "ref-pdf-form").reference)
  assert.equal(formResolved.status, "exact")
  assert.equal(formResolved.descriptor.label, "Form field: Customer")
  assert.equal(formResolved.descriptor.content.value, "Ada")

  const annotation = pdf.createPdfNativeObjectReferenceDraft(ctx, annotations[2])
  assert.equal(annotation.kind, "annotation")
  const annotationResolved = await pdf.resolvePdfReference(ctx, confirm(annotation, "ref-pdf-annot").reference)
  assert.equal(annotationResolved.status, "exact")
  assert.equal(annotationResolved.descriptor.label, "Annotation: Please review")
})

test("PDF visual recognizer results become visual locators, not native objects", async () => {
  const detection = {
    objectId: "table-1",
    kind: "table",
    region: { x: 0.1, y: 0.2, width: 0.5, height: 0.3 },
    label: "Revenue table",
    content: { text: "Region Revenue" },
    reliability: {
      semantic: { level: "likely", reasonCodes: ["visual.table-shape"] },
      boundary: { level: "likely", reasonCodes: ["visual.detected-region"] },
      hierarchy: { level: "uncertain", reasonCodes: ["visual.parent-unknown"] },
      relocation: { level: "uncertain", reasonCodes: ["visual.features"] },
    },
  }
  const ctx = context("pdf-rev-1", {
    visualObjects: [{ ...detection, pageIndex: 0 }],
  })
  const draft = pdf.createPdfVisualObjectReferenceDraft(ctx, "host-ocr", 0, detection)
  assert.equal(draft.kind, "table")
  assert.equal(draft.source, "visual")
  assert.equal(draft.locator.value.kind, "visual-object")
  assert.equal(draft.locator.value.providerId, "host-ocr")
  assert.equal(draft.reliability.semantic.level, "likely")
  assert.deepEqual(draft.fallbackRegion.rect, detection.region)

  const resolved = await pdf.resolvePdfReference(ctx, confirm(draft, "ref-pdf-visual").reference)
  assert.equal(resolved.status, "exact")
  assert.equal(resolved.descriptor.label, "Revenue table")
  assert.equal(resolved.descriptor.childrenState, "available")

  const registryProvider = pdf.PDF_VISUAL_OBJECT_PROVIDER
  const listed = registryProvider.list(ctx)
  assert.equal(listed.length, 1)
  assert.equal(listed[0].draft.locator.value.providerId, "visual")

  assert.throws(
    () => pdf.createPdfVisualObjectReferenceDraft(ctx, "host-ocr", 0, {
      ...detection,
      region: { x: 0.9, y: 0.2, width: 0.5, height: 0.3 },
    }),
    /normalized page space/u,
  )
})

test("PDF stage 4 behaviors cover link actions, form rules, annotation state and document-level state", async () => {
  const behaviors = pdf.createPdfBehaviorRecords({
    annotations,
    attachments: [{ id: "attach-1", name: "source.csv" }],
    layers: [{ id: "ocg-2", name: "Draft watermark", visible: false }],
    signatures: [{ id: "sig-1", name: "Approval", signer: "Ada", valid: true }],
    permissions: { copying: false, printing: true },
  })
  assert.deepEqual(behaviors.map((record) => record.behavior), [
    "link-action",
    "form-rule",
    "annotation-state",
    "attachment",
    "layer",
    "signature",
    "permission",
  ])
  const ctx = context("pdf-rev-1", { behaviors })

  const action = behaviors.find((record) => record.behavior === "link-action")
  const formRule = behaviors.find((record) => record.behavior === "form-rule")
  const annotationState = behaviors.find((record) => record.behavior === "annotation-state")
  assert.equal(annotationState.pageIndex, 1)

  for (const record of behaviors) {
    const descriptor = pdf.describePdfBehavior(ctx, record)
    assert.equal(interaction.isOfficeObjectVisible(descriptor), false)
    assert.equal(descriptor.draft.locator.value.kind, "behavior")
  }

  const actionDraft = pdf.createPdfBehaviorReferenceDraft(ctx, action)
  assert.equal(actionDraft.kind, "action")
  assert.equal(actionDraft.locator.value.nativeId, action.instanceId)
  const actionResolved = await pdf.resolvePdfReference(ctx, confirm(actionDraft, "ref-pdf-action").reference)
  assert.equal(actionResolved.status, "exact")
  assert.equal(actionResolved.descriptor.facets.url, "https://example.com/report")

  const formDraft = pdf.createPdfBehaviorReferenceDraft(ctx, formRule)
  assert.equal(formDraft.kind, "form-rule")
  const formResolved = await pdf.resolvePdfReference(ctx, confirm(formDraft, "ref-pdf-rule").reference)
  assert.equal(formResolved.status, "exact")
  assert.equal(formResolved.descriptor.label, "Form field: Customer")

  const localized = interaction.localizeOfficeDescriptorLabel(
    formResolved.descriptor.label,
    interaction.resolveOfficeMessages("zh-CN").common,
  )
  assert.equal(localized, "表单控件：Customer")

  const layer = behaviors.find((record) => record.behavior === "layer")
  assert.equal((await pdf.resolvePdfReference(ctx, confirm(
    pdf.createPdfBehaviorReferenceDraft(ctx, layer), "ref-pdf-layer",
  ).reference)).status, "exact")
})

test("PDF object registry lists native, visual and behavior families", () => {
  const registry = pdf.createPdfObjectRegistry()
  assert.deepEqual(registry.providers().map((provider) => provider.id), [
    "pdf.native-objects",
    "pdf.visual-objects",
    "pdf.behaviors",
  ])
  const behaviors = pdf.createPdfBehaviorRecords({ annotations })
  const ctx = context("pdf-rev-1", {
    annotations,
    behaviors,
    visualObjects: [{
      objectId: "block-1",
      pageIndex: 0,
      kind: "text-block",
      region: { x: 0.1, y: 0.1, width: 0.4, height: 0.1 },
      reliability: {
        semantic: { level: "likely", reasonCodes: ["visual.text-block"] },
        boundary: { level: "likely", reasonCodes: ["visual.region"] },
        hierarchy: { level: "uncertain", reasonCodes: ["visual.parent-unknown"] },
        relocation: { level: "uncertain", reasonCodes: ["visual.features"] },
      },
    }],
  })
  const kinds = new Set(registry.list(ctx).map((descriptor) => descriptor.draft.kind))
  for (const kind of ["link", "form-control", "annotation", "text-block", "action", "form-rule"]) {
    assert.ok(kinds.has(kind), `registry is missing ${kind}`)
  }
  const nativeOnly = registry.list(ctx, { kinds: ["link"] })
  assert.equal(nativeOnly.length, 1)
  assert.equal(nativeOnly[0].label, "Link")
})

test("PDF attachments, OCG layers, signatures and permissions resolve from host behaviorSource", async () => {
  const behaviorSource = {
    attachments: [{ id: "attach-1", name: "source.csv", description: "Original export" }],
    layers: [{ id: "ocg-2", name: "Draft watermark", visible: false }],
    signatures: [{ id: "sig-1", name: "Approval", signer: "Ada", valid: true }],
    permissions: { copying: false, printing: true },
  }
  const ctx = context("pdf-rev-1", { behaviorSource })
  const records = pdf.resolvePdfBehaviorRecords(ctx)
  assert.deepEqual(records.map((record) => record.behavior), [
    "attachment", "layer", "signature", "permission",
  ])
  assert.equal(pdf.resolvePdfBehaviorRecords(ctx), records, "behaviorSource records are cached per context")

  const expectedKinds = {
    attachment: "attachment",
    layer: "layer",
    signature: "signature",
    permission: "permission",
  }
  const registry = pdf.createPdfObjectRegistry()
  const listed = registry.list(ctx, { kinds: ["attachment", "layer", "signature", "permission"] })
  assert.equal(listed.length, 4)
  assert.ok(listed.every((descriptor) => interaction.isOfficeObjectVisible(descriptor) === false))

  for (const record of records) {
    const descriptor = pdf.describePdfBehavior(ctx, record)
    assert.equal(descriptor.draft.kind, expectedKinds[record.behavior])
    assert.equal(descriptor.draft.locator.value.kind, "behavior")
    assert.equal(descriptor.draft.locator.value.nativeId, record.instanceId)
    assert.ok(descriptor.draft.fingerprint.ancestorKeys.includes(
      `pdf.behavior-instance:${record.instanceId}`,
    ))
    assert.equal("visual" in descriptor, false)

    const draft = pdf.createPdfBehaviorReferenceDraft(ctx, record)
    const confirmed = confirm(draft, `ref-pdf-${record.behavior}`)
    const resolved = await pdf.resolvePdfReference(ctx, confirmed.reference)
    assert.equal(resolved.status, "exact")
    assert.equal(resolved.descriptor.label, record.label)
    assert.equal(resolved.descriptor.draft.kind, expectedKinds[record.behavior])
    assert.deepEqual(resolved.descriptor.path.map((entry) => entry.kind), record.path.map((entry) => entry.kind))

    const moved = await pdf.resolvePdfReference(context("pdf-rev-2", { behaviorSource }), confirmed.reference)
    assert.equal(moved.status, "relocated")
    assert.equal(moved.reference.document.revision, "pdf-rev-2")
  }

  const localized = interaction.localizeOfficeDescriptorLabel(
    listed.find((descriptor) => descriptor.draft.kind === "attachment").label,
    interaction.resolveOfficeMessages("zh-CN").common,
  )
  assert.equal(localized, "附件：source.csv")

  const withPrebuilt = context("pdf-rev-1", { behaviorSource, behaviors: [] })
  assert.deepEqual(pdf.resolvePdfBehaviorRecords(withPrebuilt), [])
  assert.throws(
    () => pdf.createPdfBehaviorReferenceDraft(context("pdf-rev-3"), records[0]),
    /not present in the reference context/u,
  )
})
