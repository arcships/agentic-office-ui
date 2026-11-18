import assert from "node:assert/strict"
import test from "node:test"

import {
  importFromDemo,
  mount,
  vue,
  walk,
} from "./vue-test-renderer.mjs"

const { createBlankDocumentModel } = await importFromDemo("@arcships/docx-core")
const { DocxDocumentSurface } = await importFromDemo("@arcships/vue-docx")

function pointerEvent(currentTarget, clientX, clientY, pointerId = 1, pointerType = "mouse") {
  return {
    button: 0,
    clientX,
    clientY,
    currentTarget,
    isPrimary: true,
    pointerId,
    pointerType,
    preventDefault() {},
  }
}

function enablePointerCapture(element) {
  const captures = new Set()
  element.setPointerCapture = (pointerId) => captures.add(pointerId)
  element.hasPointerCapture = (pointerId) => captures.has(pointerId)
  element.releasePointerCapture = (pointerId) => captures.delete(pointerId)
}

test("DocxDocumentSurface confirms page and touch-region references and resolves them", async () => {
  const confirmations = []
  const resolutions = []
  let api
  let mode
  const Harness = vue.defineComponent({
    setup() {
      const surface = vue.ref(null)
      mode = vue.ref("object")
      vue.watch(surface, (value) => { api = value }, { flush: "sync" })
      return () => vue.h(DocxDocumentSurface, {
        ref: surface,
        model: createBlankDocumentModel(),
        documentId: "surface.docx",
        selectionMode: mode.value,
        emitReferenceCandidates: true,
        onReferenceConfirm: (event) => confirmations.push(event),
        onReferenceResolve: (event) => resolutions.push(event),
      })
    },
  })
  const mounted = await mount(Harness)
  const surface = walk(mounted.root).find((node) => node.props?.class === "docx-document-surface")
  const page = walk(mounted.root).find((node) => node.props?.["data-docx-page-index"] === 0)
  assert.ok(surface)
  assert.ok(page)
  page.dataset = { docxPageIndex: "0" }
  page.closest = (selector) => selector === "[data-docx-page-index]" ? page : null
  page.getBoundingClientRect = () => ({ left: 10, top: 20, width: 900, height: 1200 })
  surface.getBoundingClientRect = () => ({ left: 0, top: 0, width: 920, height: 1240 })
  surface.ownerDocument = { elementsFromPoint: () => [page] }

  surface.props.onClick({ clientX: 100, clientY: 120, detail: 1, shiftKey: false })
  assert.equal(confirmations.at(-1).reference.kind, "page")
  assert.equal(confirmations.at(-1).reference.document.documentId, "surface.docx")
  assert.equal(api.getDocumentRevision().format, "docx")
  const pageResult = await api.resolveReference(confirmations.at(-1).reference)
  assert.equal(pageResult.status, "exact")
  assert.equal(resolutions.at(-1).referenceId, confirmations.at(-1).reference.referenceId)

  mode.value = "region"
  await vue.nextTick()
  const layer = walk(mounted.root).find((node) => node.props?.class === "docx-document-surface__region-layer")
  assert.ok(layer)
  enablePointerCapture(layer)
  layer.props.onPointerdown(pointerEvent(layer, 100, 120, 9, "touch"))
  layer.props.onPointermove(pointerEvent(layer, 300, 420, 9, "touch"))
  layer.props.onPointerup(pointerEvent(layer, 300, 420, 9, "touch"))
  assert.equal(confirmations.at(-1).reference.kind, "region")
  assert.equal(confirmations.at(-1).trigger, "touch")
  assert.equal(confirmations.at(-1).reference.locator.value.space, "page")
  assert.deepEqual(mounted.warnings, [])
  mounted.app.unmount()
})

test("DocxDocumentSurface revisions derive from the host document id and stay stable across mounts", async () => {
  const model = createBlankDocumentModel()
  const revisions = []
  function harness(documentId) {
    return vue.defineComponent({
      setup() {
        const surface = vue.ref(null)
        vue.onMounted(() => revisions.push(surface.value.getDocumentRevision()))
        return () => vue.h(DocxDocumentSurface, {
          ref: surface,
          model,
          ...(documentId ? { documentId } : {}),
        })
      },
    })
  }

  const first = await mount(harness("stable.docx"))
  first.app.unmount()
  const second = await mount(harness("stable.docx"))
  second.app.unmount()
  const third = await mount(harness())
  third.app.unmount()
  const fourth = await mount(harness())
  fourth.app.unmount()

  assert.equal(revisions.length, 4)
  assert.equal(revisions[0].documentId, "stable.docx")
  assert.equal(revisions[1].documentId, "stable.docx")
  assert.equal(revisions[0].revision, revisions[1].revision)
  assert.doesNotMatch(revisions[0].revision, /docx-surface-/u)
  assert.match(revisions[0].revision, /^stable\.docx:/u)
  assert.notEqual(revisions[2].documentId, revisions[3].documentId)
  assert.notEqual(revisions[2].revision, revisions[3].revision)
})

test("DocxDocumentSurface object hit testing ranks image, table and comment candidates and clears on leave", async () => {
  const model = createBlankDocumentModel()
  const paragraph = model.nodes.find((node) => node.type === "paragraph")
  assert.ok(paragraph)
  paragraph.children.push({ type: "image", alt: "Revenue chart", widthPx: 200, heightPx: 100 })
  const imageChildIndex = paragraph.children.length - 1
  const imageParagraphIndex = model.nodes.indexOf(paragraph)
  model.nodes.push({
    type: "table",
    rows: [{ cells: [{ type: "table-cell", nodes: [{ type: "paragraph", children: [{ type: "text", text: "Cell" }] }] }] }],
  })
  const tableIndex = model.nodes.length - 1
  model.metadata.comments = [{ id: 7, author: "Ada", text: "Please revise this sentence." }]

  const candidateEvents = []
  let api
  const Harness = vue.defineComponent({
    setup() {
      const surface = vue.ref(null)
      vue.watch(surface, (value) => { api = value }, { flush: "sync" })
      return () => vue.h(DocxDocumentSurface, {
        ref: surface,
        model,
        documentId: "candidates.docx",
        selectionMode: "object",
        emitReferenceCandidates: true,
        onReferenceCandidateChange: (event) => candidateEvents.push(event),
      })
    },
  })
  const mounted = await mount(Harness)
  const surface = walk(mounted.root).find((node) => node.props?.class === "docx-document-surface")
  assert.ok(surface)

  const page = {
    dataset: { docxPageIndex: "0" },
    closest(selector) { return selector === "[data-docx-page-index]" ? this : null },
  }
  const paragraphElement = {
    dataset: { docxParagraphNodeIndex: String(imageParagraphIndex) },
    closest(selector) {
      if (selector === "[data-docx-paragraph-node-index]") return this
      return page.closest(selector)
    },
  }
  const imageElement = {
    dataset: { docxImageChildIndex: String(imageChildIndex) },
    closest(selector) {
      if (selector === "[data-docx-image-child-index]") return this
      return paragraphElement.closest(selector)
    },
  }
  const tableElement = {
    dataset: { docxTableIndex: String(tableIndex) },
    closest(selector) {
      if (selector === "[data-docx-table-host='true']") return this
      return page.closest(selector)
    },
  }
  const commentElement = {
    dataset: { docxCommentIds: "7" },
    closest(selector) {
      if (selector === "[data-docx-comment-ids], [data-docx-comment-id]") return this
      return paragraphElement.closest(selector)
    },
  }
  surface.ownerDocument = { elementsFromPoint: () => [imageElement, tableElement, commentElement, page] }

  const candidates = api.hitTest({ clientX: 20, clientY: 20 })
  const ids = candidates.map((candidate) => candidate.candidateId)
  assert.ok(ids.includes(`docx:image:p${imageParagraphIndex}:${imageChildIndex}`), ids.join(","))
  assert.ok(ids.includes(`docx:table:${tableIndex}`), ids.join(","))
  assert.ok(ids.includes("docx:comment:7"), ids.join(","))
  assert.equal(candidates.at(-1).candidateId, "docx:page:0")
  assert.equal(candidates[0].hit, "direct")

  surface.props.onPointermove({ clientX: 20, clientY: 20 })
  assert.equal(candidateEvents.at(-1).candidates.length, 4)
  assert.equal(candidateEvents.at(-1).activeCandidateId, candidateEvents.at(-1).candidates[0].candidateId)

  surface.props.onPointerleave()
  assert.deepEqual(candidateEvents.at(-1).candidates, [])
  assert.equal(candidateEvents.at(-1).activeCandidateId, undefined)
  assert.deepEqual(mounted.warnings, [])
  mounted.app.unmount()
})

test("DocxDocumentSurface region drafts start at a point and follow the live page rect", async () => {
  const draftEvents = []
  const confirmations = []
  let mode
  const Harness = vue.defineComponent({
    setup() {
      const surface = vue.ref(null)
      mode = vue.ref("object")
      return () => vue.h(DocxDocumentSurface, {
        ref: surface,
        model: createBlankDocumentModel(),
        documentId: "region.docx",
        selectionMode: mode.value,
        onRegionDraftChange: (event) => draftEvents.push(event),
        onReferenceConfirm: (event) => confirmations.push(event),
      })
    },
  })
  const mounted = await mount(Harness)
  const surface = walk(mounted.root).find((node) => node.props?.class === "docx-document-surface")
  const page = walk(mounted.root).find((node) => node.props?.["data-docx-page-index"] === 0)
  assert.ok(surface)
  assert.ok(page)

  let bounds = { left: 10, top: 20, width: 900, height: 1200 }
  page.dataset = { docxPageIndex: "0" }
  page.closest = (selector) => selector === "[data-docx-page-index]" ? page : null
  page.getBoundingClientRect = () => bounds
  surface.getBoundingClientRect = () => ({ left: 0, top: 0, width: 920, height: 1240 })
  surface.ownerDocument = { elementsFromPoint: () => [page] }
  mode.value = "region"
  await vue.nextTick()
  const layer = walk(mounted.root).find((node) => node.props?.class === "docx-document-surface__region-layer")
  assert.ok(layer)
  enablePointerCapture(layer)

  layer.props.onPointerdown(pointerEvent(layer, 100, 120, 5, "mouse"))
  const start = draftEvents.find((event) => event.phase === "start")
  assert.ok(start)
  assert.deepEqual(start.point, { space: "page", pageIndex: 0, x: 0.1, y: (120 - 20) / 1200 })
  assert.equal("region" in start, false, "start phase must not carry a zero-area rectangle")

  bounds = { left: 10, top: -180, width: 900, height: 1200 }
  layer.props.onPointermove(pointerEvent(layer, 300, 420, 5, "mouse"))
  const change = draftEvents.at(-1)
  assert.equal(change.phase, "change")
  assert.equal(change.region.rect.x, 0.1)
  assert.equal(change.region.rect.width, (300 - 10) / 900 - 0.1)
  assert.equal(change.region.rect.y, (120 + 180) / 1200)
  assert.equal(change.region.rect.height, (420 + 180) / 1200 - (120 + 180) / 1200)
  assert.ok(change.region.rect.width > 0 && change.region.rect.height > 0)

  layer.props.onPointerup(pointerEvent(layer, 300, 420, 5, "mouse"))
  assert.equal(confirmations.at(-1).reference.kind, "region")
  assert.deepEqual(mounted.warnings, [])
  mounted.app.unmount()
})

test("DocxDocumentSurface emits declared resolve, describe and capture error codes", async () => {
  const errors = []
  const confirmations = []
  let api
  let model
  const Harness = vue.defineComponent({
    setup() {
      const surface = vue.ref(null)
      model = vue.ref(createBlankDocumentModel())
      vue.watch(surface, (value) => { api = value }, { flush: "sync" })
      return () => vue.h(DocxDocumentSurface, {
        ref: surface,
        model: model.value,
        documentId: "errors.docx",
        selectionMode: "object",
        onReferenceConfirm: (event) => confirmations.push(event),
        onReferenceError: (event) => errors.push(event),
      })
    },
  })
  const mounted = await mount(Harness)
  const surface = walk(mounted.root).find((node) => node.props?.class === "docx-document-surface")
  const page = walk(mounted.root).find((node) => node.props?.["data-docx-page-index"] === 0)
  assert.ok(surface)
  assert.ok(page)
  page.dataset = { docxPageIndex: "0" }
  page.closest = (selector) => selector === "[data-docx-page-index]" ? page : null
  page.getBoundingClientRect = () => ({ left: 10, top: 20, width: 900, height: 1200 })
  surface.getBoundingClientRect = () => ({ left: 0, top: 0, width: 920, height: 1240 })
  surface.ownerDocument = { elementsFromPoint: () => [page] }

  surface.props.onClick({ clientX: 100, clientY: 120, detail: 1, shiftKey: false })
  const confirmed = confirmations.at(-1).reference
  assert.equal(confirmed.kind, "page")
  assert.equal((await api.resolveReference(confirmed)).status, "exact")
  assert.deepEqual(errors, [])

  model.value = {
    ...model.value,
    nodes: [...model.value.nodes, { type: "paragraph", children: [{ type: "text", text: "Updated content" }] }],
    metadata: { ...model.value.metadata },
  }
  await vue.nextTick()
  const revisionResult = await api.resolveReference(confirmed)
  assert.equal(revisionResult.status, "unsupported")
  assert.equal(errors.at(-1).code, "REVISION_CHANGED")
  assert.equal(errors.at(-1).operation, "resolve")

  const wrongDocument = structuredClone(confirmed)
  wrongDocument.document.documentId = "other.docx"
  await assert.rejects(api.describeReference(wrongDocument))
  assert.equal(errors.at(-1).code, "REVISION_CHANGED")
  assert.equal(errors.at(-1).operation, "describe")

  const missingPage = structuredClone(confirmed)
  missingPage.locator.value.pageIndex = 99
  await assert.rejects(api.describeReference(missingPage))
  assert.equal(errors.at(-1).code, "DESCRIBE_FAILED")

  const paragraphStructure = structuredClone(confirmed)
  paragraphStructure.kind = "paragraph"
  paragraphStructure.locator.value = {
    kind: "structure",
    part: { kind: "body" },
    path: [{ kind: "paragraph", index: 0 }],
  }
  assert.equal((await api.resolveReference(paragraphStructure)).status, "relocated")

  const unsupported = structuredClone(confirmed)
  unsupported.locator.value = {
    kind: "behavior",
    behavior: "bookmark",
    owner: { scope: "document" },
    instanceId: "bookmark-1",
  }
  await api.resolveReference(unsupported)
  assert.equal(errors.at(-1).code, "RESOLVE_FAILED")

  await assert.rejects(api.captureReferencePreview(confirmed, { maxWidth: 20_000 }))
  assert.equal(errors.at(-1).code, "CAPTURE_LIMIT_EXCEEDED")
  assert.equal(errors.at(-1).operation, "capture")
  mounted.app.unmount()
})
