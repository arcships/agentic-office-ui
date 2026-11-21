import assert from "node:assert/strict"
import test from "node:test"

import {
  findByTestId,
  importFromDemo,
  mount,
  vue,
  walk,
  waitFor,
} from "./vue-test-renderer.mjs"

const { PdfSurface } = await importFromDemo("@arcships/vue-pdf")

const PDF_SOURCE = {
  kind: "blob",
  blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }),
}

const LINK = {
  id: "link-7",
  pageIndex: 0,
  kind: "link",
  rect: { x: 30, y: 60, width: 180, height: 20 },
  url: "https://example.com/report",
}

function createRuntime({ annotations = [LINK] } = {}) {
  const calls = { getPageAnnotations: 0 }
  const document = {
    id: "pdf-object-document",
    pageCount: 2,
    pages: [
      { index: 0, width: 600, height: 800, rotation: 0 },
      { index: 1, width: 600, height: 800, rotation: 0 },
    ],
  }
  return {
    calls,
    document,
    runtime: {
      runtimeId: "pdf-object-runtime",
      async openDocument() {
        return document
      },
      async renderPage(_document, pageIndex) {
        return new Blob([`page:${pageIndex}`], { type: "image/png" })
      },
      async renderThumbnail(_document, pageIndex) {
        return new Blob([`thumbnail:${pageIndex}`], { type: "image/png" })
      },
      async getPageLabels() {
        return ["1", "2"]
      },
      async getOutline() {
        return []
      },
      async getPageTextGeometry() {
        return { runs: [] }
      },
      async getTextSlices(_document, slices) {
        return slices.map(() => "")
      },
      async search() {
        return []
      },
      async getPageAnnotations(_document, pageIndex) {
        calls.getPageAnnotations += 1
        return annotations.filter((annotation) => annotation.pageIndex === pageIndex)
      },
      async closeDocument() {},
      async dispose() {},
    },
  }
}

function stubObjectUrls(t) {
  const originalCreate = URL.createObjectURL
  const originalRevoke = URL.revokeObjectURL
  let counter = 0
  URL.createObjectURL = () => `blob:pdf-object-${++counter}`
  URL.revokeObjectURL = () => {}
  t.after(() => {
    URL.createObjectURL = originalCreate
    URL.revokeObjectURL = originalRevoke
  })
}

function makeScrollable(node, { clientHeight = 600, clientWidth = 900 } = {}) {
  Object.defineProperty(node, "clientHeight", { configurable: true, get: () => clientHeight })
  Object.defineProperty(node, "clientWidth", { configurable: true, get: () => clientWidth })
  Object.defineProperty(node, "scrollTop", { configurable: true, writable: true, value: 0 })
  Object.defineProperty(node, "scrollLeft", { configurable: true, writable: true, value: 0 })
  return node
}

function annotationElement(node) {
  const element = {
    dataset: {
      annotationId: node.props["data-annotation-id"],
      annotationPage: node.props["data-annotation-page"],
      annotationKind: node.props["data-annotation-kind"],
    },
    getBoundingClientRect: () => ({ left: 20, top: 40, width: 180, height: 20, right: 200, bottom: 60 }),
  }
  element.closest = (selector) => selector.includes("[data-annotation-id]") ? element : null
  return element
}

async function mountSurface(t) {
  stubObjectUrls(t)
  const fake = createRuntime()
  const confirmations = []
  const candidateChanges = []
  const Harness = vue.defineComponent({
    setup() {
      return () => vue.h(PdfSurface, {
        source: PDF_SOURCE,
        runtime: fake.runtime,
        documentId: "object-mode.pdf",
        selectionMode: "object",
        emitReferenceCandidates: true,
        showThumbnails: false,
        showOutline: false,
        onReferenceConfirm: (event) => confirmations.push(event),
        onReferenceCandidateChange: (change) => candidateChanges.push(change),
      })
    },
  })
  const mounted = await mount(Harness)
  const surface = findByTestId(mounted.root, "pdf-surface")
  assert.ok(surface)
  makeScrollable(surface)
  return { ...mounted, fake, surface, confirmations, candidateChanges }
}

test("PdfSurface object mode confirms a native link with descriptor, fingerprint and outline", async (t) => {
  const mounted = await mountSurface(t)
  try {
    await waitFor(() => walk(mounted.root).some((node) => node.props?.["data-annotation-id"] === "link-7"), "annotation hit layer")
    const annotationNode = walk(mounted.root).find((node) => node.props?.["data-annotation-id"] === "link-7")
    const element = annotationElement(annotationNode)
    mounted.surface.ownerDocument = { elementFromPoint: () => element }
    mounted.surface.getBoundingClientRect = () => ({ left: 0, top: 0, width: 900, height: 600 })

    const interaction = walk(mounted.root).find((node) =>
      String(node.props?.class || "").split(/\s+/).includes("pdf-surface__interaction"))
    assert.ok(interaction, "page interaction layer")

    interaction.props.onPointermove({ clientX: 40, clientY: 50, pointerId: 1, pointerType: "mouse" })
    await vue.nextTick()
    const candidates = mounted.candidateChanges.at(-1)?.candidates ?? []
    assert.equal(candidates[0]?.draft.kind, "link")
    assert.ok(candidates[0].draft.fingerprint.contentHash.startsWith("fnv1a32:"))
    const outline = walk(mounted.root).find((node) =>
      String(node.props?.class || "").split(/\s+/).includes("pdf-surface__object-outline"))
    assert.ok(outline, "hover renders a page-space outline for the link")
    assert.equal(outline.props.style.width, "180px")
    assert.equal(outline.props.style.height, "20px")

    interaction.props.onClick({ clientX: 40, clientY: 50, detail: 1, shiftKey: false })
    const confirmation = mounted.confirmations.at(-1)
    assert.ok(confirmation, "link confirmation is emitted")
    assert.equal(confirmation.reference.kind, "link")
    assert.equal(confirmation.reference.document.documentId, "object-mode.pdf")
    assert.equal(confirmation.reference.locator.value.kind, "native-object")
    assert.equal(confirmation.reference.locator.value.objectType, "link")
    assert.equal(confirmation.reference.reliability.semantic.level, "exact")
    assert.ok(confirmation.reference.fingerprint.ancestorKeys.includes("pdf.native-object:link:link-7"))

    mounted.surface.props.onPointerleave()
    await vue.nextTick()
    assert.ok(!walk(mounted.root).some((node) =>
      String(node.props?.class || "").split(/\s+/).includes("pdf-surface__object-outline")))
    assert.deepEqual(mounted.warnings, [])
  } finally {
    mounted.app.unmount()
  }
})

test("PdfSurface shows no outline when the page is the only candidate", async (t) => {
  const mounted = await mountSurface(t)
  try {
    await waitFor(() => walk(mounted.root).some((node) => node.props?.["data-page-index"] === 0), "rendered page")
    const pageNode = walk(mounted.root).find((node) => node.props?.["data-page-index"] === "0"
      || node.props?.["data-page-index"] === 0)
    const pageElement = {
      dataset: { pageIndex: "0" },
      getBoundingClientRect: () => ({ left: 0, top: 0, width: 600, height: 800 }),
    }
    pageElement.closest = (selector) => selector.includes("[data-annotation-id]")
      ? null
      : selector.includes("[data-page-index]") ? pageElement : null
    mounted.surface.ownerDocument = { elementFromPoint: () => pageElement }
    const interaction = walk(mounted.root).find((node) =>
      String(node.props?.class || "").split(/\s+/).includes("pdf-surface__interaction"))
    interaction.props.onPointermove({ clientX: 100, clientY: 100, pointerId: 1, pointerType: "mouse" })
    await vue.nextTick()
    assert.equal(mounted.candidateChanges.at(-1)?.candidates[0]?.draft.kind, "page")
    assert.ok(!walk(mounted.root).some((node) =>
      String(node.props?.class || "").split(/\s+/).includes("pdf-surface__object-outline")))
    assert.ok(pageNode)
  } finally {
    mounted.app.unmount()
  }
})
