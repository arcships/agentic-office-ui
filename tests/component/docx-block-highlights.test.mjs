import assert from "node:assert/strict"
import test from "node:test"

import {
  appendText,
  installGeometryDom,
  setRect,
} from "../fixtures/docx-geometry-dom.mjs"
import {
  importFromDemo,
  mount,
  vue,
  walk,
} from "./vue-test-renderer.mjs"

const docx = await importFromDemo("@arcships/docx-core")
const { DocxBlockHighlights, DocxDocumentSurface, DocxWordHighlights } = await importFromDemo("@arcships/vue-docx")

function paragraph(text) {
  return { type: "paragraph", children: [{ type: "text", text }] }
}

function buildModel() {
  return {
    nodes: [
      paragraph("Blocks fixture one"),
      paragraph("Blocks fixture two"),
      paragraph("third block text"),
      paragraph("fourth block text"),
      paragraph("fifth block text"),
      paragraph("sixth block text"),
    ],
    metadata: {
      sourceParts: 1,
      warnings: [],
      headerSections: [],
      footerSections: [],
      paragraphStyles: [],
    },
  }
}

function addPage(context, pageIndex) {
  const wrapper = context.createElement("div")
  wrapper.setAttribute("data-docx-page-index", String(pageIndex))
  const surface = context.createElement("div")
  surface.setAttribute("data-docx-page-surface", "true")
  surface.computedStyle = { width: "816px", height: "1056px" }
  setRect(surface, {
    left: 100,
    top: 50,
    width: 816,
    height: 1056,
  })
  wrapper.appendChild(surface)
  context.root.appendChild(wrapper)
  return surface
}

function addParagraphHost(surface, blockId, { top = 90, text = "block text" } = {}) {
  const host = surface.ownerDocument.createElement("div")
  host.setAttribute("data-docx-paragraph-host", "true")
  host.setAttribute("data-docx-source-id", blockId)
  setRect(host, { left: 130, top, width: 400, height: 24 })
  surface.appendChild(host)
  const span = surface.ownerDocument.createElement("span")
  span.setAttribute("data-docx-text-start", "0")
  span.setAttribute("data-docx-text-end", String(text.length))
  setRect(span, { left: 130, top, width: 400, height: 24 })
  host.appendChild(span)
  appendText(
    span,
    text,
    [...text].map((_, index) => ({
      left: 130 + index * 10,
      top,
      width: 10,
      height: 20,
    })),
  )
  return host
}

async function mountHighlights(props) {
  const Harness = vue.defineComponent({
    setup() {
      return () => vue.h(DocxBlockHighlights, props)
    },
  })
  return mount(Harness)
}

async function mountWordHighlights(props) {
  const Harness = vue.defineComponent({
    setup() {
      return () => vue.h(DocxWordHighlights, props)
    },
  })
  return mount(Harness)
}

function highlightBoxes(surface) {
  return walk(surface).filter(
    (node) => node.props?.["data-docx-block-highlight"] !== undefined,
  )
}

test("DocxBlockHighlights renders absolutely positioned boxes exactly over blocks 3 and 4", async () => {
  const geometryDom = installGeometryDom()
  try {
    const surface = addPage(geometryDom, 0)
    const document = await docx.extractDocumentBlocks(buildModel())
    const block3 = document.blocks[2]
    const block4 = document.blocks[3]
    addParagraphHost(surface, block3.id, { top: 90, text: block3.content })
    addParagraphHost(surface, block4.id, { top: 130, text: block4.content })

    const geometries = []
    const mounted = await mountHighlights({
      root: geometryDom.root,
      blocks: [block3, block4],
      block: block4,
      pageIndex: 0,
      onGeometry: (value) => geometries.push(value),
    })
    await vue.nextTick()

    const boxes = highlightBoxes(surface)
    assert.equal(boxes.length, 2)
    assert.deepEqual(
      boxes.map((box) => box.props["data-docx-block-highlight"]),
      [block3.id, block4.id],
    )
    assert.deepEqual(
      boxes.map((box) => box.props["data-docx-block-type"]),
      ["text", "text"],
    )
    assert.equal(boxes[0].props.style.left, "30px")
    assert.equal(boxes[0].props.style.top, "40px")
    assert.equal(boxes[0].props.style.width, `${block3.content.length * 10}px`)
    assert.equal(boxes[0].props.style.height, "20px")
    assert.equal(boxes[0].props.style.position, "absolute")
    assert.equal(boxes[0].props.style.pointerEvents, "none")
    assert.equal(boxes[1].props.style.top, "80px")
    assert.equal(boxes[0].props["data-docx-highlight-active"], undefined)
    assert.equal(boxes[1].props["data-docx-highlight-active"], true)
    assert.match(boxes[0].props.style.border, /^1px solid/u)
    assert.match(boxes[1].props.style.border, /^2px solid/u)
    assert.ok(geometries.length >= 1)
    assert.equal(geometries.at(-1).get(block3.id).fragments.length, 1)
    assert.equal(mounted.warnings.length, 0)
    mounted.app.unmount()
  } finally {
    geometryDom.restore()
  }
})

test("DocxBlockHighlights filters by pageIndex and accepts blockIds without a block list", async () => {
  const geometryDom = installGeometryDom()
  try {
    const page0 = addPage(geometryDom, 0)
    const page1 = addPage(geometryDom, 1)
    const document = await docx.extractDocumentBlocks(buildModel())
    const block3 = document.blocks[2]
    const block4 = document.blocks[3]
    addParagraphHost(page0, block3.id, { top: 90, text: block3.content })
    const page1Host = addParagraphHost(page1, block4.id, { top: 100, text: block4.content })
    page1Host.parentElement.parentElement; // keep the page1 tree reachable

    const mounted = await mountHighlights({
      root: geometryDom.root,
      blockIds: [block3.id, block4.id],
      pageIndex: 0,
    })
    await vue.nextTick()

    const page0Boxes = highlightBoxes(page0)
    const page1Boxes = highlightBoxes(page1)
    assert.equal(page0Boxes.length, 1)
    assert.equal(page0Boxes[0].props["data-docx-block-highlight"], block3.id)
    assert.equal(page1Boxes.length, 0)
    mounted.app.unmount()
  } finally {
    geometryDom.restore()
  }
})

test("DocxWordHighlights renders word ranges and marks them as word highlights", async () => {
  const geometryDom = installGeometryDom()
  try {
    const surface = addPage(geometryDom, 0)
    const text = "word highlight fixture"
    const hostId = "docx:node:2"
    const host = addParagraphHost(surface, hostId, { top: 90, text })
    assert.ok(host)
    const start = text.indexOf("highlight")
    const end = start + "highlight".length

    const mounted = await mountWordHighlights({
      root: geometryDom.root,
      wordRanges: [{ blockId: hostId, start, end, content: "highlight" }],
      pageIndex: 0,
    })
    await vue.nextTick()

    const boxes = walk(surface).filter(
      (node) => node.props?.["data-docx-word-highlight"] !== undefined,
    )
    assert.equal(boxes.length, 1)
    assert.equal(boxes[0].props["data-docx-block-highlight"], `${hostId}/word/${start}-${end}`)
    assert.equal(boxes[0].props["data-docx-word-highlight"], `${hostId}/word/${start}-${end}`)
    // "highlight" starts after "word " (5 chars).
    assert.equal(boxes[0].props.style.left, "80px")
    assert.equal(boxes[0].props.style.width, "90px")
    mounted.app.unmount()
  } finally {
    geometryDom.restore()
  }
})

test("DocxBlockHighlights re-measures after repagination without losing existing boxes", async () => {
  const geometryDom = installGeometryDom()
  try {
    const page0 = addPage(geometryDom, 0)
    const document = await docx.extractDocumentBlocks(buildModel())
    const block = document.blocks[2]
    addParagraphHost(page0, block.id, { top: 90, text: block.content })

    const highlights = vue.ref(null)
    const Harness = vue.defineComponent({
      setup() {
        return () => vue.h(DocxBlockHighlights, {
          ref: highlights,
          root: geometryDom.root,
          blocks: [block],
        })
      },
    })
    const mounted = await mount(Harness)
    await vue.nextTick()
    assert.equal(highlightBoxes(page0).length, 1)

    // Repagination mounts the continuation of the same source block on a new
    // page; the next measurement pass adds that page's box and keeps page 0.
    const page1 = addPage(geometryDom, 1)
    addParagraphHost(page1, block.id, { top: 110, text: block.content })
    highlights.value.refresh()
    await vue.nextTick()

    assert.equal(highlightBoxes(page0).length, 1)
    assert.equal(highlightBoxes(page1).length, 1)
    assert.equal(highlightBoxes(page1)[0].props["data-docx-block-highlight"], block.id)
    assert.equal(highlightBoxes(page1)[0].props.style.top, "60px")
    mounted.app.unmount()
  } finally {
    geometryDom.restore()
  }
})

test("DocxDocumentSurface confirms block and word references through referenceConfirm", async () => {
  const base = await importFromDemo("@arcships/docx-core")
  const model = {
    ...base.createBlankDocumentModel(),
    nodes: [
      { type: "paragraph", children: [{ type: "text", text: "Block reference text" }] },
    ],
  }
  let api
  const confirmations = []
  const Harness = vue.defineComponent({
    setup() {
      const surface = vue.ref(null)
      vue.watch(surface, (value) => { api = value }, { flush: "sync" })
      return () => vue.h(DocxDocumentSurface, {
        ref: surface,
        model,
        documentId: "blocks.docx",
        onReferenceConfirm: (event) => confirmations.push(event),
      })
    },
  })
  const mounted = await mount(Harness)
  await vue.nextTick()

  const document = await api.getDocumentBlocks({ includeWords: true })
  assert.equal(document.documentId, "blocks.docx")
  const block = document.blocks[0]
  assert.equal(block.content, "Block reference text")

  // The viewer writes the extracted source id back into the DOM, and text runs
  // expose their paragraph offsets for word geometry.
  const paragraphHost = walk(mounted.root).find(
    (node) => node.props?.["data-docx-paragraph-host"] === true,
  )
  assert.ok(paragraphHost)
  assert.equal(paragraphHost.props["data-docx-source-id"], block.id)
  assert.match(String(paragraphHost.props.innerHTML), /data-docx-text-start="0"/u)
  assert.match(
    String(paragraphHost.props.innerHTML),
    new RegExp(`data-docx-text-end="${block.content.length}"`, "u"),
  )
  const viewerRoot = walk(mounted.root).find(
    (node) => node.props?.["data-docx-document-id"] !== undefined,
  )
  assert.ok(viewerRoot)
  assert.equal(viewerRoot.props["data-docx-document-id"], "blocks.docx")

  const blockDraft = api.confirmBlockReference(block)
  assert.equal(confirmations.at(-1).reference.kind, "paragraph")
  assert.equal(confirmations.at(-1).reference.document.documentId, "blocks.docx")
  assert.equal(confirmations.at(-1).reference.locator.value.kind, "structure")
  assert.equal((await api.resolveReference(confirmations.at(-1).reference)).status, "exact")
  assert.equal(blockDraft.kind, "paragraph")

  const word = block.words[0]
  api.confirmWordReference(block, word)
  assert.equal(confirmations.at(-1).reference.kind, "text-range")
  assert.equal(confirmations.at(-1).reference.locator.value.start.offset, 0)
  assert.equal(confirmations.at(-1).reference.locator.value.end.offset, "Block".length)
  const wordResult = await api.resolveReference(confirmations.at(-1).reference)
  assert.equal(wordResult.status, "exact")
  assert.equal(wordResult.descriptor.content.text, "Block")
  assert.deepEqual(mounted.warnings, [])
  mounted.app.unmount()
})
