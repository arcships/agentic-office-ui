import assert from "node:assert/strict"
import { readFileSync } from "node:fs"
import test from "node:test"

import {
  appendText,
  installGeometryDom,
  setRect,
} from "../fixtures/docx-geometry-dom.mjs"

const docx = await import(
  new URL("../../packages/docx-core/dist/index.js", import.meta.url).href
)
const interaction = await import(
  new URL("../../packages/office-interaction/dist/index.js", import.meta.url).href
)

function paragraph(text, style) {
  return {
    type: "paragraph",
    children: [{ type: "text", text }],
    ...(style ? { style } : {}),
  }
}

function imageParagraph(alt) {
  return {
    type: "paragraph",
    children: [{
      type: "image",
      src: "word/media/image1.png",
      widthPx: 40,
      heightPx: 30,
      alt,
    }],
  }
}

function tableNode(text) {
  return {
    type: "table",
    rows: [{
      type: "table-row",
      cells: [{
        type: "table-cell",
        nodes: [paragraph(text)],
      }],
    }],
  }
}

function buildModel() {
  return {
    nodes: [
      paragraph("Document title", { styleName: "Title", styleId: "Title" }),
      paragraph("Chapter", {
        styleName: "Heading 1",
        styleId: "Heading1",
        headingLevel: 1,
      }),
      paragraph("List item", {
        numbering: { numId: 1, ilvl: 0 },
      }),
      paragraph("alpha beta gamma"),
      imageParagraph("Quarterly chart"),
      paragraph(""),
      tableNode("Cell text"),
      paragraph("Tail text"),
    ],
    metadata: {
      sourceParts: 1,
      warnings: ["preserved-only: diagram"],
      headerSections: [{
        partName: "word/header1.xml",
        nodes: [paragraph("Header title")],
      }],
      footerSections: [{
        partName: "word/footer1.xml",
        nodes: [paragraph("Footer note")],
      }],
      paragraphStyles: [],
      numberingDefinitions: {
        abstracts: [{ abstractNumId: 1, levels: [{ ilvl: 0, format: "decimal" }] }],
        instances: [{ numId: 1, abstractNumId: 1 }],
      },
      footnotes: [
        { id: 1, text: "Footnote", nodes: [paragraph("Footnote body")] },
        { id: 3, text: "Loose note" },
      ],
      endnotes: [{ id: 2, text: "Endnote", nodes: [paragraph("Endnote body")] }],
    },
  }
}

test("DOCUMENT_BLOCK_TYPES exposes the complete Extend parser type set", () => {
  assert.deepEqual([...docx.DOCUMENT_BLOCK_TYPES], [
    "text",
    "heading",
    "section_heading",
    "figure",
    "table",
    "table_head",
    "table_cell",
    "key_value",
    "page_number",
    "barcode",
    "formula",
    "header",
    "footer",
    "watermark",
    "legend",
  ])
})

test("extractDocumentBlocks derives paragraphs, headings, lists, figures, tables and stories", async () => {
  const result = await docx.extractDocumentBlocks(buildModel(), {
    includeWords: true,
    documentId: "blocks.docx",
  })

  assert.equal(result.schemaVersion, 1)
  assert.equal(result.documentId, "blocks.docx")
  assert.equal(result.geometryStatus, "requires_layout")
  assert.deepEqual(result.warnings, ["preserved-only: diagram"])

  const bodyTypes = result.blocks.map((block) => block.type)
  assert.deepEqual(bodyTypes.slice(0, 3), ["section_heading", "heading", "text"])

  const title = result.blocks[0]
  assert.equal(title.content, "Document title")
  assert.equal(title.details.style.styleName, "Title")

  const heading = result.blocks[1]
  assert.equal(heading.details.headingLevel, 1)
  assert.deepEqual(heading.details.modelPath, [{ kind: "paragraph", index: 1 }])
  assert.equal(heading.sourceRefs[0].nodeId, "docx:node:1")

  const list = result.blocks[2]
  assert.deepEqual(list.details.list, { numId: 1, level: 0, format: "decimal" })
  assert.equal(list.type, "text")

  const wordsBlock = result.blocks[3]
  assert.equal(wordsBlock.type, "text")
  assert.equal(wordsBlock.content, "alpha beta gamma")
  assert.deepEqual(wordsBlock.words.map((word) => word.content), ["alpha", "beta", "gamma"])
  assert.deepEqual(wordsBlock.words.map((word) => word.id), [
    `${wordsBlock.id}/word/0-5`,
    `${wordsBlock.id}/word/6-10`,
    `${wordsBlock.id}/word/11-16`,
  ])
  assert.equal(wordsBlock.words[1].sourceRefs[0].nodeId, wordsBlock.id)
  assert.equal(wordsBlock.words[1].sourceRefs[0].start, 6)
  assert.equal(wordsBlock.words[1].sourceRefs[0].end, 10)
  assert.equal(wordsBlock.words[1].sourceRefs[0].offsetEncoding, "utf-16")

  const figure = result.blocks[4]
  assert.equal(figure.type, "figure")
  assert.equal(figure.content, "Quarterly chart")
  assert.equal(figure.id, "docx:node:4/child/0")
  assert.equal(figure.details.childIndex, 0)

  const table = result.blocks.find((block) => block.type === "table")
  assert.ok(table)
  assert.equal(table.content, "Cell text")
  assert.equal(table.details.rowCount, 1)
  assert.equal(table.details.columnCount, 1)
  assert.equal(table.children[0].type, "table_cell")
  assert.equal(table.children[0].sourceRefs[0].nodeId, `${table.id}/r0/c0`)
  assert.equal(table.children[0].children[0].content, "Cell text")

  const header = result.blocks.find((block) => block.type === "header")
  assert.ok(header)
  assert.equal(header.sourceRefs[0].partName, "word/header1.xml")
  assert.equal(header.children[0].sourceRefs[0].story, "header")
  assert.equal(header.children[0].sourceRefs[0].nodeId, "docx:header:word/header1.xml:node:0")

  const footer = result.blocks.find((block) => block.type === "footer")
  assert.ok(footer)
  assert.equal(footer.children[0].content, "Footer note")

  const footnote = result.blocks.find(
    (block) => block.sourceRefs[0]?.story === "footnote",
  )
  assert.ok(footnote)
  assert.equal(footnote.content, "Footnote body")
  const looseFootnote = result.blocks.find((block) => block.content === "Loose note")
  assert.ok(looseFootnote)
  assert.equal(looseFootnote.sourceRefs[0].story, "footnote")
  assert.ok(looseFootnote.words.length > 0)
  const endnote = result.blocks.find(
    (block) => block.sourceRefs[0]?.story === "endnote",
  )
  assert.equal(endnote.content, "Endnote body")

  // Empty paragraph wrappers promote their only meaningful child, and the
  // figure keeps the paragraph's source id so the viewer can find its DOM.
  assert.equal(figure.sourceRefs[0].nodeId, "docx:node:4/child/0")
})

test("extractDocumentBlocks omits words unless requested and skips empty paragraphs/cells", async () => {
  const result = await docx.extractDocumentBlocks(buildModel())
  assert.equal(result.blocks[3].words, undefined)
  const table = result.blocks.find((block) => block.type === "table")
  assert.equal(table.children.length, 1)
  const emptyCellModel = {
    nodes: [{
      type: "table",
      rows: [{
        type: "table-row",
        cells: [
          { type: "table-cell", nodes: [paragraph("")] },
          { type: "table-cell", nodes: [paragraph("kept")] },
        ],
      }],
    }],
    metadata: {
      sourceParts: 1,
      warnings: [],
      headerSections: [],
      footerSections: [],
      paragraphStyles: [],
    },
  }
  const emptyResult = await docx.extractDocumentBlocks(emptyCellModel)
  assert.deepEqual(emptyResult.blocks[0].children.map((cell) => cell.id), [
    `${emptyResult.blocks[0].id}/r0/c1`,
  ])
})

test("segmentDocumentWords keeps punctuation and splits CJK per character", () => {
  assert.deepEqual(
    docx.segmentDocumentWords("Hello, world!").map((word) => word.text),
    ["Hello,", "world!"],
  )
  assert.deepEqual(
    docx.segmentDocumentWords("好世界 abc").map((word) => [word.text, word.start, word.end]),
    [["好", 0, 1], ["世", 1, 2], ["界", 2, 3], ["abc", 4, 7]],
  )
  assert.deepEqual(docx.segmentDocumentWords("   "), [])
})

test("flattenDocumentBlocks and collectDuplicateDocModelBlockIds follow container scope", () => {
  const result = docx.extractDocumentBlocksSync(buildModel())
  const flat = docx.flattenDocumentBlocks(result.blocks)
  assert.ok(flat.length > result.blocks.length)
  assert.equal(flat[0].id, result.blocks[0].id)

  const model = buildModel()
  assert.deepEqual(docx.collectDuplicateDocModelBlockIds(model), [])
  model.nodes[0].blockId = "dup"
  model.nodes[1].blockId = "dup"
  assert.deepEqual(docx.collectDuplicateDocModelBlockIds(model), ["dup"])
})

function makePageDom(context, { scale = 1, pageIndex = 0, paragraphTop = 90 } = {}) {
  const wrapper = context.createElement("div")
  wrapper.setAttribute("data-docx-page-index", String(pageIndex))
  const surface = context.createElement("div")
  surface.setAttribute("data-docx-page-surface", "true")
  surface.computedStyle = { width: "816px", height: "1056px" }
  setRect(surface, {
    left: 100 * scale,
    top: 50 * scale,
    width: 816 * scale,
    height: 1056 * scale,
  })
  wrapper.appendChild(surface)
  context.root.appendChild(wrapper)
  return { wrapper, surface }
}

function addParagraphHost(surface, blockId, { scale = 1, top = 90, texts } = {}) {
  const host = surface.ownerDocument.createElement("div")
  host.setAttribute("data-docx-paragraph-host", "true")
  host.setAttribute("data-docx-source-id", blockId)
  setRect(host, {
    left: 130 * scale,
    top: top * scale,
    width: 400 * scale,
    height: 24 * scale,
  })
  surface.appendChild(host)
  for (const chunk of texts) {
    const span = surface.ownerDocument.createElement("span")
    span.setAttribute("data-docx-text-start", String(chunk.start))
    span.setAttribute("data-docx-text-end", String(chunk.end))
    setRect(span, {
      left: 130 * scale,
      top: top * scale,
      width: 400 * scale,
      height: 24 * scale,
    })
    host.appendChild(span)
    appendText(span, chunk.text, chunk.charRects(scale))
  }
  return host
}

function charRects(text, startLeft, top, charWidth = 10, height = 20) {
  return [...text].map((_, index) => ({
    left: startLeft + index * charWidth,
    top,
    width: charWidth,
    height,
  }))
}

function paragraphGeometryFixture(scale) {
  const geometryDom = installGeometryDom()
  const { root, createElement } = geometryDom
  const { surface } = makePageDom(geometryDom, { scale, paragraphTop: 90 })
  const text = "alpha beta gamma"
  const host = addParagraphHost(surface, "docx:node:3", {
    scale,
    top: 90,
    texts: [{
      text,
      start: 0,
      end: text.length,
      charRects: (factor) => charRects(text, 130 * factor, 90 * factor, 10 * factor, 20 * factor),
    }],
  })
  return { ...geometryDom, root, surface, host, text, before: () => {} }
}

test("block geometry projects DOM text boxes into page CSS pixels at 100% and 200% zoom", () => {
  const results = []
  for (const scale of [1, 2]) {
    const fixture = paragraphGeometryFixture(scale)
    try {
      const block = {
        object: "block",
        id: "docx:node:3",
        type: "text",
        content: fixture.text,
        details: { kind: "paragraph" },
        sourceRefs: [{
          documentId: "blocks.docx",
          nodeId: "docx:node:3",
          partName: "word/document.xml",
          story: "body",
          offsetEncoding: "utf-16",
        }],
        children: [],
      }
      const geometries = docx.resolveDocxBlockGeometries(fixture.root, [block])
      const geometry = geometries.get(block.id)
      assert.equal(geometry.status, "rendered")
      assert.equal(geometry.fragments.length, 1)
      const box = geometry.fragments[0].boundingBox
      // Text starts 30px after the page's left edge and 40px below its top.
      assert.ok(Math.abs(box.left - 30) < 0.01, `left=${box.left}`)
      assert.ok(Math.abs(box.top - 40) < 0.01, `top=${box.top}`)
      assert.ok(Math.abs(box.right - 190) < 0.01, `right=${box.right}`)
      assert.ok(Math.abs(box.bottom - 60) < 0.01, `bottom=${box.bottom}`)
      assert.equal(geometry.fragments[0].pageIndex, 0)
      assert.equal(geometry.fragments[0].pageWidth, 816)
      assert.equal(geometry.fragments[0].pageHeight, 1056)
      results.push(box)
    } finally {
      fixture.restore()
    }
  }
  assert.deepEqual(results[0], results[1])
})

test("word geometry maps UTF-16 offsets to per-word boxes and survives page slices", () => {
  const fixture = paragraphGeometryFixture(1)
  try {
    const text = fixture.text
    const words = docx.segmentDocumentWords(text)
    const beta = words[1]
    const word = {
      object: "word",
      id: `docx:node:3/word/${beta.start}-${beta.end}`,
      blockId: "docx:node:3",
      content: beta.text,
      sourceRefs: [{
        documentId: "",
        nodeId: "docx:node:3",
        partName: "word/document.xml",
        story: "body",
        start: beta.start,
        end: beta.end,
        offsetEncoding: "utf-16",
      }],
    }
    const geometries = docx.resolveDocxWordGeometries(fixture.root, [word])
    const geometry = geometries.get(word.id)
    assert.equal(geometry.status, "rendered")
    const box = geometry.fragments[0].boundingBox
    assert.ok(Math.abs(box.left - 90) < 0.01, `left=${box.left}`)
    assert.ok(Math.abs(box.right - 130) < 0.01, `right=${box.right}`)
    assert.ok(Math.abs(box.top - 40) < 0.01, `top=${box.top}`)

    // A changed rendered character at the same offsets is reported instead of
    // drawing a box over different text.
    fixture.host.children[0].children[0].data = "alpha BETA gamma"
    const changed = docx.resolveDocxWordGeometries(fixture.root, [word])
    assert.equal(changed.get(word.id).status, "unresolved")
    assert.deepEqual(changed.get(word.id).issues, [{
      nodeId: "docx:node:3",
      reason: "range_text_mismatch",
    }])
  } finally {
    fixture.restore()
  }
})

test("block geometry follows the page's rotation (landscape paper dimensions)", () => {
  const geometryDom = installGeometryDom()
  try {
    const { root, createElement } = geometryDom
    const wrapper = createElement("div")
    wrapper.setAttribute("data-docx-page-index", "0")
    const surface = createElement("div")
    surface.setAttribute("data-docx-page-surface", "true")
    surface.computedStyle = { width: "1056px", height: "816px" }
    setRect(surface, { left: 20, top: 10, width: 1056, height: 816 })
    wrapper.appendChild(surface)
    root.appendChild(wrapper)
    const host = createElement("div")
    host.setAttribute("data-docx-paragraph-host", "true")
    host.setAttribute("data-docx-source-id", "docx:node:0")
    setRect(host, { left: 100, top: 80, width: 300, height: 24 })
    surface.appendChild(host)
    const span = createElement("span")
    span.setAttribute("data-docx-text-start", "0")
    span.setAttribute("data-docx-text-end", "9")
    setRect(span, { left: 100, top: 80, width: 300, height: 24 })
    host.appendChild(span)
    appendText(
      span,
      "landscape",
      [..."landscape"].map((_, index) => ({
        left: 100 + index * 10,
        top: 80,
        width: 10,
        height: 20,
      })),
    )
    const block = {
      object: "block",
      id: "docx:node:0",
      type: "text",
      content: "landscape",
      details: { kind: "paragraph" },
      sourceRefs: [{
        documentId: "",
        nodeId: "docx:node:0",
        partName: "word/document.xml",
        story: "body",
        offsetEncoding: "utf-16",
      }],
      children: [],
    }
    const fragment = docx.resolveDocxBlockGeometries(root, [block])
      .get(block.id).fragments[0]
    assert.equal(fragment.pageWidth, 1056)
    assert.equal(fragment.pageHeight, 816)
    assert.ok(Math.abs(fragment.boundingBox.left - 80) < 0.01)
    assert.ok(Math.abs(fragment.boundingBox.top - 70) < 0.01)
    assert.ok(Math.abs(fragment.boundingBox.bottom - 90) < 0.01)
  } finally {
    geometryDom.restore()
  }
})

test("block geometry returns one fragment per mounted page for a split paragraph", () => {
  const geometryDom = installGeometryDom()
  try {
    const { root } = geometryDom
    const page0 = makePageDom(geometryDom, { pageIndex: 0, paragraphTop: 90 })
    addParagraphHost(page0.surface, "docx:node:5", {
      top: 90,
      texts: [{
        text: "split paragraph",
        start: 0,
        end: 15,
        charRects: (factor) => charRects("split paragraph", 130 * factor, 90 * factor),
      }],
    })
    const page1 = makePageDom(geometryDom, { pageIndex: 1, paragraphTop: 120 })
    addParagraphHost(page1.surface, "docx:node:5", {
      top: 120,
      texts: [{
        text: "split paragraph",
        start: 0,
        end: 15,
        charRects: (factor) => charRects("split paragraph", 130 * factor, 120 * factor),
      }],
    })
    const block = {
      object: "block",
      id: "docx:node:5",
      type: "text",
      content: "split paragraph",
      details: { kind: "paragraph" },
      sourceRefs: [{
        documentId: "",
        nodeId: "docx:node:5",
        partName: "word/document.xml",
        story: "body",
        offsetEncoding: "utf-16",
      }],
      children: [],
    }
    const geometry = docx.resolveDocxBlockGeometries(root, [block]).get(block.id)
    assert.equal(geometry.fragments.length, 2)
    assert.deepEqual(geometry.fragments.map((fragment) => fragment.pageIndex), [0, 1])
    assert.ok(geometry.fragments[1].boundingBox.top > geometry.fragments[0].boundingBox.top)

    const normalized = docx.normalizedDocxBoundingBox(geometry.fragments[0])
    assert.deepEqual(Object.keys(normalized).sort(), [
      "height", "left", "pageIndex", "top", "width",
    ])
    assert.equal(normalized.pageIndex, 0)
    assert.ok(normalized.left > 0 && normalized.left < 1)
    assert.ok(normalized.height > 0 && normalized.height < 1)

    const pageBlocks = docx.toDocxPageBlocks(block, geometry)
    assert.deepEqual(pageBlocks.map((page) => page.id), [
      "docx:node:5/page/0",
      "docx:node:5/page/1",
    ])
    assert.deepEqual(pageBlocks.map((page) => page.metadata.page.number), [1, 2])
  } finally {
    geometryDom.restore()
  }
})

test("extraction covers an imported fixture with unique source ids and word ranges", async () => {
  const bytes = readFileSync(
    new URL("../../apps/demo/public/samples/chinese-mixed.docx", import.meta.url),
  )
  const input = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)
  const { model } = await docx.wasmBuildDocModelFromBytes(input)
  const result = await docx.extractDocumentBlocks(model, {
    includeWords: true,
    documentId: "chinese-mixed.docx",
  })
  assert.equal(result.documentId, "chinese-mixed.docx")
  const flat = docx.flattenDocumentBlocks(result.blocks)
  assert.ok(flat.length > 0)

  const ids = flat.map((block) => block.id)
  assert.equal(new Set(ids).size, ids.length)
  assert.equal(new Set(flat.map((block) => block.sourceRefs[0].nodeId)).size, flat.length)
  for (const block of flat) {
    assert.equal(block.sourceRefs[0].nodeId, block.id)
    assert.equal(block.sourceRefs[0].documentId, "chinese-mixed.docx")
  }

  const words = flat.flatMap((block) => block.words ?? [])
  assert.ok(words.length > 0)
  for (const word of words) {
    const block = flat.find((entry) => entry.id === word.blockId)
    assert.ok(block)
    const source = word.sourceRefs[0]
    assert.equal(block.content.slice(source.start, source.end), word.content)
    assert.equal(source.nodeId, block.id)
  }

  // CJK text is split per code point so per-character highlights are possible.
  const cjkWords = words.filter((word) => /^[\p{Script=Han}]$/u.test(word.content))
  assert.ok(cjkWords.length > 0)
})

test("table, row, cell and figure blocks measure their structural element boxes", () => {
  const geometryDom = installGeometryDom()
  try {
    const { root, createElement } = geometryDom
    const { surface } = makePageDom(geometryDom, { paragraphTop: 90 })
    const result = docx.extractDocumentBlocksSync(buildModel())
    const table = result.blocks.find((block) => block.type === "table")
    const cell = table.children[0]
    const figure = result.blocks.find((block) => block.type === "figure")
    assert.ok(table && cell && figure)

    const tableElement = createElement("table")
    tableElement.setAttribute("data-docx-source-id", table.id)
    setRect(tableElement, { left: 130, top: 200, width: 300, height: 60 })
    surface.appendChild(tableElement)
    const rowElement = createElement("tr")
    rowElement.setAttribute("data-docx-source-id", `${table.id}/r0`)
    setRect(rowElement, { left: 130, top: 200, width: 300, height: 60 })
    tableElement.appendChild(rowElement)
    const cellElement = createElement("td")
    cellElement.setAttribute("data-docx-source-id", cell.id)
    setRect(cellElement, { left: 130, top: 200, width: 300, height: 60 })
    rowElement.appendChild(cellElement)

    const figureElement = createElement("img")
    figureElement.setAttribute("data-docx-source-id", figure.id)
    setRect(figureElement, { left: 150, top: 120, width: 40, height: 30 })
    surface.appendChild(figureElement)

    const targets = [
      table,
      {
        object: "block",
        id: `${table.id}/r0`,
        type: "table_head",
        content: "",
        details: { kind: "table-row" },
        sourceRefs: [{
          documentId: "",
          nodeId: `${table.id}/r0`,
          partName: "word/document.xml",
          story: "body",
          offsetEncoding: "utf-16",
        }],
        children: [],
      },
      cell,
      figure,
    ]
    const geometries = docx.resolveDocxBlockGeometries(root, targets)
    for (const target of targets) {
      assert.equal(geometries.get(target.id).status, "rendered", target.id)
      assert.equal(geometries.get(target.id).fragments.length, 1)
    }
    const tableBox = geometries.get(table.id).fragments[0].boundingBox
    assert.ok(Math.abs(tableBox.left - 30) < 0.01)
    assert.ok(Math.abs(tableBox.top - 150) < 0.01)
    assert.ok(Math.abs(tableBox.right - 330) < 0.01)
    assert.ok(Math.abs(tableBox.bottom - 210) < 0.01)
    const figureBox = geometries.get(figure.id).fragments[0].boundingBox
    assert.ok(Math.abs(figureBox.left - 50) < 0.01)
    assert.ok(Math.abs(figureBox.top - 70) < 0.01)
  } finally {
    geometryDom.restore()
  }
})

test("reference drafts carry block and word identity through the unified contract", async () => {
  const model = buildModel()
  const result = await docx.extractDocumentBlocks(model, {
    includeWords: true,
    documentId: "blocks.docx",
  })
  const flat = docx.flattenDocumentBlocks(result.blocks)
  const paragraphBlock = flat.find((block) => block.content === "alpha beta gamma")
  const headingBlock = flat.find((block) => block.content === "Chapter")
  const listBlock = flat.find((block) => block.content === "List item")
  const tableBlock = flat.find((block) => block.type === "table")
  const cellBlock = flat.find((block) => block.type === "table_cell")
  const figureBlock = flat.find((block) => block.type === "figure")
  const headerBlock = flat.find((block) => block.type === "header")
  assert.ok(paragraphBlock && headingBlock && listBlock && tableBlock && cellBlock)
  assert.ok(figureBlock && headerBlock)

  const context = {
    revision: { format: "docx", documentId: "blocks.docx", revision: "v1" },
    model,
    pageCount: 1,
  }
  const parse = (draft) =>
    interaction.parseOfficeObjectReference({ ...draft, referenceId: "ref-block" })

  const paragraphDraft = docx.createDocxBlockReferenceDraft(context, paragraphBlock)
  assert.equal(paragraphDraft.kind, "paragraph")
  assert.equal(paragraphDraft.locator.value.kind, "structure")
  assert.deepEqual(paragraphDraft.locator.value.path, [
    { kind: "paragraph", index: 3 },
  ])
  parse(paragraphDraft)
  const paragraphResult = docx.resolveDocxReference(context, parse(paragraphDraft))
  assert.equal(paragraphResult.status, "exact")
  assert.equal(paragraphResult.descriptor.content.text, "alpha beta gamma")

  const headingDraft = docx.createDocxBlockReferenceDraft(context, headingBlock)
  assert.equal(headingDraft.kind, "heading")
  const listDraft = docx.createDocxBlockReferenceDraft(context, listBlock)
  assert.equal(listDraft.kind, "list-item")
  const tableDraft = docx.createDocxBlockReferenceDraft(context, tableBlock)
  assert.equal(tableDraft.kind, "table")
  assert.equal(tableDraft.locator.value.kind, "structure")
  assert.equal(docx.resolveDocxReference(context, parse(tableDraft)).status, "exact")
  const cellDraft = docx.createDocxBlockReferenceDraft(context, cellBlock)
  assert.equal(cellDraft.kind, "table-cell")
  const cellResult = docx.resolveDocxReference(context, parse(cellDraft))
  assert.equal(cellResult.status, "exact")
  assert.match(cellResult.descriptor.label, /Cell 1, 1/u)
  const figureDraft = docx.createDocxBlockReferenceDraft(context, figureBlock)
  assert.equal(figureDraft.kind, "image")
  assert.equal(figureDraft.locator.value.kind, "image")
  const headerDraft = docx.createDocxBlockReferenceDraft(context, headerBlock)
  assert.equal(headerDraft.kind, "section")
  assert.equal(headerDraft.locator.value.part.kind, "header")
  assert.equal(docx.resolveDocxReference(context, parse(headerDraft)).status, "exact")

  const word = paragraphBlock.words.find((entry) => entry.content === "beta")
  const wordDraft = docx.createDocxWordReferenceDraft(context, paragraphBlock, word)
  assert.equal(wordDraft.kind, "text-range")
  assert.deepEqual(wordDraft.locator.value.start, {
    path: wordDraft.locator.value.start.path,
    offset: 6,
  })
  assert.equal(wordDraft.locator.value.end.offset, 10)
  assert.equal(wordDraft.fingerprint.exactText, "beta")
  const wordResult = docx.resolveDocxReference(context, parse(wordDraft))
  assert.equal(wordResult.status, "exact")
  assert.equal(wordResult.descriptor.content.text, "beta")

  // A new revision relocates by quote instead of dropping the reference.
  const changedRevision = {
    ...context,
    revision: { format: "docx", documentId: "blocks.docx", revision: "v2" },
  }
  const relocated = docx.resolveDocxReference(changedRevision, parse(wordDraft))
  assert.equal(relocated.status, "relocated")
  assert.deepEqual(relocated.reasonCodes, ["docx.revision-changed", "docx.locator-stable"])

  // Word drafts reject offsets outside the block text.
  assert.throws(
    () => docx.createDocxWordReferenceDraft(context, paragraphBlock, {
      blockId: paragraphBlock.id,
      start: 40,
      end: 50,
    }),
    /exceed the block text/u,
  )
})
