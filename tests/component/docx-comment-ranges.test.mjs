// Regression tests for the remaining bullets of issue #29:
//
//   1. `commentRangeStart` / `commentRangeEnd` are model data: character offsets
//      stored on the paragraph at parse time (not re-parsed from `sourceXml`).
//   2. Highlights and gutter cards stay alive after edits.
//   3. The narrow-width gutter becomes an overlay drawer instead of leaving the
//      paper or flowing into the page.
//
// Each block names the bullet it covers so the mapping stays checkable.

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import { importFromDemo, mount, walk } from "./vue-test-renderer.mjs";

const {
  DocxPageSurface,
  DocxViewerRoot,
  DOCX_ANNOTATION_GUTTER_DRAWER_WIDTH_PX,
  DOCX_ANNOTATION_GUTTER_WIDTH_PX,
  annotationGutterDrawerPaneStyle,
  annotationGutterUsesDrawer,
  renderParagraphRuns,
  renderStaticHtml,
  useDocxEditor,
} = await importFromDemo("@arcships/vue-docx")
const {
  collectCommentsFromModel,
  createBlankDocumentModel,
  mergeParagraphCommentRanges,
  normalizeDocModel,
  parseParagraphCommentRanges,
  remapParagraphCommentRanges,
  replaceText,
  resolveParagraphCommentMarkup,
  resolveParagraphCommentRanges,
  serializeDocx,
  shiftParagraphCommentRanges,
  splitParagraphCommentRanges,
  updateParagraphText,
} = await importFromDemo("@arcships/docx-core")
const { strFromU8, unzipSync } = await import(
  new URL("../..", import.meta.url).href + "packages/docx-core/node_modules/fflate/esm/browser.js"
)

const vue = await importFromDemo("vue")

// ── Fixtures ───────────────────────────────────────────────────────
const COMMENT_RANGE_XML =
  '<w:p><w:r><w:t>Alpha </w:t></w:r>' +
  '<w:commentRangeStart w:id="1"/>' +
  '<w:r><w:t>beta gamma</w:t></w:r>' +
  '<w:commentRangeEnd w:id="1"/>' +
  '<w:r><w:t> delta</w:t></w:r></w:p>'

const COMMENTED_TEXT = "Alpha beta gamma delta"

function paragraphWithRanges(overrides = {}) {
  return {
    type: "paragraph",
    children: [{ type: "text", text: COMMENTED_TEXT }],
    sourceXml: COMMENT_RANGE_XML,
    ...overrides,
  }
}

function modelWithComment({ keepSourceXml = true } = {}) {
  const model = createBlankDocumentModel()
  model.metadata.comments = [
    { id: 1, author: "Alex Reviewer", date: "2026-01-02T00:00:00Z", text: "Please confirm." },
  ]
  const paragraph = paragraphWithRanges(
    keepSourceXml ? {} : { sourceXml: undefined }
  )
  if (!keepSourceXml) {
    paragraph.commentRanges = [
      { commentId: 1, start: 6, end: 16, anchorText: "beta gamma" },
    ]
  }
  model.nodes = [
    paragraph,
    { type: "paragraph", children: [{ type: "text", text: "Second paragraph" }] },
  ]
  // The importer normalizes parsed models; do the same so the fixture matches
  // the real pipeline (parse -> normalize -> edit).
  return normalizeDocModel(model)
}

function nodesWithClass(root, className) {
  return walk(root).filter((node) =>
    String(node?.props?.class || "").split(/\s+/).includes(className)
  )
}

function dataNodes(root, attribute) {
  return walk(root).filter((node) => node.props?.[attribute] !== undefined)
}

// Handlers compiled from `@click.stop` expect a native event; the test renderer
// has no DOM, so supply the two methods the modifier wrapper calls.
function clickElement(node) {
  node.props.onClick({ stopPropagation() {}, preventDefault() {} })
}

// ── Bullet 1: model offsets at parse time ──────────────────────────
test("parsing stores commentRangeStart/End as character offsets, not run granularity", () => {
  const ranges = parseParagraphCommentRanges(COMMENT_RANGE_XML, COMMENTED_TEXT.length)

  assert.deepEqual(ranges, [
    { commentId: 1, start: 6, end: 16, anchorText: "beta gamma" },
  ])
  // The commented text starts inside the first run's text: a run-granular
  // mapping would have started the range at offset 0 or ended it at 22.
  assert.equal(COMMENTED_TEXT.slice(ranges[0].start, ranges[0].end), "beta gamma")
})

test("normalizeDocModel captures the offsets on the paragraph model", () => {
  const model = modelWithComment()
  const normalized = normalizeDocModel(model)

  assert.deepEqual(normalized.nodes[0].commentRanges, [
    { commentId: 1, start: 6, end: 16, anchorText: "beta gamma" },
  ])
  // Normalizing again keeps the stored offsets (they are not re-derived).
  const again = normalizeDocModel(normalized)
  assert.deepEqual(again.nodes[0].commentRanges, normalized.nodes[0].commentRanges)
})

test("resolveParagraphCommentMarkup maps model offsets to exact character windows", () => {
  const normalized = normalizeDocModel(modelWithComment())
  const markup = resolveParagraphCommentMarkup(normalized.nodes[0])

  assert.ok(markup)
  assert.deepEqual(markup.commentIdsByVisibleChildIndex, [[1]])
  assert.deepEqual(markup.rangesByVisibleChildIndex, [
    [{ commentId: 1, start: 6, end: 16 }],
  ])
})

test("the run renderer highlights only the commented characters of a run", () => {
  const normalized = normalizeDocModel(modelWithComment())
  const rendered = renderParagraphRuns(
    normalized.nodes[0],
    "p-0",
    "light",
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    { showCommentHighlights: true }
  )
  const html = (Array.isArray(rendered) ? rendered : [rendered])
    .map((vnode) => renderStaticHtml(vnode))
    .join("")

  // One span per character run with the same comment set: "Alpha ", the
  // highlighted "beta gamma", then " delta".
  const spans = [...html.matchAll(/<span[^>]*>([^<]*)<\/span>/g)].map((match) => ({
    text: match[1],
    highlighted: match[0].includes("background-color"),
  }))

  assert.deepEqual(
    spans.map((span) => span.text).join(""),
    COMMENTED_TEXT
  )
  assert.deepEqual(
    spans.filter((span) => span.highlighted).map((span) => span.text),
    ["beta gamma"],
    "only the commented characters are highlighted"
  )
  assert.match(html, /data-docx-comment-ids="1"/)
})

// ── Bullet 2: edits keep the ranges, highlights and cards ──────────
test("updateParagraphText shifts stored offsets and keeps the highlight exact", () => {
  const normalized = normalizeDocModel(modelWithComment())
  const edited = updateParagraphText(normalized, 0, `XX ${COMMENTED_TEXT}`)

  assert.deepEqual(edited.nodes[0].commentRanges, [
    { commentId: 1, start: 9, end: 19, anchorText: "beta gamma" },
  ])
  assert.equal(edited.nodes[0].sourceXml, undefined, "the edit clears sourceXml")

  const markup = resolveParagraphCommentMarkup(edited.nodes[0])
  assert.deepEqual(markup?.rangesByVisibleChildIndex, [
    [{ commentId: 1, start: 9, end: 19 }],
  ])
  // The card survives the edit too: comments no longer come from sourceXml.
  assert.deepEqual(
    collectCommentsFromModel(edited).map((comment) => comment.anchorText),
    ["beta gamma"]
  )
})

test("typing inside a comment trims the range instead of dropping it", () => {
  const normalized = normalizeDocModel(modelWithComment())
  const edited = updateParagraphText(normalized, 0, "Alpha beZZta gamma delta")

  assert.deepEqual(edited.nodes[0].commentRanges, [
    { commentId: 1, start: 6, end: 18, anchorText: "beZZta gamma" },
  ])
  assert.equal(collectCommentsFromModel(edited).length, 1)
})

test("replaceText remaps stored offsets for a find/replace edit", () => {
  const normalized = normalizeDocModel(modelWithComment())
  const replaced = replaceText(normalized, "Alpha", "Alpha one two")

  // "Alpha" -> "Alpha one two" inserts 8 characters before the range.
  assert.deepEqual(replaced.nodes[0].commentRanges, [
    { commentId: 1, start: 14, end: 24, anchorText: "beta gamma" },
  ])
})

test("a range whose offsets are lost is re-anchored by its anchorText", () => {
  const rewritten = "Rewritten text that still mentions beta gamma somewhere"
  const remapped = remapParagraphCommentRanges(
    COMMENTED_TEXT,
    rewritten,
    [{ commentId: 1, start: 6, end: 16, anchorText: "beta gamma" }]
  )
  // The edit replaced the whole paragraph, so the offsets are meaningless...
  assert.notEqual(rewritten.slice(remapped[0].start, remapped[0].end), "beta gamma")
  // ...but the anchor text survives for the read path to re-anchor on.
  assert.equal(remapped[0].anchorText, "beta gamma")

  assert.deepEqual(resolveParagraphCommentRanges(rewritten, remapped), [
    { commentId: 1, start: 35, end: 45, anchorText: "beta gamma" },
  ])
})

test("split / shift / merge helpers keep ranges correct across structural edits", () => {
  const ranges = [{ commentId: 1, start: 6, end: 16, anchorText: "beta gamma" }]

  const split = splitParagraphCommentRanges(ranges, 10)
  assert.deepEqual(split.before, [
    { commentId: 1, start: 6, end: 10, anchorText: "beta" },
  ])
  assert.deepEqual(split.after, [
    { commentId: 1, start: 0, end: 6, anchorText: " gamma" },
  ])

  assert.deepEqual(shiftParagraphCommentRanges(ranges, 5), [
    { commentId: 1, start: 11, end: 21, anchorText: "beta gamma" },
  ])

  const merged = mergeParagraphCommentRanges(
    [
      { commentId: 1, start: 6, end: 10, anchorText: "beta" },
      { commentId: 1, start: 10, end: 16, anchorText: " gamma" },
    ],
    "Alpha beta gamma delta"
  )
  assert.deepEqual(merged, [
    { commentId: 1, start: 6, end: 16, anchorText: "beta gamma" },
  ])
})

test("the editor keeps a comment on both halves when a paragraph is split", async () => {
  let controller
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: modelWithComment() })
      return () => vue.h("div")
    },
  })
  const mounted = await mount(Harness)
  try {
    const result = controller.splitParagraphAtSelection(COMMENTED_TEXT, 10)
    assert.ok(result, "the split is applied")

    const first = controller.model.nodes[0]
    const second = controller.model.nodes[1]
    assert.equal(first.children.map((child) => child.text).join(""), "Alpha beta")
    assert.equal(
      second.children.map((child) => child.text).join(""),
      " gamma delta"
    )
    assert.deepEqual(first.commentRanges, [
      { commentId: 1, start: 6, end: 10, anchorText: "beta" },
    ])
    assert.deepEqual(second.commentRanges, [
      { commentId: 1, start: 0, end: 6, anchorText: " gamma" },
    ])
  } finally {
    mounted.app.unmount()
  }
})

test("joining two paragraphs carries both comment ranges onto the merged text", async () => {
  const model = modelWithComment()
  // Second paragraph also carries a comment so the join has ranges on both sides.
  model.metadata.comments.push({
    id: 2,
    author: "Morgan Editor",
    date: "2026-01-03T00:00:00Z",
    text: "Second note.",
  })
  model.nodes[1] = normalizeDocModel({
    ...model,
    nodes: [
      { type: "paragraph", children: [{ type: "text", text: "unused" }] },
      {
        type: "paragraph",
        children: [{ type: "text", text: "Keep this sentence" }],
        commentRanges: [
          { commentId: 2, start: 5, end: 9, anchorText: "this" },
        ],
      },
    ],
  }).nodes[1]

  let controller
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model })
      return () => vue.h("div")
    },
  })
  const mounted = await mount(Harness)
  try {
    // Delete from the end of the first paragraph through the start of the
    // second: the paragraphs join, and both comments must stay anchored.
    controller.deleteExpandedSelection({
      start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 20 },
      end: { location: { kind: "paragraph", nodeIndex: 1 }, offset: 5 },
    })
    await vue.nextTick()

    const merged = controller.model.nodes[0]
    const mergedText = merged.children
      .map((child) => (child.type === "text" ? child.text : ""))
      .join("")
    assert.equal(mergedText, "Alpha beta gamma delthis sentence")
    assert.deepEqual(
      (merged.commentRanges ?? []).map((range) => ({
        commentId: range.commentId,
        text: mergedText.slice(range.start, range.end),
      })),
      [
        { commentId: 1, text: "beta gamma" },
        { commentId: 2, text: "this" },
      ]
    )
    assert.equal(
      collectCommentsFromModel(controller.model).length,
      2,
      "both cards survive the join"
    )
  } finally {
    mounted.app.unmount()
  }
})

test("the gutter still shows the card when the paragraph was edited (no sourceXml)", async () => {
  const model = modelWithComment({ keepSourceXml: false })
  assert.equal(collectCommentsFromModel(model).length, 1, "model offsets anchor the card")

  const mounted = await mount(DocxViewerRoot, {
    model,
    zoomScale: 100,
    showComments: true,
  })
  try {
    const cards = nodesWithClass(mounted.root, "docx-gutter-card")
    assert.equal(cards.length, 1)
    assert.equal(cards[0].props["data-docx-gutter-card-id"], "comment:p:0:comment:1")
    assert.deepEqual(mounted.warnings, [])
  } finally {
    mounted.app.unmount()
  }
})

// ── Bullet 1 (serializer): ranges go back into document.xml ────────
test("serialization re-emits comment range markers at the stored offsets", async () => {
  const normalized = normalizeDocModel(modelWithComment())
  const edited = updateParagraphText(normalized, 0, `XX ${COMMENTED_TEXT}`)
  assert.equal(edited.nodes[0].sourceXml, undefined)

  const bytes = await serializeDocx(edited)
  const archive = unzipSync(new Uint8Array(bytes))
  const documentXml = strFromU8(archive["word/document.xml"])

  assert.match(
    documentXml,
    /XX Alpha <\/w:t><\/w:r><w:commentRangeStart w:id="1"\/><w:r><w:rPr>|<w:t[^>]*>XX Alpha <\/w:t><\/w:r><w:commentRangeStart w:id="1"\/>/,
    "the range starts right after the two inserted characters"
  )
  assert.match(
    documentXml,
    /beta gamma<\/w:t><\/w:r><w:commentRangeEnd w:id="1"\/>/,
    "the range ends right after the commented text"
  )
  assert.match(documentXml, /<w:commentReference w:id="1"\/>/)

  // The markers must not be inside a run: `w:commentRangeStart` is
  // paragraph-level content in OOXML.
  assert.doesNotMatch(documentXml, /<w:r>[^<]*<w:commentRangeStart/)
})

// ── Bullet 3: narrow-width drawer ─────────────────────────────────
test("annotationGutterUsesDrawer switches to the drawer below the threshold", () => {
  assert.equal(annotationGutterUsesDrawer(1440), false)
  assert.equal(annotationGutterUsesDrawer(1024), false)
  assert.equal(annotationGutterUsesDrawer(1023), true)
  assert.equal(annotationGutterUsesDrawer(760), true)
  assert.equal(annotationGutterUsesDrawer(undefined), false)

  const open = annotationGutterDrawerPaneStyle(2, 1056, true)
  assert.equal(open.width, `${DOCX_ANNOTATION_GUTTER_DRAWER_WIDTH_PX}px`)
  assert.equal(open.transform, "scale(0.5)")
  assert.equal(open.right, "0px")
  assert.equal(open.left, "auto")
  assert.equal(open.height, `${1056 * 2}px`)

  const closed = annotationGutterDrawerPaneStyle(1, 1056, false)
  assert.equal(closed.width, "28px")
})

async function mountPageSurface(overrides = {}) {
  return mount(DocxPageSurface, {
    pageIndex: 0,
    pageLayout: {
      pageWidthPx: 816,
      pageHeightPx: 1056,
      marginsPx: { top: 72, right: 72, bottom: 72, left: 72 },
    },
    pageNodeSegments: [{ nodeIndex: 0 }],
    model: modelWithComment({ keepSourceXml: false }),
    pageWidthPx: 816,
    pageContentWidthPx: 672,
    pageNumber: 1,
    totalPages: 1,
    commentsEnabled: true,
    comments: collectCommentsFromModel(modelWithComment({ keepSourceXml: false })),
    ...overrides,
  })
}

test("DocxPageSurface renders the gutter as a toggleable drawer in drawer mode", async () => {
  const mounted = await mountPageSurface({ annotationDrawerMode: true })
  try {
    const gutter = dataNodes(mounted.root, "data-docx-annotation-gutter")[0]
    assert.ok(gutter, "the drawer pane renders")
    assert.equal(gutter.props["data-docx-annotation-gutter-mode"], "drawer")
    assert.equal(gutter.props["data-docx-annotation-drawer-open"], "true")
    assert.ok(
      String(gutter.props.class).split(/\s+/).includes("docx-tracked-change-gutter--drawer")
    )
    assert.equal(gutter.props.style.width, `${DOCX_ANNOTATION_GUTTER_DRAWER_WIDTH_PX}px`)

    const tab = dataNodes(mounted.root, "data-docx-gutter-drawer-tab")[0]
    assert.ok(tab, "the drawer exposes a tab")
    assert.equal(nodesWithClass(mounted.root, "docx-gutter-card").length, 1)

    clickElement(tab)
    await vue.nextTick()
    const closedGutter = dataNodes(mounted.root, "data-docx-annotation-gutter")[0]
    assert.equal(closedGutter.props["data-docx-annotation-drawer-open"], "false")
    assert.equal(nodesWithClass(mounted.root, "docx-gutter-card").length, 0)
    assert.equal(closedGutter.props.style.width, "28px")

    clickElement(dataNodes(mounted.root, "data-docx-gutter-drawer-tab")[0])
    await vue.nextTick()
    assert.equal(nodesWithClass(mounted.root, "docx-gutter-card").length, 1)
    assert.deepEqual(mounted.warnings, [])
  } finally {
    mounted.app.unmount()
  }
})

test("DocxViewerRoot switches to the drawer and reserves no gutter column on narrow viewports", async (t) => {
  const originalClientWidth = Object.getOwnPropertyDescriptor(
    Object.prototype,
    "clientWidth"
  )
  Object.defineProperty(Object.prototype, "clientWidth", {
    configurable: true,
    get: () => 760,
  })
  t.after(() => {
    if (originalClientWidth) {
      Object.defineProperty(Object.prototype, "clientWidth", originalClientWidth)
    } else {
      delete Object.prototype.clientWidth
    }
  })

  const mounted = await mount(DocxViewerRoot, {
    model: modelWithComment({ keepSourceXml: false }),
    zoomScale: 100,
    showComments: true,
  })
  try {
    await vue.nextTick()

    const gutter = dataNodes(mounted.root, "data-docx-annotation-gutter")[0]
    assert.ok(gutter, "the annotation drawer renders")
    assert.equal(
      gutter.props["data-docx-annotation-gutter-mode"],
      "drawer",
      "760px is below the drawer threshold"
    )

    const wrapper = dataNodes(mounted.root, "data-docx-page-wrapper")[0]
    const pageSurface = dataNodes(mounted.root, "data-docx-page-surface")[0]
    const pageWidth = Number.parseFloat(pageSurface.props.style.width)
    assert.equal(
      wrapper.props.style.width,
      `${pageWidth}px`,
      "drawer mode reserves no column beside the paper"
    )
    assert.notEqual(
      gutter.props.style.width,
      `${DOCX_ANNOTATION_GUTTER_WIDTH_PX}px`,
      "the drawer is not the fixed column pane"
    )
    assert.deepEqual(mounted.warnings, [])
  } finally {
    mounted.app.unmount()
  }
})

// Keep the fixture honest about the source document used in the docs.
void readFileSync
