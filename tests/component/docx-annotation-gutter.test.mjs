// Regression tests for issue #29: DOCX comment / tracked-change cards must be
// anchored to their text instead of being an unpositioned list.
//
// Three layers are covered:
//   1. docx-core: page annotation ordering, page-space anchor resolution and
//      the collision-free card layout (pure, no browser).
//   2. DocxPageSurface: cards are positioned from their anchors, numbered in
//      document order, and the pane counter-scales instead of scaling with the
//      page.
//   3. DocxViewerRoot: the zoom factor reaches the pane and the reserved
//      gutter footprint stays constant in screen px.

import assert from "node:assert/strict";
import test from "node:test";

import {
  importFromDemo,
  mount,
  walk,
} from "./vue-test-renderer.mjs";

const {
  DocxPageSurface,
  DocxViewerRoot,
  DOCX_ANNOTATION_GUTTER_GAP_PX,
  DOCX_ANNOTATION_GUTTER_WIDTH_PX,
  annotationGutterCardTopPx,
  annotationGutterPaneStyle,
  annotationGutterReservedWidthPx,
} = await importFromDemo("@arcships/vue-docx")
const {
  buildPageGutterAnnotations,
  collectCommentsFromModel,
  createBlankDocumentModel,
  findGutterAnnotationAnchorElementInPage,
  layoutTrackedChangesForPage,
  resolveGutterAnnotationAnchorsInPage,
} = await importFromDemo("@arcships/docx-core")

const PAGE_WIDTH = 816
const PAGE_HEIGHT = 1056
const MARGINS = { top: 72, bottom: 72, left: 72, right: 72 }
const PAGE_LAYOUT = {
  pageWidthPx: PAGE_WIDTH,
  pageHeightPx: PAGE_HEIGHT,
  marginsPx: MARGINS,
}

// ── Fixtures ───────────────────────────────────────────────────────
function paragraphLocation(nodeIndex) {
  return { kind: "paragraph", nodeIndex }
}

function commentFixture({ id, commentId, nodeIndex, text, author, resolved }) {
  return {
    id,
    commentId,
    author: author ?? "Alex Reviewer",
    date: "2026-01-02T00:00:00Z",
    text,
    anchorText: "Hello",
    resolved,
    nodeIndex,
    location: paragraphLocation(nodeIndex),
  }
}

function trackedChangeFixture({ id, kind, nodeIndex, text, author }) {
  return {
    id,
    kind,
    author: author ?? "Morgan Editor",
    date: "2026-01-03T00:00:00Z",
    text,
    nodeIndex,
    location: paragraphLocation(nodeIndex),
  }
}

function pageSegments(...nodeIndexes) {
  return nodeIndexes.map((nodeIndex) => ({ nodeIndex }))
}

function hasClass(node, className) {
  return String(node?.props?.class || "").split(/\s+/).includes(className)
}

function nodesWithClass(root, className) {
  return walk(root).filter((node) => hasClass(node, className))
}

function dataNodes(root, attribute) {
  return walk(root).filter((node) => node.props?.[attribute] !== undefined)
}

function modelWithParagraphs(texts) {
  const model = createBlankDocumentModel()
  return {
    ...model,
    nodes: texts.map((text) => ({
      type: "paragraph",
      children: [{ type: "text", text }],
    })),
  }
}

// ── Fake DOM for the anchor resolver ───────────────────────────────
function rect(left, top, width, height) {
  return { left, top, width, height, right: left + width, bottom: top + height }
}

function fakeAnchor(anchorRect) {
  return { getBoundingClientRect: () => anchorRect }
}

function fakeRun({ attributes, anchorRect }) {
  return {
    ...fakeAnchor(anchorRect),
    hasAttribute(name) {
      return name in attributes
    },
    getAttribute(name) {
      return attributes[name] ?? null
    },
    querySelectorAll() {
      return []
    },
  }
}

function fakePageElement({ clientRect, pageOffsetSize, matches }) {
  return {
    offsetHeight: pageOffsetSize.height,
    offsetWidth: pageOffsetSize.width,
    getBoundingClientRect: () => clientRect,
    hasAttribute: () => false,
    getAttribute: () => null,
    querySelector(selector) {
      for (const entry of matches) {
        if (entry.matches(selector)) return entry.element
      }
      return null
    },
    querySelectorAll() {
      return []
    },
  }
}

function paragraphSelector(nodeIndex) {
  return `[data-docx-paragraph-host="true"][data-docx-paragraph-node-index="${nodeIndex}"]`
}

// ── 1. docx-core: ordering ─────────────────────────────────────────
test("buildPageGutterAnnotations numbers annotations in document order, not by kind", () => {
  const annotations = buildPageGutterAnnotations({
    pageSegments: pageSegments(0, 2),
    // Deliberately unsorted input: kinds interleaved, out-of-order nodes.
    trackedChanges: [
      trackedChangeFixture({ id: "change-late", kind: "insertion", nodeIndex: 2, text: "late" }),
      trackedChangeFixture({ id: "change-early", kind: "deletion", nodeIndex: 0, text: "early" }),
    ],
    comments: [
      commentFixture({ id: "comment-late", commentId: 7, nodeIndex: 2, text: "late comment" }),
      commentFixture({ id: "comment-early", commentId: 3, nodeIndex: 0, text: "early comment" }),
    ],
  })

  assert.deepEqual(
    annotations.map((annotation) => annotation.id),
    ["comment:comment-early", "change:change-early", "comment:comment-late", "change:change-late"]
  )
  assert.deepEqual(
    annotations.map((annotation) => annotation.ordinal),
    [1, 2, 3, 4]
  )
})

test("buildPageGutterAnnotations keeps only the page's nodes and honours visibility toggles", () => {
  const input = {
    pageSegments: pageSegments(1),
    comments: [
      commentFixture({ id: "comment-on-page", commentId: 1, nodeIndex: 1, text: "on page" }),
      commentFixture({ id: "comment-elsewhere", commentId: 2, nodeIndex: 5, text: "elsewhere" }),
    ],
    trackedChanges: [
      trackedChangeFixture({ id: "change-on-page", kind: "insertion", nodeIndex: 1, text: "on page" }),
    ],
  }

  const both = buildPageGutterAnnotations(input)
  assert.deepEqual(
    both.map((annotation) => annotation.id),
    ["comment:comment-on-page", "change:change-on-page"]
  )

  assert.deepEqual(
    buildPageGutterAnnotations({ ...input, showComments: false }).map((annotation) => annotation.id),
    ["change:change-on-page"]
  )
  assert.deepEqual(
    buildPageGutterAnnotations({ ...input, showTrackedChanges: false }).map((annotation) => annotation.id),
    ["comment:comment-on-page"]
  )
  assert.deepEqual(buildPageGutterAnnotations({ ...input, pageSegments: pageSegments(4) }), [])
})

test("buildPageGutterAnnotations is deterministic for repeated calls", () => {
  const input = {
    pageSegments: pageSegments(0),
    comments: [
      commentFixture({ id: "comment-b", commentId: 9, nodeIndex: 0, text: "b" }),
      commentFixture({ id: "comment-a", commentId: 2, nodeIndex: 0, text: "a" }),
    ],
  }

  const first = buildPageGutterAnnotations(input).map((annotation) => annotation.id)
  const second = buildPageGutterAnnotations(input).map((annotation) => annotation.id)
  assert.deepEqual(first, second)
  assert.deepEqual(first, ["comment:comment-a", "comment:comment-b"])
})

test("buildPageGutterAnnotations keeps header/footer comments on matching parts and hides resolved threads", () => {
  const header = {
    ...commentFixture({ id: "comment-header", commentId: 21, nodeIndex: 0, text: "header note" }),
    region: "header",
    partName: "word/header1.xml",
  }
  const footer = {
    ...commentFixture({ id: "comment-footer", commentId: 22, nodeIndex: 0, text: "footer note" }),
    region: "footer",
    partName: "word/footer1.xml",
  }
  const resolvedRoot = commentFixture({ id: "comment-resolved", commentId: 23, nodeIndex: 0, text: "done", resolved: true })
  const resolvedReply = {
    ...commentFixture({ id: "comment-reply", commentId: 24, nodeIndex: 0, text: "ack" }),
    parentId: 23,
  }

  const annotations = buildPageGutterAnnotations({
    pageSegments: pageSegments(0),
    comments: [header, footer, resolvedRoot, resolvedReply],
    headerPartName: "word/header1.xml",
    footerPartName: "word/footer1.xml",
  })
  assert.deepEqual(
    annotations.map((annotation) => annotation.id),
    ["comment:comment-header", "comment:comment-footer"],
  )

  const withoutParts = buildPageGutterAnnotations({
    pageSegments: pageSegments(0),
    comments: [header, footer],
  })
  assert.deepEqual(withoutParts, [], "header/footer comments need their part on the page")

  const withResolved = buildPageGutterAnnotations({
    pageSegments: pageSegments(0),
    comments: [resolvedRoot, resolvedReply],
    showResolvedComments: true,
  })
  assert.deepEqual(
    withResolved.map((annotation) => annotation.id),
    ["comment:comment-resolved", "comment:comment-reply"],
  )
})

test("findGutterAnnotationAnchorElementInPage resolves header comment anchors", () => {
  const paragraphElement = {
    ...fakeAnchor(rect(72, 40, 200, 20)),
    hasAttribute: () => false,
    getAttribute: () => null,
    querySelectorAll: () => [],
  }
  const headerElement = {
    hasAttribute: () => false,
    getAttribute: () => null,
    querySelectorAll: () => [],
    querySelector: (selector) =>
      selector.includes("data-docx-header-node-index") ? paragraphElement : null,
  }
  const page = fakePageElement({
    clientRect: rect(0, 0, PAGE_WIDTH, PAGE_HEIGHT),
    pageOffsetSize: { width: PAGE_WIDTH, height: PAGE_HEIGHT },
    matches: [
      {
        matches: (selector) => selector === '[data-docx-page-header="true"]',
        element: headerElement,
      },
    ],
  })
  const header = {
    ...commentFixture({ id: "comment-header", commentId: 31, nodeIndex: 0, text: "header note" }),
    region: "header",
    partName: "word/header1.xml",
  }
  const annotations = buildPageGutterAnnotations({
    pageSegments: pageSegments(),
    comments: [header],
    headerPartName: "word/header1.xml",
  })

  assert.equal(
    findGutterAnnotationAnchorElementInPage(page, annotations[0]),
    paragraphElement,
  )
})

// ── 2. docx-core: page-space anchors ───────────────────────────────
test("resolveGutterAnnotationAnchorsInPage measures anchors in page space, not screen space", () => {
  const annotations = buildPageGutterAnnotations({
    pageSegments: pageSegments(0, 1),
    comments: [
      commentFixture({ id: "comment-with-run", commentId: 5, nodeIndex: 0, text: "has run" }),
      commentFixture({ id: "comment-without-run", commentId: 6, nodeIndex: 1, text: "lost run" }),
    ],
  })

  // The page is rendered at 50% zoom: client rects are half of page space.
  const pageClientRect = rect(100, 50, PAGE_WIDTH / 2, PAGE_HEIGHT / 2)
  const page = fakePageElement({
    clientRect: pageClientRect,
    pageOffsetSize: { width: PAGE_WIDTH, height: PAGE_HEIGHT },
    matches: [
      {
        matches: (selector) => selector === paragraphSelector(0),
        element: {
          ...fakeAnchor(rect(100, 50, PAGE_WIDTH / 2, 40)),
          hasAttribute: () => false,
          getAttribute: () => null,
          querySelectorAll: () => [
            fakeRun({
              attributes: { "data-docx-comment-ids": "5" },
              anchorRect: rect(160, 90, 40, 10),
            }),
          ],
        },
      },
      {
        matches: (selector) => selector === paragraphSelector(1),
        element: {
          ...fakeAnchor(rect(100, 150, PAGE_WIDTH / 2, 20)),
          hasAttribute: () => false,
          getAttribute: () => null,
          querySelectorAll: () => [],
        },
      },
    ],
  })

  const anchors = resolveGutterAnnotationAnchorsInPage(page, annotations)

  // Run at client (160, 90) relative to the page's client origin (100, 50)
  // → (60, 40) screen px → (120, 80) page px at 50% zoom.
  assert.deepEqual(anchors.get("comment:comment-with-run"), { x: 120, y: 80 })
  // No inline markup left (the paragraph was edited): the paragraph itself is
  // the anchor, so the card still has a position.
  assert.deepEqual(anchors.get("comment:comment-without-run"), { x: 0, y: 200 })
})

test("resolveGutterAnnotationAnchorsInPage returns nothing when the page has no layout box", () => {
  const annotations = buildPageGutterAnnotations({
    pageSegments: pageSegments(0),
    comments: [commentFixture({ id: "comment", commentId: 1, nodeIndex: 0, text: "x" })],
  })
  const page = fakePageElement({
    clientRect: rect(0, 0, 0, 0),
    pageOffsetSize: { width: PAGE_WIDTH, height: PAGE_HEIGHT },
    matches: [
      {
        matches: () => true,
        element: {
          ...fakeAnchor(rect(0, 0, 10, 10)),
          hasAttribute: () => false,
          getAttribute: () => null,
          querySelectorAll: () => [],
        },
      },
    ],
  })

  assert.equal(resolveGutterAnnotationAnchorsInPage(page, annotations).size, 0)
})

// ── 3. docx-core: card layout ──────────────────────────────────────
function annotationsForLayout() {
  return buildPageGutterAnnotations({
    pageSegments: pageSegments(0, 1, 2),
    comments: [
      commentFixture({ id: "comment-low", commentId: 1, nodeIndex: 2, text: "near the end" }),
      commentFixture({ id: "comment-high", commentId: 2, nodeIndex: 0, text: "near the top" }),
    ],
    trackedChanges: [
      trackedChangeFixture({ id: "change-middle", kind: "insertion", nodeIndex: 1, text: "middle" }),
    ],
  })
}

test("layoutTrackedChangesForPage orders cards by anchor position instead of document or kind order", () => {
  const annotations = annotationsForLayout()
  const anchors = new Map([
    ["comment:comment-high", { x: 72, y: 120 }],
    ["change:change-middle", { x: 72, y: 40 }],
    ["comment:comment-low", { x: 72, y: 700 }],
  ])

  const positioned = layoutTrackedChangesForPage(annotations, anchors, undefined, PAGE_WIDTH, PAGE_HEIGHT)

  assert.deepEqual(
    positioned.map((entry) => entry.annotation.id),
    ["change:change-middle", "comment:comment-high", "comment:comment-low"]
  )
  assert.deepEqual(
    positioned.map((entry) => entry.anchorY),
    [40, 120, 700]
  )
  // Cards start at their anchor and never overlap.
  positioned.forEach((entry) => {
    assert.ok(entry.top >= 8, "cards stay inside the page")
  })
  for (let index = 1; index < positioned.length; index += 1) {
    const previous = positioned[index - 1]
    assert.ok(
      positioned[index].top >= previous.top + previous.heightPx,
      "stacked cards must not overlap"
    )
  }
})

test("layoutTrackedChangesForPage reserves measured card heights when stacking", () => {
  const annotations = annotationsForLayout().filter((annotation) => annotation.comment)
  const anchors = new Map([
    [annotations[0].id, { x: 72, y: 100 }],
    [annotations[1].id, { x: 72, y: 160 }],
  ])

  const withoutMeasurement = layoutTrackedChangesForPage(annotations, anchors, undefined, PAGE_WIDTH, PAGE_HEIGHT)
  const tallCardId = annotations[0].id
  const withMeasurement = layoutTrackedChangesForPage(
    annotations,
    anchors,
    new Map([[tallCardId, 240]]),
    PAGE_WIDTH,
    PAGE_HEIGHT
  )

  assert.equal(
    withMeasurement.find((entry) => entry.annotation.id === tallCardId).heightPx,
    240
  )
  assert.ok(
    withMeasurement[1].top > withoutMeasurement[1].top,
    "a taller first card pushes the next card further down"
  )
})

test("layoutTrackedChangesForPage keeps same-line cards in document order and stable under jitter", () => {
  const annotations = annotationsForLayout()
  const sameLine = annotations.map((annotation) => ({
    ...annotation,
    // Same text line: 1px of measurement jitter between runs.
    anchorY: 300 + annotations.indexOf(annotation) % 2,
  }))
  const anchors = new Map(
    sameLine.map((annotation, index) => [
      annotation.id,
      { x: 72, y: 300 + (index % 2) },
    ])
  )

  const first = layoutTrackedChangesForPage(sameLine, anchors, undefined, PAGE_WIDTH, PAGE_HEIGHT)
  const jittered = new Map(
    sameLine.map((annotation, index) => [
      annotation.id,
      { x: 72, y: 301 - (index % 2) },
    ])
  )
  const second = layoutTrackedChangesForPage(sameLine, jittered, undefined, PAGE_WIDTH, PAGE_HEIGHT)

  assert.deepEqual(
    first.map((entry) => entry.annotation.id),
    sameLine.map((annotation) => annotation.id),
    "same-line cards follow document order"
  )
  assert.deepEqual(
    second.map((entry) => entry.annotation.id),
    first.map((entry) => entry.annotation.id),
    "1px anchor jitter must not reorder cards"
  )
})

test("layoutTrackedChangesForPage anchors unmeasured cards next to their neighbours", () => {
  const annotations = annotationsForLayout()
  // Only the middle annotation resolved; the others must not fall back to a
  // page-wide synthetic slot far away from it.
  const anchors = new Map([["change:change-middle", { x: 72, y: 480 }]])

  const positioned = layoutTrackedChangesForPage(annotations, anchors, undefined, PAGE_WIDTH, PAGE_HEIGHT)
  const anchorYs = positioned.map((entry) => entry.anchorY)

  assert.equal(anchorYs.length, 3)
  assert.deepEqual(anchorYs, [480, 480, 480], "unresolved cards inherit the nearest measured anchor")
  assert.deepEqual(
    positioned.map((entry) => entry.annotation.id).sort(),
    annotations.map((annotation) => annotation.id).sort()
  )
})

// ── 4. vue-docx: gutter metrics ────────────────────────────────────
test("annotation gutter metrics stay constant in screen px across zoom levels", () => {
  assert.equal(annotationGutterReservedWidthPx(), DOCX_ANNOTATION_GUTTER_GAP_PX + DOCX_ANNOTATION_GUTTER_WIDTH_PX)

  const atHalf = annotationGutterPaneStyle(0.5, PAGE_HEIGHT)
  const atDouble = annotationGutterPaneStyle(2, PAGE_HEIGHT)

  assert.equal(atHalf.width, `${DOCX_ANNOTATION_GUTTER_WIDTH_PX}px`)
  assert.equal(atDouble.width, `${DOCX_ANNOTATION_GUTTER_WIDTH_PX}px`)
  assert.equal(atHalf.transform, "scale(2)")
  assert.equal(atDouble.transform, "scale(0.5)")
  assert.equal(atHalf.left, `calc(100% + ${DOCX_ANNOTATION_GUTTER_GAP_PX / 0.5}px)`)
  assert.equal(atDouble.left, `calc(100% + ${DOCX_ANNOTATION_GUTTER_GAP_PX / 2}px)`)
  // The pane covers exactly the paper in screen px at any zoom.
  assert.equal(atHalf.height, `${PAGE_HEIGHT * 0.5}px`)
  assert.equal(atDouble.height, `${PAGE_HEIGHT * 2}px`)

  assert.equal(annotationGutterCardTopPx(200, 2), 400)
  assert.equal(annotationGutterCardTopPx(200, 0.5), 100)
})

// ── 5. DocxPageSurface wiring ──────────────────────────────────────
async function mountPageSurface(overrides = {}) {
  return mount(DocxPageSurface, {
    pageIndex: 0,
    pageLayout: PAGE_LAYOUT,
    pageNodeSegments: pageSegments(0, 1, 2),
    model: modelWithParagraphs(["First paragraph", "Second paragraph", "Third paragraph"]),
    pageWidthPx: PAGE_WIDTH,
    pageContentWidthPx: PAGE_WIDTH - MARGINS.left - MARGINS.right,
    pageNumber: 1,
    totalPages: 1,
    trackedChangesEnabled: true,
    commentsEnabled: true,
    comments: [
      commentFixture({ id: "comment-node-2", commentId: 4, nodeIndex: 2, text: "Please confirm the summary." }),
      commentFixture({ id: "comment-node-0", commentId: 1, nodeIndex: 0, text: "Check the intro." }),
    ],
    trackedChanges: [
      trackedChangeFixture({ id: "change-node-1", kind: "insertion", nodeIndex: 1, text: "inserted words" }),
    ],
    zoomFactor: 2,
    ...overrides,
  })
}

test("DocxPageSurface positions comment cards against their anchors in document order", async () => {
  const mounted = await mountPageSurface()
  try {
    const cards = nodesWithClass(mounted.root, "docx-gutter-card")
    assert.equal(cards.length, 3)
    assert.deepEqual(
      cards.map((card) => card.props["data-docx-gutter-card-index"]),
      [1, 2, 3]
    )
    // Document order: comment on node 0, change on node 1, comment on node 2.
    assert.deepEqual(
      cards.map((card) => card.props["data-docx-gutter-card-id"]),
      ["comment:comment-node-0", "change:change-node-1", "comment:comment-node-2"]
    )

    const tops = cards.map((card) => Number.parseFloat(card.props.style.top))
    assert.deepEqual(tops, [...tops].sort((left, right) => left - right))
    assert.equal(new Set(tops).size, tops.length, "cards are not stacked at one fixed offset")

    const gutter = dataNodes(mounted.root, "data-docx-annotation-gutter")[0]
    assert.ok(gutter, "the gutter pane renders")
    // Counter-scaled pane: 200% page zoom shows 100%-sized card text.
    assert.equal(gutter.props.style.transform, "scale(0.5)")
    assert.equal(gutter.props.style.width, `${DOCX_ANNOTATION_GUTTER_WIDTH_PX}px`)
    assert.notEqual(gutter.props.style.top, "72px")

    assert.deepEqual(mounted.warnings, [])
  } finally {
    mounted.app.unmount()
  }
})

test("DocxPageSurface numbers the anchor markers so they match their cards", async () => {
  const mounted = await mountPageSurface()
  try {
    const markers = dataNodes(mounted.root, "data-docx-comment-marker")
    assert.deepEqual(
      markers.map((marker) => marker.props["data-docx-comment-marker"]),
      [1, 3]
    )
    assert.deepEqual(
      markers.map((marker) => marker.props["data-docx-gutter-marker-id"]),
      ["comment:comment-node-0", "comment:comment-node-2"]
    )
    assert.deepEqual(
      dataNodes(mounted.root, "data-docx-gutter-leader-id").map((node) => node.props["data-docx-gutter-leader-id"]),
      ["comment:comment-node-0", "change:change-node-1", "comment:comment-node-2"]
    )
    assert.deepEqual(mounted.warnings, [])
  } finally {
    mounted.app.unmount()
  }
})

test("DocxPageSurface hides comment cards and markers when comments are toggled off", async () => {
  const mounted = await mountPageSurface({ commentsEnabled: false })
  try {
    const cards = nodesWithClass(mounted.root, "docx-gutter-card")
    assert.equal(cards.length, 1)
    assert.equal(cards[0].props["data-docx-gutter-card-id"], "change:change-node-1")
    assert.deepEqual(dataNodes(mounted.root, "data-docx-comment-marker"), [])
    assert.deepEqual(mounted.warnings, [])
  } finally {
    mounted.app.unmount()
  }
})

// ── 6. DocxViewerRoot zoom wiring ──────────────────────────────────
function modelWithComment() {
  const model = createBlankDocumentModel()
  return {
    ...model,
    metadata: {
      ...model.metadata,
      comments: [
        { id: 1, author: "Alex Reviewer", date: "2026-01-02T00:00:00Z", text: "Check this." },
      ],
    },
    nodes: [
      {
        type: "paragraph",
        children: [{ type: "text", text: "Hello" }],
        sourceXml:
          '<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>Hello</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>',
      },
      { type: "paragraph", children: [{ type: "text", text: "After" }] },
    ],
  }
}

test("DocxViewerRoot reserves a constant gutter footprint and counter-scales it at any zoom", async () => {
  const model = modelWithComment()
  assert.equal(collectCommentsFromModel(model).length, 1, "fixture must carry one comment")

  for (const [zoomScale, expectedScale] of [[50, "scale(2)"], [200, "scale(0.5)"]]) {
    const mounted = await mount(DocxViewerRoot, {
      model,
      zoomScale,
      showComments: true,
    })
    try {
      const wrappers = dataNodes(mounted.root, "data-docx-page-wrapper")
      assert.equal(wrappers.length, 1, "one page is visible")

      const pageSurface = dataNodes(mounted.root, "data-docx-page-surface")[0]
      assert.ok(pageSurface, "the page surface renders")
      const pageWidth = Number.parseFloat(pageSurface.props.style.width)
      assert.ok(Number.isFinite(pageWidth) && pageWidth > 0)

      const zoom = zoomScale / 100
      assert.equal(
        wrappers[0].props.style.width,
        `${pageWidth * zoom + annotationGutterReservedWidthPx()}px`,
        "the gutter footprint is a screen-space constant"
      )

      const gutter = dataNodes(mounted.root, "data-docx-annotation-gutter")[0]
      assert.ok(gutter, "the annotation pane renders beside the paper")
      assert.equal(gutter.props.style.transform, expectedScale)
      assert.equal(gutter.props.style.width, `${DOCX_ANNOTATION_GUTTER_WIDTH_PX}px`)

      assert.deepEqual(mounted.warnings, [])
    } finally {
      mounted.app.unmount()
    }
  }
})

test("DocxViewerRoot renders annotation cards as an overlay, not in flow under the page", async () => {
  const mounted = await mount(DocxViewerRoot, {
    model: modelWithComment(),
    zoomScale: 100,
    showComments: true,
  })
  try {
    const gutter = dataNodes(mounted.root, "data-docx-annotation-gutter")[0]
    assert.ok(gutter)
    // The pane must never be `position: static`: a static pane sits under the
    // body text and inflates the measured page height (issue #29, item 5).
    assert.equal(gutter.props.style.top, "0px")
    const pageSurface = dataNodes(mounted.root, "data-docx-page-surface")[0]
    assert.ok(pageSurface)
    assert.equal(nodesWithClass(mounted.root, "docx-gutter-card").length, 1)
    assert.deepEqual(mounted.warnings, [])
  } finally {
    mounted.app.unmount()
  }
})
