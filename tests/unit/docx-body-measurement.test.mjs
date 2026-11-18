// Clipped body-rect measurement (issue #135).
//
// Covers the ported `createClippedBodyRectReader` (clip intersection, per-call
// caching, zoom normalization), the `data-docx-pagination-ignore` pruning, and
// the clipped-cell fixture where the measured height must equal the visible
// height instead of the scroll area.

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";

const core = await import(
  new URL("../../packages/docx-core/dist/index.js", import.meta.url).href
);

const fixture = JSON.parse(
  readFileSync(
    new URL("../fixtures/docx/clipped-cell-measurement.json", import.meta.url),
    "utf8"
  )
);

// ---------------------------------------------------------------------------
// Fake DOM: plain elements with a rect, an inline overflow style and parent
// links, which is everything the clipped rect reader is allowed to touch.
// ---------------------------------------------------------------------------

function rect(top, left, width, height) {
  return { top, left, right: left + width, bottom: top + height, width, height };
}

function buildElement(spec, counters = { rectReads: 0 }) {
  const box = rect(
    spec.rect.top,
    spec.rect.left,
    spec.rect.right - spec.rect.left,
    spec.rect.bottom - spec.rect.top
  );
  const element = {
    name: spec.name ?? "element",
    parentElement: null,
    style: spec.overflow ? { overflow: spec.overflow } : {},
    children: [],
    hasAttribute(name) {
      return name === "data-docx-pagination-ignore" && spec.ignored === true;
    },
    getBoundingClientRect() {
      counters.rectReads += 1;
      return { ...box };
    },
    clientHeight: spec.clientHeightPx,
    scrollHeight: spec.scrollHeightPx,
  };
  for (const child of spec.children ?? []) {
    const node = buildElement(child, counters);
    node.parentElement = element;
    element.children.push(node);
  }
  return element;
}

function element(name, box, options = {}) {
  return {
    name,
    rect: box,
    overflow: options.overflow,
    ignored: options.ignored,
    children: options.children ?? [],
  };
}

function findByName(root, name) {
  if (root.name === name) return root;
  for (const child of root.children) {
    const match = findByName(child, name);
    if (match) return match;
  }
  return undefined;
}

// ---------------------------------------------------------------------------
// Reader behaviour
// ---------------------------------------------------------------------------

test("clipped body rect reader intersects every clipping ancestor", () => {
  const counters = { rectReads: 0 };
  const root = buildElement(
    element("page-body", { top: 72, left: 72, right: 744, bottom: 984 }, {
      overflow: "hidden",
      children: [
        element("cell", { top: 600, left: 72, right: 744, bottom: 850 }, {
          overflow: "hidden",
          children: [
            element("paragraph", { top: 610, left: 82, right: 734, bottom: 1400 }),
          ],
        }),
      ],
    }),
    counters
  );
  const cell = findByName(root, "cell");
  const paragraph = findByName(root, "paragraph");

  const reader = core.createClippedBodyRectReader(root);
  const clipped = reader(paragraph);

  assert.ok(clipped, "the paragraph is measurable");
  assert.equal(
    clipped.bottomPx,
    850 - 72,
    "the clipped paragraph reports the cell's visible bottom, not its own"
  );
  assert.equal(clipped.topPx, 610 - 72);
  assert.equal(clipped.heightPx, 850 - 610);
  assert.equal(reader(cell).bottomPx, 850 - 72);
  assert.equal(
    reader(root),
    undefined,
    "the measured root is not a descendant of itself"
  );

  const readsAfterFirstPass = counters.rectReads;
  assert.equal(reader(paragraph), clipped, "results are cached per element");
  assert.equal(
    counters.rectReads,
    readsAfterFirstPass,
    "a cached element does not re-read its rect"
  );
});

test("clipping ancestors above the measured root are ignored", () => {
  // A viewer scroll container also clips (overflow: auto) but a page height
  // measured through it would depend on the reader's scroll position.
  const root = buildElement(
    element("viewer", { top: 0, left: 0, right: 816, bottom: 400 }, {
      overflow: "auto",
      children: [
        element("page-body", { top: 72, left: 72, right: 744, bottom: 984 }, {
          overflow: "hidden",
          children: [
            element("paragraph", { top: 72, left: 72, right: 744, bottom: 900 }),
          ],
        }),
      ],
    })
  );
  const body = findByName(root, "page-body");
  const paragraph = findByName(root, "paragraph");

  const clipped = core.createClippedBodyRectReader(body)(paragraph);
  assert.ok(clipped);
  assert.equal(
    clipped.bottomPx,
    900 - 72,
    "the paragraph stays visible; only the body clips inside the measured root"
  );
});

test("clipped body rect reader normalizes zoomed rects back to layout pixels", () => {
  const unscaled = buildElement(
    element("page-body", { top: 72, left: 72, right: 744, bottom: 984 }, {
      overflow: "hidden",
      children: [
        element("cell", { top: 600, left: 72, right: 744, bottom: 850 }, {
          overflow: "hidden",
          children: [
            element("paragraph", { top: 610, left: 82, right: 734, bottom: 1400 }),
          ],
        }),
      ],
    })
  );
  const zoom = 1.5;
  const scale = (value) => value * zoom;
  const zoomed = buildElement(
    element(
      "page-body",
      {
        top: scale(72),
        left: scale(72),
        right: scale(744),
        bottom: scale(984),
      },
      {
        overflow: "hidden",
        children: [
          element(
            "cell",
            {
              top: scale(600),
              left: scale(72),
              right: scale(744),
              bottom: scale(850),
            },
            {
              overflow: "hidden",
              children: [
                element("paragraph", {
                  top: scale(610),
                  left: scale(82),
                  right: scale(734),
                  bottom: scale(1400),
                }),
              ],
            }
          ),
        ],
      }
    )
  );

  const plain = core.createClippedBodyRectReader(unscaled)(
    findByName(unscaled, "paragraph")
  );
  const scaled = core.createClippedBodyRectReader(zoomed, {
    zoomFactor: zoom,
  })(findByName(zoomed, "paragraph"));

  assert.deepEqual(scaled, plain, "a zoomed subtree measures the same layout box");
});

test("rects without right/bottom are completed from width/height", () => {
  const paragraph = {
    parentElement: null,
    style: {},
    children: [],
    getBoundingClientRect: () => ({ top: 10, left: 4, width: 100, height: 20 }),
  };
  const root = {
    parentElement: null,
    style: { overflow: "hidden" },
    children: [paragraph],
    getBoundingClientRect: () => ({ top: 0, left: 0, right: 800, bottom: 1000 }),
  };
  paragraph.parentElement = root;

  const measured = core.createClippedBodyRectReader(root)(paragraph);
  assert.ok(measured);
  assert.equal(measured.rightPx, 104);
  assert.equal(measured.bottomPx, 30);
});

test("elements outside the measured root are not measurable", () => {
  const root = buildElement(
    element("page-body", { top: 0, left: 0, right: 800, bottom: 1000 }, {
      overflow: "hidden",
    })
  );
  const foreign = buildElement(
    element("toolbar", { top: 0, left: 0, right: 800, bottom: 60 })
  );

  const reader = core.createClippedBodyRectReader(root);
  assert.equal(
    reader(foreign),
    undefined,
    "a node that is not a descendant of the root cannot be measured"
  );
  assert.equal(
    core.resolveClippedBodyRenderedBottomPx(root),
    undefined,
    "a root without measurable descendants has no clipped bottom"
  );
});

test("elementClipsOverflow reads computed overflow before the inline style", () => {
  const inlineVisible = { style: { overflow: "visible" } };
  const original = globalThis.getComputedStyle;
  globalThis.getComputedStyle = () => ({ overflowX: "hidden", overflowY: "auto" });
  try {
    assert.equal(
      core.elementClipsOverflow(inlineVisible),
      true,
      "computed overflow wins over the inline shorthand"
    );
  } finally {
    if (original === undefined) delete globalThis.getComputedStyle;
    else globalThis.getComputedStyle = original;
  }

  assert.equal(
    core.elementClipsOverflow({ style: { overflow: "clip" } }),
    true,
    "inline overflow: clip still clips without a layout engine"
  );
  assert.equal(core.elementClipsOverflow({ style: { overflow: "visible" } }), false);
  assert.equal(core.elementClipsOverflow({}), false);
  assert.equal(core.elementClipsOverflow(null), false);
});

test("pagination-ignore subtrees are pruned from the measurement", () => {
  const counters = { rectReads: 0 };
  const root = buildElement(
    element("page-body", { top: 0, left: 0, right: 800, bottom: 1000 }, {
      children: [
        element("paragraph", { top: 0, left: 0, right: 800, bottom: 200 }),
        element("overlay", { top: 0, left: 0, right: 800, bottom: 1400 }, {
          ignored: true,
          children: [
            element("gutter card", { top: 0, left: 0, right: 800, bottom: 1400 }),
          ],
        }),
      ],
    }),
    counters
  );
  const overlay = findByName(root, "overlay");
  const card = findByName(root, "gutter card");
  assert.ok(overlay && card, "the fixture renders the ignored overlay subtree");

  assert.deepEqual(
    core.collectPageContentMeasurementDescendants(root).map((node) => node.name),
    ["paragraph"],
    "the ignored overlay and its subtree are skipped"
  );
  assert.equal(
    core.resolveClippedBodyRenderedBottomPx(root),
    200,
    "an ignored overlay cannot extend the body"
  );
  // The reader itself is not filtered: callers decide what to measure.
  assert.equal(core.createClippedBodyRectReader(root)(card).bottomPx, 1400);
});

// ---------------------------------------------------------------------------
// Fixture: a page whose last block is clipped by a table cell
// ---------------------------------------------------------------------------

function buildFixture() {
  const root = buildElement(fixture.root);
  return {
    root,
    body: findByName(root, "page-body"),
    cell: findByName(root, "cell (exact height slice)"),
  };
}

test("clipped cell fixture: measured page height equals the visible height", () => {
  const { root, body, cell } = buildFixture();
  const expected = fixture.expected;

  // The reader intersects the cell's clip box.
  const clippedCell = core.createClippedBodyRectReader(root)(cell);
  assert.equal(clippedCell.bottomPx, expected.cellVisibleBottomPx);

  const surfaceBottom = core.resolveClippedBodyRenderedBottomPx(root);
  assert.equal(
    surfaceBottom,
    expected.surfaceVisibleContentBottomPx,
    "the page's visible content ends at the page body's clipped bottom"
  );
  assert.ok(
    surfaceBottom < expected.unclippedSurfaceContentBottomPx,
    "the paragraphs clipped by the cell must not be measured"
  );

  const pageHeight = core.resolveClippedMeasuredHeightPx({
    // What the old measurement reported: the grown box / scroll area.
    fallbackHeightPx: expected.bodyScrollHeightPx,
    minimumHeightPx: expected.pageMeasuredHeightPx,
    clippedContentBottomPx: surfaceBottom,
  });
  assert.equal(
    pageHeight,
    expected.pageMeasuredHeightPx,
    "clipped content cannot inflate the page past one paper"
  );

  const bodyBottom = core.resolveClippedBodyRenderedBottomPx(body);
  assert.equal(bodyBottom, expected.bodyVisibleContentBottomPx);

  const bodyHeight = core.resolveClippedMeasuredHeightPx({
    fallbackHeightPx: expected.bodyScrollHeightPx,
    minimumHeightPx: expected.bodyClientHeightPx,
    clippedContentBottomPx: bodyBottom,
  });
  assert.equal(
    bodyHeight,
    expected.bodyMeasuredContentHeightPx,
    "page content height is the visible body height, not scrollHeight"
  );
  assert.ok(
    expected.bodyScrollHeightPx > bodyHeight,
    "the fixture still exposes the inflated scroll area the old path measured"
  );
});

test("clipped cell fixture: unclipped rects would inflate the page", () => {
  const { body } = buildFixture();
  const unclipped = core
    .collectPageContentMeasurementDescendants(body)
    .map((element) => ({
      bottomPx: readerWithoutClipping(element, body),
      widthPx: 1,
      heightPx: 1,
    }));
  const inflatedBottom = core.resolveMeasuredBodyRenderedBottomPx(unclipped);

  assert.equal(inflatedBottom, fixture.expected.bodyUnclippedContentBottomPx);
  assert.ok(
    inflatedBottom > fixture.expected.bodyVisibleContentBottomPx,
    "without the clip intersection the overflowing paragraphs win"
  );
  assert.equal(
    core.resolveClippedMeasuredHeightPx({
      minimumHeightPx: fixture.expected.pageMeasuredHeightPx,
      clippedContentBottomPx: inflatedBottom,
    }),
    inflatedBottom,
    "the page height would follow the clipped paragraphs without the reader"
  );
});

function readerWithoutClipping(element, root) {
  const rootBox = root.getBoundingClientRect();
  const box = element.getBoundingClientRect();
  return box.bottom - rootBox.top;
}

test("clipped cell fixture: the next page starts at the paper offset", () => {
  const { root } = buildFixture();
  const measuredPageHeightPx = core.resolveClippedMeasuredHeightPx({
    fallbackHeightPx: fixture.expected.bodyScrollHeightPx,
    minimumHeightPx: fixture.expected.pageMeasuredHeightPx,
    clippedContentBottomPx: core.resolveClippedBodyRenderedBottomPx(root),
  });
  assert.equal(measuredPageHeightPx, fixture.expected.pageMeasuredHeightPx);
  assert.equal(
    measuredPageHeightPx + core.DOC_PAGE_BREAK_GAP,
    fixture.expected.nextPageOffsetPx,
    "the following page starts one paper plus the page-break gap later"
  );
});

test("fallbacks keep a page measurable without a layout engine", () => {
  assert.equal(
    core.resolveClippedMeasuredHeightPx({
      fallbackHeightPx: 1400,
      minimumHeightPx: 1056,
    }),
    1400,
    "offsetHeight is the fallback when nothing could be measured"
  );
  assert.equal(
    core.resolveClippedMeasuredHeightPx({
      fallbackHeightPx: 900,
      minimumHeightPx: 1056,
    }),
    1056,
    "a page is never shorter than one paper"
  );
  assert.equal(core.resolveClippedMeasuredHeightPx({}), undefined);
  assert.equal(core.resolveClippedBodyRenderedBottomPx(null), undefined);
  assert.deepEqual(core.collectPageContentMeasurementDescendants(undefined), []);
});
