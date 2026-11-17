import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromVueDocx = createRequire(
  new URL("../../packages/vue-docx/package.json", import.meta.url),
);
const core = await import(
  pathToFileURL(requireFromVueDocx.resolve("@arcships/docx-core")).href
);

const PAGE_ORIGIN = { left: 0, top: 0, marginLeft: 96, marginTop: 72 };
const ROTATED_90 = 5400000; // DrawingML rotation units: 60000 per degree.

function imageRun(overrides = {}) {
  return {
    type: "image",
    src: "data:image/png;base64,AAAA",
    alt: "image",
    widthPx: 120,
    heightPx: 80,
    ...overrides,
  };
}

function rotatedSourceXml(rotationRaw) {
  return `<a:xfrm rot="${rotationRaw}"><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm>`;
}

function paragraphWith(children) {
  return { type: "paragraph", children };
}

/**
 * Pretext text measurement needs a 2D canvas to measure glyph widths. Node has
 * none, so install a deterministic stub (8px per character) for the tests that
 * exercise line layout; without it `canUsePretext()` is false and every layout
 * call returns undefined.
 */
function withFakeCanvas(run) {
  const previousDocument = globalThis.document;
  globalThis.document = {
    createElement() {
      return {
        getContext() {
          return {
            font: "",
            measureText(text) {
              return { width: String(text ?? "").length * 8 };
            },
          };
        },
      };
    },
  };
  try {
    return run();
  } finally {
    if (previousDocument === undefined) {
      delete globalThis.document;
    } else {
      globalThis.document = previousDocument;
    }
  }
}

// --- 1. Page origin ------------------------------------------------------

test("floating page origin measures page anchors from the page edge and margin anchors from the margins", () => {
  const origin = core.resolveFloatingPageOriginPx(816, {
    top: 72,
    right: 72,
    bottom: 72,
    left: 96,
  });
  assert.deepEqual(origin, {
    left: 0,
    top: 0,
    marginLeft: 96,
    marginTop: 72,
    columnLeft: 96,
    columnTop: 72,
    pageWidth: 816,
  });

  assert.deepEqual(core.resolveFloatingPageOriginPx(undefined, undefined), {
    left: 0,
    top: 0,
    marginLeft: 0,
    marginTop: 0,
    columnLeft: 0,
    columnTop: 0,
    pageWidth: core.DEFAULT_DOC_PAGE_WIDTH,
  });
});

test("a margin-anchored image lands at the left margin instead of the page edge", () => {
  const image = imageRun({
    floating: {
      xPx: 10,
      yPx: 20,
      horizontalRelativeTo: "margin",
      verticalRelativeTo: "margin",
    },
  });
  const origin = core.resolveFloatingPageOriginPx(816, {
    top: 72,
    right: 72,
    bottom: 72,
    left: 96,
  });

  const anchored = core.absoluteFloatingImageStyle(image, {
    pageOriginLeft: origin.left,
    pageOriginTop: origin.top,
    marginOriginLeft: origin.marginLeft,
    marginOriginTop: origin.marginTop,
    columnOriginLeft: origin.columnLeft,
    columnOriginTop: origin.columnTop,
  });
  assert.equal(anchored.position, "absolute");
  assert.equal(anchored.left, 106, "xPx is measured from the left margin");
  assert.equal(anchored.top, 92, "yPx is measured from the top margin");

  // Without an origin every margin/page anchor collapses onto the page corner.
  const withoutOrigin = core.absoluteFloatingImageStyle(image, {});
  assert.equal(withoutOrigin.left, 10);
});

// --- 2. Behind-text z-order ---------------------------------------------

test("behind-text images get a negative z-index while in-front images stay above the text", () => {
  const behind = core.absoluteFloatingImageStyle(
    imageRun({
      floating: {
        xPx: 0,
        yPx: 0,
        horizontalRelativeTo: "page",
        verticalRelativeTo: "page",
        behindDocument: true,
      },
    }),
    PAGE_ORIGIN,
  );
  const inFront = core.absoluteFloatingImageStyle(
    imageRun({
      floating: {
        xPx: 0,
        yPx: 0,
        horizontalRelativeTo: "page",
        verticalRelativeTo: "page",
      },
    }),
    PAGE_ORIGIN,
  );

  assert.ok(Number(behind.zIndex) < 0, `behindDoc must paint under the text, got ${behind.zIndex}`);
  assert.ok(Number(inFront.zIndex) > 0, `in-front images paint above the text, got ${inFront.zIndex}`);
});

// --- 3. Rotation --------------------------------------------------------

test("a 90-degree rotated 200x100 image occupies a 100x200 frame", () => {
  const image = imageRun({
    widthPx: 200,
    heightPx: 100,
    sourceXml: rotatedSourceXml(ROTATED_90),
  });

  assert.deepEqual(core.resolveImageFrameSizePx(image, 200, 100), {
    widthPx: 100,
    heightPx: 200,
    rotationSwapped: true,
  });

  const transform = core.resolveImageRenderTransformStyle(image, {
    frameWidthPx: 100,
    frameHeightPx: 200,
    fillFrame: true,
  });
  assert.equal(transform.width, "200px", "the rotated content keeps the unrotated width");
  assert.equal(transform.height, "100px", "the rotated content keeps the unrotated height");
  assert.equal(transform.transformOrigin, "top left");
  assert.match(String(transform.transform), /rotate\(90deg\)$/);
});

test("unrotated and arbitrarily rotated images keep their declared frame", () => {
  const plain = imageRun({ widthPx: 200, heightPx: 100 });
  assert.deepEqual(core.resolveImageFrameSizePx(plain, 200, 100), {
    widthPx: 200,
    heightPx: 100,
    rotationSwapped: false,
  });

  const tilted = imageRun({
    widthPx: 200,
    heightPx: 100,
    sourceXml: rotatedSourceXml(45 * 60000),
  });
  assert.equal(
    core.resolveImageFrameSizePx(tilted, 200, 100).rotationSwapped,
    false,
    "only quarter turns swap the frame",
  );
});

// --- 4. wrapText drives the measured free band ---------------------------

function wrappedParagraph(floatingOverrides) {
  return paragraphWith([
    imageRun({
      widthPx: 120,
      heightPx: 120,
      floating: {
        wrapType: "square",
        horizontalRelativeTo: "column",
        verticalRelativeTo: "paragraph",
        xPx: 0,
        yPx: 0,
        distLPx: 0,
        distRPx: 0,
        distTPx: 0,
        distBPx: 0,
        ...floatingOverrides,
      },
    }),
    { type: "text", text: "wrapped paragraph text ".repeat(4) },
  ]);
}

test("wrapText decides which band beside a floating image the text may use", () => {
  const textOnLeft = core.resolveWrappedFloatingTextBands(
    wrappedParagraph({ wrapText: "left" }).children[0],
    { containerWidthPx: 320, imageWidthPx: 120 },
  );
  assert.deepEqual(textOnLeft, {
    side: "right",
    textOnLeft: true,
    textOnRight: false,
  });

  const textOnRight = core.resolveWrappedFloatingTextBands(
    wrappedParagraph({ wrapText: "right" }).children[0],
    { containerWidthPx: 320, imageWidthPx: 120 },
  );
  assert.deepEqual(textOnRight, {
    side: "left",
    textOnLeft: false,
    textOnRight: true,
  });
});

test("text measurement excludes the side disabled by wrapText", () => {
  withFakeCanvas(() => {
    const textOnLeft = core.resolveParagraphDualWrappedTextLayout(
      wrappedParagraph({ wrapText: "left" }),
      320,
      16,
    );
    const textOnRight = core.resolveParagraphDualWrappedTextLayout(
      wrappedParagraph({ wrapText: "right" }),
      320,
      16,
    );
    assert.ok(textOnLeft, "measurement should model the wrap exclusion");
    assert.ok(textOnRight, "measurement should model the wrap exclusion");

    const leftBand = textOnLeft.geometries[0].exclusion;
    const rightBand = textOnRight.geometries[0].exclusion;
    assert.deepEqual(
      { left: leftBand.left, right: leftBand.right },
      { left: 200, right: 320 },
      "wrapText=left: the image owns the right edge, text keeps [0, 200]",
    );
    assert.deepEqual(
      { left: rightBand.left, right: rightBand.right },
      { left: 0, right: 120 },
      "wrapText=right: the image owns the left edge, text keeps [120, 320]",
    );

    // The measured line boxes must start in the free band, not under the image.
    const firstFragmentX = (layout) => layout.layout.lines[0]?.fragments[0]?.x;
    assert.equal(firstFragmentX(textOnLeft), 0);
    assert.equal(firstFragmentX(textOnRight), 120);
  });
});

test("tight and through fall back to the square exclusion", () => {
  withFakeCanvas(() => {
    const square = core.resolveParagraphDualWrappedTextLayout(
      wrappedParagraph({ wrapType: "square", wrapText: "right" }),
      320,
      16,
    );
    for (const wrapType of ["tight", "through"]) {
      const layout = core.resolveParagraphDualWrappedTextLayout(
        wrappedParagraph({ wrapType, wrapText: "right" }),
        320,
        16,
      );
      assert.ok(layout, `${wrapType} should still wrap like square`);
      assert.deepEqual(
        layout.geometries[0].exclusion,
        square.geometries[0].exclusion,
        `${wrapType} falls back to the rectangular square exclusion`,
      );
    }
  });
});

// --- 5. Cross-paragraph exclusion ---------------------------------------

test("the paragraph after a floating image starts below it when the image outlives its anchor", () => {
  const tallFloat = imageRun({
    widthPx: 120,
    heightPx: 240,
    floating: {
      wrapType: "square",
      horizontalRelativeTo: "column",
      verticalRelativeTo: "paragraph",
      xPx: 0,
      yPx: 0,
      distTPx: 0,
      distBPx: 0,
    },
  });
  const anchor = paragraphWith([tallFloat, { type: "text", text: "anchor" }]);
  const following = paragraphWith([{ type: "text", text: "following" }]);

  const anchorHeightPx = core.estimateParagraphHeightPx(anchor, 480);
  const [anchorClearance, followingClearance] =
    core.resolveWrappedImageFlowClearancePx([anchor, following], 480);

  assert.equal(anchorClearance.marginTopPx, 0, "the anchor itself is not pushed down");
  assert.ok(
    Math.abs(followingClearance.marginTopPx - Math.max(0, 240 - anchorHeightPx)) <= 1,
    `the following paragraph clears the part of the float the anchor does not cover (got ${followingClearance.marginTopPx})`,
  );
  assert.ok(followingClearance.marginTopPx > 0);
});

test("behind-text and top-and-bottom floats do not push following paragraphs down", () => {
  const baseImage = {
    widthPx: 120,
    heightPx: 240,
    floating: {
      horizontalRelativeTo: "column",
      verticalRelativeTo: "paragraph",
      xPx: 0,
      yPx: 0,
    },
  };
  const following = paragraphWith([{ type: "text", text: "following" }]);

  const behind = paragraphWith([
    imageRun({
      ...baseImage,
      floating: { ...baseImage.floating, wrapType: "square", behindDocument: true },
    }),
    { type: "text", text: "anchor" },
  ]);
  const topAndBottom = paragraphWith([
    imageRun({
      ...baseImage,
      floating: { ...baseImage.floating, wrapType: "topAndBottom" },
    }),
    { type: "text", text: "anchor" },
  ]);

  assert.equal(
    core.resolveWrappedImageFlowClearancePx([behind, following], 480)[1].marginTopPx,
    0,
    "behind-text images are meant to sit under the text",
  );
  assert.equal(
    core.resolveWrappedImageFlowClearancePx([topAndBottom, following], 480)[1].marginTopPx,
    0,
    "top-and-bottom floats already grow the anchor paragraph",
  );
});
