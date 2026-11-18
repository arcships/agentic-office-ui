// DOCX inline line metrics (issue #127).
//
// DOCX lays a line out from the glyph boxes of its runs: ascent, descent and a
// suppressible bottom space that a page break may trim. Before this suite the
// layout only knew `fontSizePt * 1.6`, so a line mixing a 28 pt inline image
// with sub/superscript runs shared one fake baseline and paginated at a height
// Word never produces.
//
// The assertions below pin:
//   - aggregateLineMetrics / checkLineVerticalFit / constrainParagraphLineSplit
//   - canvas `fontBoundingBoxAscent/Descent` measurement (with a headless
//     fallback when the host has no canvas)
//   - aggregation of fragments into a line box, snapped to the docGrid pitch
//   - the issue's acceptance case: one baseline, Word-matching height (±1 px)

import assert from "node:assert/strict";
import { after, test } from "node:test";

// pretext and the line-metrics module both need a text-measurement canvas;
// Node has none, so provide a deterministic shim before the package loads.
// `fontBoundingBoxAscent/Descent` are reported as Calibri-like ratios so the
// assertions can reason about real Word line heights.
const CALIBRI_ASCENT_EM = 0.952;
const CALIBRI_DESCENT_EM = 0.269;
const originalOffscreenCanvas = globalThis.OffscreenCanvas;

// Which vertical metrics the shim reports. "fontBoundingBox" is the real
// browser path; "actualBoundingBox" models engines that only expose ink boxes;
// "widthOnly" models a host that can measure advances but not font boxes, which
// is the branch a canvas-less host takes as well.
let canvasMetricsMode = "fontBoundingBox";

globalThis.OffscreenCanvas = class {
  constructor(width, height) {
    this.width = width;
    this.height = height;
  }
  getContext() {
    return {
      font: "12px sans-serif",
      measureText(text) {
        const fontSizePx = Number.parseFloat(
          /(\d+(?:\.\d+)?)px/.exec(this.font)?.[1] ?? "12"
        );
        const result = { width: text.length * fontSizePx * 0.5 };
        if (canvasMetricsMode === "fontBoundingBox") {
          result.fontBoundingBoxAscent = fontSizePx * CALIBRI_ASCENT_EM;
          result.fontBoundingBoxDescent = fontSizePx * CALIBRI_DESCENT_EM;
        }
        if (canvasMetricsMode === "actualBoundingBox") {
          result.actualBoundingBoxAscent = fontSizePx * 0.8;
          result.actualBoundingBoxDescent = fontSizePx * 0.22;
        }
        return result;
      },
    };
  }
};

after(() => {
  if (originalOffscreenCanvas === undefined) {
    delete globalThis.OffscreenCanvas;
  } else {
    globalThis.OffscreenCanvas = originalOffscreenCanvas;
  }
});

const core = await import(
  new URL("../../packages/docx-core/dist/index.js", import.meta.url).href
);

const TEXT_STYLE = { fontSizePt: 12 };
// 28 pt at 96 dpi: a 37 px inline image, i.e. taller than any 12 pt text line.
const IMAGE_HEIGHT_PX = Math.round((28 * 96) / 72);

function paragraph(children, style) {
  return {
    type: "paragraph",
    children,
    ...(style ? { style } : {}),
  };
}

function textRunsOnlyParagraph(text) {
  return paragraph([{ type: "text", text, style: TEXT_STYLE }]);
}

function acceptImageParagraph() {
  return paragraph([
    { type: "text", text: "H", style: TEXT_STYLE },
    { type: "text", text: "2", style: { ...TEXT_STYLE, verticalAlign: "subscript" } },
    { type: "text", text: "+", style: { ...TEXT_STYLE, verticalAlign: "superscript" } },
    {
      type: "image",
      src: "data:image/png;base64,AAAA",
      widthPx: IMAGE_HEIGHT_PX,
      heightPx: IMAGE_HEIGHT_PX,
    },
  ]);
}

function layoutParagraph(target, lineHeightPx, options = {}) {
  const source = core.buildParagraphPretextLayoutSource(target, {
    allowExplicitLineBreakText: true,
    expandTabsForLayout: true,
  });
  assert.ok(source, "expected a pretext layout source");
  const layout = core.layoutParagraphPretextSource(
    target,
    source,
    480,
    lineHeightPx,
    options.exclusions ?? [],
    options.docGridLinePitchPx
      ? { docGridLinePitchPx: options.docGridLinePitchPx }
      : undefined
  );
  assert.ok(layout, "expected a pretext layout");
  return layout;
}

// --- aggregateLineMetrics -------------------------------------------------

test("aggregateLineMetrics takes the tallest ascent, deepest descent and keeps the ink below the baseline unsuppressible", () => {
  const metrics = core.aggregateLineMetrics([
    { ascentPx: 14.4, descentPx: 4 },
    { ascentPx: 28, descentPx: 0 },
    { ascentPx: 6, descentPx: 9, unsuppressibleDescentPx: 5 },
  ]);

  assert.equal(metrics.ascentPx, 28, "the inline image owns the ascent");
  assert.equal(metrics.descentPx, 9, "the deepest descender owns the descent");
  assert.equal(
    metrics.unsuppressibleDescentPx,
    5,
    "only the declared ink descent must stay below the baseline"
  );
  assert.equal(
    metrics.suppressibleBottomSpacePx,
    4,
    "the rest of the descent is trimmable bottom space"
  );
  assert.equal(metrics.heightPx, 37, "ascent + descent");
});

test("aggregateLineMetrics defaults the unsuppressible descent to the full descent", () => {
  const metrics = core.aggregateLineMetrics([
    { ascentPx: 12, descentPx: 3 },
    { ascentPx: 10, descentPx: 4 },
  ]);

  assert.equal(metrics.unsuppressibleDescentPx, 4);
  assert.equal(metrics.suppressibleBottomSpacePx, 0);
  assert.equal(metrics.heightPx, 16);
});

test("aggregateLineMetrics clamps an unsuppressible descent that exceeds the descent and honours declared bottom space", () => {
  const clamped = core.aggregateLineMetrics([
    { ascentPx: 10, descentPx: 2, unsuppressibleDescentPx: 8 },
  ]);
  assert.equal(clamped.unsuppressibleDescentPx, 2, "no more than the descent");
  assert.equal(clamped.suppressibleBottomSpacePx, 0);

  const declared = core.aggregateLineMetrics([
    { ascentPx: 10, descentPx: 2, suppressibleBottomSpacePx: 6 },
  ]);
  assert.equal(declared.suppressibleBottomSpacePx, 6);
  assert.equal(
    declared.heightPx,
    18,
    "declared trimmable space is part of the descent it sits in"
  );
  assert.equal(declared.descentPx, 8);
  assert.equal(
    declared.suppressibleBottomSpacePx,
    declared.descentPx - declared.unsuppressibleDescentPx,
    "the box never advertises more trimmable space than it contains"
  );
});

test("aggregateLineMetrics of no elements is an empty line box", () => {
  assert.deepEqual(core.aggregateLineMetrics([]), {
    ascentPx: 0,
    descentPx: 0,
    unsuppressibleDescentPx: 0,
    suppressibleBottomSpacePx: 0,
    heightPx: 0,
  });
});

// --- font metrics ---------------------------------------------------------

test("fontVerticalMetrics measures fontBoundingBoxAscent/Descent from canvas measureText", () => {
  core.invalidateTextMeasurements();
  const metrics = core.fontVerticalMetrics("normal normal 400 16px Calibri");

  assert.equal(metrics.ascentPx, 15.23, "16 px * 0.952, measured not guessed");
  assert.equal(metrics.descentPx, 4.3, "16 px * 0.269");
});

test("fontVerticalMetrics falls back to actualBoundingBox metrics when the font box is missing", () => {
  canvasMetricsMode = "actualBoundingBox";
  core.invalidateTextMeasurements();
  const metrics = core.fontVerticalMetrics("normal normal 400 20px ActualOnly");

  assert.equal(metrics.ascentPx, 16, "20 px * 0.8 actual ascent");
  assert.equal(metrics.descentPx, 4.4, "20 px * 0.22 actual descent");
  canvasMetricsMode = "fontBoundingBox";
  core.invalidateTextMeasurements();
});

test("fontVerticalMetrics is deterministic when the host cannot report font boxes (SSR / worker)", () => {
  canvasMetricsMode = "widthOnly";
  core.invalidateTextMeasurements();
  const metrics = core.fontVerticalMetrics("normal normal 400 16px NoCanvasProbe");

  assert.deepEqual(metrics, { ascentPx: 14.4, descentPx: 4 });
  canvasMetricsMode = "fontBoundingBox";
  core.invalidateTextMeasurements();
});

test("scriptBaselineShiftPx matches the CSS super/sub offsets Word uses", () => {
  assert.equal(
    core.scriptBaselineShiftPx("superscript", 16),
    5.28,
    "Word raises a superscript by a third of the parent size"
  );
  assert.equal(core.scriptBaselineShiftPx("subscript", 16), -3.2);
  assert.equal(core.scriptBaselineShiftPx(undefined, 16), 0);
});

test("textLineMetricElement measures the run's own (scaled) font box", () => {
  const script = core.textLineMetricElement("normal normal 400 10px Calibri");

  assert.equal(script.ascentPx, 9.52, "10 px * 0.952");
  assert.equal(script.descentPx, 2.69);
  assert.equal(
    script.unsuppressibleDescentPx,
    2.69,
    "the glyph descent can never be trimmed"
  );
});

// --- fixed baseline / doc grid -------------------------------------------

test("withFixedLineBaseline re-seats a line inside an exact line height", () => {
  const metrics = core.withFixedLineBaseline(
    {
      ascentPx: 37,
      descentPx: 6,
      unsuppressibleDescentPx: 4.3,
      suppressibleBottomSpacePx: 1.7,
      heightPx: 43,
    },
    30
  );

  assert.equal(metrics.heightPx, 30);
  assert.equal(metrics.ascentPx, 30, "the tall ascent is clipped to the box");
  assert.equal(metrics.descentPx, 0);
  assert.equal(metrics.ascentPx + metrics.descentPx, metrics.heightPx);
});

test("resolveLineGridHeightPx snaps a line box up to the docGrid line pitch", () => {
  assert.equal(core.resolveLineGridHeightPx(43, undefined), 43);
  assert.equal(core.resolveLineGridHeightPx(43, 0), 43);
  assert.equal(core.resolveLineGridHeightPx(43, 20), 60, "three grid lines");
  assert.equal(core.resolveLineGridHeightPx(20, 20), 20, "an exact pitch stays");
  assert.equal(core.resolveLineGridHeightPx(2, 20), 20, "at least one grid line");
});

// --- vertical fit / line split -------------------------------------------

test("checkLineVerticalFit rescues a line by trimming only its suppressible bottom space", () => {
  const metrics = {
    ascentPx: 37,
    descentPx: 23,
    unsuppressibleDescentPx: 4.3,
    suppressibleBottomSpacePx: 18.7,
    heightPx: 60,
  };

  const full = core.checkLineVerticalFit(metrics, 60);
  assert.equal(full.fits, true);
  assert.equal(full.rescued, false);

  const rescued = core.checkLineVerticalFit(metrics, 42);
  assert.equal(rescued.fits, true);
  assert.equal(rescued.rescued, true, "the line only fits after trimming");
  assert.ok(
    Math.abs(rescued.suppressedHeightPx - 41.3) < 0.001,
    `ink height stays at 41.3, got ${rescued.suppressedHeightPx}`
  );
  assert.ok(Math.abs(rescued.suppressedBottomSpacePx - 18.7) < 0.001);

  const overflowing = core.checkLineVerticalFit(metrics, 40);
  assert.equal(
    overflowing.fits,
    false,
    "the unsuppressible descent never gets trimmed"
  );
});

test("constrainParagraphLineSplit walks per-line metrics and only rescues the last line", () => {
  const lineMetrics = [
    // 14 px text line with 1 px of trimmable bottom space.
    {
      ascentPx: 11,
      descentPx: 3,
      unsuppressibleDescentPx: 2,
      suppressibleBottomSpacePx: 1,
      heightPx: 14,
    },
    // 43 px line grown by an inline image; 18.7 px are trimmable leading.
    {
      ascentPx: 37,
      descentPx: 6,
      unsuppressibleDescentPx: 4.3,
      suppressibleBottomSpacePx: 1.7,
      heightPx: 43,
    },
    {
      ascentPx: 11,
      descentPx: 3,
      unsuppressibleDescentPx: 2,
      suppressibleBottomSpacePx: 1,
      heightPx: 14,
    },
  ];

  assert.equal(
    core.constrainParagraphLineSplit(lineMetrics, 0, 3, 57),
    2,
    "two full lines fit"
  );
  assert.equal(
    core.constrainParagraphLineSplit(lineMetrics, 0, 3, 42),
    1,
    "the image line cannot be trimmed below its ink"
  );
  assert.equal(
    core.constrainParagraphLineSplit(lineMetrics, 0, 3, 1000),
    3
  );
  assert.equal(
    core.constrainParagraphLineSplit(lineMetrics, 1, 3, 43),
    2,
    "an exact full-height fit is not a rescue"
  );
  assert.equal(
    core.constrainParagraphLineSplit(lineMetrics, 0, 3, 0),
    0,
    "no room, no lines"
  );
  assert.equal(
    core.constrainParagraphLineSplit(lineMetrics, 0, 1, 1000),
    1,
    "the caller's maxEndLineIndex is an upper bound"
  );

  const rescuing = [
    {
      ascentPx: 20,
      descentPx: 24,
      unsuppressibleDescentPx: 4,
      suppressibleBottomSpacePx: 20,
      heightPx: 44,
    },
  ];
  assert.equal(
    core.constrainParagraphLineSplit(rescuing, 0, 1, 24),
    1,
    "a page break may trim the bottom space to keep the line"
  );
  assert.equal(
    core.constrainParagraphLineSplit(rescuing, 0, 1, 23),
    0,
    "but never below the ascent plus the ink descent"
  );
});

test("sumLineHeightsPx adds only the requested line range", () => {
  const lineMetrics = [14, 43, 14].map((heightPx) => ({
    ascentPx: heightPx - 1,
    descentPx: 1,
    unsuppressibleDescentPx: 1,
    suppressibleBottomSpacePx: 0,
    heightPx,
  }));

  assert.equal(core.sumLineHeightsPx(lineMetrics, 0, 3), 71);
  assert.equal(core.sumLineHeightsPx(lineMetrics, 1, 2), 43);
  assert.equal(core.sumLineHeightsPx(undefined, 0, 3), 0);
});

// --- fragments and line boxes --------------------------------------------

test("line fragments record the baseline of their own font", () => {
  const target = textRunsOnlyParagraph("Hello");
  const lineHeightPx = core.estimateParagraphLineHeightPx(target);
  const layout = layoutParagraph(target, lineHeightPx);
  const line = layout.lines[0];

  assert.ok(line, "expected one laid-out line");
  assert.equal(layout.hasCustomVerticalMetrics, false);
  assert.equal(
    line.metrics.heightPx,
    lineHeightPx,
    "plain text keeps the paragraph line height (no line growth)"
  );
  assert.equal(
    line.baselineYPx,
    line.y + line.metrics.ascentPx,
    "the baseline is the line top plus the ascent"
  );
  assert.equal(layout.lineMetrics.length, layout.lineCount);
  assert.deepEqual(layout.lineMetrics[0], line.metrics);
});

// --- the issue's acceptance case -----------------------------------------

test("a 12 pt paragraph with a 28 pt inline image and sub/superscript keeps one baseline (issue #127 acceptance)", () => {
  const target = acceptImageParagraph();
  const lineHeightPx = core.estimateParagraphLineHeightPx(target);
  const layout = layoutParagraph(target, lineHeightPx);
  const line = layout.lines[0];

  assert.ok(line, "expected one laid-out line");
  assert.equal(layout.lineCount, 1, "image and runs share the line");
  assert.equal(
    layout.hasCustomVerticalMetrics,
    true,
    "the image forces a metric-driven line box"
  );

  // One baseline: the image bottom and every run's glyph box hang off it.
  const baselineYPx = line.baselineYPx;
  assert.equal(baselineYPx, line.y + IMAGE_HEIGHT_PX, "the image top is the ascent");
  assert.equal(line.metrics.ascentPx, IMAGE_HEIGHT_PX);

  const imageFragment = line.fragments.find(
    (fragment) => fragment.hasCustomVerticalMetrics === true && fragment.ascentPx === IMAGE_HEIGHT_PX
  );
  assert.ok(imageFragment, "the image fragment carries the object metrics");
  assert.equal(imageFragment.descentPx, 0, "an inline image has no descent");

  const scripts = line.fragments.filter(
    (fragment) => fragment.baselineShiftPx !== undefined
  );
  assert.equal(scripts.length, 2, "both sub/superscript fragments record a baseline shift");
  const shifts = scripts
    .map((fragment) => fragment.baselineShiftPx ?? 0)
    .sort((left, right) => left - right);
  assert.deepEqual(
    shifts,
    [
      core.scriptBaselineShiftPx("subscript", 16),
      core.scriptBaselineShiftPx("superscript", 16),
    ].sort((left, right) => left - right),
    "the subscript sits below and the superscript above the shared baseline"
  );
  for (const fragment of scripts) {
    assert.equal(
      core.fragmentBaselineYPx(line, fragment),
      baselineYPx - (fragment.baselineShiftPx ?? 0),
      "fragmentBaselineYPx answers where the run's glyphs sit"
    );
  }

  // Word: 37 px image on the baseline plus the 12 pt descender (4.3 px) = 41.3.
  // A sub/superscript never grows the line box (its ink may overflow it).
  const wordHeightPx = IMAGE_HEIGHT_PX + 16 * CALIBRI_DESCENT_EM;
  assert.ok(
    Math.abs(line.metrics.heightPx - wordHeightPx) <= 1,
    `line height ${line.metrics.heightPx} must match Word's ${wordHeightPx} within 1 px`
  );
  assert.equal(layout.height, line.metrics.heightPx);

  // Pagination must agree with the layout it will render.
  assert.equal(
    core.estimateParagraphHeightPx(target, 480),
    line.metrics.heightPx,
    "the paragraph height estimate uses the same line metrics"
  );
  const resolved = core.resolveParagraphLineMetrics(target, {
    availableWidthPx: 480,
    lineHeightPx,
    expectedLineCount: 1,
  });
  assert.ok(resolved, "line metrics are available to the paginator");
  assert.equal(resolved[0].heightPx, line.metrics.heightPx);
});

test("the acceptance line snaps to the docGrid pitch and the extra leading is suppressible", () => {
  const target = acceptImageParagraph();
  const docGridLinePitchPx = 20;
  const lineHeightPx = core.estimateParagraphLineHeightPx(
    target,
    docGridLinePitchPx
  );
  const layout = layoutParagraph(target, lineHeightPx, { docGridLinePitchPx });
  const metrics = layout.lines[0].metrics;

  assert.equal(metrics.heightPx, 60, "42 px of content needs three 20 px grid lines");
  assert.equal(metrics.ascentPx, IMAGE_HEIGHT_PX, "the baseline does not move");
  assert.ok(
    Math.abs(metrics.unsuppressibleDescentPx - 16 * CALIBRI_DESCENT_EM) < 0.01,
    "the 12 pt ink descent is intact"
  );
  assert.ok(
    Math.abs(
      metrics.suppressibleBottomSpacePx -
        (60 - IMAGE_HEIGHT_PX - 16 * CALIBRI_DESCENT_EM)
    ) < 0.01,
    "everything above the ink is trimmable"
  );

  const fit = core.checkLineVerticalFit(metrics, 42);
  assert.equal(fit.fits, true);
  assert.equal(fit.rescued, true, "a 20 px docGrid line is rescued on a short page");
});

test("a larger run grows its line; a uniform paragraph keeps the paragraph line height", () => {
  const mixed = paragraph([
    { type: "text", text: "a", style: TEXT_STYLE },
    { type: "text", text: "B", style: { fontSizePt: 28 } },
  ]);
  const mixedLineHeightPx = core.estimateParagraphLineHeightPx(mixed);
  const mixedLayout = layoutParagraph(mixed, mixedLineHeightPx);
  assert.equal(mixedLayout.hasCustomVerticalMetrics, true);
  assert.ok(
    mixedLayout.lines[0].metrics.heightPx > mixedLineHeightPx,
    "a 28 pt run is taller than the 12 pt line box"
  );
  assert.equal(
    core.estimateParagraphHeightPx(mixed, 480),
    mixedLayout.height,
    "the estimator follows the grown line"
  );

  const uniform = textRunsOnlyParagraph("a uniform paragraph stays uniform");
  const uniformLineHeightPx = core.estimateParagraphLineHeightPx(uniform);
  assert.equal(
    core.estimateParagraphHeightPx(uniform, 480),
    uniformLineHeightPx,
    "plain text keeps lineHeight * lineCount"
  );
  assert.equal(
    core.resolveParagraphLineMetrics(uniform, {
      availableWidthPx: 480,
      lineHeightPx: uniformLineHeightPx,
      expectedLineCount: 1,
    }),
    undefined,
    "no custom metrics means the paginator keeps the uniform model"
  );
});

test("a superscribed run records its baseline shift without growing the line", () => {
  const target = paragraph([
    { type: "text", text: "See note", style: TEXT_STYLE },
    { type: "text", text: "1", style: { ...TEXT_STYLE, verticalAlign: "superscript" } },
  ]);
  const lineHeightPx = core.estimateParagraphLineHeightPx(target);
  const layout = layoutParagraph(target, lineHeightPx);
  const line = layout.lines[0];

  assert.equal(layout.hasCustomVerticalMetrics, false, "Word keeps the line box");
  assert.equal(line.metrics.heightPx, lineHeightPx);
  const scriptFragment = line.fragments[1];
  assert.ok(scriptFragment, "the superscript fragment is laid out");
  assert.equal(
    scriptFragment.baselineShiftPx,
    core.scriptBaselineShiftPx("superscript", 16),
    "the superscript baseline sits above the line baseline"
  );
  assert.ok(
    core.fragmentBaselineYPx(line, scriptFragment) < core.lineBaselineYPx(line),
    "the raised glyph box rises instead of resizing the line box"
  );
});

// --- pagination split -----------------------------------------------------

function mixedSizeParagraph(prefixRun) {
  const words = Array.from({ length: 120 }, (_value, index) => `w${index}`).join(" ");
  return paragraph([
    ...(prefixRun
      ? [
          { type: "text", text: "BIG", style: { fontSizePt: 28 } },
          { type: "text", text: " ", style: TEXT_STYLE },
        ]
      : []),
    { type: "text", text: words, style: TEXT_STYLE },
  ]);
}

function pageConsumedLineHeightsPx(layout, model) {
  return layout.pages.map((page) => {
    let consumedHeightPx = 0;
    for (const block of page.blocks) {
      const range = block.paragraphLineRange;
      if (!range) {
        continue;
      }
      const lineMetrics = core.resolveParagraphLineMetrics(
        model.nodes[block.nodeIndex],
        {
          availableWidthPx: page.geometry.contentBoxPx.width,
          lineHeightPx: range.lineHeightPx,
          expectedLineCount: range.totalLineCount,
        }
      );
      consumedHeightPx += lineMetrics
        ? core.sumLineHeightsPx(
            lineMetrics,
            range.startLineIndex,
            range.endLineIndex
          )
        : (range.endLineIndex - range.startLineIndex) * range.lineHeightPx;
    }
    return consumedHeightPx;
  });
}

const SPLIT_PAGE_OPTIONS = {
  pageWidthPx: 400,
  pageHeightPx: 300,
  marginPx: 72,
  allowParagraphLineSplitting: true,
};

test("pagination splits a paragraph on its aggregated line heights, not on lineCount * lineHeight", () => {
  const bigModel = {
    metadata: { numberingDefinitions: {} },
    nodes: [mixedSizeParagraph(true)],
  };
  const plainModel = {
    metadata: { numberingDefinitions: {} },
    nodes: [mixedSizeParagraph(false)],
  };
  const bigParagraph = bigModel.nodes[0];
  const plainParagraph = plainModel.nodes[0];
  const contentWidthPx = 256;

  assert.equal(
    core.paragraphLineCountWithinWidth(bigParagraph, contentWidthPx),
    core.paragraphLineCountWithinWidth(plainParagraph, contentWidthPx),
    "both paragraphs wrap into the same number of lines"
  );

  const bigPages = core.paginateDocumentLayout(bigModel, SPLIT_PAGE_OPTIONS).pages;
  const plainPages = core.paginateDocumentLayout(plainModel, SPLIT_PAGE_OPTIONS)
    .pages;
  const firstBigEnd = bigPages[0].blocks[0].paragraphLineRange.endLineIndex;
  const firstPlainEnd =
    plainPages[0].blocks[0].paragraphLineRange.endLineIndex;

  assert.ok(
    firstBigEnd < firstPlainEnd,
    `the 28 pt first line must cost the page lines (got ${firstBigEnd} vs ${firstPlainEnd})`
  );
  assert.equal(
    bigPages[0].blocks[0].paragraphLineRange.startLineIndex,
    0,
    "the tall first line stays on the first page"
  );

  const lineMetrics = core.resolveParagraphLineMetrics(bigParagraph, {
    availableWidthPx: contentWidthPx,
    expectedLineCount: core.paragraphLineCountWithinWidth(
      bigParagraph,
      contentWidthPx
    ),
  });
  assert.ok(lineMetrics, "the paginator resolves per-line metrics here");
  assert.ok(
    lineMetrics[0].heightPx > lineMetrics[1].heightPx,
    "only the line with the 28 pt run grows"
  );
  assert.equal(
    core.estimateParagraphHeightPx(bigParagraph, contentWidthPx),
    core.sumLineHeightsPx(lineMetrics, 0, lineMetrics.length),
    "the paragraph height estimate equals the sum of its line boxes"
  );
});

test("no page overflows its content box once line boxes are aggregated", () => {
  const model = {
    metadata: { numberingDefinitions: {} },
    nodes: [mixedSizeParagraph(true)],
  };
  const layout = core.paginateDocumentLayout(model, SPLIT_PAGE_OPTIONS);
  const consumedByPage = pageConsumedLineHeightsPx(layout, model);

  layout.pages.forEach((page, pageIndex) => {
    assert.ok(
      consumedByPage[pageIndex] <= page.geometry.contentBoxPx.height + 2,
      `page ${pageIndex} consumed ${consumedByPage[pageIndex]}px of ${page.geometry.contentBoxPx.height}px`
    );
  });
});

// --- selection geometry ---------------------------------------------------

test("selection geometry follows the per-line box instead of one fixed line height", () => {
  const target = acceptImageParagraph();
  const lineHeightPx = core.estimateParagraphLineHeightPx(target);
  const layout = layoutParagraph(target, lineHeightPx);
  const line = layout.lines[0];
  const caret = core.resolveCaretRectAtOffset(layout, 0);

  assert.ok(caret, "expected a caret rect");
  assert.equal(caret.top, line.y);
  assert.equal(caret.height, line.metrics.heightPx);

  const sliced = core.sliceLayoutToLineRange(layout, 0, 1);
  assert.equal(sliced.height, line.metrics.heightPx);
  assert.equal(sliced.lineMetrics[0].heightPx, line.metrics.heightPx);
});
