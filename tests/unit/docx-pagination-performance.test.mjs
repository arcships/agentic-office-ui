// DOCX repagination performance guard for the font-metric shim (issue #123
// reopened).
//
// Scroll performance is not cache reuse alone: the viewer repaginates after a
// resize, a font load or a measured-height pass. This guard builds a mixed
// Latin/CJK/RTL document, measures the cold (cache-retired) and warm
// repagination passes, and compares the styled document with a plain one. The
// budgets are deliberately generous so slow CI hosts do not flake; a real
// regression (e.g. measuring every run on every line) is far larger.

import assert from "node:assert/strict";
import { after, test } from "node:test";

const originalOffscreenCanvas = globalThis.OffscreenCanvas;

class PerformanceCanvasContext {
  font = "12px sans-serif";
  fontKerning = "auto";
  textRendering = "auto";

  measureText(text) {
    const fontSizePx = Number.parseFloat(
      /(\d+(?:\.\d+)?)px/.exec(this.font)?.[1] ?? "12"
    );
    const family = this.font.toLowerCase();
    let width = 0;
    for (const character of text) {
      if (/\s/.test(character)) {
        width += fontSizePx * 0.5;
      } else if (/[\u2e80-\u9fff\u3040-\u30ff\uac00-\ud7af]/.test(character)) {
        width += fontSizePx * (family.includes("cjk") ? 1 : 0.75);
      } else if (/[\u0590-\u08ff]/.test(character)) {
        width += fontSizePx * (family.includes("arabic") ? 0.5 : 0.6);
      } else if (/[A-Z0-9]/.test(character)) {
        width += fontSizePx * 0.625;
      } else {
        width += fontSizePx * 0.5;
      }
    }
    if (this.fontKerning !== "none") {
      width -= (text.match(/AV/g) ?? []).length * 1.2;
    }
    if (this.textRendering !== "optimizeSpeed") {
      width -= (text.match(/fi/g) ?? []).length * 2;
    }
    return {
      width,
      fontBoundingBoxAscent: fontSizePx * 0.9,
      fontBoundingBoxDescent: fontSizePx * 0.25,
    };
  }
}

globalThis.OffscreenCanvas = class {
  constructor(width, height) {
    this.width = width;
    this.height = height;
  }
  getContext() {
    return new PerformanceCanvasContext();
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

const PAGE_OPTIONS = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: 96 };
const PARAGRAPH_COUNT = 120;

const PLAIN_STYLE = {
  fontFamily: "PerfLatin",
  fontFamilyAscii: "PerfLatin",
  fontFamilyHAnsi: "PerfLatin",
  fontFamilyEastAsia: "PerfCJK",
  fontFamilyComplexScript: "PerfArabic",
  sourceFontFamily: "PerfLatin",
  fontSizePt: 12,
};

const STYLED_STYLE = {
  ...PLAIN_STYLE,
  kerningMinFontSizePt: 12,
  ligatures: "standard",
  contextualAlternates: true,
};

function paragraph(text, style) {
  return { type: "paragraph", children: [{ type: "text", text, style }] };
}

function modelWith(styles) {
  const nodes = [];
  for (let index = 0; index < PARAGRAPH_COUNT; index += 1) {
    const style = styles[index % styles.length];
    const text =
      index % 4 === 0
        ? `AVAV fi fi latin paragraph ${index}`
        : index % 4 === 1
        ? `中文段落 ${index} AV fi`
        : index % 4 === 2
        ? `مرحبا بالعالم ${index}`
        : `Plain body text for paragraph ${index}`;
    nodes.push(paragraph(text, style));
  }
  return { ...core.createBlankDocumentModel(), nodes };
}

function median(values) {
  const sorted = [...values].sort((left, right) => left - right);
  return sorted[Math.floor(sorted.length / 2)] ?? 0;
}

function measurePasses(model, passes) {
  const samples = [];
  for (let pass = 0; pass < passes; pass += 1) {
    core.invalidateFontMetrics();
    const startedAt = performance.now();
    const layout = core.paginateDocumentLayout(model, PAGE_OPTIONS);
    samples.push(performance.now() - startedAt);
    assert.ok(layout.pages.length > 0);
  }
  return samples;
}

test(
  "repagination stays within the scroll budget after the metric shim",
  { timeout: 20_000 },
  (t) => {
    const plainModel = modelWith([PLAIN_STYLE]);
    const styledModel = modelWith([PLAIN_STYLE, STYLED_STYLE]);

    // Warm the module-level caches and JIT, then measure fresh passes with the
    // metrics retired (the worst case after a web font loads).
    core.paginateDocumentLayout(plainModel, PAGE_OPTIONS);
    core.paginateDocumentLayout(styledModel, PAGE_OPTIONS);

    const plainSamples = measurePasses(plainModel, 5);
    const styledSamples = measurePasses(styledModel, 5);
    const plainMedian = median(plainSamples);
    const styledMedian = median(styledSamples);

    t.diagnostic(
      `cold repagination ${PARAGRAPH_COUNT} paragraphs: plain=${plainMedian.toFixed(
        1
      )}ms styled=${styledMedian.toFixed(1)}ms`
    );

    assert.ok(
      styledMedian <= 600,
      `styled repagination took ${styledMedian.toFixed(1)}ms (budget 600ms)`
    );
    assert.ok(
      styledMedian <= plainMedian * 4 + 75,
      `kerning/ligature metrics regressed pagination: styled=${styledMedian.toFixed(
        1
      )}ms plain=${plainMedian.toFixed(1)}ms`
    );

    // Warm scroll/repagination passes (height feedback on a live viewer) must
    // not re-measure every run; they should be no slower than the cold pass.
    const warmStartedAt = performance.now();
    for (let pass = 0; pass < 10; pass += 1) {
      core.paginateDocumentLayout(styledModel, PAGE_OPTIONS);
    }
    const warmPerPass = (performance.now() - warmStartedAt) / 10;
    t.diagnostic(`warm repagination per pass: ${warmPerPass.toFixed(1)}ms`);
    assert.ok(
      warmPerPass <= 250,
      `warm repagination took ${warmPerPass.toFixed(1)}ms per pass`
    );
    assert.ok(
      warmPerPass <= Math.max(styledMedian, 1) * 2,
      `warm repagination (${warmPerPass.toFixed(
        1
      )}ms) is slower than cold (${styledMedian.toFixed(1)}ms)`
    );
  }
);
