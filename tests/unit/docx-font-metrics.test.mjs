// DOCX font metrics, kerning, script-slot segmentation and layout agreement
// (issue #123).
//
// The fixture canvas below models the pieces of Canvas text measurement the
// subsystem relies on:
//   - per-family advances (Latin, CJK, complex script);
//   - `fontKerning` (the `kern` feature);
//   - `textRendering` as the offscreen fallback for feature settings;
//   - font-box ascent/descent for the line-metric estimator.
//
// That lets the assertions pin Word-like line-break positions and prove that
// the estimate and the pretext render use the same metrics.

import assert from "node:assert/strict";
import { after, test } from "node:test";

const originalOffscreenCanvas = globalThis.OffscreenCanvas;
const originalDocument = globalThis.document;
const originalFontFace = globalThis.FontFace;

let widthScale = 1;
let measureCalls = 0;
let lastFont = "";

function isCjk(character) {
  return /[\u2e80-\u9fff\u3040-\u30ff\uac00-\ud7af]/.test(character);
}

function isComplexScript(character) {
  return /[\u0590-\u08ff\ufb1d-\ufdff\ufe70-\ufeff]/.test(character);
}

class FixtureCanvasContext {
  font = "12px sans-serif";
  fontKerning = "auto";
  textRendering = "auto";

  measureText(text) {
    measureCalls += 1;
    lastFont = this.font;
    const fontSizePx = Number.parseFloat(
      /(\d+(?:\.\d+)?)px/.exec(this.font)?.[1] ?? "12"
    );
    const family = this.font.toLowerCase();
    const bold = /(?:^|\s)700(?=\s)/.test(this.font);
    let width = 0;
    for (const character of text) {
      if (/\s/.test(character)) {
        width += fontSizePx * 0.5;
      } else if (isCjk(character)) {
        width += fontSizePx * (family.includes("testcjk") ? 1 : 0.75);
      } else if (isComplexScript(character)) {
        width += fontSizePx * (family.includes("testcs") ? 0.5 : 0.6);
      } else if (/[A-Z0-9]/.test(character)) {
        width += fontSizePx * 0.625;
      } else {
        width += fontSizePx * 0.5;
      }
    }
    if (bold) width *= 1.05;
    if (this.fontKerning !== "none") {
      const pairCount = (text.match(/AV/g) ?? []).length;
      width -= pairCount * 1.2;
    }
    if (this.textRendering !== "optimizeSpeed" && /fi/.test(text)) {
      width -= 2;
    }
    return {
      width: width * widthScale,
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
    return new FixtureCanvasContext();
  }
};

after(() => {
  if (originalOffscreenCanvas === undefined) {
    delete globalThis.OffscreenCanvas;
  } else {
    globalThis.OffscreenCanvas = originalOffscreenCanvas;
  }
  if (originalDocument === undefined) {
    delete globalThis.document;
  } else {
    globalThis.document = originalDocument;
  }
  if (originalFontFace === undefined) {
    delete globalThis.FontFace;
  } else {
    globalThis.FontFace = originalFontFace;
  }
});

const core = await import(
  new URL("../../packages/docx-core/dist/index.js", import.meta.url).href
);

const TEXT_STYLE = {
  fontFamily: "TestLatin",
  fontFamilyAscii: "TestLatin",
  fontFamilyHAnsi: "TestLatin",
  fontFamilyEastAsia: "TestCJK",
  fontFamilyComplexScript: "TestCS",
  sourceFontFamily: "TestLatin",
  fontSizePt: 12,
  kerningMinFontSizePt: 12,
};

function paragraph(text, style = TEXT_STYLE) {
  return {
    type: "paragraph",
    children: [{ type: "text", text, style }],
  };
}

function blankModel(nodes) {
  return {
    ...core.createBlankDocumentModel(),
    nodes,
  };
}

// ── Script slots ─────────────────────────────────────────────────────

test("script slots split Latin, CJK and RTL text at UTF-16 offsets", () => {
  const text = "Hello 中文 مرحبا!";
  const segments = core.segmentTextByDocxScriptFont(text, TEXT_STYLE);

  assert.equal(segments.map((segment) => segment.text).join(""), text);
  assert.deepEqual(
    segments.map((segment) => [segment.text, segment.script, segment.fontFamily]),
    [
      ["Hello ", "ascii", "TestLatin"],
      ["中文", "eastAsia", "TestCJK"],
      [" ", "ascii", "TestLatin"],
      ["مرحبا", "complexScript", "TestCS"],
      ["!", "ascii", "TestLatin"],
    ]
  );
  for (let index = 0; index < segments.length; index += 1) {
    assert.equal(segments[index].startOffset, index === 0 ? 0 : segments[index - 1].endOffset);
    assert.equal(
      text.slice(segments[index].startOffset, segments[index].endOffset),
      segments[index].text
    );
  }
});

test("an explicit RTL/complex-script run forces the cs slot", () => {
  const style = {
    ...TEXT_STYLE,
    rightToLeft: true,
    direction: "rtl",
  };
  const segments = core.segmentTextByDocxScriptFont("abc 123", style);
  assert.equal(segments.length, 1);
  assert.equal(segments[0].script, "complexScript");
  assert.equal(segments[0].fontFamily, "TestCS");
  assert.equal(core.resolveDocxTextFontFamily("abc", style), "TestCS");
});

test("neutral non-ASCII characters inherit the previous strong script", () => {
  const segments = core.segmentTextByDocxScriptFont("中\u200d文", TEXT_STYLE);
  assert.deepEqual(
    segments.map((segment) => [segment.text, segment.script]),
    [["中\u200d文", "eastAsia"]]
  );
});

test("an edited legacy family overrides the structured slots", () => {
  const edited = {
    ...TEXT_STYLE,
    fontFamily: "EditedFamily",
    sourceFontFamily: "TestLatin",
  };
  const [segment] = core.segmentTextByDocxScriptFont("中文", edited);
  assert.equal(segment.fontFamily, "EditedFamily");
});

test("the complex-script size is used for cs segments", () => {
  const style = { ...TEXT_STYLE, fontSizeCsPt: 18 };
  const [segment] = core.segmentTextByDocxScriptFont("مرحبا", style);
  assert.equal(segment.script, "complexScript");
  assert.equal(segment.fontSizePt, 18);
  assert.equal(core.resolveDocxTextFontSizePt("مرحبا", style), 18);
});

// ── Kerning / OpenType features ─────────────────────────────────────

test("kerning follows the OOXML threshold in half-points", () => {
  assert.equal(core.resolveDocxFontKerning({}, 12), "none");
  assert.equal(
    core.resolveDocxFontKerning({ kerningMinFontSizePt: 12 }, 12),
    "normal"
  );
  assert.equal(
    core.resolveDocxFontKerning({ kerningMinFontSizePt: 12 }, 11.99),
    "none"
  );
  assert.equal(core.resolveDocxFontKerning({ kerningMinFontSizePt: 12 }), undefined);
});

test("ligature and contextual-alternate settings map to CSS features", () => {
  assert.equal(
    core.resolveDocxFontFeatureSettings({}),
    core.DISABLED_DOCX_FONT_FEATURE_SETTINGS
  );
  assert.equal(
    core.resolveDocxFontFeatureSettings({ ligatures: "standard" }),
    '"liga" 1, "clig" 0, "dlig" 0, "hlig" 0, "calt" 0'
  );
  assert.equal(
    core.resolveDocxFontFeatureSettings({ ligatures: "all", contextualAlternates: true }),
    '"liga" 1, "clig" 1, "dlig" 1, "hlig" 1, "calt" 1'
  );
  assert.equal(
    core.resolveDocxFontFeatureSettings(undefined, { enableOpenTypeFeatures: false }),
    core.DISABLED_DOCX_FONT_FEATURE_SETTINGS
  );
  assert.equal(core.measuredFontSizePt("italic 700 18px Test"), 13.5);
});

test("OOXML kerning, ligature, hint, language and cs-size properties parse", () => {
  const parsed = core.parseRunStyleFromRunXml(
    "<w:r><w:rPr>" +
      '<w:rFonts w:ascii="TestLatin" w:hAnsi="TestLatin" w:eastAsia="TestCJK" w:cs="TestCS" w:hint="eastAsia"/>' +
      '<w:kern w:val="24"/><w14:ligatures w14:val="standard"/><w14:cntxtAlts/>' +
      '<w:sz w:val="24"/><w:szCs w:val="32"/>' +
      '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>' +
      "<w:cs/>" +
      "</w:rPr><w:t>x</w:t></w:r>"
  );

  assert.equal(parsed.kerningMinFontSizePt, 12);
  assert.equal(parsed.ligatures, "standard");
  assert.equal(parsed.contextualAlternates, true);
  assert.equal(parsed.fontHint, "eastAsia");
  assert.equal(parsed.fontSizePt, 12);
  assert.equal(parsed.fontSizeCsPt, 16);
  assert.equal(parsed.language, "en-US");
  assert.equal(parsed.languageEastAsia, "ja-JP");
  assert.equal(parsed.languageBidi, "ar-SA");
  assert.equal(parsed.complexScript, true);

  const css = core.runStyleToCss(parsed, "light", "中文");
  assert.equal(css.fontKerning, "normal");
  assert.match(css.fontFeatureSettings, /"liga" 1/);
  assert.match(css.fontFamily, /TestCJK/);
});

// ── Advance widths ──────────────────────────────────────────────────

test("advance widths are keyed by family, weight, style, kerning and features", () => {
  core.invalidateFontMetrics();
  const regular = core.fontAdvanceWidthPx("400 16px TestLatin, serif", "AVAV", {
    fontKerning: "normal",
    fontFeatureSettings: core.DISABLED_DOCX_FONT_FEATURE_SETTINGS,
  });
  const bold = core.fontAdvanceWidthPx("700 16px TestLatin, serif", "AVAV", {
    fontKerning: "normal",
    fontFeatureSettings: core.DISABLED_DOCX_FONT_FEATURE_SETTINGS,
  });
  assert.ok(bold > regular, "bold advances use the bold face metrics");

  const unkerned = core.fontAdvanceWidthPx("400 16px TestLatin, serif", "AVAV", {
    fontKerning: "none",
    fontFeatureSettings: core.DISABLED_DOCX_FONT_FEATURE_SETTINGS,
  });
  assert.ok(unkerned > regular, "kerning tightens the AV pairs");

  const ligatures = core.fontAdvanceWidthPx("400 16px TestLatin, serif", "fi", {
    fontKerning: "none",
    fontFeatureSettings: '"liga" 1, "clig" 0, "dlig" 0, "hlig" 0, "calt" 0',
  });
  const disabled = core.fontAdvanceWidthPx("400 16px TestLatin, serif", "fi", {
    fontKerning: "none",
    fontFeatureSettings: core.DISABLED_DOCX_FONT_FEATURE_SETTINGS,
  });
  assert.ok(ligatures < disabled, "enabled ligatures are narrower");
});

test("measureTextWidthPx resolves kerning, script slots and the complex-script face", () => {
  const latin = core.measureTextWidthPx("AV", TEXT_STYLE, 16);
  const latinNoKerning = core.measureTextWidthPx(
    "AV",
    { ...TEXT_STYLE, kerningMinFontSizePt: undefined },
    16
  );
  assert.ok(Math.abs(latinNoKerning - latin - 1.2) < 1e-9);

  const cjk = core.measureTextWidthPx("中", TEXT_STYLE, 16);
  const latinOnlyCjk = core.measureTextWidthPx(
    "中",
    { ...TEXT_STYLE, fontFamilyEastAsia: undefined, sourceFontFamily: "TestLatin" },
    16
  );
  assert.equal(cjk, 16);
  assert.ok(latinOnlyCjk < cjk, "the CJK slot supplies the wider glyph");

  const rtl = core.measureTextWidthPx("مرحبا", TEXT_STYLE, 16);
  const latinRtl = core.measureTextWidthPx(
    "مرحبا",
    { ...TEXT_STYLE, fontFamilyComplexScript: undefined, rightToLeft: false, direction: undefined },
    16
  );
  assert.equal(rtl, 40);
  assert.ok(latinRtl > rtl, "the cs slot supplies the complex-script face");
});

test("the measurement cache and the public fontMetrics facade invalidate together", () => {
  core.invalidateFontMetrics();
  const beforeRevision = core.getFontMetricsRevision();
  const first = core.measureTextWidthPx("CacheProbe", TEXT_STYLE, 16);
  const callsAfterFirst = measureCalls;
  const second = core.measureTextWidthPx("CacheProbe", TEXT_STYLE, 16);
  assert.equal(second, first);
  assert.equal(measureCalls, callsAfterFirst, "repeat measurements hit the cache");

  const revision = core.invalidateFontMetrics();
  assert.ok(revision > beforeRevision);
  assert.ok(core.fontMetrics.revision() >= revision);

  widthScale = 1.25;
  const afterInvalidate = core.measureTextWidthPx("CacheProbe", TEXT_STYLE, 16);
  widthScale = 1;
  assert.ok(afterInvalidate > first, "invalidation retires the cached advance");
});

test("fontMetrics objects expose the parsed family, weight and style", () => {
  const metrics = core.resolveFontMetrics("italic 700 18px TestLatin, serif", {
    fontKerning: "normal",
  });
  assert.equal(metrics.family, "TestLatin, serif");
  assert.equal(metrics.weight, 700);
  assert.equal(metrics.style, "italic");
  assert.equal(metrics.sizePx, 18);
  assert.equal(metrics.kerning, "normal");
  assert.ok(metrics.advancePx("AV") > 0);
});

// ── DOM / headless fallbacks ────────────────────────────────────────

test("a document-only host measures with the DOM span fallback", () => {
  const appended = [];
  const fakeDocument = {
    body: {
      appendChild(element) {
        appended.push(element);
      },
      removeChild() {},
    },
    createElement(tag) {
      if (tag === "canvas") {
        return { getContext: () => null };
      }
      return {
        textContent: "",
        style: {},
        setAttribute() {},
        getBoundingClientRect() {
          return { width: this.textContent.length * 3 };
        },
      };
    },
  };
  delete globalThis.OffscreenCanvas;
  globalThis.document = fakeDocument;

  try {
    core.invalidateFontMetrics();
    const width = core.measureTextWidthPx("DOMFALLBACK", TEXT_STYLE, 16);
    assert.ok(width > 0);
    assert.equal(appended.length, 1, "the fallback appends one hidden span");
    assert.equal(core.measureTextWidthWithDom("16px TestLatin, serif", "abcd"), 12);
  } finally {
    globalThis.document = originalDocument;
    globalThis.OffscreenCanvas = class {
      constructor(width, height) {
        this.width = width;
        this.height = height;
      }
      getContext() {
        return new FixtureCanvasContext();
      }
    };
  }
});

test("a headless host falls back to deterministic glyph advances", () => {
  delete globalThis.OffscreenCanvas;
  delete globalThis.document;
  core.invalidateFontMetrics();
  try {
    const latin = core.fontAdvanceWidthPx("16px TestLatin, serif", "a");
    const cjk = core.fontAdvanceWidthPx("16px TestCJK, serif", "中");
    assert.ok(latin > 0);
    assert.ok(cjk > latin, "CJK advances use the ideographic width");
  } finally {
    globalThis.OffscreenCanvas = class {
      constructor(width, height) {
        this.width = width;
        this.height = height;
      }
      getContext() {
        return new FixtureCanvasContext();
      }
    };
    if (originalDocument !== undefined) globalThis.document = originalDocument;
  }
});

test("document.fonts.ready retires metrics when the face set settles", async () => {
  const listeners = new Map();
  const fakeDocument = {
    fonts: {
      addEventListener(type, listener) {
        listeners.set(type, listener);
      },
      removeEventListener(type) {
        listeners.delete(type);
      },
      ready: Promise.resolve(),
      add() {},
      delete() {},
      load: async () => [],
    },
  };
  globalThis.document = fakeDocument;
  core.invalidateFontMetrics();
  const before = core.getFontMetricsRevision();
  try {
    assert.equal(core.ensureDocumentFontInvalidation(), true);
    await Promise.resolve();
    await Promise.resolve();
    assert.ok(
      core.getFontMetricsRevision() > before,
      "resolving document.fonts.ready bumps the metric revision"
    );
    listeners.get("loadingdone")?.();
    assert.ok(core.getFontMetricsRevision() > before);
  } finally {
    globalThis.document = originalDocument;
  }
});

// ── Layout agreement and pagination ────────────────────────────────

function renderedLineWidths(paragraphNode, widthPx) {
  const source = core.buildParagraphPretextLayoutSource(paragraphNode);
  assert.ok(source);
  const layout = core.layoutParagraphPretextSource(
    paragraphNode,
    source,
    widthPx,
    20,
    []
  );
  assert.ok(layout);
  return {
    lineCount: layout.lineCount,
    widths: layout.lines.map((line) =>
      line.fragments.reduce(
        (total, fragment) => total + fragment.width,
        0
      )
    ),
    texts: layout.lines.map((line) =>
      line.fragments.map((fragment) => fragment.text).join("")
    ),
  };
}

test("Latin line breaks agree between the estimator and the render layout at a 1px edge", () => {
  const latinParagraph = paragraph("AVAV AVAV AVAV");
  assert.equal(core.paragraphLineCountWithinWidth(latinParagraph, 50), 3);

  const narrow = renderedLineWidths(latinParagraph, 50);
  assert.equal(narrow.lineCount, 3);
  for (const lineText of narrow.texts) {
    const measured = core.measureTextWidthPx(lineText, TEXT_STYLE, 16);
    assert.ok(measured <= 50 + 1, `line "${lineText}" fits within 1px`);
  }

  // Two "AVAV" words plus the space are 83.2px wide: 83px does not fit,
  // 84px does — the acceptance case's 1px boundary.
  assert.equal(core.paragraphLineCountWithinWidth(latinParagraph, 83), 3);
  assert.equal(core.paragraphLineCountWithinWidth(latinParagraph, 84), 2);
  assert.equal(renderedLineWidths(latinParagraph, 84).lineCount, 2);
});

test("CJK line breaks match between estimator and render layout", () => {
  const cjkParagraph = paragraph("中文中文中文");
  assert.equal(core.paragraphLineCountWithinWidth(cjkParagraph, 40), 3);
  const layout = renderedLineWidths(cjkParagraph, 40);
  assert.equal(layout.lineCount, 3);
  for (const lineText of layout.texts) {
    assert.ok(core.measureTextWidthPx(lineText, TEXT_STYLE, 16) <= 40 + 1);
  }
});

test("RTL line breaks match between estimator and render layout", () => {
  const rtlStyle = {
    ...TEXT_STYLE,
    rightToLeft: true,
    direction: "rtl",
  };
  const rtlParagraph = paragraph("مرحبا مرحبا", rtlStyle);
  const estimated = core.paragraphLineCountWithinWidth(rtlParagraph, 48);
  const layout = renderedLineWidths(rtlParagraph, 48);
  assert.equal(layout.lineCount, estimated);
  for (const lineText of layout.texts) {
    assert.ok(core.measureTextWidthPx(lineText, rtlStyle, 16) <= 48 + 1);
  }
});

test("layout-engine line height uses the platform-neutral font metrics", () => {
  const pages = core.layoutDocument(blankModel([paragraph("Metrics")]), {
    pageWidth: 400,
    pageHeight: 400,
    margin: 10,
    minLineHeight: 1,
    paragraphSpacing: 0,
    tableCellPadding: 0,
  });
  const block = pages[0]?.blocks[0];
  assert.equal(block?.kind, "paragraph");
  // 12 pt = 16 px; 0.9 em ascent + 0.25 em descent = 18.4 px.
  assert.equal(block.height, 18);
});

test("page count is stable when metrics are invalidated at any zoom profile", () => {
  const model = blankModel(
    Array.from({ length: 40 }, (_, index) =>
      paragraph(`AVAV AVAV paragraph ${index}`)
    )
  );
  const options = {
    pageWidthPx: 816,
    pageHeightPx: 1056,
    marginPx: 96,
  };
  const baseline = core.paginateDocumentLayout(model, options).pages.length;
  assert.ok(baseline > 0);

  core.invalidateFontMetrics();
  const afterInvalidate = core.paginateDocumentLayout(model, options).pages.length;
  assert.equal(afterInvalidate, baseline);

  // Font metrics do not depend on the zoom transform: the same page geometry
  // must produce the same segmentation after a profile invalidation.
  core.invalidateFontMetrics();
  assert.equal(
    core.paginateDocumentLayout(model, options).pages.length,
    baseline
  );
});

test("repagination reuses cached advances (scroll hot path)", () => {
  const model = blankModel(
    Array.from({ length: 12 }, (_, index) =>
      paragraph(`AVAV AVAV cache probe ${index}`)
    )
  );
  const options = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: 96 };
  core.invalidateFontMetrics();
  core.paginateDocumentLayout(model, options);
  const callsAfterFirstPass = measureCalls;
  core.paginateDocumentLayout(model, options);
  assert.equal(
    measureCalls,
    callsAfterFirstPass,
    "a second pagination pass must read cached advances"
  );
});

// ── Font table + document fonts ────────────────────────────────────
const FONT_TABLE_XML =
  "<w:fonts>" +
  '<w:font w:name="Arial"/>' +
  '<w:font w:name="Calibri">' +
  '<w:embedRegular r:id="rId1" w:fontKey="{00000000-0000-0000-0000-000000000000}"/>' +
  "</w:font>" +
  "</w:fonts>";

test("fontTableEntries does not let a self-closing font borrow the next entry", () => {
  const entries = core.fontTableEntries(FONT_TABLE_XML);
  assert.deepEqual(
    entries.map((entry) => entry.family),
    ["Calibri"]
  );
  assert.match(entries[0].xml, /embedRegular/);
});

test("embedded font descriptors come from the font-table entries", () => {
  const pkg = {
    parts: new Map([
      ["word/fontTable.xml", { name: "word/fontTable.xml", content: FONT_TABLE_XML }],
      [
        "word/_rels/fontTable.xml.rels",
        {
          name: "word/_rels/fontTable.xml.rels",
          content:
            '<Relationships><Relationship Id="rId1" Target="fonts/font1.odttf"/></Relationships>',
        },
      ],
    ]),
    binaryAssets: new Map([
      ["word/fonts/font1.odttf", new Uint8Array([1, 2, 3, 4])],
    ]),
  };
  const descriptors = core.collectEmbeddedFontFaceDescriptors(pkg);
  assert.equal(descriptors.length, 1);
  assert.equal(descriptors[0].family, "Calibri");
  assert.equal(descriptors[0].weight, "400");
});

test("documentFontRequests covers body, CJK, RTL, numbering, notes and headers", () => {
  const model = blankModel([
    {
      type: "paragraph",
      style: { numbering: { numId: 1, ilvl: 0 } },
      children: [{ type: "text", text: "Body 中文 مرحبا", style: TEXT_STYLE }],
    },
  ]);
  model.metadata.paragraphStyles = [
    { id: "Normal", name: "Normal", isDefault: true, runStyle: TEXT_STYLE },
  ];
  model.metadata.numberingDefinitions = {
    abstracts: [
      {
        abstractNumId: 1,
        levels: [{ ilvl: 0, text: "%1.", runStyle: TEXT_STYLE }],
      },
    ],
    instances: [{ numId: 1, abstractNumId: 1 }],
  };
  model.metadata.footnotes = [{ id: 1, text: "Note مرحبا" }];
  model.metadata.comments = [{ id: 1, text: "Comment 中文" }];
  model.metadata.headerSections = [
    {
      partName: "word/header1.xml",
      nodes: [{ type: "paragraph", children: [{ type: "text", text: "Header" }] }],
    },
  ];

  const requests = core.documentFontRequests(model, ["FancyTableFont"]);
  const families = requests.map((request) => request.family.toLowerCase());
  assert.ok(families.some((family) => family.includes("testlatin")));
  assert.ok(families.some((family) => family.includes("testcjk")));
  assert.ok(families.some((family) => family.includes("testcs")));
  assert.ok(families.some((family) => family.includes("fancytablefont")));
  const header = requests.find((request) =>
    request.characters.has("H") && request.characters.has("d")
  );
  assert.ok(header, "body and header text are collected");
});

test("awaitDocumentFonts preloads families, embedded faces and waits for ready", async () => {
  const loaded = [];
  const added = [];
  const fakeDocument = {
    fonts: {
      add(face) {
        added.push(face);
      },
      delete() {},
      async load(font, text) {
        loaded.push([font, text]);
        return [];
      },
      ready: Promise.resolve(),
    },
  };
  globalThis.FontFace = class {
    constructor(family, source, descriptors) {
      this.family = family;
      this.source = source;
      this.descriptors = descriptors;
    }
    async load() {
      return this;
    }
  };

  const model = blankModel([
    {
      type: "paragraph",
      children: [{ type: "text", text: "Hello 中文", style: TEXT_STYLE }],
    },
  ]);
  const result = await core.awaitDocumentFonts(model, {
    document: fakeDocument,
    extraFamilies: ["FancyTableFont"],
    embeddedFonts: [
      {
        family: "EmbeddedFace",
        style: "normal",
        weight: "400",
        source: new ArrayBuffer(4),
      },
    ],
  });

  assert.equal(result.faces.length, 1);
  assert.equal(result.faces[0].family, "EmbeddedFace");
  assert.equal(added.length, 1);
  assert.ok(
    loaded.some(([font, chars]) => font.includes("FancyTableFont") && chars.includes("M"))
  );
  assert.ok(
    loaded.some(([font, chars]) => font.includes("TestCJK") && chars.includes("中")),
    "CJK characters are preloaded for the east-Asian face"
  );
});
