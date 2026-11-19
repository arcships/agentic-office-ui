// DOCX automatic paragraph spacing and the document default tab stop
// (issue #129).
//
// Two layout inputs were missing from the model:
//
//   * `w:spacing/@w:beforeAutospacing` and `@w:afterAutospacing` — Word's "auto"
//     paragraph spacing. Word honours the flags only in compatibility mode 12
//     (Word 2007) and 15 (Word 2013+), inserts a fixed 14 pt gap (measured from
//     a LibreOffice/Word render: 14.0 pt for 10 pt, 10.5 pt, 12 pt and 24 pt
//     text alike) and collapses the gap between adjacent items of the same list.
//   * `w:defaultTabStop` — the fallback tab grid. Word's default is 720 twips
//     (48 px at 96 dpi); CJK templates commonly ship 420 twips (28 px).
//
// The assertions pin the model fields, the resolution rules, the measurement
// (including the cache keys that separate documents), and the tab geometry
// against the numbers a Word/LibreOffice render produces.

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

const WORD_NS =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// 14 pt at 96 dpi: 280 twips / 15 twips-per-pixel.
const AUTO_SPACING_PX = 280 / 15;

function settingsXml({ compatibilityMode, defaultTabStopTwips } = {}) {
  const compat =
    compatibilityMode === undefined
      ? ""
      : `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${compatibilityMode}"/></w:compat>`;
  const tabStop =
    defaultTabStopTwips === undefined
      ? ""
      : `<w:defaultTabStop w:val="${defaultTabStopTwips}"/>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="${WORD_NS}">${tabStop}${compat}</w:settings>`;
}

function paragraph({
  text = "Paragraph",
  spacing = 'w:before="0" w:after="0"',
  numbering,
  sizeHalfPoints = 21,
} = {}) {
  const pPrChildren = [
    numbering
      ? `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numbering}"/></w:numPr>`
      : "",
    `<w:spacing ${spacing}/>`,
  ].join("");
  return `<w:p><w:pPr>${pPrChildren}</w:pPr><w:r><w:rPr><w:sz w:val="${sizeHalfPoints}"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

function documentXml(paragraphs) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${WORD_NS}"><w:body>${paragraphs.join("")}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`;
}

async function importModel(paragraphs, settings) {
  let pkg = core.createMinimalDocxPackage(documentXml(paragraphs));
  if (settings) {
    pkg = core.withPart(pkg, {
      name: "word/settings.xml",
      content: settingsXml(settings),
    });
  }
  return core.buildDocModel(pkg);
}

function paragraphNode({ spacing, numbering, text = "Paragraph" } = {}) {
  return {
    type: "paragraph",
    children: [{ type: "text", text }],
    ...(spacing ? { style: { spacing } } : {}),
    ...(numbering ? { style: { spacing, numbering } } : {}),
  };
}

function docModel(nodes, metadata = {}) {
  return {
    nodes,
    metadata: {
      sourceParts: 0,
      warnings: [],
      headerSections: [],
      footerSections: [],
      paragraphStyles: [],
      ...metadata,
    },
  };
}

// ── settings.xml parsing ───────────────────────────────────────────

test("w:defaultTabStop is parsed from settings.xml", () => {
  assert.equal(
    core.parseDefaultTabStopTwips(settingsXml({ defaultTabStopTwips: 420 })),
    420,
  );
  assert.equal(
    core.parseDefaultTabStopTwips(settingsXml({ defaultTabStopTwips: 720 })),
    720,
  );
  assert.equal(core.parseDefaultTabStopTwips(settingsXml()), undefined);
  assert.equal(core.parseDefaultTabStopTwips(undefined), undefined);
  // A zero/negative value is not a grid; Word's default applies instead.
  assert.equal(
    core.parseDefaultTabStopTwips('<w:defaultTabStop w:val="0"/>'),
    undefined,
  );
});

test("w:compatSetting compatibilityMode is parsed and other settings are ignored", () => {
  assert.equal(core.parseCompatibilityMode(settingsXml({ compatibilityMode: 15 })), 15);
  assert.equal(core.parseCompatibilityMode(settingsXml({ compatibilityMode: 12 })), 12);
  assert.equal(core.parseCompatibilityMode(settingsXml()), undefined);
  assert.equal(core.parseCompatibilityMode(undefined), undefined);

  const mixed = `<w:settings xmlns:w="${WORD_NS}"><w:compat>
    <w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>
  </w:compat></w:settings>`;
  assert.equal(core.parseCompatibilityMode(mixed), 14);
});

test("w:spacing autospacing flags are parsed with OOXML on/off values", () => {
  assert.deepEqual(
    core.parseParagraphAutoSpacing(
      '<w:pPr><w:spacing w:beforeAutospacing="1" w:afterAutospacing="true"/></w:pPr>',
    ),
    { beforeAutoSpacing: true, afterAutoSpacing: true },
  );
  assert.deepEqual(
    core.parseParagraphAutoSpacing(
      '<w:pPr><w:spacing w:beforeAutospacing="on" w:afterAutospacing="0"/></w:pPr>',
    ),
    { beforeAutoSpacing: true, afterAutoSpacing: false },
  );
  assert.deepEqual(
    core.parseParagraphAutoSpacing(
      '<w:pPr><w:spacing w:beforeAutospacing="1"/></w:pPr>',
    ),
    { beforeAutoSpacing: true, afterAutoSpacing: undefined },
  );
  // Explicit numeric spacing only, an empty pPr, or no pPr at all: nothing to add.
  assert.equal(
    core.parseParagraphAutoSpacing('<w:pPr><w:spacing w:before="240"/></w:pPr>'),
    undefined,
  );
  assert.equal(core.parseParagraphAutoSpacing("<w:pPr/>"), undefined);
  assert.equal(core.parseParagraphAutoSpacing(undefined), undefined);
});

// ── Import ─────────────────────────────────────────────────────────

test("import puts auto spacing, compatibility mode and the default tab stop on the model", async () => {
  const model = await importModel(
    [paragraph({ spacing: 'w:beforeAutospacing="1" w:afterAutospacing="1"' })],
    { compatibilityMode: 15, defaultTabStopTwips: 420 },
  );

  assert.equal(model.metadata.compatibility?.compatibilityMode, 15);
  assert.equal(model.metadata.defaultTabStopTwips, 420);
  const spacing = model.nodes[0].style?.spacing;
  assert.equal(spacing?.beforeAutoSpacing, true);
  assert.equal(spacing?.afterAutoSpacing, true);
  // The auto flags name no value, so they never become explicit w:before/after.
  assert.equal(spacing?.beforeTwips, undefined);
  assert.equal(spacing?.afterTwips, undefined);
});

test("import leaves a document without settings.xml on Word's defaults", async () => {
  const model = await importModel([
    paragraph({ spacing: 'w:beforeAutospacing="1"' }),
  ]);
  assert.equal(model.metadata.compatibility?.compatibilityMode, undefined);
  assert.equal(model.metadata.defaultTabStopTwips, undefined);
  assert.equal(model.nodes[0].style?.spacing?.beforeAutoSpacing, true);
});

// ── Auto spacing rules ─────────────────────────────────────────────

test("auto spacing is honoured only in compatibility mode 12 and 15", () => {
  const spacing = { beforeAutoSpacing: true, afterAutoSpacing: true };
  for (const mode of [12, 15]) {
    assert.deepEqual(
      core.resolveParagraphAutoSpacingPx(paragraphNode({ spacing }), {
        compatibilityMode: mode,
      }),
      { beforePx: AUTO_SPACING_PX, afterPx: AUTO_SPACING_PX },
      `mode ${mode}`,
    );
  }
  for (const mode of [undefined, 11, 13, 14]) {
    assert.deepEqual(
      core.resolveParagraphAutoSpacingPx(paragraphNode({ spacing }), {
        compatibilityMode: mode,
      }),
      { beforePx: 0, afterPx: 0 },
      `mode ${String(mode)}`,
    );
  }
});

test("the automatic gap is Word's 14 pt, independent of the font size", () => {
  assert.equal(core.HTML_AUTO_SPACING_TWIPS, 280);
  assert.equal(core.HTML_AUTO_SPACING_PX, AUTO_SPACING_PX);
  assert.ok(Math.abs(core.HTML_AUTO_SPACING_PX - 18.6667) < 0.001);

  // The same 14 pt gap for 10 pt and 24 pt text (Word/LibreOffice render).
  const measure = (spacing) =>
    core.estimateParagraphHeightPx(
      paragraphNode({ spacing, text: "Measure me" }),
      400,
      undefined,
      undefined,
      false,
      { compatibilityMode: 15 },
    );
  const plain = measure({ lineTwips: 240, lineRule: "auto" });
  const auto = measure({
    lineTwips: 240,
    lineRule: "auto",
    beforeAutoSpacing: true,
    afterAutoSpacing: true,
  });
  assert.ok(
    Math.abs(auto - plain - 2 * AUTO_SPACING_PX) < 0.01,
    `expected a 2 x 14pt gap, got ${auto - plain}`,
  );
});

test("adjacent items of the same list collapse the automatic gap", () => {
  const first = paragraphNode({
    spacing: { beforeAutoSpacing: true, afterAutoSpacing: true },
    numbering: { numId: 3, ilvl: 0 },
  });
  const second = paragraphNode({
    spacing: { beforeAutoSpacing: true, afterAutoSpacing: true },
    numbering: { numId: 3, ilvl: 0 },
  });
  const otherList = paragraphNode({
    spacing: { beforeAutoSpacing: true, afterAutoSpacing: true },
    numbering: { numId: 4, ilvl: 0 },
  });
  const body = paragraphNode({
    spacing: { beforeAutoSpacing: true, afterAutoSpacing: true },
  });

  assert.equal(core.paragraphsShareListNumbering(first, second), true);
  assert.equal(core.paragraphsShareListNumbering(first, otherList), false);
  assert.equal(core.paragraphsShareListNumbering(first, body), false);

  const defaults = { compatibilityMode: 15 };
  assert.deepEqual(core.resolveParagraphAutoSpacingInStory([first, second], 1, second, defaults), {
    beforePx: 0,
    afterPx: AUTO_SPACING_PX,
  });
  assert.deepEqual(core.resolveParagraphAutoSpacingInStory([first, second], 0, first, defaults), {
    beforePx: AUTO_SPACING_PX,
    afterPx: 0,
  });
  assert.deepEqual(
    core.resolveParagraphAutoSpacingInStory([first, otherList], 1, otherList, defaults),
    { beforePx: AUTO_SPACING_PX, afterPx: AUTO_SPACING_PX },
  );
  assert.deepEqual(
    core.resolveParagraphAutoSpacingInStory([body, body], 1, body, defaults),
    { beforePx: AUTO_SPACING_PX, afterPx: AUTO_SPACING_PX },
  );
});

test("pagination spacing helpers add the automatic gap to explicit spacing", () => {
  const paragraph = paragraphNode({
    spacing: {
      beforeTwips: 240,
      afterTwips: 120,
      beforeAutoSpacing: true,
      afterAutoSpacing: true,
    },
  });
  const defaults = { compatibilityMode: 15 };

  assert.equal(
    core.paragraphBeforeSpacingPx(paragraph, defaults),
    240 / 15 + AUTO_SPACING_PX,
  );
  assert.equal(
    core.paragraphAfterSpacingPx(paragraph, defaults),
    120 / 15 + AUTO_SPACING_PX,
  );
  // Without a compatibility mode the flags are ignored, as in Word 2010.
  assert.equal(core.paragraphBeforeSpacingPx(paragraph), 240 / 15);
  assert.equal(core.paragraphAfterSpacingPx(paragraph), 120 / 15);
});

test("the flow drops only the automatic half of the gap between list items", () => {
  const first = paragraphNode({
    spacing: { beforeTwips: 240, afterTwips: 60, afterAutoSpacing: true },
    numbering: { numId: 7, ilvl: 0 },
  });
  const second = paragraphNode({
    spacing: { beforeTwips: 120, beforeAutoSpacing: true },
    numbering: { numId: 7, ilvl: 0 },
  });
  const model = docModel([first, second], {
    compatibility: { compatibilityMode: 15 },
  });

  // Explicit w:after (4 px) survives; the automatic 14 pt gap is collapsed.
  assert.equal(core.effectiveParagraphAfterSpacingPx(model, 0, first), 4);
  // Explicit w:before (8 px) survives; the automatic before is collapsed.
  assert.equal(
    core.effectiveParagraphBeforeSpacingPx(model, 1, second, 10, false),
    8,
  );
});

// ── Measurement and cache keys ──────────────────────────────────────

test("a compatibility-15 CJK paragraph gets Word's automatic rhythm", async () => {
  const paragraphs = [
    paragraph({ text: "第一段", sizeHalfPoints: 21 }),
    paragraph({
      text: "第二段",
      sizeHalfPoints: 21,
      spacing: 'w:beforeAutospacing="1" w:afterAutospacing="1"',
    }),
    paragraph({ text: "第三段", sizeHalfPoints: 21 }),
  ];
  const compatible = await importModel(paragraphs, {
    compatibilityMode: 15,
    defaultTabStopTwips: 420,
  });
  const legacy = await importModel(paragraphs, {
    compatibilityMode: 14,
    defaultTabStopTwips: 420,
  });

  const width = 468;
  const heightsOf = (model) =>
    model.nodes.map((node) =>
      core.estimateParagraphHeightPx(
        node,
        width,
        undefined,
        undefined,
        false,
        core.resolveDocumentLayoutDefaults(model.metadata),
      ),
    );
  const compatibleHeights = heightsOf(compatible);
  const legacyHeights = heightsOf(legacy);

  // Mode 14 ignores the flags: every paragraph keeps the plain line height.
  assert.equal(legacyHeights[0], legacyHeights[1]);
  assert.equal(legacyHeights[1], legacyHeights[2]);
  // Mode 15 adds Word's 14 pt gap before and after the middle paragraph only.
  assert.equal(compatibleHeights[0], legacyHeights[0]);
  assert.equal(compatibleHeights[2], legacyHeights[2]);
  assert.ok(
    Math.abs(compatibleHeights[1] - legacyHeights[1] - 2 * AUTO_SPACING_PX) <
      0.01,
    `expected two 14pt gaps, got ${compatibleHeights[1] - legacyHeights[1]}`,
  );
  // Baseline rhythm: in the compatibility-15 document the first and second
  // paragraph baselines sit two gaps further apart than in the old one.
  assert.ok(
    Math.abs(
      compatibleHeights[0] +
        compatibleHeights[1] -
        (legacyHeights[0] + legacyHeights[1]) -
        2 * AUTO_SPACING_PX,
    ) < 0.01,
  );
});

test("automatic paragraph spacing changes the page rhythm end to end", () => {
  const autoParagraph = (index) => ({
    type: "paragraph",
    children: [{ type: "text", text: `Paragraph ${index}` }],
    style: {
      spacing: {
        beforeAutoSpacing: true,
        afterAutoSpacing: true,
        lineTwips: 240,
        lineRule: "auto",
      },
    },
  });
  const model = (compatibilityMode) =>
    docModel(
      Array.from({ length: 30 }, (_, index) => autoParagraph(index)),
      {
        compatibility: { compatibilityMode },
        defaultTabStopTwips: 420,
      },
    );
  const options = {
    pageWidthPx: 612,
    pageHeightPx: 300,
    marginPx: 36,
    allowParagraphLineSplitting: false,
  };

  const compatible = core.paginateDocumentLayout(model(15), options);
  const legacy = core.paginateDocumentLayout(model(14), options);

  // Word 2013 rhythm needs roughly one page per paragraph here; Word 2010
  // packs them because it ignores the flags.
  assert.ok(
    compatible.pages.length > legacy.pages.length,
    `compatibility 15 should paginate taller (${compatible.pages.length} vs ${legacy.pages.length})`,
  );
  const firstPageHeights = (layout) =>
    layout.pages[0].blocks.map((block) => block.framePx.height);
  const compatibleFirstHeight = firstPageHeights(compatible)[0];
  const legacyFirstHeight = firstPageHeights(legacy)[0];
  // The page-top paragraph drops its before spacing, so the consumed height
  // gains one gap (both gaps minus the dropped before) — and rounding of the
  // raw estimate keeps it within a pixel of that.
  const gained = compatibleFirstHeight - legacyFirstHeight;
  assert.ok(
    gained > AUTO_SPACING_PX - 1 && gained <= AUTO_SPACING_PX + 1,
    `the page-top paragraph should gain one 14pt gap (got ${gained})`,
  );
});

test("the default tab stop is part of the measurement cache key", async () => {
  // Same paragraph XML, two documents with different tab grids: the cached
  // estimate of the first must not be reused for the second.
  const tabbed = paragraph({ text: "AAAA\tBB" });
  const wide = await importModel([tabbed], {
    compatibilityMode: 15,
    defaultTabStopTwips: 720,
  });
  const narrow = await importModel([tabbed], {
    compatibilityMode: 15,
    defaultTabStopTwips: 420,
  });
  assert.equal(wide.nodes[0].sourceXml, narrow.nodes[0].sourceXml);

  // At 70 px the 48 px grid still keeps "AAAA\tBB" on one line, while the
  // 28 px grid wraps it: Word/LibreOffice differ here too.
  const width = 70;
  const wideHeight = core.estimateParagraphHeightPx(
    wide.nodes[0],
    width,
    undefined,
    undefined,
    false,
    core.resolveDocumentLayoutDefaults(wide.metadata),
  );
  const narrowHeight = core.estimateParagraphHeightPx(
    narrow.nodes[0],
    width,
    undefined,
    undefined,
    false,
    core.resolveDocumentLayoutDefaults(narrow.metadata),
  );

  assert.notEqual(
    wideHeight,
    narrowHeight,
    "a 28 px grid wraps tabbed text differently from a 48 px grid",
  );
  // Re-measuring the first document still returns its own value.
  assert.equal(
    core.estimateParagraphHeightPx(
      wide.nodes[0],
      width,
      undefined,
      undefined,
      false,
      core.resolveDocumentLayoutDefaults(wide.metadata),
    ),
    wideHeight,
  );
});

// ── Default tab stop geometry ──────────────────────────────────────

test("resolveDefaultTabStopPx maps twips to Word's pixel grid", () => {
  assert.equal(core.resolveDefaultTabStopPx(420), 28);
  assert.equal(core.resolveDefaultTabStopPx(720), core.DEFAULT_TAB_STOP_PX);
  assert.equal(core.DEFAULT_TAB_STOP_PX, 48);
  assert.equal(core.resolveDefaultTabStopPx(undefined), 48);
  assert.equal(core.resolveDefaultTabStopPx(0), 48);
  assert.equal(core.resolveDefaultTabStopPx(-5), 48);
});

test("tabs advance on the document's default grid", () => {
  // Word places a tab at the next multiple of the default stop from the line
  // origin; 420 twips is 21 pt = 28 px.
  assert.equal(core.resolveNextTabStopPx(0, [], 28), 28);
  assert.equal(core.resolveNextTabStopPx(20, [], 28), 28);
  assert.equal(core.resolveNextTabStopPx(28, [], 28), 56);
  assert.equal(core.resolveNextTabStopPx(29, [], 28), 56);
  // Default grid stays 48 px when the document names none.
  assert.equal(core.resolveNextTabStopPx(20, []), 48);
  // Explicit stops still win.
  assert.equal(core.resolveNextTabStopPx(20, [40, 80], 28), 40);
  assert.equal(core.resolveNextTabStopPx(50, [40, 80], 28), 80);
});

test("a tabbed run is laid out at the document's default stop", () => {
  const paragraph = paragraphNode({
    text: "AA\tZZ",
    spacing: { lineTwips: 240, lineRule: "auto" },
  });
  const measureTabWidth = (defaultTabStopPx) => {
    const source = core.buildParagraphPretextLayoutSource(paragraph, {
      expandTabsForLayout: true,
      defaultTabStopPx,
    });
    const tabRun = source.runs.find((run) => run.kind === "tab");
    return tabRun?.tabWidthPx;
  };
  const narrow = measureTabWidth(28);
  const wide = measureTabWidth(48);

  // Word/LibreOffice: the tab lands 28 px (420 twips) or 48 px (720 twips)
  // from the line origin, so the gap is the stop minus the text already set —
  // the two grids ask for gaps exactly 20 px apart.
  assert.equal(wide - narrow, 48 - 28);
  assert.ok(narrow > 0 && narrow < 28, `narrow: ${narrow}`);
  assert.ok(wide > 0 && wide < 48, `wide: ${wide}`);
});

test("a wrapped left tab re-aligns to the next default stop", () => {
  // Centre/right anchored stops keep their absolute zone layout.
  assert.equal(core.allowsWrappedLeftTabRefit(undefined), true);
  assert.equal(core.allowsWrappedLeftTabRefit([{ alignment: "left" }]), true);
  assert.equal(core.allowsWrappedLeftTabRefit([{ alignment: "center" }]), false);
  assert.equal(core.allowsWrappedLeftTabRefit([{ alignment: "right" }]), false);
  assert.equal(core.allowsWrappedLeftTabRefit([{ alignment: "decimal" }]), false);
  assert.equal(
    core.allowsWrappedLeftTabRefit([{ alignment: "left" }, { alignment: "right" }]),
    false,
  );

  // The tab stop sits beyond the 120 px line, so the tab wraps. On the new line
  // a left tab re-aligns to the first default stop (28 px) before the following
  // word, which then no longer fits; a decimal stop keeps the old behaviour of
  // dropping the tab.
  const lineWidthPx = 120;
  const defaults = { defaultTabStopTwips: 420 };
  const paragraphWith = (alignment) => ({
    type: "paragraph",
    children: [
      { type: "text", text: "AAA" },
      { type: "text", text: "\t" },
      { type: "text", text: "MMMMMMMMMM" },
    ],
    style: { tabStops: [{ alignment, positionTwips: (lineWidthPx + 40) * 15 }] },
  });

  const refitLineCount = core.paragraphLineCountWithinWidth(
    paragraphWith("left"),
    lineWidthPx,
    undefined,
    defaults,
  );
  const collapsedLineCount = core.paragraphLineCountWithinWidth(
    paragraphWith("decimal"),
    lineWidthPx,
    undefined,
    defaults,
  );
  assert.equal(collapsedLineCount, 2);
  assert.equal(refitLineCount, 3);
});
