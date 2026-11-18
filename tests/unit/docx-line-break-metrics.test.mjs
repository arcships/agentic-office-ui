// DOCX line-break agreement against captured font metrics (issue #123
// reopened).
//
// The fixture stores HarfBuzz-shaped advances for a Word-compatible serif
// (Nimbus Roman), Noto Sans CJK JP and Noto Naskh Arabic at 16 px. HarfBuzz is
// the OpenType shaper used by LibreOffice and Chromium, so the expected widths
// include the real kerning and `fi` ligature substitutions rather than a
// synthetic character table.
//
// The canvas below reads that fixture and honours `fontKerning` /
// `textRendering` exactly like the production measurement canvas. The layout
// must therefore reproduce the captured widths even though pretext's own
// canvas cannot carry a per-run feature string.

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { after, test } from "node:test";

const FIXTURE = JSON.parse(
  readFileSync(
    new URL("../fixtures/docx-font-metrics.harfbuzz.json", import.meta.url),
    "utf8"
  )
);

const originalOffscreenCanvas = globalThis.OffscreenCanvas;

function fixtureFontForCanvasFont(font) {
  const lower = String(font).toLowerCase();
  for (const [key, entry] of Object.entries(FIXTURE.fonts)) {
    if (lower.includes(entry.family.toLowerCase())) {
      return entry;
    }
  }
  return FIXTURE.fonts.TestLatin;
}

function measurementMode(context) {
  const kerning = context.fontKerning !== "none";
  const features = context.textRendering !== "optimizeSpeed";
  if (kerning && features) return "kernLiga";
  if (kerning) return "kern";
  if (features) return "liga";
  return "default";
}

function fallbackWidth(entry, text, mode) {
  const ligaturesEnabled = mode === "liga" || mode === "kernLiga";
  const kerningEnabled = mode === "kern" || mode === "kernLiga";
  if (ligaturesEnabled) {
    for (const [ligature, width] of Object.entries(entry.ligatures ?? {})) {
      if (text.includes(ligature)) {
        const [before, after = ""] = text.split(ligature);
        return (
          fallbackWidth(entry, before, mode) +
          width +
          fallbackWidth(entry, after, mode)
        );
      }
    }
  }

  let width = 0;
  let previous = undefined;
  for (const character of text) {
    width += entry.advances[character] ?? 0;
    if (kerningEnabled && previous !== undefined) {
      width += entry.kerning?.[previous + character] ?? 0;
    }
    previous = character;
  }
  return width;
}

function shapedWidth(entry, text, mode) {
  const exact = entry.stringWidths?.[text]?.[mode];
  return exact ?? fallbackWidth(entry, text, mode);
}

class FixtureCanvasContext {
  font = "12px sans-serif";
  fontKerning = "auto";
  textRendering = "auto";

  measureText(text) {
    const sizePx = Number.parseFloat(
      /(\d+(?:\.\d+)?)px/.exec(this.font)?.[1] ?? "16"
    );
    const entry = fixtureFontForCanvasFont(this.font);
    return {
      width: shapedWidth(entry, String(text), measurementMode(this)),
      fontBoundingBoxAscent: sizePx * 0.9,
      fontBoundingBoxDescent: sizePx * 0.25,
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
});

const core = await import(
  new URL("../../packages/docx-core/dist/index.js", import.meta.url).href
);

const LATIN_FAMILY = FIXTURE.fonts.TestLatin.family;
const CJK_FAMILY = FIXTURE.fonts.TestCJK.family;
const CS_FAMILY = FIXTURE.fonts.TestCS.family;

function styleForCase(fontKey) {
  const base = {
    fontFamily: LATIN_FAMILY,
    fontFamilyAscii: LATIN_FAMILY,
    fontFamilyHAnsi: LATIN_FAMILY,
    fontFamilyEastAsia: CJK_FAMILY,
    fontFamilyComplexScript: CS_FAMILY,
    sourceFontFamily: LATIN_FAMILY,
    fontSizePt: 12,
  };
  if (fontKey === "TestCJK") {
    return { ...base, sourceFontFamily: LATIN_FAMILY };
  }
  if (fontKey === "TestCS") {
    return {
      ...base,
      rightToLeft: true,
      direction: "rtl",
    };
  }
  return { ...base, fontFamilyEastAsia: undefined, fontFamilyComplexScript: undefined };
}

function styleForMode(fontKey, mode) {
  const style = styleForCase(fontKey);
  if (mode === "kern" || mode === "kernLiga") {
    style.kerningMinFontSizePt = 12;
  }
  if (mode === "liga" || mode === "kernLiga") {
    style.ligatures = "standard";
    style.contextualAlternates = true;
  }
  return style;
}

function paragraph(text, style) {
  return { type: "paragraph", children: [{ type: "text", text, style }] };
}

function layoutLines(paragraphNode, widthPx) {
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
  return layout.lines.map((line) => {
    const text = line.fragments.map((fragment) => fragment.text).join("");
    const width = line.fragments.reduce(
      (total, fragment) => total + fragment.width,
      0
    );
    return { text, normalizedText: text.trimEnd(), width };
  });
}

test("captured metrics reproduce default, kern, liga and kern+liga advances", () => {
  for (const entry of Object.values(FIXTURE.fonts)) {
    for (const [mode, expected] of Object.entries(
      entry.stringWidths?.["AVAV fi fi"] ?? {}
    )) {
      if (entry === FIXTURE.fonts.TestCJK || entry === FIXTURE.fonts.TestCS) {
        continue;
      }
      const style = styleForMode("TestLatin", mode);
      assert.ok(
        Math.abs(core.measureTextWidthPx("AVAV fi fi", style, 16) - expected) <=
          1,
        `${entry.family} ${mode}: expected ${expected}`
      );
    }
  }
});

for (const fixtureCase of FIXTURE.cases) {
  test(`captured line breaks: ${fixtureCase.name}`, () => {
    const style = styleForMode(fixtureCase.font, fixtureCase.mode);
    const paragraphNode = paragraph(fixtureCase.text, style);

    const measured = core.measureTextWidthPx(
      fixtureCase.text,
      style,
      FIXTURE.source.sizePx
    );
    assert.ok(
      Math.abs(measured - fixtureCase.expectedWidthPx) <= 1,
      `${fixtureCase.name}: ${measured} vs ${fixtureCase.expectedWidthPx}`
    );

    const lines = layoutLines(paragraphNode, fixtureCase.widthPx);
    assert.deepEqual(
      lines.map((line) => line.normalizedText),
      fixtureCase.expectedLines.map((line) => line.text)
    );
    assert.equal(
      core.paragraphLineCountWithinWidth(paragraphNode, fixtureCase.widthPx),
      fixtureCase.expectedLines.length
    );

    lines.forEach((line, index) => {
      const expected = fixtureCase.expectedLines[index];
      const lineWidth = core.measureTextWidthPx(
        line.normalizedText,
        style,
        FIXTURE.source.sizePx
      );
      assert.ok(
        Math.abs(lineWidth - expected.widthPx) <= 1,
        `${fixtureCase.name} line ${index}: ${lineWidth} vs ${expected.widthPx}`
      );
      assert.ok(
        line.width <= fixtureCase.widthPx + 1,
        `${fixtureCase.name} line ${index} overflows the container`
      );
    });
  });
}

test("kerning and ligatures change the line break, not just the recorded width", () => {
  const defaultCase = FIXTURE.cases.find(
    (fixtureCase) => fixtureCase.name === "latin-default"
  );
  const kernLigaCase = FIXTURE.cases.find(
    (fixtureCase) => fixtureCase.name === "latin-kernliga"
  );
  assert.ok(defaultCase);
  assert.ok(kernLigaCase);

  const defaultLines = layoutLines(
    paragraph(defaultCase.text, styleForMode("TestLatin", "default")),
    defaultCase.widthPx
  );
  const kernLigaLines = layoutLines(
    paragraph(kernLigaCase.text, styleForMode("TestLatin", "kernLiga")),
    kernLigaCase.widthPx
  );

  assert.equal(defaultLines.length, 2);
  assert.equal(kernLigaLines.length, 1);
  const defaultTotal = defaultLines.reduce((sum, line) => sum + line.width, 0);
  const kernLigaTotal = kernLigaLines.reduce((sum, line) => sum + line.width, 0);
  assert.ok(
    kernLigaTotal < defaultTotal - 4,
    "kerning/ligatures narrow the whole laid-out text"
  );
});
