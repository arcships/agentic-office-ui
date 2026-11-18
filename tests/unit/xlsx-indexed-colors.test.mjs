import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

// Regression coverage for issue #105: legacy OOXML indexed colours were parsed
// into `{ indexed }` but never resolved, so the 56-colour Excel palette (and a
// workbook's own `xl/styles.xml` `<indexedColors>` override) painted as the
// fallback black/transparent. The fixtures are the shared review-corpus files.

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const {
  DOMParser: XmlDomParser,
  Node: XmlNode,
  XMLSerializer: XmlSerializer,
} = requireFromXlsxCore("@xmldom/xmldom");
globalThis.DOMParser ??= XmlDomParser;
globalThis.Node ??= XmlNode;
globalThis.XMLSerializer ??= XmlSerializer;

const core = await import(
  pathToFileURL(requireFromXlsxCore.resolve("@arcships/xlsx-core")).href
);

function corpusBytes(name) {
  return new Uint8Array(readFileSync(new URL(`../../test-data/corpus/${name}`, import.meta.url)));
}

const standardBytes = corpusBytes("xlsx-indexed-colors.xlsx");
const customBytes = corpusBytes("xlsx-indexed-custom-palette.xlsx");

// ECMA-376 18.8.27, as listed by openpyxl's `COLOR_INDEX`.
const STANDARD_INDEXED = [
  "#000000", "#ffffff", "#ff0000", "#00ff00", "#0000ff", "#ffff00", "#ff00ff", "#00ffff",
  "#000000", "#ffffff", "#ff0000", "#00ff00", "#0000ff", "#ffff00", "#ff00ff", "#00ffff",
  "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#c0c0c0", "#808080",
  "#9999ff", "#993366", "#ffffcc", "#ccffff", "#660066", "#ff8080", "#0066cc", "#ccccff",
  "#000080", "#ff00ff", "#ffff00", "#00ffff", "#800080", "#800000", "#008080", "#0000ff",
  "#00ccff", "#ccffff", "#ccffcc", "#ffff99", "#99ccff", "#ff99cc", "#cc99ff", "#ffcc99",
  "#3366ff", "#33cccc", "#99cc00", "#ffcc00", "#ff9900", "#ff6600", "#666699", "#969696",
  "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333",
];

function fillColor(structure, styleId) {
  const style = structure.styleById[styleId];
  assert.ok(style?.fill, `style ${styleId} keeps its fill`);
  return core.resolveWorkbookFillStyle(style.fill, structure.themePalette).backgroundColor;
}

function fontColor(structure, styleId) {
  const style = structure.styleById[styleId];
  assert.ok(style?.font?.color, `style ${styleId} keeps its font colour`);
  return core.resolveWorkbookColor(style.font.color, structure.themePalette);
}

test("standard indexed colours resolve without a workbook palette", () => {
  STANDARD_INDEXED.forEach((expected, index) => {
    assert.equal(
      core.resolveWorkbookColor({ indexed: index }),
      expected,
      `indexed ${index} resolves to ${expected}`,
    );
  });

  // Excel reserves 64/65 for the system foreground/background.
  assert.equal(core.resolveWorkbookColor({ indexed: 64 }), "#000000");
  assert.equal(core.resolveWorkbookColor({ indexed: 65 }), "#ffffff");

  // Index 0 is black, so resolution must test for a colour rather than truthiness.
  assert.equal(core.resolveWorkbookColor({ indexed: 0 }), "#000000");

  // Values outside the palette and non-integer indexes stay unresolved.
  assert.equal(core.resolveWorkbookColor({ indexed: 66 }), null);
  assert.equal(core.resolveWorkbookColor({ indexed: -1 }), null);
  assert.equal(core.resolveWorkbookColor({ indexed: 10.5 }), null);
  assert.equal(core.resolveWorkbookColor({ indexed: "not-a-number" }), null);
});

test("workbook indexedColors override the standard palette", () => {
  const standard = core.parseWorkbookStructureAssets(standardBytes);
  assert.equal(standard.themePalette.indexedColors[10], "#ff0000");
  assert.equal(standard.themePalette.indexedColors[22], "#c0c0c0");
  assert.equal(standard.themePalette.indexedColors[64], "#000000");
  assert.equal(standard.themePalette.indexedColors[65], "#ffffff");
  assert.equal(
    Object.keys(standard.themePalette.indexedColors).length,
    66,
    "all 64 indexed colours plus the 2 system colours are present",
  );

  const custom = core.parseWorkbookStructureAssets(customBytes);
  assert.equal(custom.themePalette.indexedColors[8], "#ff00ff");
  assert.equal(custom.themePalette.indexedColors[10], "#112233");
  assert.equal(custom.themePalette.indexedColors[22], "#abcdef");
  assert.equal(
    custom.themePalette.indexedColors[9],
    "#ffffff",
    "entries the workbook does not change keep the standard value",
  );

  // A partial caller-supplied palette still falls back per index.
  assert.equal(
    core.resolveWorkbookColor(
      { indexed: 10 },
      { colorsByIndex: {}, indexedColors: { 10: "#123456" } },
    ),
    "#123456",
  );
  assert.equal(
    core.resolveWorkbookColor(
      { indexed: 22 },
      { colorsByIndex: {}, indexedColors: { 10: "#123456" } },
    ),
    "#c0c0c0",
  );
});

test("engine-shaped indexed colours keep the workbook palette", () => {
  const custom = core.parseWorkbookStructureAssets(customBytes);
  // `worksheet.getCellStyleAt()` hands the grid an already-resolved default
  // `hex` plus the palette index. The custom palette must still win.
  const engineColor = { colorType: "indexed", hex: "000000", paletteIndex: 8 };
  assert.equal(core.resolveWorkbookColor(engineColor, custom.themePalette), "#ff00ff");
  assert.equal(
    core.resolveWorkbookFillStyle({ color: engineColor, fillType: "solid" }, custom.themePalette)
      .backgroundColor,
    "#ff00ff",
  );
  assert.equal(
    core.resolveWorkbookColor(
      { colorType: "indexed", hex: "FF0000", paletteIndex: 10 },
      custom.themePalette,
    ),
    "#112233",
  );

  // Without an indexed index, a direct hex still resolves as before.
  assert.equal(core.resolveWorkbookColor({ colorType: "rgb", hex: "FF0000" }), "#ff0000");
});

test("indexed fills and fonts resolve to Excel's colours", () => {
  const standard = core.parseWorkbookStructureAssets(standardBytes);

  assert.deepEqual(standard.styleById[1].fill.color, { indexed: 10 });
  assert.equal(fillColor(standard, 1), "#ff0000");
  assert.equal(fillColor(standard, 2), "#c0c0c0");
  assert.equal(fillColor(standard, 3), "#969696");
  assert.equal(fillColor(standard, 4), "#333333");
  assert.equal(fillColor(standard, 5), "#ffffff");
  assert.equal(fillColor(standard, 6), "#000000");
  assert.equal(fillColor(standard, 7), "#ffffff");
  assert.equal(fillColor(standard, 8), "#000000");

  // Tint applies to indexed colours with the same HSL rule as theme colours.
  assert.equal(fillColor(standard, 9), "#ff8080");
  assert.equal(fillColor(standard, 10), "#800000");

  assert.deepEqual(standard.styleById[11].font.color, { indexed: 10 });
  assert.equal(fontColor(standard, 11), "#ff0000");
});

test("a custom indexed palette reaches fills, fonts and the viewer asset path", () => {
  const custom = core.parseWorkbookStructureAssets(customBytes);
  assert.equal(fillColor(custom, 1), "#ff00ff");
  assert.equal(fillColor(custom, 2), "#112233");
  assert.equal(fillColor(custom, 3), "#ffffff");
  assert.equal(fontColor(custom, 4), "#112233");

  // `parseWorkbookImageAssets()` is what the viewer/controller consumes.
  const assets = core.parseWorkbookImageAssets(customBytes);
  assert.equal(assets.themePalette.indexedColors[8], "#ff00ff");
  assert.equal(assets.themePalette.indexedColors[10], "#112233");
  assert.equal(
    core.resolveWorkbookFillStyle(assets.styleById[2].fill, assets.themePalette).backgroundColor,
    "#112233",
  );
  assert.equal(
    core.resolveWorkbookColor(assets.styleById[4].font.color, assets.themePalette),
    "#112233",
  );
});
