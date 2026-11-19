import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";

import ts from "../../packages/xlsx-core/node_modules/typescript/lib/typescript.js";

const ROOT = path.resolve(import.meta.dirname, "../..");
const GRID_DOM_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/grid-dom.ts");
const RICH_TEXT_LAYOUT_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/rich-text-layout.ts");

function dataUrl(source) {
  const output = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.ESNext,
      target: ts.ScriptTarget.ES2022,
    },
  }).outputText;
  return `data:text/javascript;base64,${Buffer.from(output).toString("base64")}`;
}

// `grid-dom` shares the superscript/subscript scale with the rich-text layout
// helper. A data-URL module cannot resolve a relative specifier, so the import
// is pointed at the transpiled helper (the only runtime import in the module).
const richTextUrl = dataUrl(readFileSync(RICH_TEXT_LAYOUT_PATH, "utf8"));
const gridDomUrl = dataUrl(
  readFileSync(GRID_DOM_PATH, "utf8").replace(
    /from "\.\/rich-text-layout"/,
    `from ${JSON.stringify(richTextUrl)}`,
  ),
);
const dom = await import(gridDomUrl);

const BASE_STYLE = {
  alignment: "left",
  backgroundColor: "#ffffff",
  color: "#18181b",
  fontFamily: "Calibri",
  fontSizePx: 16,
  fontStyle: "normal",
  fontWeight: "400",
  rotation: null,
  shrinkToFit: false,
  strike: false,
  underline: false,
  vertical: "bottom",
  wrap: false,
};

const BORDER_RECORD = {
  bottom: { color: { hex: "FFFF0000" }, style: "mediumDashed" },
  top: { color: { hex: "FF0000FF" }, style: "hair" },
};

test("renderMode normalizes to the canvas default and only accepts an explicit dom", () => {
  assert.equal(dom.resolveGridRenderMode("dom"), "dom");
  assert.equal(dom.resolveGridRenderMode("canvas"), "canvas");
  assert.equal(dom.resolveGridRenderMode(undefined), "canvas");
  assert.equal(dom.resolveGridRenderMode(null), "canvas");
  assert.equal(dom.resolveGridRenderMode("DOM"), "canvas");
  assert.equal(dom.DEFAULT_GRID_RENDER_MODE, "canvas");
});

test("workbook border edges map to the same CSS widths and dash styles the canvas draws", () => {
  assert.equal(dom.canvasBorderEdgeCss(undefined, null), undefined);
  assert.equal(dom.canvasBorderEdgeCss({ style: "none" }, null), undefined);
  assert.equal(dom.canvasBorderEdgeCss({ style: "thin" }, "#112233"), "1px solid #112233");
  assert.equal(dom.canvasBorderEdgeCss({ style: "dashed" }, "#112233"), "1px dashed #112233");
  assert.equal(dom.canvasBorderEdgeCss({ style: "mediumDashed" }, "#112233"), "2px dashed #112233");
  assert.equal(dom.canvasBorderEdgeCss({ style: "dashDot" }, "#112233"), "1px dashed #112233");
  assert.equal(dom.canvasBorderEdgeCss({ style: "dotted" }, "#112233"), "1px dotted #112233");
  assert.equal(dom.canvasBorderEdgeCss({ style: "hair" }, "#112233"), "1px dotted #112233");
  assert.equal(dom.canvasBorderEdgeCss({ style: "thick" }, "#112233"), "3px solid #112233");
  assert.equal(dom.canvasBorderEdgeCss({ style: "double" }, "#112233"), "3px double #112233");
  assert.equal(dom.canvasBorderEdgeCss({ style: "medium" }, null), "2px solid #52525b");
});

test("a resolved canvas cell style becomes the same DOM background, font and alignment", () => {
  const style = {
    ...BASE_STYLE,
    alignment: "right",
    backgroundColor: "#fff2cc",
    color: "#c00000",
    fontStyle: "italic",
    fontWeight: "700",
    strike: true,
    underline: true,
    vertical: "middle",
    wrap: true,
  };
  const css = dom.canvasCellStyleToCss(style, {
    borders: {
      bottom: "2px dashed #ff0000",
      left: undefined,
      right: undefined,
      top: "1px dotted #0000ff",
    },
    gridColor: "#d9d9d9",
    gridLines: true,
  });

  assert.equal(css.backgroundColor, "#fff2cc");
  assert.equal(css.color, "#c00000");
  assert.equal(css.fontWeight, "700");
  assert.equal(css.fontStyle, "italic");
  assert.equal(css.fontFamily, "\"Calibri\", -apple-system, BlinkMacSystemFont, \"Segoe UI\", sans-serif");
  assert.equal(css.fontSize, "16px");
  assert.equal(css.lineHeight, "20px");
  assert.equal(css.display, "flex");
  assert.equal(css.justifyContent, "flex-end");
  assert.equal(css.alignItems, "center");
  assert.equal(css.textAlign, "right");
  assert.equal(css.boxSizing, "border-box");
  assert.equal(css.padding, "1px 4px");
  assert.equal(css.borderBottom, "2px dashed #ff0000");
  assert.equal(css.borderTop, "1px dotted #0000ff");
  assert.equal(css.boxShadow, "inset -1px 0 0 0 #d9d9d9, inset 0 -1px 0 0 #d9d9d9");
  assert.match(css.textDecoration, /underline/);
  assert.match(css.textDecoration, /line-through/);

  const left = dom.canvasCellStyleToCss({ ...BASE_STYLE, alignment: "center", vertical: "top" });
  assert.equal(left.justifyContent, "center");
  assert.equal(left.alignItems, "flex-start");
  assert.equal(left.boxShadow, undefined, "gridlines are off unless requested");

  const bottom = dom.canvasCellStyleToCss({ ...BASE_STYLE, alignment: "left", vertical: "bottom" });
  assert.equal(bottom.justifyContent, "flex-start");
  assert.equal(bottom.alignItems, "flex-end");
});

test("DOM cell style options carry conditional fills, hyperlinks, shrink and merge gridlines", () => {
  const scale = dom.canvasCellStyleToCss(BASE_STYLE, { backgroundColor: "#dbeafe" });
  assert.equal(scale.backgroundColor, "#dbeafe");

  const link = dom.canvasCellStyleToCss({ ...BASE_STYLE, color: "#18181b" }, { hyperlink: true });
  assert.equal(link.color, "#2563eb");
  assert.match(link.textDecoration, /underline/);
  const customLink = dom.canvasCellStyleToCss(BASE_STYLE, { hyperlink: true, hyperlinkColor: "#0f766e" });
  assert.equal(customLink.color, "#0f766e");

  const shrunk = dom.canvasCellStyleToCss({ ...BASE_STYLE, shrinkToFit: true }, { fontSizePx: 9 });
  assert.equal(shrunk.fontSize, "9px");
  assert.equal(shrunk.lineHeight, "12px", "the line height never drops below the canvas minimum");

  const merged = dom.canvasCellStyleToCss(BASE_STYLE, { gridColor: "#d9d9d9", gridLines: true, merged: true });
  assert.equal(merged.boxShadow, "inset 0 0 0 1px #d9d9d9", "a merge paints its own outline");
});

test("DOM cell text keeps the canvas clipping and rotation rules and stays selectable", () => {
  const flat = dom.canvasCellTextStyle(BASE_STYLE);
  assert.equal(flat.whiteSpace, "nowrap");
  assert.equal(flat.textOverflow, "ellipsis");
  assert.equal(flat.overflow, "hidden");
  assert.equal(flat.userSelect, "text");
  assert.equal(flat.fontSize, "16px");
  assert.equal(flat.maxWidth, "100%");
  assert.equal(flat.position, "relative");
  assert.equal(flat.zIndex, 1);

  const wrapped = dom.canvasCellTextStyle({ ...BASE_STYLE, wrap: true }, { fontSizePx: 10 });
  assert.equal(wrapped.whiteSpace, "pre-wrap");
  assert.equal(wrapped.textOverflow, "clip");
  assert.equal(wrapped.overflowWrap, "break-word");
  assert.equal(wrapped.fontSize, "10px");
  assert.equal(wrapped.lineHeight, "12.5px");

  const icon = dom.canvasCellTextStyle(BASE_STYLE, { maxWidthPx: 42 });
  assert.equal(icon.maxWidth, "42px");
  assert.equal(icon.marginRight, 5, "the text stops short of the icon column");

  const rotated = dom.canvasCellTextStyle({ ...BASE_STYLE, rotation: -45 });
  assert.equal(rotated.position, "absolute");
  assert.equal(rotated.left, "50%");
  assert.equal(rotated.top, "50%");
  assert.equal(rotated.width, "max-content");
  assert.equal(rotated.maxWidth, "none");
  assert.equal(rotated.whiteSpace, "nowrap");
  assert.equal(rotated.transform, "translate(-50%, -50%) rotate(-45deg)");
});

test("rich-text runs keep their own font, colour and baseline inside the DOM cell", () => {
  const plain = dom.richTextRunCss(BASE_STYLE, { bold: true, color: undefined }, { color: "#123456" });
  assert.equal(plain.fontWeight, "700");
  assert.equal(plain.color, "#123456");
  assert.equal(plain.fontSize, "16px");

  const decorated = dom.richTextRunCss(BASE_STYLE, {
    italic: true,
    name: "Aptos",
    strikethrough: true,
    underline: true,
  });
  assert.equal(decorated.fontStyle, "italic");
  assert.equal(decorated.fontFamily, "Aptos");
  assert.match(decorated.textDecoration, /underline/);
  assert.match(decorated.textDecoration, /line-through/);

  const superscript = dom.richTextRunCss(BASE_STYLE, { verticalAlign: "superscript" }, { fontSizePx: 20 });
  assert.equal(superscript.verticalAlign, "super");
  assert.equal(superscript.fontSize, "14px", "superscript runs shrink to 0.7×");
  assert.equal(superscript.lineHeight, "0.7");

  const subscript = dom.richTextRunCss(BASE_STYLE, { verticalAlign: "subscript" });
  assert.equal(subscript.verticalAlign, "sub");
  assert.equal(subscript.fontSize, "11px", "16px × 0.7 rounds like the canvas font string");
});

test("the DOM axis window pins the frozen slots and never renders a cell twice", () => {
  assert.deepEqual(dom.resolveGridDomAxisWindow({ end: 5, start: 0 }, 2, 10), {
    frozen: [0, 1],
    scrolling: [2, 3, 4, 5],
  });
  assert.deepEqual(dom.resolveGridDomAxisWindow({ end: 5, start: 0 }, 0, 10), {
    frozen: [],
    scrolling: [0, 1, 2, 3, 4, 5],
  });
  assert.deepEqual(dom.resolveGridDomAxisWindow({ end: 5, start: 0 }, 3, 3), {
    frozen: [0, 1, 2],
    scrolling: [],
  });
  assert.deepEqual(dom.resolveGridDomAxisWindow({ end: 30, start: 20 }, 2, 10), {
    frozen: [0, 1],
    scrolling: [],
  });
  assert.deepEqual(dom.resolveGridDomAxisWindow({ end: -1, start: 0 }, 0, 0), {
    frozen: [],
    scrolling: [],
  });
  assert.deepEqual(dom.resolveGridDomAxisWindow({ end: 4, start: 2 }, 99, 6), {
    frozen: [0, 1, 2, 3, 4, 5],
    scrolling: [],
  });
});

test("selection outlines are clipped to their pane and never allocate an unbounded box", () => {
  const inside = dom.resolveGridDomOutlineStyle(
    { bottom: 48, left: 0, right: 80, top: 0 },
    { bottom: 600, left: 0, right: 800, top: 0 },
    { color: "#2563eb" },
  );
  assert.deepEqual(inside, {
    border: "2px solid #2563eb",
    boxSizing: "border-box",
    height: "46px",
    left: "1px",
    pointerEvents: "none",
    position: "absolute",
    top: "1px",
    width: "78px",
    zIndex: 2,
  });

  const dashed = dom.resolveGridDomOutlineStyle(
    { bottom: 30, left: 10, right: 20, top: 20 },
    { bottom: 600, left: 0, right: 800, top: 0 },
    { color: "#71717a", dashed: true, inset: 0.5, lineWidth: 1.5, zIndex: 4 },
  );
  assert.equal(dashed.border, "1.5px dashed #71717a");
  assert.equal(dashed.zIndex, 4);

  assert.equal(
    dom.resolveGridDomOutlineStyle(
      { bottom: 5000, left: 0, right: 80, top: 4900 },
      { bottom: 100, left: 0, right: 800, top: 0 },
      { color: "#2563eb" },
    ),
    null,
    "a range far above the pane draws nothing",
  );
  // Within the clamp margin the box is still returned so its real border stays
  // put; the pane's overflow clipping is what hides it.
  const nearby = dom.resolveGridDomOutlineStyle(
    { bottom: 200, left: 0, right: 80, top: 100 },
    { bottom: 100, left: 0, right: 800, top: 0 },
    { color: "#2563eb" },
  );
  assert.equal(nearby.top, "101px");
  assert.equal(nearby.height, "98px");
  assert.equal(
    dom.resolveGridDomOutlineStyle(
      { bottom: 10, left: 40, right: 20, top: 0 },
      { bottom: 600, left: 0, right: 800, top: 0 },
      { color: "#2563eb" },
    ),
    null,
    "an inverted rect draws nothing",
  );

  // A whole-row selection is hundreds of thousands of pixels tall; the box is
  // clamped far outside the pane, which still clips it invisibly.
  const huge = dom.resolveGridDomOutlineStyle(
    { bottom: 24_000_000, left: 0, right: 80, top: 0 },
    { bottom: 600, left: 0, right: 800, top: 0 },
    { color: "#2563eb" },
  );
  assert.ok(huge, "the visible edge stays rendered");
  assert.equal(huge.height, "4599px");
  assert.equal(huge.top, "1px");
});

test("the selection tint matches the canvas inset fill", () => {
  assert.deepEqual(dom.gridDomSelectionStyle("rgba(37, 99, 235, 0.08)"), {
    background: "rgba(37, 99, 235, 0.08)",
    inset: "1px",
    pointerEvents: "none",
    position: "absolute",
    zIndex: 0,
  });
});
