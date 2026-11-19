import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import test from "node:test";
import { pathToFileURL } from "node:url";

import ts from "../../packages/xlsx-core/node_modules/typescript/lib/typescript.js";

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const { strFromU8, strToU8, unzipSync, zipSync } = await import(
  pathToFileURL(requireFromXlsxCore.resolve("fflate")).href
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

const ROOT = path.resolve(import.meta.dirname, "../..");
const PANE_GEOMETRY_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/grid-pane-geometry.ts");
const GRID_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/XlsxGrid.vue");
const CONTEXT_MENU_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/XlsxContextMenu.vue");
const GRID_SOURCE = readFileSync(GRID_PATH, "utf8");

function loadTsModule(source) {
  const output = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.ESNext,
      target: ts.ScriptTarget.ES2022,
    },
  }).outputText;
  return import(`data:text/javascript;base64,${Buffer.from(output).toString("base64")}`);
}

const geometry = await loadTsModule(readFileSync(PANE_GEOMETRY_PATH, "utf8"));
const {
  axisLayoutBoundary,
  axisLayoutOffset,
  axisLayoutScroll,
  mirrorBodyOffset,
  resolvePaneContentPoint,
  resolvePaneOverlayRect,
  resolvePaneViewPoint,
} = geometry;

const samples = new URL("../../apps/demo/public/samples/", import.meta.url);

/**
 * The charts/images fixture with Excel's right-to-left sheet view and the
 * frozen row/column pair the acceptance case names.
 */
function rightToLeftWorkbookBytes(rightToLeft = true) {
  const archive = unzipSync(new Uint8Array(readFileSync(new URL("charts-images.xlsx", samples))));
  const sheet = strFromU8(archive["xl/worksheets/sheet1.xml"]);
  const view = rightToLeft
    ? '<sheetViews><sheetView workbookViewId="0" rightToLeft="1"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>'
    : '<sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>';
  archive["xl/worksheets/sheet1.xml"] = strToU8(
    sheet.replace(/<sheetViews>.*?<\/sheetViews>/s, view),
  );
  return zipSync(archive);
}

test("XLSX parser reads sheetView/@rightToLeft", () => {
  const rtl = core.parseWorkbookStructureAssets(rightToLeftWorkbookBytes(true));
  assert.equal(rtl.sheetStatesByWorkbookSheetIndex[0].rightToLeft, true);
  assert.deepEqual(rtl.sheetStatesByWorkbookSheetIndex[0].freezePanes, { col: 1, row: 1 });

  const ltr = core.parseWorkbookStructureAssets(rightToLeftWorkbookBytes(false));
  assert.equal(ltr.sheetStatesByWorkbookSheetIndex[0].rightToLeft, false);
});

test("RTL column offsets mirror within the content width", () => {
  // Three 80 px columns: A at the right edge, then B and C to its left.
  const total = 240;
  assert.equal(axisLayoutOffset(0, 80, total, true), 160, "A starts 160 px from the left");
  assert.equal(axisLayoutOffset(80, 160, total, true), 80, "B starts 80 px from the left");
  assert.equal(axisLayoutOffset(160, 240, total, true), 0, "C starts at the left edge");
  assert.equal(axisLayoutOffset(0, 80, total, false), 0, "LTR keeps the content offset");

  assert.equal(axisLayoutBoundary(0, total, true), 240, "boundary 0 is the right edge");
  assert.equal(axisLayoutBoundary(80, total, true), 160);
  assert.equal(axisLayoutBoundary(240, total, true), 0, "boundary 3 is the left edge");
  assert.equal(axisLayoutBoundary(80, total, false), 80);
});

test("RTL scrolling starts at the right edge and ends at column 0", () => {
  // 240 px of columns in a 120 px viewport.
  assert.equal(axisLayoutScroll(0, 120, 240, true), 120, "scroll 0 shows the right edge");
  assert.equal(axisLayoutScroll(120, 120, 240, true), 0, "the last scroll shows the left edge");
  assert.equal(axisLayoutScroll(0, 120, 240, false), 0);
  assert.equal(mirrorBodyOffset(10, 120, true), 110, "pointer x is measured from the right");
  assert.equal(mirrorBodyOffset(10, 120, false), 10);
});

test("RTL hit testing resolves the pinned right-hand strip", () => {
  // Body is 200 px wide, column A is frozen at 0..80 px from the right.
  const frozen = { width: 80, height: 24 };
  const main = resolvePaneContentPoint(mirrorBodyOffset(30, 200, true), 10, 40, 0, frozen.width, frozen.height);
  assert.equal(main.pane, "frozen-rows");
  assert.equal(main.contentX, 170 + 40, "the scrolling pane adds scrollLeft");

  const pinned = resolvePaneContentPoint(mirrorBodyOffset(190, 200, true), 10, 40, 0, frozen.width, frozen.height);
  assert.equal(pinned.pane, "frozen-corner");
  assert.equal(pinned.contentX, 10, "column A keeps its content coordinate");

  // The same point on an LTR sheet is column A/B territory, not the pinned
  // right-hand strip.
  const ltr = resolvePaneContentPoint(30, 10, 40, 0, frozen.width, frozen.height);
  assert.equal(ltr.pane, "frozen-corner");
  assert.equal(ltr.contentX, 30);
});

test("RTL overlays mirror into the pinned right-hand pane", () => {
  const base = {
    frozenCols: 1,
    frozenRows: 1,
    frozenWidth: 80,
    frozenHeight: 24,
    viewportWidth: 400,
    viewportHeight: 200,
    scrollLeft: 50,
    scrollTop: 30,
    rtl: true,
  };
  assert.deepEqual(
    resolvePaneViewPoint({ ...base, contentX: 0, contentY: 0, width: 80, displayCol: 0, displayRow: 0 }),
    { x: 320, y: 0 },
    "column A is pinned to the right edge",
  );
  assert.deepEqual(
    resolvePaneViewPoint({ ...base, contentX: 160, contentY: 0, width: 80, displayCol: 2, displayRow: 0 }),
    { x: 400 - (160 + 80 - 50), y: 0 },
    "a scrolling cell mirrors its content interval",
  );
  assert.deepEqual(
    resolvePaneOverlayRect({
      ...base, contentX: 0, contentY: 0, width: 80, height: 24, displayCol: 0, displayRow: 0,
    }),
    { left: 320, top: 0, width: 80, height: 24 },
    "a pinned overlay keeps the right strip",
  );
  assert.deepEqual(
    resolvePaneOverlayRect({
      ...base, contentX: 100, contentY: 0, width: 80, height: 24, displayCol: 1, displayRow: 0,
    }),
    { left: 270, top: 0, width: 50, height: 24 },
    "an overlay never crosses into the pinned strip",
  );
});

test("XLSX grid maps the mirrored layout through one set of helpers", () => {
  for (const fragment of [
    /const isRightToLeft = computed\(\(\) => activeSheet\.value\?\.rightToLeft === true\)/,
    /function getColLayoutX\(displayCol: number\)/,
    /function getColLayoutBoundary\(boundary: number\)/,
    /function getColLayoutScroll\(bodyWidth: number\)/,
    /function paneCellViewX\(contentX: number, width: number, frozen: boolean\)/,
  ]) {
    assert.match(GRID_SOURCE, fragment);
  }

  const paintBody = GRID_SOURCE.slice(
    GRID_SOURCE.indexOf("function paintBody("),
    GRID_SOURCE.indexOf("function paintFrozenPanes("),
  );
  assert.match(paintBody, /ctx\.translate\(-getColLayoutScroll\(bw\), -scrollTop\.value\)/);
  assert.match(paintBody, /const colX = mergeEnd && isRightToLeft\.value/);
  assert.match(paintBody, /getColLayoutBoundary\(c\)/);
  assert.match(paintBody, /style\.alignment !== trailingAlignment/);
  assert.match(GRID_SOURCE, /target\.alignment = isRightToLeft\.value && horizontalAlignment !== "center"/);

  const resolvePanePoint = GRID_SOURCE.slice(
    GRID_SOURCE.indexOf("function resolvePanePoint("),
    GRID_SOURCE.indexOf("/** Body viewport size"),
  );
  assert.match(resolvePanePoint, /mirrorBodyOffset\(/);

  const axis = GRID_SOURCE.slice(
    GRID_SOURCE.indexOf("function hitTestAxis("),
    GRID_SOURCE.indexOf("function selectAxisRange("),
  );
  assert.match(axis, /mirrorBodyOffset\(/);

  const scroll = GRID_SOURCE.slice(
    GRID_SOURCE.indexOf("function scrollToCell("),
    GRID_SOURCE.indexOf("function onGridKeydown("),
  );
  assert.match(scroll, /const maxVisibleScroll = colX - frozen\.width/);

  assert.match(GRID_SOURCE, /ArrowLeft: \(\) => moveCell\(0, isRightToLeft\.value \? 1 : -1/);
  assert.match(GRID_SOURCE, /ArrowRight: \(\) => moveCell\(0, isRightToLeft\.value \? -1 : 1/);
  assert.match(GRID_SOURCE, /isRightToLeft\.value \? -1 : 1\)/);

  const contextMenu = readFileSync(CONTEXT_MENU_PATH, "utf8");
  assert.match(contextMenu, /sheet\.rightToLeft === true/);
  assert.match(
    contextMenu,
    /Math\.max\(0, rect\.width - ROW_HEADER_WIDTH\) - localX/,
    "right-click cell resolution mirrors the pointer offset",
  );
});
