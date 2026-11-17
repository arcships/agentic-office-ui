import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";

import ts from "../../packages/xlsx-core/node_modules/typescript/lib/typescript.js";

// Issue #13: the XLSX grid must not do data-sized work per paint, must not
// allocate a worksheet handle per section, and must not precompute a thumbnail
// per sheet before the first hover.

const ROOT = path.resolve(import.meta.dirname, "../..");
const GRID_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/XlsxGrid.vue");
const STATS_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/conditional-format-stats.ts");
const THUMBNAILS_PATH = path.join(
  ROOT,
  "packages/vue-xlsx/src/composables/useXlsxViewerThumbnails.ts",
);
const SHEET_TABS_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/XlsxSheetTabs.vue");

const THUMBNAIL_MAX_ROWS = 60;
const THUMBNAIL_MAX_COLS = 26;

function transpile(source) {
  return ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.ESNext,
      target: ts.ScriptTarget.ES2022,
    },
  }).outputText;
}

function moduleUrl(code) {
  return `data:text/javascript;base64,${Buffer.from(code).toString("base64")}`;
}

// The composable only needs `computed`; a data-URL shim with Vue's memoizing
// behaviour keeps the test free of the workspace build.
const VUE_SHIM = moduleUrl(`
  export const computed = (getter) => {
    let cached;
    let computedValue = false;
    return {
      get value() {
        if (!computedValue) {
          cached = getter();
          computedValue = true;
        }
        return cached;
      },
    };
  };
`);

async function loadStatsModule() {
  return import(moduleUrl(transpile(readFileSync(STATS_PATH, "utf8"))));
}

async function loadThumbnailsModule() {
  const code = transpile(readFileSync(THUMBNAILS_PATH, "utf8"))
    .replace(/from\s*"vue"/u, `from "${VUE_SHIM}"`);
  return import(moduleUrl(code));
}

function scriptSource(file) {
  const source = readFileSync(file, "utf8");
  if (!file.endsWith(".vue")) return source;
  return [...source.matchAll(/<script(?:\s+setup)?[^>]*>([\s\S]*?)<\/script>/gu)]
    .map((match) => match[1])
    .join("\n");
}

function parseScript(file) {
  return ts.createSourceFile(
    file,
    scriptSource(file),
    ts.ScriptTarget.Latest,
    true,
    ts.ScriptKind.TS,
  );
}

function descendants(node, predicate) {
  const matches = [];
  function visit(current) {
    if (predicate(current)) matches.push(current);
    ts.forEachChild(current, visit);
  }
  visit(node);
  return matches;
}

function functionNode(ast, name) {
  return descendants(
    ast,
    (node) => ts.isFunctionDeclaration(node) && node.name?.text === name,
  )[0];
}

function cellsIn(node, ast, name) {
  return descendants(node, (child) => {
    if (!ts.isCallExpression(child)) return false;
    const callee = child.expression;
    if (ts.isPropertyAccessExpression(callee)) return callee.name.text === name;
    return ts.isIdentifier(callee) && callee.text === name;
  }).map((call) => call.getText(ast));
}

function range(startRow, startCol, endRow, endCol) {
  return { start: { row: startRow, col: startCol }, end: { row: endRow, col: endCol } };
}

// ── Conditional-format values ────────────────────────────────────────

test("XLSX conditional values scan small rule ranges exactly and skip blank cells", async () => {
  const stats = await loadStatsModule();
  const reads = [];
  const readValue = (cell) => {
    reads.push(cell);
    if ((cell.row + cell.col) % 5 === 0) return "   ";
    return String(cell.row * 10 + cell.col);
  };
  const ranges = stats.normalizeConditionalFormatRanges([range(0, 0, 4, 9)]);

  const values = stats.collectConditionalFormatValues(
    ranges,
    { maxUsedCol: 999, maxUsedRow: 999 },
    readValue,
  );

  assert.equal(reads.length, 50, "a 50-cell rule range must be scanned cell by cell");
  assert.deepEqual(
    values,
    reads
      .filter((cell) => (cell.row + cell.col) % 5 !== 0)
      .map((cell) => cell.row * 10 + cell.col),
    "blank cells never enter a rule's scale",
  );
});

test("XLSX conditional values cap the reads for data-sized rule ranges", async () => {
  const stats = await loadStatsModule();
  let reads = 0;
  const readValue = (cell) => {
    reads += 1;
    return String(cell.row * 100 + cell.col);
  };
  // 500 × 100 = 50 000 cells, the shape that used to cost one formatted lookup
  // per cell on every repaint.
  const ranges = stats.normalizeConditionalFormatRanges([range(0, 0, 499, 99)]);
  const bounds = { maxUsedCol: 99, maxUsedRow: 499 };

  const values = stats.collectConditionalFormatValues(ranges, bounds, readValue);
  reads = 0;
  const second = stats.collectConditionalFormatValues(ranges, bounds, readValue);

  assert.ok(reads <= stats.CONDITIONAL_FORMAT_CELL_BUDGET + 1, `sampled ${reads} cells`);
  assert.ok(reads < 50_000 / 4, `sampling must stay far below the range size, read ${reads}`);
  assert.equal(reads > 0, true, "sampling must still read the range");
  assert.deepEqual(second, values, "sampling is deterministic");
  assert.ok(values.length <= stats.CONDITIONAL_FORMAT_CELL_BUDGET, `kept ${values.length} values`);
  const min = Math.min(...values);
  const max = Math.max(...values);
  assert.ok(min >= 0 && max <= 49_999, "sampled values stay inside the data");
  assert.ok(max > min, "sampled values keep the data spread");
});

test("XLSX conditional values stay bounded for ranges that would overflow a Math.min spread", async () => {
  const stats = await loadStatsModule();
  let reads = 0;
  const readValue = () => {
    reads += 1;
    return "1,5";
  };
  const ranges = stats.normalizeConditionalFormatRanges([range(0, 0, 1999, 99)]);

  const values = stats.collectConditionalFormatValues(
    ranges,
    { maxUsedCol: 99, maxUsedRow: 1999 },
    readValue,
  );

  assert.equal(values.every((value) => value === 15), true, "display grouping is stripped");
  assert.ok(values.length <= stats.CONDITIONAL_FORMAT_CELL_BUDGET, `kept ${values.length} values`);
  assert.ok(reads <= stats.CONDITIONAL_FORMAT_CELL_BUDGET + 1, `sampled ${reads} cells`);
});

test("XLSX conditional values clamp rule ranges to the sheet used extent", async () => {
  const stats = await loadStatsModule();
  const reads = [];
  const ranges = stats.normalizeConditionalFormatRanges([range(0, 0, 999, 4)]);

  stats.collectConditionalFormatValues(ranges, { maxUsedCol: 4, maxUsedRow: 9 }, (cell) => {
    reads.push(cell);
    return "1";
  });

  assert.equal(reads.length, 50, "only the used 10 × 5 window is read");
  assert.deepEqual(reads.at(-1), { col: 4, row: 9 });
});

test("XLSX conditional ranges normalize reversed ranges like the per-cell test did", async () => {
  const stats = await loadStatsModule();
  const [normalized] = stats.normalizeConditionalFormatRanges([range(4, 3, 1, 0)]);

  assert.deepEqual(normalized, { endCol: 3, endRow: 4, startCol: 0, startRow: 1 });
  assert.equal(stats.cellInNormalizedRange({ col: 0, row: 1 }, normalized), true);
  assert.equal(stats.cellInNormalizedRange({ col: 4, row: 4 }, normalized), false);
});

test("XLSX scoped cache keeps computed rule data across repaints and drops it per revision", async () => {
  const stats = await loadStatsModule();
  const cache = stats.createScopedCache();
  const rule = { ranges: [range(0, 0, 0, 0)] };
  const revisionOne = { revision: 1 };
  const revisionTwo = { revision: 2 };
  let computed = 0;
  const resolveWith = (scope) => cache.resolve(scope, rule, () => {
    computed += 1;
    return { min: computed, max: computed };
  });

  for (let paint = 0; paint < 60; paint += 1) resolveWith(revisionOne);
  assert.equal(computed, 1, "60 repaints on one revision must compute the stats once");
  assert.equal(cache.size, 1);

  assert.deepEqual(resolveWith(revisionOne), { min: 1, max: 1 });
  assert.equal(computed, 1);

  assert.deepEqual(resolveWith(revisionTwo), { min: 2, max: 2 });
  assert.equal(computed, 2, "a new revision invalidates the cached stats");

  const secondRule = { ranges: [range(0, 1, 0, 1)] };
  cache.resolve(revisionTwo, secondRule, () => {
    computed += 1;
    return { min: 9, max: 9 };
  });
  assert.equal(computed, 3, "each rule keeps its own entry");
  assert.equal(cache.size, 2);
});

test("XLSX grid keeps conditional-format work out of the paint path", () => {
  const ast = parseScript(GRID_PATH);
  const paintBody = functionNode(ast, "paintBody");
  const paintGrid = functionNode(ast, "paintGrid");
  assert.ok(paintBody && paintGrid, "paintBody/paintGrid were not found");

  const resets = descendants(paintBody, (node) => {
    if (!ts.isBinaryExpression(node) || node.operatorToken.kind !== ts.SyntaxKind.EqualsToken) {
      return false;
    }
    const target = node.left;
    if (ts.isIdentifier(target)) return /conditional/i.test(target.text);
    if (ts.isPropertyAccessExpression(target)) return /conditional/i.test(target.name.text);
    return false;
  });
  assert.deepEqual(
    resets.map((node) => node.getText(ast)),
    [],
    "paintBody must not reset the conditional-format cache; it is revision scoped",
  );

  const scope = functionNode(ast, "currentConditionalFormatScope");
  assert.ok(scope, "currentConditionalFormatScope was not found");
  const scopeBody = scope.getText(ast);
  assert.match(scopeBody, /revision\.value/u, "the cache scope must follow the sheet revision");
  assert.match(
    scopeBody,
    /workerRowsRevision\.value/u,
    "worker-hydrated rows must also invalidate the cache scope",
  );
  assert.match(scopeBody, /workbookSheetIndex/u, "the cache scope must follow the active sheet");
  assert.match(scopeBody, /themePalette/u, "a theme swap must invalidate resolved conditional colours");

  // Rule environments, per-cell bounds and the value scan all hang off scoped
  // caches, and the bounded scan is what the grid feeds them.
  const gridSource = readFileSync(GRID_PATH, "utf8");
  assert.match(
    gridSource,
    /const conditionalRuleEnvironments = createScopedCache<ConditionalEnvironment>\(\)/u,
    "rule environments must not be a per-paint Map",
  );
  assert.match(gridSource, /const conditionalRuleRanges = createScopedCache<NormalizedRange\[\]>\(\)/u);
  assert.match(gridSource, /collectConditionalFormatValues\(/u);
  assert.match(
    gridSource,
    /conditionalRuleEnvironments\.resolve\(scope, rule/u,
    "rule environments must come from the revision-scoped cache",
  );
});

// ── Worksheet handle allocation ───────────────────────────────────────

test("XLSX grid allocates and frees one worksheet handle per paint", () => {
  const ast = parseScript(GRID_PATH);
  const paintGrid = functionNode(ast, "paintGrid");
  const paintBody = functionNode(ast, "paintBody");
  const paintFrozenPanes = functionNode(ast, "paintFrozenPanes");
  assert.ok(paintGrid && paintBody && paintFrozenPanes, "paint functions were not found");

  assert.equal(
    cellsIn(paintGrid, ast, "getActiveWorksheet").length,
    1,
    "paintGrid must open exactly one worksheet handle for the whole frame",
  );
  assert.deepEqual(cellsIn(paintBody, ast, "getActiveWorksheet"), []);
  assert.deepEqual(cellsIn(paintFrozenPanes, ast, "getActiveWorksheet"), []);
  assert.deepEqual(
    cellsIn(paintBody, ast, "free"),
    [],
    "paintBody must not free the handle it does not own",
  );
  assert.deepEqual(
    cellsIn(paintFrozenPanes, ast, "free"),
    [],
    "paintFrozenPanes must not free the handle it does not own",
  );

  const freesInFinally = descendants(paintGrid, (node) => {
    if (!ts.isTryStatement(node) || !node.finallyBlock) return false;
    return cellsIn(node.finallyBlock, ast, "free").length > 0;
  });
  assert.equal(
    freesInFinally.length,
    1,
    "the single handle must be released in a finally so early returns cannot leak it",
  );
});

test("XLSX sheet tabs paint previews on hover instead of preloading every sheet", () => {
  const source = readFileSync(SHEET_TABS_PATH, "utf8");
  assert.equal(
    /onMounted\s*\(/u.test(source),
    false,
    "XlsxSheetTabs must not preload every thumbnail on mount",
  );
  assert.equal(
    /document\.createElement\(\s*"canvas"\s*\)/u.test(source),
    false,
    "offscreen canvases must not be created for sheets nobody hovered",
  );
  assert.match(source, /thumb\.paint\(canvas as HTMLCanvasElement\)/u, "hover must paint the preview");
});

// ── Sheet thumbnails ──────────────────────────────────────────────────

function createThumbnailProbe(sheetCount) {
  const counters = { formattedReads: 0, handlesOpened: 0, handlesFreed: 0, rowHeights: 0, columnWidths: 0 };
  let revision = 0;
  const controller = {
    get revision() {
      return revision;
    },
    sheets: Array.from({ length: sheetCount }, (_, index) => ({
      maxUsedCol: 20,
      maxUsedRow: 199,
      minUsedCol: 0,
      minUsedRow: 0,
      name: `Sheet${index + 1}`,
      workbookSheetIndex: index,
    })),
    workbook: {
      getSheet() {
        counters.handlesOpened += 1;
        return {
          free() {
            counters.handlesFreed += 1;
          },
          getColumnWidth() {
            counters.columnWidths += 1;
            return 80;
          },
          getFormattedValueAt(row, col) {
            counters.formattedReads += 1;
            return `r${row}c${col}`;
          },
          getRowHeight() {
            counters.rowHeights += 1;
            return 24;
          },
        };
      },
    },
  };
  return {
    bumpRevision() {
      revision += 1;
    },
    controllers: controller,
    counters,
  };
}

function fakeCanvas() {
  const context = {
    beginPath() {},
    clip() {},
    fillRect() {},
    fillText() {},
    lineTo() {},
    moveTo() {},
    rect() {},
    restore() {},
    save() {},
    setTransform() {},
    stroke() {},
  };
  return {
    getContext: () => context,
    height: 0,
    style: {},
    width: 0,
  };
}

test("XLSX sheet thumbnails read nothing until a preview is painted, then cache per revision", async () => {
  const { useXlsxViewerThumbnails } = await loadThumbnailsModule();
  const probe = createThumbnailProbe(40);
  const api = useXlsxViewerThumbnails(() => probe.controllers);

  const thumbnails = api.thumbnails;
  assert.equal(thumbnails.length, 40);
  assert.deepEqual(
    probe.counters,
    { columnWidths: 0, formattedReads: 0, handlesFreed: 0, handlesOpened: 0, rowHeights: 0 },
    "listing a 40-sheet workbook must not read any cell or open any handle",
  );
  assert.deepEqual(thumbnails[7].sourceRange, {
    end: { col: 22, row: 59 },
    start: { col: 0, row: 0 },
  });

  const previewRows = THUMBNAIL_MAX_ROWS;
  const previewCols = 23;

  assert.equal(thumbnails[7].height > 0, true);
  assert.equal(thumbnails[7].height > 0, true);
  assert.equal(probe.counters.handlesOpened, 1, "size lookups must reuse one handle");
  assert.equal(probe.counters.formattedReads, 0, "size lookups must not read cell values");

  const canvas = fakeCanvas();
  assert.equal(api.paintThumbnail(7, canvas), true);
  assert.equal(probe.counters.formattedReads, previewRows * previewCols);
  assert.equal(canvas.width > 0 && canvas.height > 0, true);
  assert.equal(
    probe.counters.handlesOpened,
    probe.counters.handlesFreed,
    "every opened worksheet handle is freed",
  );

  const readsAfterFirstPaint = probe.counters.formattedReads;
  assert.equal(api.paintThumbnail(7, fakeCanvas()), true);
  assert.equal(
    probe.counters.formattedReads,
    readsAfterFirstPaint,
    "repainting the same revision must reuse the cached cell values",
  );

  probe.bumpRevision();
  assert.equal(api.paintThumbnail(7, fakeCanvas()), true);
  assert.equal(
    probe.counters.formattedReads,
    readsAfterFirstPaint * 2,
    "a revision bump refreshes the cached preview",
  );

  assert.equal(api.paintThumbnail(0, fakeCanvas()), true);
  assert.equal(
    probe.counters.formattedReads,
    readsAfterFirstPaint * 2 + previewRows * previewCols,
    "only the hovered sheet is read",
  );
});
