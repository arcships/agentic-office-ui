import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";
import { fileURLToPath } from "node:url";

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const ts = requireFromXlsxCore("typescript");

function loadTranspiledSource(relativePath) {
  const source = readFileSync(new URL(relativePath, import.meta.url), "utf8");
  const emitted = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.ESNext,
      target: ts.ScriptTarget.ES2022,
    },
  }).outputText;
  return import(`data:text/javascript;base64,${Buffer.from(emitted).toString("base64")}`);
}

const {
  FORMULA_REFERENCE_COLORS,
  collectFormulaReferences,
  formulaReferenceColor,
  formulaReferenceSegments,
  parseFormulaReferences,
  resolveNamedRangeTarget,
} = await loadTranspiledSource(
  "../../packages/vue-xlsx/src/composables/formula-reference-highlight.ts",
);
const { resolvePaintedReferenceRange } = await loadTranspiledSource(
  "../../packages/vue-xlsx/src/components/grid-paint.ts",
);

const GRID_PATH = new URL(
  "../../packages/vue-xlsx/src/components/XlsxGrid.vue",
  import.meta.url,
);
const GRID_SOURCE = readFileSync(GRID_PATH, "utf8");
const GRID_SCRIPT = [...GRID_SOURCE.matchAll(/<script(?:\s+setup)?[^>]*>([\s\S]*?)<\/script>/g)]
  .map((match) => match[1])
  .join("\n");
const GRID_AST = ts.createSourceFile(
  fileURLToPath(GRID_PATH),
  GRID_SCRIPT,
  ts.ScriptTarget.Latest,
  true,
  ts.ScriptKind.TS,
);

function gridFunction(name) {
  let found = null;
  function visit(node) {
    if (found) return;
    if (ts.isFunctionDeclaration(node) && node.name?.text === name) {
      found = node;
      return;
    }
    ts.forEachChild(node, visit);
  }
  visit(GRID_AST);
  return found;
}

async function loadReferenceBorderPainter() {
  const source = gridFunction("paintFormulaReferenceBorders")?.getText(GRID_AST);
  assert.ok(source, "XlsxGrid.paintFormulaReferenceBorders exists");
  const prelude = `
    const formulaReferenceRanges = { value: [] };
    const activeFormulaReference = { value: null };
    const sheetRowToDisplayIndex = { value: new Map() };
    const sheetColToDisplayIndex = { value: new Map() };
    const displayRows = { value: [0, 1, 2, 3] };
    const displayCols = { value: [0, 1, 2] };
    const drawn = [];
    function resolvePaintedReferenceRange(range) {
      return { start: { ...range.start }, end: { ...range.end } };
    }
    function drawRangeBorder(ctx, range, style) { drawn.push({ range, style }); }
  `;
  const output = ts.transpileModule(
    `${prelude}\n${source}\nexport { formulaReferenceRanges, activeFormulaReference, drawn, paintFormulaReferenceBorders };`,
    {
      compilerOptions: {
        module: ts.ModuleKind.ESNext,
        target: ts.ScriptTarget.ES2022,
      },
    },
  ).outputText;
  return import(`data:text/javascript;base64,${Buffer.from(output).toString("base64")}`);
}

function addresses(references) {
  return references.map((reference) => reference.rangeAddress);
}

function caretInside(formula, token, offset = 1) {
  const start = formula.indexOf(token);
  assert.notEqual(start, -1, `missing ${token} in ${formula}`);
  return start + offset;
}

test("a formula's range and cell references are parsed with distinct colours", () => {
  const formula = "=SUM(A1:A3)+B2";
  const { references, activeIndex, activeReference } = collectFormulaReferences(
    formula,
    caretInside(formula, "A1:A3"),
  );

  assert.deepEqual(addresses(references), ["A1:A3", "B2"]);
  assert.equal(references[0].kind, "range");
  assert.deepEqual(references[0].range, {
    start: { row: 0, col: 0 },
    end: { row: 2, col: 0 },
  });
  assert.equal(references[1].kind, "cell");
  assert.deepEqual(references[1].range, {
    start: { row: 1, col: 1 },
    end: { row: 1, col: 1 },
  });
  assert.notEqual(
    references[0].color,
    references[1].color,
    "two references must be tellable apart by colour",
  );
  assert.equal(FORMULA_REFERENCE_COLORS.includes(references[0].color), true);
  assert.equal(activeIndex, 0);
  assert.equal(activeReference, references[0]);
});

test("moving the caret changes which reference is active", () => {
  const formula = "=SUM(A1:A3)+B2";
  const first = collectFormulaReferences(formula, caretInside(formula, "A1:A3"));
  const second = collectFormulaReferences(formula, caretInside(formula, "B2"));
  const none = collectFormulaReferences(formula, formula.indexOf("SUM"));

  assert.equal(first.activeReference?.rangeAddress, "A1:A3");
  assert.equal(second.activeReference?.rangeAddress, "B2");
  assert.equal(second.activeIndex, 1);
  assert.equal(none.activeIndex, -1, "a caret outside every reference activates none");
  assert.equal(none.activeReference, null);
});

test("repeated references keep one colour and regard $ anchors as the same target", () => {
  const { references } = collectFormulaReferences("=A1+$A$1+A1", 0);
  assert.deepEqual(addresses(references), ["A1", "A1", "A1"]);
  assert.deepEqual(
    references.map((reference) => reference.text),
    ["A1", "$A$1", "A1"],
    "the typed text keeps its anchors even though the target range is canonical",
  );
  assert.equal(references[0].color, references[1].color);
  assert.equal(references[1].color, references[2].color);
});

test("sheet-qualified references keep their sheet name and quoted spaces", () => {
  const formula = "=SUM('My Sheet'!$A$1:$B$2)+Sheet2!C3";
  const { references } = collectFormulaReferences(formula, 0);

  assert.equal(references.length, 2);
  assert.equal(references[0].sheetName, "My Sheet");
  assert.equal(references[0].rangeAddress, "A1:B2");
  assert.equal(references[0].text, "'My Sheet'!$A$1:$B$2");
  assert.equal(references[1].sheetName, "Sheet2");
  assert.equal(references[1].rangeAddress, "C3");
});

test("unicode sheet names and defined names parse like ASCII ones", () => {
  const calls = [];
  const { references } = collectFormulaReferences("=SUM(工作表1!A1:B2)+总计", 0, {
    resolveName(name) {
      calls.push(name);
      return { range: { start: { row: 2, col: 2 }, end: { row: 2, col: 2 } } };
    },
  });

  assert.equal(references.length, 2);
  assert.equal(references[0].sheetName, "工作表1");
  assert.equal(references[0].rangeAddress, "A1:B2");
  assert.equal(references[1].name, "总计");
  assert.equal(references[1].rangeAddress, "C3");
  assert.deepEqual(calls, ["总计"]);
});

test("whole-column and whole-row references cover Excel's bounds", () => {
  const { references } = collectFormulaReferences("=SUM(A:A)+SUM(3:3)", 0);
  assert.deepEqual(addresses(references), ["A:A", "3:3"]);
  assert.deepEqual(references[0].range, {
    start: { row: 0, col: 0 },
    end: { row: 1_048_575, col: 0 },
  });
  assert.deepEqual(references[1].range, {
    start: { row: 2, col: 0 },
    end: { row: 2, col: 16_383 },
  });
});

test("string literals, function names and external workbooks are not references", () => {
  const formula = '="A1"&LOG10(2)+SUM(B1:C1)';
  assert.deepEqual(addresses(parseFormulaReferences(formula)), ["B1:C1"]);

  const external = collectFormulaReferences("=[Book1.xlsx]Sheet1!A1+B1", 0);
  assert.deepEqual(addresses(external.references), ["B1"]);

  // LOG10 looks like a cell address but is a function call.
  assert.deepEqual(parseFormulaReferences("=LOG10(100)"), []);
});

test("ranges tolerate Excel's whitespace around the range operator", () => {
  const { references } = collectFormulaReferences("=SUM(A1 : A3)", 4);
  assert.equal(references.length, 1);
  assert.equal(references[0].rangeAddress, "A1:A3");
  assert.equal(references[0].text, "A1 : A3");
  assert.equal(parseFormulaReferences("=XFE1").length, 0, "XFE is outside 16,384 columns");
  assert.deepEqual(addresses(parseFormulaReferences("=@A1#")), ["A1"]);
});

test("defined names resolve through the supplied resolver and unknown names stay plain", () => {
  const calls = [];
  const options = {
    resolveName(name, sheetName) {
      calls.push([name, sheetName]);
      return name === "Total"
        ? { range: { start: { row: 0, col: 0 }, end: { row: 2, col: 0 } }, sheetName: null }
        : null;
    },
  };
  const resolved = collectFormulaReferences("=Total*2+Missing", 1, options);
  assert.deepEqual(addresses(resolved.references), ["A1:A3"]);
  assert.equal(resolved.references[0].kind, "name");
  assert.equal(resolved.references[0].name, "Total");
  assert.deepEqual(calls[0], ["Total", null]);

  const scoped = collectFormulaReferences("=Sheet2!Total", 3, options);
  assert.equal(scoped.references[0]?.sheetName, "Sheet2");
  assert.deepEqual(calls.at(-1), ["Total", "Sheet2"]);

  const elsewhere = collectFormulaReferences("=OtherName", 1, {
    resolveName: () => ({
      range: { start: { row: 4, col: 1 }, end: { row: 4, col: 1 } },
      sheetName: "Other Sheet",
    }),
  });
  assert.equal(elsewhere.references[0]?.sheetName, "Other Sheet");
  assert.equal(elsewhere.references[0]?.rangeAddress, "B5");
});

test("defined-name expressions resolve to their first static range", () => {
  assert.deepEqual(resolveNamedRangeTarget("Sheet1!$A$1:$C$3"), {
    range: { start: { row: 0, col: 0 }, end: { row: 2, col: 2 } },
    sheetName: "Sheet1",
  });
  assert.deepEqual(resolveNamedRangeTarget("SUM(Sheet2!$B$2:$B$4)", "Sheet9"), {
    range: { start: { row: 1, col: 1 }, end: { row: 3, col: 1 } },
    sheetName: "Sheet2",
  });
  assert.deepEqual(resolveNamedRangeTarget("$D$4", "Sheet3"), {
    range: { start: { row: 3, col: 3 }, end: { row: 3, col: 3 } },
    sheetName: "Sheet3",
  });
  assert.equal(resolveNamedRangeTarget("{1,2;3,4}"), null);
  assert.equal(resolveNamedRangeTarget("#REF!"), null);
});

test("segments preserve the formula text and mark only references", () => {
  const formula = "=SUM(A1:A3)+B2";
  const { references, activeIndex } = collectFormulaReferences(formula, caretInside(formula, "B2"));
  const segments = formulaReferenceSegments(formula, references, activeIndex);

  assert.equal(segments.map((segment) => segment.text).join(""), formula);
  assert.deepEqual(
    segments.filter((segment) => segment.reference).map((segment) => [
      segment.text,
      segment.reference?.rangeAddress,
      segment.active,
    ]),
    [
      ["A1:A3", "A1:A3", false],
      ["B2", "B2", true],
    ],
  );
  assert.equal(segments[0].color, null);
  assert.equal(segments.find((segment) => segment.text === "A1:A3")?.color, references[0].color);
});

test("the palette cycles after eight distinct references", () => {
  const refs = parseFormulaReferences("=A1+B2+C3+D4+E5+F6+G7+H8+I9");
  assert.equal(refs.length, 9);
  assert.equal(refs[8].color, refs[0].color);
  assert.equal(formulaReferenceColor(8), FORMULA_REFERENCE_COLORS[0]);
  assert.equal(formulaReferenceColor(-1), FORMULA_REFERENCE_COLORS[0]);
});

test("plain values are not formulas and produce no references", () => {
  assert.deepEqual(parseFormulaReferences("A1"), []);
  assert.deepEqual(parseFormulaReferences("SUM(A1:A3)"), []);
  assert.deepEqual(parseFormulaReferences(""), []);
});

test("reference outlines clamp whole rows/columns to the visible display", () => {
  const rowMap = new Map([[0, 0], [1, 1], [2, 2], [3, 3]]);
  const colMap = new Map([[0, 0], [1, 1], [2, 2]]);

  assert.deepEqual(
    resolvePaintedReferenceRange(
      { start: { row: 0, col: 0 }, end: { row: 1_048_575, col: 0 } },
      rowMap,
      colMap,
      4,
      3,
    ),
    { start: { row: 0, col: 0 }, end: { row: 3, col: 0 } },
    "a whole column paints to the last displayed row",
  );
  assert.deepEqual(
    resolvePaintedReferenceRange(
      { start: { row: 2, col: 0 }, end: { row: 2, col: 16_383 } },
      rowMap,
      colMap,
      4,
      3,
    ),
    { start: { row: 2, col: 0 }, end: { row: 2, col: 2 } },
    "a whole row paints to the last displayed column",
  );
  assert.equal(
    resolvePaintedReferenceRange(
      { start: { row: 900, col: 900 }, end: { row: 900, col: 900 } },
      rowMap,
      colMap,
      4,
      3,
    ),
    null,
    "a single off-screen cell is not stretched over the viewport",
  );
  assert.equal(
    resolvePaintedReferenceRange(
      { start: { row: 900, col: 1 }, end: { row: 950, col: 1 } },
      rowMap,
      colMap,
      4,
      3,
    ),
    null,
    "a multi-row range below the viewport is not stretched either",
  );
});

test("the grid outlines every reference with its own colour and thickens the active one", async () => {
  const painter = await loadReferenceBorderPainter();
  const first = {
    color: "#1d4ed8",
    range: { start: { row: 0, col: 0 }, end: { row: 2, col: 0 } },
    rangeAddress: "A1:A3",
  };
  const second = {
    color: "#15803d",
    range: { start: { row: 1, col: 1 }, end: { row: 1, col: 1 } },
    rangeAddress: "B2",
  };
  painter.formulaReferenceRanges.value = [first, second];
  painter.activeFormulaReference.value = first;
  painter.paintFormulaReferenceBorders({});

  assert.equal(painter.drawn.length, 2);
  assert.deepEqual(
    painter.drawn.map((entry) => entry.style.color),
    ["#1d4ed8", "#15803d"],
    "each range keeps its palette colour",
  );
  assert.equal(painter.drawn[0].style.lineWidth, 3, "the caret's range is the thick outline");
  assert.equal(painter.drawn[1].style.lineWidth, 1.5);
  assert.equal(painter.drawn[0].style.displayed, true, "the clamp already returned display coordinates");
});
