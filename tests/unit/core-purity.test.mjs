import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import test from "node:test";
import { fileURLToPath, pathToFileURL } from "node:url";

const rootDir = resolve(dirname(fileURLToPath(import.meta.url)), "../..");
const docxCoreUrl = pathToFileURL(
  resolve(rootDir, "packages/docx-core/src/core.ts"),
).href;
const xlsxCoreUrl = pathToFileURL(
  resolve(rootDir, "packages/xlsx-core/src/core.ts"),
).href;
const resolveHookUrl = `data:text/javascript,${encodeURIComponent(`
  export async function resolve(specifier, context, nextResolve) {
    try {
      return await nextResolve(specifier, context);
    } catch (error) {
      if (specifier.startsWith(".") && !specifier.endsWith(".ts")) {
        return nextResolve(specifier + ".ts", context);
      }
      throw error;
    }
  }
`)}`;
const registerHookUrl = `data:text/javascript,${encodeURIComponent(`
  import { register } from "node:module";
  register(${JSON.stringify(resolveHookUrl)}, import.meta.url);
`)}`;

function runNode(args) {
  return spawnSync(process.execPath, args, {
    cwd: rootDir,
    encoding: "utf8",
    env: { ...process.env, NODE_NO_WARNINGS: "1" },
  });
}

function declarationGraph(entryPath) {
  const files = new Map();
  const visit = (filePath) => {
    if (files.has(filePath)) return;
    const source = readFileSync(filePath, "utf8");
    files.set(filePath, source);
    for (const match of source.matchAll(/(?:from\s+|import\s*)["'](\.[^"']+)["']/g)) {
      const specifier = match[1];
      const declaration = resolve(
        dirname(filePath),
        specifier.replace(/\.js$/u, ".d.ts"),
      );
      visit(declaration);
    }
  };
  visit(entryPath);
  return files;
}

function withoutComments(source) {
  return source
    .replace(/\/\*[\s\S]*?\*\//g, "")
    .replace(/\/\/.*$/gm, "");
}

test("source core entries stay deterministic without browser capabilities", () => {
  const childScript = `
    import assert from "node:assert/strict";

    for (const name of ["window", "document", "DOMParser", "XMLSerializer", "Worker"]) {
      Object.defineProperty(globalThis, name, {
        configurable: true,
        get() { throw new Error("forbidden browser global: " + name); },
      });
    }
    Object.defineProperty(URL, "createObjectURL", {
      configurable: true,
      get() { throw new Error("forbidden browser global: URL.createObjectURL"); },
    });

    function deepFreeze(value, seen = new WeakSet()) {
      if (value === null || typeof value !== "object" || seen.has(value)) return value;
      seen.add(value);
      for (const key of Reflect.ownKeys(value)) deepFreeze(value[key], seen);
      return Object.freeze(value);
    }

    const docx = await import(${JSON.stringify(docxCoreUrl)});
    const xlsx = await import(${JSON.stringify(xlsxCoreUrl)});

    const model = deepFreeze({
      metadata: {
        footerSections: [],
        headerSections: [],
        paragraphStyles: [],
        sourceParts: 1,
        warnings: [],
      },
      nodes: [{
        children: [
          { style: { bold: true }, text: "stable", type: "text" },
          { data: [1, 2, 3], mimeType: "image/png", type: "image" },
        ],
        style: { align: "left" },
        type: "paragraph",
      }, {
        rows: [{
          cells: [{
            nodes: [{
              children: [{ text: "cell stable", type: "text" }],
              type: "paragraph",
            }],
            type: "table-cell",
          }],
          type: "table-row",
        }],
        type: "table",
      }],
    });
    const modelBefore = JSON.stringify(model);
    const normalizeThenClone = docx.cloneDocModel(docx.normalizeDocModel(model));
    const cloneThenNormalize = docx.normalizeDocModel(docx.cloneDocModel(model));
    assert.deepEqual(normalizeThenClone, cloneThenNormalize);
    assert.deepEqual(docx.normalizeDocModel(model), docx.normalizeDocModel(model));
    assert.notEqual(normalizeThenClone, model);
    assert.notEqual(normalizeThenClone.nodes[0], model.nodes[0]);
    assert.equal(JSON.stringify(model), modelBefore);
    assert.equal(Object.isFrozen(model.nodes[0].children), true);

    const paragraphUpdated = docx.updateParagraphText(model, 0, "updated");
    const tableUpdated = docx.updateTableCellText(model, 1, 0, 0, "table value");
    const styled = docx.applyRunStyle(model, 0, 0, { italic: true });
    assert.equal(paragraphUpdated.nodes[0].children[0].text, "updated");
    assert.equal(tableUpdated.nodes[1].rows[0].cells[0].nodes[0].children[0].text, "table value");
    assert.equal(styled.nodes[0].children[0].style.italic, true);
    const commandForward = docx.replaceText(
      docx.updateTableCellText(model, 1, 0, 0, "table value"),
      "stable",
      "changed",
    );
    const commandReverse = docx.updateTableCellText(
      docx.replaceText(model, "stable", "changed"),
      1,
      0,
      0,
      "table value",
    );
    assert.equal(JSON.stringify(commandForward), JSON.stringify(commandReverse));
    assert.deepEqual(
      docx.layoutDocument(model, { pageHeight: 180, pageWidth: 240 }),
      docx.layoutDocument(model, { pageHeight: 180, pageWidth: 240 }),
    );
    assert.equal(JSON.stringify(model), modelBefore);

    const palette = deepFreeze({
      colorsByIndex: ["#000000", "#ffffff", "#336699"],
      majorLatinFont: "Major",
      minorLatinFont: "Minor",
    });
    const color = deepFreeze({ theme: 2, tint: 0.25 });
    const fill = deepFreeze({
      degree: 30,
      fillType: "gradient",
      gradientType: "linear",
      stops: [
        { color: { hex: "ff112233" }, position: 1 },
        { color: { rgb: "445566" }, position: 0 },
      ],
    });
    const xlsxBefore = JSON.stringify({ color, fill, palette });

    const forwardColor = xlsx.resolveWorkbookColor(color, palette);
    const forwardFill = xlsx.resolveWorkbookFillStyle(fill, palette);
    const reverseFill = xlsx.resolveWorkbookFillStyle(fill, palette);
    const reverseColor = xlsx.resolveWorkbookColor(color, palette);
    assert.equal(forwardColor, reverseColor);
    assert.deepEqual(forwardFill, reverseFill);
    assert.deepEqual(
      xlsx.resolveWorkbookFillStyle(fill, palette),
      xlsx.resolveWorkbookFillStyle(fill, palette),
    );
    assert.equal(JSON.stringify({ color, fill, palette }), xlsxBefore);
    assert.equal(forwardFill.backgroundColor, "#112233");
    assert.match(forwardFill.backgroundImage, /^linear-gradient\\(/);

    const reversedRange = deepFreeze({
      end: { col: 0, row: 0 },
      start: { col: 27, row: 4 },
    });
    assert.equal(xlsx.columnLabel(0), "A");
    assert.equal(xlsx.columnLabel(27), "AB");
    assert.equal(xlsx.rangeToA1(reversedRange), "A1:AB5");
    assert.equal(
      xlsx.rangeToA1({ start: reversedRange.end, end: reversedRange.start }),
      "A1:AB5",
    );
    assert.equal(
      xlsx.buildA1RangeFormula("O'Brien", { col: 0, row: 0 }, { col: 1, row: 2 }),
      "'O''Brien'!$A$1:$B$3",
    );

    const chart = deepFreeze({
      chartType: "Line",
      series: [{
        categories: [],
        categoriesRef: { formula: "'Sheet 1'!$A$1:$A$2" },
        dataPoints: [],
        id: "series-1",
        name: "Revenue, Q1",
        values: [],
        valuesRef: { formula: "'Sheet 1'!$B$1:$B$2" },
      }],
    });
    const seriesFormula = xlsx.buildChartSeriesFormula(chart, 0);
    assert.deepEqual(
      xlsx.parseChartSeriesFormula(seriesFormula, chart),
      xlsx.parseChartSeriesFormula(seriesFormula, chart),
    );
    assert.equal(xlsx.parseChartSeriesFormula(seriesFormula, chart).nameLiteral, "Revenue, Q1");

    const absoluteAnchor = deepFreeze({
      kind: "absolute",
      positionEmu: { x: 9525, y: 19050 },
      sizeEmu: { cx: 28575, cy: 38100 },
    });
    const absoluteRect = xlsx.anchorToRect(absoluteAnchor);
    assert.deepEqual(xlsx.rectToAbsoluteAnchor(absoluteRect), absoluteAnchor);
    assert.equal(xlsx.pixelsToEmu(xlsx.emuToPixels(28575)), 28575);
    const group = deepFreeze({ chX: 5, chY: 10, scaleX: 2, scaleY: 3, x: 7, y: 11 });
    assert.deepEqual(
      xlsx.applyGroupTransform(absoluteRect, group),
      xlsx.applyGroupTransform(absoluteRect, group),
    );
    assert.equal(JSON.stringify({ color, fill, palette }), xlsxBefore);

    process.stdout.write(JSON.stringify({
      docxExports: Object.keys(docx).sort(),
      status: "PASS",
      xlsxExports: Object.keys(xlsx).sort(),
    }));
  `;
  const result = runNode([
    "--experimental-strip-types",
    "--import",
    registerHookUrl,
    "--input-type=module",
    "--eval",
    childScript,
  ]);
  assert.equal(result.status, 0, result.stderr || result.stdout);
  const evidence = JSON.parse(result.stdout);
  assert.equal(evidence.status, "PASS");
  for (const name of [
    "applyRunStyle",
    "cloneDocModel",
    "layoutDocument",
    "normalizeDocModel",
    "replaceText",
    "updateParagraphText",
    "updateTableCellText",
  ]) {
    assert.ok(evidence.docxExports.includes(name), `missing DOCX core export ${name}`);
  }
  for (const name of [
    "anchorToRect",
    "buildChartSeriesFormula",
    "columnLabel",
    "parseChartSeriesFormula",
    "rangeToA1",
    "resolveWorkbookColor",
  ]) {
    assert.ok(evidence.xlsxExports.includes(name), `missing XLSX core export ${name}`);
  }
});

test("core dependency gate reports a clean recursive runtime graph", () => {
  const result = runNode(["scripts/ci/verify-core-boundaries.mjs"]);
  assert.equal(result.status, 0, result.stderr || result.stdout);
  const evidence = JSON.parse(result.stdout);
  assert.equal(evidence.status, "PASS");
  assert.deepEqual(evidence.violations, []);
  assert.deepEqual(evidence.externalRuntimeImports, []);
  assert.ok(evidence.runtimeFiles.includes("packages/docx-core/src/core.ts"));
  assert.ok(evidence.runtimeFiles.includes("packages/xlsx-core/src/core.ts"));
  assert.ok(evidence.runtimeFiles.includes("packages/docx-core/src/engine/clone.ts"));
  assert.ok(evidence.runtimeFiles.includes("packages/xlsx-core/src/colors.ts"));
});

test("published core and runtime subpaths preserve the intended declaration boundary", () => {
  const packages = ["docx-core", "xlsx-core"];
  const forbiddenCoreDeclaration =
    /@dukelib\/sheets-wasm|@extend-ai\/office-runtime|\b(?:AbortSignal|Blob|DOMParser|File|HTMLCanvasElement|Worker|XMLSerializer|fetch)\b/u;

  for (const packageName of packages) {
    const packageDir = resolve(rootDir, "packages", packageName);
    const manifest = JSON.parse(readFileSync(resolve(packageDir, "package.json"), "utf8"));
    assert.equal(manifest.exports["./core"].types, "./dist/core.d.ts");
    assert.equal(manifest.exports["./runtime"].types, "./dist/runtime.d.ts");

    const declarations = declarationGraph(resolve(packageDir, "dist/core.d.ts"));
    for (const [filePath, source] of declarations) {
      assert.doesNotMatch(
        withoutComments(source),
        forbiddenCoreDeclaration,
        `${packageName} pure declaration graph leaked through ${filePath}`,
      );
    }

    const runtimeDeclaration = readFileSync(resolve(packageDir, "dist/runtime.d.ts"), "utf8");
    assert.doesNotMatch(runtimeDeclaration, /export \* from ["']\.\/index\.js["']/u);
    assert.doesNotMatch(runtimeDeclaration, /setWasmSource/u);

    const wasmUrlDeclaration = readFileSync(resolve(packageDir, "dist/wasm-url.d.ts"), "utf8");
    assert.doesNotMatch(wasmUrlDeclaration, /from ["']\.\/index\.js["']/u);
  }
});
