import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import test from "node:test";

import ts from "../../packages/xlsx-core/node_modules/typescript/lib/typescript.js";

// The XLSX clipboard and search fixes are verified against their real sources:
// clipboard-text.ts has no imports and runs from a data: URL, while
// useXlsxSearch.ts is loaded with its "vue" and "@arcships/office-runtime"
// specifiers pointed at the real modules.
const ROOT = path.resolve(import.meta.dirname, "../..");
const CLIPBOARD_TEXT_PATH = path.join(ROOT, "packages/vue-xlsx/src/composables/clipboard-text.ts");
const CLIPBOARD_PATH = path.join(ROOT, "packages/vue-xlsx/src/composables/clipboard.ts");
const SEARCH_PATH = path.join(ROOT, "packages/vue-xlsx/src/composables/useXlsxSearch.ts");
const SEARCH_SESSION_PATH = path.join(ROOT, "packages/office-runtime/src/surface-search.ts");
const GRID_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/XlsxGrid.vue");
const CONTEXT_MENU_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/XlsxContextMenu.vue");

const requireFromDemo = createRequire(new URL("../../apps/demo/package.json", import.meta.url));

function transpile(source, fileName) {
  return ts.transpileModule(source, {
    fileName,
    compilerOptions: {
      module: ts.ModuleKind.ESNext,
      target: ts.ScriptTarget.ES2022,
    },
  }).outputText;
}

async function importTranspiled(source, fileName) {
  const output = transpile(source, fileName);
  return import(`data:text/javascript;base64,${Buffer.from(output, "utf8").toString("base64")}`);
}

function vueModuleUrl() {
  // vue's node ESM entry, so the generated module can import the real runtime.
  const manifest = requireFromDemo.resolve("vue/package.json");
  return pathToFileURL(path.join(path.dirname(manifest), "index.mjs")).href;
}

function scriptSource(filePath) {
  const text = readFileSync(filePath, "utf8");
  return [...text.matchAll(/<script(?:\s+setup)?[^>]*>([\s\S]*?)<\/script>/g)]
    .map((match) => match[1])
    .join("\n");
}

function parseTypeScript(filePath, text) {
  return ts.createSourceFile(filePath, text, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
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

function functionBody(ast, name) {
  const declaration = descendants(
    ast,
    (node) => ts.isFunctionDeclaration(node) && node.name?.text === name,
  )[0];
  return declaration?.body?.getText(ast) ?? "";
}

async function loadSearchModule() {
  const vueUrl = vueModuleUrl();
  const sessionUrl = pathToFileURL(SEARCH_SESSION_PATH).href;
  const source = readFileSync(SEARCH_PATH, "utf8");
  assert.notEqual(source, source.replace('from "vue"', `from ${JSON.stringify(vueUrl)}`), "useXlsxSearch must import the vue runtime");
  const rewritten = source
    .replace('from "vue"', `from ${JSON.stringify(vueUrl)}`)
    .replace('from "@arcships/office-runtime"', `from ${JSON.stringify(sessionUrl)}`);
  // Node strips the remaining (type-only) imports when loading the module.
  const directory = mkdtempSync(path.join(tmpdir(), "xlsx-search-"));
  const file = path.join(directory, "useXlsxSearch.generated.ts");
  writeFileSync(file, rewritten);
  try {
    return await import(pathToFileURL(file).href);
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
}

const clipboardText = await importTranspiled(
  readFileSync(CLIPBOARD_TEXT_PATH, "utf8"),
  CLIPBOARD_TEXT_PATH,
);
const vue = await import(vueModuleUrl());
const { useXlsxSearch } = await loadSearchModule();

test("TSV export quotes fields that would otherwise split the grid", () => {
  assert.equal(clipboardText.escapeTsvField("plain"), "plain");
  assert.equal(clipboardText.escapeTsvField("two\tcells"), "\"two\tcells\"");
  assert.equal(clipboardText.escapeTsvField("two\nlines"), "\"two\nlines\"");
  assert.equal(clipboardText.escapeTsvField("say \"hi\""), "\"say \"\"hi\"\"\"");

  assert.equal(
    clipboardText.serializeClipboardText([
      ["alpha", "beta\tgamma"],
      ["line1\nline2", "quote\"d"],
    ]),
    "alpha\t\"beta\tgamma\"\r\n\"line1\nline2\"\t\"quote\"\"d\"",
  );
});

test("parseClipboardText round-trips escaped TSV and quoted external cells", () => {
  const rows = [
    ["alpha", "beta\tgamma"],
    ["line1\nline2", "quote\"d"],
    ["", "trailing"],
  ];
  assert.deepEqual(
    clipboardText.parseClipboardText(clipboardText.serializeClipboardText(rows)),
    rows,
  );

  // Quoted TSV copied from Excel / Sheets keeps its embedded separators.
  assert.deepEqual(
    clipboardText.parseClipboardText('a\t"b\tc"\r\n"d\r\ne"\tf\r\n'),
    [["a", "b\tc"], ["d\ne", "f"]],
  );
  assert.deepEqual(
    clipboardText.parseClipboardText('"say ""hi"""\tplain'),
    [['say "hi"', "plain"]],
  );
  // A quote inside an unquoted field stays literal.
  assert.deepEqual(clipboardText.parseClipboardText('a"b\tc'), [['a"b', "c"]]);
});

test("parseClipboardText keeps the single trailing row and empty input shape", () => {
  assert.deepEqual(clipboardText.parseClipboardText("a\tb\n"), [["a", "b"]]);
  assert.deepEqual(clipboardText.parseClipboardText("a\tb\t"), [["a", "b", ""]]);
  assert.deepEqual(clipboardText.parseClipboardText(""), [[""]]);
  assert.deepEqual(clipboardText.parseClipboardText("a\rb\rc"), [["a"], ["b"], ["c"]]);
});

function fakeClipboardEnvironment(overrides = {}) {
  const calls = { write: [], writeText: [] };
  const items = [];
  const environment = {
    mimeType: "application/x-test-range+json",
    clipboard: {
      async write(entries) {
        calls.write.push(entries);
        if (overrides.rejectWrite) throw new Error("NotAllowedError");
      },
      async writeText(text) {
        calls.writeText.push(text);
        if (overrides.rejectWriteText) throw new Error("NotAllowedError");
      },
    },
    ClipboardItem: overrides.ClipboardItem === undefined
      ? class {
        constructor(data) {
          items.push(data);
          if (overrides.rejectConstruct) throw new Error("NotSupportedError");
          return { data };
        }
      }
      : overrides.ClipboardItem,
  };
  return { environment, calls, items };
}

const PAYLOAD = { html: "<table><tr><td>a</td></tr></table>", structured: '{"cells":[]}', text: "a" };

test("clipboard write falls back to text when the composite item is rejected", async () => {
  const { environment, calls } = fakeClipboardEnvironment({ rejectWrite: true });

  assert.equal(await clipboardText.writeClipboardPayload(PAYLOAD, environment), true);
  assert.equal(calls.write.length, 1);
  assert.deepEqual(calls.writeText, ["a"]);
});

test("clipboard write falls back to text when the browser cannot build an item", async () => {
  const { environment, calls } = fakeClipboardEnvironment({ rejectConstruct: true });

  assert.equal(await clipboardText.writeClipboardPayload(PAYLOAD, environment), true);
  assert.deepEqual(calls.writeText, ["a"]);

  const textOnly = fakeClipboardEnvironment({ ClipboardItem: null });
  assert.equal(await clipboardText.writeClipboardPayload(PAYLOAD, textOnly.environment), true);
  assert.deepEqual(textOnly.calls.writeText, ["a"]);
});

test("clipboard write sends every flavour when the browser accepts the item", async () => {
  const { environment, calls, items } = fakeClipboardEnvironment();

  assert.equal(await clipboardText.writeClipboardPayload(PAYLOAD, environment), true);
  assert.equal(calls.write.length, 1);
  assert.deepEqual(calls.writeText, []);
  assert.deepEqual(Object.keys(items[0]).sort(), [
    "application/x-test-range+json",
    "text/html",
    "text/plain",
  ]);
  assert.equal(await items[0]["application/x-test-range+json"].text(), PAYLOAD.structured);
  assert.equal(await items[0]["text/plain"].text(), "a");
});

test("clipboard write reports false instead of rejecting when every flavour is refused", async () => {
  const refusedWrite = fakeClipboardEnvironment({ rejectWrite: true, rejectWriteText: true });
  assert.equal(await clipboardText.writeClipboardPayload(PAYLOAD, refusedWrite.environment), false);

  const missingClipboard = { mimeType: "application/x-test-range+json", clipboard: null };
  assert.equal(await clipboardText.writeClipboardPayload(PAYLOAD, missingClipboard), false);

  const readOnlyClipboard = { mimeType: "application/x-test-range+json", clipboard: { writeText: undefined } };
  assert.equal(await clipboardText.writeClipboardPayload(PAYLOAD, readOnlyClipboard), false);
});

test("copy selection escapes the text flavour and reports the clipboard outcome", () => {
  const clipboardAst = parseTypeScript(CLIPBOARD_PATH, readFileSync(CLIPBOARD_PATH, "utf8"));
  assert.match(
    functionBody(clipboardAst, "getClipboardData"),
    /text:\s*serializeClipboardText\(rows\)/u,
  );
  assert.match(
    functionBody(clipboardAst, "getClipboardData"),
    /rows\.push\(textCells\)/u,
  );
  const copyBody = functionBody(clipboardAst, "copySelectionToClipboard");
  assert.match(copyBody, /writeClipboardPayload\(/u);
  assert.doesNotMatch(copyBody, /navigator\.clipboard\.writeText/u);

  for (const [componentPath, handler] of [
    [GRID_PATH, "onGridKeydown"],
    [CONTEXT_MENU_PATH, "onCopy"],
  ]) {
    const source = readFileSync(componentPath, "utf8");
    const ast = parseTypeScript(componentPath, scriptSource(componentPath));
    assert.match(source, /useClipboardCopyNotice\(/u, `${path.basename(componentPath)} must report copy failures`);
    assert.match(source, /<XlsxClipboardNotice /u, `${path.basename(componentPath)} must render the copy notice`);
    assert.match(functionBody(ast, handler), /copySelection\(\)/u);
    assert.doesNotMatch(
      functionBody(ast, handler),
      /props\.controller\.copySelectionToClipboard\(\)/u,
      `${handler} must await the copy result instead of dropping the promise`,
    );
  }
});

function searchFixture(values = ["Needle", "needle"]) {
  const sheet = {
    name: "Data",
    workbookSheetIndex: 0,
    minUsedRow: 0,
    maxUsedRow: 0,
    minUsedCol: 0,
    maxUsedCol: 1,
    visibleRows: [0],
    visibleCols: [0, 1],
    mergedRegions: [],
  };
  const worksheet = {
    getFormattedValueAt: (_row, col) => values[col] ?? "",
    getFormulaAt: () => "",
    getCalculatedValueAt: () => ({ is_empty: true, toString: () => "", free() {} }),
    free() {},
  };
  const revision = vue.ref(0);
  const activated = [];
  const controller = {
    activeSheetIndex: 0,
    isWorkerBacked: false,
    sheets: [sheet],
    workbook: { getSheet: () => worksheet },
    get revision() { return revision.value; },
    setActiveSheetIndex(index) { this.activeSheetIndex = index; },
    selectCell: (cell) => activated.push(cell),
  };
  return { controller, revision, activated };
}

test("useXlsxSearch keeps the case-insensitive default and honours caseSensitive", async () => {
  const insensitive = searchFixture();
  const insensitiveScope = vue.effectScope();
  const insensitiveSearch = insensitiveScope.run(() => useXlsxSearch(() => insensitive.controller));
  const loose = await insensitiveSearch.search("needle");
  assert.equal(loose.status, "ready");
  assert.equal(loose.matches.length, 2);
  insensitiveScope.stop();

  const sensitive = searchFixture();
  const sensitiveScope = vue.effectScope();
  const sensitiveSearch = sensitiveScope.run(() =>
    useXlsxSearch(() => sensitive.controller, { caseSensitive: true }));
  const exact = await sensitiveSearch.search("needle");
  assert.equal(exact.status, "ready");
  assert.deepEqual(exact.matches.map((match) => match.text), ["needle"]);
  assert.equal(exact.matches[0].address, "B1");

  const exactUpper = await sensitiveSearch.search("Needle");
  assert.deepEqual(exactUpper.matches.map((match) => match.text), ["Needle"]);
  sensitiveScope.stop();
});

test("useXlsxSearch accepts per-call whole-word and regular-expression options", async () => {
  const words = searchFixture(["cat", "category"]);
  const wordScope = vue.effectScope();
  const wordSearch = wordScope.run(() => useXlsxSearch(() => words.controller));
  const loose = await wordSearch.search("cat");
  assert.equal(loose.matches.length, 2);
  const whole = await wordSearch.search("cat", { wholeWord: true });
  assert.deepEqual(whole.matches.map((match) => match.text), ["cat"]);
  assert.equal(whole.matches[0].address, "A1");
  wordScope.stop();

  const versions = searchFixture(["1.2.3", "Version 4.5.6"]);
  const versionScope = vue.effectScope();
  const versionSearch = versionScope.run(() => useXlsxSearch(() => versions.controller));
  const regex = await versionSearch.search("\\d+\\.\\d+\\.\\d+", { regex: true });
  assert.deepEqual(regex.matches.map((match) => match.text), ["1.2.3", "Version 4.5.6"]);
  const exact = await versionSearch.search("VERSION\\s+4", { regex: true, caseSensitive: true });
  assert.equal(exact.matches.length, 0);
  const insensitive = await versionSearch.search("VERSION\\s+4", { regex: true });
  assert.equal(insensitive.matches.length, 1);
  assert.equal(insensitive.matches[0].address, "B1");
  await assert.rejects(
    versionSearch.search("(", { regex: true }),
    /Invalid regular expression/u,
  );
  versionScope.stop();
});

test("useXlsxSearch drops stale matches when the workbook revision changes", async () => {
  const fixture = searchFixture();
  const scope = vue.effectScope();
  const search = scope.run(() => useXlsxSearch(() => fixture.controller));

  const searched = await search.search("needle");
  assert.equal(searched.matches.length, 2);

  // An edit / undo / sort on the active sheet bumps the controller revision.
  fixture.revision.value += 1;
  await vue.nextTick();

  const cleared = search.getSearchState();
  assert.equal(cleared.status, "idle");
  assert.equal(cleared.query, "");
  assert.deepEqual(cleared.matches, []);
  assert.equal(cleared.activeIndex, -1);
  await assert.rejects(search.activateSearchMatch(0), RangeError);

  // Replacing the controller invalidates the session the same way.
  const replacement = vue.shallowRef(searchFixture().controller);
  const swappedScope = vue.effectScope();
  const swapped = swappedScope.run(() => useXlsxSearch(() => replacement.value));
  assert.equal((await swapped.search("needle")).matches.length, 2);
  replacement.value = searchFixture().controller;
  await vue.nextTick();
  assert.equal(swapped.getSearchState().status, "idle");

  scope.stop();
  swappedScope.stop();
});
