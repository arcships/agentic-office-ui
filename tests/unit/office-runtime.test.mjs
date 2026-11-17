import assert from "node:assert/strict";
import { readFileSync, readdirSync } from "node:fs";
import { test } from "node:test";
import { createRequire } from "node:module";

const packageRoot = new URL("../../packages/office-runtime/", import.meta.url);
const runtime = await import(new URL("dist/index.js", packageRoot).href);
const requireFromRuntime = createRequire(new URL("package.json", packageRoot));
const { strToU8, zipSync } = await import(
  new URL(`file://${requireFromRuntime.resolve("fflate")}`).href
);

const DOCX_CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const XLSX_CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>`;

const OOXML_VARIANTS = [
  ["xlsx", "xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", "xl/workbook.xml"],
  ["xlsx", "xlsb", "application/vnd.ms-excel.sheet.binary.macroEnabled.main", "xl/workbook.bin"],
  ["xlsx", "xlsm", "application/vnd.ms-excel.sheet.macroEnabled.main+xml", "xl/workbook.xml"],
  ["xlsx", "xltx", "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml", "xl/workbook.xml"],
  ["xlsx", "xltm", "application/vnd.ms-excel.template.macroEnabled.main+xml", "xl/workbook.xml"],
  ["docx", "docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "word/document.xml"],
  ["docx", "docm", "application/vnd.ms-word.document.macroEnabled.main+xml", "word/document.xml"],
  ["docx", "dotx", "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml", "word/document.xml"],
  ["docx", "dotm", "application/vnd.ms-word.template.macroEnabledTemplate.main+xml", "word/document.xml"],
  ["pptx", "pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml", "ppt/presentation.xml"],
  ["pptx", "pptm", "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml", "ppt/presentation.xml"],
  ["pptx", "ppsx", "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml", "ppt/presentation.xml"],
  ["pptx", "ppsm", "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml", "ppt/presentation.xml"],
  ["pptx", "potx", "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml", "ppt/presentation.xml"],
  ["pptx", "potm", "application/vnd.ms-powerpoint.template.macroEnabled.main+xml", "ppt/presentation.xml"],
];

function makeArchive(entries, contentTypes = DOCX_CONTENT_TYPES) {
  return zipSync({
    "[Content_Types].xml": strToU8(contentTypes),
    ...Object.fromEntries(Object.entries(entries).map(([name, value]) => [name, strToU8(value)])),
  });
}

function sourceFiles(directory) {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const url = new URL(`${entry.name}${entry.isDirectory() ? "/" : ""}`, directory);
    return entry.isDirectory() ? sourceFiles(url) : entry.name.endsWith(".ts") ? [url] : [];
  });
}

function createCountedAbortSignal() {
  const controller = new AbortController();
  const signal = controller.signal;
  const originalAdd = signal.addEventListener.bind(signal);
  const originalRemove = signal.removeEventListener.bind(signal);
  const activeAbortListeners = new Set();
  let addCalls = 0;
  let removeCalls = 0;

  signal.addEventListener = (type, listener, options) => {
    if (type === "abort") {
      addCalls += 1;
      activeAbortListeners.add(listener);
    }
    return originalAdd(type, listener, options);
  };
  signal.removeEventListener = (type, listener, options) => {
    if (type === "abort") {
      removeCalls += 1;
      activeAbortListeners.delete(listener);
    }
    return originalRemove(type, listener, options);
  };

  return {
    controller,
    signal,
    counts: () => ({
      addCalls,
      removeCalls,
      active: activeAbortListeners.size,
    }),
  };
}

function assertStale(task) {
  assert.equal(task.isCurrent(), false);
  assert.throws(
    () => task.assertCurrent(),
    (error) => error?.name === "OfficeLoadError" && error?.code === "STALE_RESULT",
  );
}

test("office-runtime is private, framework-free and has no UI DOM dependency", () => {
  const manifest = JSON.parse(readFileSync(new URL("package.json", packageRoot), "utf8"));
  assert.equal(manifest.private, true);
  assert.equal(manifest.publishConfig, undefined);
  assert.equal(JSON.stringify(manifest).includes("vue"), false);

  for (const file of sourceFiles(new URL("src/", packageRoot))) {
    const source = readFileSync(file, "utf8");
    assert.doesNotMatch(source, /(?:from\s+["']vue["']|\.vue["'])/);
    assert.doesNotMatch(source, /\b(?:window|document|HTMLElement|HTMLCanvasElement|ImageBitmap)\b/);
  }
});

test("surface zoom math clamps the shared range and normalizes wheel deltas", () => {
  assert.equal(runtime.clampSurfaceZoom(Number.NaN), 1);
  assert.equal(runtime.clampSurfaceZoom(0.1), 0.5);
  assert.equal(runtime.clampSurfaceZoom(4), 2);
  assert.equal(runtime.clampSurfaceZoom(1.234), 1.23);

  assert.equal(runtime.normalizeWheelDelta(2, 0), 2);
  assert.equal(runtime.normalizeWheelDelta(2, 1), 32);
  assert.equal(runtime.normalizeWheelDelta(2, 2), 1600);
  assert.ok(runtime.nextSurfaceZoom(1, -20) > 1);
  assert.ok(runtime.nextSurfaceZoom(1, 20) < 1);
  assert.equal(runtime.nextSurfaceZoom(1, -0.1), 1.01);
  assert.equal(runtime.nextSurfaceZoom(1, 0.1), 0.99);
  assert.equal(runtime.nextSurfaceZoom(2, -100), 2);
  assert.equal(runtime.nextSurfaceZoom(0.5, 100), 0.5);
});

test("surface search keeps only the latest query and activates its first match", async () => {
  const searches = new Map();
  const activations = [];
  const session = runtime.createSurfaceSearchSession({
    search(query, { signal }) {
      return new Promise((resolve, reject) => {
        const onAbort = () => {
          const error = new Error("replaced");
          error.name = "AbortError";
          reject(error);
        };
        signal.addEventListener("abort", onAbort, { once: true });
        searches.set(query, (matches) => {
          signal.removeEventListener("abort", onAbort);
          resolve(matches);
        });
      });
    },
    activate(match, index) {
      activations.push([match, index]);
    },
  });

  const stale = session.search("old");
  const latest = session.search("new");
  await assert.rejects(stale, (error) => error?.name === "AbortError");
  searches.get("new")(["new:1", "new:2"]);

  const state = await latest;
  assert.equal(state.status, "ready");
  assert.equal(state.query, "new");
  assert.deepEqual(state.matches, ["new:1", "new:2"]);
  assert.equal(state.activeIndex, 0);
  assert.deepEqual(activations, [["new:1", 0]]);
  session.dispose();
});

test("surface search navigation wraps and clear resets the public state", async () => {
  const activations = [];
  let clearCount = 0;
  const session = runtime.createSurfaceSearchSession({
    search: () => ["a", "b", "c"],
    activate: (match, index) => activations.push([match, index]),
    clear: () => { clearCount += 1; },
  });

  await session.search("x");
  await session.searchPrevious();
  assert.equal(session.getSearchState().activeIndex, 2);
  await session.searchNext();
  assert.equal(session.getSearchState().activeIndex, 0);
  session.clearSearch();
  assert.deepEqual(session.getSearchState(), {
    status: "idle",
    query: "",
    matches: [],
    activeIndex: -1,
  });
  assert.deepEqual(activations, [["a", 0], ["c", 2], ["a", 0]]);
  assert.equal(clearCount, 1);
  session.dispose();
});

test("surface search resolves superseded activation with the latest state instead of rejecting", async () => {
  let firstActivationStarted;
  const activationStarted = new Promise((resolve) => { firstActivationStarted = resolve; });
  const activations = [];
  const session = runtime.createSurfaceSearchSession({
    search: (query) => [query],
    activate(match, index, { signal }) {
      activations.push([match, index]);
      if (activations.length > 1) return;
      firstActivationStarted();
      return new Promise((_resolve, reject) => {
        signal.addEventListener("abort", () => {
          const error = new Error("activation replaced");
          error.name = "AbortError";
          reject(error);
        }, { once: true });
      });
    },
  });

  const first = session.search("first");
  await activationStarted;
  const second = session.search("second");

  const firstState = await first;
  assert.notEqual(firstState, undefined, "superseded search must resolve with state");
  assert.equal(firstState.status !== "error", true);
  assert.equal(firstState.query, "second");

  const secondState = await second;
  assert.equal(secondState.status, "ready");
  assert.deepEqual(secondState.matches, ["second"]);
  assert.equal(secondState.activeIndex, 0);
  assert.deepEqual(activations, [["first", 0], ["second", 0]]);
  session.dispose();
});

test("surface search reports activation failure through state without rejecting the host", async () => {
  const session = runtime.createSurfaceSearchSession({
    search: () => ["virtualized"],
    activate() { throw new Error("match page is not mounted"); },
  });

  const state = await session.search("virtualized");
  assert.equal(state.status, "error");
  assert.equal(state.query, "virtualized");
  assert.deepEqual(state.matches, ["virtualized"]);
  assert.equal(state.error.code, "ACTIVATION_FAILED");
  assert.match(state.error.message, /not mounted/u);
  assert.equal(session.getSearchState(), state);
  session.dispose();
});

test("archive budgets validate central metadata and actual streamed extraction at exact boundaries", () => {
  const archive = makeArchive({
    "word/document.xml": "<w:document xmlns:w=\"urn:w\"><w:body><w:p>ok</w:p></w:body></w:document>",
  });
  const inspected = runtime.inspectOfficeArchive(archive, {});
  const exact = {
    maxInputBytes: archive.byteLength,
    maxArchiveEntries: inspected.entryCount,
    maxUncompressedBytes: inspected.uncompressedBytes,
    maxSingleEntryBytes: Math.max(...inspected.entries.map((entry) => entry.uncompressedBytes)),
    maxCompressionRatio: inspected.maxCompressionRatio,
    maxArchivePathLength: 1024,
    maxXmlBytes: inspected.uncompressedBytes,
    maxSingleXmlBytes: Math.max(...inspected.entries.map((entry) => entry.uncompressedBytes)),
    maxXmlDepth: 8,
    maxXmlAttributeBytes: 1024,
    maxTextNodeBytes: 1024,
    maxRelationships: 10,
  };
  const validated = runtime.validateOfficeArchive(archive, exact, { format: "docx" });
  assert.equal(validated.entryCount, 2);
  assert.equal(validated.uncompressedBytes, inspected.uncompressedBytes);

  for (const [limit, value] of [
    ["maxInputBytes", exact.maxInputBytes],
    ["maxArchiveEntries", exact.maxArchiveEntries],
    ["maxUncompressedBytes", exact.maxUncompressedBytes],
    ["maxSingleEntryBytes", exact.maxSingleEntryBytes],
    ["maxCompressionRatio", exact.maxCompressionRatio],
  ]) {
    assert.throws(
      () => runtime.validateOfficeArchive(archive, { ...exact, [limit]: value - 1 }, { format: "docx" }),
      (error) => error?.code === "LIMIT_EXCEEDED" && error?.limit === limit,
    );
  }
});

test("archive budgets reject unsafe paths, unsafe XML, wrong MIME and aggregate XLSX counts", () => {
  const baseLimits = {
    maxArchiveEntries: 20,
    maxUncompressedBytes: 1024 * 1024,
    maxSingleEntryBytes: 1024 * 1024,
    maxCompressionRatio: 200,
    maxArchivePathLength: 1024,
    maxXmlBytes: 1024 * 1024,
    maxSingleXmlBytes: 1024 * 1024,
    maxXmlDepth: 20,
    maxXmlAttributeBytes: 1024,
    maxTextNodeBytes: 1024,
    maxRelationships: 20,
  };
  assert.throws(
    () => runtime.validateOfficeArchive(
      makeArchive({ "../evil.xml": "<evil/>" }),
      baseLimits,
      { format: "docx" },
    ),
    (error) => error?.code === "INVALID_SOURCE" && error?.phase === "archive",
  );
  assert.throws(
    () => runtime.validateOfficeArchive(
      makeArchive({ "word/document.xml": "<!DOCTYPE x [<!ENTITY y 'boom'>]><x>&y;</x>" }),
      baseLimits,
      { format: "docx" },
    ),
    (error) => error?.code === "INVALID_SOURCE" && error?.phase === "xml",
  );
  assert.throws(
    () => runtime.validateOfficeArchive(
      makeArchive({ "word/document.xml": "<x/>" }, XLSX_CONTENT_TYPES),
      baseLimits,
      { format: "docx" },
    ),
    (error) => error?.code === "INVALID_SOURCE" && error?.phase === "archive",
  );

  const xlsx = makeArchive({
    "xl/workbook.xml": "<workbook><sheets><sheet name=\"A\"/><sheet name=\"B\"/></sheets></workbook>",
    "xl/sharedStrings.xml": "<sst><si><t>A</t></si><si><t>B</t></si></sst>",
    "xl/worksheets/sheet1.xml": "<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><f>1+1</f></c></row></sheetData></worksheet>",
  }, XLSX_CONTENT_TYPES);
  for (const [limit, allowed] of [
    ["maxWorksheets", 1],
    ["maxSharedStrings", 1],
    ["maxFormulaCount", 0.5],
  ]) {
    assert.throws(
      () => runtime.validateOfficeArchive(xlsx, { ...baseLimits, [limit]: allowed }, { format: "xlsx" }),
      (error) => error?.code === "LIMIT_EXCEEDED" && error?.limit === limit,
    );
  }
});

test("source format detection and archive validation recognize supported OOXML families", () => {
  for (const [family, format, contentType, mainPart] of OOXML_VARIANTS) {
    const contentTypes = `<?xml version="1.0" encoding="UTF-8"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Override PartName="/${mainPart}" ContentType="${contentType}"/>
      </Types>`;
    const archive = makeArchive({ [mainPart]: "<root/>" }, contentTypes);
    const detected = runtime.detectSourceFormat({
      bytes: archive,
      fileName: `sample.${format}`,
      ooxmlContentTypes: contentTypes,
    });
    assert.equal(detected?.family, family);
    assert.equal(detected?.format, format);
    assert.equal(detected?.confidence, "high");

    const validated = runtime.validateOfficeArchive(archive, {}, { format: family });
    assert.equal(validated.sourceFormat, format);
  }
});

test("source format detection uses signatures for PDF and legacy XLS", () => {
  const pdf = runtime.detectSourceFormat({
    bytes: new TextEncoder().encode("%PDF-1.7"),
    fileName: "wrong.docx",
  });
  assert.deepEqual(
    { family: pdf?.family, format: pdf?.format, confidence: pdf?.confidence },
    { family: "pdf", format: "pdf", confidence: "high" },
  );

  const xls = runtime.detectSourceFormat({
    bytes: Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]),
    fileName: "legacy.xls",
  });
  assert.deepEqual(
    { family: xls?.family, format: xls?.format, confidence: xls?.confidence },
    { family: "xlsx", format: "xls", confidence: "medium" },
  );
});

test("PPTX archive validation accepts the real playback fixture", () => {
  const fixture = readFileSync(new URL("../fixtures/pptx/playback-controlled.pptx", import.meta.url));
  const validated = runtime.validateOfficeArchive(fixture, {}, { format: "pptx" });
  assert.equal(validated.sourceFormat, "pptx");
  assert.ok(validated.entryCount > 0);
});

test("XLSX archive budgets count elements by local name when namespace prefixes are present", () => {
  const limits = {
    maxArchiveEntries: 20,
    maxUncompressedBytes: 1024 * 1024,
    maxSingleEntryBytes: 1024 * 1024,
    maxCompressionRatio: 200,
    maxArchivePathLength: 1024,
    maxXmlBytes: 1024 * 1024,
    maxSingleXmlBytes: 1024 * 1024,
    maxXmlDepth: 20,
    maxXmlAttributeBytes: 1024,
    maxTextNodeBytes: 1024,
    maxRelationships: 10,
    maxWorksheets: 10,
    maxSharedStrings: 10,
    maxFormulaCount: 10,
    maxWorksheetRows: 10,
    maxWorksheetColumns: 10,
  };
  const worksheetXml = `<x:worksheet xmlns:x="urn:x"><x:sheetData>
    <x:row r="1"/><x:row r="2"/>
    <x:row r="3"><x:c r="B3"><x:f>1</x:f></x:c><x:c r="C3"><x:f>2</x:f></x:c></x:row>
  </x:sheetData></x:worksheet>`;
  const xlsx = makeArchive({
    "xl/workbook.xml": `<x:workbook xmlns:x="urn:x"><x:sheets><x:sheet/><x:sheet/></x:sheets></x:workbook>`,
    "xl/sharedStrings.xml": `<x:sst xmlns:x="urn:x"><x:si/><x:si/></x:sst>`,
    "xl/worksheets/sheet1.xml": worksheetXml,
    "xl/_rels/workbook.xml.rels": `<r:Relationships xmlns:r="urn:r"><r:Relationship/><r:Relationship/></r:Relationships>`,
  }, XLSX_CONTENT_TYPES);

  const worksheet = runtime.validateOfficeXmlEntry(
    "xl/worksheets/sheet1.xml",
    new TextEncoder().encode(worksheetXml),
    limits,
  );
  assert.deepEqual(worksheet, {
    relationships: 0,
    rows: 3,
    maxRow: 3,
    maxColumn: 3,
    worksheets: 0,
    sharedStrings: 0,
    formulas: 2,
  });
  const validated = runtime.validateOfficeArchive(xlsx, limits, { format: "xlsx" });
  assert.equal(validated.relationships, 2);
  assert.equal(validated.worksheets, 2);
  assert.equal(validated.sharedStrings, 2);
  assert.equal(validated.formulas, 2);

  for (const [limit, allowed] of [
    ["maxRelationships", 1],
    ["maxWorksheets", 1],
    ["maxSharedStrings", 1],
    ["maxFormulaCount", 1],
    ["maxWorksheetRows", 2],
    ["maxWorksheetColumns", 2],
  ]) {
    assert.throws(
      () => runtime.validateOfficeArchive(xlsx, { ...limits, [limit]: allowed }, { format: "xlsx" }),
      (error) => error?.code === "LIMIT_EXCEEDED" && error?.limit === limit,
    );
  }
});

test("archive validation observes an already-aborted signal before extraction", () => {
  const controller = new AbortController();
  controller.abort();
  assert.throws(
    () => runtime.validateOfficeArchive(makeArchive({ "word/document.xml": "<x/>" }), {}, {
      format: "docx",
      signal: controller.signal,
    }),
    (error) => error?.code === "ABORTED",
  );
});

test("load context snapshots source, resources, limits and stable task identity", () => {
  const resources = { wasmUrl: "https://assets.example/docx.wasm", workerUrl: "https://assets.example/docx-worker.js" };
  const limits = { maxInputBytes: 1024 };
  const sequence = runtime.createOfficeTaskSequence("runtime-a");
  const context = runtime.createOfficeLoadContext({
    runtimeId: "runtime-a",
    taskId: sequence.next(),
    source: { kind: "bytes", bytes: new ArrayBuffer(8), name: "sample.docx" },
    resources,
    limits,
    signal: new AbortController().signal,
  });

  resources.wasmUrl = "https://mutated.invalid/wasm";
  limits.maxInputBytes = 1;
  assert.equal(context.runtimeId, "runtime-a");
  assert.equal(context.taskId, "runtime-a:1");
  assert.equal(context.sourceKind, "bytes");
  assert.equal(context.resources.wasmUrl, "https://assets.example/docx.wasm");
  assert.equal(context.limits.maxInputBytes, 1024);
  assert.equal(Object.isFrozen(context), true);
  assert.equal(Object.isFrozen(context.resources), true);
  assert.equal(Object.isFrozen(context.limits), true);

  const coordinator = runtime.createLatestTaskCoordinator(sequence);
  const task = coordinator.start(
    { kind: "bytes", bytes: new ArrayBuffer(4) },
    {
      limits: { maxInputBytes: 4 },
      limitDefaults: { maxInputBytes: 100, maxArchiveEntries: 10 },
      limitHardCaps: { maxInputBytes: 50, maxUncompressedBytes: 7 },
    },
  );
  assert.deepEqual(task.context.limits, {
    maxInputBytes: 4,
    maxArchiveEntries: 10,
    maxUncompressedBytes: 7,
  });
  coordinator.dispose();
  sequence.dispose();
});

test("task sequences are instance-owned and shared by loaders of one runtime", () => {
  const sequenceA = runtime.createOfficeTaskSequence("runtime-a");
  const sequenceB = runtime.createOfficeTaskSequence("runtime-b");
  const loaderA = runtime.createLatestTaskCoordinator(sequenceA);
  const loaderB = runtime.createLatestTaskCoordinator(sequenceA);
  const otherRuntimeLoader = runtime.createLatestTaskCoordinator(sequenceB);

  const first = loaderA.start({ kind: "bytes", bytes: new ArrayBuffer(1) });
  const second = loaderB.start({ kind: "bytes", bytes: new ArrayBuffer(1) });
  const isolated = otherRuntimeLoader.start({ kind: "bytes", bytes: new ArrayBuffer(1) });
  assert.equal(first.context.taskId, "runtime-a:1");
  assert.equal(second.context.taskId, "runtime-a:2");
  assert.equal(isolated.context.taskId, "runtime-b:1");
  assert.equal(first.isCurrent(), true);
  assert.equal(second.isCurrent(), true);

  const replacement = loaderA.start({ kind: "bytes", bytes: new ArrayBuffer(1) });
  assert.equal(replacement.context.taskId, "runtime-a:3");
  assert.equal(first.signal.aborted, true);
  assert.equal(first.isCurrent(), false);
  assert.throws(() => first.assertCurrent(), (error) => error?.code === "STALE_RESULT");

  loaderA.dispose();
  loaderB.dispose();
  otherRuntimeLoader.dispose();
  sequenceA.dispose();
  sequenceB.dispose();
});

test("external and internal cancellation combine without mutating the caller signal", () => {
  const sequence = runtime.createOfficeTaskSequence("cancel-runtime");
  const loader = runtime.createLatestTaskCoordinator(sequence);
  const external = new AbortController();
  const task = loader.start(
    { kind: "bytes", bytes: new ArrayBuffer(1) },
    { signal: external.signal },
  );
  loader.cancel();
  assert.equal(task.signal.aborted, true);
  assert.equal(external.signal.aborted, false);

  const alreadyAborted = new AbortController();
  alreadyAborted.abort();
  const cancelledAtStart = loader.start(
    { kind: "bytes", bytes: new ArrayBuffer(1) },
    { signal: alreadyAborted.signal },
  );
  assert.equal(cancelledAtStart.signal.aborted, true);
  loader.dispose();
  assert.throws(
    () => loader.start({ kind: "bytes", bytes: new ArrayBuffer(1) }),
    (error) => error?.code === "RUNTIME_DISPOSED",
  );
  sequence.dispose();
});

test("finished tasks stop being current and remove the external abort listener once", () => {
  const sequence = runtime.createOfficeTaskSequence("finish-lifecycle");
  const coordinator = runtime.createLatestTaskCoordinator(sequence);
  const external = createCountedAbortSignal();
  const task = coordinator.start(
    { kind: "bytes", bytes: new ArrayBuffer(1) },
    { signal: external.signal },
  );

  assert.equal(task.isCurrent(), true);
  assert.deepEqual(external.counts(), { addCalls: 1, removeCalls: 0, active: 1 });
  task.finish();
  assertStale(task);
  assert.deepEqual(external.counts(), { addCalls: 1, removeCalls: 1, active: 0 });

  task.finish();
  coordinator.cancel();
  coordinator.dispose();
  coordinator.dispose();
  sequence.dispose();
  sequence.dispose();
  assert.deepEqual(external.counts(), { addCalls: 1, removeCalls: 1, active: 0 });
});

test("replacing a task removes its external abort listener and makes it stale", () => {
  const sequence = runtime.createOfficeTaskSequence("replace-lifecycle");
  const coordinator = runtime.createLatestTaskCoordinator(sequence);
  const oldExternal = createCountedAbortSignal();
  const nextExternal = createCountedAbortSignal();
  const oldTask = coordinator.start(
    { kind: "bytes", bytes: new ArrayBuffer(1) },
    { signal: oldExternal.signal },
  );
  const nextTask = coordinator.start(
    { kind: "bytes", bytes: new ArrayBuffer(1) },
    { signal: nextExternal.signal },
  );

  assert.equal(oldTask.signal.aborted, true);
  assertStale(oldTask);
  assert.deepEqual(oldExternal.counts(), { addCalls: 1, removeCalls: 1, active: 0 });
  assert.equal(nextTask.isCurrent(), true);
  assert.deepEqual(nextExternal.counts(), { addCalls: 1, removeCalls: 0, active: 1 });

  oldTask.finish();
  oldTask.finish();
  nextTask.finish();
  nextTask.finish();
  coordinator.dispose();
  sequence.dispose();
  assert.deepEqual(oldExternal.counts(), { addCalls: 1, removeCalls: 1, active: 0 });
  assert.deepEqual(nextExternal.counts(), { addCalls: 1, removeCalls: 1, active: 0 });
});

test("cancel is idempotent and removes the external abort listener once", () => {
  const sequence = runtime.createOfficeTaskSequence("cancel-lifecycle");
  const coordinator = runtime.createLatestTaskCoordinator(sequence);
  const external = createCountedAbortSignal();
  const task = coordinator.start(
    { kind: "bytes", bytes: new ArrayBuffer(1) },
    { signal: external.signal },
  );

  coordinator.cancel();
  coordinator.cancel();
  assert.equal(task.signal.aborted, true);
  assert.equal(external.signal.aborted, false);
  assertStale(task);
  assert.deepEqual(external.counts(), { addCalls: 1, removeCalls: 1, active: 0 });

  task.finish();
  task.finish();
  coordinator.dispose();
  sequence.dispose();
  assert.deepEqual(external.counts(), { addCalls: 1, removeCalls: 1, active: 0 });
});

test("dispose is idempotent and removes the external abort listener once", () => {
  const sequence = runtime.createOfficeTaskSequence("dispose-lifecycle");
  const coordinator = runtime.createLatestTaskCoordinator(sequence);
  const external = createCountedAbortSignal();
  const task = coordinator.start(
    { kind: "bytes", bytes: new ArrayBuffer(1) },
    { signal: external.signal },
  );

  coordinator.dispose();
  coordinator.dispose();
  assert.equal(task.signal.aborted, true);
  assertStale(task);
  assert.deepEqual(external.counts(), { addCalls: 1, removeCalls: 1, active: 0 });
  assert.throws(
    () => coordinator.start({ kind: "bytes", bytes: new ArrayBuffer(1) }),
    (error) => error?.code === "RUNTIME_DISPOSED",
  );

  task.finish();
  task.finish();
  sequence.dispose();
  sequence.dispose();
  assert.deepEqual(external.counts(), { addCalls: 1, removeCalls: 1, active: 0 });
});

test("URL policy rejects dangerous sources before fetch and strips secrets", async () => {
  let requestCount = 0;
  const policy = {
    enabled: true,
    baseUrl: "https://app.example/viewer/",
    allowRelativeUrl: true,
    allowedProtocols: ["https:"],
    allowedOrigins: ["https://app.example"],
    fetch: async () => {
      requestCount += 1;
      throw new Error("must not fetch");
    },
  };

  await assert.rejects(
    runtime.loadOfficeSource({ kind: "url", url: "javascript:alert(1)" }, { urlPolicy: policy }),
    (error) => error?.name === "OfficeLoadError" && error?.code === "SOURCE_NOT_ALLOWED",
  );
  await assert.rejects(
    runtime.loadOfficeSource({ kind: "url", url: "https://other.example/a.xlsx?token=secret#x" }, { urlPolicy: policy }),
    (error) => error?.code === "SOURCE_NOT_ALLOWED" && !JSON.stringify(error).includes("secret"),
  );
  assert.equal(requestCount, 0);
  assert.equal(
    runtime.sanitizeOfficeUrl("https://app.example/a.xlsx?token=secret#x"),
    "https://app.example/a.xlsx",
  );
});

test("link policy accepts safe hrefs and rejects executable schemes", () => {
  for (const target of [
    "https://example.com/path?q=1#frag",
    "http://example.com/path",
    "HTTPS://example.com/path",
    "mailto:person@example.com",
    "#bookmark",
    "/relative/path",
    "./relative/path",
    "relative/path?q=1",
    "//example.com/path",
  ]) {
    assert.equal(runtime.sanitizeOfficeLink(target), target);
    assert.equal(runtime.isSafeOfficeLink(target), true);
  }

  for (const target of [
    "javascript:alert(1)",
    "JaVaScRiPt:alert(1)",
    " javascript:alert(1)",
    "java\nscript:alert(1)",
    "java\tscript:alert(1)",
    "data:text/html,<script>alert(1)</script>",
    "vbscript:msgbox(1)",
    "file:///etc/passwd",
    "tel:+10000000000",
    "https://",
    "mailto:",
    "",
  ]) {
    assert.equal(runtime.sanitizeOfficeLink(target), undefined);
    assert.equal(runtime.isSafeOfficeLink(target), false);
  }

  assert.deepEqual([...runtime.OFFICE_LINK_PROTOCOLS], ["http:", "https:", "mailto:"]);
});

test("source loading uses controlled fetch and enforces the input byte limit", async () => {
  const calls = [];
  const bytes = new Uint8Array([0x50, 0x4b, 0x03, 0x04]);
  const policy = {
    enabled: true,
    baseUrl: "https://app.example/viewer/",
    allowRelativeUrl: true,
    allowedProtocols: ["https:"],
    allowedOrigins: ["https://app.example"],
    fetch: async (url, init) => {
      calls.push({ url, init });
      return {
        ok: true,
        status: 200,
        url,
        headers: { get: (name) => name.toLowerCase() === "content-length" ? String(bytes.byteLength) : null },
        arrayBuffer: async () => bytes.buffer.slice(0),
      };
    },
  };

  const loaded = await runtime.loadOfficeSource(
    { kind: "url", url: "book.xlsx?token=secret" },
    { urlPolicy: policy, limits: { maxInputBytes: bytes.byteLength } },
  );
  assert.deepEqual([...new Uint8Array(loaded.buffer)], [...bytes]);
  assert.equal(loaded.sourceKind, "url");
  assert.equal(loaded.resolvedUrl, "https://app.example/viewer/book.xlsx");
  assert.equal(calls[0].init.redirect, "follow");
  assert.equal(calls[0].init.credentials, "omit");

  await assert.rejects(
    runtime.loadOfficeSource(
      { kind: "url", url: "book.xlsx" },
      { urlPolicy: policy, limits: { maxInputBytes: bytes.byteLength - 1 } },
    ),
    (error) => error?.code === "LIMIT_EXCEEDED" && error?.limit === "maxInputBytes",
  );
});

test("limit resolution keeps hard caps without defaults and rejects invalid values", () => {
  const tightened = runtime.resolveOfficeLimits(
    { maxInputBytes: 100, maxXmlDepth: 10 },
    { maxInputBytes: 50, maxFormulaCount: 3 },
    { maxInputBytes: 80, maxFormulaCount: 200, maxDocxPages: 500 },
  );
  assert.equal(tightened.maxInputBytes, 50);
  assert.equal(tightened.maxXmlDepth, 10);
  assert.equal(tightened.maxFormulaCount, 3);
  assert.equal(tightened.maxDocxPages, 500);
  assert.equal(Object.isFrozen(tightened), true);

  const capped = runtime.resolveOfficeLimits(
    { maxInputBytes: 100 },
    { maxInputBytes: 500 },
    { maxInputBytes: 80, maxArchiveEntries: 20 },
  );
  assert.equal(capped.maxInputBytes, 80);
  assert.equal(capped.maxArchiveEntries, 20);

  for (const invalid of [0, -1, Number.NaN, Number.POSITIVE_INFINITY]) {
    assert.throws(
      () => runtime.resolveOfficeLimits(
        { maxInputBytes: 10 },
        { maxInputBytes: invalid },
        { maxInputBytes: 20 },
      ),
      (error) => error?.name === "OfficeLoadError"
        && error?.code === "INVALID_ARGUMENT"
        && error?.phase === "config"
        && error?.limit === "maxInputBytes",
    );
  }
});

test("non-positive maxInputBytes is a configuration error instead of an unlimited guard", async () => {
  let fetches = 0;
  const policy = {
    enabled: true,
    baseUrl: "https://app.example/viewer/",
    allowRelativeUrl: true,
    allowedProtocols: ["https:"],
    allowedOrigins: ["https://app.example"],
    fetch: async () => {
      fetches += 1;
      throw new Error("must not fetch");
    },
  };

  for (const invalid of [0, -1, Number.NaN, Number.POSITIVE_INFINITY]) {
    await assert.rejects(
      runtime.loadOfficeSource(
        { kind: "bytes", bytes: new ArrayBuffer(4) },
        { limits: { maxInputBytes: invalid } },
      ),
      (error) => error?.code === "INVALID_ARGUMENT"
        && error?.phase === "config"
        && error?.limit === "maxInputBytes",
    );
    await assert.rejects(
      runtime.loadOfficeSource(
        { kind: "url", url: "book.xlsx" },
        { urlPolicy: policy, limits: { maxInputBytes: invalid } },
      ),
      (error) => error?.code === "INVALID_ARGUMENT" && error?.limit === "maxInputBytes",
    );
  }
  assert.equal(fetches, 0);
});

test("URL loading cancels a chunked response as soon as maxInputBytes is exceeded", async () => {
  const chunks = [
    new Uint8Array([0x50, 0x4b, 0x03, 0x04]),
    new Uint8Array([1, 2, 3, 4]),
    new Uint8Array([5, 6, 7, 8]),
  ];
  let reads = 0;
  let cancelCalls = 0;
  let releaseCalls = 0;
  let arrayBufferCalls = 0;
  const policy = {
    enabled: true,
    baseUrl: "https://app.example/viewer/",
    allowRelativeUrl: true,
    allowedProtocols: ["https:"],
    allowedOrigins: ["https://app.example"],
    fetch: async (url) => ({
      ok: true,
      status: 200,
      url,
      headers: { get: () => null },
      body: {
        getReader: () => ({
          read: async () => {
            const value = chunks[reads];
            reads += 1;
            return value ? { done: false, value } : { done: true };
          },
          cancel: async () => { cancelCalls += 1; },
          releaseLock: () => { releaseCalls += 1; },
        }),
      },
      arrayBuffer: async () => {
        arrayBufferCalls += 1;
        return new ArrayBuffer(0);
      },
    }),
  };

  await assert.rejects(
    runtime.loadOfficeSource(
      { kind: "url", url: "book.xlsx" },
      { urlPolicy: policy, limits: { maxInputBytes: 5 } },
    ),
    (error) => error?.code === "LIMIT_EXCEEDED"
      && error?.phase === "input"
      && error?.limit === "maxInputBytes"
      && error?.actual === 8
      && error?.allowed === 5,
  );
  assert.equal(reads, 2);
  assert.equal(cancelCalls, 1);
  assert.equal(releaseCalls, 1);
  assert.equal(arrayBufferCalls, 0);
});

test("URL loading assembles a chunked response that stays within maxInputBytes", async () => {
  const chunks = [new Uint8Array([0x50, 0x4b]), new Uint8Array([0x03, 0x04])];
  let reads = 0;
  let cancelCalls = 0;
  const policy = {
    enabled: true,
    baseUrl: "https://app.example/viewer/",
    allowRelativeUrl: true,
    allowedProtocols: ["https:"],
    allowedOrigins: ["https://app.example"],
    fetch: async (url) => ({
      ok: true,
      status: 200,
      url,
      headers: { get: () => null },
      body: {
        getReader: () => ({
          read: async () => {
            const value = chunks[reads];
            reads += 1;
            return value ? { done: false, value } : { done: true };
          },
          cancel: async () => { cancelCalls += 1; },
          releaseLock: () => {},
        }),
      },
      arrayBuffer: async () => { throw new Error("the stream must be preferred"); },
    }),
  };

  const loaded = await runtime.loadOfficeSource(
    { kind: "url", url: "book.xlsx" },
    { urlPolicy: policy, limits: { maxInputBytes: 4 } },
  );
  assert.deepEqual([...new Uint8Array(loaded.buffer)], [0x50, 0x4b, 0x03, 0x04]);
  assert.equal(reads, 3);
  assert.equal(cancelCalls, 0);
});

test("URL loading rejects an oversized Content-Length before reading the stream", async () => {
  let readerCalls = 0;
  let cancelCalls = 0;
  let arrayBufferCalls = 0;
  const policy = {
    enabled: true,
    baseUrl: "https://app.example/viewer/",
    allowRelativeUrl: true,
    allowedProtocols: ["https:"],
    allowedOrigins: ["https://app.example"],
    fetch: async (url) => ({
      ok: true,
      status: 200,
      url,
      headers: { get: (name) => name.toLowerCase() === "content-length" ? "8" : null },
      body: {
        getReader: () => {
          readerCalls += 1;
          throw new Error("the body must not be read");
        },
        cancel: async () => { cancelCalls += 1; },
      },
      arrayBuffer: async () => {
        arrayBufferCalls += 1;
        return new ArrayBuffer(8);
      },
    }),
  };

  await assert.rejects(
    runtime.loadOfficeSource(
      { kind: "url", url: "book.xlsx" },
      { urlPolicy: policy, limits: { maxInputBytes: 4 } },
    ),
    (error) => error?.code === "LIMIT_EXCEEDED" && error?.actual === 8 && error?.allowed === 4,
  );
  assert.equal(readerCalls, 0);
  assert.equal(cancelCalls, 1);
  assert.equal(arrayBufferCalls, 0);
});

test("URL loading still enforces maxInputBytes for fetchers that only expose arrayBuffer", async () => {
  let reads = 0;
  const policy = {
    enabled: true,
    baseUrl: "https://app.example/viewer/",
    allowRelativeUrl: true,
    allowedProtocols: ["https:"],
    allowedOrigins: ["https://app.example"],
    fetch: async (url) => ({
      ok: true,
      status: 200,
      url,
      headers: { get: () => null },
      arrayBuffer: async () => {
        reads += 1;
        return new Uint8Array(8).buffer;
      },
    }),
  };

  await assert.rejects(
    runtime.loadOfficeSource(
      { kind: "url", url: "book.xlsx" },
      { urlPolicy: policy, limits: { maxInputBytes: 5 } },
    ),
    (error) => error?.code === "LIMIT_EXCEEDED" && error?.actual === 8 && error?.allowed === 5,
  );
  assert.equal(reads, 1);
});

test("URL loading cancels the stream and reports ABORTED when cancelled mid-read", async () => {
  const controller = new AbortController();
  let reads = 0;
  let cancelCalls = 0;
  let releaseCalls = 0;
  const policy = {
    enabled: true,
    baseUrl: "https://app.example/viewer/",
    allowRelativeUrl: true,
    allowedProtocols: ["https:"],
    allowedOrigins: ["https://app.example"],
    fetch: async (url) => ({
      ok: true,
      status: 200,
      url,
      headers: { get: () => null },
      body: {
        getReader: () => ({
          read: async () => {
            reads += 1;
            if (reads === 1) return { done: false, value: new Uint8Array([0x50]) };
            controller.abort();
            return { done: false, value: new Uint8Array([0x4b]) };
          },
          cancel: async () => { cancelCalls += 1; },
          releaseLock: () => { releaseCalls += 1; },
        }),
      },
      arrayBuffer: async () => new ArrayBuffer(0),
    }),
  };

  await assert.rejects(
    runtime.loadOfficeSource(
      { kind: "url", url: "book.xlsx" },
      { urlPolicy: policy, signal: controller.signal, limits: { maxInputBytes: 10 } },
    ),
    (error) => error?.code === "ABORTED" && error?.phase === "read",
  );
  assert.equal(reads, 2);
  assert.equal(cancelCalls, 1);
  assert.equal(releaseCalls, 1);
});

test("URL loading follows same-origin redirects and revalidates the final address", async () => {
  const calls = [];
  const bytes = new Uint8Array([0x50, 0x4b, 0x03, 0x04]);
  const policy = {
    enabled: true,
    baseUrl: "https://app.example/viewer/",
    allowRelativeUrl: true,
    allowedProtocols: ["https:"],
    allowedOrigins: ["https://app.example"],
    fetch: async (url, init) => {
      calls.push({ url, init });
      return {
        ok: true,
        status: 200,
        url: "https://app.example/cdn/final.xlsx?token=secret",
        headers: { get: () => null },
        arrayBuffer: async () => bytes.buffer.slice(0),
      };
    },
  };

  const loaded = await runtime.loadOfficeSource(
    { kind: "url", url: "book.xlsx" },
    { urlPolicy: policy, limits: { maxInputBytes: 4 } },
  );
  assert.equal(calls[0].url, "https://app.example/viewer/book.xlsx");
  assert.equal(calls[0].init.redirect, "follow");
  assert.equal(loaded.resolvedUrl, "https://app.example/cdn/final.xlsx");
  assert.equal(loaded.fileName, "final.xlsx");
});

test("URL loading rejects redirects whose final address violates the URL policy", async () => {
  for (const finalUrl of ["https://other.example/final.xlsx", "http://app.example/final.xlsx"]) {
    let reads = 0;
    const policy = {
      enabled: true,
      baseUrl: "https://app.example/viewer/",
      allowRelativeUrl: true,
      allowedProtocols: ["https:"],
      allowedOrigins: ["https://app.example"],
      fetch: async () => ({
        ok: true,
        status: 200,
        url: finalUrl,
        headers: { get: () => null },
        arrayBuffer: async () => {
          reads += 1;
          return new ArrayBuffer(4);
        },
      }),
    };

    await assert.rejects(
      runtime.loadOfficeSource(
        { kind: "url", url: "book.xlsx" },
        { urlPolicy: policy, limits: { maxInputBytes: 4 } },
      ),
      (error) => error?.code === "SOURCE_NOT_ALLOWED" && error?.phase === "source",
    );
    assert.equal(reads, 0);
  }
});

test("all formats receive the same stable abort error and safe diagnostics", () => {
  const abort = new Error("cancelled");
  abort.name = "AbortError";
  for (const format of ["docx", "xlsx", "pdf"]) {
    const error = runtime.toOfficeLoadError(abort, {
      fallbackCode: "INVALID_SOURCE",
      format,
      sourceKind: "url",
      taskId: `${format}:1`,
      url: "https://app.example/file?token=secret",
    });
    assert.equal(error.name, "OfficeLoadError");
    assert.equal(error.code, "ABORTED");
    assert.equal(error.format, format);
    assert.equal(error.taskId, `${format}:1`);
    assert.equal(error.url, "https://app.example/file");
    assert.equal(JSON.stringify(error).includes("secret"), false);
    assert.equal(JSON.stringify(error).includes("stack"), false);
  }

  const event = runtime.createOfficeDiagnostic({
    type: "load-error",
    runtimeId: "safe-runtime",
    taskId: "safe-runtime:1",
    sourceKind: "url",
    url: "https://app.example/file?token=secret#fragment",
  });
  assert.equal(event.url, "https://app.example/file");
  assert.equal(JSON.stringify(event).includes("secret"), false);
  assert.equal(Object.isFrozen(event), true);
});

test("archive parts validation matches the streaming validator without a second unzip", () => {
  const archive = makeArchive({
    "word/document.xml": "<w:document xmlns:w=\"urn:w\"><w:body><w:p>ok</w:p></w:body></w:document>",
    "word/_rels/document.xml.rels": "<Relationships><Relationship Id=\"r1\" Type=\"t\" Target=\"x\"/></Relationships>",
  });
  const validated = runtime.validateOfficeArchive(archive, {}, { format: "docx" });
  const parts = [
    { name: "[Content_Types].xml", text: DOCX_CONTENT_TYPES },
    { name: "word/document.xml", text: "<w:document xmlns:w=\"urn:w\"><w:body><w:p>ok</w:p></w:body></w:document>" },
    { name: "word/_rels/document.xml.rels", text: "<Relationships><Relationship Id=\"r1\" Type=\"t\" Target=\"x\"/></Relationships>" },
  ];
  const fromParts = runtime.validateOfficeArchiveParts(parts, {}, { format: "docx" });
  assert.equal(fromParts.sourceFormat, validated.sourceFormat);
  assert.equal(fromParts.entryCount, validated.entryCount);
  assert.equal(fromParts.uncompressedBytes, validated.uncompressedBytes);
  assert.equal(fromParts.xmlBytes, validated.xmlBytes);
  assert.equal(fromParts.relationships, validated.relationships);
  assert.equal(Object.isFrozen(fromParts), true);

  // Explicit byte lengths win over measuring the text, so binary parts can pass sizes.
  assert.equal(
    runtime.validateOfficeArchiveParts(
      [{ name: "[Content_Types].xml", text: DOCX_CONTENT_TYPES }, { name: "word/media/a.bin", bytes: 4 }],
      {},
      { format: "docx" },
    ).uncompressedBytes,
    Buffer.byteLength(DOCX_CONTENT_TYPES) + 4,
  );
  assert.throws(
    () => runtime.validateOfficeArchiveParts(
      [{ name: "[Content_Types].xml", text: DOCX_CONTENT_TYPES }, { name: "word/media/a.bin" }],
      {},
      { format: "docx" },
    ),
    (error) => error?.code === "INVALID_ARGUMENT",
  );
});

test("archive parts validation enforces XML budgets, duplicates and the OOXML family", () => {
  const parts = [
    { name: "[Content_Types].xml", text: DOCX_CONTENT_TYPES },
    { name: "word/document.xml", text: `<w:document>${"<a>".repeat(12)}x${"</a>".repeat(12)}</w:document>` },
    { name: "word/_rels/document.xml.rels", text: "<Relationships><Relationship Id=\"r1\" Type=\"t\" Target=\"x\"/></Relationships>" },
  ];
  for (const [limit, value] of [
    ["maxArchiveEntries", 1],
    ["maxSingleEntryBytes", 10],
    ["maxUncompressedBytes", 10],
    ["maxXmlBytes", 10],
    ["maxSingleXmlBytes", 10],
    ["maxXmlDepth", 4],
    ["maxRelationships", 0],
  ]) {
    assert.throws(
      () => runtime.validateOfficeArchiveParts(parts, { [limit]: value }, { format: "docx" }),
      (error) => error?.code === "LIMIT_EXCEEDED" && error?.limit === limit,
      `expected ${limit} to be enforced`,
    );
  }
  assert.throws(
    () => runtime.validateOfficeArchiveParts(
      [{ name: "word/document.xml", text: "<x/>" }, { name: "word/document.xml", text: "<x/>" }],
      {},
      { format: "docx" },
    ),
    (error) => error?.code === "INVALID_SOURCE" && error?.phase === "archive",
  );
  assert.throws(
    () => runtime.validateOfficeArchiveParts(
      [{ name: "../evil.xml", text: "<x/>" }],
      {},
      { format: "docx" },
    ),
    (error) => error?.code === "INVALID_SOURCE" && error?.phase === "archive",
  );
  assert.throws(
    () => runtime.validateOfficeArchiveParts(
      [{ name: "[Content_Types].xml", text: XLSX_CONTENT_TYPES }],
      {},
      { format: "docx" },
    ),
    (error) => error?.code === "INVALID_SOURCE" && error?.phase === "archive",
  );
  assert.throws(
    () => runtime.validateOfficeArchiveParts(
      [{ name: "word/document.xml", text: "<!DOCTYPE x><x/>" }],
      {},
      { format: "docx" },
    ),
    (error) => error?.code === "INVALID_SOURCE" && error?.phase === "xml",
  );
  const controller = new AbortController();
  controller.abort();
  assert.throws(
    () => runtime.validateOfficeArchiveParts(parts, {}, { format: "docx", signal: controller.signal }),
    (error) => error?.code === "ABORTED",
  );
});

test("byte sources are read in place instead of copied", async () => {
  const bytes = zipSync({ "a.txt": strToU8("ok") });
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  const loaded = await runtime.loadOfficeSource({ kind: "bytes", bytes: buffer, name: "a.docx" });
  assert.equal(loaded.buffer, buffer);
  assert.equal(loaded.fileName, "a.docx");
  assert.equal(runtime.utf8ByteLength("a\u00e9\u4e2d\ud83d\ude00"), 10);
});

test("resolved limits include hard-cap-only entries and never drop caps", () => {
  const resolved = runtime.resolveOfficeLimits(
    { maxInputBytes: 1000 },
    { maxInputBytes: 5000, maxArchiveEntries: 5 },
    { maxInputBytes: 100, maxArchiveEntries: 10, maxUncompressedBytes: 7 },
  );
  assert.equal(resolved.maxInputBytes, 100);
  assert.equal(resolved.maxArchiveEntries, 5);
  assert.equal(resolved.maxUncompressedBytes, 7);
  assert.equal("maxParseMs" in resolved, false);
});

test("load context resolves partial limits against defaults and hard caps", () => {
  const context = runtime.createOfficeLoadContext({
    runtimeId: "limits-runtime",
    taskId: "limits-runtime:1",
    source: { kind: "bytes", bytes: new ArrayBuffer(4) },
    limits: { maxInputBytes: 4 },
    limitDefaults: { maxInputBytes: 100, maxArchiveEntries: 10 },
    limitHardCaps: { maxInputBytes: 50, maxUncompressedBytes: 7 },
    signal: new AbortController().signal,
  });
  assert.deepEqual(context.limits, {
    maxInputBytes: 4,
    maxArchiveEntries: 10,
    maxUncompressedBytes: 7,
  });
  assert.equal(Object.isFrozen(context.limits), true);
});

test("streamed URL sources are rejected before an oversized body is buffered", async () => {
  let arrayBufferReads = 0;
  const policy = {
    enabled: true,
    baseUrl: "https://files.example/",
    allowedProtocols: ["https:"],
    allowedOrigins: ["https://files.example"],
    fetch: async (url) => ({
      ok: true,
      status: 200,
      url,
      headers: { get: () => null },
      body: new ReadableStream({
        start(controller) {
          controller.enqueue(new Uint8Array(6));
          controller.enqueue(new Uint8Array(6));
          controller.close();
        },
      }),
      arrayBuffer: async () => {
        arrayBufferReads += 1;
        return new ArrayBuffer(12);
      },
    }),
  };

  await assert.rejects(
    runtime.loadOfficeSource(
      { kind: "url", url: "https://files.example/large.bin" },
      { urlPolicy: policy, limits: { maxInputBytes: 8 } },
    ),
    (error) => error?.code === "LIMIT_EXCEEDED"
      && error?.limit === "maxInputBytes"
      && error?.actual === 12
      && error?.allowed === 8,
  );
  assert.equal(arrayBufferReads, 0);
});

test("streamed URL sources accept bodies at the exact byte limit", async () => {
  const bytes = Uint8Array.from([1, 2, 3, 4, 5, 6, 7, 8]);
  const policy = {
    enabled: true,
    baseUrl: "https://files.example/",
    allowedProtocols: ["https:"],
    allowedOrigins: ["https://files.example"],
    fetch: async (url) => ({
      ok: true,
      status: 200,
      url,
      headers: { get: () => null },
      body: new ReadableStream({
        start(controller) {
          controller.enqueue(bytes.slice(0, 3));
          controller.enqueue(bytes.slice(3));
          controller.close();
        },
      }),
      arrayBuffer: async () => {
        throw new Error("streaming responses must not be buffered twice");
      },
    }),
  };

  const loaded = await runtime.loadOfficeSource(
    { kind: "url", url: "https://files.example/exact.bin" },
    { urlPolicy: policy, limits: { maxInputBytes: bytes.byteLength } },
  );
  assert.deepEqual([...new Uint8Array(loaded.buffer)], [...bytes]);
});

test("touch gesture math keeps pinch and double-tap zoom inside the shared range", () => {
  assert.equal(runtime.pointerDistance({ clientX: 0, clientY: 0 }, { clientX: 3, clientY: 4 }), 5);
  assert.equal(runtime.pinchSurfaceZoom(1, 100, 200), 2);
  assert.equal(runtime.pinchSurfaceZoom(1, 100, 50), 0.5);
  assert.equal(runtime.pinchSurfaceZoom(1, 100, 400), 2);
  assert.equal(runtime.pinchSurfaceZoom(1, 100, 1), 0.5);
  assert.equal(runtime.pinchSurfaceZoom(1, 0, 200), 1);
  assert.equal(runtime.pinchSurfaceZoom(Number.NaN, 100, 200), 1);
  assert.equal(runtime.nextDoubleTapZoom(1), 2);
  assert.equal(runtime.nextDoubleTapZoom(0.5), 2);
  assert.equal(runtime.nextDoubleTapZoom(1.25), 1);
  assert.equal(runtime.nextDoubleTapZoom(2), 1);
  assert.equal(runtime.TOUCH_MIN_TARGET_PX, 44);
});

function touchPointer(overrides = {}) {
  return {
    pointerId: 1,
    pointerType: "touch",
    clientX: 0,
    clientY: 0,
    ...overrides,
  };
}

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

test("touch tracker maps a two-finger pinch to zoom updates around the midpoint", () => {
  const zooms = [];
  const tracker = runtime.createTouchGestureTracker({
    getZoom: () => 1,
    onZoom: (zoom, center) => zooms.push({ zoom, center }),
  });

  tracker.pointerDown(touchPointer({ pointerId: 1, clientX: 0, clientY: 0 }));
  tracker.pointerDown(touchPointer({ pointerId: 2, clientX: 100, clientY: 0 }));
  assert.equal(tracker.isPinching, true);

  tracker.pointerMove(touchPointer({ pointerId: 2, clientX: 200, clientY: 0 }));
  assert.deepEqual(zooms, [{ zoom: 2, center: { clientX: 100, clientY: 0 } }]);

  // The final finger-up after a pinch must not become a tap.
  tracker.pointerUp(touchPointer({ pointerId: 2, clientX: 200, clientY: 0 }));
  tracker.pointerUp(touchPointer({ pointerId: 1, clientX: 0, clientY: 0 }));
  tracker.pointerDown(touchPointer({ pointerId: 1, clientX: 0, clientY: 0 }));
  tracker.pointerUp(touchPointer({ pointerId: 1, clientX: 0, clientY: 0 }));
  assert.equal(zooms.length, 1);
});

test("touch tracker fires a long press once, suppresses the native menu, and cancels on movement", async () => {
  const presses = [];
  const tracker = runtime.createTouchGestureTracker({
    longPressMs: 10,
    onLongPress: (pointer) => presses.push(pointer),
  });

  tracker.pointerDown(touchPointer({ clientX: 40, clientY: 50 }));
  await wait(30);
  assert.equal(presses.length, 1);
  assert.equal(presses[0].clientX, 40);
  assert.equal(presses[0].clientY, 50);
  assert.equal(tracker.shouldSuppressContextMenu(42, 52), true);
  assert.equal(tracker.shouldSuppressContextMenu(240, 52), false);
  let prevented = 0;
  const pressRelease = touchPointer({ clientX: 40, clientY: 50 });
  tracker.pointerUp({ ...pressRelease, preventDefault: () => { prevented += 1; } });
  assert.equal(prevented, 1, "long press suppresses the trailing click");

  tracker.pointerDown(touchPointer({ clientX: 10, clientY: 10 }));
  tracker.pointerMove(touchPointer({ clientX: 60, clientY: 10 }));
  await wait(30);
  assert.equal(presses.length, 1, "moving before the delay cancels the long press");
});

test("touch tracker recognises a double tap and prevents the trailing click", () => {
  const taps = [];
  let prevented = 0;
  const tracker = runtime.createTouchGestureTracker({
    doubleTapMs: 1_000,
    onDoubleTap: (pointer) => taps.push(pointer),
  });
  const tap = (overrides = {}) => {
    const event = touchPointer({ clientX: 20, clientY: 30, ...overrides });
    tracker.pointerDown(event);
    tracker.pointerUp({ ...event, preventDefault: () => { prevented += 1; } });
  };

  tap();
  assert.equal(taps.length, 0);
  tap();
  assert.equal(taps.length, 1);
  assert.equal(taps[0].clientX, 20);
  assert.equal(prevented, 1);

  // A second double tap pair starts a new sequence rather than firing twice.
  tap();
  assert.equal(taps.length, 1);
  tap();
  assert.equal(taps.length, 2);
});

test("touch tracker can keep double tap for editing targets", () => {
  const taps = [];
  const tracker = runtime.createTouchGestureTracker({
    doubleTapMs: 1_000,
    isDoubleTapEnabled: () => false,
    onDoubleTap: (pointer) => taps.push(pointer),
  });
  const tap = () => {
    tracker.pointerDown(touchPointer());
    tracker.pointerUp(touchPointer());
  };
  tap();
  tap();
  assert.equal(taps.length, 0);
});

test("touch tracker ignores mouse pointers and disabled surfaces", async () => {
  let presses = 0;
  let enabled = false;
  const tracker = runtime.createTouchGestureTracker({
    longPressMs: 5,
    isEnabled: () => enabled,
    onLongPress: () => { presses += 1; },
  });

  tracker.pointerDown(touchPointer({ pointerType: "mouse" }));
  await wait(15);
  tracker.pointerDown(touchPointer());
  await wait(15);
  assert.equal(presses, 0);

  enabled = true;
  tracker.pointerDown(touchPointer());
  await wait(15);
  assert.equal(presses, 1);
  tracker.reset();
});
