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
  assert.equal(calls[0].init.redirect, "error");
  assert.equal(calls[0].init.credentials, "omit");

  await assert.rejects(
    runtime.loadOfficeSource(
      { kind: "url", url: "book.xlsx" },
      { urlPolicy: policy, limits: { maxInputBytes: bytes.byteLength - 1 } },
    ),
    (error) => error?.code === "LIMIT_EXCEEDED" && error?.limit === "maxInputBytes",
  );
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
