import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";

const requireFromDemo = createRequire(
  new URL("../../apps/demo/package.json", import.meta.url),
);
const typescript = requireFromDemo("typescript");
const helperSource = readFileSync(
  new URL("../../apps/demo/src/runtime-limits.ts", import.meta.url),
  "utf8",
);
const helperModule = await import(`data:text/javascript;base64,${Buffer.from(
  typescript.transpileModule(helperSource, {
    compilerOptions: {
      module: typescript.ModuleKind.ES2022,
      target: typescript.ScriptTarget.ES2022,
    },
  }).outputText,
).toString("base64")}`);

const sampleBytes = readFileSync(
  new URL("../../apps/demo/public/samples/invoice-table.docx", import.meta.url),
).byteLength;

test("runtime demo input limit rejects from file.size alone", () => {
  assert.ok(helperModule.ISOLATION_REJECT_LIMIT_BYTES < sampleBytes);
  assert.ok(helperModule.ISOLATION_SUCCESS_LIMIT_BYTES > sampleBytes);

  const rejected = helperModule.inputSizeLimitError(
    sampleBytes,
    helperModule.ISOLATION_REJECT_LIMIT_BYTES,
  );
  assert.deepEqual(rejected, {
    code: "LIMIT_EXCEEDED",
    message: `输入大小 ${sampleBytes} 字节超过允许值 ${helperModule.ISOLATION_REJECT_LIMIT_BYTES} 字节。`,
    phase: "input",
    limit: "maxInputBytes",
    actual: sampleBytes,
    allowed: helperModule.ISOLATION_REJECT_LIMIT_BYTES,
  });
  assert.equal(helperModule.inputSizeLimitError(sampleBytes, sampleBytes), undefined);
  assert.equal(helperModule.inputSizeLimitError(sampleBytes, undefined), undefined);
  assert.equal(helperModule.inputSizeLimitError(undefined, 1), undefined);
  assert.equal(helperModule.inputSizeLimitError(sampleBytes, 0), undefined);
});

test("runtime demo normalizes structured limit errors", () => {
  const error = Object.assign(new Error("too large"), {
    code: "LIMIT_EXCEEDED",
    phase: "input",
    limit: "maxInputBytes",
    actual: 10,
    allowed: 5,
  });
  const normalized = helperModule.normalizeRuntimeError(error);
  assert.deepEqual(normalized, {
    code: "LIMIT_EXCEEDED",
    message: "too large",
    phase: "input",
    limit: "maxInputBytes",
    actual: 10,
    allowed: 5,
  });
  assert.equal(helperModule.isInputLimitError(normalized), true);
  assert.equal(helperModule.isInputLimitError({ code: "LIMIT_EXCEEDED", message: "" }), false);
  assert.deepEqual(helperModule.normalizeRuntimeError("boom"), {
    code: "PARSE_FAILED",
    message: "boom",
  });
});

test("runtime limits page checks file.size before reading the upload", () => {
  const source = readFileSync(
    new URL("../../apps/demo/src/pages/RuntimeLimitsPage.vue", import.meta.url),
    "utf8",
  );
  const sizeCheck = source.indexOf(
    "const preReadError = inputSizeLimitError(file.size, activeLimits.value.maxInputBytes)",
  );
  const read = source.indexOf("const buffer = await file.arrayBuffer()");
  assert.ok(sizeCheck >= 0, "file.size precheck is missing");
  assert.ok(read >= 0, "file read is missing");
  assert.ok(sizeCheck < read, "file.size must be checked before arrayBuffer()");
  assert.match(source, /data-error-source/);
  assert.match(source, /data-read-bytes/);
});

test("runtime isolation and worker pages exercise their configured limits", () => {
  const isolation = readFileSync(
    new URL("../../apps/demo/src/pages/RuntimeIsolationPage.vue", import.meta.url),
    "utf8",
  );
  assert.match(isolation, /ISOLATION_REJECT_LIMIT_BYTES/);
  assert.match(isolation, /data-demo-rejection/);
  assert.match(isolation, /runtime-b-error/);

  const worker = readFileSync(
    new URL("../../apps/demo/src/pages/RuntimeWorkerPage.vue", import.meta.url),
    "utf8",
  );
  assert.match(worker, /worker-error-limit/);
  assert.match(worker, /worker-error-actual/);
  assert.match(worker, /worker-error-allowed/);
  assert.match(worker, /worker-configured-parse-limit/);
});
