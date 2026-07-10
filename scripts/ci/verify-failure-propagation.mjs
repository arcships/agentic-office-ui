#!/usr/bin/env node

import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";

const root = path.resolve(import.meta.dirname, "../..");
const pnpm = process.platform === "win32" ? "pnpm.cmd" : "pnpm";
const evidenceRoot = path.resolve(
  process.env.CI_EVIDENCE_ROOT ||
    path.join(root, "output", "acceptance", "ci-self-test"),
  "failure-probe",
);
mkdirSync(evidenceRoot, { recursive: true });

const result = spawnSync(pnpm, ["test:unit"], {
  cwd: root,
  encoding: "utf8",
  env: {
    ...process.env,
    CI_PREBUILT: "1",
    CI_RUN_ID: process.env.CI_RUN_ID || "failure-probe",
    CI_EVIDENCE_ROOT: evidenceRoot,
    CI_FAILURE_PROBE: "unit:docx-integration",
  },
  maxBuffer: 16 * 1024 * 1024,
});
const output = `${result.stdout || ""}${result.stderr || ""}`;
writeFileSync(path.join(evidenceRoot, "command.log"), output);

assert.notEqual(result.status, 0, "a child failure must make the suite fail");
const summaryPath = path.join(evidenceRoot, "unit", "summary.json");
const summary = JSON.parse(readFileSync(summaryPath, "utf8"));
assert.equal(summary.result, "FAIL");
const failedIndex = summary.results.findIndex((item) => item.id === "docx-integration");
assert.ok(failedIndex >= 0, "the requested failing step must be reported");
assert.equal(summary.results[failedIndex].status, "FAIL");
assert.equal(summary.results[failedIndex].failureProbe, true);
assert.ok(
  summary.results.slice(0, failedIndex).every((item) => item.status === "PASS"),
  "steps before the failure must pass",
);
assert.ok(
  summary.results.slice(failedIndex + 1).every((item) => item.status === "NOT_RUN"),
  "steps after the failure must not run",
);

const typeProbe = path.join(root, "apps", "demo", "src", "__ci_type_failure_probe__.vue");
assert.equal(existsSync(typeProbe), false, "the temporary type probe path must be unused");
let typeResult;
try {
  writeFileSync(
    typeProbe,
    '<script setup lang="ts">\nconst mustBeText: string = 42;\n</script>\n<template><div>{{ mustBeText }}</div></template>\n',
  );
  typeResult = spawnSync(pnpm, ["typecheck"], {
    cwd: root,
    encoding: "utf8",
    maxBuffer: 16 * 1024 * 1024,
  });
  writeFileSync(
    path.join(evidenceRoot, "type-error-command.log"),
    `${typeResult.stdout || ""}${typeResult.stderr || ""}`,
  );
} finally {
  rmSync(typeProbe, { force: true });
}
assert.notEqual(typeResult.status, 0, "a real Vue type error must fail root typecheck");
const typeOutput = `${typeResult.stdout || ""}${typeResult.stderr || ""}`;
assert.match(typeOutput, /__ci_type_failure_probe__\.vue/);
assert.match(typeOutput, /TS2322/);
assert.equal(existsSync(typeProbe), false, "the temporary type probe must be cleaned");

const report = {
  result: "PASS",
  assertion: "registered child failures propagate as non-zero and identify the failing step",
  childExitCode: result.status,
  childSummary: summaryPath,
  typeErrorProbe: {
    exitCode: typeResult.status,
    errorCode: "TS2322",
    cleaned: !existsSync(typeProbe),
  },
};
writeFileSync(
  path.join(evidenceRoot, "self-test-summary.json"),
  `${JSON.stringify(report, null, 2)}\n`,
);
console.log(`PASS: CI failure propagation verified at ${summaryPath}`);
