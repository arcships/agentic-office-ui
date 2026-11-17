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

// Issue #57: a BLOCKED step (exit code 2) must never be reported as a pass.
function runUnitProbe(id, extraEnv) {
  const probeRoot = path.resolve(evidenceRoot, id);
  const probe = spawnSync(pnpm, ["test:unit"], {
    cwd: root,
    encoding: "utf8",
    env: {
      ...process.env,
      CI_PREBUILT: "1",
      CI_RUN_ID: process.env.CI_RUN_ID || "blocked-probe",
      CI_EVIDENCE_ROOT: probeRoot,
      CI_FAILURE_PROBE: "unit:docx-integration",
      ...extraEnv,
    },
    maxBuffer: 16 * 1024 * 1024,
  });
  const summaryPath = path.join(probeRoot, "unit", "summary.json");
  writeFileSync(
    path.join(evidenceRoot, `${id}-command.log`),
    `${probe.stdout || ""}${probe.stderr || ""}`,
  );
  return { probe, summary: JSON.parse(readFileSync(summaryPath, "utf8")) };
}

const blockedProbe = runUnitProbe("blocked", {
  CI_FAILURE_PROBE_EXIT: "2",
  CI_PROBE_BLOCKED: "1",
});
assert.equal(blockedProbe.probe.status, 2, "a BLOCKED step must exit non-zero");
assert.equal(blockedProbe.summary.result, "INCOMPLETE", "blocked steps must not pass the suite");
assert.equal(blockedProbe.summary.allowBlocked, false);
assert.deepEqual(blockedProbe.summary.blockedSteps, ["docx-integration"]);
assert.equal(
  blockedProbe.summary.results.find((item) => item.id === "docx-integration")?.status,
  "BLOCKED",
);

const allowedProbe = runUnitProbe("blocked-allowed", {
  CI_FAILURE_PROBE_EXIT: "2",
  CI_PROBE_BLOCKED: "1",
  CI_ALLOW_BLOCKED: "1",
});
assert.equal(allowedProbe.probe.status, 0, "CI_ALLOW_BLOCKED=1 is the only opt-out");
assert.equal(allowedProbe.summary.result, "PASS");
assert.equal(allowedProbe.summary.allowBlocked, true);
assert.deepEqual(allowedProbe.summary.blockedSteps, ["docx-integration"]);

// The acceptance shard registry must stay callable and must reject typos.
const shardList = spawnSync(
  process.execPath,
  [path.join(root, "scripts", "ci", "run-acceptance-shard.mjs"), "--list"],
  { cwd: root, encoding: "utf8", maxBuffer: 4 * 1024 * 1024 },
);
assert.equal(shardList.status, 0, "the acceptance shard registry must list its shards");
for (const shard of ["docs", "pack", "self-test", "release-pack"]) {
  assert.ok(shardList.stdout.includes(shard), `acceptance shard ${shard} must be registered`);
}
const unknownShard = spawnSync(
  process.execPath,
  [path.join(root, "scripts", "ci", "run-acceptance-shard.mjs"), "not-a-shard"],
  { cwd: root, encoding: "utf8", maxBuffer: 4 * 1024 * 1024 },
);
assert.equal(unknownShard.status, 2, "an unknown shard must fail fast");

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
  blockedProbe: {
    exitCode: blockedProbe.probe.status,
    result: blockedProbe.summary.result,
    blockedSteps: blockedProbe.summary.blockedSteps,
  },
  allowBlockedProbe: {
    exitCode: allowedProbe.probe.status,
    result: allowedProbe.summary.result,
  },
  acceptanceShards: shardList.stdout
    .split("\n")
    .map((line) => line.trim().split(/\s+/)[1])
    .filter(Boolean),
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
