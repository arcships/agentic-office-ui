#!/usr/bin/env node

/**
 * Acceptance shards for continuous integration.
 *
 * Issue #57: the workflow built packages but never verified them, and several
 * verifiers could finish without asserting anything. This runner registers the
 * verification work that must survive CI into four independent shards that
 * parallelise on GitHub Actions.
 *
 * Every shard is offline by design: no upstream checkout, no preview server and
 * no browser. Browser blackbox suites stay opt-in (`pnpm test:blackbox`,
 * `pnpm test:consumer`, `pnpm test:release`) because they need a Playwright
 * Chromium download and, for upstream UX parity, a fixed upstream worktree.
 * `assertOffline` rejects any registered step that would break that promise.
 *
 * Usage:
 *   node scripts/ci/run-acceptance-shard.mjs --list
 *   node scripts/ci/run-acceptance-shard.mjs docs
 *   node scripts/ci/run-acceptance-shard.mjs 2
 *   node scripts/ci/run-acceptance-shard.mjs all
 *
 * Environment:
 *   CI_PREBUILT=1        reuse an existing workspace build instead of building
 *   CI_EVIDENCE_ROOT     evidence base directory (default output/acceptance/...)
 *   CI_RUN_ID            evidence run identifier
 */

import { spawnSync } from "node:child_process";
import { mkdirSync, writeFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";

const root = path.resolve(import.meta.dirname, "../..");
const pnpm = process.platform === "win32" ? "pnpm.cmd" : "pnpm";
const prebuilt = process.env.CI_PREBUILT === "1";

function timestamp() {
  return new Date().toISOString().replace(/[-:]/g, "").replace(/\.\d{3}Z$/, "Z");
}

function readGitHead() {
  const result = spawnSync("git", ["rev-parse", "HEAD"], {
    cwd: root,
    encoding: "utf8",
  });
  return result.status === 0 ? result.stdout.trim() : "no-git-head";
}

function command(id, executable, args, options = {}) {
  return { id, executable, args, ...options };
}

// The unit suite inspects apps/demo/dist (published chunk separation), so the
// build step mirrors `pnpm build` instead of building packages alone.
const buildStep = command("build", pnpm, ["build"]);

const shards = [
  {
    id: "docs",
    alias: "1",
    title: "Documentation and registration contract",
    // Registers every required file, suite step and documentation snippet; it
    // reads the repository only and therefore needs no build.
    needsBuild: false,
    steps: [
      command("documentation-contract", process.execPath, ["scripts/ci/verify-docs.mjs"]),
      command("open-source-attribution", process.execPath, [
        "scripts/ci/verify-attribution.mjs",
      ]),
    ],
  },
  {
    id: "pack",
    alias: "2",
    title: "Published package manifests",
    needsBuild: true,
    steps: [
      command("real-tgz-manifests", process.execPath, ["scripts/ci/pack-manifests.mjs"]),
    ],
  },
  {
    id: "self-test",
    alias: "3",
    title: "CI failure propagation self-test",
    // Proves that a failing child suite exits non-zero and that a real type
    // error fails the root typecheck: the guard against green-but-empty CI.
    needsBuild: true,
    steps: [
      command("failure-propagation", process.execPath, [
        "scripts/ci/verify-failure-propagation.mjs",
      ]),
    ],
  },
  {
    id: "release-pack",
    alias: "4",
    title: "Reproducible two-clean-build package verification",
    // The script copies a clean snapshot and runs its own install and build
    // twice, so a workspace build here would only be wasted time.
    needsBuild: false,
    steps: [
      command("p4-reproducible-pack", process.execPath, [
        "scripts/ci/p4-reproducible-pack.mjs",
      ]),
    ],
  },
];

// Shards must stay runnable on a plain ubuntu-latest runner: no upstream
// checkout, no preview server, no browser download.
const offlineViolations = [
  { pattern: /tests[\\/]blackbox/i, reason: "browser blackbox tests are opt-in (pnpm test:blackbox)" },
  { pattern: /run-python\.mjs/i, reason: "Python steps need Playwright, which is opt-in" },
  { pattern: /playwright/i, reason: "Playwright browsers are opt-in" },
  { pattern: /ux_parity|extend-hq/i, reason: "upstream UX parity needs a fixed upstream worktree" },
  { pattern: /pack-consumer|verify-pptx-consumer|compatibility-matrix/i, reason: "consumer checks install tarballs from the registry" },
  { pattern: /https?:\/\//i, reason: "shards must not depend on network services" },
];

function assertOffline(definition) {
  for (const step of definition.steps) {
    const rendered = [step.executable, ...step.args].join(" ");
    for (const violation of offlineViolations) {
      if (violation.pattern.test(rendered)) {
        console.error(
          `Shard ${definition.id} registers an offline-unfriendly step ${step.id}: ${rendered}\n${violation.reason}`,
        );
        process.exit(2);
      }
    }
  }
}

function findShard(argument) {
  return shards.find((item) => item.id === argument || item.alias === argument);
}

function usage(message) {
  if (message) console.error(message);
  console.error(
    `Unknown acceptance shard. Expected one of: ${shards
      .map((item) => `${item.alias}=${item.id}`)
      .join(", ")}, or all.`,
  );
  process.exit(2);
}

const argument = process.argv[2];

if (!argument || argument === "--help" || argument === "-h") {
  usage(argument ? undefined : "Missing acceptance shard name.");
}

if (argument === "--list" || argument === "list") {
  for (const shard of shards) {
    console.log(`${shard.alias}  ${shard.id.padEnd(12)} ${shard.title}`);
    for (const step of shard.steps) {
      console.log(`     - ${step.id}: ${[step.executable, ...step.args].join(" ")}`);
    }
  }
  process.exit(0);
}

const selected = argument === "all" ? shards : [findShard(argument)];
if (selected.some((item) => !item)) usage();

for (const shard of selected) assertOffline(shard);

const commit = readGitHead();
const runId = process.env.CI_RUN_ID || timestamp();
const evidenceRoot = path.resolve(
  process.env.CI_EVIDENCE_ROOT ||
    path.join(root, "output", "acceptance", commit, "P1-CI-01", runId),
);
const baseEnv = {
  ...process.env,
  CI_RUN_ID: runId,
  CI_EVIDENCE_ROOT: evidenceRoot,
};

let shardFailed = false;

for (const shard of selected) {
  const shardDir = path.join(evidenceRoot, `shard-${shard.id}`);
  mkdirSync(shardDir, { recursive: true });
  writeFileSync(
    path.join(shardDir, "environment.json"),
    `${JSON.stringify(
      {
        commit,
        cwd: root,
        node: process.version,
        platform: process.platform,
        arch: process.arch,
        runId,
        shard: shard.id,
        prebuilt,
      },
      null,
      2,
    )}\n`,
  );

  const steps = [];
  if (shard.needsBuild && !prebuilt) steps.push(buildStep);
  steps.push(...shard.steps);

  const results = [];
  let failed = false;

  for (const step of steps) {
    if (failed) {
      results.push({ id: step.id, status: "NOT_RUN", exitCode: null });
      continue;
    }

    const startedAt = new Date().toISOString();
    console.log(`\n[shard:${shard.id}] ${step.id}`);
    const result = spawnSync(step.executable, step.args, {
      cwd: root,
      encoding: "utf8",
      env: { ...baseEnv, ...(step.env || {}) },
      maxBuffer: 64 * 1024 * 1024,
    });
    const output = `${result.stdout || ""}${result.stderr || ""}`;
    if (output) process.stdout.write(output);
    if (result.error) console.error(result.error);

    const status = result.status ?? 1;
    writeFileSync(path.join(shardDir, `${step.id}.log`), output);
    results.push({
      id: step.id,
      command: [step.executable, ...step.args],
      status: status === 0 ? "PASS" : "FAIL",
      exitCode: status,
      startedAt,
      finishedAt: new Date().toISOString(),
    });
    if (status !== 0) failed = true;
  }

  writeFileSync(
    path.join(shardDir, "summary.json"),
    `${JSON.stringify(
      {
        shard: shard.id,
        title: shard.title,
        result: failed ? "FAIL" : "PASS",
        runId,
        commit,
        evidenceRoot: shardDir,
        results,
      },
      null,
      2,
    )}\n`,
  );

  console.log(`\n[shard:${shard.id}] ${failed ? "FAIL" : "PASS"}: ${shardDir}`);
  if (failed) shardFailed = true;
}

process.exit(shardFailed ? 1 : 0);
