#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";

const root = path.resolve(import.meta.dirname, "../..");
const headResult = spawnSync("git", ["rev-parse", "HEAD"], {
  cwd: root,
  encoding: "utf8",
});
const commit = headResult.status === 0 ? headResult.stdout.trim() : "no-git-head";
const runId = process.env.CI_RUN_ID || new Date().toISOString().replace(/[-:]/g, "");
const output = path.resolve(
  process.env.COMPATIBILITY_EVIDENCE_DIR ||
    path.join(root, "output", "acceptance", commit, "P4-MATRIX-01", runId, "matrix"),
);
const suppliedManifestPath = process.env.COMPATIBILITY_MANIFEST_PATH
  ? path.resolve(process.env.COMPATIBILITY_MANIFEST_PATH)
  : null;
const manifestDir = suppliedManifestPath
  ? path.dirname(suppliedManifestPath)
  : path.join(output, "real-tgz-manifests");
const manifestPath = suppliedManifestPath || path.join(manifestDir, "summary.json");
const browsers = (process.env.COMPATIBILITY_BROWSERS || "chromium,firefox,webkit")
  .split(",")
  .map((value) => value.trim())
  .filter(Boolean);
const requestedStackIds = (process.env.COMPATIBILITY_STACKS || "current,vue-minimum")
  .split(",")
  .map((value) => value.trim())
  .filter(Boolean);
const requestedStacks = new Set(requestedStackIds);
const allowPartial = process.env.COMPATIBILITY_ALLOW_PARTIAL === "1";
const allStacks = [
  {
    id: "current",
    versions: {
      vue: "3.5.39",
      vite: "6.4.3",
      typescript: "5.9.3",
      vueTsc: "3.3.6",
      viteVuePlugin: "5.2.4",
    },
  },
  {
    id: "vue-minimum",
    versions: {
      vue: "3.2.25",
      vite: "6.4.3",
      typescript: "5.9.3",
      vueTsc: "3.3.6",
      viteVuePlugin: "5.2.4",
    },
  },
];
const stacks = allStacks.filter((stack) => requestedStacks.has(stack.id));
const allowedBrowsers = new Set(["chromium", "firefox", "webkit"]);
const allowedStacks = new Set(allStacks.map((stack) => stack.id));
const requiredBrowsers = ["chromium", "firefox", "webkit"];
const requiredStackIds = ["current", "vue-minimum"];
const isFullMatrix =
  browsers.length === requiredBrowsers.length &&
  new Set(browsers).size === browsers.length &&
  requiredBrowsers.every((browser) => browsers.includes(browser)) &&
  requestedStackIds.length === requiredStackIds.length &&
  requestedStacks.size === requestedStackIds.length &&
  requiredStackIds.every((stack) => requestedStacks.has(stack));

rmSync(output, { recursive: true, force: true });
mkdirSync(output, { recursive: true });

const commands = [];
const failures = [];

function run(id, executable, args, env = {}) {
  const startedAt = new Date().toISOString();
  const result = spawnSync(executable, args, {
    cwd: root,
    encoding: "utf8",
    env: { ...process.env, ...env },
    maxBuffer: 128 * 1024 * 1024,
  });
  const outputText = `${result.stdout || ""}${result.stderr || ""}${result.error ? `\n${result.error.stack || result.error.message}` : ""}`;
  writeFileSync(path.join(output, `${id}.log`), outputText);
  commands.push({
    id,
    command: [executable, ...args],
    exitCode: result.status ?? 1,
    startedAt,
    finishedAt: new Date().toISOString(),
    error: result.error?.message || null,
  });
  if (result.status !== 0) {
    failures.push(`${id} exited ${result.status ?? 1}`);
    return false;
  }
  return true;
}

function readJson(file, label) {
  try {
    return JSON.parse(readFileSync(file, "utf8"));
  } catch (error) {
    failures.push(`${label} could not be read: ${error instanceof Error ? error.message : String(error)}`);
    return null;
  }
}

if (
  browsers.length === 0 ||
  new Set(browsers).size !== browsers.length ||
  browsers.some((browser) => !allowedBrowsers.has(browser)) ||
  stacks.length === 0 ||
  requestedStacks.size !== requestedStackIds.length ||
  [...requestedStacks].some((stack) => !allowedStacks.has(stack))
) {
  failures.push(
    `invalid matrix selection: browsers=${JSON.stringify(browsers)}, stacks=${JSON.stringify([...requestedStacks])}`,
  );
}
if (failures.length === 0 && !isFullMatrix && !allowPartial) {
  failures.push("partial compatibility selection requires COMPATIBILITY_ALLOW_PARTIAL=1");
}

if (failures.length === 0 && process.env.CI_PREBUILT !== "1") {
  run("build", "pnpm", ["build"]);
}
if (failures.length === 0 && !suppliedManifestPath) {
  run("pack-manifests", process.execPath, ["scripts/ci/pack-manifests.mjs"], {
    PACK_EVIDENCE_DIR: manifestDir,
  });
}

let manifest = null;
if (failures.length === 0) {
  if (!existsSync(manifestPath)) {
    failures.push(`candidate manifest does not exist: ${manifestPath}`);
  } else {
    manifest = readJson(manifestPath, "candidate manifest");
  }
  const publicApi = readJson(
    path.join(root, "scripts", "ci", "public-api-contract.json"),
    "public API contract",
  );
  const expectedPackages = publicApi?.packages?.map((entry) => entry.name).sort() ?? [];
  const manifestPackages = manifest?.entries?.map((entry) => entry.package).sort() ?? [];
  if (
    !manifest ||
    manifest.result !== "PASS" ||
    expectedPackages.length === 0 ||
    JSON.stringify(manifestPackages) !== JSON.stringify(expectedPackages)
  ) {
    failures.push(
      `public-package tgz manifest did not pass: ${JSON.stringify({ expectedPackages, manifestPackages })}`,
    );
  }
}

const results = [];
const preflightPassed = failures.length === 0;
for (const [index, stack] of stacks.entries()) {
  if (!preflightPassed) break;
  const stackDir = path.join(output, stack.id);
  const ok = run(
    `consumer-${stack.id}`,
    process.execPath,
    ["scripts/ci/pack-consumer.mjs"],
    {
      PACK_MANIFEST_PATH: manifestPath,
      PACK_CONSUMER_EVIDENCE_DIR: stackDir,
      PACK_CONSUMER_BROWSERS: browsers.join(","),
      PACK_CONSUMER_PORT: String(4291 + index),
      PACK_CONSUMER_VUE_VERSION: stack.versions.vue,
      PACK_CONSUMER_VITE_VERSION: stack.versions.vite,
      PACK_CONSUMER_TYPESCRIPT_VERSION: stack.versions.typescript,
      PACK_CONSUMER_VUE_TSC_VERSION: stack.versions.vueTsc,
      PACK_CONSUMER_VITE_VUE_PLUGIN_VERSION: stack.versions.viteVuePlugin,
    },
  );
  const consumerSummaryPath = path.join(stackDir, "summary.json");
  const consumer = existsSync(consumerSummaryPath)
    ? readJson(consumerSummaryPath, `${stack.id} consumer summary`) || { result: "FAIL", failure: "consumer summary was invalid" }
    : { result: "FAIL", failure: "consumer summary was not written" };
  const browserResults = browsers.map((browser) => {
    const evidenceName = browsers.length === 1 && browser === "chromium"
      ? "browser"
      : `browser-${browser}`;
    const resultPath = path.join(stackDir, evidenceName, "result.json");
    return existsSync(resultPath)
      ? readJson(resultPath, `${stack.id}/${browser} result`) || {
          browserName: browser,
          browser: null,
          browserSource: null,
          status: "NOT_RUN",
          checks: {},
        }
      : {
          browserName: browser,
          browser: null,
          browserSource: null,
          status: "NOT_RUN",
          checks: {},
        };
  });
  results.push({
    stack: stack.id,
    versions: stack.versions,
    result: ok && consumer.result === "PASS" && browserResults.every((item) => item.status === "PASS")
      ? "PASS"
      : "FAIL",
    consumer,
    browsers: browserResults.map((item) => ({
      name: item.browserName,
      version: item.browser,
      source: item.browserSource,
      status: item.status,
      checks: item.checks,
    })),
  });
}

if (results.some((entry) => entry.result !== "PASS")) {
  failures.push("at least one declared compatibility combination failed");
}
if (preflightPassed && results.length !== stacks.length) {
  failures.push(`only ${results.length}/${stacks.length} compatibility stacks produced results`);
}

const summaryResult = failures.length > 0 ? "FAIL" : isFullMatrix ? "PASS" : "PARTIAL_PASS";

const summary = {
  suite: "P4-COMPATIBILITY-MATRIX",
  result: summaryResult,
  releaseEligible: summaryResult === "PASS" && isFullMatrix,
  commit,
  environment: {
    node: process.version,
    platform: process.platform,
    arch: process.arch,
    locale: "zh-CN",
    timezone: "Asia/Shanghai",
    viewport: "1440x900",
    deviceScaleFactor: 1,
  },
  requiredBrowsers,
  requiredStacks: requiredStackIds,
  selectedBrowsers: browsers,
  selectedStacks: stacks.map((stack) => stack.id),
  allowPartial,
  isFullMatrix,
  manifestPath,
  reusedCandidateManifest: Boolean(suppliedManifestPath),
  packageArchives: manifest?.entries?.map((entry) => ({
    package: entry.package,
    version: entry.version,
    filename: entry.filename,
    sha256: entry.sha256,
  })) || [],
  commands,
  results,
  failures,
};
writeFileSync(path.join(output, "summary.json"), `${JSON.stringify(summary, null, 2)}\n`);
console.log(`${summary.result}: compatibility matrix at ${output}`);
for (const failure of failures) console.error(`- ${failure}`);
process.exit(summary.result === "FAIL" ? 1 : 0);
