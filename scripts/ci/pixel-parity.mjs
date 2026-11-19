#!/usr/bin/env node

/**
 * Extend UI <-> @arcships/vue-ui pixel-parity harness entry point.
 *
 *   pnpm ui:parity                    # tokens + build + full matrix + gate
 *   pnpm ui:parity -- --only button,file-upload --viewports 1440x900
 *   pnpm ui:parity:report             # same, always exit 0 (artifact for a PR)
 *
 * Steps:
 *   1. `docs/design/extend-tokens.json` is regenerated (and must stay in sync
 *      with the pinned Extend CSS).
 *   2. Both harness apps are built with their package build script.
 *   3. scripts/ci/pixel-parity.py drives Chromium at DPR 2 and writes
 *      output/ui-parity/{report.json,report.html,SUMMARY.md,cases/**}.
 *
 * The Python interpreter is resolved by scripts/ci/run-python.mjs (the same
 * helper the other browser suites use), so `PYTHON=` or a venv on PATH works.
 */
import { spawnSync } from "node:child_process";
import path from "node:path";
import process from "node:process";

const root = path.resolve(import.meta.dirname, "../..");
const pythonScript = path.join(root, "scripts/ci/pixel-parity.py");

function parseArgs(argv) {
  const options = {
    only: undefined,
    viewports: undefined,
    themes: undefined,
    out: process.env.UI_PARITY_OUT || path.join(root, "output/ui-parity"),
    build: true,
    tokens: true,
    reportOnly: false,
    selfCheck: false,
    passthrough: [],
  };
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    const next = () => argv[++index];
    if (arg === "--") continue;
    else if (arg === "--only") options.only = next();
    else if (arg === "--viewports") options.viewports = next();
    else if (arg === "--themes") options.themes = next();
    else if (arg === "--out") options.out = path.resolve(next());
    else if (arg === "--no-build") options.build = false;
    else if (arg === "--skip-tokens") options.tokens = false;
    else if (arg === "--report-only") options.reportOnly = true;
    else if (arg === "--self-check") options.selfCheck = true;
    else if (arg.startsWith("--only=")) options.only = arg.slice(7);
    else if (arg.startsWith("--viewports=")) options.viewports = arg.slice(11);
    else if (arg.startsWith("--themes=")) options.themes = arg.slice(9);
    else if (arg.startsWith("--out=")) options.out = path.resolve(arg.slice(6));
    else options.passthrough.push(arg);
  }
  return options;
}

function run(command, args, extraEnv = {}) {
  const result = spawnSync(command, args, {
    cwd: root,
    stdio: "inherit",
    env: { ...process.env, ...extraEnv },
  });
  if (result.error) throw result.error;
  return result.status ?? 1;
}

function build(filter) {
  console.log(`\n▸ build ${filter}`);
  const status = run("pnpm", ["--filter", filter, "build"]);
  if (status !== 0) {
    console.error(`build failed: ${filter}`);
    process.exit(status);
  }
}

function main() {
  const options = parseArgs(process.argv.slice(2));

  if (options.tokens) {
    console.log("▸ extract Extend tokens");
    const status = run(process.execPath, ["scripts/ci/extract-extend-tokens.mjs"]);
    if (status !== 0) process.exit(status);
  }

  if (options.build) {
    build("@aou/extend-reference");
    build("@aou/vue-reference");
  }

  const pythonArgs = [
    "scripts/ci/run-python.mjs",
    pythonScript,
    "--out",
    options.out,
  ];
  if (options.only) pythonArgs.push("--only", options.only);
  if (options.viewports) pythonArgs.push("--viewports", options.viewports);
  if (options.themes) pythonArgs.push("--themes", options.themes);
  if (options.reportOnly) pythonArgs.push("--report-only");
  if (options.selfCheck) pythonArgs.push("--self-check");
  pythonArgs.push(...options.passthrough);

  console.log("\n▸ capture + diff");
  const status = run(process.execPath, pythonArgs, {
    RUN_PYTHON_MODULES: process.env.RUN_PYTHON_MODULES || "playwright,PIL",
  });
  console.log(`\nartifacts: ${path.relative(root, options.out)}/report.html`);
  process.exit(status);
}

main();
