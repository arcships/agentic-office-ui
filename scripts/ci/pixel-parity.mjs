#!/usr/bin/env node

/**
 * Extend UI <-> @arcships/vue-ui pixel-parity harness entry point.
 *
 *   pnpm ui:parity                    # tokens + build + full matrix + gate
 *   pnpm ui:parity -- --only button,file-upload --viewports 1440x900
 *   pnpm ui:parity -- --only core --gate-mode ready --baseline
 *   pnpm ui:parity:report             # always exit 0 (artifact for a PR)
 *
 * Modes (see docs/testing/ui-parity-ci.md):
 *   --gate-mode all|ready|off   `all` (default) fails on every parity failure,
 *                               `ready` only on surfaces with
 *                               `gateStatus: blocking`, `off` never fails
 *   --baseline                  diff the subject captures against the
 *                               committed baselines in tests/visual/extend
 *   --write-baseline            write captures into the baseline tree; use
 *                               `pnpm ui:baseline` so the approval flow and
 *                               manifest hashes stay in sync
 *
 * Steps:
 *   1. `docs/design/extend-tokens.json` is regenerated (and must stay in sync
 *      with the pinned Extend CSS).
 *   2. Workspace packages the reference apps import by package name are built
 *      when their dist is missing, then both harness apps are built.
 *   3. scripts/ci/pixel-parity.py drives Chromium at DPR 2 and writes
 *      output/ui-parity/{report.json,report.html,SUMMARY.md,cases/**}.
 *
 * The Python interpreter is resolved by scripts/ci/run-python.mjs (the same
 * helper the other browser suites use), so `PYTHON=` or a venv on PATH works.
 */
import { spawnSync } from "node:child_process";
import path from "node:path";
import process from "node:process";
import { buildReferenceApp } from "./reference-build.mjs";

const root = path.resolve(import.meta.dirname, "../..");
const pythonScript = path.join(root, "scripts/ci/pixel-parity.py");

/**
 * The comparison app consumes `packages/vue-*` sources; the shared helper also
 * builds the `@arcships/*` dist those sources import by package name.
 */
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
    gateMode: undefined,
    baseline: false,
    writeBaseline: false,
    baselineDir: undefined,
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
    else if (arg === "--gate-mode") options.gateMode = next();
    else if (arg === "--baseline") options.baseline = true;
    else if (arg === "--write-baseline") options.writeBaseline = true;
    else if (arg === "--baseline-dir") options.baselineDir = next();
    else if (arg.startsWith("--only=")) options.only = arg.slice(7);
    else if (arg.startsWith("--viewports=")) options.viewports = arg.slice(11);
    else if (arg.startsWith("--themes=")) options.themes = arg.slice(9);
    else if (arg.startsWith("--out=")) options.out = path.resolve(arg.slice(6));
    else if (arg.startsWith("--gate-mode=")) options.gateMode = arg.slice(12);
    else if (arg.startsWith("--baseline-dir=")) options.baselineDir = arg.slice(15);
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
  if (!buildReferenceApp(filter)) process.exit(1);
}

function main() {
  const options = parseArgs(process.argv.slice(2));

  if (options.tokens) {
    console.log("▸ extract Extend tokens");
    const status = run(process.execPath, ["scripts/ci/extend-tokens.mjs"]);
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
  if (options.gateMode) pythonArgs.push("--gate-mode", options.gateMode);
  if (options.baseline) pythonArgs.push("--baseline");
  if (options.writeBaseline) pythonArgs.push("--write-baseline");
  if (options.baselineDir) pythonArgs.push("--baseline-dir", options.baselineDir);
  pythonArgs.push(...options.passthrough);

  console.log("\n▸ capture + diff");
  const status = run(process.execPath, pythonArgs, {
    RUN_PYTHON_MODULES: process.env.RUN_PYTHON_MODULES || "playwright,PIL",
  });
  console.log(`\nartifacts: ${path.relative(root, options.out)}/report.html`);
  process.exit(status);
}

main();
