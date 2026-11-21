#!/usr/bin/env node

/**
 * L2 product-route parity entry point (issue #182).
 *
 *   pnpm ui:parity:routes                    # report-only matrix for every route
 *   pnpm ui:parity:routes -- --only pdf-viewer --themes light
 *   pnpm ui:parity:routes -- --gate-mode ready
 *   pnpm ui:parity:routes -- --self-check    # reference vs itself (determinism)
 *
 * Builds the pinned upstream reference app (`tools/extend-reference`, route
 * entries under `route.html`) and the subject demo (`apps/demo`), then runs
 * `scripts/ci/route-parity.py`, which serves both builds, captures the product
 * region for every route x viewport x theme and writes
 * `output/ui-parity-routes/{report.json,report.html,SUMMARY.md,routes-matrix.json}`
 * plus reference/subject/diff/heatmap PNGs per case.
 *
 * Region selectors, exclusions and classified differences live in
 * `tools/parity-fixtures/routes.json`; the rationale is in
 * `docs/testing/ui-parity-routes.md`.
 */
import { spawnSync } from "node:child_process";
import path from "node:path";
import process from "node:process";

import { buildReferenceApp } from "./reference-build.mjs";

const root = path.resolve(import.meta.dirname, "../..");

function parseArgs(argv) {
  const options = {
    build: true,
    out: process.env.UI_PARITY_ROUTES_OUT || path.join(root, "output/ui-parity-routes"),
    passthrough: [],
  };
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    const next = () => argv[++index];
    if (arg === "--") continue;
    else if (arg === "--no-build") options.build = false;
    else if (arg === "--out") options.out = path.resolve(next());
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

function main() {
  const options = parseArgs(process.argv.slice(2));

  if (options.build) {
    // The demo route imports the workspace packages by name, exactly like
    // `pnpm dev`; building the dependency tree keeps the subject in sync with
    // the sources under test instead of a stale dist.
    if (!buildReferenceApp("@aou/extend-reference")) process.exit(1);
    console.log("\n▸ build demo");
    if (run("pnpm", ["--filter", "demo...", "build"]) !== 0) process.exit(1);
  }

  const pythonArgs = [
    "scripts/ci/run-python.mjs",
    path.join(root, "scripts/ci/route-parity.py"),
    "--out",
    options.out,
    ...options.passthrough,
  ];
  const status = run(process.execPath, pythonArgs, {
    RUN_PYTHON_MODULES: process.env.RUN_PYTHON_MODULES || "playwright,PIL",
  });
  console.log(
    `\nartifacts: ${path.relative(root, options.out)}/report.html`,
  );
  process.exit(status);
}

main();
