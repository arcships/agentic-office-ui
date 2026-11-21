#!/usr/bin/env node

/**
 * Core component accessibility gate (#165): axe-core WCAG A/AA rules plus a
 * keyboard smoke test (Tab reachability, visible focus ring, accessible name,
 * select open/close) against the built parity comparison app.
 *
 *   pnpm test:a11y                    # gate; fails on unlisted violations
 *   pnpm test:a11y:report             # same run, always exit 0 (PR artifact)
 *   pnpm test:a11y -- --only button,input --theme dark
 *
 * New violations either get fixed or recorded in `tests/a11y/allowlist.json`
 * with an owner issue and a reason. `pnpm test:a11y -- --build` builds the
 * comparison app first (CI builds it as an explicit step).
 */
import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import path from "node:path";
import process from "node:process";
import { buildReferenceApp } from "./reference-build.mjs";

const root = path.resolve(import.meta.dirname, "../..");
const dist = path.join(root, "tools/vue-reference/dist");

const options = {
  build: false,
  passthrough: [],
};
const argv = process.argv.slice(2);
for (let index = 0; index < argv.length; index += 1) {
  const arg = argv[index];
  if (arg === "--build") options.build = true;
  else if (arg === "--") continue;
  else options.passthrough.push(arg);
}

if (!existsSync(path.join(dist, "index.html"))) {
  if (!options.build) {
    console.error(
      [
        "test:a11y: tools/vue-reference/dist/index.html is missing.",
        "Build it first (`pnpm --filter @aou/vue-reference build`), or pass --build.",
      ].join("\n"),
    );
    process.exit(2);
  }
}
if (options.build) {
  const built = buildReferenceApp("@aou/vue-reference");
  if (!built) process.exit(1);
}

const status = spawnSync(
  process.execPath,
  [
    "scripts/ci/run-python.mjs",
    "scripts/ci/a11y-check.py",
    ...options.passthrough,
  ],
  {
    cwd: root,
    stdio: "inherit",
    env: { ...process.env, RUN_PYTHON_MODULES: process.env.RUN_PYTHON_MODULES || "playwright" },
  },
).status ?? 1;
process.exit(status);
