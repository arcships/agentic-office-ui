#!/usr/bin/env node

/**
 * Build helpers for the two parity harness apps (`tools/extend-reference`,
 * `tools/vue-reference`).
 *
 * The comparison app imports `packages/vue-*` sources, and those sources
 * import `@arcships/*` core packages by package name. Their dist has to exist
 * before vite resolves the app, otherwise a clean checkout cannot build it —
 * which is why the harness builds these workspace packages itself instead of
 * relying on a repository-wide build.
 */
import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import path from "node:path";
import process from "node:process";

const root = path.resolve(import.meta.dirname, "../..");

const workspaceDependencies = [
  { filter: "@arcships/office-runtime", probe: "packages/office-runtime/dist/index.js" },
  { filter: "@arcships/docx-core...", probe: "packages/docx-core/dist/index.js" },
  { filter: "@arcships/xlsx-core", probe: "packages/xlsx-core/dist/index.js" },
];

function runBuild(filter) {
  console.log(`\n▸ build ${filter}`);
  const result = spawnSync("pnpm", ["--filter", filter, "build"], {
    cwd: root,
    stdio: "inherit",
    env: process.env,
  });
  if (result.error) throw result.error;
  if ((result.status ?? 1) !== 0) {
    console.error(`build failed: ${filter}`);
    return false;
  }
  return true;
}

/** Build workspace packages the reference apps import by name (once). */
export function buildWorkspaceDependencies() {
  for (const { filter, probe } of workspaceDependencies) {
    if (existsSync(path.join(root, probe))) {
      console.log(`▸ ${filter}: dist present, skipping`);
      continue;
    }
    if (!runBuild(filter)) return false;
  }
  return true;
}

/** Build one harness app (`@aou/extend-reference` or `@aou/vue-reference`). */
export function buildReferenceApp(filter) {
  return buildWorkspaceDependencies() && runBuild(filter);
}
