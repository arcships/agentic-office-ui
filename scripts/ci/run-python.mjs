#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import process from "node:process";

const args = process.argv.slice(2);
if (args.length === 0) {
  console.error("Usage: node scripts/ci/run-python.mjs <script.py> [...args]");
  process.exit(2);
}

const candidates = [
  process.env.PYTHON,
  process.env.CONDA_PREFIX && path.join(process.env.CONDA_PREFIX, "bin", "python"),
  "python3",
  "python",
  path.join(os.homedir(), "miniconda3", "bin", "python"),
].filter(Boolean);

let selected = null;
for (const candidate of [...new Set(candidates)]) {
  if (candidate.includes(path.sep) && !existsSync(candidate)) continue;
  const probe = spawnSync(candidate, ["-c", "import playwright"], {
    encoding: "utf8",
  });
  if (probe.status === 0) {
    selected = candidate;
    break;
  }
}

if (!selected) {
  console.error(
    "No Python interpreter with Playwright was found. Install requirements-ci.txt first.",
  );
  process.exit(2);
}

console.log(`Using Python: ${selected}`);
const result = spawnSync(selected, args, {
  stdio: "inherit",
  env: process.env,
});
if (result.error) console.error(result.error);
process.exit(result.status ?? 1);
