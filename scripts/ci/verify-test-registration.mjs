#!/usr/bin/env node

/**
 * Guards the explicit `node --test` file lists in scripts/ci/run-suite.mjs.
 *
 * Node is given file paths (no glob), so a new test file that nobody registers
 * is dead code that still looks like coverage. Issue #62 lost two component
 * files that way; every tests/unit and tests/component file must appear in a
 * suite, and every registered path must still exist.
 */

import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";

const root = path.resolve(import.meta.dirname, "../..");
const runSuitePath = path.join(root, "scripts", "ci", "run-suite.mjs");
const source = readFileSync(runSuitePath, "utf8");

const registered = new Set();
for (const match of source.matchAll(/"((?:tests|packages)\/[^"]+\.(?:test\.mjs|mjs))"/g)) {
  registered.add(match[1]);
}

const directories = ["tests/unit", "tests/component"];
const failures = [];
const discovered = [];

for (const directory of directories) {
  const absolute = path.join(root, directory);
  for (const entry of readdirSync(absolute, { withFileTypes: true })) {
    if (!entry.isFile() || !entry.name.endsWith(".test.mjs")) continue;
    const relative = `${directory}/${entry.name}`;
    discovered.push(relative);
    if (!registered.has(relative)) {
      failures.push(`${relative}: not registered in scripts/ci/run-suite.mjs`);
    }
  }
}

for (const file of registered) {
  if (!existsSync(path.join(root, file))) {
    failures.push(`${file}: registered in run-suite.mjs but missing on disk`);
  }
}

discovered.sort();
const summary = {
  suite: "P1-CI-TEST-REGISTRATION",
  result: failures.length ? "FAIL" : "PASS",
  registered: registered.size,
  discovered: discovered.length,
  files: discovered,
  failures,
};
const evidenceDir = process.env.TEST_REGISTRATION_EVIDENCE_DIR;
if (evidenceDir) {
  mkdirSync(evidenceDir, { recursive: true });
  writeFileSync(
    path.join(evidenceDir, "summary.json"),
    `${JSON.stringify(summary, null, 2)}\n`,
  );
}

if (failures.length) {
  console.error(`test registration FAIL (${failures.length} problem(s)):`);
  for (const failure of failures) console.error(`  ${failure}`);
  process.exit(1);
}

console.log(
  `test registration PASS: ${discovered.length} files wired into ${registered.size} suite paths`,
);
