#!/usr/bin/env node

import { createHash } from "node:crypto";
import {
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  writeFileSync,
} from "node:fs";
import path from "node:path";
import process from "node:process";

const root = path.resolve(import.meta.dirname, "../..");
const manifestPath = path.resolve(process.argv[2] || path.join(root, "test-data", "manifest.json"));
const output = path.resolve(
  process.env.FIXTURE_EVIDENCE_DIR ||
    path.join(root, "output", "acceptance", "fixture-manifest"),
);
mkdirSync(output, { recursive: true });

const failures = [];
const manifest = JSON.parse(readFileSync(manifestPath, "utf8"));
const samplesRoot = path.resolve(root, manifest.root || "");
const categories = new Set();
const seenFiles = new Set();
const byHash = new Map();
const results = [];

if (manifest.version !== 1) failures.push("manifest version must be 1");
if (manifest.containsSensitiveData !== false) {
  failures.push("manifest must explicitly declare containsSensitiveData=false");
}
if (!Array.isArray(manifest.entries) || manifest.entries.length === 0) {
  failures.push("manifest entries must be non-empty");
}

for (const [sourceId, source] of Object.entries(manifest.sources || {})) {
  if (source.kind !== "generated") failures.push(`${sourceId}: source must be generated`);
  const generator = path.resolve(root, source.generator || "");
  if (!existsSync(generator)) failures.push(`${sourceId}: generator does not exist`);
  else if (!readFileSync(generator, "utf8").includes(`random.seed(${source.seed})`)) {
    failures.push(`${sourceId}: generator does not declare seed ${source.seed}`);
  }
}

for (const entry of manifest.entries || []) {
  const entryFailures = [];
  if (!entry.file || path.basename(entry.file) !== entry.file) {
    entryFailures.push("file must be a basename");
  }
  if (seenFiles.has(entry.file)) entryFailures.push("duplicate file entry");
  seenFiles.add(entry.file);
  categories.add(entry.category);
  if (!manifest.sources?.[entry.source]) entryFailures.push("unknown source");
  if (!entry.expected) entryFailures.push("missing expected behavior");
  if (!/^[a-f0-9]{64}$/.test(entry.sha256 || "")) entryFailures.push("invalid sha256");

  const filePath = path.resolve(samplesRoot, entry.file || "");
  if (!filePath.startsWith(`${samplesRoot}${path.sep}`) || !existsSync(filePath)) {
    entryFailures.push("file missing or outside fixture root");
  } else {
    const bytes = readFileSync(filePath);
    const sha256 = createHash("sha256").update(bytes).digest("hex");
    if (bytes.byteLength !== entry.bytes) entryFailures.push(`bytes ${bytes.byteLength} != ${entry.bytes}`);
    if (sha256 !== entry.sha256) entryFailures.push(`sha256 ${sha256} != ${entry.sha256}`);
    const firstEight = bytes.subarray(0, 8);
    const isCorrupted = entry.file.startsWith("corrupted.");
    if ((entry.type === "docx" || entry.type === "xlsx") && !isCorrupted) {
      if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) entryFailures.push("missing ZIP signature");
    } else if (entry.type === "pdf" && !isCorrupted) {
      if (bytes.subarray(0, 4).toString("ascii") !== "%PDF") entryFailures.push("missing PDF signature");
    } else if (entry.type === "png") {
      if (!firstEight.equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]))) entryFailures.push("missing PNG signature");
    } else if (entry.type === "json") {
      try {
        JSON.parse(bytes.toString("utf8"));
      } catch {
        entryFailures.push("invalid JSON");
      }
    }
    const group = byHash.get(sha256) || [];
    group.push(entry);
    byHash.set(sha256, group);
  }
  for (const failure of entryFailures) failures.push(`${entry.file}: ${failure}`);
  results.push({ file: entry.file, result: entryFailures.length ? "FAIL" : "PASS", failures: entryFailures });
}

const actualFiles = readdirSync(samplesRoot)
  .filter((name) => name !== "manifest.json")
  .sort();
const declaredFiles = [...seenFiles].sort();
if (JSON.stringify(actualFiles) !== JSON.stringify(declaredFiles)) {
  failures.push(`fixture file set differs: actual=${actualFiles.join(",")} declared=${declaredFiles.join(",")}`);
}

for (const group of byHash.values()) {
  if (group.length < 2) continue;
  const primary = group.find((entry) => !entry.aliasOf);
  if (!primary) {
    failures.push(`duplicate sha ${group[0].sha256} has no primary entry`);
    continue;
  }
  for (const entry of group) {
    if (entry === primary) continue;
    if (entry.aliasOf !== primary.file) {
      failures.push(`${entry.file}: duplicate content must declare aliasOf=${primary.file}`);
    }
  }
}

for (const required of ["normal", "boundary", "malicious", "large"]) {
  if (!categories.has(required)) failures.push(`missing fixture category ${required}`);
}

const summary = {
  suite: "P1-FIXTURE-01",
  result: failures.length ? "FAIL" : "PASS",
  manifest: manifestPath,
  fixtureRoot: samplesRoot,
  counts: {
    entries: manifest.entries?.length || 0,
    categories: [...categories].sort(),
  },
  results,
  failures,
};
writeFileSync(path.join(output, "summary.json"), `${JSON.stringify(summary, null, 2)}\n`);
console.log(`${summary.result}: ${summary.counts.entries} fixtures verified at ${output}`);
if (failures.length) {
  for (const failure of failures) console.error(`- ${failure}`);
  process.exit(1);
}
