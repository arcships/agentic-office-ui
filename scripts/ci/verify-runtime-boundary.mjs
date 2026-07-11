#!/usr/bin/env node

import { mkdirSync, readFileSync, readdirSync, statSync, writeFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";

const root = path.resolve(import.meta.dirname, "../..");
const output = path.resolve(
  process.env.RUNTIME_BOUNDARY_EVIDENCE_DIR ||
    path.join(root, "output", "acceptance", "runtime-boundary"),
);
const runtimeRoot = path.join(root, "packages", "office-runtime");
const publicPackages = ["docx-core", "xlsx-core", "vue-docx", "vue-xlsx", "vue-pdf", "vue-ui"];
const failures = [];

function walk(directory) {
  return readdirSync(directory).flatMap((name) => {
    const absolute = path.join(directory, name);
    return statSync(absolute).isDirectory() ? walk(absolute) : [absolute];
  });
}

const manifest = JSON.parse(readFileSync(path.join(runtimeRoot, "package.json"), "utf8"));
if (manifest.private !== true) failures.push("office-runtime must be private");
if (manifest.publishConfig !== undefined) failures.push("office-runtime must not define publishConfig");
if (/vue|@vue|vue-pdf|vue-ui|vue-docx|vue-xlsx/.test(JSON.stringify(manifest))) {
  failures.push("office-runtime manifest contains a Vue dependency");
}

const forbiddenSource = /(?:from\s+["']vue["']|\.vue["']|\bwindow\b|\bdocument\b|\bHTMLElement\b|\bHTMLCanvasElement\b|\bImageBitmap\b)/;
for (const file of walk(path.join(runtimeRoot, "src"))) {
  if (!file.endsWith(".ts")) continue;
  const source = readFileSync(file, "utf8");
  if (forbiddenSource.test(source)) failures.push(`forbidden UI dependency in ${path.relative(root, file)}`);
}

const privateReferenceFiles = [];
for (const packageName of publicPackages) {
  const packageRoot = path.join(root, "packages", packageName);
  const candidates = [path.join(packageRoot, "package.json"), ...walk(path.join(packageRoot, "dist"))]
    .filter((file) => /(?:package\.json|\.(?:js|d\.ts))$/.test(file));
  for (const file of candidates) {
    if (readFileSync(file, "utf8").includes("@arcships/office-runtime")) {
      privateReferenceFiles.push(path.relative(root, file));
    }
  }
}
if (privateReferenceFiles.length) {
  failures.push(`public artifacts reference private runtime: ${privateReferenceFiles.join(", ")}`);
}

const packScript = readFileSync(path.join(root, "scripts", "ci", "pack-manifests.mjs"), "utf8");
if (packScript.includes('"packages/office-runtime"')) {
  failures.push("private runtime was added to the public pack list");
}

const runtime = await import(path.join(runtimeRoot, "dist", "index.js"));
const sequenceA = runtime.createOfficeTaskSequence("boundary-a");
const sequenceB = runtime.createOfficeTaskSequence("boundary-b");
const instanceIsolation = [sequenceA.next(), sequenceA.next(), sequenceB.next()];
if (instanceIsolation.join(",") !== "boundary-a:1,boundary-a:2,boundary-b:1") {
  failures.push(`task sequences are not instance-owned: ${instanceIsolation.join(",")}`);
}
sequenceA.dispose();
sequenceB.dispose();

mkdirSync(output, { recursive: true });
const summary = {
  suite: "P2-RUNTIME-BOUNDARY",
  result: failures.length ? "FAIL" : "PASS",
  private: manifest.private === true,
  publicPackages,
  privateReferenceFiles,
  instanceIsolation,
  failures,
};
writeFileSync(path.join(output, "summary.json"), `${JSON.stringify(summary, null, 2)}\n`);
console.log(`${summary.result}: office runtime boundary checked at ${output}`);
if (failures.length) {
  for (const failure of failures) console.error(`- ${failure}`);
  process.exit(1);
}
