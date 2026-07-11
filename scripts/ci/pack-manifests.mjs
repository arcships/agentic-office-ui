#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";

const root = path.resolve(import.meta.dirname, "../..");
const output = path.resolve(
  process.env.PACK_EVIDENCE_DIR ||
    path.join(root, "output", "acceptance", "pack-manifests"),
);
const tgzDir = path.join(output, "tgz");
const packageDirs = [
  "packages/docx-core",
  "packages/xlsx-core",
  "packages/vue-docx",
  "packages/vue-xlsx",
  "packages/vue-pdf",
  "packages/vue-ui",
];

rmSync(output, { recursive: true, force: true });
mkdirSync(tgzDir, { recursive: true });

function collectExportTargets(value, result = []) {
  if (typeof value === "string") result.push(value);
  else if (value && typeof value === "object") {
    for (const child of Object.values(value)) collectExportTargets(child, result);
  }
  return result;
}

const entries = [];
const failures = [];
for (const packageDir of packageDirs) {
  const absoluteDir = path.join(root, packageDir);
  const packed = spawnSync(
    "npm",
    ["pack", "--json", "--pack-destination", tgzDir],
    { cwd: absoluteDir, encoding: "utf8", maxBuffer: 32 * 1024 * 1024 },
  );
  writeFileSync(
    path.join(output, `${path.basename(packageDir)}-npm-pack.log`),
    `${packed.stdout || ""}${packed.stderr || ""}`,
  );
  if (packed.status !== 0) {
    failures.push(`${packageDir}: npm pack exited ${packed.status}`);
    continue;
  }

  let report;
  try {
    const jsonStart = packed.stdout.lastIndexOf("\n[\n");
    const jsonEnd = packed.stdout.lastIndexOf("]");
    if (jsonStart < 0 || jsonEnd < jsonStart) {
      throw new Error("npm pack JSON array was not found in command output");
    }
    report = JSON.parse(packed.stdout.slice(jsonStart + 1, jsonEnd + 1)).at(0);
  } catch (error) {
    failures.push(`${packageDir}: invalid npm pack JSON: ${error.message}`);
    continue;
  }
  const manifest = JSON.parse(
    readFileSync(path.join(absoluteDir, "package.json"), "utf8"),
  );
  const files = new Set(report.files.map((file) => file.path));
  const exportTargets = collectExportTargets(manifest.exports).filter(
    (target) => target.startsWith("./"),
  );
  const missingTargets = exportTargets
    .map((target) => target.slice(2))
    .filter((target) => !files.has(target));
  const hasSource = [...files].some((file) => file.startsWith("src/"));
  const hasTests = [...files].some((file) => /^(test|tests)\//.test(file));
  const hasWorkspaceSpecifier = JSON.stringify(manifest).includes("workspace:");
  const isPublicPackage =
    manifest.private !== true && manifest.publishConfig?.access === "public";
  const isVuePackage = manifest.name.startsWith("@arcships/vue-");
  const isCorePackage = /\/(docx|xlsx)-core$/.test(manifest.name);
  const requiresWasm = isCorePackage || manifest.name === "@arcships/vue-pdf";
  const styleTarget = manifest.exports?.["./style.css"];
  const hasPublicStyle =
    !isVuePackage ||
    (typeof styleTarget === "string" && files.has(styleTarget.replace(/^\.\//, "")));
  const hasWorker = [...files].some((file) => /worker\.js$/.test(file));
  const hasWasm = [...files].some((file) => file.endsWith(".wasm"));

  for (const target of missingTargets) {
    failures.push(`${manifest.name}: missing exported file ${target}`);
  }
  if (hasSource) failures.push(`${manifest.name}: package contains src/`);
  if (hasTests) failures.push(`${manifest.name}: package contains tests/`);
  if (hasWorkspaceSpecifier) failures.push(`${manifest.name}: contains workspace: specifier`);
  if (!isPublicPackage) {
    failures.push(`${manifest.name}: must be publishable with publishConfig.access=public`);
  }
  if (!hasPublicStyle) failures.push(`${manifest.name}: missing public style.css`);
  if (isCorePackage && !hasWorker) failures.push(`${manifest.name}: missing Worker`);
  if (requiresWasm && !hasWasm) failures.push(`${manifest.name}: missing WASM`);

  const tgzPath = path.join(tgzDir, report.filename);
  const resourceFiles = [...new Set([
    manifest.exports?.["./worker"],
    ...Object.entries(manifest.exports || {})
      .filter(([key]) => key.startsWith("./assets/") && key.endsWith(".wasm"))
      .map(([, target]) => target),
  ])]
    .filter((target) => typeof target === "string")
    .map((target) => target.replace(/^\.\//, ""));
  const resources = resourceFiles.map((file) => {
    const extracted = spawnSync("tar", ["-xOf", tgzPath, `package/${file}`], {
      maxBuffer: 16 * 1024 * 1024,
    });
    if (extracted.status !== 0) {
      failures.push(`${manifest.name}: could not hash archive resource ${file}`);
      return { path: file, sha256: null };
    }
    return {
      path: file,
      sha256: createHash("sha256").update(extracted.stdout).digest("hex"),
    };
  });
  entries.push({
    package: manifest.name,
    version: manifest.version,
    filename: report.filename,
    sha256: createHash("sha256").update(readFileSync(tgzPath)).digest("hex"),
    fileCount: report.files.length,
    unpackedSize: report.unpackedSize,
    missingTargets,
    hasSource,
    hasTests,
    hasWorkspaceSpecifier,
    isPublicPackage,
    hasPublicStyle,
    hasWorker,
    hasWasm,
    resources,
    files: report.files,
  });
}

if (entries.length !== packageDirs.length) {
  failures.push(`expected ${packageDirs.length} packages, packed ${entries.length}`);
}
const summary = {
  suite: "P1-CI-CONSUMER-PREFLIGHT",
  result: failures.length === 0 ? "PASS" : "FAIL",
  scope: "real tgz manifests; external installation is expanded by P1-PACK-TEST-01",
  entries,
  failures,
};
writeFileSync(path.join(output, "summary.json"), `${JSON.stringify(summary, null, 2)}\n`);
console.log(`${summary.result}: ${entries.length} real package archives checked at ${output}`);
if (failures.length) {
  for (const failure of failures) console.error(`- ${failure}`);
  process.exit(1);
}
