#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const root = path.resolve(import.meta.dirname, "../..");
const expectedPackages = [
  "@arcships/docx-core",
  "@arcships/xlsx-core",
  "@arcships/vue-docx",
  "@arcships/vue-xlsx",
  "@arcships/vue-pdf",
  "@arcships/vue-ui",
];
const requiredSuites = [
  "unit",
  "component",
  "blackbox",
  "stress",
  "docs",
];

function fail(message) {
  throw new Error(message);
}

function readJson(file, label = file) {
  if (!existsSync(file)) fail(`${label} is missing: ${file}`);
  try {
    return JSON.parse(readFileSync(file, "utf8"));
  } catch (error) {
    fail(`${label} is not valid JSON: ${error instanceof Error ? error.message : String(error)}`);
  }
}

function hashFile(file, algorithm = "sha256", encoding = "hex") {
  return createHash(algorithm).update(readFileSync(file)).digest(encoding);
}

function integrity(file) {
  return `sha512-${hashFile(file, "sha512", "base64")}`;
}

function gitHead() {
  const result = spawnSync("git", ["rev-parse", "HEAD"], {
    cwd: root,
    encoding: "utf8",
  });
  if (result.status !== 0) fail("git rev-parse HEAD failed");
  return result.stdout.trim();
}

function inspectArchive(file) {
  const result = spawnSync("tar", ["-xOf", file, "package/package.json"], {
    encoding: "utf8",
    maxBuffer: 4 * 1024 * 1024,
  });
  if (result.status !== 0) fail(`cannot read package/package.json from ${file}`);
  try {
    return JSON.parse(result.stdout);
  } catch (error) {
    fail(`archive package.json is invalid in ${file}: ${error instanceof Error ? error.message : String(error)}`);
  }
}

function collectReleaseResources(archiveManifest) {
  const exports = archiveManifest.exports || {};
  return [...new Set([
    exports["./worker"],
    ...Object.entries(exports)
      .filter(([key]) => key.startsWith("./assets/") && key.endsWith(".wasm"))
      .map(([, target]) => target),
  ])]
    .filter((target) => typeof target === "string")
    .map((target) => target.replace(/^\.\//, ""));
}

function copyEvidence(source, destinationDirectory, id, commit, evidence) {
  const summary = readJson(source, `${id} summary`);
  if (summary.result !== "PASS") fail(`${id} summary did not pass`);
  if (summary.commit !== undefined && summary.commit !== commit) {
    fail(`${id} summary belongs to ${summary.commit}, expected ${commit}`);
  }
  const relativePath = path.posix.join("evidence", `${id}.json`);
  const destination = path.join(destinationDirectory, relativePath);
  copyFileSync(source, destination);
  evidence.push({
    id,
    path: relativePath,
    sha256: hashFile(destination),
  });
  return summary;
}

function assertTagMatchesVersion(version) {
  const expectedTag = process.env.RELEASE_EXPECTED_TAG ||
    (process.env.GITHUB_REF_TYPE === "tag" ? process.env.GITHUB_REF_NAME : "");
  if (expectedTag && expectedTag !== `v${version}`) {
    fail(`release tag ${expectedTag} does not match package version v${version}`);
  }
}

function materialize() {
  const commit = gitHead();
  const reproduciblePath = path.resolve(
    process.env.RELEASE_REPRODUCIBLE_SUMMARY ||
      path.join(root, "output", "acceptance", "p4-reproducible-pack", "summary.json"),
  );
  const reproducible = readJson(reproduciblePath, "P4 reproducible pack summary");
  if (reproducible.result !== "PASS") fail("P4 reproducible pack did not pass");
  if (reproducible.source?.head !== commit) {
    fail(`P4 reproducible pack belongs to ${reproducible.source?.head}, expected ${commit}`);
  }
  const sourceRun = reproducible.runs?.[0];
  if (sourceRun?.status !== "PASS" || sourceRun.packages?.length !== expectedPackages.length) {
    fail("P4 reproducible pack run 1 is incomplete");
  }
  if (
    reproducible.comparisons?.length !== expectedPackages.length ||
    reproducible.comparisons.some((entry) => !entry.sameEffectiveContent)
  ) {
    fail("P4 reproducible pack runs do not have identical effective content");
  }

  const manifestPath = path.resolve(
    process.env.RELEASE_PACK_MANIFEST ||
      path.join(path.dirname(reproduciblePath), "candidate-manifest", "summary.json"),
  );
  const manifestDirectory = path.dirname(manifestPath);
  const tgzDirectory = path.join(manifestDirectory, "tgz");
  rmSync(manifestDirectory, { recursive: true, force: true });
  mkdirSync(tgzDirectory, { recursive: true });

  const packages = new Map(sourceRun.packages.map((entry) => [entry.package, entry]));
  if (packages.size !== expectedPackages.length || expectedPackages.some((name) => !packages.has(name))) {
    fail("P4 reproducible pack run 1 does not contain the exact six public packages");
  }
  const entries = [];
  for (const packageName of expectedPackages) {
    const entry = packages.get(packageName);
    const source = path.resolve(path.dirname(reproduciblePath), entry.tgz);
    if (!existsSync(source) || hashFile(source) !== entry.tgzSha256) {
      fail(`${packageName} reproducible tgz is missing or changed`);
    }
    const archiveManifest = inspectArchive(source);
    if (archiveManifest.name !== packageName || archiveManifest.version !== entry.version) {
      fail(`${packageName} reproducible tgz identity is invalid`);
    }
    if (archiveManifest.private === true || archiveManifest.publishConfig?.access !== "public") {
      fail(`${packageName} reproducible tgz is not configured as a public package`);
    }
    const destination = path.join(tgzDirectory, entry.filename);
    copyFileSync(source, destination);
    const resources = collectReleaseResources(archiveManifest).map((resource) => {
      const file = entry.files?.find((item) => item.path === resource);
      if (!file?.sha256) fail(`${packageName} resource ${resource} is missing from the tgz`);
      return { path: resource, sha256: file.sha256 };
    });
    entries.push({
      package: packageName,
      version: entry.version,
      filename: entry.filename,
      sha256: entry.tgzSha256,
      fileCount: entry.fileCount,
      unpackedSize: entry.unpackedSize,
      resources,
      files: entry.files,
    });
  }
  const versions = new Set(entries.map((entry) => entry.version));
  if (versions.size !== 1) fail("P4 reproducible tgz versions are inconsistent");
  const summary = {
    suite: "P4-RELEASE-CANDIDATE-PACK",
    result: "PASS",
    commit,
    sourceSnapshotSha256: reproducible.source.snapshotSha256,
    sourceRun: 1,
    entries,
    failures: [],
  };
  writeFileSync(manifestPath, `${JSON.stringify(summary, null, 2)}\n`);
  console.log(`PASS: exact candidate pack materialized at ${manifestDirectory}`);
}

function prepare() {
  const commit = gitHead();
  const expectedCommit = process.env.RELEASE_EXPECTED_COMMIT || process.env.GITHUB_SHA || commit;
  if (commit !== expectedCommit) fail(`checked-out commit ${commit} does not match ${expectedCommit}`);
  const ciCandidate = process.env.RELEASE_CI_CANDIDATE === "1";

  const runId = process.env.CI_RUN_ID || process.env.GITHUB_RUN_ID || "local";
  const evidenceRoot = path.resolve(
    process.env.RELEASE_EVIDENCE_ROOT ||
      path.join(root, "output", "acceptance", commit, "BB-RELEASE", runId),
  );
  const candidateDir = path.resolve(
    process.env.RELEASE_CANDIDATE_DIR || path.join(evidenceRoot, "candidate"),
  );
  const consumerManifestPath = path.resolve(
    process.env.RELEASE_PACK_MANIFEST ||
      path.join(evidenceRoot, "candidate-source", "real-tgz-manifests", "summary.json"),
  );
  const sourceTgzDir = path.join(path.dirname(consumerManifestPath), "tgz");

  rmSync(candidateDir, { recursive: true, force: true });
  mkdirSync(path.join(candidateDir, "tgz"), { recursive: true });
  mkdirSync(path.join(candidateDir, "evidence"), { recursive: true });

  const evidence = [];
  let matrixDetail = null;
  if (!ciCandidate) {
    for (const suite of requiredSuites) {
      copyEvidence(
        path.join(evidenceRoot, suite, "summary.json"),
        candidateDir,
        `${suite}-summary`,
        commit,
        evidence,
      );
    }

    matrixDetail = copyEvidence(
      path.join(evidenceRoot, "matrix", "compatibility-matrix", "summary.json"),
      candidateDir,
      "matrix-detail",
      commit,
      evidence,
    );
    copyEvidence(
      path.join(evidenceRoot, "consumer", "external-tgz-consumer", "summary.json"),
      candidateDir,
      "consumer-detail",
      commit,
      evidence,
    );
  }

  copyEvidence(
    consumerManifestPath,
    candidateDir,
    "candidate-pack-manifest",
    commit,
    evidence,
  );
  if (!ciCandidate) {
    copyEvidence(
      process.env.RELEASE_REPRODUCIBLE_SUMMARY ||
        path.join(evidenceRoot, "release", "p4-reproducible-pack", "summary.json"),
      candidateDir,
      "p4-reproducible-pack",
      commit,
      evidence,
    );
    copyEvidence(
      process.env.RELEASE_READINESS_SUMMARY ||
        path.join(evidenceRoot, "release", "p4-release-readiness", "summary.json"),
      candidateDir,
      "p4-release-readiness",
      commit,
      evidence,
    );
  }

  const packSummary = readJson(consumerManifestPath, "consumer tgz manifest");
  if (
    packSummary.result !== "PASS" ||
    (packSummary.commit !== undefined && packSummary.commit !== commit)
  ) {
    fail("consumer tgz manifest did not pass for this commit");
  }
  if (!Array.isArray(packSummary.entries) || packSummary.entries.length !== expectedPackages.length) {
    fail(`consumer tgz manifest must contain exactly ${expectedPackages.length} packages`);
  }

  const byPackage = new Map(packSummary.entries.map((entry) => [entry.package, entry]));
  if (byPackage.size !== expectedPackages.length || expectedPackages.some((name) => !byPackage.has(name))) {
    fail("consumer tgz manifest does not contain the exact six public packages");
  }
  const versions = new Set(packSummary.entries.map((entry) => entry.version));
  if (versions.size !== 1) fail("the six public packages do not use one version");
  const version = [...versions][0];
  if (typeof version !== "string" || !version) fail("candidate version is missing");
  assertTagMatchesVersion(version);

  const archives = [];
  for (const packageName of expectedPackages) {
    const entry = byPackage.get(packageName);
    if (path.basename(entry.filename || "") !== entry.filename || !entry.filename.endsWith(".tgz")) {
      fail(`${packageName} has an unsafe archive filename`);
    }
    const source = path.join(sourceTgzDir, entry.filename);
    if (!existsSync(source) || !statSync(source).isFile()) fail(`${packageName} archive is missing`);
    const actualSha256 = hashFile(source);
    if (actualSha256 !== entry.sha256) fail(`${packageName} archive SHA-256 does not match its manifest`);
    const archiveManifest = inspectArchive(source);
    if (archiveManifest.name !== packageName || archiveManifest.version !== version) {
      fail(`${entry.filename} contains ${archiveManifest.name}@${archiveManifest.version}, expected ${packageName}@${version}`);
    }
    const relativePath = path.posix.join("tgz", entry.filename);
    const destination = path.join(candidateDir, relativePath);
    copyFileSync(source, destination);
    archives.push({
      package: packageName,
      version,
      filename: entry.filename,
      path: relativePath,
      sha256: actualSha256,
      integrity: integrity(destination),
    });
  }

  if (matrixDetail) {
    const matrixArchives = new Map(
      (matrixDetail.packageArchives || []).map((entry) => [entry.package, entry]),
    );
    for (const archive of archives) {
      const tested = matrixArchives.get(archive.package);
      if (!tested || tested.version !== archive.version || tested.sha256 !== archive.sha256) {
        fail(`${archive.package} matrix archive is not the exact candidate tgz`);
      }
    }
  }

  const verifierName = "prepare-release-artifact.mjs";
  const verifierPath = path.join(candidateDir, verifierName);
  copyFileSync(fileURLToPath(import.meta.url), verifierPath);
  const manifest = {
    schemaVersion: 1,
    result: "PASS",
    commit,
    runId,
    version,
    createdAt: new Date().toISOString(),
    archives,
    evidence,
    verifier: {
      path: verifierName,
      sha256: hashFile(verifierPath),
    },
  };
  writeFileSync(
    path.join(candidateDir, "candidate-manifest.json"),
    `${JSON.stringify(manifest, null, 2)}\n`,
  );
  verify(candidateDir, { expectedCommit: commit, expectedTag: process.env.RELEASE_EXPECTED_TAG });
  console.log(`PASS: release candidate prepared at ${candidateDir}`);
}

function walk(directory) {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const absolute = path.join(directory, entry.name);
    return entry.isDirectory() ? walk(absolute) : [absolute];
  });
}

function verify(candidateDir, options = {}) {
  const directory = path.resolve(candidateDir);
  const manifestPath = path.join(directory, "candidate-manifest.json");
  const manifest = readJson(manifestPath, "candidate manifest");
  if (manifest.schemaVersion !== 1 || manifest.result !== "PASS") fail("candidate manifest is not a passing schema v1 manifest");
  const expectedCommit = options.expectedCommit || process.env.RELEASE_EXPECTED_COMMIT;
  if (expectedCommit && manifest.commit !== expectedCommit) {
    fail(`candidate commit ${manifest.commit} does not match ${expectedCommit}`);
  }
  const expectedTag = options.expectedTag || process.env.RELEASE_EXPECTED_TAG;
  if (expectedTag && expectedTag !== `v${manifest.version}`) {
    fail(`candidate version v${manifest.version} does not match tag ${expectedTag}`);
  }
  if (!Array.isArray(manifest.archives) || manifest.archives.length !== expectedPackages.length) {
    fail("candidate must contain exactly six archives");
  }
  const names = manifest.archives.map((entry) => entry.package);
  if (new Set(names).size !== expectedPackages.length || expectedPackages.some((name) => !names.includes(name))) {
    fail("candidate archive package set is invalid");
  }
  if (manifest.archives.some((entry) => entry.version !== manifest.version)) {
    fail("candidate archive versions are inconsistent");
  }

  const declaredFiles = new Set(["candidate-manifest.json"]);
  for (const archive of manifest.archives) {
    const file = path.resolve(directory, archive.path);
    if (!file.startsWith(`${directory}${path.sep}`)) fail(`archive path escapes candidate: ${archive.path}`);
    if (!existsSync(file) || hashFile(file) !== archive.sha256 || integrity(file) !== archive.integrity) {
      fail(`${archive.package} archive hash check failed`);
    }
    const archiveManifest = inspectArchive(file);
    if (archiveManifest.name !== archive.package || archiveManifest.version !== archive.version) {
      fail(`${archive.package} archive identity check failed`);
    }
    declaredFiles.add(path.relative(directory, file));
  }
  for (const item of manifest.evidence || []) {
    const file = path.resolve(directory, item.path);
    if (!file.startsWith(`${directory}${path.sep}`) || !existsSync(file) || hashFile(file) !== item.sha256) {
      fail(`${item.id} evidence hash check failed`);
    }
    const summary = readJson(file, `${item.id} evidence`);
    if (summary.result !== "PASS") fail(`${item.id} evidence no longer reports PASS`);
    if (summary.commit !== undefined && summary.commit !== manifest.commit) {
      fail(`${item.id} evidence belongs to another commit`);
    }
    declaredFiles.add(path.relative(directory, file));
  }
  const verifier = path.resolve(directory, manifest.verifier?.path || "");
  if (!verifier.startsWith(`${directory}${path.sep}`) || !existsSync(verifier) || hashFile(verifier) !== manifest.verifier.sha256) {
    fail("candidate verifier hash check failed");
  }
  declaredFiles.add(path.relative(directory, verifier));

  const actualFiles = walk(directory).map((file) => path.relative(directory, file));
  const extras = actualFiles.filter((file) => !declaredFiles.has(file));
  const missing = [...declaredFiles].filter((file) => !actualFiles.includes(file));
  if (extras.length || missing.length) {
    fail(`candidate file set differs from manifest; extra=${extras.join(",")}; missing=${missing.join(",")}`);
  }
  return manifest;
}

function runNpm(args, { allowMissing = false } = {}) {
  const result = spawnSync("npm", args, {
    encoding: "utf8",
    stdio: allowMissing ? "pipe" : "inherit",
    maxBuffer: 8 * 1024 * 1024,
  });
  if (allowMissing) {
    const output = `${result.stdout || ""}${result.stderr || ""}`;
    if (result.status !== 0 && /(?:E404|404 Not Found)/i.test(output)) return null;
    if (result.status !== 0) fail(`npm ${args.join(" ")} failed: ${output.trim()}`);
    try {
      return JSON.parse(result.stdout.trim());
    } catch (error) {
      fail(`npm ${args.join(" ")} returned invalid JSON: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
  if (result.status !== 0) fail(`npm ${args.join(" ")} exited ${result.status ?? 1}`);
  return true;
}

function registryIntegrity(packageName, version) {
  return runNpm(["view", `${packageName}@${version}`, "dist.integrity", "--json"], {
    allowMissing: true,
  });
}

function registryDistTags(packageName) {
  return runNpm(["view", packageName, "dist-tags", "--json"], {
    allowMissing: true,
  }) || {};
}

function publish(candidateDir) {
  const manifest = verify(candidateDir);
  const stagingTag = process.env.RELEASE_STAGING_TAG;
  const targetTag = process.env.RELEASE_TARGET_TAG || "latest";
  const validTag = /^[a-z][a-z0-9._-]{0,63}$/;
  if (!stagingTag || !validTag.test(stagingTag) || stagingTag === targetTag) {
    fail("RELEASE_STAGING_TAG must be a safe, non-final npm dist-tag");
  }
  if (!validTag.test(targetTag)) fail("RELEASE_TARGET_TAG is invalid");

  const useProvenance = process.env.RELEASE_PROVENANCE !== "false";
  for (const archive of manifest.archives) {
    const existing = registryIntegrity(archive.package, archive.version);
    if (existing === null) {
      const publishArgs = [
        "publish",
        path.resolve(candidateDir, archive.path),
        "--access",
        "public",
        "--tag",
        stagingTag,
      ];
      if (useProvenance) publishArgs.push("--provenance");
      runNpm(publishArgs);
    } else if (existing !== archive.integrity) {
      fail(`${archive.package}@${archive.version} already exists with different integrity`);
    } else {
      console.log(`SKIP: ${archive.package}@${archive.version} already has the candidate integrity`);
      runNpm(["dist-tag", "add", `${archive.package}@${archive.version}`, stagingTag]);
    }
  }

  for (const archive of manifest.archives) {
    const existing = registryIntegrity(archive.package, archive.version);
    if (existing !== archive.integrity) {
      fail(`${archive.package}@${archive.version} staging verification failed`);
    }
  }
  const previousTargets = new Map(
    manifest.archives.map((archive) => [
      archive.package,
      registryDistTags(archive.package)[targetTag] || null,
    ]),
  );
  const promoted = [];
  try {
    for (const archive of manifest.archives) {
      runNpm(["dist-tag", "add", `${archive.package}@${archive.version}`, targetTag]);
      promoted.push(archive);
    }
  } catch (error) {
    const rollbackErrors = [];
    for (const archive of [...promoted].reverse()) {
      const previous = previousTargets.get(archive.package);
      try {
        if (previous) {
          runNpm(["dist-tag", "add", `${archive.package}@${previous}`, targetTag]);
        } else {
          runNpm(["dist-tag", "rm", archive.package, targetTag]);
        }
      } catch (rollbackError) {
        rollbackErrors.push(
          `${archive.package}: ${rollbackError instanceof Error ? rollbackError.message : String(rollbackError)}`,
        );
      }
    }
    const reason = error instanceof Error ? error.message : String(error);
    fail(
      `final tag promotion failed and was rolled back: ${reason}` +
        (rollbackErrors.length ? `; rollback errors: ${rollbackErrors.join(" | ")}` : ""),
    );
  }
  console.log(`PASS: ${manifest.version} promoted from ${stagingTag} to ${targetTag}`);
}

try {
  const mode = process.argv[2] || "prepare";
  if (mode === "prepare") prepare();
  else if (mode === "materialize") materialize();
  else if (mode === "verify") {
    const candidateDir = process.argv[3] || process.env.RELEASE_CANDIDATE_DIR || process.cwd();
    const manifest = verify(candidateDir);
    console.log(`PASS: verified ${manifest.version} candidate for ${manifest.commit}`);
  } else if (mode === "publish") {
    const candidateDir = process.argv[3] || process.env.RELEASE_CANDIDATE_DIR || process.cwd();
    publish(candidateDir);
  } else {
    fail(`unknown mode ${mode}; expected materialize, prepare, verify, or publish`);
  }
} catch (error) {
  console.error(error instanceof Error ? error.stack || error.message : String(error));
  process.exit(1);
}
