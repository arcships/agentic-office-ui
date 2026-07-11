#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import {
  chmodSync,
  copyFileSync,
  cpSync,
  existsSync,
  lstatSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readlinkSync,
  readdirSync,
  rmSync,
  symlinkSync,
  writeFileSync,
} from "node:fs";
import os from "node:os";
import path from "node:path";
import process from "node:process";

const root = path.resolve(import.meta.dirname, "../..");
const output = path.resolve(
  process.env.P4_REPRO_EVIDENCE_DIR ||
    path.join(root, "output", "acceptance", "p4-reproducible-pack"),
);
const packageDirectories = [
  "packages/docx-core",
  "packages/xlsx-core",
  "packages/vue-docx",
  "packages/vue-xlsx",
  "packages/vue-pdf",
  "packages/vue-ui",
];
const generatedDirectoryNames = new Set(["node_modules", "dist", "output"]);

function sha256(value) {
  return createHash("sha256").update(value).digest("hex");
}

function collectExportTargets(value, result = []) {
  if (typeof value === "string") result.push(value);
  else if (value && typeof value === "object") {
    for (const child of Object.values(value)) collectExportTargets(child, result);
  }
  return result;
}

function safeOutputDirectory(directory) {
  const resolved = path.resolve(directory);
  const rootFromOutput = path.relative(resolved, root);
  const containsWorkspace = rootFromOutput === "" ||
    (!rootFromOutput.startsWith("..") && !path.isAbsolute(rootFromOutput));
  return !containsWorkspace && resolved !== path.parse(resolved).root;
}

function gitValue(args) {
  const result = spawnSync("git", args, {
    cwd: root,
    encoding: "utf8",
    maxBuffer: 16 * 1024 * 1024,
  });
  return result.status === 0 ? result.stdout.trim() : null;
}

function currentSourceFiles() {
  const result = spawnSync(
    "git",
    ["ls-files", "--cached", "--others", "--exclude-standard", "-z"],
    {
      cwd: root,
      encoding: "buffer",
      maxBuffer: 128 * 1024 * 1024,
    },
  );
  if (result.status !== 0) {
    throw new Error(`无法列出当前源码文件：${result.stderr?.toString("utf8") || "git 执行失败"}`);
  }
  return [...new Set(
    result.stdout
      .toString("utf8")
      .split("\0")
      .filter(Boolean),
  )].sort();
}

function copyCurrentSource(files, destination) {
  mkdirSync(destination, { recursive: true });
  for (const relative of files) {
    const source = path.join(root, relative);
    if (!existsSync(source)) continue;
    const target = path.join(destination, relative);
    const stat = lstatSync(source);
    mkdirSync(path.dirname(target), { recursive: true });
    if (stat.isSymbolicLink()) {
      symlinkSync(readlinkSync(source), target);
      continue;
    }
    if (!stat.isFile()) continue;
    copyFileSync(source, target);
    chmodSync(target, stat.mode & 0o777);
  }
}

function removeGeneratedDirectories(directory, removed = [], base = directory) {
  if (!existsSync(directory)) return removed;
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const absolute = path.join(directory, entry.name);
    if (generatedDirectoryNames.has(entry.name)) {
      rmSync(absolute, { recursive: true, force: true });
      removed.push(path.relative(base, absolute));
      continue;
    }
    if (entry.isDirectory()) removeGeneratedDirectories(absolute, removed, base);
  }
  return removed;
}

function walkFiles(directory, prefix = "") {
  if (!existsSync(directory)) return [];
  const result = [];
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const relative = prefix ? `${prefix}/${entry.name}` : entry.name;
    const absolute = path.join(directory, entry.name);
    if (entry.isDirectory()) result.push(...walkFiles(absolute, relative));
    else result.push(relative);
  }
  return result.sort();
}

function describeSourceFile(base, relative) {
  const absolute = path.join(base, relative);
  const stat = lstatSync(absolute);
  if (stat.isSymbolicLink()) {
    const target = readlinkSync(absolute);
    return {
      path: relative,
      kind: "symlink",
      mode: stat.mode & 0o777,
      size: Buffer.byteLength(target),
      sha256: sha256(target),
    };
  }
  const bytes = readFileSync(absolute);
  return {
    path: relative,
    kind: "file",
    mode: stat.mode & 0o777,
    size: bytes.byteLength,
    sha256: sha256(bytes),
  };
}

function snapshotManifest(directory) {
  return walkFiles(directory).map((relative) => describeSourceFile(directory, relative));
}

function writeCommandLog(file, command, args, result, elapsedMs) {
  const rendered = [command, ...args].map((item) => JSON.stringify(item)).join(" ");
  const body = [
    `$ ${rendered}`,
    `exitCode=${result.status ?? "null"}`,
    `signal=${result.signal ?? ""}`,
    `elapsedMs=${elapsedMs}`,
    "",
    result.stdout || "",
    result.stderr || "",
  ].join("\n");
  writeFileSync(file, body.endsWith("\n") ? body : `${body}\n`);
}

function runLogged(runEvidence, label, command, args, cwd) {
  const started = Date.now();
  const result = spawnSync(command, args, {
    cwd,
    encoding: "utf8",
    env: {
      ...process.env,
      CI: process.env.CI || "1",
      HUSKY: "0",
    },
    maxBuffer: 256 * 1024 * 1024,
  });
  const elapsedMs = Date.now() - started;
  const log = path.join(runEvidence, `${label}.log`);
  writeCommandLog(log, command, args, result, elapsedMs);
  return {
    ok: !result.error && result.status === 0,
    record: {
      label,
      command: [command, ...args],
      exitCode: result.status,
      signal: result.signal,
      elapsedMs,
      log: path.relative(output, log),
      error: result.error?.message || null,
    },
    stdout: result.stdout || "",
    stderr: result.stderr || "",
  };
}

function parsePackReport(stdout) {
  const start = stdout.indexOf("[");
  const end = stdout.lastIndexOf("]");
  if (start < 0 || end < start) throw new Error("npm pack 没有返回 JSON 数组");
  const report = JSON.parse(stdout.slice(start, end + 1)).at(0);
  if (!report?.filename || !Array.isArray(report.files)) {
    throw new Error("npm pack JSON 缺少文件名或文件清单");
  }
  return report;
}

function extractArchiveFile(tgzPath, relative) {
  const result = spawnSync("tar", ["-xOf", tgzPath, `package/${relative}`], {
    encoding: "buffer",
    maxBuffer: 256 * 1024 * 1024,
  });
  if (result.status !== 0) {
    throw new Error(`无法从 ${path.basename(tgzPath)} 读取 ${relative}`);
  }
  return result.stdout;
}

function effectiveHash(files) {
  const hash = createHash("sha256");
  for (const file of files) {
    hash.update(file.path);
    hash.update("\0");
    hash.update(String(file.mode ?? ""));
    hash.update("\0");
    hash.update(file.sha256);
    hash.update("\0");
  }
  return hash.digest("hex");
}

function packOne(runNumber, sourceDirectory, runEvidence, packageDirectory) {
  const manifest = JSON.parse(
    readFileSync(path.join(sourceDirectory, packageDirectory, "package.json"), "utf8"),
  );
  const tgzDirectory = path.join(runEvidence, "tgz");
  const manifestDirectory = path.join(runEvidence, "manifests");
  mkdirSync(tgzDirectory, { recursive: true });
  mkdirSync(manifestDirectory, { recursive: true });
  const label = `npm-pack-${path.basename(packageDirectory)}`;
  const packed = runLogged(
    runEvidence,
    label,
    "npm",
    ["pack", "--ignore-scripts", "--json", "--pack-destination", tgzDirectory],
    path.join(sourceDirectory, packageDirectory),
  );
  if (!packed.ok) {
    throw Object.assign(new Error(`${manifest.name}: npm pack 失败`), {
      commandRecord: packed.record,
    });
  }
  const report = parsePackReport(packed.stdout);
  const tgzPath = path.join(tgzDirectory, report.filename);
  const files = report.files
    .map((file) => ({
      path: file.path,
      size: file.size,
      mode: file.mode ?? null,
      sha256: sha256(extractArchiveFile(tgzPath, file.path)),
    }))
    .sort((left, right) => left.path.localeCompare(right.path));
  const normalizedManifest = {
    package: manifest.name,
    version: manifest.version,
    files,
  };
  const normalizedManifestText = `${JSON.stringify(normalizedManifest, null, 2)}\n`;
  const manifestText = `${JSON.stringify({ run: runNumber, ...normalizedManifest }, null, 2)}\n`;
  const manifestFile = path.join(
    manifestDirectory,
    `${manifest.name.replace(/^@/, "").replaceAll("/", "-")}.json`,
  );
  writeFileSync(manifestFile, manifestText);
  return {
    package: manifest.name,
    version: manifest.version,
    filename: report.filename,
    fileCount: files.length,
    unpackedSize: report.unpackedSize,
    tgzSha256: sha256(readFileSync(tgzPath)),
    manifestSha256: sha256(normalizedManifestText),
    manifestFileSha256: sha256(manifestText),
    effectiveSha256: effectiveHash(files),
    files,
    command: packed.record,
    tgz: path.relative(output, tgzPath),
    manifest: path.relative(output, manifestFile),
  };
}

function executeRun(runNumber, sourceSnapshot, temporaryRoot) {
  const runEvidence = path.join(output, `run-${runNumber}`);
  const sourceDirectory = path.join(temporaryRoot, `run-${runNumber}`, "source");
  mkdirSync(runEvidence, { recursive: true });
  mkdirSync(path.dirname(sourceDirectory), { recursive: true });
  cpSync(sourceSnapshot, sourceDirectory, {
    recursive: true,
    dereference: false,
    preserveTimestamps: true,
    verbatimSymlinks: true,
  });
  const removedGeneratedDirectories = removeGeneratedDirectories(sourceDirectory);
  const run = {
    run: runNumber,
    status: "FAIL",
    sourceDirectory: "temporary-clean-snapshot",
    removedGeneratedDirectories,
    commands: [],
    packages: [],
    error: null,
  };
  try {
    const install = runLogged(
      runEvidence,
      "pnpm-install",
      "pnpm",
      ["install", "--frozen-lockfile"],
      sourceDirectory,
    );
    run.commands.push(install.record);
    if (!install.ok) throw new Error("pnpm install --frozen-lockfile 失败");

    const build = runLogged(runEvidence, "pnpm-build", "pnpm", ["build"], sourceDirectory);
    run.commands.push(build.record);
    if (!build.ok) throw new Error("pnpm build 失败");

    for (const packageDirectory of packageDirectories) {
      const packed = packOne(runNumber, sourceDirectory, runEvidence, packageDirectory);
      run.commands.push(packed.command);
      run.packages.push(packed);
    }
    run.status = "PASS";
  } catch (error) {
    if (error?.commandRecord) run.commands.push(error.commandRecord);
    run.error = error instanceof Error ? error.message : String(error);
  }
  return run;
}

function compareFileLists(firstFiles = [], secondFiles = []) {
  const first = new Map(firstFiles.map((file) => [file.path, file]));
  const second = new Map(secondFiles.map((file) => [file.path, file]));
  const paths = [...new Set([...first.keys(), ...second.keys()])].sort();
  const missingFromRunOne = [];
  const missingFromRunTwo = [];
  const changed = [];
  for (const file of paths) {
    const left = first.get(file);
    const right = second.get(file);
    if (!left) {
      missingFromRunOne.push(file);
      continue;
    }
    if (!right) {
      missingFromRunTwo.push(file);
      continue;
    }
    if (
      left.sha256 !== right.sha256 ||
      left.size !== right.size ||
      left.mode !== right.mode
    ) {
      changed.push({ path: file, runOne: left, runTwo: right });
    }
  }
  return { missingFromRunOne, missingFromRunTwo, changed };
}

function compareRuns(runOne, runTwo) {
  return packageDirectories.map((packageDirectory) => {
    const expectedName = JSON.parse(
      readFileSync(path.join(root, packageDirectory, "package.json"), "utf8"),
    ).name;
    const first = runOne.packages.find((entry) => entry.package === expectedName);
    const second = runTwo.packages.find((entry) => entry.package === expectedName);
    const fileDiff = compareFileLists(first?.files, second?.files);
    const sameEffectiveContent = Boolean(
      first &&
      second &&
      first.effectiveSha256 === second.effectiveSha256 &&
      fileDiff.missingFromRunOne.length === 0 &&
      fileDiff.missingFromRunTwo.length === 0 &&
      fileDiff.changed.length === 0,
    );
    return {
      package: expectedName,
      presentInBothRuns: Boolean(first && second),
      tgzSha256: [first?.tgzSha256 ?? null, second?.tgzSha256 ?? null],
      manifestSha256: [first?.manifestSha256 ?? null, second?.manifestSha256 ?? null],
      effectiveSha256: [first?.effectiveSha256 ?? null, second?.effectiveSha256 ?? null],
      sameTgz: Boolean(first && second && first.tgzSha256 === second.tgzSha256),
      sameManifest: Boolean(first && second && first.manifestSha256 === second.manifestSha256),
      sameEffectiveContent,
      fileDiff,
    };
  });
}

function writeCandidateManifest(run, overallResult) {
  const entries = run.packages.map((packed) => {
    const packageDirectory = packageDirectories.find((directory) => {
      const manifest = JSON.parse(readFileSync(path.join(root, directory, "package.json"), "utf8"));
      return manifest.name === packed.package;
    });
    if (!packageDirectory) throw new Error(`找不到候选包目录：${packed.package}`);
    const manifest = JSON.parse(readFileSync(path.join(root, packageDirectory, "package.json"), "utf8"));
    const resourceTargets = [...new Set([
      manifest.exports?.["./worker"],
      ...Object.entries(manifest.exports || {})
        .filter(([key]) => key.startsWith("./assets/") && key.endsWith(".wasm"))
        .map(([, target]) => target),
    ])]
      .filter((target) => typeof target === "string")
      .map((target) => target.replace(/^\.\//, ""));
    const resources = resourceTargets.map((resourcePath) => {
      const file = packed.files.find((entry) => entry.path === resourcePath);
      return { path: resourcePath, sha256: file?.sha256 || null };
    });
    return {
      package: packed.package,
      version: packed.version,
      filename: packed.filename,
      sha256: packed.tgzSha256,
      effectiveSha256: packed.effectiveSha256,
      fileCount: packed.fileCount,
      unpackedSize: packed.unpackedSize,
      resources,
      files: packed.files,
    };
  });
  const manifest = {
    suite: "P4-REPRODUCIBLE-CANDIDATE-MANIFEST",
    result: overallResult,
    sourceHead: gitValue(["rev-parse", "HEAD"]),
    sourceSnapshotSha256: summary.source.snapshotSha256,
    entries,
    failures: overallResult === "PASS" ? [] : ["two-build reproducibility gate failed"],
  };
  const manifestPath = path.join(output, `run-${run.run}`, "summary.json");
  writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`);
  return path.relative(output, manifestPath);
}

if (!safeOutputDirectory(output)) {
  console.error(`FAIL: 不安全的证据目录：${output}`);
  process.exit(1);
}

rmSync(output, { recursive: true, force: true });
mkdirSync(output, { recursive: true });
const temporaryRoot = mkdtempSync(path.join(os.tmpdir(), "agentic-office-p4-pack-"));
const summary = {
  suite: "P4-PACK-TWO-CLEAN-BUILDS",
  result: "FAIL",
  source: {
    root: "current-workspace-snapshot",
    head: gitValue(["rev-parse", "HEAD"]),
    status: gitValue(["status", "--short"]),
    fileCount: 0,
    snapshotSha256: null,
    removedGeneratedDirectories: [],
  },
  requirements: {
    install: "pnpm install --frozen-lockfile",
    build: "pnpm build",
    packages: packageDirectories,
  },
  runs: [],
  comparisons: [],
  failures: [],
};

try {
  const sourceSnapshot = path.join(temporaryRoot, "snapshot");
  copyCurrentSource(currentSourceFiles(), sourceSnapshot);
  summary.source.removedGeneratedDirectories = removeGeneratedDirectories(sourceSnapshot);
  const sourceFiles = snapshotManifest(sourceSnapshot);
  const sourceManifestText = `${JSON.stringify(sourceFiles, null, 2)}\n`;
  writeFileSync(path.join(output, "source-manifest.json"), sourceManifestText);
  summary.source.fileCount = sourceFiles.length;
  summary.source.snapshotSha256 = sha256(sourceManifestText);

  const runOne = executeRun(1, sourceSnapshot, temporaryRoot);
  const runTwo = executeRun(2, sourceSnapshot, temporaryRoot);
  summary.runs = [runOne, runTwo];
  summary.comparisons = compareRuns(runOne, runTwo);
  for (const run of summary.runs) {
    if (run.status !== "PASS") summary.failures.push(`run-${run.run}: ${run.error || "失败"}`);
  }
  for (const comparison of summary.comparisons) {
    if (!comparison.sameEffectiveContent) {
      summary.failures.push(`${comparison.package}: 两次有效内容不同`);
    }
  }
  if (
    summary.failures.length === 0 &&
    summary.runs.every((run) => run.packages.length === packageDirectories.length)
  ) {
    summary.result = "PASS";
  }
  if (summary.runs[0]) {
    summary.candidateManifest = writeCandidateManifest(summary.runs[0], summary.result);
  }
} catch (error) {
  summary.failures.push(error instanceof Error ? error.stack || error.message : String(error));
} finally {
  rmSync(temporaryRoot, { recursive: true, force: true });
  writeFileSync(path.join(output, "summary.json"), `${JSON.stringify(summary, null, 2)}\n`);
}

console.log(`${summary.result}: two isolated builds checked at ${output}`);
for (const failure of summary.failures) console.error(`- ${failure}`);
process.exit(summary.result === "PASS" ? 0 : 1);
