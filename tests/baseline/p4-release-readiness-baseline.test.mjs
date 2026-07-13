import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";

const ROOT = path.resolve(import.meta.dirname, "../..");
const PACKAGE_DIRS = [
  "packages/docx-core",
  "packages/xlsx-core",
  "packages/vue-docx",
  "packages/vue-xlsx",
  "packages/vue-pdf",
  "packages/vue-ui",
  "packages/pptx-core",
  "packages/vue-pptx",
];
const CONTRACT_PATH = path.join(ROOT, "scripts/ci/public-api-contract.json");
const API_DOCUMENT_PATH = path.join(ROOT, "docs/api/public-api-contract.md");
const PRIVATE_RUNTIME = "@arcships/office-runtime";

function gitHead() {
  const result = spawnSync("git", ["rev-parse", "HEAD"], {
    cwd: ROOT,
    encoding: "utf8",
  });
  return result.status === 0 ? result.stdout.trim() : "unknown-head";
}

const HEAD = gitHead();
const EVIDENCE = path.resolve(
  process.env.P4_BASELINE_EVIDENCE_DIR ||
    path.join(ROOT, "output/acceptance", HEAD, "P4-READINESS-BASELINE", "local"),
);

function stable(value) {
  if (Array.isArray(value)) return value.map(stable);
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.keys(value).sort().map((key) => [key, stable(value[key])]),
    );
  }
  return value;
}

function collectExportTargets(value, result = []) {
  if (typeof value === "string") result.push(value);
  else if (value && typeof value === "object") {
    for (const child of Object.values(value)) collectExportTargets(child, result);
  }
  return result;
}

function walk(directory) {
  if (!existsSync(directory)) return [];
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const absolute = path.join(directory, entry.name);
    return entry.isDirectory() ? walk(absolute) : [absolute];
  });
}

function parsePackReport(output) {
  const start = output.indexOf("[");
  const end = output.lastIndexOf("]");
  if (start < 0 || end < start) throw new Error("npm pack did not return JSON");
  return JSON.parse(output.slice(start, end + 1))[0];
}

function extractArchiveFile(tgzPath, file) {
  const result = spawnSync("tar", ["-xOf", tgzPath, `package/${file}`], {
    maxBuffer: 128 * 1024 * 1024,
  });
  if (result.status !== 0) {
    throw new Error(`cannot extract ${file} from ${path.basename(tgzPath)}`);
  }
  return result.stdout;
}

function hashBytes(value) {
  return createHash("sha256").update(value).digest("hex");
}

function effectiveArchiveHash(tgzPath, files) {
  const hash = createHash("sha256");
  for (const file of [...files].sort()) {
    hash.update(file);
    hash.update("\0");
    hash.update(hashBytes(extractArchiveFile(tgzPath, file)));
    hash.update("\0");
  }
  return hash.digest("hex");
}

function packCurrentDist(runDirectory, packageDir) {
  mkdirSync(runDirectory, { recursive: true });
  const result = spawnSync(
    "npm",
    ["pack", "--ignore-scripts", "--json", "--pack-destination", runDirectory],
    {
      cwd: path.join(ROOT, packageDir),
      encoding: "utf8",
      maxBuffer: 64 * 1024 * 1024,
    },
  );
  if (result.status !== 0) {
    throw new Error(`${packageDir}: npm pack exited ${result.status}: ${result.stderr}`);
  }
  const report = parsePackReport(result.stdout);
  const tgzPath = path.join(runDirectory, report.filename);
  const files = report.files.map((entry) => entry.path);
  return {
    report,
    tgzPath,
    sha256: hashBytes(readFileSync(tgzPath)),
    effectiveSha256: effectiveArchiveHash(tgzPath, files),
  };
}

function scanArchive(packageName, packed, expectedExports) {
  const files = packed.report.files.map((entry) => entry.path);
  const targetFiles = collectExportTargets(expectedExports)
    .filter((target) => target.startsWith("./"))
    .map((target) => target.slice(2));
  const missingTargets = targetFiles.filter((target) => !files.includes(target));
  const forbiddenFiles = files.filter((file) =>
    /^(?:src|test|tests|fixtures?|samples?|output)\//i.test(file) ||
    /\.(?:docx|xlsx|pdf)$/i.test(file),
  );
  const licenseFiles = files.filter((file) => /(?:^|\/)(?:LICENSE|LICENCE)(?:\..*)?$/i.test(file));
  const absolutePathLeaks = [];
  const privateRuntimeLeaks = [];
  const workspaceLeaks = [];
  const textual = /\.(?:js|mjs|cjs|d\.ts|json|css|map|md|txt)$/i;
  for (const file of files.filter((entry) => textual.test(entry))) {
    const text = extractArchiveFile(packed.tgzPath, file).toString("utf8");
    if (
      text.includes(ROOT) ||
      /(?:file:\/\/\/Users\/|\/Users\/[A-Za-z0-9._-]+\/|\/home\/runner\/work\/|[A-Za-z]:\\Users\\)/.test(text)
    ) {
      absolutePathLeaks.push(file);
    }
    if (text.includes(PRIVATE_RUNTIME)) privateRuntimeLeaks.push(file);
    if (text.includes("workspace:")) workspaceLeaks.push(file);
  }
  return {
    package: packageName,
    fileCount: files.length,
    unpackedSize: packed.report.unpackedSize,
    files: packed.report.files,
    missingTargets,
    forbiddenFiles,
    licenseFiles,
    absolutePathLeaks,
    privateRuntimeLeaks,
    workspaceLeaks,
  };
}

test("P4 release readiness baseline", () => {
  mkdirSync(EVIDENCE, { recursive: true });
  const temporary = mkdtempSync(path.join(os.tmpdir(), "agentic-office-p4-audit-"));
  const checks = [];
  const implementation = [];
  const add = (id, passed, details, minimumFix) => {
    checks.push({ id, result: passed ? "PASS" : "FAIL", details });
    if (!passed && minimumFix) implementation.push({ id, minimumFix });
  };

  try {
    const contract = JSON.parse(readFileSync(CONTRACT_PATH, "utf8"));
    const apiDocument = readFileSync(API_DOCUMENT_PATH, "utf8");
    const manifests = PACKAGE_DIRS.map((directory) => ({
      directory,
      manifest: JSON.parse(readFileSync(path.join(ROOT, directory, "package.json"), "utf8")),
    }));

    const exportChecks = manifests.map(({ directory, manifest }) => {
      const expected = contract.packages.find((entry) => entry.name === manifest.name);
      return {
        package: manifest.name,
        contractPresent: Boolean(expected),
        exact: Boolean(expected) &&
          JSON.stringify(stable(manifest.exports)) === JSON.stringify(stable(expected.exports)),
        currentTargetsPresent: collectExportTargets(manifest.exports).every((target) =>
          !target.startsWith("./") || existsSync(path.join(ROOT, directory, target.slice(2))),
        ),
      };
    });
    add(
      "P4-API-EXPORTS",
      exportChecks.every((entry) => entry.contractPresent && entry.exact && entry.currentTargetsPresent),
      exportChecks,
      "在最终合并后刷新八包导出快照，并确保合同、package.json 与正式 dist 完全一致。",
    );

    const versions = Object.fromEntries(
      manifests.map(({ manifest }) => [manifest.name, manifest.version]),
    );
    add(
      "P4-API-VERSION-CONSISTENCY",
      new Set(Object.values(versions)).size === 1,
      versions,
      "八个公开包必须使用同一候选版本。",
    );
    add(
      "P4-API-VERSION-POLICY",
      apiDocument.includes("0.2.0") &&
        apiDocument.includes("1.0.0") &&
        apiDocument.includes("removeInVersion") &&
        apiDocument.includes("removeByDate"),
      { document: path.relative(ROOT, API_DOCUMENT_PATH) },
      "补齐公开弃用和内部桥接的版本/日期登记。",
    );
    add(
      "P4-API-RELEASE-VERSION",
      Object.values(versions).every((version) => version === "0.5.4"),
      { planned: "0.5.4", actual: versions },
      "在发布流程冻结候选制品前一次性把八包版本升级为已选定的 0.5.4；不得提前零散改版本。",
    );

    const privateManifest = JSON.parse(
      readFileSync(path.join(ROOT, "packages/office-runtime/package.json"), "utf8"),
    );
    add(
      "P4-PRIVATE-RUNTIME",
      privateManifest.private === true &&
        manifests.every(({ manifest }) => !JSON.stringify(manifest).includes(PRIVATE_RUNTIME)),
      {
        private: privateManifest.private,
        publicManifestReferences: manifests
          .filter(({ manifest }) => JSON.stringify(manifest).includes(PRIVATE_RUNTIME))
          .map(({ manifest }) => manifest.name),
      },
      "保持 office-runtime 为私有包，并将所需实现打入公开包，不能加入公开依赖。",
    );

    const runOne = [];
    const runTwo = [];
    for (const { directory, manifest } of manifests) {
      runOne.push({
        package: manifest.name,
        ...packCurrentDist(path.join(temporary, "run-1"), directory),
      });
      runTwo.push({
        package: manifest.name,
        ...packCurrentDist(path.join(temporary, "run-2"), directory),
      });
    }
    const reproducibility = runOne.map((first) => {
      const second = runTwo.find((entry) => entry.package === first.package);
      return {
        package: first.package,
        tgzSha256: [first.sha256, second?.sha256],
        effectiveSha256: [first.effectiveSha256, second?.effectiveSha256],
        sameTgz: first.sha256 === second?.sha256,
        sameEffectiveContent: first.effectiveSha256 === second?.effectiveSha256,
      };
    });
    add(
      "P4-PACK-SAME-DIST-REPEAT",
      reproducibility.every((entry) => entry.sameEffectiveContent),
      reproducibility,
      "先保证相同 dist 重复打包的有效内容稳定。",
    );
    const suppliedReproducibleSummary = process.env.P4_REPRO_SUMMARY_PATH
      ? path.resolve(process.env.P4_REPRO_SUMMARY_PATH)
      : null;
    const reproduciblePackEvidence = suppliedReproducibleSummary
      ? path.dirname(suppliedReproducibleSummary)
      : path.join(EVIDENCE, "two-clean-builds");
    const reproduciblePack = suppliedReproducibleSummary
      ? { status: 0, signal: null, error: null, stdout: "reused tested summary", stderr: "" }
      : spawnSync(
          process.execPath,
          ["scripts/ci/p4-reproducible-pack.mjs"],
          {
            cwd: ROOT,
            encoding: "utf8",
            env: {
              ...process.env,
              P4_REPRO_EVIDENCE_DIR: reproduciblePackEvidence,
            },
            maxBuffer: 16 * 1024 * 1024,
          },
        );
    const reproduciblePackSummaryPath = suppliedReproducibleSummary ||
      path.join(reproduciblePackEvidence, "summary.json");
    let reproduciblePackSummary = null;
    let reproduciblePackSummaryError = null;
    try {
      reproduciblePackSummary = JSON.parse(readFileSync(reproduciblePackSummaryPath, "utf8"));
    } catch (error) {
      reproduciblePackSummaryError = error instanceof Error ? error.message : String(error);
    }
    const reproducibleComparisons = reproduciblePackSummary?.comparisons || [];
    add(
      "P4-PACK-TWO-CLEAN-BUILDS",
      reproduciblePack.status === 0 &&
        reproduciblePackSummary?.result === "PASS" &&
        reproduciblePackSummary?.runs?.length === 2 &&
        reproduciblePackSummary.runs.every((run) => run.status === "PASS") &&
        reproducibleComparisons.length === PACKAGE_DIRS.length &&
        reproducibleComparisons.every((entry) => entry.sameEffectiveContent),
      {
        exitCode: reproduciblePack.status,
        signal: reproduciblePack.signal,
        error: reproduciblePack.error?.message || null,
        stdout: reproduciblePack.stdout.trim(),
        stderr: reproduciblePack.stderr.trim(),
        evidence: path.relative(ROOT, reproduciblePackEvidence),
        reusedExistingSummary: Boolean(suppliedReproducibleSummary),
        summaryError: reproduciblePackSummaryError,
        summaryResult: reproduciblePackSummary?.result || null,
        source: reproduciblePackSummary?.source || null,
        failures: reproduciblePackSummary?.failures || [],
        comparisons: reproducibleComparisons.map((entry) => ({
          package: entry.package,
          sameTgz: entry.sameTgz,
          sameManifest: entry.sameManifest,
          sameEffectiveContent: entry.sameEffectiveContent,
          fileDiff: entry.fileDiff,
        })),
      },
      "隔离复制当前源码两次，冻结锁文件安装并正式构建六包；逐文件比较有效内容并保存差异。",
    );

    const archiveScans = runOne.map((packed) => {
      const expected = contract.packages.find((entry) => entry.name === packed.package);
      return scanArchive(packed.package, packed, expected?.exports || {});
    });
    add(
      "P4-PACK-CONTENT-SAFETY",
      archiveScans.every((entry) =>
        entry.missingTargets.length === 0 &&
        entry.forbiddenFiles.length === 0 &&
        entry.absolutePathLeaks.length === 0 &&
        entry.privateRuntimeLeaks.length === 0 &&
        entry.workspaceLeaks.length === 0
      ),
      archiveScans,
      "将入口、绝对路径、源码/测试材料、workspace 和私有运行包扫描加入候选 tgz 硬门禁。",
    );
    add(
      "P4-PACK-LICENSE",
      manifests.every(({ manifest }) => typeof manifest.license === "string") &&
        archiveScans.every((entry) => entry.licenseFiles.length > 0),
      manifests.map(({ manifest }) => ({
        package: manifest.name,
        licenseField: manifest.license || null,
        packagedLicenseFiles:
          archiveScans.find((entry) => entry.package === manifest.name)?.licenseFiles || [],
      })),
      "确认授权来源后，为六包补 license 字段和实际 LICENSE 文件，并验证进入 tgz。",
    );

    writeFileSync(
      path.join(EVIDENCE, "current-dist-pack-report.json"),
      `${JSON.stringify({ reproducibility, archiveScans }, null, 2)}\n`,
    );

    const publicApiEvidence = path.join(EVIDENCE, "public-api-contract");
    const publicApi = spawnSync(process.execPath, ["scripts/ci/verify-public-api.mjs"], {
      cwd: ROOT,
      encoding: "utf8",
      env: { ...process.env, PUBLIC_API_EVIDENCE_DIR: publicApiEvidence },
      maxBuffer: 64 * 1024 * 1024,
    });
    add(
      "P4-API-BUILT-CONTRACT",
      publicApi.status === 0,
      {
        exitCode: publicApi.status,
        stdout: publicApi.stdout.trim(),
        stderr: publicApi.stderr.trim(),
        evidence: path.relative(ROOT, publicApiEvidence),
      },
      "全部 P3 接口冻结后重建六包并刷新经人工确认的公开声明指纹；不能用自动接受变化代替复核。",
    );

    const workflowFiles = walk(path.join(ROOT, ".github/workflows"));
    const workflowText = workflowFiles.map((file) => readFileSync(file, "utf8")).join("\n");
    const rootManifest = JSON.parse(readFileSync(path.join(ROOT, "package.json"), "utf8"));
    add(
      "P4-RELEASE-PIPELINE",
      /npm\s+publish/.test(workflowText) &&
        /download-artifact/.test(workflowText) &&
        /github\.sha/.test(workflowText),
      {
        workflows: workflowFiles.map((file) => path.relative(ROOT, file)),
        hasPublish: /npm\s+publish/.test(workflowText),
        uploadsEvidence: /upload-artifact/.test(workflowText),
        downloadsTestedArtifact: /download-artifact/.test(workflowText),
        bindsCommit: /github\.sha/.test(workflowText),
      },
      "新增只发布同一提交已测试 tgz 的受保护流程：先产出并留存制品，后续发布只下载该制品，不重新构建。",
    );
    add(
      "P4-RELEASE-SUITE",
      rootManifest.scripts?.["test:release"] &&
        !readFileSync(path.join(ROOT, "scripts/ci/run-suite.mjs"), "utf8")
          .includes('release: [\n    command("p1-check"'),
      { command: rootManifest.scripts?.["test:release"] || null },
      "把 test:release 扩为真正 BB-RELEASE，而不是 P1 check 的别名。",
    );

    const compatibilityFiles = [
      "docs/testing/compatibility-matrix.md",
      "docs/compatibility-matrix.md",
    ].filter((file) => existsSync(path.join(ROOT, file)));
    const compatibilityScriptText = readFileSync(
      path.join(ROOT, "scripts/ci/compatibility-matrix.mjs"),
      "utf8",
    );
    const standaloneMatrixAvailable =
      rootManifest.scripts?.["test:matrix"] === "node scripts/ci/run-suite.mjs matrix";
    add(
      "P4-COMPATIBILITY-MATRIX",
      compatibilityFiles.length > 0 &&
        standaloneMatrixAvailable &&
        /chromium/i.test(compatibilityScriptText) &&
        /firefox/i.test(compatibilityScriptText) &&
        /webkit/i.test(compatibilityScriptText),
      {
        matrixDocuments: compatibilityFiles,
        standaloneMatrixAvailable,
        supportedBrowsers: {
          chromium: /chromium/i.test(compatibilityScriptText),
          firefox: /firefox/i.test(compatibilityScriptText),
          webkit: /webkit/i.test(compatibilityScriptText),
        },
      },
      "登记 Vue/Vite/TypeScript/Node/浏览器支持范围，并对每个声明组合运行安装、类型、构建和核心黑盒。",
    );

    const releaseDocuments = [
      "CHANGELOG.md",
      "RELEASE_NOTES.md",
      "docs/release-notes.md",
      "docs/migration-0.2.md",
    ].filter((file) => existsSync(path.join(ROOT, file)));
    add(
      "P4-RELEASE-NOTES",
      releaseDocuments.some((file) => /release|changelog/i.test(file)) &&
        releaseDocuments.some((file) => /migration/i.test(file)),
      { documents: releaseDocuments },
      "新增 0.2.0 发布说明和迁移文档，列出新入口、弃用项、限制、错误码、资源配置和回退办法。",
    );
  } catch (error) {
    add(
      "P4-AUDIT-EXECUTION",
      false,
      { error: error instanceof Error ? error.stack || error.message : String(error) },
      "修复审计脚本自身后重新运行，不能把未执行当成产品通过。",
    );
  } finally {
    rmSync(temporary, { recursive: true, force: true });
  }

  const failures = checks.filter((entry) => entry.result === "FAIL");
  const summary = {
    suite: "P4-RELEASE-READINESS-BASELINE",
    commit: HEAD,
    result: failures.length === 0 ? "PASS" : "FAIL",
    scope: "static manifests/current dist/two no-build packs; not a P4 clean-build or release verdict",
    checks,
    failures: failures.map((entry) => entry.id),
    minimumImplementation: implementation,
  };
  writeFileSync(path.join(EVIDENCE, "summary.json"), `${JSON.stringify(summary, null, 2)}\n`);
  assert.equal(
    failures.length,
    0,
    `P4 release readiness has ${failures.length} red gates: ${failures.map((entry) => entry.id).join(", ")}`,
  );
});
