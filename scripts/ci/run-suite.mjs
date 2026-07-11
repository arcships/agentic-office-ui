#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import { mkdirSync, writeFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";

const root = path.resolve(import.meta.dirname, "../..");
const suiteName = process.argv[2];
const pnpm = process.platform === "win32" ? "pnpm.cmd" : "pnpm";

function timestamp() {
  return new Date().toISOString().replace(/[-:]/g, "").replace(/\.\d{3}Z$/, "Z");
}

function readGitHead() {
  const result = spawnSync("git", ["rev-parse", "HEAD"], {
    cwd: root,
    encoding: "utf8",
  });
  return result.status === 0 ? result.stdout.trim() : "no-git-head";
}

const commit = readGitHead();
const runId = process.env.CI_RUN_ID || timestamp();
const evidenceSuite = {
  matrix: "P4-MATRIX-01",
  performance: "P3-PERF-BASELINE-01",
  release: "BB-RELEASE",
  stress: "BB-STRESS",
}[suiteName] || "P1-CI-01";
const evidenceRoot = path.resolve(
  process.env.CI_EVIDENCE_ROOT ||
    path.join(root, "output", "acceptance", commit, evidenceSuite, runId),
);
const suiteDir = path.join(evidenceRoot, suiteName || "unknown-suite");
const prebuilt = process.env.CI_PREBUILT === "1";

function command(id, executable, args, options = {}) {
  return { id, executable, args, ...options };
}

function childSuite(id, script) {
  return command(id, pnpm, [script], {
    env: { CI_PREBUILT: "1" },
  });
}

function buildStep() {
  return prebuilt ? [] : [command("build", pnpm, ["build"])];
}

const suites = {
  unit: [
    ...buildStep(),
    command("core-behavior", process.execPath, [
      "--test",
      "--test-concurrency=1",
      "tests/unit/office-runtime.test.mjs",
      "tests/unit/image-budget.test.mjs",
      "tests/unit/docx-core.test.mjs",
      "tests/unit/docx-history-cache.test.mjs",
      "tests/unit/xlsx-core.test.mjs",
      "packages/vue-xlsx/test/lazy-loading.mjs",
      "tests/unit/pdf-source.test.mjs",
      "tests/unit/pptx-core.test.mjs",
      "tests/unit/core-purity.test.mjs",
    ]),
    command("core-boundary", process.execPath, [
      "scripts/ci/verify-core-boundaries.mjs",
    ]),
    command("dead-placeholder-boundary", process.execPath, [
      "scripts/ci/verify-dead-placeholders.mjs",
    ]),
    command("public-api-contract", process.execPath, [
      "scripts/ci/verify-public-api.mjs",
    ], {
      env: {
        PUBLIC_API_EVIDENCE_DIR: path.join(suiteDir, "public-api-contract"),
      },
    }),
    command("runtime-boundary", process.execPath, [
      "scripts/ci/verify-runtime-boundary.mjs",
    ], {
      env: {
        RUNTIME_BOUNDARY_EVIDENCE_DIR: path.join(suiteDir, "runtime-boundary"),
      },
    }),
    command("docx-integration", process.execPath, [
      "packages/vue-docx/tests/verify-integration.mjs",
    ]),
    command("xlsx-structure", process.execPath, [
      "packages/vue-xlsx/test/structure.mjs",
    ]),
    command("fixture-manifest", process.execPath, [
      "scripts/ci/verify-fixtures.mjs",
    ], {
      env: {
        FIXTURE_EVIDENCE_DIR: path.join(suiteDir, "fixture-manifest"),
      },
    }),
  ],
  component: [
    ...buildStep(),
    command("vue-component-behavior", process.execPath, [
      "--test",
      "--test-concurrency=1",
      "tests/component/vue-component-behavior.test.mjs",
      "tests/component/vue-components.test.mjs",
      "tests/component/pptx-preview.test.mjs",
      "tests/component/xlsx-race.test.mjs",
      "tests/component/docx-render-parity.test.mjs",
      "tests/component/docx-edit-interactions.test.mjs",
      "tests/component/docx-selection-input.test.mjs",
      "tests/component/docx-thumbnail-panel.test.mjs",
      "tests/component/docx-product-interactions.test.mjs",
      "tests/component/docx-page-budget.test.mjs",
      "tests/component/docx-history-budget.test.mjs",
      "tests/component/xlsx-history-budget.test.mjs",
      "tests/component/xlsx-grid-scroll.test.mjs",
      "tests/component/image-decode-state.test.mjs",
    ]),
    command("docx-xlsx-component-smoke", process.execPath, [
      "scripts/ci/component-smoke.mjs",
    ]),
    command("p3-regression-baselines", process.execPath, [
      "--test",
      "--test-concurrency=1",
      "tests/baseline/p3-cache-baseline.test.mjs",
      "tests/baseline/p3-image-decode-baseline.test.mjs",
      "tests/baseline/p3-worker-lifecycle-baseline.test.mjs",
      "tests/baseline/p3-xlsx-performance-baseline.test.mjs",
    ]),
  ],
  blackbox: [
    ...buildStep(),
    command("formal-preview-routes", process.execPath, [
      "scripts/ci/run-python.mjs",
      "tests/blackbox/routes_smoke.py",
    ], {
      env: {
        BLACKBOX_EVIDENCE_DIR: path.join(suiteDir, "formal-preview-routes"),
      },
    }),
    command("formal-preview-workflows", process.execPath, [
      "scripts/ci/run-python.mjs",
      "tests/blackbox/e2e_workflows.py",
    ], {
      env: {
        BLACKBOX_EVIDENCE_DIR: path.join(suiteDir, "formal-preview-workflows"),
      },
    }),
    command("formal-pdf-workflows", process.execPath, [
      "scripts/ci/run-python.mjs",
      "tests/blackbox/pdf_workflows.py",
    ], {
      env: {
        BLACKBOX_EVIDENCE_DIR: path.join(suiteDir, "formal-pdf-workflows"),
      },
    }),
    command("formal-docx-render-parity", process.execPath, [
      "scripts/ci/run-python.mjs",
      "tests/blackbox/docx_parity.py",
    ], {
      env: {
        BLACKBOX_EVIDENCE_DIR: path.join(suiteDir, "formal-docx-render-parity"),
      },
    }),
    command("formal-docx-editing", process.execPath, [
      "scripts/ci/run-python.mjs",
      "tests/blackbox/docx_editing.py",
    ], {
      env: {
        BLACKBOX_EVIDENCE_DIR: path.join(suiteDir, "formal-docx-editing"),
      },
    }),
    command("formal-ux-parity", process.execPath, [
      "scripts/ci/run-python.mjs",
      "tests/blackbox/ux_parity.py",
    ], {
      env: {
        BLACKBOX_EVIDENCE_DIR: path.join(suiteDir, "formal-ux-parity"),
      },
      allowBlocked: true,
    }),
    command("formal-fidelity-content", process.execPath, [
      "scripts/ci/run-python.mjs",
      "tests/blackbox/fidelity_content.py",
    ], {
      env: {
        BLACKBOX_EVIDENCE_DIR: path.join(suiteDir, "formal-fidelity-content"),
      },
    }),
    command("formal-security-workflows", process.execPath, [
      "scripts/ci/run-python.mjs",
      "tests/blackbox/security_workflows.py",
    ], {
      env: {
        BLACKBOX_EVIDENCE_DIR: path.join(suiteDir, "formal-security-workflows"),
      },
    }),
    command("formal-config-isolation", process.execPath, [
      "scripts/ci/run-python.mjs",
      "tests/blackbox/config_isolation.py",
    ], {
      env: {
        BLACKBOX_EVIDENCE_DIR: path.join(suiteDir, "formal-config-isolation"),
      },
    }),
    command("formal-race-regression", process.execPath, [
      "scripts/ci/run-python.mjs",
      "tests/blackbox/race_workflows.py",
    ], {
      env: {
        BLACKBOX_EVIDENCE_DIR: path.join(suiteDir, "formal-race-regression"),
      },
      allowBlocked: true,
    }),
  ],
  stress: [
    ...buildStep(),
    command("formal-stress-workflows", process.execPath, [
      "scripts/ci/run-python.mjs",
      "tests/blackbox/stress_workflows.py",
    ], {
      env: {
        BLACKBOX_EVIDENCE_DIR: path.join(suiteDir, "formal-stress-workflows"),
      },
    }),
    command("formal-worker-workflows", process.execPath, [
      "scripts/ci/run-python.mjs",
      "tests/blackbox/worker_workflows.py",
    ], {
      env: {
        BLACKBOX_EVIDENCE_DIR: path.join(suiteDir, "formal-worker-workflows"),
      },
    }),
  ],
  performance: [
    ...buildStep(),
    command("formal-scroll-range-regression", process.execPath, [
      "scripts/ci/run-python.mjs",
      "tests/baseline/p3-xlsx-scroll-baseline.py",
    ], {
      env: {
        BLACKBOX_EVIDENCE_DIR: path.join(suiteDir, "formal-scroll-range-regression"),
      },
    }),
    command("formal-performance-baseline", process.execPath, [
      "scripts/ci/run-python.mjs",
      "tests/blackbox/performance_baseline.py",
    ], {
      env: {
        BLACKBOX_EVIDENCE_DIR: path.join(suiteDir, "formal-performance-baseline"),
      },
    }),
  ],
  matrix: [
    ...buildStep(),
    command("compatibility-matrix", process.execPath, [
      "scripts/ci/compatibility-matrix.mjs",
    ], {
      env: {
        CI_PREBUILT: "1",
        COMPATIBILITY_EVIDENCE_DIR: path.join(suiteDir, "compatibility-matrix"),
      },
    }),
  ],
  consumer: [
    ...buildStep(),
    command("real-tgz-manifests", process.execPath, [
      "scripts/ci/pack-manifests.mjs",
    ], {
      env: {
        PACK_EVIDENCE_DIR: path.join(suiteDir, "real-tgz-manifests"),
      },
    }),
    command("external-tgz-consumer", process.execPath, [
      "scripts/ci/pack-consumer.mjs",
    ], {
      env: {
        PACK_MANIFEST_PATH: path.join(suiteDir, "real-tgz-manifests", "summary.json"),
        PACK_CONSUMER_EVIDENCE_DIR: path.join(suiteDir, "external-tgz-consumer"),
      },
    }),
  ],
  docs: [
    command("documentation-contract", process.execPath, [
      "scripts/ci/verify-docs.mjs",
    ], {
      env: {
        DOCS_EVIDENCE_DIR: path.join(suiteDir, "documentation-contract"),
      },
    }),
  ],
  test: [
    ...(prebuilt ? [] : [command("build", pnpm, ["build"])]),
    childSuite("unit", "test:unit"),
    childSuite("component", "test:component"),
  ],
  check: [
    command("typecheck", pnpm, ["typecheck"]),
    command("build", pnpm, ["-r", "build"]),
    childSuite("unit", "test:unit"),
    childSuite("component", "test:component"),
    childSuite("blackbox", "test:blackbox"),
    childSuite("stress", "test:stress"),
    childSuite("consumer", "test:consumer"),
    childSuite("docs", "test:docs"),
    command("failure-propagation", pnpm, ["ci:self-test"], {
      env: { CI_PREBUILT: "1" },
    }),
    command("diff-check", "git", ["diff", "--check"]),
  ],
  release: [
    command("typecheck", pnpm, ["typecheck"]),
    command("build", pnpm, ["-r", "build"]),
    childSuite("unit", "test:unit"),
    childSuite("component", "test:component"),
    childSuite("blackbox", "test:blackbox"),
    childSuite("stress", "test:stress"),
    childSuite("performance", "test:performance"),
    command("p4-reproducible-pack", process.execPath, [
      "scripts/ci/p4-reproducible-pack.mjs",
    ], {
      env: {
        P4_REPRO_EVIDENCE_DIR: path.join(suiteDir, "p4-reproducible-pack"),
      },
    }),
    command("materialize-candidate-pack", process.execPath, [
      "scripts/ci/prepare-release-artifact.mjs",
      "materialize",
    ], {
      env: {
        RELEASE_REPRODUCIBLE_SUMMARY: path.join(
          suiteDir,
          "p4-reproducible-pack",
          "summary.json",
        ),
        RELEASE_PACK_MANIFEST: path.join(
          evidenceRoot,
          "candidate-source",
          "real-tgz-manifests",
          "summary.json",
        ),
      },
    }),
    command("consumer", process.execPath, [
      "scripts/ci/pack-consumer.mjs",
    ], {
      env: {
        PACK_MANIFEST_PATH: path.join(
          evidenceRoot,
          "candidate-source",
          "real-tgz-manifests",
          "summary.json",
        ),
        PACK_CONSUMER_EVIDENCE_DIR: path.join(
          evidenceRoot,
          "consumer",
          "external-tgz-consumer",
        ),
      },
    }),
    command("matrix", process.execPath, [
      "scripts/ci/compatibility-matrix.mjs",
    ], {
      env: {
        CI_PREBUILT: "1",
        COMPATIBILITY_MANIFEST_PATH: path.join(
          evidenceRoot,
          "candidate-source",
          "real-tgz-manifests",
          "summary.json",
        ),
        COMPATIBILITY_EVIDENCE_DIR: path.join(
          evidenceRoot,
          "matrix",
          "compatibility-matrix",
        ),
      },
    }),
    childSuite("docs", "test:docs"),
    command("p4-release-readiness", process.execPath, [
      "--test",
      "--test-concurrency=1",
      "tests/baseline/p4-release-readiness-baseline.test.mjs",
    ], {
      env: {
        P4_BASELINE_EVIDENCE_DIR: path.join(suiteDir, "p4-release-readiness"),
        P4_REPRO_SUMMARY_PATH: path.join(
          suiteDir,
          "p4-reproducible-pack",
          "summary.json",
        ),
      },
    }),
    command("prepare-release-artifact", process.execPath, [
      "scripts/ci/prepare-release-artifact.mjs",
    ], {
      env: {
        RELEASE_EVIDENCE_ROOT: evidenceRoot,
        RELEASE_CANDIDATE_DIR: path.join(evidenceRoot, "candidate"),
        RELEASE_PACK_MANIFEST: path.join(
          evidenceRoot,
          "candidate-source",
          "real-tgz-manifests",
          "summary.json",
        ),
        RELEASE_REPRODUCIBLE_SUMMARY: path.join(
          suiteDir,
          "p4-reproducible-pack",
          "summary.json",
        ),
        RELEASE_READINESS_SUMMARY: path.join(
          suiteDir,
          "p4-release-readiness",
          "summary.json",
        ),
      },
    }),
    command("diff-check", "git", ["diff", "--check"]),
  ],
};

if (!suiteName || !Object.hasOwn(suites, suiteName)) {
  console.error(
    `Unknown CI suite ${JSON.stringify(suiteName)}. Expected one of: ${Object.keys(suites).join(", ")}`,
  );
  process.exit(2);
}

const steps = suites[suiteName];
if (steps.length === 0) {
  console.error(`CI suite ${suiteName} has no registered steps.`);
  process.exit(2);
}

mkdirSync(suiteDir, { recursive: true });
writeFileSync(
  path.join(evidenceRoot, "environment.json"),
  `${JSON.stringify({
    commit,
    cwd: root,
    node: process.version,
    platform: process.platform,
    arch: process.arch,
    runId,
  }, null, 2)}\n`,
);

const baseEnv = {
  ...process.env,
  CI_RUN_ID: runId,
  CI_EVIDENCE_ROOT: evidenceRoot,
};
const failureProbe = process.env.CI_FAILURE_PROBE || "";
const results = [];
let failed = false;

for (const step of steps) {
  if (failed) {
    results.push({ id: step.id, status: "NOT_RUN", exitCode: null });
    continue;
  }

  const startedAt = new Date().toISOString();
  const probeMatches =
    failureProbe === step.id || failureProbe === `${suiteName}:${step.id}`;
  console.log(`\n[${suiteName}] ${step.id}`);

  let status;
  let signal = null;
  let output = "";
  if (probeMatches) {
    status = 97;
    output = `Intentional CI failure probe at ${suiteName}:${step.id}\n`;
    console.error(output.trim());
  } else {
    const result = spawnSync(step.executable, step.args, {
      cwd: root,
      encoding: "utf8",
      env: { ...baseEnv, ...(step.env || {}) },
      maxBuffer: 64 * 1024 * 1024,
    });
    status = result.status ?? 1;
    signal = result.signal || null;
    output = `${result.stdout || ""}${result.stderr || ""}`;
    if (output) process.stdout.write(output);
    if (result.error) {
      output += `\n${result.error.stack || result.error.message}\n`;
      console.error(result.error);
    }
  }

  const finishedAt = new Date().toISOString();
  const blockedOk = step.allowBlocked && status === 2;
  const stepStatus = blockedOk ? "BLOCKED" : status === 0 ? "PASS" : "FAIL";
  writeFileSync(path.join(suiteDir, `${step.id}.log`), output);
  results.push({
    id: step.id,
    command: [step.executable, ...step.args],
    status: stepStatus,
    exitCode: status,
    signal,
    startedAt,
    finishedAt,
    failureProbe: probeMatches,
  });
  if (!blockedOk && status !== 0) failed = true;
}

const summary = {
  suite: suiteName,
  result: failed ? "FAIL" : "PASS",
  runId,
  commit,
  evidenceRoot,
  results,
};
writeFileSync(
  path.join(suiteDir, "summary.json"),
  `${JSON.stringify(summary, null, 2)}\n`,
);
console.log(`\n[${suiteName}] ${summary.result}: ${suiteDir}`);
process.exit(failed ? 1 : 0);
