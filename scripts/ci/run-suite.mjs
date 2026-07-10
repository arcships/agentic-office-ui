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
const evidenceRoot = path.resolve(
  process.env.CI_EVIDENCE_ROOT ||
    path.join(root, "output", "acceptance", commit, "P1-CI-01", runId),
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
      "tests/unit/docx-core.test.mjs",
      "tests/unit/xlsx-core.test.mjs",
      "tests/unit/pdf-source.test.mjs",
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
      "tests/component/xlsx-race.test.mjs",
      "tests/component/docx-render-parity.test.mjs",
      "tests/component/docx-edit-interactions.test.mjs",
    ]),
    command("docx-xlsx-component-smoke", process.execPath, [
      "scripts/ci/component-smoke.mjs",
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
    childSuite("consumer", "test:consumer"),
    childSuite("docs", "test:docs"),
    command("failure-propagation", pnpm, ["ci:self-test"], {
      env: { CI_PREBUILT: "1" },
    }),
    command("diff-check", "git", ["diff", "--check"]),
  ],
  release: [
    command("p1-check", pnpm, ["check"]),
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
  const stepStatus = status === 0 ? "PASS" : "FAIL";
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
  if (status !== 0) failed = true;
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
