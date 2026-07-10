#!/usr/bin/env node

import { mkdirSync, readFileSync, writeFileSync, existsSync } from "node:fs";
import path from "node:path";
import process from "node:process";

const root = path.resolve(import.meta.dirname, "../..");
const output = path.resolve(
  process.env.DOCS_EVIDENCE_DIR || path.join(root, "output", "acceptance", "docs-contract"),
);
const failures = [];
const checks = [];

function read(relative) {
  return readFileSync(path.join(root, relative), "utf8");
}

function check(name, condition, details = undefined) {
  checks.push({ name, status: condition ? "PASS" : "FAIL", details });
  if (!condition) failures.push(details ? `${name}: ${details}` : name);
}

const packageJson = JSON.parse(read("package.json"));
const expectedScripts = {
  "test:unit": "node scripts/ci/run-suite.mjs unit",
  "test:component": "node scripts/ci/run-suite.mjs component",
  "test:blackbox": "node scripts/ci/run-suite.mjs blackbox",
  "test:e2e": "node scripts/ci/run-suite.mjs blackbox",
  "test:consumer": "node scripts/ci/run-suite.mjs consumer",
  "test:pack": "node scripts/ci/run-suite.mjs consumer",
  "test:docs": "node scripts/ci/run-suite.mjs docs",
  check: "node scripts/ci/run-suite.mjs check",
};
for (const [name, value] of Object.entries(expectedScripts)) {
  check(`root script ${name}`, packageJson.scripts?.[name] === value, packageJson.scripts?.[name]);
}

const requiredFiles = [
  "requirements-ci.txt",
  ".github/workflows/ci.yml",
  "test-data/manifest.json",
  "tests/unit/docx-core.test.mjs",
  "tests/unit/xlsx-core.test.mjs",
  "tests/component/vue-components.test.mjs",
  "tests/blackbox/browser_evidence.py",
  "tests/blackbox/console_allowlist.json",
  "tests/blackbox/routes_smoke.py",
  "tests/blackbox/e2e_workflows.py",
  "tests/blackbox/fault_server.py",
  "tests/blackbox/race_workflows.py",
  "tests/blackbox/pack_consumer.py",
  "tests/consumer/template/src/App.vue",
  "scripts/ci/pack-manifests.mjs",
  "scripts/ci/pack-consumer.mjs",
  "scripts/ci/verify-fixtures.mjs",
  "scripts/ci/verify-docs.mjs",
];
if (process.env.DOCS_PROBE_MISSING === "1") {
  requiredFiles.push("tests/blackbox/__intentional_missing_documented_script__.py");
}
for (const file of requiredFiles) check(`documented file ${file}`, existsSync(path.join(root, file)));

const plan = read("docs/end-to-end-blackbox-test-plan.md");
const runbook = read("docs/testing/agent-execution-runbook.md");
const roadmap = read("docs/plan/stabilization-roadmap.md");
const combined = `${plan}\n${runbook}`;
for (const snippet of [
  "pnpm test:unit",
  "pnpm test:component",
  "pnpm test:blackbox",
  "pnpm test:consumer",
  "pnpm test:docs",
  "test-data/manifest.json",
  "tests/blackbox/race_workflows.py",
  "tests/consumer/template/",
  "output/acceptance/<commit>/<suite-id>/<YYYYMMDDTHHMMSS+0800>/",
]) {
  check(`documentation contains ${snippet}`, combined.includes(snippet));
}
for (const obsolete of [
  "P1-UNIT-01 继续补齐",
  "P1-VUE-01 继续补齐",
  "P1-E2E-01 继续补齐",
  "外部安装矩阵由 P1-PACK-TEST-01 完成",
  "tests/blackbox/fixture-expectations.json",
  "tests/consumer/vue-vite",
  "最终应把它固化到仓库测试目录",
]) {
  check(`documentation removed obsolete text ${obsolete}`, !combined.includes(obsolete));
}

const runner = read("scripts/ci/run-suite.mjs");
for (const registered of [
  '"formal-preview-routes"',
  '"formal-preview-workflows"',
  '"formal-race-regression"',
  '"external-tgz-consumer"',
  '"documentation-contract"',
  'childSuite("docs", "test:docs")',
]) {
  check(`suite registration ${registered}`, runner.includes(registered));
}
check("roadmap P1-DOC record exists", roadmap.includes("| P1-DOC-TEST-01 | Codex `/root`"));
check("CI runs root check", read(".github/workflows/ci.yml").includes("pnpm check"));
check("Playwright dependency pinned", /^playwright==\d+\.\d+\.\d+$/m.test(read("requirements-ci.txt")));

const allowlist = JSON.parse(read("tests/blackbox/console_allowlist.json"));
check(
  "browser allowlist is explicit and empty",
  ["console", "pageErrors", "requestFailures", "responses"].every(
    (key) => Array.isArray(allowlist[key]) && allowlist[key].length === 0,
  ),
);
const fixtures = JSON.parse(read("test-data/manifest.json"));
check("fixture manifest has 20 entries", fixtures.entries?.length === 20, fixtures.entries?.length);
check("fixture manifest excludes sensitive data", fixtures.containsSensitiveData === false);
check(
  "consumer template has no source alias",
  !read("tests/consumer/template/vite.config.mjs").includes("alias"),
);

mkdirSync(output, { recursive: true });
const summary = {
  suite: "P1-DOCUMENTATION-CONTRACT",
  result: failures.length ? "FAIL" : "PASS",
  checks,
  failures,
};
writeFileSync(path.join(output, "summary.json"), `${JSON.stringify(summary, null, 2)}\n`);
console.log(`${summary.result}: ${checks.length} documentation contract checks at ${output}`);
if (failures.length) {
  for (const failure of failures) console.error(`- ${failure}`);
  process.exit(1);
}
