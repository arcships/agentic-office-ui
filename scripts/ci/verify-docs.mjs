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
  "test:stress": "node scripts/ci/run-suite.mjs stress",
  "test:performance": "node scripts/ci/run-suite.mjs performance",
  "test:matrix": "node scripts/ci/run-suite.mjs matrix",
  "test:docs": "node scripts/ci/run-suite.mjs docs",
  "test:release": "node scripts/ci/run-suite.mjs release",
  check: "node scripts/ci/run-suite.mjs check",
};
for (const [name, value] of Object.entries(expectedScripts)) {
  check(`root script ${name}`, packageJson.scripts?.[name] === value, packageJson.scripts?.[name]);
}

const requiredFiles = [
  "README.md",
  "RELEASE_NOTES.md",
  "requirements-ci.txt",
  ".github/workflows/ci.yml",
  ".github/workflows/release.yml",
  "docs/INDEX.md",
  "docs/api/public-api-contract.md",
  "docs/migration-0.2.md",
  "docs/testing/compatibility-matrix.md",
  "scripts/ci/compatibility-matrix.mjs",
  "scripts/ci/p4-reproducible-pack.mjs",
  "scripts/ci/prepare-release-artifact.mjs",
  "test-data/manifest.json",
  "tests/unit/docx-core.test.mjs",
  "tests/unit/xlsx-core.test.mjs",
  "tests/component/vue-components.test.mjs",
  "tests/component/docx-selection-input.test.mjs",
  "tests/blackbox/browser_evidence.py",
  "tests/blackbox/console_allowlist.json",
  "tests/blackbox/routes_smoke.py",
  "tests/blackbox/e2e_workflows.py",
  "tests/blackbox/fault_server.py",
  "tests/blackbox/race_workflows.py",
  "tests/blackbox/pack_consumer.py",
  "tests/blackbox/ux_parity.py",
  "tests/consumer/template/src/App.vue",
  "scripts/ci/pack-manifests.mjs",
  "scripts/ci/pack-consumer.mjs",
  "scripts/ci/verify-fixtures.mjs",
  "scripts/ci/verify-docs.mjs",
  "tests/baseline/README.md",
];
if (process.env.DOCS_PROBE_MISSING === "1") {
  requiredFiles.push("tests/blackbox/__intentional_missing_documented_script__.py");
}
for (const file of requiredFiles) check(`documented file ${file}`, existsSync(path.join(root, file)));

const plan = read("docs/end-to-end-blackbox-test-plan.md");
const runbook = read("docs/testing/agent-execution-runbook.md");
const roadmap = read("docs/plan/stabilization-roadmap.md");
const readme = read("README.md");
const index = read("docs/INDEX.md");
const releaseNotes = read("RELEASE_NOTES.md");
const migration = read("docs/migration-0.2.md");
const compatibility = read("docs/testing/compatibility-matrix.md");
const combined = `${plan}\n${runbook}`;
for (const snippet of [
  "pnpm test:unit",
  "pnpm test:component",
  "pnpm test:blackbox",
  "pnpm test:consumer",
  "pnpm test:matrix",
  "pnpm test:docs",
  "pnpm test:release",
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

for (const [name, content] of Object.entries({ readme, releaseNotes, migration })) {
  check(`${name} says 0.2.0 is not released`, content.includes("0.2.0") && content.includes("尚未发布"));
}
check(
  "README keeps PDF usable",
  ["翻页", "缩放", "搜索", "下载", "50 MiB", "PDF_TOO_LARGE"].every((term) => readme.includes(term)),
);
check(
  "release notes cover required migration facts",
  [
    "createDocxRuntime",
    "createXlsxRuntime",
    "弃用",
    "1.0.0",
    "50 MiB",
    "PDF_TOO_LARGE",
    "Worker",
    "WASM",
    "回退办法",
  ].every((term) => releaseNotes.includes(term)),
);
check(
  "migration uses public package resources",
  [
    "@arcships/docx-core/worker?worker&url",
    "@arcships/xlsx-core/worker?worker&url",
    "@arcships/vue-docx/style.css",
    "@arcships/vue-xlsx/style.css",
    "@arcships/vue-pdf/style.css",
    "PDF_TOO_LARGE",
  ].every((term) => migration.includes(term)),
);
check(
  "compatibility matrix declares tools and browsers",
  ["Node.js", "Vue", "Vite", "TypeScript", "Chromium", "Firefox", "WebKit", "待验证"].every(
    (term) => compatibility.includes(term),
  ),
);
check(
  "documentation index links candidate documents",
  ["RELEASE_NOTES.md", "migration-0.2.md", "testing/compatibility-matrix.md"].every((term) =>
    index.includes(term),
  ),
);

for (const file of [
  "docs/docx-migration-architecture.md",
  "docs/xlsx-migration-architecture.md",
  "docs/upstream-docx-feature-alignment.md",
  "docs/upstream-xlsx-feature-alignment.md",
]) {
  check(`historical notice ${file}`, read(file).includes("历史资料"));
}
check("historical notice docs/plan/README.md", read("docs/plan/README.md").includes("历史迁移计划"));

for (const file of [
  "README.md",
  "RELEASE_NOTES.md",
  "docs/INDEX.md",
  "docs/migration-0.2.md",
  "docs/testing/compatibility-matrix.md",
]) {
  const content = read(file);
  for (const match of content.matchAll(/\[[^\]]+\]\(([^)]+)\)/g)) {
    const rawTarget = match[1].trim().replace(/^<|>$/g, "");
    if (/^(?:https?:|mailto:|#)/i.test(rawTarget)) continue;
    const target = decodeURIComponent(rawTarget.split("#", 1)[0]);
    if (!target) continue;
    const absolute = path.resolve(root, path.dirname(file), target);
    check(`relative link ${file} -> ${rawTarget}`, existsSync(absolute));
  }
}

const runner = read("scripts/ci/run-suite.mjs");
for (const registered of [
  '"formal-preview-routes"',
  '"formal-preview-workflows"',
  '"formal-race-regression"',
  '"compatibility-matrix"',
  '"external-tgz-consumer"',
  '"documentation-contract"',
  '"p3-regression-baselines"',
  '"formal-scroll-range-regression"',
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
check("fixture manifest has 21 entries", fixtures.entries?.length === 21, fixtures.entries?.length);
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
