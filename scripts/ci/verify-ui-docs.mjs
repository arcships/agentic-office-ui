#!/usr/bin/env node

/**
 * UI documentation contract (#165).
 *
 * Keeps the 0.8 UI docs honest without a browser:
 *   1. Every primitive in `packages/vue-ui/src/components/ui/*.vue` and every
 *      parity surface in `tools/parity-fixtures/manifest.json` is listed in
 *      `docs/components/ui-component-map.md`, so `docs/components` is a
 *      complete entry point.
 *   2. Every `--office-*` token declared in the token layer appears in the
 *      token table (`docs/design/office-theme.md`).
 *   3. The index links the component map, the style migration guide, the CI
 *      gate doc and the release checklist.
 *   4. The release checklist still requires ui-parity, a11y and style results.
 *   5. The gate implementation (workflow, stylelint config, a11y allowlist)
 *      still matches what the docs promise.
 *
 * Environment:
 *   UI_DOCS_EVIDENCE_DIR  evidence directory (default output/acceptance/docs-contract/ui-docs)
 */
import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";

const root = path.resolve(import.meta.dirname, "../..");
const evidenceDir = path.resolve(
  process.env.UI_DOCS_EVIDENCE_DIR ||
    path.join(root, "output/acceptance/docs-contract/ui-docs"),
);
const failures = [];
const checks = [];

function read(relative) {
  return readFileSync(path.join(root, relative), "utf8");
}

function readJson(relative) {
  return JSON.parse(read(relative));
}

function check(name, condition, details = undefined) {
  checks.push({ name, status: condition ? "PASS" : "FAIL", ...(details ? { details } : {}) });
  if (!condition) failures.push(details ? `${name}: ${details}` : name);
}

const componentMap = read("docs/components/ui-component-map.md");
const tokenTable = read("docs/design/office-theme.md");
const index = read("docs/INDEX.md");
const componentReadme = read("docs/components/README.md");
const releaseChecklist = read("docs/plan/release-checklist.md");
const ciDoc = read("docs/testing/ui-parity-ci.md");
const harnessDoc = read("docs/design/extend-parity-harness.md");
const workflow = read(".github/workflows/ui-parity.yml");
const stylelintConfig = read("stylelint.config.mjs");
const parityManifest = readJson("tools/parity-fixtures/manifest.json");
const parityLiteralAllowlist = readJson("tests/style/parity-literal-allowlist.json");

// 1. Every primitive component is discoverable from docs/components.
const primitiveFiles = readdirSync(path.join(root, "packages/vue-ui/src/components/ui"))
  .filter((file) => file.startsWith("Office") && file.endsWith(".vue"))
  .map((file) => path.basename(file, ".vue"));
for (const component of primitiveFiles) {
  check(
    `component map documents ${component}`,
    componentMap.includes(`\`${component}\``),
    `packages/vue-ui/src/components/ui/${component}.vue`,
  );
}

// 2. Every parity surface has a row in the map and a doc reference.
for (const surface of parityManifest.surfaces) {
  check(
    `component map documents parity surface ${surface.id}`,
    componentMap.includes(`\`${surface.id}\``),
  );
}

// 3. Every --office-* token declaration is in the token table.
const tokenSource = ["tokens.css", "theme.css"]
  .map((file) => read(`packages/vue-ui/src/styles/${file}`))
  .join("\n");
const declaredTokens = new Set(
  [...tokenSource.matchAll(/^\s*(--office-[a-z0-9-]+)\s*:/gm)].map((match) => match[1]),
);
const missingTokens = [...declaredTokens].filter((token) => !tokenTable.includes(token));
check(
  "token table documents every --office-* declaration",
  missingTokens.length === 0,
  missingTokens.length ? `missing: ${missingTokens.slice(0, 10).join(", ")}` : undefined,
);

// 4. Navigation: docs/components and the index reach the new documents.
for (const target of [
  "ui-component-map.md",
  "../guide/ui-style-migration.md",
  "../testing/ui-parity-ci.md",
  "../plan/release-checklist.md",
]) {
  check(`components README links ${target}`, componentReadme.includes(target));
}
for (const target of [
  "components/ui-component-map.md",
  "guide/ui-style-migration.md",
  "testing/ui-parity-ci.md",
  "plan/release-checklist.md",
  "design/office-theme.md",
  "components/ui-primitives.md",
]) {
  check(`docs index links ${target}`, index.includes(target));
}
for (const target of [
  "../testing/ui-parity-ci.md",
  "../guide/ui-style-migration.md",
  "../components/ui-component-map.md",
]) {
  check(`parity harness doc links ${target}`, harnessDoc.includes(target));
}

// 5. The release checklist must require the UI evidence.
for (const requirement of [
  "ui-parity",
  "pnpm ui:baseline:verify",
  "pnpm lint:style",
  "pnpm test:a11y",
  "pnpm ui:tokens:check",
  "ui-parity-core",
]) {
  check(`release checklist requires ${requirement}`, releaseChecklist.includes(requirement));
}

// 6. CI doc keeps the switch documented.
for (const snippet of [
  "ui-parity-core",
  "ui-parity-nightly",
  "ui-contracts",
  "ui-a11y",
  "--gate-mode",
  "gateStatus",
  "strict",
  "topDiffRegions",
]) {
  check(`CI doc documents ${snippet}`, ciDoc.includes(snippet));
}
for (const snippet of [
  "--gate-mode",
  "--baseline",
  "test:a11y",
  "lint:style",
  "ui:baseline:verify",
  "ui:tokens:check",
  "gateStatus: blocking",
]) {
  check(`CI doc keeps the ${snippet} command`, ciDoc.includes(snippet));
}

// 7. The workflow actually implements what the docs say.
for (const snippet of [
  "ui-parity-core",
  "ui-parity-nightly",
  "ui-contracts",
  "ui-a11y",
  "--gate-mode",
  "--baseline",
  "pnpm lint:style",
  "pnpm ui:baseline:verify",
  "pnpm test:a11y",
  "schedule:",
  "upload-artifact",
]) {
  check(`ui-parity workflow contains ${snippet}`, workflow.includes(snippet));
}

// 8. Style guards: the four token rules exist, the token layer is exempt.
for (const rule of [
  "color-no-hex",
  "declaration-property-unit-disallowed-list",
  "declaration-property-value-disallowed-list",
]) {
  check(`stylelint config enables ${rule}`, stylelintConfig.includes(rule));
}
for (const exemption of [
  "packages/*/src/styles/tokens.css",
  "packages/*/src/styles/theme.css",
]) {
  check(`stylelint config exempts the token layer (${exemption})`, stylelintConfig.includes(exemption));
}
for (const rule of ["color-no-hex", "declaration-property-unit-disallowed-list", "declaration-property-value-disallowed-list"]) {
  check(
    `style baseline records ${rule}`,
    read("tests/style/component-style-baseline.json").includes(rule),
  );
}
check(
  "style baseline carries an approval",
  /"approved": \{[\s\S]*?"reviewer":/.test(read("tests/style/component-style-baseline.json")),
);

// #175: the two ledgers are distinct and the parity allowlist is exact. A
// glob or wildcard here would silently exempt new debt, so the docs contract
// rejects it even before the style gate runs.
check(
  "parity-literal allowlist has schemaVersion 1",
  parityLiteralAllowlist.schemaVersion === 1 && Array.isArray(parityLiteralAllowlist.entries),
);
for (const entry of parityLiteralAllowlist.entries || []) {
  check(
    `parity-literal ${entry.id || "<missing id>"} is exact and owned`,
    Boolean(
      entry.id &&
        entry.file &&
        !/[*?[\]{}]/u.test(entry.file) &&
        entry.rule &&
        entry.value &&
        Number.isInteger(entry.count) &&
        entry.count > 0 &&
        entry.upstreamPath &&
        entry.upstreamCommit &&
        entry.issue &&
        entry.reason &&
        entry.reviewBy,
    ),
    entry.file,
  );
}
for (const snippet of [
  "component-style-baseline.json",
  "parity-literal-allowlist.json",
  "upstreamCommit",
  "reviewBy",
  "stale",
]) {
  check(`CI doc explains the two style ledgers (${snippet})`, ciDoc.includes(snippet));
}
check(
  "lint-styles mentions both style ledgers",
  read("scripts/ci/lint-styles.mjs").includes("parity-literal-allowlist.json") &&
    read("scripts/ci/lint-styles.mjs").includes("component-style-baseline.json"),
);

// 9. a11y allowlist entries have an owner issue and a reason.
const a11yConfig = readJson("tests/a11y/config.json");
const a11yAllowlist = readJson("tests/a11y/allowlist.json");
for (const entry of a11yAllowlist.entries || []) {
  check(
    `a11y allowlist entry ${entry.surface}/${entry.rule || entry.check} has an issue and reason`,
    Boolean(entry.issue && entry.reason && entry.surface && (entry.rule || entry.check)),
  );
}
for (const surface of a11yConfig.surfaces || []) {
  check(
    `a11y config covers parity surface ${surface.id}`,
    parityManifest.surfaces.some((candidate) => candidate.id === surface.id),
    `unknown surface ${surface.id}`,
  );
}

// 10. Baselines exist and are named by component/state/theme/viewport.
const baselineManifestPath = "tests/visual/extend/manifest.json";
if (existsSync(path.join(root, baselineManifestPath))) {
  const baselines = readJson(baselineManifestPath);
  check(
    "visual baseline naming is component/state/theme-viewport",
    baselines.naming === parityManifest.baselines.naming,
  );
  for (const file of baselines.files || []) {
    check(
      `baseline file name ${file.path}`,
      /^[a-z0-9-]+\/[a-z0-9-]+\/(light|dark)-[0-9]+x[0-9]+\.png$/.test(file.path),
    );
  }
  check(
    "visual baseline approvals exist",
    Array.isArray(baselines.approvals) && baselines.approvals.length > 0,
  );
  check(
    "visual baseline README explains the approval flow",
    read("tests/visual/extend/README.md").includes("--approve"),
  );
} else {
  check(`visual baseline manifest ${baselineManifestPath}`, false);
}

mkdirSync(evidenceDir, { recursive: true });
const summary = {
  suite: "P1-UI-DOCUMENTATION-CONTRACT",
  result: failures.length ? "FAIL" : "PASS",
  checks,
  failures,
};
writeFileSync(path.join(evidenceDir, "summary.json"), `${JSON.stringify(summary, null, 2)}\n`);
console.log(`${summary.result}: ${checks.length} UI documentation checks at ${evidenceDir}`);
if (failures.length) {
  for (const failure of failures) console.error(`- ${failure}`);
  process.exit(1);
}
