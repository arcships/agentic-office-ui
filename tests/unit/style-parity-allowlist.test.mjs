// Issue #175 — parity-literal allowlist vs. legacy style baseline.
//
// The style gate has two ledgers:
//   - tests/style/component-style-baseline.json freezes historical debt only;
//   - tests/style/parity-literal-allowlist.json names exact upstream literals.
//
// These tests pin the policy:
//   - warning.text parsing is robust for hex/unit/value messages;
//   - allowlist entries are exact (no globs/wildcards), owned and time-boxed;
//   - an allowlisted literal passes and appears in the evidence;
//   - a bare value outside both ledgers fails;
//   - stale entries are reported, expired/invalid entries fail closed;
//   - the legacy baseline still absorbs its frozen warnings.
//
// The behavioural checks spawn the real CLI against a temporary fixture under
// `output/`, so the exact stylelint message text and the evidence file are
// exercised, not just a mocked parser.

import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import {
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import path from "node:path";
import test from "node:test";
import {
  PINNED_UPSTREAM_COMMIT,
  matchParityLiterals,
  parseStylelintWarningText,
  validateParityAllowlist,
} from "../../scripts/ci/style-parity-allowlist.mjs";

const ROOT_URL = new URL("../../", import.meta.url);
const rootPath = path.resolve(new URL(".", ROOT_URL).pathname);
const today = "2026-11-21";

function readJson(relativePath) {
  return JSON.parse(readFileSync(new URL(relativePath, ROOT_URL), "utf8"));
}

function validEntry(overrides = {}) {
  return {
    id: "sample-entry",
    file: "packages/vue-pptx/src/PptxViewerToolbar.vue",
    rule: "declaration-property-unit-disallowed-list",
    property: "box-shadow",
    value: "px",
    count: 16,
    upstreamPath: "apps/v4/components/extend/pptx-viewer.tsx",
    upstreamCommit: PINNED_UPSTREAM_COMMIT,
    issue: "#161",
    reason: "The pinned upstream Tailwind shadow utility supplies this pixel geometry.",
    reviewBy: "2027-05-21",
    ...overrides,
  };
}

function relativeToRoot(absolute) {
  return path.relative(rootPath, absolute).split(path.sep).join("/");
}

function fixture(t) {
  mkdirSync(path.join(rootPath, "output"), { recursive: true });
  const directory = mkdtempSync(path.join(rootPath, "output/style-parity-test-"));
  t.after(() => rmSync(directory, { recursive: true, force: true }));
  return {
    directory,
    write(name, content) {
      const absolute = path.join(directory, name);
      writeFileSync(absolute, content);
      return absolute;
    },
    writeJson(name, value) {
      return this.write(name, `${JSON.stringify(value, null, 2)}\n`);
    },
  };
}

function runGate({ files, baseline, allowlist, evidence, args = [] }) {
  const result = spawnSync(process.execPath, ["scripts/ci/lint-styles.mjs", ...args], {
    cwd: rootPath,
    encoding: "utf8",
    timeout: 120_000,
    env: {
      ...process.env,
      STYLE_GUARD_FILES: files.join(","),
      STYLE_GUARD_BASELINE: baseline,
      STYLE_GUARD_ALLOWLIST: allowlist,
      STYLE_GUARD_EVIDENCE_DIR: evidence,
      STYLE_GUARD_TODAY: today,
    },
  });
  const summaryPath = path.join(rootPath, evidence, "summary.json");
  let summary = null;
  try {
    summary = JSON.parse(readFileSync(summaryPath, "utf8"));
  } catch {
    summary = null;
  }
  return { ...result, summary };
}

test("stylelint warning text parses hex, unit and value messages", () => {
  assert.deepEqual(
    parseStylelintWarningText('Unexpected hex color "#fff" (color-no-hex)'),
    { kind: "hex", property: null, value: "#fff", unit: null },
  );
  assert.deepEqual(
    parseStylelintWarningText(
      'Unexpected unit "px" for property "border-radius" (declaration-property-unit-disallowed-list)',
    ),
    { kind: "unit", property: "border-radius", value: "px", unit: "px" },
  );
  assert.deepEqual(
    parseStylelintWarningText(
      'Unexpected value "0 12px 40px rgba(0, 0, 0, 0.12)" for property "box-shadow" (declaration-property-value-disallowed-list)',
    ),
    {
      kind: "value",
      property: "box-shadow",
      value: "0 12px 40px rgba(0, 0, 0, 0.12)",
      unit: null,
    },
  );
  assert.deepEqual(parseStylelintWarningText("some future message"), {
    kind: "unknown",
    property: null,
    value: null,
    unit: null,
  });
  assert.deepEqual(parseStylelintWarningText(""), {
    kind: "unknown",
    property: null,
    value: null,
    unit: null,
  });
});

test("allowlist schema rejects globs, wildcards, missing owners and expired reviews", () => {
  const wildcardFile = validateParityAllowlist(
    { schemaVersion: 1, entries: [validEntry({ file: "packages/vue-pptx/src/*.vue" })] },
    { root: rootPath, today },
  );
  assert.equal(wildcardFile.entries.length, 0);
  assert.match(JSON.stringify(wildcardFile.problems), /no globs|does not exist/);

  const wildcardProperty = validateParityAllowlist(
    { schemaVersion: 1, entries: [validEntry({ property: "*" })] },
    { root: rootPath, today },
  );
  assert.equal(wildcardProperty.entries.length, 0);
  assert.match(JSON.stringify(wildcardProperty.problems), /no wildcards/);

  const wildcardValue = validateParityAllowlist(
    { schemaVersion: 1, entries: [validEntry({ value: "p*" })] },
    { root: rootPath, today },
  );
  assert.equal(wildcardValue.entries.length, 0);
  assert.match(JSON.stringify(wildcardValue.problems), /no wildcards/);

  const missingOwner = validateParityAllowlist(
    {
      schemaVersion: 1,
      entries: [validEntry({ reason: "too short", reviewBy: undefined, issue: undefined })],
    },
    { root: rootPath, today },
  );
  assert.equal(missingOwner.entries.length, 0);
  assert.match(JSON.stringify(missingOwner.problems), /issue/);
  assert.match(JSON.stringify(missingOwner.problems), /reason/);
  assert.match(JSON.stringify(missingOwner.problems), /reviewBy/);

  const expired = validateParityAllowlist(
    { schemaVersion: 1, entries: [validEntry({ reviewBy: "2020-01-01" })] },
    { root: rootPath, today },
  );
  assert.equal(expired.entries.length, 0);
  assert.equal(expired.expired.length, 1);
  assert.match(expired.expired[0].message, /reviewBy 2020-01-01 has passed/);
});

test("a new bare value fails against an empty baseline and allowlist", (t) => {
  const files = fixture(t);
  const fixturePath = files.write("fixture.css", ".x {\n  border-radius: 8px;\n}\n");
  const baseline = files.writeJson("baseline.json", { schemaVersion: 1, counts: {} });
  const allowlist = files.writeJson("allowlist.json", { schemaVersion: 1, entries: [] });
  const evidence = relativeToRoot(path.join(files.directory, "evidence"));

  const result = runGate({
    files: [fixturePath],
    baseline: relativeToRoot(baseline),
    allowlist: relativeToRoot(allowlist),
    evidence,
  });

  assert.equal(result.status, 1, `${result.stdout}\n${result.stderr}`);
  assert.match(result.stderr, /new component style violations/u);
  assert.match(result.stderr, /border-radius/u);
  assert.equal(result.summary.result, "FAIL");
  assert.equal(result.summary.totals.new, 1);
  assert.equal(result.summary.totals.parityAllowlisted, 0);
  assert.equal(result.summary.newViolations[0].file, relativeToRoot(fixturePath));
});

test("an allowlisted upstream literal passes and appears in the evidence", (t) => {
  const files = fixture(t);
  const fixturePath = files.write("fixture.css", ".x {\n  transition: color 150ms ease;\n}\n");
  const baseline = files.writeJson("baseline.json", { schemaVersion: 1, counts: {} });
  const allowlist = files.writeJson("allowlist.json", {
    schemaVersion: 1,
    entries: [
      validEntry({
        id: "fixture-transition",
        file: relativeToRoot(fixturePath),
        rule: "declaration-property-unit-disallowed-list",
        property: "transition",
        value: "ms",
        count: 1,
      }),
    ],
  });
  const evidence = relativeToRoot(path.join(files.directory, "evidence"));

  const result = runGate({
    files: [fixturePath],
    baseline: relativeToRoot(baseline),
    allowlist: relativeToRoot(allowlist),
    evidence,
  });

  assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);
  assert.match(result.stdout, /1 parity-literal/u);
  assert.match(result.stdout, /0 new/u);
  assert.equal(result.summary.result, "PASS");
  assert.equal(result.summary.totals.warnings, 1);
  assert.equal(result.summary.totals.parityAllowlisted, 1);
  assert.equal(result.summary.totals.baselined, 0);
  assert.equal(result.summary.totals.new, 0);
  assert.equal(result.summary.parityAllowlist.matched, 1);
  assert.equal(result.summary.parityAllowlist.matches[0].id, "fixture-transition");
  assert.deepEqual(result.summary.parityAllowlist.matches[0].warnings, [
    {
      file: relativeToRoot(fixturePath),
      line: 2,
      column: 21,
      rule: "declaration-property-unit-disallowed-list",
      property: "transition",
      value: "ms",
    },
  ]);
  assert.deepEqual(result.summary.parityAllowlist.stale, []);
  assert.deepEqual(result.summary.parityAllowlist.invalid, []);
  assert.deepEqual(result.summary.parityAllowlist.expired, []);
});

test("stale allowlist entries are reported without blocking", (t) => {
  const files = fixture(t);
  const fixturePath = files.write("fixture.css", ".x {\n  border-radius: 8px;\n}\n");
  const baseline = files.writeJson("baseline.json", { schemaVersion: 1, counts: {} });
  const allowlist = files.writeJson("allowlist.json", {
    schemaVersion: 1,
    entries: [
      validEntry({
        id: "fixture-stale",
        file: relativeToRoot(fixturePath),
        property: "border-radius",
        value: "px",
        count: 2,
      }),
    ],
  });
  const evidence = relativeToRoot(path.join(files.directory, "evidence"));

  const result = runGate({
    files: [fixturePath],
    baseline: relativeToRoot(baseline),
    allowlist: relativeToRoot(allowlist),
    evidence,
  });

  assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);
  assert.match(result.stdout, /stale/u);
  assert.equal(result.summary.result, "PASS");
  assert.equal(result.summary.totals.parityAllowlisted, 1);
  assert.equal(result.summary.parityAllowlist.stale.length, 1);
  assert.deepEqual(result.summary.parityAllowlist.stale[0], {
    id: "fixture-stale",
    file: relativeToRoot(fixturePath),
    rule: "declaration-property-unit-disallowed-list",
    property: "border-radius",
    value: "px",
    expected: 2,
    matched: 1,
    missing: 1,
  });
});

test("expired allowlist entries fail closed instead of exempting code", (t) => {
  const files = fixture(t);
  const fixturePath = files.write("fixture.css", ".x {\n  border-radius: 8px;\n}\n");
  const baseline = files.writeJson("baseline.json", { schemaVersion: 1, counts: {} });
  const allowlist = files.writeJson("allowlist.json", {
    schemaVersion: 1,
    entries: [
      validEntry({
        id: "fixture-expired",
        file: relativeToRoot(fixturePath),
        property: "border-radius",
        value: "px",
        count: 1,
        reviewBy: "2020-01-01",
      }),
    ],
  });
  const evidence = relativeToRoot(path.join(files.directory, "evidence"));

  const result = runGate({
    files: [fixturePath],
    baseline: relativeToRoot(baseline),
    allowlist: relativeToRoot(allowlist),
    evidence,
  });

  assert.equal(result.status, 1, `${result.stdout}\n${result.stderr}`);
  assert.match(result.stderr, /reviewBy 2020-01-01 has passed/u);
  assert.equal(result.summary.result, "FAIL");
  assert.equal(result.summary.parityAllowlist.expired.length, 1);
  assert.equal(result.summary.totals.parityAllowlisted, 0);
  assert.equal(result.summary.totals.new, 1);
});

test("the legacy baseline still absorbs warnings after parity deduction", (t) => {
  const files = fixture(t);
  const fixturePath = files.write(
    "fixture.css",
    ".a {\n  border-radius: 8px;\n}\n.b {\n  border-radius: 4px;\n}\n",
  );
  const baseline = files.writeJson("baseline.json", {
    schemaVersion: 1,
    counts: {
      [relativeToRoot(fixturePath)]: {
        "declaration-property-unit-disallowed-list": 1,
      },
    },
  });
  const allowlist = files.writeJson("allowlist.json", {
    schemaVersion: 1,
    entries: [
      validEntry({
        id: "fixture-parity",
        file: relativeToRoot(fixturePath),
        property: "border-radius",
        value: "px",
        count: 1,
      }),
    ],
  });
  const evidence = relativeToRoot(path.join(files.directory, "evidence"));

  const result = runGate({
    files: [fixturePath],
    baseline: relativeToRoot(baseline),
    allowlist: relativeToRoot(allowlist),
    evidence,
  });

  assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);
  assert.equal(result.summary.result, "PASS");
  assert.equal(result.summary.totals.warnings, 2);
  assert.equal(result.summary.totals.parityAllowlisted, 1);
  assert.equal(result.summary.totals.baselined, 1);
  assert.equal(result.summary.totals.new, 0);
  assert.equal(result.summary.parityAllowlist.matches[0].matched, 1);
});

test("write-baseline excludes parity literals instead of freezing them as legacy debt", (t) => {
  const files = fixture(t);
  const fixturePath = files.write("fixture.css", ".x {\n  transition: color 150ms ease;\n}\n");
  const baseline = files.writeJson("baseline.json", { schemaVersion: 1, counts: {} });
  const allowlist = files.writeJson("allowlist.json", {
    schemaVersion: 1,
    entries: [
      validEntry({
        id: "fixture-next-baseline",
        file: relativeToRoot(fixturePath),
        rule: "declaration-property-unit-disallowed-list",
        property: "transition",
        value: "ms",
        count: 1,
      }),
    ],
  });
  const evidence = relativeToRoot(path.join(files.directory, "evidence"));

  const result = runGate({
    files: [fixturePath],
    baseline: relativeToRoot(baseline),
    allowlist: relativeToRoot(allowlist),
    evidence,
    args: [
      "--write-baseline",
      "--approve",
      "--reviewer",
      "tester",
      "--issue",
      "#175",
      "--reason",
      "exercise the refresh path without absorbing parity literals",
    ],
  });

  assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);
  const written = JSON.parse(readFileSync(baseline, "utf8"));
  assert.deepEqual(written.counts, {});
  assert.equal(written.totals.warnings, 0);
  assert.equal(result.summary.result, "BASELINE-WRITTEN");
  assert.equal(result.summary.parityAllowlist.matched, 1);
  assert.equal(result.summary.parityAllowlisted, 1);
});

test("a new bare value in an already allowlisted file still fails when the count is exhausted", (t) => {
  const files = fixture(t);
  const fixturePath = files.write(
    "fixture.css",
    ".a {\n  border-radius: 8px;\n}\n.b {\n  border-radius: 4px;\n}\n",
  );
  const baseline = files.writeJson("baseline.json", { schemaVersion: 1, counts: {} });
  const allowlist = files.writeJson("allowlist.json", {
    schemaVersion: 1,
    entries: [
      validEntry({
        id: "fixture-one-of-two",
        file: relativeToRoot(fixturePath),
        property: "border-radius",
        value: "px",
        count: 1,
      }),
    ],
  });
  const evidence = relativeToRoot(path.join(files.directory, "evidence"));

  const result = runGate({
    files: [fixturePath],
    baseline: relativeToRoot(baseline),
    allowlist: relativeToRoot(allowlist),
    evidence,
  });

  assert.equal(result.status, 1, `${result.stdout}\n${result.stderr}`);
  assert.equal(result.summary.result, "FAIL");
  assert.equal(result.summary.totals.parityAllowlisted, 1);
  assert.equal(result.summary.totals.new, 1);
  assert.equal(result.summary.newViolations.length, 1);
});

test("write-baseline refuses to freeze a new bare value as legacy debt", (t) => {
  const files = fixture(t);
  const fixturePath = files.write("fixture.css", ".x {\n  border-radius: 8px;\n}\n");
  const baseline = files.writeJson("baseline.json", { schemaVersion: 1, counts: {} });
  const allowlist = files.writeJson("allowlist.json", { schemaVersion: 1, entries: [] });
  const evidence = relativeToRoot(path.join(files.directory, "evidence"));

  const result = runGate({
    files: [fixturePath],
    baseline: relativeToRoot(baseline),
    allowlist: relativeToRoot(allowlist),
    evidence,
    args: [
      "--write-baseline",
      "--approve",
      "--reviewer",
      "tester",
      "--issue",
      "#175",
      "--reason",
      "try to freeze a new violation as legacy debt",
    ],
  });

  assert.equal(result.status, 2, `${result.stdout}\n${result.stderr}`);
  assert.match(result.stderr, /shrink-only/u);
  assert.equal(result.summary.result, "FAIL");
  assert.equal(result.summary.growth.length, 1);
  assert.deepEqual(JSON.parse(readFileSync(baseline, "utf8")).counts, {});
});

test("the committed ledgers cover the #161b parity literals and keep legacy debt separate", () => {
  const document = readJson("tests/style/parity-literal-allowlist.json");
  const validation = validateParityAllowlist(document, { root: rootPath, today });
  assert.deepEqual(validation.problems, []);
  assert.deepEqual(validation.expired, []);
  assert.equal(document.schemaVersion, 1);
  assert.equal(validation.entries.length, 3);
  assert.equal(
    validation.entries.reduce((total, entry) => total + entry.count, 0),
    22,
  );
  for (const entry of validation.entries) {
    assert.ok(entry.file.startsWith("packages/"), `${entry.id} must name component CSS`);
    assert.ok(!entry.file.includes("*"), `${entry.id} must not use a glob`);
    assert.ok(entry.upstreamCommit, `${entry.id} must pin the upstream commit`);
    assert.ok(entry.issue.startsWith("#"), `${entry.id} must name an issue`);
    assert.ok(entry.reason.length >= 40, `${entry.id} must explain the exception`);
    assert.match(entry.reviewBy, /^\d{4}-\d{2}-\d{2}$/u);
  }
  assert.deepEqual(
    validation.entries.map((entry) => entry.id).sort(),
    [
      "pptx-slide-card-transition",
      "pptx-viewer-toolbar-ring-radius",
      "pptx-viewer-toolbar-ring-shadow",
    ],
  );

  const baseline = readJson("tests/style/component-style-baseline.json");
  assert.ok(!baseline.counts["packages/vue-pptx/src/PptxViewerToolbar.vue"]);
  assert.ok(!baseline.counts["packages/vue-pptx/src/PptxSlideCard.vue"]);
  assert.match(baseline.approved.issue, /^#\d+$/u);
});

test("matching is exact: file, rule, property and value all participate", () => {
  const warnings = [
    {
      file: "packages/vue-pptx/src/PptxViewerToolbar.vue",
      rule: "declaration-property-unit-disallowed-list",
      line: 1,
      column: 1,
      text: 'Unexpected unit "px" for property "box-shadow" (x)',
    },
    {
      file: "packages/vue-pptx/src/PptxViewerToolbar.vue",
      rule: "declaration-property-unit-disallowed-list",
      line: 2,
      column: 1,
      text: 'Unexpected unit "px" for property "border-radius" (x)',
    },
  ];
  const entry = validEntry({ id: "shadow-only", property: "box-shadow", count: 1 });
  const result = matchParityLiterals(warnings, [entry]);
  assert.equal(result.consumed.size, 1);
  assert.equal(result.matches[0].warnings[0].line, 1);
  assert.equal(result.stale.length, 0);

  const wrongValue = matchParityLiterals(warnings, [validEntry({ value: "ms", count: 1 })]);
  assert.equal(wrongValue.consumed.size, 0);
  assert.equal(wrongValue.stale[0].missing, 1);
});

