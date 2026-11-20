// Issue #165 — visual-regression gate, style guards and a11y contract.
//
// These assertions pin the infrastructure that makes a broken viewer fail in
// CI: the gate modes and thresholds in the parity manifest, the committed
// baseline tree and its approval contract, the stylelint token rules with
// their frozen baseline, the axe/keyboard configuration, and the workflow that
// wires them together.
//
// The behavioural checks spawn the real CLIs:
//   - `scripts/ci/lint-styles.mjs --report-only` must pass against the frozen
//     baseline (a new violation fails it, which is the guard).
//   - `scripts/ci/visual-baseline.mjs verify --require-approval` must pass
//     (an edited PNG without an approval record fails it).
//   - both capture paths must refuse to write without an explicit approval.

import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { existsSync, readFileSync, readdirSync } from "node:fs";
import path from "node:path";
import test from "node:test";

const ROOT = new URL("../../", import.meta.url);
const rootPath = path.resolve(new URL(".", ROOT).pathname);

function read(relativePath) {
  return readFileSync(new URL(relativePath, ROOT), "utf8");
}

function readJson(relativePath) {
  return JSON.parse(read(relativePath));
}

function run(args, env = {}) {
  return spawnSync(process.execPath, args, {
    cwd: rootPath,
    encoding: "utf8",
    env: { ...process.env, ...env },
    timeout: 120_000,
  });
}

const manifest = readJson("tools/parity-fixtures/manifest.json");
const baselineManifest = readJson("tests/visual/extend/manifest.json");
const styleBaseline = readJson("tests/style/component-style-baseline.json");
const a11yConfig = readJson("tests/a11y/config.json");
const a11yAllowlist = readJson("tests/a11y/allowlist.json");
const workflow = read(".github/workflows/ui-parity.yml");
const stylelintConfig = read("stylelint.config.mjs");

test("parity thresholds stay at geometry 0px, pixels <=0.1%, top-N regions", () => {
  assert.equal(manifest.thresholds.geometryTolerancePx, 0);
  assert.equal(manifest.thresholds.maxDiffPixelRatio, 0.001);
  assert.equal(manifest.thresholds.maxChannelDelta, 2);
  assert.equal(manifest.thresholds.topDiffRegions, 5);
  // Gate modes: default `all` keeps `pnpm ui:parity` behaviour, CI uses `ready`.
  assert.deepEqual(Object.keys(manifest.gate.modes).sort(), ["all", "off", "ready"]);
  assert.equal(manifest.gate.prMode, "ready");
  assert.equal(manifest.gate.finalMode, "all");
});

test("manifest marks every surface core/blocking with an owner for pending ones", () => {
  const core = manifest.surfaces.filter((surface) => surface.core).map((surface) => surface.id);
  assert.deepEqual(core, manifest.coreSubset.surfaces);
  assert.ok(core.includes("button"));
  assert.ok(core.includes("input"));
  assert.ok(core.includes("docx-toolbar"), "closed DOCX chrome gates on pull requests");
  assert.ok(core.includes("xlsx-toolbar"), "closed XLSX chrome gates on pull requests");
  assert.ok(!core.includes("bounding-box-citations"), "the large block matrix stays nightly-only");
  assert.ok(!core.includes("document-sidebar"), "the pending sidebar stays out of the PR subset");

  for (const surface of manifest.surfaces) {
    assert.ok(["blocking", "pending"].includes(surface.gateStatus), `${surface.id} gateStatus`);
    if (surface.gateStatus === "pending") {
      assert.match(
        surface.gatePendingIssue ?? "",
        /^#\d+$/,
        `${surface.id} is pending without a gatePendingIssue`,
      );
    } else {
      assert.equal(surface.gatePendingIssue, undefined);
    }
  }

  // The surfaces that are green today must gate; the known deltas must not.
  const blocking = manifest.surfaces
    .filter((surface) => surface.gateStatus === "blocking")
    .map((surface) => surface.id)
    .sort();
  assert.deepEqual(blocking, [
    "bounding-box-citations",
    "button",
    "dialog",
    "document-splits",
    "docx-sidebar",
    "docx-toolbar",
    "e-signature",
    "file-thumbnail",
    "file-upload",
    "input",
    "layout-blocks",
    "pdf-sidebar",
    "pdf-toolbar",
    "select",
    "tooltip",
    "xlsx-sheet-tabs",
    "xlsx-status-bar",
    "xlsx-toolbar",
  ]);
  const pending = manifest.surfaces
    .filter((surface) => surface.gateStatus === "pending")
    .map((surface) => `${surface.id}:${surface.gatePendingIssue}`)
    .sort();
  // Only the canonical DocumentViewerSidebar (#158) is still in flight.
  assert.deepEqual(pending, ["document-sidebar:#158"]);
});

test("committed baselines cover exactly the gate-blocking matrix", () => {
  assert.equal(manifest.baselines.directory, "tests/visual/extend");
  assert.equal(manifest.baselines.manifest, "tests/visual/extend/manifest.json");
  assert.equal(manifest.baselines.naming, "<surface>/<state>/<theme>-<viewport>.png");
  assert.deepEqual(manifest.baselines.viewports, ["1440x900", "390x844"]);
  assert.deepEqual(manifest.baselines.themes, ["light", "dark"]);
  assert.equal(manifest.baselines.capturePaddingPx, 12);

  assert.equal(baselineManifest.naming, manifest.baselines.naming);
  assert.deepEqual(baselineManifest.viewports, manifest.baselines.viewports);
  assert.deepEqual(baselineManifest.themes, manifest.baselines.themes);
  assert.equal(baselineManifest.capturePaddingPx, 12);
  assert.equal(baselineManifest.deviceScaleFactor, manifest.deviceScaleFactor);
  assert.equal(baselineManifest.referenceCommit, manifest.referenceCommit);

  const expected = new Set();
  for (const surface of manifest.surfaces.filter((item) => item.gateStatus === "blocking")) {
    for (const state of surface.states) {
      for (const theme of manifest.baselines.themes) {
        for (const viewport of manifest.baselines.viewports) {
          expected.add(`${surface.id}/${state.id}/${theme}-${viewport}.png`);
        }
      }
    }
  }
  const recorded = new Set(baselineManifest.files.map((file) => file.path));
  assert.deepEqual([...recorded].sort(), [...expected].sort());
  assert.equal(baselineManifest.files.length, 348);

  for (const file of baselineManifest.files) {
    assert.match(file.path, /^[a-z0-9-]+\/[a-z0-9-]+\/(light|dark)-[0-9]+x[0-9]+\.png$/);
    const absolute = new URL(`tests/visual/extend/${file.path}`, ROOT);
    assert.ok(existsSync(absolute), `${file.path} is missing on disk`);
    const sha256 = createHash("sha256").update(readFileSync(absolute)).digest("hex");
    assert.equal(sha256, file.sha256, `${file.path} does not match the recorded hash`);
    assert.ok(file.width > 0 && file.height > 0);
  }

  // No stray PNGs: a hand-dropped image would otherwise silently join the tree.
  const diskFiles = readdirSync(new URL("tests/visual/extend", ROOT), {
    recursive: true,
    withFileTypes: true,
  })
    .filter((entry) => entry.isFile() && entry.name.endsWith(".png"))
    .map((entry) => path.relative(rootPath, path.join(entry.parentPath, entry.name)).split(path.sep).join("/"))
    .map((file) => file.replace(/^tests\/visual\/extend\//, ""))
    .sort();
  assert.deepEqual(diskFiles, [...recorded].sort());
});

test("baseline approvals cover the committed tree", () => {
  const approvals = baselineManifest.approvals;
  assert.ok(Array.isArray(approvals) && approvals.length >= 1);
  const latest = approvals.at(-1);
  for (const field of ["date", "reviewer", "issue", "reason", "treeSha256"]) {
    assert.ok(latest[field], `approval is missing ${field}`);
  }
  assert.match(latest.issue, /^#\d+$/);
  assert.ok(latest.reason.length >= 8);
  const lines = baselineManifest.files
    .map((file) => `${file.path} ${file.sha256}`)
    .sort()
    .join("\n");
  const treeSha256 = createHash("sha256").update(`${lines}\n`).digest("hex");
  assert.equal(baselineManifest.treeSha256, treeSha256);
  assert.equal(latest.treeSha256, treeSha256);
  assert.match(read("tests/visual/extend/APPROVALS.md"), /--approve/);

  const verify = run(["scripts/ci/visual-baseline.mjs", "verify", "--require-approval"]);
  assert.equal(verify.status, 0, `${verify.stdout}\n${verify.stderr}`);
});

test("baseline capture refuses to write without an approval", () => {
  const capture = run(["scripts/ci/visual-baseline.mjs", "capture", "--only", "button"]);
  assert.notEqual(capture.status, 0);
  assert.match(`${capture.stdout}${capture.stderr}`, /requires an approval/);

  const partial = run([
    "scripts/ci/visual-baseline.mjs",
    "capture",
    "--only",
    "button",
    "--approve",
    "--reviewer",
    "someone",
  ]);
  assert.notEqual(partial.status, 0);
  assert.match(`${partial.stdout}${partial.stderr}`, /--issue/);
});

test("component style guards keep the four token rules and freeze only existing debt", () => {
  assert.match(stylelintConfig, /"color-no-hex": true/);
  assert.match(stylelintConfig, /"declaration-property-unit-disallowed-list"/);
  assert.match(stylelintConfig, /"declaration-property-value-disallowed-list"/);
  for (const property of ["radius", "font-size", "box-shadow", "transition"]) {
    assert.match(stylelintConfig, new RegExp(property), `${property} is not guarded`);
  }
  // Token files are the one documented exemption.
  assert.match(stylelintConfig, /packages\/\*\/src\/styles\/tokens\.css/);
  assert.match(stylelintConfig, /packages\/\*\/src\/styles\/theme\.css/);

  assert.equal(styleBaseline.schemaVersion, 1);
  assert.ok(styleBaseline.approved.reviewer);
  assert.match(styleBaseline.approved.issue, /^#\d+$/);
  assert.ok(styleBaseline.approved.reason.length >= 8);
  assert.deepEqual(styleBaseline.rules.sort(), [
    "color-no-hex",
    "declaration-property-unit-disallowed-list",
    "declaration-property-value-disallowed-list",
  ]);
  const counts = Object.values(styleBaseline.counts);
  assert.ok(counts.length > 0);
  for (const rules of counts) {
    for (const value of Object.values(rules)) assert.ok(Number.isInteger(value) && value > 0);
  }

  const lint = run(["scripts/ci/lint-styles.mjs", "--report-only"], {
    STYLE_GUARD_EVIDENCE_DIR: path.join(rootPath, "output/style-guard-test"),
  });
  assert.equal(lint.status, 0, `${lint.stdout}\n${lint.stderr}`);
  assert.match(lint.stdout, /0 new/);
});

test("style baseline cannot grow without an approval", () => {
  const grow = run(["scripts/ci/lint-styles.mjs", "--write-baseline"]);
  assert.notEqual(grow.status, 0);
  assert.match(`${grow.stdout}${grow.stderr}`, /requires an approval/);
});

test("a11y config covers the core surfaces and allowlists only owned findings", () => {
  assert.deepEqual(
    a11yConfig.surfaces.map((surface) => surface.id),
    ["button", "input", "select", "dialog", "tooltip", "file-upload", "file-thumbnail", "pdf-toolbar", "pdf-sidebar"],
  );
  assert.ok(a11yConfig.axe.tags.includes("wcag2a"));
  assert.ok(a11yConfig.axe.tags.includes("wcag21aa"));
  assert.deepEqual(a11yConfig.themes, ["light", "dark"]);
  for (const surface of a11yConfig.surfaces) {
    assert.ok(
      manifest.surfaces.some((candidate) => candidate.id === surface.id),
      `${surface.id} is not a parity surface`,
    );
    for (const check of surface.checks) {
      assert.ok(
        ["focusable", "focus-indicator", "accessible-name", "select-opens-on-arrow"].includes(check),
        `unknown keyboard check ${check}`,
      );
    }
  }
  for (const entry of a11yAllowlist.entries) {
    assert.ok(entry.surface);
    assert.ok(entry.rule || entry.check);
    assert.match(entry.issue, /^#\d+$/);
    assert.ok(entry.reason.length >= 20, "allowlist reasons must explain the finding");
  }
});

test("ui-parity workflow wires the PR subset, nightly matrix, artifacts and the strict switch", () => {
  for (const job of ["ui-contracts", "ui-parity-core", "ui-parity-nightly", "ui-a11y"]) {
    assert.match(workflow, new RegExp(`\\n  ${job}:`), `missing job ${job}`);
  }
  assert.match(workflow, /schedule:\s*\n(?:\s*#[^\n]*\n)*\s*- cron:/);
  assert.match(workflow, /--only "\$\{UI_PARITY_SUBSET\}"/);
  assert.match(workflow, /UI_PARITY_SUBSET: core/);
  assert.match(workflow, /--gate-mode "\$\{\{ inputs.strict && 'all' \|\| 'ready' \}\}"/);
  assert.match(workflow, /--baseline/);
  assert.match(workflow, /pnpm ui:baseline:verify/);
  assert.match(workflow, /pnpm lint:style/);
  assert.match(workflow, /pnpm test:a11y/);
  const uploads = workflow.match(/actions\/upload-artifact@v7/g) || [];
  assert.ok(uploads.length >= 4, "every job must upload its report");
  assert.match(workflow, /path: output\/ui-parity\//);
  assert.match(workflow, /path: output\/a11y\//);
});

test("gate docs and release checklist stay in sync with the implementation", () => {
  const checklist = read("docs/plan/release-checklist.md");
  for (const requirement of ["ui-parity-core", "pnpm ui:baseline:verify", "pnpm lint:style", "pnpm test:a11y"]) {
    assert.ok(checklist.includes(requirement), `release checklist lost ${requirement}`);
  }
  const gateDoc = read("docs/testing/ui-parity-ci.md");
  assert.ok(gateDoc.includes("topDiffRegions"));
  assert.ok(gateDoc.includes("gatePendingIssue"));
  assert.ok(gateDoc.includes("--gate-mode"));
  const docs = run(["scripts/ci/verify-ui-docs.mjs"], {
    UI_DOCS_EVIDENCE_DIR: path.join(rootPath, "output/acceptance/docs-contract/ui-docs-test"),
  });
  assert.equal(docs.status, 0, `${docs.stdout}\n${docs.stderr}`);
});
