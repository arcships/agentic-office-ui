// Issue #155 — Extend reference parity harness contract.
//
// These assertions pin the harness wiring that the later UI issues depend on:
// the shared fixture/manifest contract, the reference/vue surface coverage, the
// generated token report and its provenance, and the `pnpm ui:parity` entry
// point. They read source files so they run without a browser or a build; the
// capture itself runs through `pnpm ui:parity`.

import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { existsSync, readFileSync } from "node:fs";
import test from "node:test";

const ROOT = new URL("../../", import.meta.url);

function read(relativePath) {
  return readFileSync(new URL(relativePath, ROOT), "utf8");
}

function readJson(relativePath) {
  return JSON.parse(read(relativePath));
}

const manifest = readJson("tools/parity-fixtures/manifest.json");
const fixtures = readJson("tools/parity-fixtures/fixtures.json");
const sources = readJson("docs/design/extend-token-sources.json");
const tokens = readJson("docs/design/extend-tokens.json");
const packageJson = readJson("package.json");
const rootPackage = packageJson;

const EXTEND_COMMIT = "70fe70c6f1e1eee53954c6923f3d9f583557a57e";

test("manifest pins the reference commit and the four-viewport matrix", () => {
  assert.equal(manifest.referenceCommit, EXTEND_COMMIT);
  assert.equal(manifest.referenceUrl, "https://github.com/extend-hq/ui");
  assert.equal(manifest.deviceScaleFactor, 2);
  assert.deepEqual(
    manifest.viewports.map((viewport) => viewport.id),
    ["1440x900", "1280x720", "768x1024", "390x844"],
  );
  assert.deepEqual(manifest.themes, ["light", "dark"]);
  assert.equal(manifest.thresholds.geometryTolerancePx, 0);
  assert.equal(manifest.thresholds.maxChannelDelta, 2);
  assert.equal(manifest.thresholds.maxDiffPixelRatio, 0.001);
});

test("manifest declares an 8-state capable surface set with stable ids", () => {
  const ids = manifest.surfaces.map((surface) => surface.id);
  assert.deepEqual(
    ids,
    [
      "button",
      "input",
      "select",
      "dialog",
      "tooltip",
      "file-upload",
      "file-thumbnail",
      "pdf-toolbar",
      "pdf-sidebar",
      "docx-toolbar",
      "docx-sidebar",
      "pptx-toolbar",
      "pptx-sorter",
      "pptx-presenter",
      "document-sidebar",
      "bounding-box-citations",
      "layout-blocks",
      "e-signature",
      "document-splits",
      "xlsx-toolbar",
      "xlsx-sheet-tabs",
      "xlsx-status-bar",
    ],
  );
  assert.equal(new Set(ids).size, ids.length, "surface ids must be unique");

  const matrixStates = new Set();
  for (const surface of manifest.surfaces) {
    for (const state of surface.states) {
      matrixStates.add(state.matrixState);
      assert.ok(state.id, `${surface.id} state without id`);
      assert.ok(state.matrixState, `${surface.id}/${state.id} without matrixState`);
    }
  }
  // The issue matrix: default/hover/focus/active/disabled/loading/empty/error.
  assert.deepEqual(
    [...matrixStates].sort(),
    ["active", "default", "disabled", "empty", "error", "focus", "hover", "loading"],
  );

  for (const surface of manifest.surfaces.filter((entry) => entry.states.length > 0)) {
    assert.ok(surface.cellWidth > 0, `${surface.id} needs a cellWidth`);
    if (surface.referenceStatus === "unreachable") {
      assert.ok(surface.unreachableReason?.length > 20, `${surface.id} needs an unreachableReason`);
    } else {
      assert.ok(surface.referenceComponent, `${surface.id} needs a referenceComponent`);
    }
    if (surface.subjectStatus === "available") {
      assert.ok(surface.subjectComponent, `${surface.id} needs a subjectComponent`);
      if (surface.referenceStatus !== "unreachable") {
        assert.equal(surface.pendingIssue ?? null, null);
      } else {
        assert.ok(surface.pendingIssue, `${surface.id} must name the owning issue`);
      }
    } else {
      assert.equal(surface.subjectStatus, "missing");
      assert.ok(surface.pendingIssue, `${surface.id} must name the blocking issue`);
    }
  }

  const button = manifest.surfaces.find((surface) => surface.id === "button");
  assert.equal(button?.subjectStatus, "available");
  assert.equal(button?.subjectComponent, "packages/vue-ui/src/components/ui/OfficeButton.vue");
  assert.equal(button?.pendingIssue ?? null, null);
});

test("both harness apps render every manifest surface and honour the DOM contract", () => {
  const referenceApp = read("tools/extend-reference/src/App.tsx");
  const vueApp = read("tools/vue-reference/src/App.vue");
  const referenceContract = read("tools/extend-reference/src/parity/ParitySurface.tsx");
  const vueCell = read("tools/vue-reference/src/parity/ParityCell.vue");
  const vueSurface = read("tools/vue-reference/src/parity/ParitySurface.vue");

  for (const surface of manifest.surfaces.filter(
    (entry) => !["fixtures-only", "unreachable"].includes(entry.referenceStatus),
  )) {
    assert.match(
      referenceApp,
      new RegExp(`(^|\\n)\\s*"?${surface.id}"?\\s*:`),
      `reference app misses ${surface.id}`,
    );
    if (surface.subjectStatus === "available") {
      assert.match(vueApp, new RegExp(`'${surface.id}'|"${surface.id}"`), `vue app misses ${surface.id}`);
    } else {
      assert.ok(
        vueApp.includes("MissingSurface"),
        `vue app needs the missing-subject placeholder for ${surface.id}`,
      );
    }
  }

  for (const attribute of [
    "data-parity-surface",
    "data-parity-theme",
    "data-parity-cell",
    "data-parity-host",
    "data-parity-matrix-state",
  ]) {
    assert.ok(
      referenceContract.includes(attribute) || referenceApp.includes(attribute),
      `reference app does not emit ${attribute}`,
    );
    assert.ok(
      [vueCell, vueSurface, vueApp].some((source) => source.includes(attribute)),
      `vue app does not emit ${attribute}`,
    );
  }
});

test("shared fixtures cover the document formats and block data from the issue", () => {
  const documentIds = fixtures.documents.map((document) => document.id);
  for (const id of ["pdf", "docx", "xlsx", "pptx", "png"]) {
    assert.ok(documentIds.includes(id), `missing fixture document: ${id}`);
  }
  assert.equal(fixtures.reference.commit, EXTEND_COMMIT);

  for (const block of ["citations", "layoutBlocks", "signature", "splits", "fileSystem"]) {
    assert.ok(fixtures.blocks[block], `missing block fixture: ${block}`);
  }

  for (const key of ["title", "description", "browseLabel", "draggingLabel", "rejection"]) {
    assert.equal(typeof fixtures.copy.fileUpload[key], "string", `missing upload copy: ${key}`);
  }
  assert.ok(fixtures.fileUpload.selectedFile.name.endsWith(".pdf"));
  assert.ok(fixtures.fileUpload.rejectedFile.name.endsWith(".txt"));
});

test("shell asks for the full state coverage on reference surfaces", () => {
  const referenceApp = read("tools/extend-reference/src/App.tsx");
  for (const surfaceId of [
    "ButtonSurface",
    "InputSurface",
    "SelectSurface",
    "DialogSurface",
    "TooltipSurface",
    "FileUploadSurface",
    "FileThumbnailSurface",
    "PdfToolbarSurface",
    "DocxToolbarSurface",
    "PptxToolbarSurface",
    "PptxSorterSurface",
    "DocumentSidebarSurface",
  ]) {
    assert.ok(referenceApp.includes(surfaceId), `reference app does not mount ${surfaceId}`);
  }
  const buttonSurface = read("tools/extend-reference/src/surfaces/ButtonSurface.tsx");
  for (const state of ["default", "hover", "focus", "active", "disabled", "loading"]) {
    assert.ok(buttonSurface.includes(`state: "${state}"`), `button surface misses ${state}`);
  }
  const uploadSurface = read("tools/extend-reference/src/surfaces/FileUploadSurface.tsx");
  for (const state of ["default", "hover", "focus", "dragging", "rejected", "file-selected"]) {
    assert.ok(uploadSurface.includes(`state: "${state}"`), `file upload surface misses ${state}`);
  }
});

test("vendored Extend CSS keeps the pinned upstream body", () => {
  const css = read("tools/extend-reference/src/styles/extend-globals.css");
  const marker = sources.bodyMarker;
  const body = css.slice(css.indexOf(marker));
  const digest = createHash("sha256").update(body).digest("hex");
  assert.equal(digest, sources.bodySha256, "vendored globals.css body changed");
  assert.equal(
    sources.globalsSha256,
    "f929d2c1de31312f6fc53bd9a7fff9f6a3feff46d4d5fd65805fea2deddbc759",
  );
  assert.equal(sources.commit, EXTEND_COMMIT);
});

test("vendored components carry the MIT attribution header", () => {
  const vendored = [
    "tools/extend-reference/src/vendor/ui/button.tsx",
    "tools/extend-reference/src/vendor/ui/spinner.tsx",
    "tools/extend-reference/src/vendor/ui/card.tsx",
    "tools/extend-reference/src/vendor/ui/separator.tsx",
    "tools/extend-reference/src/vendor/ui/select.tsx",
    "tools/extend-reference/src/vendor/ui/input.tsx",
    "tools/extend-reference/src/vendor/ui/dialog.tsx",
    "tools/extend-reference/src/vendor/ui/tooltip.tsx",
    "tools/extend-reference/src/vendor/extend/file-upload.tsx",
    "tools/extend-reference/src/vendor/extend/file-thumbnail.tsx",
    "tools/extend-reference/src/vendor/extend/document-viewer-sidebar.tsx",
    "tools/extend-reference/src/vendor/extend/docx-toolbar.tsx",
    "tools/extend-reference/src/vendor/extend/pdf-toolbar.tsx",
    "tools/extend-reference/src/lib/cn.ts",
    "tools/extend-reference/src/styles/extend-globals.css",
    "tools/extend-reference/src/styles/extend-shadcn-tailwind.css",
  ];
  for (const file of vendored) {
    const header = read(file).split(/\r?\n/).slice(0, 15).join("\n");
    assert.match(header, /extend-hq\/ui @? ?70fe70c6f1e1eee53954c6923f3d9f583557a57e/, `${file}: commit`);
    assert.match(header, /MIT/, `${file}: license id`);
    assert.match(header, /Copyright \(c\) 2026 CrowdView Inc, dba Extend/, `${file}: copyright`);
  }
  const notice = read("tools/extend-reference/src/vendor/NOTICE.md");
  for (const marker of ["MIT License", "CrowdView Inc, dba Extend", "docs/upstream/licenses/extend-ui-MIT.txt"]) {
    assert.ok(notice.includes(marker), `NOTICE.md misses ${marker}`);
  }
  assert.ok(existsSync(new URL("tools/extend-reference/UPSTREAM.md", ROOT)));
  assert.ok(existsSync(new URL("tools/extend-reference/README.md", ROOT)));
});

test("token report covers colors, radius, type, spacing, shadow, ring and motion", () => {
  assert.equal(tokens.provenance.commit, EXTEND_COMMIT);
  assert.deepEqual(Object.keys(tokens.colors).sort(), ["dark", "light"]);
  assert.equal(Object.keys(tokens.colors.light).length, 46);
  assert.equal(tokens.radius.basePx, "10px");
  assert.equal(tokens.radius.scale.lg.px, "10px");
  assert.equal(tokens.radius.scale.xl.px, "14px");
  assert.equal(tokens.spacing.base.px, 4);
  assert.equal(tokens.typography.scale.sm.fontSizePx, 14);
  assert.equal(tokens.typography.scale.sm.lineHeightPx, 20);
  assert.equal(tokens.ring.focusRing.width, "2px");
  assert.equal(tokens.ring.focusRing.offset, "1px");
  assert.ok(tokens.shadows.xs?.raw.includes("0 1px 2px 0 rgb(0 0 0 / 0.05)"));
  assert.equal(tokens.motion.defaultDuration, "150ms");
  assert.ok(tokens.motion.durations.some((duration) => duration.ms === 200));
  assert.ok(tokens.motion.easings.length > 0);

  const background = tokens.colors.light.background;
  assert.equal(background.hex, "#ffffff");
  const border = tokens.colors.light.border;
  assert.equal(border.hex, "#000000");
  assert.equal(border.alpha, 0.08);
});

test("token mapping compares the pinned palette with the #156 office layer", () => {
  assert.equal(tokens.mapping.length, 46);
  const byName = new Map(tokens.mapping.map((entry) => [entry.extendToken, entry]));
  assert.equal(byName.get("--background").officeToken, "--office-background");
  assert.equal(byName.get("--background").officeTokenStatus, "present");
  assert.equal(byName.get("--border").officeTokenStatus, "present");
  assert.equal(byName.get("--border").lightDelta.maxChannelDelta, 0);
  assert.equal(byName.get("--border").lightDelta.alphaDelta, 0);
  // #156 reopened: every pinned semantic colour, including the --code* family,
  // is represented by an --office-* token with no channel delta.
  for (const name of ["--code", "--code-foreground", "--code-highlight", "--code-number"]) {
    const entry = byName.get(name);
    assert.ok(entry, `${name} must be in the mapping`);
    assert.equal(entry.officeTokenStatus, "present", `${name} must be declared`);
    assert.equal(entry.lightDelta.matches, true, `${name} light`);
    assert.equal(entry.darkDelta.matches, true, `${name} dark`);
  }
  assert.ok(
    tokens.mapping.every((entry) => entry.officeTokenStatus === "present"),
    "every Extend semantic colour must be mapped",
  );

  assert.ok(tokens.tokenLayer.entries.length > 20);
  assert.deepEqual(tokens.tokenLayer.summary, {
    total: 27,
    matched: 27,
    mismatched: 0,
    missing: 0,
  });
  const padding = tokens.tokenLayer.entries.find((entry) =>
    entry.token.startsWith("control-padding-inline"),
  );
  // The token is the full 12px inset (11px padding + 1px border); consumers
  // subtract --office-border-width when they paint the border.
  assert.equal(padding.extend.px, 12);
  assert.equal(padding.office.px, 12);
  assert.equal(padding.match, true);
  const radius = tokens.tokenLayer.entries.find((entry) => entry.token === "--radius");
  assert.equal(radius.match, true);
});

test("pnpm ui:parity wires the mjs entry point to the Playwright helper", () => {
  assert.equal(rootPackage.scripts["ui:parity"], "node scripts/ci/pixel-parity.mjs");
  assert.equal(rootPackage.scripts["ui:parity:report"], "node scripts/ci/pixel-parity.mjs --report-only");
  assert.equal(rootPackage.scripts["ui:tokens"], "node scripts/ci/extract-extend-tokens.mjs");
  assert.equal(
    rootPackage.scripts["ui:tokens:check"],
    "node scripts/ci/extract-extend-tokens.mjs --check",
  );
  assert.equal(
    rootPackage.scripts["ui:parity:self-check"],
    "node scripts/ci/pixel-parity.mjs --self-check --out output/ui-parity-selfcheck",
  );
  assert.ok(
    read("pnpm-workspace.yaml").includes('"tools/*"'),
    "tools/* must be a pnpm workspace glob",
  );

  const orchestrator = read("scripts/ci/pixel-parity.mjs");
  assert.ok(orchestrator.includes("scripts/ci/pixel-parity.py"));
  assert.ok(orchestrator.includes("@aou/extend-reference"));
  assert.ok(orchestrator.includes("@aou/vue-reference"));
  assert.ok(orchestrator.includes("--report-only"));

  const helper = read("scripts/ci/pixel-parity.py");
  assert.ok(helper.includes("device_scale_factor"));
  assert.ok(helper.includes("document.fonts.status"));
  assert.ok(helper.includes("animation: none !important"));
  assert.ok(helper.includes("maxChannelDelta"));
  assert.ok(helper.includes("maxDiffPixelRatio"));
  assert.ok(helper.includes("maxContiguousDiffPx") || helper.includes("maxContiguousRunPx"));
  assert.ok(helper.includes("heatmap"));
  assert.ok(helper.includes("--self-check"), "the helper needs a determinism self-check mode");
  // Lazy offscreen thumbnails must be loaded and decoded before a capture, or
  // only one side of the matrix renders them and the case looks flaky.
  assert.ok(helper.includes("def settle_images(page)"), "capture must settle lazy images");
  assert.ok(helper.includes("settle_images(page)"), "capture_side must call settle_images");
  assert.ok(helper.includes("image.decode()"), "settle_images must wait for decode");

  const referenceVite = read("tools/extend-reference/vite.config.ts");
  assert.ok(referenceVite.includes("@tailwindcss/vite"));
  const vueVite = read("tools/vue-reference/vite.config.ts");
  assert.ok(vueVite.includes("theme.css"), "the subject must consume the #156 theme layer");
});

test("harness docs and CI artifact wiring exist", () => {
  const doc = read("docs/design/extend-parity-harness.md");
  for (const snippet of [
    "pnpm ui:parity",
    "pnpm ui:parity:self-check",
    "docs/design/extend-tokens.json",
    "1440x900",
    "device_scale_factor",
    "ui-parity-report",
    "Button",
    "FileUpload",
    "PDF",
  ]) {
    assert.ok(doc.includes(snippet), `harness doc misses ${snippet}`);
  }
  const workflow = read(".github/workflows/ui-parity.yml");
  assert.ok(workflow.includes("pnpm ui:parity"));
  // #165 split the single `ui-parity-report` artifact into the PR core subset
  // and the nightly full matrix; both still carry report.json/report.html.
  assert.ok(workflow.includes("ui-parity-core-"));
  assert.ok(workflow.includes("ui-parity-full-"));
  assert.ok(workflow.includes("actions/upload-artifact"));
  assert.ok(workflow.includes("playwright install"));
  assert.ok(workflow.includes("ui:parity:self-check"));
});
