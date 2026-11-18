#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
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
// A blocked step (exit code 2) is never a pass. `allowBlocked` only marks a step
// as legitimately blockable in this environment; the suite still reports
// INCOMPLETE and exits non-zero. CI_ALLOW_BLOCKED=1 downgrades that to a warning
// for local runs that intentionally lack the upstream checkout.
const allowBlocked = process.env.CI_ALLOW_BLOCKED === "1";

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
      "tests/unit/viewer-keys.test.mjs",
      "tests/unit/runtime-demo-limits.test.mjs",
      "tests/unit/office-interaction.test.mjs",
      "tests/unit/i18n.test.mjs",
      "tests/unit/image-budget.test.mjs",
      "tests/unit/docx-core.test.mjs",
      "tests/unit/docx-search.test.mjs",
      "tests/unit/docx-replace.test.mjs",
      "tests/unit/docx-fixture-fidelity.test.mjs",
      "tests/unit/docx-form-field-required.test.mjs",
      "tests/unit/docx-model-preservation.test.mjs",
      "tests/unit/docx-export-prepare-model.test.mjs",
      "tests/unit/docx-notes.test.mjs",
      "tests/unit/docx-page-appearance.test.mjs",
      "tests/unit/docx-body-measurement.test.mjs",
      "tests/unit/docx-tracked-changes.test.mjs",
      "tests/unit/docx-revision-display.test.mjs",
      "tests/unit/docx-typography.test.mjs",
      "tests/unit/docx-paragraph-spacing.test.mjs",
      "tests/unit/docx-line-metrics.test.mjs",
      "tests/unit/docx-pagination-fidelity.test.mjs",
      "tests/unit/docx-pagination-completion.test.mjs",
      "tests/unit/docx-pagination-measure-loop.test.mjs",
      "tests/unit/docx-section-break.test.mjs",
      "tests/unit/docx-section-page-numbers.test.mjs",
      "tests/unit/docx-page-setup.test.mjs",
      "tests/unit/docx-floating-image.test.mjs",
      "tests/unit/docx-image-layer-order.test.mjs",
      "tests/unit/docx-reference-adapter.test.mjs",
      "tests/unit/docx-document-blocks.test.mjs",
      "tests/unit/docx-table-selection.test.mjs",
      "tests/unit/docx-history-cache.test.mjs",
      "tests/unit/docx-edit-structural-sharing.test.mjs",
      "tests/unit/docx-pagination-runtime.test.mjs",
      "tests/unit/docx-thumbnail-signature.test.mjs",
      "tests/unit/docx-clipboard.test.mjs",
      "tests/unit/docx-comments.test.mjs",
      "tests/unit/docx-bookmarks.test.mjs",
      "tests/unit/docx-svg-pictures.test.mjs",
      "tests/unit/docx-table-floating.test.mjs",
      "tests/unit/docx-drawing-objects.test.mjs",
      "tests/unit/xlsx-core.test.mjs",
      "tests/unit/xlsx-print.test.mjs",
      "tests/unit/xlsx-worker.test.mjs",
      "tests/unit/xlsx-anchor-geometry.test.mjs",
      "tests/unit/xlsx-chart-data.test.mjs",
      "tests/unit/xlsx-chart-type-edit.test.mjs",
      "tests/unit/xlsx-drawing-fidelity.test.mjs",
      "tests/unit/xlsx-typed-dates.test.mjs",
      "tests/unit/xlsx-reference-adapter.test.mjs",
      "tests/unit/xlsx-sheet-rename.test.mjs",
      "tests/unit/xlsx-sheet-order.test.mjs",
      "tests/unit/xlsx-sheet-visibility.test.mjs",
      "tests/unit/xlsx-authoring.test.mjs",
      "tests/unit/xlsx-structure.test.mjs",
      "packages/vue-xlsx/test/lazy-loading.mjs",
      "packages/vue-xlsx/test/chart-date-format.mjs",
      "tests/unit/pdf-source.test.mjs",
      "tests/unit/pdf-render-runtime.test.mjs",
      "tests/unit/pdf-page-viewport.test.mjs",
      "tests/unit/pdf-bitmap-cache.test.mjs",
      "tests/unit/pdf-page-raster.test.mjs",
      "tests/unit/pdf-reference-adapter.test.mjs",
      "tests/unit/pptx-core.test.mjs",
      "tests/unit/pptx-media-transitions.test.mjs",
      "tests/unit/pptx-browser-session.test.mjs",
      "tests/unit/pptx-playback-controller.test.mjs",
      "tests/unit/pptx-morph-geometry.test.mjs",
      "tests/unit/pptx-track-player.test.mjs",
      "tests/unit/pptx-baseline-contract.test.mjs",
      "tests/unit/pptx-reference-adapter.test.mjs",
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
    command("xlsx-export-openpyxl", process.execPath, [
      "scripts/ci/run-python.mjs",
      "tests/blackbox/xlsx_export_openpyxl.py",
    ], {
      env: { RUN_PYTHON_MODULES: "openpyxl", NODE: process.execPath },
    }),
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
      "tests/component/pdf-page-lifecycle.test.mjs",
      "tests/component/i18n-source.test.mjs",
      "tests/component/i18n-runtime.test.mjs",
      "tests/component/office-selection-ui.test.mjs",
      "tests/component/office-reference-surfaces.test.mjs",
      "tests/component/docx-reference-surfaces.test.mjs",
      "tests/component/pptx-preview.test.mjs",
      "tests/component/xlsx-race.test.mjs",
      "tests/component/xlsx-export.test.mjs",
      "tests/component/xlsx-deferred-load.test.mjs",
      "tests/component/docx-render-parity.test.mjs",
      "tests/component/docx-paragraph-formatting.test.mjs",
      "tests/component/docx-page-overflow.test.mjs",
      "tests/component/docx-page-clipped-measurement.test.mjs",
      "tests/component/docx-page-appearance.test.mjs",
      "tests/component/docx-floating-image-layout.test.mjs",
      "tests/component/docx-edit-interactions.test.mjs",
      "tests/component/docx-clipboard-rich.test.mjs",
      "tests/component/docx-link-safety.test.mjs",
      "tests/component/docx-toolbar-state.test.mjs",
      "tests/component/docx-table-editing.test.mjs",
      "tests/component/docx-command-wiring.test.mjs",
      "tests/component/docx-edit-model-preservation.test.mjs",
      "tests/component/docx-selection-input.test.mjs",
      "tests/component/docx-thumbnail-panel.test.mjs",
      "tests/component/docx-column-pagination.test.mjs",
      "tests/component/docx-multipage-shell.test.mjs",
      "tests/component/docx-page-setup.test.mjs",
      "tests/component/docx-section-page-numbers.test.mjs",
      "tests/component/docx-page-thumbnails-cache.test.mjs",
      "tests/component/docx-thumbnail-raster.test.mjs",
      "tests/component/docx-product-interactions.test.mjs",
      "tests/component/docx-search-refresh.test.mjs",
      "tests/component/docx-search-replace.test.mjs",
      "tests/component/docx-annotation-gutter.test.mjs",
      "tests/component/docx-fixture-fidelity.test.mjs",
      "tests/component/docx-comment-ranges.test.mjs",
      "tests/component/docx-comments-crud.test.mjs",
      "tests/component/docx-bookmarks.test.mjs",
      "tests/component/docx-table-presentation.test.mjs",
      "tests/component/docx-file-handling.test.mjs",
      "tests/component/docx-export-bytes.test.mjs",
      "tests/component/docx-tracked-changes.test.mjs",
      "tests/component/docx-revision-display.test.mjs",
      "tests/component/docx-context-menu-form-fields.test.mjs",
      "tests/component/docx-host-render-hooks.test.mjs",
      "tests/component/docx-image-layer-order.test.mjs",
      "tests/component/docx-editor-affordances.test.mjs",
      "tests/component/docx-table-floating.test.mjs",
      "tests/component/docx-drawing-objects.test.mjs",
      "tests/component/docx-block-highlights.test.mjs",
      "tests/component/docx-page-budget.test.mjs",
      "tests/component/pdf-surface-lazy-render.test.mjs",
      "tests/component/docx-notes.test.mjs",
      "tests/component/docx-history-budget.test.mjs",
      "tests/component/docx-edit-clone-sharing.test.mjs",
      "tests/component/docx-text-commit-burst.test.mjs",
      "tests/component/xlsx-history-budget.test.mjs",
      "tests/component/xlsx-replace.test.mjs",
      "tests/component/xlsx-number-formats.test.mjs",
      "tests/component/xlsx-number-locale.test.mjs",
      "tests/component/xlsx-number-format-undo.test.mjs",
      "tests/component/xlsx-worker-cell-format.test.mjs",
      "tests/component/xlsx-sheet-sync.test.mjs",
      "tests/component/xlsx-sheet-rename.test.mjs",
      "tests/component/xlsx-sheet-order.test.mjs",
      "tests/component/xlsx-sheet-tabs.test.mjs",
      "tests/component/xlsx-viewer-tab-surface.test.mjs",
      "tests/component/xlsx-sheet-visibility.test.mjs",
      "tests/component/xlsx-drawing-fidelity.test.mjs",
      "tests/component/xlsx-calculation-safety.test.mjs",
      "tests/component/xlsx-grid-scroll.test.mjs",
      "tests/component/xlsx-cell-editing.test.mjs",
      "tests/component/xlsx-clipboard-search.test.mjs",
      "tests/component/xlsx-conditional-format.test.mjs",
      "tests/component/xlsx-grid-freeze-panes.test.mjs",
      "tests/component/xlsx-grid-fidelity.test.mjs",
      "tests/component/xlsx-render-perf.test.mjs",
      "tests/component/xlsx-data-validation.test.mjs",
      "tests/component/xlsx-chart-rendering.test.mjs",
      "tests/component/xlsx-chart-axis.test.mjs",
      "tests/component/xlsx-chart-type-menu.test.mjs",
      "tests/component/xlsx-authoring-ui.test.mjs",
      "tests/component/xlsx-print.test.mjs",
      "tests/component/xlsx-structure.test.mjs",
      "tests/component/xlsx-formatting-ui.test.mjs",
      "tests/component/office-a11y.test.mjs",
      "tests/component/viewer-keyboard.test.mjs",
      "tests/component/touch-interactions.test.mjs",
      "tests/component/image-decode-state.test.mjs",
      "tests/component/docx-svg-pictures.test.mjs",
      "tests/component/viewer-print-theme.test.mjs",
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
    command("formal-touch-mobile", process.execPath, [
      "scripts/ci/run-python.mjs",
      "tests/blackbox/touch_mobile.py",
    ], {
      env: {
        BLACKBOX_EVIDENCE_DIR: path.join(suiteDir, "formal-touch-mobile"),
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
    command("formal-docx-export-bytes", process.execPath, [
      "scripts/ci/run-python.mjs",
      "tests/blackbox/docx_export_bytes.py",
    ], {
      env: {
        BLACKBOX_EVIDENCE_DIR: path.join(suiteDir, "formal-docx-export-bytes"),
      },
    }),
    // The only automated coverage for the DOCX comment/revision gutter.
    command("formal-docx-annotations", process.execPath, [
      "scripts/ci/run-python.mjs",
      "tests/blackbox/docx_annotations.py",
    ], {
      env: {
        BLACKBOX_EVIDENCE_DIR: path.join(suiteDir, "formal-docx-annotations"),
      },
    }),
    command("formal-ux-parity", process.execPath, [
      "scripts/ci/run-python.mjs",
      "tests/blackbox/ux_parity.py",
    ], {
      env: {
        BLACKBOX_EVIDENCE_DIR: path.join(suiteDir, "formal-ux-parity"),
      },
      // Blockable because the upstream reference worktree is an opt-in checkout,
      // but a BLOCKED result makes the suite INCOMPLETE (issue #57), not PASS.
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
  pdf: [
    ...buildStep(),
    // Opens every fixture with a real parser: page counts, /Rotate metadata,
    // image-only pages, encryption and CJK extraction before any browser runs.
    command("material-verification", process.execPath, [
      "scripts/ci/run-python.mjs",
      "scripts/verify_test_materials.py",
    ], {
      env: { RUN_PYTHON_MODULES: "pypdf,docx,openpyxl,PIL" },
    }),
    command("fixture-manifest", process.execPath, [
      "scripts/ci/verify-fixtures.mjs",
    ], {
      env: {
        FIXTURE_EVIDENCE_DIR: path.join(suiteDir, "fixture-manifest"),
      },
    }),
    command("formal-preview-routes", process.execPath, [
      "scripts/ci/run-python.mjs",
      "tests/blackbox/routes_smoke.py",
    ], {
      env: {
        BLACKBOX_EVIDENCE_DIR: path.join(suiteDir, "formal-preview-routes"),
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
    command("formal-fidelity-content", process.execPath, [
      "scripts/ci/run-python.mjs",
      "tests/blackbox/fidelity_content.py",
    ], {
      env: {
        BLACKBOX_EVIDENCE_DIR: path.join(suiteDir, "formal-fidelity-content"),
      },
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
// The probe exit code lets scripts/ci/verify-failure-propagation.mjs exercise the
// BLOCKED path (exit code 2) without a browser or an upstream checkout.
const failureProbeExit = Number.parseInt(process.env.CI_FAILURE_PROBE_EXIT || "97", 10);
const probeBlocked = process.env.CI_PROBE_BLOCKED === "1";
const results = [];
const blockedSteps = [];
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
    status = failureProbeExit;
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
  const blockable = step.allowBlocked || (probeMatches && probeBlocked);
  const blocked = blockable && status === 2;
  const stepStatus = blocked ? "BLOCKED" : status === 0 ? "PASS" : "FAIL";

  // When a Playwright blackbox test fails with no stdout, try reading its
  // summary.json from the evidence directory so the failure reason shows up
  // in CI logs.
  if (!blocked && status !== 0 && !output.trim()) {
    const summaryPath = path.join(suiteDir, step.id, "summary.json");
    try {
      const summary = JSON.parse(readFileSync(summaryPath, "utf8"));
      const compact = JSON.stringify(summary, null, 2).slice(0, 8000);
      output = `${compact}\n`;
      process.stdout.write(output);
    } catch { /* no summary available */ }
  }

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
  if (blocked) blockedSteps.push(step.id);
  else if (status !== 0) failed = true;
}

const result = failed
  ? "FAIL"
  : blockedSteps.length > 0 && !allowBlocked
    ? "INCOMPLETE"
    : "PASS";
const summary = {
  suite: suiteName,
  result,
  runId,
  commit,
  evidenceRoot,
  allowBlocked,
  blockedSteps,
  results,
};
writeFileSync(
  path.join(suiteDir, "summary.json"),
  `${JSON.stringify(summary, null, 2)}\n`,
);

if (blockedSteps.length > 0) {
  const detail = allowBlocked
    ? "counted as a warning because CI_ALLOW_BLOCKED=1"
    : "not counted as a pass; set CI_ALLOW_BLOCKED=1 only for local runs without the upstream checkout";
  console.warn(`\n[${suiteName}] BLOCKED steps (${detail}): ${blockedSteps.join(", ")}`);
}
console.log(`\n[${suiteName}] ${result}: ${suiteDir}`);
if (failed) process.exit(1);
process.exit(result === "INCOMPLETE" ? 2 : 0);
