import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import {
  DEFAULT_ICON_PACKAGE,
  PARITY_ICON_PACKAGE,
  collectImports,
  iconImportsIn,
  iconPackageName,
  isThirdPartyIconLibrary,
  validateIconAllowlist,
} from "../../scripts/ci/icon-allowlist.mjs";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const verifier = path.join(root, "scripts/ci/verify-icons.mjs");
const ledgerPath = path.join(root, "tests/icons/allowlist.json");

const SURFACE = "pdf-toolbar";
const UPSTREAM_SNAPSHOT = "tools/extend-reference/src/vendor/extend/pdf-toolbar.tsx";
const ALLOWED_TOOLBAR = "packages/vue-pdf/src/components/PdfViewerToolbar.vue";
const ADAPTER = "packages/vue-ui/src/icons.ts";

const HUGE_SOURCE = `<script setup lang="ts">
import { CirclePlusIcon } from "@hugeicons/core-free-icons"
import { HugeIcon } from "@arcships/vue-ui"
</script>
<template><HugeIcon :icon="CirclePlusIcon" /></template>
`;

const LUCIDE_SOURCE = `import { Search } from "@lucide/vue"

export const ICONS = Object.freeze({ search: Search })
`;

function write(file, content) {
  mkdirSync(path.dirname(file), { recursive: true });
  writeFileSync(file, content);
}

function parityEntry(overrides = {}) {
  return {
    file: ALLOWED_TOOLBAR,
    iconPackages: [PARITY_ICON_PACKAGE],
    track: "office-chrome",
    surface: SURFACE,
    upstreamSource: UPSTREAM_SNAPSHOT,
    issue: "#162",
    reason: "fixture parity exception",
    reviewBy: "2027-05-21",
    ...overrides,
  };
}

function adapterEntry(overrides = {}) {
  return {
    file: ADAPTER,
    iconPackages: [DEFAULT_ICON_PACKAGE],
    track: "default-adapter",
    issue: "#180",
    reason: "fixture registry adapter",
    reviewBy: "2027-05-21",
    ...overrides,
  };
}

function policyDoc(entries = [parityEntry(), adapterEntry()]) {
  const rows = entries.map((entry) => `| \`${entry.file}\` | fixture row |`).join("\n");
  return [
    "# Fixture icon policy",
    "",
    "`OfficeIcon` renders @lucide/vue; `HugeIcon` renders @hugeicons/core-free-icons.",
    "",
    "| File | Meaning |",
    "|---|---|",
    rows,
    "",
  ].join("\n");
}

function writeLedger(fixture, entries) {
  write(
    path.join(fixture, "tests/icons/allowlist.json"),
    JSON.stringify({ schemaVersion: 1, entries }, null, 2),
  );
  write(path.join(fixture, "docs/design/icon-policy.md"), policyDoc(entries));
}

/**
 * Minimal workspace that mirrors the real one: one parity package importing
 * Hugeicons data, the vue-ui adapter importing Lucide, the vendored upstream
 * snapshot the ledger points at, and the parity manifest that names a surface.
 */
function createFixture({
  entries = [parityEntry(), adapterEntry()],
  doc = policyDoc(entries),
  files = {},
  packages = {},
} = {}) {
  const fixture = mkdtempSync(path.join(tmpdir(), "icon-policy-"));
  write(
    path.join(fixture, "packages/vue-pdf/package.json"),
    JSON.stringify({ name: "@fixture/vue-pdf", dependencies: { [PARITY_ICON_PACKAGE]: "4.4.0" } }, null, 2),
  );
  write(path.join(fixture, ALLOWED_TOOLBAR), HUGE_SOURCE);
  write(
    path.join(fixture, "packages/vue-ui/package.json"),
    JSON.stringify({ name: "@fixture/vue-ui", dependencies: { [DEFAULT_ICON_PACKAGE]: "1.76.0" } }, null, 2),
  );
  write(path.join(fixture, ADAPTER), LUCIDE_SOURCE);
  write(path.join(fixture, UPSTREAM_SNAPSHOT), "// pinned upstream snapshot\n");
  write(
    path.join(fixture, "tools/parity-fixtures/manifest.json"),
    JSON.stringify({ surfaces: [{ id: SURFACE }] }, null, 2),
  );
  write(path.join(fixture, "docs/design/icon-policy.md"), doc);
  write(
    path.join(fixture, "tests/icons/allowlist.json"),
    JSON.stringify({ schemaVersion: 1, entries }, null, 2),
  );
  for (const [file, content] of Object.entries(files)) write(path.join(fixture, file), content);
  for (const [file, content] of Object.entries(packages)) write(path.join(fixture, file), content);
  return fixture;
}

function runVerifier(cwd, env = {}) {
  return spawnSync(process.execPath, [verifier], {
    cwd: root,
    encoding: "utf8",
    env: {
      ...process.env,
      ICON_GUARD_ROOT: cwd,
      ICON_GUARD_EVIDENCE_DIR: path.join(cwd, "output/icons"),
      ...env,
    },
  });
}

function withFixture(options, run) {
  const fixture = createFixture(options);
  try {
    return run(fixture);
  } finally {
    rmSync(fixture, { recursive: true, force: true });
  }
}

function readSummary(fixture) {
  return JSON.parse(readFileSync(path.join(fixture, "output/icons/summary.json"), "utf8"));
}

test("icon helpers reduce specifiers to packages and spot third icon systems", () => {
  assert.equal(iconPackageName("@hugeicons/core-free-icons"), "@hugeicons/core-free-icons");
  assert.equal(iconPackageName("@hugeicons/core-free-icons/dist/foo"), "@hugeicons/core-free-icons");
  assert.equal(iconPackageName("lucide-vue-next/dist/icon"), "lucide-vue-next");
  assert.equal(iconPackageName("./local"), null);
  assert.equal(iconPackageName("/absolute"), null);

  assert.equal(isThirdPartyIconLibrary(DEFAULT_ICON_PACKAGE), false);
  assert.equal(isThirdPartyIconLibrary(PARITY_ICON_PACKAGE), false);
  assert.equal(isThirdPartyIconLibrary(PARITY_ICON_PACKAGE, []), false);
  assert.equal(isThirdPartyIconLibrary("@hugeicons/react"), true);
  assert.equal(isThirdPartyIconLibrary("@tabler/icons-vue"), true);
  assert.equal(isThirdPartyIconLibrary("@fortawesome/free-solid-svg-icons"), true);
  assert.equal(isThirdPartyIconLibrary("bootstrap-icons"), true);
  assert.equal(isThirdPartyIconLibrary("lucide-react"), true);
  assert.equal(isThirdPartyIconLibrary("@scope/icons-vue"), true);
  assert.equal(isThirdPartyIconLibrary("reka-ui"), false);
  assert.equal(isThirdPartyIconLibrary("@icon-name-generator/core", ["@icon-name-generator/core"]), false);

  const imports = collectImports(`${HUGE_SOURCE}${LUCIDE_SOURCE}`);
  assert.deepEqual(
    imports.filter((entry) => [DEFAULT_ICON_PACKAGE, PARITY_ICON_PACKAGE].includes(entry.package))
      .map((entry) => entry.package),
    [PARITY_ICON_PACKAGE, DEFAULT_ICON_PACKAGE],
  );
  assert.deepEqual(iconImportsIn(HUGE_SOURCE).map((entry) => entry.line), [2]);
});

test("icon allowlist validation rejects globs, unknown tracks and missing owners", () => {
  const fixture = createFixture();
  try {
    const globbed = validateIconAllowlist(
      { schemaVersion: 1, entries: [parityEntry({ file: "packages/vue-pdf/src/components/*.vue" })] },
      { root: fixture, knownSurfaces: [SURFACE], today: "2026-11-21" },
    );
    assert.deepEqual(globbed.failures.map((failure) => failure.code), ["ICON_ALLOWLIST_FILE"]);

    const unknownTrack = validateIconAllowlist(
      { schemaVersion: 1, entries: [parityEntry({ track: "whatever" })] },
      { root: fixture, knownSurfaces: [SURFACE], today: "2026-11-21" },
    );
    assert.ok(unknownTrack.failures.some((failure) => failure.code === "ICON_ALLOWLIST_TRACK"));

    const missingOwner = validateIconAllowlist(
      { schemaVersion: 1, entries: [parityEntry({ issue: "180", reason: "" })] },
      { root: fixture, knownSurfaces: [SURFACE], today: "2026-11-21" },
    );
    assert.ok(missingOwner.failures.every((failure) => failure.code === "ICON_ALLOWLIST_OWNER"));
    assert.equal(missingOwner.failures.length, 2);
  } finally {
    rmSync(fixture, { recursive: true, force: true });
  }
});

test("icon guard accepts the allowlisted dual track and reports it", () => {
  withFixture({}, (fixture) => {
    const result = runVerifier(fixture);
    assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);
    const summary = readSummary(fixture);
    assert.equal(summary.status, "PASS");
    assert.deepEqual(
      summary.allowedFiles.map((entry) => [entry.file, entry.track, entry.iconPackages]),
      [
        [ALLOWED_TOOLBAR, "office-chrome", [PARITY_ICON_PACKAGE]],
        [ADAPTER, "default-adapter", [DEFAULT_ICON_PACKAGE]],
      ],
    );
    assert.equal(summary.scannedSourceFiles, 2);
    assert.deepEqual(summary.declaredDependencies, [
      { file: "packages/vue-pdf/package.json", package: PARITY_ICON_PACKAGE },
      { file: "packages/vue-ui/package.json", package: DEFAULT_ICON_PACKAGE },
    ]);
    assert.match(result.stdout, /icon policy PASS: 2 allowed file/);
    assert.match(result.stdout, new RegExp(ALLOWED_TOOLBAR.replaceAll("/", "\\/")));
  });
});

test("icon guard fails a new file that imports Hugeicons data directly", () => {
  withFixture(
    {
      files: {
        "packages/vue-pdf/src/components/NewToolbar.vue": HUGE_SOURCE,
      },
    },
    (fixture) => {
      const result = runVerifier(fixture);
      assert.equal(result.status, 1, result.stdout);
      assert.match(result.stderr, /ICON_IMPORT_NOT_ALLOWLISTED packages\/vue-pdf\/src\/components\/NewToolbar\.vue:2/);
      assert.match(result.stderr, /outside tests\/icons\/allowlist\.json/);
    },
  );
});

test("icon guard fails Lucide and Hugeicons in the same file unless the entry is mixed", () => {
  const mixedSource = `${HUGE_SOURCE}<script lang="ts">\nimport { Search } from "@lucide/vue"\n</script>\n`;
  withFixture(
    {
      files: { [ALLOWED_TOOLBAR]: mixedSource },
      entries: [parityEntry(), adapterEntry()],
      packages: {
        "packages/vue-pdf/package.json": JSON.stringify(
          {
            name: "@fixture/vue-pdf",
            dependencies: { [PARITY_ICON_PACKAGE]: "4.4.0", [DEFAULT_ICON_PACKAGE]: "1.76.0" },
          },
          null,
          2,
        ),
      },
    },
    (fixture) => {
      const rejected = runVerifier(fixture);
      assert.equal(rejected.status, 1, rejected.stdout);
      assert.match(rejected.stderr, /ICON_MIXED_ICON_SYSTEMS/);
      assert.match(rejected.stderr, /ICON_IMPORT_PACKAGE_NOT_LISTED/);

      const mixedEntry = parityEntry({
        iconPackages: [PARITY_ICON_PACKAGE, DEFAULT_ICON_PACKAGE],
        mixed: true,
        mixedReason: "fixture shares one toolbar with the registry adapter",
      });
      writeLedger(fixture, [mixedEntry, adapterEntry()]);
      const accepted = runVerifier(fixture);
      assert.equal(accepted.status, 0, `${accepted.stdout}\n${accepted.stderr}`);
      assert.equal(
        readSummary(fixture).allowedFiles.find((entry) => entry.file === ALLOWED_TOOLBAR).mixed,
        true,
      );
    },
  );
});

test("icon guard rejects any third icon library in source and dependencies", () => {
  withFixture(
    {
      files: {
        "packages/vue-pdf/src/components/ThirdParty.vue": `<script setup lang="ts">\nimport { IconArrow } from "@tabler/icons-vue"\n</script>\n`,
      },
      packages: {
        "packages/vue-pdf/package.json": JSON.stringify(
          {
            name: "@fixture/vue-pdf",
            dependencies: { [PARITY_ICON_PACKAGE]: "4.4.0" },
            devDependencies: { "@fortawesome/free-solid-svg-icons": "^6.0.0" },
          },
          null,
          2,
        ),
      },
    },
    (fixture) => {
      const result = runVerifier(fixture);
      assert.equal(result.status, 1, result.stdout);
      assert.match(result.stderr, /ICON_THIRD_PARTY_LIBRARY packages\/vue-pdf\/src\/components\/ThirdParty\.vue:2/);
      assert.match(result.stderr, /ICON_THIRD_PARTY_DEPENDENCY packages\/vue-pdf\/package\.json/);
    },
  );
});

test("icon guard fails closed on stale, unused and expired ledger rows", () => {
  withFixture(
    {
      entries: [parityEntry(), adapterEntry({ file: "packages/vue-ui/src/gone.ts" })],
    },
    (fixture) => {
      const stale = runVerifier(fixture);
      assert.equal(stale.status, 1, stale.stdout);
      assert.match(stale.stderr, /ICON_ALLOWLIST_STALE packages\/vue-ui\/src\/gone\.ts/);

      write(
        path.join(fixture, "packages/vue-pdf/src/components/PlainToolbar.vue"),
        "<template><div /></template>\n",
      );
      writeLedger(fixture, [
        parityEntry({ file: "packages/vue-pdf/src/components/PlainToolbar.vue" }),
        adapterEntry(),
      ]);
      const unused = runVerifier(fixture);
      assert.equal(unused.status, 1, unused.stdout);
      assert.match(
        unused.stderr,
        /ICON_ALLOWLIST_ENTRY_UNUSED packages\/vue-pdf\/src\/components\/PlainToolbar\.vue/,
      );

      writeLedger(fixture, [parityEntry(), adapterEntry()]);
      const expired = runVerifier(fixture, { ICON_GUARD_TODAY: "2030-01-01" });
      assert.equal(expired.status, 1, expired.stdout);
      assert.match(expired.stderr, /ICON_ALLOWLIST_EXPIRED/);
    },
  );
});

test("icon guard keeps the policy document and the ledger in sync", () => {
  withFixture(
    {
      doc: "# Fixture icon policy\n\nOfficeIcon and HugeIcon exist.\n\n@lucide/vue @hugeicons/core-free-icons\n",
    },
    (fixture) => {
      const result = runVerifier(fixture);
      assert.equal(result.status, 1, result.stdout);
      assert.match(result.stderr, /ICON_POLICY_DOC_DRIFT docs\/design\/icon-policy\.md/);
      assert.match(result.stderr, /packages\/vue-pdf\/src\/components\/PdfViewerToolbar\.vue/);
    },
  );
});

test("icon guard rejects a declared icon dependency with no allowlisted importer", () => {
  withFixture(
    {
      packages: {
        "packages/vue-docx/package.json": JSON.stringify(
          { name: "@fixture/vue-docx", dependencies: { [PARITY_ICON_PACKAGE]: "4.4.0" } },
          null,
          2,
        ),
        "packages/vue-docx/src/components/PlainToolbar.vue": "<template><div /></template>\n",
      },
    },
    (fixture) => {
      const result = runVerifier(fixture);
      assert.equal(result.status, 1, result.stdout);
      assert.match(result.stderr, /ICON_ORPHAN_DEPENDENCY packages\/vue-docx\/package\.json/);
    },
  );
});

test("the repository ledger matches the real icon imports", () => {
  const result = runVerifier(root, { ICON_GUARD_EVIDENCE_DIR: "" });
  assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);

  const ledger = JSON.parse(readFileSync(ledgerPath, "utf8"));
  const parityFiles = ledger.entries.filter((entry) =>
    entry.iconPackages.includes(PARITY_ICON_PACKAGE),
  );
  const chromeFiles = parityFiles.filter((entry) => entry.track === "office-chrome");

  // #180: the viewer chrome subset that motivated the policy is exactly these five.
  assert.deepEqual(
    chromeFiles.map((entry) => entry.file).sort(),
    [
      "packages/vue-docx/src/components/DocxViewerToolbar.vue",
      "packages/vue-pdf/src/components/PdfViewerToolbar.vue",
      "packages/vue-pptx/src/PptxViewerToolbar.vue",
      "packages/vue-xlsx/src/components/XlsxRibbon.vue",
      "packages/vue-xlsx/src/components/XlsxToolbar.vue",
    ],
  );
  assert.ok(parityFiles.every((entry) => entry.surface && entry.upstreamSource));
  assert.deepEqual(
    ledger.entries
      .filter((entry) => entry.track === "default-adapter")
      .map((entry) => entry.file)
      .sort(),
    ["packages/vue-ui/src/components/FileThumbnail.vue", "packages/vue-ui/src/icons.ts"],
  );

  const policyDocText = readFileSync(path.join(root, "docs/design/icon-policy.md"), "utf8");
  for (const entry of ledger.entries) {
    assert.ok(policyDocText.includes(entry.file), `${entry.file} is missing from the policy doc`);
  }
});
