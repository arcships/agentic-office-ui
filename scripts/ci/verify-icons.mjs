#!/usr/bin/env node

/**
 * Icon policy guard (#180).
 *
 * The project runs a controlled dual icon track (docs/design/icon-policy.md):
 *
 * 1. `OfficeIcon` + `@lucide/vue` is the default. Components ask for a registry
 *    name; only `packages/vue-ui/src/icons.ts` imports Lucide directly.
 * 2. `HugeIcon` + `@hugeicons/core-free-icons` is allowed only in the exact
 *    files listed in `tests/icons/allowlist.json`, because those files mirror a
 *    pinned upstream glyph set for pixel parity.
 *
 * The guard fails when a scanned source file imports `@hugeicons/*` outside the
 * ledger, mixes Lucide and Hugeicons without `mixed: true`, imports any third
 * icon library, or when a ledger entry no longer matches the code (stale,
 * unused, expired, or not backed by the vendored upstream snapshot). It also
 * guards `package.json`: a third icon library in any dependency block fails, a
 * declared icon dependency must be used by an allowlisted file, and an
 * allowlisted file must declare the package it imports.
 *
 * Environment overrides (used by the unit tests):
 *   ICON_GUARD_ROOT          repository (or fixture) root
 *   ICON_GUARD_ALLOWLIST     path to the ledger, default tests/icons/allowlist.json
 *   ICON_GUARD_EVIDENCE_DIR  directory for summary.json
 *   ICON_GUARD_TODAY         ISO date used for reviewBy expiry
 */
import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";

import {
  KNOWN_ICON_PACKAGES,
  PARITY_ICON_PACKAGE,
  collectImports,
  isThirdPartyIconLibrary,
  validateIconAllowlist,
} from "./icon-allowlist.mjs";

const root = path.resolve(process.env.ICON_GUARD_ROOT || path.join(import.meta.dirname, "../.."));
const allowlistPath = path.resolve(
  process.env.ICON_GUARD_ALLOWLIST || path.join(root, "tests", "icons", "allowlist.json"),
);
const today = process.env.ICON_GUARD_TODAY || new Date().toISOString().slice(0, 10);
const evidenceDir = process.env.ICON_GUARD_EVIDENCE_DIR;
const policyDocPath = path.resolve(
  process.env.ICON_GUARD_POLICY_DOC || path.join(root, "docs", "design", "icon-policy.md"),
);
const jsonOnly = process.argv.includes("--json");

const SOURCE_EXTENSIONS = new Set([".cjs", ".cts", ".js", ".jsx", ".mjs", ".mts", ".ts", ".tsx", ".vue"]);
const SCAN_ROOTS = ["packages", "apps"];
const DEPENDENCY_FIELDS = [
  "dependencies",
  "devDependencies",
  "optionalDependencies",
  "peerDependencies",
];

const violations = [];
const sourceFiles = [];
const dependencyFiles = [];
const declaredDependencies = [];

function repoPath(absolute) {
  return path.relative(root, absolute).split(path.sep).join("/");
}

function addViolation(code, file, message, line = null) {
  violations.push({ code, file: file ?? null, line, message });
}

function read(relative) {
  return readFileSync(path.join(root, relative), "utf8");
}

function readJson(relative) {
  return JSON.parse(read(relative));
}

function walk(directory, result = []) {
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const absolute = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      if (entry.name === "node_modules") continue;
      walk(absolute, result);
      continue;
    }
    if (SOURCE_EXTENSIONS.has(path.extname(entry.name))) result.push(absolute);
  }
  return result;
}

function workspaceDirectories() {
  const directories = [];
  for (const group of SCAN_ROOTS) {
    const groupPath = path.join(root, group);
    if (!existsSync(groupPath)) continue;
    for (const entry of readdirSync(groupPath, { withFileTypes: true })) {
      if (!entry.isDirectory() || entry.name === "node_modules") continue;
      const absolute = path.join(groupPath, entry.name);
      if (existsSync(path.join(absolute, "package.json"))) directories.push(absolute);
    }
  }
  return directories;
}

// 1. The ledger itself. A malformed, stale or expired entry fails closed.
let ledger = { entries: [], exemptDependencies: [] };
const ledgerRelative = path.relative(root, allowlistPath).split(path.sep).join("/");
if (!existsSync(allowlistPath)) {
  addViolation("ICON_ALLOWLIST_MISSING", ledgerRelative, "the icon allowlist does not exist");
} else {
  let parsed;
  try {
    parsed = JSON.parse(readFileSync(allowlistPath, "utf8"));
  } catch (error) {
    addViolation(
      "ICON_ALLOWLIST_SCHEMA",
      ledgerRelative,
      `invalid JSON: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
  if (parsed !== undefined) {
    const manifestPath = path.join(root, "tools", "parity-fixtures", "manifest.json");
    const knownSurfaces = existsSync(manifestPath)
      ? (JSON.parse(readFileSync(manifestPath, "utf8")).surfaces || []).map((surface) => surface.id)
      : [];
    const validated = validateIconAllowlist(parsed, { knownSurfaces, root, today });
    ledger = { entries: validated.entries, exemptDependencies: validated.exemptDependencies };
    for (const failure of validated.failures) {
      addViolation(failure.code, failure.file ?? ledgerRelative, failure.message);
    }
  }
}

const exemptDependencies = ledger.exemptDependencies ?? [];
const entryByFile = new Map(ledger.entries.map((entry) => [entry.file, entry]));

// 2. Source imports in packages/*/src and apps/*/src.
const workspaceRoots = workspaceDirectories();
for (const directory of workspaceRoots) {
  const sourceRoot = path.join(directory, "src");
  if (!existsSync(sourceRoot)) continue;
  for (const absolute of walk(sourceRoot)) {
    const relative = repoPath(absolute);
    const text = readFileSync(absolute, "utf8");
    const iconImports = collectImports(text).filter((entry) =>
      KNOWN_ICON_PACKAGES.includes(entry.package) ||
      isThirdPartyIconLibrary(entry.package, exemptDependencies),
    );
    const thirdParty = iconImports.filter(
      (entry) => isThirdPartyIconLibrary(entry.package, exemptDependencies),
    );
    for (const entry of thirdParty) {
      addViolation(
        "ICON_THIRD_PARTY_LIBRARY",
        relative,
        `${entry.specifier} is a third icon system; use <OfficeIcon name="..."> (Lucide) or the HugeIcon parity exception`,
        entry.line,
      );
    }
    const known = iconImports.filter((entry) => KNOWN_ICON_PACKAGES.includes(entry.package));
    if (known.length === 0) continue;

    const packages = [...new Set(known.map((entry) => entry.package))].sort();
    sourceFiles.push({ file: relative, packages });

    const entry = entryByFile.get(relative);
    if (!entry) {
      const first = known[0];
      addViolation(
        "ICON_IMPORT_NOT_ALLOWLISTED",
        relative,
        `${first.specifier} is imported outside tests/icons/allowlist.json; use <OfficeIcon name="...">, or add an exact ledger row with a parity surface and an upstream snapshot`,
        first.line,
      );
      continue;
    }

    for (const packageName of packages) {
      if (!entry.iconPackages.includes(packageName)) {
        const first = known.find((candidate) => candidate.package === packageName);
        addViolation(
          "ICON_IMPORT_PACKAGE_NOT_LISTED",
          relative,
          `the ledger entry does not list ${packageName} for this file`,
          first?.line ?? null,
        );
      }
    }

    if (packages.includes(PARITY_ICON_PACKAGE) && packages.length > 1 && entry.mixed !== true) {
      addViolation(
        "ICON_MIXED_ICON_SYSTEMS",
        relative,
        `this file imports both ${KNOWN_ICON_PACKAGES.join(" and ")}; split the file or declare mixed: true with a mixedReason`,
      );
    }

    for (const packageName of entry.iconPackages) {
      if (!packages.includes(packageName)) {
        addViolation(
          "ICON_ALLOWLIST_ENTRY_UNUSED_PACKAGE",
          relative,
          `the ledger allows ${packageName} here, but the file no longer imports it; shrink the entry`,
        );
      }
    }
  }
}

// 3. Ledger rows must describe a real import, not a reservation.
for (const entry of ledger.entries) {
  const file = sourceFiles.find((candidate) => candidate.file === entry.file);
  if (file) continue;
  if (!existsSync(path.join(root, entry.file))) continue; // already reported as stale
  addViolation(
    "ICON_ALLOWLIST_ENTRY_UNUSED",
    entry.file,
    "the ledger lists this file but it imports no icon data; delete the entry instead of reserving it",
  );
}

// 4. The policy doc is the human half of the ledger and must stay in sync.
const policyDocRelative = repoPath(policyDocPath);
const policyDoc = existsSync(policyDocPath) ? readFileSync(policyDocPath, "utf8") : null;
if (policyDoc === null) {
  addViolation("ICON_POLICY_DOC_MISSING", policyDocRelative, "the icon policy document does not exist");
} else {
  for (const entry of ledger.entries) {
    if (policyDoc.includes(entry.file)) continue;
    addViolation(
      "ICON_POLICY_DOC_DRIFT",
      policyDocRelative,
      `the allowlist row for ${entry.file} has no matching row in the policy doc`,
    );
  }
  for (const token of [PARITY_ICON_PACKAGE, KNOWN_ICON_PACKAGES[0], "OfficeIcon", "HugeIcon"]) {
    if (policyDoc.includes(token)) continue;
    addViolation("ICON_POLICY_DOC_DRIFT", policyDocRelative, `the policy must name ${token}`);
  }
}

// 5. package.json: no third icon library, no orphan declaration, no undeclared use.
for (const directory of workspaceRoots) {
  const relative = repoPath(path.join(directory, "package.json"));
  const manifest = JSON.parse(readFileSync(path.join(directory, "package.json"), "utf8"));
  dependencyFiles.push(relative);
  const declared = new Set();
  for (const field of DEPENDENCY_FIELDS) {
    for (const name of Object.keys(manifest[field] || {})) {
      declared.add(name);
      if (isThirdPartyIconLibrary(name, exemptDependencies)) {
        addViolation(
          "ICON_THIRD_PARTY_DEPENDENCY",
          relative,
          `${field} declares ${name}, a third icon system`,
        );
      }
    }
  }
  for (const packageName of KNOWN_ICON_PACKAGES) {
    if (declared.has(packageName)) declaredDependencies.push({ file: relative, package: packageName });
  }

  const localEntries = ledger.entries.filter((entry) =>
    entry.file.startsWith(`${repoPath(directory)}/`),
  );
  for (const packageName of KNOWN_ICON_PACKAGES) {
    if (!declared.has(packageName)) continue;
    if (!localEntries.some((entry) => entry.iconPackages.includes(packageName))) {
      addViolation(
        "ICON_ORPHAN_DEPENDENCY",
        relative,
        `${packageName} is declared but no allowlisted file in this package imports it; drop the dependency or use it through the shared renderer`,
      );
    }
  }
  for (const entry of localEntries) {
    for (const packageName of entry.iconPackages) {
      if (declared.has(packageName)) continue;
      addViolation(
        "ICON_UNDECLARED_DEPENDENCY",
        relative,
        `${entry.file} imports ${packageName} but this package does not declare it`,
      );
    }
  }
}

const allowedFiles = ledger.entries
  .map((entry) => ({
    file: entry.file,
    iconPackages: entry.iconPackages,
    issue: entry.issue,
    mixed: entry.mixed,
    reviewBy: entry.reviewBy,
    surface: entry.surface,
    track: entry.track,
    upstreamSource: entry.upstreamSource,
  }))
  .sort((left, right) => left.file.localeCompare(right.file));

const summary = {
  allowedFiles,
  declaredDependencies: declaredDependencies.sort((left, right) =>
    `${left.package}:${left.file}`.localeCompare(`${right.package}:${right.file}`),
  ),
  dependencyFiles: dependencyFiles.sort(),
  exemptDependencies,
  policyDoc: policyDocRelative,
  root,
  scannedSourceFiles: sourceFiles.length,
  status: violations.length === 0 ? "PASS" : "FAIL",
  today,
  violations: violations.sort((left, right) =>
    `${left.file}:${left.line ?? 0}:${left.code}`.localeCompare(
      `${right.file}:${right.line ?? 0}:${right.code}`,
    ),
  ),
};

if (evidenceDir) {
  mkdirSync(evidenceDir, { recursive: true });
  writeFileSync(path.join(evidenceDir, "summary.json"), `${JSON.stringify(summary, null, 2)}\n`);
}

if (jsonOnly) {
  process.stdout.write(`${JSON.stringify(summary, null, 2)}\n`);
} else {
  const lines = [
    `icon policy ${summary.status}: ${allowedFiles.length} allowed file(s), ${sourceFiles.length} source file(s) with icon imports, ${violations.length} violation(s)`,
  ];
  for (const entry of allowedFiles) {
    const surface = entry.surface ? ` surface=${entry.surface}` : "";
    lines.push(`  ${entry.track.padEnd(15)} ${entry.file} [${entry.iconPackages.join(", ")}]${surface}`);
  }
  if (violations.length > 0) {
    lines.push("", `icon policy FAIL (${violations.length} problem(s)):`);
    for (const violation of violations) {
      const location = violation.line ? `${violation.file}:${violation.line}` : violation.file;
      lines.push(`  ${violation.code} ${location} — ${violation.message}`);
    }
    process.stderr.write(`${lines.join("\n")}\n`);
  } else {
    process.stdout.write(`${lines.join("\n")}\n`);
  }
}

if (violations.length > 0) process.exitCode = 1;
