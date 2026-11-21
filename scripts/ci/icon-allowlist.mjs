/**
 * Icon allowlist helpers (#180).
 *
 * The icon policy is a controlled dual track, documented in
 * `docs/design/icon-policy.md`:
 *
 * - `OfficeIcon` is the default adapter. It renders `@lucide/vue` icons from
 *   the shared `OFFICE_ICONS` registry, so a component asks for a *name* and
 *   the adapter owns the package import.
 * - `HugeIcon` renders `@hugeicons/core-free-icons` data. It exists only for
 *   viewer chrome and Extend-parity components that must match the pinned
 *   upstream glyphs pixel for pixel. Every direct import of that data lives in
 *   `tests/icons/allowlist.json`, one exact file at a time.
 *
 * `scripts/ci/verify-icons.mjs` turns that ledger into a gate: a direct
 * `@hugeicons/*` import outside the ledger fails, a file that mixes Lucide and
 * Hugeicons data fails unless the entry declares `mixed`, and any third icon
 * library fails everywhere — in source and in `package.json`.
 *
 * The helpers here are pure apart from the optional existence checks, so the
 * policy can be unit tested without a workspace build.
 */
import { existsSync } from "node:fs";
import path from "node:path";

/** The pinned Extend UI commit every `upstreamSource` snapshot comes from. */
export { PINNED_UPSTREAM_COMMIT } from "./style-parity-allowlist.mjs";

/** Package that backs the default `OfficeIcon` adapter. */
export const DEFAULT_ICON_PACKAGE = "@lucide/vue";

/** Package that backs the pixel-parity `HugeIcon` exception. */
export const PARITY_ICON_PACKAGE = "@hugeicons/core-free-icons";

/** The only icon data packages allowed in `packages/*\/src` and `apps/*\/src`. */
export const KNOWN_ICON_PACKAGES = Object.freeze([
  DEFAULT_ICON_PACKAGE,
  PARITY_ICON_PACKAGE,
]);

/**
 * `default-adapter` is the single Lucide entry point.
 * `office-chrome` is the viewer toolbar/ribbon subset.
 * `extend-parity` is the ported Extend component set.
 */
export const ICON_TRACKS = Object.freeze([
  "default-adapter",
  "office-chrome",
  "extend-parity",
]);

/** Vendored upstream snapshots that anchor a parity exception. */
export const UPSTREAM_SNAPSHOT_PREFIX = "tools/extend-reference/src/vendor/";

/** Scoped icon packages, matched by their `@scope/` prefix. */
const THIRD_PARTY_ICON_SCOPES = Object.freeze([
  "@ant-design/icons",
  "@atlaskit/icon",
  "@blueprintjs/icons",
  "@carbon/icons",
  "@chakra-ui/icons",
  "@coreui/icons",
  "@element-plus/icons-vue",
  "@fluentui/react-icons",
  "@fortawesome/",
  "@heroicons/",
  "@hugeicons/",
  "@icon-park/",
  "@iconify/",
  "@iconoir/",
  "@lucide/",
  "@mdi/",
  "@mui/icons-material",
  "@nutui/icons-vue",
  "@phosphor-icons/",
  "@radix-ui/react-icons",
  "@remixicon/",
  "@tabler/",
  "@vicons/",
]);

/** Unscoped icon packages, matched exactly. */
const THIRD_PARTY_ICON_PACKAGES = Object.freeze([
  "bootstrap-icons",
  "boxicons",
  "css.gg",
  "eva-icons",
  "feather-icons",
  "font-awesome",
  "iconify",
  "ionicons",
  "line-awesome",
  "lucide",
  "lucide-react",
  "lucide-vue",
  "lucide-vue-next",
  "material-icons",
  "material-symbols",
  "octicons",
  "open-iconic",
  "phosphor-icons",
  "primeicons",
  "react-icons",
  "remixicon",
  "simple-icons",
  "unplugin-icons",
  "vue-feather-icons",
]);

const GLOB_PATTERN = /[*?[\]{}]/;
const SPECIFIER_PATTERNS = [
  /\bfrom\s+["']([^"']+)["']/gu,
  /\bimport\s*\(\s*["']([^"']+)["']\s*\)/gu,
  /\bimport\s+["']([^"']+)["']/gu,
  /\brequire\s*\(\s*["']([^"']+)["']\s*\)/gu,
];

/**
 * Reduce an import specifier to its package name.
 *
 * `@scope/pkg/sub/path` → `@scope/pkg`; `pkg/sub` → `pkg`; relative and
 * absolute specifiers → `null`.
 */
export function iconPackageName(specifier) {
  const value = String(specifier ?? "").trim();
  if (!value || value.startsWith(".") || value.startsWith("/") || value.startsWith("#")) {
    return null;
  }
  const segments = value.split("/");
  return value.startsWith("@") ? segments.slice(0, 2).join("/") : segments[0];
}

/**
 * Is this package name an icon library outside the two allowed tracks?
 *
 * Exact known packages win first, so `@hugeicons/core-free-icons` stays valid.
 * A declared `exemptDependencies` entry (a non-icon package whose name merely
 * contains "icon") also opts out; that list is part of the reviewed ledger.
 */
export function isThirdPartyIconLibrary(name, exemptDependencies = []) {
  const packageName = String(name ?? "").trim();
  if (!packageName) return false;
  if (KNOWN_ICON_PACKAGES.includes(packageName)) return false;
  if (exemptDependencies.includes(packageName)) return false;
  if (THIRD_PARTY_ICON_PACKAGES.includes(packageName)) return true;
  if (THIRD_PARTY_ICON_SCOPES.some((scope) => packageName.startsWith(scope))) return true;
  return /icon/iu.test(packageName);
}

/**
 * Every import specifier in a source file that resolves to a package.
 *
 * The regexes intentionally cover `.vue` SFCs without a full parser: only
 * module specifiers matter for this guard, and a false positive would have to
 * be a literal `from "@scope/icon-package"` inside a string.
 */
export function collectImports(sourceText) {
  const text = String(sourceText ?? "");
  const imports = [];
  const seen = new Set();
  for (const pattern of SPECIFIER_PATTERNS) {
    for (const match of text.matchAll(pattern)) {
      const specifier = match[1];
      if (!specifier) continue;
      const key = `${match.index}:${specifier}`;
      if (seen.has(key)) continue;
      seen.add(key);
      const packageName = iconPackageName(specifier);
      if (!packageName) continue;
      const line = text.slice(0, match.index ?? 0).split("\n").length;
      imports.push({ line, package: packageName, specifier });
    }
  }
  return imports.sort(
    (left, right) => left.line - right.line || left.specifier.localeCompare(right.specifier),
  );
}

/** Icon packages a source file imports, in declaration order. */
export function iconImportsIn(sourceText) {
  return collectImports(sourceText).filter((entry) =>
    KNOWN_ICON_PACKAGES.includes(entry.package),
  );
}

function isPlainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function isSafeRelativeFile(file) {
  return (
    typeof file === "string" &&
    file.length > 0 &&
    !file.startsWith("/") &&
    !file.includes("\\") &&
    !file.split("/").includes("..") &&
    !GLOB_PATTERN.test(file)
  );
}

function isValidDate(value) {
  if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}$/u.test(value)) return false;
  const parsed = new Date(`${value}T00:00:00Z`);
  return !Number.isNaN(parsed.getTime()) && parsed.toISOString().slice(0, 10) === value;
}

function nonEmptyString(value) {
  return typeof value === "string" && value.trim().length > 0;
}

/**
 * Validate the icon ledger.
 *
 * Every failure is a `{code, file, message}` object; callers print them and
 * exit non-zero. Nothing here "repairs" an entry: a malformed, stale or
 * expired entry fails closed instead of silently exempting a file.
 *
 * @param {unknown} data parsed `tests/icons/allowlist.json`
 * @param {object} options
 * @param {string} options.root repository (or fixture) root for existence checks
 * @param {string[]} [options.knownSurfaces] parity surface ids from the manifest
 * @param {string} [options.today] ISO date used for `reviewBy` expiry
 */
export function validateIconAllowlist(data, options = {}) {
  const {
    root = process.cwd(),
    knownSurfaces = [],
    today = new Date().toISOString().slice(0, 10),
  } = options;
  const failures = [];
  const entries = [];
  const fail = (code, file, message) =>
    failures.push({ code, file: file ?? null, message });

  if (!isPlainObject(data)) {
    fail("ICON_ALLOWLIST_SCHEMA", null, "tests/icons/allowlist.json must be a JSON object");
    return { entries, failures, exemptDependencies: [] };
  }
  if (data.schemaVersion !== 1) {
    fail("ICON_ALLOWLIST_SCHEMA", null, "schemaVersion must be 1");
  }

  const exemptDependencies = Array.isArray(data.exemptDependencies) ? data.exemptDependencies : [];
  if (data.exemptDependencies !== undefined && !Array.isArray(data.exemptDependencies)) {
    fail("ICON_ALLOWLIST_SCHEMA", null, "exemptDependencies must be an array of package names");
  }
  for (const name of exemptDependencies) {
    if (!nonEmptyString(name)) {
      fail("ICON_ALLOWLIST_SCHEMA", null, "exemptDependencies entries must be non-empty strings");
    }
  }

  if (!Array.isArray(data.entries) || data.entries.length === 0) {
    fail("ICON_ALLOWLIST_SCHEMA", null, "entries must be a non-empty array");
    return { entries, failures, exemptDependencies };
  }

  const seenFiles = new Set();
  let previousFile = "";
  for (const [index, raw] of data.entries.entries()) {
    const label = isPlainObject(raw) && typeof raw.file === "string" ? raw.file : `entries[${index}]`;
    if (!isPlainObject(raw)) {
      fail("ICON_ALLOWLIST_SCHEMA", label, "entry must be an object");
      continue;
    }

    const {
      file,
      iconPackages,
      track,
      surface,
      upstreamSource,
      mixed,
      mixedReason,
      issue,
      reason,
      reviewBy,
    } = raw;
    let valid = true;

    if (!isSafeRelativeFile(file)) {
      fail(
        "ICON_ALLOWLIST_FILE",
        label,
        "file must be an exact repository-relative path (no globs, no .., no absolute paths)",
      );
      valid = false;
    } else if (seenFiles.has(file)) {
      fail("ICON_ALLOWLIST_DUPLICATE", file, "duplicate entry for the same file");
      valid = false;
    } else {
      seenFiles.add(file);
      if (file < previousFile) {
        fail(
          "ICON_ALLOWLIST_ORDER",
          file,
          "entries must stay sorted by file so the ledger diffs cleanly",
        );
      }
      previousFile = file;
      if (!existsSync(path.join(root, file))) {
        fail("ICON_ALLOWLIST_STALE", file, "allowlisted file does not exist; delete the stale entry");
        valid = false;
      }
    }

    if (!Array.isArray(iconPackages) || iconPackages.length === 0) {
      fail("ICON_ALLOWLIST_PACKAGES", label, "iconPackages must list at least one icon package");
      valid = false;
    } else {
      for (const packageName of iconPackages) {
        if (!KNOWN_ICON_PACKAGES.includes(packageName)) {
          fail(
            "ICON_ALLOWLIST_PACKAGES",
            label,
            `${packageName} is not a valid icon package; use ${KNOWN_ICON_PACKAGES.join(" or ")}`,
          );
          valid = false;
        }
      }
      if (new Set(iconPackages).size !== iconPackages.length) {
        fail("ICON_ALLOWLIST_PACKAGES", label, "iconPackages contains duplicates");
        valid = false;
      }
    }

    if (!ICON_TRACKS.includes(track)) {
      fail("ICON_ALLOWLIST_TRACK", label, `track must be one of ${ICON_TRACKS.join(", ")}`);
      valid = false;
    }

    const packages = Array.isArray(iconPackages) ? iconPackages : [];
    const usesParity = packages.includes(PARITY_ICON_PACKAGE);
    const usesDefault = packages.includes(DEFAULT_ICON_PACKAGE);

    if (track === "default-adapter") {
      if (packages.length !== 1 || !usesDefault) {
        fail("ICON_ALLOWLIST_TRACK", label, "default-adapter entries may only declare the Lucide package");
        valid = false;
      }
      if (isSafeRelativeFile(file) && !file.startsWith("packages/vue-ui/src/")) {
        fail(
          "ICON_ALLOWLIST_TRACK",
          file,
          "the default Lucide adapter lives in packages/vue-ui/src; other packages must use OfficeIcon names",
        );
        valid = false;
      }
    } else if (track !== undefined && !usesParity) {
      fail(
        "ICON_ALLOWLIST_TRACK",
        label,
        `${track} entries must declare ${PARITY_ICON_PACKAGE}; only default-adapter may be Lucide-only`,
      );
      valid = false;
    }

    if (usesParity && usesDefault) {
      if (mixed !== true) {
        fail(
          "ICON_ALLOWLIST_MIXED",
          label,
          "a file that imports both icon data packages needs `mixed: true` and a mixedReason",
        );
        valid = false;
      } else if (!nonEmptyString(mixedReason)) {
        fail("ICON_ALLOWLIST_MIXED", label, "mixed entries need a mixedReason");
        valid = false;
      }
    }

    if (track !== "default-adapter") {
      if (!nonEmptyString(surface)) {
        fail("ICON_ALLOWLIST_SURFACE", label, "parity entries must name the parity surface they follow");
        valid = false;
      } else if (knownSurfaces.length > 0 && !knownSurfaces.includes(surface)) {
        fail(
          "ICON_ALLOWLIST_SURFACE",
          label,
          `${surface} is not a surface in tools/parity-fixtures/manifest.json`,
        );
        valid = false;
      }
      if (!nonEmptyString(upstreamSource) || !upstreamSource.startsWith(UPSTREAM_SNAPSHOT_PREFIX)) {
        fail(
          "ICON_ALLOWLIST_UPSTREAM",
          label,
          `parity entries must point at a pinned snapshot under ${UPSTREAM_SNAPSHOT_PREFIX}`,
        );
        valid = false;
      } else if (!existsSync(path.join(root, upstreamSource))) {
        fail(
          "ICON_ALLOWLIST_UPSTREAM",
          label,
          `${upstreamSource} is not a vendored snapshot in this repository`,
        );
        valid = false;
      }
    }

    if (!nonEmptyString(issue) || !/^#\d+$/u.test(issue)) {
      fail("ICON_ALLOWLIST_OWNER", label, "issue must be a GitHub issue reference like #180");
      valid = false;
    }
    if (!nonEmptyString(reason)) {
      fail("ICON_ALLOWLIST_OWNER", label, "reason must explain why the exception is needed");
      valid = false;
    }
    if (!isValidDate(reviewBy)) {
      fail("ICON_ALLOWLIST_OWNER", label, "reviewBy must be an ISO date (YYYY-MM-DD)");
      valid = false;
    } else if (reviewBy < today) {
      fail(
        "ICON_ALLOWLIST_EXPIRED",
        label,
        `reviewBy ${reviewBy} has passed; re-review or remove the entry`,
      );
      valid = false;
    }

    if (valid) {
      entries.push({
        file,
        iconPackages: [...packages],
        issue,
        mixed: mixed === true,
        reason,
        reviewBy,
        surface: surface ?? null,
        track,
        upstreamSource: upstreamSource ?? null,
      });
    }
  }

  return { entries, failures, exemptDependencies };
}
