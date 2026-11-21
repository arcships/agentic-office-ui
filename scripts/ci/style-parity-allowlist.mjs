/**
 * Parity-literal allowlist helpers (#175).
 *
 * `tests/style/component-style-baseline.json` is the legacy-debt ledger: it
 * only counts component CSS violations that already existed before the token
 * migration. `tests/style/parity-literal-allowlist.json` is the deliberate
 * exception ledger: a small, exact list of upstream Tailwind literals that are
 * kept bare because the token layer cannot reproduce their pixel geometry.
 *
 * The two ledgers are deliberately different:
 *
 * - a baseline entry is a per-file + rule count and can only shrink;
 * - an allowlist entry names one exact `file` + `rule` + `property` + `value`
 *   plus a bounded `count`, the pinned upstream path/commit, the owning issue,
 *   a reason and a `reviewBy` date. Globs, prefixes and wildcards are rejected.
 *
 * `scripts/ci/lint-styles.mjs` subtracts allowlisted warnings before it
 * compares the remaining warnings with the baseline, so an upstream parity
 * literal never has to grow the generic baseline and a new bare value fails.
 *
 * The functions here are pure (except `validateParityAllowlist`'s optional
 * existence check) so the policy can be unit tested without running stylelint
 * again.
 */
import { existsSync } from "node:fs";
import path from "node:path";

/** The three stylelint rules guarded by `stylelint.config.mjs`. */
export const STYLE_GUARD_RULES = [
  "color-no-hex",
  "declaration-property-unit-disallowed-list",
  "declaration-property-value-disallowed-list",
];

/** The pinned Extend UI commit every upstreamPath/review refers to. */
export const PINNED_UPSTREAM_COMMIT = "70fe70c6f1e1eee53954c6923f3d9f583557a57e";

const UNIT_VALUES = new Set(["px", "rem", "em", "pt", "%", "vw", "vh", "s", "ms"]);
const GLOB_PATTERN = /[*?[\]{}]/;

/**
 * Parse the structured part of a stylelint warning.
 *
 * stylelint emits three message shapes for the rules above:
 *
 *   Unexpected hex color "#fff" (color-no-hex)
 *   Unexpected unit "px" for property "border-radius" (...unit...)
 *   Unexpected value "0 12px 40px rgba(0, 0, 0, 0.12)" for property "box-shadow"
 *
 * The unit message reports the offending unit, so `value` is the unit for
 * unit rules and the full literal for value rules. Hex warnings have no CSS
 * property, so `property` is `null`; the allowlist schema mirrors that.
 *
 * Unknown shapes return `kind: "unknown"` with `property`/`value` set to
 * `null`; such warnings can never be allowlisted and therefore fail as new
 * violations instead of being silently absorbed.
 */
export function parseStylelintWarningText(text) {
  const message = String(text ?? "").trim();
  let match = message.match(/^Unexpected hex color "([^"]+)"/u);
  if (match) {
    return { kind: "hex", property: null, value: match[1], unit: null };
  }
  match = message.match(/^Unexpected unit "([^"]+)" for property "([^"]+)"/u);
  if (match) {
    return { kind: "unit", property: match[2], value: match[1], unit: match[1] };
  }
  match = message.match(/^Unexpected value "(.*)" for property "([^"]+)"/u);
  if (match) {
    return { kind: "value", property: match[2], value: match[1], unit: null };
  }
  return { kind: "unknown", property: null, value: null, unit: null };
}

function toIsoDate(value) {
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  return String(value).slice(0, 10);
}

function isValidIsoDate(value) {
  if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}$/u.test(value)) return false;
  const parsed = new Date(`${value}T00:00:00Z`);
  return !Number.isNaN(parsed.valueOf()) && parsed.toISOString().slice(0, 10) === value;
}

function isNonEmptyString(value) {
  return typeof value === "string" && value.trim().length > 0;
}

/**
 * Validate an allowlist document and return the entries the gate may use.
 *
 * Invalid and expired entries are *not* returned: their warnings stay visible
 * as new violations, so a typo can never silently exempt code. The gate turns
 * `problems` and `expired` into a failure; `stale` (declared `count` larger
 * than the actual matches) is reported but does not block, because it means
 * debt was removed and the ledger should shrink.
 *
 * @param {unknown} document parsed `parity-literal-allowlist.json`
 * @param {{root?: string, today?: string|Date}} options
 * @returns {{entries: object[], problems: object[], expired: object[]}}
 */
export function validateParityAllowlist(document, { root, today = new Date() } = {}) {
  const problems = [];
  const expired = [];
  const entries = [];
  const todayIso = toIsoDate(today);
  const seenIds = new Set();

  if (!document || typeof document !== "object" || Array.isArray(document)) {
    return {
      entries,
      problems: [{ index: null, id: null, messages: ["allowlist must be a JSON object"] }],
      expired,
    };
  }
  if (document.schemaVersion !== 1) {
    problems.push({
      index: null,
      id: null,
      messages: [`schemaVersion must be 1 (received ${JSON.stringify(document.schemaVersion)})`],
    });
  }
  if (!Array.isArray(document.entries)) {
    problems.push({ index: null, id: null, messages: ["entries must be an array"] });
    return { entries, problems, expired };
  }

  document.entries.forEach((raw, index) => {
    const id = isNonEmptyString(raw?.id) ? raw.id.trim() : null;
    const label = id ? `entries[${index}] (${id})` : `entries[${index}]`;
    const messages = [];

    if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
      problems.push({ index, id, messages: [`${label} must be an object`] });
      return;
    }
    if (!id || !/^[a-z0-9][a-z0-9-]*$/u.test(id)) {
      messages.push("id must be lowercase kebab-case");
    } else if (seenIds.has(id)) {
      messages.push(`duplicate id ${id}`);
    }
    seenIds.add(id);

    if (!isNonEmptyString(raw.file)) {
      messages.push("file must be a repository-relative path");
    } else {
      const file = raw.file.trim();
      if (
        file.includes("\\") ||
        file.startsWith("/") ||
        file.split("/").includes("..") ||
        GLOB_PATTERN.test(file)
      ) {
        messages.push("file must be an exact repository-relative path (no globs, no '..')");
      } else if (root && !existsSync(path.join(root, file))) {
        messages.push(`file does not exist: ${file}`);
      }
    }

    if (!STYLE_GUARD_RULES.includes(raw.rule)) {
      messages.push(`rule must be one of ${STYLE_GUARD_RULES.join(", ")}`);
    }

    if (raw.rule === "color-no-hex") {
      if (raw.property !== null) {
        messages.push("property must be null for color-no-hex");
      }
    } else if (!isNonEmptyString(raw.property)) {
      messages.push("property must name the CSS property");
    } else if (GLOB_PATTERN.test(raw.property)) {
      messages.push("property must be exact (no wildcards)");
    }

    if (!isNonEmptyString(raw.value)) {
      messages.push("value must be the exact stylelint literal");
    } else if (GLOB_PATTERN.test(raw.value)) {
      messages.push("value must be exact (no wildcards)");
    } else if (
      raw.rule === "declaration-property-unit-disallowed-list" &&
      !UNIT_VALUES.has(raw.value)
    ) {
      messages.push(`value must be a guarded unit for unit rules: ${[...UNIT_VALUES].join(", ")}`);
    } else if (raw.rule === "color-no-hex" && !/^#[0-9a-f]{3,8}$/iu.test(raw.value)) {
      messages.push("value must be an exact hex colour for color-no-hex");
    }

    if (!Number.isInteger(raw.count) || raw.count <= 0) {
      messages.push("count must be a positive integer");
    } else if (raw.count > 1000) {
      messages.push("count must stay bounded (<= 1000)");
    }

    if (!isNonEmptyString(raw.upstreamPath) || raw.upstreamPath.trim().length < 5) {
      messages.push("upstreamPath must name the pinned upstream file");
    } else if (raw.upstreamPath.split("/").includes("..")) {
      messages.push("upstreamPath must not escape the upstream repository");
    }
    if (!isNonEmptyString(raw.upstreamCommit) || !/^[0-9a-f]{7,40}$/iu.test(raw.upstreamCommit)) {
      messages.push("upstreamCommit must be a 7–40 character hex commit");
    }

    if (
      !isNonEmptyString(raw.issue) ||
      !(/^#\d+$/u.test(raw.issue) || /^https?:\/\//u.test(raw.issue))
    ) {
      messages.push("issue must look like #175 or an issue URL");
    }
    if (!isNonEmptyString(raw.reason) || raw.reason.trim().length < 40) {
      messages.push("reason must explain why the literal is parity, not legacy debt");
    }
    if (!isValidIsoDate(raw.reviewBy)) {
      messages.push("reviewBy must be an ISO date (YYYY-MM-DD)");
    }
    if (
      raw.source !== undefined &&
      raw.source !== "upstream" &&
      raw.source !== "product-affordance"
    ) {
      messages.push("source must be upstream or product-affordance when present");
    }

    if (messages.length > 0) {
      problems.push({ index, id, messages: [`${label}: ${messages.join("; ")}`] });
      return;
    }
    if (raw.reviewBy < todayIso) {
      expired.push({
        index,
        id,
        file: raw.file,
        reviewBy: raw.reviewBy,
        issue: raw.issue,
        message: `${label}: reviewBy ${raw.reviewBy} has passed (${todayIso})`,
      });
      return;
    }

    entries.push({
      id,
      file: raw.file,
      rule: raw.rule,
      property: raw.property,
      value: raw.value,
      count: raw.count,
      upstreamPath: raw.upstreamPath,
      upstreamCommit: raw.upstreamCommit,
      issue: raw.issue,
      reason: raw.reason.trim(),
      reviewBy: raw.reviewBy,
      source: raw.source || "upstream",
    });
  });

  return { entries, problems, expired };
}

/**
 * Subtract allowlisted warnings from the stylelint result.
 *
 * Each entry consumes at most `count` warnings with exactly the same file,
 * rule, property and value. Warnings are consumed in stylelint order, so the
 * evidence is deterministic. A `stale` record means the entry declared more
 * occurrences than exist now; the gate reports it so the ledger can shrink.
 *
 * @param {object[]} warnings normalised warnings (`file`, `rule`, `text`, …)
 * @param {object[]} entries validated allowlist entries
 */
export function matchParityLiterals(warnings, entries) {
  const consumed = new Set();
  const matches = [];
  const stale = [];

  for (const entry of entries) {
    const matched = [];
    for (let index = 0; index < warnings.length && matched.length < entry.count; index += 1) {
      if (consumed.has(index)) continue;
      const warning = warnings[index];
      if (warning.file !== entry.file || warning.rule !== entry.rule) continue;
      const parsed = warning.parsed || parseStylelintWarningText(warning.text);
      if (parsed.property !== entry.property || parsed.value !== entry.value) continue;
      consumed.add(index);
      matched.push({
        file: warning.file,
        line: warning.line,
        column: warning.column,
        rule: warning.rule,
        property: parsed.property,
        value: parsed.value,
      });
    }

    if (matched.length !== entry.count) {
      stale.push({
        id: entry.id,
        file: entry.file,
        rule: entry.rule,
        property: entry.property,
        value: entry.value,
        expected: entry.count,
        matched: matched.length,
        missing: entry.count - matched.length,
      });
    }
    matches.push({
      id: entry.id,
      file: entry.file,
      rule: entry.rule,
      property: entry.property,
      value: entry.value,
      issue: entry.issue,
      upstreamPath: entry.upstreamPath,
      upstreamCommit: entry.upstreamCommit,
      count: entry.count,
      matched: matched.length,
      warnings: matched,
    });
  }

  return { consumed, matches, stale };
}

/**
 * Build the evidence/report block for the gate.
 */
export function summarizeParityAllowlist({ path: allowlistPath, document, validation, matching }) {
  const matched = matching.matches.reduce((total, entry) => total + entry.matched, 0);
  return {
    path: allowlistPath,
    schemaVersion: document?.schemaVersion ?? null,
    entries: validation.entries.length,
    declared: validation.entries.reduce((total, entry) => total + entry.count, 0),
    matched,
    matches: matching.matches,
    stale: matching.stale,
    expired: validation.expired,
    invalid: validation.problems,
  };
}
