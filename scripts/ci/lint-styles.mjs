#!/usr/bin/env node

/**
 * Component CSS token guards (#165, #175).
 *
 * Runs stylelint (`stylelint.config.mjs`) over every component style in
 * `packages/*\/src` and polices two separate ledgers:
 *
 * 1. `tests/style/component-style-baseline.json` — historical debt.
 *    Counts are per file + rule. When a file already freezes N occurrences of
 *    a rule, the N reported lines are treated as frozen, so a new occurrence
 *    of that same rule in the same file fails the run. Fixing violations never
 *    fails; the summary reports entries that can shrink. The baseline is
 *    shrink-only: `pnpm lint:style:baseline --approve ...` records a refresh,
 *    but it refuses to add a new file, rule or occurrence.
 *
 * 2. `tests/style/parity-literal-allowlist.json` — deliberate upstream
 *    pixel-parity literals (#175). Each entry names one exact file + rule +
 *    property + value with a bounded count, the pinned upstream path/commit,
 *    the owning issue, a reason and a review date. Allowlisted warnings are
 *    deducted *before* the baseline comparison, so an upstream Tailwind
 *    literal never grows the generic baseline. Globs, prefixes and wildcards
 *    are rejected by `scripts/ci/style-parity-allowlist.mjs`.
 *
 *    A warning that matches neither ledger is a new violation and fails.
 *    Entries that no longer match are reported as stale (shrink the ledger);
 *    invalid or expired entries fail closed, so a typo or a missed review can
 *    never silently exempt code.
 *
 * The rules are never disabled for a directory: the token files are exempt in
 * `stylelint.config.mjs` because defining raw values is their whole job.
 *
 * Environment overrides (used by the unit tests):
 *   STYLE_GUARD_FILES           comma-separated files/globs to lint
 *   STYLE_GUARD_BASELINE        path to the legacy baseline
 *   STYLE_GUARD_ALLOWLIST       path to the parity-literal allowlist
 *   STYLE_GUARD_EVIDENCE_DIR    evidence directory
 *   STYLE_GUARD_TODAY           ISO date used for reviewBy expiry
 */
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import stylelint from "stylelint";
import {
  matchParityLiterals,
  parseStylelintWarningText,
  summarizeParityAllowlist,
  validateParityAllowlist,
} from "./style-parity-allowlist.mjs";

const root = path.resolve(import.meta.dirname, "../..");
const baselinePath = process.env.STYLE_GUARD_BASELINE
  ? path.resolve(root, process.env.STYLE_GUARD_BASELINE)
  : path.join(root, "tests/style/component-style-baseline.json");
const allowlistPath = process.env.STYLE_GUARD_ALLOWLIST
  ? path.resolve(root, process.env.STYLE_GUARD_ALLOWLIST)
  : path.join(root, "tests/style/parity-literal-allowlist.json");
const evidenceDir = path.resolve(
  process.env.STYLE_GUARD_EVIDENCE_DIR || path.join(root, "output/style-guard"),
);
const today = process.env.STYLE_GUARD_TODAY || new Date().toISOString().slice(0, 10);
const files = process.env.STYLE_GUARD_FILES
  ? process.env.STYLE_GUARD_FILES.split(",").map((file) => path.resolve(root, file.trim()))
  : ["packages/*/src/**/*.{vue,css}"];
const suite = "P1-STYLE-TOKEN-GUARD";

function parseFlag(argv, name) {
  const index = argv.indexOf(name);
  if (index === -1) return undefined;
  const value = argv[index + 1];
  if (value === undefined || value.startsWith("--")) {
    console.error(`lint-styles: ${name} needs a value`);
    process.exit(2);
  }
  return value;
}

async function lint() {
  const result = await stylelint.lint({
    files,
    configFile: path.join(root, "stylelint.config.mjs"),
    formatter: "json",
  });
  const warnings = [];
  for (const file of result.results) {
    const source = path.relative(root, file.source).split(path.sep).join("/");
    if (file.errored && file.warnings.some((warning) => !warning.rule)) {
      const parseErrors = file.warnings.filter((warning) => !warning.rule);
      warnings.push(
        ...parseErrors.map((warning) => ({
          file: source,
          line: warning.line || 1,
          column: warning.column || 1,
          rule: "parse-error",
          text: warning.text,
          parsed: parseStylelintWarningText(warning.text),
        })),
      );
      continue;
    }
    warnings.push(
      ...file.warnings.map((warning) => ({
        file: source,
        line: warning.line || 1,
        column: warning.column || 1,
        rule: warning.rule || "unknown",
        text: warning.text,
        parsed: parseStylelintWarningText(warning.text),
      })),
    );
  }
  warnings.sort((left, right) =>
    left.file.localeCompare(right.file) || left.line - right.line || left.column - right.column,
  );
  return warnings;
}

function countByFileAndRule(warnings) {
  const counts = {};
  for (const warning of warnings) {
    counts[warning.file] ??= {};
    counts[warning.file][warning.rule] = (counts[warning.file][warning.rule] || 0) + 1;
  }
  return counts;
}

function loadJson(filePath, fallback) {
  try {
    return JSON.parse(readFileSync(filePath, "utf8"));
  } catch (error) {
    if (error.code === "ENOENT") return fallback;
    throw error;
  }
}

function loadBaseline() {
  return loadJson(baselinePath, { counts: {}, approved: null, totals: {} });
}

function loadAllowlist() {
  let text;
  try {
    text = readFileSync(allowlistPath, "utf8");
  } catch (error) {
    if (error.code !== "ENOENT") throw error;
    return {
      document: { schemaVersion: 1, entries: [] },
      error: {
        index: null,
        id: null,
        messages: [`allowlist not found: ${path.relative(root, allowlistPath)}`],
      },
    };
  }
  try {
    return { document: JSON.parse(text), error: null };
  } catch (error) {
    return {
      document: { schemaVersion: 1, entries: [] },
      error: {
        index: null,
        id: null,
        messages: [
          `allowlist is not valid JSON (${path.relative(root, allowlistPath)}): ${error.message}`,
        ],
      },
    };
  }
}

function countBaselineEntries(counts) {
  return Object.values(counts).reduce(
    (total, rules) => total + Object.values(rules).reduce((sum, value) => sum + value, 0),
    0,
  );
}

function requireApproval(argv) {
  if (!argv.includes("--approve")) {
    console.error(
      [
        "lint-styles: refreshing the style baseline requires an approval.",
        "",
        "  pnpm lint:style:baseline --approve --reviewer <github-handle> \\",
        "    --issue <#nnn> --reason <why>",
        "",
        "Fix the violation instead whenever the value can come from",
        "packages/vue-ui/src/styles/tokens.css. If it is an intentional",
        "upstream parity literal, add an exact entry to",
        "tests/style/parity-literal-allowlist.json instead. The legacy baseline",
        "is shrink-only and cannot absorb new debt.",
      ].join("\n"),
    );
    process.exit(2);
  }
  const reviewer = parseFlag(argv, "--reviewer");
  const issue = parseFlag(argv, "--issue");
  const reason = parseFlag(argv, "--reason");
  if (!reviewer || !issue || !reason) {
    console.error("lint-styles: --approve needs --reviewer, --issue and --reason");
    process.exit(2);
  }
  if (reason.trim().length < 8) {
    console.error("lint-styles: --reason must explain why the debt is acceptable");
    process.exit(2);
  }
  if (!issue.startsWith("#") && !/^https?:/.test(issue)) {
    console.error(`lint-styles: --issue must look like #165 (received ${issue})`);
    process.exit(2);
  }
  return { reviewer, issue, reason: reason.trim(), date: new Date().toISOString().slice(0, 10) };
}

function main() {
  const argv = process.argv.slice(2);
  const writeBaseline = argv.includes("--write-baseline");
  const reportOnly = argv.includes("--report-only");

  return lint().then((warnings) => {
    const loadedAllowlist = loadAllowlist();
    const allowlistValidation = validateParityAllowlist(loadedAllowlist.document, {
      root,
      today,
    });
    if (loadedAllowlist.error) {
      allowlistValidation.problems.unshift(loadedAllowlist.error);
    }
    const parity = matchParityLiterals(warnings, allowlistValidation.entries);
    const paritySummary = summarizeParityAllowlist({
      path: path.relative(root, allowlistPath),
      document: loadedAllowlist.document,
      validation: allowlistValidation,
      matching: parity,
    });

    // Only warnings that are not parity literals are compared with the legacy
    // baseline. This is the #175 separation: upstream literals never enter the
    // generic baseline, and a new bare value still fails.
    const remainingWarnings = warnings.filter((_warning, index) => !parity.consumed.has(index));
    const counts = countByFileAndRule(remainingWarnings);
    const baseline = loadBaseline();
    const approvedCounts = baseline.counts || {};

    const newViolations = [];
    for (const [file, rules] of Object.entries(counts)) {
      for (const [rule, count] of Object.entries(rules)) {
        const allowed = approvedCounts[file]?.[rule] || 0;
        if (count > allowed) {
          const fresh = remainingWarnings.filter(
            (warning) => warning.file === file && warning.rule === rule,
          );
          newViolations.push(...fresh.slice(allowed));
        }
      }
    }

    const shrinkable = [];
    let baselined = 0;
    for (const [file, rules] of Object.entries(approvedCounts)) {
      for (const [rule, allowed] of Object.entries(rules)) {
        const actual = counts[file]?.[rule] || 0;
        baselined += Math.min(actual, allowed);
        if (actual < allowed) shrinkable.push({ file, rule, allowed, actual });
      }
    }

    const policyProblems = [
      ...allowlistValidation.problems,
      ...allowlistValidation.expired.map((entry) => ({
        index: entry.index,
        id: entry.id,
        messages: [entry.message],
      })),
    ];
    const parityAllowlisted = paritySummary.matched;
    const totals = {
      files: new Set(warnings.map((warning) => warning.file)).size,
      warnings: warnings.length,
      byRule: warnings.reduce((accumulator, warning) => {
        accumulator[warning.rule] = (accumulator[warning.rule] || 0) + 1;
        return accumulator;
      }, {}),
      parityAllowlisted,
      baselined,
      new: newViolations.length,
    };

    mkdirSync(evidenceDir, { recursive: true });

    if (writeBaseline) {
      if (policyProblems.length) {
        const summary = {
          suite,
          result: "FAIL",
          reason: "parity-literal allowlist has invalid or expired entries",
          parityAllowlist: paritySummary,
        };
        writeFileSync(
          path.join(evidenceDir, "summary.json"),
          `${JSON.stringify(summary, null, 2)}\n`,
        );
        console.error("lint-styles: refusing to write the baseline — fix the parity allowlist:");
        for (const problem of policyProblems) {
          for (const message of problem.messages) console.error(`  ${message}`);
        }
        return 2;
      }
      const approval = requireApproval(argv);
      const growth = [];
      for (const [file, rules] of Object.entries(counts)) {
        for (const [rule, count] of Object.entries(rules)) {
          const allowed = approvedCounts[file]?.[rule] || 0;
          if (count > allowed) growth.push({ file, rule, allowed, actual: count });
        }
      }
      if (growth.length) {
        const summary = {
          suite,
          result: "FAIL",
          reason: "the legacy style baseline is shrink-only",
          growth,
          parityAllowlist: paritySummary,
        };
        writeFileSync(
          path.join(evidenceDir, "summary.json"),
          `${JSON.stringify(summary, null, 2)}\n`,
        );
        console.error(
          "lint-styles: refusing to write the baseline — the legacy baseline is shrink-only. " +
            "New bare values must be tokenised or added to tests/style/parity-literal-allowlist.json:",
        );
        for (const entry of growth) {
          console.error(
            `  ${entry.file} ${entry.rule}: ${entry.allowed} -> ${entry.actual}`,
          );
        }
        return 2;
      }
      const document = {
        schemaVersion: 1,
        generatedBy: "node scripts/ci/lint-styles.mjs --write-baseline",
        rules: [
          "color-no-hex",
          "declaration-property-unit-disallowed-list",
          "declaration-property-value-disallowed-list",
        ],
        approved: approval,
        counts: Object.fromEntries(
          Object.entries(counts)
            .sort(([left], [right]) => left.localeCompare(right))
            .map(([file, rules]) => [
              file,
              Object.fromEntries(
                Object.entries(rules).sort(([left], [right]) => left.localeCompare(right)),
              ),
            ]),
        ),
        totals: {
          files: new Set(remainingWarnings.map((warning) => warning.file)).size,
          warnings: remainingWarnings.length,
          byRule: remainingWarnings.reduce((accumulator, warning) => {
            accumulator[warning.rule] = (accumulator[warning.rule] || 0) + 1;
            return accumulator;
          }, {}),
        },
      };
      mkdirSync(dirname(baselinePath), { recursive: true });
      writeFileSync(baselinePath, `${JSON.stringify(document, null, 2)}\n`);
      console.log(
        `style guards: baseline written — ${remainingWarnings.length} frozen legacy violation(s) in ${document.totals.files} file(s), ` +
          `${parityAllowlisted} parity literal(s) excluded`,
      );
      console.log(`approval: ${approval.reviewer} ${approval.issue} — ${approval.reason}`);
      writeFileSync(
        path.join(evidenceDir, "summary.json"),
        `${JSON.stringify(
          {
            suite,
            result: "BASELINE-WRITTEN",
            ...totals,
            parityAllowlist: paritySummary,
            approval,
          },
          null,
          2,
        )}\n`,
      );
      return 0;
    }

    const failed = newViolations.length > 0 || policyProblems.length > 0;
    const summary = {
      suite,
      result: failed ? "FAIL" : "PASS",
      reportOnly,
      baseline: path.relative(root, baselinePath),
      baselineApproval: baseline.approved,
      baselineAllowed: countBaselineEntries(approvedCounts),
      parityAllowlist: paritySummary,
      totals,
      shrinkable,
      newViolations,
    };
    writeFileSync(path.join(evidenceDir, "summary.json"), `${JSON.stringify(summary, null, 2)}\n`);

    console.log(
      `style guards: ${summary.result} — ${totals.warnings} violation(s) in ${totals.files} file(s), ` +
        `${totals.baselined} baselined, ${totals.parityAllowlisted} parity-literal, ${totals.new} new`,
    );
    for (const [rule, count] of Object.entries(totals.byRule).sort((a, b) => b[1] - a[1])) {
      console.log(`  ${rule}: ${count}`);
    }
    if (parityAllowlisted > 0) {
      console.log(
        `\nparity literals allowlisted: ${parityAllowlisted} (${paritySummary.path})`,
      );
      for (const entry of paritySummary.matches) {
        if (entry.matched > 0) {
          console.log(`  ${entry.id}: ${entry.matched}/${entry.count} ${entry.file}`);
        }
      }
    }
    if (paritySummary.stale.length) {
      console.log(
        `\n${paritySummary.stale.length} parity-literal entr(ies) are stale; shrink or remove them:`,
      );
      for (const entry of paritySummary.stale) {
        console.log(
          `  ${entry.id}: expected ${entry.expected}, matched ${entry.matched} (${entry.file})`,
        );
      }
    }
    if (policyProblems.length) {
      console.error("\nparity-literal allowlist problems:");
      for (const problem of policyProblems) {
        for (const message of problem.messages) console.error(`  ${message}`);
      }
      console.error(
        "\nFix tests/style/parity-literal-allowlist.json (exact file/property/value, issue, reason, reviewBy).",
      );
    }
    if (newViolations.length) {
      console.error("\nnew component style violations:");
      for (const violation of newViolations) {
        console.error(
          `  ${violation.file}:${violation.line}:${violation.column} ${violation.rule} — ${violation.text}`,
        );
      }
      console.error(
        "\nUse an --office-* token. If the literal is an intentional upstream parity value, " +
          "add an exact entry to tests/style/parity-literal-allowlist.json " +
          "(file/rule/property/value/count, upstreamPath/commit, issue, reason, reviewBy).",
      );
    }
    if (shrinkable.length) {
      console.log(
        `\n${shrinkable.length} baseline entr(ies) can shrink; refresh with \`pnpm lint:style:baseline --approve ...\`.`,
      );
    }
    console.log(`evidence: ${path.relative(root, path.join(evidenceDir, "summary.json"))}`);
    return failed && !reportOnly ? 1 : 0;
  });
}

const invokedDirectly =
  process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (invokedDirectly) {
  main().then((code) => process.exit(code));
}

export { countByFileAndRule, lint };
