#!/usr/bin/env node

/**
 * Component CSS token guards (#165).
 *
 * Runs stylelint (`stylelint.config.mjs`) over every component style in
 * `packages/*\/src` and compares the result with the frozen baseline in
 * `tests/style/component-style-baseline.json`.
 *
 * - A new violation (a new file, or more occurrences than the baseline holds)
 *   fails the run and is printed with `file:line:column`.
 * - Baseline counts are per file + rule. When a file already freezes N
 *   occurrences of a rule, the N reported lines are treated as frozen, so a
 *   new occurrence of that same rule in the same file fails the run but may be
 *   printed on top of an existing line. Files without a baseline entry are
 *   always reported exactly.
 * - Fixing violations never fails; the summary reports the baseline entries
 *   that can shrink.
 * - Growing the baseline is an approval, not a `--update`:
 *
 *     pnpm lint:style:baseline --approve --reviewer <handle> \
 *       --issue "#nnn" --reason <why>
 *
 * The rules are never disabled for a directory: the token files are exempt in
 * `stylelint.config.mjs` because defining raw values is their whole job.
 */
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import path from "node:path";
import process from "node:process";
import stylelint from "stylelint";

const root = path.resolve(import.meta.dirname, "../..");
const baselinePath = path.join(root, "tests/style/component-style-baseline.json");
const evidenceDir = path.resolve(
  process.env.STYLE_GUARD_EVIDENCE_DIR || path.join(root, "output/style-guard"),
);
const files = ["packages/*/src/**/*.{vue,css}"];
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

function loadBaseline() {
  try {
    return JSON.parse(readFileSync(baselinePath, "utf8"));
  } catch (error) {
    if (error.code === "ENOENT") return { counts: {}, approved: null, totals: {} };
    throw error;
  }
}

function requireApproval(argv) {
  if (!argv.includes("--approve")) {
    console.error(
      [
        "lint-styles: growing the style baseline requires an approval.",
        "",
        "  pnpm lint:style:baseline --approve --reviewer <github-handle> \\",
        "    --issue <#nnn> --reason <why>",
        "",
        "Fix the violation instead whenever the value can come from",
        "packages/vue-ui/src/styles/tokens.css.",
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
    const counts = countByFileAndRule(warnings);
    const baseline = loadBaseline();
    const approvedCounts = baseline.counts || {};

    const newViolations = [];
    for (const [file, rules] of Object.entries(counts)) {
      for (const [rule, count] of Object.entries(rules)) {
        const allowed = approvedCounts[file]?.[rule] || 0;
        if (count > allowed) {
          const fresh = warnings.filter(
            (warning) => warning.file === file && warning.rule === rule,
          );
          newViolations.push(...fresh.slice(allowed));
        }
      }
    }

    const shrinkable = [];
    for (const [file, rules] of Object.entries(approvedCounts)) {
      for (const [rule, allowed] of Object.entries(rules)) {
        const actual = counts[file]?.[rule] || 0;
        if (actual < allowed) shrinkable.push({ file, rule, allowed, actual });
      }
    }

    const totals = {
      files: new Set(warnings.map((warning) => warning.file)).size,
      warnings: warnings.length,
      byRule: warnings.reduce((accumulator, warning) => {
        accumulator[warning.rule] = (accumulator[warning.rule] || 0) + 1;
        return accumulator;
      }, {}),
      baselined: Object.values(approvedCounts).reduce(
        (total, rules) => total + Object.values(rules).reduce((sum, value) => sum + value, 0),
        0,
      ),
      new: newViolations.length,
    };

    mkdirSync(evidenceDir, { recursive: true });

    if (writeBaseline) {
      const approval = requireApproval(argv);
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
        totals: { files: totals.files, warnings: totals.warnings, byRule: totals.byRule },
      };
      mkdirSync(dirname(baselinePath), { recursive: true });
      writeFileSync(baselinePath, `${JSON.stringify(document, null, 2)}\n`);
      console.log(
        `style guards: baseline written — ${totals.warnings} frozen violation(s) in ${totals.files} file(s)`,
      );
      console.log(`approval: ${approval.reviewer} ${approval.issue} — ${approval.reason}`);
      writeFileSync(
        path.join(evidenceDir, "summary.json"),
        `${JSON.stringify({ suite, result: "BASELINE-WRITTEN", ...totals, approval }, null, 2)}\n`,
      );
      return 0;
    }

    const summary = {
      suite,
      result: newViolations.length ? "FAIL" : "PASS",
      reportOnly,
      baseline: path.relative(root, baselinePath),
      baselineApproval: baseline.approved,
      totals,
      shrinkable,
      newViolations,
    };
    writeFileSync(path.join(evidenceDir, "summary.json"), `${JSON.stringify(summary, null, 2)}\n`);

    console.log(
      `style guards: ${summary.result} — ${totals.warnings} violation(s) in ${totals.files} file(s), ` +
        `${totals.baselined} baselined, ${totals.new} new`,
    );
    for (const [rule, count] of Object.entries(totals.byRule).sort((a, b) => b[1] - a[1])) {
      console.log(`  ${rule}: ${count}`);
    }
    if (newViolations.length) {
      console.error("\nnew component style violations:");
      for (const violation of newViolations) {
        console.error(
          `  ${violation.file}:${violation.line}:${violation.column} ${violation.rule} — ${violation.text}`,
        );
      }
      console.error(
        "\nUse an --office-* token, or approve the debt with `pnpm lint:style:baseline --approve ...`.",
      );
    }
    if (shrinkable.length) {
      console.log(
        `\n${shrinkable.length} baseline entr(ies) can shrink; refresh with \`pnpm lint:style:baseline --approve ...\`.`,
      );
    }
    console.log(`evidence: ${path.relative(root, path.join(evidenceDir, "summary.json"))}`);
    return newViolations.length && !reportOnly ? 1 : 0;
  });
}

main().then((code) => process.exit(code));
