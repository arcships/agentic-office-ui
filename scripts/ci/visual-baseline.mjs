#!/usr/bin/env node

/**
 * Committed visual baselines for the Extend parity harness (#165).
 *
 * The baselines live in `tests/visual/extend/` and are named
 * `<surface>/<state>/<theme>-<viewport>.png`. They capture the *subject*
 * (this repository's Vue components) so a PR that breaks a toolbar's spacing
 * or colour fails even when the pinned reference app is unchanged.
 *
 * Updating a baseline is an approval, not a `--update` flag:
 *
 *   pnpm ui:baseline status
 *   pnpm ui:baseline capture --only button \
 *     --approve --reviewer <github-handle> --issue "#165" \
 *     --reason "OfficeButton radius now uses --office-radius-md"
 *   pnpm ui:baseline verify --require-approval
 *
 * `capture` refuses to write anything without `--approve`, a reviewer, an
 * issue reference and a reason. Every capture records a `treeSha256` covering
 * the whole baseline tree; `verify` fails when the committed PNGs no longer
 * match their recorded hashes or when no approval covers the current tree, so
 * hand-edited or casually refreshed images cannot land in CI.
 */
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import {
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  statSync,
  writeFileSync,
} from "node:fs";
import path from "node:path";
import process from "node:process";

const root = path.resolve(import.meta.dirname, "../..");
const pixelManifestPath = path.join(root, "tools/parity-fixtures/manifest.json");
const pixelManifest = JSON.parse(readFileSync(pixelManifestPath, "utf8"));
const baselineConfig = pixelManifest.baselines;
const baselineDir = path.join(root, baselineConfig.directory);
const baselineManifestPath = path.join(root, baselineConfig.manifest);
const approvalsPath = path.join(baselineDir, "APPROVALS.md");

const blockingSurfaces = pixelManifest.surfaces.filter(
  (surface) => surface.gateStatus === "blocking" && surface.states?.length,
);

function fail(message) {
  console.error(`visual-baseline: ${message}`);
  process.exit(1);
}

function sha256(buffer) {
  return createHash("sha256").update(buffer).digest("hex");
}

/** PNG IHDR: signature(8) + length(4) + "IHDR"(4) + width(4) + height(4). */
function pngSize(buffer) {
  if (buffer.length < 24 || buffer.toString("ascii", 12, 16) !== "IHDR") {
    fail("baseline is not a PNG (IHDR chunk missing)");
  }
  return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) };
}

function relative(absolute) {
  return path.relative(baselineDir, absolute).split(path.sep).join("/");
}

function walk(directory) {
  if (!existsSync(directory)) return [];
  const files = [];
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const absolute = path.join(directory, entry.name);
    if (entry.isDirectory()) files.push(...walk(absolute));
    else if (entry.isFile() && entry.name.endsWith(".png")) files.push(absolute);
  }
  return files.sort();
}

/** `<surface>/<state>/<theme>-<viewport>.png` */
function parseName(relativePath) {
  const match = /^([^/]+)\/([^/]+)\/([a-z]+)-([0-9]+x[0-9]+)\.png$/.exec(relativePath);
  if (!match) return null;
  return {
    surface: match[1],
    state: match[2],
    theme: match[3],
    viewport: match[4],
  };
}

function collectFiles() {
  const files = [];
  for (const absolute of walk(baselineDir)) {
    const relativePath = relative(absolute);
    const parsed = parseName(relativePath);
    if (!parsed) fail(`baseline filename does not follow <surface>/<state>/<theme>-<viewport>.png: ${relativePath}`);
    const buffer = readFileSync(absolute);
    const { width, height } = pngSize(buffer);
    files.push({
      path: relativePath,
      ...parsed,
      width,
      height,
      bytes: statSync(absolute).size,
      sha256: sha256(buffer),
    });
  }
  return files;
}

function treeHash(files) {
  const lines = files
    .map((file) => `${file.path} ${file.sha256}`)
    .sort()
    .join("\n");
  return sha256(`${lines}\n`);
}

function expectedFiles() {
  const expected = new Set();
  for (const surface of blockingSurfaces) {
    for (const state of surface.states) {
      for (const theme of baselineConfig.themes) {
        for (const viewport of baselineConfig.viewports) {
          expected.add(`${surface.id}/${state.id}/${theme}-${viewport}.png`);
        }
      }
    }
  }
  return expected;
}

function readManifest() {
  if (!existsSync(baselineManifestPath)) return null;
  return JSON.parse(readFileSync(baselineManifestPath, "utf8"));
}

function loadGitHead() {
  const result = spawnSync("git", ["rev-parse", "HEAD"], { cwd: root, encoding: "utf8" });
  return result.status === 0 ? result.stdout.trim() : "no-git-head";
}

function writeApprovals(manifest) {
  const lines = [
    "# Visual baseline approvals",
    "",
    "Every capture of `tests/visual/extend/**` must be recorded here with a",
    "reviewer, an issue and a reason. `pnpm ui:baseline verify --require-approval`",
    "fails when the committed tree is not covered by the most recent record.",
    "",
    "| Approved | Reviewer | Issue | Reason | Files | Tree |",
    "|---|---|---|---|---|---|",
  ];
  for (const approval of manifest.approvals) {
    lines.push(
      `| ${approval.date} | ${approval.reviewer} | ${approval.issue} | ${approval.reason.replaceAll("|", "\\|")} | ${approval.files} | \`${approval.treeSha256.slice(0, 12)}\` |`,
    );
  }
  lines.push(
    "",
    "`pnpm ui:baseline capture --only <surface> --approve --reviewer <handle>`",
    "`--issue <#nnn> --reason <why>` is the only supported way to add a row.",
    "",
  );
  writeFileSync(approvalsPath, lines.join("\n"));
}

function statusCommand() {
  const files = collectFiles();
  const manifest = readManifest();
  const expected = expectedFiles();
  const present = new Set(files.map((file) => file.path));
  const missing = [...expected].filter((file) => !present.has(file)).sort();
  const extra = files.map((file) => file.path).filter((file) => !expected.has(file));
  console.log(`directory: ${path.relative(root, baselineDir)}`);
  console.log(`files: ${files.length} (${(files.reduce((total, file) => total + file.bytes, 0) / 1024 / 1024).toFixed(2)} MiB)`);
  console.log(`expected: ${expected.size} for ${blockingSurfaces.length} gate-blocking surface(s)`);
  if (missing.length) console.log(`missing baselines (${missing.length}):\n  ${missing.slice(0, 10).join("\n  ")}`);
  if (extra.length) console.log(`unexpected baselines (${extra.length}):\n  ${extra.slice(0, 10).join("\n  ")}`);
  if (!manifest) {
    console.log("manifest: missing (run pnpm ui:baseline capture with --approve)");
    return missing.length || extra.length ? 1 : 0;
  }
  const currentTree = treeHash(files);
  const matches = manifest.treeSha256 === currentTree;
  console.log(`manifest: ${matches ? "in sync" : "STALE"} (tree ${currentTree.slice(0, 12)})`);
  const latest = manifest.approvals?.at(-1);
  console.log(
    `approval: ${latest ? `${latest.date} ${latest.reviewer} ${latest.issue} ${latest.treeSha256.slice(0, 12)}` : "none"}`,
  );
  return missing.length || extra.length || !matches ? 1 : 0;
}

function parseFlag(argv, name) {
  const index = argv.indexOf(name);
  if (index === -1) return undefined;
  const value = argv[index + 1];
  if (value === undefined || value.startsWith("--")) fail(`${name} needs a value`);
  return value;
}

function captureCommand(argv) {
  const approve = argv.includes("--approve");
  const reviewer = parseFlag(argv, "--reviewer");
  const issue = parseFlag(argv, "--issue");
  const reason = parseFlag(argv, "--reason");
  const only = parseFlag(argv, "--only");
  const passthrough = [];
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (["--approve", "--reviewer", "--issue", "--reason", "--only"].includes(arg)) {
      index += arg === "--approve" ? 0 : 1;
      continue;
    }
    if (arg.startsWith("--viewports") || arg.startsWith("--themes")) {
      passthrough.push(arg);
      const next = argv[index + 1];
      if (next && !next.startsWith("--")) {
        passthrough.push(next);
        index += 1;
      }
      continue;
    }
    passthrough.push(arg);
  }

  if (!approve) {
    fail(
      [
        "writing baselines requires an approval.",
        "",
        "  pnpm ui:baseline capture --only <surface> --approve \\",
        "    --reviewer <github-handle> --issue <#nnn> --reason <why>",
        "",
        "Then review the PNG diff in the pull request; `pnpm ui:baseline verify`",
        "enforces the recorded tree hash. See tests/visual/extend/README.md.",
      ].join("\n"),
    );
  }
  if (!reviewer || !issue || !reason) {
    fail("--approve needs --reviewer, --issue and --reason");
  }
  if (reason.trim().length < 8) {
    fail("--reason must explain the visual change (at least 8 characters)");
  }
  if (!issue.startsWith("#") && !/^https?:/.test(issue)) {
    fail(`--issue must be a GitHub issue reference like #165 (received ${issue})`);
  }

  const surfaces = only
    ? only.split(",").map((value) => value.trim()).filter(Boolean)
    : blockingSurfaces.map((surface) => surface.id);

  console.log(
    `▸ capture ${surfaces.length} surface(s) × ${baselineConfig.viewports.join(",")} × ${baselineConfig.themes.join(",")}`,
  );
  const args = [
    "scripts/ci/pixel-parity.mjs",
    "--write-baseline",
    "--only",
    surfaces.join(","),
    "--viewports",
    baselineConfig.viewports.join(","),
    "--themes",
    baselineConfig.themes.join(","),
    ...passthrough,
  ];
  const result = spawnSync(process.execPath, args, { cwd: root, stdio: "inherit" });
  if (result.status !== 0) fail(`parity capture failed with exit code ${result.status}`);

  const files = collectFiles();
  const expected = expectedFiles();
  const present = new Set(files.map((file) => file.path));
  const missing = [...expected].filter((file) => !present.has(file));
  if (missing.length) {
    fail(`capture left ${missing.length} expected baseline(s) missing, e.g. ${missing.slice(0, 5).join(", ")}`);
  }

  const previous = readManifest();
  const previousHashes = new Map((previous?.files || []).map((file) => [file.path, file.sha256]));
  const added = files.filter((file) => !previousHashes.has(file.path)).length;
  const updated = files.filter(
    (file) => previousHashes.has(file.path) && previousHashes.get(file.path) !== file.sha256,
  ).length;
  const removed = [...previousHashes.keys()].filter((file) => !present.has(file)).length;
  const currentTree = treeHash(files);

  const manifest = {
    schemaVersion: 1,
    harness: "extend-ui-pixel-parity",
    referenceCommit: pixelManifest.referenceCommit,
    generatedBy: "node scripts/ci/visual-baseline.mjs capture",
    naming: baselineConfig.naming,
    deviceScaleFactor: pixelManifest.deviceScaleFactor,
    capturePaddingPx: baselineConfig.capturePaddingPx,
    viewports: baselineConfig.viewports,
    themes: baselineConfig.themes,
    surfaces: Object.fromEntries(
      blockingSurfaces
        .filter((surface) => surfaces.includes(surface.id))
        .map((surface) => [
          surface.id,
          {
            subjectComponent: surface.subjectComponent,
            states: surface.states.map((state) => state.id),
            files: files.filter((file) => file.surface === surface.id).length,
          },
        ]),
    ),
    treeSha256: currentTree,
    files,
    approvals: [
      ...(previous?.approvals || []),
      {
        date: new Date().toISOString().slice(0, 10),
        reviewer,
        issue,
        reason: reason.trim(),
        treeSha256: currentTree,
        files: files.length,
        added,
        updated,
        removed,
        gitHead: loadGitHead(),
      },
    ],
  };
  mkdirSync(baselineDir, { recursive: true });
  writeFileSync(baselineManifestPath, `${JSON.stringify(manifest, null, 2)}\n`);
  writeApprovals(manifest);
  console.log(
    `▸ baselines: ${files.length} files (${added} added, ${updated} updated, ${removed} removed) — tree ${currentTree.slice(0, 12)}`,
  );
  console.log(`▸ approval: ${reviewer} ${issue} — ${reason.trim()}`);
  return 0;
}

function verifyCommand(argv) {
  const requireApproval = argv.includes("--require-approval");
  const manifest = readManifest();
  if (!manifest) fail(`missing ${path.relative(root, baselineManifestPath)}`);
  const problems = [];

  if (manifest.viewports?.join(",") !== baselineConfig.viewports.join(",")) {
    problems.push("manifest viewports differ from tools/parity-fixtures/manifest.json");
  }
  if (manifest.themes?.join(",") !== baselineConfig.themes.join(",")) {
    problems.push("manifest themes differ from tools/parity-fixtures/manifest.json");
  }
  if (manifest.capturePaddingPx !== baselineConfig.capturePaddingPx) {
    problems.push("manifest capturePaddingPx differs from tools/parity-fixtures/manifest.json");
  }

  const recorded = new Map((manifest.files || []).map((file) => [file.path, file.sha256]));
  const actual = collectFiles();
  const actualByPath = new Map(actual.map((file) => [file.path, file.sha256]));
  for (const [file, hash] of recorded) {
    if (!actualByPath.has(file)) problems.push(`baseline missing on disk: ${file}`);
    else if (actualByPath.get(file) !== hash) problems.push(`baseline differs from the approved hash: ${file}`);
  }
  for (const file of actualByPath.keys()) {
    if (!recorded.has(file)) problems.push(`baseline file is not recorded in the manifest: ${file}`);
  }
  const expected = expectedFiles();
  for (const file of expected) {
    if (!recorded.has(file)) problems.push(`gate-blocking surface has no committed baseline: ${file}`);
  }
  const currentTree = treeHash(actual);
  if (manifest.treeSha256 !== currentTree) {
    problems.push(`tree hash mismatch: manifest ${manifest.treeSha256.slice(0, 12)} vs disk ${currentTree.slice(0, 12)}`);
  }
  if (requireApproval) {
    const approvals = manifest.approvals || [];
    if (!approvals.length) problems.push("no approval record");
    else {
      const latest = approvals.at(-1);
      if (latest.treeSha256 !== currentTree) {
        problems.push(
          `the current baseline tree is not covered by an approval (latest ${latest.treeSha256.slice(0, 12)} vs ${currentTree.slice(0, 12)})`,
        );
      }
      for (const approval of approvals) {
        if (!approval.reviewer || !approval.issue || !approval.reason) {
          problems.push(`approval ${approval.date} is missing reviewer/issue/reason`);
        }
      }
    }
  }

  if (problems.length) {
    for (const problem of problems) console.error(`- ${problem}`);
    fail(`${problems.length} baseline contract problem(s)`);
  }
  console.log(
    `visual baselines: PASS — ${actual.length} files, tree ${currentTree.slice(0, 12)}, ${(manifest.approvals || []).length} approval(s)`,
  );
  return 0;
}

function main() {
  const [command = "status", ...rest] = process.argv.slice(2);
  if (command === "status") process.exit(statusCommand());
  if (command === "capture") process.exit(captureCommand(rest));
  if (command === "verify") process.exit(verifyCommand(rest));
  fail(`unknown command ${command}; expected status|capture|verify`);
}

main();
