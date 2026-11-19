#!/usr/bin/env node

/**
 * Open-source attribution contract (#167).
 *
 * Verifies, without weakening any other gate:
 *   1. THIRD_PARTY_NOTICES.md, the upstream index and every license copy exist
 *      and keep the exact upstream license text markers.
 *   2. The root LICENSE stays Apache-2.0.
 *   3. README.md and packages/vue-ui/README.md point at THIRD_PARTY_NOTICES.md.
 *   4. docs/upstream/extend-ui.md lists every index row in
 *      docs/upstream/attribution.json with the same target and status, and
 *      uses only the four allowed statuses.
 *   5. Every existing adapted/copied target carries the Extend header. Planned
 *      targets that do not exist yet are reported as pending, not as failures;
 *      the file must carry the header as soon as it is created.
 *   6. Every source file that already carries the Extend header is listed in
 *      the index, so copies cannot drift in without attribution.
 *   7. Upstream URLs are reachable, or have a reachable archived fallback.
 *
 * Environment:
 *   ATTRIBUTION_ROOT          repository root to verify (default: this repository; used by tests)
 *   ATTRIBUTION_EVIDENCE_DIR  evidence directory (default output/acceptance/docs-contract/attribution)
 *   ATTRIBUTION_SKIP_URLS=1   skip the network checks (offline runs only; the summary records the skip)
 *   ATTRIBUTION_URL_TIMEOUT_MS  per-request timeout (default 15000)
 */

import { existsSync, mkdirSync, readFileSync, readdirSync, statSync, writeFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";

const root = path.resolve(
  process.env.ATTRIBUTION_ROOT || path.resolve(import.meta.dirname, "../.."),
);
const evidenceDir = path.resolve(
  process.env.ATTRIBUTION_EVIDENCE_DIR ||
    path.join(root, "output", "acceptance", "docs-contract", "attribution"),
);
const skipUrls = process.env.ATTRIBUTION_SKIP_URLS === "1" || process.argv.includes("--skip-urls");
const urlTimeoutMs = Number.parseInt(process.env.ATTRIBUTION_URL_TIMEOUT_MS || "15000", 10);
const allowedStatuses = new Set(["reference-only", "adapted", "copied", "not-used"]);
const headerRequiringStatuses = new Set(["adapted", "copied"]);
const headerExtensions = new Set([".ts", ".tsx", ".vue", ".css", ".js", ".mjs", ".cjs"]);
const skipDirs = new Set(["node_modules", "dist", "output", "coverage", ".git", ".turbo", ".next"]);

const failures = [];
const checks = [];
const warnings = [];

function rel(absolute) {
  return path.relative(root, absolute).split(path.sep).join("/");
}

function read(relative) {
  return readFileSync(path.join(root, relative), "utf8");
}

function check(name, condition, details = undefined) {
  checks.push({ name, status: condition ? "PASS" : "FAIL", ...(details ? { details } : {}) });
  if (!condition) failures.push(details ? `${name}: ${details}` : name);
  return condition;
}

function warn(message) {
  warnings.push(message);
}

function walk(directory, collector) {
  let entries;
  try {
    entries = readdirSync(directory, { withFileTypes: true });
  } catch {
    return;
  }
  for (const entry of entries) {
    if (entry.isDirectory()) {
      if (skipDirs.has(entry.name)) continue;
      walk(path.join(directory, entry.name), collector);
    } else if (entry.isFile()) {
      collector(path.join(directory, entry.name));
    }
  }
}

function firstLines(content, count = 15) {
  return content.split(/\r?\n/).slice(0, count).join("\n");
}

function hasAttributionHeader(content, header) {
  const head = firstLines(content);
  const markers = (header.markers || []).map((marker) =>
    typeof marker === "string" ? { name: marker, pattern: marker } : marker,
  );
  const hasMarkers = markers.every((marker) => new RegExp(marker.pattern, "i").test(head));
  if (!hasMarkers) return false;
  const markerLine = head
    .split(/\r?\n/)
    .find((line) => new RegExp(markers[0].pattern, "i").test(line));
  if (!markerLine) return false;
  const trimmed = markerLine.trimStart();
  return header.commentPrefixes.some((prefix) => trimmed.startsWith(prefix));
}

function stripCell(cell) {
  return cell.replace(/`/g, "").replace(/\s+/g, " ").trim();
}

function parseTables(markdown) {
  const rows = [];
  for (const line of markdown.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed.startsWith("|")) continue;
    const cells = trimmed.replace(/^\|/, "").replace(/\|$/, "").split("|").map(stripCell);
    if (cells.length < 5) continue;
    if (cells.every((cell) => /^:?-{3,}:?$/.test(cell) || cell === "")) continue;
    if (cells[0] === "Extend 文件" || cells[3] === "状态") continue;
    rows.push(cells);
  }
  return rows;
}

function normalizeTarget(value) {
  return value.replace(/^\.\//, "").replace(/\/+$/, "");
}

function targetMatches(cell, target) {
  const normalizedCell = normalizeTarget(cell);
  const normalizedTarget = normalizeTarget(target);
  if (!normalizedTarget) return cell === "" || cell === "—" || cell === "-";
  return normalizedCell === normalizedTarget || normalizedCell.includes(normalizedTarget);
}

// ---------------------------------------------------------------------------
// 1. Manifest
// ---------------------------------------------------------------------------

if (!existsSync(path.join(root, "docs/upstream/attribution.json"))) {
  console.error("FAIL: docs/upstream/attribution.json is missing");
  process.exit(1);
}

let manifest;
try {
  manifest = JSON.parse(read("docs/upstream/attribution.json"));
} catch (error) {
  console.error(`FAIL: docs/upstream/attribution.json is not valid JSON: ${error.message}`);
  process.exit(1);
}

check("manifest schema version", manifest.schemaVersion === 1, String(manifest.schemaVersion));
check("manifest pins the Extend commit", /^[0-9a-f]{40}$/.test(manifest.extend?.commit || ""), manifest.extend?.commit);
check("manifest declares the shared header", typeof manifest.header?.required === "string" && manifest.header.required.includes(manifest.extend?.commit));
check(
  "manifest declares header markers",
  Array.isArray(manifest.header?.markers) && manifest.header.markers.length >= 2,
);
check(
  "manifest lists the required upstream projects",
  ["extend-ui", "shadcn-vue", "reka-ui", "lucide", "geist"].every((id) =>
    (manifest.upstreams || []).some((upstream) => upstream.id === id),
  ),
);
check("manifest index is not empty", Array.isArray(manifest.index) && manifest.index.length > 0);
check("manifest URL checks are declared", Array.isArray(manifest.urlChecks) && manifest.urlChecks.length > 0);

for (const entry of manifest.index || []) {
  check(
    `index status ${entry.upstreamFile} -> ${entry.target || "(not used)"}`,
    allowedStatuses.has(entry.status),
    entry.status,
  );
  check(
    `index issue reference ${entry.upstreamFile} -> ${entry.target || "(not used)"}`,
    Array.isArray(entry.issues) && entry.issues.length > 0,
  );
}

// ---------------------------------------------------------------------------
// 2. Required files and license text
// ---------------------------------------------------------------------------

const requiredFiles = [
  "THIRD_PARTY_NOTICES.md",
  "docs/upstream/extend-ui.md",
  "LICENSE",
  "README.md",
  "packages/vue-ui/README.md",
];
for (const file of requiredFiles) {
  check(`required file ${file}`, existsSync(path.join(root, file)));
}

const rootLicense = existsSync(path.join(root, "LICENSE")) ? read("LICENSE") : "";
check(
  "root LICENSE stays Apache-2.0",
  rootLicense.includes("Apache License") &&
    rootLicense.includes("Version 2.0, January 2004") &&
    !rootLicense.includes("CrowdView Inc, dba Extend"),
);

const notices = existsSync(path.join(root, "THIRD_PARTY_NOTICES.md"))
  ? read("THIRD_PARTY_NOTICES.md")
  : "";
check("THIRD_PARTY_NOTICES.md is not empty", notices.trim().length > 200);

for (const upstream of manifest.upstreams || []) {
  const licensePath = path.join(root, upstream.licenseFile);
  const exists = existsSync(licensePath);
  check(`license file ${upstream.id}`, exists, upstream.licenseFile);
  if (!exists) continue;
  const text = readFileSync(licensePath, "utf8");
  for (const marker of upstream.licenseMarkers || []) {
    check(
      `license text ${upstream.id} contains ${JSON.stringify(marker)}`,
      text.includes(marker),
      upstream.licenseFile,
    );
  }
  check(
    `notice covers ${upstream.id}`,
    notices.includes(upstream.name) || (upstream.aliases || []).some((alias) => notices.includes(alias)),
  );
  check(`notice names the license for ${upstream.id}`, notices.includes(upstream.license), upstream.license);
  check(`notice keeps the copyright for ${upstream.id}`, notices.includes(upstream.copyright));
  check(`notice links the license copy for ${upstream.id}`, notices.includes(upstream.licenseFile));
  check(`notice links the source for ${upstream.id}`, notices.includes(upstream.sourceUrl));
}

check(
  "THIRD_PARTY_NOTICES.md carries the full Extend MIT text",
  ["MIT License", "CrowdView Inc, dba Extend", "Portions Copyright (c) 2023 shadcn", "THE SOFTWARE IS PROVIDED \"AS IS\""].every(
    (marker) => notices.includes(marker),
  ),
);

// ---------------------------------------------------------------------------
// 3. README pointers
// ---------------------------------------------------------------------------

for (const file of ["README.md", "packages/vue-ui/README.md"]) {
  if (!existsSync(path.join(root, file))) continue;
  const content = read(file);
  check(
    `attribution pointer ${file}`,
    /open-source attribution/i.test(content) && content.includes("THIRD_PARTY_NOTICES.md"),
  );
}

// ---------------------------------------------------------------------------
// 4. Human index vs manifest
// ---------------------------------------------------------------------------

const indexDoc = existsSync(path.join(root, "docs/upstream/extend-ui.md"))
  ? read("docs/upstream/extend-ui.md")
  : "";
check("index doc pins the Extend commit", indexDoc.includes(manifest.extend.commit));
check("index doc links the Extend repository", indexDoc.includes(manifest.extend.repository));
const extendLicenseUrl = (manifest.upstreams || []).find((item) => item.id === "extend-ui")?.licenseUrl;
check("index doc links the pinned Extend license", !extendLicenseUrl || indexDoc.includes(extendLicenseUrl));
check(
  "index doc states the header format",
  indexDoc.includes(manifest.header.required),
);

const docRows = parseTables(indexDoc);
for (const row of docRows) {
  const statusCell = row[3];
  if (!allowedStatuses.has(statusCell)) {
    check(`index doc status vocabulary ${JSON.stringify(row[0])}`, false, statusCell);
  }
}
for (const entry of manifest.index || []) {
  const match = docRows.some(
    (row) =>
      row[0] === entry.upstreamFile &&
      (entry.status === "not-used" ? row[2] === "" || row[2] === "—" || row[2] === "-" : targetMatches(row[2], entry.target)) &&
      row[3] === entry.status,
  );
  check(
    `index doc row ${entry.upstreamFile} -> ${entry.target || "(not used)"} [${entry.status}]`,
    match,
    "docs/upstream/extend-ui.md is out of sync with docs/upstream/attribution.json",
  );
}

check("docs index links the upstream index", read("docs/INDEX.md").includes("upstream/extend-ui.md"));

// ---------------------------------------------------------------------------
// 5. Headers on existing adapted/copied targets + pending adapter report
// ---------------------------------------------------------------------------

const pending = [];
const headerChecked = [];
const coveredTargets = [];

function targetFiles(entry) {
  const excluded = new Set(entry.exclude || []);
  if (entry.target.endsWith("/")) {
    const directory = path.join(root, entry.target);
    if (!existsSync(directory) || !statSync(directory).isDirectory()) return [];
    const files = [];
    walk(directory, (file) => {
      if (excluded.has(path.basename(file)) || excluded.has(rel(file))) return;
      if (headerExtensions.has(path.extname(file).toLowerCase())) files.push(file);
    });
    return files;
  }
  const file = path.join(root, entry.target);
  return existsSync(file) ? [file] : [];
}

for (const entry of manifest.index || []) {
  if (!headerRequiringStatuses.has(entry.status) || !entry.target) continue;
  coveredTargets.push(entry.target);
  const files = targetFiles(entry);
  if (files.length === 0) {
    pending.push({ target: entry.target, status: entry.status, issues: entry.issues });
    continue;
  }
  for (const file of files) {
    // A file can match both a file target and a directory guard; check it once.
    if (headerChecked.includes(rel(file))) continue;
    const content = readFileSync(file, "utf8");
    const ok = hasAttributionHeader(content, manifest.header);
    headerChecked.push(rel(file));
    check(
      `attribution header ${rel(file)}`,
      ok,
      `expected: ${manifest.header.required}`,
    );
  }
}

function isCoveredByIndex(file) {
  const relative = rel(file);
  return coveredTargets.some((target) =>
    target.endsWith("/") ? relative.startsWith(normalizeTarget(target) + "/") : normalizeTarget(target) === relative,
  );
}

// ---------------------------------------------------------------------------
// 6. Reverse check: no unindexed file carries the Extend header
// ---------------------------------------------------------------------------

const marker = manifest.extend.commit;
const scanRoots = ["packages", "apps", "tools"];
const unindexed = [];
for (const directory of scanRoots) {
  const absolute = path.join(root, directory);
  if (!existsSync(absolute)) continue;
  walk(absolute, (file) => {
    if (!headerExtensions.has(path.extname(file).toLowerCase())) return;
    const content = readFileSync(file, "utf8");
    if (content.includes(marker) && !isCoveredByIndex(file)) unindexed.push(rel(file));
  });
}
check(
  "every file with an Extend header is in the index",
  unindexed.length === 0,
  unindexed.join(", "),
);

// A vendored snapshot must ship the MIT notice with the copies, not only a header.
const vendorRoot = path.join(root, "tools/extend-reference/src/vendor");
if (existsSync(vendorRoot)) {
  const candidates = [
    "tools/extend-reference/src/vendor/NOTICE.md",
    "tools/extend-reference/UPSTREAM.md",
    "tools/extend-reference/LICENSE-extend-ui.txt",
    "tools/extend-reference/LICENSE",
  ].filter((file) => existsSync(path.join(root, file)));
  const noticeText = candidates.map((file) => read(file)).join("\n");
  check(
    "vendored reference snapshot keeps the Extend MIT notice",
    candidates.length > 0 &&
      /\bMIT\b/i.test(noticeText) &&
      noticeText.includes("CrowdView Inc, dba Extend") &&
      (noticeText.includes("THE SOFTWARE IS PROVIDED") ||
        noticeText.includes("docs/upstream/licenses/extend-ui-MIT.txt")),
    "tools/extend-reference/src/vendor needs NOTICE.md or a LICENSE file with the Extend MIT text",
  );
}

// ---------------------------------------------------------------------------
// 7. Upstream URL reachability (or archived fallback)
// ---------------------------------------------------------------------------

async function probe(url) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), urlTimeoutMs);
  try {
    let response = await fetch(url, {
      method: "HEAD",
      redirect: "follow",
      signal: controller.signal,
      headers: { "user-agent": "agentic-office-ui attribution check" },
    });
    if ([403, 405].includes(response.status)) {
      response = await fetch(url, {
        method: "GET",
        redirect: "follow",
        signal: controller.signal,
        headers: { "user-agent": "agentic-office-ui attribution check", range: "bytes=0-1024" },
      });
    }
    return { ok: response.status < 400, status: response.status };
  } catch (error) {
    return { ok: false, error: error.name === "AbortError" ? `timeout after ${urlTimeoutMs}ms` : error.message };
  } finally {
    clearTimeout(timer);
  }
}

async function probeWithRetry(url) {
  let last = await probe(url);
  if (!last.ok) last = await probe(url);
  return last;
}

const urlResults = [];
if (skipUrls) {
  warn("URL reachability check skipped because ATTRIBUTION_SKIP_URLS=1");
} else {
  const queue = [...(manifest.urlChecks || [])];
  const concurrency = 6;
  async function worker() {
    while (queue.length > 0) {
      const item = queue.shift();
      const primary = await probeWithRetry(item.url);
      if (primary.ok) {
        urlResults.push({ url: item.url, status: "PASS", httpStatus: primary.status });
        continue;
      }
      if (item.archivedUrl) {
        const fallback = await probeWithRetry(item.archivedUrl);
        if (fallback.ok) {
          urlResults.push({
            url: item.url,
            status: "ARCHIVED",
            httpStatus: primary.status ?? primary.error,
            archivedUrl: item.archivedUrl,
          });
          warn(`${item.url} unreachable; archived snapshot verified at ${item.archivedUrl}`);
          continue;
        }
      }
      urlResults.push({
        url: item.url,
        status: "FAIL",
        httpStatus: primary.status ?? primary.error,
        archivedUrl: item.archivedUrl ?? null,
      });
      check(
        `upstream URL reachable ${item.url}`,
        false,
        `${primary.status ?? primary.error}${item.archivedUrl ? `; archived fallback also failed` : "; no archived URL recorded"}`,
      );
    }
  }
  await Promise.all(Array.from({ length: concurrency }, worker));
}

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------

mkdirSync(evidenceDir, { recursive: true });
const summary = {
  suite: "ATTRIBUTION-CONTRACT",
  result: failures.length ? "FAIL" : "PASS",
  extendCommit: manifest.extend.commit,
  checkedAt: new Date().toISOString(),
  checks,
  failures,
  warnings,
  pendingTargets: pending,
  headerChecked,
  urlChecks: urlResults,
  urlChecksSkipped: skipUrls,
};
writeFileSync(path.join(evidenceDir, "summary.json"), `${JSON.stringify(summary, null, 2)}\n`);

const headerCheckedText = headerChecked.length > 0 ? `${headerChecked.length} adapted/copied file(s)` : "no adapted/copied file exists yet";
console.log(
  `attribution: ${headerCheckedText}, ${pending.length} planned target(s) pending, ${checks.length} checks at ${rel(evidenceDir)}`,
);
if (pending.length > 0) {
  for (const item of pending) {
    console.log(`  pending ${item.status}: ${item.target} (${item.issues.join(", ")})`);
  }
}
for (const message of warnings) console.warn(`  warning: ${message}`);
if (failures.length > 0) {
  console.error(`\nFAIL: ${failures.length} attribution check(s) failed`);
  for (const failure of failures) console.error(`- ${failure}`);
  process.exit(1);
}
console.log("PASS: attribution contract holds");
