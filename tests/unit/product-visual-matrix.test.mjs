import assert from "node:assert/strict"
import { accessSync, constants, readFileSync, statSync } from "node:fs"
import path from "node:path"
import test from "node:test"
import { fileURLToPath } from "node:url"

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..")
const baseline = path.join(root, "tests/baseline/visual")
const manifest = JSON.parse(readFileSync(path.join(baseline, "manifest.json"), "utf8"))
const config = JSON.parse(readFileSync(path.join(baseline, "config.json"), "utf8"))
const verdicts = JSON.parse(readFileSync(path.join(baseline, "verdicts.json"), "utf8"))
const ci = readFileSync(path.join(root, ".github/workflows/ci.yml"), "utf8")
const scriptPath = path.join(root, "scripts/ci/capture-visual-matrix.py")

const REQUIRED_ISSUES = [13, 14, 15, 27, 28, 29, 30, 36, 44, 48, 67, 70]
const MATRIX_ROUTES = [
  "docx-viewer",
  "docx-editor",
  "docx-surface",
  "docx-parity",
  "xlsx-viewer",
  "xlsx-surface",
  "pdf-viewer",
  "pdf-surface",
  "pptx-viewer",
  "pptx-surface",
]

function fileSize(relative) {
  const absolute = path.join(root, relative)
  accessSync(absolute)
  return statSync(absolute).size
}

test("visual matrix baseline manifest is complete and self-consistent", () => {
  assert.equal(manifest.kind, "product-visual-matrix-baseline")
  assert.equal(manifest.label, "pre-ui-migration-snapshot")
  assert.deepEqual(manifest.counts, { matrix: 12, evidence: 17, total: 29 })
  assert.deepEqual(manifest.routes, MATRIX_ROUTES)
  assert.deepEqual(manifest.viewports, [
    "desktop-1440x900",
    "laptop-1280x720",
    "tablet-768x1024",
    "mobile-390x844",
  ])
  assert.deepEqual(manifest.zooms, [50, 100, 200])

  const keys = new Set()
  for (const entry of manifest.entries) {
    const key = `${entry.kind}/${entry.id}/${entry.viewport}/${entry.theme}/${entry.zoom}`
    assert.ok(!keys.has(key), `duplicate baseline entry ${key}`)
    keys.add(key)
    assert.ok(MATRIX_ROUTES.includes(entry.route.replace(/^\//, "")), `unknown route ${entry.route}`)
    if (entry.kind === "matrix") {
      assert.ok(MATRIX_ROUTES.includes(entry.id), `unknown matrix route ${entry.id}`)
    }
    assert.ok(fileSize(entry.screenshot) > 0, `missing screenshot ${entry.screenshot}`)
    assert.ok(fileSize(entry.structure) > 0, `missing structure ${entry.structure}`)
    const header = readFileSync(path.join(root, entry.screenshot)).subarray(0, 12)
    assert.equal(header.subarray(0, 4).toString("ascii"), "RIFF")
    assert.equal(header.subarray(8, 12).toString("ascii"), "WEBP")
  }

  const matrix = manifest.entries.filter((entry) => entry.kind === "matrix")
  assert.equal(matrix.length, 12)
  assert.deepEqual(
    [...new Set(matrix.map((entry) => entry.id))].sort(),
    ["docx-viewer", "pdf-viewer", "pptx-viewer", "xlsx-viewer"],
  )
  for (const entry of matrix) {
    assert.equal(entry.zoom, 100)
    assert.ok(
      (entry.viewport === "desktop-1440x900" && ["light", "dark"].includes(entry.theme)) ||
        (entry.viewport === "mobile-390x844" && entry.theme === "light"),
      `unexpected subset scenario ${entry.viewport}/${entry.theme}`,
    )
  }
})

test("visual matrix config keeps pixel diffs non-blocking and structures blocking", () => {
  assert.equal(config.structure.assertionsBlocking, true)
  assert.equal(config.pixelDiff.blocking, false)
  assert.equal(config.pixelDiff.nightly, true)
  assert.ok(config.pixelDiff.threshold > 0)
  assert.ok(config.pixelDiff.maxDiffRatio > 0 && config.pixelDiff.maxDiffRatio < 0.01)
  assert.equal(config.matrix.fullMatrixPending, true)
  assert.deepEqual(config.matrix.routes, MATRIX_ROUTES)
  assert.deepEqual(config.matrix.zooms, [50, 100, 200])
})

test("every historical rendering finding has an observed verdict", () => {
  const issues = verdicts.map((verdict) => verdict.issue).sort((a, b) => a - b)
  assert.deepEqual(issues, [...REQUIRED_ISSUES].sort((a, b) => a - b))
  const allowed = new Set(["confirmed", "refuted", "needs-another-fixture"])
  for (const verdict of verdicts) {
    assert.ok(allowed.has(verdict.verdict), `bad verdict for #${verdict.issue}`)
    assert.ok(verdict.title.length > 0)
    assert.ok(verdict.observed.length > 0)
    assert.ok(verdict.gap.length > 0)
    for (const evidence of verdict.evidence) {
      assert.ok(fileSize(path.join("tests/baseline/visual", evidence)) > 0)
    }
  }
})

test("nightly CI captures the visual matrix without a blocking pixel gate", () => {
  assert.match(ci, /Capture and compare the product visual matrix/)
  assert.match(ci, /scripts\/ci\/capture-visual-matrix\.py/)
  assert.match(ci, /--skip-build --check --output output\/visual-matrix/)
  assert.match(ci, /visual-matrix-\$\{\{ github\.sha \}\}/)
  assert.match(ci, /RUN_PYTHON_MODULES: "playwright,PIL"/)
  assert.ok(ci.includes("schedule:"), "the workflow must stay nightly/manual")
  assert.ok(!ci.includes("--fail-on-pixel"), "pixel diffs must stay non-blocking until #165")
})

test("the capture script is executable and exposes its baseline modes", () => {
  accessSync(scriptPath, constants.X_OK)
  const source = readFileSync(scriptPath, "utf8")
  for (const token of [
    "--update-baseline",
    "--check",
    "--report-only",
    "--pixel-threshold",
    "--max-diff-ratio",
    "--fail-on-pixel",
    "document.fonts.ready",
    "DETERMINISM_CSS",
  ]) {
    assert.ok(source.includes(token), `capture script is missing ${token}`)
  }
})
