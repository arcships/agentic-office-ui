import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const rootDir = resolve(dirname(fileURLToPath(import.meta.url)), "../..");
const read = (relative) => readFileSync(resolve(rootDir, relative), "utf8");
const readJson = (relative) => JSON.parse(read(relative));

const PLAYBACK_FIXTURES = "tests/fixtures/pptx/playback";
const manifest = readJson(`${PLAYBACK_FIXTURES}/manifest.json`);
const morphMethods = new Set(["explicit-name", "creation-id", "name-and-geometry", "text-and-geometry"]);
const morphConfidences = new Set(["strong", "medium", "weak"]);
const dispositions = new Set(["strict", "approximate", "static", "unparsed"]);

function assertNumber(value, label) {
  assert.equal(typeof value, "number", `${label} must be a number`);
  assert.ok(Number.isFinite(value), `${label} must be finite`);
}

function checkBaseline(entry) {
  const baseline = readJson(`${PLAYBACK_FIXTURES}/${entry.baseline}`);
  const label = entry.file;

  assert.equal(baseline.schemaVersion, 1, `${label} schemaVersion`);
  assert.equal(baseline.fixture, entry.file, `${label} fixture`);
  assert.equal(baseline.powerPointVersion, entry.powerpointVersion, `${label} powerPointVersion`);
  assert.equal(baseline.focusSlideIndex, entry.focusSlideIndex, `${label} focusSlideIndex`);
  assert.equal(typeof baseline.focus, "string", `${label} focus`);
  assert.ok(baseline.focus.length > 0, `${label} focus text`);

  const document = baseline.document;
  for (const key of ["slides", "objects", "nodes", "media", "actions"]) {
    assert.ok(Number.isInteger(document[key]) && document[key] >= 0, `${label} document.${key}`);
  }
  assert.ok(Array.isArray(document.hiddenSlides), `${label} hiddenSlides`);
  for (const index of document.hiddenSlides) {
    assert.ok(Number.isInteger(index) && index >= 0 && index < document.slides, `${label} hidden slide ${index}`);
  }

  const capability = baseline.capability;
  assert.equal(
    capability.discovered,
    capability.strict + capability.approximate + capability.static + capability.unparsed,
    `${label} capability counters`,
  );
  assert.ok(Array.isArray(capability.features), `${label} capability features`);
  const featureIds = new Set();
  for (const feature of capability.features) {
    assert.ok(typeof feature.id === "string" && feature.id.length > 0, `${label} feature id`);
    assert.equal(featureIds.has(feature.id), false, `${label} duplicate feature ${feature.id}`);
    featureIds.add(feature.id);
    assert.ok(Number.isInteger(feature.slideIndex), `${label} feature slideIndex`);
    assert.ok(dispositions.has(feature.disposition), `${label} feature disposition ${feature.disposition}`);
    if (feature.feature.startsWith("morph:")) {
      assert.ok(morphMethods.has(feature.feature.slice("morph:".length)), `${label} morph method ${feature.feature}`);
    }
  }

  const focused = baseline.focusedSlide;
  assert.equal(focused.index, baseline.focusSlideIndex, `${label} focusedSlide.index`);
  assert.equal(typeof focused.hidden, "boolean", `${label} focusedSlide.hidden`);
  assert.ok(Number.isInteger(focused.nodeCount), `${label} focusedSlide.nodeCount`);
  if (focused.transition != null) {
    assert.equal(typeof focused.transition.kind, "string", `${label} transition kind`);
    assertNumber(focused.transition.durationMs, `${label} transition durationMs`);
    assert.equal(typeof focused.transition.advanceOnClick, "boolean", `${label} transition advanceOnClick`);
  }

  assert.ok(Array.isArray(focused.media), `${label} focusedSlide.media`);
  for (const item of focused.media) {
    assert.ok(item.kind === "audio" || item.kind === "video", `${label} media kind`);
    assert.equal(typeof item.embedded, "boolean", `${label} media embedded`);
    // 自动播放是播放解析结果的一部分，基准里必须保留。
    assert.equal(typeof item.autoplay, "boolean", `${label} media ${item.id} autoplay`);
    assert.equal(typeof item.loop, "boolean", `${label} media ${item.id} loop`);
    assertNumber(item.volume, `${label} media ${item.id} volume`);
    assert.ok(item.volume >= 0 && item.volume <= 1, `${label} media ${item.id} volume range`);
    assert.ok(Array.isArray(item.bookmarks), `${label} media ${item.id} bookmarks`);
    for (const bookmark of item.bookmarks) {
      assert.equal(typeof bookmark.name, "string", `${label} bookmark name`);
      assertNumber(bookmark.timeMs, `${label} bookmark timeMs`);
    }
  }

  assert.ok(Array.isArray(focused.morphFromPrevious), `${label} morphFromPrevious`);
  const pairs = new Set();
  for (const match of focused.morphFromPrevious) {
    assert.ok(morphMethods.has(match.method), `${label} morph match method ${match.method}`);
    assert.ok(morphConfidences.has(match.confidence), `${label} morph match confidence ${match.confidence}`);
    assert.equal(typeof match.unique, "boolean", `${label} morph match unique`);
    assertNumber(match.score, `${label} morph match score`);
    assert.ok(match.score > 0 && match.score <= 1, `${label} morph match score range`);
    assert.notEqual(match.from, match.to, `${label} morph match reuses one object`);
    const pair = `${match.from}\u0000${match.to}`;
    assert.equal(pairs.has(pair), false, `${label} duplicate morph pair`);
    pairs.add(pair);
  }

  assert.ok(Array.isArray(baseline.schedule), `${label} schedule`);
  for (const group of baseline.schedule) {
    assert.equal(typeof group.key, "string", `${label} schedule key`);
    assert.ok(
      group.durationMs === "indefinite" || typeof group.durationMs === "number",
      `${label} schedule duration`,
    );
    assert.ok(Array.isArray(group.effects), `${label} schedule effects`);
    for (const effect of group.effects) {
      assert.equal(effect.triggerKey, group.key, `${label} effect triggerKey`);
      assertNumber(effect.startMs, `${label} effect startMs`);
      assert.ok(effect.endMs === "indefinite" || Number.isFinite(effect.endMs), `${label} effect endMs`);
    }
  }

  assert.ok(Array.isArray(baseline.eventSequence), `${label} eventSequence`);
  for (const event of baseline.eventSequence) {
    assert.equal(typeof event.type, "string", `${label} event type`);
    assert.notEqual(event.type, "statechange", `${label} volatile statechange event`);
    assert.notEqual(event.type, "capability", `${label} volatile capability event`);
  }

  const assertions = baseline.assertions;
  if (assertions.mediaPolicy) {
    assert.equal(typeof assertions.mediaPolicy.nativeLoop, "boolean", `${label} mediaPolicy.nativeLoop`);
    assert.equal(typeof assertions.mediaPolicy.autoplay, "boolean", `${label} mediaPolicy.autoplay`);
    assert.equal(typeof assertions.mediaPolicy.loop, "boolean", `${label} mediaPolicy.loop`);
  }
  if (assertions.mediaTracks) {
    assert.equal(assertions.mediaTracks["media-playback"], "play", `${label} media-playback track`);
    assertNumber(assertions.mediaTracks["media-volume"], `${label} media-volume track`);
  }
  return baseline;
}

test("PPTX browser baselines keep the parsed model, morph and media contract", () => {
  assert.ok(Array.isArray(manifest.entries) && manifest.entries.length >= 13, "manifest entries");
  const readyEntries = manifest.entries.filter((entry) => entry.status === "ready");
  const controlledEntries = manifest.entries.filter((entry) => entry.status === "controlled");
  assert.ok(readyEntries.length >= 13, `ready manifest entries: ${readyEntries.length}`);
  for (const entry of controlledEntries) {
    assert.ok(entry.file.endsWith(".pptx"), `${entry.file} controlled extension`);
    assert.match(entry.source, /generate-pptx-v2-fixtures\.mjs/u, `${entry.file} controlled source`);
    assert.ok(Number.isInteger(entry.focusSlideIndex), `${entry.file} controlled focusSlideIndex`);
    assert.ok(Array.isArray(entry.expectations) && entry.expectations.length > 0, `${entry.file} controlled expectations`);
    assert.ok(Array.isArray(entry.xmlMarkers) && entry.xmlMarkers.length > 0, `${entry.file} controlled xml markers`);
  }
  const fixtures = new Set();
  const mediaFixtures = [];
  const morphFixtures = [];
  for (const entry of readyEntries) {
    assert.equal(entry.status, "ready", `${entry.file} status`);
    assert.equal(entry.baseline, `${entry.file.replace(/\.pptx$/u, "")}/expected-events.json`, `${entry.file} baseline path`);
    assert.ok(!fixtures.has(entry.file), `${entry.file} duplicate manifest entry`);
    fixtures.add(entry.file);
    const baseline = checkBaseline(entry);
    if (baseline.focusedSlide.media.length > 0) mediaFixtures.push(entry.file);
    const hasMorphFeature = baseline.capability.features.some((feature) => feature.feature.startsWith("morph:"));
    if (hasMorphFeature) {
      morphFixtures.push(entry.file);
      assert.ok(
        baseline.focusedSlide.transition?.kind === "morph"
          || baseline.capability.features.some((feature) => feature.feature === "transition:morph"),
        `${entry.file} morph capability without a morph transition`,
      );
    }
  }
  // 这两类证据必须至少各有一份基准，否则浏览器捕获覆盖不到新行为。
  assert.ok(mediaFixtures.length >= 2, `media baselines: ${mediaFixtures.join(", ")}`);
  assert.ok(
    mediaFixtures.every((file) => {
      const baseline = readJson(`${PLAYBACK_FIXTURES}/${file.replace(/\.pptx$/u, "")}/expected-events.json`);
      return baseline.assertions.mediaPolicy?.autoplay === true && baseline.assertions.mediaTracks;
    }),
    "media baselines must capture autoplay and media tracks",
  );
  assert.ok(morphFixtures.length >= 1, `morph baselines: ${morphFixtures.join(", ")}`);
  const morphBaseline = readJson(
    `${PLAYBACK_FIXTURES}/${morphFixtures[0].replace(/\.pptx$/u, "")}/expected-events.json`,
  );
  assert.ok(
    morphBaseline.capability.features.some((feature) => feature.feature === "morph:explicit-name"),
    "morph baseline must record the explicit-name match method",
  );
  assert.ok(
    morphBaseline.capability.features.some((feature) => feature.feature === "transition:morph"),
    "morph baseline must record the morph transition",
  );
});

test("PPTX baseline capture and browser verification stay wired together", () => {
  const capture = read("scripts/fixtures/capture-pptx-formal-baselines.py");
  assert.match(capture, /--check/u);
  assert.match(capture, /def baseline_diff/u);
  assert.match(capture, /"mediaTracks"/u);
  assert.match(capture, /unified_diff/u);

  const probe = read("tests/blackbox/pptx_playback_model.py");
  assert.match(probe, /--morph-check/u);
  assert.match(probe, /morphWordSpans/u);
  assert.match(probe, /morphAnimatedObjects/u);
  assert.match(probe, /morphGeometryCount/u);
  assert.match(probe, /morphTextCount/u);
  assert.match(probe, /data-pptx-morph-word/u);
  assert.match(probe, /compilePptxSlideTracks/u);

  const browser = read("scripts/ci/verify-pptx-playback-browser.mjs");
  assert.match(browser, /--morph-check/u);
  assert.match(browser, /capture-pptx-formal-baselines\.py/u);
  assert.match(browser, /--check/u);

  const scripts = readJson("package.json").scripts;
  assert.match(scripts["test:pptx:baselines"], /capture-pptx-formal-baselines\.py --check/u);
});

// 无浏览器时也能验证捕获映射：用基准反推探针输出，再交给捕获脚本重建快照。
const snapshotScript = `
import importlib.util
import json
import sys
from pathlib import Path

root = Path.cwd()
module_path = root / "scripts/fixtures/capture-pptx-formal-baselines.py"
spec = importlib.util.spec_from_file_location("capture_baselines", module_path)
capture = importlib.util.module_from_spec(spec)
spec.loader.exec_module(capture)

fixtures = root / "tests/fixtures/pptx/playback"
failures = []
for file_name, fixture_spec in capture.SPECS.items():
    target = fixtures / Path(file_name).stem / "expected-events.json"
    stored = json.loads(target.read_text(encoding="utf-8"))
    raw = {
        "playbackSlides": stored["document"]["slides"],
        "objects": stored["document"]["objects"],
        "nodes": stored["document"]["nodes"],
        "media": stored["document"]["media"],
        "hiddenSlides": stored["document"]["hiddenSlides"],
        "actions": stored["document"]["actions"],
        "discovered": stored["capability"]["discovered"],
        "strict": stored["capability"]["strict"],
        "approximate": stored["capability"]["approximate"],
        "static": stored["capability"]["static"],
        "unparsed": stored["capability"]["unparsed"],
        "capabilityFeatures": stored["capability"]["features"],
        "focusedSlide": stored["focusedSlide"],
        "schedule": stored["schedule"],
        "eventSequence": stored["eventSequence"],
        **stored["assertions"],
    }
    rebuilt = capture.stable_snapshot(file_name, fixture_spec, raw)
    if rebuilt != stored:
        failures.append(file_name)
    if fixture_spec["focus"] != stored["focus"]:
        failures.append(f"{file_name} focus")
    if fixture_spec["slide"] != stored["focusSlideIndex"]:
        failures.append(f"{file_name} focusSlideIndex")
    unchanged = capture.baseline_diff(file_name, fixture_spec, dict(stored))
    if unchanged:
        failures.append(f"{file_name} diff")
    changed = {
        **raw,
        "objects": stored["document"]["objects"] + 1,
    }
    altered = capture.stable_snapshot(file_name, fixture_spec, changed)
    if not capture.baseline_diff(file_name, fixture_spec, altered):
        failures.append(f"{file_name} diff-detects-change")

if failures:
    print("FAIL " + ", ".join(failures))
    sys.exit(1)
print("PASS")
`;

test("PPTX baseline capture rebuilds every stored snapshot and detects drift", (t) => {
  const result = spawnSync("python3", ["-"], {
    cwd: rootDir,
    input: snapshotScript,
    encoding: "utf8",
  });
  if (result.error?.code === "ENOENT") {
    t.skip("python3 is not available");
    return;
  }
  assert.equal(result.status, 0, result.stderr || result.stdout);
  assert.equal(result.stdout.trim(), "PASS");
});
