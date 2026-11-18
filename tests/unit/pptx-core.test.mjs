import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";

import { importFromDemo } from "../component/vue-test-renderer.mjs";
import { createTsModuleLoader } from "./ts-source-loader.mjs";

const requireFromPptxCore = createRequire(
  new URL("../../packages/pptx-core/package.json", import.meta.url),
);
const { strToU8, zipSync } = await import(
  new URL(`file://${requireFromPptxCore.resolve("fflate")}`).href
);

const {
  DEFAULT_PPTX_PREVIEW_LIMITS,
  PptxPlaybackError,
  PptxPreviewError,
  createPptxCapabilityReport,
  createPptxObjectKey,
  compilePptxSlideSchedule,
  compilePptxSlideTracks,
  evaluatePptxTrack,
  evaluatePptxTriggerGroup,
  matchPptxMorphObjects,
  rebuildPptxStateAtBoundary,
} = await importFromDemo("@arcships/pptx-core");
const {
  PPTX_ARCHIVE_READER,
  createPptxDocumentSession,
  createPptxPreviewSession,
  loadPptxArchive,
  resolvePptxExternalMediaUrl,
} = await importFromDemo("@arcships/pptx-core/browser");
const typescript = requireFromPptxCore("typescript");
// The track index is not part of the public bundle, so load its small source
// graph through data URLs instead of relying on Node's TypeScript loader.
const loadSource = createTsModuleLoader(
  new URL("../../packages/pptx-core/src/", import.meta.url),
  typescript,
);
const trackPlayer = await import(loadSource("browser/track-player.ts"));

test("PPTX core root stays browser-free and exposes fixed safe defaults", () => {
  assert.equal(Object.isFrozen(DEFAULT_PPTX_PREVIEW_LIMITS), true);
  assert.equal(DEFAULT_PPTX_PREVIEW_LIMITS.maxInputBytes, 50 * 1024 * 1024);
  assert.equal(DEFAULT_PPTX_PREVIEW_LIMITS.maxArchiveEntries, 4_000);
  assert.equal(DEFAULT_PPTX_PREVIEW_LIMITS.maxSingleEntryBytes, 32 * 1024 * 1024);
  assert.equal(DEFAULT_PPTX_PREVIEW_LIMITS.maxUncompressedBytes, 256 * 1024 * 1024);
  assert.equal(DEFAULT_PPTX_PREVIEW_LIMITS.maxMediaBytes, 192 * 1024 * 1024);
  assert.equal(DEFAULT_PPTX_PREVIEW_LIMITS.maxCompressionRatio, 100);
  assert.equal(DEFAULT_PPTX_PREVIEW_LIMITS.maxArchivePathLength, 1_024);
  assert.equal(DEFAULT_PPTX_PREVIEW_LIMITS.maxXmlBytes, 256 * 1024 * 1024);
  assert.equal(DEFAULT_PPTX_PREVIEW_LIMITS.maxSingleXmlBytes, 32 * 1024 * 1024);
  assert.equal(DEFAULT_PPTX_PREVIEW_LIMITS.maxXmlDepth, 256);
  assert.equal(DEFAULT_PPTX_PREVIEW_LIMITS.maxXmlAttributeBytes, 1 * 1024 * 1024);
  assert.equal(DEFAULT_PPTX_PREVIEW_LIMITS.maxTextNodeBytes, 16 * 1024 * 1024);
  assert.equal(DEFAULT_PPTX_PREVIEW_LIMITS.maxRelationships, 100_000);

  const error = new PptxPreviewError("LIMIT_EXCEEDED", "too large");
  assert.equal(error.name, "PptxPreviewError");
  assert.equal(error.code, "LIMIT_EXCEEDED");

  const source = readFileSync(new URL("../../packages/pptx-core/dist/index.js", import.meta.url), "utf8");
  assert.doesNotMatch(source, /@aiden0z\/pptx-renderer/);
  assert.doesNotMatch(source, /\b(?:window|document|DOMParser|Worker)\b/);
});

test("PPTX Morph grades strong identity ahead of name/text geometry matches", () => {
  const object = (key, values = {}) => ({
    key,
    slidePath: key.startsWith("from") ? "ppt/slides/slide1.xml" : "ppt/slides/slide2.xml",
    source: "slide",
    shapeId: key,
    groupPath: [],
    nodeType: "shape",
    ...values,
  });
  const matches = matchPptxMorphObjects(
    [
      object("from-name", { explicitMorphName: "!!hero" }),
      object("from-creation", { creationId: "ABC-123" }),
      object("from-duplicate-a", { explicitMorphName: "!!duplicate" }),
      object("from-duplicate-b", { explicitMorphName: "!!duplicate" }),
      object("from-text", { name: "普通名称" }),
    ],
    [
      object("to-name", { explicitMorphName: "!!hero" }),
      object("to-creation", { creationId: "abc-123" }),
      object("to-duplicate", { explicitMorphName: "!!duplicate" }),
      object("to-text", { name: "普通名称" }),
    ],
  );
  assert.deepEqual(
    matches.map((match) => [match.from, match.to, match.method, match.confidence]),
    [
      ["from-name", "to-name", "explicit-name", "strong"],
      ["from-creation", "to-creation", "creation-id", "strong"],
    ],
  );

  const geometric = matchPptxMorphObjects(
    [
      object("from-title", {
        name: "Title 1",
        geometry: { x: 0, y: 0, width: 400, height: 100 },
        text: "开局",
      }),
      object("from-weak", {
        name: "old-label",
        geometry: { x: 0, y: 200, width: 300, height: 100 },
        text: "季度总结",
      }),
    ],
    [
      object("to-title", {
        name: "Title 1",
        geometry: { x: 40, y: 10, width: 400, height: 100 },
        text: "收尾",
      }),
      object("to-weak", {
        name: "new-label",
        geometry: { x: 20, y: 220, width: 300, height: 100 },
        text: "季度总结",
      }),
    ],
  );
  assert.deepEqual(
    geometric.map((match) => [match.from, match.to, match.method, match.confidence]),
    [
      ["from-title", "to-title", "name-and-geometry", "medium"],
      ["from-weak", "to-weak", "text-and-geometry", "weak"],
    ],
  );
  assert.ok(geometric.every((match) => match.unique === true));
  assert.ok(geometric.every((match) => match.score > 0 && match.score <= 1));

  // 没有几何信息时不会凭空产生中等或弱匹配。
  assert.deepEqual(
    matchPptxMorphObjects(
      [object("from-plain", { name: "Title 1" })],
      [object("to-plain", { name: "Title 1" })],
    ),
    [],
  );
});

test("PPTX playback root provides stable object keys and capability counts", () => {
  const topLevel = createPptxObjectKey({
    slidePath: "ppt/slides/slide4.xml",
    source: "slide",
    shapeId: "12",
  });
  const grouped = createPptxObjectKey({
    slidePath: "/ppt/slides/slide4.xml/",
    source: "slide",
    shapeId: "12",
    groupPath: ["8"],
  });
  assert.equal(topLevel, "ppt/slides/slide4.xml|slide|shape:12");
  assert.equal(grouped, "ppt/slides/slide4.xml|slide|group:8/shape:12");
  assert.notEqual(topLevel, grouped);
  assert.throws(
    () => createPptxObjectKey({ slidePath: "ppt/slides/slide1.xml", source: "slide", shapeId: "1/2" }),
    TypeError,
  );

  const report = createPptxCapabilityReport([
    { id: "1", slideIndex: 0, feature: "fade", disposition: "strict" },
    { id: "2", slideIndex: 0, feature: "morph", disposition: "approximate" },
    { id: "3", slideIndex: 0, feature: "chart-build", disposition: "static" },
    { id: "4", slideIndex: 0, feature: "unknown", disposition: "unparsed" },
  ]);
  assert.deepEqual(
    {
      discovered: report.discovered,
      strict: report.strict,
      approximate: report.approximate,
      static: report.static,
      unparsed: report.unparsed,
    },
    { discovered: 4, strict: 1, approximate: 1, static: 1, unparsed: 1 },
  );
  assert.equal(Object.isFrozen(report), true);
  assert.equal(Object.isFrozen(report.features), true);

  const error = new PptxPlaybackError("TARGET_NOT_FOUND", "missing", {
    slideIndex: 2,
    objectKey: grouped,
  });
  assert.equal(error.name, "PptxPlaybackError");
  assert.equal(error.slideIndex, 2);
  assert.equal(error.objectKey, grouped);
});

function timeNode(id, values = {}) {
  return {
    id,
    container: "parallel",
    kind: "group",
    delayMs: 0,
    durationMs: 0,
    autoReverse: false,
    fill: "hold",
    restart: "always",
    acceleration: 0,
    deceleration: 0,
    conditions: [],
    childIds: [],
    ...values,
  };
}

function playbackSlide(nodes, rootNodeId = "root") {
  return {
    index: 0,
    hidden: false,
    objects: [],
    morphFromPrevious: [],
    rootNodeId,
    nodes,
    media: [],
    actions: [],
    capability: createPptxCapabilityReport(),
  };
}

test("PPTX time tree compiles click, with-previous and after-previous groups", () => {
  const target = "ppt/slides/slide1.xml|slide|shape:5";
  const slide = playbackSlide({
    root: timeNode("root", { childIds: ["click-1", "click-2"] }),
    "click-1": timeNode("click-1", { nodeType: "clickEffect", childIds: ["fade", "scale"] }),
    fade: timeNode("fade", {
      kind: "effect",
      nodeType: "withEffect",
      durationMs: 500,
      effect: { id: "fade", kind: "fade-in", targetObjectKey: target, presetClass: "entr", values: {} },
    }),
    scale: timeNode("scale", {
      kind: "effect",
      nodeType: "afterEffect",
      durationMs: 500,
      effect: {
        id: "scale",
        kind: "scale",
        targetObjectKey: target,
        presetClass: "entr",
        values: { scaleFromX: 0, scaleFromY: 0, scaleToX: 1, scaleToY: 1 },
      },
    }),
    "click-2": timeNode("click-2", { nodeType: "clickEffect", childIds: ["hide"] }),
    hide: timeNode("hide", {
      kind: "effect",
      nodeType: "withEffect",
      effect: { id: "hide", kind: "disappear", targetObjectKey: target, presetClass: "exit", values: {} },
    }),
  });
  const schedule = compilePptxSlideSchedule(slide);
  assert.equal(schedule.clickBoundaryCount, 2);
  assert.deepEqual(schedule.groups.map((group) => group.key), ["click:1", "click:2"]);
  assert.deepEqual(
    schedule.groups[0].effects.map((effect) => [effect.nodeId, effect.startMs, effect.endMs]),
    [["fade", 0, 500], ["scale", 500, 1000]],
  );

  const compiled = compilePptxSlideTracks(slide);
  assert.equal(compiled.initialState[target].display, false);
  assert.equal(compiled.initialState[target].opacity, 0);
  assert.equal(compiled.initialState[target]["scale-x"], 0);
  const first = compiled.groups[0];
  const at250 = evaluatePptxTriggerGroup(first, 250);
  const at750 = evaluatePptxTriggerGroup(first, 750);
  assert.equal(at250[target].opacity, 0.5);
  assert.equal(at750[target]["scale-x"], 0.5);
  const beforeClick = rebuildPptxStateAtBoundary(compiled, 0);
  const afterFirstClick = rebuildPptxStateAtBoundary(compiled, 1);
  assert.equal(beforeClick[target].display, false);
  assert.equal(afterFirstClick[target].display, true);
  assert.equal(afterFirstClick[target].opacity, 1);
  assert.equal(afterFirstClick[target]["scale-x"], 1);
});

test("PPTX main sequence onNext children become separate click steps", () => {
  const target = "ppt/slides/slide1.xml|slide|shape:5";
  const effect = (id) => timeNode(id, {
    kind: "effect",
    nodeType: "afterEffect",
    durationMs: 400,
    effect: { id, kind: "fade-in", targetObjectKey: target, presetClass: "entr", values: {} },
  });
  const slide = playbackSlide({
    root: timeNode("root", { childIds: ["main"] }),
    main: timeNode("main", {
      nodeType: "mainSeq",
      container: "sequence",
      conditions: [{ source: "next", event: "on-next", delayMs: 0 }],
      childIds: ["generated-step-1", "generated-step-2", "standard-wrapper"],
    }),
    "generated-step-1": timeNode("generated-step-1", {
      conditions: [{ source: "start", event: "delay", delayMs: 0 }],
      childIds: ["generated-effect-1"],
    }),
    "generated-effect-1": effect("generated-effect-1"),
    "generated-step-2": timeNode("generated-step-2", {
      conditions: [{ source: "start", event: "delay", delayMs: 0 }],
      childIds: ["generated-effect-2"],
    }),
    "generated-effect-2": effect("generated-effect-2"),
    "standard-wrapper": timeNode("standard-wrapper", { childIds: ["standard-click"] }),
    "standard-click": timeNode("standard-click", {
      nodeType: "clickEffect",
      childIds: ["standard-effect"],
    }),
    "standard-effect": effect("standard-effect"),
  });

  const schedule = compilePptxSlideSchedule(slide);
  assert.equal(schedule.clickBoundaryCount, 3);
  assert.deepEqual(schedule.groups.map((group) => group.key), ["click:1", "click:2", "click:3"]);
  assert.deepEqual(
    schedule.groups.map((group) => group.effects.map((item) => item.nodeId)),
    [["generated-effect-1"], ["generated-effect-2"], ["standard-effect"]],
  );
});

test("PPTX property tracks resolve simultaneous writes by stable node order", () => {
  const target = "ppt/slides/slide1.xml|slide|shape:7";
  const slide = playbackSlide({
    root: timeNode("root", { childIds: ["click"] }),
    click: timeNode("click", { nodeType: "clickEffect", childIds: ["fade-in", "fade-out"] }),
    "fade-in": timeNode("fade-in", {
      kind: "effect",
      nodeType: "withEffect",
      durationMs: 500,
      effect: { id: "fade-in", kind: "fade-in", targetObjectKey: target, presetClass: "entr", values: {} },
    }),
    "fade-out": timeNode("fade-out", {
      kind: "effect",
      nodeType: "withEffect",
      durationMs: 500,
      effect: { id: "fade-out", kind: "fade-out", targetObjectKey: target, presetClass: "exit", values: {} },
    }),
  });
  const compiled = compilePptxSlideTracks(slide);
  const finalState = rebuildPptxStateAtBoundary(compiled, 1);
  assert.equal(finalState[target].opacity, 0);
  assert.equal(finalState[target].display, false);
});

test("PPTX property tracks reveal a lower active effect after a later remove-fill overlap", () => {
  const target = "ppt/slides/slide1.xml|slide|shape:8";
  const slide = playbackSlide({
    root: timeNode("root", { childIds: ["click"] }),
    click: timeNode("click", { nodeType: "clickEffect", childIds: ["base", "overlay"] }),
    base: timeNode("base", {
      kind: "effect",
      nodeType: "withEffect",
      durationMs: 1000,
      effect: {
        id: "base",
        kind: "emphasis",
        targetObjectKey: target,
        presetClass: "emph",
        values: { emphasisProperty: "opacity", emphasisFrom: 0, emphasisTo: 1 },
      },
    }),
    overlay: timeNode("overlay", {
      kind: "effect",
      nodeType: "withEffect",
      durationMs: 200,
      fill: "remove",
      conditions: [{ source: "start", event: "with-previous", delayMs: 500 }],
      effect: {
        id: "overlay",
        kind: "emphasis",
        targetObjectKey: target,
        presetClass: "emph",
        values: { emphasisProperty: "opacity", emphasisFrom: 1, emphasisTo: 0 },
      },
    }),
  });
  const group = compilePptxSlideTracks(slide).groups[0];
  assert.equal(evaluatePptxTriggerGroup(group, 600)[target].opacity, 0.5);
  assert.equal(evaluatePptxTriggerGroup(group, 700)[target].opacity, 0.7);
  assert.equal(evaluatePptxTriggerGroup(group, 900)[target].opacity, 0.9);
});

test("PPTX unsupported entrance effects keep their static initial state", () => {
  const target = "ppt/slides/slide1.xml|slide|shape:14";
  const slide = playbackSlide({
    root: timeNode("root", { childIds: ["unsupported-wipe"] }),
    "unsupported-wipe": timeNode("unsupported-wipe", {
      kind: "effect",
      durationMs: 500,
      effect: {
        id: "unsupported-wipe",
        kind: "wipe",
        targetObjectKey: target,
        presetClass: "entr",
        filter: "wipe(up)",
        values: {},
      },
    }),
  });
  const compiled = compilePptxSlideTracks(slide, Object.freeze({
    [target]: Object.freeze({ display: true, opacity: 1, "clip-path": "none" }),
  }));
  assert.equal(compiled.initialState[target].display, true);
  assert.equal(compiled.initialState[target]["clip-path"], "none");
  assert.equal(compiled.groups[0].tracks.length, 0);
  assert.equal(compiled.groups[0].effects.length, 0);
  assert.equal(compiled.groups[0].durationMs, 0);
});

test("PPTX time tree does not count inherited click events on behavior nodes", () => {
  const target = "ppt/slides/slide1.xml|slide|shape:9";
  const slide = playbackSlide({
    root: timeNode("root", { childIds: ["click-1", "click-2", "click-3"] }),
    "click-1": timeNode("click-1", { nodeType: "clickEffect", childIds: ["fade-1"] }),
    "fade-1": timeNode("fade-1", {
      kind: "effect",
      nodeType: "withEffect",
      childIds: ["behavior-1"],
      effect: { id: "fade-1", kind: "fade-in", targetObjectKey: target, presetClass: "entr", values: {} },
    }),
    "behavior-1": timeNode("behavior-1"),
    "click-2": timeNode("click-2", { nodeType: "clickEffect", childIds: ["fade-2"] }),
    "fade-2": timeNode("fade-2", {
      kind: "effect",
      nodeType: "withEffect",
      childIds: ["behavior-2"],
      effect: { id: "fade-2", kind: "fade-in", targetObjectKey: target, presetClass: "entr", values: {} },
    }),
    "behavior-2": timeNode("behavior-2"),
    "click-3": timeNode("click-3", { nodeType: "clickEffect", childIds: ["fade-3"] }),
    "fade-3": timeNode("fade-3", {
      kind: "effect",
      nodeType: "withEffect",
      effect: { id: "fade-3", kind: "fade-in", targetObjectKey: target, presetClass: "entr", values: {} },
    }),
  });
  const schedule = compilePptxSlideSchedule(slide);
  assert.equal(schedule.clickBoundaryCount, 3);
  assert.deepEqual(schedule.groups.map((group) => group.key), ["click:1", "click:2", "click:3"]);
});

test("PPTX time tree keeps object clicks and media bookmarks in separate trigger groups", () => {
  const trigger = "ppt/slides/slide1.xml|slide|shape:9";
  const target = "ppt/slides/slide1.xml|slide|shape:10";
  const slide = playbackSlide({
    root: timeNode("root", { childIds: ["shape-group", "bookmark-group"] }),
    "shape-group": timeNode("shape-group", {
      childIds: ["shape-effect"],
      conditions: [{ source: "start", event: "on-shape-click", delayMs: 0, targetObjectKey: trigger }],
    }),
    "shape-effect": timeNode("shape-effect", {
      kind: "effect",
      durationMs: 200,
      effect: { id: "shape-effect", kind: "fade-in", targetObjectKey: target, presetClass: "entr", values: {} },
    }),
    "bookmark-group": timeNode("bookmark-group", {
      childIds: ["bookmark-effect"],
      conditions: [{ source: "start", event: "on-media-bookmark", delayMs: 0, bookmarkName: "middle" }],
    }),
    "bookmark-effect": timeNode("bookmark-effect", {
      kind: "effect",
      durationMs: 300,
      effect: { id: "bookmark-effect", kind: "appear", targetObjectKey: target, presetClass: "entr", values: {} },
    }),
  });
  const schedule = compilePptxSlideSchedule(slide);
  assert.deepEqual(schedule.groups.map((group) => group.key), [
    `shape:${trigger}`,
    "bookmark:middle",
  ]);
  assert.equal(schedule.clickBoundaryCount, 0);
});

test("PPTX time tree keeps unsupported trigger events out of executable groups", () => {
  const target = "ppt/slides/slide1.xml|slide|shape:11";
  const slide = playbackSlide({
    root: timeNode("root", { childIds: ["hover-group"] }),
    "hover-group": timeNode("hover-group", {
      childIds: ["hover-effect"],
      conditions: [{ source: "start", event: "unknown", delayMs: 0, rawEvent: "onMouseOver" }],
    }),
    "hover-effect": timeNode("hover-effect", {
      kind: "effect",
      durationMs: 200,
      effect: { id: "hover-effect", kind: "fade-in", targetObjectKey: target, presetClass: "entr", values: {} },
    }),
  });
  const schedule = compilePptxSlideSchedule(slide);
  assert.deepEqual(schedule.groups, []);
  assert.ok(schedule.unsupportedNodeIds.includes("hover-group"));
  assert.ok(schedule.unsupportedNodeIds.includes("hover-effect"));
});

test("PPTX tracks execute finite repeats, indefinite repeats, auto-reverse, easing and fill removal", () => {
  const target = "ppt/slides/slide1.xml|slide|shape:12";
  const repeated = playbackSlide({
    root: timeNode("root", { childIds: ["repeat"] }),
    repeat: timeNode("repeat", {
      kind: "effect",
      durationMs: 100,
      repeatCount: 2,
      autoReverse: true,
      acceleration: 0.5,
      fill: "remove",
      effect: { id: "repeat", kind: "fade-in", targetObjectKey: target, presetClass: "entr", values: {} },
    }),
  });
  const repeatedCompiled = compilePptxSlideTracks(repeated);
  assert.equal(repeatedCompiled.groups[0].durationMs, 400);
  const opacity = repeatedCompiled.groups[0].tracks.find((track) => track.property === "opacity");
  const display = repeatedCompiled.groups[0].tracks.find((track) => track.property === "display");
  assert.ok(opacity);
  assert.ok(evaluatePptxTrack(opacity, 50) > 0 && evaluatePptxTrack(opacity, 50) < 0.5);
  assert.equal(evaluatePptxTrack(opacity, 100), 1);
  assert.ok(evaluatePptxTrack(opacity, 150) > 0.5 && evaluatePptxTrack(opacity, 150) < 1);
  assert.equal(evaluatePptxTrack(opacity, 200), 0);
  assert.equal(evaluatePptxTrack(opacity, 400), 0);
  assert.equal(evaluatePptxTrack(display, 400), false);

  const indefinite = playbackSlide({
    root: timeNode("root", { childIds: ["loop"] }),
    loop: timeNode("loop", {
      kind: "effect",
      durationMs: 100,
      repeatCount: "indefinite",
      autoReverse: true,
      effect: { id: "loop", kind: "scale", targetObjectKey: target, presetClass: "emph", values: { scaleFromX: 1, scaleFromY: 1, scaleToX: 2, scaleToY: 2 } },
    }),
  });
  const loopTrack = compilePptxSlideTracks(indefinite).groups[0].tracks.find((track) => track.property === "scale-x");
  assert.ok(loopTrack?.repeatIndefinite);
  assert.equal(evaluatePptxTrack(loopTrack, 50), 1.5);
  assert.equal(evaluatePptxTrack(loopTrack, 150), 1.5);
  assert.equal(evaluatePptxTrack(loopTrack, 250), 1.5);
});

test("PPTX wipe and normalized straight motion paths interpolate in page coordinates", () => {
  const target = "ppt/slides/slide1.xml|slide|shape:13";
  const slide = playbackSlide({
    root: timeNode("root", { childIds: ["wipe", "motion"] }),
    wipe: timeNode("wipe", {
      kind: "effect",
      durationMs: 100,
      effect: { id: "wipe", kind: "wipe", targetObjectKey: target, presetClass: "entr", filter: "wipe(left)", values: {} },
    }),
    motion: timeNode("motion", {
      kind: "effect",
      nodeType: "withEffect",
      durationMs: 100,
      effect: {
        id: "motion",
        kind: "motion-path",
        targetObjectKey: target,
        motionPath: "M 0 0 L 0.25 -0.5 E",
        values: { motionScaleX: 960, motionScaleY: 540 },
      },
    }),
  });
  const tracks = compilePptxSlideTracks(slide).groups[0].tracks;
  const wipe = tracks.find((track) => track.property === "clip-path");
  const x = tracks.find((track) => track.property === "translate-x");
  const y = tracks.find((track) => track.property === "translate-y");
  assert.equal(evaluatePptxTrack(wipe, 50), "inset(0% 50% 0% 0%)");
  assert.equal(evaluatePptxTrack(x, 100), 240);
  assert.equal(evaluatePptxTrack(y, 100), -270);
});

test("PPTX preview separates vertical browse from slide playback and keeps one archive pass", () => {
  assert.equal(typeof createPptxDocumentSession, "function");
  assert.equal(typeof createPptxPreviewSession, "function");
  const source = readFileSync(
    new URL("../../packages/pptx-core/src/browser.ts", import.meta.url),
    "utf8",
  );
  assert.match(source, /loadPptxArchive\(resolved\.buffer/);
  // The streaming validator inflates every entry, so the browser path must not
  // call it next to the renderer parse.
  assert.doesNotMatch(source, /validateOfficeArchive\(/);
  assert.doesNotMatch(source, /parseZip/);
  assert.match(source, /buildPresentation\(files/);
  assert.match(source, /viewer\.load\(presentation\)/);
  assert.match(source, /parsePptxPlaybackDocument\(files, presentation/);
  assert.match(source, /viewer\.renderList\(/);
  assert.match(source, /viewer\.renderSlide\(initialSlide\)/);
  assert.match(source, /renderMode: options\.renderMode \?\? "list"/);
  assert.match(source, /resolvePptxExternalMediaUrl\(path, externalMediaPolicy\)/);
  assert.doesNotMatch(source, /viewer\.open\(/);

  const archiveSource = readFileSync(
    new URL("../../packages/pptx-core/src/archive.ts", import.meta.url),
    "utf8",
  );
  assert.match(archiveSource, /reader\.parseLazyMedia\(buffer/);
  assert.match(archiveSource, /inspectOfficeArchive\(buffer/);
  assert.match(archiveSource, /validateOfficeArchiveParts\(pptxArchiveParts\(files/);

  const patch = readFileSync(
    new URL("../../patches/@aiden0z__pptx-renderer@1.2.4.patch", import.meta.url),
    "utf8",
  );
  assert.match(patch, /dataset\.pptxNodeId = t\.id/);
  assert.match(patch, /dataset\.pptxNodeType = t\.nodeType/);
  assert.match(patch, /dataset\.pptxPartPath/);
  assert.match(patch, /dataset\.pptxParagraphIndex = String\(m\)/);
});

test("PPTX time tree inherits the delay of every enclosing container", () => {
  const target = "ppt/slides/slide1.xml|slide|shape:21";
  const fade = (id, values = {}) => timeNode(id, {
    kind: "effect",
    durationMs: 200,
    effect: { id, kind: "fade-in", targetObjectKey: target, presetClass: "entr", values: {} },
    ...values,
  });
  const delayedBatch = playbackSlide({
    root: timeNode("root", { childIds: ["batch"] }),
    batch: timeNode("batch", { delayMs: 500, childIds: ["fade-a", "fade-b"] }),
    "fade-a": fade("fade-a"),
    "fade-b": fade("fade-b", {
      conditions: [{ source: "start", event: "with-previous", delayMs: 0 }],
    }),
  });
  assert.deepEqual(
    compilePptxSlideSchedule(delayedBatch).groups[0].effects.map((effect) => [effect.nodeId, effect.startMs]),
    [["fade-a", 500], ["fade-b", 500]],
  );

  const nestedBatch = playbackSlide({
    root: timeNode("root", { childIds: ["batch"] }),
    batch: timeNode("batch", { delayMs: 500, childIds: ["inner"] }),
    inner: timeNode("inner", { delayMs: 250, childIds: ["nested-fade"] }),
    "nested-fade": fade("nested-fade"),
  });
  assert.deepEqual(
    compilePptxSlideSchedule(nestedBatch).groups[0].effects.map((effect) => [effect.nodeId, effect.startMs]),
    [["nested-fade", 750]],
  );
});

test("PPTX exit fade keeps the element visible until every repeat pass ends", () => {
  const target = "ppt/slides/slide1.xml|slide|shape:22";
  const slide = playbackSlide({
    root: timeNode("root", { childIds: ["fade-out"] }),
    "fade-out": timeNode("fade-out", {
      kind: "effect",
      durationMs: 200,
      repeatCount: 2,
      autoReverse: true,
      effect: { id: "fade-out", kind: "fade-out", targetObjectKey: target, presetClass: "exit", values: {} },
    }),
  });
  const group = compilePptxSlideTracks(slide).groups[0];
  assert.equal(group.durationMs, 800);
  const display = group.tracks.find((track) => track.property === "display");
  const opacity = group.tracks.find((track) => track.property === "opacity");
  assert.equal(evaluatePptxTrack(display, 199), true);
  assert.equal(evaluatePptxTrack(display, 400), true);
  assert.ok(evaluatePptxTrack(opacity, 700) > 0);
  assert.equal(evaluatePptxTrack(display, 799), true);
  assert.equal(evaluatePptxTrack(display, 800), false);
});

test("PPTX simultaneous transform behaviours compose instead of dropping the earlier one", () => {
  const target = "ppt/slides/slide1.xml|slide|shape:23";
  const grow = (id, scale) => timeNode(id, {
    kind: "effect",
    nodeType: "withEffect",
    durationMs: 500,
    effect: {
      id,
      kind: "scale",
      targetObjectKey: target,
      presetClass: "emph",
      values: { scaleFromX: 1, scaleFromY: 1, scaleToX: scale, scaleToY: scale },
    },
  });
  const move = (id, x, nodeType) => timeNode(id, {
    kind: "effect",
    nodeType,
    durationMs: 1000,
    effect: { id, kind: "motion-path", targetObjectKey: target, motionPath: `M 0 0 L ${x} 0 E`, values: {} },
  });
  const slide = playbackSlide({
    root: timeNode("root", { childIds: ["click"] }),
    click: timeNode("click", {
      nodeType: "clickEffect",
      childIds: ["grow-a", "grow-b", "move-a", "move-b"],
    }),
    "grow-a": grow("grow-a", 2),
    "grow-b": grow("grow-b", 3),
    "move-a": move("move-a", 100, "withEffect"),
    "move-b": move("move-b", 50, "afterEffect"),
  });
  const group = compilePptxSlideTracks(slide).groups[0];
  const scaleX = group.tracks.find((track) => track.property === "scale-x");
  const translateX = group.tracks.find((track) => track.property === "translate-x");
  assert.equal(evaluatePptxTrack(scaleX, 250), 2.5);
  assert.equal(evaluatePptxTrack(scaleX, 500), 4);
  // Sequential motion keeps composing from the settled position instead of restarting.
  assert.equal(evaluatePptxTrack(translateX, 500), 50);
  assert.equal(evaluatePptxTrack(translateX, 1500), 125);
});

function fakeParagraph(index, computed = {}) {
  return {
    dataset: { pptxParagraphIndex: String(index) },
    style: {},
    computed: {
      visibility: "visible",
      display: "block",
      opacity: "1",
      clipPath: "none",
      filter: "none",
      transform: "none",
      translate: "none",
      scale: "none",
      rotate: "none",
      ...computed,
    },
    querySelectorAll() {
      return [];
    },
  };
}

function fakeElement(objectKey, { computed = {}, paragraphs = [] } = {}) {
  return {
    dataset: { pptxObjectKey: objectKey },
    style: {},
    paragraphs,
    computed: {
      visibility: "visible",
      display: "block",
      opacity: "1",
      clipPath: "none",
      filter: "none",
      transform: "none",
      translate: "none",
      scale: "none",
      rotate: "none",
      ...computed,
    },
    querySelectorAll(selector) {
      return selector === "[data-pptx-paragraph-index]" ? paragraphs : [];
    },
  };
}

function fakeRoot(elements) {
  const root = {
    queries: 0,
    querySelectorAll(selector) {
      root.queries += 1;
      if (selector === "[data-pptx-object-key]") return elements;
      return elements.flatMap((element) => element.querySelectorAll(selector));
    },
  };
  return root;
}

function withComputedStyle(run) {
  const previous = globalThis.getComputedStyle;
  globalThis.getComputedStyle = (element) => element.computed;
  try {
    return run();
  } finally {
    if (previous === undefined) delete globalThis.getComputedStyle;
    else globalThis.getComputedStyle = previous;
  }
}

function approximately(actual, expected, tolerance = 1e-4) {
  return typeof actual === "number" && Math.abs(actual - expected) <= tolerance;
}

test("PPTX static state capture keeps the transform the renderer already applied", () => {
  withComputedStyle(() => {
    const objectKey = "ppt/slides/slide1.xml|slide|shape:24";
    const element = fakeElement(objectKey, {
      computed: {
        opacity: "0.9",
        // rotate(30deg) scale(2) translate(10px, 20px)
        transform: "matrix(1.7320508, 1, -1, 1.7320508, 10, 20)",
      },
    });
    const root = fakeRoot([element]);
    const captured = trackPlayer.capturePptxStaticState(root);
    const state = captured[objectKey];
    assert.ok(approximately(state.rotate, 30));
    assert.ok(approximately(state["scale-x"], 2));
    assert.ok(approximately(state["scale-y"], 2));
    assert.equal(state["translate-x"], 10);
    assert.equal(state["translate-y"], 20);
    assert.equal(state.opacity, 0.9);

    // Resting on the captured state must not stack a second copy on top of the renderer.
    trackPlayer.applyPptxObjectState(root, captured);
    assert.equal(element.style.rotate, "0deg");
    assert.equal(element.style.scale, "1 1");
    assert.equal(element.style.translate, "0px 0px");

    // An entrance motion animates relative to the position the renderer painted.
    trackPlayer.applyPptxObjectState(root, {
      [objectKey]: { ...state, "translate-x": 250, "translate-y": -250 },
    });
    assert.equal(element.style.translate, "240px -270px");
    assert.equal(element.style.rotate, "0deg");
  });
});

test("PPTX track targets resolve once per render and animate duplicate object keys", () => {
  withComputedStyle(() => {
    const duplicateKey = "ppt/slides/slide1.xml|slide|shape:25";
    const uniqueKey = "ppt/slides/slide1.xml|slide|shape:26";
    const first = fakeElement(duplicateKey);
    first.isConnected = true;
    const second = fakeElement(duplicateKey);
    const unique = fakeElement(uniqueKey, {
      paragraphs: [fakeParagraph(0), fakeParagraph(2)],
    });
    const root = fakeRoot([first, second, unique]);
    const targets = trackPlayer.createPptxTrackTargets(root);
    assert.equal(root.queries, 1);
    assert.deepEqual(targets.elements.get(duplicateKey).length, 2);
    assert.deepEqual(
      [...targets.elements.keys()].filter((key) => key.startsWith(`${uniqueKey}#`)),
      [
        `${uniqueKey}#paragraph:0-0`,
        `${uniqueKey}#paragraph:2-2`,
        `${uniqueKey}#paragraph:0-2`,
      ],
    );

    const captured = trackPlayer.capturePptxStaticState(root, targets);
    assert.equal(root.queries, 1);
    assert.ok(captured[duplicateKey]);
    assert.ok(captured[`${uniqueKey}#paragraph:0-2`]);

    // Resting on the captured state keeps every duplicate at identity.
    trackPlayer.applyPptxObjectState(root, captured, targets);
    assert.equal(first.style.scale, "1 1");
    assert.equal(second.style.scale, "1 1");

    const state = {
      [duplicateKey]: { display: false, opacity: 0.25 },
      [`${uniqueKey}#paragraph:0-2`]: { opacity: 0.5 },
    };
    for (let frame = 0; frame < 100; frame += 1) {
      trackPlayer.applyPptxObjectState(root, state, targets);
    }
    assert.equal(root.queries, 1);
    for (const element of [first, second]) {
      assert.equal(element.style.visibility, "hidden");
      assert.equal(element.style.opacity, "0.25");
    }
    assert.equal(unique.paragraphs[0].style.opacity, "0.5");
    assert.equal(unique.paragraphs[1].style.opacity, "0.5");

    // A re-rendered slide detaches the cached elements, so the index is resolved again.
    first.isConnected = false;
    trackPlayer.applyPptxObjectState(root, state, targets);
    assert.equal(root.queries, 2);
  });
});

const PPTX_CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
</Types>`;

function pptxArchive(entries = {}) {
  return zipSync({
    "[Content_Types].xml": strToU8(PPTX_CONTENT_TYPES),
    ...Object.fromEntries(Object.entries(entries).map(([name, value]) => [name, strToU8(value)])),
  });
}

function countingReader(calls) {
  return {
    async parse(buffer, limits) {
      calls.push("parse");
      return PPTX_ARCHIVE_READER.parse(buffer, limits);
    },
    async parseLazyMedia(buffer, limits) {
      calls.push("parseLazyMedia");
      return PPTX_ARCHIVE_READER.parseLazyMedia(buffer, limits);
    },
  };
}

function bufferOf(bytes) {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

test("PPTX archive loading decompresses the container once and validates the decoded parts", async () => {
  const fixture = readFileSync(new URL("../fixtures/pptx/playback-controlled.pptx", import.meta.url));
  const calls = [];
  const loaded = await loadPptxArchive(bufferOf(fixture), {
    limits: DEFAULT_PPTX_PREVIEW_LIMITS,
    lazyMedia: true,
    reader: countingReader(calls),
  });
  // The reader is the only decompression pass: no second validator unzip.
  assert.deepEqual(calls, ["parseLazyMedia"]);
  assert.equal(loaded.validation.sourceFormat, "pptx");
  assert.ok(loaded.validation.entryCount > 0);
  assert.ok(loaded.validation.xmlBytes > 0);
  assert.ok(loaded.validation.entryCount <= loaded.summary.entryCount);
  assert.ok(loaded.summary.uncompressedBytes >= loaded.validation.uncompressedBytes);
  assert.equal(Object.isFrozen(loaded), true);
  assert.ok(loaded.files.slides.size > 0);

  const eagerCalls = [];
  const eager = await loadPptxArchive(bufferOf(fixture), {
    limits: DEFAULT_PPTX_PREVIEW_LIMITS,
    reader: countingReader(eagerCalls),
  });
  assert.deepEqual(eagerCalls, ["parse"]);
  assert.ok(eager.validation.uncompressedBytes >= loaded.validation.uncompressedBytes);
});

test("PPTX archive budgets reject ratio bombs before decompressing and deep XML after it", async () => {
  // A metadata-level ratio bomb never reaches the reader.
  const calls = [];
  const bomb = pptxArchive({
    "ppt/presentation.xml": "<p:presentation/>",
    "ppt/media/blob.bin": "A".repeat(4 * 1024 * 1024),
  });
  await assert.rejects(
    loadPptxArchive(bufferOf(bomb), {
      limits: { ...DEFAULT_PPTX_PREVIEW_LIMITS, maxCompressionRatio: 50 },
      reader: countingReader(calls),
    }),
    (error) => error?.code === "LIMIT_EXCEEDED" && error?.limit === "maxCompressionRatio",
  );
  assert.deepEqual(calls, []);

  // Deep XML is caught on the parts the renderer already decoded.
  const deep = `<p:presentation>${"<a>".repeat(40)}x${"</a>".repeat(40)}</p:presentation>`;
  const deepCalls = [];
  await assert.rejects(
    loadPptxArchive(bufferOf(pptxArchive({ "ppt/presentation.xml": deep })), {
      limits: { ...DEFAULT_PPTX_PREVIEW_LIMITS, maxXmlDepth: 8 },
      lazyMedia: true,
      reader: countingReader(deepCalls),
    }),
    (error) => error?.code === "LIMIT_EXCEEDED" && error?.limit === "maxXmlDepth",
  );
  assert.deepEqual(deepCalls, ["parseLazyMedia"]);

  await assert.rejects(
    loadPptxArchive(bufferOf(pptxArchive({ "ppt/presentation.xml": "<p:presentation/>" })), {
      limits: { ...DEFAULT_PPTX_PREVIEW_LIMITS, maxSingleXmlBytes: 8 },
      reader: countingReader([]),
    }),
    (error) => error?.code === "LIMIT_EXCEEDED" && error?.limit === "maxSingleXmlBytes",
  );

  // Long entry paths are rejected from the central directory, before inflating.
  const longPathCalls = [];
  await assert.rejects(
    loadPptxArchive(bufferOf(pptxArchive({
      "ppt/presentation.xml": "<p:presentation/>",
      [`ppt/media/${"long".repeat(40)}.bin`]: "x",
    })), {
      limits: { ...DEFAULT_PPTX_PREVIEW_LIMITS, maxArchivePathLength: 32 },
      reader: countingReader(longPathCalls),
    }),
    (error) => error?.code === "LIMIT_EXCEEDED" && error?.limit === "maxArchivePathLength",
  );
  assert.deepEqual(longPathCalls, []);

  // Relationship budgets are accumulated over the decoded `.rels` parts.
  const relationships = pptxArchive({
    "ppt/presentation.xml": "<p:presentation/>",
    "ppt/_rels/presentation.xml.rels": "<Relationships>"
      + ["1", "2", "3"].map((id) => `<Relationship Id="r${id}" Type="t" Target="x"/>`).join("")
      + "</Relationships>",
  });
  await assert.rejects(
    loadPptxArchive(bufferOf(relationships), {
      limits: { ...DEFAULT_PPTX_PREVIEW_LIMITS, maxRelationships: 2 },
      lazyMedia: true,
      reader: countingReader([]),
    }),
    (error) => error?.code === "LIMIT_EXCEEDED" && error?.limit === "maxRelationships",
  );
  const accepted = await loadPptxArchive(bufferOf(relationships), {
    limits: { ...DEFAULT_PPTX_PREVIEW_LIMITS, maxRelationships: 3 },
    lazyMedia: true,
    reader: countingReader([]),
  });
  assert.equal(accepted.validation.relationships, 3);
});

test("PPTX external media needs both the opt-in and a URL policy match", () => {
  assert.equal(resolvePptxExternalMediaUrl("https://cdn.example/a.png", undefined), undefined);
  const policy = {
    baseUrl: "https://deck.example/slides/",
    allowedProtocols: ["https:"],
    allowedOrigins: ["https://cdn.example"],
  };
  assert.equal(
    resolvePptxExternalMediaUrl("https://cdn.example/webfont/tracker.gif", policy),
    "https://cdn.example/webfont/tracker.gif",
  );
  assert.equal(resolvePptxExternalMediaUrl("https://tracker.example/pixel.gif", policy), undefined);
  assert.equal(resolvePptxExternalMediaUrl("javascript:alert(1)", policy), undefined);
  assert.equal(resolvePptxExternalMediaUrl("data:image/png;base64,AAAA", policy), undefined);
  assert.equal(resolvePptxExternalMediaUrl("https://cdn.example/x.png", { ...policy, enabled: false }), undefined);
});

test("PPTX per-open limits rebuild the viewer used by lazy slide and media parsing", () => {
  const source = readFileSync(
    new URL("../../packages/pptx-core/src/browser.ts", import.meta.url),
    "utf8",
  );
  assert.match(source, /const createViewer = \(/);
  assert.match(source, /zipLimits: toZipLimits\(limits\)/);
  assert.match(source, /currentLimits = limits/);
  assert.match(source, /if \(!sameLimits\(limits, viewerLimits\)\)/);
  assert.match(source, /viewer = createViewer\(limits, zoomPercent\)/);

  const types = readFileSync(
    new URL("../../packages/pptx-core/src/browser-types.ts", import.meta.url),
    "utf8",
  );
  assert.match(types, /readonly limits: Readonly<PptxPreviewLimits>/);
});
