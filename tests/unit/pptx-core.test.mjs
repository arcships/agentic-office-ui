import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import { importFromDemo } from "../component/vue-test-renderer.mjs";

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
  createPptxDocumentSession,
  createPptxPreviewSession,
} = await importFromDemo("@arcships/pptx-core/browser");

test("PPTX core root stays browser-free and exposes fixed safe defaults", () => {
  assert.equal(Object.isFrozen(DEFAULT_PPTX_PREVIEW_LIMITS), true);
  assert.equal(DEFAULT_PPTX_PREVIEW_LIMITS.maxInputBytes, 50 * 1024 * 1024);
  assert.equal(DEFAULT_PPTX_PREVIEW_LIMITS.maxArchiveEntries, 4_000);
  assert.equal(DEFAULT_PPTX_PREVIEW_LIMITS.maxSingleEntryBytes, 32 * 1024 * 1024);
  assert.equal(DEFAULT_PPTX_PREVIEW_LIMITS.maxUncompressedBytes, 256 * 1024 * 1024);
  assert.equal(DEFAULT_PPTX_PREVIEW_LIMITS.maxMediaBytes, 192 * 1024 * 1024);

  const error = new PptxPreviewError("LIMIT_EXCEEDED", "too large");
  assert.equal(error.name, "PptxPreviewError");
  assert.equal(error.code, "LIMIT_EXCEEDED");

  const source = readFileSync(new URL("../../packages/pptx-core/dist/index.js", import.meta.url), "utf8");
  assert.doesNotMatch(source, /@aiden0z\/pptx-renderer/);
  assert.doesNotMatch(source, /\b(?:window|document|DOMParser|Worker)\b/);
});

test("PPTX Morph only creates strong unique identity matches", () => {
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

test("PPTX preview uses one archive parse and separates vertical browse from slide playback", () => {
  assert.equal(typeof createPptxDocumentSession, "function");
  assert.equal(typeof createPptxPreviewSession, "function");
  const source = readFileSync(
    new URL("../../packages/pptx-core/src/browser.ts", import.meta.url),
    "utf8",
  );
  assert.match(source, /parseZipLazyMedia\(resolved\.buffer/);
  assert.match(source, /buildPresentation\(files/);
  assert.match(source, /viewer\.load\(presentation\)/);
  assert.match(source, /parsePptxPlaybackDocument\(files, presentation/);
  assert.match(source, /viewer\.renderList\(/);
  assert.match(source, /viewer\.renderSlide\(initialSlide\)/);
  assert.match(source, /renderMode: options\.renderMode \?\? "list"/);
  assert.doesNotMatch(source, /viewer\.open\(/);

  const patch = readFileSync(
    new URL("../../patches/@aiden0z__pptx-renderer@1.2.4.patch", import.meta.url),
    "utf8",
  );
  assert.match(patch, /dataset\.pptxNodeId = t\.id/);
  assert.match(patch, /dataset\.pptxNodeType = t\.nodeType/);
  assert.match(patch, /dataset\.pptxPartPath/);
  assert.match(patch, /dataset\.pptxParagraphIndex = String\(m\)/);
});
