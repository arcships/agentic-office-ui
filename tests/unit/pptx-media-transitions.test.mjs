import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import test from "node:test";
import { fileURLToPath, pathToFileURL } from "node:url";

const rootDir = resolve(dirname(fileURLToPath(import.meta.url)), "../..");
const sourceUrl = (relative) => pathToFileURL(resolve(rootDir, relative)).href;
const resolveHookUrl = `data:text/javascript,${encodeURIComponent(`
  export async function resolve(specifier, context, nextResolve) {
    try {
      return await nextResolve(specifier, context);
    } catch (error) {
      if (specifier.startsWith(".") && !specifier.endsWith(".ts")) {
        return nextResolve(specifier + ".ts", context);
      }
      throw error;
    }
  }
`)}`;
const registerHookUrl = `data:text/javascript,${encodeURIComponent(`
  import { register } from "node:module";
  register(${JSON.stringify(resolveHookUrl)}, import.meta.url);
`)}`;

function runNode(args) {
  return spawnSync(process.execPath, args, {
    cwd: rootDir,
    encoding: "utf8",
    env: { ...process.env, NODE_NO_WARNINGS: "1" },
  });
}

const modules = {
  labTiming: sourceUrl("apps/pptx-playback-lab/src/timing.ts"),
  mediaCommand: sourceUrl("packages/pptx-core/src/playback/media-command.ts"),
  timingValues: sourceUrl("packages/pptx-core/src/playback/timing-values.ts"),
  trackCompiler: sourceUrl("packages/pptx-core/src/playback/track-compiler.ts"),
  transitionPlayer: sourceUrl("packages/pptx-core/src/browser/transition-player.ts"),
};

const childScript = `
  import assert from "node:assert/strict";

  const [labTiming, mediaCommand, timingValues, trackCompiler, transitionPlayer] =
    await Promise.all(${JSON.stringify(Object.values(modules))}.map((url) => import(url)));

  const objectKey = "ppt/slides/slide1.xml|slide|shape:5";

  function mediaItem(values = {}) {
    return Object.freeze({
      id: "rId2",
      objectKey,
      relationId: "rId2",
      kind: "video",
      embedded: true,
      autoplay: false,
      loop: false,
      volume: 1,
      bookmarks: Object.freeze([]),
      ...values,
    });
  }

  function timeNode(id, values = {}) {
    return {
      id,
      container: "parallel",
      kind: "effect",
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

  function slideOf({ nodes, media = [], rootNodeId = "root" }) {
    return {
      index: 0,
      hidden: false,
      objects: [],
      morphFromPrevious: [],
      rootNodeId,
      nodes,
      media,
      actions: [],
      capability: { discovered: 0, strict: 0, approximate: 0, static: 0, unparsed: 0, features: [] },
    };
  }

  // 1. repeatCount 是千分位编码（ECMA-376 §19.5.33），不再按 1000 猜测。
  assert.equal(timingValues.repeatCountValue("1000"), 1);
  assert.equal(timingValues.repeatCountValue("2000"), 2);
  assert.equal(timingValues.repeatCountValue("500"), 0.5);
  assert.equal(timingValues.repeatCountValue("1001"), 1.001);
  assert.equal(timingValues.repeatCountValue("indefinite"), "indefinite");
  assert.equal(timingValues.repeatCountValue("0"), undefined);
  assert.equal(timingValues.repeatCountValue(undefined), undefined);
  assert.match(timingValues.repeatCountNote("2000", 2), /2000/u);
  assert.match(timingValues.repeatCountNote("2000", 2), /2 次/u);
  assert.equal(timingValues.repeatCountNote("indefinite", "indefinite"), undefined);

  // 2. 千分位换算直接决定轨道时长：2000 千分位 = 两次，而不是两千次。
  const repeatSlide = slideOf({
    nodes: {
      root: timeNode("root", { kind: "group", childIds: ["repeat"] }),
      repeat: timeNode("repeat", {
        durationMs: 500,
        repeatCount: timingValues.repeatCountValue("2000"),
        conditions: [{ source: "start", event: "with-previous", delayMs: 0 }],
        effect: { id: "repeat", kind: "fade-in", targetObjectKey: objectKey, presetClass: "entr", values: {} },
      }),
    },
  });
  const repeatGroup = trackCompiler.compilePptxSlideTracks(repeatSlide).groups[0];
  assert.equal(repeatGroup.durationMs, 1000);
  const repeatOpacity = repeatGroup.tracks.find((track) => track.property === "opacity");
  assert.ok(repeatOpacity);
  assert.ok(trackCompiler.evaluatePptxTrack(repeatOpacity, 250) > 0);
  assert.ok(trackCompiler.evaluatePptxTrack(repeatOpacity, 250) < 1);
  assert.equal(trackCompiler.evaluatePptxTrack(repeatOpacity, 1000), 1);

  const singleRepeat = slideOf({
    nodes: {
      root: timeNode("root", { kind: "group", childIds: ["repeat"] }),
      repeat: timeNode("repeat", {
        durationMs: 500,
        repeatCount: timingValues.repeatCountValue("1000"),
        conditions: [{ source: "start", event: "with-previous", delayMs: 0 }],
        effect: { id: "repeat", kind: "fade-in", targetObjectKey: objectKey, presetClass: "entr", values: {} },
      }),
    },
  });
  assert.equal(trackCompiler.compilePptxSlideTracks(singleRepeat).groups[0].durationMs, 500);

  // 3. 自动播放媒体在时间零就产生 media-playback/media-volume 轨道。
  const autoplaySlide = slideOf({ media: [mediaItem({ autoplay: true, volume: 0.4 })], nodes: {} });
  const autoplayCompiled = trackCompiler.compilePptxSlideTracks(autoplaySlide);
  assert.equal(autoplayCompiled.groups[0].key, "auto");
  assert.equal(autoplayCompiled.groups[0].durationMs, 0);
  const autoplayState = trackCompiler.evaluatePptxTriggerGroup(autoplayCompiled.groups[0], 0)[objectKey];
  assert.equal(autoplayState["media-playback"], "play");
  assert.equal(autoplayState["media-volume"], 0.4);

  // 没有自动播放时不会凭空建立自动组，也不重复生成播放轨道。
  const silentSlide = slideOf({ media: [mediaItem({ autoplay: false, volume: 0.7 })], nodes: {} });
  assert.deepEqual(trackCompiler.compilePptxSlideTracks(silentSlide).groups, []);

  // 4. playFrom 命令生成 media-time 定位轨道。
  const seekSlide = slideOf({
    media: [mediaItem({ autoplay: true })],
    nodes: {
      root: timeNode("root", { kind: "group", childIds: ["seek"] }),
      seek: timeNode("seek", {
        kind: "media",
        durationMs: 1,
        conditions: [{ source: "start", event: "with-previous", delayMs: 0 }],
        effect: {
          id: "seek",
          kind: "media-command",
          targetObjectKey: objectKey,
          presetClass: "mediacall",
          command: "playFrom(1.5)",
          values: {},
        },
      }),
    },
  });
  const seekGroup = trackCompiler.compilePptxSlideTracks(seekSlide).groups[0];
  assert.equal(seekGroup.key, "auto");
  const seekState = trackCompiler.evaluatePptxTriggerGroup(seekGroup, 1)[objectKey];
  assert.equal(seekState["media-time"], 1500);
  assert.equal(seekState["media-playback"], "playFrom(1.5)");
  assert.equal(mediaCommand.normalizeMediaCommand(seekState["media-playback"]), "play");
  assert.equal(seekGroup.tracks.filter((track) => track.property === "media-playback").length, 1);

  // 5. 媒体命令与循环策略。
  assert.equal(mediaCommand.normalizeMediaCommand("togglePause"), "toggle-pause");
  assert.equal(mediaCommand.normalizeMediaCommand("playFrom(2.5)"), "play");
  assert.equal(mediaCommand.normalizeMediaCommand("resume"), "play");
  assert.equal(mediaCommand.normalizeMediaCommand("stop"), "stop");
  assert.equal(mediaCommand.normalizeMediaCommand(""), undefined);
  assert.equal(mediaCommand.mediaSeekMs("playFrom(1.5)"), 1500);
  assert.equal(mediaCommand.mediaSeekMs(" playFrom( 2 ) "), 2000);
  assert.equal(mediaCommand.mediaSeekMs("playFrom(-3)"), 0);
  assert.equal(mediaCommand.mediaSeekMs("play"), undefined);
  assert.equal(mediaCommand.isMediaPlayCommand("play"), true);
  assert.equal(mediaCommand.isMediaPlayCommand("playFrom(4)"), true);
  assert.equal(mediaCommand.isMediaPlayCommand("togglePause"), false);
  assert.equal(mediaCommand.isMediaPlayCommand("stop"), false);

  assert.deepEqual(mediaCommand.pptxMediaLoopPolicy({ loop: true }), { nativeLoop: true });
  assert.deepEqual(
    mediaCommand.pptxMediaLoopPolicy({ loop: true, trimStartMs: 1000 }),
    { nativeLoop: false, restartMs: 1000 },
  );
  assert.deepEqual(
    mediaCommand.pptxMediaLoopPolicy({ loop: true, trimEndMs: 1100 }),
    { nativeLoop: false, restartMs: 0 },
  );
  assert.deepEqual(
    mediaCommand.pptxMediaLoopPolicy({ loop: true, trimStartMs: 1000, trimEndMs: 1100 }),
    { nativeLoop: false, restartMs: 1000 },
  );
  assert.deepEqual(mediaCommand.pptxMediaLoopPolicy({ loop: false, trimStartMs: 1000 }), { nativeLoop: false });

  // 自动播放判定使用与时间表一致的规则：只有落在自动组里的播放命令算自动播放。
  const autoplayMedia = [mediaItem({ autoplay: false })];
  const autoplayNodes = {
    root: timeNode("root", { kind: "group", childIds: ["play"] }),
    play: timeNode("play", {
      kind: "media",
      durationMs: 1,
      conditions: [{ source: "start", event: "with-previous", delayMs: 0 }],
      effect: { id: "play", kind: "media-command", targetObjectKey: objectKey, presetClass: "mediacall", command: "play", values: {} },
    }),
  };
  assert.equal(mediaCommand.withAutoplayMedia(autoplayMedia, autoplayNodes, "root")[0].autoplay, true);

  const clickNodes = {
    root: timeNode("root", { kind: "group", childIds: ["click"] }),
    click: timeNode("click", {
      kind: "group",
      conditions: [{ source: "start", event: "on-click", delayMs: 0 }],
      childIds: ["play"],
    }),
    play: timeNode("play", {
      kind: "media",
      durationMs: 1,
      conditions: [{ source: "start", event: "with-previous", delayMs: 0 }],
      effect: { id: "play", kind: "media-command", targetObjectKey: objectKey, presetClass: "mediacall", command: "play", values: {} },
    }),
  };
  assert.equal(mediaCommand.withAutoplayMedia(autoplayMedia, clickNodes, "root")[0].autoplay, false);
  assert.equal(mediaCommand.withAutoplayMedia(autoplayMedia, {}, "root")[0].autoplay, false);

  // 6. Morph 粒度、匹配可信度和逐词配对。
  assert.equal(transitionPlayer.pptxMorphGranularity({ kind: "morph", option: "byWord", durationMs: 1000, advanceOnClick: true }), "word");
  assert.equal(transitionPlayer.pptxMorphGranularity({ kind: "morph", option: "byChar", durationMs: 1000, advanceOnClick: true }), "word");
  assert.equal(transitionPlayer.pptxMorphGranularity({ kind: "morph", option: "byObject", durationMs: 1000, advanceOnClick: true }), "object");
  assert.equal(transitionPlayer.pptxMorphGranularity(undefined), "object");

  const match = (values) => ({ from: "from", to: "to", method: "name-and-geometry", confidence: "medium", score: 0.5, unique: true, ...values });
  assert.equal(transitionPlayer.isPptxMorphMatchAnimatable(match({ confidence: "strong" })), true);
  assert.equal(transitionPlayer.isPptxMorphMatchAnimatable(match({ confidence: "medium" })), true);
  assert.equal(transitionPlayer.isPptxMorphMatchAnimatable(match({ confidence: "weak" })), false);
  assert.equal(transitionPlayer.isPptxMorphMatchAnimatable(match({ confidence: "medium", unique: false })), false);

  assert.deepEqual(transitionPlayer.matchPptxMorphWords("alpha beta", "beta alpha"), [
    { from: 1, to: 0, text: "beta" },
    { from: 0, to: 1, text: "alpha" },
  ]);
  assert.deepEqual(transitionPlayer.matchPptxMorphWords("the the cat", "the cat"), [
    { from: 0, to: 0, text: "the" },
    { from: 2, to: 1, text: "cat" },
  ]);
  assert.deepEqual(transitionPlayer.matchPptxMorphWords("完全不同的文字", "totally different"), []);

  // 逐词位移：进度 0 回到旧页同名词的位置，进度 1 归零；不同词得到不同位移。
  assert.deepEqual(
    transitionPlayer.pptxMorphWordTranslate({ left: 100, top: 50 }, { left: 40, top: 10 }, 0),
    { x: 60, y: 40 },
  );
  assert.deepEqual(
    transitionPlayer.pptxMorphWordTranslate({ left: 100, top: 50 }, { left: 40, top: 10 }, 1),
    { x: 0, y: 0 },
  );
  assert.deepEqual(
    transitionPlayer.pptxMorphWordTranslate({ left: 100, top: 50 }, { left: 40, top: 10 }, 0.5),
    { x: 30, y: 20 },
  );
  const firstWord = transitionPlayer.pptxMorphWordTranslate({ left: 10, top: 0 }, { left: 0, top: 0 }, 0);
  const secondWord = transitionPlayer.pptxMorphWordTranslate({ left: 90, top: 0 }, { left: 30, top: 0 }, 0);
  assert.notDeepEqual(firstWord, secondWord);

  // 7. 实验室：无限重复用 Infinity 表示，批次不再自行结束。
  const indefiniteEffect = (values = {}) => ({
    id: "indefinite",
    targetId: "2",
    presetClass: "emph",
    presetId: "6",
    nodeType: "withEffect",
    trigger: "withPrevious",
    duration: 500,
    delay: 0,
    start: 0,
    filter: "",
    transition: "",
    command: "",
    kind: "emphasis",
    motionPath: "",
    rotationBy: 0,
    scaleBy: null,
    scaleFrom: null,
    scaleTo: null,
    attributes: [],
    values: [],
    paragraphRange: null,
    repeatCount: "indefinite",
    autoReverse: false,
    acceleration: 0,
    deceleration: 0,
    ...values,
  });

  assert.equal(labTiming.parseRepeatCount("2000"), 2);
  assert.equal(labTiming.parseRepeatCount("500"), 0.5);
  assert.equal(labTiming.parseRepeatCount("indefinite"), "indefinite");
  assert.equal(labTiming.parseRepeatCount(null), null);
  assert.equal(labTiming.effectiveDuration(indefiniteEffect()), Infinity);
  assert.equal(labTiming.effectiveDuration(indefiniteEffect({ autoReverse: true })), Infinity);
  assert.equal(labTiming.effectiveDuration(indefiniteEffect({ repeatCount: 2, autoReverse: true })), 2000);

  const indefiniteBatch = labTiming.scheduleBatch(0, [indefiniteEffect()]);
  assert.equal(indefiniteBatch.indefinite, true);
  assert.equal(indefiniteBatch.duration, Infinity);
  const followed = labTiming.scheduleBatch(0, [
    indefiniteEffect(),
    indefiniteEffect({ id: "after", trigger: "afterPrevious", repeatCount: null }),
  ]);
  assert.equal(followed.indefinite, true);
  assert.equal(followed.effects[1].start, Infinity);
  const finiteBatch = labTiming.scheduleBatch(1, [indefiniteEffect({ repeatCount: null })]);
  assert.equal(finiteBatch.indefinite, false);
  assert.equal(finiteBatch.duration, 500);

  process.stdout.write(JSON.stringify({ status: "PASS" }));
`;

test("PPTX repeat, media command and morph contracts follow the ECMA timing encoding", () => {
  const result = runNode([
    "--experimental-strip-types",
    "--import",
    registerHookUrl,
    "--input-type=module",
    "--eval",
    childScript,
  ]);
  assert.equal(result.status, 0, result.stderr || result.stdout);
  const evidence = JSON.parse(result.stdout);
  assert.equal(evidence.status, "PASS");
});

const mediaPlayerScript = `
  import assert from "node:assert/strict";

  const { createPptxMediaPlayer } = await import(${JSON.stringify(sourceUrl("packages/pptx-core/src/browser/media-player.ts"))});

  const objectKey = "ppt/slides/slide1.xml|slide|shape:5";

  function fakeMediaElement() {
    let currentTime = 0;
    let seeks = 0;
    const listeners = new Map();
    const element = {
      dataset: { pptxObjectKey: objectKey },
      loop: false,
      volume: 1,
      paused: true,
      duration: 1.2,
      playCalls: 0,
      matches: () => true,
      querySelector: () => null,
      get currentTime() { return currentTime; },
      set currentTime(value) { currentTime = value; seeks += 1; },
      get seekCount() { return seeks; },
      play() { element.playCalls += 1; element.paused = false; return Promise.resolve(); },
      pause() { element.paused = true; },
      addEventListener(type, listener) {
        listeners.set(type, [...(listeners.get(type) ?? []), listener]);
      },
      removeEventListener(type, listener) {
        listeners.set(type, (listeners.get(type) ?? []).filter((candidate) => candidate !== listener));
      },
      dispatch(type) {
        for (const listener of [...(listeners.get(type) ?? [])]) listener();
      },
    };
    return element;
  }

  function setup(itemValues = {}) {
    const element = fakeMediaElement();
    const item = Object.freeze({
      id: "rId2",
      objectKey,
      relationId: "rId2",
      kind: "audio",
      embedded: true,
      autoplay: false,
      loop: false,
      volume: 1,
      bookmarks: Object.freeze([]),
      ...itemValues,
    });
    const root = { querySelectorAll: (selector) => (selector === "[data-pptx-object-key]" ? [element] : []) };
    const player = createPptxMediaPlayer(root, {
      onBlocked: () => undefined,
      onError: (_id, error) => { throw error; },
      onBookmark: () => undefined,
      onStateChange: () => undefined,
    });
    player.prepare({ index: 0, hidden: false, objects: [], morphFromPrevious: [], nodes: {}, media: [item], actions: [], capability: {} });
    return { element, item, player };
  }

  // 循环 + 只有裁剪起点：关闭原生 loop，每次循环都回到裁剪起点。
  {
    const { element } = setup({ loop: true, trimStartMs: 1000 });
    assert.equal(element.loop, false);
    assert.equal(element.currentTime, 1);
    element.currentTime = 1.19;
    const seeksBefore = element.seekCount;
    element.dispatch("ended");
    assert.equal(element.currentTime, 1);
    assert.equal(element.seekCount, seeksBefore + 1);
    assert.equal(element.playCalls, 1);
  }

  // 无裁剪的循环保留原生 loop。
  {
    const { element } = setup({ loop: true });
    assert.equal(element.loop, true);
  }

  // 裁剪终点仍然由 timeupdate 手动回绕。
  {
    const { element } = setup({ loop: true, trimStartMs: 100, trimEndMs: 1100 });
    assert.equal(element.loop, false);
    assert.equal(element.currentTime, 0.1);
    element.currentTime = 0.15;
    element.dispatch("timeupdate");
    assert.equal(element.currentTime, 0.1);
    assert.equal(element.playCalls, 1);
  }

  // 音量轨道改变元素音量。
  {
    const { element, player } = setup({ volume: 1 });
    player.apply({ [objectKey]: { "media-volume": 0.5 } });
    assert.equal(element.volume, 0.5);
  }

  // media-time 只在请求值变化时定位一次，不会每帧把播放拉回起点。
  {
    const { element, player } = setup();
    const seeksBefore = element.seekCount;
    player.apply({ [objectKey]: { "media-time": 5000 } });
    assert.equal(element.currentTime, 5);
    player.apply({ [objectKey]: { "media-time": 5000 } });
    assert.equal(element.seekCount, seeksBefore + 1);
    player.apply({ [objectKey]: { "media-time": 6000 } });
    assert.equal(element.currentTime, 6);
    assert.equal(element.seekCount, seeksBefore + 2);
  }

  // 播放命令去重，重复求值不会反复调用 play。
  {
    const { element, player } = setup({ trimStartMs: 100 });
    player.apply({ [objectKey]: { "media-playback": "play" } });
    player.apply({ [objectKey]: { "media-playback": "play" } });
    assert.equal(element.playCalls, 1);
    player.apply({ [objectKey]: { "media-playback": "stop" } });
    assert.equal(element.paused, true);
    assert.equal(element.currentTime, 0.1);
  }

  process.stdout.write(JSON.stringify({ status: "PASS" }));
`;

test("PPTX media player loops trimmed media from the trim start and applies volume/seek tracks", () => {
  const result = runNode([
    "--experimental-strip-types",
    "--import",
    registerHookUrl,
    "--input-type=module",
    "--eval",
    mediaPlayerScript,
  ]);
  assert.equal(result.status, 0, result.stderr || result.stdout);
  const evidence = JSON.parse(result.stdout);
  assert.equal(evidence.status, "PASS");
});

test("PPTX playback lab keeps indefinite batches alive until the next click", () => {
  const main = readFileSync(resolve(rootDir, "apps/pptx-playback-lab/src/main.ts"), "utf8");
  assert.match(main, /effect\.repeatCount === "indefinite" \? Number\.POSITIVE_INFINITY/u);
  assert.doesNotMatch(main, /effect\.repeatCount === "indefinite" \? 2/u);
  assert.match(main, /if \(batch\.indefinite\) \{/u);
  assert.match(main, /timing\.batches\.some\(\(batch\) => batch\.indefinite\)/u);

  const timing = readFileSync(resolve(rootDir, "apps/pptx-playback-lab/src/timing.ts"), "utf8");
  assert.match(timing, /Number\.POSITIVE_INFINITY/u);
  assert.doesNotMatch(timing, /repeatCount === "indefinite" \? 2 :/u);
  assert.match(timing, /thousandths \/ 1000/u);
});
