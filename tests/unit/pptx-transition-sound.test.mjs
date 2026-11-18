import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
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

const playerUrl = sourceUrl("packages/pptx-core/src/browser/transition-sound.ts");

const script = `
  import assert from "node:assert/strict";
  const { createPptxTransitionSoundPlayer } = await import(${JSON.stringify(playerUrl)});

  const created = [];
  class FakeAudio {
    constructor() {
      this.src = "";
      this.loop = false;
      this.volume = 1;
      this.paused = true;
      this.playCalls = 0;
      this.rejectWith = undefined;
      this.loadCalls = 0;
      created.push(this);
    }
    play() {
      this.playCalls += 1;
      if (this.rejectWith) return Promise.reject(this.rejectWith);
      this.paused = false;
      return Promise.resolve();
    }
    pause() {
      this.paused = true;
    }
    load() {
      this.loadCalls += 1;
    }
  }
  globalThis.document = { createElement: () => new FakeAudio() };

  function slide(index, sound = true) {
    return {
      index,
      transition: sound
        ? {
            kind: "fade",
            durationMs: 500,
            advanceOnClick: true,
            sound: {
              relationId: "rId9",
              sourcePath: "ppt/media/applause.wav",
              contentType: "audio/wav",
              embedded: true,
              loop: false,
              volume: 0.6,
            },
          }
        : undefined,
    };
  }

  const blocked = [];
  const errors = [];
  let states = 0;
  const released = [];
  const player = createPptxTransitionSoundPlayer({
    onBlocked: (id) => blocked.push(id),
    onError: (id, reason) => errors.push([id, reason.message]),
    onStateChange: () => { states += 1; },
    resolveSource: async (item) => {
      assert.equal(item.kind, "audio");
      assert.equal(item.embedded, true);
      return "blob:transition-sound";
    },
    releaseSource: (_item, url) => released.push(url),
  });

  // 1. Prepare + play resolves the source and starts the audio element.
  player.prepare(slide(0));
  await player.play();
  assert.equal(created.length, 1);
  assert.equal(created[0].src, "blob:transition-sound");
  assert.equal(created[0].playCalls, 1);
  assert.equal(created[0].volume, 0.6);
  assert.equal(blocked.length, 0);
  assert.equal(errors.length, 0);

  // 2. NotAllowedError is surfaced exactly like blocked media and retried.
  player.prepare(slide(1));
  const second = created[1];
  second.rejectWith = new DOMException("autoplay blocked", "NotAllowedError");
  await player.play();
  assert.deepEqual([...player.blockedMediaIds], ["transition-sound:1"]);
  assert.deepEqual(blocked, ["transition-sound:1"]);
  assert.equal(second.playCalls, 1);
  second.rejectWith = undefined;
  await player.resumeBlocked("transition-sound:1");
  assert.equal(second.playCalls, 2);
  assert.deepEqual([...player.blockedMediaIds], []);

  // 3. Controller pause/resume keeps the element paused until resume.
  player.pause();
  assert.equal(second.paused, true);
  await player.resume();
  assert.equal(second.paused, false);

  // 4. A new slide stops and releases the previous source.
  player.prepare(slide(2, false));
  assert.deepEqual(released, ["blob:transition-sound", "blob:transition-sound"]);
  assert.equal(second.paused, true);
  player.dispose();
  assert.equal(states > 0, true);

  process.stdout.write(JSON.stringify({ status: "PASS" }));
`;

test("PPTX transition sounds share the media resolver and autoplay gate", () => {
  const result = spawnSync(process.execPath, [
    "--experimental-strip-types",
    "--import",
    registerHookUrl,
    "--input-type=module",
    "--eval",
    script,
  ], {
    cwd: rootDir,
    encoding: "utf8",
    env: { ...process.env, NODE_NO_WARNINGS: "1" },
  });
  assert.equal(result.status, 0, result.stderr || result.stdout);
  assert.equal(JSON.parse(result.stdout).status, "PASS");
});
