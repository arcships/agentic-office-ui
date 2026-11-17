import assert from "node:assert/strict";
import test from "node:test";

import { importFromDemo } from "../component/vue-test-renderer.mjs";

const {
  createPptxPlaybackController,
} = await importFromDemo("@arcships/pptx-core/browser");

function createRoot() {
  return {
    firstElementChild: null,
    querySelectorAll() {
      return [];
    },
  };
}

function capability() {
  return {
    approximate: 0,
    discovered: 0,
    features: [],
    static: 0,
    strict: 0,
    unparsed: 0,
  };
}

function slide(index, extra = {}) {
  return {
    actions: [],
    capability: capability(),
    hidden: false,
    index,
    media: [],
    morphFromPrevious: [],
    objects: [],
    ...extra,
  };
}

function documentWithSlides(count) {
  return {
    capability: capability(),
    slides: Array.from({ length: count }, (_, index) => slide(index)),
  };
}

function createHost(document, overrides = {}) {
  return {
    document,
    initialSlideIndex: 0,
    renderSlide: async () => {},
    root: createRoot(),
    ...overrides,
  };
}

function deferred() {
  let resolve;
  const promise = new Promise((nextResolve) => {
    resolve = nextResolve;
  });
  return { promise, resolve };
}

test("superseded navigation commands reject with a cancelled error", async () => {
  const firstRender = deferred();
  const rendered = [];
  const host = createHost(documentWithSlides(4), {
    renderSlide(index) {
      rendered.push(index);
      return index === 1 ? firstRender.promise : Promise.resolve();
    },
  });
  const controller = createPptxPlaybackController(host, { autoplay: false });

  try {
    const first = controller.goToSlide(1);
    const superseded = controller.goToSlide(2);
    const last = controller.goToSlide(3);

    await assert.rejects(superseded, (error) => error.code === "COMMAND_CANCELLED");
    firstRender.resolve();
    await first;
    await last;

    assert.deepEqual(rendered, [1, 3]);
    assert.equal(controller.snapshot.slideIndex, 3);
    assert.equal(controller.snapshot.status, "ready");
  } finally {
    controller.dispose();
  }
});

test("successful navigation transitions from transitioning to ready", async () => {
  const render = deferred();
  const host = createHost(documentWithSlides(3), {
    renderSlide() {
      return render.promise;
    },
  });
  const controller = createPptxPlaybackController(host, { autoplay: false });
  const events = [];
  controller.subscribe((event) => events.push(event));

  try {
    const navigation = controller.goToSlide(1);
    assert.equal(controller.snapshot.status, "transitioning");
    assert.equal(controller.snapshot.generation, 2);

    render.resolve();
    await navigation;

    assert.equal(controller.snapshot.status, "ready");
    assert.equal(controller.snapshot.slideIndex, 1);
    const slidechange = events.find((event) => event.type === "slidechange");
    assert.deepEqual(
      { from: slidechange?.from, reason: slidechange?.reason, to: slidechange?.to },
      { from: 0, reason: "direct", to: 1 },
    );
  } finally {
    controller.dispose();
  }
});

test("disposing rejects navigation commands that are still queued", async () => {
  const render = deferred();
  const host = createHost(documentWithSlides(3), {
    renderSlide() {
      return render.promise;
    },
  });
  const controller = createPptxPlaybackController(host, { autoplay: false });

  const running = controller.goToSlide(1);
  const queued = controller.goToSlide(2);
  controller.dispose();
  render.resolve();

  await assert.rejects(running, (error) => error.code === "PLAYBACK_DISPOSED");
  await assert.rejects(queued, (error) => error.code === "PLAYBACK_DISPOSED");
  assert.equal(controller.snapshot.status, "disposed");
});

test("a failed slide render restores the previous status and emits an error", async () => {
  let fail = true;
  const host = createHost(documentWithSlides(3), {
    async renderSlide(index) {
      if (index === 1 && fail) throw new Error("render exploded");
    },
  });
  const controller = createPptxPlaybackController(host, { autoplay: false });
  const events = [];
  controller.subscribe((event) => events.push(event));

  try {
    await assert.rejects(
      controller.goToSlide(1),
      (error) => error.code === "RENDER_FAILED" && error.slideIndex === 1,
    );
    assert.equal(controller.snapshot.status, "ready");
    assert.equal(controller.snapshot.slideIndex, 0);
    const failure = events.find((event) => event.type === "error");
    assert.equal(failure?.error.code, "RENDER_FAILED");
    assert.equal(failure?.error.slideIndex, 1);
    assert.equal(events.some((event) => event.type === "slidechange"), false);

    fail = false;
    await controller.goToSlide(1);
    assert.equal(controller.snapshot.status, "ready");
    assert.equal(controller.snapshot.slideIndex, 1);
  } finally {
    controller.dispose();
  }
});

test("retained playback events are bounded and replay at most the cap", () => {
  const warningCount = 1000;
  const nodes = {
    root: {
      acceleration: 0,
      autoReverse: false,
      childIds: Array.from({ length: warningCount }, (_, index) => `effect-${index}`),
      conditions: [],
      container: "root",
      deceleration: 0,
      delayMs: 0,
      durationMs: 0,
      fill: "hold",
      id: "root",
      kind: "group",
      restart: "never",
    },
  };
  for (let index = 0; index < warningCount; index += 1) {
    nodes[`effect-${index}`] = {
      acceleration: 0,
      autoReverse: false,
      childIds: [],
      conditions: [{
        delayMs: 0,
        event: index === 0 ? "with-previous" : "after-previous",
        source: "start",
      }],
      container: "behavior",
      deceleration: 0,
      delayMs: 0,
      durationMs: 0,
      effect: {
        id: `animation-${index}`,
        kind: "appear",
        targetObjectKey: `missing-object-${index}`,
        values: {},
      },
      fill: "hold",
      id: `effect-${index}`,
      kind: "effect",
      restart: "never",
    };
  }
  const host = createHost({
    capability: capability(),
    slides: [slide(0, { nodes, rootNodeId: "root" })],
  });
  const controller = createPptxPlaybackController(host, { autoplay: false });

  try {
    const warnings = [];
    controller.subscribe((event) => {
      if (event.type === "warning") warnings.push(event.warning);
    });
    assert.equal(warnings.length, 50);
    assert.equal(warnings[0].objectKey, `missing-object-${warningCount - 50}`);
    assert.equal(warnings.at(-1).objectKey, `missing-object-${warningCount - 1}`);

    const secondWarnings = [];
    controller.subscribe((event) => {
      if (event.type === "warning") secondWarnings.push(event.warning);
    });
    assert.equal(secondWarnings.length, 50);
  } finally {
    controller.dispose();
  }
});
