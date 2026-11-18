import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { dirname, resolve } from "node:path";
import test from "node:test";
import { fileURLToPath, pathToFileURL } from "node:url";

const rootDir = resolve(dirname(fileURLToPath(import.meta.url)), "../..");
const sourceUrl = (relative) => pathToFileURL(resolve(rootDir, relative)).href;
// The PPTX source package imports its own modules without file extensions, so
// the child process registers the resolver hook the other PPTX unit tests use.
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

const helpersScript = `
  import assert from "node:assert/strict";

  const transitions = await import(${JSON.stringify(sourceUrl("packages/pptx-core/src/playback/transitions.ts"))});

  // 1. Capability table: the four originals plus cover/uncover/split are exact,
  //    dissolve and morph are approximate, 3D kinds stay static, unknown kinds
  //    only run under approximation: "safe".
  for (const kind of ["cut", "fade", "push", "wipe", "cover", "uncover", "split"]) {
    const capability = transitions.pptxTransitionCapability(kind, "off");
    assert.equal(capability.disposition, "strict", kind);
    assert.equal(capability.feature, \`transition:\${kind}\`, kind);
    assert.equal(capability.executable, true, kind);
    assert.equal(capability.reason, undefined, kind);
    assert.equal(transitions.isPptxTransitionExecutable(kind, "off"), true, kind);
    assert.equal(transitions.isPptxTransitionExecutable(kind.toUpperCase(), "off"), true, kind);
  }
  for (const kind of ["dissolve", "morph"]) {
    const strict = transitions.pptxTransitionCapability(kind, "off");
    assert.equal(strict.disposition, "static", kind);
    assert.equal(strict.executable, false, kind);
    const safe = transitions.pptxTransitionCapability(kind, "safe");
    assert.equal(safe.disposition, "approximate", kind);
    assert.equal(safe.executable, true, kind);
    assert.equal(safe.feature, \`transition:\${kind}\`, kind);
    assert.ok(safe.reason, kind);
  }
  // 3D transitions never degrade to a 2D fade, even with approximations on.
  for (const kind of ["cube", "gallery", "pan"]) {
    const capability = transitions.pptxTransitionCapability(kind, "safe");
    assert.equal(capability.disposition, "static", kind);
    assert.equal(capability.executable, false, kind);
    assert.equal(capability.feature, \`transition:3d:\${kind}\`, kind);
  }
  const unknown = transitions.pptxTransitionCapability("zoom", "safe");
  assert.equal(unknown.disposition, "approximate");
  assert.equal(unknown.executable, true);
  assert.equal(transitions.pptxTransitionCapability("zoom", "off").executable, false);
  assert.equal(transitions.isPptxTransitionExecutable(undefined, "safe"), false);

  // 2. Direction letters name the edge the incoming slide enters from.
  assert.deepEqual(transitions.pptxTransitionDirection(undefined), { x: 1, y: 0 });
  assert.deepEqual(transitions.pptxTransitionDirection("l"), { x: 1, y: 0 });
  assert.deepEqual(transitions.pptxTransitionDirection("RIGHT"), { x: -1, y: 0 });
  assert.deepEqual(transitions.pptxTransitionDirection("u"), { x: 0, y: 1 });
  assert.deepEqual(transitions.pptxTransitionDirection("down"), { x: 0, y: -1 });
  assert.deepEqual(transitions.pptxTransitionDirection("lu"), { x: 1, y: 1 });
  assert.deepEqual(transitions.pptxTransitionDirection("ld"), { x: 1, y: -1 });
  assert.deepEqual(transitions.pptxTransitionDirection("ru"), { x: -1, y: 1 });
  assert.deepEqual(transitions.pptxTransitionDirection("rd"), { x: -1, y: -1 });

  // 3. Wipe keeps the band it already painted on the named edge.
  assert.equal(transitions.pptxTransitionWipeClip(undefined, 0), "inset(0% 100% 0% 0%)");
  assert.equal(transitions.pptxTransitionWipeClip("l", 0.5), "inset(0% 50% 0% 0%)");
  assert.equal(transitions.pptxTransitionWipeClip("r", 0.5), "inset(0% 0% 0% 50%)");
  assert.equal(transitions.pptxTransitionWipeClip("u", 0.25), "inset(75% 0% 0% 0%)");
  assert.equal(transitions.pptxTransitionWipeClip("d", 1), "inset(0% 0% 0% 0%)");
  assert.equal(transitions.pptxTransitionWipeClip("lu", 0.5), "inset(0% 50% 50% 0%)");
  assert.equal(transitions.pptxTransitionWipeClip("ld", 0.5), "inset(50% 50% 0% 0%)");
  assert.equal(transitions.pptxTransitionWipeClip("ru", 0.5), "inset(0% 0% 50% 50%)");
  assert.equal(transitions.pptxTransitionWipeClip("rd", 0.5), "inset(50% 0% 0% 50%)");
  assert.deepEqual(transitions.pptxTransitionWipeInsets("rd", 0), [100, 0, 0, 100]);

  // 4. Split: the orientation names the split line, the direction grows
  //    outwards or inwards.
  assert.equal(
    transitions.pptxTransitionSplitMask("vert", "out", 0.5),
    "linear-gradient(to right, transparent 0 25%, #000 25% 75%, transparent 75% 100%)",
  );
  assert.equal(
    transitions.pptxTransitionSplitMask(undefined, "out", 1),
    "linear-gradient(to right, transparent 0 0%, #000 0% 100%, transparent 100% 100%)",
  );
  assert.equal(
    transitions.pptxTransitionSplitMask("horz", "in", 0.25),
    "linear-gradient(to bottom, #000 0 12.5%, transparent 12.5% 87.5%, #000 87.5% 100%)",
  );
  assert.equal(
    transitions.pptxTransitionSplitMask("vert", "in", 0),
    "linear-gradient(to right, #000 0 0%, transparent 0% 100%, #000 100% 100%)",
  );

  // 5. Progress is clamped into 0…1 and a zero duration counts as finished.
  assert.equal(transitions.pptxTransitionProgress(-50, 1000), 0);
  assert.equal(transitions.pptxTransitionProgress(500, 1000), 0.5);
  assert.equal(transitions.pptxTransitionProgress(5000, 1000), 1);
  assert.equal(transitions.pptxTransitionProgress(500, 0), 1);

  process.stdout.write(JSON.stringify({ status: "PASS" }));
`;

const playerScript = `
  import assert from "node:assert/strict";

  const { mountPptxTransition } = await import(${JSON.stringify(sourceUrl("packages/pptx-core/src/browser/transition-player.ts"))});

  // 最小 DOM：只覆盖 transition-player 使用的节点、样式和矩形。
  function createDom() {
    const style = () => ({
      opacity: "", translate: "", scale: "", clipPath: "", maskImage: "", maskRepeat: "",
      maskSize: "", willChange: "", position: "", zIndex: "", inset: "", pointerEvents: "",
    });
    function element(tag) {
      const node = {
        tag,
        style: style(),
        dataset: {},
        children: [],
        parentNode: null,
        get firstElementChild() {
          return node.children[0] ?? null;
        },
        appendChild(child) {
          child.parentNode = node;
          node.children.push(child);
          return child;
        },
        remove() {
          if (!node.parentNode) return;
          const index = node.parentNode.children.indexOf(node);
          if (index >= 0) node.parentNode.children.splice(index, 1);
          node.parentNode = null;
        },
        querySelectorAll(selector) {
          const result = [];
          const visit = (current) => {
            for (const child of current.children) {
              if (selector === "[data-pptx-object-key]" && child.dataset.pptxObjectKey) result.push(child);
              visit(child);
            }
          };
          visit(node);
          return result;
        },
        getBoundingClientRect() {
          return { left: 0, top: 0, width: 960, height: 540, right: 960, bottom: 540 };
        },
      };
      return node;
    }
    return {
      createElement: (tag) => element(tag),
      createTreeWalker: () => ({ currentNode: null, nextNode: () => null }),
    };
  }

  const doc = createDom();
  globalThis.document = doc;
  globalThis.getComputedStyle = () => ({ position: "static" });

  function setup(transition) {
    const root = doc.createElement("div");
    const incoming = doc.createElement("section");
    root.appendChild(incoming);
    const outgoing = doc.createElement("div");
    const capture = { layer: outgoing, objects: new Map() };
    return { root, incoming, outgoing, capture, transition: { durationMs: 1000, advanceOnClick: true, ...transition } };
  }

  // 1. Cover: the new slide slides in over the old one, which never moves.
  {
    const { root, incoming, outgoing, capture, transition } = setup({ kind: "cover", direction: "l" });
    const handle = mountPptxTransition(root, capture, transition);
    assert.ok(handle);
    handle.apply(0);
    assert.equal(incoming.style.translate, "960px 0px");
    assert.equal(outgoing.style.translate, "0px 0px");
    handle.apply(500);
    assert.equal(incoming.style.translate, "480px 0px");
    handle.apply(1000);
    assert.equal(incoming.style.translate, "0px 0px");
    handle.finish();
    assert.equal(incoming.style.translate, "");
    assert.equal(outgoing.parentNode, null);
  }

  // 2. Uncover: the old slide slides away and reveals the new one underneath.
  {
    const { root, incoming, outgoing, capture, transition } = setup({ kind: "uncover", direction: "rd" });
    const handle = mountPptxTransition(root, capture, transition);
    assert.ok(handle);
    handle.apply(0);
    assert.equal(incoming.style.translate, "0px 0px");
    assert.equal(outgoing.style.translate, "0px 0px");
    handle.apply(500);
    assert.equal(outgoing.style.translate, "480px 270px");
    handle.apply(1000);
    assert.equal(outgoing.style.translate, "960px 540px");
    handle.finish();
  }

  // 3. Push keeps both layers moving and now follows diagonal directions.
  {
    const { root, incoming, outgoing, capture, transition } = setup({ kind: "push", direction: "up" });
    const handle = mountPptxTransition(root, capture, transition);
    assert.ok(handle);
    handle.apply(0);
    assert.equal(incoming.style.translate, "0px 540px");
    assert.equal(outgoing.style.translate, "0px 0px");
    handle.apply(750);
    assert.equal(incoming.style.translate, "0px 135px");
    handle.finish();
  }

  // 4. Wipe covers every corner direction through the shared insets.
  {
    const { root, incoming, capture, transition } = setup({ kind: "wipe", direction: "lu" });
    const handle = mountPptxTransition(root, capture, transition);
    assert.ok(handle);
    handle.apply(0);
    assert.equal(incoming.style.clipPath, "inset(0% 100% 100% 0%)");
    handle.apply(500);
    assert.equal(incoming.style.clipPath, "inset(0% 50% 50% 0%)");
    handle.apply(1000);
    assert.equal(incoming.style.clipPath, "inset(0% 0% 0% 0%)");
    handle.finish();
    assert.equal(incoming.style.clipPath, "");
  }

  // 5. Split masks the incoming slide and restores the mask when it finishes.
  {
    const { root, incoming, capture, transition } = setup({ kind: "split", orientation: "horz", direction: "out" });
    const handle = mountPptxTransition(root, capture, transition);
    assert.ok(handle);
    handle.apply(500);
    assert.equal(
      incoming.style.maskImage,
      "linear-gradient(to bottom, transparent 0 25%, #000 25% 75%, transparent 75% 100%)",
    );
    assert.equal(incoming.style.maskRepeat, "no-repeat");
    assert.equal(incoming.style.maskSize, "100% 100%");
    handle.finish();
    assert.equal(incoming.style.maskImage, "");
    assert.equal(incoming.style.maskRepeat, "");
    assert.equal(incoming.style.maskSize, "");
  }

  // 6. Dissolve is a cross fade; unknown kinds fall back to the same fade.
  for (const kind of ["dissolve", "zoom"]) {
    const { root, incoming, outgoing, capture, transition } = setup({ kind });
    const handle = mountPptxTransition(root, capture, transition);
    assert.ok(handle, kind);
    handle.apply(0);
    assert.equal(incoming.style.opacity, "0", kind);
    assert.equal(outgoing.style.opacity, "1", kind);
    handle.apply(500);
    assert.equal(incoming.style.opacity, "0.5", kind);
    assert.equal(outgoing.style.opacity, "0.5", kind);
    handle.apply(1000);
    assert.equal(incoming.style.opacity, "1", kind);
    handle.finish();
  }

  // 7. A cut (or a zero-length transition) swaps slides without a layer.
  {
    const { root, capture, transition } = setup({ kind: "cut" });
    assert.equal(mountPptxTransition(root, capture, transition), null);
    assert.equal(mountPptxTransition(root, capture, { ...transition, kind: "Fade", durationMs: 0 }), null);
    assert.equal(mountPptxTransition(root, null, transition), null);
  }

  process.stdout.write(JSON.stringify({ status: "PASS" }));
`;

test("PPTX transition capability, directions and split masks are exact", () => {
  const result = runNode([
    "--experimental-strip-types",
    "--import",
    registerHookUrl,
    "--input-type=module",
    "--eval",
    helpersScript,
  ]);
  assert.equal(result.status, 0, result.stderr || result.stdout);
  assert.equal(JSON.parse(result.stdout).status, "PASS");
});

test("cover, uncover, split, dissolve and diagonal wipes render on the transition layer", () => {
  const result = runNode([
    "--experimental-strip-types",
    "--import",
    registerHookUrl,
    "--input-type=module",
    "--eval",
    playerScript,
  ]);
  assert.equal(result.status, 0, result.stderr || result.stdout);
  assert.equal(JSON.parse(result.stdout).status, "PASS");
});
