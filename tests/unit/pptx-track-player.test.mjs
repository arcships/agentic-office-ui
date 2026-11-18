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

function runNode(args) {
  return spawnSync(process.execPath, args, {
    cwd: rootDir,
    encoding: "utf8",
    env: { ...process.env, NODE_NO_WARNINGS: "1" },
  });
}

const trackPlayerUrl = sourceUrl("packages/pptx-core/src/browser/track-player.ts");

const script = `
  import assert from "node:assert/strict";

  const track = await import(${JSON.stringify(trackPlayerUrl)});

  function element(kind) {
    const node = {
      kind,
      style: {
        visibility: "", opacity: "", pointerEvents: "", translate: "", scale: "", rotate: "",
        clipPath: "", filter: "", backgroundColor: "", borderColor: "", color: "", display: "",
      },
      dataset: {},
      children: [],
      parentElement: null,
      isConnected: true,
      appendChild(child) { child.parentElement = node; node.children.push(child); return child },
      querySelectorAll(selector) {
        const matches = [];
        const visit = (current) => {
          for (const child of current.children) {
            if (child.kind === "text") continue;
            if (selector === "[data-pptx-object-key]" && child.dataset.pptxObjectKey !== undefined) matches.push(child);
            if (selector === "[data-pptx-paragraph-index]" && child.dataset.pptxParagraphIndex !== undefined) matches.push(child);
            visit(child);
          }
        };
        visit(node);
        return matches;
      },
    };
    return node;
  }

  function setup() {
    const objectKey = "ppt/slides/slide12.xml|slide|shape:3";
    const root = element("div");
    const object = element("div");
    object.dataset.pptxObjectKey = objectKey;
    root.appendChild(object);
    const paragraphs = [0, 1, 2].map((index) => {
      const paragraph = element("div");
      paragraph.dataset.pptxParagraphIndex = String(index);
      object.appendChild(paragraph);
      return paragraph;
    });
    globalThis.getComputedStyle = () => ({
      visibility: "visible",
      display: "block",
      opacity: "1",
      clipPath: "",
      filter: "",
      transform: "",
      translate: "",
      scale: "",
      rotate: "",
    });
    return { objectKey, root, object, paragraphs };
  }

  // 1. 目标索引同时提供对象、单段和段落范围。
  {
    const { objectKey, root } = setup();
    const keys = [...track.createPptxTrackTargets(root).elements.keys()];
    assert.deepEqual(keys.sort(), [
      objectKey,
      objectKey + "#paragraph:0-0",
      objectKey + "#paragraph:0-1",
      objectKey + "#paragraph:0-2",
      objectKey + "#paragraph:1-1",
      objectKey + "#paragraph:1-2",
      objectKey + "#paragraph:2-2",
    ].sort());
    const ranges = track.createPptxTrackTargets(root).elements;
    assert.equal(ranges.get(objectKey + "#paragraph:1-2").length, 2);
  }

  // 2. 静态状态同时包含对象、单段和段落范围（范围是求值目标）。
  {
    const { objectKey, root } = setup();
    const captured = track.capturePptxStaticState(root);
    assert.deepEqual(Object.keys(captured).sort(), [
      objectKey,
      objectKey + "#paragraph:0-0",
      objectKey + "#paragraph:0-1",
      objectKey + "#paragraph:0-2",
      objectKey + "#paragraph:1-1",
      objectKey + "#paragraph:1-2",
      objectKey + "#paragraph:2-2",
    ].sort());
    assert.equal(captured[objectKey + "#paragraph:0-0"].display, true);
  }

  // 3. 单段落状态只作用于该段落，不会波及同一对象的其他段落。
  {
    const { objectKey, root, paragraphs } = setup();
    track.applyPptxObjectState(root, {
      [objectKey + "#paragraph:0-0"]: { display: false, opacity: 0 },
    });
    assert.equal(paragraphs[0].style.visibility, "hidden");
    assert.equal(paragraphs[0].style.opacity, "0");
    assert.equal(paragraphs[1].style.visibility, "");
    assert.equal(paragraphs[2].style.visibility, "");
  }

  // 4. 回归场景：范围状态先应用，单段隐藏状态最后写入，时间零三段都应隐藏。
  {
    const { objectKey, root, paragraphs } = setup();
    const state = track.capturePptxStaticState(root);
    const prepared = {
      ...state,
      [objectKey + "#paragraph:0-0"]: { ...state[objectKey + "#paragraph:0-0"], display: false, opacity: 0 },
      [objectKey + "#paragraph:1-1"]: { ...state[objectKey + "#paragraph:1-1"], display: false, opacity: 0 },
      [objectKey + "#paragraph:2-2"]: { ...state[objectKey + "#paragraph:2-2"], display: false, opacity: 0 },
    };
    track.applyPptxObjectState(root, prepared);
    assert.deepEqual(
      paragraphs.map((paragraph) => [paragraph.style.visibility, paragraph.style.opacity]),
      [["hidden", "0"], ["hidden", "0"], ["hidden", "0"]],
    );
  }

  // 5. 对象级状态仍然作用于对象元素本身。
  {
    const { objectKey, root, object } = setup();
    track.applyPptxObjectState(root, { [objectKey]: { display: false, opacity: 0.25 } });
    assert.equal(object.style.visibility, "hidden");
    assert.equal(object.style.opacity, "0.25");
  }

  process.stdout.write(JSON.stringify({ status: "PASS" }));
`;

test("PPTX element state applies paragraph ranges before the more specific single paragraph", () => {
  const result = runNode([
    "--experimental-strip-types",
    "--import",
    registerHookUrl,
    "--input-type=module",
    "--eval",
    script,
  ]);
  assert.equal(result.status, 0, result.stderr || result.stdout);
  assert.equal(JSON.parse(result.stdout).status, "PASS");
});
