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

const identityUrl = sourceUrl("packages/pptx-core/src/playback/identity.ts");
const transitionUrl = sourceUrl("packages/pptx-core/src/browser/transition-player.ts");

const matchingScript = `
  import assert from "node:assert/strict";

  const identity = await import(${JSON.stringify(identityUrl)});

  const geometry = (x, y, width, height) => ({ x, y, width, height });
  const object = (key, values = {}) => ({
    key,
    slidePath: key.startsWith("from") ? "ppt/slides/slide4.xml" : "ppt/slides/slide5.xml",
    source: "slide",
    shapeId: key,
    groupPath: [],
    nodeType: "shape",
    ...values,
  });

  // 1. 几何从句柄属性构造。
  assert.deepEqual(
    identity.createPptxObjectGeometry({ x: "914400", y: "411480", cx: "7315200", cy: "640080" }),
    { x: 914400, y: 411480, width: 7315200, height: 640080 },
  );
  assert.deepEqual(
    identity.createPptxObjectGeometry({ x: 0, y: 0, cx: 10, cy: 20 }),
    { x: 0, y: 0, width: 10, height: 20 },
  );
  assert.equal(identity.createPptxObjectGeometry({ x: "0", y: "0" }), undefined);
  assert.equal(identity.createPptxObjectGeometry({ x: "0", y: "0", cx: "0", cy: "20" }), undefined);
  assert.equal(identity.createPptxObjectGeometry({ x: "nope", y: "0", cx: "10", cy: "20" }), undefined);

  // 2. 文字归一化：折叠空白、拒绝空值和超长文字。
  assert.equal(identity.normalizePptxObjectText("  Morph\\n  强身份 "), "Morph 强身份");
  assert.equal(identity.normalizePptxObjectText("   "), undefined);
  assert.equal(identity.normalizePptxObjectText(undefined), undefined);
  assert.equal(identity.normalizePptxObjectText("a".repeat(241)), undefined);
  assert.equal(identity.normalizePptxObjectText("a".repeat(240)).length, 240);

  // 3. 几何分数：相同几何高分，位移过大或尺寸差异过大时不可接受。
  const same = identity.pptxMorphGeometryScore(geometry(0, 0, 100, 100), geometry(0, 0, 100, 100));
  assert.ok(same > 0.9 && same <= 1);
  const moved = identity.pptxMorphGeometryScore(geometry(0, 0, 100, 100), geometry(20, 10, 100, 100));
  assert.ok(moved < same && moved > 0);
  assert.equal(identity.pptxMorphGeometryScore(geometry(0, 0, 100, 100), geometry(100000, 100000, 100, 100)), undefined);
  assert.equal(identity.pptxMorphGeometryScore(geometry(0, 0, 100, 100), geometry(0, 0, 1000, 1000)), undefined);
  assert.equal(identity.pptxMorphGeometryScore(undefined, geometry(0, 0, 100, 100)), undefined);
  assert.equal(
    identity.pptxMorphGeometryScore(geometry(0, 0, 100, 100), geometry(0, 0, 120, 120), { maxDistanceRatio: 6, minScaleRatio: 0.9 }),
    undefined,
  );

  // 4. 强匹配优先：显式名称和创建编号仍然最先匹配。
  const strong = identity.matchPptxMorphObjects(
    [
      object("from-hero", { explicitMorphName: "!!hero", geometry: geometry(0, 0, 100, 100) }),
      object("from-creation", { creationId: "ABC-1", geometry: geometry(0, 0, 100, 100) }),
    ],
    [
      object("to-hero", { explicitMorphName: "!!hero", geometry: geometry(500, 500, 400, 400) }),
      object("to-creation", { creationId: "abc-1", geometry: geometry(0, 0, 100, 100) }),
    ],
  );
  assert.deepEqual(strong.map((match) => [match.from, match.to, match.method, match.confidence]), [
    ["from-hero", "to-hero", "explicit-name", "strong"],
    ["from-creation", "to-creation", "creation-id", "strong"],
  ]);
  assert.equal(strong.every((match) => match.score === 1), true);
  assert.equal(strong.every((match) => match.unique === true), true);

  // 5. 中等匹配：相同普通名称 + 对象类型 + 接近几何。
  const medium = identity.matchPptxMorphObjects(
    [object("from-title", { name: "Title 1", geometry: geometry(914400, 411480, 7315200, 640080), text: "旧标题" })],
    [object("to-title", { name: "Title 1", geometry: geometry(1016000, 411480, 7000000, 640080), text: "新标题" })],
  );
  assert.equal(medium.length, 1);
  assert.equal(medium[0].method, "name-and-geometry");
  assert.equal(medium[0].confidence, "medium");
  assert.ok(medium[0].score > 0 && medium[0].score < 1);
  assert.equal(medium[0].unique, true);

  // 6. 没有几何、类型不同或名称不同时不产生中等匹配。
  assert.deepEqual(
    identity.matchPptxMorphObjects([object("from-a", { name: "Title 1" })], [object("to-a", { name: "Title 1" })]),
    [],
  );
  assert.deepEqual(
    identity.matchPptxMorphObjects(
      [object("from-a", { name: "Title 1", nodeType: "shape", geometry: geometry(0, 0, 100, 100) })],
      [object("to-a", { name: "Title 1", nodeType: "picture", geometry: geometry(0, 0, 100, 100) })],
    ),
    [],
  );
  assert.deepEqual(
    identity.matchPptxMorphObjects(
      [object("from-a", { name: "Title 1", geometry: geometry(0, 0, 100, 100) })],
      [object("to-a", { name: "Title 2", geometry: geometry(0, 0, 100, 100) })],
    ),
    [],
  );

  // 7. 中等匹配不是“看到同名就算”：几何差异过大时保留为未匹配。
  assert.deepEqual(
    identity.matchPptxMorphObjects(
      [object("from-a", { name: "Dot", geometry: geometry(0, 0, 10000, 10000) })],
      [object("to-a", { name: "Dot", geometry: geometry(5000000, 3000000, 900, 900) })],
    ),
    [],
  );

  // 8. 弱匹配：文字 + 类型 + 接近几何，只进入报告。
  const weak = identity.matchPptxMorphObjects(
    [object("from-text", { name: "old-name", text: "季度总结", geometry: geometry(0, 0, 100, 100) })],
    [object("to-text", { name: "new-name", text: "季度总结", geometry: geometry(10, 10, 110, 110) })],
  );
  assert.equal(weak.length, 1);
  assert.equal(weak[0].method, "text-and-geometry");
  assert.equal(weak[0].confidence, "weak");
  assert.ok(weak[0].score > 0);

  // 9. 唯一性：同名或同文字出现多次时不猜测对应关系。
  assert.deepEqual(
    identity.matchPptxMorphObjects(
      [
        object("from-a", { name: "Rectangle", geometry: geometry(0, 0, 100, 100) }),
        object("from-b", { name: "Rectangle", geometry: geometry(10, 10, 100, 100) }),
      ],
      [object("to-a", { name: "Rectangle", geometry: geometry(0, 0, 100, 100) })],
    ),
    [],
  );
  assert.deepEqual(
    identity.matchPptxMorphObjects(
      [object("from-a", { text: "重复文字", geometry: geometry(0, 0, 100, 100) })],
      [
        object("to-a", { text: "重复文字", geometry: geometry(0, 0, 100, 100) }),
        object("to-b", { text: "重复文字", geometry: geometry(10, 10, 100, 100) }),
      ],
    ),
    [],
  );

  // 10. 显式名称对象不会同时被当成按普通名称的中等匹配。
  const mixed = identity.matchPptxMorphObjects(
    [object("from-a", { name: "!!hero", explicitMorphName: "!!hero", geometry: geometry(0, 0, 100, 100) })],
    [object("to-a", { name: "!!hero", explicitMorphName: "!!other", geometry: geometry(0, 0, 100, 100) })],
  );
  assert.deepEqual(mixed, []);

  // 11. 一个对象只匹配一次：强匹配占用的对象不再进入弱匹配。
  const exclusive = identity.matchPptxMorphObjects(
    [
      object("from-a", { explicitMorphName: "!!one", text: "共享文字", geometry: geometry(0, 0, 100, 100) }),
      object("from-b", { name: "other", text: "共享文字", geometry: geometry(0, 0, 100, 100) }),
    ],
    [
      object("to-a", { explicitMorphName: "!!one", text: "共享文字", geometry: geometry(0, 0, 100, 100) }),
      object("to-b", { name: "other2", text: "共享文字", geometry: geometry(0, 0, 100, 100) }),
    ],
  );
  assert.deepEqual(exclusive.map((match) => [match.from, match.to, match.confidence]), [
    ["from-a", "to-a", "strong"],
    ["from-b", "to-b", "weak"],
  ]);

  // 12. 自定义阈值可以收紧或放宽判定。
  const from = [object("from-a", { name: "Box", geometry: geometry(0, 0, 100, 100) })];
  const to = [object("to-a", { name: "Box", geometry: geometry(300, 0, 100, 100) })];
  assert.equal(identity.matchPptxMorphObjects(from, to, { maxDistanceRatio: 6, minScaleRatio: 0.2 }).length, 1);
  assert.equal(identity.matchPptxMorphObjects(from, to, { maxDistanceRatio: 1, minScaleRatio: 0.2 }).length, 0);

  process.stdout.write(JSON.stringify({ status: "PASS" }));
`;

test("PPTX morph matching grades strong, medium and weak identity with geometry", () => {
  const result = runNode([
    "--experimental-strip-types",
    "--import",
    registerHookUrl,
    "--input-type=module",
    "--eval",
    matchingScript,
  ]);
  assert.equal(result.status, 0, result.stderr || result.stdout);
  assert.equal(JSON.parse(result.stdout).status, "PASS");
});

const transitionScript = `
  import assert from "node:assert/strict";

  const { mountPptxTransition } = await import(${JSON.stringify(transitionUrl)});
  const { isPptxMorphMatchAnimatable } = await import(${JSON.stringify(transitionUrl)});

  // 最小 DOM：只需要 transition-player 使用的节点、样式、文字树和矩形。
  function createDom() {
    const rects = { from: {}, to: {} };
    const style = () => ({
      opacity: "", translate: "", scale: "", clipPath: "", willChange: "",
      position: "", zIndex: "", display: "", inset: "", pointerEvents: "",
    });
    const doc = { rects, created: [] };
    const sideOf = (node) => node.side;
    const textOf = (node) => node.children.map((child) => (child.kind === "text" ? child.data : textOf(child))).join("");
    const applySide = (target, side) => {
      target.ownerDocument = doc;
      if (target.side === undefined && side !== undefined) target.side = side;
      target.children?.forEach((child) => applySide(child, side));
    };

    function insert(node, child, anchor) {
      child.parentNode?.children.splice(child.parentNode.children.indexOf(child), 1);
      child.parentNode = node;
      const index = anchor ? node.children.indexOf(anchor) : -1;
      if (index < 0) node.children.push(child);
      else node.children.splice(index, 0, child);
      applySide(child, node.side);
    }

    function element(kind) {
      const node = {
        kind,
        parentNode: null,
        ownerDocument: doc,
        side: undefined,
        style: style(),
        dataset: {},
        children: [],
        rect: undefined,
        get firstElementChild() {
          return node.children.find((child) => child.kind !== "text") ?? null;
        },
        get textContent() { return textOf(node); },
        set textContent(value) {
          node.children = [];
          if (value !== "") insert(node, doc.createTextNode(value));
        },
        appendChild(child) { insert(node, child); return child; },
        insertBefore(child, anchor) { insert(node, child, anchor); return child; },
        removeChild(child) {
          const index = node.children.indexOf(child);
          if (index >= 0) node.children.splice(index, 1);
          child.parentNode = null;
          return child;
        },
        remove() { node.parentNode?.removeChild(node); },
        replaceWith(replacement) {
          const parent = node.parentNode;
          if (!parent) return;
          const index = parent.children.indexOf(node);
          const inserted = replacement.kind === "fragment" ? replacement.children : [replacement];
          parent.children.splice(index, 1, ...inserted);
          for (const child of inserted) {
            child.parentNode = parent;
            applySide(child, parent.side);
          }
          if (replacement.kind === "fragment") replacement.children = [];
          node.parentNode = null;
        },
        querySelectorAll(selector) {
          const matches = [];
          const visit = (current) => {
            for (const child of current.children ?? []) {
              if (child.kind === "text") continue;
              if (selector === "[data-pptx-object-key]" && child.dataset.pptxObjectKey) matches.push(child);
              visit(child);
            }
          };
          visit(node);
          return matches;
        },
        getBoundingClientRect() {
          const values = node.rect ?? doc.rects[node.side]?.[textOf(node)] ?? { left: 0, top: 0, width: 0, height: 0 };
          return { ...values, right: values.left + values.width, bottom: values.top + values.height };
        },
      };
      doc.created.push(node);
      return node;
    }

    doc.createElement = (tag) => element(tag);
    doc.createTextNode = (data) => {
      const node = element("text");
      node.data = data;
      return node;
    };
    doc.createDocumentFragment = () => element("fragment");
    doc.createTreeWalker = (root) => {
      const nodes = [];
      const visit = (node) => {
        for (const child of node.children ?? []) {
          if (child.kind === "text") nodes.push(child);
          else visit(child);
        }
      };
      visit(root);
      let index = -1;
      return {
        get currentNode() { return nodes[index]; },
        nextNode() { index += 1; return index < nodes.length ? nodes[index] : null; },
      };
    };
    doc.element = element;
    return doc;
  }

  function markedObject(doc, key, side, text, rect) {
    const object = doc.createElement("div");
    object.dataset.pptxObjectKey = key;
    object.side = side;
    object.rect = rect;
    if (text) object.textContent = text;
    for (const child of object.children) child.side = side;
    return object;
  }

  const key = "ppt/slides/slide5.xml|slide|shape:3";

  function setup({ option, confidence = "medium", method = "name-and-geometry" }) {
    const doc = createDom();
    globalThis.document = doc;
    globalThis.getComputedStyle = () => ({ position: "static" });

    const incoming = doc.createElement("section");
    incoming.side = "to";
    const incomingObject = markedObject(doc, key, "to", "alpha beta", { left: 100, top: 100, width: 400, height: 80 });
    incoming.appendChild(incomingObject);

    const outgoing = doc.createElement("section");
    outgoing.side = "from";
    const outgoingObject = markedObject(doc, key, "from", "alpha beta", { left: 100, top: 100, width: 400, height: 80 });
    outgoing.appendChild(outgoingObject);

    const root = doc.createElement("div");
    root.side = "from";
    root.appendChild(incoming);

    doc.rects.to.alpha = { left: 300, top: 40, width: 60, height: 30 };
    doc.rects.to.beta = { left: 380, top: 40, width: 50, height: 30 };
    doc.rects.from.alpha = { left: 120, top: 10, width: 60, height: 30 };
    doc.rects.from.beta = { left: 260, top: 10, width: 50, height: 30 };

    const match = { from: key, to: key, method, confidence, score: 0.5, unique: true };
    const capture = { layer: outgoing, objects: new Map([[key, { left: 120, top: 20, width: 380, height: 90 }]]) };
    const transition = { kind: "morph", option, durationMs: 1000, advanceOnClick: true };
    return { doc, root, incomingObject, outgoing, match, capture, transition };
  }

  function wordSpans(doc) {
    return doc.created.filter((node) => node.kind === "span" && "pptxMorphWord" in node.dataset);
  }

  function incomingWordSpans(doc) {
    return wordSpans(doc).filter((node) => node.side === "to");
  }

  // 逐词 Morph：byWord 把每个词从旧页同名词的位置移回新位置，结束后恢复 DOM。
  {
    const { doc, root, incomingObject, match, capture, transition } = setup({ option: "byWord" });
    const handle = mountPptxTransition(root, capture, transition, [match]);
    assert.ok(handle);
    assert.equal(handle.durationMs, 1000);
    assert.equal(incomingObject.textContent, "alpha beta");

    const spans = incomingWordSpans(doc);
    assert.equal(spans.length, 2);
    handle.apply(0);
    assert.equal(spans[0].style.translate, "-180px -30px");
    assert.equal(spans[1].style.translate, "-120px -30px");
    assert.equal(spans[0].style.display, "inline-block");

    handle.apply(500);
    assert.equal(spans[0].style.translate, "-90px -15px");
    handle.apply(1000);
    assert.equal(spans[0].style.translate, "0px 0px");
    assert.equal(spans[1].style.translate, "0px 0px");

    handle.finish();
    assert.equal(incomingWordSpans(doc).filter((span) => span.parentNode).length, 0);
    assert.equal(incomingObject.textContent, "alpha beta");
    assert.equal(incomingObject.children.every((child) => child.kind === "text"), true);
  }

  // byObject（或未声明 option）仍然按对象整体补间。
  {
    const { doc, root, incomingObject, match, capture, transition } = setup({ option: "byObject", confidence: "strong", method: "explicit-name" });
    const handle = mountPptxTransition(root, capture, transition, [match]);
    assert.ok(handle);
    assert.equal(incomingWordSpans(doc).length, 0);
    handle.apply(0);
    assert.equal(incomingObject.style.translate, "20px -80px");
    assert.equal(incomingObject.style.scale, "0.95 1.125");
    handle.apply(1000);
    assert.equal(incomingObject.style.translate, "0px 0px");
    handle.finish();
  }

  // 弱匹配不执行任何补间。
  {
    const { doc, root, incomingObject, match, capture, transition } = setup({ option: "byWord", confidence: "weak", method: "text-and-geometry" });
    assert.equal(isPptxMorphMatchAnimatable(match), false);
    const handle = mountPptxTransition(root, capture, transition, [match]);
    assert.ok(handle);
    assert.equal(incomingWordSpans(doc).length, 0);
    handle.apply(0);
    assert.equal(incomingObject.style.translate, "");
    assert.equal(incomingObject.style.scale, "");
    handle.finish();
  }

  // 中等匹配在对象粒度上同样执行补间。
  {
    const { root, incomingObject, match, capture, transition } = setup({ option: "byObject" });
    const handle = mountPptxTransition(root, capture, transition, [match]);
    assert.ok(handle);
    handle.apply(0);
    assert.equal(incomingObject.style.translate, "20px -80px");
    handle.finish();
    assert.equal(incomingObject.style.translate, "");
  }

  process.stdout.write(JSON.stringify({ status: "PASS" }));
`;

test("PPTX morph transition applies byWord and medium matches, skips weak ones", () => {
  const result = runNode([
    "--experimental-strip-types",
    "--import",
    registerHookUrl,
    "--input-type=module",
    "--eval",
    transitionScript,
  ]);
  assert.equal(result.status, 0, result.stderr || result.stdout);
  assert.equal(JSON.parse(result.stdout).status, "PASS");
});
