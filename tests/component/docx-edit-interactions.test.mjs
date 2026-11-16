import assert from "node:assert/strict";
import test from "node:test";

import {
  importFromDemo,
  mount,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

const { DocxEditor, DocxImageLayer, useDocxEditor } = await importFromDemo("@arcships/vue-docx");
const { createBlankDocumentModel } = await importFromDemo("@arcships/docx-core");

function paragraphText(node) {
  if (!node || node.type !== "paragraph") return "";
  return node.children.map((child) => child.type === "text" ? child.text : "").join("");
}

function imageAt(model, nodeIndex, childIndex) {
  const node = model.nodes[nodeIndex];
  const child = node?.type === "paragraph" ? node.children[childIndex] : undefined;
  assert.equal(child?.type, "image");
  return child;
}

function deepFreeze(value, seen = new Set()) {
  if (!value || typeof value !== "object" || seen.has(value)) return value;
  seen.add(value);
  for (const child of Object.values(value)) deepFreeze(child, seen);
  return Object.freeze(value);
}

function starterModel(nodes) {
  return {
    ...createBlankDocumentModel(),
    nodes,
  };
}

async function mountController(model) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model });
      return () => vue.h("div", { "data-testid": "controller-harness" }, controller.status);
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, get controller() { return controller; } };
}

function createListenerTarget() {
  const listeners = new Map();
  return {
    addEventListener(type, listener) {
      const current = listeners.get(type) ?? new Set();
      current.add(listener);
      listeners.set(type, current);
    },
    removeEventListener(type, listener) {
      listeners.get(type)?.delete(listener);
    },
    count(type) {
      return listeners.get(type)?.size ?? 0;
    },
    total() {
      return [...listeners.values()].reduce((sum, entries) => sum + entries.size, 0);
    },
  };
}

function installDomEnvironment(t) {
  const originalWindow = globalThis.window;
  const originalDocument = globalThis.document;
  const originalResizeObserver = globalThis.ResizeObserver;
  const originalRect = Object.getOwnPropertyDescriptor(Object.prototype, "getBoundingClientRect");
  const originalClientHeight = Object.getOwnPropertyDescriptor(Object.prototype, "clientHeight");
  const windowListeners = createListenerTarget();
  const documentListeners = createListenerTarget();

  class FakeText {
    constructor(text) { this.textContent = String(text); }
    get outerHTML() { return this.textContent; }
  }

  class FakeElement {
    constructor(tagName) {
      this.tagName = String(tagName).toLowerCase();
      this.attributes = new Map();
      this.children = [];
      this.style = { cssText: "", setProperty(name, value) { this[name] = value; } };
    }
    appendChild(child) { this.children.push(child); return child; }
    setAttribute(name, value) { this.attributes.set(String(name), String(value)); }
    get innerHTML() { return this.children.map((child) => child.outerHTML ?? "").join(""); }
    get outerHTML() { return `<${this.tagName}>${this.innerHTML}</${this.tagName}>`; }
    getContext() { return { measureText: (text) => ({ width: String(text).length * 8 }) }; }
    querySelector() { return null; }
  }

  globalThis.window = {
    addEventListener: windowListeners.addEventListener,
    removeEventListener: windowListeners.removeEventListener,
    getSelection() { return null; },
  };
  globalThis.document = {
    addEventListener: documentListeners.addEventListener,
    removeEventListener: documentListeners.removeEventListener,
    body: { append() {} },
    createElement(tagName) { return new FakeElement(tagName); },
    createTextNode(text) { return new FakeText(text); },
    getSelection() { return null; },
    execCommand() { return false; },
  };
  globalThis.ResizeObserver = class ResizeObserver {
    observe() {}
    disconnect() {}
  };
  Object.defineProperty(Object.prototype, "getBoundingClientRect", {
    configurable: true,
    value: () => ({ height: 1056, width: 816, top: 0, left: 0, right: 816, bottom: 1056 }),
  });
  Object.defineProperty(Object.prototype, "clientHeight", {
    configurable: true,
    get: () => 800,
  });

  t.after(() => {
    if (originalWindow === undefined) delete globalThis.window;
    else globalThis.window = originalWindow;
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
    if (originalResizeObserver === undefined) delete globalThis.ResizeObserver;
    else globalThis.ResizeObserver = originalResizeObserver;
    if (originalRect) Object.defineProperty(Object.prototype, "getBoundingClientRect", originalRect);
    else delete Object.prototype.getBoundingClientRect;
    if (originalClientHeight) Object.defineProperty(Object.prototype, "clientHeight", originalClientHeight);
    else delete Object.prototype.clientHeight;
  });

  return { windowListeners, documentListeners };
}

test("DocxEditor dynamically owns no editing listeners while read-only", async (t) => {
  const { windowListeners, documentListeners } = installDomEnvironment(t);
  const editable = vue.ref(false);
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({
        starterModel: starterModel([
          { type: "paragraph", children: [{ type: "text", text: "listener boundary" }] },
        ]),
      });
      return () => vue.h(DocxEditor, {
        editor: controller,
        editable: editable.value,
        showToolbar: false,
        showThumbnails: false,
      });
    },
  });
  const mounted = await mount(Harness);

  assert.equal(windowListeners.total(), 0, "read-only mount must not own a window editing listener");
  assert.equal(documentListeners.total(), 0, "read-only mount must not own a document editing listener");

  editable.value = true;
  await vue.nextTick();
  assert.equal(windowListeners.count("keydown"), 1, "editable mode must own exactly one shortcut listener");
  assert.equal(documentListeners.total(), 0, "idle editable mode must not leak document drag listeners");

  editable.value = false;
  await vue.nextTick();
  assert.equal(windowListeners.total(), 0, "switching to read-only must remove editing listeners");
  assert.equal(documentListeners.total(), 0);

  editable.value = true;
  await vue.nextTick();
  assert.equal(windowListeners.count("keydown"), 1, "re-enabling must not duplicate listeners");
  mounted.app.unmount();
  assert.equal(windowListeners.total(), 0, "unmount must release the listener");
  assert.equal(documentListeners.total(), 0);
  assert.deepEqual(mounted.warnings, []);
});

test("text and format commands keep inputs immutable and undo/redo restore selection", async () => {
  const source = deepFreeze(starterModel([
    { type: "paragraph", children: [{ type: "text", text: "abcdef" }] },
  ]));
  const mounted = await mountController(source);
  const { controller } = mounted;
  const expandedRange = {
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 2 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 4 },
  };

  controller.selectParagraph(0);
  await vue.nextTick();
  controller.setActiveTextRange(expandedRange);
  await vue.nextTick();

  const beforeReplace = controller.model;
  const beforeReplaceJson = JSON.stringify(beforeReplace);
  const collapsed = controller.replaceExpandedSelection("中A", expandedRange);
  await vue.nextTick();

  assert.notStrictEqual(controller.model, beforeReplace);
  assert.equal(JSON.stringify(beforeReplace), beforeReplaceJson, "replace must not mutate its input model");
  assert.equal(paragraphText(controller.model.nodes[0]), "ab中Aef");
  assert.deepEqual(collapsed, {
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 4 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 4 },
  });
  assert.equal(controller.canUndo, true);

  controller.undo();
  await vue.nextTick();
  assert.equal(paragraphText(controller.model.nodes[0]), "abcdef");
  assert.deepEqual(controller.selection, { kind: "paragraph", nodeIndex: 0 });
  assert.deepEqual(controller.activeTextRange, expandedRange);
  assert.deepEqual(controller.historyRestoreRequest?.activeTextRange, expandedRange);
  assert.equal(controller.canRedo, true);

  controller.redo();
  await vue.nextTick();
  assert.equal(paragraphText(controller.model.nodes[0]), "ab中Aef");
  assert.deepEqual(controller.activeTextRange, collapsed);
  assert.deepEqual(controller.historyRestoreRequest?.activeTextRange, collapsed);

  const formatRange = {
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 2 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 4 },
  };
  controller.setActiveTextRange(formatRange);
  await vue.nextTick();
  const beforeFormat = controller.model;
  const beforeFormatJson = JSON.stringify(beforeFormat);
  controller.toggleBold();
  await vue.nextTick();
  assert.notStrictEqual(controller.model, beforeFormat);
  assert.equal(JSON.stringify(beforeFormat), beforeFormatJson, "format command must not mutate its input model");
  const boldText = controller.model.nodes[0].children
    .filter((child) => child.type === "text" && child.style?.bold)
    .map((child) => child.text)
    .join("");
  assert.equal(boldText, "中A");

  controller.undo();
  await vue.nextTick();
  assert.deepEqual(controller.activeTextRange, formatRange);
  assert.deepEqual(controller.historyRestoreRequest?.activeTextRange, formatRange);
  controller.redo();
  await vue.nextTick();
  assert.equal(controller.model.nodes[0].children.some((child) => child.type === "text" && child.style?.bold), true);

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("character formatting applies to every paragraph in a cross-paragraph selection", async () => {
  const source = deepFreeze(starterModel([
    { type: "paragraph", sourceXml: "<w:p id=\"p0\"/>", children: [{ type: "text", text: "alpha" }] },
    { type: "paragraph", sourceXml: "<w:p id=\"p1\"/>", children: [{ type: "text", text: "beta" }] },
  ]));
  const mounted = await mountController(source);
  const { controller } = mounted;

  controller.selectParagraph(0);
  await vue.nextTick();
  controller.setActiveTextRange({
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 2 },
    end: { location: { kind: "paragraph", nodeIndex: 1 }, offset: 2 },
  });
  await vue.nextTick();

  controller.toggleBold();
  await vue.nextTick();

  const boldText = (node) => node.children
    .filter((child) => child.type === "text" && child.style?.bold)
    .map((child) => child.text)
    .join("");
  assert.equal(boldText(controller.model.nodes[0]), "pha", "tail of the first paragraph must be bold");
  assert.equal(boldText(controller.model.nodes[1]), "be", "head of the last paragraph must be bold");
  assert.equal(paragraphText(controller.model.nodes[0]), "alpha");
  assert.equal(paragraphText(controller.model.nodes[1]), "beta");
  assert.equal(controller.model.nodes[0].sourceXml, undefined);
  assert.equal(controller.model.nodes[1].sourceXml, undefined);
  assert.equal(controller.canUndo, true);

  controller.undo();
  await vue.nextTick();
  assert.equal(boldText(controller.model.nodes[0]), "");
  assert.equal(boldText(controller.model.nodes[1]), "");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("collapsed caret formatting stays pending, adds no undo entry, and styles the next typed run", async () => {
  const source = deepFreeze(starterModel([
    {
      type: "paragraph",
      sourceXml: "<w:p id=\"pending\"/>",
      children: [{ type: "text", text: "abcdef", style: { fontSizePt: 11 } }],
    },
  ]));
  const mounted = await mountController(source);
  const { controller } = mounted;

  controller.selectParagraph(0);
  await vue.nextTick();
  controller.setActiveTextRange({
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 2 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 2 },
  });
  await vue.nextTick();

  const beforeModel = controller.model;
  controller.toggleBold();
  await vue.nextTick();
  assert.strictEqual(controller.model, beforeModel, "a caret style change must not clone the model");
  assert.equal(controller.canUndo, false, "a caret style change must not add an empty undo entry");
  assert.equal(controller.pendingRunStyle?.bold, true);

  controller.commitParagraphText(0, "abXcdef");
  await vue.nextTick();
  assert.equal(paragraphText(controller.model.nodes[0]), "abXcdef");
  const runs = controller.model.nodes[0].children.filter((child) => child.type === "text");
  const inserted = runs.find((run) => run.text === "X");
  assert.ok(inserted, "typed text should produce a run");
  assert.equal(inserted.style?.bold, true, "the pending bold style must style the typed run");
  assert.equal(inserted.style?.fontSizePt, 11, "the typed run must inherit the surrounding run style");
  assert.equal(runs.filter((run) => run.text === "ab").every((run) => !run.style?.bold), true);
  assert.equal(runs.filter((run) => run.text === "cdef").every((run) => !run.style?.bold), true);

  controller.selectParagraph(0);
  await vue.nextTick();
  assert.equal(controller.pendingRunStyle, undefined, "moving the caret must drop the pending style");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("floating image drag and resize convert screen pixels to page coordinates at 50% zoom", async (t) => {
  const originalDocument = globalThis.document;
  const listeners = new Map();
  globalThis.document = {
    addEventListener(type, listener) {
      const entries = listeners.get(type) ?? new Set();
      entries.add(listener);
      listeners.set(type, entries);
    },
    removeEventListener(type, listener) {
      listeners.get(type)?.delete(listener);
    },
  };
  t.after(() => {
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
  });
  const dispatchDocument = (type, event) => {
    for (const listener of listeners.get(type) ?? []) listener(event);
  };

  const source = deepFreeze(starterModel([
    {
      type: "paragraph",
      children: [
        { type: "text", text: "anchor" },
        {
          type: "image",
          src: "data:image/svg+xml;base64,PHN2Zy8+",
          alt: "drag-image",
          widthPx: 100,
          heightPx: 60,
          floating: { xPx: 20, yPx: 30, wrapType: "square", behindDocument: false, zIndex: 1 },
        },
      ],
    },
  ]));
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: source });
      return () => vue.h(DocxImageLayer, {
        pageIndex: 0,
        pageWidthPx: 816,
        pageHeightPx: 1056,
        controller,
        pageNodeSegments: [{ nodeIndex: 0 }],
        editable: true,
        zoomFactor: 0.5,
      });
    },
  });
  const mounted = await mount(Harness);
  const floatingImage = walk(mounted.root).find(
    (node) => node.props?.["data-testid"] === "docx-floating-image"
  );
  assert.ok(floatingImage, "floating image should render");

  floatingImage.props.onMousedown({ preventDefault() {}, clientX: 100, clientY: 100 });
  dispatchDocument("mousemove", { type: "mousemove", clientX: 150, clientY: 120 });
  dispatchDocument("mouseup", { type: "mouseup", clientX: 150, clientY: 120 });
  await vue.nextTick();
  const dragged = controller.model.nodes[0].children[1];
  assert.deepEqual(
    { xPx: dragged.floating.xPx, yPx: dragged.floating.yPx },
    { xPx: 120, yPx: 70 },
    "a 50px pointer delta at 50% zoom must move the image 100 page pixels",
  );

  const resizeHandle = walk(mounted.root).find(
    (node) => String(node.props?.class ?? "").includes("docx-image-resize-handle--se")
  );
  assert.ok(resizeHandle, "selected image should render the south-east resize handle");
  resizeHandle.props.onMousedown({
    preventDefault() {},
    stopPropagation() {},
    clientX: 100,
    clientY: 100,
  });
  dispatchDocument("mousemove", { type: "mousemove", clientX: 150, clientY: 125 });
  dispatchDocument("mouseup", { type: "mouseup", clientX: 150, clientY: 125 });
  await vue.nextTick();
  const resized = controller.model.nodes[0].children[1];
  assert.equal(resized.widthPx, 200, "50px horizontal pointer delta must add 100 page pixels");
  assert.equal(resized.heightPx, 110, "25px vertical pointer delta must add 50 page pixels");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("image and table public operations commit real model transactions and history", async () => {
  const source = deepFreeze(starterModel([
    {
      type: "paragraph",
      children: [
        { type: "text", text: "source" },
        {
          type: "image",
          src: "data:image/svg+xml;base64,PHN2Zy8+",
          alt: "transaction-image",
          widthPx: 100,
          heightPx: 60,
        },
      ],
    },
    { type: "paragraph", children: [{ type: "text", text: "target" }] },
  ]));
  const mounted = await mountController(source);
  const { controller } = mounted;
  const imageLocation = { kind: "paragraph", nodeIndex: 0, childIndex: 1 };

  const beforeWrap = controller.model;
  const beforeWrapJson = JSON.stringify(beforeWrap);
  controller.setImageWrapMode(imageLocation, "square", { xPx: 12, yPx: 18 });
  await vue.nextTick();
  assert.notStrictEqual(controller.model, beforeWrap);
  assert.equal(JSON.stringify(beforeWrap), beforeWrapJson);
  assert.equal(imageAt(controller.model, 0, 1).floating?.wrapType, "square");

  controller.moveFloatingImage(imageLocation, { xPx: 44, yPx: 55 });
  await vue.nextTick();
  assert.equal(imageAt(controller.model, 0, 1).floating?.xPx, 44);
  assert.equal(imageAt(controller.model, 0, 1).floating?.yPx, 55);

  controller.resizeImage(imageLocation, 180, 90);
  await vue.nextTick();
  assert.equal(imageAt(controller.model, 0, 1).widthPx, 180);
  assert.equal(imageAt(controller.model, 0, 1).heightPx, 90);
  assert.equal(controller.canUndo, true);
  controller.undo();
  await vue.nextTick();
  assert.equal(imageAt(controller.model, 0, 1).widthPx, 100, "undo restores pre-resize dimensions");
  controller.redo();
  await vue.nextTick();
  assert.equal(imageAt(controller.model, 0, 1).widthPx, 180);

  controller.moveImage(imageLocation, { kind: "paragraph", nodeIndex: 1, childIndex: 1 });
  await vue.nextTick();
  assert.equal(controller.model.nodes[0].children.some((child) => child.type === "image"), false);
  assert.equal(controller.model.nodes[1].children[1]?.type, "image");
  controller.undo();
  await vue.nextTick();
  assert.equal(controller.model.nodes[0].children[1]?.type, "image");
  controller.redo();
  await vue.nextTick();
  assert.equal(controller.model.nodes[1].children[1]?.type, "image");

  const beforeTable = controller.model;
  const beforeTableJson = JSON.stringify(beforeTable);
  controller.insertTable();
  await vue.nextTick();
  const tableIndex = controller.model.nodes.findIndex((node) => node.type === "table");
  assert.ok(tableIndex >= 0);
  assert.equal(JSON.stringify(beforeTable), beforeTableJson);
  controller.insertTableRow(tableIndex, 0, "below");
  await vue.nextTick();
  controller.insertTableColumn(tableIndex, 0, "right");
  await vue.nextTick();
  controller.setTableColumnWidths(tableIndex, [90, 110, 130, 150]);
  await vue.nextTick();
  const table = controller.model.nodes[tableIndex];
  assert.equal(table.type, "table");
  assert.equal(table.rows.length, 4);
  assert.equal(table.rows[0].cells.length, 4);
  assert.deepEqual(table.style?.columnWidthsTwips, [1350, 1650, 1950, 2250]);
  assert.equal(controller.canUndo, true);
  controller.undo();
  await vue.nextTick();
  assert.notDeepEqual(controller.model.nodes[tableIndex].style?.columnWidthsTwips, [1350, 1650, 1950, 2250]);
  controller.redo();
  await vue.nextTick();
  assert.deepEqual(controller.model.nodes[tableIndex].style?.columnWidthsTwips, [1350, 1650, 1950, 2250]);

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});
