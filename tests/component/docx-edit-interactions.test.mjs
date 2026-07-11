import assert from "node:assert/strict";
import test from "node:test";

import {
  importFromDemo,
  mount,
  vue,
} from "./vue-test-renderer.mjs";

const { DocxEditor, useDocxEditor } = await importFromDemo("@arcships/vue-docx");
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
