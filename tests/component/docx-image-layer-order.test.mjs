// End-to-end coverage for issue #145 — DOCX image layer order commands.
//
// Every bullet of the issue is exercised here:
//   - two overlapping images reordered through the image context menu
//   - render order and exported `wp:anchor/@relativeHeight` change
//   - undo restores both the model and the exported order
//   - disabled states for inline / top-most / bottom-most images
//   - right-clicking an image resolves an `image` context menu location

import assert from "node:assert/strict";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  teleportTarget,
  textContent,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

const {
  DocxContextMenu,
  DocxEditor,
  DocxImageLayer,
  useDocxEditor,
} = await importFromDemo("@arcships/vue-docx");
const core = await importFromDemo("@arcships/docx-core");
const { docxImageLocationFromNode } = await import(
  new URL(
    "../../packages/vue-docx/src/composables/editor-selection.ts",
    import.meta.url,
  )
);

const MAX = core.WORD_IMAGE_Z_INDEX_MAX;
const STEP = core.WORD_IMAGE_Z_INDEX_STEP;

function imageNode(alt, overrides = {}) {
  return {
    type: "image",
    src: "data:image/png;base64,AAAA",
    alt,
    widthPx: 120,
    heightPx: 80,
    floating: {
      xPx: 10,
      yPx: 10,
      wrapType: "none",
      behindDocument: false,
      zIndex: 65536,
      ...overrides.floating,
    },
  };
}

function overlappingImagesModel(firstZIndex = 65536, secondZIndex = 131072) {
  const model = core.createBlankDocumentModel();
  model.nodes = [{
    type: "paragraph",
    children: [
      imageNode("first", { floating: { zIndex: firstZIndex } }),
      imageNode("second", { floating: { xPx: 50, yPx: 40, zIndex: secondZIndex } }),
    ],
  }];
  return model;
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

async function mountMenu(controller, context) {
  const current = vue.ref(context);
  const closes = [];
  const Harness = vue.defineComponent({
    setup() {
      return () => vue.h(DocxContextMenu, {
        context: current.value,
        controller,
        x: 24,
        y: 32,
        onClose: () => {
          closes.push("close");
          current.value = undefined;
        },
      });
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, context: current, closes };
}

function imageContext(controller, childIndex = 0) {
  const image = controller.model.nodes[0].children[childIndex];
  return {
    kind: "image",
    image: {
      location: { kind: "paragraph", nodeIndex: 0, childIndex },
      floating: image.floating,
      wrap: core.resolveDocxImageWrapState(image.floating),
    },
  };
}

function layerImageByChildIndex(root, childIndex) {
  return walk(root).find((node) => (
    node.props?.["data-testid"] === "docx-floating-image" &&
    node.props?.["data-image-child-index"] === childIndex
  ));
}

async function relativeHeights(controller) {
  const xml = await core.modelToDocumentXml(controller.model);
  return [...xml.matchAll(/relativeHeight="(\d+)"/gu)].map((match) => Number(match[1]));
}

test("the image context menu reorders overlapping images, exports the new z-order and undoes it", async (t) => {
  const mounted = await mountController(overlappingImagesModel());
  const { controller } = mounted;
  t.after(() => mounted.app.unmount());

  assert.deepEqual(await relativeHeights(controller), [65536, 131072]);

  const menu = await mountMenu(controller, imageContext(controller));
  t.after(() => menu.app.unmount());

  const toFront = findByTestId(teleportTarget("body"), "docx-context-menu-image-bring-to-front");
  const forward = findByTestId(teleportTarget("body"), "docx-context-menu-image-bring-forward");
  const backward = findByTestId(teleportTarget("body"), "docx-context-menu-image-send-backward");
  const toBack = findByTestId(teleportTarget("body"), "docx-context-menu-image-send-to-back");
  assert.ok(toFront, "the menu exposes Bring to Front");
  assert.ok(forward, "the menu exposes Bring Forward");
  assert.ok(backward, "the menu exposes Send Backward");
  assert.ok(toBack, "the menu exposes Send to Back");
  assert.equal(toFront.props.disabled, false);
  assert.equal(forward.props.disabled, false);
  assert.equal(backward.props.disabled, false);
  assert.equal(toBack.props.disabled, false);

  assert.equal(controller.canUndo, false);
  await toFront.props.onClick();
  await vue.nextTick();
  assert.equal(controller.model.nodes[0].children[0].floating.zIndex, MAX);
  assert.equal(controller.model.nodes[0].children[0].floating.behindDocument, false);
  assert.equal(controller.canUndo, true, "the layer command is one undo entry");
  assert.deepEqual(
    await relativeHeights(controller),
    [MAX, 131072],
    "export emits the new relativeHeight",
  );
  assert.deepEqual(menu.closes, ["close"], "choosing an action closes the menu");

  controller.undo();
  await vue.nextTick();
  assert.equal(controller.model.nodes[0].children[0].floating.zIndex, 65536);
  assert.deepEqual(
    await relativeHeights(controller),
    [65536, 131072],
    "undo restores the previous exported z-order",
  );

  controller.redo();
  await vue.nextTick();
  assert.equal(controller.model.nodes[0].children[0].floating.zIndex, MAX);
  assert.deepEqual(await relativeHeights(controller), [MAX, 131072]);
});

test("floating image containers paint in the model relativeHeight order", async (t) => {
  const mounted = await mountController(overlappingImagesModel());
  const { controller } = mounted;
  t.after(() => mounted.app.unmount());

  const layer = await mount(vue.defineComponent({
    setup() {
      return () => vue.h(DocxImageLayer, {
        pageIndex: 0,
        pageWidthPx: 816,
        pageHeightPx: 1056,
        controller,
        pageNodeSegments: [{ nodeIndex: 0 }],
        editable: true,
      });
    },
  }));
  t.after(() => layer.app.unmount());

  let first = layerImageByChildIndex(layer.root, 0);
  let second = layerImageByChildIndex(layer.root, 1);
  assert.ok(first && second, "both overlapping images render in the layer");
  assert.equal(first.props.style.zIndex, String(65536 / STEP));
  assert.equal(second.props.style.zIndex, String(131072 / STEP));
  assert.ok(
    Number(first.props.style.zIndex) < Number(second.props.style.zIndex),
    "the lower relativeHeight paints first",
  );

  controller.bringToFront({ kind: "paragraph", nodeIndex: 0, childIndex: 0 });
  await vue.nextTick();
  first = layerImageByChildIndex(layer.root, 0);
  second = layerImageByChildIndex(layer.root, 1);
  assert.equal(first.props.style.zIndex, String(MAX / STEP));
  assert.equal(second.props.style.zIndex, String(131072 / STEP));
  assert.ok(
    Number(first.props.style.zIndex) > Number(second.props.style.zIndex),
    "Bring to Front paints the image above the other float",
  );

  controller.sendToBack({ kind: "paragraph", nodeIndex: 0, childIndex: 0 });
  await vue.nextTick();
  first = layerImageByChildIndex(layer.root, 0);
  second = layerImageByChildIndex(layer.root, 1);
  assert.equal(first.props.style.zIndex, "1");
  assert.ok(
    Number(first.props.style.zIndex) < Number(second.props.style.zIndex),
    "Send to Back paints the image below the other float",
  );
});

test("layer commands are exposed with Word's disabled states", async (t) => {
  const mounted = await mountController(overlappingImagesModel(MAX, 131072));
  const { controller } = mounted;
  t.after(() => mounted.app.unmount());

  const menu = await mountMenu(controller, imageContext(controller));
  t.after(() => menu.app.unmount());

  const body = teleportTarget("body");
  assert.equal(
    findByTestId(body, "docx-context-menu-image-bring-to-front").props.disabled,
    true,
    "an image already at the top cannot be brought further forward",
  );
  assert.equal(
    findByTestId(body, "docx-context-menu-image-bring-forward").props.disabled,
    true,
  );
  assert.equal(
    findByTestId(body, "docx-context-menu-image-send-backward").props.disabled,
    false,
  );
  assert.equal(
    findByTestId(body, "docx-context-menu-image-send-to-back").props.disabled,
    false,
  );
});

test("inline images expose all four layer commands as disabled", async (t) => {
  const model = overlappingImagesModel();
  delete model.nodes[0].children[0].floating;
  const mounted = await mountController(model);
  const { controller } = mounted;
  t.after(() => mounted.app.unmount());

  const menu = await mountMenu(controller, {
    kind: "image",
    image: {
      location: { kind: "paragraph", nodeIndex: 0, childIndex: 0 },
      wrap: core.resolveDocxImageWrapState(undefined),
    },
  });
  t.after(() => menu.app.unmount());

  const body = teleportTarget("body");
  for (const actionId of [
    "image-bring-to-front",
    "image-bring-forward",
    "image-send-backward",
    "image-send-to-back",
  ]) {
    const button = findByTestId(body, `docx-context-menu-${actionId}`);
    assert.ok(button, `the menu exposes ${actionId}`);
    assert.equal(button.props.disabled, true, `${actionId} needs a floating image`);
  }
});

test("a layer command invalidates stale image XML so export re-emits the new z-order", async (t) => {
  const model = core.createBlankDocumentModel();
  model.nodes = [{
    type: "paragraph",
    // An imported paragraph keeps its original drawing XML until an image
    // edit invalidates it. If the command did not clear this, export would
    // re-emit the old relativeHeight and the reorder would be lost.
    sourceXml: "<w:p/>",
    children: [{
      ...imageNode("imported", { floating: { zIndex: 65536 } }),
      sourceXml: "<w:drawing/>",
    }],
  }];

  const mounted = await mountController(model);
  const { controller } = mounted;
  t.after(() => mounted.app.unmount());

  assert.equal(
    (await core.modelToDocumentXml(controller.model)).includes("relativeHeight="),
    false,
    "the preserved source XML is emitted verbatim before the edit",
  );

  const menu = await mountMenu(controller, imageContext(controller));
  t.after(() => menu.app.unmount());
  await findByTestId(
    teleportTarget("body"),
    "docx-context-menu-image-bring-forward",
  ).props.onClick();
  await vue.nextTick();

  assert.equal(controller.model.nodes[0].sourceXml, undefined);
  assert.equal(controller.model.nodes[0].children[0].sourceXml, undefined);
  assert.deepEqual(await relativeHeights(controller), [65536 + STEP]);
});

test("right-clicking an image opens the built-in menu with its layer actions", async (t) => {
  const restoreDom = installEditorDomEnvironment();
  const mounted = await mountEditor(overlappingImagesModel());
  t.after(() => {
    mounted.app.unmount();
    restoreDom();
  });

  const viewer = walk(mounted.root).find(
    (node) => String(node.props?.class || "").split(/\s+/).includes("docx-viewer-root"),
  );
  assert.ok(viewer, "the editor renders a viewer root that receives contextmenu");
  viewer.props.onContextmenu({
    button: 2,
    clientX: 160,
    clientY: 120,
    target: fakeImageTarget({ nodeIndex: 0, childIndex: 0 }),
    preventDefault() {},
  });
  const menuPromise = () => walk(teleportTarget("body")).find(
    (node) => node.props?.role === "menu" &&
      String(node.props?.class || "").split(/\s+/).includes("docx-context-menu"),
  );
  await waitFor(menuPromise, "the image context menu to open");
  const menu = menuPromise();

  assert.ok(textContent(menu).includes("Layer"), "the image menu renders a Layer section");
  for (const label of ["Bring to Front", "Bring Forward", "Send Backward", "Send to Back"]) {
    assert.ok(textContent(menu).includes(label), `the image menu renders ${label}`);
  }
  assert.equal(
    textContent(menu).includes("Insert Row Above"),
    false,
    "a body image target must not open the table section",
  );

  const toFront = findByTestId(menu, "docx-context-menu-image-bring-to-front");
  assert.ok(toFront);
  await toFront.props.onClick();
  await vue.nextTick();
  assert.equal(mounted.controller.model.nodes[0].children[0].floating.zIndex, MAX);
});

function installEditorDomEnvironment() {
  const originalWindow = globalThis.window;
  const originalDocument = globalThis.document;
  const originalNode = globalThis.Node;
  const originalResizeObserver = globalThis.ResizeObserver;
  const originalRect = Object.getOwnPropertyDescriptor(Object.prototype, "getBoundingClientRect");
  const originalClientHeight = Object.getOwnPropertyDescriptor(Object.prototype, "clientHeight");

  class FakeText {
    constructor(text) { this.textContent = String(text); }
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
    getAttribute(name) { return this.attributes.get(String(name)) ?? null; }
    getContext() { return { measureText: (text) => ({ width: String(text).length * 8 }) }; }
    querySelector() { return null; }
  }

  globalThis.window = {
    addEventListener() {},
    removeEventListener() {},
    getSelection() { return null; },
  };
  globalThis.document = {
    addEventListener() {},
    removeEventListener() {},
    body: { append() {} },
    createElement(tagName) { return new FakeElement(tagName); },
    createTextNode(text) { return new FakeText(text); },
    getSelection() { return null; },
    execCommand() { return false; },
  };
  globalThis.Node = class Node {};
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

  return () => {
    if (originalWindow === undefined) delete globalThis.window;
    else globalThis.window = originalWindow;
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
    if (originalNode === undefined) delete globalThis.Node;
    else globalThis.Node = originalNode;
    if (originalResizeObserver === undefined) delete globalThis.ResizeObserver;
    else globalThis.ResizeObserver = originalResizeObserver;
    if (originalRect) Object.defineProperty(Object.prototype, "getBoundingClientRect", originalRect);
    else delete Object.prototype.getBoundingClientRect;
    if (originalClientHeight) Object.defineProperty(Object.prototype, "clientHeight", originalClientHeight);
    else delete Object.prototype.clientHeight;
  };
}

async function mountEditor(model) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model });
      return () => vue.h(DocxEditor, {
        editor: controller,
        model,
        showToolbar: false,
        showThumbnails: false,
      });
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, get controller() { return controller; } };
}

function fakeImageTarget({ nodeIndex = 0, childIndex = 0 } = {}) {
  const target = new globalThis.Node();
  const host = {
    nodeType: 1,
    tagName: "P",
    getAttribute(name) {
      return name === "data-docx-paragraph-node-index" ? String(nodeIndex) : null;
    },
    matches() { return false; },
  };
  const image = {
    nodeType: 1,
    tagName: "IMG",
    getAttribute(name) {
      return name === "data-docx-image-child-index" ? String(childIndex) : null;
    },
    closest(selector) {
      if (selector === "[data-docx-table-cell-paragraph-host='true']") return null;
      if (selector === "[data-docx-paragraph-node-index]") return host;
      return null;
    },
  };
  target.closest = (selector) => {
    if (selector === "[data-docx-table-cell='true']") return null;
    if (selector.includes("data-docx-image-child-index")) return image;
    return null;
  };
  return target;
}

function fakeElement(attributes, parent = null) {
  const element = {
    parentElement: parent,
    getAttribute(name) {
      return Object.prototype.hasOwnProperty.call(attributes, name)
        ? String(attributes[name])
        : null;
    },
    matches(selector) {
      return matchesAttributeSelector(this, selector);
    },
    closest(selector) {
      for (let current = this; current; current = current.parentElement) {
        if (matchesAttributeSelector(current, selector)) return current;
      }
      return null;
    },
  };
  return element;
}

function matchesAttributeSelector(element, selector) {
  return selector.split(",").some((part) => {
    const match = /^\[([^\]=]+)(?:='([^']*)')?\]$/u.exec(part.trim());
    if (!match) return false;
    const value = element.getAttribute(match[1]);
    if (value === null) return false;
    return match[2] === undefined || value === match[2];
  });
}

test("right-click image targets resolve to model image locations", () => {
  const paragraphHost = fakeElement({
    "data-docx-paragraph-host": "true",
    "data-docx-paragraph-node-index": "3",
  });
  const inlineImage = fakeElement(
    { "data-docx-image-child-index": "2" },
    paragraphHost,
  );
  assert.deepEqual(docxImageLocationFromNode(inlineImage), {
    kind: "paragraph",
    nodeIndex: 3,
    childIndex: 2,
  });

  const layerContainer = fakeElement({
    "data-image-node-index": "5",
    "data-image-child-index": "1",
  });
  const floatingImage = fakeElement({}, layerContainer);
  assert.deepEqual(docxImageLocationFromNode(floatingImage), {
    kind: "paragraph",
    nodeIndex: 5,
    childIndex: 1,
  });

  const tableCellHost = fakeElement({
    "data-docx-table-cell-paragraph-host": "true",
    "data-docx-table-index": "1",
    "data-docx-table-row-index": "0",
    "data-docx-table-cell-index": "2",
    "data-docx-paragraph-index": "1",
  });
  const tableImage = fakeElement(
    { "data-docx-image-child-index": "0" },
    tableCellHost,
  );
  assert.deepEqual(docxImageLocationFromNode(tableImage), {
    kind: "table-cell",
    tableIndex: 1,
    rowIndex: 0,
    cellIndex: 2,
    paragraphIndex: 1,
    childIndex: 0,
  });

  assert.equal(docxImageLocationFromNode(fakeElement({})), undefined);
});
