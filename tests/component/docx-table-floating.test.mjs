// Issue #150 — DOCX table move handle and floating-table drag.
//
// Every block names the bullet it covers so the mapping stays checkable:
//
//   1. move handle + drag ghost .................... DocxTableHost
//   2. handle selects the whole table .............. selectTableCellRange
//   3. drag → `tblpPr` with drop position + wrap ... setTableFloating
//   4. export/reopen keeps it, undo restores inline . modelToDocumentXml / undo
//   5. disabled in read-only and off-body tables .... editable / nested / header

import assert from "node:assert/strict";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

import { importFromDemo, mount, vue, walk } from "./vue-test-renderer.mjs";

const requireFromDocxCore = createRequire(
  new URL("../../packages/docx-core/package.json", import.meta.url),
);
const core = await import(
  pathToFileURL(requireFromDocxCore.resolve("@arcships/docx-core")).href
);

const { DocxTableHost, useDocxEditor } = await importFromDemo("@arcships/vue-docx");
const { createBlankDocumentModel } = await importFromDemo("@arcships/docx-core");

// ── Fixtures ───────────────────────────────────────────────────────
function paragraph(text) {
  return { type: "paragraph", children: [{ type: "text", text }] };
}

function tableNode(rows = 2) {
  return {
    type: "table",
    style: {
      layout: "fixed",
      widthTwips: 6000,
      columnWidthsTwips: [3000, 3000],
    },
    rows: Array.from({ length: rows }, (_, rowIndex) => ({
      type: "table-row",
      style: { heightTwips: 600 },
      cells: [0, 1].map((cellIndex) => ({
        type: "table-cell",
        style: { widthTwips: 3000 },
        nodes: [paragraph(`r${rowIndex}c${cellIndex}`)],
      })),
    })),
  };
}

function modelWithNodes(nodes) {
  return { ...createBlankDocumentModel(), nodes };
}

const PAGE_ORIGIN = {
  left: 0,
  top: 0,
  marginLeft: 96,
  marginTop: 96,
  columnLeft: 96,
  columnTop: 96,
  pageWidth: 816,
};

// ── DOM environment ────────────────────────────────────────────────
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
    dispatch(type, event) {
      for (const listener of [...(listeners.get(type) ?? [])]) listener(event);
    },
    count(type) {
      return listeners.get(type)?.size ?? 0;
    },
  };
}

function installDomEnvironment() {
  const originalWindow = globalThis.window;
  const originalDocument = globalThis.document;
  const originalResizeObserver = globalThis.ResizeObserver;
  const documentListeners = createListenerTarget();

  class FakeElement {
    constructor(tagName) {
      this.tagName = String(tagName).toLowerCase();
      this.attributes = new Map();
      this.children = [];
      this.style = { cssText: "", setProperty(name, value) { this[name] = value; } };
    }
    appendChild(child) {
      this.children.push(child);
      return child;
    }
    setAttribute(name, value) {
      this.attributes.set(String(name), String(value));
    }
    get innerHTML() {
      return this.children.map((child) => child.outerHTML ?? "").join("");
    }
    get outerHTML() {
      return `<${this.tagName}>${this.innerHTML}</${this.tagName}>`;
    }
    getContext() {
      return { measureText: (text) => ({ width: String(text).length * 8 }) };
    }
    querySelector() {
      return null;
    }
    querySelectorAll() {
      return [];
    }
  }

  globalThis.window = {
    addEventListener() {},
    removeEventListener() {},
    getSelection() {
      return null;
    },
  };
  globalThis.document = {
    addEventListener: documentListeners.addEventListener,
    removeEventListener: documentListeners.removeEventListener,
    body: { append() {} },
    createElement(tagName) {
      return new FakeElement(tagName);
    },
    createTextNode(text) {
      return { textContent: String(text) };
    },
    getSelection() {
      return null;
    },
    execCommand() {
      return false;
    },
  };
  globalThis.ResizeObserver = class ResizeObserver {
    observe() {}
    disconnect() {}
  };

  function restore() {
    if (originalWindow === undefined) delete globalThis.window;
    else globalThis.window = originalWindow;
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
    if (originalResizeObserver === undefined) delete globalThis.ResizeObserver;
    else globalThis.ResizeObserver = originalResizeObserver;
  }

  return { documentListeners, restore };
}

// ── Mount helpers ──────────────────────────────────────────────────
async function mountController(model) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model });
      return () => vue.h("div", { "data-testid": "controller-harness" }, controller.status);
    },
  });
  const mounted = await mount(Harness);
  return {
    ...mounted,
    get controller() {
      return controller;
    },
  };
}

async function mountTableHost(controller, props = {}) {
  const Harness = vue.defineComponent({
    setup() {
      return () =>
        vue.h(DocxTableHost, {
          table: controller.model.nodes[0],
          tableIndex: 0,
          editable: true,
          controller,
          maxAvailableWidthPx: 600,
          floatingPageOriginPx: PAGE_ORIGIN,
          ...props,
        });
    },
  });
  return mount(Harness);
}

function moveHandles(root) {
  return walk(root).filter(
    (node) => node.props?.["data-docx-table-move-handle"] !== undefined,
  );
}

function dragGhosts(root) {
  return walk(root).filter(
    (node) => node.props?.["data-docx-table-drag-ghost"] !== undefined,
  );
}

function tableHostElement(root) {
  return walk(root).find((node) => node.props?.["data-docx-table-host"] !== undefined);
}

async function hoverTable(host) {
  tableHostElement(host.root).props.onPointerenter({});
  await vue.nextTick();
}

// ── 1. move handle and drag ghost ──────────────────────────────────
test("hovering a body table shows its move handle and the drag renders a ghost", async (t) => {
  const env = installDomEnvironment();
  const mounted = await mountController(modelWithNodes([tableNode(2)]));
  const host = await mountTableHost(mounted.controller);
  t.after(() => {
    host.app.unmount();
    mounted.app.unmount();
    env.restore();
  });

  assert.equal(moveHandles(host.root).length, 0, "the handle waits for the hover");
  await hoverTable(host);
  const handles = moveHandles(host.root);
  assert.equal(handles.length, 1, "one move handle per table");
  assert.equal(handles[0].props["data-docx-thumbnail-exclude"], "true");
  assert.equal(handles[0].props.title, "Drag to move or float this table");
  assert.equal(handles[0].props.class, "docx-table-move-handle");

  handles[0].props.onMousedown({ preventDefault() {}, button: 0, clientX: 40, clientY: 60 });
  await vue.nextTick();
  assert.equal(moveHandles(host.root).length, 1, "the handle stays while dragging");
  assert.equal(dragGhosts(host.root).length, 1, "the drag shows a ghost");

  env.documentListeners.dispatch("mousemove", { clientX: 70, clientY: 72 });
  await vue.nextTick();
  assert.equal(
    dragGhosts(host.root)[0].props.style.transform,
    "translate(30px, 12px)",
    "the ghost follows the pointer",
  );

  env.documentListeners.dispatch("mouseup", { clientX: 70, clientY: 72 });
  await vue.nextTick();
  assert.equal(dragGhosts(host.root).length, 0, "the ghost disappears on drop");
  assert.deepEqual(mounted.warnings, []);
});

// ── 2. the handle selects the whole table ──────────────────────────
test("pressing the move handle selects every cell of the table", async (t) => {
  const env = installDomEnvironment();
  const mounted = await mountController(modelWithNodes([tableNode(3)]));
  const host = await mountTableHost(mounted.controller);
  t.after(() => {
    host.app.unmount();
    mounted.app.unmount();
    env.restore();
  });

  await hoverTable(host);
  moveHandles(host.root)[0].props.onMousedown({
    preventDefault() {},
    button: 0,
    clientX: 10,
    clientY: 10,
  });
  await vue.nextTick();

  const selection = mounted.controller.selection;
  assert.equal(selection.kind, "table-cell");
  assert.equal(selection.tableIndex, 0);
  assert.deepEqual(selection.cellRange, {
    tableIndex: 0,
    anchorRowIndex: 0,
    anchorCellIndex: 0,
    focusRowIndex: 2,
    focusCellIndex: 1,
  });
  env.documentListeners.dispatch("mouseup", { clientX: 10, clientY: 10 });
});

// ── 3. drag → tblpPr with drop position and wrap ───────────────────
test("dragging the handle floats the table at the drop position", async (t) => {
  const env = installDomEnvironment();
  const mounted = await mountController(modelWithNodes([tableNode(2)]));
  const host = await mountTableHost(mounted.controller);
  t.after(() => {
    host.app.unmount();
    mounted.app.unmount();
    env.restore();
  });

  await hoverTable(host);
  moveHandles(host.root)[0].props.onMousedown({
    preventDefault() {},
    button: 0,
    clientX: 100,
    clientY: 100,
  });
  env.documentListeners.dispatch("mousemove", { clientX: 196, clientY: 124 });
  env.documentListeners.dispatch("mouseup", { clientX: 196, clientY: 124 });
  await vue.nextTick();

  const table = mounted.controller.model.nodes[0];
  // 96px right and 24px down = 1440 / 360 twips, anchored to the margin.
  assert.deepEqual(table.style.floating, {
    xTwips: 1440,
    yTwips: 360,
    horizontalAnchor: "margin",
    verticalAnchor: "text",
  });
  assert.equal(table.sourceXml, undefined, "export rebuilds the table");
  assert.equal(mounted.controller.status, "Floated table");
  assert.equal(mounted.controller.canUndo, true);

  const xml = await core.modelToDocumentXml(mounted.controller.model);
  assert.match(
    xml,
    /<w:tblPr><w:tblpPr w:vertAnchor="text" w:horzAnchor="margin" w:tblpX="1440" w:tblpY="360"\/>/,
    "the drop position reaches w:tblpPr",
  );
  assert.deepEqual(mounted.warnings, []);
});

test("the drop is clamped so the table stays on the page", async (t) => {
  const env = installDomEnvironment();
  const mounted = await mountController(modelWithNodes([tableNode(1)]));
  const host = await mountTableHost(mounted.controller, { maxAvailableWidthPx: 200 });
  t.after(() => {
    host.app.unmount();
    mounted.app.unmount();
    env.restore();
  });

  await hoverTable(host);
  moveHandles(host.root)[0].props.onMousedown({
    preventDefault() {},
    button: 0,
    clientX: 0,
    clientY: 0,
  });
  env.documentListeners.dispatch("mousemove", { clientX: 5000, clientY: 0 });
  env.documentListeners.dispatch("mouseup", { clientX: 5000, clientY: 0 });
  await vue.nextTick();

  assert.equal(
    mounted.controller.model.nodes[0].style.floating.xTwips,
    (PAGE_ORIGIN.pageWidth - 42) * 15,
    "the left edge stops 42px inside the page",
  );
});

// ── 4. export/reopen keeps it, undo restores inline ────────────────
test("a dragged table reopens floating and undo restores the inline layout", async (t) => {
  const env = installDomEnvironment();
  const mounted = await mountController(modelWithNodes([tableNode(2)]));
  const host = await mountTableHost(mounted.controller);
  t.after(() => {
    host.app.unmount();
    mounted.app.unmount();
    env.restore();
  });

  await hoverTable(host);
  moveHandles(host.root)[0].props.onMousedown({
    preventDefault() {},
    button: 0,
    clientX: 100,
    clientY: 100,
  });
  env.documentListeners.dispatch("mousemove", { clientX: 130, clientY: 112 });
  env.documentListeners.dispatch("mouseup", { clientX: 130, clientY: 112 });
  await vue.nextTick();

  const bytes = await core.serializeDocx(mounted.controller.model);
  const reopened = (await core.buildDocModelFromBytes(bytes)).model;
  assert.deepEqual(reopened.nodes[0].style.floating, {
    xTwips: 450,
    yTwips: 180,
    horizontalAnchor: "margin",
    verticalAnchor: "text",
  });

  mounted.controller.undo();
  await vue.nextTick();
  assert.equal(
    mounted.controller.model.nodes[0].style.floating,
    undefined,
    "one undo puts the table back into the text flow",
  );
  const inlineXml = await core.modelToDocumentXml(mounted.controller.model);
  assert.doesNotMatch(inlineXml, /<w:tblpPr/);
});

// ── 5. controller command details ──────────────────────────────────
test("a click on the handle only selects; it does not float the table", async (t) => {
  const env = installDomEnvironment();
  const mounted = await mountController(modelWithNodes([tableNode(2)]));
  const host = await mountTableHost(mounted.controller);
  t.after(() => {
    host.app.unmount();
    mounted.app.unmount();
    env.restore();
  });

  await hoverTable(host);
  moveHandles(host.root)[0].props.onMousedown({
    preventDefault() {},
    button: 0,
    clientX: 40,
    clientY: 40,
  });
  env.documentListeners.dispatch("mousemove", { clientX: 41, clientY: 40 });
  env.documentListeners.dispatch("mouseup", { clientX: 41, clientY: 40 });
  await vue.nextTick();

  assert.equal(mounted.controller.model.nodes[0].style.floating, undefined);
  assert.equal(mounted.controller.canUndo, false, "a click opens no undo entry");
  assert.equal(mounted.controller.selection.kind, "table-cell");
});

test("setTableFloating keeps wrap distances, reports refusals and clears to inline", async (t) => {
  const mounted = await mountController(modelWithNodes([tableNode(1)]));
  t.after(() => mounted.app.unmount());
  const { controller } = mounted;

  controller.setTableFloating(0, {
    xTwips: 720,
    yTwips: 0,
    horizontalAnchor: "margin",
    verticalAnchor: "text",
    leftFromTextTwips: 180,
    rightFromTextTwips: 180,
  });
  controller.setTableFloating(0, { xTwips: 1440 });
  assert.deepEqual(controller.model.nodes[0].style.floating, {
    xTwips: 1440,
    yTwips: 0,
    horizontalAnchor: "margin",
    verticalAnchor: "text",
    leftFromTextTwips: 180,
    rightFromTextTwips: 180,
  });
  assert.equal(controller.status, "Floated table");

  // A second call with the same position is a no-op and opens no undo entry.
  const undoBefore = controller.canUndo;
  controller.setTableFloating(0, { xTwips: 1440 });
  assert.equal(controller.canUndo, undoBefore);

  // Unrepresentable values never reach the model.
  controller.setTableFloating(0, { xTwips: Number.NaN, horizontalAnchor: "somewhere" });
  assert.equal(controller.model.nodes[0].style.floating.xTwips, 1440);

  controller.setTableFloating(0, null);
  assert.equal(controller.model.nodes[0].style.floating, undefined);
  assert.equal(controller.status, "Table back in the text flow");

  controller.setTableFloating(9, { xTwips: 0 });
  assert.equal(controller.status, "Cannot float: table not found");
  assert.deepEqual(mounted.warnings, []);
});

// ── 6. read-only and off-body tables ───────────────────────────────
test("read-only, nested, header and opted-out hosts render no move handle", async (t) => {
  const env = installDomEnvironment();
  const mounted = await mountController(modelWithNodes([tableNode(2)]));
  t.after(() => {
    mounted.app.unmount();
    env.restore();
  });

  const readOnly = await mountTableHost(mounted.controller, { editable: false });
  await hoverTable(readOnly);
  assert.equal(moveHandles(readOnly.root).length, 0, "read-only tables have no handle");
  readOnly.app.unmount();

  const nested = await mountTableHost(mounted.controller, {
    tableRuntimeKey: "body-cell-nested-table-0-0-0",
  });
  await hoverTable(nested);
  assert.equal(moveHandles(nested.root).length, 0, "nested tables cannot float");
  nested.app.unmount();

  const header = await mountTableHost(mounted.controller, {
    withinHeaderFooter: true,
    headerFooterRegion: "header",
  });
  await hoverTable(header);
  assert.equal(moveHandles(header.root).length, 0, "header tables cannot float");
  header.app.unmount();

  const optedOut = await mountTableHost(mounted.controller, {
    showTableMoveHandle: false,
  });
  await hoverTable(optedOut);
  assert.equal(moveHandles(optedOut.root).length, 0, "hosts can opt out");
  optedOut.app.unmount();

  const editable = await mountTableHost(mounted.controller);
  await hoverTable(editable);
  assert.equal(moveHandles(editable.root).length, 1, "an editable body table shows it");
  editable.app.unmount();
});
