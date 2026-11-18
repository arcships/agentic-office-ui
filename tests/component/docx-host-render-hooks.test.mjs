// Host render hooks for DOCX cards and context menus (#151).
//
// Upstream exposes `renderCommentCard`, `renderTrackedChangeCard`,
// `renderContextMenu` and `renderTableContextMenu`. These tests cover every
// bullet of the gap:
//
//   1. a host can replace comment and tracked-change cards, with accept /
//      reject / resolve callbacks wired to the controller commands;
//   2. a host can replace the text, image and table context menus, receiving
//      the built-in actions and a `runAction()` bridge;
//   3. read-only editing sessions still open a Copy-only default menu;
//   4. the default appearance is unchanged when no hook is provided.

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
  DocxDocumentSurface,
  DocxEditor,
  DocxPageSurface,
  DocxTrackedChangeGutter,
  useDocxEditor,
} = await importFromDemo("@arcships/vue-docx");
const core = await importFromDemo("@arcships/docx-core");

// ── Fixtures ───────────────────────────────────────────────────────
const PAGE_LAYOUT = {
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginsPx: { top: 72, right: 72, bottom: 72, left: 72 },
};

const PARAGRAPH_REVISION_XML =
  "<w:p>" +
  '<w:r><w:t xml:space="preserve">Response time changed from </w:t></w:r>' +
  '<w:del w:id="1" w:author="Alex Reviewer" w:date="2026-07-11T00:00:00Z">' +
  "<w:r><w:delText>five business days</w:delText></w:r></w:del>" +
  '<w:r><w:t xml:space="preserve"> to </w:t></w:r>' +
  '<w:ins w:id="2" w:author="Morgan Editor" w:date="2026-07-11T00:05:00Z">' +
  "<w:r><w:t>two business days</w:t></w:r></w:ins>" +
  "<w:r><w:t>.</w:t></w:r>" +
  "</w:p>";

function paragraph(text) {
  return { type: "paragraph", children: [{ type: "text", text }] };
}

function starterModel(...texts) {
  return {
    ...core.createBlankDocumentModel(),
    nodes: texts.map(paragraph),
  };
}

function revisionModel() {
  return {
    ...core.createBlankDocumentModel(),
    nodes: [
      {
        type: "paragraph",
        children: [
          { type: "text", text: "Response time changed from " },
          { type: "text", text: " to " },
          { type: "text", text: "two business days" },
          { type: "text", text: "." },
        ],
        sourceXml: PARAGRAPH_REVISION_XML,
      },
    ],
  };
}

function textRange(nodeIndex, start, end) {
  return {
    start: { location: { kind: "paragraph", nodeIndex }, offset: start },
    end: { location: { kind: "paragraph", nodeIndex }, offset: end },
  };
}

function paragraphText(node) {
  return node.children.map((child) => child.text ?? "").join("");
}

function fakeEvent() {
  return { stopPropagation() {}, preventDefault() {} };
}

/** Cards are positioned after mount, so give the layout a few ticks. */
async function settleGutter() {
  await vue.nextTick();
  await vue.nextTick();
  await vue.nextTick();
}

function hasClass(node, className) {
  return String(node?.props?.class || "").split(/\s+/).includes(className);
}

function gutterCards(root) {
  return walk(root).filter(
    (node) => node.props?.["data-docx-gutter-card-id"] !== undefined,
  );
}

async function mountPageSurface(options) {
  const { model, controller, ...props } = options;
  const mounted = await mount(DocxPageSurface, {
    pageIndex: 0,
    pageLayout: PAGE_LAYOUT,
    pageNodeSegments: model.nodes.map((_, nodeIndex) => ({ nodeIndex })),
    model,
    pageWidthPx: PAGE_LAYOUT.pageWidthPx,
    pageContentWidthPx: 672,
    pageNumber: 1,
    totalPages: 1,
    commentsEnabled: true,
    trackedChangesEnabled: true,
    editable: true,
    controller,
    ...props,
  });
  await settleGutter();
  return mounted;
}

async function commentSurface({ editable = true, ...hooks } = {}) {
  const controller = useDocxEditor({
    starterModel: starterModel("Alpha beta gamma delta"),
  });
  controller.addComment(textRange(0, 6, 16), "Please confirm.");
  const mounted = await mountPageSurface({
    model: controller.model,
    controller,
    editable,
    ...hooks,
  });
  return { ...mounted, controller };
}

async function revisionSurface({ editable = true, ...hooks } = {}) {
  const controller = useDocxEditor({ starterModel: revisionModel() });
  controller.setShowTrackedChanges(true);
  const mounted = await mountPageSurface({
    model: controller.model,
    controller,
    editable,
    revisionDisplay: "balloons",
    ...hooks,
  });
  return { ...mounted, controller };
}

// ── Bullet 1: custom cards ─────────────────────────────────────────
test("gutter cards keep their default markup when no host renderer is provided", async (t) => {
  const mounted = await commentSurface();
  t.after(() => mounted.app.unmount());

  const [card] = gutterCards(mounted.root);
  assert.ok(card, "the comment gutter card renders");
  assert.ok(
    walk(card).some((node) => hasClass(node, "docx-gutter-card-type")),
    "the default card header is unchanged",
  );
  assert.match(textContent(card), /Please confirm\./u);
  assert.ok(findByTestId(card, "docx-gutter-comment-resolve"), "default resolve stays");
  assert.equal(findByTestId(mounted.root, "custom-comment-card"), undefined);
  assert.deepEqual(mounted.warnings, []);
});

test("a host comment card receives the comment, its page and resolve() wired to the controller", async (t) => {
  const seen = [];
  const mounted = await commentSurface({
    renderCommentCard: (card) => {
      seen.push(card);
      return vue.h(
        "div",
        { "data-testid": "custom-comment-card" },
        [
          vue.h("span", { class: "custom-author" }, card.comment.author),
          vue.h("span", { class: "custom-snippet" }, card.snippet),
          vue.h(
            "button",
            {
              "data-testid": "custom-resolve",
              onClick: () => card.setResolved(true),
            },
            "Resolve thread",
          ),
        ],
      );
    },
  });
  t.after(() => mounted.app.unmount());

  const custom = findByTestId(mounted.root, "custom-comment-card");
  assert.ok(custom, "the host card replaces the comment card body");

  const [card] = gutterCards(mounted.root);
  assert.ok(
    !walk(card).some((node) => hasClass(node, "docx-gutter-card-type")),
    "the default card body is not rendered alongside the host card",
  );
  assert.equal(
    findByTestId(mounted.root, "docx-gutter-comment-resolve"),
    undefined,
    "the default comment actions are replaced",
  );

  const entry = seen.at(-1);
  assert.equal(entry.comment.text, "Please confirm.");
  assert.equal(entry.snippet, "Please confirm.");
  assert.equal(entry.pageIndex, 0);
  assert.equal(entry.accentColor.length > 0, true);
  assert.equal(typeof entry.style.top, "string");
  assert.equal(entry.canModify, true);
  assert.equal(typeof entry.activate, "function");

  findByTestId(mounted.root, "custom-resolve").props.onClick(fakeEvent());
  await vue.nextTick();
  assert.equal(mounted.controller.comments[0].resolved, true, "resolve() resolves the thread");
  assert.equal(mounted.controller.model.metadata.comments[0].resolved, true);
  assert.deepEqual(mounted.warnings, []);
});

test("a read-only revision card reports canDecide=false and accept/reject are no-ops", async (t) => {
  const mounted = await revisionSurface({
    editable: false,
    renderTrackedChangeCard: (card) =>
      vue.h("button", {
        "data-testid": "custom-decide",
        onClick: () => {
          card.accept();
          card.reject();
        },
      }, String(card.canDecide)),
  });
  t.after(() => mounted.app.unmount());

  const button = findByTestId(mounted.root, "custom-decide");
  assert.equal(textContent(button), "false");
  button.props.onClick(fakeEvent());
  await vue.nextTick();
  assert.equal(
    paragraphText(mounted.controller.model.nodes[0]),
    "Response time changed from  to two business days.",
    "the model is unchanged while the revision is still displayed",
  );
  assert.equal(mounted.controller.canUndo, false);
  assert.deepEqual(mounted.warnings, []);
});

test("a read-only comment card reports canModify=false and resolve() is a no-op", async (t) => {
  const mounted = await commentSurface({
    editable: false,
    renderCommentCard: (card) =>
      vue.h("button", {
        "data-testid": "custom-resolve",
        onClick: () => card.setResolved(true),
      }, String(card.canModify)),
  });
  t.after(() => mounted.app.unmount());

  const button = findByTestId(mounted.root, "custom-resolve");
  assert.equal(textContent(button), "false");
  button.props.onClick(fakeEvent());
  await vue.nextTick();
  assert.equal(mounted.controller.comments[0].resolved, undefined);
  assert.deepEqual(mounted.warnings, []);
});

test("a host revision card receives the change and accept/reject run the controller commands", async (t) => {
  const seen = [];
  const mounted = await revisionSurface({
    renderTrackedChangeCard: (card) => {
      seen.push(card);
      return vue.h(
        "div",
        { "data-testid": "custom-change-card" },
        [
          vue.h("span", { class: "custom-kind" }, card.kindLabel),
          vue.h("span", { class: "custom-snippet" }, card.snippet),
          vue.h(
            "button",
            {
              "data-testid": "custom-accept",
              onClick: () => card.accept(),
            },
            "Accept",
          ),
          vue.h(
            "button",
            {
              "data-testid": "custom-reject",
              onClick: () => card.reject(),
            },
            "Reject",
          ),
        ],
      );
    },
  });
  t.after(() => mounted.app.unmount());

  assert.ok(findByTestId(mounted.root, "custom-change-card"));
  assert.equal(
    findByTestId(mounted.root, "docx-gutter-accept"),
    undefined,
    "the default revision actions are replaced",
  );

  const entry = seen.at(-1);
  assert.equal(entry.change.kind, "deletion");
  assert.match(entry.kindLabel, /Deleted/u);
  assert.match(entry.snippet, /five business days/u);
  assert.equal(entry.canDecide, true);
  assert.equal(entry.pageIndex, 0);

  findByTestId(mounted.root, "custom-accept").props.onClick(fakeEvent());
  await vue.nextTick();
  assert.equal(
    paragraphText(mounted.controller.model.nodes[0]),
    "Response time changed from  to two business days.",
    "accept() resolves the revision like the built-in card",
  );
  assert.equal(mounted.controller.canUndo, true);
  assert.deepEqual(gutterCards(mounted.root), []);
  assert.deepEqual(mounted.warnings, []);
});

test("a host that only overrides one card type keeps the default card for the other", async (t) => {
  const mounted = await revisionSurface({
    renderCommentCard: () =>
      vue.h("div", { "data-testid": "custom-comment-card" }),
  });
  t.after(() => mounted.app.unmount());

  assert.equal(findByTestId(mounted.root, "custom-comment-card"), undefined);
  assert.ok(findByTestId(mounted.root, "docx-gutter-accept"), "revisions keep the default card");
  assert.deepEqual(mounted.warnings, []);
});

test("host card callbacks emit the card commands with their annotation payloads", async (t) => {
  const comment = {
    id: "comment-1",
    commentId: 7,
    author: "Alex Reviewer",
    text: "Please confirm.",
    nodeIndex: 0,
    location: { kind: "paragraph", nodeIndex: 0 },
  };
  const change = {
    id: "change-1",
    kind: "deletion",
    author: "Morgan Editor",
    text: "five business days",
    nodeIndex: 0,
    location: { kind: "paragraph", nodeIndex: 0 },
  };
  const positioned = [
    {
      annotation: { id: "comment:comment-1", comment, ordinal: 1 },
      anchorX: 0,
      anchorY: 0,
      top: 10,
      heightPx: 20,
      connectorLane: 0,
    },
    {
      annotation: { id: "change:change-1", trackedChange: change, ordinal: 2 },
      anchorX: 0,
      anchorY: 0,
      top: 40,
      heightPx: 20,
      connectorLane: 0,
    },
  ];
  const emitted = [];
  const mounted = await mount(DocxTrackedChangeGutter, {
    positioned,
    zoomFactor: 1,
    pageHeightPx: 500,
    pageIndex: 2,
    commentActions: true,
    decidable: true,
    renderCommentCard: (card) => {
      assert.equal(card.pageIndex, 2, "the pane reports its page to host cards");
      return vue.h("button", {
        "data-testid": "emit-resolve",
        onClick: () => {
          card.activate();
          card.setResolved(true);
        },
      }, "Resolve");
    },
    renderTrackedChangeCard: (card) =>
      vue.h("button", {
        "data-testid": "emit-reject",
        onClick: () => card.reject(),
      }, "Reject"),
    onActivate: (annotationId) => emitted.push(["activate", annotationId]),
    onResolve: (commentId, resolved) => emitted.push(["resolve", commentId, resolved]),
    onDecide: (annotationId, decision) => emitted.push(["decide", annotationId, decision]),
  });
  t.after(() => mounted.app.unmount());

  findByTestId(mounted.root, "emit-resolve").props.onClick(fakeEvent());
  findByTestId(mounted.root, "emit-reject").props.onClick(fakeEvent());
  assert.deepEqual(emitted, [
    ["activate", "comment:comment-1"],
    ["resolve", 7, true],
    ["decide", "change:change-1", "reject"],
  ]);
  assert.deepEqual(mounted.warnings, []);
});

test("custom cards flow through DocxDocumentSurface", async (t) => {
  const controller = useDocxEditor({
    starterModel: starterModel("Alpha beta gamma delta"),
  });
  controller.addComment(textRange(0, 6, 16), "Please confirm.");
  const mounted = await mount(DocxDocumentSurface, {
    model: controller.model,
    controller,
    renderCommentCard: (card) =>
      vue.h("div", { "data-testid": "surface-custom-comment-card" }, card.comment.text),
  });
  await settleGutter();
  t.after(() => mounted.app.unmount());

  const card = findByTestId(mounted.root, "surface-custom-comment-card");
  assert.ok(card, "DocxDocumentSurface forwards the card renderer to the page surface");
  assert.match(textContent(card), /Please confirm\./u);
  assert.deepEqual(mounted.warnings, []);
});

// ── Bullet 2: custom context menus ─────────────────────────────────
function menuElement() {
  return walk(teleportTarget("body")).find(
    (node) => hasClass(node, "docx-context-menu") && node.props.role === "menu",
  );
}

function hostMenuElement() {
  return findInBody("docx-context-menu-host");
}

function findInBody(testId) {
  return walk(teleportTarget("body")).find(
    (node) => node.props?.["data-testid"] === testId,
  );
}

async function mountMenu({ controller, context, ...props }) {
  const current = vue.ref(context);
  const closes = [];
  const Harness = vue.defineComponent({
    setup() {
      return () =>
        vue.h(DocxContextMenu, {
          context: current.value,
          controller,
          x: 120,
          y: 90,
          onClose: () => {
            closes.push("close");
            current.value = undefined;
          },
          ...props,
        });
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, context: current, closes };
}

function stubMenuController() {
  const calls = [];
  return {
    calls,
    controller: {
      documentTheme: "light",
      selectedRunStyle: {},
      copy: async () => {
        calls.push("copy");
      },
      setHighlight: (value) => calls.push(`highlight:${value}`),
      setTextColor: (value) => calls.push(`textColor:${value}`),
    },
  };
}

test("a host menu replaces the default menu and receives position, actions and closeMenu", async (t) => {
  const { controller, calls } = stubMenuController();
  const seen = [];
  const mounted = await mountMenu({
    controller,
    context: { kind: "text" },
    renderContextMenu: (menu) => {
      seen.push(menu);
      return vue.h("div", { "data-testid": "custom-menu" }, [
        vue.h(
          "button",
          { "data-testid": "custom-menu-copy", onClick: () => menu.runAction("copy") },
          "Copy",
        ),
        vue.h(
          "button",
          { "data-testid": "custom-menu-close", onClick: () => menu.closeMenu() },
          "Close",
        ),
      ]);
    },
  });
  t.after(() => mounted.app.unmount());
  await waitFor(hostMenuElement, "the host menu to render");

  assert.equal(menuElement(), undefined, "the built-in menu is replaced");
  assert.ok(findInBody("custom-menu"), "the host menu renders inside the pointer layer");

  const menu = seen.at(-1);
  assert.deepEqual(menu.position, { x: 120, y: 90 });
  assert.equal(menu.context.kind, "text");
  assert.equal(menu.readOnly, false);
  assert.equal(menu.documentTheme, "light");
  assert.deepEqual(
    menu.actions.slice(0, 3).map((action) => action.id),
    ["cut", "copy", "paste"],
    "host menus receive the built-in actions in menu order",
  );
  assert.ok(
    menu.actions.some((action) => action.id === "bold"),
    "formatting actions are included",
  );

  findInBody("custom-menu-copy").props.onClick(fakeEvent());
  await waitFor(() => calls.includes("copy"), "runAction('copy') to copy");
  assert.deepEqual(mounted.closes, ["close"], "running an action closes the menu");
  assert.deepEqual(mounted.warnings, []);
});

test("a host menu's closeMenu closes the menu without running an action", async (t) => {
  const { controller, calls } = stubMenuController();
  const mounted = await mountMenu({
    controller,
    context: { kind: "text" },
    renderContextMenu: (menu) =>
      vue.h("button", {
        "data-testid": "custom-menu-close",
        onClick: () => menu.closeMenu(),
      }, "Close"),
  });
  t.after(() => mounted.app.unmount());
  await waitFor(hostMenuElement, "the host menu to render");

  findInBody("custom-menu-close").props.onClick(fakeEvent());
  await waitFor(() => mounted.closes.length === 1, "closeMenu to close the menu");
  assert.deepEqual(calls, []);
  assert.deepEqual(mounted.warnings, []);
});

test("the table hook wins for table targets and runAction runs the built-in table command", async (t) => {
  const controller = useDocxEditor({
    starterModel: {
      ...core.createBlankDocumentModel(),
      nodes: [
        {
          type: "table",
          rows: [
            { cells: [{ nodes: [paragraph("A1")] }] },
            { cells: [{ nodes: [paragraph("A2")] }] },
          ],
        },
      ],
    },
  });
  const seen = [];
  const mounted = await mountMenu({
    controller,
    context: {
      kind: "table",
      tableContext: { tableIndex: 0, rowIndex: 0, cellIndex: 0 },
    },
    renderContextMenu: () => vue.h("div", { "data-testid": "general-menu" }),
    renderTableContextMenu: (menu) => {
      seen.push(menu);
      return vue.h(
        "button",
        { "data-testid": "custom-table-menu", onClick: () => menu.runAction("insert-row-below") },
        "Insert row below",
      );
    },
  });
  t.after(() => mounted.app.unmount());
  await waitFor(hostMenuElement, "the host table menu to render");

  assert.equal(findInBody("general-menu"), undefined, "the table hook wins");
  const menu = seen.at(-1);
  assert.deepEqual(menu.context, { tableIndex: 0, rowIndex: 0, cellIndex: 0 });
  assert.ok(
    menu.actions.some((action) => action.id === "insert-row-above"),
    "table actions are exposed to the host",
  );

  findInBody("custom-table-menu").props.onClick(fakeEvent());
  await waitFor(
    () => controller.model.nodes[0].rows.length === 3,
    "runAction to insert a table row",
  );
  assert.equal(controller.model.nodes[0].rows.length, 3);
  assert.deepEqual(mounted.warnings, []);
});

test("a host menu receives filtered actions and readOnly=true in a read-only document", async (t) => {
  const { controller } = stubMenuController();
  const seen = [];
  const mounted = await mountMenu({
    controller,
    context: { kind: "text" },
    readOnly: true,
    renderContextMenu: (menu) => {
      seen.push(menu);
      return vue.h("div", { "data-testid": "custom-menu" });
    },
  });
  t.after(() => mounted.app.unmount());
  await waitFor(hostMenuElement, "the host menu to render");

  const menu = seen.at(-1);
  assert.equal(menu.readOnly, true);
  assert.deepEqual(menu.actions.map((action) => action.id), ["copy"]);
  assert.deepEqual(mounted.warnings, []);
});

test("read-only table targets skip the table hook so editing commands cannot run", async (t) => {
  const { controller } = stubMenuController();
  let tableMenuCalls = 0;
  const mounted = await mountMenu({
    controller,
    context: {
      kind: "table",
      tableContext: { tableIndex: 0, rowIndex: 0, cellIndex: 0 },
    },
    readOnly: true,
    renderTableContextMenu: () => {
      tableMenuCalls += 1;
      return vue.h("div", { "data-testid": "custom-table-menu" });
    },
  });
  t.after(() => mounted.app.unmount());
  await waitFor(menuElement, "the default read-only menu to render");

  assert.equal(tableMenuCalls, 0, "the table hook is not called in a read-only document");
  assert.equal(hostMenuElement(), undefined);
  assert.ok(findInBody("docx-context-menu-copy"), "the Copy menu is still available");
  assert.deepEqual(
    walk(menuElement())
      .filter((node) => node.props?.role === "menuitem")
      .map((node) => node.props["data-testid"]),
    ["docx-context-menu-copy"],
  );
  assert.deepEqual(mounted.warnings, []);
});

test("the default menu keeps its appearance when no menu hook is provided", async (t) => {
  const { controller } = stubMenuController();
  const mounted = await mountMenu({ controller, context: { kind: "text", position: { x: 30, y: 40 } } });
  t.after(() => mounted.app.unmount());
  await waitFor(menuElement, "the default menu to render");

  const menu = menuElement();
  assert.ok(menu);
  assert.equal(menu.props["data-docx-context-menu-read-only"], undefined);
  const items = walk(menu).filter((node) => node.props?.role === "menuitem");
  assert.ok(items.length > 10, "every built-in action is still rendered");
  assert.ok(textContent(menu).includes("Bold"));
  assert.equal(hostMenuElement(), undefined);
  assert.deepEqual(mounted.warnings, []);
});

// ── Bullet 3: read-only Copy menu in the editor ────────────────────
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

function fakeParagraphTarget({ text, caret, nodeIndex = 0 }) {
  const target = new globalThis.Node();
  const textNode = { nodeType: 3, textContent: text };
  const host = {
    nodeType: 1,
    tagName: "P",
    textContent: text,
    ownerDocument: {
      caretRangeFromPoint: () => ({ startContainer: textNode, startOffset: caret }),
      createRange: () => ({
        setStart() {},
        setEnd() {},
        cloneContents: () => ({ textContent: text.slice(0, caret) }),
      }),
    },
    contains(node) { return node === host || node === textNode; },
    matches() { return false; },
    getAttribute(name) {
      return name === "data-docx-paragraph-node-index" ? String(nodeIndex) : null;
    },
  };
  target.closest = (selector) => {
    if (selector === "[data-docx-table-cell='true']") return null;
    return selector.includes("[data-docx-paragraph-host='true']") ? host : null;
  };
  return target;
}

function viewerRootNode(root) {
  return walk(root).find((node) => hasClass(node, "docx-viewer-root"));
}

async function mountEditor({ editable, ...props }) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({
        starterModel: starterModel("Hello world", "Second paragraph"),
      });
      return () =>
        vue.h(DocxEditor, {
          editor: controller,
          showToolbar: false,
          showThumbnails: false,
          editable,
          ...props,
        });
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, get controller() { return controller; } };
}

test("a read-only editor still opens a Copy-only context menu", async (t) => {
  const restoreDom = installEditorDomEnvironment();
  const mounted = await mountEditor({ editable: false });
  t.after(() => {
    mounted.app.unmount();
    restoreDom();
  });
  const copied = [];
  mounted.controller.copy = async () => {
    copied.push("copy");
  };

  const viewer = viewerRootNode(mounted.root);
  assert.ok(viewer, "the editor renders a viewer root that receives contextmenu");
  viewer.props.onContextmenu({
    button: 2,
    clientX: 160,
    clientY: 120,
    target: fakeParagraphTarget({ text: "Hello world", caret: 5 }),
    preventDefault() {},
  });
  await waitFor(
    () => findInBody("docx-context-menu-copy"),
    "the read-only Copy menu to open",
  );

  const menu = menuElement();
  assert.equal(menu.props["data-docx-context-menu-read-only"], "true");
  assert.deepEqual(
    walk(menu)
      .filter((node) => node.props?.role === "menuitem")
      .map((node) => textContent(node).trim()),
    ["📋Copy⌘C"],
    "read-only documents only offer Copy",
  );

  findByTestId(menu, "docx-context-menu-copy").props.onClick(fakeEvent());
  await waitFor(() => copied.length === 1, "Copy to run");
  assert.deepEqual(copied, ["copy"]);
  await vue.nextTick();
  assert.equal(menuElement(), undefined, "the menu closes after Copy");
  assert.deepEqual(mounted.warnings, []);
});

test("DocxEditor forwards the context menu renderer and its readOnly filter", async (t) => {
  const restoreDom = installEditorDomEnvironment();
  const seen = [];
  const copied = [];
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({
        starterModel: starterModel("Hello world", "Second paragraph"),
      });
      return () =>
        vue.h(DocxEditor, {
          editor: controller,
          showToolbar: false,
          showThumbnails: false,
          editable: false,
          renderContextMenu: (menu) => {
            seen.push(menu);
            return vue.h(
              "button",
              {
                "data-testid": "editor-custom-menu-copy",
                onClick: () => menu.runAction("copy"),
              },
              menu.actions.map((action) => action.id).join(","),
            );
          },
        });
    },
  });
  const mounted = await mount(Harness);
  t.after(() => {
    mounted.app.unmount();
    restoreDom();
  });
  controller.copy = async () => {
    copied.push("copy");
  };

  const viewer = viewerRootNode(mounted.root);
  viewer.props.onContextmenu({
    button: 2,
    clientX: 160,
    clientY: 120,
    target: fakeParagraphTarget({ text: "Hello world", caret: 5 }),
    preventDefault() {},
  });
  await waitFor(
    () => findByTestId(teleportTarget("body"), "editor-custom-menu-copy"),
    "the editor to render the host context menu",
  );

  const menu = seen.at(-1);
  assert.equal(menu.readOnly, true);
  assert.deepEqual(menu.actions.map((action) => action.id), ["copy"]);
  assert.equal(menuElement(), undefined, "the built-in menu is replaced");

  findByTestId(teleportTarget("body"), "editor-custom-menu-copy").props.onClick(fakeEvent());
  await waitFor(() => copied.length === 1, "the host menu's runAction('copy') to copy");
  assert.deepEqual(copied, ["copy"]);
  assert.deepEqual(mounted.warnings, []);
});

test("an editable editor keeps the full default menu", async (t) => {
  const restoreDom = installEditorDomEnvironment();
  const mounted = await mountEditor({ editable: true });
  t.after(() => {
    mounted.app.unmount();
    restoreDom();
  });

  const viewer = viewerRootNode(mounted.root);
  viewer.props.onContextmenu({
    button: 2,
    clientX: 160,
    clientY: 120,
    target: fakeParagraphTarget({ text: "Hello world", caret: 5 }),
    preventDefault() {},
  });
  await waitFor(menuElement, "the default menu to open");

  const menu = menuElement();
  assert.equal(menu.props["data-docx-context-menu-read-only"], undefined);
  assert.ok(
    walk(menu).some((node) => textContent(node).trim() === "Bold"),
    "editing actions stay available",
  );
  assert.deepEqual(mounted.warnings, []);
});

test("DocxEditor forwards custom card renderers to the page gutter", async (t) => {
  const restoreDom = installEditorDomEnvironment();
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({
        starterModel: revisionModel(),
        initialShowTrackedChanges: true,
      });
      return () =>
        vue.h(DocxEditor, {
          editor: controller,
          showToolbar: false,
          showThumbnails: false,
          editable: true,
          renderTrackedChangeCard: (card) =>
            vue.h("div", { "data-testid": "editor-custom-change-card" }, card.kindLabel),
        });
    },
  });
  const mounted = await mount(Harness);
  t.after(() => {
    mounted.app.unmount();
    restoreDom();
  });
  await waitFor(
    () => findByTestId(mounted.root, "editor-custom-change-card"),
    "the editor to render the host revision card",
  );

  const card = findByTestId(mounted.root, "editor-custom-change-card");
  assert.match(textContent(card), /Deleted/u);
  assert.deepEqual(mounted.warnings, []);
});
