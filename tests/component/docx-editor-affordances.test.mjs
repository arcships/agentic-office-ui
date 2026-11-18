import assert from "node:assert/strict";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  textContent,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

const {
  DocxEditor,
  DocxFormFieldDialog,
  DocxLinkHoverCard,
  DocxStylePreviewPopover,
  DocxTableHost,
  DocxToolbar,
  docxReorderNodeIndexFromElement,
  docxStylePreviewCss,
  parseDocxFormFieldLocationKey,
  resolveDocxReorderInsertionIndex,
  resolveDocxReorderTarget,
  resolveFormFieldAt,
  useDocxEditor,
  useDocxNodeReorder,
} = await importFromDemo("@arcships/vue-docx");
const { createBlankDocumentModel, paragraphText } = await importFromDemo("@arcships/docx-core");

// ── Fixtures ───────────────────────────────────────────────────────
function modelWithNodes(nodes) {
  return { ...createBlankDocumentModel(), nodes };
}

function paragraph(text, extra = {}) {
  return { type: "paragraph", children: [{ type: "text", text }], ...extra };
}

function linkedParagraph(text, link) {
  return {
    type: "paragraph",
    children: [{ type: "text", text, link }],
  };
}

function bookmarkParagraph(name, text) {
  return {
    type: "paragraph",
    sourceXml:
      `<w:p><w:bookmarkStart w:id="9" w:name="${name}"/>` +
      `<w:r><w:t>${text}</w:t></w:r><w:bookmarkEnd w:id="9"/></w:p>`,
    children: [{ type: "text", text }],
  };
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

function formFieldParagraph() {
  return {
    type: "paragraph",
    children: [
      {
        type: "form-field",
        fieldType: "dropdown",
        title: "Status",
        value: "open",
        options: [
          { displayText: "Open", value: "open" },
          { displayText: "Closed", value: "closed" },
        ],
      },
    ],
  };
}

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
  const originalRect = Object.getOwnPropertyDescriptor(Object.prototype, "getBoundingClientRect");
  const originalClientHeight = Object.getOwnPropertyDescriptor(Object.prototype, "clientHeight");
  const documentListeners = createListenerTarget();
  const windowListeners = createListenerTarget();

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
    querySelectorAll() { return []; }
  }

  const opened = [];
  globalThis.window = {
    addEventListener: windowListeners.addEventListener,
    removeEventListener: windowListeners.removeEventListener,
    getSelection() { return null; },
    open(url, name, features) {
      opened.push([url, name, features]);
    },
  };
  globalThis.document = {
    addEventListener: documentListeners.addEventListener,
    removeEventListener: documentListeners.removeEventListener,
    body: { append() {} },
    createElement(tagName) { return new FakeElement(tagName); },
    createTextNode(text) { return { textContent: String(text) }; },
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

  function restore() {
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
  }

  return { documentListeners, windowListeners, opened, restore };
}

/**
 * Unmounts the given apps and restores the fake DOM afterwards. The restore has
 * to run after unmount: components remove their document/window listeners there.
 */
function cleanupAfterTest(t, env, ...apps) {
  t.after(() => {
    for (const app of apps) app?.app?.unmount?.();
    env.restore();
  });
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
  return { ...mounted, get controller() { return controller; } };
}

async function mountEditor(model, props = {}) {
  let controller;
  let editorRef;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model });
      return () =>
        vue.h(DocxEditor, {
          ref: (instance) => {
            editorRef = instance;
          },
          editor: controller,
          editable: true,
          showToolbar: false,
          showThumbnails: false,
          ...props,
        });
    },
  });
  const mounted = await mount(Harness);
  return {
    ...mounted,
    get controller() {
      return controller;
    },
    get editor() {
      return editorRef;
    },
  };
}

function editorRoot(root) {
  return findByTestId(root, "docx-editor");
}

/** Types into a `v-model` input the way the DOM does: value + input event. */
function typeInto(node, value) {
  assert.ok(node, "input must render");
  node.value = value;
  node.dispatchEvent({ type: "input", target: node });
}

function classNames(node) {
  return String(node?.props?.class || "").split(/\s+/).filter(Boolean);
}

function hasClass(node, className) {
  return classNames(node).includes(className);
}

function tableRowHandles(root) {
  return walk(root).filter((node) => node.props?.["data-docx-table-row-handle"] !== undefined);
}

function fakeTextHost(nodeIndex) {
  return {
    matches: () => false,
    getAttribute(name) {
      return name === "data-docx-paragraph-node-index" ? String(nodeIndex) : null;
    },
  };
}

function fakeAnchor({ href, start, end, nodeIndex = 0 }) {
  const anchor = {
    closest(selector) {
      if (selector === "a[href]") return anchor;
      if (selector.includes("data-docx-paragraph-host")) return fakeTextHost(nodeIndex);
      return null;
    },
    getAttribute(name) {
      if (name === "href") return href;
      if (name === "data-docx-link-start") return start === undefined ? null : String(start);
      if (name === "data-docx-link-end") return end === undefined ? null : String(end);
      return null;
    },
    getBoundingClientRect() {
      return { left: 40, bottom: 80, top: 60, width: 120, height: 20 };
    },
  };
  return anchor;
}

function fakeReorderHost(nodeIndex, rect = { top: 10, bottom: 60, left: 20, width: 300 }) {
  const host = {
    getAttribute(name) {
      if (name === "data-docx-paragraph-node-index") return String(nodeIndex);
      return null;
    },
    closest(selector) {
      if (selector.includes("data-docx-paragraph-host")) return host;
      return null;
    },
    getBoundingClientRect() {
      return rect;
    },
  };
  return {
    host,
    // The editor asks for the hovered host with the node selector and for its
    // own handle with a class selector, so the fake element has to answer both.
    target: {
      closest(selector) {
        return selector.includes("data-docx-paragraph-host") ? host : null;
      },
    },
  };
}

// ── Form field configuration dialog ───────────────────────────────
test("updateFormField persists value, options and widget settings in one undo entry", async (t) => {
  const mounted = await mountController(modelWithNodes([formFieldParagraph()]));
  t.after(() => mounted.app.unmount());
  const { controller } = mounted;
  const location = { kind: "paragraph", nodeIndex: 0, childIndex: 0 };

  const saved = controller.updateFormField(location, {
    fieldType: "dropdown",
    title: "Ticket status",
    placeholder: "Pick one",
    required: true,
    value: "closed",
    options: [
      { displayText: "Open", value: "open" },
      { displayText: "Closed", value: "closed" },
      { displayText: "Escalated", value: "escalated" },
    ],
    widget: { name: "Status" },
  });
  assert.equal(saved, true);

  let field = controller.model.nodes[0].children[0];
  assert.equal(field.fieldType, "dropdown");
  assert.equal(field.title, "Ticket status");
  assert.equal(field.placeholder, "Pick one");
  assert.equal(field.required, true);
  assert.equal(field.value, "closed");
  assert.deepEqual(field.options.map((option) => option.displayText), [
    "Open",
    "Closed",
    "Escalated",
  ]);
  assert.equal(field.widget.name, "Status");

  // The dialog save is one history entry.
  assert.equal(controller.canUndo, true);
  controller.undo();
  await vue.nextTick();
  field = controller.model.nodes[0].children[0];
  assert.equal(field.value, "open");
  assert.equal(field.title, "Status");
  assert.equal(field.options.length, 2);
  assert.equal(field.required, undefined);

  // Changing the field type drops properties the new type cannot render.
  assert.equal(
    controller.updateFormField(location, { fieldType: "text", value: "free text" }),
    true
  );
  field = controller.model.nodes[0].children[0];
  assert.equal(field.fieldType, "text");
  assert.equal(field.value, "free text");
  assert.equal(field.options, undefined, "a text field keeps no dropdown options");

  assert.deepEqual(mounted.warnings, []);
});

test("updateFormField refuses values that violate the field and unknown targets", async (t) => {
  const mounted = await mountController(modelWithNodes([formFieldParagraph()]));
  t.after(() => mounted.app.unmount());
  const { controller } = mounted;
  const location = { kind: "paragraph", nodeIndex: 0, childIndex: 0 };

  assert.equal(
    controller.updateFormField(location, { value: "not-an-option" }),
    false,
    "a dropdown value outside the option list must be refused"
  );
  assert.equal(controller.model.nodes[0].children[0].value, "open");

  assert.equal(
    controller.updateFormField({ kind: "paragraph", nodeIndex: 4, childIndex: 0 }, { value: "x" }),
    false
  );
  assert.deepEqual(mounted.warnings, []);
});

test("form field helpers resolve the node a location points at", () => {
  const model = modelWithNodes([formFieldParagraph()]);
  assert.deepEqual(parseDocxFormFieldLocationKey("p:0:0"), {
    kind: "paragraph",
    nodeIndex: 0,
    childIndex: 0,
  });
  assert.deepEqual(parseDocxFormFieldLocationKey("t:1:2:3:4:5"), {
    kind: "table-cell",
    tableIndex: 1,
    rowIndex: 2,
    cellIndex: 3,
    paragraphIndex: 4,
    childIndex: 5,
  });
  assert.equal(parseDocxFormFieldLocationKey("p:0"), undefined);
  assert.equal(parseDocxFormFieldLocationKey("nonsense"), undefined);
  assert.equal(parseDocxFormFieldLocationKey(""), undefined);

  const field = resolveFormFieldAt(model, { kind: "paragraph", nodeIndex: 0, childIndex: 0 });
  assert.equal(field?.title, "Status");
  assert.equal(resolveFormFieldAt(model, { kind: "paragraph", nodeIndex: 0, childIndex: 3 }), undefined);
});

test("DocxFormFieldDialog edits value and dropdown options and saves them", async (t) => {
  const mounted = await mountController(modelWithNodes([formFieldParagraph()]));
  t.after(() => mounted.app.unmount());
  const { controller } = mounted;
  const location = { kind: "paragraph", nodeIndex: 0, childIndex: 0 };
  controller.selectFormField(location);

  const closes = [];
  const Harness = vue.defineComponent({
    setup() {
      return () =>
        vue.h(DocxFormFieldDialog, {
          field: controller.model.nodes[0].children[0],
          location,
          controller,
          onClose: () => closes.push("close"),
        });
    },
  });
  const dialog = await mount(Harness);
  t.after(() => dialog.app.unmount());

  const title = findByTestId(dialog.root, "docx-form-field-dialog-title");
  assert.ok(title, "the dialog exposes a name field");
  assert.equal(title.value, "Status");

  const optionInput = findByTestId(dialog.root, "docx-form-field-dialog-option-1");
  assert.ok(optionInput, "existing dropdown options are editable");
  assert.equal(optionInput.value, "Closed");

  // Add a third option and rename the second one.
  findByTestId(dialog.root, "docx-form-field-dialog-add-option").props.onClick();
  await vue.nextTick();
  const addedOption = findByTestId(dialog.root, "docx-form-field-dialog-option-2");
  assert.ok(addedOption, "add option appends a row");
  typeInto(addedOption, "Escalated");
  await vue.nextTick();

  // A dropdown's value is chosen from its options, not typed: the dialog
  // renders a select so the saved value always exists in the option list.
  const valueSelect = findByTestId(dialog.root, "docx-form-field-dialog-value");
  assert.equal(valueSelect.type, "select");
  assert.deepEqual(
    walk(valueSelect)
      .filter((node) => node.type === "option")
      .map((node) => node.props.value),
    ["", "open", "closed", "Escalated"]
  );
  // A native select commits through its `v-model` assignment.
  valueSelect.props["onUpdate:modelValue"]("Escalated");
  await vue.nextTick();

  findByTestId(dialog.root, "docx-form-field-dialog-save").props.onClick();
  await vue.nextTick();

  const field = controller.model.nodes[0].children[0];
  assert.equal(field.value, "Escalated");
  assert.deepEqual(field.options.map((option) => option.displayText), [
    "Open",
    "Closed",
    "Escalated",
  ]);
  assert.deepEqual(closes, ["close"]);
  assert.deepEqual(mounted.warnings, []);
});

test("DocxEditor opens the form-field dialog on a double click and closes it again", async (t) => {
  const env = installDomEnvironment();
  const mounted = await mountEditor(modelWithNodes([formFieldParagraph()]));
  cleanupAfterTest(t, env, mounted);

  const root = editorRoot(mounted.root);
  assert.ok(root);
  let prevented = 0;
  const target = {
    closest(selector) {
      if (selector === "[data-docx-form-field-location]") {
        return { getAttribute: () => "p:0:0" };
      }
      return null;
    },
  };
  root.props.onDblclick({ target, preventDefault: () => { prevented += 1; } });
  await waitFor(
    () => findByTestId(mounted.root, "docx-form-field-dialog"),
    "the form field dialog to open"
  );
  assert.equal(prevented, 1, "the double click is consumed by the dialog");
  assert.deepEqual(
    mounted.controller.selectedFormField?.location,
    { kind: "paragraph", nodeIndex: 0, childIndex: 0 }
  );

  findByTestId(mounted.root, "docx-form-field-dialog-cancel").props.onClick();
  await vue.nextTick();
  assert.equal(findByTestId(mounted.root, "docx-form-field-dialog"), undefined);
  assert.deepEqual(mounted.warnings, []);
});

test("DocxEditor leaves the form-field dialog closed when the affordance is disabled", async (t) => {
  const env = installDomEnvironment();
  const mounted = await mountEditor(modelWithNodes([formFieldParagraph()]), {
    showFormFieldDialog: false,
  });
  cleanupAfterTest(t, env, mounted);

  const target = {
    closest: () => ({ getAttribute: () => "p:0:0" }),
  };
  editorRoot(mounted.root).props.onDblclick({ target, preventDefault() {} });
  await vue.nextTick();
  assert.equal(findByTestId(mounted.root, "docx-form-field-dialog"), undefined);
});

// ── Link hover card ───────────────────────────────────────────────
test("DocxLinkHoverCard exposes open/edit/remove and reports an empty address as remove", async (t) => {
  const events = [];
  const mounted = await mount(DocxLinkHoverCard, {
    href: "https://example.com/page",
    x: 12,
    y: 34,
    onOpen: (href) => events.push(["open", href]),
    onEdit: (href) => events.push(["edit", href]),
    onRemove: (href) => events.push(["remove", href]),
    onClose: () => events.push(["close"]),
  });
  t.after(() => mounted.app.unmount());

  const card = findByTestId(mounted.root, "docx-link-card");
  assert.ok(card);
  assert.equal(card.props.role, "dialog");
  assert.equal(card.props.style.position, "fixed");
  assert.equal(card.props.style.left, "12px");
  assert.equal(card.props.style.top, "34px");
  assert.equal(findByTestId(mounted.root, "docx-link-card-url").props.title, "https://example.com/page");

  findByTestId(mounted.root, "docx-link-card-open").props.onClick();
  assert.deepEqual(events.at(-1), ["open", "https://example.com/page"]);

  findByTestId(mounted.root, "docx-link-card-edit").props.onClick();
  await vue.nextTick();
  const input = findByTestId(mounted.root, "docx-link-card-input");
  assert.ok(input, "edit reveals the address field");
  assert.equal(input.value, "https://example.com/page");
  typeInto(input, "https://example.com/next");
  await vue.nextTick();
  findByTestId(mounted.root, "docx-link-card-save").props.onClick();
  assert.deepEqual(events.at(-1), ["edit", "https://example.com/next"]);
  await vue.nextTick();

  findByTestId(mounted.root, "docx-link-card-edit").props.onClick();
  await vue.nextTick();
  typeInto(findByTestId(mounted.root, "docx-link-card-input"), "");
  await vue.nextTick();
  findByTestId(mounted.root, "docx-link-card-save").props.onClick();
  assert.deepEqual(events.at(-1), ["remove", "https://example.com/page"]);
  assert.deepEqual(mounted.warnings, []);
});

test("DocxLinkHoverCard hides editing actions on links without an editable range", async (t) => {
  const mounted = await mount(DocxLinkHoverCard, {
    href: "#bookmark",
    editable: false,
  });
  t.after(() => mounted.app.unmount());
  assert.ok(findByTestId(mounted.root, "docx-link-card-open"));
  assert.equal(findByTestId(mounted.root, "docx-link-card-edit"), undefined);
  assert.equal(findByTestId(mounted.root, "docx-link-card-remove"), undefined);
  assert.equal(
    textContent(findByTestId(mounted.root, "docx-link-card-url")),
    "Links to a place in this document"
  );
});

test("hovering a link in DocxEditor shows its card and remove clears the model link", async (t) => {
  const env = installDomEnvironment();
  const mounted = await mountEditor(
    modelWithNodes([linkedParagraph("Example", "https://example.com")])
  );
  cleanupAfterTest(t, env, mounted);
  const root = editorRoot(mounted.root);

  root.props.onPointerover({
    target: fakeAnchor({ href: "https://example.com", start: 0, end: 7 }),
  });
  await waitFor(
    () => findByTestId(mounted.root, "docx-link-card"),
    "the link card to render"
  );

  findByTestId(mounted.root, "docx-link-card-remove").props.onClick();
  await vue.nextTick();

  let run = mounted.controller.model.nodes[0].children[0];
  assert.equal(run.link, undefined, "remove clears the link of exactly the hovered run");
  assert.equal(paragraphText(mounted.controller.model.nodes[0]), "Example");
  assert.ok(mounted.controller.canUndo);

  mounted.controller.undo();
  await vue.nextTick();
  run = mounted.controller.model.nodes[0].children[0];
  assert.equal(run.link, "https://example.com");
  assert.deepEqual(mounted.warnings, []);
});

test("the link card opens external links and defers internal bookmarks to the host", async (t) => {
  const env = installDomEnvironment();
  const openedLinks = [];
  const mounted = await mountEditor(
    modelWithNodes([linkedParagraph("Example", "https://example.com")]),
    { onLinkOpen: (href) => openedLinks.push(href) }
  );
  cleanupAfterTest(t, env, mounted);
  const root = editorRoot(mounted.root);

  root.props.onPointerover({
    target: fakeAnchor({ href: "https://example.com", start: 0, end: 7 }),
  });
  await waitFor(
    () => findByTestId(mounted.root, "docx-link-card-open"),
    "the open action"
  );
  findByTestId(mounted.root, "docx-link-card-open").props.onClick();
  await vue.nextTick();
  assert.deepEqual(env.opened, [["https://example.com", "_blank", "noopener,noreferrer"]]);
  assert.deepEqual(openedLinks, ["https://example.com"]);

  root.props.onPointerover({
    target: fakeAnchor({ href: "#introduction", start: 0, end: 7 }),
  });
  await waitFor(
    () => findByTestId(mounted.root, "docx-link-card-open"),
    "the card for an internal link"
  );
  findByTestId(mounted.root, "docx-link-card-open").props.onClick();
  await vue.nextTick();
  assert.deepEqual(openedLinks, ["https://example.com", "#introduction"]);
  assert.equal(env.opened.length, 1, "an internal bookmark is not opened as a URL");
  assert.deepEqual(mounted.warnings, []);
});

test("the link card's internal open action navigates the document and still notifies the host", async (t) => {
  const env = installDomEnvironment();
  const openedLinks = [];
  const mounted = await mountEditor(
    modelWithNodes([
      linkedParagraph("Jump", "#Target_Heading"),
      bookmarkParagraph("Target_Heading", "Destination"),
    ]),
    { onLinkOpen: (href) => openedLinks.push(href) }
  );
  cleanupAfterTest(t, env, mounted);
  const root = editorRoot(mounted.root);

  root.props.onPointerover({
    target: fakeAnchor({ href: "#Target_Heading", start: 0, end: 4 }),
  });
  await waitFor(
    () => findByTestId(mounted.root, "docx-link-card-open"),
    "the card for an internal link"
  );
  findByTestId(mounted.root, "docx-link-card-open").props.onClick();

  await waitFor(
    () => mounted.controller.selection.kind === "paragraph" &&
      mounted.controller.selection.nodeIndex === 1,
    "the bookmark target to take the caret"
  );
  assert.deepEqual(openedLinks, ["#Target_Heading"]);
  assert.equal(env.opened.length, 0, "the bookmark is not opened as an external URL");

  const revealed = await mounted.editor.revealBookmark("Target_Heading");
  assert.equal(revealed.status, "revealed");
  assert.equal(revealed.nodeIndex, 1);
  assert.equal(revealed.text, "Destination");
  assert.deepEqual(mounted.warnings, []);
});

test("editing a hovered link rewrites its address through setLink", async (t) => {
  const env = installDomEnvironment();
  const mounted = await mountEditor(
    modelWithNodes([linkedParagraph("Example", "https://example.com")])
  );
  cleanupAfterTest(t, env, mounted);
  const root = editorRoot(mounted.root);

  root.props.onPointerover({
    target: fakeAnchor({ href: "https://example.com", start: 0, end: 7 }),
  });
  await waitFor(
    () => findByTestId(mounted.root, "docx-link-card-edit"),
    "the link card actions"
  );
  findByTestId(mounted.root, "docx-link-card-edit").props.onClick();
  await vue.nextTick();
  typeInto(
    findByTestId(mounted.root, "docx-link-card-input"),
    "https://example.org/updated"
  );
  await vue.nextTick();
  findByTestId(mounted.root, "docx-link-card-save").props.onClick();
  await vue.nextTick();

  assert.equal(
    mounted.controller.model.nodes[0].children[0].link,
    "https://example.org/updated"
  );
  assert.deepEqual(mounted.warnings, []);
});

// ── Table row resize ──────────────────────────────────────────────
test("DocxTableHost renders a handle per row and commits a dragged row height", async (t) => {
  const { documentListeners, restore } = installDomEnvironment();
  const mounted = await mountController(modelWithNodes([tableNode(3)]));
  const { controller } = mounted;

  const Harness = vue.defineComponent({
    setup() {
      return () =>
        vue.h(DocxTableHost, {
          table: controller.model.nodes[0],
          tableIndex: 0,
          editable: true,
          controller,
          maxAvailableWidthPx: 600,
        });
    },
  });
  const host = await mount(Harness);
  t.after(() => {
    host.app.unmount();
    mounted.app.unmount();
    restore();
  });
  const tableHostEl = walk(host.root).find(
    (node) => node.props?.["data-docx-table-host"] !== undefined
  );
  assert.ok(tableHostEl, "the table host renders");
  assert.equal(
    tableRowHandles(host.root).length,
    0,
    "row handles stay hidden until the table is hovered or active"
  );
  tableHostEl.props.onPointerenter({});
  await vue.nextTick();
  assert.equal(tableRowHandles(host.root).length, 3, "hovering the table shows its row handles");

  controller.selectTableCell(0, 0, 0);
  await vue.nextTick();

  const handles = tableRowHandles(host.root);
  assert.equal(handles.length, 3, "one handle per rendered row");
  assert.ok(
    walk(host.root).some((node) => hasClass(node, "docx-table-row-handles")),
    "the row handle layer renders"
  );

  handles[0].props.onMousedown({ preventDefault() {}, clientY: 100 });
  documentListeners.dispatch("mousemove", { clientY: 140 });
  documentListeners.dispatch("mouseup", { clientY: 140 });
  await vue.nextTick();

  const row = controller.model.nodes[0].rows[0];
  // 600 twips = 40px; dragging 40px down gives an 80px minimum height.
  assert.equal(row.style.heightTwips, 1200);
  assert.equal(row.style.heightRule, "atLeast");
  assert.equal(controller.canUndo, true);

  controller.undo();
  await vue.nextTick();
  assert.equal(controller.model.nodes[0].rows[0].style.heightTwips, 600);
  assert.deepEqual(mounted.warnings, []);
});

test("setTableRowHeight refuses non-positive heights and missing rows", async (t) => {
  const mounted = await mountController(modelWithNodes([tableNode(1)]));
  t.after(() => mounted.app.unmount());
  const { controller } = mounted;

  controller.setTableRowHeight(0, 0, 0);
  assert.equal(controller.model.nodes[0].rows[0].style.heightTwips, 600);
  controller.setTableRowHeight(0, 4, 100);
  assert.equal(controller.model.nodes[0].rows.length, 1);
  assert.deepEqual(mounted.warnings, []);
});

// ── Paragraph style preview ───────────────────────────────────────
test("docxStylePreviewCss maps a style's run and paragraph properties", () => {
  const styles = createBlankDocumentModel().metadata.paragraphStyles;
  const heading = styles.find((style) => style.id === "Heading1");
  const preview = docxStylePreviewCss(heading, "light");
  assert.equal(preview.text.fontWeight, "700");
  assert.match(String(preview.text.fontFamily), /Calibri Light/);
  assert.ok(String(preview.text.fontSize).endsWith("px"));
  assert.deepEqual(docxStylePreviewCss(undefined), { text: {}, paragraph: {} });

  const indented = docxStylePreviewCss(
    {
      id: "Quote",
      name: "Quote",
      align: "center",
      spacing: { beforeTwips: 240, afterTwips: 120 },
      indent: { leftTwips: 720 },
    },
    "light"
  );
  assert.equal(indented.paragraph.textAlign, "center");
  assert.equal(indented.paragraph.marginTop, "16px");
  assert.equal(indented.paragraph.marginBottom, "8px");
  assert.equal(indented.paragraph.marginLeft, "48px");
});

test("hovering a style in the toolbar previews it and clicking applies it", async (t) => {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({
        starterModel: modelWithNodes([paragraph("styled paragraph")]),
      });
      return () => vue.h(DocxToolbar, { controller, currentPage: 1, totalPages: 1 });
    },
  });
  const mounted = await mount(Harness);
  t.after(() => mounted.app.unmount());
  controller.selectParagraph(0);
  await vue.nextTick();

  assert.equal(findByTestId(mounted.root, "docx-style-preview"), undefined);

  const picker = findByTestId(mounted.root, "editor-paragraph-style-picker");
  assert.ok(picker, "the style picker wrapper renders");
  picker.props.onPointerenter();
  await waitFor(
    () => findByTestId(mounted.root, "docx-style-preview"),
    "the style preview to open"
  );

  const sample = findByTestId(mounted.root, "docx-style-preview-sample");
  assert.ok(sample, "the preview renders a sample paragraph");

  const headingItem = findByTestId(mounted.root, "docx-style-preview-item-Heading1");
  assert.ok(headingItem, "available styles are listed");
  headingItem.props.onPointerenter();
  await vue.nextTick();

  assert.equal(
    textContent(findByTestId(mounted.root, "docx-style-preview-name")),
    "Heading 1"
  );
  const styledSample = findByTestId(mounted.root, "docx-style-preview-sample");
  assert.equal(styledSample.props.style.fontWeight, "700");
  assert.match(String(styledSample.props.style.fontFamily), /Calibri Light/);

  headingItem.props.onClick();
  await vue.nextTick();
  assert.equal(controller.selectedParagraphStyleId, "Heading1");
  assert.equal(
    controller.model.nodes[0].style?.styleId,
    "Heading1",
    "clicking a preview applies the style"
  );
  assert.equal(findByTestId(mounted.root, "docx-style-preview"), undefined);
  assert.deepEqual(mounted.warnings, []);
});

test("DocxStylePreviewPopover can be rendered on its own", async (t) => {
  const styles = createBlankDocumentModel().metadata.paragraphStyles;
  const selects = [];
  const mounted = await mount(DocxStylePreviewPopover, {
    styles,
    previewStyleId: "Heading2",
    x: 10,
    y: 20,
    onSelect: (styleId) => selects.push(styleId),
  });
  t.after(() => mounted.app.unmount());

  const popover = findByTestId(mounted.root, "docx-style-preview");
  assert.ok(popover);
  assert.equal(popover.props.style.left, "10px");
  assert.equal(popover.props.style.top, "20px");
  assert.equal(
    textContent(findByTestId(mounted.root, "docx-style-preview-name")),
    "Heading 2"
  );
  findByTestId(mounted.root, "docx-style-preview-item-normal").props.onClick();
  assert.deepEqual(selects, [undefined]);
});

test("the toolbar hides the style preview when disabled", async (t) => {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: modelWithNodes([paragraph("plain")]) });
      return () =>
        vue.h(DocxToolbar, {
          controller,
          currentPage: 1,
          totalPages: 1,
          showStylePreview: false,
        });
    },
  });
  const mounted = await mount(Harness);
  t.after(() => mounted.app.unmount());
  findByTestId(mounted.root, "editor-paragraph-style-picker").props.onPointerenter();
  await vue.nextTick();
  assert.equal(findByTestId(mounted.root, "docx-style-preview"), undefined);
});

// ── Drag-to-reorder ──────────────────────────────────────────────
test("reorder helpers resolve the hovered node and a no-op drop", () => {
  const frames = [
    { nodeIndex: 0, top: 0, bottom: 40 },
    { nodeIndex: 1, top: 40, bottom: 90 },
    { nodeIndex: 2, top: 90, bottom: 120 },
  ];
  assert.equal(resolveDocxReorderInsertionIndex(-10, frames), 0);
  assert.equal(resolveDocxReorderInsertionIndex(10, frames), 0);
  assert.equal(resolveDocxReorderInsertionIndex(50, frames), 1);
  assert.equal(resolveDocxReorderInsertionIndex(200, frames), 3);
  assert.equal(resolveDocxReorderInsertionIndex(10, []), undefined);

  assert.equal(resolveDocxReorderTarget(1, 0), 0);
  assert.equal(resolveDocxReorderTarget(1, 3), 2);
  assert.equal(resolveDocxReorderTarget(1, 1), undefined, "before itself is a no-op");
  assert.equal(resolveDocxReorderTarget(1, 2), undefined, "after itself is a no-op");
  assert.equal(resolveDocxReorderTarget(-1, 2), undefined);

  const { target } = fakeReorderHost(2);
  assert.equal(docxReorderNodeIndexFromElement(target), 2);
  assert.equal(
    docxReorderNodeIndexFromElement({ closest: () => null }),
    undefined,
    "chrome outside a body node is not draggable"
  );
});

test("moveBodyNode reorders paragraphs and tables as one history entry", async (t) => {
  const mounted = await mountController(
    modelWithNodes([paragraph("one"), paragraph("two"), tableNode(1), paragraph("three")])
  );
  t.after(() => mounted.app.unmount());
  const { controller } = mounted;

  assert.equal(controller.moveBodyNode(0, 3), true);
  assert.deepEqual(
    controller.model.nodes.map((node) =>
      node.type === "paragraph" ? paragraphText(node) : "table"
    ),
    ["two", "table", "one", "three"]
  );
  assert.equal(controller.canUndo, true);

  controller.undo();
  await vue.nextTick();
  assert.deepEqual(
    controller.model.nodes.map((node) =>
      node.type === "paragraph" ? paragraphText(node) : "table"
    ),
    ["one", "two", "table", "three"]
  );

  // Dropping a node onto its own slot changes nothing and adds no history.
  assert.equal(controller.canUndo, false, "undo is exhausted after the undo above");
  assert.equal(controller.moveBodyNode(0, 1), false);
  assert.equal(controller.moveBodyNode(0, 0), false);
  assert.equal(controller.canUndo, false, "a no-op drop must not add a history entry");
  assert.deepEqual(mounted.warnings, []);
});

test("useDocxNodeReorder drags a node to the drop slot and supports keyboard moves", async (t) => {
  const mounted = await mountController(
    modelWithNodes([paragraph("one"), paragraph("two"), paragraph("three")])
  );
  t.after(() => mounted.app.unmount());
  const { controller } = mounted;

  const frames = [
    { nodeIndex: 0, top: 0, bottom: 40 },
    { nodeIndex: 1, top: 40, bottom: 80 },
    { nodeIndex: 2, top: 80, bottom: 120 },
  ];
  const reorder = useDocxNodeReorder(controller, {
    collectFrames: () => frames,
    rectOf: () => ({ top: 0, bottom: 40, left: 10, width: 200 }),
  });

  const { target } = fakeReorderHost(0);
  reorder.onHover({ clientX: 20, clientY: 10, target });
  assert.equal(reorder.hover.value?.nodeIndex, 0);
  assert.equal(reorder.hover.value?.left, 10);

  reorder.startDrag(0);
  assert.equal(reorder.dragging.value, true);
  reorder.updateDrag({ clientY: 100 });
  assert.equal(reorder.dropIndex.value, 3);
  assert.equal(reorder.dropTop.value, 120);
  assert.equal(reorder.finishDrag(), true);
  await vue.nextTick();
  assert.deepEqual(
    controller.model.nodes.map((node) => paragraphText(node)),
    ["two", "three", "one"]
  );
  assert.equal(reorder.dragging.value, false);

  assert.equal(reorder.moveBy(2, -1), true);
  await vue.nextTick();
  assert.deepEqual(
    controller.model.nodes.map((node) => paragraphText(node)),
    ["two", "one", "three"]
  );
  assert.equal(reorder.moveBy(0, -1), false, "moving the first node up is refused");
  assert.deepEqual(mounted.warnings, []);
});

test("DocxEditor drags the hovered node handle and reorders with one history entry", async (t) => {
  const { documentListeners, restore } = installDomEnvironment();
  const mounted = await mountEditor(
    modelWithNodes([paragraph("one"), paragraph("two"), paragraph("three")])
  );
  t.after(() => {
    mounted.app.unmount();
    restore();
  });
  const root = editorRoot(mounted.root);

  // The drag resolves the node frames through `document.querySelectorAll`, so
  // the fake document reports one host per body node.
  const frames = [
    { nodeIndex: 0, top: 10, bottom: 50, left: 20, width: 300 },
    { nodeIndex: 1, top: 50, bottom: 90, left: 20, width: 300 },
    { nodeIndex: 2, top: 90, bottom: 130, left: 20, width: 300 },
  ];
  const hosts = frames.map((frame) => fakeReorderHost(frame.nodeIndex, frame).host);
  globalThis.document.querySelectorAll = () => hosts;

  const { target } = fakeReorderHost(0, frames[0]);
  root.props.onPointermove({ target });
  await waitFor(
    () => findByTestId(mounted.root, "docx-node-reorder-handle"),
    "the reorder handle to appear"
  );

  const handle = findByTestId(mounted.root, "docx-node-reorder-handle");
  assert.equal(handle.props["data-docx-node-index"], 0);
  assert.equal(handle.props.style.position, "fixed");

  // Moving from the paragraph onto the handle keeps the handle (and the node it
  // will move) alive.
  root.props.onPointermove({
    target: { closest: (selector) => (selector === ".docx-node-reorder-handle" ? {} : null) },
  });
  await vue.nextTick();
  assert.ok(
    findByTestId(mounted.root, "docx-node-reorder-handle"),
    "the handle survives the pointer leaving the paragraph for the handle"
  );

  handle.props.onMousedown({ preventDefault() {}, stopPropagation() {} });

  // Simulate the drag: node 0 moves past nodes 1 and 2.
  documentListeners.dispatch("mousemove", { clientY: 400 });
  await vue.nextTick();
  assert.ok(
    findByTestId(mounted.root, "docx-node-drop-indicator"),
    "the drop indicator follows the drag"
  );
  documentListeners.dispatch("mouseup", {});
  await vue.nextTick();

  assert.deepEqual(
    mounted.controller.model.nodes.map((node) => paragraphText(node)),
    ["two", "three", "one"]
  );
  assert.equal(mounted.controller.canUndo, true);
  assert.equal(findByTestId(mounted.root, "docx-node-reorder-handle"), undefined);
  assert.deepEqual(mounted.warnings, []);
});
