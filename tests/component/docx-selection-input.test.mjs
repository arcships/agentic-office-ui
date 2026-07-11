import assert from "node:assert/strict";
import test from "node:test";

import {
  importFromDemo,
  mount,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

const { DocxEditor, useDocxEditor } = await importFromDemo("@arcships/vue-docx");
const { createBlankDocumentModel } = await importFromDemo("@arcships/docx-core");

function paragraphText(node) {
  if (!node || node.type !== "paragraph") return "";
  return node.children
    .filter((child) => child.type === "text")
    .map((child) => child.text)
    .join("");
}

function starterModel(texts) {
  return {
    ...createBlankDocumentModel(),
    nodes: texts.map((text) => ({
      type: "paragraph",
      children: [{ type: "text", text }],
    })),
  };
}

function createListenerTarget() {
  const listeners = new Map();
  return {
    addEventListener(type, listener) {
      const entries = listeners.get(type) ?? new Set();
      entries.add(listener);
      listeners.set(type, entries);
    },
    removeEventListener(type, listener) {
      listeners.get(type)?.delete(listener);
    },
    dispatch(type, event = { type }) {
      for (const listener of listeners.get(type) ?? []) listener(event);
    },
  };
}

function classNames(node) {
  const value = node?.props?.class;
  if (typeof value === "string") return value.split(/\s+/).filter(Boolean);
  if (Array.isArray(value)) return value.flatMap((entry) => String(entry).split(/\s+/));
  if (value && typeof value === "object") {
    return Object.entries(value).filter(([, enabled]) => enabled).map(([name]) => name);
  }
  return [];
}

function attributeValue(node, name) {
  const value = node?.props?.[name];
  return value === undefined || value === null ? null : String(value);
}

function matchesSelector(node, selector) {
  return selector.split(",").some((rawPart) => {
    const part = rawPart.trim();
    const classes = [...part.matchAll(/\.([a-zA-Z0-9_-]+)/g)].map((match) => match[1]);
    if (classes.some((name) => !classNames(node).includes(name))) return false;

    const attributes = [...part.matchAll(/\[([^=\]]+)(?:=["']?([^"'\]]+)["']?)?\]/g)];
    for (const [, name, expected] of attributes) {
      const actual = attributeValue(node, name);
      if (actual === null || (expected !== undefined && actual !== expected)) return false;
    }
    return classes.length > 0 || attributes.length > 0;
  });
}

function containsNode(owner, candidate) {
  let current = candidate;
  while (current) {
    if (current === owner) return true;
    current = current.parentElement ?? current.parentNode ?? current.parent ?? null;
  }
  return false;
}

function installDomEnvironment(t) {
  const originalWindow = globalThis.window;
  const originalDocument = globalThis.document;
  const originalResizeObserver = globalThis.ResizeObserver;
  const originalRect = Object.getOwnPropertyDescriptor(Object.prototype, "getBoundingClientRect");
  const originalClientHeight = Object.getOwnPropertyDescriptor(Object.prototype, "clientHeight");
  const windowListeners = createListenerTarget();
  const documentListeners = createListenerTarget();
  let selectedRange;
  let cleanupBeforeRestore;
  let collapseSelectionOnFocus = false;

  class StaticText {
    constructor(text) {
      this.textContent = String(text);
    }
    get outerHTML() { return this.textContent; }
  }

  class StaticElement {
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

  function nearestTextHost(node) {
    let current = node;
    while (current) {
      if (
        attributeValue(current, "data-docx-paragraph-host") === "true" ||
        attributeValue(current, "data-docx-table-cell-paragraph-host") === "true"
      ) return current;
      current = current.parentElement ?? current.parentNode ?? current.parent ?? null;
    }
    return null;
  }

  function textLeaves(host) {
    const result = [];
    function visit(node, withinNumberingLabel = false) {
      const numbered = withinNumberingLabel || matchesSelector(
        node,
        "[data-docx-numbering-label='true']",
      );
      if (node?.nodeType === 3) {
        result.push({ node, text: node.textContent ?? "", numbered });
        return;
      }
      for (const child of node?.childNodes ?? []) visit(child, numbered);
    }
    visit(host);
    return result;
  }

  function subtreeTextLength(node) {
    if (node?.nodeType === 3) return node.textContent?.length ?? 0;
    return [...(node?.childNodes ?? [])]
      .reduce((sum, child) => sum + subtreeTextLength(child), 0);
  }

  function boundaryOffset(host, node, offset) {
    if (node?.nodeType === 3) {
      let total = 0;
      for (const leaf of textLeaves(host)) {
        if (leaf.node === node) return total + Math.max(0, Math.min(offset, leaf.text.length));
        total += leaf.text.length;
      }
      return total;
    }

    let total = 0;
    let found = false;
    function visit(current) {
      if (found) return;
      if (current === node) {
        const children = [...(current.childNodes ?? [])];
        total += children.slice(0, Math.max(0, Math.min(offset, children.length)))
          .reduce((sum, child) => sum + subtreeTextLength(child), 0);
        found = true;
        return;
      }
      for (const child of current?.childNodes ?? []) {
        if (containsNode(child, node)) {
          visit(child);
          return;
        }
        total += subtreeTextLength(child);
      }
    }
    visit(host);
    return total;
  }

  function fragmentForRange(range) {
    const host = nearestTextHost(range.startContainer) ?? nearestTextHost(range.endContainer);
    if (!host) {
      return { textContent: "", querySelectorAll: () => [] };
    }
    const start = boundaryOffset(host, range.startContainer, range.startOffset);
    const end = boundaryOffset(host, range.endContainer, range.endOffset);
    const pieces = [];
    let traversed = 0;
    for (const leaf of textLeaves(host)) {
      const leafStart = traversed;
      const leafEnd = traversed + leaf.text.length;
      const from = Math.max(start, leafStart);
      const to = Math.min(end, leafEnd);
      if (to > from) {
        pieces.push({
          text: leaf.text.slice(from - leafStart, to - leafStart),
          numbered: leaf.numbered,
          removed: false,
        });
      }
      traversed = leafEnd;
    }
    return {
      get textContent() {
        return pieces.filter((piece) => !piece.removed).map((piece) => piece.text).join("");
      },
      querySelectorAll(selector) {
        if (selector === "[data-docx-numbering-label='true']") {
          return pieces.filter((piece) => piece.numbered).map((piece) => ({
            remove() { piece.removed = true; },
          }));
        }
        return [];
      },
    };
  }

  function createRange() {
    return {
      startContainer: null,
      startOffset: 0,
      endContainer: null,
      endOffset: 0,
      setStart(node, offset) { this.startContainer = node; this.startOffset = offset; },
      setEnd(node, offset) { this.endContainer = node; this.endOffset = offset; },
      cloneContents() { return fragmentForRange(this); },
      selectNodeContents(node) {
        this.setStart(node, 0);
        this.setEnd(node, node.childNodes?.length ?? 0);
      },
      collapse(toStart) {
        if (toStart) this.setEnd(this.startContainer, this.startOffset);
        else this.setStart(this.endContainer, this.endOffset);
      },
    };
  }

  const selection = {
    get rangeCount() { return selectedRange ? 1 : 0; },
    get isCollapsed() {
      return !selectedRange || (
        selectedRange.startContainer === selectedRange.endContainer &&
        selectedRange.startOffset === selectedRange.endOffset
      );
    },
    getRangeAt() {
      if (!selectedRange) throw new RangeError("No selection range");
      return selectedRange;
    },
    removeAllRanges() { selectedRange = undefined; },
    addRange(range) { selectedRange = range; },
  };

  const documentObject = {
    addEventListener: documentListeners.addEventListener,
    removeEventListener: documentListeners.removeEventListener,
    body: { append() {} },
    activeElement: null,
    createElement(tagName) { return new StaticElement(tagName); },
    createTextNode(text) { return new StaticText(text); },
    createRange,
    getSelection() { return selection; },
    focusElement(node) {
      if (this.activeElement === node) return;
      this.activeElement = node;
      if (!collapseSelectionOnFocus) return;
      const host = nearestTextHost(node);
      if (host) {
        const collapsed = createRange();
        collapsed.setStart(host, 0);
        collapsed.setEnd(host, 0);
        selectedRange = collapsed;
      }
      node.props?.onFocus?.();
    },
    execCommand() { return false; },
  };
  globalThis.document = documentObject;
  globalThis.window = {
    addEventListener: windowListeners.addEventListener,
    removeEventListener: windowListeners.removeEventListener,
    getSelection() { return selection; },
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
    cleanupBeforeRestore?.();
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

  return {
    document: documentObject,
    dispatchDocument: documentListeners.dispatch,
    setCleanup(callback) { cleanupBeforeRestore = callback; },
    setCollapseSelectionOnFocus(value) { collapseSelectionOnFocus = Boolean(value); },
    setSelection(startContainer, startOffset, endContainer = startContainer, endOffset = startOffset) {
      const range = createRange();
      range.setStart(startContainer, startOffset);
      range.setEnd(endContainer, endOffset);
      selectedRange = range;
    },
    selectionOffsets(host) {
      if (!selectedRange) return undefined;
      return {
        start: boundaryOffset(host, selectedRange.startContainer, selectedRange.startOffset),
        end: boundaryOffset(host, selectedRange.endContainer, selectedRange.endOffset),
      };
    },
  };
}

function decorateElement(node, documentObject) {
  if (!node || node.type?.startsWith?.("#")) return node;
  node.nodeType = 1;
  node.tagName = String(node.type).toUpperCase();
  node.ownerDocument = documentObject;
  node.getAttribute = (name) => attributeValue(node, name);
  node.matches = (selector) => matchesSelector(node, selector);
  node.contains = (candidate) => containsNode(node, candidate);
  node.closest = (selector) => {
    let current = node;
    while (current) {
      if (matchesSelector(current, selector)) return current;
      current = current.parentElement ?? current.parent ?? null;
    }
    return null;
  };
  node.addEventListener ??= () => {};
  node.removeEventListener ??= () => {};
  node.focus = () => { documentObject.focusElement(node); };
  if (!Object.prototype.hasOwnProperty.call(node, "childNodes")) node.childNodes = [];
  return node;
}

function setHostDomText(host, text, documentObject, numberingLabel = "") {
  decorateElement(host, documentObject);
  const contentText = {
    nodeType: 3,
    textContent: text,
    childNodes: [],
    parentElement: host,
    parentNode: host,
  };
  const children = [];
  if (numberingLabel) {
    const label = {
      type: "span",
      props: { "data-docx-numbering-label": "true" },
      childNodes: [],
      parentElement: host,
      parentNode: host,
    };
    decorateElement(label, documentObject);
    const labelText = {
      nodeType: 3,
      textContent: numberingLabel,
      childNodes: [],
      parentElement: label,
      parentNode: label,
    };
    label.childNodes = [labelText];
    children.push(label);
  }
  children.push(contentText);
  host.childNodes = children;
  return contentText;
}

function installRendererDom(mounted, documentObject) {
  for (const node of walk(mounted.root)) decorateElement(node, documentObject);
  const editorRoot = walk(mounted.root).find((node) => node.props?.["data-testid"] === "docx-editor");
  assert.ok(editorRoot, "editor root should render");
  editorRoot.querySelector = (selector) => {
    for (const node of walk(mounted.root)) {
      decorateElement(node, documentObject);
      if (matchesSelector(node, selector)) return node;
    }
    return null;
  };
  return editorRoot;
}

function eventHandler(node, eventName) {
  const expected = `on${eventName}`.toLowerCase();
  const entry = Object.entries(node.props ?? {})
    .find(([name, value]) => name.toLowerCase() === expected && typeof value === "function");
  assert.ok(entry, `${eventName} handler should render`);
  return entry[1];
}

function paragraphHosts(root) {
  return walk(root).filter((node) => node.props?.["data-testid"] === "editor-paragraph");
}

async function mountEditableDocument(t, texts) {
  const dom = installDomEnvironment(t);
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: starterModel(texts) });
      return () => vue.h(DocxEditor, {
        editor: controller,
        editable: true,
        showToolbar: false,
        showThumbnails: false,
      });
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => paragraphHosts(mounted.root).length >= texts.length,
    "editable DOCX paragraphs",
  );
  installRendererDom(mounted, dom.document);
  dom.setCleanup(() => mounted.app.unmount());
  return { ...mounted, dom, get controller() { return controller; } };
}

test("real DOM Range updates activeTextRange and ignores the numbering label", async (t) => {
  const mounted = await mountEditableDocument(t, ["abcdef"]);
  const [paragraph] = paragraphHosts(mounted.root);
  const contentText = setHostDomText(paragraph, "abcdef", mounted.dom.document, "1. ");
  const pageBody = walk(mounted.root)
    .find((node) => node.props?.["data-docx-page-body"] === "true");
  assert.ok(pageBody, "page body should render");

  eventHandler(pageBody, "pointerdown")({ target: paragraph });
  mounted.dom.setSelection(contentText, 2, contentText, 5);
  mounted.dom.dispatchDocument("selectionchange");
  await vue.nextTick();

  assert.deepEqual(mounted.controller.activeTextRange, {
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 2 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 5 },
  });
  assert.deepEqual(mounted.warnings, []);
});

test("format undo and redo restore the expanded model and DOM selection after focus", async (t) => {
  const mounted = await mountEditableDocument(t, ["abcdef"]);
  const [paragraph] = paragraphHosts(mounted.root);
  const contentText = setHostDomText(paragraph, "abcdef", mounted.dom.document);
  const pageBody = walk(mounted.root)
    .find((node) => node.props?.["data-docx-page-body"] === "true");
  assert.ok(pageBody, "page body should render");

  eventHandler(pageBody, "pointerdown")({ target: paragraph });
  mounted.dom.setSelection(contentText, 2, contentText, 5);
  mounted.dom.dispatchDocument("selectionchange");
  await vue.nextTick();

  const expectedRange = {
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 2 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 5 },
  };
  assert.deepEqual(mounted.controller.activeTextRange, expectedRange);
  mounted.dom.setCollapseSelectionOnFocus(true);

  mounted.controller.toggleBold();
  await vue.nextTick();
  await vue.nextTick();
  const boldText = mounted.controller.model.nodes[0].children
    .filter((child) => child.type === "text" && child.style?.bold)
    .map((child) => child.text)
    .join("");
  assert.equal(boldText, "cde");

  mounted.controller.undo();
  await vue.nextTick();
  await vue.nextTick();
  assert.deepEqual(mounted.controller.activeTextRange, expectedRange);
  assert.deepEqual(mounted.dom.selectionOffsets(paragraph), { start: 2, end: 5 });

  mounted.controller.redo();
  await vue.nextTick();
  await vue.nextTick();
  assert.deepEqual(mounted.controller.activeTextRange, expectedRange);
  assert.deepEqual(mounted.dom.selectionOffsets(paragraph), { start: 2, end: 5 });
  assert.deepEqual(mounted.warnings, []);
});

test("Enter splits at the caret and one undo restores the original paragraph", async (t) => {
  const mounted = await mountEditableDocument(t, ["abcdef"]);
  const [paragraph] = paragraphHosts(mounted.root);
  const textNode = setHostDomText(paragraph, "abcdef", mounted.dom.document);
  mounted.dom.setSelection(textNode, 3);
  eventHandler(paragraph, "focus")();

  let prevented = false;
  eventHandler(paragraph, "keydown")({
    key: "Enter",
    shiftKey: false,
    isComposing: false,
    preventDefault() { prevented = true; },
  });
  await vue.nextTick();

  assert.equal(prevented, true);
  assert.deepEqual(mounted.controller.model.nodes.map(paragraphText), ["abc", "def"]);
  assert.deepEqual(mounted.controller.activeTextRange, {
    start: { location: { kind: "paragraph", nodeIndex: 1 }, offset: 0 },
    end: { location: { kind: "paragraph", nodeIndex: 1 }, offset: 0 },
  });
  assert.equal(mounted.controller.canUndo, true);

  mounted.controller.undo();
  await vue.nextTick();
  assert.deepEqual(mounted.controller.model.nodes.map(paragraphText), ["abcdef"]);
  assert.equal(mounted.controller.canUndo, false, "split should create exactly one history entry");
  assert.equal(mounted.controller.canRedo, true);
  assert.deepEqual(mounted.warnings, []);
});

test("Backspace at the start of a paragraph merges it with the previous paragraph", async (t) => {
  const mounted = await mountEditableDocument(t, ["left", "right"]);
  const [first, second] = paragraphHosts(mounted.root);
  setHostDomText(first, "left", mounted.dom.document);
  const secondText = setHostDomText(second, "right", mounted.dom.document);
  mounted.dom.setSelection(secondText, 0);
  eventHandler(second, "focus")();

  let prevented = false;
  eventHandler(second, "keydown")({
    key: "Backspace",
    shiftKey: false,
    isComposing: false,
    preventDefault() { prevented = true; },
  });
  await vue.nextTick();

  assert.equal(prevented, true);
  assert.deepEqual(mounted.controller.model.nodes.map(paragraphText), ["leftright"]);
  assert.deepEqual(mounted.controller.activeTextRange, {
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 4 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 4 },
  });
  assert.deepEqual(mounted.warnings, []);
});

test("composition text is committed on blur after compositionend", async (t) => {
  const mounted = await mountEditableDocument(t, ["start"]);
  const [paragraph] = paragraphHosts(mounted.root);
  setHostDomText(paragraph, "start", mounted.dom.document);

  eventHandler(paragraph, "compositionstart")({});
  const textNode = setHostDomText(paragraph, "输入完成", mounted.dom.document);
  mounted.dom.setSelection(textNode, 4);
  eventHandler(paragraph, "input")({});
  assert.equal(paragraphText(mounted.controller.model.nodes[0]), "start");

  eventHandler(paragraph, "compositionend")({});
  assert.equal(mounted.controller.selectionSessionKind, "idle");
  eventHandler(paragraph, "blur")();
  await vue.nextTick();

  assert.equal(paragraphText(mounted.controller.model.nodes[0]), "输入完成");
  assert.deepEqual(mounted.warnings, []);
});
