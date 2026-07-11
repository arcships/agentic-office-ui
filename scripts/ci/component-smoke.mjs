#!/usr/bin/env node

import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";

const requireFromDemo = createRequire(
  new URL("../../apps/demo/package.json", import.meta.url),
);

async function importFromDemo(specifier) {
  return import(pathToFileURL(requireFromDemo.resolve(specifier)).href);
}

const vue = await importFromDemo("vue");
const { DocxViewer } = await importFromDemo("@arcships/vue-docx");
const { XlsxViewer } = await importFromDemo("@arcships/vue-xlsx");

function createHostNode(type, text = "") {
  const listeners = new Map();
  return {
    type,
    text,
    props: {},
    children: [],
    parent: null,
    addEventListener(event, listener) {
      const entries = listeners.get(event) ?? new Set();
      entries.add(listener);
      listeners.set(event, entries);
    },
    removeEventListener(event, listener) {
      listeners.get(event)?.delete(listener);
    },
  };
}

function insertHostNode(node, parent, anchor = null) {
  if (node.parent) {
    const oldIndex = node.parent.children.indexOf(node);
    if (oldIndex >= 0) node.parent.children.splice(oldIndex, 1);
  }
  node.parent = parent;
  if (!anchor) parent.children.push(node);
  else parent.children.splice(parent.children.indexOf(anchor), 0, node);
}

const renderer = vue.createRenderer({
  patchProp(node, key, _previous, next) {
    if (next == null) delete node.props[key];
    else node.props[key] = next;
  },
  insert(node, parent, anchor = null) {
    insertHostNode(node, parent, anchor);
  },
  remove(node) {
    if (!node.parent) return;
    const index = node.parent.children.indexOf(node);
    if (index >= 0) node.parent.children.splice(index, 1);
    node.parent = null;
  },
  createElement(type) {
    return createHostNode(type);
  },
  createText(text) {
    return createHostNode("#text", text);
  },
  createComment(text) {
    return createHostNode("#comment", text);
  },
  setText(node, text) {
    node.text = text;
  },
  setElementText(node, text) {
    const child = createHostNode("#text", text);
    child.parent = node;
    node.children = [child];
  },
  parentNode(node) {
    return node.parent;
  },
  nextSibling(node) {
    if (!node.parent) return null;
    const index = node.parent.children.indexOf(node);
    return node.parent.children[index + 1] || null;
  },
  setScopeId() {},
  cloneNode(node) {
    return { ...node, props: { ...node.props }, children: [...node.children] };
  },
  insertStaticContent(content, parent, anchor) {
    const node = createHostNode("#static", content);
    insertHostNode(node, parent, anchor);
    return [node, node];
  },
});

function walk(node, result = []) {
  result.push(node);
  for (const child of node.children || []) walk(child, result);
  return result;
}

function textContent(node) {
  return walk(node)
    .filter((item) => item.type === "#text" || item.type === "#static")
    .map((item) => item.text)
    .join("");
}

async function mount(component, props) {
  const root = createHostNode("#root");
  const warnings = [];
  const app = renderer.createApp({
    render: () => vue.h(component, props),
  });
  app.config.warnHandler = (message) => warnings.push(message);
  app.mount(root);
  await vue.nextTick();
  return { app, root, warnings };
}

const docx = await mount(DocxViewer, { emptyState: "DOCX CI empty state" });
const docxNodes = walk(docx.root);
const docxRoot = docxNodes.find(
  (node) => node.props?.["data-testid"] === "docx-viewer",
);
assert.ok(docxRoot, "DocxViewer renders its public root element");
assert.equal(docxRoot.props["data-state"], "idle");
assert.match(textContent(docx.root), /DOCX CI empty state/);
assert.deepEqual(docx.warnings, [], "DocxViewer renders without Vue warnings");
docx.app.unmount();

const emptyController = {
  readOnly: true,
  isLoading: false,
  error: null,
  activeSheet: null,
  tabs: [],
};
const xlsx = await mount(XlsxViewer, {
  controller: emptyController,
  height: "320px",
  showDefaultToolbar: false,
  showRibbon: false,
  showFormulaBar: false,
  showImages: false,
});
const xlsxRoot = walk(xlsx.root).find((node) => {
  const value = node.props?.class;
  return typeof value === "string" && value.split(/\s+/).includes("xlsx-viewer");
});
assert.ok(xlsxRoot, "XlsxViewer renders its public root element");
assert.match(textContent(xlsx.root), /打开一个 XLSX 文件开始查看/);
assert.deepEqual(xlsx.warnings, [], "XlsxViewer renders without Vue warnings");
xlsx.app.unmount();

console.log("PASS: DOCX/XLSX public components render through Vue without DOM or private state access.");
