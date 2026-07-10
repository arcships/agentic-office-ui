import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";

const requireFromDemo = createRequire(
  new URL("../../apps/demo/package.json", import.meta.url),
);

export async function importFromDemo(specifier) {
  return import(pathToFileURL(requireFromDemo.resolve(specifier)).href);
}

export const vue = await importFromDemo("vue");

function createHostNode(type, text = "") {
  const listeners = new Map();
  return {
    type,
    text,
    props: {},
    children: [],
    parent: null,
    get options() {
      return type === "select" ? this.children.filter((child) => child.type === "option") : undefined;
    },
    getRootNode() {
      let current = this;
      while (current.parent) current = current.parent;
      return current;
    },
    addEventListener(event, listener) {
      listeners.set(event, listener);
    },
    removeEventListener(event, listener) {
      if (listeners.get(event) === listener) listeners.delete(event);
    },
    dispatchEvent(event) {
      listeners.get(event.type)?.(event);
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
    if (next == null) {
      delete node.props[key];
      if (key === "value" || key === "multiple" || key === "selected") delete node[key];
    } else {
      node.props[key] = next;
      if (key === "value" || key === "multiple" || key === "selected") node[key] = next;
    }
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

export function walk(node, result = []) {
  result.push(node);
  for (const child of node.children || []) walk(child, result);
  return result;
}

export function textContent(node) {
  return walk(node)
    .filter((item) => item.type === "#text" || item.type === "#static")
    .map((item) => item.text)
    .join("");
}

export function findByTestId(root, testId) {
  return walk(root).find((node) => node.props?.["data-testid"] === testId);
}

export async function mount(component, props = {}) {
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

export async function waitFor(predicate, message, timeoutMs = 2_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    await vue.nextTick();
    if (predicate()) return;
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
  throw new Error(`Timed out waiting for ${message}`);
}
