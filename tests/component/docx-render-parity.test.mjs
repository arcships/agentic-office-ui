import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import {
  importFromDemo,
  mount,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

const { DocxEditor, DocxViewer } = await importFromDemo("@extend-ai/vue-docx");
const { wasmBuildDocModelFromBytes } = await importFromDemo("@extend-ai/docx-core");

function classes(node) {
  return String(node?.props?.class || "").split(/\s+/).filter(Boolean);
}

function hasClass(node, name) {
  return classes(node).includes(name);
}

function descendants(node) {
  return walk(node, []).slice(1);
}

function decodeHtml(value) {
  return value
    .replace(/&#(\d+);/g, (_match, code) => String.fromCodePoint(Number(code)))
    .replace(/&#x([\da-f]+);/gi, (_match, code) => String.fromCodePoint(Number.parseInt(code, 16)))
    .replace(/&nbsp;/gi, " ")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&amp;/gi, "&");
}

function normalizeText(value) {
  return String(value || "").replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
}

function htmlText(value) {
  return normalizeText(decodeHtml(String(value || "").replace(/<[^>]*>/g, " ")));
}

function nodeText(node) {
  if (!node) return "";
  if (node.type === "#text" || node.type === "#static") return normalizeText(node.text);
  const ownHtml = node.props?.innerHTML ? htmlText(node.props.innerHTML) : "";
  const childText = (node.children || []).map((child) => nodeText(child)).join(" ");
  return normalizeText(`${ownHtml} ${childText}`);
}

function sourceKind(source) {
  if (!source) return "missing";
  if (source.startsWith("blob:")) return "blob";
  if (source.startsWith("data:")) return "data";
  if (source.startsWith("http://") || source.startsWith("https://")) return "http";
  return "relative";
}

function nodeImages(node) {
  const images = [];
  for (const item of [node, ...descendants(node)]) {
    if (item.type === "img") images.push({ sourceKind: sourceKind(String(item.props?.src || "")) });
    const html = String(item.props?.innerHTML || "");
    for (const match of html.matchAll(/<img\b[^>]*>/gi)) {
      const source = match[0].match(/\bsrc=["']([^"']*)["']/i)?.[1] || "";
      images.push({ sourceKind: sourceKind(source) });
    }
  }
  return images;
}

function directChildren(node, predicate) {
  return (node?.children || []).filter(predicate);
}

function normalizeTable(node) {
  const rows = [node, ...descendants(node)]
    .filter((item) => item.type === "tr")
    .map((row) => directChildren(row, (cell) => cell.type === "td" || cell.type === "th")
      .map((cell) => ({ text: nodeText(cell), images: nodeImages(cell) })));
  return {
    type: "table",
    text: normalizeText(rows.flat().map((cell) => cell.text).join(" ")),
    rows,
    images: nodeImages(node),
  };
}

function normalizeParagraph(node) {
  return { type: "paragraph", text: nodeText(node), images: nodeImages(node) };
}

function sectionShape(node) {
  return { text: nodeText(node), images: node ? nodeImages(node) : [] };
}

function pageShape(index, nodes, header, footer, root) {
  return {
    index,
    nodes,
    text: normalizeText(nodes.map((node) => node.text).join(" ")),
    tables: nodes.filter((node) => node.type === "table").length,
    cells: nodes.reduce((total, node) => total + (node.type === "table" ? node.rows.flat().length : 0), 0),
    images: nodeImages(root).length,
    header,
    footer,
  };
}

function viewerPages(root) {
  const legacyPages = walk(root)
    .filter((node) => node.props?.["data-testid"] === "docx-page")
    .map((page, index) => {
      const nodes = directChildren(page, (child) =>
        hasClass(child, "docx-viewer-paragraph") || hasClass(child, "docx-viewer-table"))
        .map((node) => hasClass(node, "docx-viewer-table") ? normalizeTable(node) : normalizeParagraph(node));
      const header = directChildren(page, (node) => hasClass(node, "docx-page-section--header"))[0];
      const footer = directChildren(page, (node) => hasClass(node, "docx-page-section--footer"))[0];
      return pageShape(index, nodes, sectionShape(header), sectionShape(footer), page);
    });
  return legacyPages.length > 0 ? legacyPages : surfacePages(root);
}

function surfacePages(root) {
  return walk(root)
    .filter((node) => node.props?.["data-docx-page-wrapper"] === true || node.props?.["data-docx-page-wrapper"] === "true")
    .map((wrapper) => {
      const surface = descendants(wrapper).find((node) =>
        node.props?.["data-docx-page-surface"] === true || node.props?.["data-docx-page-surface"] === "true");
      const body = descendants(surface).find((node) =>
        node.props?.["data-docx-page-body"] === true || node.props?.["data-docx-page-body"] === "true");
      const nodes = directChildren(body, (node) =>
        hasClass(node, "docx-paragraph-host") || hasClass(node, "docx-table-host"))
        .map((node) => hasClass(node, "docx-table-host") ? normalizeTable(node) : normalizeParagraph(node));
      const header = descendants(surface).find((node) =>
        node.props?.["data-docx-page-header"] === true ||
        node.props?.["data-docx-page-header"] === "true" ||
        hasClass(node, "docx-page-header"));
      const footer = descendants(surface).find((node) =>
        node.props?.["data-docx-page-footer"] === true ||
        node.props?.["data-docx-page-footer"] === "true" ||
        hasClass(node, "docx-page-footer"));
      return pageShape(Number(wrapper.props?.["data-docx-page-index"] || 0), nodes, sectionShape(header), sectionShape(footer), surface);
    })
    .sort((left, right) => left.index - right.index);
}

function interaction(root) {
  const nodes = walk(root);
  return {
    contentEditableTrue: nodes.filter((node) => node.props?.contenteditable === "true").length,
    inputs: nodes.filter((node) => node.type === "input").length,
    selects: nodes.filter((node) => node.type === "select").length,
    resizeHandles: nodes.filter((node) => classes(node).some((name) =>
      /resize-handle|column-handle|move-handle|add-row/.test(name))).length,
  };
}

function structure(pages, interactionState) {
  const nodes = pages.flatMap((page) => page.nodes);
  return {
    pageCount: pages.length,
    pages,
    totals: {
      nodes: nodes.length,
      paragraphs: nodes.filter((node) => node.type === "paragraph").length,
      tables: nodes.filter((node) => node.type === "table").length,
      cells: nodes.reduce((total, node) => total + (node.type === "table" ? node.rows.flat().length : 0), 0),
      images: pages.reduce((total, page) => total + page.images, 0),
    },
    interaction: interactionState,
  };
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function installDomEnvironment(t) {
  const originalWindow = globalThis.window;
  const originalDocument = globalThis.document;
  const originalResizeObserver = globalThis.ResizeObserver;
  const originalRect = Object.getOwnPropertyDescriptor(Object.prototype, "getBoundingClientRect");
  const originalClientHeight = Object.getOwnPropertyDescriptor(Object.prototype, "clientHeight");

  class FakeText {
    constructor(text) { this.textContent = String(text); }
    get outerHTML() { return escapeHtml(this.textContent); }
  }

  class FakeElement {
    constructor(tagName) {
      this.tagName = String(tagName).toLowerCase();
      this.attributes = new Map();
      this.children = [];
      this.style = {
        cssText: "",
        setProperty(name, value, priority = "") {
          this[name] = `${value}${priority ? ` !${priority}` : ""}`;
        },
      };
    }
    appendChild(child) { this.children.push(child); return child; }
    setAttribute(name, value) { this.attributes.set(String(name), String(value)); }
    get innerHTML() { return this.children.map((child) => child.outerHTML || "").join(""); }
    get outerHTML() {
      const attributes = [...this.attributes].map(([name, value]) =>
        value === "" ? ` ${name}` : ` ${name}="${escapeHtml(value)}"`).join("");
      const styleEntries = Object.entries(this.style)
        .filter(([name, value]) => name !== "setProperty" && name !== "cssText" && value)
        .map(([name, value]) => `${name}:${value}`);
      const styleText = this.style.cssText || styleEntries.join(";");
      const style = styleText ? ` style="${escapeHtml(styleText)}"` : "";
      return `<${this.tagName}${attributes}${style}>${this.innerHTML}</${this.tagName}>`;
    }
    getContext() { return { measureText: (text) => ({ width: String(text).length * 8 }) }; }
  }

  globalThis.window = {
    addEventListener() {},
    removeEventListener() {},
    getSelection() { return null; },
  };
  globalThis.document = {
    body: { append() {} },
    createElement(tagName) { return new FakeElement(tagName); },
    createTextNode(text) { return new FakeText(text); },
    getSelection() { return null; },
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
}

test("DocxViewer and read-only DocxEditor render one public DocModel with identical structure", async (t) => {
  const file = readFileSync(new URL("../../apps/demo/public/samples/invoice-table.docx", import.meta.url));
  const bytes = file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength);
  const { model } = await wasmBuildDocModelFromBytes(bytes);
  installDomEnvironment(t);

  const Harness = vue.defineComponent({
    setup() {
      return () => vue.h("div", { "data-testid": "docx-parity-harness" }, [
        vue.h("section", { "data-testid": "parity-viewer" }, [vue.h(DocxViewer, { model })]),
        vue.h("section", { "data-testid": "parity-editor" }, [vue.h(DocxEditor, {
          model,
          editable: false,
          showToolbar: false,
          showThumbnails: false,
        })]),
      ]);
    },
  });
  const mounted = await mount(Harness);
  await vue.nextTick();

  const viewerRoot = walk(mounted.root).find((node) => node.props?.["data-testid"] === "parity-viewer");
  const editorRoot = walk(mounted.root).find((node) => node.props?.["data-testid"] === "parity-editor");
  const viewer = structure(viewerPages(viewerRoot), interaction(viewerRoot));
  const editor = structure(surfacePages(editorRoot), interaction(editorRoot));

  assert.equal(viewer.totals.cells, 34, "the real invoice fixture must expose all 34 cells");
  assert.equal(editor.totals.cells, 34, "the read-only editor must expose all 34 cells");
  assert.deepEqual(editor.interaction, {
    contentEditableTrue: 0,
    inputs: 0,
    selects: 0,
    resizeHandles: 0,
  });
  assert.deepEqual(editor, viewer, "Viewer and read-only Editor must normalize to the same rendered structure");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});
