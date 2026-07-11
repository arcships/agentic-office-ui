import assert from "node:assert/strict";
import test from "node:test";

import {
  importFromDemo,
  mount,
  textContent,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

const { DocxThumbnailPanel } = await importFromDemo("@arcships/vue-docx");
const { cloneDocModel, defaultStarterModel } = await importFromDemo("@arcships/docx-core");

function paragraph(text) {
  return {
    type: "paragraph",
    children: [{ type: "text", text }],
  };
}

function classNames(node) {
  return String(node?.props?.class || "").split(/\s+/).filter(Boolean);
}

function nodesWithClass(root, className) {
  return walk(root).filter((node) => classNames(node).includes(className));
}

function installCanvasEnvironment(t) {
  const originalDataset = Object.getOwnPropertyDescriptor(Object.prototype, "dataset");
  const originalGetContext = Object.getOwnPropertyDescriptor(Object.prototype, "getContext");
  const operations = [];

  Object.defineProperty(Object.prototype, "dataset", {
    configurable: true,
    get() {
      if (!Object.hasOwn(this, "__thumbnailDataset")) {
        Object.defineProperty(this, "__thumbnailDataset", {
          configurable: true,
          value: {},
        });
      }
      return this.__thumbnailDataset;
    },
  });
  Object.defineProperty(Object.prototype, "getContext", {
    configurable: true,
    value(kind) {
      if (this.type !== "canvas" || kind !== "2d") return null;
      if (Object.hasOwn(this, "__thumbnailContext")) return this.__thumbnailContext;
      const record = (name, ...args) => operations.push({ canvas: this, name, args });
      const context = {
        setTransform: (...args) => record("setTransform", ...args),
        clearRect: (...args) => record("clearRect", ...args),
        fillRect: (...args) => record("fillRect", ...args),
        strokeRect: (...args) => record("strokeRect", ...args),
        save: () => record("save"),
        restore: () => record("restore"),
        scale: (...args) => record("scale", ...args),
        beginPath: () => record("beginPath"),
        rect: (...args) => record("rect", ...args),
        clip: () => record("clip"),
        fillText: (...args) => record("fillText", ...args),
        moveTo: (...args) => record("moveTo", ...args),
        lineTo: (...args) => record("lineTo", ...args),
        stroke: () => record("stroke"),
        measureText: (value) => ({ width: String(value).length * 8 }),
      };
      Object.defineProperty(this, "__thumbnailContext", {
        configurable: true,
        value: context,
      });
      return context;
    },
  });

  t.after(() => {
    if (originalDataset) Object.defineProperty(Object.prototype, "dataset", originalDataset);
    else delete Object.prototype.dataset;
    if (originalGetContext) Object.defineProperty(Object.prototype, "getContext", originalGetContext);
    else delete Object.prototype.getContext;
  });

  return operations;
}

test("DOCX thumbnail rail keeps long-document scroll space and paints table text", async (t) => {
  const operations = installCanvasEnvironment(t);
  const model = cloneDocModel(defaultStarterModel);
  model.nodes = [
    {
      type: "table",
      rows: [{ cells: [{ nodes: [paragraph("Quarterly total")], style: { backgroundColor: "E5E7EB" } }] }],
    },
    ...Array.from({ length: 49 }, (_, index) => paragraph(`Planning item ${index + 1}`)),
  ];

  const mounted = await mount(DocxThumbnailPanel, {
    model,
    layoutOptions: {
      pageWidth: 200,
      pageHeight: 100,
      margin: 10,
      minLineHeight: 30,
      paragraphSpacing: 10,
    },
    showToggle: false,
  });
  await vue.nextTick();
  await vue.nextTick();

  assert.equal(textContent(nodesWithClass(mounted.root, "docx-thumbnail-count")[0]), "26");
  assert.equal(nodesWithClass(mounted.root, "docx-thumbnail-item").length, 20);
  const bottomSpacer = nodesWithClass(mounted.root, "docx-thumbnail-spacer").at(-1);
  assert.equal(bottomSpacer.props.style.height, "636px");

  const firstCanvas = walk(mounted.root).find((node) => node.type === "canvas");
  assert.equal(firstCanvas.dataset.thumbnailState, "ready");
  const firstCanvasOperations = operations.filter((operation) => operation.canvas === firstCanvas);
  const contentScale = firstCanvasOperations.findIndex((operation) => operation.name === "scale");
  const tableFill = firstCanvasOperations.findIndex(
    (operation, index) => index > contentScale && operation.name === "fillRect"
  );
  const tableStroke = firstCanvasOperations.findIndex(
    (operation, index) => index > tableFill && operation.name === "strokeRect"
  );
  assert.ok(tableFill > contentScale && tableStroke > tableFill, "table fill must precede its border");
  assert.ok(firstCanvasOperations.some((operation) => operation.name === "fillText"));

  const list = nodesWithClass(mounted.root, "docx-thumbnail-list")[0];
  list.props.onScrollPassive({ target: { scrollTop: 24 * 106 } });
  await vue.nextTick();
  await vue.nextTick();

  const visibleItems = nodesWithClass(mounted.root, "docx-thumbnail-item");
  assert.equal(visibleItems[0].props["data-thumbnail-page-index"], 19);
  assert.equal(visibleItems.at(-1).props["data-thumbnail-page-index"], 25);
  const topSpacer = nodesWithClass(mounted.root, "docx-thumbnail-spacer")[0];
  assert.equal(topSpacer.props.style.height, `${19 * 106}px`);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});
