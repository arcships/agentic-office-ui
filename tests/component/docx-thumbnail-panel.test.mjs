import assert from "node:assert/strict";
import test from "node:test";

import {
  importFromDemo,
  mount,
  textContent,
  vue,
  waitFor,
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
        fillRect: (...args) => operations.push({
          canvas: this,
          name: "fillRect",
          args,
          fillStyle: context.fillStyle,
        }),
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
        drawImage: (...args) => record("drawImage", ...args),
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
    ...Array.from({ length: 249 }, (_, index) => paragraph(`Planning item ${index + 1}`)),
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

  const list = nodesWithClass(mounted.root, "docx-thumbnail-list")[0];
  const pageCount = Number(list.props["data-page-count"]);
  assert.ok(pageCount > 20, `expected a long document, got ${pageCount} pages`);
  assert.equal(nodesWithClass(mounted.root, "docx-thumbnail-item").length, 20);
  // Upstream row stride: a 92px card at this 2:1 page ratio (46px) plus the
  // 40px option box and 12px `pb-3` row gap, with 16px track padding.
  const rowStride = 46 + 52;
  const track = nodesWithClass(mounted.root, "docx-thumbnail-track")[0];
  assert.equal(track.props.style.height, `${pageCount * rowStride + 32}px`);

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

  list.props.onScrollPassive({ target: { scrollTop: 24 * rowStride } });
  await vue.nextTick();
  await vue.nextTick();

  const visibleItems = nodesWithClass(mounted.root, "docx-thumbnail-item");
  assert.equal(visibleItems[0].props["data-thumbnail-page-index"], 19);
  assert.equal(visibleItems.at(-1).props["data-thumbnail-page-index"], Math.min(19 + 20, pageCount) - 1);
  const rows = nodesWithClass(mounted.root, "docx-thumbnail-row");
  assert.equal(
    rows[0].props.style.transform,
    `translateY(${16 + 19 * rowStride}px)`,
    "the first rendered row must sit at its virtual offset inside the track",
  );
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DOCX thumbnail paints the document page background colour", async (t) => {
  const operations = installCanvasEnvironment(t);
  const model = cloneDocModel(defaultStarterModel);
  model.metadata.documentBackgroundColor = "DDEBF7";
  model.nodes = [paragraph("Tinted page")];

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

  const firstCanvas = walk(mounted.root).find((node) => node.type === "canvas");
  assert.equal(firstCanvas.dataset.thumbnailState, "ready");
  const pageFill = operations.find((operation) =>
    operation.canvas === firstCanvas &&
    operation.name === "fillRect" &&
    operation.fillStyle === "#DDEBF7"
  );
  assert.ok(pageFill, "the thumbnail canvas must use the document background colour");
  assert.deepEqual(pageFill.args.slice(0, 2), [0, 0]);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DOCX thumbnail paints a patterned page background", async (t) => {
  const operations = installCanvasEnvironment(t);
  const model = cloneDocModel(defaultStarterModel);
  model.metadata.documentBackgroundColor = "DDEBF7";
  model.metadata.documentBackgroundPattern = {
    preset: "pct20",
    foregroundColor: "#2F5597",
    backgroundColor: "#DDEBF7",
  };
  model.nodes = [paragraph("Patterned page")];

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

  const firstCanvas = walk(mounted.root).find((node) => node.type === "canvas");
  assert.ok(
    operations.some((operation) =>
      operation.canvas === firstCanvas &&
      operation.name === "fillRect" &&
      operation.fillStyle === "#DDEBF7" &&
      operation.args[2] > 1 &&
      operation.args[3] > 1
    ),
    "the thumbnail canvas must paint the pattern background colour"
  );
  assert.ok(
    operations.some((operation) =>
      operation.canvas === firstCanvas &&
      operation.name === "fillRect" &&
      operation.fillStyle === "#2F5597" &&
      operation.args[2] === 1 &&
      operation.args[3] === 1
    ),
    "the thumbnail canvas must paint pattern dots"
  );
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DOCX thumbnail rail wraps body text and paints headers, footers and notes", async (t) => {
  const operations = installCanvasEnvironment(t);
  const model = cloneDocModel(defaultStarterModel);
  model.nodes = [
    {
      type: "paragraph",
      children: [{ type: "text", text: "word ".repeat(40) }],
    },
  ];
  model.metadata.headerSections = [
    {
      partName: "header1.xml",
      referenceType: "default",
      nodes: [paragraph("Confidential header")],
    },
  ];
  model.metadata.footerSections = [
    {
      partName: "footer1.xml",
      referenceType: "default",
      nodes: [paragraph("Footer disclaimer")],
    },
  ];
  model.metadata.endnotes = [{ id: 1, text: "Endnote details" }];

  const mounted = await mount(DocxThumbnailPanel, {
    model,
    layoutOptions: {
      pageWidth: 300,
      pageHeight: 400,
      margin: 30,
      minLineHeight: 80,
      paragraphSpacing: 4,
    },
    showToggle: false,
  });
  await vue.nextTick();
  await vue.nextTick();

  const firstCanvas = walk(mounted.root).find((node) => node.type === "canvas");
  const paintedText = operations
    .filter((operation) => operation.canvas === firstCanvas && operation.name === "fillText")
    .map((operation) => String(operation.args[0]));

  assert.ok(
    paintedText.filter((text) => text === "word").length > 1,
    "wrapped body text must paint more than one line",
  );
  assert.ok(paintedText.some((text) => text.includes("Confidential")));
  assert.ok(paintedText.some((text) => text.includes("Footer")));
  assert.ok(paintedText.some((text) => text.includes("Endnote")));
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DOCX thumbnail rail draws embedded images instead of an X placeholder", async (t) => {
  const operations = installCanvasEnvironment(t);
  const originalImage = globalThis.Image;
  globalThis.Image = class {
    constructor() {
      this.naturalWidth = 120;
      this.naturalHeight = 80;
      this.width = 120;
      this.height = 80;
    }

    set src(value) {
      this._src = value;
      queueMicrotask(() => this.onload?.());
    }

    get src() {
      return this._src;
    }
  };
  t.after(() => {
    if (originalImage === undefined) delete globalThis.Image;
    else globalThis.Image = originalImage;
  });

  const model = cloneDocModel(defaultStarterModel);
  model.nodes = [
    {
      type: "paragraph",
      children: [
        {
          type: "image",
          src: "data:image/png;base64,AAAA",
          widthPx: 120,
          heightPx: 80,
        },
      ],
    },
  ];

  const mounted = await mount(DocxThumbnailPanel, {
    model,
    layoutOptions: { pageWidth: 200, pageHeight: 300, margin: 20, minLineHeight: 20 },
    showToggle: false,
  });

  const firstCanvas = walk(mounted.root).find((node) => node.type === "canvas");
  await waitFor(
    () => operations.some((operation) => operation.canvas === firstCanvas && operation.name === "drawImage"),
    "embedded image draw",
  );
  assert.ok(
    operations.some(
      (operation) =>
        operation.canvas === firstCanvas &&
        operation.name === "drawImage" &&
        operation.args[0]?.naturalWidth === 120,
    ),
  );
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DOCX thumbnail rail follows device pixel ratio changes", async (t) => {
  installCanvasEnvironment(t);
  const originalWindow = globalThis.window;
  const resizeListeners = new Set();
  globalThis.window = {
    devicePixelRatio: 1,
    addEventListener(type, listener) {
      if (type === "resize") resizeListeners.add(listener);
    },
    removeEventListener(type, listener) {
      if (type === "resize") resizeListeners.delete(listener);
    },
    matchMedia() {
      return {
        addEventListener() {},
        removeEventListener() {},
      };
    },
  };
  t.after(() => {
    if (originalWindow === undefined) delete globalThis.window;
    else globalThis.window = originalWindow;
  });

  const model = cloneDocModel(defaultStarterModel);
  model.nodes = [paragraph("Resolution test")];
  const mounted = await mount(DocxThumbnailPanel, {
    model,
    layoutOptions: { pageWidth: 200, pageHeight: 300, margin: 20 },
    showToggle: false,
  });
  await vue.nextTick();
  await vue.nextTick();

  const firstCanvas = walk(mounted.root).find((node) => node.type === "canvas");
  assert.equal(firstCanvas.props.width, 92);
  assert.equal(firstCanvas.props.height, firstCanvas.props.width * (300 / 200));

  globalThis.window.devicePixelRatio = 2;
  for (const listener of resizeListeners) listener();
  await vue.nextTick();
  await vue.nextTick();

  assert.equal(firstCanvas.props.width, 184);
  assert.equal(firstCanvas.props.height, 184 * (300 / 200));
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});
