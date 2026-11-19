import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import {
  importFromDemo,
  mount,
  textContent,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

const {
  DocxDocumentSurface,
  DocxThumbnailPanel,
  DocxViewerRoot,
  DocxPageSurface,
  ensureDocxViewerPageSurfaceRegistry,
  useDocxEditor,
} = await importFromDemo("@arcships/vue-docx");
const {
  buildLayoutSnapshot,
  cloneDocModel,
  defaultStarterModel,
} = await importFromDemo("@arcships/docx-core");

const VUE_DOCX_COMPONENTS = new URL(
  "../../packages/vue-docx/src/components/",
  import.meta.url,
);

function paragraph(text) {
  return { type: "paragraph", children: [{ type: "text", text }] };
}

function modelWithParagraphs(count, prefix = "Planning item") {
  const model = cloneDocModel(defaultStarterModel);
  model.nodes = Array.from({ length: count }, (_, index) => paragraph(`${prefix} ${index + 1}`));
  return model;
}

const SMALL_LAYOUT = {
  pageWidth: 200,
  pageHeight: 100,
  margin: 10,
  minLineHeight: 30,
  paragraphSpacing: 10,
};

// Page geometry the viewer derives from section properties must match the
// layout options used for pagination, otherwise page boxes and pagination
// disagree in the test renderer (which cannot measure real DOM).
const SECTION_PROPERTIES =
  '<w:sectPr><w:pgSz w:w="3000" w:h="4500"/>' +
  '<w:pgMar w:top="150" w:right="150" w:bottom="150" w:left="150"/></w:sectPr>';
const VIEWER_LAYOUT = {
  pageWidth: 200,
  pageHeight: 300,
  margin: 10,
  minLineHeight: 30,
  paragraphSpacing: 10,
};

function modelForViewer(count, prefix = "Viewer item") {
  const model = modelWithParagraphs(count, prefix);
  model.metadata.sectionPropertiesXml = SECTION_PROPERTIES;
  return model;
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
      const noop = () => {};
      const context = {
        setTransform: noop,
        clearRect: noop,
        fillRect: noop,
        strokeRect: noop,
        save: noop,
        restore: noop,
        scale: noop,
        beginPath: noop,
        rect: noop,
        clip: noop,
        fillText: noop,
        moveTo: noop,
        lineTo: noop,
        stroke: noop,
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
}

function fakeController(model, overrides = {}) {
  return {
    model,
    documentLoadNonce: 1,
    syncPaginationInfo() {},
    ...overrides,
  };
}

function findComponentVNodes(vnode, name, result = []) {
  if (!vnode || typeof vnode !== "object") return result;
  if (vnode.type?.__name === name || vnode.type?.name === name) result.push(vnode);
  if (Array.isArray(vnode.children)) {
    for (const child of vnode.children) findComponentVNodes(child, name, result);
  }
  if (vnode.component?.subTree && vnode.component.subTree !== vnode) {
    findComponentVNodes(vnode.component.subTree, name, result);
  }
  return result;
}

test("thumbnail rail uses per-page page sizes and keeps the current page window", async (t) => {
  installCanvasEnvironment(t);
  const model = modelWithParagraphs(10);
  const snapshot = buildLayoutSnapshot(model, SMALL_LAYOUT);
  const controller = fakeController(model);
  const registry = ensureDocxViewerPageSurfaceRegistry(controller);
  registry.layoutPageCount = 40;
  registry.layoutPageSizes = new Map(
    Array.from({ length: 40 }, (_, pageIndex) => [
      pageIndex,
      pageIndex === 25
        ? { widthPx: 1056, heightPx: 816 }
        : { widthPx: 816, heightPx: 1056 },
    ]),
  );
  registry.layoutSnapshot = {
    ...snapshot,
    pages: Array.from({ length: 40 }, (_, pageIndex) =>
      snapshot.pages[pageIndex % snapshot.pages.length],
    ),
  };

  const mounted = await mount(DocxThumbnailPanel, {
    controller,
    model,
    currentPage: 30,
    totalPages: 40,
    showToggle: false,
  });
  await vue.nextTick();
  await vue.nextTick();

  const items = nodesWithClass(mounted.root, "docx-thumbnail-item");
  assert.equal(textContent(nodesWithClass(mounted.root, "docx-thumbnail-count")[0]), "40");
  assert.equal(
    items[0].props["data-thumbnail-page-index"],
    19,
    "the rail window starts around the current page instead of page 1",
  );
  const active = items.find((item) => classNames(item).includes("docx-thumbnail-item--active"));
  assert.equal(active?.props["data-thumbnail-page-index"], 29);
  assert.equal(textContent(active), "30");

  const canvases = walk(mounted.root).filter((node) => node.type === "canvas");
  assert.equal(canvases[0].props.height, 181, "portrait page keeps its own aspect ratio");
  assert.equal(canvases[6].props.height, 108, "landscape page is not squashed to portrait height");
  assert.equal(
    nodesWithClass(mounted.root, "docx-thumbnail-canvas-wrapper")[6].props.style.height,
    "108px",
  );
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("thumbnail rail derives per-page aspect ratios from section layout", async (t) => {
  installCanvasEnvironment(t);
  const model = modelWithParagraphs(12);
  const portrait =
    '<w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="3000" w:h="4500"/>' +
    '<w:pgMar w:top="150" w:right="150" w:bottom="150" w:left="150"/></w:sectPr>';
  const landscape =
    '<w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="4500" w:h="3000" w:orient="landscape"/>' +
    '<w:pgMar w:top="150" w:right="150" w:bottom="150" w:left="150"/></w:sectPr>';
  model.metadata.sections = [
    { startNodeIndex: 0, sectionPropertiesXml: portrait },
    { startNodeIndex: 6, sectionPropertiesXml: landscape },
  ];
  model.metadata.sectionPropertiesXml = portrait;

  const mounted = await mount(DocxThumbnailPanel, {
    model,
    layoutOptions: VIEWER_LAYOUT,
    totalPages: 2,
    showToggle: false,
  });
  await vue.nextTick();
  await vue.nextTick();

  const canvases = walk(mounted.root).filter((node) => node.type === "canvas");
  assert.equal(canvases.length, 2);
  assert.equal(canvases[0].props.height, 210, "portrait section uses its own height");
  assert.equal(canvases[1].props.height, 93, "landscape section uses its own height");
  mounted.app.unmount();
});

test("thumbnail rail does not re-paginate when a controller owns the layout", async (t) => {
  installCanvasEnvironment(t);
  const model = modelWithParagraphs(30);
  const controller = fakeController(model);

  const mounted = await mount(DocxThumbnailPanel, {
    controller,
    model,
    totalPages: 3,
    showToggle: false,
  });
  await vue.nextTick();

  assert.equal(
    textContent(nodesWithClass(mounted.root, "docx-thumbnail-count")[0]),
    "3",
    "the rail must wait for the viewer's shared page count instead of running its own layout",
  );
  assert.equal(nodesWithClass(mounted.root, "docx-thumbnail-item").length, 3);
  mounted.app.unmount();
});

test("viewer drops measured page heights when the document changes", async () => {
  const layoutOptions = VIEWER_LAYOUT;
  const firstModel = modelForViewer(10, "First document");
  const secondModel = modelForViewer(10, "Second document");
  const activeModel = vue.shallowRef(firstModel);
  const Harness = vue.defineComponent({
    setup() {
      return () => vue.h(DocxViewerRoot, {
        model: activeModel.value,
        layoutOptions,
      });
    },
  });
  const mounted = await mount(Harness);
  await vue.nextTick();

  const root = walk(mounted.root).find((node) => node.props?.["data-docx-page-count"]);
  const spacer = root.children[0];
  const initialHeight = parseFloat(spacer.props.style.height);

  const pageSurfaces = findComponentVNodes(mounted.app._instance.subTree, "DocxPageSurface");
  assert.ok(pageSurfaces.length > 0, "the viewer should render page surfaces");
  pageSurfaces[0].props.onMeasure({ heightPx: 999 });
  await vue.nextTick();

  assert.equal(
    parseFloat(spacer.props.style.height),
    initialHeight + (999 - VIEWER_LAYOUT.pageHeight),
    "a measured height should affect the virtualized scroll length",
  );

  activeModel.value = secondModel;
  await vue.nextTick();
  await vue.nextTick();

  assert.equal(
    parseFloat(spacer.props.style.height),
    initialHeight,
    "pages that were never re-measured must not keep the previous document's height",
  );
  mounted.app.unmount();
});

test("page offsets include the scroll container padding", async () => {
  const model = modelForViewer(60);
  const mounted = await mount(DocxViewerRoot, {
    model,
    layoutOptions: VIEWER_LAYOUT,
  });
  await vue.nextTick();

  const root = walk(mounted.root).find((node) => node.props?.["data-docx-page-count"]);
  const spacer = root.children[0];
  const pageWrappers = walk(root).filter((node) => node.props?.["data-docx-page-wrapper"] === true);
  const pageCount = Number(root.props["data-docx-page-count"]);

  assert.ok(pageWrappers.length >= 2, "multiple pages should be visible in the test viewport");
  assert.equal(
    pageWrappers[0].props.style.top,
    "0px",
    "the spacer already starts after the 24px container padding",
  );
  assert.equal(
    pageWrappers[1].props.style.top,
    `${VIEWER_LAYOUT.pageHeight + 28}px`,
    "later pages are offset from the padded start",
  );
  assert.equal(
    parseFloat(spacer.props.style.height),
    24 + pageCount * (VIEWER_LAYOUT.pageHeight + 28),
    "scroll-space offsets include the container padding",
  );

  mounted.app.unmount();
});

test("fit-width reserves the annotation gutter", async () => {
  const model = modelForViewer(10);
  const mounted = await mount(DocxViewerRoot, {
    model,
    layoutOptions: VIEWER_LAYOUT,
    fitWidth: true,
    showComments: true,
  });
  await vue.nextTick();

  const root = walk(mounted.root).find((node) => node.props?.["data-docx-page-count"]);
  const spacer = root.children[0];
  const pageWrappers = walk(root).filter((node) => node.props?.["data-docx-page-wrapper"] === true);

  const fittedWidth = parseFloat(spacer.props.style.minWidth);
  assert.ok(
    fittedWidth <= 800.5,
    `fit-width footer (${fittedWidth}px) must fit the 800px container including the gutter`,
  );
  const wrapperWidth = parseFloat(pageWrappers[0].props.style.width);
  assert.ok(wrapperWidth < 800, "the paper plus gutter footprint must fit inside the container");

  mounted.app.unmount();
});

test("drag-over outline clears after a single dragleave", async () => {
  const model = modelForViewer(2);
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model });
      return () => vue.h(DocxViewerRoot, {
        model: controller.model,
        controller,
        editable: true,
        layoutOptions: VIEWER_LAYOUT,
      });
    },
  });
  const mounted = await mount(Harness);
  await vue.nextTick();

  const root = walk(mounted.root).find((node) => node.props?.["data-docx-page-count"]);
  const event = {
    preventDefault() {},
    dataTransfer: {
      types: ["Files"],
      files: [{ name: "dropped.docx" }],
      dropEffect: "none",
    },
  };
  root.props.onDragover(event);
  await vue.nextTick();
  root.props.onDragover(event);
  await vue.nextTick();
  assert.equal(root.props.style.outline, "3px dashed #2563eb");

  root.props.onDragleave({ relatedTarget: null });
  await vue.nextTick();
  assert.equal(root.props.style.outline, "none", "the dashed outline must not stick after dragleave");

  mounted.app.unmount();
  controller = undefined;
});

test("comment gutter stays outside the clipped paper box", async () => {
  const model = modelWithParagraphs(1);
  const mounted = await mount(DocxPageSurface, {
    pageIndex: 0,
    pageLayout: {
      pageWidthPx: 816,
      pageHeightPx: 1056,
      marginsPx: { top: 72, right: 72, bottom: 72, left: 72 },
    },
    pageNodeSegments: [{ nodeIndex: 0 }],
    model,
    pageWidthPx: 816,
    pageContentWidthPx: 672,
    pageNumber: 1,
    totalPages: 1,
    commentsEnabled: true,
    comments: [{
      id: "comment-1",
      nodeIndex: 0,
      author: "Reviewer",
      date: "2026-01-01T00:00:00Z",
      text: "Tighten this paragraph.",
      resolved: false,
    }],
  });
  await vue.nextTick();

  const frame = nodesWithClass(mounted.root, "docx-page-frame")[0];
  const surface = walk(mounted.root).find(
    (node) => node.props?.["data-docx-page-surface"] === "true",
  );
  const gutter = nodesWithClass(mounted.root, "docx-tracked-change-gutter")[0];
  assert.ok(frame && surface && gutter, "frame, clipped surface and gutter must render");
  assert.equal(surface.parent, frame, "the paper surface is clipped inside the frame");
  assert.equal(gutter.parent, frame, "the gutter is a sibling of the clipped surface");
  assert.equal(
    walk(surface).includes(gutter),
    false,
    "the gutter must not sit inside the clipped page box",
  );

  const surfaceSource = readFileSync(
    new URL("DocxPageSurface.vue", VUE_DOCX_COMPONENTS),
    "utf8",
  );
  const gutterSource = readFileSync(
    new URL("DocxTrackedChangeGutter.vue", VUE_DOCX_COMPONENTS),
    "utf8",
  );
  assert.match(
    surfaceSource,
    /class="docx-page-frame"[^>]*>[\s\S]*data-docx-page-surface="true"[\s\S]*<DocxTrackedChangeGutter/s,
    "the gutter template must stay outside the clipped surface",
  );
  assert.match(surfaceSource, /overflow:\s*"hidden"/, "the paper surface must clip overflow");
  assert.match(
    gutterSource,
    /\.docx-tracked-change-gutter\s*\{[^}]*left:\s*calc\(100%\s*\+\s*16px\)/s,
    "the gutter is absolutely positioned beside the paper",
  );
  assert.doesNotMatch(
    gutterSource,
    /position:\s*static/,
    "the narrow-width gutter must stay out of the page's flow",
  );
  mounted.app.unmount();
});

test("search highlights translate cached geometry while scrolling", async (t) => {
  const originalNode = globalThis.Node;
  const originalDocument = globalThis.document;
  const hostRect = () => ({ left: 0, top: 0, right: 600, bottom: 400, width: 600, height: 400 });
  const rangeRect = () => ({ left: 20, top: 100, right: 120, bottom: 116, width: 100, height: 16 });
  globalThis.Node = { TEXT_NODE: 3 };
  globalThis.document = {
    createElement() {
      return {
        style: {},
        dataset: {},
        width: 0,
        height: 0,
        getContext: () => null,
        setAttribute() {},
        appendChild() {},
      };
    },
    createTextNode(text) {
      return { nodeType: 3, textContent: String(text) };
    },
    createDocumentFragment() {
      return { appendChild() {}, childNodes: [] };
    },
    createRange() {
      return {
        setStart() {},
        setEnd() {},
        getClientRects: () => [rangeRect()],
        getBoundingClientRect: rangeRect,
      };
    },
  };
  t.after(() => {
    if (originalNode === undefined) delete globalThis.Node;
    else globalThis.Node = originalNode;
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
  });

  const model = modelWithParagraphs(1);
  model.nodes[0].children[0].text = "alpha beta alpha";
  const mounted = await mount(DocxDocumentSurface, { model });
  await vue.nextTick();

  const surfaceRoot = walk(mounted.root).find((node) =>
    classNames(node).includes("docx-document-surface"),
  );
  const container = walk(mounted.root).find((node) => node.props?.["data-docx-page-count"]);
  assert.ok(surfaceRoot);
  assert.ok(container);
  surfaceRoot.getBoundingClientRect = hostRect;
  container.getBoundingClientRect = hostRect;
  container.scrollTop = 0;
  container.scrollLeft = 0;

  let queryCount = 0;
  const textNode = { nodeType: 3, textContent: "alpha beta alpha" };
  const paragraphHost = {
    nodeType: 1,
    getBoundingClientRect: hostRect,
    childNodes: [textNode],
  };
  surfaceRoot.querySelectorAll = (selector) => {
    queryCount += 1;
    return /\[data-docx-paragraph-node-index="0"\]/u.test(selector) ? [paragraphHost] : [];
  };

  const surface = mounted.app._instance.subTree.component.exposed;
  await surface.search("alpha");
  await vue.nextTick();

  const readHighlights = () => walk(mounted.root).filter((node) =>
    classNames(node).includes("docx-document-surface__search-highlight"),
  );
  assert.equal(readHighlights().length, 2);
  assert.equal(parseFloat(readHighlights()[0].props.style.top), 100);
  const queriesAfterSearch = queryCount;
  assert.ok(queriesAfterSearch > 0);

  container.scrollTop = 100;
  container.dispatchEvent({ type: "scroll" });
  await vue.nextTick();

  assert.equal(
    queryCount,
    queriesAfterSearch,
    "scrolling must reuse cached ranges instead of walking the DOM per match",
  );
  assert.equal(parseFloat(readHighlights()[0].props.style.top), 0, "cached rects translate with scroll");

  mounted.app.unmount();
});

test("thumbnail surface cache invalidates per page instead of clearing on every publish", () => {
  const source = readFileSync(
    new URL(
      "../../packages/vue-docx/src/composables/useDocxPageThumbnails.ts",
      import.meta.url,
    ),
    "utf8",
  );
  const start = source.indexOf("const renderTrackedPages");
  const end = source.indexOf("const unsubscribe");
  assert.ok(start >= 0 && end > start, "renderTrackedPages must exist");
  const renderTrackedPages = source.slice(start, end);
  assert.doesNotMatch(
    renderTrackedPages,
    /surfaceCache\.clear\(\)/,
    "surface publishes must not discard every cached thumbnail",
  );
  assert.match(
    renderTrackedPages,
    /surfaceCache\.delete\(String\(pageIndex\)\)/,
    "changed page content signatures should invalidate only that page",
  );
});
