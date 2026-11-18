import assert from "node:assert/strict";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

// Vue's native v-model directive checks these browser constructors during an
// update. The custom component renderer intentionally has no browser DOM.
if (globalThis.Document === undefined) globalThis.Document = class Document {};
if (globalThis.ShadowRoot === undefined) globalThis.ShadowRoot = class ShadowRoot {};

const { DocxViewer, DocxViewerRoot } = await importFromDemo("@arcships/vue-docx");
const { cloneDocModel, defaultStarterModel } = await importFromDemo(
  "@arcships/docx-core",
);

function paragraph(text) {
  return { type: "paragraph", children: [{ type: "text", text }] };
}

const SECTION_PROPERTIES =
  '<w:sectPr><w:pgSz w:w="3000" w:h="4500"/>' +
  '<w:pgMar w:top="150" w:right="150" w:bottom="150" w:left="150"/></w:sectPr>';

// 200 x 300 px pages with a 30px minimum line height: the four paragraphs of
// the fixtures below paginate one paragraph per page.
const VIEWER_LAYOUT = {
  pageWidth: 200,
  pageHeight: 300,
  margin: 10,
  minLineHeight: 30,
  paragraphSpacing: 10,
};

function modelForViewer(count, prefix = "Viewer item") {
  const model = cloneDocModel(defaultStarterModel);
  model.nodes = Array.from({ length: count }, (_, index) =>
    paragraph(`${prefix} ${index + 1}`),
  );
  model.metadata.sectionPropertiesXml = SECTION_PROPERTIES;
  return model;
}

function viewerRoot(root) {
  return walk(root).find((node) => node.props?.["data-docx-page-count"] !== undefined);
}

function pageWrappers(root) {
  return walk(root).filter((node) => node.props?.["data-docx-page-wrapper"] === true);
}

function spacerOf(root) {
  const container = viewerRoot(root);
  return container?.children[0];
}

/** Gives the fake DOM node the scrollbox properties DocxViewerRoot reads. */
function makeScrollable(node, { clientHeight = 600, clientWidth = 800, scrollTop = 0 } = {}) {
  const state = { scrollTop };
  Object.defineProperty(node, "clientHeight", {
    configurable: true,
    get: () => clientHeight,
  });
  Object.defineProperty(node, "clientWidth", {
    configurable: true,
    get: () => clientWidth,
  });
  Object.defineProperty(node, "scrollTop", {
    configurable: true,
    get: () => state.scrollTop,
    set: (value) => {
      state.scrollTop = value;
    },
  });
  const scrolls = [];
  node.scrollTo = (options) => {
    scrolls.push(options);
    if (typeof options?.top === "number") state.scrollTop = options.top;
  };
  node.getBoundingClientRect = () => ({
    bottom: clientHeight,
    height: clientHeight,
    left: 0,
    right: clientWidth,
    top: 0,
    width: clientWidth,
  });
  return { scrolls, state };
}

test("DOCX spread mode renders two pages side by side with their own page numbers", async () => {
  const model = modelForViewer(120);
  const mounted = await mount(DocxViewerRoot, {
    model,
    layoutOptions: VIEWER_LAYOUT,
    viewMode: "spread",
  });
  await vue.nextTick();

  const root = viewerRoot(mounted.root);
  assert.equal(root.props["data-docx-view-mode"], "spread");
  assert.equal(root.props["data-docx-page-count"], 6, "120 paragraphs paginate onto 6 pages");

  const wrappers = pageWrappers(root);
  assert.ok(wrappers.length >= 2, "the first spread mounts both of its pages");
  assert.equal(wrappers.length % 2, 0, "spread rows always mount in pairs");
  assert.deepEqual(
    wrappers.slice(0, 2).map((wrapper) => wrapper.props["data-docx-page-index"]),
    [0, 1],
  );
  assert.deepEqual(
    wrappers.slice(0, 2).map((wrapper) => wrapper.props["data-docx-page-number"]),
    [1, 2],
    "each spread page keeps its own page number",
  );
  assert.equal(
    wrappers[0].props.style.top,
    wrappers[1].props.style.top,
    "both pages of a spread start on the same row",
  );
  for (let index = 0; index < wrappers.length; index += 2) {
    assert.equal(
      wrappers[index].props.style.top,
      wrappers[index + 1].props.style.top,
      `row ${index / 2 + 1} keeps its pair on one top offset`,
    );
  }
  // 200px pages with a 24px spread gap: the pair is centered as one row.
  assert.equal(wrappers[0].props.style.left, "calc(50% + -112px)");
  assert.equal(wrappers[1].props.style.left, "calc(50% + 112px)");

  const spacer = spacerOf(root);
  assert.equal(
    parseFloat(spacer.props.style.height),
    24 + 3 * (VIEWER_LAYOUT.pageHeight + 28),
    "the scroll length counts spread rows, not pages",
  );
  assert.equal(
    parseFloat(spacer.props.style.minWidth),
    200 + 24 + 200 + 48,
    "the scroll width covers the whole spread",
  );
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DOCX single mode mounts one page and snaps the scroll offset to its row", async () => {
  const model = modelForViewer(120);
  const mounted = await mount(DocxViewerRoot, {
    model,
    layoutOptions: VIEWER_LAYOUT,
    viewMode: "single",
  });
  await vue.nextTick();

  const root = viewerRoot(mounted.root);
  assert.equal(root.props["data-docx-view-mode"], "single");
  const scroll = makeScrollable(root, { clientHeight: 400, scrollTop: 300 });
  await vue.nextTick();

  // Page rows start at 24, 352, 680 and 1008; a 400px viewport centered at
  // 500px sits inside the second page, so the reader rests on page 2.
  root.props.onScrollPassive();
  await waitFor(
    () => scroll.state.scrollTop === 352,
    "single-page snapping pulled the viewport to the page boundary",
  );
  const wrappers = pageWrappers(root);
  assert.equal(wrappers.length, 1, "only one page is mounted at a time");
  assert.equal(wrappers[0].props["data-docx-page-index"], 1);
  assert.equal(wrappers[0].props["data-docx-page-number"], 2);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DOCX context menu resolves the left and right page of a spread", async () => {
  const model = modelForViewer(120);
  const contexts = [];
  const mounted = await mount(DocxViewerRoot, {
    model,
    layoutOptions: VIEWER_LAYOUT,
    viewMode: "spread",
    onContextMenu: (ctx) => contexts.push(ctx),
  });
  await vue.nextTick();

  const root = viewerRoot(mounted.root);
  makeScrollable(root, { clientHeight: 600, clientWidth: 800 });
  await vue.nextTick();

  // The 424px spread is centered in the 800px container: pages cover
  // 188..388 and 412..612, both inside the first row (0..300).
  root.props.onContextmenu({ clientX: 250, clientY: 100, preventDefault() {} });
  root.props.onContextmenu({ clientX: 500, clientY: 100, preventDefault() {} });
  assert.deepEqual(contexts.map((ctx) => ctx.pageIndex), [0, 1]);
  mounted.app.unmount();
});

test("DOCX viewer switching modes keeps the current page", async () => {
  const model = modelForViewer(120);
  const viewMode = vue.ref("continuous");
  const mounted = await mount(
    vue.defineComponent({
      setup() {
        return () =>
          vue.h(DocxViewer, {
            model,
            layoutOptions: VIEWER_LAYOUT,
            viewMode: viewMode.value,
            "onUpdate:viewMode": (mode) => {
              viewMode.value = mode;
            },
          });
      },
    }),
  );
  await waitFor(
    () => findByTestId(mounted.root, "docx-page-current") !== undefined,
    "DOCX viewer toolbar",
  );
  const root = viewerRoot(mounted.root);
  const scroll = makeScrollable(root, { clientHeight: 400, scrollTop: 0 });
  await vue.nextTick();

  // Navigate to page 3 through the toolbar page field.
  const pageInput = findByTestId(mounted.root, "docx-page-current");
  pageInput.value = "3";
  pageInput.dispatchEvent({ type: "input", target: pageInput });
  const keydownHandlers = Array.isArray(pageInput.props.onKeydown)
    ? pageInput.props.onKeydown
    : [pageInput.props.onKeydown];
  for (const handler of keydownHandlers) handler?.({ key: "Enter", preventDefault() {} });
  await waitFor(
    () => scroll.scrolls.at(-1)?.top === 680,
    "page 3 scroll request",
  );
  root.props.onScrollPassive();
  await vue.nextTick();

  // Switching to a spread keeps page 3 selected and scrolls to its row.
  await findByTestId(mounted.root, "docx-view-mode-spread").props.onClick();
  await waitFor(
    () => root.props["data-docx-view-mode"] === "spread",
    "spread mode applied",
  );
  assert.equal(
    findByTestId(mounted.root, "docx-page-current").value,
    "3",
    "switching modes must not rewind the current page",
  );
  await waitFor(
    () => scroll.scrolls.at(-1)?.top === 352,
    "the spread row containing page 3",
  );

  // Switching back to single page mode keeps page 3 as well.
  await findByTestId(mounted.root, "docx-view-mode-single").props.onClick();
  await waitFor(
    () => root.props["data-docx-view-mode"] === "single",
    "single mode applied",
  );
  assert.equal(findByTestId(mounted.root, "docx-page-current").value, "3");
  await waitFor(
    () => scroll.scrolls.at(-1)?.top === 680,
    "the single page row for page 3",
  );
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DOCX fit-page keeps the whole paper in the viewport", async (t) => {
  const originalClientWidth = Object.getOwnPropertyDescriptor(
    Object.prototype,
    "clientWidth",
  );
  const originalClientHeight = Object.getOwnPropertyDescriptor(
    Object.prototype,
    "clientHeight",
  );
  // A short, wide viewport: fit-width fills 200px with a 400px cap, while the
  // whole page can only grow to (400 - 64) / 300.
  Object.defineProperty(Object.prototype, "clientWidth", {
    configurable: true,
    get() {
      return 800;
    },
  });
  Object.defineProperty(Object.prototype, "clientHeight", {
    configurable: true,
    get() {
      return 400;
    },
  });
  t.after(() => {
    if (originalClientWidth) {
      Object.defineProperty(Object.prototype, "clientWidth", originalClientWidth);
    } else {
      delete Object.prototype.clientWidth;
    }
    if (originalClientHeight) {
      Object.defineProperty(Object.prototype, "clientHeight", originalClientHeight);
    } else {
      delete Object.prototype.clientHeight;
    }
  });

  const model = modelForViewer(2);
  const fitWidth = await mount(DocxViewerRoot, {
    model,
    layoutOptions: VIEWER_LAYOUT,
    fitWidth: true,
  });
  await vue.nextTick();
  const fitWidthWrapper = pageWrappers(viewerRoot(fitWidth.root))[0];
  fitWidth.app.unmount();

  const fitPage = await mount(DocxViewerRoot, {
    model,
    layoutOptions: VIEWER_LAYOUT,
    fitMode: "page",
  });
  await vue.nextTick();
  const root = viewerRoot(fitPage.root);
  const fitPageWrapper = pageWrappers(root)[0];

  assert.equal(root.props["data-docx-fit-mode"], "page");
  assert.equal(parseFloat(fitWidthWrapper.props.style.width), 400, "fit width caps at 200%");
  const fitPageWidth = parseFloat(fitPageWrapper.props.style.width);
  assert.ok(
    fitPageWidth > 200 && fitPageWidth < 400,
    `fit page (${fitPageWidth}px) is tighter than fit width so the page height fits`,
  );
  assert.deepEqual(fitPage.warnings, []);
  fitPage.app.unmount();
});
