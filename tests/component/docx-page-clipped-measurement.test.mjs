// Clipped page measurement in the rendered viewer (issue #135).
//
// The headless test renderer has no layout engine, so each test installs the
// geometry a browser would produce: a page surface that grew past the paper, a
// page body that clips its content, and paragraphs whose boxes run past the
// clipping ancestor. The components must report the visible height, not
// `offsetHeight`/`scrollHeight` of the inflated box.

import assert from "node:assert/strict";
import test from "node:test";

import { importFromDemo, mount, vue, walk } from "./vue-test-renderer.mjs";

const { DocxPageBody, DocxPageWrapper, DocxViewerRoot } = await importFromDemo(
  "@arcships/vue-docx"
);
const { cloneDocModel, defaultStarterModel, DOC_PAGE_BREAK_GAP } =
  await importFromDemo("@arcships/docx-core");

// Mirrors DocxViewerRoot's PAGE_TOP_PADDING: wrapper `top` is the page offset
// minus the scroll container's padding.
const PAGE_TOP_PADDING = 24;

const SECTION_PROPERTIES =
  '<w:sectPr><w:pgSz w:w="3000" w:h="4500"/>' +
  '<w:pgMar w:top="150" w:right="150" w:bottom="150" w:left="150"/></w:sectPr>';

// Page geometry that the viewer derives from the section properties above
// (3000x4500 twips = 200x300px, 150 twips margins = 10px).
const VIEWER_LAYOUT = {
  pageWidth: 200,
  pageHeight: 300,
  margin: 10,
  minLineHeight: 30,
  paragraphSpacing: 10,
};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function paragraph(text) {
  return { type: "paragraph", children: [{ type: "text", text }] };
}

function modelForViewer(count, prefix = "Clipped page item") {
  const model = cloneDocModel(defaultStarterModel);
  model.nodes = Array.from({ length: count }, (_, index) =>
    paragraph(`${prefix} ${index + 1}`)
  );
  model.metadata.sectionPropertiesXml = SECTION_PROPERTIES;
  return model;
}

function prop(node, name) {
  return node?.props?.[name];
}

function hasProp(node, name) {
  return prop(node, name) !== undefined;
}

/**
 * Give a host node the rect/box metrics a browser would report. The renderer's
 * nodes are plain objects, so the DOM surface the measurement code relies on is
 * added here: a rect, a parent link and an inline overflow style.
 */
function installGeometry(
  node,
  { rect, overflow, offsetHeight, scrollHeight, clientHeight }
) {
  Object.defineProperty(node, "parentElement", {
    configurable: true,
    get() {
      return this.parent ?? null;
    },
  });
  const style = overflow ? { overflow } : {};
  Object.defineProperty(node, "style", {
    configurable: true,
    get() {
      return style;
    },
  });
  node.getBoundingClientRect = () => ({
    ...rect,
    width: rect.right - rect.left,
    height: rect.bottom - rect.top,
  });
  if (offsetHeight !== undefined) node.offsetHeight = offsetHeight;
  if (scrollHeight !== undefined) node.scrollHeight = scrollHeight;
  if (clientHeight !== undefined) node.clientHeight = clientHeight;
}

/** ResizeObserver stub that records callbacks so a test can re-measure. */
function installResizeObserver() {
  const callbacks = [];
  const previous = globalThis.ResizeObserver;
  globalThis.ResizeObserver = class ResizeObserver {
    constructor(callback) {
      callbacks.push(callback);
    }
    observe() {}
    unobserve() {}
    disconnect() {}
  };
  return {
    fire() {
      for (const callback of [...callbacks]) callback([]);
    },
    restore() {
      if (previous === undefined) delete globalThis.ResizeObserver;
      else globalThis.ResizeObserver = previous;
    },
  };
}

async function settle() {
  await vue.nextTick();
  await vue.nextTick();
}

// ---------------------------------------------------------------------------
// Viewer: page offsets follow the visible page, not the grown box
// ---------------------------------------------------------------------------

test("clipped page content does not inflate the measured page height or the next page offset", async (t) => {
  const observer = installResizeObserver();
  t.after(() => observer.restore());

  const mounted = await mount(DocxViewerRoot, {
    model: modelForViewer(40),
    layoutOptions: VIEWER_LAYOUT,
  });
  await settle();

  const wrappers = walk(mounted.root).filter((node) =>
    hasProp(node, "data-docx-page-wrapper")
  );
  assert.ok(wrappers.length >= 2, "the test viewport mounts at least two pages");

  for (const wrapper of wrappers) {
    const pageIndex = Number(prop(wrapper, "data-docx-page-index"));
    // Distinct coordinate spaces keep per-page rects independent.
    const origin = 10_000 * pageIndex;
    const surface = walk(wrapper).find((node) =>
      hasProp(node, "data-docx-page-surface")
    );
    const body = walk(wrapper).find((node) =>
      hasProp(node, "data-docx-page-body")
    );
    assert.ok(surface, `page ${pageIndex} renders a surface`);
    assert.ok(body, `page ${pageIndex} renders a body`);

    // The buggy render let the page box grow to 400px while the visible paper
    // edge stayed at 300px; pagination followed the 400px box.
    installGeometry(surface, {
      rect: { top: origin, left: 0, right: 200, bottom: origin + 300 },
      overflow: "hidden",
      offsetHeight: 400,
    });
    installGeometry(body, {
      rect: {
        top: origin + 10,
        left: 10,
        right: 190,
        bottom: origin + 290,
      },
      overflow: "hidden",
      clientHeight: 280,
      scrollHeight: 1400,
    });
    for (const node of walk(body)) {
      if (!hasProp(node, "data-docx-paragraph-host")) continue;
      installGeometry(node, {
        rect: {
          top: origin + 20,
          left: 10,
          right: 190,
          bottom: origin + 400,
        },
      });
    }
  }

  observer.fire();
  await settle();

  const measured = walk(mounted.root).filter((node) =>
    hasProp(node, "data-docx-page-wrapper")
  );
  const pageTop = (pageIndex) => {
    const wrapper = measured.find(
      (node) => Number(prop(node, "data-docx-page-index")) === pageIndex
    );
    assert.ok(wrapper, `page ${pageIndex} is still mounted`);
    return Number.parseFloat(prop(wrapper, "style")?.top);
  };
  const pageHeight = (pageIndex) => {
    const wrapper = measured.find(
      (node) => Number(prop(node, "data-docx-page-index")) === pageIndex
    );
    return Number.parseFloat(prop(wrapper, "style")?.height);
  };

  assert.equal(
    pageHeight(0),
    VIEWER_LAYOUT.pageHeight,
    "a page whose content is clipped stays exactly one paper tall"
  );
  assert.equal(
    pageTop(1) - pageTop(0),
    VIEWER_LAYOUT.pageHeight + DOC_PAGE_BREAK_GAP,
    "the next page starts one paper plus the page-break gap below, not below the inflated box"
  );
  assert.equal(
    pageTop(0),
    0,
    "page offsets stay relative to the scroll container's padding"
  );
  assert.ok(PAGE_TOP_PADDING > 0);

  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

// ---------------------------------------------------------------------------
// Page body: content height follows the visible body
// ---------------------------------------------------------------------------

test("page body reports the visible content height instead of its scroll area", async () => {
  const model = modelForViewer(1);
  const segments = vue.shallowRef([{ nodeIndex: 0 }]);
  const measured = [];
  const Harness = vue.defineComponent({
    setup() {
      return () =>
        vue.h(DocxPageBody, {
          pageIndex: 0,
          pageNodeSegments: segments.value,
          pageLayout: {
            pageWidthPx: 816,
            pageHeightPx: 1056,
            marginsPx: { top: 72, right: 72, bottom: 72, left: 72 },
          },
          pageContentWidthPx: 672,
          model,
          pageNumber: 1,
          totalPages: 1,
          onContentMeasure: (heightPx) => measured.push(heightPx),
        });
    },
  });

  const mounted = await mount(Harness);
  await settle();
  assert.deepEqual(
    measured,
    [],
    "a body without a layout engine reports no phantom content height"
  );

  const body = walk(mounted.root).find((node) =>
    hasProp(node, "data-docx-page-body")
  );
  const paragraphHost = walk(mounted.root).find((node) =>
    hasProp(node, "data-docx-paragraph-host")
  );
  assert.ok(body, "the page body is rendered");
  assert.ok(paragraphHost, "the paragraph host is rendered");

  installGeometry(body, {
    rect: { top: 72, left: 72, right: 744, bottom: 984 },
    overflow: "hidden",
    clientHeight: 912,
    // What the old measurement reported: content clipped at the paper edge but
    // still counted inside the scroll area.
    scrollHeight: 1400,
  });
  installGeometry(paragraphHost, {
    rect: { top: 82, left: 82, right: 734, bottom: 1400 },
  });

  // Re-render the same page to trigger the post-flush measurement.
  segments.value = [...segments.value];
  await settle();

  assert.equal(
    measured.at(-1),
    912,
    "the body reports its visible height, not the 1400px scroll area"
  );
  assert.notEqual(measured.at(-1), 1400);

  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

// ---------------------------------------------------------------------------
// Page wrapper: clipped cell content is excluded, visible overflow still grows
// ---------------------------------------------------------------------------

function mountWrapper(measured) {
  const Harness = vue.defineComponent({
    setup() {
      return () =>
        vue.h(
          DocxPageWrapper,
          {
            page: { number: 1 },
            pageWidthPx: 816,
            pageHeightPx: 1056,
            onMeasure: (event) => measured.push(event.heightPx),
          },
          {
            default: () => [
              vue.h(
                "div",
                { class: "docx-test-cell", style: { overflow: "hidden" } },
                [vue.h("div", { class: "docx-test-paragraph" })]
              ),
            ],
          }
        );
    },
  });
  return mount(Harness);
}

test("page wrapper measures the visible page instead of its scroll area", async (t) => {
  const observer = installResizeObserver();
  t.after(() => observer.restore());

  const measured = [];
  const mounted = await mountWrapper(measured);
  await settle();

  const wrapper = walk(mounted.root).find((node) =>
    hasProp(node, "data-docx-page-wrapper")
  );
  const cell = walk(mounted.root).find((node) =>
    String(prop(node, "class") ?? "").includes("docx-test-cell")
  );
  const block = walk(mounted.root).find((node) =>
    String(prop(node, "class") ?? "").includes("docx-test-paragraph")
  );
  assert.ok(wrapper, "the wrapper is rendered");
  assert.ok(cell, "the clipped cell is rendered");
  assert.ok(block, "the clipped paragraph is rendered");

  installGeometry(wrapper, {
    rect: { top: 0, left: 0, right: 816, bottom: 1056 },
    offsetHeight: 1056,
    // A clipped cell still contributes to the wrapper's scroll area.
    scrollHeight: 1400,
  });
  installGeometry(cell, {
    rect: { top: 72, left: 72, right: 744, bottom: 300 },
    overflow: "hidden",
  });
  installGeometry(block, {
    rect: { top: 82, left: 82, right: 734, bottom: 1400 },
  });

  observer.fire();
  await settle();

  assert.equal(
    measured.at(-1),
    1056,
    "the wrapper reports one paper, not the clipped content's scroll area"
  );
  assert.notEqual(measured.at(-1), 1400);

  // Visible content taller than the paper still grows the measured page.
  installGeometry(cell, {
    rect: { top: 72, left: 72, right: 744, bottom: 1300 },
  });
  installGeometry(block, {
    rect: { top: 82, left: 82, right: 734, bottom: 1300 },
  });

  observer.fire();
  await settle();

  assert.equal(
    measured.at(-1),
    1300,
    "content that is not clipped still defines the page height"
  );

  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});
