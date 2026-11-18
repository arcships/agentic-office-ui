// DOCX bookmark navigation (issue #148).
//
// Internal `#bookmark` links used to render as anchors whose click did nothing:
// the run renderer had an `onInternalLinkClick` hook but both hosts passed
// `undefined`, and static HTML strips inline handlers anyway. The coverage below
// keeps every link in the chain alive:
//
//   1. clicking a TOC entry reveals its bookmark page + target paragraph ... surface click
//   2. the revealed paragraph is highlighted and takes the caret .......... surface click
//   3. `revealBookmark(name)` resolves a target ............................. exposed API
//   4. a missing bookmark shows a non-blocking notice ...................... exposed API
//   5. read-only paragraph/table hosts forward internal link clicks ........ host wiring
//   6. `DocxViewer` exposes the same navigation ............................ viewer API

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
  DocxDocumentSurface,
  DocxParagraphHost,
  DocxTableHost,
  DocxViewer,
  useDocxEditor,
} = await importFromDemo("@arcships/vue-docx");
const {
  createBlankDocumentModel,
} = await importFromDemo("@arcships/docx-core");

const SECTION_PROPERTIES =
  '<w:sectPr><w:pgSz w:w="3000" w:h="4500"/>' +
  '<w:pgMar w:top="150" w:right="150" w:bottom="150" w:left="150"/></w:sectPr>';
// Page geometry the viewer derives from the section properties: about seven
// short paragraphs per page, so a late target is genuinely off-screen.
const VIEWER_LAYOUT = {
  pageWidth: 200,
  pageHeight: 300,
  margin: 10,
  minLineHeight: 30,
  paragraphSpacing: 10,
};

function paragraph(text) {
  return { type: "paragraph", children: [{ type: "text", text }] };
}

function linkedParagraph(text, href) {
  return { type: "paragraph", children: [{ type: "text", text, link: href }] };
}

function bookmarkParagraph(name, text) {
  return {
    type: "paragraph",
    sourceXml:
      `<w:p><w:bookmarkStart w:id="7" w:name="${name}"/>` +
      `<w:r><w:t>${text}</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>`,
    children: [{ type: "text", text }],
  };
}

function bookmarkModel() {
  const nodes = Array.from({ length: 40 }, (_, index) => {
    if (index === 0) return linkedParagraph("Jump to target", "#Target_Section");
    if (index === 30) return bookmarkParagraph("Target_Section", "Target paragraph");
    return paragraph(`Body paragraph ${index}`);
  });
  const model = { ...createBlankDocumentModel(), nodes };
  model.metadata.sectionPropertiesXml = SECTION_PROPERTIES;
  return model;
}

function paragraphHost(root, nodeIndex) {
  return walk(root).find(
    (node) =>
      node.props?.["data-docx-paragraph-host"] === true &&
      String(node.props?.["data-docx-paragraph-node-index"]) === String(nodeIndex),
  );
}

function classNames(node) {
  return String(node?.props?.class || "").split(/\s+/).filter(Boolean);
}

function hasClass(node, className) {
  return classNames(node).includes(className);
}

/**
 * A cancelable click on an internal link. The hosts find the anchor through
 * `Element.closest`, which the shared host renderer does not implement, so the
 * fake target mirrors just that API.
 */
function internalLinkClick(href) {
  const anchor = {
    closest(selector) {
      return selector === "a[href]" ? anchor : null;
    },
    getAttribute(name) {
      return name === "href" ? href : null;
    },
  };
  let prevented = 0;
  let stopped = 0;
  return {
    event: {
      target: {
        closest(selector) {
          return anchor.closest(selector);
        },
      },
      preventDefault() {
        prevented += 1;
      },
      stopPropagation() {
        stopped += 1;
      },
    },
    prevented: () => prevented,
    stopped: () => stopped,
  };
}

function fakeTargetElement() {
  return {
    scrollIntoViewCalls: [],
    focusCalls: 0,
    scrollIntoView(options) {
      this.scrollIntoViewCalls.push(options);
    },
    focus() {
      this.focusCalls += 1;
    },
  };
}

async function mountSurface(model = bookmarkModel()) {
  let controller;
  const surfaceRef = vue.ref();
  const requestedPages = [];
  const missingNames = [];
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model });
      return () =>
        vue.h(DocxDocumentSurface, {
          ref: surfaceRef,
          model: controller.model,
          controller,
          editable: true,
          layoutOptions: VIEWER_LAYOUT,
          onRequestPageReveal: (pageIndex) => requestedPages.push(pageIndex),
          onBookmarkMissing: (name) => missingNames.push(name),
        });
    },
  });
  const mounted = await mount(Harness);
  await vue.nextTick();
  return {
    ...mounted,
    get controller() {
      return controller;
    },
    get surface() {
      return surfaceRef.value;
    },
    requestedPages,
    missingNames,
  };
}

test("clicking an internal link scrolls to the bookmark, highlights it and takes the caret", async (t) => {
  const mounted = await mountSurface();
  t.after(() => mounted.app.unmount());
  const { controller, surface } = mounted;
  assert.ok(surface, "the surface exposes its navigation API");

  const container = surface.scrollContainer;
  assert.ok(container, "the viewer root is the scroll container");
  container.scrollTop = 0;
  const targetElement = fakeTargetElement();
  container.querySelector = (selector) =>
    selector.includes('[data-docx-paragraph-node-index="30"]') ? targetElement : null;

  const tocHost = paragraphHost(mounted.root, 0);
  assert.ok(tocHost, "the TOC paragraph renders");
  assert.ok(hasClass(tocHost, "docx-paragraph-host"));

  const click = internalLinkClick("#Target_Section");
  tocHost.props.onClick(click.event);

  assert.equal(click.prevented(), 1, "the browser's fragment jump is cancelled");
  assert.equal(click.stopped(), 1, "the surface reference click does not also run");

  await waitFor(
    () =>
      controller.selection.kind === "paragraph" &&
      controller.selection.nodeIndex === 30,
    "the bookmarked paragraph becomes the selection/caret target",
  );

  await waitFor(
    () => hasClass(paragraphHost(mounted.root, 30), "docx-paragraph-host--bookmark-target"),
    "the revealed paragraph is highlighted",
  );
  assert.equal(paragraphHost(mounted.root, 0)?.props?.class?.includes("bookmark-target"), false);

  assert.ok(targetElement.focusCalls > 0, "the editable target takes focus/caret");
  assert.deepEqual(
    targetElement.scrollIntoViewCalls.at(-1),
    { block: "center", behavior: "smooth" },
  );
  assert.deepEqual(mounted.warnings, []);
});

test("revealBookmark resolves a target through the exposed surface API", async (t) => {
  const model = {
    ...createBlankDocumentModel(),
    nodes: Array.from({ length: 200 }, (_, index) =>
      index === 190
        ? bookmarkParagraph("Target_Section", "Target paragraph")
        : paragraph(`Body paragraph ${index}`)
    ),
  };
  model.metadata.sectionPropertiesXml = SECTION_PROPERTIES;

  const surfaceRef = vue.ref();
  const requestedPages = [];
  const Harness = vue.defineComponent({
    setup() {
      return () =>
        vue.h(DocxDocumentSurface, {
          ref: surfaceRef,
          model,
          editable: false,
          layoutOptions: VIEWER_LAYOUT,
          onRequestPageReveal: (pageIndex) => requestedPages.push(pageIndex),
        });
    },
  });
  const mounted = await mount(Harness);
  t.after(() => mounted.app.unmount());
  await vue.nextTick();

  const surface = surfaceRef.value;
  const result = await surface.revealBookmark("target_section");
  assert.equal(result.status, "revealed");
  assert.equal(result.name, "Target_Section");
  assert.equal(result.nodeIndex, 190);
  assert.equal(result.text, "Target paragraph");
  assert.ok(result.pageIndex > 0, "navigation knows which page lays out the target");
  assert.ok(
    requestedPages.includes(result.pageIndex),
    "the host is asked to reveal the target page",
  );
  assert.ok(surface.scrollContainer.scrollTop > 0, "the viewer scrolls to the target page");
  await waitFor(
    () => hasClass(paragraphHost(mounted.root, 190), "docx-paragraph-host--bookmark-target"),
    "the mounted target page is highlighted",
  );
  assert.deepEqual(mounted.warnings, []);
});

test("a table-cell bookmark targets the cell paragraph before the table host", async (t) => {
  const table = {
    type: "table",
    rows: [
      {
        type: "table-row",
        cells: [
          { type: "table-cell", nodes: [paragraph("left")] },
          {
            type: "table-cell",
            nodes: [bookmarkParagraph("Cell_Target", "Cell target")],
          },
        ],
      },
    ],
  };
  const model = { ...createBlankDocumentModel(), nodes: [table] };
  const mounted = await mountSurface(model);
  t.after(() => mounted.app.unmount());

  const container = mounted.surface.scrollContainer;
  const queries = [];
  const cellElement = fakeTargetElement();
  container.querySelector = (selector) => {
    queries.push(selector);
    return selector.includes("[data-docx-table-cell-paragraph-host") ? cellElement : null;
  };

  const result = await mounted.surface.revealBookmark("Cell_Target");
  assert.equal(result.status, "revealed");
  assert.equal(result.nodeIndex, 0);
  assert.equal(result.tableRowIndex, 0);
  assert.equal(result.tableCellIndex, 1);
  assert.match(queries[0], /data-docx-table-cell-paragraph-host/);
  assert.match(queries[0], /data-docx-table-row-index="0"/);
  assert.match(queries[0], /data-docx-table-cell-index="1"/);
  assert.ok(cellElement.focusCalls > 0, "the cell paragraph takes the caret");
  assert.deepEqual(
    cellElement.scrollIntoViewCalls.at(-1),
    { block: "center", behavior: "smooth" },
  );

  const tableHost = walk(mounted.root).find(
    (node) => node.props?.["data-docx-table-host"] === true,
  );
  assert.ok(hasClass(tableHost, "docx-table-host--bookmark-target"));
  const targetCell = walk(mounted.root).find(
    (node) =>
      node.props?.["data-docx-table-cell"] === true &&
      String(node.props?.["data-docx-table-row-index"]) === "0" &&
      String(node.props?.["data-docx-table-cell-index"]) === "1",
  );
  assert.ok(hasClass(targetCell, "docx-table-cell--bookmark-target"));
  assert.deepEqual(mounted.warnings, []);
});

test("a missing bookmark shows a non-blocking notice instead of throwing", async (t) => {
  const mounted = await mountSurface();
  t.after(() => mounted.app.unmount());

  const result = await mounted.surface.revealBookmark("#Not_There");
  assert.deepEqual(result, { status: "missing", name: "Not_There" });
  await vue.nextTick();

  const notice = findByTestId(mounted.root, "docx-bookmark-notice");
  assert.ok(notice, "the notice renders");
  assert.match(textContent(notice), /Not_There/);
  assert.match(textContent(notice), /not found/i);
  assert.deepEqual(mounted.missingNames, ["Not_There"]);
  // A missing target must not move the caret off the current paragraph.
  assert.equal(mounted.controller.selection.kind, "paragraph");
  assert.equal(mounted.controller.selection.nodeIndex, 0);
  assert.deepEqual(mounted.warnings, []);
});

test("read-only paragraph and table hosts forward internal link clicks", async (t) => {
  const calls = [];
  const paragraphHosts = await mount(DocxParagraphHost, {
    paragraph: linkedParagraph("Jump", "#Read_Only_Target"),
    paragraphIndex: 0,
    editable: false,
    onInternalLinkClick: (name) => calls.push(name),
  });
  t.after(() => paragraphHosts.app.unmount());

  const host = walk(paragraphHosts.root).find(
    (node) => node.props?.["data-docx-paragraph-host"] === true,
  );
  const paragraphClick = internalLinkClick("#Read_Only_Target");
  host.props.onClick(paragraphClick.event);
  assert.deepEqual(calls, ["read_only_target"]);
  assert.equal(paragraphClick.prevented(), 1);
  assert.equal(paragraphClick.stopped(), 1);

  const tableCalls = [];
  const table = {
    type: "table",
    rows: [
      {
        type: "table-row",
        cells: [
          {
            type: "table-cell",
            nodes: [linkedParagraph("Cell jump", "#Cell_Target")],
          },
        ],
      },
    ],
  };
  const tableHosts = await mount(DocxTableHost, {
    table,
    tableIndex: 0,
    editable: false,
    onInternalLinkClick: (name) => tableCalls.push(name),
  });
  t.after(() => tableHosts.app.unmount());

  const cell = walk(tableHosts.root).find(
    (node) => node.props?.["data-docx-table-cell"] === true,
  );
  assert.ok(cell, "the cell renders");
  const cellClick = internalLinkClick("#Cell_Target");
  cell.props.onClick(cellClick.event);
  assert.deepEqual(tableCalls, ["cell_target"]);
  assert.equal(cellClick.prevented(), 1);
  // `@click.stop` plus the handler's own stopPropagation both fire.
  assert.ok(cellClick.stopped() >= 1, "the cell link click stops at the cell");
});

test("DocxViewer exposes bookmark navigation for embedded read-only surfaces", async (t) => {
  const model = bookmarkModel();
  const viewerRef = vue.ref();
  const Harness = vue.defineComponent({
    setup() {
      return () =>
        vue.h(DocxViewer, {
          ref: viewerRef,
          model,
          showToolbar: false,
        });
    },
  });
  const mounted = await mount(Harness);
  t.after(() => mounted.app.unmount());
  await vue.nextTick();

  const result = await viewerRef.value.revealBookmark("#Target_Section");
  assert.equal(result.status, "revealed");
  assert.equal(result.nodeIndex, 30);
  assert.ok(result.pageIndex > 0);
  assert.deepEqual(mounted.warnings, []);
});
