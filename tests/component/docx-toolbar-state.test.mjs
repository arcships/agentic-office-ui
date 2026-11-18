import assert from "node:assert/strict";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

if (globalThis.Document === undefined) globalThis.Document = class Document {};
if (globalThis.ShadowRoot === undefined) globalThis.ShadowRoot = class ShadowRoot {};

const { DocxToolbar, DocxViewer, useDocxEditor } = await importFromDemo("@arcships/vue-docx");
const { createBlankDocumentModel } = await importFromDemo("@arcships/docx-core");

function styledModel() {
  return {
    ...createBlankDocumentModel(),
    nodes: [
      {
        type: "paragraph",
        style: { align: "center" },
        children: [{
          type: "text",
          text: "Styled",
          style: { bold: true, italic: true, underline: true, strike: true },
        }],
      },
      {
        type: "paragraph",
        style: { align: "left" },
        children: [{ type: "text", text: "Left aligned" }],
      },
    ],
  };
}

function classNames(node) {
  return String(node?.props?.class || "").split(/\s+/).filter(Boolean);
}

function hasClass(node, className) {
  return classNames(node).includes(className);
}

function selectOptionValues(root, testId) {
  const select = findByTestId(root, testId);
  assert.ok(select, `${testId} must render`);
  return walk(select)
    .filter((node) => node.type === "option")
    .map((node) => Number(node.props?.value));
}

function findComponentType(vnode, name) {
  if (!vnode || typeof vnode !== "object") return undefined;
  if (vnode.type?.__name === name || vnode.type?.name === name) return vnode.type;
  for (const child of vnode.children ?? []) {
    const found = findComponentType(child, name);
    if (found) return found;
  }
  if (vnode.component?.subTree) return findComponentType(vnode.component.subTree, name);
  return undefined;
}

function typePage(input, value) {
  input.value = value;
  input.dispatchEvent({ type: "input", target: input });
}

function pressPageKey(input, key) {
  const handlers = Array.isArray(input.props.onKeydown)
    ? input.props.onKeydown
    : [input.props.onKeydown];
  for (const handler of handlers) handler({ key, preventDefault() {} });
}

test("DocxToolbar keeps a gesture zoom value selectable and exposes toggle state", async () => {
  let controller;
  const zoomPercent = vue.ref(100);
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: styledModel() });
      return () => vue.h(DocxToolbar, {
        controller,
        zoom: zoomPercent.value,
        currentPage: 1,
        totalPages: 1,
      });
    },
  });

  const mounted = await mount(Harness);
  controller.selectParagraph(0);
  await vue.nextTick();

  zoomPercent.value = 113;
  await vue.nextTick();
  const zoom = findByTestId(mounted.root, "editor-zoom");
  assert.ok(zoom);
  assert.equal(zoom.props.value, 113, "the select must keep the current gesture zoom");
  assert.ok(selectOptionValues(mounted.root, "editor-zoom").includes(113));

  const bold = findByTestId(mounted.root, "editor-bold");
  assert.ok(bold);
  assert.equal(bold.props["aria-label"], "Bold");
  assert.equal(bold.props["aria-pressed"], true);
  assert.equal(hasClass(bold, "active"), true);

  for (const testId of ["editor-italic", "editor-underline", "editor-strike"]) {
    const toggle = findByTestId(mounted.root, testId);
    assert.ok(toggle, `${testId} must render`);
    assert.equal(toggle.props["aria-pressed"], true, `${testId} must expose its active state`);
    assert.equal(hasClass(toggle, "active"), true);
  }

  const center = findByTestId(mounted.root, "editor-align-center");
  const left = findByTestId(mounted.root, "editor-align-left");
  assert.ok(center);
  assert.ok(left);
  assert.equal(center.props["aria-label"], "Align center");
  assert.equal(center.props["aria-pressed"], true);
  assert.equal(hasClass(center, "active"), true);
  assert.equal(left.props["aria-pressed"], false);
  assert.equal(hasClass(left, "active"), false);

  controller.selectParagraph(1);
  await vue.nextTick();
  const leftAfterCaretMove = findByTestId(mounted.root, "editor-align-left");
  const centerAfterCaretMove = findByTestId(mounted.root, "editor-align-center");
  assert.equal(leftAfterCaretMove.props["aria-pressed"], true);
  assert.equal(hasClass(leftAfterCaretMove, "active"), true);
  assert.equal(centerAfterCaretMove.props["aria-pressed"], false);
  assert.equal(hasClass(centerAfterCaretMove, "active"), false);

  const link = findByTestId(mounted.root, "editor-link");
  assert.ok(link);
  assert.equal(link.props["aria-label"], "Link");
  assert.equal(link.props["aria-pressed"], false);

  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DocxViewer shows gesture zoom in its toolbar and announces search status", async () => {
  const mounted = await mount(DocxViewer, {
    model: styledModel(),
    defaultZoom: 100,
  });

  const scrollContainer = walk(mounted.root).find((node) =>
    classNames(node).includes("docx-viewer-root")
  );
  assert.ok(scrollContainer, "the viewer root must expose its scroll container");
  scrollContainer.dispatchEvent({
    type: "wheel",
    ctrlKey: true,
    deltaY: -100,
    deltaMode: 0,
    clientX: 0,
    clientY: 0,
    preventDefault() {},
  });
  await vue.nextTick();
  await vue.nextTick();

  const zoom = findByTestId(mounted.root, "docx-zoom-select");
  assert.ok(zoom);
  const values = selectOptionValues(mounted.root, "docx-zoom-select");
  assert.ok(values.includes(zoom.props.value), "the select value must have a matching option");
  assert.notEqual(zoom.props.value, 100, "gesture zoom must update the select");

  const status = findByTestId(mounted.root, "docx-search-status");
  assert.ok(status);
  assert.equal(status.props.role, "status");
  assert.equal(status.props["aria-live"], "polite");
  assert.equal(status.props["aria-atomic"], "true");

  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DocxViewerToolbar keeps a focused page draft while currentPage changes", async () => {
  const seed = await mount(DocxViewer, { model: styledModel() });
  const Toolbar = findComponentType(seed.app._instance?.subTree, "DocxViewerToolbar");
  assert.ok(Toolbar, "the viewer must render its toolbar internally");
  seed.app.unmount();

  const currentPage = vue.ref(10);
  const selectedPages = [];
  const Harness = vue.defineComponent({
    setup() {
      return () => vue.h(Toolbar, {
        canDownload: false,
        disabled: false,
        fileName: "page.docx",
        isDark: false,
        sidebarOpen: false,
        showTrackedChanges: true,
        showComments: true,
        trackedChangeCount: 0,
        commentCount: 0,
        currentPage: currentPage.value,
        totalPages: 300,
        zoom: 100,
        searchQuery: "",
        searchResultCount: 0,
        searchResultIndex: -1,
        onSelectPage: (page) => selectedPages.push(page),
      });
    },
  });

  const mounted = await mount(Harness);
  let input = findByTestId(mounted.root, "docx-page-current");
  assert.ok(input);
  assert.equal(input.value, "10");

  input.props.onFocus();
  typePage(input, "120");
  await vue.nextTick();
  currentPage.value = 200;
  await vue.nextTick();

  input = findByTestId(mounted.root, "docx-page-current");
  assert.equal(
    input.value,
    "120",
    "a focused page draft must not be overwritten by currentPage changes",
  );

  pressPageKey(input, "Enter");
  await vue.nextTick();
  assert.deepEqual(selectedPages, [120], "Enter commits the focused draft");

  input = findByTestId(mounted.root, "docx-page-current");
  typePage(input, "130");
  currentPage.value = 240;
  await vue.nextTick();
  input.props.onBlur();
  await vue.nextTick();
  assert.deepEqual(selectedPages, [120, 130], "blur commits the pending draft");

  input = findByTestId(mounted.root, "docx-page-current");
  input.props.onFocus();
  typePage(input, "140");
  await vue.nextTick();
  pressPageKey(input, "Escape");
  await vue.nextTick();
  assert.equal(
    findByTestId(mounted.root, "docx-page-current").value,
    "240",
    "Escape restores the current page instead of committing the draft",
  );
  assert.deepEqual(selectedPages, [120, 130]);

  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});
