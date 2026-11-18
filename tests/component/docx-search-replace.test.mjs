import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  textContent,
  vue,
  waitFor,
} from "./vue-test-renderer.mjs";

const {
  DocxEditor,
  DocxToolbar,
  docxReplaceTargetFromMatch,
  canReplaceDocxSearchMatch,
  findDocxSearchMatches,
  useDocxEditor,
  useDocxFindReplace,
} = await importFromDemo("@arcships/vue-docx");
const {
  createBlankDocumentModel,
  importDocxBuffer,
  paragraphText,
  serializeDocx,
} = await importFromDemo("@arcships/docx-core");
const { OfficeFindBar } = await importFromDemo("@arcships/vue-ui");

function paragraph(text, style) {
  return {
    type: "paragraph",
    children: [{ type: "text", text, ...(style ? { style } : undefined) }],
  };
}

function modelWithStory() {
  const base = createBlankDocumentModel();
  return {
    ...base,
    nodes: [
      paragraph("alpha needle beta"),
      {
        type: "table",
        rows: [{
          type: "table-row",
          cells: [{ type: "table-cell", nodes: [paragraph("cell needle")] }],
        }],
      },
      paragraph("needle omega"),
    ],
    metadata: {
      ...base.metadata,
      headerSections: [{
        partName: "word/header1.xml",
        referenceType: "default",
        nodes: [paragraph("header needle")],
      }],
      footnotes: [{ id: 1, text: "note needle" }],
    },
  };
}

async function mountController(model) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model });
      return () => vue.h("div", { "data-testid": "controller-harness" });
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, get controller() { return controller; } };
}

function documentText(model) {
  return model.nodes
    .map((node) => node.type === "paragraph"
      ? paragraphText(node)
      : node.rows.flatMap((row) => row.cells.map((cell) => paragraphText(cell.nodes[0]))).join("|"))
    .join("\n");
}

test("replaceMatch rewrites one match and undo restores it", async () => {
  const { app, controller } = await mountController(modelWithStory());
  try {
    const matches = findDocxSearchMatches(controller.model, "needle");
    assert.equal(matches.length, 5);
    const target = docxReplaceTargetFromMatch(matches[1]);
    assert.ok(target, "table-cell results are replaceable");

    assert.equal(controller.replaceMatch(target, "pin"), true);
    assert.equal(
      paragraphText(controller.model.nodes[1].rows[0].cells[0].nodes[0]),
      "cell pin",
      "only the targeted match changes",
    );
    assert.equal(paragraphText(controller.model.nodes[0]), "alpha needle beta");
    assert.equal(controller.canUndo, true);

    controller.undo();
    await vue.nextTick();
    assert.equal(
      paragraphText(controller.model.nodes[1].rows[0].cells[0].nodes[0]),
      "cell needle",
    );
    assert.equal(controller.canUndo, false, "one replace is one undo entry");

    // Replacing with identical text is not a new revision.
    assert.equal(controller.replaceMatch(target, "needle"), false);
    assert.equal(controller.canUndo, false);
  } finally {
    app.unmount();
  }
});

test("replaceMatch refuses read-only stories instead of rewriting them", async () => {
  const { app, controller } = await mountController(modelWithStory());
  try {
    const [headerMatch, footnoteMatch] = ["header", "footnote"].map((region) =>
      findDocxSearchMatches(controller.model, "needle").find((match) => match.region === region));
    assert.ok(headerMatch && footnoteMatch);
    assert.equal(canReplaceDocxSearchMatch(headerMatch), false);
    assert.equal(canReplaceDocxSearchMatch(footnoteMatch), false);
    assert.equal(docxReplaceTargetFromMatch(headerMatch), undefined);

    const before = controller.model;
    assert.equal(
      controller.replaceMatch({
        region: "header",
        location: { kind: "paragraph", nodeIndex: 0 },
        start: 7,
        end: 13,
      }, "pin"),
      false,
    );
    assert.equal(controller.model, before, "a rejected target must not touch the model");
    assert.match(controller.status, /header, footer or note/u);
  } finally {
    app.unmount();
  }
});

test("replaceAll rewrites every editable match as one undo entry and reports skipped stories", async () => {
  const { app, controller } = await mountController(modelWithStory());
  try {
    const summary = controller.replaceAll("needle", "pin");
    assert.deepEqual(summary, { replaced: 3, skipped: 2 });
    assert.equal(
      documentText(controller.model),
      "alpha pin beta\ncell pin\npin omega",
    );
    assert.equal(
      controller.model.metadata.headerSections[0].nodes[0].children[0].text,
      "header needle",
      "header text stays untouched",
    );
    assert.equal(controller.model.metadata.footnotes[0].text, "note needle");

    controller.undo();
    await vue.nextTick();
    assert.equal(documentText(controller.model), "alpha needle beta\ncell needle\nneedle omega");
    assert.equal(
      controller.canUndo,
      false,
      "replace-all records exactly one undo entry no matter how many paragraphs it touches",
    );

    controller.redo();
    await vue.nextTick();
    assert.equal(
      paragraphText(controller.model.nodes[0]),
      "alpha pin beta",
      "redo re-applies the whole batch",
    );
  } finally {
    app.unmount();
  }
});

test("replaceAll honours case and whole-word options", async () => {
  const base = createBlankDocumentModel();
  const { app, controller } = await mountController({
    ...base,
    nodes: [paragraph("Needle needle needles")],
  });
  try {
    // Case-insensitive is the default, like Excel: every occurrence counts,
    // including the one inside "needles".
    assert.deepEqual(controller.replaceAll("needle", "pin"), { replaced: 3, skipped: 0 });
    assert.equal(paragraphText(controller.model.nodes[0]), "pin pin pins");
    controller.undo();
    await vue.nextTick();

    assert.deepEqual(
      controller.replaceAll("needle", "pin", { caseSensitive: true }),
      { replaced: 2, skipped: 0 },
    );
    assert.equal(paragraphText(controller.model.nodes[0]), "Needle pin pins");
    controller.undo();
    await vue.nextTick();

    assert.deepEqual(
      controller.replaceAll("needle", "pin", { wholeWord: true }),
      { replaced: 2, skipped: 0 },
    );
    assert.equal(paragraphText(controller.model.nodes[0]), "pin pin needles");
  } finally {
    app.unmount();
  }
});

test("a replaced document survives export and reopen with its run styles", async () => {
  const base = createBlankDocumentModel();
  const model = {
    ...base,
    nodes: [
      {
        type: "paragraph",
        children: [
          { type: "text", text: "before " },
          { type: "text", text: "needle", style: { bold: true } },
          { type: "text", text: " after" },
        ],
      },
    ],
  };
  const { app, controller } = await mountController(model);
  try {
    assert.deepEqual(controller.replaceAll("needle", "pin"), { replaced: 1, skipped: 0 });
    const bytes = await serializeDocx(controller.model);
    const reopened = await importDocxBuffer(bytes, { useWorker: false });
    const text = paragraphText(reopened.model.nodes[0]);
    assert.equal(text, "before pin after", "the exported document carries the replacement");
    assert.doesNotMatch(text, /needle/u);

    const boldRun = reopened.model.nodes[0].children
      .find((child) => child.type === "text" && child.text === "pin");
    assert.equal(boldRun?.style?.bold, true, "the replaced text keeps the matched run's style");
  } finally {
    app.unmount();
  }
});

test("replace all on a real fixture survives export and reopen", async () => {
  const bytes = readFileSync(
    new URL("../../apps/demo/public/samples/demo.docx", import.meta.url),
  );
  const imported = await importDocxBuffer(
    bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    { useWorker: false },
  );
  const firstBodyText = imported.model.nodes
    .filter((node) => node.type === "paragraph")
    .map((node) => paragraphText(node))
    .find((text) => /\b[A-Za-z]{5,}\b/u.test(text));
  assert.ok(firstBodyText, "the fixture must have a paragraph with a replaceable word");
  const token = firstBodyText.match(/\b[A-Za-z]{5,}\b/u)[0];

  const { app, controller } = await mountController(imported.model);
  try {
    const summary = controller.replaceAll(token, "REPLACED", { caseSensitive: true });
    assert.ok(summary.replaced > 0, "the fixture token appears in the body");

    const exported = await serializeDocx(controller.model);
    const reopened = await importDocxBuffer(exported, { useWorker: false });
    const bodyText = reopened.model.nodes
      .filter((node) => node.type === "paragraph")
      .map((node) => paragraphText(node))
      .join("\n");
    assert.doesNotMatch(bodyText, new RegExp(`\\b${token}\\b`, "u"));
    assert.match(bodyText, /REPLACED/u);
  } finally {
    app.unmount();
  }
});

// ── Find bar UI ────────────────────────────────────────────────────
function installDomEnvironment(t) {
  const originalWindow = globalThis.window;
  const originalDocument = globalThis.document;
  const originalResizeObserver = globalThis.ResizeObserver;
  const originalRect = Object.getOwnPropertyDescriptor(Object.prototype, "getBoundingClientRect");
  const originalClientHeight = Object.getOwnPropertyDescriptor(Object.prototype, "clientHeight");
  const originalQuerySelectorAll = Object.getOwnPropertyDescriptor(Object.prototype, "querySelectorAll");

  class FakeElement {
    constructor(tagName) {
      this.tagName = String(tagName).toLowerCase();
      this.children = [];
      this.style = { cssText: "", setProperty(name, value) { this[name] = value; } };
    }
    appendChild(child) { this.children.push(child); return child; }
    setAttribute() {}
    getContext() { return { measureText: (text) => ({ width: String(text).length * 8 }) }; }
    querySelector() { return null; }
    querySelectorAll() { return []; }
  }

  globalThis.window = {
    addEventListener() {},
    removeEventListener() {},
    getSelection() { return null; },
  };
  globalThis.document = {
    addEventListener() {},
    removeEventListener() {},
    body: { append() {} },
    createElement(tagName) { return new FakeElement(tagName); },
    createTextNode(text) { return { textContent: String(text) }; },
    getSelection() { return null; },
    execCommand() { return false; },
  };
  globalThis.ResizeObserver = class ResizeObserver {
    observe() {}
    disconnect() {}
  };
  Object.defineProperty(Object.prototype, "getBoundingClientRect", {
    configurable: true,
    value: () => ({ height: 1056, width: 816, top: 0, left: 0, right: 816, bottom: 1056 }),
  });
  Object.defineProperty(Object.prototype, "querySelectorAll", {
    configurable: true,
    value: () => [],
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
    if (originalQuerySelectorAll) Object.defineProperty(Object.prototype, "querySelectorAll", originalQuerySelectorAll);
    else delete Object.prototype.querySelectorAll;
  });
}

async function mountEditor(t) {
  installDomEnvironment(t);
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({
        starterModel: {
          ...createBlankDocumentModel(),
          nodes: [paragraph("alpha needle beta"), paragraph("needle omega")],
        },
      });
      return () => vue.h(DocxEditor, {
        editor: controller,
        editable: true,
        showToolbar: true,
        showThumbnails: false,
      });
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => findByTestId(mounted.root, "editor-find-toggle"),
    "the editor toolbar renders its find toggle",
  );
  return { ...mounted, get controller() { return controller; } };
}

function inputEvent(value) {
  return { target: { value } };
}

test("the editor find bar opens, searches and replaces all through the toolbar", async (t) => {
  const mounted = await mountEditor(t);
  try {
    findByTestId(mounted.root, "editor-find-toggle").props.onClick();
    await vue.nextTick();

    const searchInput = findByTestId(mounted.root, "docx-editor-search-input");
    assert.ok(searchInput, "the find field renders inside the editor toolbar");
    searchInput.props.onInput(inputEvent("needle"));
    // Headless mounts have no DOM ranges, so highlight activation cannot place
    // the first result; the result list itself still drives the replace half.
    await waitFor(
      () => findByTestId(mounted.root, "office-find-replace-all").props.disabled === false,
      "the search reports results to the find bar",
    );

    const replaceInput = findByTestId(mounted.root, "office-find-replace-input");
    assert.ok(replaceInput, "the replace field renders next to the query");
    replaceInput.props.onInput(inputEvent("pin"));
    await vue.nextTick();

    const replaceAll = findByTestId(mounted.root, "office-find-replace-all");
    assert.equal(replaceAll.props.disabled, false, "replace all is available while results exist");
    replaceAll.props.onClick();
    await waitFor(
      () => paragraphText(mounted.controller.model.nodes[0]) === "alpha pin beta",
      "replace all rewrites the document",
    );
    assert.equal(paragraphText(mounted.controller.model.nodes[1]), "pin omega");
    await waitFor(
      () => findByTestId(mounted.root, "office-find-replace-all").props.disabled === true,
      "the result list follows the replaced text",
    );

    mounted.controller.undo();
    await vue.nextTick();
    assert.equal(paragraphText(mounted.controller.model.nodes[0]), "alpha needle beta");

    findByTestId(mounted.root, "office-find-close").props.onClick();
    await vue.nextTick();
    assert.equal(findByTestId(mounted.root, "docx-editor-search-input"), undefined);
  } finally {
    mounted.app.unmount();
  }
});

test("Ctrl/Cmd+F opens the find bar and Escape closes it", async (t) => {
  const mounted = await mountEditor(t);
  try {
    const root = findByTestId(mounted.root, "docx-editor");
    root.props.onKeydown({
      defaultPrevented: false,
      isComposing: false,
      key: "f",
      ctrlKey: true,
      metaKey: false,
      shiftKey: false,
      target: null,
      preventDefault() {},
    });
    await vue.nextTick();
    assert.ok(findByTestId(mounted.root, "docx-editor-search-input"));

    root.props.onKeydown({
      defaultPrevented: false,
      isComposing: false,
      key: "Escape",
      ctrlKey: false,
      metaKey: false,
      shiftKey: false,
      target: null,
      preventDefault() {},
    });
    await vue.nextTick();
    assert.equal(findByTestId(mounted.root, "docx-editor-search-input"), undefined);
  } finally {
    mounted.app.unmount();
  }
});

test("Replace in the find bar rewrites the active result", async () => {
  // Surface activation needs real DOM ranges, which the headless renderer does
  // not have; drive the bar with a stub surface that reports an active result
  // so the Replace button exercises the same composable + controller path.
  const model = {
    ...createBlankDocumentModel(),
    nodes: [paragraph("alpha needle beta")],
  };
  let controller;
  let find;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model });
      find = useDocxFindReplace({
        controller: () => controller,
        surface: () => ({
          search: async (query) => {
            const matches = findDocxSearchMatches(controller.model, query);
            find.syncSearchState({ status: "ready", query, matches, activeIndex: 0 });
            return { status: "ready", query, matches, activeIndex: 0 };
          },
          refreshSearch: async () => ({
            status: "ready",
            query: find.state.value.query,
            matches: findDocxSearchMatches(controller.model, find.state.value.query),
            activeIndex: 0,
          }),
          searchNext: async () => {},
          searchPrevious: async () => {},
          clearSearch: () => {},
          getSearchState: () => ({
            status: "idle",
            query: "",
            matches: [],
            activeIndex: -1,
          }),
        }),
      });
      return () => vue.h(DocxToolbar, {
        controller,
        findOpen: find.state.value.open,
        onToggleFind: find.toggleFind,
      }, {
        find: () => find.state.value.open
          ? vue.h(OfficeFindBar, {
            query: find.state.value.query,
            status: find.state.value.status,
            activeIndex: find.state.value.activeIndex,
            resultCount: find.state.value.matches.length,
            inputTestId: "docx-editor-search-input",
            showReplace: true,
            replacement: find.state.value.replacement,
            replaceDisabled: !find.canReplace.value,
            replaceAllDisabled: !find.canReplaceAll.value,
            "onUpdate:query": find.setQuery,
            "onUpdate:replacement": find.setReplacement,
            onReplace: find.replaceActive,
          })
          : null,
      });
    },
  });
  const mounted = await mount(Harness);
  try {
    findByTestId(mounted.root, "editor-find-toggle").props.onClick();
    await vue.nextTick();
    findByTestId(mounted.root, "docx-editor-search-input").props.onInput(inputEvent("needle"));
    await waitFor(
      () => find.state.value.matches.length === 1 && find.state.value.activeIndex === 0,
      "the stub surface reports one active result",
    );
    findByTestId(mounted.root, "office-find-replace-input").props.onInput(inputEvent("pin"));
    await vue.nextTick();
    assert.equal(find.canReplace.value, true);

    findByTestId(mounted.root, "office-find-replace").props.onClick();
    await waitFor(
      () => paragraphText(controller.model.nodes[0]) === "alpha pin beta",
      "Replace rewrites the active result",
    );
    assert.equal(controller.canUndo, true);
    controller.undo();
    await vue.nextTick();
    assert.equal(paragraphText(controller.model.nodes[0]), "alpha needle beta");
  } finally {
    mounted.app.unmount();
  }
});
