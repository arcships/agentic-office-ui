// OpenDocument text in the Vue components — issue #144.
//
// Opening `.odt` must render the document read-only with the upstream
// view-only reason, disable download/export, and leave `.docx` behaviour
// unchanged. The fixture is parsed by the real reader; the component-level
// runtime is a small loader stub so these tests stay browser-free.

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  textContent,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

if (globalThis.Document === undefined) globalThis.Document = class Document {};
if (globalThis.ShadowRoot === undefined) globalThis.ShadowRoot = class ShadowRoot {};

const core = await importFromDemo("@arcships/docx-core");
const {
  DocxEditor,
  DocxViewer,
  OPEN_DOCUMENT_VIEW_ONLY_REASON,
  isOpenDocumentTextFileName,
  isSupportedDocumentFileName,
  isWordFileName,
  useDocxEditor,
} = await importFromDemo("@arcships/vue-docx");

const nativeSetTimeout = globalThis.setTimeout;
const FIXTURE = new URL("../fixtures/docx/opendocument/sample-html-features.odt", import.meta.url);
const ODT_BYTES = (() => {
  const bytes = readFileSync(FIXTURE);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
})();

let odtModelPromise;
function odtModel() {
  odtModelPromise ??= core.buildDocModelFromBytes(ODT_BYTES).then((result) => result.model);
  return odtModelPromise;
}

async function waitUntil(predicate, message, timeoutMs = 5_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    await vue.nextTick();
    if (predicate()) return;
    await new Promise((resolve) => nativeSetTimeout(resolve, 5));
  }
  throw new Error(`Timed out waiting for ${message}`);
}

function emptyPackage() {
  return { parts: new Map(), binaryAssets: new Map() };
}

function fakeRuntime(model) {
  return {
    limits: {},
    createLoader: () => ({
      async load() {
        return {
          package: emptyPackage(),
          model,
          source: "main-thread",
          timings: { totalMs: 0, parseMs: 0, buildModelMs: 0 },
        };
      },
      cancel() {},
      dispose() {},
    }),
  };
}

function fileInput(root) {
  return walk(root).find((node) => node.type === "input" && node.props?.type === "file");
}

async function uploadViewerFile(root, name) {
  await fileInput(root).props.onChange({
    target: {
      files: [{
        name,
        size: 64,
        arrayBuffer: async () => ODT_BYTES,
      }],
      value: name,
    },
  });
}

test("the public file policy names both Word and OpenDocument text families", () => {
  for (const name of ["a.docx", "a.docm", "a.dotx", "a.dotm", "A.FODT", "a.odt", "a.ott", "a.odm", "a.otm", "a.oth"]) {
    assert.equal(isSupportedDocumentFileName(name), true, name);
  }
  for (const name of ["a.doc", "a.ods", "a.odp", "a.pdf", "odt"]) {
    assert.equal(isSupportedDocumentFileName(name), false, name);
  }
  assert.equal(isWordFileName("a.docx"), true);
  assert.equal(isWordFileName("a.odt"), false);
  assert.equal(isOpenDocumentTextFileName("a.odt"), true);
  assert.equal(isOpenDocumentTextFileName("a.docx"), false);
  assert.match(OPEN_DOCUMENT_VIEW_ONLY_REASON, /view-only/);
});

test("DocxViewer shows the view-only reason and disables download for an OpenDocument model", async () => {
  const model = await odtModel();
  const mounted = await mount(DocxViewer, {
    model,
    fileName: "sample.odt",
  });

  try {
    await waitUntil(
      () => findByTestId(mounted.root, "docx-viewer")?.props["data-state"] === "ready",
      "OpenDocument viewer ready",
    );
    const root = findByTestId(mounted.root, "docx-viewer");
    assert.equal(root.props["data-view-only"], "true");
    const reason = findByTestId(mounted.root, "docx-view-only-reason");
    assert.ok(reason, "the reason must be visible, not only exposed as state");
    assert.equal(textContent(reason).trim(), OPEN_DOCUMENT_VIEW_ONLY_REASON);
    assert.equal(
      walk(mounted.root).some((node) => node.props?.contenteditable === "true"),
      false,
      "the viewer never becomes editable",
    );
    assert.equal(findByTestId(mounted.root, "docx-download").props.disabled, true);
    assert.ok(
      walk(mounted.root).some((node) => node.props?.["data-docx-page-index"] !== undefined),
      "the OpenDocument document must render at least one page",
    );
    assert.ok(
      walk(mounted.root).filter((node) => node.props?.["data-docx-paragraph-host"] === true).length >= 10,
      "the imported paragraphs must reach the page surface",
    );
  } finally {
    mounted.app.unmount();
  }
});

test("opening a real .odt renders read-only and leaves .docx download enabled", async () => {
  const odt = await odtModel();
  const openDocumentViewer = await mount(DocxViewer, { runtime: fakeRuntime(odt) });

  try {
    await uploadViewerFile(openDocumentViewer.root, "sample.odt");
    await waitUntil(
      () => findByTestId(openDocumentViewer.root, "docx-viewer")?.props["data-state"] === "ready",
      "OpenDocument upload ready",
    );
    assert.equal(findByTestId(openDocumentViewer.root, "docx-viewer").props["data-view-only"], "true");
    assert.ok(findByTestId(openDocumentViewer.root, "docx-view-only-reason"));
    assert.equal(findByTestId(openDocumentViewer.root, "docx-download").props.disabled, true);
  } finally {
    openDocumentViewer.app.unmount();
  }

  const word = await mount(DocxViewer, {
    runtime: fakeRuntime(core.createBlankDocumentModel()),
  });
  try {
    await uploadViewerFile(word.root, "sample.docx");
    await waitUntil(
      () => findByTestId(word.root, "docx-viewer")?.props["data-state"] === "ready",
      "DOCX upload ready",
    );
    assert.equal(findByTestId(word.root, "docx-viewer").props["data-view-only"], undefined);
    assert.equal(findByTestId(word.root, "docx-view-only-reason"), undefined);
    assert.equal(findByTestId(word.root, "docx-download").props.disabled, false);
  } finally {
    word.app.unmount();
  }
});

test("the editor opens OpenDocument text view-only and refuses edits and export", async () => {
  const model = await odtModel();
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model, initialFileName: "sample.odt" });
      return () => vue.h("div", { "data-testid": "odt-editor-harness" });
    },
  });
  const mounted = await mount(Harness);
  const editor = await mount(DocxEditor, {
    editor: controller,
    editable: true,
    showToolbar: true,
  });

  try {
    await vue.nextTick();
    assert.equal(controller.viewOnlyReason, OPEN_DOCUMENT_VIEW_ONLY_REASON);
    const reason = findByTestId(editor.root, "docx-editor-view-only-reason");
    assert.ok(reason);
    assert.equal(textContent(reason).trim(), OPEN_DOCUMENT_VIEW_ONLY_REASON);
    assert.equal(
      walk(editor.root).some((node) => node.props?.contenteditable === "true"),
      false,
      "the editing surface must not be contenteditable",
    );

    const before = controller.model.nodes.length;
    const first = controller.appendParagraph("one");
    const second = controller.appendParagraph("two");
    assert.equal(second, first, "a blocked append reports the same place twice");
    assert.equal(controller.model.nodes.length, before, "the view-only model never grows");
    assert.equal(controller.status, OPEN_DOCUMENT_VIEW_ONLY_REASON);

    await assert.rejects(
      controller.getDocxBytes(),
      (error) => /view-only/.test(error.message),
    );
    controller.exportDocx();
    await waitUntil(
      () => controller.status === OPEN_DOCUMENT_VIEW_ONLY_REASON,
      "view-only export status",
    );

    const wordController = useDocxEditor({ starterModel: core.createBlankDocumentModel() });
    assert.equal(wordController.viewOnlyReason, undefined);
    const wordBefore = wordController.model.nodes.length;
    wordController.appendParagraph("editable");
    assert.equal(wordController.model.nodes.length, wordBefore + 1);
  } finally {
    editor.app.unmount();
    mounted.app.unmount();
  }
});
