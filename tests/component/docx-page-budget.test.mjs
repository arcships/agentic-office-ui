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

const { DocxEditor, DocxViewer } = await importFromDemo("@arcships/vue-docx");
const {
  createBlankDocumentModel,
  createDocxRuntime,
} = await importFromDemo("@arcships/docx-core");

function twoPageModel() {
  const model = createBlankDocumentModel();
  return {
    ...model,
    metadata: { ...model.metadata, documentPageCount: 2 },
    nodes: Array.from({ length: 60 }, (_, index) => ({
      type: "paragraph",
      children: [{ type: "text", text: `段落 ${index + 1} ${"内容".repeat(100)}` }],
    })),
  };
}

function assertPageLimitError(node, emittedError) {
  assert.ok(node, "the public component must expose a stable page-limit error state");
  assert.equal(node.props["data-error-code"], "LIMIT_EXCEEDED");
  assert.equal(node.props["data-error-phase"], "layout");
  assert.equal(node.props["data-error-limit"], "maxDocxPages");
  assert.equal(node.props["data-error-actual"], 2);
  assert.equal(node.props["data-error-allowed"], 1);
  assert.equal(emittedError?.code, "LIMIT_EXCEEDED");
  assert.equal(emittedError?.phase, "layout");
  assert.equal(emittedError?.limit, "maxDocxPages");
  assert.equal(emittedError?.actual, 2);
  assert.equal(emittedError?.allowed, 1);
}

test("DocxViewer reports ready only after a normal document page count is accepted", async () => {
  const successes = [];
  const runtime = {
    id: "docx-page-budget-normal",
    limits: { maxDocxPages: 2 },
    createLoader() {
      return {
        async load() {
          return { model: twoPageModel(), package: {}, source: "worker" };
        },
        cancel() {},
        dispose() {},
      };
    },
    dispose() {},
  };
  const mounted = await mount(DocxViewer, {
    file: new ArrayBuffer(16),
    runtime,
    onLoadSuccess: (result) => successes.push(result),
  });

  try {
    await waitFor(
      () => findByTestId(mounted.root, "docx-viewer")?.props["data-state"] === "ready",
      "normal DOCX Viewer ready state after page validation",
    );
    assert.equal(successes.length, 1);
    assert.equal(successes[0].source, "worker");
    assert.equal(findByTestId(mounted.root, "load-error"), undefined);
  } finally {
    mounted.app.unmount();
  }
});

test("DocxViewer hides an over-limit document and recovers with a new runtime limit", async () => {
  const lowRuntime = createDocxRuntime({ limits: { maxDocxPages: 1 } });
  const normalRuntime = createDocxRuntime({ limits: { maxDocxPages: 2 } });
  const activeRuntime = vue.shallowRef(lowRuntime);
  const errors = [];
  const Harness = vue.defineComponent({
    setup() {
      return () => vue.h(DocxViewer, {
        model: twoPageModel(),
        runtime: activeRuntime.value,
        onLoadError: (error) => errors.push(error),
      });
    },
  });
  const mounted = await mount(Harness);

  try {
    await waitFor(
      () => findByTestId(mounted.root, "docx-viewer")?.props["data-state"] === "error",
      "DOCX Viewer page-limit error",
    );
    assertPageLimitError(findByTestId(mounted.root, "load-error"), errors[0]);
    assert.equal(
      walk(mounted.root).some((node) => node.props?.["data-testid"] === "docx-page"),
      false,
      "an over-limit document must not remain visible",
    );

    activeRuntime.value = normalRuntime;
    await vue.nextTick();
    await waitFor(
      () => findByTestId(mounted.root, "docx-viewer")?.props["data-state"] === "ready",
      "DOCX Viewer recovery with a normal page limit",
    );
    assert.equal(findByTestId(mounted.root, "load-error"), undefined);
    assert.equal(
      walk(mounted.root).some((node) => node.props?.["data-testid"] === "docx-page"),
      true,
      "the valid document should render again after applying a normal limit",
    );
  } finally {
    mounted.app.unmount();
    lowRuntime.dispose();
    normalRuntime.dispose();
  }
});

test("DocxEditor shared surface hides a document that exceeds runtime maxDocxPages", async () => {
  const runtime = createDocxRuntime({ limits: { maxDocxPages: 1 } });
  const errors = [];
  const mounted = await mount(DocxEditor, {
    editable: false,
    model: twoPageModel(),
    runtime,
    showThumbnails: true,
    showToolbar: false,
    onLoadError: (error) => errors.push(error),
  });

  try {
    await waitFor(
      () => Boolean(findByTestId(mounted.root, "editor-load-error")),
      "DOCX Editor page-limit error",
    );
    assertPageLimitError(findByTestId(mounted.root, "editor-load-error"), errors[0]);
    assert.equal(
      walk(mounted.root).some((node) => node.props?.["data-testid"] === "docx-page"),
      false,
      "the shared Editor surface must not display over-limit pages",
    );
  } finally {
    mounted.app.unmount();
    runtime.dispose();
  }
});
