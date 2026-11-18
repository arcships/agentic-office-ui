import assert from "node:assert/strict";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

const {
  DocxViewer,
  DocxViewerRoot,
  useDocxEditor,
} = await importFromDemo("@arcships/vue-docx");
const { createBlankDocumentModel } = await importFromDemo("@arcships/docx-core");

const UNSUPPORTED_MESSAGE =
  "Only Word (.docx, .docm, .dotx, .dotm) and OpenDocument text (.odt, .ott, .odm, .otm, .oth, .fodt) files are supported";
const SUPPORTED_EXTENSIONS = [
  ".docx",
  ".docm",
  ".dotx",
  ".dotm",
  ".odt",
  ".ott",
  ".odm",
  ".otm",
  ".oth",
  ".fodt",
];
const ACCEPT_ATTRIBUTE = [
  ...SUPPORTED_EXTENSIONS,
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "application/vnd.ms-word.document.macroEnabled.12",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
  "application/vnd.ms-word.template.macroEnabled.12",
  "application/vnd.oasis.opendocument.text",
  "application/vnd.oasis.opendocument.text-template",
  "application/vnd.oasis.opendocument.text-master",
  "application/vnd.oasis.opendocument.text-master-template",
  "application/vnd.oasis.opendocument.text-web",
  "application/vnd.oasis.opendocument.text-flat-xml",
].join(",");

const nativeSetTimeout = globalThis.setTimeout;

function sampleFile(name) {
  return new File([new Uint8Array([1, 2, 3])], name);
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

async function mountController(options = {}) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({
        starterModel: createBlankDocumentModel(),
        ...options,
      });
      return () => vue.h("div", { "data-testid": "controller-harness" }, controller.status);
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, get controller() { return controller; } };
}

// A successful Worker response is enough to exercise the import lifecycle;
// the point of this test is file-type handling and names, not DOCX parsing.
function installSuccessfulWorker(t) {
  const originalWorker = globalThis.Worker;
  class SuccessfulWorker {
    constructor() {
      this.listeners = new Map();
    }
    addEventListener(type, listener) {
      const current = this.listeners.get(type) ?? [];
      current.push(listener);
      this.listeners.set(type, current);
    }
    removeEventListener(type, listener) {
      this.listeners.set(type, (this.listeners.get(type) ?? []).filter((entry) => entry !== listener));
    }
    postMessage(request) {
      queueMicrotask(() => {
        for (const listener of this.listeners.get("message") ?? []) {
          listener({
            data: {
              id: request.id,
              type: "success",
              package: {},
              model: createBlankDocumentModel(),
              timings: { totalMs: 1, parseMs: 1, buildModelMs: 0 },
            },
          });
        }
      });
    }
    terminate() {}
  }
  globalThis.Worker = SuccessfulWorker;
  t.after(() => {
    if (originalWorker === undefined) delete globalThis.Worker;
    else globalThis.Worker = originalWorker;
  });
}

function installDownloadEnvironment(t) {
  const originalDocument = globalThis.document;
  const originalCreateObjectURL = URL.createObjectURL;
  const originalRevokeObjectURL = URL.revokeObjectURL;
  const anchors = [];

  globalThis.document = {
    body: { append() {} },
    createElement(tagName) {
      assert.equal(tagName, "a");
      const anchor = {
        style: {},
        download: "",
        href: "",
        click() {},
        remove() {},
      };
      anchors.push(anchor);
      return anchor;
    },
  };
  URL.createObjectURL = () => "blob:docx-file-handling-test";
  URL.revokeObjectURL = () => {};

  t.after(() => {
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
    URL.createObjectURL = originalCreateObjectURL;
    URL.revokeObjectURL = originalRevokeObjectURL;
  });

  return { anchors };
}

function findDropTarget(root) {
  return walk(root).find((node) =>
    String(node.props?.class || "").split(/\s+/).includes("docx-viewer-root")
  );
}

test("New restores the initial file name and the next export uses it", async (t) => {
  installSuccessfulWorker(t);
  const mounted = await mountController({ initialFileName: "starter.docx" });
  const { anchors } = installDownloadEnvironment(t);

  try {
    await mounted.controller.importDocxFile(sampleFile("imported.docx"));
    assert.equal(mounted.controller.fileName, "imported.docx");

    mounted.controller.newDocument();
    assert.equal(mounted.controller.fileName, "starter.docx");

    mounted.controller.exportDocx();
    await waitUntil(
      () => mounted.controller.status === "Exported DOCX",
      "DOCX export after New",
    );
    assert.equal(anchors.at(-1)?.download, "starter-edited.docx");
  } finally {
    mounted.app.unmount();
  }
});

test("New falls back to the starter display name without an initialFileName", async () => {
  const mounted = await mountController();
  try {
    mounted.controller.newDocument();
    assert.equal(mounted.controller.fileName, "(new document)");
  } finally {
    mounted.app.unmount();
  }
});

test("the importer accepts every Word and OpenDocument text extension and rejects legacy .doc", async (t) => {
  installSuccessfulWorker(t);
  const mounted = await mountController({ initialFileName: "starter.docx" });

  try {
    for (const extension of SUPPORTED_EXTENSIONS) {
      await mounted.controller.importDocxFile(sampleFile(`sample${extension}`));
      assert.equal(mounted.controller.fileName, `sample${extension}`);
      assert.equal(mounted.controller.status, "Ready");
    }

    await mounted.controller.importDocxFile(sampleFile("legacy.doc"));
    assert.equal(mounted.controller.status, UNSUPPORTED_MESSAGE);
    assert.equal(mounted.controller.isImporting, false);
    assert.equal(mounted.controller.fileName, `sample${SUPPORTED_EXTENSIONS.at(-1)}`);
  } finally {
    mounted.app.unmount();
  }
});

test("drag-and-drop accepts every picker extension and reports rejected drops", async () => {
  let controller;
  const imported = [];
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: createBlankDocumentModel() });
      controller.importDocxFile = async (file) => {
        imported.push(file.name);
      };
      return () => vue.h(DocxViewerRoot, {
        model: controller.model,
        controller,
        editable: true,
      });
    },
  });
  const mounted = await mount(Harness);

  try {
    const dropTarget = findDropTarget(mounted.root);
    assert.ok(dropTarget, "viewer root must own the drop target");

    for (const extension of SUPPORTED_EXTENSIONS) {
      const fileName = `sample${extension}`;
      const dataTransfer = {
        types: ["Files"],
        files: [sampleFile(fileName)],
        dropEffect: "none",
      };
      dropTarget.props.onDragover({ dataTransfer, preventDefault() {} });
      assert.equal(dataTransfer.dropEffect, "copy", `${extension} must be accepted on dragover`);
      await dropTarget.props.onDrop({ dataTransfer, preventDefault() {} });
    }

    assert.deepEqual(
      imported,
      SUPPORTED_EXTENSIONS.map((extension) => `sample${extension}`),
    );

    await dropTarget.props.onDrop({
      dataTransfer: { files: [sampleFile("legacy.doc")] },
      preventDefault() {},
    });
    assert.equal(controller.status, UNSUPPORTED_MESSAGE);
    assert.equal(imported.length, SUPPORTED_EXTENSIONS.length);
  } finally {
    mounted.app.unmount();
  }
});

test("DocxViewer picker and upload path share the accepted extension policy", async () => {
  const errors = [];
  const mounted = await mount(DocxViewer, {
    emptyState: "DOCX empty",
    onLoadError: (error) => errors.push(error),
  });

  try {
    const fileInput = walk(mounted.root).find(
      (node) => node.type === "input" && node.props?.type === "file",
    );
    assert.ok(fileInput, "viewer must render a file input");
    assert.equal(fileInput.props.accept, ACCEPT_ATTRIBUTE);
    assert.doesNotMatch(fileInput.props.accept, /\.doc(?:,|$)/);

    await fileInput.props.onChange({
      target: { files: [sampleFile("legacy.doc")], value: "legacy.doc" },
    });
    assert.equal(errors.length, 1);
    assert.equal(errors[0].message, UNSUPPORTED_MESSAGE);
    assert.equal(findByTestId(mounted.root, "docx-viewer").props["data-state"], "error");
  } finally {
    mounted.app.unmount();
  }
});
