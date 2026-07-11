import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import {
  importFromDemo,
  mount,
  vue,
  waitFor,
} from "../component/vue-test-renderer.mjs";

const {
  DocxThumbnailSurfaceCache,
  applyEditorTransactionV2,
  createBlankDocumentModel,
  createEditorStateV2,
  getDownscaledThumbnailImageDataUri,
} = await importFromDemo("@arcships/docx-core");
const {
  revokeWorkbookImageAssets,
  setWasmSource,
} = await importFromDemo("@arcships/xlsx-core");
const { useXlsxViewerController } = await importFromDemo("@arcships/vue-xlsx");

const TEST_HISTORY_BUDGET_BYTES = 64 * 1024;
const LARGE_HISTORY_TEXT_BYTES = 16 * 1024;
const TEST_THUMBNAIL_BUDGET_BYTES = 4 * 1024 * 1024;

function docxModelWithText(text) {
  return {
    ...createBlankDocumentModel(),
    nodes: [{ type: "paragraph", children: [{ type: "text", text }] }],
  };
}

function utf8JsonBytes(value) {
  return new TextEncoder().encode(JSON.stringify(value)).byteLength;
}

function workbookBuffer() {
  const bytes = readFileSync(
    new URL("../../apps/demo/public/samples/sales-table.xlsx", import.meta.url),
  );
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function configureXlsxWasm() {
  const wasmBytes = readFileSync(
    new URL(
      "../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm",
      import.meta.url,
    ),
  );
  setWasmSource(
    wasmBytes.buffer.slice(
      wasmBytes.byteOffset,
      wasmBytes.byteOffset + wasmBytes.byteLength,
    ),
  );
}

async function mountEditableXlsxController() {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: workbookBuffer(),
        fileName: "p3-cache-baseline.xlsx",
        historyMaxBytes: TEST_HISTORY_BUDGET_BYTES,
        skipXmlParsing: true,
        useWorker: false,
      });
      return () => vue.h("div", { "data-state": controller.sourceState });
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    "editable XLSX terminal state",
  );
  assert.ifError(controller.error);
  return { ...mounted, controller };
}

test("P3-CACHE-01 regression: DOCX history is limited by retained bytes", () => {
  const selection = { kind: "paragraph", nodeIndex: 0 };
  let state = createEditorStateV2({
    historyMaxBytes: TEST_HISTORY_BUDGET_BYTES,
    historyMaxEntries: 100,
    model: docxModelWithText("initial"),
    selection,
  });

  for (let index = 0; index < 12; index += 1) {
    state = applyEditorTransactionV2(state, {
      model: docxModelWithText(`${index}:${"D".repeat(LARGE_HISTORY_TEXT_BYTES)}`),
      selection,
    });
  }

  const retainedBytes = state.history.past.reduce(
    (total, entry) => total + utf8JsonBytes(entry),
    0,
  );
  assert.ok(
    retainedBytes <= TEST_HISTORY_BUDGET_BYTES,
    `DOCX history retained ${retainedBytes} bytes; it is capped only by entry count`,
  );
});

test("P3-CACHE-01 regression: XLSX public undo history is limited by retained bytes", async () => {
  configureXlsxWasm();
  const mounted = await mountEditableXlsxController();
  try {
    for (let index = 0; index < 12; index += 1) {
      mounted.controller.setCellValue(
        { col: 0, row: 0 },
        `${index}:${"X".repeat(LARGE_HISTORY_TEXT_BYTES)}`,
      );
    }
    await vue.nextTick();

    let undoCount = 0;
    while (mounted.controller.canUndo && undoCount < 20) {
      mounted.controller.undo();
      undoCount += 1;
      await vue.nextTick();
    }

    const maximumEntriesUnderTestBudget = Math.floor(
      TEST_HISTORY_BUDGET_BYTES / (LARGE_HISTORY_TEXT_BYTES * 2),
    );
    assert.ok(
      undoCount <= maximumEntriesUnderTestBudget,
      `XLSX retained ${undoCount} large undo entries; it is capped only by entry count`,
    );
  } finally {
    mounted.app.unmount();
  }
});

test("P3-CACHE-01 regression: DOCX thumbnail cache evicts by pixel bytes", () => {
  const cache = new DocxThumbnailSurfaceCache({
    estimateBytes: (surface) => surface.width * surface.height * 4,
    maxBytes: TEST_THUMBNAIL_BUDGET_BYTES,
    maxEntries: 8,
  });
  const surfaces = Array.from({ length: 3 }, (_, index) => ({
    id: `surface-${index}`,
    height: 1024,
    width: 1024,
  }));
  surfaces.forEach((surface) => cache.set(surface.id, surface));

  const retained = surfaces
    .map((surface) => cache.get(surface.id))
    .filter(Boolean);
  const retainedBytes = retained.reduce(
    (total, surface) => total + surface.width * surface.height * 4,
    0,
  );
  assert.ok(
    retainedBytes <= TEST_THUMBNAIL_BUDGET_BYTES,
    `thumbnail cache retained ${retainedBytes} pixel bytes under an entry-only limit`,
  );
});

test("P3-CACHE-01 regression: DOCX thumbnail consumers do not share module cache", async (t) => {
  const originalDocument = globalThis.document;
  const originalImage = globalThis.Image;
  let renderVersion = "first-consumer";

  class FakeImage {
    naturalHeight = 1024;
    naturalWidth = 1024;
    height = 1024;
    width = 1024;
    async decode() {}
  }

  globalThis.Image = FakeImage;
  globalThis.document = {
    createElement(tagName) {
      assert.equal(tagName, "canvas");
      return {
        height: 0,
        width: 0,
        getContext() {
          return {
            drawImage() {},
            imageSmoothingEnabled: false,
            imageSmoothingQuality: "low",
          };
        },
        toDataURL() {
          return `data:image/png;base64,${renderVersion}`;
        },
      };
    },
  };
  t.after(() => {
    if (originalImage === undefined) delete globalThis.Image;
    else globalThis.Image = originalImage;
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
  });

  const source = `data:image/png;base64,${"A".repeat(40_000)}`;
  const firstConsumerResult = await getDownscaledThumbnailImageDataUri(source);
  renderVersion = "second-consumer";
  const secondConsumerResult = await getDownscaledThumbnailImageDataUri(source);

  assert.match(firstConsumerResult, /first-consumer$/);
  assert.match(
    secondConsumerResult,
    /second-consumer$/,
    "the second consumer received a value retained by a module-level cache",
  );
});

test("P3-CACHE-01 regression: XLSX object URL cleanup is idempotent", (t) => {
  const originalRevokeObjectURL = URL.revokeObjectURL;
  const revoked = [];
  URL.revokeObjectURL = (url) => revoked.push(url);
  t.after(() => {
    URL.revokeObjectURL = originalRevokeObjectURL;
  });

  const assets = {
    objectUrls: ["blob:p3-cache-first", "blob:p3-cache-second"],
  };
  revokeWorkbookImageAssets(assets);
  revokeWorkbookImageAssets(assets);

  assert.deepEqual(
    revoked,
    ["blob:p3-cache-first", "blob:p3-cache-second"],
    "repeated unload cleanup revoked the same object URLs more than once",
  );
});
