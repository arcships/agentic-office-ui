import assert from "node:assert/strict";
import { test } from "node:test";

import {
  importFromDemo,
  mount,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

const { renderParagraphRuns } = await importFromDemo("@arcships/vue-docx");
const { XlsxImageLayer } = await importFromDemo("@arcships/vue-xlsx");

function imageRun(src, alt) {
  return {
    type: "image",
    src,
    alt,
    widthPx: 120,
    heightPx: 80,
  };
}

function xlsxImage(id) {
  return {
    id,
    name: id,
    mimeType: "image/png",
    sheetIndex: 0,
    workbookSheetIndex: 0,
    src: `data:image/png;base64,${id}`,
    zIndex: 0,
    anchor: {
      kind: "one-cell",
      from: { col: 0, colOffsetEmu: 0, row: 0, rowOffsetEmu: 0 },
      sizeEmu: { cx: 1_143_000, cy: 762_000 },
    },
  };
}

test("DOCX image decode errors expose a stable state and public callback", () => {
  const diagnostics = [];
  const rendered = renderParagraphRuns(
    { type: "paragraph", children: [imageRun("data:image/png;base64,broken", "broken")] },
    "decode-state",
    "light",
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    { onImageDecodeError: (error) => diagnostics.push(error) },
  );
  const image = (Array.isArray(rendered) ? rendered : [rendered]).find((node) => node.type === "img");
  assert.ok(image);
  const attributes = new Map();
  const target = {
    dataset: { imageState: "loading" },
    setAttribute(name, value) { attributes.set(name, value); },
  };
  image.props.onError({ currentTarget: target });
  assert.equal(target.dataset.imageState, "error");
  assert.equal(attributes.get("aria-invalid"), "true");
  assert.deepEqual(diagnostics, [{
    code: "IMAGE_DECODE_FAILED",
    phase: "image-decode",
    imageKey: "decode-state-run-0",
    message: "DOCX 图片无法解码。",
  }]);
});

test("XLSX decodes only the configured number of images and releases slots on load or error", async () => {
  const diagnostics = [];
  const controller = {
    selectedImageId: null,
    images: [xlsxImage("one"), xlsxImage("two"), xlsxImage("three")],
    maxConcurrentImageDecodes: 1,
    selectImage() {},
    reportImageDecodeError(id, message) { diagnostics.push({ id, message }); },
  };
  const mounted = await mount(XlsxImageLayer, { controller, showImages: true });
  try {
    const images = () => walk(mounted.root).filter((node) => node.type === "img");
    assert.deepEqual(images().map((node) => node.props.alt), ["one"]);

    images()[0].props.onLoad({ currentTarget: { dataset: {} } });
    await vue.nextTick();
    assert.deepEqual(images().map((node) => node.props.alt), ["one", "two"]);

    const second = images().find((node) => node.props.alt === "two");
    second.props.onError({ currentTarget: { dataset: {} } });
    await vue.nextTick();
    assert.deepEqual(images().map((node) => node.props.alt), ["one", "three"]);
    assert.deepEqual(diagnostics, [{ id: "two", message: undefined }]);
    assert.ok(
      walk(mounted.root).some(
        (node) => node.props?.["data-image-state"] === "error",
      ),
    );
  } finally {
    mounted.app.unmount();
  }
});
