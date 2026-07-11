import assert from "node:assert/strict";
import test from "node:test";

import {
  importFromDemo,
  mount,
  walk,
} from "../component/vue-test-renderer.mjs";

const { renderParagraphRuns } = await importFromDemo("@arcships/vue-docx");
const { XlsxImageLayer } = await importFromDemo("@arcships/vue-xlsx");

function declaredOversizedPngDataUri(width, height) {
  // Deliberately tiny and incomplete. It declares a huge IHDR size without
  // allocating the corresponding pixel buffer, as required by STRESS-004.
  const bytes = new Uint8Array(24);
  bytes.set([137, 80, 78, 71, 13, 10, 26, 10]);
  bytes.set([0, 0, 0, 13, 73, 72, 68, 82], 8);
  const view = new DataView(bytes.buffer);
  view.setUint32(16, width);
  view.setUint32(20, height);
  return `data:image/png;base64,${Buffer.from(bytes).toString("base64")}`;
}

function findRenderedImage(root) {
  return walk(root).find((node) => node.type === "img");
}

test("P3-IMAGE-01 regression: DOCX reports browser image decode failure", () => {
  const rendered = renderParagraphRuns(
    {
      type: "paragraph",
      children: [
        {
          type: "image",
          src: declaredOversizedPngDataUri(100_000, 100_000),
          alt: "oversized image metadata",
          widthPx: 120,
          heightPx: 80,
        },
      ],
    },
    "p3-image-baseline",
  );
  const image = (Array.isArray(rendered) ? rendered : [rendered]).find(
    (node) => node.type === "img",
  );

  assert.ok(image, "DOCX should render the image node before browser decoding");
  assert.equal(
    typeof image.props?.onError,
    "function",
    "DOCX currently has no decode-error handler, stable error state, or diagnostic event",
  );
});

test("P3-IMAGE-01 regression: XLSX reports browser image decode failure", async () => {
  const controller = {
    selectedImageId: null,
    images: [
      {
        id: "oversized-image",
        name: "oversized image metadata",
        mimeType: "image/png",
        sheetIndex: 0,
        workbookSheetIndex: 0,
        src: declaredOversizedPngDataUri(100_000, 100_000),
        zIndex: 0,
        anchor: {
          kind: "one-cell",
          from: { col: 0, colOffsetEmu: 0, row: 0, rowOffsetEmu: 0 },
          sizeEmu: { cx: 1_143_000, cy: 762_000 },
        },
      },
    ],
    selectImage() {},
  };
  const mounted = await mount(XlsxImageLayer, {
    controller,
    showImages: true,
  });

  try {
    const image = findRenderedImage(mounted.root);
    assert.ok(image, "XLSX should render the image node before browser decoding");
    assert.equal(
      typeof image.props.onError,
      "function",
      "XLSX currently has no decode-error handler, stable error state, or diagnostic event",
    );
  } finally {
    mounted.app.unmount();
  }
});
