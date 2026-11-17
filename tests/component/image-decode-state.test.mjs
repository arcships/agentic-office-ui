import assert from "node:assert/strict";
import { test } from "node:test";

import {
  importFromDemo,
  mount,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

const { DocxImageLayer, DocxTableHost, renderParagraphRuns } =
  await importFromDemo("@arcships/vue-docx");
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

function floatingImageRun(src, alt, overrides = {}) {
  return {
    ...imageRun(src, alt),
    floating: {
      xPx: 8,
      yPx: 8,
      wrapType: "none",
      behindDocument: false,
      ...overrides,
    },
  };
}

function imageModel(image) {
  return {
    nodes: [{ type: "paragraph", children: [image] }],
    metadata: {
      sourceParts: 0,
      warnings: [],
      headerSections: [],
      footerSections: [],
      paragraphStyles: [],
    },
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

test("DOCX floating image failure resets when the document switches", async () => {
  const controller = vue.reactive({
    model: imageModel(floatingImageRun("data:image/png;base64,broken", "broken")),
    selection: { kind: "paragraph", nodeIndex: 0 },
    selectParagraph() {},
    moveFloatingImage() {},
    resizeImage() {},
    setImageWrapMode() {},
  });
  const mounted = await mount(DocxImageLayer, {
    pageIndex: 0,
    pageWidthPx: 816,
    pageHeightPx: 1056,
    controller,
    pageNodeSegments: [{ nodeIndex: 0 }],
  });
  try {
    const images = () => walk(mounted.root).filter((node) => node.type === "img");
    assert.equal(images().length, 1);

    images()[0].props.onError({ currentTarget: { dataset: {} } });
    await vue.nextTick();
    assert.equal(images().length, 0);
    assert.ok(
      walk(mounted.root).some(
        (node) => node.props?.["data-image-state"] === "error",
      ),
    );

    // Doc B reuses node 0 / child 0 with a valid image. Position-keyed error
    // state would keep it hidden; source-keyed state must render it.
    controller.model = imageModel(
      floatingImageRun("data:image/png;base64,dmFsaWQ=", "valid"),
    );
    await vue.nextTick();
    const nextImages = images();
    assert.equal(nextImages.length, 1);
    assert.equal(nextImages[0].props.src, "data:image/png;base64,dmFsaWQ=");
  } finally {
    mounted.app.unmount();
  }
});

test("DOCX floating image with layoutInCell stays inside its table cell", () => {
  const paragraph = {
    type: "paragraph",
    children: [
      floatingImageRun("data:image/png;base64,Y2VsbA==", "cell image", {
        xPx: 120,
        yPx: 10,
        layoutInCell: true,
      }),
    ],
  };
  paragraph.children[0].widthPx = 300;
  paragraph.children[0].heightPx = 120;
  const rendered = renderParagraphRuns(
    paragraph,
    "layout-in-cell",
    "light",
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    { tableCellLayout: { widthPx: 140 } },
  );
  const image = (Array.isArray(rendered) ? rendered : [rendered]).find(
    (node) => node.type === "img",
  );
  assert.ok(image);
  assert.equal(image.props["data-docx-layout-in-cell"], "true");
  assert.equal(image.props.style.maxWidth, "140px");
  assert.equal(image.props.style.left, 0);
  assert.equal(image.props.style.top, 10);

  const unconstrained = renderParagraphRuns(
    {
      ...paragraph,
      children: [
        {
          ...paragraph.children[0],
          floating: { ...paragraph.children[0].floating, layoutInCell: false },
        },
      ],
    },
    "layout-out-of-cell",
    "light",
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    { tableCellLayout: { widthPx: 140 } },
  );
  const outOfCell = (Array.isArray(unconstrained) ? unconstrained : [unconstrained]).find(
    (node) => node.type === "img",
  );
  assert.ok(outOfCell);
  assert.equal(outOfCell.props["data-docx-layout-in-cell"], undefined);
  assert.equal(outOfCell.props.style.maxWidth, undefined);
});

test("DOCX locked floating image anchors expose no resize handles", async () => {
  const controller = vue.reactive({
    model: imageModel(
      floatingImageRun("data:image/png;base64,bG9jaw==", "locked", {
        lockedAnchor: true,
      }),
    ),
    selection: { kind: "paragraph", nodeIndex: 0 },
    selectParagraph() {},
    moveFloatingImage() {},
    resizeImage() {},
    setImageWrapMode() {},
  });
  const mounted = await mount(DocxImageLayer, {
    pageIndex: 0,
    pageWidthPx: 816,
    pageHeightPx: 1056,
    controller,
    pageNodeSegments: [{ nodeIndex: 0 }],
    editable: true,
  });
  try {
    assert.equal(
      walk(mounted.root).filter(
        (node) => node.props?.["data-testid"] === "docx-image-resize-handle",
      ).length,
      0,
    );

    const container = walk(mounted.root).find(
      (node) => node.props?.["data-testid"] === "docx-floating-image",
    );
    assert.ok(container);
    container.props.onMousedown({
      preventDefault() {},
      clientX: 20,
      clientY: 20,
    });
    await vue.nextTick();
    assert.equal(
      walk(mounted.root).every(
        (node) => node.props?.["data-testid"] !== "docx-image-resize-handle",
      ),
      true,
    );
  } finally {
    mounted.app.unmount();
  }
});

test("DOCX table host clips layoutInCell floating images to the cell", async () => {
  const table = {
    type: "table",
    rows: [
      {
        type: "table-row",
        cells: [
          {
            type: "table-cell",
            nodes: [
              {
                type: "paragraph",
                children: [
                  {
                    ...floatingImageRun(
                      "data:image/png;base64,Y2VsbA==",
                      "cell image",
                      { layoutInCell: true },
                    ),
                    widthPx: 300,
                    heightPx: 120,
                  },
                ],
              },
            ],
          },
        ],
      },
    ],
  };
  const mounted = await mount(DocxTableHost, { table, tableIndex: 0 });
  try {
    const cell = walk(mounted.root).find(
      (node) => node.props?.["data-docx-table-cell"] === true,
    );
    assert.ok(cell);
    assert.match(cell.props.class, /docx-table-cell--contain-floating/);

    const paragraphHost = walk(mounted.root).find(
      (node) => node.props?.["data-docx-table-cell-paragraph-host"] === true,
    );
    assert.ok(paragraphHost);
    const maxWidthMatch = paragraphHost.props.innerHTML.match(
      /max-width:(\d+)px/,
    );
    assert.ok(maxWidthMatch, "layoutInCell image should receive a cell width cap");
    const maxWidthPx = Number(maxWidthMatch[1]);
    assert.ok(maxWidthPx > 0 && maxWidthPx < 300);
    assert.match(paragraphHost.props.innerHTML, /left:0/);
    assert.match(paragraphHost.props.innerHTML, /data-docx-layout-in-cell="true"/);
  } finally {
    mounted.app.unmount();
  }
});
