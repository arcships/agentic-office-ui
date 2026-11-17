import assert from "node:assert/strict";
import test from "node:test";

import {
  importFromDemo,
  mount,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

const { DocxPageSurface, DocxViewerRoot, useDocxEditor } = await importFromDemo("@arcships/vue-docx");
const {
  createBlankDocumentModel,
} = await importFromDemo("@arcships/docx-core");

const PAGE_LAYOUT = {
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginsPx: { top: 72, right: 72, bottom: 72, left: 72 },
  headerDistancePx: 36,
  footerDistancePx: 36,
};

const PAGE_BORDER_XML = [
  "<w:sectPr>",
  '<w:pgBorders w:offsetFrom="page">',
  '<w:top w:val="double" w:sz="12" w:space="24" w:color="FF0000" w:shadow="1"/>',
  '<w:right w:val="single" w:sz="6" w:space="18" w:color="00FF00"/>',
  '<w:bottom w:val="dashed" w:sz="6" w:space="24" w:color="0000FF"/>',
  '<w:left w:val="single" w:sz="6" w:space="20" w:color="112233"/>',
  "</w:pgBorders>",
  "</w:sectPr>",
].join("");

function paragraph(children) {
  return { type: "paragraph", children };
}

function imageRun(overrides) {
  return {
    type: "image",
    src: "data:image/png;base64,iVBORw0KGgo=",
    widthPx: 800,
    heightPx: 1000,
    ...overrides,
  };
}

function mountSurface(model, props = {}) {
  return mount(DocxPageSurface, {
    pageIndex: 0,
    pageLayout: PAGE_LAYOUT,
    pageNodeSegments: model.nodes.map((_node, nodeIndex) => ({ nodeIndex })),
    model,
    editable: false,
    pageWidthPx: PAGE_LAYOUT.pageWidthPx,
    pageContentWidthPx: PAGE_LAYOUT.pageWidthPx - 144,
    pageNumber: 1,
    totalPages: 1,
    theme: "light",
    ...props,
  });
}

function isTrue(value) {
  return value === true || value === "true";
}

function surfaceNode(root) {
  return walk(root).find((node) => isTrue(node.props?.["data-docx-page-surface"]));
}

function borderEdges(root) {
  const overlay = walk(root).find(
    (node) => isTrue(node.props?.["data-docx-page-border-overlay"])
  );
  assert.ok(overlay, "page border overlay must render");
  return walk(overlay).filter((node) =>
    typeof node.props?.["data-docx-page-border-edge"] === "string"
  );
}

test("DOCX page surface paints document background, pgBorders and a full-page cover behind the body", async () => {
  const model = createBlankDocumentModel();
  model.metadata.documentBackgroundColor = "FFF2CC";
  model.metadata.sectionPropertiesXml = PAGE_BORDER_XML;
  model.nodes = [
    paragraph([
      imageRun({
        floating: {
          behindDocument: true,
          wrapType: "none",
          verticalRelativeTo: "page",
          xPx: 0,
          yPx: 0,
        },
      }),
    ]),
    paragraph([{ type: "text", text: "Body text stays above the cover" }]),
  ];

  const mounted = await mountSurface(model);

  const surface = surfaceNode(mounted.root);
  assert.equal(surface.props.style.backgroundColor, "#FFF2CC");
  assert.equal(surface.props.style.isolation, "isolate");

  const coverLayer = walk(mounted.root).find(
    (node) => isTrue(node.props?.["data-docx-page-cover-layer"])
  );
  assert.ok(coverLayer, "full-page behind-text image must render in the cover layer");
  const coverImage = walk(coverLayer).find((node) => node.type === "img");
  assert.ok(coverImage);
  assert.equal(coverImage.props.src, "data:image/png;base64,iVBORw0KGgo=");
  assert.equal(coverImage.props.style.left, 0);
  assert.equal(coverImage.props.style.top, 0);
  assert.equal(coverImage.props.style.width, 816);
  assert.equal(coverImage.props.style.height, 1056);

  const paragraphHosts = walk(mounted.root).filter(
    (node) => node.props?.["data-docx-paragraph-host"] === true
  );
  assert.equal(
    paragraphHosts[0].props.innerHTML.includes("<img"),
    false,
    "the cover image must not be rendered a second time in the body"
  );
  assert.match(paragraphHosts[1].props.innerHTML, /Body text stays above the cover/);

  const edges = borderEdges(mounted.root);
  assert.deepEqual(
    edges.map((edge) => edge.props["data-docx-page-border-edge"]),
    ["top", "right", "bottom", "left"]
  );
  const top = edges.find((edge) => edge.props["data-docx-page-border-edge"] === "top");
  assert.equal(top.props.style.top, "32px");
  assert.equal(top.props.style.borderTop, "2px double #FF0000");
  assert.equal(top.props.style.color, "#FF0000");
  assert.equal(top.props.style.filter, "drop-shadow(1px 1px 0 currentColor)");

  const bottom = edges.find((edge) => edge.props["data-docx-page-border-edge"] === "bottom");
  assert.equal(bottom.props.style.bottom, "32px");
  assert.equal(bottom.props.style.borderBottom, "1px dashed #0000FF");

  const overlay = walk(mounted.root).find(
    (node) => isTrue(node.props?.["data-docx-page-border-overlay"])
  );
  assert.equal(overlay.props.style.zIndex, "5");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DOCX viewer root passes document background and section borders into every page surface", async () => {
  const model = createBlankDocumentModel();
  model.metadata.documentBackgroundColor = "FFF2CC";
  model.metadata.sectionPropertiesXml = PAGE_BORDER_XML;
  model.nodes = [paragraph([{ type: "text", text: "Bordered page" }])];

  const mounted = await mount(DocxViewerRoot, { model });
  await vue.nextTick();

  const surface = surfaceNode(mounted.root);
  assert.ok(surface);
  assert.equal(surface.props.style.backgroundColor, "#FFF2CC");

  const edges = borderEdges(mounted.root);
  assert.equal(edges.length, 4);
  const top = edges.find((edge) => edge.props["data-docx-page-border-edge"] === "top");
  assert.equal(top.props.style.top, "32px");
  assert.equal(top.props.style.borderTop, "2px double #FF0000");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DOCX viewer root exposes a full-page behind-text cover from the paginated body", async () => {
  const model = createBlankDocumentModel();
  model.nodes = [
    paragraph([
      imageRun({
        floating: {
          behindDocument: true,
          wrapType: "none",
          verticalRelativeTo: "page",
          xPx: 0,
          yPx: 0,
        },
      }),
    ]),
    paragraph([{ type: "text", text: "Covered body" }]),
  ];

  const mounted = await mount(DocxViewerRoot, { model });
  await vue.nextTick();

  const coverLayer = walk(mounted.root).find(
    (node) => isTrue(node.props?.["data-docx-page-cover-layer"])
  );
  assert.ok(coverLayer, "the viewer must detect a paginated full-page cover image");
  const coverImage = walk(coverLayer).find((node) => node.type === "img");
  assert.ok(coverImage);
  assert.equal(coverImage.props.src, "data:image/png;base64,iVBORw0KGgo=");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DOCX behind-text watermark renders below body text, not in the front image layer", async () => {
  const model = createBlankDocumentModel();
  model.nodes = [
    paragraph([
      imageRun({
        widthPx: 300,
        heightPx: 200,
        floating: {
          behindDocument: true,
          wrapType: "none",
          verticalRelativeTo: "page",
          xPx: 100,
          yPx: 120,
        },
      }),
    ]),
    paragraph([{ type: "text", text: "Watermark must stay behind this text" }]),
  ];

  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model });
      return () =>
        vue.h(DocxPageSurface, {
          pageIndex: 0,
          pageLayout: PAGE_LAYOUT,
          pageNodeSegments: model.nodes.map((_node, nodeIndex) => ({ nodeIndex })),
          model: controller.model,
          controller,
          editable: true,
          pageWidthPx: PAGE_LAYOUT.pageWidthPx,
          pageContentWidthPx: PAGE_LAYOUT.pageWidthPx - 144,
          pageNumber: 1,
          totalPages: 1,
          theme: "light",
        });
    },
  });

  const mounted = await mount(Harness);
  await vue.nextTick();

  const behindLayer = walk(mounted.root).find(
    (node) => node.props?.["data-docx-image-layer"] === "behind"
  );
  assert.ok(behindLayer, "behind-text image must use the behind image layer");
  assert.equal(behindLayer.props.style.zIndex, "-1");

  const behindImage = walk(behindLayer).find((node) =>
    node.props?.["data-image-behind"] === "true"
  );
  assert.ok(behindImage);
  assert.equal(behindImage.props.style.left, "100px");
  assert.equal(behindImage.props.style.top, "120px");
  assert.equal(behindImage.props.style.pointerEvents, "none");

  const frontLayer = walk(mounted.root).find(
    (node) => node.props?.["data-docx-image-layer"] === "front"
  );
  assert.equal(frontLayer, undefined, "a behind-text image must not render in the front layer");

  const paragraphHosts = walk(mounted.root).filter(
    (node) => node.props?.["data-docx-paragraph-host"] === true
  );
  assert.equal(
    paragraphHosts[0].props.innerHTML.includes("<img"),
    false,
    "editable body content leaves floating images to the image layer"
  );
  assert.match(paragraphHosts[1].props.innerHTML, /Watermark must stay behind this text/);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});
