import assert from "node:assert/strict";
import { test } from "node:test";

import {
  importFromDemo,
  mount,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

const { DocxImageLayer, DocxPageSurface, renderParagraphRuns } = await importFromDemo("@arcships/vue-docx");

const PAGE_ORIGIN = {
  left: 0,
  top: 0,
  marginLeft: 96,
  marginTop: 72,
  columnLeft: 96,
  columnTop: 72,
  pageWidth: 816,
};
const ROTATED_90 = 5400000; // DrawingML rotation units: 60000 per degree.

function imageRun(overrides = {}) {
  return {
    type: "image",
    src: "data:image/png;base64,AAAA",
    alt: "image",
    widthPx: 120,
    heightPx: 80,
    ...overrides,
  };
}

function paragraphWith(children) {
  return { type: "paragraph", children };
}

function renderRuns(paragraph, options) {
  const rendered = renderParagraphRuns(
    paragraph,
    options?.key ?? "floating-image",
    "light",
    undefined,
    undefined,
    options?.origin,
    undefined,
    undefined,
    undefined,
    options?.runOptions,
  );
  return Array.isArray(rendered) ? rendered : [rendered];
}

function flatten(vnodes) {
  return vnodes.flatMap((node) =>
    Array.isArray(node?.children) ? [node, ...flatten(node.children)] : [node],
  );
}

function imageNode(vnodes) {
  const node = flatten(vnodes).find((child) => child?.type === "img");
  assert.ok(node, "expected a rendered image node");
  return node;
}

// --- Page origin ---------------------------------------------------------

function pageModel(nodes) {
  return { metadata: { numberingDefinitions: {} }, nodes };
}

function pageLayout() {
  return {
    pageWidthPx: 816,
    pageHeightPx: 1056,
    marginsPx: { top: 72, right: 72, bottom: 72, left: 96 },
  };
}

function pageSegments(count) {
  return Array.from({ length: count }, (_value, index) => ({ nodeIndex: index }));
}

function mountSurface(model, overrides = {}) {
  return mount(DocxPageSurface, {
    pageIndex: 0,
    pageLayout: pageLayout(),
    pageNodeSegments: pageSegments(model.nodes.length),
    model,
    pageWidthPx: 816,
    pageContentWidthPx: 624,
    pageNumber: 1,
    totalPages: 1,
    editable: false,
    ...overrides,
  });
}

function paragraphHosts(root) {
  return walk(root).filter(
    (node) => node.props?.["data-docx-paragraph-host"] === true,
  );
}

test("the page surface hands the page origin down to every paragraph host", async () => {
  const model = pageModel([
    paragraphWith([
      imageRun({
        floating: {
          xPx: 10,
          yPx: 20,
          horizontalRelativeTo: "margin",
          verticalRelativeTo: "margin",
        },
      }),
    ]),
  ]);
  const mounted = await mountSurface(model);
  try {
    const [host] = paragraphHosts(mounted.root);
    assert.ok(host, "expected a rendered paragraph host");
    const html = String(host.props.innerHTML ?? "");
    assert.match(
      html,
      /left:106px/,
      `margin-anchored images must be offset by the left margin, got: ${html}`,
    );
    assert.match(html, /top:92px/, "and by the top margin");
  } finally {
    mounted.app.unmount();
  }
});

// --- Page origin (render helper) ----------------------------------------

test("renderParagraphRuns places margin-anchored floats from the page margin", () => {
  const paragraph = paragraphWith([
    imageRun({
      floating: {
        xPx: 10,
        yPx: 20,
        horizontalRelativeTo: "margin",
        verticalRelativeTo: "margin",
      },
    }),
  ]);

  const anchored = imageNode(renderRuns(paragraph, { origin: PAGE_ORIGIN }));
  assert.equal(anchored.props.style.position, "absolute");
  assert.equal(anchored.props.style.left, 106, "xPx must be measured from the left margin");
  assert.equal(anchored.props.style.top, 92, "yPx must be measured from the top margin");

  const unanchored = imageNode(renderRuns(paragraph));
  assert.equal(unanchored.props.style.left, 10, "without an origin the anchor collapses to the page edge");
});

// --- Z-order in view mode ------------------------------------------------

test("a behindDoc image is painted under the paragraph text in view mode", () => {
  const paragraph = paragraphWith([
    { type: "text", text: "body text" },
    imageRun({
      floating: {
        xPx: 0,
        yPx: 0,
        horizontalRelativeTo: "page",
        verticalRelativeTo: "page",
        behindDocument: true,
      },
    }),
  ]);

  const nodes = flatten(renderRuns(paragraph, { origin: PAGE_ORIGIN }));
  const image = imageNode(nodes);
  const text = nodes.find((node) => node?.type === "span" && node.children === "body text");
  assert.ok(text, "expected the text run span");

  assert.ok(
    Number(image.props.style.zIndex) < 0,
    `behindDoc must paint below the text, got z-index ${image.props.style.zIndex}`,
  );
  assert.equal(
    text.props.style?.zIndex,
    undefined,
    "text runs must not be raised above the behind-text image",
  );
});

// --- Rotation ------------------------------------------------------------

test("a 90-degree rotated image is laid out in a swapped frame", () => {
  const paragraph = paragraphWith([
    imageRun({
      widthPx: 200,
      heightPx: 100,
      sourceXml: `<a:xfrm rot="${ROTATED_90}"><a:off x="0" y="0"/></a:ext cx="1905000" cy="952500"/></a:xfrm>`,
    }),
  ]);

  const nodes = renderRuns(paragraph);
  const frame = nodes.find((node) => node?.props?.style?.width === "100px");
  assert.ok(frame, "the rotated frame should be 100px wide");
  assert.equal(frame.props.style.height, "200px", "the rotated frame should be 200px tall");
  assert.equal(frame.props["data-docx-image-layout"], "inline");

  const image = imageNode(nodes);
  assert.equal(image.props.style.width, "200px", "the content keeps its unrotated width");
  assert.equal(image.props.style.height, "100px", "the content keeps its unrotated height");
  assert.match(String(image.props.style.transform), /rotate\(90deg\)/);
});

// --- Cross-paragraph exclusion -----------------------------------------

test("a float that outlives its anchor paragraph pushes the next paragraph down", async () => {
  const model = pageModel([
    paragraphWith([
      imageRun({
        widthPx: 120,
        heightPx: 240,
        floating: {
          wrapType: "square",
          wrapText: "right",
          horizontalRelativeTo: "column",
          verticalRelativeTo: "paragraph",
          xPx: 0,
          yPx: 0,
        },
      }),
      { type: "text", text: "anchor" },
    ]),
    paragraphWith([{ type: "text", text: "following paragraph" }]),
  ]);
  const mounted = await mountSurface(model);
  try {
    const hosts = paragraphHosts(mounted.root);
    assert.equal(hosts.length, 2);
    assert.equal(
      hosts[0].props.style?.marginTop,
      undefined,
      "the anchor paragraph keeps its own position",
    );
    const clearance = Number.parseInt(String(hosts[1].props.style?.marginTop ?? "0"), 10);
    assert.ok(
      clearance > 0,
      `the following paragraph must clear the float, got margin-top ${hosts[1].props.style?.marginTop}`,
    );
  } finally {
    mounted.app.unmount();
  }
});

// --- Edit-mode layer split ----------------------------------------------

function controllerWith(children) {
  const model = vue.reactive({
    nodes: [{ type: "paragraph", children }],
    metadata: { numberingDefinitions: {} },
  });
  return {
    model,
    selection: { kind: "document" },
    selectParagraph() {},
    setImageWrapMode() {},
    moveFloatingImage() {},
    resizeImage() {},
  };
}

function mountLayer(controller, extraProps = {}) {
  return mount(DocxImageLayer, {
    pageIndex: 0,
    pageWidthPx: 816,
    pageHeightPx: 1056,
    controller,
    pageNodeSegments: [{ nodeIndex: 0 }],
    editable: true,
    ...extraProps,
  });
}

async function flushUpdates() {
  for (let index = 0; index < 5; index += 1) {
    await vue.nextTick();
  }
}

test("edit mode splits floating images into a behind-text layer and a front layer", async () => {
  const controller = controllerWith([
    imageRun({
      floating: {
        xPx: 0,
        yPx: 0,
        horizontalRelativeTo: "page",
        verticalRelativeTo: "page",
        behindDocument: true,
      },
    }),
    imageRun({
      floating: {
        xPx: 40,
        yPx: 60,
        horizontalRelativeTo: "page",
        verticalRelativeTo: "page",
      },
    }),
  ]);
  const mounted = await mountLayer(controller);
  try {
    const nodes = walk(mounted.root);
    const layers = nodes.filter((node) => node.props?.["data-docx-image-layer"]);
    assert.deepEqual(
      layers.map((layer) => layer.props["data-docx-image-layer"]),
      ["behind", "front"],
      "behind-text images render in their own layer below the text",
    );

    const [behindLayer, frontLayer] = layers;
    assert.ok(
      Number(behindLayer.props.style.zIndex) < 0,
      `the behind layer must sit under the text, got z-index ${behindLayer.props.style.zIndex}`,
    );
    assert.ok(
      Number(frontLayer.props.style.zIndex) > 0,
      `the front layer must sit above the text, got z-index ${frontLayer.props.style.zIndex}`,
    );

    const containers = nodes.filter(
      (node) => node.props?.["data-testid"] === "docx-floating-image",
    );
    assert.equal(containers.length, 2);

    const behindContainer = containers.find(
      (node) => node.props["data-image-behind"] === "true",
    );
    const frontContainer = containers.find(
      (node) => node.props["data-image-behind"] !== "true",
    );
    assert.ok(behindContainer, "the behindDoc image keeps its identity");
    assert.ok(frontContainer);

    assert.equal(behindContainer.parent, behindLayer, "behindDoc image belongs to the behind layer");
    assert.equal(frontContainer.parent, frontLayer, "in-front image belongs to the front layer");
    assert.equal(behindContainer.props.style.pointerEvents, "none", "behind-text images never steal text clicks");
    assert.equal(frontContainer.props.style.pointerEvents, "auto");
  } finally {
    mounted.app.unmount();
  }
});

// --- Hit box re-measurement --------------------------------------------

function installLayoutEnvironment(t) {
  const originalWindow = globalThis.window;
  const originalResizeObserver = globalThis.ResizeObserver;
  const resizeListeners = new Set();

  globalThis.window = {
    addEventListener(type, listener) {
      if (type === "resize") resizeListeners.add(listener);
    },
    removeEventListener(type, listener) {
      if (type === "resize") resizeListeners.delete(listener);
    },
  };
  globalThis.ResizeObserver = class ResizeObserver {
    observe() {}
    unobserve() {}
    disconnect() {}
  };

  t.after(() => {
    if (originalWindow === undefined) delete globalThis.window;
    else globalThis.window = originalWindow;
    if (originalResizeObserver === undefined) delete globalThis.ResizeObserver;
    else globalThis.ResizeObserver = originalResizeObserver;
  });

  return {
    fireResize() {
      for (const listener of [...resizeListeners]) listener({ type: "resize" });
    },
  };
}

test("the wrapped drag target follows the paragraph after a text edit", async (t) => {
  const environment = installLayoutEnvironment(t);
  const controller = controllerWith([
    imageRun({
      widthPx: 120,
      heightPx: 80,
      floating: {
        wrapType: "square",
        horizontalRelativeTo: "column",
        verticalRelativeTo: "paragraph",
        xPx: 0,
        yPx: 0,
        wrapText: "right",
      },
    }),
    { type: "text", text: "first line" },
  ]);

  const mounted = await mountLayer(controller);
  try {
    const root = walk(mounted.root).find(
      (node) => node.props?.class === "docx-image-layer-root",
    );
    assert.ok(root, "expected the image layer root");

    const wrappedRect = { left: 96, top: 120, width: 120, height: 80 };
    const wrappedElement = { getBoundingClientRect: () => ({ ...wrappedRect }) };
    const surface = {
      offsetWidth: 816,
      getBoundingClientRect: () => ({
        left: 0,
        top: 0,
        right: 816,
        bottom: 1056,
        width: 816,
        height: 1056,
      }),
      querySelector: () => ({ querySelector: () => wrappedElement }),
    };
    // The test renderer's host nodes have no DOM, so the component resolves the
    // page surface through the element it received via its template ref.
    root.closest = (selector) =>
      selector === "[data-docx-page-surface='true']" ? surface : null;

    environment.fireResize();
    await flushUpdates();

    const containerNode = () =>
      walk(mounted.root).find(
        (node) => node.props?.["data-testid"] === "docx-floating-image",
      );
    const firstBox = containerNode().props.style;
    assert.equal(firstBox.left, "96px", "the drag target matches the wrapped image");
    assert.equal(firstBox.top, "120px");
    assert.equal(firstBox.width, "120px");
    assert.equal(firstBox.height, "80px");

    // A text edit above the image moves it down; the cached hit box must follow
    // without waiting for an x/y change on the image itself.
    wrappedRect.top = 168;
    controller.model.nodes[0].children.push({ type: "text", text: " edited" });
    await flushUpdates();

    const nextBox = containerNode().props.style;
    assert.equal(nextBox.top, "168px", "the drag target must move with the paragraph after a text edit");
    assert.equal(nextBox.left, "96px");
  } finally {
    mounted.app.unmount();
  }
});
