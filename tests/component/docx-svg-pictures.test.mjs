import assert from "node:assert/strict";
import { test } from "node:test";

import {
  importFromDemo,
  mount,
  textContent,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

const { DocxImageLayer, DocxPageSurface, renderParagraphRuns } =
  await importFromDemo("@arcships/vue-docx");

const SVG_MARKUP =
  '<svg xmlns="http://www.w3.org/2000/svg" width="96" height="96" viewBox="0 0 96 96">' +
  '<rect width="96" height="96" fill="#2563eb"/>' +
  '<circle cx="48" cy="48" r="22" fill="#fbbf24"/>' +
  "</svg>";

const HOSTILE_MARKUP =
  '<svg xmlns="http://www.w3.org/2000/svg" width="160" height="96" viewBox="0 0 160 96">' +
  "<script>window.__pwned = true;</script>" +
  '<rect width="160" height="96" fill="#fee2e2" onload="window.__pwned = true"/>' +
  '<circle cx="80" cy="48" r="28" fill="#fbbf24"/>' +
  '<image href="https://example.invalid/tracker.png" width="8" height="8"/>' +
  "</svg>";

const FOREIGN_OBJECT_MARKUP =
  '<svg xmlns="http://www.w3.org/2000/svg" width="160" height="96" viewBox="0 0 160 96">' +
  '<foreignObject x="0" y="0" width="160" height="96">' +
  '<div xmlns="http://www.w3.org/1999/xhtml">HTML</div>' +
  "</foreignObject></svg>";

const FILTER_MARKUP =
  '<svg xmlns="http://www.w3.org/2000/svg" width="160" height="96" viewBox="0 0 160 96">' +
  '<defs><filter id="blur"><feGaussianBlur stdDeviation="4"/></filter></defs>' +
  '<circle cx="80" cy="48" r="28" fill="#6366f1" filter="url(#blur)"/>' +
  "</svg>";

function svgDataUri(markup) {
  return `data:image/svg+xml;base64,${Buffer.from(markup, "utf8").toString("base64")}`;
}

function imageRun(markup, overrides = {}) {
  return {
    type: "image",
    src: svgDataUri(markup),
    contentType: "image/svg+xml",
    partName: "word/media/picture.svg",
    alt: "svg picture",
    widthPx: 160,
    heightPx: 96,
    ...overrides,
  };
}

function floatingImageRun(markup, overrides = {}) {
  return {
    ...imageRun(markup, { widthPx: 200, heightPx: 120, ...overrides }),
    floating: {
      xPx: 24,
      yPx: 16,
      wrapType: "none",
      behindDocument: false,
      ...overrides.floating,
    },
  };
}

function paragraphWith(children) {
  return { type: "paragraph", children };
}

function renderRuns(paragraph, options) {
  const rendered = renderParagraphRuns(
    paragraph,
    options?.key ?? "svg-picture",
    "light",
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    options?.runOptions,
  );
  return Array.isArray(rendered) ? rendered : [rendered];
}

function vnodeText(value, result = []) {
  if (Array.isArray(value)) {
    for (const item of value) vnodeText(item, result);
    return result;
  }
  if (!value || typeof value !== "object") {
    if (typeof value === "string" && value.trim()) result.push(value.trim());
    return result;
  }
  vnodeText(value.children, result);
  return result;
}

function flatten(vnodes) {
  return vnodes.flatMap((node) =>
    Array.isArray(node?.children) ? [node, ...flatten(node.children)] : [node],
  );
}

function decodeSvgSource(source) {
  assert.match(source, /^data:image\/svg\+xml;charset=utf-8,/u);
  return decodeURIComponent(source.slice(source.indexOf(",") + 1));
}

function imageModel(image) {
  return {
    nodes: [paragraphWith([image])],
    metadata: {
      sourceParts: 0,
      warnings: [],
      headerSections: [],
      footerSections: [],
      paragraphStyles: [],
    },
  };
}

function pageLayout() {
  return {
    pageWidthPx: 816,
    pageHeightPx: 1056,
    marginsPx: { top: 72, right: 72, bottom: 72, left: 96 },
  };
}

function mountSurface(model, overrides = {}) {
  return mount(DocxPageSurface, {
    pageIndex: 0,
    pageLayout: pageLayout(),
    pageNodeSegments: [{ nodeIndex: 0 }],
    model,
    pageWidthPx: 816,
    pageContentWidthPx: 624,
    pageNumber: 1,
    totalPages: 1,
    editable: false,
    ...overrides,
  });
}

test("DOCX inline SVG pictures render as sanitised vector images at the anchored size", () => {
  const diagnostics = [];
  const rendered = renderRuns(
    paragraphWith([imageRun(HOSTILE_MARKUP)]),
    {
      runOptions: {
        onImageUnsupported: (diagnostic) => diagnostics.push(diagnostic),
      },
    },
  );
  const image = flatten(rendered).find((node) => node?.type === "img");
  assert.ok(image, "expected an <img> for a supported SVG picture");
  assert.equal(image.props.alt, "svg picture");
  assert.equal(image.props.style.width, "160px");
  assert.equal(image.props.style.height, "96px");
  assert.match(image.props.src, /^data:image\/svg\+xml;charset=utf-8,/u);
  const markup = decodeSvgSource(image.props.src);
  assert.match(markup, /<circle/u);
  assert.doesNotMatch(markup, /<script/iu);
  assert.doesNotMatch(markup, /onload/iu);
  assert.doesNotMatch(markup, /https?:\/\/(?!www\.w3\.org\/)/u);
  assert.deepEqual(diagnostics, []);

  // Non-SVG images keep their original source untouched.
  const png = flatten(
    renderRuns(
      paragraphWith([
        {
          type: "image",
          src: "data:image/png;base64,AAAA",
          contentType: "image/png",
          alt: "png",
          widthPx: 40,
          heightPx: 20,
        },
      ]),
    ),
  ).find((node) => node?.type === "img");
  assert.equal(png.props.src, "data:image/png;base64,AAAA");
});

test("DOCX unsupported SVG features render a placeholder and report a diagnostic", () => {
  const diagnostics = [];
  for (const [markup, reason] of [
    [FOREIGN_OBJECT_MARKUP, "foreignObject"],
    [FILTER_MARKUP, "filter"],
  ]) {
    diagnostics.length = 0;
    const rendered = renderRuns(paragraphWith([imageRun(markup)]), {
      key: `svg-${reason}`,
      runOptions: {
        onImageUnsupported: (diagnostic) => diagnostics.push(diagnostic),
      },
    });
    assert.equal(
      flatten(rendered).some((node) => node?.type === "img"),
      false,
      `${reason}: an unsupported picture must not reach an <img>`,
    );
    const placeholder = flatten(rendered).find(
      (node) => node?.type === "span" && node.props?.role === "img",
    );
    assert.ok(placeholder, `${reason}: expected a placeholder`);
    assert.equal(placeholder.props["aria-label"], "svg picture");
    assert.ok(vnodeText(placeholder).join(" ").includes("SVG"), `${reason}: placeholder label`);
    assert.deepEqual(diagnostics, [
      {
        code: "IMAGE_UNSUPPORTED_FEATURE",
        phase: "image-render",
        imageKey: `svg-${reason}-run-0`,
        reason,
        partName: "word/media/picture.svg",
        message: `DOCX SVG picture uses unsupported ${reason} content; showing a placeholder.`,
      },
    ]);
  }

  // A host can supply its own localized template.
  const localized = [];
  renderRuns(paragraphWith([imageRun(FOREIGN_OBJECT_MARKUP)]), {
    runOptions: {
      onImageUnsupported: (diagnostic) => localized.push(diagnostic),
      imageUnsupportedText: "Bild nicht darstellbar: {feature}",
    },
  });
  assert.equal(localized[0]?.message, "Bild nicht darstellbar: foreignObject");
});

test("DOCX floating SVG pictures stay vector at zoom and fall back when unsupported", async () => {
  const controller = vue.reactive({
    model: imageModel(floatingImageRun(SVG_MARKUP)),
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
    zoomFactor: 2,
  });
  try {
    const image = walk(mounted.root).find((node) => node.type === "img");
    assert.ok(image, "expected the floating SVG to render");
    assert.match(
      image.props.src,
      /^data:image\/svg\+xml;charset=utf-8,/u,
      "a 200% zoom must not switch the picture to a raster source",
    );
    assert.equal(image.props.alt, "svg picture");
    const container = walk(mounted.root).find(
      (node) => node.props?.["data-testid"] === "docx-floating-image",
    );
    assert.equal(container.props.style.width, "200px");
    assert.equal(container.props.style.height, "120px");

    // Switch to a picture the pipeline cannot paint: the layer keeps the
    // anchored box and labels it instead of drawing a broken icon.
    controller.model = imageModel(floatingImageRun(FOREIGN_OBJECT_MARKUP));
    await vue.nextTick();
    assert.equal(
      walk(mounted.root).some((node) => node.type === "img"),
      false,
    );
    const fallback = walk(mounted.root).find(
      (node) => node.props?.["data-image-state"] === "unsupported",
    );
    assert.ok(fallback, "expected an unsupported-image fallback");
    assert.ok(textContent(fallback).includes("SVG"));
  } finally {
    mounted.app.unmount();
  }
});

test("DOCX page surface keeps SVG pictures vector at 200% zoom", async () => {
  const model = {
    ...imageModel(imageRun(SVG_MARKUP, { widthPx: 96, heightPx: 96 })),
    metadata: {
      ...imageModel(imageRun(SVG_MARKUP)).metadata,
      paragraphStyles: [],
    },
  };
  const mounted = await mountSurface(model, { zoomFactor: 2 });
  try {
    const host = walk(mounted.root).find(
      (node) => node.props?.["data-docx-paragraph-host"] === true,
    );
    assert.ok(host, "expected the paragraph host to render");
    assert.match(host.props.innerHTML, /<img[^>]+data:image\/svg\+xml;charset=utf-8,/u);
    assert.match(host.props.innerHTML, /width:96px/u);
    assert.match(host.props.innerHTML, /height:96px/u);
    assert.doesNotMatch(host.props.innerHTML, /data:image\/png/u);
  } finally {
    mounted.app.unmount();
  }
});
