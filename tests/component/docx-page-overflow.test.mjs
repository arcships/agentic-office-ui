import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import { importFromDemo, mount, vue, walk } from "./vue-test-renderer.mjs";

const { DocxViewer } = await importFromDemo("@arcships/vue-docx");
const { wasmBuildDocModelFromBytes } = await importFromDemo("@arcships/docx-core");

const PAGE_HEIGHT_PX = 1056;

function prop(node, name) {
  return node?.props?.[name];
}

function isPageWrapper(node) {
  const value = prop(node, "data-docx-page-wrapper");
  return value === true || value === "true";
}

function isPageSurface(node) {
  const value = prop(node, "data-docx-page-surface");
  return value === true || value === "true";
}

function descendants(node) {
  return walk(node, []).slice(1);
}

function classes(node) {
  return String(node?.props?.class || "").split(/\s+/).filter(Boolean);
}

async function loadModel(fixture) {
  const file = readFileSync(
    new URL(`../../apps/demo/public/samples/${fixture}`, import.meta.url),
  );
  const bytes = file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength);
  const { model } = await wasmBuildDocModelFromBytes(bytes);
  return model;
}

test("rendered DOCX pages are exactly one paper tall and clip pagination overflow", async () => {
  // legal-contract.docx carries more paragraphs than the estimator assigns to
  // each page, which is the fixture that used to grow the surface past 1525px.
  const model = await loadModel("legal-contract.docx");
  const mounted = await mount(DocxViewer, { model, showToolbar: false });
  await vue.nextTick();
  await vue.nextTick();

  const root = walk(mounted.root).find((node) =>
    prop(node, "data-docx-page-count") !== undefined,
  );
  assert.ok(root, "DocxViewer renders its scrollable page root");
  const pageCount = Number(prop(root, "data-docx-page-count"));
  assert.ok(pageCount > 1, `legal fixture must span pages, received ${pageCount}`);

  const wrappers = walk(mounted.root).filter(isPageWrapper);
  assert.ok(wrappers.length > 0, "at least the first page is mounted");

  for (const wrapper of wrappers) {
    const pageIndex = Number(prop(wrapper, "data-docx-page-index"));
    const surface = descendants(wrapper).find(isPageSurface);
    assert.ok(surface, `page ${pageIndex} renders a page surface`);

    const surfaceStyle = prop(surface, "style") || {};
    assert.equal(
      surfaceStyle.height,
      `${PAGE_HEIGHT_PX}px`,
      `page ${pageIndex} surface must have a fixed paper height`,
    );
    assert.equal(
      surfaceStyle.minHeight,
      undefined,
      `page ${pageIndex} must not use minHeight, which lets content grow the paper`,
    );
    assert.equal(
      surfaceStyle.overflow,
      "hidden",
      `page ${pageIndex} surface must clip content at the paper edge`,
    );
    assert.equal(surfaceStyle.boxSizing, "border-box");
    assert.equal(
      Number(prop(surface, "data-docx-page-paper-height")),
      PAGE_HEIGHT_PX,
      `page ${pageIndex} exposes the paper height used for layout`,
    );

    const frame = wrapper.children?.find((node) =>
      descendants(node).includes(surface),
    );
    assert.ok(frame, `page ${pageIndex} surface has a framing element`);
    assert.equal(
      prop(frame, "style")?.height,
      `${PAGE_HEIGHT_PX}px`,
      `page ${pageIndex} frame is one paper tall`,
    );

    const wrapperStyle = prop(wrapper, "style") || {};
    const wrapperHeight = Number.parseFloat(wrapperStyle.height);
    assert.ok(
      Math.abs(wrapperHeight - PAGE_HEIGHT_PX) <= 1,
      `page ${pageIndex} wrapper height ${wrapperStyle.height} must match the paper height`,
    );

    const body = descendants(surface).find(
      (node) => prop(node, "data-docx-page-body") === "true",
    );
    assert.ok(body, `page ${pageIndex} renders a page body`);
    const bodyStyle = prop(body, "style") || {};
    assert.equal(
      bodyStyle.overflow,
      "hidden",
      `page ${pageIndex} body clips content instead of drawing over the next page`,
    );
    assert.equal(bodyStyle.minHeight, "0px");
    assert.equal(bodyStyle.flex, "0 1 auto");
  }

  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("annotation gutter stays outside the clipped page surface", async () => {
  const model = await loadModel("review-comments.docx");
  const mounted = await mount(DocxViewer, {
    model,
    showToolbar: false,
    showTrackedChanges: true,
    showComments: true,
  });
  await vue.nextTick();
  await vue.nextTick();

  const surface = walk(mounted.root).find(isPageSurface);
  assert.ok(surface, "a page surface is mounted");

  const gutter = walk(mounted.root).find((node) =>
    classes(node).includes("docx-tracked-change-gutter"),
  );
  assert.ok(gutter, "the tracked-change gutter is rendered");

  assert.ok(
    !descendants(surface).includes(gutter),
    "the gutter must not be clipped by the page surface",
  );
  assert.equal(
    gutter.parent,
    surface.parent,
    "the gutter is a sibling of the paper inside the page frame",
  );

  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});
