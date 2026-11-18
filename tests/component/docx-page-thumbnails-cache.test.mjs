import assert from "node:assert/strict";
import test from "node:test";

import { importFromDemo, vue } from "./vue-test-renderer.mjs";

const vueDocx = await importFromDemo("@arcships/vue-docx");

class FakeCanvas {
  constructor() {
    this.width = 0;
    this.height = 0;
    this.style = {};
    this.context2d = null;
  }

  getContext(kind) {
    if (kind !== "2d") return null;
    if (!this.context2d) {
      this.context2d = {
        setTransform() {},
        clearRect() {},
        fillRect() {},
        strokeRect() {},
        save() {},
        restore() {},
        scale() {},
        beginPath() {},
        rect() {},
        clip() {},
        fillText() {},
        moveTo() {},
        lineTo() {},
        stroke() {},
        drawImage() {},
        measureText: (value) => ({ width: String(value).length * 8 }),
        imageSmoothingEnabled: false,
        imageSmoothingQuality: "",
      };
    }
    return this.context2d;
  }
}

function fakePageElement() {
  return {
    cloneNode() {
      return { querySelectorAll: () => [] };
    },
  };
}

function installBrowserEnvironment(t) {
  const originalDocument = globalThis.document;
  const originalWindow = globalThis.window;
  const originalImage = globalThis.Image;
  const originalXmlSerializer = globalThis.XMLSerializer;
  const createdCanvases = [];

  globalThis.document = {
    createElement(tagName) {
      if (tagName !== "canvas") {
        throw new Error(`Unexpected element: ${tagName}`);
      }
      const canvas = new FakeCanvas();
      createdCanvases.push(canvas);
      return canvas;
    },
  };
  globalThis.window = { devicePixelRatio: 1 };
  globalThis.XMLSerializer = class {
    serializeToString() {
      return "<div></div>";
    }
  };
  globalThis.Image = class {
    constructor() {
      this.naturalWidth = 100;
      this.naturalHeight = 200;
      this.width = 100;
      this.height = 200;
    }

    decode() {
      return Promise.resolve();
    }

    set src(value) {
      this._src = value;
    }

    get src() {
      return this._src;
    }
  };

  t.after(() => {
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
    if (originalWindow === undefined) delete globalThis.window;
    else globalThis.window = originalWindow;
    if (originalImage === undefined) delete globalThis.Image;
    else globalThis.Image = originalImage;
    if (originalXmlSerializer === undefined) delete globalThis.XMLSerializer;
    else globalThis.XMLSerializer = originalXmlSerializer;
  });

  return createdCanvases;
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function waitFor(predicate, message, timeoutMs = 2_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    await vue.nextTick();
    if (predicate()) return;
    await delay(5);
  }
  throw new Error(`Timed out waiting for ${message}`);
}

test("DOCX page surface registry reports per-page changes and has no dead snapshot map", () => {
  const registry = vueDocx.createDocxViewerPageSurfaceRegistry();
  assert.equal(
    Object.hasOwn(registry, "pageThumbnailSnapshots"),
    false,
    "the unwired snapshot registry path must be removed",
  );

  const editor = { syncPaginationInfo() {} };
  const changes = [];
  const unsubscribe = vueDocx.subscribeDocxViewerPageSurfaces(editor, (change) => {
    changes.push(change);
  });
  vueDocx.notifyDocxViewerPageSurfaceListeners(editor, {
    pageIndex: 3,
    kind: "publish",
    contentKey: "page-3",
  });
  vueDocx.notifyDocxViewerPageSurfaceListeners(editor, {
    pageIndex: 4,
    kind: "unpublish",
  });
  unsubscribe();
  vueDocx.notifyDocxViewerPageSurfaceListeners(editor, {
    pageIndex: 5,
    kind: "publish",
  });

  assert.deepEqual(changes, [
    { pageIndex: 3, kind: "publish", contentKey: "page-3" },
    { pageIndex: 4, kind: "unpublish" },
  ]);
});

test("DOCX page thumbnails re-raster on resolution changes and invalidate per page", async (t) => {
  const createdCanvases = installBrowserEnvironment(t);
  const editor = { syncPaginationInfo() {} };
  const registry = vueDocx.ensureDocxViewerPageSurfaceRegistry(editor);
  registry.pageElements.set(0, fakePageElement());
  registry.pageElements.set(1, fakePageElement());
  registry.pageContentKeys.set(0, "page-0-v1");
  registry.pageContentKeys.set(1, "page-1-v1");
  registry.pageSizes.set(0, { widthPx: 100, heightPx: 200 });
  registry.pageSizes.set(1, { widthPx: 100, heightPx: 200 });

  const thumbnails = vueDocx.useDocxPageThumbnails(editor, {
    sourceWidthPx: 100,
    sourceHeightPx: 200,
    maxWidthPx: 50,
  });
  const canvas0 = new FakeCanvas();
  const canvas1 = new FakeCanvas();
  thumbnails.attachCanvasForPage(0, canvas0);
  thumbnails.attachCanvasForPage(1, canvas1);

  await waitFor(
    () =>
      thumbnails.pageThumbnailStates.get(0)?.status === "ready" &&
      thumbnails.pageThumbnailStates.get(1)?.status === "ready",
    "initial thumbnail rasters",
  );
  assert.equal(canvas0.width, 50);
  assert.equal(canvas0.height, 100);
  assert.equal(createdCanvases.length, 2);

  // A cache hit at the same content key and resolution must not re-raster.
  await thumbnails.renderPageThumbnailToCanvas(1, canvas1);
  assert.equal(createdCanvases.length, 2);

  // Same content key, different page size: the cached surface has the wrong
  // aspect ratio and must be replaced.
  registry.pageSizes.set(0, { widthPx: 200, heightPx: 200 });
  await thumbnails.renderPageThumbnailToCanvas(0, canvas0);
  assert.equal(canvas0.width, 50);
  assert.equal(canvas0.height, 50);
  assert.equal(createdCanvases.length, 3);

  // A DPR change must produce new pixel dimensions instead of reusing the
  // cached surface.
  globalThis.window.devicePixelRatio = 2;
  await thumbnails.renderPageThumbnailToCanvas(0, canvas0);
  assert.equal(canvas0.width, 100);
  assert.equal(canvas0.height, 100);
  assert.equal(createdCanvases.length, 4);

  // Coalesce repeated publish notifications for one page into one raster, and
  // leave the other page's cache alone.
  const beforePublish = createdCanvases.length;
  registry.pageContentKeys.set(0, "page-0-v2");
  vueDocx.notifyDocxViewerPageSurfaceListeners(editor, {
    pageIndex: 0,
    kind: "publish",
    contentKey: "page-0-v2",
  });
  vueDocx.notifyDocxViewerPageSurfaceListeners(editor, {
    pageIndex: 0,
    kind: "publish",
    contentKey: "page-0-v2",
  });
  await waitFor(
    () => createdCanvases.length === beforePublish + 1,
    "coalesced page 0 raster",
  );
  await delay(40);
  assert.equal(createdCanvases.length, beforePublish + 1);
  assert.equal(canvas1.width, 50, "page 1 must keep its cached surface");
  assert.equal(canvas1.height, 100);

  // Unpublishing a page keeps its last good thumbnail instead of painting the
  // unavailable placeholder.
  const beforeUnpublish = createdCanvases.length;
  vueDocx.notifyDocxViewerPageSurfaceListeners(editor, {
    pageIndex: 0,
    kind: "unpublish",
  });
  await delay(40);
  assert.equal(createdCanvases.length, beforeUnpublish);
  assert.equal(canvas0.width, 100);
  assert.equal(thumbnails.pageThumbnailStates.get(0)?.status, "ready");
});
