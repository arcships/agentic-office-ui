import assert from "node:assert/strict";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  textContent,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

const { PdfRenderError, PdfSurface, PdfViewer } = await importFromDemo("@arcships/vue-pdf");

const PDF_SOURCE = {
  kind: "blob",
  blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }),
};

function createRuntime({ pageCount = 3, failOpen } = {}) {
  const calls = { dispose: [], openDocument: [], renderPage: [], renderThumbnail: [] };
  const document = {
    id: "pdf-chrome-document",
    pageCount,
    pages: Array.from({ length: pageCount }, (_, index) => ({
      index,
      width: 600,
      height: 800,
      rotation: 0,
    })),
  };
  return {
    calls,
    document,
    runtime: {
      runtimeId: "pdf-chrome-runtime",
      async openDocument() {
        calls.openDocument.push(true);
        if (failOpen) throw failOpen;
        return document;
      },
      async renderPage(_document, pageIndex) {
        calls.renderPage.push(pageIndex);
        return new Blob([`page:${pageIndex}`], { type: "image/png" });
      },
      async renderThumbnail(_document, pageIndex) {
        calls.renderThumbnail.push(pageIndex);
        return new Blob([`thumbnail:${pageIndex}`], { type: "image/png" });
      },
      async getPageLabels() {
        return ["1", "2", "3"].slice(0, pageCount);
      },
      async getOutline() {
        return [];
      },
      async getPageTextGeometry() {
        return { runs: [] };
      },
      async getTextSlices(_document, slices) {
        return slices.map(() => "");
      },
      async search() {
        return [];
      },
      async closeDocument() {},
      async dispose() {
        calls.dispose.push(true);
      },
    },
  };
}

function stubObjectUrls(t) {
  const originalCreate = URL.createObjectURL;
  const originalRevoke = URL.revokeObjectURL;
  let counter = 0;
  URL.createObjectURL = () => `blob:pdf-chrome-${++counter}`;
  URL.revokeObjectURL = () => {};
  t.after(() => {
    URL.createObjectURL = originalCreate;
    URL.revokeObjectURL = originalRevoke;
  });
}

function classNames(node) {
  return String(node?.props?.class || "").split(/\s+/).filter(Boolean);
}

function nodesWithClass(root, className) {
  return walk(root).filter((node) => classNames(node).includes(className));
}

function allByTestId(root, testId) {
  return walk(root).filter((node) => node.props?.["data-testid"] === testId);
}

test("PdfViewer toolbar is the shared OfficeToolbar with every PDF command", async (t) => {
  stubObjectUrls(t);
  const fake = createRuntime();
  const mounted = await mount(PdfViewer, {
    source: PDF_SOURCE,
    runtime: fake.runtime,
    defaultZoom: 1,
  });
  try {
    await waitFor(
      () => findByTestId(mounted.root, "pdf-status")?.props["data-state"] === "ready",
      "rendered PDF page",
    );

    const toolbar = findByTestId(mounted.root, "pdf-toolbar");
    assert.ok(toolbar, "PDF toolbar");
    assert.equal(toolbar.props.role, "toolbar");
    assert.ok(classNames(toolbar).includes("office-toolbar"), "toolbar uses the shared chrome class");
    assert.ok(classNames(toolbar).includes("pdf-toolbar"), "the public print/touch hook stays");
    assert.ok(walk(toolbar).some((node) => classNames(node).includes("office-toolbar__leading")));
    assert.ok(walk(toolbar).some((node) => classNames(node).includes("office-toolbar__trailing")));
    assert.ok(
      nodesWithClass(toolbar, "office-icon-button").length >= 10,
      "sidebar, paging, zoom, rotate, search, print and download are shared icon buttons",
    );

    const pageInput = findByTestId(mounted.root, "pdf-page-input");
    assert.equal(pageInput.type, "input");
    assert.equal(pageInput.props.type, "number");
    assert.equal(pageInput.props["aria-label"], "Page number");
    assert.equal(pageInput.value, "1");
    assert.equal(findByTestId(mounted.root, "pdf-page-count").children[0].text, "3");
    assert.equal(findByTestId(mounted.root, "pdf-zoom-select").props["aria-label"], "Zoom level");
    assert.equal(findByTestId(mounted.root, "pdf-zoom-select").options.length, 6);
    assert.ok(findByTestId(mounted.root, "pdf-rotate"));
    assert.ok(findByTestId(mounted.root, "pdf-search-input"));
    assert.ok(findByTestId(mounted.root, "pdf-search-submit"));
    assert.ok(findByTestId(mounted.root, "pdf-print"));
    assert.ok(findByTestId(mounted.root, "pdf-download"));

    // The hidden live region is unchanged: status is still announced.
    assert.equal(findByTestId(mounted.root, "pdf-status").props.role, "status");
    assert.equal(findByTestId(mounted.root, "pdf-status").props["data-state"], "ready");
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("PdfViewer print command rasterizes every page through the toolbar", async (t) => {
  stubObjectUrls(t);
  const previousWindow = globalThis.window;
  let printed = 0;
  globalThis.window = {
    print: () => { printed += 1; },
    matchMedia: () => ({ matches: true, addEventListener() {}, removeEventListener() {} }),
  };
  try {
    const fake = createRuntime();
    const mounted = await mount(PdfViewer, { source: PDF_SOURCE, runtime: fake.runtime });
    await waitFor(
      () => findByTestId(mounted.root, "pdf-status")?.props["data-state"] === "ready",
      "rendered PDF page",
    );
    findByTestId(mounted.root, "pdf-print").props.onClick();
    await waitFor(() => printed === 1, "browser print dialog");
    assert.equal(fake.calls.renderPage.length, 1 + 3, "print renders every page, not just the active one");
    assert.ok(findByTestId(mounted.root, "pdf-print-pages"), "the print snapshot page set is mounted");
    assert.deepEqual(mounted.warnings, []);
    mounted.app.unmount();
  } finally {
    if (previousWindow === undefined) delete globalThis.window;
    else globalThis.window = previousWindow;
  }
});

test("PdfViewer sidebar is the shared document sidebar and keeps the thumbnail strip", async (t) => {
  stubObjectUrls(t);
  const fake = createRuntime();
  const mounted = await mount(PdfViewer, { source: PDF_SOURCE, runtime: fake.runtime });
  try {
    await waitFor(() => allByTestId(mounted.root, "pdf-thumbnail-image").length === 3, "thumbnails");
    const sidebar = findByTestId(mounted.root, "pdf-thumbnails-sidebar");
    assert.ok(sidebar, "viewer thumbnail sidebar");
    assert.ok(classNames(sidebar).includes("office-document-sidebar"));
    assert.equal(sidebar.props["data-sidebar-mode"], "inline");
    assert.equal(sidebar.props["data-sidebar-open"], "true");
    assert.match(sidebar.props.style["--office-ui-sidebar-width"], /160/);
    assert.ok(
      findByTestId(mounted.root, "pdf-thumbnails"),
      "the existing thumbnail strip stays inside the shared sidebar",
    );
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("PdfViewer renders the shared status bar on request", async (t) => {
  stubObjectUrls(t);
  const fake = createRuntime();
  const mounted = await mount(PdfViewer, {
    source: PDF_SOURCE,
    runtime: fake.runtime,
    fileName: "report.pdf",
    showStatusBar: true,
  });
  try {
    await waitFor(
      () => findByTestId(mounted.root, "pdf-status")?.props["data-state"] === "ready",
      "rendered PDF page",
    );
    const bar = findByTestId(mounted.root, "pdf-status-bar");
    assert.ok(bar, "visible status bar");
    assert.ok(classNames(bar).includes("office-status-bar"));
    assert.match(textContent(findByTestId(mounted.root, "office-status-bar-page")), /1.*3/);
    assert.match(textContent(findByTestId(mounted.root, "office-status-bar-file")), /report\.pdf/);
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("PdfViewer renders empty, loading, error and encrypted states with Office states", async (t) => {
  stubObjectUrls(t);
  const empty = await mount(PdfViewer, {});
  assert.ok(classNames(findByTestId(empty.root, "pdf-empty")).includes("office-state--empty"));
  assert.match(textContent(findByTestId(empty.root, "pdf-empty")), /No PDF loaded/);
  empty.app.unmount();

  const encrypted = await mount(PdfViewer, {
    source: PDF_SOURCE,
    runtime: createRuntime({
      failOpen: new PdfRenderError("ENCRYPTED", "password required"),
    }).runtime,
  });
  try {
    await waitFor(
      () => findByTestId(encrypted.root, "load-error") !== undefined,
      "encrypted PDF error",
    );
    const error = findByTestId(encrypted.root, "load-error");
    assert.ok(classNames(error).includes("office-state--error"));
    assert.equal(error.props.role, "alert");
    assert.equal(error.props["data-error-code"], "ENCRYPTED");
    assert.equal(error.props["data-encrypted"], "true");
    assert.match(textContent(error), /encrypted and requires a password/);
    assert.match(textContent(error), /Enter the document password/);
    assert.deepEqual(encrypted.warnings, []);
  } finally {
    encrypted.app.unmount();
  }
});

test("PdfSurface sidebar, states and status use the shared chrome", async (t) => {
  stubObjectUrls(t);
  const fake = createRuntime();
  const mounted = await mount(PdfSurface, {
    source: PDF_SOURCE,
    runtime: fake.runtime,
    showStatusBar: true,
  });
  try {
    await waitFor(() => findByTestId(mounted.root, "pdf-surface-thumbnail") !== undefined, "surface rail");
    const nav = findByTestId(mounted.root, "pdf-surface-nav");
    assert.ok(classNames(nav).includes("office-document-sidebar"));
    assert.ok(
      findByTestId(mounted.root, "pdf-surface-thumbnails"),
      "surface thumbnails stay inside the shared sidebar",
    );
    const bar = findByTestId(mounted.root, "pdf-surface-status-bar");
    assert.ok(bar, "surface status bar");
    assert.ok(classNames(bar).includes("office-status-bar"));
    assert.match(textContent(findByTestId(mounted.root, "office-status-bar-page")), /1.*3/);
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }

  const failed = await mount(PdfSurface, {
    source: { kind: "blob", blob: new Blob(["not pdf"], { type: "text/plain" }) },
    runtime: createRuntime().runtime,
  });
  try {
    await waitFor(
      () => findByTestId(failed.root, "pdf-surface-load-error") !== undefined,
      "surface load error",
    );
    const error = findByTestId(failed.root, "pdf-surface-load-error");
    assert.ok(classNames(error).includes("office-state--error"));
    assert.equal(error.props.role, "alert");
    assert.equal(error.props["data-error-code"], "INVALID_PDF");
    assert.deepEqual(failed.warnings, []);
  } finally {
    failed.app.unmount();
  }
});
