import assert from "node:assert/strict";
import test from "node:test";

import {
  importFromDemo,
  mount,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

const {
  DOCX_THUMBNAIL_EXCLUDE_ATTRIBUTE,
  createBlankDocumentModel,
  getDownscaledThumbnailImageDataUri,
  rasterizeDocxThumbnailSurface,
  renderDocxThumbnailSnapshotSurface,
} = await importFromDemo("@arcships/docx-core");
const { DocxImageLayer, DocxTableHost } = await importFromDemo(
  "@arcships/vue-docx",
);

// ---------------------------------------------------------------------------
// Minimal fake DOM
//
// The raster helpers are browser code; these fakes model only the DOM surface
// they touch so the pipeline can run under `node --test`.
// ---------------------------------------------------------------------------

class FakeStyle {
  constructor(initial = {}) {
    this.values = new Map(Object.entries(initial));
  }

  setProperty(name, value) {
    this.values.set(String(name), String(value));
  }

  getPropertyValue(name) {
    return this.values.get(String(name)) ?? "";
  }

  get cssText() {
    return [...this.values]
      .map(([name, value]) => `${name}:${value}`)
      .join(";");
  }
}

class FakeClassList {
  constructor(names = []) {
    this.names = new Set(names);
  }

  contains(name) {
    return this.names.has(name);
  }

  add(...names) {
    for (const name of names) this.names.add(name);
  }

  remove(...names) {
    for (const name of names) this.names.delete(name);
  }
}

class FakeTextNode {
  constructor(text) {
    this.nodeType = 3;
    this.textContent = text;
    this.parentNode = null;
  }
}

class FakeElement {
  constructor(tagName, options = {}) {
    this.nodeType = 1;
    this.tagName = tagName;
    this.attributes = { ...(options.attributes ?? {}) };
    this.style = new FakeStyle(options.style);
    this.computedStyle = new FakeStyle(options.computedStyle);
    this.classList = new FakeClassList(options.classes ?? []);
    this.children = [];
    this.parentNode = null;
    for (const child of options.children ?? []) this.appendChild(child);
  }

  get childNodes() {
    return this.children;
  }

  appendChild(child) {
    child.parentNode = this;
    this.children.push(child);
    return child;
  }

  getAttribute(name) {
    return this.attributes[name] ?? null;
  }

  setAttribute(name, value) {
    this.attributes[name] = String(value);
  }

  matches(selector) {
    return matchesSelector(this, selector);
  }

  querySelectorAll(selector) {
    const matches = [];
    const visit = (node) => {
      if (node.nodeType !== 1) return;
      if (node !== this && matchesSelector(node, selector)) matches.push(node);
      for (const child of node.children) visit(child);
    };
    for (const child of this.children) visit(child);
    return matches;
  }

  remove() {
    if (!this.parentNode) return;
    const index = this.parentNode.children.indexOf(this);
    if (index >= 0) this.parentNode.children.splice(index, 1);
    this.parentNode = null;
  }

  cloneNode(deep) {
    const clone = new FakeElement(this.tagName, {
      attributes: this.attributes,
      classes: [...this.classList.names],
    });
    for (const [name, value] of this.style.values) {
      clone.style.setProperty(name, value);
    }
    if (deep) {
      for (const child of this.children) {
        clone.appendChild(
          child.nodeType === 3
            ? new FakeTextNode(child.textContent)
            : child.cloneNode(true),
        );
      }
    }
    return clone;
  }
}

function matchesSelector(element, selector) {
  return selector
    .split(",")
    .some((part) => matchesSingleSelector(element, part.trim()));
}

function matchesSingleSelector(element, selector) {
  const attributeMatch = selector.match(/^\[([\w-]+)(?:="([^"]*)")?\]$/);
  if (attributeMatch) {
    const [, name, expected] = attributeMatch;
    const actual = element.getAttribute(name);
    if (actual === null) return false;
    return expected === undefined || actual === expected;
  }
  return element.tagName === selector;
}

function escapeXmlText(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

class FakeXmlSerializer {
  serializeToString(node) {
    if (node.nodeType === 3) return escapeXmlText(node.textContent);
    const attributes = [];
    if (node.classList.names.size > 0) {
      attributes.push(`class="${[...node.classList.names].join(" ")}"`);
    }
    for (const [name, value] of Object.entries(node.attributes)) {
      attributes.push(`${name}="${value}"`);
    }
    if (node.style.values.size > 0) {
      attributes.push(`style="${node.style.cssText}"`);
    }
    const open = `<${node.tagName}${attributes
      .map((attribute) => ` ${attribute}`)
      .join("")}>`;
    return `${open}${node.children
      .map((child) => this.serializeToString(child))
      .join("")}</${node.tagName}>`;
  }
}

function restoreGlobal(name, value) {
  if (value === undefined) delete globalThis[name];
  else globalThis[name] = value;
}

/**
 * Installs `document`, `Image`, `XMLSerializer`, `window` and
 * `getComputedStyle` fakes. `imageBehavior` controls how fake images settle:
 * `resolve` (decode succeeds), `error` (the error event fires) or `pending`
 * (the decoder never settles, exercising the timeout).
 */
function installRasterEnvironment(t, imageBehavior = "resolve") {
  const originals = {
    document: globalThis.document,
    window: globalThis.window,
    Image: globalThis.Image,
    XMLSerializer: globalThis.XMLSerializer,
    getComputedStyle: globalThis.getComputedStyle,
  };
  const operations = [];

  class FakeImage {
    static instances = [];

    constructor() {
      this.decoding = "";
      this.naturalWidth = 40;
      this.naturalHeight = 30;
      this.width = 40;
      this.height = 30;
      this.onload = null;
      this.onerror = null;
      this._src = "";
      FakeImage.instances.push(this);
    }

    get src() {
      return this._src;
    }

    set src(value) {
      this._src = value;
      if (imageBehavior === "error") {
        queueMicrotask(() => this.onerror?.(new Error("fake decode error")));
      } else if (imageBehavior === "resolve") {
        queueMicrotask(() => this.onload?.(new Event("load")));
      }
    }

    decode() {
      if (imageBehavior === "resolve") return Promise.resolve();
      if (imageBehavior === "error") {
        return Promise.reject(new Error("fake decode error"));
      }
      return new Promise(() => {});
    }
  }

  const context = {
    fillStyle: "#000000",
    strokeStyle: "#000000",
    lineWidth: 1,
    font: "",
    textBaseline: "alphabetic",
    imageSmoothingEnabled: false,
    imageSmoothingQuality: "low",
    setTransform(...args) {
      operations.push({ name: "setTransform", args });
    },
    clearRect(...args) {
      operations.push({ name: "clearRect", args });
    },
    fillRect(...args) {
      operations.push({ name: "fillRect", args });
    },
    strokeRect(...args) {
      operations.push({ name: "strokeRect", args });
    },
    drawImage(...args) {
      operations.push({ name: "drawImage", args });
    },
    save() {
      operations.push({ name: "save", args: [] });
    },
    restore() {
      operations.push({ name: "restore", args: [] });
    },
    beginPath() {
      operations.push({ name: "beginPath", args: [] });
    },
    rect(...args) {
      operations.push({ name: "rect", args });
    },
    clip() {
      operations.push({ name: "clip", args: [] });
    },
    fillText(...args) {
      operations.push({ name: "fillText", args });
    },
    measureText(value) {
      return { width: String(value).length * 6 };
    },
  };

  globalThis.window = { devicePixelRatio: 1 };
  globalThis.Image = FakeImage;
  globalThis.XMLSerializer = FakeXmlSerializer;
  globalThis.getComputedStyle = (element) =>
    element.computedStyle ?? new FakeStyle();
  globalThis.document = {
    createElement(tagName) {
      assert.equal(tagName, "canvas");
      return {
        width: 0,
        height: 0,
        style: {},
        getContext(kind) {
          assert.equal(kind, "2d");
          return context;
        },
      };
    },
  };

  t.after(() => {
    restoreGlobal("document", originals.document);
    restoreGlobal("window", originals.window);
    restoreGlobal("Image", originals.Image);
    restoreGlobal("XMLSerializer", originals.XMLSerializer);
    restoreGlobal("getComputedStyle", originals.getComputedStyle);
  });

  return { FakeImage, operations };
}

function classNames(node) {
  return String(node?.props?.class ?? "")
    .split(/\s+/)
    .filter(Boolean);
}

function paragraph(text) {
  return {
    type: "paragraph",
    children: [{ type: "text", text }],
  };
}

const DOCUMENT_COMPUTED_STYLE = {
  "font-family": "Calibri, sans-serif",
  "font-size": "14.6667px",
  "line-height": "18px",
  "white-space": "pre-wrap",
  color: "rgb(17, 24, 39)",
  "text-align": "start",
};

function fakeTableCell(classes, text, background) {
  return new FakeElement("td", {
    classes,
    computedStyle: {
      "background-color": background,
      "border-top-width": "1px",
      "border-top-style": "solid",
      "border-top-color": "rgb(209, 213, 219)",
      "border-right-width": "1px",
      "border-right-style": "solid",
      "border-right-color": "rgb(209, 213, 219)",
      "border-bottom-width": "1px",
      "border-bottom-style": "solid",
      "border-bottom-color": "rgb(209, 213, 219)",
      "border-left-width": "1px",
      "border-left-style": "solid",
      "border-left-color": "rgb(209, 213, 219)",
      "padding-top": "8px",
      "padding-right": "8px",
      "padding-bottom": "8px",
      "padding-left": "8px",
      "vertical-align": "top",
      ...DOCUMENT_COMPUTED_STYLE,
    },
    children: [new FakeTextNode(text)],
  });
}

function fakePageSurface() {
  const activeParagraph = new FakeElement("div", {
    classes: ["docx-paragraph-host", "docx-paragraph-host--active"],
    computedStyle: {
      ...DOCUMENT_COMPUTED_STYLE,
      display: "block",
      "background-color": "rgba(59, 130, 246, 0.04)",
    },
    children: [
      new FakeTextNode("Active paragraph"),
      new FakeElement("span", {
        computedStyle: {
          ...DOCUMENT_COMPUTED_STYLE,
          "background-color": "rgba(0, 0, 0, 0)",
        },
        children: [new FakeTextNode(" run")],
      }),
    ],
  });
  const table = new FakeElement("table", {
    computedStyle: {
      display: "table",
      "border-collapse": "collapse",
      "border-spacing": "0px",
      "table-layout": "fixed",
      ...DOCUMENT_COMPUTED_STYLE,
    },
    children: [
      fakeTableCell(
        ["docx-table-cell", "docx-table-cell--active"],
        "Active cell",
        "rgba(59, 130, 246, 0.04)",
      ),
      fakeTableCell(
        ["docx-table-cell"],
        "Shaded cell",
        "rgb(229, 231, 235)",
      ),
    ],
  });
  const addRowButton = new FakeElement("div", {
    classes: ["docx-table-add-row"],
    attributes: { [DOCX_THUMBNAIL_EXCLUDE_ATTRIBUTE]: "true" },
    children: [new FakeTextNode("+ Add row")],
  });
  const resizeHandle = new FakeElement("div", {
    classes: ["docx-image-resize-handle"],
    attributes: { "data-image-resize-handle": "true" },
  });

  return new FakeElement("div", {
    attributes: { "data-docx-page-surface": "true" },
    computedStyle: {
      display: "flex",
      "flex-direction": "column",
      position: "relative",
      "box-sizing": "border-box",
      "background-color": "rgb(255, 255, 255)",
      "font-family": "Arial, sans-serif",
      "font-size": "16px",
      "line-height": "normal",
      color: "rgb(17, 24, 39)",
    },
    children: [
      activeParagraph,
      table,
      addRowButton,
      resizeHandle,
    ],
  });
}

test("DOCX thumbnail raster inlines computed page styles and strips editor chrome", async (t) => {
  const { FakeImage } = installRasterEnvironment(t);
  const page = fakePageSurface();

  const surface = await rasterizeDocxThumbnailSurface({
    pageElement: page,
    sourceWidthPx: 816,
    sourceHeightPx: 1056,
    widthPx: 140,
    heightPx: 181,
    pixelWidthPx: 280,
    pixelHeightPx: 362,
  });

  assert.equal(surface.width, 280);
  assert.equal(surface.height, 362);
  const svgImage = FakeImage.instances.find((image) =>
    image.src.startsWith("data:image/svg+xml"),
  );
  assert.ok(svgImage, "the raster pipeline must load the serialized SVG");
  const serialized = decodeURIComponent(
    svgImage.src.slice(svgImage.src.indexOf(",") + 1),
  );

  // Scoped CSS values now travel as inline styles.
  assert.match(serialized, /line-height:18px/);
  assert.match(serialized, /font-family:Calibri, sans-serif/);
  assert.match(serialized, /white-space:pre-wrap/);
  assert.match(serialized, /flex-direction:column/);
  assert.match(serialized, /border-collapse:collapse/);
  assert.match(serialized, /table-layout:fixed/);
  assert.match(serialized, /padding-top:8px/);
  assert.match(serialized, /border-top-color:rgb\(209, 213, 219\)/);
  assert.match(serialized, /background-color:rgb\(229, 231, 235\)/);

  // Inherited values are written once per value change (root -> paragraph and
  // root -> table), not repeated on every child run/cell.
  assert.equal(
    (serialized.match(/font-size:14\.6667px/g) ?? []).length,
    2,
    "inherited computed font-size should not be repeated on child runs",
  );

  // Editor chrome is gone from the clone.
  assert.doesNotMatch(serialized, /\+ Add row/);
  assert.doesNotMatch(serialized, /docx-image-resize-handle/);
  assert.doesNotMatch(serialized, /docx-paragraph-host--active/);
  assert.doesNotMatch(serialized, /docx-table-cell--active/);
  assert.doesNotMatch(serialized, /rgba\(59, 130, 246, 0\.04\)/);
  assert.match(serialized, /Active paragraph/);
  assert.match(serialized, /Shaded cell/);

  // The live page is not mutated by the raster pass.
  assert.equal(
    page.querySelectorAll('[data-image-resize-handle="true"]').length,
    1,
  );
  assert.equal(
    page.querySelectorAll(`[${DOCX_THUMBNAIL_EXCLUDE_ATTRIBUTE}="true"]`).length,
    1,
  );
});

test("DOCX thumbnail image decode falls back to a placeholder after the deadline", async (t) => {
  const { operations } = installRasterEnvironment(t, "pending");
  t.mock.timers.enable({ apis: ["setTimeout"] });

  const pendingSurface = renderDocxThumbnailSnapshotSurface({
    snapshot: {
      key: "decode-timeout",
      sourceWidthPx: 816,
      sourceHeightPx: 1056,
      pageBackgroundColor: "#ffffff",
      elements: [
        {
          kind: "image",
          xPx: 10,
          yPx: 20,
          widthPx: 100,
          heightPx: 50,
          src: "data:image/png;base64,stalled",
        },
      ],
    },
    widthPx: 140,
    heightPx: 181,
    pixelWidthPx: 280,
    pixelHeightPx: 362,
  });

  // Nothing but the decode deadline can settle this render.
  await Promise.resolve();
  assert.equal(operations.some((operation) => operation.name === "drawImage"), false);
  t.mock.timers.tick(10_000);

  const surface = await pendingSurface;
  assert.ok(surface);
  assert.ok(
    operations.some(
      (operation) =>
        operation.name === "fillRect" &&
        operation.args[0] === 10 &&
        operation.args[1] === 20 &&
        operation.args[2] === 100 &&
        operation.args[3] === 50,
    ),
    "a timed-out image must fall back to the placeholder box",
  );
  assert.ok(operations.some((operation) => operation.name === "strokeRect"));
});

test("DOCX thumbnail downscale decode timeout resolves as no replacement", async (t) => {
  installRasterEnvironment(t, "pending");
  const started = Date.now();
  const source = `data:image/png;base64,${"A".repeat(40_000)}`;

  const result = await getDownscaledThumbnailImageDataUri(source, {
    decodeTimeoutMs: 25,
  });

  assert.equal(result, undefined);
  assert.ok(
    Date.now() - started < 1_000,
    "a stalled image decoder must not pin the downscale request",
  );
});

test("DOCX table chrome advertises itself for thumbnail exclusion", async () => {
  const table = {
    type: "table",
    rows: [
      {
        cells: [
          {
            nodes: [paragraph("Cell content")],
          },
        ],
      },
    ],
  };
  const model = {
    ...createBlankDocumentModel(),
    nodes: [table],
  };
  const controller = {
    model,
    selection: {
      kind: "table-cell",
      tableIndex: 0,
      rowIndex: 0,
      cellIndex: 0,
    },
    showTrackedChanges: false,
    registerTextDraftFlusher() {
      return () => {};
    },
  };
  const Harness = vue.defineComponent({
    setup() {
      return () =>
        vue.h(DocxTableHost, {
          table,
          tableIndex: 0,
          editable: true,
          controller,
        });
    },
  });

  const mounted = await mount(Harness);
  const handles = walk(mounted.root).find((node) =>
    classNames(node).includes("docx-table-column-handles"),
  );
  const addRowButton = walk(mounted.root).find((node) =>
    classNames(node).includes("docx-table-add-row"),
  );

  assert.ok(handles, "active table must render column handles");
  assert.ok(addRowButton, "editable table must render the add-row button");
  assert.equal(handles.props[DOCX_THUMBNAIL_EXCLUDE_ATTRIBUTE], "true");
  assert.equal(addRowButton.props[DOCX_THUMBNAIL_EXCLUDE_ATTRIBUTE], "true");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DOCX floating-image chrome advertises itself for thumbnail exclusion", async () => {
  const imageRun = {
    type: "image",
    src: "data:image/png;base64,AAAA",
    alt: "floating picture",
    widthPx: 120,
    heightPx: 80,
    floating: { xPx: 10, yPx: 20 },
  };
  const model = {
    ...createBlankDocumentModel(),
    nodes: [
      {
        type: "paragraph",
        children: [imageRun],
      },
    ],
  };
  const controller = {
    model,
    selection: { kind: "paragraph", nodeIndex: 0 },
  };
  const Harness = vue.defineComponent({
    setup() {
      return () =>
        vue.h(DocxImageLayer, {
          pageIndex: 0,
          pageWidthPx: 816,
          pageHeightPx: 1056,
          controller,
          pageNodeSegments: [{ nodeIndex: 0 }],
          editable: true,
        });
    },
  });

  const mounted = await mount(Harness);
  const toolbar = walk(mounted.root).find((node) =>
    classNames(node).includes("docx-image-wrap-toolbar"),
  );
  const handles = walk(mounted.root).filter(
    (node) => node.props?.["data-image-resize-handle"] === "true",
  );

  assert.ok(toolbar, "selected floating image must render its wrap toolbar");
  assert.equal(toolbar.props[DOCX_THUMBNAIL_EXCLUDE_ATTRIBUTE], "true");
  assert.equal(handles.length, 4, "all resize handles must be marked");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DOCX thumbnail raster keeps SVG picture sources vector", async (t) => {
  const { FakeImage } = installRasterEnvironment(t);
  const svgSource = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(
    '<svg xmlns="http://www.w3.org/2000/svg">' +
      "<rect/>".repeat(6_000) +
      "</svg>",
  )}`;
  assert.ok(
    svgSource.length > 32_768,
    "the fixture source must be large enough to qualify for raster downscaling",
  );

  const page = new FakeElement("div", {
    attributes: { "data-docx-page-surface": "true" },
    computedStyle: {
      ...DOCUMENT_COMPUTED_STYLE,
      display: "block",
      position: "relative",
      "box-sizing": "border-box",
      "background-color": "rgb(255, 255, 255)",
    },
    children: [
      new FakeElement("img", {
        attributes: { src: svgSource, alt: "svg picture" },
        computedStyle: { ...DOCUMENT_COMPUTED_STYLE, display: "block" },
      }),
    ],
  });

  const surface = await rasterizeDocxThumbnailSurface({
    pageElement: page,
    sourceWidthPx: 816,
    sourceHeightPx: 1056,
    widthPx: 140,
    heightPx: 181,
    pixelWidthPx: 280,
    pixelHeightPx: 362,
  });
  assert.equal(surface.width, 280);

  // The only image decode is the thumbnail's own outer SVG. The picture is
  // not re-decoded to a downscaled PNG, so the thumbnail keeps it vector until
  // the final surface rasterisation.
  assert.equal(FakeImage.instances.length, 1);
  const outer = FakeImage.instances[0];
  assert.match(outer.src, /^data:image\/svg\+xml/u);
  const serialized = decodeURIComponent(outer.src.slice(outer.src.indexOf(",") + 1));
  assert.match(serialized, /<img[^>]+src="data:image\/svg\+xml;charset=utf-8,/u);
  assert.doesNotMatch(serialized, /data:image\/png/u);
});
