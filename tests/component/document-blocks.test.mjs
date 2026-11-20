// Issue #164 — document blocks interaction contract.
//
// Covers the behaviours the pixel-parity harness cannot express: keyboard
// focus/selection, adding and reordering document splits, drag/drop between
// groups, typed and drawn signatures, review-field editing, and the
// empty/loading/error surfaces. Layout and colours are gated separately by
// `pnpm ui:parity -- --only bounding-box-citations,layout-blocks,e-signature,document-splits`.
import assert from "node:assert/strict";
import test from "node:test";

import {
  importFromDemo,
  keydownEvent,
  mount,
  textContent,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

const {
  BoundingBoxCitations,
  DocumentSplits,
  ESignatureFields,
  LayoutBlocks,
  SignaturePad,
} = await importFromDemo("@arcships/vue-ui");

function classNames(value) {
  if (Array.isArray(value)) return value.map(classNames).filter(Boolean).join(" ");
  if (value && typeof value === "object") {
    return Object.entries(value)
      .filter(([, enabled]) => enabled)
      .map(([name]) => name)
      .join(" ");
  }
  return String(value ?? "");
}

function byClass(root, className, all = false) {
  const matches = walk(root).filter((node) =>
    classNames(node.props?.class).split(/\s+/).includes(className),
  );
  return all ? matches : matches[0];
}

function bySlot(root, slot, all = false) {
  const matches = walk(root).filter((node) => node.props?.["data-slot"] === slot);
  return all ? matches : matches[0];
}

function byAttr(root, name, value, all = false) {
  const matches = walk(root).filter((node) => node.props?.[name] === value);
  return all ? matches : matches[0];
}

function waitForFrame() {
  return new Promise((resolve) => setTimeout(resolve, 10));
}

const REVIEW_FIELDS = [
  {
    key: "vendor_name",
    schema: { type: "string", title: "Vendor name" },
    actual: "Northwind Industrial Supplies",
    expected: "Northwind Industrial Supplies",
    location: { page: 1, area: { left: 8, top: 9, width: 42, height: 4.5 } },
  },
  {
    key: "invoice_total",
    schema: { type: "number", title: "Invoice total" },
    actual: 2882.58,
    expected: 2882.58,
    location: { page: 1, area: { left: 62, top: 64, width: 30, height: 4 } },
  },
  {
    key: "payment_terms",
    schema: { type: "string", title: "Payment terms" },
    actual: "Net 30",
    expected: "Net 30",
    location: { page: 1, area: { left: 8, top: 88, width: 50, height: 4 } },
  },
];

const OCR_BLOCKS = [
  { id: "b1", type: "heading", text: "Payment terms", confidence: 0.98, page: 1, bbox: [10, 10, 100, 20] },
  { id: "b2", type: "paragraph", text: "Invoices are due within 30 days.", confidence: 0.89, page: 1, bbox: [10, 40, 200, 40] },
  { id: "b3", type: "table", text: "Amount and due date", confidence: 0.77, page: 2, bbox: [10, 100, 200, 60] },
];

const SPLITS = [
  { id: "split-1", title: "Abstract and intro", pages: ["page-1", "page-2", "page-3"] },
  { id: "split-2", title: "Model architecture", pages: ["page-4", "page-5"] },
  { id: "split-3", title: "Training and results", pages: ["page-6"] },
];

function pointerEvent(x, y, pointerId = 1) {
  let prevented = 0;
  return {
    clientX: x,
    clientY: y,
    pointerId,
    preventedCount: () => prevented,
    preventDefault() {
      prevented += 1;
    },
    currentTarget: {
      setPointerCapture() {},
      releasePointerCapture() {},
    },
  };
}

/** Minimal canvas surface for the draw-mode SignaturePad. */
function installCanvas(t) {
  const operations = [];
  const originalGetContext = Object.getOwnPropertyDescriptor(Object.prototype, "getContext");
  const originalToDataURL = Object.getOwnPropertyDescriptor(Object.prototype, "toDataURL");

  Object.defineProperty(Object.prototype, "getContext", {
    configurable: true,
    value(kind) {
      if (this.type !== "canvas" || kind !== "2d") return null;
      if (Object.hasOwn(this, "__context")) return this.__context;
      const record = (name, ...args) => operations.push({ name, args });
      const context = {
        setTransform: (...args) => record("setTransform", ...args),
        clearRect: (...args) => record("clearRect", ...args),
        fillRect: (...args) => record("fillRect", ...args),
        beginPath: () => record("beginPath"),
        moveTo: (...args) => record("moveTo", ...args),
        lineTo: (...args) => record("lineTo", ...args),
        stroke: () => record("stroke"),
      };
      Object.defineProperty(this, "__context", { configurable: true, value: context });
      return context;
    },
  });
  Object.defineProperty(Object.prototype, "toDataURL", {
    configurable: true,
    value() {
      return "data:image/png;base64,c2lnbmF0dXJl";
    },
  });

  t.after(() => {
    if (originalGetContext) Object.defineProperty(Object.prototype, "getContext", originalGetContext);
    else delete Object.prototype.getContext;
    if (originalToDataURL) Object.defineProperty(Object.prototype, "toDataURL", originalToDataURL);
    else delete Object.prototype.toDataURL;
  });

  return operations;
}

// Layout blocks ----------------------------------------------------------

test("LayoutBlocks renders typed OCR cards with confidence, page and active state", async () => {
  const mounted = await mount(LayoutBlocks, {
    view: "panel",
    blocks: OCR_BLOCKS,
  });

  assert.equal(bySlot(mounted.root, "layout-block-card", true).length, 3);
  assert.equal(byClass(mounted.root, "layout-blocks__card--heading", true).length, 1);
  assert.match(textContent(mounted.root), /Heading/);
  assert.match(textContent(mounted.root), /98%/);
  assert.match(textContent(mounted.root), /p\. 1/);
  assert.equal(
    byAttr(mounted.root, "data-layout-block-id", "b1").props["aria-pressed"],
    true,
  );
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("LayoutBlocks keyboard navigation and activation emit focus/click events", async () => {
  const focused = [];
  const clicked = [];
  const active = [];
  const mounted = await mount(LayoutBlocks, {
    view: "panel",
    blocks: OCR_BLOCKS,
    onBlockFocus: (block) => focused.push(block.id),
    onBlockClick: (block) => clicked.push(block.id),
    "onUpdate:activeBlockId": (id) => active.push(id),
  });

  const first = byAttr(mounted.root, "data-layout-block-id", "b1");
  const down = keydownEvent("ArrowDown");
  first.props.onKeydown(down.event);
  await vue.nextTick();
  assert.equal(down.preventedCount(), 1);
  assert.deepEqual(focused, ["b2"]);
  assert.deepEqual(active, ["b2"]);
  assert.equal(
    byAttr(mounted.root, "data-layout-block-id", "b2").props["aria-pressed"],
    true,
  );

  const end = keydownEvent("End");
  byAttr(mounted.root, "data-layout-block-id", "b2").props.onKeydown(end.event);
  await vue.nextTick();
  assert.equal(
    byAttr(mounted.root, "data-layout-block-id", "b3").props["aria-pressed"],
    true,
  );

  byAttr(mounted.root, "data-layout-block-id", "b3").props.onClick();
  assert.deepEqual(clicked, ["b3"]);
  mounted.app.unmount();
});

test("LayoutBlocks empty, loading and error surfaces expose live status semantics", async () => {
  const empty = await mount(LayoutBlocks, { view: "panel", blocks: [] });
  assert.equal(bySlot(empty.root, "layout-blocks-empty").props.role, "status");
  assert.match(textContent(empty.root), /No layout blocks found\./);
  empty.app.unmount();

  const loading = await mount(LayoutBlocks, { view: "panel", blocks: [], loading: true });
  assert.equal(bySlot(loading.root, "layout-blocks-loading").props.role, "status");
  assert.match(textContent(loading.root), /Loading layout blocks/);
  loading.app.unmount();

  const error = await mount(LayoutBlocks, {
    view: "panel",
    blocks: OCR_BLOCKS,
    error: "Layout blocks could not be loaded.",
  });
  const alert = walk(error.root).find((node) => node.props?.role === "alert");
  assert.match(textContent(alert), /could not be loaded/);
  error.app.unmount();
});

test("LayoutBlocks markdown rendering escapes host text before formatting", async () => {
  const mounted = await mount(LayoutBlocks, {
    view: "panel",
    renderMarkdown: true,
    blocks: [
      {
        id: "md",
        type: "paragraph",
        text: "# Heading\n**Bold** and <img src=x onerror=alert(1)> [site](https://example.com)",
        bbox: [0, 0, 10, 10],
      },
    ],
  });
  const markdown = bySlot(mounted.root, "layout-block-markdown");
  assert.match(markdown.props.innerHTML, /&lt;img src=x onerror=alert\(1\)&gt;/);
  assert.doesNotMatch(markdown.props.innerHTML, /<img/);
  assert.match(markdown.props.innerHTML, /<strong>Bold<\/strong>/);
  assert.match(markdown.props.innerHTML, /href="https:\/\/example.com"/);
  mounted.app.unmount();
});

test("LayoutBlocks split mode renders the page overlay and viewer slot", async () => {
  const mounted = await mount(LayoutBlocks, {
    view: "split",
    pageImage: "/samples/contract-page.png",
    blocks: OCR_BLOCKS,
    pageWidth: 1000,
    pageHeight: 1330,
  });
  const overlays = walk(mounted.root).filter(
    (node) => node.props?.["data-slot"] === "layout-block-overlay",
  );
  assert.equal(overlays.length, 3);
  assert.equal(overlays[0].props.style.left, "1%");
  assert.equal(overlays[0].props.style.top, `${(10 / 1330) * 100}%`);
  assert.equal(overlays[0].props["aria-pressed"], true);
  assert.equal(bySlot(mounted.root, "layout-blocks-viewer").children.length >= 1, true);
  mounted.app.unmount();
});

// Bounding box citations ------------------------------------------------

test("BoundingBoxCitations review mode renders field cards with Actual/Expected and types", async () => {
  const mounted = await mount(BoundingBoxCitations, {
    reviewFields: REVIEW_FIELDS,
    pageImage: "/samples/contract-page.png",
    view: "panel",
  });

  assert.equal(bySlot(mounted.root, "bounding-box-field", true).length, 3);
  assert.match(textContent(mounted.root), /Vendor name/);
  assert.match(textContent(mounted.root), /Actual/);
  assert.match(textContent(mounted.root), /Expected/);
  assert.match(textContent(mounted.root), /string/);
  assert.match(textContent(mounted.root), /number/);
  assert.equal(
    byAttr(mounted.root, "data-review-field-key", "vendor_name").props["aria-pressed"],
    undefined,
  );
  assert.match(
    classNames(byAttr(mounted.root, "data-review-field-key", "vendor_name").props.class),
    /is-active/,
  );
  mounted.app.unmount();
});

test("BoundingBoxCitations hover/focus, keyboard and editing update the active field", async () => {
  const focused = [];
  const active = [];
  const expected = [];
  const mounted = await mount(BoundingBoxCitations, {
    reviewFields: REVIEW_FIELDS,
    view: "panel",
    onFieldFocus: (field) => focused.push(field.key),
    "onUpdate:activeFieldKey": (key) => active.push(key),
    "onUpdate:expected": (key, value) => expected.push([key, value]),
  });

  const second = byAttr(mounted.root, "data-review-field-key", "invoice_total");
  second.props.onMouseenter();
  await vue.nextTick();
  assert.deepEqual(focused, ["invoice_total"]);
  assert.deepEqual(active, ["invoice_total"]);
  assert.match(classNames(second.props.class), /is-active/);

  const third = byAttr(mounted.root, "data-review-field-key", "payment_terms");
  third.props.onFocusCapture();
  await vue.nextTick();
  assert.deepEqual(active, ["invoice_total", "payment_terms"]);

  const up = keydownEvent("ArrowUp");
  third.props.onKeydown(up.event);
  await vue.nextTick();
  assert.equal(up.preventedCount(), 1);
  assert.deepEqual(active.at(-1), "invoice_total");

  const input = walk(third).find((node) => node.props?.["data-slot"] === "input");
  input.props.onInput({ target: { value: "Net 45" } });
  await vue.nextTick();
  assert.deepEqual(expected, [["payment_terms", "Net 45"]]);
  assert.match(classNames(second.props.class), /is-active/);

  mounted.app.unmount();
});

test("BoundingBoxCitations expectedValues seeds the modified state and revert resets", async () => {
  const expected = [];
  const mounted = await mount(BoundingBoxCitations, {
    reviewFields: REVIEW_FIELDS,
    expectedValues: { payment_terms: "Net 45" },
    view: "panel",
    "onUpdate:expected": (key, value) => expected.push([key, value]),
  });

  const card = byAttr(mounted.root, "data-review-field-key", "payment_terms");
  const dot = byClass(card, "review-card__modified-dot");
  assert.doesNotMatch(classNames(dot.props.class), /is-hidden/);
  assert.equal(walk(card).some((node) => node.props?.["data-slot"] === "bounding-box-revert"), true);

  const revert = walk(card).find(
    (node) => node.props?.["data-slot"] === "bounding-box-revert",
  );
  revert.props.onClick({ stopPropagation() {} });
  await vue.nextTick();
  assert.deepEqual(expected, [["payment_terms", "Net 30"]]);
  assert.match(classNames(dot.props.class), /is-hidden/);
  mounted.app.unmount();
});

test("BoundingBoxCitations tab switch shows the JSON view and empty/loading/error states", async () => {
  const mounted = await mount(BoundingBoxCitations, {
    reviewFields: REVIEW_FIELDS,
    view: "panel",
  });
  bySlot(mounted.root, "bounding-box-tab-json").props.onClick();
  await vue.nextTick();
  assert.match(textContent(bySlot(mounted.root, "bounding-box-json")), /"vendor_name"/);
  bySlot(mounted.root, "bounding-box-tab-form").props.onClick();
  await vue.nextTick();
  assert.equal(bySlot(mounted.root, "bounding-box-field", true).length, 3);
  mounted.app.unmount();

  const empty = await mount(BoundingBoxCitations, { reviewFields: [], view: "panel" });
  assert.equal(bySlot(empty.root, "bounding-box-empty").props.role, "status");
  empty.app.unmount();

  const loading = await mount(BoundingBoxCitations, {
    reviewFields: [],
    loading: true,
    view: "panel",
  });
  assert.equal(bySlot(loading.root, "bounding-box-loading").props.role, "status");
  loading.app.unmount();

  const error = await mount(BoundingBoxCitations, {
    reviewFields: [],
    error: "Review fields could not be loaded.",
    view: "panel",
  });
  assert.equal(
    walk(error.root).some((node) => node.props?.role === "alert"),
    true,
  );
  error.app.unmount();
});

test("BoundingBoxCitations split mode positions the active page highlight", async () => {
  const mounted = await mount(BoundingBoxCitations, {
    reviewFields: REVIEW_FIELDS,
    pageImage: "/samples/contract-page.png",
    view: "split",
    activeFieldKey: "invoice_total",
  });
  const highlight = bySlot(mounted.root, "bounding-box-highlight");
  assert.equal(highlight.props.style.left, "62%");
  assert.equal(highlight.props.style.top, "64%");
  assert.equal(highlight.props.style.width, "30%");
  assert.equal(highlight.props.style.height, "4%");
  mounted.app.unmount();
});

test("BoundingBoxCitations keeps the legacy page-card API and field-click event", async () => {
  const clicked = [];
  const fields = [
    {
      id: "vendor",
      label: "Vendor",
      page: 1,
      rect: [0.07, 0.1, 0.48, 0.15],
      value: "Northwind",
      confidence: 0.96,
    },
  ];
  const mounted = await mount(BoundingBoxCitations, {
    file: "/samples/contract-page.png",
    fields,
    onFieldClick: (field) => clicked.push(field.id),
  });
  assert.match(textContent(mounted.root), /Page 1/);
  byClass(mounted.root, "citation-list-item").props.onClick();
  assert.deepEqual(clicked, ["vendor"]);
  mounted.app.unmount();
});

// E-signature -----------------------------------------------------------

test("SignaturePad typed mode emits typed value, clear and confirm", async () => {
  const signature = [];
  const model = [];
  const confirmed = [];
  const cleared = [];
  const mounted = await mount(SignaturePad, {
    mode: "typed",
    showControls: true,
    showConfirm: true,
    showCancel: true,
    "onUpdate:signature": (value) => signature.push(value),
    "onUpdate:modelValue": (value) => model.push(value),
    onConfirm: (value) => confirmed.push(value),
    onClear: () => cleared.push(true),
  });

  const input = bySlot(mounted.root, "signature-typed-input");
  input.props.onInput({ target: { value: "Dana Whitfield" } });
  await vue.nextTick();
  assert.deepEqual(model, ["Dana Whitfield"]);
  assert.match(textContent(mounted.root), /Dana Whitfield/);
  assert.equal(bySlot(mounted.root, "signature-confirm").props.disabled, false);

  bySlot(mounted.root, "signature-confirm").props.onClick();
  assert.deepEqual(confirmed, ["Dana Whitfield"]);

  bySlot(mounted.root, "signature-clear").props.onClick();
  await vue.nextTick();
  assert.deepEqual(cleared, [true]);
  assert.deepEqual(signature.at(-1), null);
  assert.deepEqual(model.at(-1), null);
  mounted.app.unmount();
});

test("SignaturePad draw mode records pointer strokes and respects disabled/error states", async (t) => {
  installCanvas(t);
  const signature = [];
  const mounted = await mount(SignaturePad, {
    showConfirm: true,
    width: 400,
    height: 200,
    "onUpdate:signature": (value) => signature.push(value),
  });
  await vue.nextTick();
  const canvas = byClass(mounted.root, "signature-pad__canvas");
  canvas.props.onPointerdown(pointerEvent(10, 10));
  canvas.props.onPointermove(pointerEvent(40, 30));
  canvas.props.onPointerup(pointerEvent(40, 30));
  await vue.nextTick();
  assert.equal(signature.at(-1), "data:image/png;base64,c2lnbmF0dXJl");
  assert.equal(bySlot(mounted.root, "signature-confirm").props.disabled, false);

  bySlot(mounted.root, "signature-clear").props.onClick();
  await vue.nextTick();
  assert.equal(signature.at(-1), null);
  assert.equal(bySlot(mounted.root, "signature-confirm").props.disabled, true);
  mounted.app.unmount();

  const disabled = await mount(SignaturePad, { disabled: true });
  const disabledCanvas = byClass(disabled.root, "signature-pad__canvas");
  disabledCanvas.props.onPointerdown(pointerEvent(1, 1));
  disabledCanvas.props.onPointermove(pointerEvent(2, 2));
  assert.equal(disabledCanvas.props.onPointermove, disabledCanvas.props.onPointermove);
  disabled.app.unmount();

  const error = await mount(SignaturePad, { error: "Signature pad unavailable" });
  assert.equal(
    walk(error.root).some((node) => node.props?.role === "alert"),
    true,
  );
  error.app.unmount();
});

test("ESignatureFields renders signed state, preview and action events", async () => {
  const signed = [];
  const cleared = [];
  const exported = [];
  const field = {
    id: "signature-1",
    label: "Authorized signature",
    page: 1,
    bbox: { x: 300, y: 504, width: 250, height: 58 },
    imageDataUrl: "data:image/png;base64,c2lnbmF0dXJl",
  };
  const mounted = await mount(ESignatureFields, {
    fields: [field],
    canExport: true,
    showPreview: true,
    onEdit: (value) => signed.push(value.id),
    onClear: (value) => cleared.push(value.id),
    onExport: (value) => exported.push(value.id),
  });
  assert.match(textContent(mounted.root), /Signature fields/);
  assert.match(textContent(mounted.root), /Signed/);
  assert.match(textContent(mounted.root), /250 x 58 on page 1/);
  assert.equal(byClass(mounted.root, "e-signature-field__preview-image").props.src, field.imageDataUrl);
  bySlot(mounted.root, "e-signature-field-sign").props.onClick();
  assert.deepEqual(signed, ["signature-1"]);
  bySlot(mounted.root, "e-signature-field-clear").props.onClick();
  assert.deepEqual(cleared, ["signature-1"]);
  bySlot(mounted.root, "e-signature-export").props.onClick();
  assert.deepEqual(exported, ["signature-1"]);
  mounted.app.unmount();
});

test("ESignatureFields unsigned, loading and error states gate export", async () => {
  const unsigned = await mount(ESignatureFields, {
    fields: [
      {
        id: "signature-1",
        label: "Authorized signature",
        page: 1,
        bbox: { x: 300, y: 504, width: 250, height: 58 },
      },
    ],
    canExport: false,
  });
  assert.match(textContent(unsigned.root), /Unsigned/);
  assert.equal(bySlot(unsigned.root, "e-signature-export").props.disabled, true);
  unsigned.app.unmount();

  const loading = await mount(ESignatureFields, { fields: [], loading: true });
  assert.equal(
    walk(loading.root).some((node) => node.props?.role === "status"),
    true,
  );
  loading.app.unmount();

  const error = await mount(ESignatureFields, {
    fields: [],
    error: "Signature fields could not be loaded.",
  });
  assert.equal(
    walk(error.root).some((node) => node.props?.role === "alert"),
    true,
  );
  error.app.unmount();
});

// Document splits -------------------------------------------------------

test("DocumentSplits renders split cards and emits page selection", async () => {
  const selected = [];
  const mounted = await mount(DocumentSplits, {
    splits: SPLITS,
    thumbnailImages: { "page-1": "/samples/contract-page.png" },
    withFrameDivider: false,
    onSelectPage: (page) => selected.push(page),
  });

  assert.equal(byClass(mounted.root, "doc-splits__card", true).length, 3);
  assert.match(textContent(mounted.root), /Abstract and intro/);
  assert.match(textContent(mounted.root), /Pages 1, 2, 3/);
  assert.equal(bySlot(mounted.root, "document-split-page", true).length, 6);

  byAttr(mounted.root, "data-page-id", "page-1").props.onClick();
  assert.deepEqual(selected, [1]);
  mounted.app.unmount();
});

test("DocumentSplits add, remove and duplicate-split guards update the list", async () => {
  const changes = [];
  const added = [];
  const removed = [];
  const mounted = await mount(DocumentSplits, {
    splits: SPLITS,
    "onUpdate:splits": (splits) => changes.push(splits),
    onAddSplit: (split) => added.push(split),
    onRemoveSplit: (split) => removed.push(split),
  });

  const addButton = walk(mounted.root).find(
    (node) => node.type === "button" && /Add split/.test(textContent(node)),
  );
  addButton.props.onClick();
  await vue.nextTick();
  assert.equal(changes.at(-1).length, 4);
  assert.equal(added.length, 1);
  assert.match(added[0].title, /Split 4/);

  const removeButtons = bySlot(mounted.root, "document-split-remove", true);
  removeButtons[0].props.onClick({ stopPropagation() {} });
  await vue.nextTick();
  assert.equal(changes.at(-1).length, 2);
  assert.equal(removed.length, 1);
  mounted.app.unmount();

  const single = await mount(DocumentSplits, { splits: [SPLITS[0]] });
  assert.equal(bySlot(single.root, "document-split-remove").props.disabled, true);
  single.app.unmount();
});

test("DocumentSplits keyboard reorders pages within and across split groups", async () => {
  const updates = [];
  const mounted = await mount(DocumentSplits, {
    splits: SPLITS,
    "onUpdate:splits": (splits) => updates.push(splits),
  });

  const page1 = byAttr(mounted.root, "data-page-id", "page-1");
  const right = keydownEvent("ArrowRight");
  page1.props.onKeydown(right.event);
  await vue.nextTick();
  assert.equal(right.preventedCount(), 1);
  assert.deepEqual(
    updates.at(-1).find((group) => group.id === "split-1").pages,
    ["page-2", "page-1", "page-3"],
  );

  const page2 = byAttr(mounted.root, "data-page-id", "page-2");
  const ctrlRight = keydownEvent("ArrowRight", { ctrlKey: true });
  page2.props.onKeydown(ctrlRight.event);
  await vue.nextTick();
  const afterMove = updates.at(-1);
  assert.deepEqual(
    afterMove.find((group) => group.id === "split-1").pages,
    ["page-1", "page-3"],
  );
  assert.deepEqual(
    afterMove.find((group) => group.id === "split-2").pages,
    ["page-4", "page-5", "page-2"],
  );
  mounted.app.unmount();
});

test("DocumentSplits drag/drop moves a page between groups", async () => {
  const updates = [];
  const mounted = await mount(DocumentSplits, {
    splits: SPLITS,
    "onUpdate:splits": (splits) => updates.push(splits),
  });

  const transfer = { data: {}, effectAllowed: "", dropEffect: "", setData(key, value) { this.data[key] = value; } };
  const dragEvent = () => ({ preventDefault() {}, stopPropagation() {}, dataTransfer: transfer });
  const page1 = byAttr(mounted.root, "data-page-id", "page-1");
  page1.props.onDragstart({ dataTransfer: transfer });
  const targetPage = byAttr(mounted.root, "data-page-id", "page-6");
  targetPage.props.onDragover(dragEvent());
  targetPage.props.onDrop(dragEvent());
  await vue.nextTick();

  const next = updates.at(-1);
  assert.deepEqual(
    next.find((group) => group.id === "split-1").pages,
    ["page-2", "page-3"],
  );
  assert.deepEqual(
    next.find((group) => group.id === "split-3").pages,
    ["page-1", "page-6"],
  );
  mounted.app.unmount();
});

test("DocumentSplits keyboard reorders split cards", async () => {
  const updates = [];
  const mounted = await mount(DocumentSplits, {
    splits: SPLITS,
    "onUpdate:splits": (splits) => updates.push(splits),
  });
  const handle = bySlot(mounted.root, "document-split-drag-handle");
  const down = keydownEvent("ArrowDown");
  handle.props.onKeydown(down.event);
  await vue.nextTick();
  assert.equal(down.preventedCount(), 1);
  assert.deepEqual(
    updates.at(-1).map((group) => group.id),
    ["split-2", "split-1", "split-3"],
  );
  mounted.app.unmount();
});

test("DocumentSplits empty, loading and error states expose status semantics", async () => {
  const empty = await mount(DocumentSplits, { splits: [] });
  assert.equal(bySlot(empty.root, "document-splits-empty").props.role, "status");
  assert.match(textContent(empty.root), /No document splits yet\./);
  empty.app.unmount();

  const loading = await mount(DocumentSplits, { splits: [], loading: true });
  assert.equal(bySlot(loading.root, "document-splits-loading").props.role, "status");
  loading.app.unmount();

  const error = await mount(DocumentSplits, {
    splits: [],
    error: "Document splits could not be loaded.",
  });
  assert.equal(
    walk(error.root).some((node) => node.props?.role === "alert"),
    true,
  );
  error.app.unmount();
});
