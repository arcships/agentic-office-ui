// Issue #173 — the DOCX form-field dialog and link hover card run on the
// shared Reka primitives (`OfficeDialog` / `OfficePopover`).
//
// Reka needs a real DOM (focus scope, portals, dismissable layers), so this
// file boots happy-dom and drives the components with real events. The
// fake-renderer DOCX suite keeps the model/pagination coverage; these tests
// replace its old hand-rolled overlay assertions.

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test, { after } from "node:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register({ url: "http://localhost/" });

const { importFromDemo, vue, waitFor } = await import("./vue-test-renderer.mjs");
const {
  DocxEditor,
  DocxFormFieldDialog,
  DocxLinkHoverCard,
  useDocxEditor,
} = await importFromDemo("@arcships/vue-docx");
const { createBlankDocumentModel, paragraphText } = await importFromDemo("@arcships/docx-core");

const { createApp, h, nextTick, ref } = vue;

after(() => {
  GlobalRegistrator.unregister();
});

async function flush(ms = 80) {
  await nextTick();
  await new Promise((resolve) => setTimeout(resolve, ms));
  await nextTick();
}

function byTestId(testId, root = document) {
  return root.querySelector(`[data-testid="${testId}"]`);
}

function allByTestId(testId, root = document) {
  return [...root.querySelectorAll(`[data-testid="${testId}"]`)];
}

function click(testId) {
  const element = byTestId(testId);
  assert.ok(element, `${testId} is rendered`);
  element.click();
}

function setInput(testId, value) {
  const element = byTestId(testId);
  assert.ok(element, `${testId} is rendered`);
  element.value = value;
  element.dispatchEvent(new window.Event("input", { bubbles: true }));
}

function setSelect(testId, value) {
  const element = byTestId(testId);
  assert.ok(element, `${testId} is rendered`);
  element.value = value;
  element.dispatchEvent(new window.Event("change", { bubbles: true }));
}

function pressEscape(element) {
  assert.ok(element, "Escape target is rendered");
  element.dispatchEvent(
    new window.KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }),
  );
}

function pressBackdrop() {
  document.body.dispatchEvent(
    new window.PointerEvent("pointerdown", { bubbles: true, cancelable: true, button: 0 }),
  );
}

function mountSetup(setup) {
  const host = document.createElement("div");
  document.body.append(host);
  const app = createApp({ setup });
  app.config.warnHandler = (message) => {
    throw new Error(`Vue warning: ${message}`);
  };
  app.mount(host);
  return {
    app,
    host,
    cleanup() {
      app.unmount();
      host.remove();
    },
  };
}

function mountHost(render) {
  return mountSetup(() => render);
}

// ── Fixtures ───────────────────────────────────────────────────────
function modelWithNodes(nodes) {
  return { ...createBlankDocumentModel(), nodes };
}

function formFieldParagraph(overrides = {}) {
  return {
    type: "paragraph",
    children: [
      {
        type: "form-field",
        fieldType: "dropdown",
        title: "Status",
        value: "open",
        options: [
          { displayText: "Open", value: "open" },
          { displayText: "Closed", value: "closed" },
        ],
        ...overrides,
      },
    ],
  };
}

function linkedParagraph(text, link) {
  return {
    type: "paragraph",
    children: [{ type: "text", text, link }],
  };
}

function bookmarkParagraph(name, text) {
  return {
    type: "paragraph",
    sourceXml:
      `<w:p><w:bookmarkStart w:id="9" w:name="${name}"/>` +
      `<w:r><w:t>${text}</w:t></w:r><w:bookmarkEnd w:id="9"/></w:p>`,
    children: [{ type: "text", text }],
  };
}

// ── Form-field dialog ──────────────────────────────────────────────
function mountFormFieldDialog(model = modelWithNodes([formFieldParagraph()])) {
  const location = { kind: "paragraph", nodeIndex: 0, childIndex: 0 };
  const open = ref(false);
  const trigger = ref();
  const closes = [];
  let controller;
  const mounted = mountSetup(() => {
    controller = useDocxEditor({ starterModel: model });
    return () => h("div", [
      h(
        "button",
        {
          ref: trigger,
          type: "button",
          "data-testid": "dialog-trigger",
          onClick: () => {
            open.value = true;
          },
        },
        "Open dialog",
      ),
      open.value
        ? h(DocxFormFieldDialog, {
            field: controller.model.nodes[0].children[0],
            location,
            controller,
            onClose: () => {
              closes.push("close");
              open.value = false;
            },
            onSaved: (field) => closes.push(["saved", field.value]),
          })
        : null,
    ]);
  });
  return {
    ...mounted,
    get controller() {
      return controller;
    },
    get trigger() {
      return trigger.value;
    },
    closes,
    async openDialog() {
      trigger.value.focus();
      trigger.value.click();
      await flush(120);
    },
  };
}

test("DocxFormFieldDialog renders on OfficeDialog and keeps the surface contract", async (t) => {
  const mounted = mountFormFieldDialog();
  t.after(() => mounted.cleanup());

  await mounted.openDialog();
  const dialog = byTestId("docx-form-field-dialog");
  assert.ok(dialog, "the dialog surface is portalled");
  assert.equal(dialog.getAttribute("role"), "dialog");
  assert.equal(dialog.getAttribute("aria-modal"), "true");
  assert.equal(dialog.getAttribute("data-docx-form-field-dialog"), "true");
  assert.equal(dialog.getAttribute("data-office-theme"), "light");
  assert.ok(document.querySelector(".office-overlay__backdrop"), "the shared backdrop renders");
  assert.ok(byTestId("docx-form-field-dialog-close"), "the shared close button keeps an id");
  assert.equal(
    document.querySelector("[data-slot='dialog-title']").textContent.trim(),
    "Form field properties",
  );

  // Reka focuses its first tabbable control, then the component moves focus to
  // the name field (the historical first-field contract).
  assert.equal(document.activeElement?.getAttribute("data-testid"), "docx-form-field-dialog-title");
});

test("DocxFormFieldDialog edits dropdown options and saves through the controller", async (t) => {
  const mounted = mountFormFieldDialog();
  t.after(() => mounted.cleanup());
  await mounted.openDialog();

  assert.equal(byTestId("docx-form-field-dialog-title").value, "Status");
  click("docx-form-field-dialog-add-option");
  await flush();
  setInput("docx-form-field-dialog-option-2", "Escalated");
  await flush();
  setSelect("docx-form-field-dialog-value", "Escalated");
  await flush();
  click("docx-form-field-dialog-save");
  await flush(120);

  const field = mounted.controller.model.nodes[0].children[0];
  assert.equal(field.value, "Escalated");
  assert.deepEqual(field.options.map((option) => option.displayText), [
    "Open",
    "Closed",
    "Escalated",
  ]);
  assert.deepEqual(mounted.closes, [["saved", "Escalated"], "close"]);
});

test("DocxFormFieldDialog keeps field validation inside the dialog", async (t) => {
  const mounted = mountFormFieldDialog(
    modelWithNodes([
      formFieldParagraph({
        fieldType: "text",
        required: true,
        value: "abcdef",
        widget: { text: { maxLength: 3 } },
      }),
    ]),
  );
  t.after(() => mounted.cleanup());
  await mounted.openDialog();

  click("docx-form-field-dialog-save");
  await flush();

  const error = document.querySelector(".docx-form-field-dialog__error");
  assert.ok(error, "an invalid value keeps the dialog open");
  assert.equal(error.getAttribute("role"), "alert");
  assert.ok(error.textContent.trim().length > 0);
  assert.ok(byTestId("docx-form-field-dialog"), "the dialog did not close");
});

test("DocxFormFieldDialog closes on Escape and the backdrop and returns focus", async (t) => {
  const mounted = mountFormFieldDialog();
  t.after(() => mounted.cleanup());
  await mounted.openDialog();

  pressEscape(byTestId("docx-form-field-dialog"));
  await flush(120);
  assert.equal(byTestId("docx-form-field-dialog"), null, "Escape closes");
  assert.equal(document.activeElement, mounted.trigger, "focus returns to the opener");

  await mounted.openDialog();
  pressBackdrop();
  await flush(120);
  assert.equal(byTestId("docx-form-field-dialog"), null, "backdrop pointer-down closes");
  assert.equal(mounted.closes.filter((entry) => entry === "close").length, 2);
});

test("DocxFormFieldDialog follows the document theme", async (t) => {
  const mounted = mountFormFieldDialog();
  t.after(() => mounted.cleanup());
  mounted.controller.setDocumentTheme("dark");
  await flush();
  await mounted.openDialog();

  const dialog = byTestId("docx-form-field-dialog");
  assert.equal(dialog.getAttribute("data-office-theme"), "dark");
  assert.ok(document.querySelector(".docx-form-field-dialog--dark"));
});

test("DocxEditor opens the form-field dialog from the field affordance", async (t) => {
  const model = modelWithNodes([formFieldParagraph()]);
  let controller;
  let editor;
  const mounted = mountSetup(() => {
    controller = useDocxEditor({ starterModel: model });
    return () => h(DocxEditor, {
      ref: (instance) => {
        editor = instance;
      },
      editor: controller,
      editable: true,
      showToolbar: false,
      showThumbnails: false,
    });
  });
  t.after(() => mounted.cleanup());
  await flush(200);

  const root = byTestId("docx-editor");
  assert.ok(root);
  const field = document.createElement("div");
  field.setAttribute("data-docx-form-field-location", "p:0:0");
  root.append(field);
  field.dispatchEvent(new window.MouseEvent("dblclick", { bubbles: true, cancelable: true }));
  await flush(120);

  assert.deepEqual(controller.selectedFormField?.location, {
    kind: "paragraph",
    nodeIndex: 0,
    childIndex: 0,
  });
  assert.ok(byTestId("docx-form-field-dialog"), "double click opens the dialog");
  assert.ok(editor);
  click("docx-form-field-dialog-cancel");
  await flush(120);
  assert.equal(byTestId("docx-form-field-dialog"), null);
});

test("DocxEditor keeps the form-field dialog off when the affordance is disabled", async () => {
  const model = modelWithNodes([formFieldParagraph()]);
  let controller;
  const mounted = mountSetup(() => {
    controller = useDocxEditor({ starterModel: model });
    return () => h(DocxEditor, {
      editor: controller,
      editable: true,
      showToolbar: false,
      showThumbnails: false,
      showFormFieldDialog: false,
    });
  });
  try {
    await flush(200);
    const root = byTestId("docx-editor");
    const field = document.createElement("div");
    field.setAttribute("data-docx-form-field-location", "p:0:0");
    root.append(field);
    field.dispatchEvent(new window.MouseEvent("dblclick", { bubbles: true, cancelable: true }));
    await flush(100);
    assert.equal(byTestId("docx-form-field-dialog"), null);
  } finally {
    mounted.cleanup();
  }
});

// ── Link hover card ────────────────────────────────────────────────
function mountLinkCard(props = {}) {
  const events = [];
  const mounted = mountHost(() =>
    h(DocxLinkHoverCard, {
      href: "https://example.com/page",
      x: 12,
      y: 34,
      onOpen: (href) => events.push(["open", href]),
      onEdit: (href) => events.push(["edit", href]),
      onRemove: (href) => events.push(["remove", href]),
      onClose: () => events.push(["close"]),
      onPointerenter: () => events.push(["pointerenter"]),
      onPointerleave: () => events.push(["pointerleave"]),
      ...props,
    }),
  );
  return { ...mounted, events };
}

test("DocxLinkHoverCard renders on OfficePopover with a virtual anchor", async (t) => {
  const mounted = mountLinkCard({ documentTheme: "dark" });
  t.after(() => mounted.cleanup());
  await flush(160);

  const card = byTestId("docx-link-card");
  assert.ok(card, "the hover card surface is portalled");
  assert.equal(card.getAttribute("role"), "dialog");
  assert.equal(card.getAttribute("aria-label"), "Link actions");
  assert.equal(card.getAttribute("aria-labelledby"), null);
  assert.equal(card.getAttribute("data-docx-link-card"), "true");
  assert.equal(card.getAttribute("data-office-theme"), "dark");
  assert.equal(document.querySelectorAll("[data-slot='popover-anchor']").length, 1);
  assert.equal(document.querySelector("[data-slot='popover-trigger']"), null);
  assert.equal(card.querySelectorAll(".docx-link-card__action").length, 3);
  assert.ok(card.querySelector(".docx-link-card--dark"));
  assert.equal(byTestId("docx-link-card-url").getAttribute("title"), "https://example.com/page");
});

test("DocxLinkHoverCard exposes open/edit/remove and pointer events", async (t) => {
  const mounted = mountLinkCard();
  t.after(() => mounted.cleanup());
  await flush(120);
  const card = byTestId("docx-link-card");

  card.dispatchEvent(new window.PointerEvent("pointerenter"));
  card.dispatchEvent(new window.PointerEvent("pointerleave"));
  assert.deepEqual(mounted.events.slice(0, 2), [["pointerenter"], ["pointerleave"]]);

  click("docx-link-card-open");
  assert.deepEqual(mounted.events.at(-1), ["open", "https://example.com/page"]);

  click("docx-link-card-edit");
  await flush();
  assert.equal(byTestId("docx-link-card-input").value, "https://example.com/page");
  setInput("docx-link-card-input", "https://example.com/next");
  await flush();
  click("docx-link-card-save");
  assert.deepEqual(mounted.events.at(-1), ["edit", "https://example.com/next"]);
  await flush();

  click("docx-link-card-edit");
  await flush();
  setInput("docx-link-card-input", "");
  await flush();
  click("docx-link-card-save");
  assert.deepEqual(mounted.events.at(-1), ["remove", "https://example.com/page"]);
});

test("DocxLinkHoverCard hides editing actions for non-editable links", async (t) => {
  const mounted = mountLinkCard({ href: "#bookmark", editable: false });
  t.after(() => mounted.cleanup());
  await flush(120);

  assert.ok(byTestId("docx-link-card-open"));
  assert.equal(byTestId("docx-link-card-edit"), null);
  assert.equal(byTestId("docx-link-card-remove"), null);
  assert.equal(
    byTestId("docx-link-card-url").textContent.trim(),
    "Links to a place in this document",
  );
});

test("DocxLinkHoverCard closes on Escape", async (t) => {
  const mounted = mountLinkCard();
  t.after(() => mounted.cleanup());
  await flush(120);

  pressEscape(byTestId("docx-link-card"));
  await flush();
  assert.deepEqual(mounted.events.filter((event) => event[0] === "close"), [["close"]]);
});

// ── Editor integration ─────────────────────────────────────────────
function mountEditor(model) {
  let controller;
  const mounted = mountSetup(() => {
    controller = useDocxEditor({ starterModel: model });
    return () => h(DocxEditor, {
      editor: controller,
      editable: true,
      showToolbar: false,
      showThumbnails: false,
    });
  });
  return {
    ...mounted,
    get controller() {
      return controller;
    },
  };
}

function appendLinkAnchor(editorRoot, { href, start = 0, end = 7, nodeIndex = 0 }) {
  const host = document.createElement("div");
  host.setAttribute("data-docx-paragraph-host", "true");
  host.setAttribute("data-docx-paragraph-node-index", String(nodeIndex));
  const anchor = document.createElement("a");
  anchor.setAttribute("href", href);
  anchor.setAttribute("data-docx-link-start", String(start));
  anchor.setAttribute("data-docx-link-end", String(end));
  anchor.textContent = "Example";
  host.append(anchor);
  editorRoot.append(host);
  return anchor;
}

test("hovering an editor link shows the card and remove clears the model link", async (t) => {
  const mounted = mountEditor(modelWithNodes([linkedParagraph("Example", "https://example.com")]));
  t.after(() => mounted.cleanup());
  await flush(200);

  const anchor = appendLinkAnchor(byTestId("docx-editor"), {
    href: "https://example.com",
  });
  anchor.dispatchEvent(new window.PointerEvent("pointerover", { bubbles: true }));
  await flush(160);

  assert.ok(byTestId("docx-link-card"), "the link card renders");
  click("docx-link-card-remove");
  await flush(120);

  const run = mounted.controller.model.nodes[0].children[0];
  assert.equal(run.link, undefined);
  assert.equal(paragraphText(mounted.controller.model.nodes[0]), "Example");
  assert.ok(mounted.controller.canUndo);
});

test("the link card opens external links and reports internal bookmarks to the host", async (t) => {
  const opened = [];
  const openedLinks = [];
  const originalOpen = window.open;
  window.open = (...args) => {
    opened.push(args);
    return null;
  };
  const model = modelWithNodes([linkedParagraph("Example", "https://example.com")]);
  let controller;
  const mounted = mountSetup(() => {
    controller = useDocxEditor({ starterModel: model });
    return () => h(DocxEditor, {
      editor: controller,
      editable: true,
      showToolbar: false,
      showThumbnails: false,
      onLinkOpen: (href) => openedLinks.push(href),
    });
  });
  t.after(() => {
    mounted.cleanup();
    window.open = originalOpen;
  });
  await flush(200);

  const root = byTestId("docx-editor");
  const anchor = appendLinkAnchor(root, { href: "https://example.com" });
  anchor.dispatchEvent(new window.PointerEvent("pointerover", { bubbles: true }));
  await flush(160);
  click("docx-link-card-open");
  await flush();

  assert.deepEqual(opened, [["https://example.com", "_blank", "noopener,noreferrer"]]);
  assert.deepEqual(openedLinks, ["https://example.com"]);
});

test("the link card navigates internal bookmarks and notifies the host", async (t) => {
  const openedLinks = [];
  const model = modelWithNodes([
    linkedParagraph("Jump", "#Target_Heading"),
    bookmarkParagraph("Target_Heading", "Destination"),
  ]);
  let controller;
  const mounted = mountSetup(() => {
    controller = useDocxEditor({ starterModel: model });
    return () => h(DocxEditor, {
      editor: controller,
      editable: true,
      showToolbar: false,
      showThumbnails: false,
      onLinkOpen: (href) => openedLinks.push(href),
    });
  });
  t.after(() => mounted.cleanup());
  await flush(200);

  const anchor = appendLinkAnchor(byTestId("docx-editor"), {
    href: "#Target_Heading",
    start: 0,
    end: 4,
  });
  anchor.dispatchEvent(new window.PointerEvent("pointerover", { bubbles: true }));
  await flush(160);
  click("docx-link-card-open");
  await flush(200);

  assert.deepEqual(openedLinks, ["#Target_Heading"]);
  await waitFor(
    () => controller.selection.kind === "paragraph" && controller.selection.nodeIndex === 1,
    "the bookmark target to take the caret",
  );
});

// ── Source guards ─────────────────────────────────────────────────
test("DOCX overlays render shared primitives instead of hand-rolled shells", () => {
  const dialogSource = readFileSync(
    new URL("../../packages/vue-docx/src/components/DocxFormFieldDialog.vue", import.meta.url),
    "utf8",
  );
  assert.match(dialogSource, /<OfficeDialog/);
  assert.match(dialogSource, /content-test-id="docx-form-field-dialog"/);
  assert.match(dialogSource, /data-docx-form-field-dialog/);
  assert.equal(/docx-form-field-dialog-backdrop/.test(dialogSource), false);
  assert.equal(/role="dialog"/.test(dialogSource), false);
  assert.equal(/<OfficeCard/.test(dialogSource), false);
  assert.equal(/<Teleport/.test(dialogSource), false);

  const cardSource = readFileSync(
    new URL("../../packages/vue-docx/src/components/DocxLinkHoverCard.vue", import.meta.url),
    "utf8",
  );
  assert.match(cardSource, /<OfficePopover/);
  assert.match(cardSource, /content-test-id="docx-link-card"/);
  assert.match(cardSource, /:virtual-ref="virtualAnchor"/);
  assert.match(cardSource, /side-offset="0"/);
  assert.match(cardSource, /Math\.max\(0, Math\.round\(props\.x\)\)/);
  assert.equal(/role="dialog"/.test(cardSource), false);
  assert.equal(/<OfficeCard/.test(cardSource), false);

  const popoverSource = readFileSync(
    new URL("../../packages/vue-ui/src/components/ui/OfficePopover.vue", import.meta.url),
    "utf8",
  );
  assert.match(popoverSource, /PopoverAnchor/);
  assert.match(popoverSource, /contentTestId/);
  assert.match(popoverSource, /virtualRef/);
});
