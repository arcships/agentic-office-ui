// Issue #157 — primitive layer runtime coverage.
//
// The repository's other component tests drive a minimal host renderer. Reka UI
// needs a real DOM (focus, pointer, portals, resize observers), so this file
// boots happy-dom before importing Vue and the built package. It covers every
// primitive, the keyboard/ARIA contract and the interactive focus trap.

import assert from "node:assert/strict";
import test, { after } from "node:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register({ url: "http://localhost/" });

const { createRequire } = await import("node:module");
const { pathToFileURL } = await import("node:url");

const requireFromDemo = createRequire(
  new URL("../../apps/demo/package.json", import.meta.url),
);
const vue = await import(pathToFileURL(requireFromDemo.resolve("vue")).href);
const ui = await import(
  pathToFileURL(requireFromDemo.resolve("@arcships/vue-ui")).href
);
const { renderToString } = await import(
  pathToFileURL(requireFromDemo.resolve("vue/server-renderer")).href
);

const { createApp, defineComponent, h, nextTick, ref } = vue;
const {
  OfficeBadge,
  OfficeButton,
  OfficeCard,
  OfficeContextMenu,
  OfficeDialog,
  OfficeDropdownMenu,
  OfficeIconButton,
  OfficeInput,
  OfficePopover,
  OfficeScrollArea,
  OfficeSelect,
  OfficeSeparator,
  OfficeSheet,
  OfficeSkeleton,
  OfficeSpinner,
  OfficeTabs,
  OfficeTextarea,
  OfficeThemeProvider,
  OfficeToastProvider,
  OfficeToggle,
  OfficeTooltip,
  useOfficeToast,
} = ui;

after(() => {
  GlobalRegistrator.unregister();
});

function mountRoot(render) {
  const host = document.createElement("div");
  document.body.append(host);
  const app = createApp({ render });
  app.mount(host);
  const cleanup = () => {
    app.unmount();
    host.remove();
    document.body.innerHTML = "";
  };
  return { app, host, cleanup };
}

function mount(component, { props = {}, slots = {} } = {}) {
  return mountRoot(() => h(component, props, slots));
}

async function flush(ms = 60) {
  await nextTick();
  await new Promise((resolve) => setTimeout(resolve, ms));
  await nextTick();
}

const bySlot = (slot, root = document) =>
  root.querySelector(`[data-slot="${slot}"]`);
const allBySlot = (slot, root = document) =>
  [...root.querySelectorAll(`[data-slot="${slot}"]`)];

function openSelect(trigger) {
  // Select opens on pointerdown and treats the following click as the same
  // gesture; an extra pointerup would toggle it closed again.
  trigger.dispatchEvent(
    new PointerEvent("pointerdown", { bubbles: true, button: 0 }),
  );
  trigger.click();
}

function click(target) {
  target.dispatchEvent(
    new PointerEvent("pointerdown", { bubbles: true, button: 0 }),
  );
  target.dispatchEvent(
    new PointerEvent("pointerup", { bubbles: true, button: 0 }),
  );
  target.click();
}

// ── Every primitive renders ────────────────────────────────────────────────

test("every primitive mounts and exposes its data-slot", async (t) => {
  const tabValue = ref("one");
  const selectValue = ref("docx");
  const toggleValue = ref(false);
  const dialogOpen = ref(false);
  const sheetOpen = ref(false);

  const { cleanup } = mountRoot(() =>
    h(OfficeToastProvider, null, () => [
      h(OfficeButton, null, () => "Button"),
      h(OfficeIconButton, { label: "More", icon: "more-horizontal" }),
      h(OfficeInput, { modelValue: "input" }),
      h(OfficeTextarea, { modelValue: "textarea" }),
      h(OfficeSelect, {
        modelValue: selectValue.value,
        "onUpdate:modelValue": (value) => (selectValue.value = value),
        options: [{ value: "docx", label: "DOCX" }],
      }),
      h(OfficeDropdownMenu, { items: [{ label: "Copy" }], label: "Menu" }),
      h(OfficeContextMenu, { items: [{ label: "Copy" }] }, () =>
        h("div", "Context target"),
      ),
      h(
        OfficeDialog,
        {
          open: dialogOpen.value,
          "onUpdate:open": (value) => (dialogOpen.value = value),
          title: "Dialog",
        },
        {
          trigger: () => h("button", "Open dialog"),
          default: () => "Dialog body",
        },
      ),
      h(
        OfficeSheet,
        {
          open: sheetOpen.value,
          "onUpdate:open": (value) => (sheetOpen.value = value),
          title: "Sheet",
        },
        {
          trigger: () => h("button", "Open sheet"),
          default: () => "Sheet body",
        },
      ),
      h(OfficePopover, { label: "Popover" }, () => "Popover body"),
      h(OfficeTooltip, { content: "Tooltip" }, () => h("button", "Tip")),
      h(
        OfficeTabs,
        {
          modelValue: tabValue.value,
          "onUpdate:modelValue": (value) => (tabValue.value = value),
          tabs: [
            { value: "one", label: "One" },
            { value: "two", label: "Two" },
          ],
        },
        {
          "panel-one": () => "Panel one",
          "panel-two": () => "Panel two",
        },
      ),
      h(
        OfficeToggle,
        {
          modelValue: toggleValue.value,
          "onUpdate:modelValue": (value) => (toggleValue.value = value),
          label: "Bold",
        },
        () => "B",
      ),
      h(OfficeBadge, null, () => "Badge"),
      h(OfficeCard, { title: "Card" }, () => "Card body"),
      h(OfficeSeparator),
      h(OfficeSkeleton),
      h(OfficeSpinner),
      h(OfficeScrollArea, null, () => h("div", "Scroll body")),
    ]),
  );
  t.after(cleanup);
  await flush();

  for (const slot of [
    "button",
    "icon-button",
    "input",
    "textarea",
    "select-trigger",
    "dropdown-menu-trigger",
    "context-menu-trigger",
    "dialog-trigger",
    "sheet-trigger",
    "popover-trigger",
    "tooltip-trigger",
    "tabs",
    "toggle",
    "badge",
    "card",
    "separator",
    "skeleton",
    "spinner",
    "scroll-area",
    "toast-viewport",
  ]) {
    assert.ok(bySlot(slot), `missing data-slot="${slot}"`);
  }
});

// ── Button / IconButton ───────────────────────────────────────────────────

test("Button applies variants, sizes, disabled and loading", async (t) => {
  const clicks = [];
  const { cleanup } = mountRoot(() =>
    h("div", [
      h(
        OfficeButton,
        {
          variant: "destructive",
          size: "lg",
          onClick: () => clicks.push("click"),
        },
        () => "Delete",
      ),
      h(OfficeButton, { disabled: true, onClick: () => clicks.push("disabled") }, () =>
        "Disabled",
      ),
      h(OfficeButton, { loading: true, onClick: () => clicks.push("loading") }, () =>
        "Saving",
      ),
      h(OfficeButton, { asChild: true }, () =>
        h("a", { href: "#target", "data-testid": "button-link" }, "Link"),
      ),
    ]),
  );
  t.after(cleanup);
  await flush();

  const [primary, disabled, loading] = [...document.querySelectorAll("button")];
  assert.match(primary.className, /office-button--destructive/);
  assert.match(primary.className, /office-button--size-lg/);
  primary.click();
  assert.deepEqual(clicks, ["click"]);

  assert.equal(disabled.disabled, true);
  assert.equal(disabled.getAttribute("data-disabled"), "true");
  disabled.click();
  assert.deepEqual(clicks, ["click"], "disabled button must not emit click");

  assert.equal(loading.getAttribute("aria-busy"), "true");
  assert.equal(loading.getAttribute("data-loading"), "true");
  assert.equal(loading.disabled, true);
  assert.ok(loading.querySelector('[data-slot="spinner"]'));
  assert.match(loading.textContent, /Saving/);

  const link = document.querySelector('[data-testid="button-link"]');
  assert.equal(link.tagName, "A");
  assert.equal(link.getAttribute("href"), "#target");
  assert.match(link.className, /office-button--default/);
});

test("IconButton keeps an accessible label and pressed state", async (t) => {
  const { cleanup } = mountRoot(() =>
    h(OfficeIconButton, {
      label: "More actions",
      icon: "more-horizontal",
      active: true,
    }),
  );
  t.after(cleanup);
  await flush();

  const button = bySlot("icon-button");
  assert.equal(button.getAttribute("aria-label"), "More actions");
  assert.equal(button.getAttribute("title"), "More actions");
  assert.equal(button.getAttribute("aria-pressed"), "true");
  assert.equal(button.getAttribute("data-active"), "true");
  assert.ok(button.querySelector("svg"));
});

// ── Input / Textarea ──────────────────────────────────────────────────────

test("Input and Textarea support v-model, sizes, invalid and disabled", async (t) => {
  const text = ref("start");
  const notes = ref("note");
  const { cleanup } = mountRoot(() =>
    h("div", [
      h(OfficeInput, {
        modelValue: text.value,
        "onUpdate:modelValue": (value) => (text.value = value),
        size: "sm",
        "data-testid": "input",
      }),
      h(OfficeTextarea, {
        modelValue: notes.value,
        "onUpdate:modelValue": (value) => (notes.value = value),
        "data-testid": "textarea",
      }),
      h(OfficeInput, { modelValue: "bad", invalid: true, "data-testid": "invalid" }),
      h(OfficeInput, { modelValue: "off", disabled: true, "data-testid": "off" }),
    ]),
  );
  t.after(cleanup);
  await flush();

  const input = document.querySelector('[data-testid="input"]');
  assert.match(
    input.closest("[data-slot=input-control]").className,
    /office-input-control--sm/,
  );
  input.value = "changed";
  input.dispatchEvent(new Event("input", { bubbles: true }));
  await flush(0);
  assert.equal(text.value, "changed");

  const textarea = document.querySelector('[data-testid="textarea"]');
  textarea.value = "new note";
  textarea.dispatchEvent(new Event("input", { bubbles: true }));
  await flush(0);
  assert.equal(notes.value, "new note");

  assert.equal(
    document.querySelector('[data-testid="invalid"]').getAttribute("aria-invalid"),
    "true",
  );
  assert.equal(document.querySelector('[data-testid="off"]').disabled, true);
});

// ── Select ────────────────────────────────────────────────────────────────

test("Select opens a portal, exposes listbox ARIA and commits a value", async (t) => {
  const value = ref("docx");
  const { cleanup } = mountRoot(() =>
    h(OfficeSelect, {
      id: "format-select",
      modelValue: value.value,
      "onUpdate:modelValue": (next) => (value.value = next),
      options: [
        {
          type: "group",
          label: "Documents",
          options: [
            { value: "docx", label: "Word" },
            { value: "pdf", label: "PDF" },
          ],
        },
        { type: "separator" },
        { value: "pptx", label: "Slides", disabled: true },
      ],
      placeholder: "Choose",
    }),
  );
  t.after(cleanup);
  await flush();

  const trigger = bySlot("select-trigger");
  assert.equal(trigger.getAttribute("role"), "combobox");
  assert.equal(trigger.getAttribute("aria-expanded"), "false");
  assert.equal(trigger.textContent.trim(), "Word");

  openSelect(trigger);
  await flush();

  assert.equal(trigger.getAttribute("aria-expanded"), "true");
  const listbox = document.querySelector('[role="listbox"]');
  assert.ok(listbox, "select content must render a listbox");
  assert.ok(bySlot("select-group"));
  assert.ok(bySlot("select-separator"));
  assert.equal(allBySlot("select-item").length, 3);
  assert.equal(
    allBySlot("select-item")[2].hasAttribute("data-disabled"),
    true,
  );

  click(allBySlot("select-item")[1]);
  await flush();

  assert.equal(value.value, "pdf");
  assert.equal(trigger.textContent.trim(), "PDF");
  assert.equal(trigger.getAttribute("aria-expanded"), "false");
});

test("Select supports keyboard selection and a disabled trigger", async (t) => {
  const value = ref("a");
  const { cleanup } = mountRoot(() =>
    h("div", [
      h(OfficeSelect, {
        modelValue: value.value,
        "onUpdate:modelValue": (next) => (value.value = next),
        options: [
          { value: "a", label: "Alpha" },
          { value: "b", label: "Beta" },
        ],
        id: "enabled",
      }),
      h(OfficeSelect, {
        modelValue: "a",
        options: [{ value: "a", label: "Alpha" }],
        disabled: true,
        id: "disabled",
      }),
    ]),
  );
  t.after(cleanup);
  await flush();

  const trigger = document.querySelector("#enabled");
  trigger.focus();
  trigger.dispatchEvent(
    new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }),
  );
  await flush();
  assert.equal(trigger.getAttribute("aria-expanded"), "true");

  // Reka moves focus to the highlighted option; keyboard navigation is then
  // handled on the option itself.
  const activeOption = document.activeElement;
  activeOption.dispatchEvent(
    new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true, cancelable: true }),
  );
  await flush(20);
  document.activeElement.dispatchEvent(
    new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }),
  );
  await flush();
  assert.equal(value.value, "b");

  const disabled = document.querySelector("#disabled");
  assert.equal(disabled.disabled, true);
  assert.equal(disabled.hasAttribute("data-disabled"), true);
});

// ── Menus ─────────────────────────────────────────────────────────────────

test("DropdownMenu opens, renders entries and reports selection", async (t) => {
  const selected = [];
  const open = ref(false);
  const { cleanup } = mountRoot(() =>
    h(OfficeDropdownMenu, {
      open: open.value,
      "onUpdate:open": (value) => (open.value = value),
      label: "Open menu",
      items: [
        { label: "Copy", shortcut: "⌘C" },
        { type: "separator" },
        { type: "label", label: "Danger" },
        { label: "Delete", variant: "destructive", disabled: true },
      ],
      onSelect: (entry) => selected.push(entry.label),
    }),
  );
  t.after(cleanup);
  await flush();

  const trigger = bySlot("dropdown-menu-trigger");
  click(trigger);
  await flush();

  assert.equal(open.value, true);
  assert.ok(document.querySelector('[role="menu"]'));
  assert.equal(allBySlot("menu-item").length, 2);
  assert.ok(bySlot("menu-separator"));
  assert.ok(bySlot("menu-label"));
  assert.equal(trigger.getAttribute("aria-expanded"), "true");

  click(allBySlot("menu-item")[0]);
  await flush();

  assert.deepEqual(selected, ["Copy"]);
  assert.equal(open.value, false);
});

test("ContextMenu opens on contextmenu and closes with Escape", async (t) => {
  const selected = [];
  const { cleanup } = mountRoot(() =>
    h(
      OfficeContextMenu,
      {
        items: [{ label: "Inspect" }, { type: "separator" }, { label: "Copy" }],
        onSelect: (entry) => selected.push(entry.label),
      },
      () => h("div", { "data-testid": "target" }, "Right click"),
    ),
  );
  t.after(cleanup);
  await flush();

  const target = document.querySelector('[data-testid="target"]');
  target.dispatchEvent(
    new MouseEvent("contextmenu", {
      bubbles: true,
      cancelable: true,
      clientX: 12,
      clientY: 12,
    }),
  );
  await flush();

  assert.ok(document.querySelector('[role="menu"]'));
  const firstItem = allBySlot("menu-item")[0];
  click(firstItem);
  await flush();
  assert.deepEqual(selected, ["Inspect"]);

  target.dispatchEvent(
    new MouseEvent("contextmenu", {
      bubbles: true,
      cancelable: true,
      clientX: 14,
      clientY: 14,
    }),
  );
  await flush();
  const content = bySlot("context-menu-content");
  content.dispatchEvent(
    new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }),
  );
  await flush();
  assert.equal(document.querySelector('[role="menu"]'), null);
});

// ── Dialog / Sheet ────────────────────────────────────────────────────────

test("Dialog traps focus, labels itself and closes on Escape", async (t) => {
  const open = ref(false);
  const { cleanup } = mountRoot(() =>
    h(
      OfficeDialog,
      {
        open: open.value,
        "onUpdate:open": (value) => (open.value = value),
        title: "Rename document",
        description: "Choose a new name.",
      },
      {
        trigger: () => h("button", { "data-testid": "dialog-trigger" }, "Open"),
        default: () =>
          h("div", [
            h("button", { "data-testid": "dialog-body-button" }, "Body action"),
          ]),
        footer: () => h("button", { "data-testid": "dialog-footer-button" }, "Save"),
      },
    ),
  );
  t.after(cleanup);
  await flush();

  document.querySelector('[data-testid="dialog-trigger"]').click();
  await flush();

  assert.equal(open.value, true);
  const dialog = document.querySelector('[role="dialog"]');
  assert.ok(dialog);
  assert.equal(dialog.getAttribute("data-slot"), "dialog-content");
  const title = bySlot("dialog-title");
  const description = bySlot("dialog-description");
  assert.equal(
    dialog.getAttribute("aria-labelledby"),
    title.getAttribute("id"),
  );
  assert.equal(
    dialog.getAttribute("aria-describedby"),
    description.getAttribute("id"),
  );
  assert.equal(title.textContent.trim(), "Rename document");
  assert.equal(description.textContent.trim(), "Choose a new name.");
  assert.ok(dialog.contains(document.activeElement), "focus must move into the dialog");

  // Tab from the last focusable element must stay inside the dialog.
  const close = bySlot("dialog-close");
  close.focus();
  close.dispatchEvent(
    new KeyboardEvent("keydown", { key: "Tab", bubbles: true, cancelable: true }),
  );
  await flush(20);
  assert.ok(dialog.contains(document.activeElement), "Tab must not escape the dialog");

  dialog.dispatchEvent(
    new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }),
  );
  await flush();
  assert.equal(open.value, false);
  assert.equal(document.querySelector('[role="dialog"]'), null);
});

test("Dialog closeOnEscape and closeOnBackdrop can be disabled", async (t) => {
  const open = ref(true);
  const { cleanup } = mountRoot(() =>
    h(
      OfficeDialog,
      {
        open: open.value,
        "onUpdate:open": (value) => (open.value = value),
        title: "Sticky dialog",
        closeOnEscape: false,
        closeOnBackdrop: false,
      },
      { default: () => "Body" },
    ),
  );
  t.after(cleanup);
  await flush();

  const dialog = document.querySelector('[role="dialog"]');
  dialog.dispatchEvent(
    new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }),
  );
  document.body.dispatchEvent(
    new PointerEvent("pointerdown", { bubbles: true, button: 0 }),
  );
  await flush();
  assert.equal(open.value, true, "dialog must stay open");
  assert.ok(document.querySelector('[role="dialog"]'));
});

test("Dialog portals into the themed subtree instead of document.body", async (t) => {
  const { cleanup } = mountRoot(() =>
    h(OfficeThemeProvider, { theme: "dark" }, () =>
      h(
        OfficeDialog,
        { open: true, title: "Themed dialog" },
        { default: () => "Body" },
      ),
    ),
  );
  t.after(cleanup);
  await flush();

  const content = bySlot("dialog-content");
  assert.ok(content);
  assert.equal(
    content.parentElement.classList.contains("office-theme__portal"),
    true,
    "dialog must teleport into the provider's portal host",
  );
  assert.equal(
    content.closest("[data-office-theme]")?.getAttribute("data-office-theme"),
    "dark",
    "portalled content must inherit theme tokens",
  );
});

test("Sheet renders on the requested edge and closes with Escape", async (t) => {
  const open = ref(false);
  const { cleanup } = mountRoot(() =>
    h(
      OfficeSheet,
      {
        open: open.value,
        "onUpdate:open": (value) => (open.value = value),
        title: "Details",
        side: "left",
      },
      {
        trigger: () => h("button", { "data-testid": "sheet-trigger" }, "Open sheet"),
        default: () => "Panel body",
      },
    ),
  );
  t.after(cleanup);
  await flush();

  document.querySelector('[data-testid="sheet-trigger"]').click();
  await flush();

  const sheet = bySlot("sheet-content");
  assert.ok(sheet);
  assert.match(sheet.className, /office-sheet__content--left/);
  assert.equal(sheet.getAttribute("role"), "dialog");
  assert.equal(bySlot("sheet-title").textContent.trim(), "Details");

  sheet.dispatchEvent(
    new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }),
  );
  await flush();
  assert.equal(open.value, false);
});

// ── Popover / Tooltip ─────────────────────────────────────────────────────

test("Popover opens on the trigger and closes on Escape", async (t) => {
  const open = ref(false);
  const { cleanup } = mountRoot(() =>
    h(
      OfficePopover,
      {
        open: open.value,
        "onUpdate:open": (value) => (open.value = value),
        label: "Open popover",
      },
      () => h("p", { "data-testid": "popover-body" }, "Popover body"),
    ),
  );
  t.after(cleanup);
  await flush();

  click(bySlot("popover-trigger"));
  await flush();

  const content = bySlot("popover-content");
  assert.ok(content);
  assert.equal(content.textContent.trim(), "Popover body");
  assert.equal(open.value, true);

  content.dispatchEvent(
    new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }),
  );
  await flush();
  assert.equal(open.value, false);
});

test("Tooltip shows on focus, links itself and honours disabled", async (t) => {
  const { cleanup } = mountRoot(() =>
    h("div", [
      h(
        OfficeTooltip,
        { content: "Copy to clipboard" },
        () => h("button", { "data-testid": "tooltip-trigger" }, "Copy"),
      ),
      h(
        OfficeTooltip,
        { content: "Never shown", disabled: true },
        () => h("button", { "data-testid": "tooltip-disabled" }, "Disabled"),
      ),
    ]),
  );
  t.after(cleanup);
  await flush();

  const trigger = document.querySelector('[data-testid="tooltip-trigger"]');
  trigger.focus();
  trigger.dispatchEvent(new FocusEvent("focus", { bubbles: true }));
  await flush(400);

  const tooltip = bySlot("tooltip-content");
  assert.ok(tooltip, "visible tooltip bubble must mount");
  assert.match(tooltip.textContent, /Copy to clipboard/);
  const description = document.querySelector('[role="tooltip"]');
  assert.ok(description, "Reka renders the accessible description");
  assert.equal(
    trigger.getAttribute("aria-describedby"),
    description.getAttribute("id"),
  );

  trigger.blur();
  trigger.dispatchEvent(new FocusEvent("blur", { bubbles: false }));
  await flush(200);

  const disabled = document.querySelector('[data-testid="tooltip-disabled"]');
  disabled.focus();
  disabled.dispatchEvent(new FocusEvent("focus", { bubbles: true }));
  await flush(400);
  assert.equal(bySlot("tooltip-content"), null);
});

// ── Tabs / Toggle ─────────────────────────────────────────────────────────

test("Tabs follow the roving tabindex and keyboard activation contract", async (t) => {
  const value = ref("one");
  const { cleanup } = mountRoot(() =>
    h(
      OfficeTabs,
      {
        modelValue: value.value,
        "onUpdate:modelValue": (next) => (value.value = next),
        tabs: [
          { value: "one", label: "One" },
          { value: "two", label: "Two" },
          { value: "three", label: "Three", disabled: true },
        ],
      },
      {
        "panel-one": () => "Panel one",
        "panel-two": () => "Panel two",
        "panel-three": () => "Panel three",
      },
    ),
  );
  t.after(cleanup);
  await flush();

  const tablist = document.querySelector('[role="tablist"]');
  const tabs = [...document.querySelectorAll('[role="tab"]')];
  assert.ok(tablist);
  assert.equal(tablist.getAttribute("aria-orientation"), "horizontal");
  assert.equal(tabs.length, 3);
  assert.equal(tabs[0].getAttribute("aria-selected"), "true");
  assert.equal(tabs[0].getAttribute("aria-controls"), bySlot("tabs-content").id);
  assert.equal(tabs[2].disabled, true);

  tabs[0].focus();
  tabs[0].dispatchEvent(
    new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true, cancelable: true }),
  );
  await flush(20);
  assert.equal(document.activeElement, tabs[1]);
  assert.equal(tabs[1].getAttribute("aria-selected"), "true");

  tabs[1].dispatchEvent(
    new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }),
  );
  await flush();
  assert.equal(value.value, "two");
  const visiblePanel = document.querySelector('[role="tabpanel"]:not([hidden])');
  assert.equal(visiblePanel.textContent.trim(), "Panel two");
});

test("Toggle updates aria-pressed and the controlled model", async (t) => {
  const pressed = ref(false);
  const { cleanup } = mountRoot(() =>
    h(
      OfficeToggle,
      {
        modelValue: pressed.value,
        "onUpdate:modelValue": (value) => (pressed.value = value),
        label: "Show grid",
      },
      () => "Grid",
    ),
  );
  t.after(cleanup);
  await flush();

  const toggle = bySlot("toggle");
  assert.equal(toggle.getAttribute("aria-pressed"), "false");
  assert.equal(toggle.getAttribute("data-state"), "off");
  toggle.click();
  await flush();
  assert.equal(pressed.value, true);
  assert.equal(toggle.getAttribute("aria-pressed"), "true");
  assert.equal(toggle.getAttribute("data-state"), "on");
});

// ── Display primitives ────────────────────────────────────────────────────

test("Badge, Card, Separator, Skeleton, Spinner and ScrollArea render their contract", async (t) => {
  const { cleanup } = mountRoot(() =>
    h("div", [
      h(OfficeBadge, { variant: "secondary", size: "lg" }, () => "Draft"),
      h(OfficeCard, { title: "Card title", description: "Card description" }, () => "Body"),
      h(OfficeSeparator),
      h(OfficeSkeleton, { "data-testid": "skeleton" }),
      h(OfficeSpinner, { label: "Loading data" }),
      h(OfficeScrollArea, { orientation: "both" }, () => h("div", "Long content")),
    ]),
  );
  t.after(cleanup);
  await flush();

  const badge = bySlot("badge");
  assert.match(badge.className, /office-badge--secondary/);
  assert.match(badge.className, /office-badge--lg/);

  const card = bySlot("card");
  assert.equal(bySlot("card-header").textContent.includes("Card title"), true);
  assert.equal(bySlot("card-content").textContent.trim(), "Body");

  const separator = bySlot("separator");
  assert.equal(separator.getAttribute("data-orientation"), "horizontal");
  assert.equal(separator.getAttribute("role"), "none");

  assert.equal(document.querySelector('[data-testid="skeleton"]').getAttribute("aria-hidden"), "true");

  const spinner = bySlot("spinner");
  assert.equal(spinner.getAttribute("role"), "status");
  assert.equal(spinner.getAttribute("aria-label"), "Loading data");

  const viewport = bySlot("scroll-area-viewport");
  assert.equal(viewport.getAttribute("tabindex"), "0");
  assert.ok(bySlot("scroll-area"));
});

// ── Toast ─────────────────────────────────────────────────────────────────

test("Toast provider renders status toasts and dismisses them", async (t) => {
  let toastApi;
  const Binder = defineComponent({
    setup() {
      toastApi = useOfficeToast();
      return () => null;
    },
  });
  const { cleanup } = mountRoot(() => h(OfficeToastProvider, null, () => h(Binder)));
  t.after(cleanup);
  await flush();

  assert.ok(bySlot("toast-viewport"));
  assert.equal(document.querySelectorAll('[data-slot="toast"]').length, 0);

  toastApi.success("Saved", "Stored in the workspace.");
  await flush();

  const toast = bySlot("toast");
  assert.ok(toast);
  assert.equal(toast.getAttribute("data-status"), "success");
  assert.equal(bySlot("toast-title").textContent.trim(), "Saved");
  assert.equal(
    bySlot("toast-description").textContent.trim(),
    "Stored in the workspace.",
  );
  assert.ok(bySlot("toast-close"));

  bySlot("toast-close").click();
  await flush(250);
  assert.equal(document.querySelectorAll('[data-slot="toast"]').length, 0);
});

test("useOfficeToast fails fast outside the provider", () => {
  const Outside = defineComponent({
    setup() {
      useOfficeToast();
      return () => null;
    },
  });
  const originalWarn = console.warn;
  console.warn = () => {};
  try {
    assert.throws(
      () => {
        const host = document.createElement("div");
        document.body.append(host);
        createApp({ render: () => h(Outside) }).mount(host);
      },
      /OfficeToastProvider/,
    );
  } finally {
    console.warn = originalWarn;
  }
  document.body.innerHTML = "";
});

// ── SSR ───────────────────────────────────────────────────────────────────

test("every primitive can be server-rendered", async () => {
  const cases = [
    ["OfficeButton", OfficeButton, {}, { default: () => "Button" }, "button"],
    [
      "OfficeIconButton",
      OfficeIconButton,
      { label: "More", icon: "more-horizontal" },
      {},
      "icon-button",
    ],
    ["OfficeInput", OfficeInput, { modelValue: "x" }, {}, "input"],
    ["OfficeTextarea", OfficeTextarea, { modelValue: "x" }, {}, "textarea"],
    [
      "OfficeSelect",
      OfficeSelect,
      { modelValue: "a", options: [{ value: "a", label: "A" }] },
      {},
      "select-trigger",
    ],
    [
      "OfficeDropdownMenu",
      OfficeDropdownMenu,
      { items: [{ label: "Copy" }], label: "Menu" },
      {},
      "dropdown-menu-trigger",
    ],
    [
      "OfficeContextMenu",
      OfficeContextMenu,
      { items: [{ label: "Copy" }] },
      { default: () => h("div", "Target") },
      "context-menu-trigger",
    ],
    [
      "OfficeDialog",
      OfficeDialog,
      { title: "Dialog" },
      { trigger: () => h("button", "Open"), default: () => "Body" },
      "dialog-trigger",
    ],
    [
      "OfficeSheet",
      OfficeSheet,
      { title: "Sheet" },
      { trigger: () => h("button", "Open"), default: () => "Body" },
      "sheet-trigger",
    ],
    [
      "OfficePopover",
      OfficePopover,
      { label: "Popover" },
      { default: () => "Body" },
      "popover-trigger",
    ],
    [
      "OfficeTooltip",
      OfficeTooltip,
      { content: "Tip" },
      { default: () => h("button", "Trigger") },
      "tooltip-trigger",
    ],
    [
      "OfficeTabs",
      OfficeTabs,
      { modelValue: "a", tabs: [{ value: "a", label: "A" }] },
      { "panel-a": () => "Panel" },
      "tabs",
    ],
    ["OfficeToggle", OfficeToggle, { label: "Toggle" }, { default: () => "T" }, "toggle"],
    ["OfficeBadge", OfficeBadge, {}, { default: () => "Badge" }, "badge"],
    ["OfficeCard", OfficeCard, { title: "Card" }, { default: () => "Body" }, "card"],
    ["OfficeSeparator", OfficeSeparator, {}, {}, "separator"],
    ["OfficeSkeleton", OfficeSkeleton, {}, {}, "skeleton"],
    ["OfficeSpinner", OfficeSpinner, {}, {}, "spinner"],
    ["OfficeScrollArea", OfficeScrollArea, {}, { default: () => "Body" }, "scroll-area"],
  ];

  for (const [name, component, props, slots, slot] of cases) {
    const html = await renderToString(h(component, props, slots));
    assert.match(
      html,
      new RegExp(`data-slot="${slot}"`),
      `${name} must server-render its root marker`,
    );
  }

  const toastHtml = await renderToString(
    h(OfficeToastProvider, null, () => h("span", "App")),
  );
  assert.match(toastHtml, /App/);
});
