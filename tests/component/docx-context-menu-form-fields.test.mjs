import assert from "node:assert/strict";
import test from "node:test";

import {
  importFromDemo,
  mount,
  teleportTarget,
  textContent,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

const { DocxContextMenu, DocxFormFieldLayer, useDocxEditor } = await importFromDemo(
  "@arcships/vue-docx",
);
const { createBlankDocumentModel } = await importFromDemo("@arcships/docx-core");

function modelWithNodes(nodes) {
  return { ...createBlankDocumentModel(), nodes };
}

function hasClass(node, className) {
  return String(node?.props?.class || "").split(/\s+/).includes(className);
}

function menuElement() {
  return walk(teleportTarget("body")).find(
    (node) => node.props?.role === "menu" && hasClass(node, "docx-context-menu"),
  );
}

function menuItems() {
  const menu = menuElement();
  return menu ? walk(menu).filter((node) => node.props?.role === "menuitem") : [];
}

async function mountContextMenu() {
  const context = vue.ref(undefined);
  const closes = [];
  const Harness = vue.defineComponent({
    setup() {
      return () =>
        vue.h(DocxContextMenu, {
          context: context.value,
          controller: { selectedRunStyle: {} },
          onClose: () => {
            closes.push("close");
            context.value = undefined;
          },
        });
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, context, closes };
}

function spyOnItems(items) {
  const focused = [];
  for (const item of items) {
    item.focus = () => focused.push(item);
    item.setAttribute = (name, value) => {
      item.props[name] = value;
    };
  }
  return focused;
}

function keydown(menu, key) {
  const prevented = [];
  const stopped = [];
  menu.props.onKeydown({
    key,
    preventDefault: () => prevented.push(key),
    stopPropagation: () => stopped.push(key),
  });
  return { prevented, stopped };
}

test("DocxContextMenu anchors at pointer coordinates and exposes menu semantics", async (t) => {
  const mounted = await mountContextMenu();
  t.after(() => mounted.app.unmount());
  mounted.context.value = { kind: "text", position: { x: 120, y: 90 } };
  await waitFor(menuElement, "context menu to render");

  const menu = menuElement();
  assert.equal(menu.props.role, "menu");
  assert.equal(menu.props["aria-orientation"], "vertical");
  assert.equal(menu.props.tabindex, "-1");
  assert.equal(menu.props.style.position, "fixed");
  assert.equal(menu.props.style.left, "120px");
  assert.equal(menu.props.style.top, "90px");

  const items = menuItems();
  assert.ok(items.length > 10, "all built-in actions are menu items");
  assert.ok(items.every((item) => item.props.role === "menuitem"));
  assert.ok(items.every((item) => item.props.tabindex === "-1"));
  assert.ok(
    walk(menu).some((node) => node.props?.role === "separator"),
    "separators carry separator semantics",
  );

  assert.deepEqual(mounted.warnings, []);
});

test("DocxContextMenu focuses the first item and roves focus with arrows, Home and End", async (t) => {
  const mounted = await mountContextMenu();
  t.after(() => mounted.app.unmount());
  mounted.context.value = { kind: "text", position: { x: 40, y: 40 } };
  await waitFor(menuElement, "context menu to render");

  const items = menuItems();
  const focused = spyOnItems(items);
  // A fresh context is how the menu reopens on another click.
  mounted.context.value = { kind: "text", position: { x: 44, y: 46 } };
  await waitFor(() => focused.length > 0, "first menu item to receive focus");
  assert.equal(focused.at(-1), items[0]);
  assert.equal(items[0].props.tabindex, "0");
  assert.ok(items.slice(1).every((item) => item.props.tabindex === "-1"));
  assert.equal(menuElement().props.style.left, "44px");
  assert.equal(menuElement().props.style.top, "46px");

  keydown(menuElement(), "ArrowDown");
  assert.equal(focused.at(-1), items[1]);

  keydown(menuElement(), "End");
  assert.equal(focused.at(-1), items.at(-1));

  keydown(menuElement(), "ArrowDown");
  assert.equal(focused.at(-1), items[0], "ArrowDown wraps to the first item");

  keydown(menuElement(), "ArrowUp");
  assert.equal(focused.at(-1), items.at(-1), "ArrowUp wraps to the last item");

  keydown(menuElement(), "Home");
  assert.equal(focused.at(-1), items[0]);
  assert.equal(items[0].props.tabindex, "0");
  assert.ok(items.slice(1).every((item) => item.props.tabindex === "-1"));

  const escape = keydown(menuElement(), "Escape");
  assert.deepEqual(escape.prevented, ["Escape"]);
  assert.deepEqual(escape.stopped, ["Escape"]);
  assert.deepEqual(mounted.closes, ["close"]);
  await vue.nextTick();
  assert.equal(menuElement(), undefined, "Escape closes the menu");

  assert.deepEqual(mounted.warnings, []);
});

test("DocxContextMenu clamps using the measured menu box, not a constant height", async (t) => {
  const originalWindow = globalThis.window;
  globalThis.window = { innerWidth: 800, innerHeight: 600 };
  t.after(() => {
    if (originalWindow === undefined) delete globalThis.window;
    else globalThis.window = originalWindow;
  });

  const mounted = await mountContextMenu();
  t.after(() => mounted.app.unmount());
  mounted.context.value = { kind: "text", position: { x: 780, y: 580 } };
  await waitFor(menuElement, "context menu to render");

  let menu = menuElement();
  // Fallback box: 240 x 420, viewport gap 8.
  assert.equal(menu.props.style.left, "552px");
  assert.equal(menu.props.style.top, "172px");

  menu.offsetWidth = 200;
  menu.offsetHeight = 300;
  mounted.context.value = { kind: "text", position: { x: 780, y: 580 } };
  await waitFor(
    () => menuElement()?.props.style.left === "592px",
    "menu position to clamp against the measured box",
  );
  menu = menuElement();
  assert.equal(menu.props.style.top, "292px");

  assert.deepEqual(mounted.warnings, []);
});

async function mountFormFieldLayer(model, pageNodeSegments) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model });
      return () =>
        vue.h(DocxFormFieldLayer, {
          pageIndex: 0,
          controller,
          pageNodeSegments,
        });
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, get controller() { return controller; } };
}

function fieldNodes(root) {
  return walk(root).filter((node) => node.props?.["data-docx-form-field"] === true);
}

function dropdownTableModel() {
  return modelWithNodes([
    {
      type: "table",
      rows: [
        {
          cells: [{
            nodes: [{
              type: "paragraph",
              children: [{
                type: "form-field",
                fieldType: "dropdown",
                options: [
                  { displayText: "Alpha" },
                  { displayText: "Beta", value: "beta" },
                ],
              }],
            }],
          }],
        },
        {
          cells: [{
            nodes: [{
              type: "paragraph",
              children: [{
                type: "form-field",
                fieldType: "text",
                value: "typed",
              }],
            }],
          }],
        },
      ],
    },
    {
      type: "paragraph",
      children: [{
        type: "form-field",
        fieldType: "checkbox",
        checked: true,
        title: "Accept",
      }],
    },
  ]);
}

test("DocxFormFieldLayer collects table-cell fields and renders dropdown displayText values", async () => {
  const model = dropdownTableModel();
  const mounted = await mountFormFieldLayer(model, [
    { nodeIndex: 0 },
    { nodeIndex: 1 },
  ]);

  const fields = fieldNodes(mounted.root);
  assert.deepEqual(
    fields.map((field) => field.props["data-docx-form-field-location"]),
    ["t:0:0:0:0:0", "t:0:1:0:0:0", "p:1:0"],
    "table-cell fields are collected with table locations, in document order",
  );

  const select = walk(mounted.root).find((node) => node.type === "select");
  assert.ok(select, "the dropdown inside the table renders as a select");
  const options = select.options;
  assert.deepEqual(
    options.map((option) => textContent(option).trim()),
    ["Alpha", "Beta"],
    "option labels come from displayText",
  );
  assert.deepEqual(
    options.map((option) => option.props.value),
    ["Alpha", "beta"],
    "option values fall back to displayText when the model has no value",
  );

  const input = walk(mounted.root).find((node) => node.type === "input");
  assert.ok(input, "text fields inside table cells render too");
  assert.equal(input.props.value, "typed");

  select.props.onChange({ target: { value: "beta" } });
  await vue.nextTick();
  const updated = mounted.controller.model.nodes[0];
  assert.equal(updated.type, "table");
  assert.equal(updated.rows[0].cells[0].nodes[0].children[0].value, "beta");

  select.props.onFocus();
  await vue.nextTick();
  assert.deepEqual(mounted.controller.selectedFormField.location, {
    kind: "table-cell",
    tableIndex: 0,
    rowIndex: 0,
    cellIndex: 0,
    paragraphIndex: 0,
    childIndex: 0,
  });
  const selectedField = fieldNodes(mounted.root).find(
    (field) => field.props["data-docx-form-field-location"] === "t:0:0:0:0:0",
  );
  assert.equal(selectedField.props.style.outline, "2px solid #3b82f6");

  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DocxFormFieldLayer only renders table fields whose rows are on the page", async () => {
  const firstPage = await mountFormFieldLayer(dropdownTableModel(), [
    { nodeIndex: 0, tableRowRange: { startRowIndex: 0, endRowIndex: 1 } },
    { nodeIndex: 1 },
  ]);
  assert.equal(walk(firstPage.root).filter((node) => node.type === "select").length, 1);
  assert.equal(
    walk(firstPage.root).filter((node) => node.type === "input" && node.props.type === "text").length,
    0,
    "the text field lives in a row rendered on the next page",
  );
  assert.deepEqual(firstPage.warnings, []);
  firstPage.app.unmount();

  const secondPage = await mountFormFieldLayer(dropdownTableModel(), [
    { nodeIndex: 0, tableRowRange: { startRowIndex: 1, endRowIndex: 2 } },
    { nodeIndex: 1 },
  ]);
  assert.equal(walk(secondPage.root).filter((node) => node.type === "select").length, 0);
  assert.equal(
    walk(secondPage.root).filter((node) => node.type === "input" && node.props.type === "text").length,
    1,
  );
  assert.deepEqual(
    fieldNodes(secondPage.root).map((field) => field.props["data-docx-form-field-location"]),
    ["t:0:1:0:0:0", "p:1:0"],
  );
  assert.deepEqual(secondPage.warnings, []);
  secondPage.app.unmount();
});
