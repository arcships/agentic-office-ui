// Issue #179 — SchemaBuilder block interaction contract.
//
// Covers the behaviours a static render cannot express: field add/remove/move,
// nested object and array editors, JSON Schema import/export, sibling duplicate
// and empty-name validation, read-only suppression, and the persisted schema
// model helpers exported from `@arcships/vue-ui`.
import assert from "node:assert/strict";
import test, { after } from "node:test";
import { readFileSync } from "node:fs";
import path from "node:path";
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

const { createApp, h, nextTick } = vue;
const {
  SchemaBuilder,
  createSchemaBuilderField,
  findSchemaBuilderField,
  nextSchemaBuilderFieldName,
  parseJsonSchema,
  parseSchemaBuilderJsonText,
  serializeSchemaBuilderSchema,
  validateSchemaBuilderSchema,
  SchemaBuilderParseError,
} = ui;

after(() => {
  GlobalRegistrator.unregister();
});

const ROOT = path.resolve(import.meta.dirname, "../..");

const SCHEMA = {
  properties: [
    {
      id: "field-vendor",
      name: "vendor_name",
      type: "string",
      required: true,
      description: "Supplier on the invoice.",
    },
    { id: "field-total", name: "invoice_total", type: "number" },
    {
      id: "field-remit",
      name: "remit_to",
      type: "object",
      properties: [{ id: "field-city", name: "city", type: "string" }],
    },
    {
      id: "field-lines",
      name: "line_items",
      type: "array",
      itemType: "object",
      properties: [{ id: "field-sku", name: "sku", type: "string", required: true }],
    },
    { id: "field-tags", name: "tags", type: "array", itemType: "string" },
  ],
};

function mount(props = {}) {
  document.body.innerHTML = "";
  const host = document.createElement("div");
  document.body.append(host);
  const warnings = [];
  const app = createApp({ render: () => h(SchemaBuilder, props) });
  app.config.warnHandler = (message) => warnings.push(message);
  app.mount(host);
  const cleanup = async () => {
    app.unmount();
    await nextTick();
    host.remove();
    document.body.innerHTML = "";
  };
  return { app, host, warnings, cleanup };
}

async function flush(ms = 30) {
  await nextTick();
  await new Promise((resolve) => setTimeout(resolve, ms));
  await nextTick();
}

const bySlot = (slot, root = document) => root.querySelector(`[data-slot="${slot}"]`);
const allBySlot = (slot, root = document) => [...root.querySelectorAll(`[data-slot="${slot}"]`)];
const fieldRows = () => allBySlot("schema-builder-field");
const fieldRow = (id) =>
  document.querySelector(`[data-slot="schema-builder-field"][data-field-id="${id}"]`);
const nameInput = (row) => row.querySelector('input:not([type="checkbox"])');
const typeTrigger = (row) => row.querySelector('[role="combobox"]');
const descriptionInput = (row) =>
  [...row.querySelectorAll('input:not([type="checkbox"])')][1];
const requiredBox = (row) => row.querySelector('input[type="checkbox"]');

function click(target) {
  target.dispatchEvent(
    new PointerEvent("pointerdown", { bubbles: true, button: 0 }),
  );
  target.dispatchEvent(
    new PointerEvent("pointerup", { bubbles: true, button: 0 }),
  );
  target.click();
}

function openSelect(trigger) {
  trigger.dispatchEvent(
    new PointerEvent("pointerdown", { bubbles: true, button: 0 }),
  );
  trigger.click();
}

function typeInto(target, value) {
  target.value = value;
  target.dispatchEvent(new Event("input", { bubbles: true }));
}

async function selectOption(trigger, label) {
  openSelect(trigger);
  await flush();
  const option = [...document.querySelectorAll('[data-slot="select-item"]')].find(
    (node) => node.textContent.trim() === label,
  );
  assert.ok(
    option,
    `option ${label} is rendered in ${[...document.querySelectorAll('[data-slot="select-item"]')].map((node) => node.textContent.trim()).join(", ")}`,
  );
  click(option);
  await flush();
}

// Pure model -------------------------------------------------------------

test("schema model serializes, imports and validates JSON Schema documents", () => {
  const serialized = serializeSchemaBuilderSchema(SCHEMA);
  assert.deepEqual(serialized.required, ["vendor_name"]);
  assert.deepEqual(serialized.properties.remit_to, {
    type: "object",
    properties: { city: { type: "string" } },
  });
  assert.deepEqual(serialized.properties.line_items, {
    type: "array",
    items: { type: "object", properties: { sku: { type: "string" } }, required: ["sku"] },
  });
  assert.deepEqual(serialized.properties.tags, { type: "array", items: { type: "string" } });

  const roundTrip = parseJsonSchema(serialized);
  assert.deepEqual(roundTrip.warnings, []);
  const reparsed = serializeSchemaBuilderSchema(roundTrip.schema);
  assert.deepEqual(reparsed, serialized, "serialize → parse → serialize is stable");

  const tolerant = parseJsonSchema({
    type: "object",
    properties: {
      count: { type: "integer" },
      nested: { type: "array", items: { type: "array" } },
    },
  });
  assert.equal(tolerant.schema.properties[0].type, "number");
  assert.equal(tolerant.schema.properties[1].itemType, "string");
  assert.deepEqual(
    tolerant.warnings.map((warning) => warning.code),
    ["unsupported-type", "unsupported-item-type"],
  );

  assert.throws(() => parseJsonSchema("nope"), (error) => {
    return error instanceof SchemaBuilderParseError && error.code === "invalid-root";
  });
  assert.throws(() => parseJsonSchema({ type: "object", properties: [] }), (error) => {
    return error.code === "invalid-properties";
  });

  const issues = validateSchemaBuilderSchema({
    properties: [
      { id: "a", name: "", type: "string" },
      { id: "b", name: "same", type: "string" },
      { id: "c", name: "same", type: "number" },
      { id: "d", name: "weird", type: "date" },
    ],
  });
  assert.deepEqual(
    issues.map((issue) => [issue.code, issue.fieldId]),
    [
      ["empty-name", "a"],
      ["duplicate-name", "c"],
      ["invalid-type", "d"],
    ],
  );

  const field = createSchemaBuilderField("x", "object");
  assert.equal(field.name, "");
  assert.equal(
    nextSchemaBuilderFieldName([{ id: "1", name: "field_1", type: "string" }]),
    "field_2",
  );

  const located = findSchemaBuilderField(roundTrip.schema, roundTrip.schema.properties[2].id);
  assert.equal(located.path, "remit_to");
  assert.equal(parseSchemaBuilderJsonText('{"type":"object","properties":{}}').schema.properties.length, 0);
});

// Rendering --------------------------------------------------------------

test("SchemaBuilder renders typed rows with nested object and array children", async () => {
  const mounted = mount({ schema: structuredClone(SCHEMA) });
  await flush();

  assert.equal(fieldRows().length, 7, "five root fields plus two nested children");
  const vendor = fieldRow("field-vendor");
  assert.equal(nameInput(vendor).value, "vendor_name");
  assert.equal(descriptionInput(vendor).value, "Supplier on the invoice.");
  assert.equal(requiredBox(vendor).checked, true);
  assert.equal(typeTrigger(vendor).textContent.trim(), "String");

  const remit = fieldRow("field-remit");
  assert.equal(typeTrigger(remit).textContent.trim(), "Object");
  assert.ok(remit.querySelector('[data-slot="schema-builder-field-children"]'));
  assert.ok(fieldRow("field-city"));

  const lines = fieldRow("field-lines");
  assert.equal(typeTrigger(lines).textContent.trim(), "Array");
  assert.ok(fieldRow("field-sku"));
  assert.ok(lines.querySelector('[data-slot="schema-builder-field-array"]'));

  const tags = fieldRow("field-tags");
  assert.equal(tags.querySelector('[data-slot="schema-builder-field-children"]'), null);
  assert.ok(tags.querySelector('[data-slot="schema-builder-field-array"]'));

  assert.deepEqual(mounted.warnings, []);
  await mounted.cleanup();
});

test("SchemaBuilder edits names, descriptions and the required flag", async () => {
  const changes = [];
  const schema = structuredClone(SCHEMA);
  const mounted = mount({
    schema,
    "onUpdate:schema": (next) => changes.push(next),
    "onSchema-change": (next) => (schema.properties = next.properties),
  });
  await flush();

  typeInto(nameInput(fieldRow("field-total")), "grand_total");
  await flush();
  assert.equal(changes.at(-1).properties[1].name, "grand_total");

  typeInto(descriptionInput(fieldRow("field-total")), "Total due");
  await flush();
  assert.equal(changes.at(-1).properties[1].description, "Total due");

  const paid = requiredBox(fieldRow("field-total"));
  assert.equal(paid.checked, false);
  paid.checked = true;
  paid.dispatchEvent(new Event("change", { bubbles: true }));
  await flush();
  assert.equal(changes.at(-1).properties[1].required, true);
  assert.equal(changes.at(-1).properties[0].required, true, "other fields keep their flags");

  assert.ok(changes.length >= 3);
  await mounted.cleanup();
});

test("SchemaBuilder changes a field type and creates nested containers", async () => {
  const changes = [];
  const mounted = mount({
    schema: structuredClone(SCHEMA),
    "onUpdate:schema": (next) => changes.push(next),
  });
  await flush();

  await selectOption(typeTrigger(fieldRow("field-total")), "Object");
  await flush();
  const total = changes.at(-1).properties[1];
  assert.equal(total.type, "object");
  assert.deepEqual(total.properties, []);
  assert.ok(fieldRow("field-total").querySelector('[data-slot="schema-builder-field-children"]'));

  await selectOption(typeTrigger(fieldRow("field-tags")), "Array");
  await flush();
  await mounted.cleanup();
});

test("SchemaBuilder adds root and nested fields with unique names", async () => {
  const changes = [];
  const mounted = mount({
    schema: structuredClone(SCHEMA),
    "onUpdate:schema": (next) => changes.push(next),
  });
  await flush();

  click(bySlot("schema-builder-add"));
  await flush();
  const added = changes.at(-1).properties.at(-1);
  assert.equal(added.name, "field_1");
  assert.equal(added.type, "string");
  assert.equal(fieldRows().length, 8);

  click(bySlot("schema-builder-add"));
  await flush();
  assert.equal(changes.at(-1).properties.at(-1).name, "field_2");

  const remit = fieldRow("field-remit");
  click(remit.querySelector('[data-slot="schema-builder-field-add-child"]'));
  await flush();
  const remitFields = changes.at(-1).properties[2].properties;
  assert.equal(remitFields.length, 2);
  assert.equal(remitFields[1].name, "field_1");
  assert.equal(fieldRow("field-remit").querySelectorAll('[data-slot="schema-builder-field"]').length, 2);
  await mounted.cleanup();
});

test("SchemaBuilder removes and reorders sibling fields", async () => {
  const changes = [];
  const mounted = mount({
    schema: structuredClone(SCHEMA),
    "onUpdate:schema": (next) => changes.push(next),
  });
  await flush();

  click(fieldRow("field-vendor").querySelector('[data-slot="schema-builder-field-move-down"]'));
  await flush();
  assert.deepEqual(
    changes.at(-1).properties.slice(0, 2).map((field) => field.id),
    ["field-total", "field-vendor"],
  );
  assert.deepEqual(
    fieldRows().slice(0, 2).map((row) => row.getAttribute("data-field-id")),
    ["field-total", "field-vendor"],
  );

  click(fieldRow("field-vendor").querySelector('[data-slot="schema-builder-field-move-up"]'));
  await flush();
  assert.deepEqual(
    changes.at(-1).properties.slice(0, 2).map((field) => field.id),
    ["field-vendor", "field-total"],
  );

  const firstMoveUp = fieldRow("field-vendor").querySelector(
    '[data-slot="schema-builder-field-move-up"]',
  );
  assert.equal(firstMoveUp.disabled, true, "the first sibling cannot move up");

  click(fieldRow("field-vendor").querySelector('[data-slot="schema-builder-field-remove"]'));
  await flush();
  assert.equal(changes.at(-1).properties.length, 4);
  assert.equal(fieldRow("field-vendor"), null);
  await mounted.cleanup();
});

// Validation -------------------------------------------------------------

test("SchemaBuilder reports empty names, duplicates and invalid types", async () => {
  const reported = [];
  const mounted = mount({
    schema: {
      properties: [
        { id: "empty", name: "  ", type: "string" },
        { id: "dup-1", name: "same", type: "string" },
        { id: "dup-2", name: "same", type: "number" },
        { id: "bad", name: "when", type: "date" },
      ],
    },
    "onValidation-change": (issues) => reported.push(issues),
  });
  await flush();

  const summary = bySlot("schema-builder-issues");
  assert.equal(summary.getAttribute("role"), "alert");
  assert.equal(summary.querySelectorAll("li").length, 3);
  assert.deepEqual(
    [...summary.querySelectorAll("li")].map((item) => item.getAttribute("data-issue-code")),
    ["empty-name", "duplicate-name", "invalid-type"],
  );

  assert.match(
    fieldRow("dup-2").querySelector('[data-slot="schema-builder-field-issue"]').textContent,
    /Duplicate field name: same/,
  );
  assert.match(
    fieldRow("empty").querySelector('[data-slot="schema-builder-field-issue"]').textContent,
    /cannot be empty/,
  );
  assert.equal(
    nameInput(fieldRow("dup-2")).getAttribute("aria-invalid"),
    "true",
    "duplicate names mark the name input invalid",
  );
  assert.equal(fieldRow("dup-2").className.includes("is-invalid"), true);

  assert.deepEqual(
    reported.at(-1).map((issue) => issue.code),
    ["empty-name", "duplicate-name", "invalid-type"],
  );
  await mounted.cleanup();
});

// JSON import / export ---------------------------------------------------

test("SchemaBuilder exports JSON Schema and emits the copy payload", async () => {
  const exported = [];
  const mounted = mount({
    schema: structuredClone(SCHEMA),
    "onExport": (value) => exported.push(value),
  });
  await flush();

  assert.equal(bySlot("schema-builder-json"), null);
  click(bySlot("schema-builder-json-toggle"));
  await flush();

  const pre = bySlot("schema-builder-json");
  assert.ok(pre);
  const text = pre.textContent;
  const parsed = JSON.parse(text);
  assert.equal(parsed.type, "object");
  assert.deepEqual(parsed.required, ["vendor_name"]);
  assert.deepEqual(parsed.properties.line_items.items, {
    type: "object",
    properties: { sku: { type: "string" } },
    required: ["sku"],
  });

  click(bySlot("schema-builder-export"));
  await flush();
  assert.equal(exported.length, 1);
  assert.deepEqual(exported[0], parsed);

  click(bySlot("schema-builder-json-toggle"));
  await flush();
  assert.equal(bySlot("schema-builder-json"), null);
  await mounted.cleanup();
});

test("SchemaBuilder imports a JSON Schema document and reports warnings", async () => {
  const changes = [];
  const mounted = mount({
    schema: structuredClone(SCHEMA),
    "onUpdate:schema": (next) => changes.push(next),
  });
  await flush();

  click(bySlot("schema-builder-import-toggle"));
  await flush();
  const textarea = bySlot("schema-builder-import").querySelector("textarea");
  assert.ok(textarea);

  typeInto(
    textarea,
    JSON.stringify({
      type: "object",
      required: ["purchase_order"],
      properties: {
        purchase_order: { type: "string", description: "PO number" },
        terms: {
          type: "object",
          properties: { net_days: { type: "integer" } },
          required: ["net_days"],
        },
        approvals: {
          type: "array",
          items: { type: "object", properties: { approver: { type: "string" } } },
        },
      },
    }),
  );
  click(bySlot("schema-builder-import-apply"));
  await flush();

  const imported = changes.at(-1);
  assert.deepEqual(imported.properties.map((field) => field.name), [
    "purchase_order",
    "terms",
    "approvals",
  ]);
  assert.equal(imported.properties[0].required, true);
  assert.equal(imported.properties[1].properties[0].type, "number");
  assert.equal(imported.properties[2].itemType, "object");

  const warnings = bySlot("schema-builder-import-warnings");
  assert.ok(warnings, "unsupported `integer` is reported instead of dropped");
  assert.match(warnings.textContent, /purchase_order|terms|capital/);
  assert.match(warnings.textContent, /not supported/);
  await mounted.cleanup();
});

test("SchemaBuilder rejects invalid JSON and keeps the current schema", async () => {
  const changes = [];
  const mounted = mount({
    schema: structuredClone(SCHEMA),
    "onUpdate:schema": (next) => changes.push(next),
  });
  await flush();

  click(bySlot("schema-builder-import-toggle"));
  await flush();
  let textarea = bySlot("schema-builder-import").querySelector("textarea");

  click(bySlot("schema-builder-import-apply"));
  await flush();
  assert.match(bySlot("schema-builder-import-error").textContent, /Paste a JSON Schema object first/);
  assert.equal(changes.length, 0);

  typeInto(textarea, "{ not json");
  click(bySlot("schema-builder-import-apply"));
  await flush();
  const error = bySlot("schema-builder-import-error");
  assert.equal(error.getAttribute("role"), "alert");
  assert.match(error.textContent, /Not valid JSON/);
  assert.equal(changes.length, 0);

  typeInto(textarea, JSON.stringify({ type: "object", properties: [] }));
  click(bySlot("schema-builder-import-apply"));
  await flush();
  assert.match(bySlot("schema-builder-import-error").textContent, /properties map/);

  typeInto(textarea, JSON.stringify({ type: "array", properties: {} }));
  click(bySlot("schema-builder-import-apply"));
  await flush();
  assert.match(bySlot("schema-builder-import-error").textContent, /properties map/);

  click(bySlot("schema-builder-import-cancel"));
  await flush();
  assert.equal(bySlot("schema-builder-import"), null, "cancel closes the panel");

  click(bySlot("schema-builder-import-toggle"));
  await flush();
  textarea = bySlot("schema-builder-import").querySelector("textarea");
  typeInto(textarea, "{}");
  click(bySlot("schema-builder-import-apply"));
  await flush();
  assert.equal(changes.at(-1).properties.length, 0, "an empty document clears the builder");
  assert.equal(bySlot("schema-builder-import"), null, "a clean import closes the panel");
  await mounted.cleanup();
});

// Read-only --------------------------------------------------------------

test("SchemaBuilder read-only disables every mutation but keeps JSON export", async () => {
  const changes = [];
  const mounted = mount({
    schema: structuredClone(SCHEMA),
    readOnly: true,
    "onUpdate:schema": (next) => changes.push(next),
  });
  await flush();

  assert.equal(bySlot("schema-builder").getAttribute("data-read-only"), "true");
  assert.equal(bySlot("schema-builder-add"), null);
  assert.equal(bySlot("schema-builder-import-toggle"), null);
  assert.equal(allBySlot("schema-builder-field-remove").length, 0);
  assert.equal(allBySlot("schema-builder-field-move-up").length, 0);
  assert.equal(allBySlot("schema-builder-field-add-child").length, 0);

  const vendor = fieldRow("field-vendor");
  assert.equal(nameInput(vendor).readOnly, true);
  assert.equal(descriptionInput(vendor).readOnly, true);
  assert.equal(requiredBox(vendor).disabled, true);
  assert.equal(typeTrigger(vendor).disabled, true);

  click(bySlot("schema-builder-json-toggle"));
  await flush();
  assert.ok(bySlot("schema-builder-json"));
  click(bySlot("schema-builder-export"));
  await flush();
  assert.equal(changes.length, 0);
  await mounted.cleanup();
});

// States ----------------------------------------------------------------

test("SchemaBuilder empty, loading and error surfaces expose live status semantics", async () => {
  const empty = mount({ schema: { properties: [] } });
  await flush();
  assert.equal(bySlot("schema-builder-empty").getAttribute("role"), "status");
  assert.match(bySlot("schema-builder-empty").textContent, /No fields yet/);
  await empty.cleanup();

  const custom = mount({ schema: { properties: [] }, emptyMessage: "Start from scratch." });
  await flush();
  assert.match(bySlot("schema-builder-empty").textContent, /Start from scratch/);
  await custom.cleanup();

  const loading = mount({ schema: { properties: [] }, loading: true });
  await flush();
  assert.equal(bySlot("schema-builder-loading").getAttribute("role"), "status");
  assert.equal(bySlot("schema-builder-empty"), null);
  await loading.cleanup();

  const error = mount({ schema: structuredClone(SCHEMA), error: "The schema failed to load." });
  await flush();
  assert.equal(bySlot("schema-builder-error").getAttribute("role"), "alert");
  assert.match(bySlot("schema-builder-error").textContent, /failed to load/);
  assert.equal(bySlot("schema-builder-fields"), null);
  await error.cleanup();
});

test("SchemaBuilder stays dependency-free and documents its responsive breakpoints", () => {
  const source = readFileSync(
    path.join(ROOT, "packages/vue-ui/src/components/SchemaBuilder.vue"),
    "utf8",
  );
  const row = readFileSync(
    path.join(ROOT, "packages/vue-ui/src/components/SchemaBuilderFieldRow.vue"),
    "utf8",
  );
  assert.doesNotMatch(source + row, /dnd-kit|@pierre\/diffs/u);
  assert.ok(source.includes("640px"));
  assert.ok(row.includes("640px"));

  const manifest = JSON.parse(
    readFileSync(path.join(ROOT, "packages/vue-ui/package.json"), "utf8"),
  );
  const dependencies = {
    ...manifest.dependencies,
    ...manifest.devDependencies,
    ...manifest.peerDependencies,
  };
  assert.equal(dependencies["@dnd-kit/core"], undefined);
  assert.equal(dependencies["@pierre/diffs"], undefined);
});
