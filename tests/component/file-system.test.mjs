// Issue #179 — FileSystem block interaction contract.
//
// happy-dom is required because the block relies on real focus, checkbox and
// double-click semantics that the repository's minimal host renderer does not
// model. Layout/colour are not asserted: the block has no pixel-parity
// surface, and the exact deltas against the upstream Finder are recorded in
// docs/components/file-system.md.
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

const { createApp, h, nextTick, ref } = vue;
const {
  FileSystem,
  buildFileSystemIndex,
  fileSystemAncestors,
  fileSystemEntries,
  fileSystemPreviewUrl,
  formatFileSystemSize,
} = ui;

after(() => {
  GlobalRegistrator.unregister();
});

const ROOT = path.resolve(import.meta.dirname, "../..");

// A 1x1 transparent PNG. happy-dom decodes data URLs; a relative demo asset
// would fail to load and FileThumbnail would drop the <img> after `error`.
const PIXEL_PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

const ITEMS = [
  {
    id: "invoices",
    name: "Invoices",
    kind: "folder",
    children: [
      { id: "invoice-pdf", name: "invoice-0881.pdf", kind: "file", size: 38_912, mime: "application/pdf" },
      {
        id: "invoice-scan",
        name: "invoice-scan.png",
        kind: "file",
        size: 286_720,
        mime: "image/png",
        thumbnailUrl: PIXEL_PNG,
      },
    ],
  },
  { id: "archive", name: "Archive", kind: "folder", children: [] },
  { id: "beta", name: "beta-notes.txt", kind: "file", size: 2048, mime: "text/plain", updatedAt: "2026-11-18T07:20:00.000Z" },
  { id: "alpha", name: "alpha-contract.pdf", kind: "file", size: 1_310_720, mime: "application/pdf" },
];

function mount(component, props = {}) {
  // A failed assertion must not leave a previous app mounted: every lookup
  // below uses `document.querySelector`, so stale rows would win.
  document.body.innerHTML = "";
  const host = document.createElement("div");
  document.body.append(host);
  const warnings = [];
  const app = createApp({ render: () => h(component, props) });
  app.config.warnHandler = (message) => warnings.push(message);
  app.mount(host);
  const cleanup = () => {
    app.unmount();
    host.remove();
    document.body.innerHTML = "";
  };
  return { app, host, warnings, cleanup };
}

async function flush(ms = 20) {
  await nextTick();
  await new Promise((resolve) => setTimeout(resolve, ms));
  await nextTick();
}

const bySlot = (slot, root = document) => root.querySelector(`[data-slot="${slot}"]`);
const allBySlot = (slot, root = document) => [...root.querySelectorAll(`[data-slot="${slot}"]`)];
const itemRows = () => allBySlot("file-system-item");
const rowFor = (id) => document.querySelector(`[data-slot="file-system-item"][data-item-id="${id}"]`);

function mainButton(row) {
  return row.querySelector('[data-slot="file-system-item-main"]');
}

function click(target) {
  target.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
}

function doubleClick(target) {
  target.dispatchEvent(new MouseEvent("dblclick", { bubbles: true, cancelable: true }));
}

function keydown(target, key, init = {}) {
  target.dispatchEvent(
    new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true, ...init }),
  );
}

// Pure model -------------------------------------------------------------

test("file system index resolves children, breadcrumbs and formatting", () => {
  const index = buildFileSystemIndex(ITEMS);
  assert.deepEqual(
    fileSystemEntries(index, null).map((node) => node.item.id),
    ["archive", "invoices", "alpha", "beta"],
    "folders sort first, then files by natural name order",
  );
  const nested = fileSystemEntries(index, "invoices").map((node) => node.item.id);
  assert.deepEqual(nested, ["invoice-pdf", "invoice-scan"]);

  const ancestors = fileSystemAncestors(index, "invoices");
  assert.deepEqual(ancestors.map((node) => node.path), ["Invoices"]);

  assert.equal(formatFileSystemSize(2048), "2 KB");
  assert.equal(formatFileSystemSize(1_310_720), "1.3 MB");
  assert.equal(formatFileSystemSize(undefined), null);
  assert.equal(fileSystemPreviewUrl(index.nodes.get("invoice-scan").item), PIXEL_PNG);
  assert.equal(fileSystemPreviewUrl(index.nodes.get("invoice-pdf").item), null);

  const duplicated = buildFileSystemIndex([
    { id: "dup", name: "one", kind: "file" },
    { id: "dup", name: "two", kind: "file" },
  ]);
  assert.deepEqual(duplicated.duplicateIds, ["dup"]);
  assert.equal(duplicated.nodes.get("dup").item.name, "one");
});

// Rendering --------------------------------------------------------------

test("FileSystem renders the root folder, breadcrumb and grid view", async () => {
  const mounted = mount(FileSystem, { items: ITEMS, title: "Workspace" });
  await flush();

  assert.equal(bySlot("file-system").getAttribute("role"), "group");
  assert.equal(bySlot("file-system").getAttribute("aria-label"), "Workspace");
  assert.equal(bySlot("file-system").getAttribute("data-view"), "grid");
  assert.equal(bySlot("file-system").getAttribute("data-current-folder-id"), "");

  const crumbs = allBySlot("file-system-crumb");
  assert.equal(crumbs.length, 1);
  assert.equal(crumbs[0].textContent.trim(), "Workspace");
  assert.equal(crumbs[0].getAttribute("aria-current"), "page");

  assert.ok(bySlot("file-system-grid"), "grid view is the default");
  assert.equal(bySlot("file-system-list"), null);
  assert.deepEqual(
    itemRows().map((row) => row.getAttribute("data-item-id")),
    ["archive", "invoices", "alpha", "beta"],
  );

  const rootList = bySlot("file-system-grid");
  assert.equal(rootList.getAttribute("role"), "list");
  assert.equal(itemRows()[0].getAttribute("role"), "listitem");

  assert.equal(rowFor("invoice-pdf"), null, "nested entries are not rendered at the root");
  assert.deepEqual(mounted.warnings, []);
  mounted.cleanup();
});

test("FileSystem toggles list and grid views and reports the change", async () => {
  const views = [];
  const mounted = mount(FileSystem, {
    items: ITEMS,
    view: undefined,
    "onUpdate:view": (value) => views.push(value),
  });
  await flush();

  const listButton = bySlot("file-system-view-list");
  const gridButton = bySlot("file-system-view-grid");
  assert.equal(listButton.getAttribute("aria-pressed"), "false");
  assert.equal(gridButton.getAttribute("aria-pressed"), "true");

  click(listButton);
  await flush();

  assert.deepEqual(views, ["list"]);
  assert.ok(bySlot("file-system-list"), "list view renders after the toggle");
  assert.equal(bySlot("file-system-grid"), null);
  assert.equal(bySlot("file-system").getAttribute("data-view"), "list");
  assert.equal(bySlot("file-system-view-list").getAttribute("aria-pressed"), "true");

  click(bySlot("file-system-view-grid"));
  await flush();
  assert.deepEqual(views, ["list", "grid"]);
  mounted.cleanup();
});

test("FileSystem selects on click and opens folders into breadcrumb navigation", async () => {
  const selections = [];
  const selectedIds = [];
  const folders = [];
  const mounted = mount(FileSystem, {
    items: ITEMS,
    "onSelection-change": (item) => selections.push(item?.id ?? null),
    "onUpdate:selectedId": (id) => selectedIds.push(id),
    "onOpen-folder": (item) => folders.push(item.id),
  });
  await flush();

  const folderButton = mainButton(rowFor("invoices"));
  click(folderButton);
  await flush();

  assert.deepEqual(selections, ["invoices"]);
  assert.deepEqual(selectedIds, ["invoices"]);
  assert.equal(mainButton(rowFor("invoices")).getAttribute("aria-current"), "true");
  assert.equal(mainButton(rowFor("archive")).getAttribute("aria-current"), null);

  doubleClick(folderButton);
  await flush();

  assert.deepEqual(folders, ["invoices"]);
  const crumbs = allBySlot("file-system-crumb");
  assert.deepEqual(crumbs.map((crumb) => crumb.textContent.trim()), ["Files", "Invoices"]);
  assert.equal(crumbs[1].getAttribute("aria-current"), "page");
  assert.equal(bySlot("file-system").getAttribute("data-current-folder-id"), "invoices");
  assert.deepEqual(
    itemRows().map((row) => row.getAttribute("data-item-id")),
    ["invoice-pdf", "invoice-scan"],
  );

  click(crumbs[0]);
  await flush();
  assert.equal(bySlot("file-system").getAttribute("data-current-folder-id"), "");
  assert.deepEqual(
    itemRows().map((row) => row.getAttribute("data-item-id")),
    ["archive", "invoices", "alpha", "beta"],
  );
  mounted.cleanup();
});

test("FileSystem back and forward walk the visited folder history", async () => {
  const folders = [];
  const mounted = mount(FileSystem, {
    items: ITEMS,
    "onUpdate:currentFolderId": (id) => folders.push(id),
  });
  await flush();

  const back = document.querySelector('button[aria-label="Back"]');
  const forward = document.querySelector('button[aria-label="Forward"]');
  assert.equal(back.disabled, true, "no history yet");
  assert.equal(forward.disabled, true);

  doubleClick(mainButton(rowFor("invoices")));
  await flush();
  assert.equal(back.disabled, false);

  click(back);
  await flush();
  assert.deepEqual(folders, ["invoices", null]);
  assert.equal(bySlot("file-system").getAttribute("data-current-folder-id"), "");
  assert.equal(forward.disabled, false);

  click(forward);
  await flush();
  assert.deepEqual(folders, ["invoices", null, "invoices"]);
  assert.equal(bySlot("file-system").getAttribute("data-current-folder-id"), "invoices");
  mounted.cleanup();
});

test("FileSystem opens files and emits the resolved file item", async () => {
  const opened = [];
  const mounted = mount(FileSystem, {
    items: ITEMS,
    "onOpen-file": (item) => opened.push({ id: item.id, name: item.name }),
  });
  await flush();

  doubleClick(mainButton(rowFor("alpha")));
  await flush();
  assert.deepEqual(opened, [{ id: "alpha", name: "alpha-contract.pdf" }]);
  assert.equal(bySlot("file-system").getAttribute("data-current-folder-id"), "");
  mounted.cleanup();
});

test("FileSystem renders FileThumbnail previews for image entries", async () => {
  const mounted = mount(FileSystem, {
    items: ITEMS,
    currentFolderId: "invoices",
  });
  await flush();

  const imageRow = rowFor("invoice-scan");
  const thumbnail = imageRow.querySelector('[data-slot="file-thumbnail"]');
  assert.ok(thumbnail, "image entries render FileThumbnail");
  const image = imageRow.querySelector("img");
  assert.ok(image, "FileThumbnail renders the preview image");
  assert.equal(image.getAttribute("src"), PIXEL_PNG);
  assert.equal(rowFor("invoice-pdf").querySelector("img"), null);
  mounted.cleanup();
});

// Keyboard ---------------------------------------------------------------

test("FileSystem keyboard navigation moves selection, opens and goes up", async () => {
  const selectedIds = [];
  const folders = [];
  const files = [];
  const mounted = mount(FileSystem, {
    items: ITEMS,
    "onUpdate:selectedId": (id) => selectedIds.push(id),
    "onOpen-folder": (item) => folders.push(item.id),
    "onOpen-file": (item) => files.push(item.id),
  });
  await flush();

  const first = mainButton(rowFor("archive"));
  assert.equal(first.getAttribute("tabindex"), "0", "first row is in the tab order");
  assert.equal(mainButton(rowFor("invoices")).getAttribute("tabindex"), "-1");

  keydown(first, "ArrowDown");
  await flush();
  assert.deepEqual(selectedIds, ["invoices"]);
  assert.equal(document.activeElement, mainButton(rowFor("invoices")), "focus follows selection");

  keydown(mainButton(rowFor("invoices")), "End");
  await flush();
  assert.deepEqual(selectedIds, ["invoices", "beta"]);

  keydown(mainButton(rowFor("beta")), "Home");
  await flush();
  assert.deepEqual(selectedIds, ["invoices", "beta", "archive"]);

  keydown(mainButton(rowFor("archive")), "Enter");
  await flush();
  assert.deepEqual(folders, ["archive"]);
  assert.equal(bySlot("file-system").getAttribute("data-current-folder-id"), "archive");
  assert.match(bySlot("file-system-empty").textContent, /empty/);

  mounted.cleanup();
});

test("FileSystem keyboard Enter/Space opens files and Left/Backspace leaves a folder", async () => {
  const files = [];
  const folders = [];
  const mounted = mount(FileSystem, {
    items: ITEMS,
    "onOpen-file": (item) => files.push(item.id),
    "onUpdate:currentFolderId": (id) => folders.push(id),
  });
  await flush();

  doubleClick(mainButton(rowFor("invoices")));
  await flush();
  keydown(mainButton(rowFor("invoice-pdf")), "Enter");
  await flush();
  assert.deepEqual(files, ["invoice-pdf"]);

  keydown(mainButton(rowFor("invoice-pdf")), " ");
  await flush();
  assert.deepEqual(files, ["invoice-pdf", "invoice-pdf"]);

  keydown(mainButton(rowFor("invoice-pdf")), "ArrowLeft");
  await flush();
  assert.deepEqual(folders, ["invoices", null]);
  assert.equal(bySlot("file-system").getAttribute("data-current-folder-id"), "");
  mounted.cleanup();
});

test("FileSystem keyboard ArrowRight opens a folder and Delete removes an entry", async () => {
  const folders = [];
  const deleted = [];
  const mounted = mount(FileSystem, {
    items: ITEMS,
    "onOpen-folder": (item) => folders.push(item.id),
    "onDelete-item": (item) => deleted.push(item.id),
  });
  await flush();

  keydown(mainButton(rowFor("invoices")), "ArrowRight");
  await flush();
  assert.deepEqual(folders, ["invoices"]);

  click(bySlot("file-system-crumb"));
  await flush();

  keydown(mainButton(rowFor("beta")), "Delete");
  await flush();
  assert.deepEqual(deleted, ["beta"]);
  mounted.cleanup();
});

// States ----------------------------------------------------------------

test("FileSystem empty, loading and error surfaces expose live status semantics", async () => {
  const empty = mount(FileSystem, { items: ITEMS, currentFolderId: "archive" });
  await flush();
  assert.equal(bySlot("file-system-empty").getAttribute("role"), "status");
  assert.equal(empty.host.querySelector('[data-slot="file-system-empty"]').textContent.trim(), "This folder is empty.");
  empty.cleanup();

  const customEmpty = mount(FileSystem, { items: [], emptyMessage: "Nothing here yet." });
  await flush();
  assert.equal(
    customEmpty.host.querySelector('[data-slot="file-system-empty"]').textContent.trim(),
    "Nothing here yet.",
  );
  customEmpty.cleanup();

  const loading = mount(FileSystem, { items: [], loading: true });
  await flush();
  assert.equal(bySlot("file-system-loading").getAttribute("role"), "status");
  assert.match(bySlot("file-system-loading").textContent, /Loading files/);
  loading.cleanup();

  const error = mount(FileSystem, { items: ITEMS, error: "The drive is unreachable." });
  await flush();
  assert.equal(bySlot("file-system-error").getAttribute("role"), "alert");
  assert.match(bySlot("file-system-error").textContent, /unreachable/);
  assert.equal(bySlot("file-system-grid"), null, "the error state replaces the entries");
  error.cleanup();
});

test("FileSystem read-only hides the delete action and ignores Delete", async () => {
  const deleted = [];
  const mounted = mount(FileSystem, {
    items: ITEMS,
    readOnly: true,
    "onDelete-item": (item) => deleted.push(item.id),
  });
  await flush();

  assert.equal(bySlot("file-system").getAttribute("data-read-only"), "true");
  assert.equal(allBySlot("file-system-item-delete").length, 0);

  keydown(mainButton(rowFor("beta")), "Delete");
  await flush();
  assert.deepEqual(deleted, []);
  mounted.cleanup();

  const editable = mount(FileSystem, {
    items: ITEMS,
    "onDelete-item": (item) => deleted.push(item.id),
  });
  await flush();
  const deleteButton = rowFor("beta").querySelector('[data-slot="file-system-item-delete"]');
  assert.ok(deleteButton);
  assert.equal(deleteButton.getAttribute("aria-label"), "Delete beta-notes.txt");
  click(deleteButton);
  await flush();
  assert.deepEqual(deleted, ["beta"]);
  editable.cleanup();
});

test("FileSystem honours controlled folder and selection props", async () => {
  const selected = ref("alpha");
  const folder = ref("invoices");
  const emitted = [];
  const mounted = mount(FileSystem, {
    items: ITEMS,
    currentFolderId: folder.value,
    selectedId: selected.value,
    "onUpdate:selectedId": (id) => emitted.push(id),
  });
  await flush();

  assert.equal(bySlot("file-system").getAttribute("data-current-folder-id"), "invoices");
  assert.deepEqual(
    itemRows().map((row) => row.getAttribute("data-item-id")),
    ["invoice-pdf", "invoice-scan"],
  );

  click(mainButton(rowFor("invoice-pdf")));
  await flush();
  assert.deepEqual(emitted, ["invoice-pdf"]);
  assert.equal(
    mainButton(rowFor("invoice-pdf")).getAttribute("aria-current"),
    null,
    "the controlled selection does not move until the host updates it",
  );

  mounted.cleanup();

  // A stale controlled folder id resolves to the root instead of rendering a
  // broken breadcrumb.
  const stale = mount(FileSystem, { items: ITEMS, currentFolderId: "deleted-folder" });
  await flush();
  assert.equal(bySlot("file-system").getAttribute("data-current-folder-id"), "");
  stale.cleanup();
});

test("FileSystem keeps the documented responsive breakpoints and tree-free dependency set", () => {
  const source = readFileSync(
    path.join(ROOT, "packages/vue-ui/src/components/FileSystem.vue"),
    "utf8",
  );
  for (const width of ["640px", "768px", "1023px"]) {
    assert.ok(source.includes(width), `FileSystem.vue keeps the ${width} breakpoint`);
  }
  assert.doesNotMatch(source, /@pierre\/trees/u);

  const manifest = JSON.parse(
    readFileSync(path.join(ROOT, "packages/vue-ui/package.json"), "utf8"),
  );
  const dependencies = {
    ...manifest.dependencies,
    ...manifest.devDependencies,
    ...manifest.peerDependencies,
  };
  assert.equal(dependencies["@pierre/trees"], undefined);
  assert.equal(dependencies["@dnd-kit/core"], undefined);
});
