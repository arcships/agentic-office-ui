// Issue #160 — XLSX chrome sits on the shared Office design system.
//
// The behaviour suites already cover every command, menu item and dialog form.
// This file pins the *migration contract*: which shared chrome components and
// shared surface classes each XLSX surface renders, so a future restyle cannot
// silently drop back to local chrome markup.

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  textContent,
  walk,
} from "./vue-test-renderer.mjs";

const { XlsxRibbon, XlsxStatusBar, XlsxToolbar, XlsxViewer } =
  await importFromDemo("@arcships/vue-xlsx");

function readComponent(name) {
  return readFileSync(
    new URL(`../../packages/vue-xlsx/src/components/${name}.vue`, import.meta.url),
    "utf8",
  );
}

const CHROME_READINESS = "viewer chrome migration";

function classesOf(node) {
  const value = node?.props?.class;
  if (typeof value === "string") return value.split(/\s+/).filter(Boolean);
  return [];
}

function nodesWithClass(root, name) {
  return walk(root).filter((node) => classesOf(node).includes(name));
}

function toolbarController(overrides = {}) {
  const calls = [];
  return {
    calls,
    controller: {
      activeCell: { row: 0, col: 0 },
      activeCellAddress: "A1",
      activeSheet: {
        colStyleIds: [],
        maxUsedCol: 1,
        maxUsedRow: 1,
        name: "Sheet1",
        rowStyleIds: [],
        styleById: {},
        visibleCols: [0, 1],
        visibleRows: [0, 1],
      },
      canDownload: true,
      canExport: true,
      canRedo: false,
      canUndo: true,
      canZoomIn: true,
      canZoomOut: true,
      displayFileName: "financial-model.xlsx",
      download: () => calls.push("download"),
      exportXlsx: () => calls.push("export"),
      freezePanes: null,
      getActiveWorksheet: () => null,
      isPageBreakPreview: false,
      isLoading: false,
      isWorkerBacked: false,
      lastExportWarning: "",
      locale: "en",
      openPrintPreview: () => calls.push("print-preview"),
      readOnly: false,
      revision: 0,
      selectedRangeAddress: "A1",
      selection: { end: { row: 0, col: 0 }, start: { row: 0, col: 0 } },
      setPageBreakPreview: (value) => calls.push(["page-breaks", value]),
      setZoomScale: (value) => calls.push(["zoom", value]),
      tabs: [{ name: "Sheet1" }],
      toggleFreezePanesAtActiveCell: () => calls.push("freeze"),
      undo: () => calls.push("undo"),
      redo: () => calls.push("redo"),
      zoomIn: () => calls.push("zoom-in"),
      zoomOut: () => calls.push("zoom-out"),
      zoomScale: 100,
      ...overrides,
    },
  };
}

test("XlsxToolbar renders the shared Office toolbar and keeps its controls", async () => {
  const { controller } = toolbarController();
  const mounted = await mount(XlsxToolbar, { controller, showUpload: true });
  try {
    const root = findByTestId(mounted.root, "xlsx-toolbar");
    assert.ok(root, "the toolbar keeps its public test id");
    const classes = classesOf(root);
    assert.ok(classes.includes("office-toolbar"), `shared toolbar shell: ${classes}`);
    assert.ok(
      classes.includes("office-ui"),
      "the chrome root keeps the shared Office scope class",
    );

    const iconButtons = nodesWithClass(mounted.root, "office-icon-button");
    assert.ok(
      iconButtons.length >= 5,
      `toolbar icon controls use OfficeIconButton (${iconButtons.length})`,
    );
    for (const testId of [
      "xlsx-zoom-out",
      "xlsx-zoom-in",
      "xlsx-search-toggle",
      "xlsx-toolbar-upload",
      "xlsx-toolbar-page-breaks",
      "xlsx-toolbar-print-preview",
      "xlsx-toolbar-download",
      "xlsx-toolbar-export",
    ]) {
      assert.ok(findByTestId(mounted.root, testId), `${testId} is still rendered`);
    }
    assert.ok(
      nodesWithClass(mounted.root, "office-toolbar__separator").length >= 1,
      "the zoom/search cluster keeps a shared separator",
    );
    // The blackbox touch suite measures this class directly.
    assert.equal(
      findByTestId(mounted.root, "xlsx-toolbar-upload").props.class.includes("xlsx-toolbar__icon-button"),
      true,
      "the legacy icon-button class survives for the touch matrix",
    );
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("XlsxRibbon is a shared toolbar of clusters and honours the read-only gate", async () => {
  const { controller } = toolbarController();
  const mounted = await mount(XlsxRibbon, { controller, readOnly: false });
  try {
    const root = findByTestId(mounted.root, "xlsx-format-toolbar");
    const classes = classesOf(root);
    assert.ok(classes.includes("office-toolbar"), `shared ribbon shell: ${classes}`);
    assert.ok(root.props.role === "toolbar", "the ribbon exposes the toolbar role");

    assert.ok(
      nodesWithClass(mounted.root, "office-toolbar__group").length >= 8,
      "every ribbon cluster is an OfficeToolbarGroup",
    );
    assert.ok(
      nodesWithClass(mounted.root, "office-toolbar__separator").length >= 6,
      "clusters stay separated by OfficeToolbarSeparator",
    );
    assert.ok(
      nodesWithClass(mounted.root, "office-button").length >= 20,
      "ribbon actions and toggles use OfficeButton/OfficeIconButton",
    );

    const freeze = findByTestId(mounted.root, "xlsx-ribbon-freeze");
    assert.equal(freeze.props["aria-pressed"], false, "an inactive toggle has no pressed state");
    assert.ok(findByTestId(mounted.root, "xlsx-ribbon-read-only"), "read-only switch is kept");
  } finally {
    mounted.app.unmount();
  }

  const locked = await mount(XlsxRibbon, { controller, readOnly: true });
  try {
    // The old `<fieldset disabled>` gate now lives on each control.
    const cut = findByTestId(locked.root, "xlsx-ribbon-cut");
    assert.equal(cut.props.disabled, true, "a read-only ribbon disables edit commands");
    const readOnly = findByTestId(locked.root, "xlsx-ribbon-read-only");
    assert.equal(readOnly.props.checked, true);
  } finally {
    locked.app.unmount();
  }
});

test("XlsxStatusBar renders OfficeStatusBar slots with the spreadsheet stats", async () => {
  const { controller } = toolbarController({
    selectedRangeAddress: "B3:C4",
    selection: { end: { row: 3, col: 2 }, start: { row: 2, col: 1 } },
  });
  const mounted = await mount(XlsxStatusBar, { controller });
  try {
    const root = findByTestId(mounted.root, "xlsx-status-bar");
    const classes = classesOf(root);
    assert.ok(classes.includes("office-status-bar"), `shared status bar: ${classes}`);
    assert.equal(root.props.role, undefined, "the status bar keeps its footer landmark");
    assert.ok(findByTestId(mounted.root, "xlsx-status-range"), "range address is kept");
    assert.ok(findByTestId(mounted.root, "xlsx-status-note"), "the selection note is kept");
    assert.match(textContent(root), /Selected/);
  } finally {
    mounted.app.unmount();
  }
});

test("XlsxViewer states come from the shared empty/loading/error components", async () => {
  const states = [
    [{ isLoading: true }, "office-state--loading", /Loading/],
    [{ error: new Error("XLSX failed") }, "office-state--error", /XLSX failed/],
    [{}, "office-state--empty", /Open an XLSX file to get started/],
  ];
  for (const [overrides, stateClass, expected] of states) {
    const { controller } = toolbarController({
      activeCell: null,
      activeSheet: null,
      selection: null,
      tabs: [],
      ...overrides,
    });
    const mounted = await mount(XlsxViewer, {
      controller,
      height: "320px",
      showDefaultToolbar: false,
      showFormulaBar: false,
      showImages: false,
      showRibbon: false,
    });
    try {
      assert.equal(
        nodesWithClass(mounted.root, "office-state").length,
        1,
        `${CHROME_READINESS}: one shared state surface`,
      );
      assert.equal(
        nodesWithClass(mounted.root, stateClass).length,
        1,
        `${stateClass} is rendered`,
      );
      assert.match(textContent(mounted.root), expected);
      assert.deepEqual(mounted.warnings, []);
    } finally {
      mounted.app.unmount();
    }
  }
});

test("menus and the authoring dialog render the shared surface contract", () => {
  const contextMenu = readComponent("XlsxContextMenu");
  assert.match(
    contextMenu,
    /class="office-ui office-menu__content office-menu__viewport xlsx-contextmenu"/,
    "the cell menu is the shared menu surface",
  );
  assert.match(contextMenu, /office-menu__item xlsx-contextmenu__item/);
  assert.match(contextMenu, /office-menu__separator/);
  assert.match(contextMenu, /office-menu__item-shortcut/);
  assert.match(
    contextMenu,
    /useOfficeMenuKeyboard/,
    "menu keyboard semantics stay in the shared composable",
  );

  const headerMenu = readComponent("XlsxHeaderMenu");
  assert.match(headerMenu, /office-ui office-menu__content office-menu__viewport xlsx-headermenu/);
  assert.match(headerMenu, /office-menu__item xlsx-headermenu__item/);
  assert.match(headerMenu, /resolveOfficeMenuKeyCommand/);

  const autoFilter = readComponent("XlsxAutoFilterMenu");
  assert.match(autoFilter, /office-ui office-popover__content xlsx-autofilter/);
  assert.match(autoFilter, /office-menu__item xlsx-autofilter__item/);

  const dialog = readComponent("XlsxAuthoringDialog");
  assert.match(dialog, /office-overlay__backdrop xlsx-authoring-backdrop/);
  assert.match(dialog, /office-dialog__content office-dialog__content--lg xlsx-authoring/);
  assert.match(dialog, /office-dialog__header xlsx-authoring__header/);
  assert.match(dialog, /office-dialog__title xlsx-authoring__title/);
  assert.match(dialog, /office-dialog__close xlsx-authoring__close/);
  assert.match(dialog, /data-testid="xlsx-authoring-close"/);
  assert.match(
    dialog,
    /office-button office-button--default office-button--size-sm xlsx-authoring__button/,
    "primary dialog actions use the shared button",
  );
});

test("XLSX chrome resolves its palette from Office tokens, not local hex values", () => {
  for (const name of ["XlsxToolbar", "XlsxRibbon", "XlsxStatusBar", "XlsxContextMenu"]) {
    const source = readComponent(name);
    // Bare hex colours are what the migration removed; the remaining hex values
    // are only `var(--office-*, <fallback>)` defaults for hosts without tokens.
    const styleStart = source.indexOf("<style");
    const styles = source.slice(styleStart);
    for (const match of styles.matchAll(/(?:background|color|border(?:-color)?):\s*([^;]+);/g)) {
      const value = match[1];
      assert.equal(
        /(^|[\s,(])#[0-9a-f]{3,8}\b/i.test(value) && !value.includes("var(--office-"),
        false,
        `${name} hardcodes a colour: ${match[0].trim()}`,
      );
    }
    assert.match(styles, /var\(--office-/, `${name} resolves shared tokens`);
  }
});
