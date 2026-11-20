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
  keydownEvent,
  mount,
  teleportTarget,
  textContent,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

const {
  XlsxAuthoringDialog,
  XlsxContextMenu,
  XlsxRibbon,
  XlsxSheetTabs,
  XlsxStatusBar,
  XlsxToolbar,
  XlsxViewer,
} = await importFromDemo("@arcships/vue-xlsx");

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

function sheetTabsController(overrides = {}) {
  const calls = [];
  return {
    calls,
    controller: {
      activeSheet: { name: "Sheet1", colStyleIds: [], rowStyleIds: [], styleById: {} },
      activeTabIndex: 0,
      isEditingLocked: false,
      readOnly: false,
      sheets: [{ hidden: false, index: 0, name: "Sheet1" }, { hidden: false, index: 1, name: "Sheet2" }],
      tabs: [
        { hidden: false, index: 0, kind: "worksheet", name: "Sheet1" },
        { hidden: false, index: 1, kind: "worksheet", name: "Sheet2" },
      ],
      setActiveTabIndex: (index) => calls.push(["tab", index]),
      ...overrides,
    },
  };
}

/** Cell-menu controller: enough sheet geometry for `openAt` to hit a cell. */
function cellMenuController(overrides = {}) {
  const calls = [];
  const sheet = {
    colStyleIds: [],
    colWidths: [],
    defaultColWidthPx: 80,
    defaultRowHeightPx: 24,
    maxUsedCol: 4,
    maxUsedRow: 6,
    name: "Sheet1",
    rightToLeft: false,
    rowHeights: [],
    rowStyleIds: [],
    styleById: {},
    visibleCols: [0, 1, 2, 3],
    visibleRows: [0, 1, 2, 3],
  };
  return {
    calls,
    controller: {
      activeCell: { row: 0, col: 0 },
      activeCellAddress: "A1",
      activeSheet: sheet,
      clearSelectedCells: () => calls.push("clear"),
      cutSelection: () => calls.push("cut"),
      duplicateDrawing: () => calls.push("duplicate"),
      deleteSelectedDrawing: () => calls.push("delete"),
      freezePanes: null,
      getCellComment: () => null,
      getCellHyperlink: () => null,
      getTableAtCell: () => null,
      insertColumns: () => {},
      insertRows: () => {},
      isEditingLocked: false,
      pasteFromClipboard: async () => {},
      readOnly: false,
      redo: () => calls.push("redo"),
      revision: 0,
      selectAll: () => calls.push("selectAll"),
      selectCell: (cell) => calls.push(["cell", cell.row, cell.col]),
      selectColumn: () => calls.push("selectColumn"),
      selectRow: () => calls.push("selectRow"),
      selection: { end: { row: 0, col: 0 }, start: { row: 0, col: 0 } },
      sortTable: () => {},
      tables: [],
      undo: () => calls.push("undo"),
      zoomScale: 100,
      ...overrides,
    },
  };
}

const targetElement = {
  addEventListener: () => {},
  getBoundingClientRect: () => ({ height: 600, left: 0, top: 0, width: 800 }),
  removeEventListener: () => {},
  scrollLeft: 0,
  scrollTop: 0,
};

async function mountCellMenu(controller) {
  let menu;
  const Harness = vue.defineComponent({
    setup() {
      return () => vue.h(XlsxContextMenu, {
        controller,
        ref: (value) => { menu = value; },
        targetElement,
      });
    },
  });
  const mounted = await mount(Harness);
  const body = teleportTarget("body");
  return {
    ...mounted,
    body,
    findInBody: (testId) => walk(body).find((node) => node.props?.["data-testid"] === testId),
    menu: () => menu,
  };
}

test("sheet tabs resolve their palette from the shared theme and tokens", async () => {
  const { calls, controller } = sheetTabsController();
  const mounted = await mount(XlsxSheetTabs, { controller, isDark: true });
  try {
    const theme = walk(mounted.root).find((node) => node.props?.["data-office-theme"] === "dark");
    assert.ok(theme, "the strip provides the dark Office theme to its popups");
    assert.ok(
      nodesWithClass(mounted.root, "office-theme").length === 1,
      "the provider host uses the display:contents shell",
    );
    const strip = nodesWithClass(mounted.root, "xlsx-sheettabs")[0];
    assert.ok(strip, "the strip keeps its public class");
    assert.match(
      String(strip.props.style.backgroundColor),
      /var\(--office-muted/,
      "the strip background is a token, not a hardcoded zinc value",
    );
    const tabs = nodesWithClass(mounted.root, "xlsx-sheettabs__tab");
    assert.ok(tabs.length >= 2, "tabs still render");
    assert.match(String(tabs[0].props.style.color), /inherit/);
    assert.match(
      String(tabs.at(-1).props.style.color),
      /var\(--office-muted-foreground/,
      "inactive tabs use the muted token",
    );
    tabs.at(-1).props.onClick();
    await vue.nextTick();
    assert.deepEqual(calls, [["tab", 1]], "tab switching still goes through the controller");
  } finally {
    mounted.app.unmount();
  }

  const source = readComponent("XlsxSheetTabs");
  assert.equal(
    /props\.isDark \?/.test(source),
    false,
    "the strip no longer branches its palette on isDark",
  );
  assert.match(source, /var\(--office-/, "the strip resolves shared tokens");
});

test("the cell context menu closes on Escape and Tab through the shared keyboard layer", async () => {
  const { controller } = cellMenuController();
  const mounted = await mountCellMenu(controller);
  try {
    assert.equal(mounted.menu().openAt(100, 60), true, "the menu opens on a cell");
    await vue.nextTick();
    const menu = () => mounted.findInBody("xlsx-context-menu");
    assert.ok(menu(), "the shared menu surface is rendered");
    assert.match(
      String(menu().props.class),
      /office-menu__content/,
      "the menu keeps the shared surface class",
    );

    const escape = keydownEvent("Escape");
    menu().props.onKeydown(escape.event);
    assert.equal(escape.preventedCount(), 1, "Escape is consumed by the shared keyboard helper");
    await vue.nextTick();
    assert.equal(menu(), undefined, "Escape closes the menu");

    mounted.menu().openAt(120, 80);
    await vue.nextTick();
    assert.ok(menu(), "the menu reopens");
    const tab = keydownEvent("Tab");
    menu().props.onKeydown(tab.event);
    assert.equal(tab.preventedCount(), 1, "Tab closes the menu instead of leaving it open");
    await vue.nextTick();
    assert.equal(menu(), undefined);

    mounted.menu().openAt(120, 80);
    await vue.nextTick();
    // A key the menu does not own must not be swallowed.
    const arrowDown = keydownEvent("ArrowDown");
    menu().props.onKeydown(arrowDown.event);
    assert.equal(arrowDown.preventedCount(), 0, "unowned keys keep their default");
    assert.ok(menu(), "the menu stays open on an unowned key");
  } finally {
    mounted.app.unmount();
  }
});

test("a read-only cell menu disables editing commands and keeps the read-only offer", async () => {
  const { controller } = cellMenuController({ isEditingLocked: true, readOnly: true });
  const mounted = await mountCellMenu(controller);
  try {
    mounted.menu().openAt(100, 60);
    await vue.nextTick();
    const cut = mounted.findInBody("xlsx-context-cut");
    assert.ok(cut, "cut stays visible so the shortcut is discoverable");
    assert.equal(cut.props.disabled, true, "cut is disabled in a read-only workbook");
    assert.equal(mounted.findInBody("xlsx-context-copy").props.disabled, undefined, "copy stays available");
    assert.equal(mounted.findInBody("xlsx-context-paste"), undefined, "paste is not offered");
    assert.equal(mounted.findInBody("xlsx-context-insert-row-above"), undefined, "row edits are not offered");
    assert.ok(mounted.findInBody("xlsx-context-select-row"), "selection commands stay available");
  } finally {
    mounted.app.unmount();
  }
});

test("the authoring dialog uses the shared dialog shell and keeps its close paths", async () => {
  const applied = [];
  const controller = {
    activeCell: { row: 0, col: 0 },
    activeCellAddress: "A1",
    revision: 0,
    selectedRangeAddress: "A1",
    selection: null,
    setSelectionNumberFormat: (format) => {
      applied.push(format);
      return true;
    },
  };
  const closes = [];
  const mounted = await mount(XlsxAuthoringDialog, {
    controller,
    kind: "numberFormat",
    onClose: () => closes.push("close"),
  });
  const body = teleportTarget("body");
  try {
    const dialog = walk(body).find((node) => node.props?.["data-testid"] === "xlsx-authoring-dialog");
    assert.ok(dialog, "the dialog renders");
    const classes = String(dialog.props.class);
    assert.match(classes, /office-ui/, "shared primitive scope");
    assert.match(classes, /office-dialog__content/, "shared dialog surface");
    assert.match(classes, /office-dialog__content--lg/, "authoring uses the wide surface");
    assert.ok(
      walk(body).some((node) => String(node.props?.class ?? "").includes("office-dialog__body")),
      "the form lives in the shared scroll body",
    );
    assert.ok(
      walk(body).some((node) => String(node.props?.class ?? "").includes("office-overlay__backdrop")),
      "the backdrop is the shared overlay",
    );

    const section = walk(body).find((node) => node.props?.["data-testid"] === "xlsx-number-format-section");
    assert.ok(section, "the number-format form still renders inside the shared shell");

    const escape = keydownEvent("Escape");
    dialog.props.onKeydown(escape.event);
    assert.equal(closes.length, 1, "Escape closes the dialog");

    const backdrop = walk(body).find((node) => node.props?.["data-testid"] === "xlsx-authoring-backdrop");
    backdrop.props.onPointerdown();
    assert.deepEqual(closes, ["close", "close"], "the backdrop closes the dialog");

    const close = walk(body).find((node) => node.props?.["data-testid"] === "xlsx-authoring-close");
    close.props.onClick();
    assert.deepEqual(closes, ["close", "close", "close"], "the header button closes the dialog");

    const apply = walk(body).find((node) => node.props?.["data-testid"] === "xlsx-number-format-apply");
    apply.props.onClick();
    await vue.nextTick();
    assert.equal(applied.length, 1, "applying the format still reaches the controller");
  } finally {
    mounted.app.unmount();
  }
});
