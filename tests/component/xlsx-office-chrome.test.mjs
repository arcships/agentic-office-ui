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

const SHEET_TABS_SOURCE = readComponent("XlsxSheetTabs");
const TOOLBAR_SOURCE = readComponent("XlsxToolbar");
const RIBBON_SOURCE = readComponent("XlsxRibbon");

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

test("XlsxToolbar renders the upstream row and keeps its controls", async () => {
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

    // Default row (issue #160): zoom cluster, search and the workbook actions
    // trigger. Everything upstream keeps in the ⋯ menu stays in the same menu.
    const iconButtons = nodesWithClass(mounted.root, "office-icon-button");
    assert.equal(
      iconButtons.length,
      4,
      `the resting row is zoom out/in, search and actions (${iconButtons.length})`,
    );
    for (const testId of [
      "xlsx-zoom-out",
      "xlsx-zoom-in",
      "xlsx-search-toggle",
      "xlsx-toolbar-actions",
    ]) {
      assert.ok(findByTestId(mounted.root, testId), `${testId} is rendered in the row`);
    }
    assert.ok(
      nodesWithClass(mounted.root, "office-toolbar__separator").length >= 2,
      "the zoom/search/actions clusters keep shared separators",
    );
    const menu = findByTestId(mounted.root, "xlsx-toolbar-menu");
    assert.ok(menu, "the workbook actions menu is rendered");
    assert.equal(menu.props.role, "menu");
    for (const testId of [
      "xlsx-toolbar-upload",
      "xlsx-toolbar-page-breaks",
      "xlsx-toolbar-print-preview",
      "xlsx-toolbar-download",
      "xlsx-toolbar-export",
    ]) {
      const item = findByTestId(mounted.root, testId);
      assert.ok(item, `${testId} is still rendered inside the actions menu`);
      assert.equal(item.props.role === "menuitem" || item.props.role === "menuitemcheckbox", true);
    }
    // The blackbox touch suite measures this class directly.
    assert.equal(
      classesOf(findByTestId(mounted.root, "xlsx-zoom-out")).includes("xlsx-toolbar__icon-button"),
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

    // Issue #170: the ribbon draws every glyph through `HugeIcon` (the icon set
    // the pinned Extend editor resolves), so the source keeps no inline `<svg>`
    // and the rendered icons are Hugeicons icon data (24x24, stroked paths).
    assert.equal(RIBBON_SOURCE.includes("<svg"), false, "the ribbon keeps no inline SVG");
    assert.match(RIBBON_SOURCE, /HugeIcon/, "the ribbon renders Hugeicons");
    assert.doesNotMatch(RIBBON_SOURCE, /icon="(undo-2|redo-2|align-left|align-center|align-right|link|table|check)"/, "no Lucide registry glyph survives in the ribbon");
    const icons = walk(mounted.root).filter((node) => node.type === "svg");
    assert.ok(icons.length >= 29, `every ribbon glyph is an icon: ${icons.length}`);
    for (const icon of icons) {
      assert.equal(icon.props.viewBox, "0 0 24 24", "Hugeicons keep the upstream 24x24 viewBox");
      assert.ok(
        walk(icon).some((node) => node.type === "path" && node.props.stroke === "currentColor"),
        "Hugeicons carry their own stroke data instead of CSS overrides",
      );
    }
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

test("XLSX chrome resolves its palette from Office tokens, not local hex values", () => {
  // The menus and the authoring dialog now render shared primitives, so only
  // the spreadsheet-specific chrome owns style rules.
  for (const name of ["XlsxToolbar", "XlsxRibbon", "XlsxStatusBar"]) {
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
    assert.equal(strip.props.style, undefined, "the strip is styled by CSS, not inline values");
    assert.match(
      SHEET_TABS_SOURCE,
      /\.xlsx-sheettabs \{[^}]*background: color-mix\(in oklab, var\(--office-muted\) 40%/,
      "the strip background is the reference's muted token wash",
    );
    const tabs = nodesWithClass(mounted.root, "xlsx-sheettabs__tab");
    assert.ok(tabs.length >= 2, "tabs still render");
    assert.match(
      SHEET_TABS_SOURCE,
      /\.xlsx-sheettabs__tab \{[^}]*color: inherit/,
      "tabs inherit the tablist colour like the upstream trigger",
    );
    assert.match(
      SHEET_TABS_SOURCE,
      /\.xlsx-sheettabs__list \{[^}]*color: color-mix\(in oklab, var\(--office-muted-foreground\) 72%/,
      "inactive tabs use the muted token at the reference alpha",
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

