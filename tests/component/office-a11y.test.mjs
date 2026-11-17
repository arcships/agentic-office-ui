import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";

import ts from "../../packages/xlsx-core/node_modules/typescript/lib/typescript.js";
import {
  findByTestId,
  importFromDemo,
  mount,
  textContent,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

// Vue's v-model directive checks these browser constructors; the shared
// component renderer intentionally has no browser DOM.
if (globalThis.Document === undefined) globalThis.Document = class Document {};
if (globalThis.ShadowRoot === undefined) globalThis.ShadowRoot = class ShadowRoot {};

const ROOT = path.resolve(import.meta.dirname, "../..");
const read = (relativePath) => readFileSync(path.join(ROOT, relativePath), "utf8");

const LIVE_STATUS_PATH = "packages/vue-ui/src/components/OfficeLiveStatus.vue";
const MENU_KEYBOARD_PATH = "packages/vue-ui/src/composables/office-menu-keyboard.ts";
const VUE_UI_INDEX_PATH = "packages/vue-ui/src/index.ts";
const DOCX_VIEWER_PATH = "packages/vue-docx/src/components/DocxViewer.vue";
const DOCX_EDITOR_PATH = "packages/vue-docx/src/components/DocxEditor.vue";
const DOCX_MENU_PATH = "packages/vue-docx/src/components/DocxContextMenu.vue";
const PPTX_VIEWER_PATH = "packages/vue-pptx/src/PptxViewer.vue";
const PDF_VIEWER_PATH = "packages/vue-pdf/src/components/PdfViewer.vue";
const XLSX_VIEWER_PATH = "packages/vue-xlsx/src/components/XlsxViewer.vue";
const XLSX_GRID_PATH = "packages/vue-xlsx/src/components/XlsxGrid.vue";
const XLSX_MENU_PATH = "packages/vue-xlsx/src/components/XlsxContextMenu.vue";

const MENU_KEYBOARD_SOURCE = read(MENU_KEYBOARD_PATH);
const MENU_KEYBOARD_FILE = ts.createSourceFile(
  MENU_KEYBOARD_PATH,
  MENU_KEYBOARD_SOURCE,
  ts.ScriptTarget.Latest,
  true,
  ts.ScriptKind.TS,
);

function topLevelDeclaration(name) {
  for (const statement of MENU_KEYBOARD_FILE.statements) {
    if (ts.isFunctionDeclaration(statement) && statement.name?.text === name) return statement;
    if (ts.isVariableStatement(statement)) {
      const declaration = statement.declarationList.declarations.find(
        (candidate) => ts.isIdentifier(candidate.name) && candidate.name.text === name,
      );
      if (declaration) return declaration;
    }
  }
  return undefined;
}

async function loadMenuKeyboard() {
  const selector = topLevelDeclaration("OFFICE_MENU_ITEM_SELECTOR");
  const resolver = topLevelDeclaration("resolveOfficeMenuKeyCommand");
  assert.ok(selector, "office-menu-keyboard.ts 必须导出 OFFICE_MENU_ITEM_SELECTOR");
  assert.ok(resolver, "office-menu-keyboard.ts 必须导出 resolveOfficeMenuKeyCommand");
  const output = ts.transpileModule(
    `const ${selector.getText(MENU_KEYBOARD_FILE)};\n` +
      `${resolver.getText(MENU_KEYBOARD_FILE)}\n` +
      "export { OFFICE_MENU_ITEM_SELECTOR, resolveOfficeMenuKeyCommand };",
    {
      compilerOptions: {
        module: ts.ModuleKind.ESNext,
        target: ts.ScriptTarget.ES2022,
      },
    },
  ).outputText;
  return import(`data:text/javascript;base64,${Buffer.from(output).toString("base64")}`);
}

const menuKeyboard = await loadMenuKeyboard();

async function loadMenuComposable() {
  const source = read(MENU_KEYBOARD_PATH).replace(
    /import \{[^}]*\} from "vue"/,
    "const getCurrentInstance = () => null;\nconst onBeforeUnmount = () => {};",
  );
  const output = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.ESNext,
      target: ts.ScriptTarget.ES2022,
    },
  }).outputText;
  return import(`data:text/javascript;base64,${Buffer.from(output).toString("base64")}`);
}

function computedBody(filePath, name) {
  const source = read(filePath);
  const match = new RegExp(`const ${name} = computed\\(([\\s\\S]*?)\\n\\}\\)`).exec(source);
  assert.ok(match, `${filePath} 必须定义 ${name}`);
  return match[1];
}

function pptxSearchHit(slideIndex, text) {
  return {
    slideIndex,
    nodeId: `node-${slideIndex}`,
    text,
    snippet: text,
    matchStart: 0,
    matchEnd: text.length,
  };
}

// Minimal preview session: the viewer only needs open/render/search here.
function searchSession() {
  let currentSlideIndex = 0;
  const session = {
    document: null,
    get currentSlideIndex() { return currentSlideIndex; },
    get zoomPercent() { return 100; },
    async open(_source, options) {
      session.document = {
        fileName: "a11y-search.pptx",
        height: 540,
        width: 960,
        slides: [0, 1, 2].map((index) => ({ hidden: false, index, number: index + 1 })),
        warnings: [],
      };
      currentSlideIndex = options?.initialSlide ?? 0;
      return session.document;
    },
    async renderSlide(index) { currentSlideIndex = index; },
    renderThumbnail() { return null; },
    searchText(query) {
      return query === "needle"
        ? [pptxSearchHit(0, "needle"), pptxSearchHit(2, "needle")]
        : [];
    },
    async highlightSearchResult() { return null; },
    clearSearchHighlights() {},
    async setZoom() {},
    cancel() {},
    dispose() {},
  };
  return session;
}

test("OfficeLiveStatus is one stable shared live region for DOCX/PPTX/XLSX", () => {
  const component = read(LIVE_STATUS_PATH);
  assert.match(
    component,
    /:role="politeness === 'assertive' \? 'alert' : 'status'"/,
    "assertive announcements must use role=alert, polite ones role=status",
  );
  assert.match(component, /:aria-live="politeness"/);
  assert.match(component, /aria-atomic="true"/);
  assert.match(component, /:aria-busy="busy \? 'true' : undefined"/);
  assert.equal(
    component.match(/\{\{ message \}\}/gu)?.length,
    1,
    "the live region must keep one text node so updates are re-announced, not re-created",
  );
  assert.match(component, /politeness: "polite"/, "default politeness must be polite");
  assert.match(
    component,
    /\.office-live-status \{[\s\S]*?clip-path: inset\(50%\)/,
    "the region must be visually hidden but stay in the accessibility tree",
  );

  const index = read(VUE_UI_INDEX_PATH);
  assert.match(index, /export \{ default as OfficeLiveStatus \} from "\.\/components\/OfficeLiveStatus\.vue"/);
  assert.match(index, /useOfficeMenuKeyboard/);
  assert.match(index, /resolveOfficeMenuKeyCommand/);
});

test("DOCX viewer announces loading, error, page and empty states", () => {
  const viewer = read(DOCX_VIEWER_PATH);
  assert.match(viewer, /import \{ OfficeLiveStatus \} from "@arcships\/vue-ui"/);
  assert.match(viewer, /<OfficeLiveStatus\s+test-id="docx-live-status"\s+:message="liveStatusMessage"\s+:busy="liveStatusBusy"/);
  assert.ok(
    viewer.indexOf('test-id="docx-live-status"') < viewer.indexOf("<DocxViewerToolbar"),
    "the live region must be part of the viewer root, not inside the toolbar",
  );

  const message = computedBody(DOCX_VIEWER_PATH, "liveStatusMessage");
  assert.match(message, /viewerState\.value === "loading"/);
  assert.match(message, /m\.value\.docx\.viewer\.live\.loading/);
  assert.match(message, /formatOfficeMessage\(m\.value\.docx\.viewer\.live\.error, \{ message: error\.value\.message \}\)/);
  assert.match(message, /m\.value\.docx\.viewer\.emptyTitle/);
  assert.match(message, /formatOfficeMessage\(m\.value\.docx\.viewer\.live\.readyPages, \{ page, total \}\)/, "page turns must be announced");
  assert.match(viewer, /const liveStatusBusy = computed\(\(\) => viewerState\.value === "loading"\)/);
});

test("DOCX editor mirrors its status bar into the shared live region", () => {
  const editor = read(DOCX_EDITOR_PATH);
  assert.match(editor, /import \{ OfficeLiveStatus \} from "@arcships\/vue-ui"/);
  assert.match(editor, /<OfficeLiveStatus test-id="editor-live-status" :message="liveStatusMessage" \/>/);
  assert.match(editor, /const localizedStatus = computed\(\(\) => String\(controller\.status/);
  const message = computedBody(DOCX_EDITOR_PATH, "liveStatusMessage");
  assert.match(message, /pageLimitError\.value/);
  assert.match(message, /formatOfficeMessage\(m\.value\.common\.page\.indicator/);
  assert.match(message, /m\.value\.docx\.editor\.liveReady/);
  assert.match(message, /localizedStatus\.value/);
});

test("PPTX and XLSX viewers use the same shared live region", () => {
  const pptx = read(PPTX_VIEWER_PATH);
  assert.match(pptx, /import \{ OfficeLiveStatus \} from "@arcships\/vue-ui"/);
  assert.match(pptx, /<OfficeLiveStatus\s+test-id="pptx-live-status"\s+:message="liveStatusMessage"\s+:busy="state === 'loading'"/);
  const pptxMessage = computedBody(PPTX_VIEWER_PATH, "liveStatusMessage");
  assert.match(pptxMessage, /m\.value\.pptx\.viewer\.live\.empty/);
  assert.match(pptxMessage, /m\.value\.pptx\.viewer\.live\.loading/);
  assert.match(pptxMessage, /formatOfficeMessage\(m\.value\.pptx\.viewer\.live\.error/);
  assert.match(pptxMessage, /formatOfficeMessage\(m\.value\.pptx\.viewer\.live\.ready, \{/, "slide changes must be announced");
  assert.match(pptxMessage, /current: activeIndex\.value \+ 1/);

  const xlsx = read(XLSX_VIEWER_PATH);
  assert.match(xlsx, /import \{ OfficeLiveStatus \} from "@arcships\/vue-ui"/);
  assert.match(xlsx, /<OfficeLiveStatus\s+test-id="xlsx-live-status"\s+:message="liveStatusMessage"\s+:busy="controller\.isLoading"/);
  const xlsxMessage = computedBody(XLSX_VIEWER_PATH, "liveStatusMessage");
  assert.match(xlsxMessage, /m\.value\.xlsx\.viewer\.live\.loading/);
  assert.match(xlsxMessage, /formatOfficeMessage\(m\.value\.xlsx\.viewer\.live\.error/);
  assert.match(xlsxMessage, /formatOfficeMessage\(m\.value\.xlsx\.viewer\.live\.sheetReady, \{ name: sheet\.name \}\)/);
});

test("PPTX announces search hits in a live region like the other viewers", () => {
  const pptx = read(PPTX_VIEWER_PATH);
  assert.match(
    pptx,
    /<OfficeLiveStatus test-id="pptx-search-status" :message="searchStatusMessage" \/>/,
    "the slide search hit counter needs its own live region",
  );
  assert.match(
    pptx,
    /class="pptx-viewer__search-count" aria-hidden="true"/,
    "the visual counter must not be announced twice",
  );

  const message = computedBody(PPTX_VIEWER_PATH, "searchStatusMessage");
  assert.match(message, /pptxDocument\.searchState\.value/);
  assert.match(message, /m\.value\.common\.search\.searching/);
  assert.match(message, /m\.value\.common\.search\.error/);
  assert.match(message, /m\.value\.common\.search\.noResults/);
  assert.match(
    message,
    /formatOfficeMessage\(m\.value\.common\.search\.resultOf, \{ current, total \}\)/,
    "the announcement must reuse the DOCX/XLSX find-bar wording",
  );
  assert.match(message, /search\.activeIndex \+ 1/, "the announced hit must follow the active result");
  assert.match(message, /Math\.min\(Math\.max\(search\.activeIndex \+ 1, 1\), total\)/);
});

test("PPTX search live region announces counts, no results and clearing", async (t) => {
  const { PptxViewer } = await importFromDemo("@arcships/vue-pptx");
  const mounted = await mount(PptxViewer, {
    source: new Uint8Array([1, 2, 3]),
    sessionFactory: () => searchSession(),
    showSearch: true,
    showSidebar: false,
  });
  t.after(() => mounted.app.unmount());

  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "ready",
    "PPTX ready state",
  );
  const form = walk(mounted.root).find(
    (node) => node.type === "form" && node.props?.role === "search",
  );
  assert.ok(form, "the search form must render");
  const input = walk(form).find((node) => node.type === "input");
  const statusText = () => textContent(findByTestId(mounted.root, "pptx-search-status"));

  assert.equal(statusText(), "", "an untouched viewer must not announce anything");
  assert.equal(findByTestId(mounted.root, "pptx-search-status").props.role, "status");

  async function submit(query) {
    input.value = query;
    input.dispatchEvent({ type: "input", target: input });
    await vue.nextTick();
    await form.props.onSubmit({ preventDefault() {} });
  }

  await submit("needle");
  await waitFor(() => /1 of 2/.test(statusText()), "announced first hit", 4_000);

  const next = walk(form).find(
    (node) => node.type === "button" && node.props?.["aria-label"] === "Next result",
  );
  assert.ok(next, "the next-hit button must render once there are results");
  next.props.onClick();
  await waitFor(() => /2 of 2/.test(statusText()), "announced second hit", 4_000);

  await submit("missing");
  await waitFor(() => statusText() === "No results", "announced empty result set", 4_000);

  input.props.onKeydown({ key: "Escape", preventDefault() {} });
  await waitFor(() => statusText() === "", "cleared search announcement", 4_000);
  assert.deepEqual(mounted.warnings, []);
});

test("vue-pdf announces viewer state and search hits through live regions", () => {
  const pdf = read(PDF_VIEWER_PATH);
  assert.match(
    pdf,
    /data-testid="pdf-status"[^>]*role="status"[^>]*aria-live="polite"[^>]*aria-atomic="true"/,
    "the PDF viewer state line must be a polite, atomic live region",
  );
  assert.match(pdf, /class="pdf-status"[^>]*>\{\{ viewerState \}\}</, "the region must speak the viewer state");
  assert.match(
    pdf,
    /\.pdf-status \{[^}]*clip-path: inset\(50%\)/,
    "the PDF state region must be visually hidden but announced",
  );
  assert.match(
    pdf,
    /data-testid="pdf-search-result"[^>]*role="status"[^>]*aria-atomic="true"/,
    "PDF search hits need an atomic status region",
  );
  // The announced hit carries the current index, the total and the page.
  assert.match(pdf, /m\.value\.pdf\.viewer\.searchResult, \{\s*current: 1,\s*total: hits\.length,\s*page: hits\[0\]\.pageIndex \+ 1,/);
  assert.match(pdf, /m\.value\.pdf\.viewer\.searchResult, \{\s*current: next \+ 1,\s*total: searchHits\.value\.length,\s*page: hit\.pageIndex \+ 1,/);
  assert.match(pdf, /m\.value\.pdf\.viewer\.searchNoResults/);
});

test("XLSX canvas grid exposes role=grid and a live active descendant", () => {
  const grid = read(XLSX_GRID_PATH);
  assert.match(grid, /role="grid"/);
  assert.match(grid, /:aria-label="gridAriaLabel"/);
  assert.match(grid, /:aria-rowcount="gridRowCount"/);
  assert.match(grid, /:aria-colcount="gridColCount"/);
  assert.match(grid, /:aria-readonly="gridAriaReadOnly"/);
  assert.match(grid, /:aria-activedescendant="activeCellDescendantId"/);

  // aria-activedescendant must reference an element owned by the grid.
  const rowIndex = grid.indexOf('role="row"');
  const gridCellIndex = grid.indexOf('role="gridcell"');
  const scrollSpaceIndex = grid.indexOf('class="xlsx-grid__scroll-space"');
  assert.ok(rowIndex > 0 && rowIndex < scrollSpaceIndex, "the a11y row must be a child of the grid element");
  assert.ok(gridCellIndex > rowIndex && gridCellIndex < scrollSpaceIndex, "the a11y cell must live inside that row");

  assert.match(grid, /:id="activeCellDescendantId"/);
  assert.match(grid, /:aria-rowindex="activeCellAria\.rowIndex"/);
  assert.match(grid, /:aria-colindex="activeCellAria\.colIndex"/);
  assert.match(grid, /:aria-label="activeCellAria\.label"/);
  assert.match(grid, /aria-selected="true"/);
  assert.match(grid, /\{\{ activeCellAria\.text \}\}/);

  // The active descendant carries the workbook address (A1-style) and its value.
  const descriptor = computedBody(XLSX_GRID_PATH, "activeCellAria");
  assert.match(descriptor, /columnLabel\(cell\.col\)/);
  assert.match(descriptor, /cell\.row \+ 1/);
  assert.match(descriptor, /props\.controller\.getCellDisplayValue\(cell\)/);
  assert.match(descriptor, /colIndex: cell\.col \+ 1/);
  assert.match(descriptor, /rowIndex: cell\.row \+ 1/);
  assert.match(descriptor, /m\.value\.xlsx\.grid\.blank/, "empty cells need a spoken placeholder");

  const descendantId = computedBody(XLSX_GRID_PATH, "activeCellDescendantId");
  assert.match(descendantId, /\$\{gridA11yId\}-active-cell/);
  assert.match(
    grid,
    /import \{ nextGridA11yId \} from "\.\/grid-a11y"/,
    "grid ids must come from a module-scope sequence so two grids never share an id",
  );
  assert.match(grid, /const gridA11yId = nextGridA11yId\(\)/);
  assert.doesNotMatch(
    grid,
    /let gridInstanceCount/,
    "a setup()-scoped counter would restart for every mounted grid",
  );
  const label = computedBody(XLSX_GRID_PATH, "gridAriaLabel");
  assert.match(label, /formatOfficeMessage\(m\.value\.xlsx\.grid\.worksheetNamed, \{ name \}\)/, "the grid label must name the worksheet");
  assert.match(grid, /const declared = activeSheet\.value\?\.rowCount \?\? 0/);
  assert.match(grid, /const declared = activeSheet\.value\?\.colCount \?\? 0/);

  assert.match(
    grid,
    /\.xlsx-grid__a11y, \.xlsx-grid__a11y-cell \{[\s\S]*?clip-path: inset\(50%\)/,
    "the a11y cell must be visually hidden without leaving the accessibility tree",
  );
  assert.doesNotMatch(grid, /\.xlsx-grid__a11y[^{]*\{[^}]*display: none/);
});

test("shared menu keyboard resolves Escape, Tab and arrow navigation the same way", () => {
  assert.equal(
    menuKeyboard.OFFICE_MENU_ITEM_SELECTOR,
    '[role="menuitem"]:not([disabled]):not([aria-disabled="true"])',
  );
  assert.deepEqual(menuKeyboard.resolveOfficeMenuKeyCommand("Escape", 4, 2), {
    type: "close",
    restoreFocus: true,
  });
  assert.deepEqual(menuKeyboard.resolveOfficeMenuKeyCommand("Tab", 4, 2), {
    type: "close",
    restoreFocus: true,
  });
  assert.deepEqual(
    menuKeyboard.resolveOfficeMenuKeyCommand("Escape", 0, -1),
    { type: "close", restoreFocus: true },
    "Escape must close a menu that has no enabled items",
  );
  assert.deepEqual(menuKeyboard.resolveOfficeMenuKeyCommand("ArrowDown", 4, -1), { type: "focus", index: 0 });
  assert.deepEqual(menuKeyboard.resolveOfficeMenuKeyCommand("ArrowDown", 4, 3), { type: "focus", index: 0 });
  assert.deepEqual(menuKeyboard.resolveOfficeMenuKeyCommand("ArrowUp", 4, -1), { type: "focus", index: 3 });
  assert.deepEqual(menuKeyboard.resolveOfficeMenuKeyCommand("ArrowUp", 4, 0), { type: "focus", index: 3 });
  assert.deepEqual(menuKeyboard.resolveOfficeMenuKeyCommand("Home", 4, 2), { type: "focus", index: 0 });
  assert.deepEqual(menuKeyboard.resolveOfficeMenuKeyCommand("End", 4, 1), { type: "focus", index: 3 });
  assert.equal(menuKeyboard.resolveOfficeMenuKeyCommand("ArrowDown", 0, -1), null);
  assert.equal(menuKeyboard.resolveOfficeMenuKeyCommand("Enter", 4, 0), null);
  assert.equal(menuKeyboard.resolveOfficeMenuKeyCommand("a", 4, 0), null);
});

test("DOCX and XLSX context menus share focus entry, arrow keys, Escape and focus return", () => {
  for (const [name, filePath] of [["DOCX", DOCX_MENU_PATH], ["XLSX", XLSX_MENU_PATH]]) {
    const menu = read(filePath);
    assert.match(
      menu,
      /from "@arcships\/vue-ui"/,
      `${name} menu must consume the shared Office menu keyboard`,
    );
    assert.match(menu, /role="menu"/, `${name} menu needs role=menu`);
    assert.match(menu, /aria-label="/, `${name} menu needs an accessible name`);
    assert.match(menu, /role="menuitem"/, `${name} items need role=menuitem`);
    assert.match(menu, /role="separator"/, `${name} separators need role=separator`);
    assert.match(menu, /@keydown="onMenuKeydown"/);
    assert.match(menu, /function onMenuKeydown\(event: KeyboardEvent\): void/);
    assert.match(menu, /captureInvoker\(\)/, `${name} menu must remember the invoking element`);
    assert.doesNotMatch(
      menu,
      /window\.addEventListener\(["']keydown/,
      `${name} menu must not own a private Escape listener that diverges from the shared semantics`,
    );
  }

  // The XLSX menu has no non-DOM test-renderer contract, so it hands the whole
  // gesture set to the shared composable.
  const xlsxMenu = read(XLSX_MENU_PATH);
  assert.match(xlsxMenu, /import \{ useOfficeMenuKeyboard \} from "@arcships\/vue-ui"/);
  assert.match(xlsxMenu, /menuKeyboard\.onKeydown\(event\)/);
  assert.match(xlsxMenu, /menu: \(\) => menuRef\.value/);
  assert.match(xlsxMenu, /focusFirst\(\)/);
  const invokerIndex = xlsxMenu.indexOf("menuKeyboard.captureInvoker()");
  const visibleIndex = xlsxMenu.indexOf("visible.value = true");
  assert.ok(invokerIndex > 0, "XLSX menu must capture the invoker");
  assert.ok(invokerIndex < visibleIndex, "the invoker must be captured before the menu becomes visible");
  assert.match(xlsxMenu, /void nextTick\(\(\) => menuKeyboard\.focusFirst\(\)\)/);
  assert.match(
    xlsxMenu,
    /restoreTarget: \(\) => props\.targetElement \?\? null/,
    "XLSX menu must return focus to the grid that opened it",
  );

  // The DOCX menu keeps its node-agnostic item walk (the component test renderer
  // has no DOM), but the key semantics come from the same shared resolver.
  const docxMenu = read(DOCX_MENU_PATH);
  assert.match(docxMenu, /import \{ resolveOfficeMenuKeyCommand \} from "@arcships\/vue-ui"/);
  assert.match(docxMenu, /resolveOfficeMenuKeyCommand\(\s*event\.key,\s*collectMenuItems\(\)\.length,\s*activeItemIndex\.value,?\s*\)/);
  assert.match(docxMenu, /if \(command\.type === "focus"\) \{\s*focusMenuItem\(command\.index\)/);
  assert.match(docxMenu, /function focusMenuItem\(index: number\): void/);
  assert.match(docxMenu, /item\.setAttribute\?\.\("tabindex"/, "DOCX keeps its roving tabindex");
  assert.match(docxMenu, /function restoreInvokerFocus\(\): void/);
  assert.match(docxMenu, /function close\(\): void \{\s*restoreInvokerFocus\(\)\s*emit\("close"\)/);
  assert.match(docxMenu, /captureInvoker\(\)\n\s*await nextTick\(\)\n\s*measureMenu\(\)\n\s*focusMenuItem\(0\)/);
});

test("the shared menu keyboard keeps Escape working when focus never enters the menu", () => {
  const source = MENU_KEYBOARD_SOURCE;
  assert.match(source, /function restoreFocus\(\): void/);
  assert.match(source, /invoker = activeElement\(\)/);
  assert.match(source, /window\.addEventListener\("keydown", onWindowKeydown\)/);
  assert.match(source, /onBeforeUnmount\(\(\) => window\.removeEventListener\("keydown", onWindowKeydown\)\)/);
  assert.match(
    source,
    /if \(focused && typeof menu\.contains === "function" && menu\.contains\(focused\)\) return/,
    "the window fallback must not shadow the menu's own keydown handling",
  );
  assert.match(source, /if \(options\.windowFallback !== false/);
  assert.match(source, /options\.close\(\)\n    restoreFocus\(\)/, "closing must always return focus");
});

test("the shared menu composable focuses items, closes on Escape and returns focus", async (t) => {
  const { useOfficeMenuKeyboard } = await loadMenuComposable();
  const previousWindow = globalThis.window;
  const previousDocument = globalThis.document;

  const windowListeners = new Map();
  globalThis.window = {
    addEventListener(type, listener) { windowListeners.set(type, listener); },
    removeEventListener(type, listener) { windowListeners.delete(type, listener); },
  };

  let focused = null;
  let active = null;
  const item = (name) => ({
    name,
    focus() {
      focused = name;
      active = this;
    },
  });
  const items = [item("copy"), item("paste"), item("clear")];
  const menuElement = {
    contains: (node) => items.includes(node),
    querySelectorAll: () => items,
  };
  const invoker = item("grid");
  active = invoker;
  globalThis.document = { get activeElement() { return active; } };

  t.after(() => {
    if (previousWindow === undefined) delete globalThis.window;
    else globalThis.window = previousWindow;
    if (previousDocument === undefined) delete globalThis.document;
    else globalThis.document = previousDocument;
  });

  let closed = 0;
  let visible = true;
  const keyboard = useOfficeMenuKeyboard({
    menu: () => (visible ? menuElement : null),
    close: () => { closed += 1; visible = false; },
    restoreTarget: () => invoker,
  });

  assert.ok(windowListeners.has("keydown"), "the Escape fallback must be registered");

  keyboard.captureInvoker();
  keyboard.focusFirst();
  assert.equal(focused, "copy", "opening the menu must focus the first item");

  const event = (key) => ({
    key,
    defaultPrevented: false,
    preventDefault() { this.defaultPrevented = true; },
    stopPropagation() {},
  });

  active = items[0];
  keyboard.onKeydown(event("ArrowDown"));
  assert.equal(focused, "paste", "ArrowDown must move to the next item");
  keyboard.onKeydown(event("End"));
  assert.equal(focused, "clear", "End must focus the last item");
  keyboard.onKeydown(event("ArrowDown"));
  assert.equal(focused, "copy", "ArrowDown must wrap to the first item");

  const escape = event("Escape");
  active = items[0];
  keyboard.onKeydown(escape);
  assert.equal(closed, 1, "Escape must close the menu");
  assert.equal(escape.defaultPrevented, true);
  assert.equal(focused, "grid", "closing must return focus to the invoking surface");

  // Re-open through the window fallback while focus is outside the menu.
  visible = true;
  keyboard.captureInvoker();
  active = invoker;
  const outsideEscape = event("Escape");
  windowListeners.get("keydown")(outsideEscape);
  assert.equal(closed, 2, "Escape from outside the menu must still close it");
  assert.equal(outsideEscape.defaultPrevented, true);

  // A focused item is handled by the menu's own keydown, not twice.
  visible = true;
  active = items[1];
  windowListeners.get("keydown")(event("Escape"));
  assert.equal(closed, 2, "the window fallback must not duplicate the menu handler");

  const restoreEscape = event("Escape");
  active = items[2];
  keyboard.onKeydown(restoreEscape);
  assert.equal(closed, 3);
  assert.equal(focused, "grid", "focus must always return to the invoker");
});
