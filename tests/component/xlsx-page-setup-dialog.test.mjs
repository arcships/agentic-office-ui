// Issue #171 — the XLSX page setup dialog on the shared `OfficeDialog`.
//
// The dialog is an Office-reference surface (Excel's Page Setup), so it has no
// upstream component to diff. These tests cover the product contract instead:
// values come from the worksheet, every field maps onto one controller patch,
// cancel/Escape write nothing, and the live summary prints the real pagination.

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
const xlsx = await import(pathToFileURL(requireFromDemo.resolve("@arcships/vue-xlsx")).href);

const { createApp, h, nextTick } = vue;
const { XlsxAuthoringDialog, XlsxPageSetupDialog } = xlsx;

after(() => {
  GlobalRegistrator.unregister();
});

async function flush(ms = 60) {
  await nextTick();
  await new Promise((resolve) => setTimeout(resolve, ms));
  await nextTick();
}

function byTestId(testId) {
  return document.querySelector(`[data-testid="${testId}"]`);
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

function setCheckbox(testId, checked) {
  const element = byTestId(testId);
  assert.ok(element, `${testId} is rendered`);
  element.checked = checked;
  element.dispatchEvent(new window.Event("change", { bubbles: true }));
}

/** Worksheet stub carrying the print records the dialog reads. */
function sheet(overrides = {}) {
  return {
    colWidthOverridesPx: {},
    defaultColWidthPx: 64,
    defaultRowHeightPx: 20,
    colCount: 8,
    printAreas: [{ end: { col: 3, row: 19 }, start: { col: 0, row: 0 } }],
    printOptions: {
      gridLines: false,
      headings: false,
      horizontalCentered: false,
      verticalCentered: false,
    },
    printTitles: { columns: [], rows: [] },
    pageMargins: { bottom: 0.75, footer: 0.3, header: 0.3, left: 0.7, right: 0.7, top: 0.75 },
    pageSetup: {
      blackAndWhite: false,
      cellComments: "none",
      copies: 1,
      draft: false,
      errors: "displayed",
      firstPageNumber: 1,
      fitToHeight: 1,
      fitToPage: false,
      fitToWidth: 1,
      orientation: "portrait",
      pageOrder: "downThenOver",
      paperSize: 1,
      scale: 100,
      useFirstPageNumber: false,
    },
    rowHeightOverridesPx: {},
    colBreaks: [],
    rowBreaks: [],
    rowStyleIds: {},
    colStyleIds: {},
    hiddenCols: [],
    hiddenRows: [],
    maxUsedCol: 5,
    maxUsedRow: 29,
    minUsedCol: 0,
    minUsedRow: 0,
    rowCount: 100,
    name: "Assumptions",
    workbookSheetIndex: 0,
    ...overrides,
  };
}

/** A worksheet stored as A4 landscape at 85% with printed gridlines. */
function landscapeSheet() {
  const base = sheet();
  return sheet({
    pageMargins: { bottom: 0.8, footer: 0.2, header: 0.25, left: 0.5, right: 0.6, top: 0.7 },
    printOptions: {
      gridLines: true,
      headings: true,
      horizontalCentered: true,
      verticalCentered: false,
    },
    pageSetup: {
      ...base.pageSetup,
      orientation: "landscape",
      pageOrder: "overThenDown",
      paperSize: 9,
      scale: 85,
    },
  });
}

function mountDialog(overrides = {}) {
  const calls = [];
  const controller = {
    activeCell: { row: 2, col: 1 },
    activeSheet: landscapeSheet(),
    isEditingLocked: false,
    isWorkerBacked: false,
    readOnly: false,
    revision: 0,
    selection: { end: { col: 3, row: 19 }, start: { col: 0, row: 0 } },
    selectedRangeAddress: "A1:D20",
    setPageSetup: (input) => {
      calls.push(["setPageSetup", input]);
      return true;
    },
    ...overrides,
  };
  const host = document.createElement("div");
  document.body.append(host);
  const closes = [];
  const app = createApp({
    render: () => h(XlsxPageSetupDialog, {
      controller,
      onClose: () => closes.push("close"),
    }),
  });
  app.mount(host);
  return {
    app,
    calls,
    closes,
    cleanup: () => {
      app.unmount();
      host.remove();
      document.body.innerHTML = "";
    },
  };
}

test("the dialog opens on the shared modal surface with the worksheet's current values", async () => {
  const mounted = mountDialog();
  try {
    await flush(120);
    const dialog = byTestId("xlsx-page-setup-dialog");
    assert.ok(dialog, "the dialog carries the public test id");
    assert.equal(dialog.getAttribute("role"), "dialog", "Reka renders a real dialog");
    assert.equal(dialog.getAttribute("data-kind"), "pageSetup");
    assert.ok(dialog.contains(document.activeElement), "focus moves into the dialog on open");
    assert.equal(
      byTestId("xlsx-page-setup-close")?.getAttribute("aria-label") ?? "",
      "Close",
      "the shared close button keeps the dialog close label",
    );
    assert.equal(byTestId("xlsx-page-setup-sheet").textContent.trim(), "Assumptions");

    // Loaded workbook values win over OOXML defaults.
    assert.equal(byTestId("xlsx-page-setup-orientation").value, "landscape");
    assert.equal(byTestId("xlsx-page-setup-paper").value, "9");
    assert.equal(byTestId("xlsx-page-setup-scale-mode").value, "adjust");
    assert.equal(byTestId("xlsx-page-setup-scale").value, "85");
    assert.equal(byTestId("xlsx-page-setup-margin-top").value, "0.7");
    assert.equal(byTestId("xlsx-page-setup-margin-left").value, "0.5");
    assert.equal(byTestId("xlsx-page-setup-margin-footer").value, "0.2");
    assert.equal(byTestId("xlsx-page-setup-page-order").value, "overThenDown");
    assert.equal(byTestId("xlsx-page-setup-grid-lines").checked, true);
    assert.equal(byTestId("xlsx-page-setup-headings").checked, true);
    assert.equal(byTestId("xlsx-page-setup-center-horizontal").checked, true);
    assert.equal(byTestId("xlsx-page-setup-center-vertical").checked, false);
    assert.equal(byTestId("xlsx-page-setup-print-area").value, "A1:D20");

    // The summary is real pagination, not a restatement of the inputs.
    const summary = byTestId("xlsx-page-setup-summary").textContent.trim();
    assert.match(summary, /A4 landscape/);
    assert.match(summary, /85%/);
    assert.match(summary, /^1 page/);
  } finally {
    mounted.cleanup();
  }
});

test("editing the fields sends exactly one page-setup patch on OK", async () => {
  const mounted = mountDialog();
  try {
    await flush(120);
    setSelect("xlsx-page-setup-orientation", "portrait");
    setSelect("xlsx-page-setup-paper", "5");
    setSelect("xlsx-page-setup-page-order", "downThenOver");
    setInput("xlsx-page-setup-scale", "120");
    setInput("xlsx-page-setup-margin-top", "1.25");
    setCheckbox("xlsx-page-setup-grid-lines", false);
    setCheckbox("xlsx-page-setup-center-vertical", true);
    await flush();

    click("xlsx-page-setup-apply");
    await flush();

    assert.equal(mounted.calls.length, 1, "OK is one controller call");
    assert.equal(mounted.calls[0][0], "setPageSetup");
    const input = mounted.calls[0][1];
    assert.equal(input.pageSetup.orientation, "portrait");
    assert.equal(input.pageSetup.paperSize, 5);
    assert.equal(input.pageSetup.pageOrder, "downThenOver");
    assert.equal(input.pageSetup.scale, 120);
    assert.equal(input.pageSetup.fitToPage, false);
    assert.equal(input.pageMargins.top, 1.25);
    assert.equal(input.pageMargins.left, 0.5, "untouched margins keep the loaded value");
    assert.equal(input.printOptions.gridLines, false);
    assert.equal(input.printOptions.verticalCentered, true);
    assert.deepEqual(input.printAreas, [{ end: { col: 3, row: 19 }, start: { col: 0, row: 0 } }]);
    assert.equal(mounted.closes.length, 1, "OK closes the dialog");
  } finally {
    mounted.cleanup();
  }
});

test("fit-to scaling writes page counts instead of a percentage", async () => {
  const mounted = mountDialog();
  try {
    await flush(120);
    setSelect("xlsx-page-setup-scale-mode", "fit");
    await flush();
    assert.equal(byTestId("xlsx-page-setup-scale"), null, "the percentage box is hidden");
    setInput("xlsx-page-setup-fit-width", "1");
    setInput("xlsx-page-setup-fit-height", "0");
    await flush();
    assert.match(
      byTestId("xlsx-page-setup-summary").textContent,
      /自动|as many pages|page/,
      "the summary re-paginates while the draft changes",
    );

    click("xlsx-page-setup-apply");
    await flush();
    const input = mounted.calls[0][1];
    assert.equal(input.pageSetup.fitToPage, true);
    assert.equal(input.pageSetup.fitToWidth, 1);
    assert.equal(input.pageSetup.fitToHeight, 0, "0 keeps Excel's automatic height");
  } finally {
    mounted.cleanup();
  }
});

test("cancel, the close button and Escape write nothing", async () => {
  const dismissals = [
    async () => click("xlsx-page-setup-cancel"),
    async () => click("xlsx-page-setup-close"),
    async () => {
      document.dispatchEvent(new window.KeyboardEvent("keydown", { bubbles: true, key: "Escape" }));
    },
  ];
  for (const dismiss of dismissals) {
    const mounted = mountDialog();
    try {
      await flush(120);
      setSelect("xlsx-page-setup-orientation", "portrait");
      await flush();
      await dismiss();
      await flush(120);
      assert.deepEqual(mounted.calls, [], "dismissing never writes the workbook");
      assert.equal(mounted.closes.length, 1, "the dialog emits close exactly once");
    } finally {
      mounted.cleanup();
    }
  }
});

test("the print area field accepts selections and rejects malformed references", async () => {
  const mounted = mountDialog({
    selection: { end: { col: 6, row: 11 }, start: { col: 1, row: 2 } },
  });
  try {
    await flush(120);
    click("xlsx-page-setup-print-area-selection");
    await flush();
    assert.equal(byTestId("xlsx-page-setup-print-area").value, "B3:G12");

    setInput("xlsx-page-setup-print-area", "A1:");
    await flush();
    assert.ok(byTestId("xlsx-page-setup-print-area-error"), "malformed references are reported");
    assert.equal(byTestId("xlsx-page-setup-apply").disabled, true, "OK is blocked");

    setInput("xlsx-page-setup-print-area", "A1:B2,C3");
    await flush();
    assert.equal(byTestId("xlsx-page-setup-print-area-error"), null);
    click("xlsx-page-setup-apply");
    await flush();
    assert.deepEqual(mounted.calls[0][1].printAreas, [
      { end: { col: 1, row: 1 }, start: { col: 0, row: 0 } },
      { end: { col: 2, row: 2 }, start: { col: 2, row: 2 } },
    ]);

    setInput("xlsx-page-setup-print-area", "");
    await flush();
    click("xlsx-page-setup-apply");
    await flush();
    assert.deepEqual(mounted.calls[1][1].printAreas, [], "an empty field clears the print area");
  } finally {
    mounted.cleanup();
  }
});

test("a locked sheet turns the page setup dialog read-only", async () => {
  for (const overrides of [{ readOnly: true }, { isEditingLocked: true }, { isWorkerBacked: true }]) {
    const mounted = mountDialog(overrides);
    try {
      await flush(120);
      assert.ok(byTestId("xlsx-page-setup-locked"), "the read-only reason is shown");
      assert.equal(byTestId("xlsx-page-setup-apply").disabled, true);
      assert.equal(byTestId("xlsx-page-setup-orientation").disabled, true);
      assert.equal(byTestId("xlsx-page-setup-print-area").disabled, true);
      click("xlsx-page-setup-apply");
      await flush();
      assert.deepEqual(mounted.calls, [], "a locked sheet is never written");
    } finally {
      mounted.cleanup();
    }
  }
});

test("the authoring dialog delegates the pageSetup kind to this dialog", async () => {
  const controller = {
    activeSheet: landscapeSheet(),
    isEditingLocked: false,
    isWorkerBacked: false,
    readOnly: false,
    revision: 0,
    selection: null,
    setPageSetup: () => true,
  };
  const host = document.createElement("div");
  document.body.append(host);
  const closes = [];
  const app = createApp({
    render: () => h(XlsxAuthoringDialog, {
      controller,
      kind: "pageSetup",
      onClose: () => closes.push("close"),
    }),
  });
  app.mount(host);
  try {
    await flush(120);
    assert.ok(byTestId("xlsx-page-setup-dialog"), "the page setup surface renders");
    assert.equal(byTestId("xlsx-authoring-dialog"), null, "the cell-range shell is not mounted");
    click("xlsx-page-setup-cancel");
    await flush(120);
    assert.equal(closes.length, 1, "the authoring dialog forwards the close event");
  } finally {
    app.unmount();
    host.remove();
    document.body.innerHTML = "";
  }
});

test("the dialog keeps the shared-surface contract in source", async () => {
  const { readFileSync } = await import("node:fs");
  const source = readFileSync(
    new URL("../../packages/vue-xlsx/src/components/XlsxPageSetupDialog.vue", import.meta.url),
    "utf8",
  );
  assert.match(source, /<OfficeDialog/, "the dialog renders OfficeDialog");
  assert.match(source, /content-test-id="xlsx-page-setup-dialog"/, "the surface has a stable id");
  assert.match(source, /close-test-id="xlsx-page-setup-close"/, "the close button has a stable id");
  assert.match(source, /class="xlsx-authoring xlsx-page-setup"/, "the shared dialog width applies");
  assert.match(
    source,
    /grid-template-columns: repeat\(auto-fit, minmax\(120px, 1fr\)\)/,
    "sections wrap into columns instead of a fixed two-column layout",
  );
  assert.equal(
    /(?:min-)?(?:width|inline-size):\s*\d{3,}px/.test(source),
    false,
    "no fixed width can overflow a 390px viewport",
  );
  assert.equal(/<Teleport/.test(source), false, "no hand-rolled portal/backdrop");
  assert.equal(
    /buildWorksheetPrintLayout/.test(source),
    true,
    "the summary reuses core pagination",
  );
  assert.equal(
    /controller\.setPageSetup/.test(source),
    true,
    "OK goes through the controller API",
  );
});
