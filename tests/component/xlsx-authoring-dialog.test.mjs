// Issue #160 — the XLSX authoring dialog on the shared `OfficeDialog` (Reka).
//
// The dialog is a real primitive-layer modal now, so it needs a browser-like
// DOM: this file boots happy-dom and drives the form through real input/click
// events. It replaces the fake-renderer dialog assertions.

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
const { XlsxAuthoringDialog } = xlsx;

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

/** Set a text/number input through the DOM so `v-model` sees the change. */
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

function mountDialog(kind, overrides = {}) {
  const calls = [];
  const base = {
    activeCell: { row: 2, col: 1 },
    activeCellAddress: "B3",
    activeSheet: { name: "Sheet1" },
    addConditionalFormatRule: (draft) => {
      calls.push(["addConditionalFormatRule", draft]);
      return true;
    },
    createTable: (input) => {
      calls.push(["createTable", input]);
      return { name: input?.name ?? "Table1" };
    },
    deleteChart: (id) => {
      calls.push(["deleteChart", id]);
      return true;
    },
    deleteConditionalFormatRule: (index) => {
      calls.push(["deleteConditionalFormatRule", index]);
      return true;
    },
    getCellComment: () => null,
    getCellDataValidation: () => null,
    getCellHyperlink: () => null,
    getChartById: () => null,
    getConditionalFormatRules: () => [],
    getTableAtCell: () => null,
    insertChart: (input) => {
      calls.push(["insertChart", input]);
      return { chartType: "Line", id: "chart-9-0", title: input.title };
    },
    removeTable: (name) => {
      calls.push(["removeTable", name]);
      return true;
    },
    revision: 0,
    selectChart: (id) => calls.push(["selectChart", id]),
    selectedChartId: null,
    selectedRangeAddress: "B3:D5",
    selection: { end: { row: 4, col: 3 }, start: { row: 2, col: 1 } },
    setCellComment: (cell, input) => {
      calls.push(["setCellComment", cell, input]);
      return Boolean(input);
    },
    setCellHyperlink: (cell, input) => {
      calls.push(["setCellHyperlink", cell, input]);
      return Boolean(input);
    },
    setSelectionDataValidation: (input) => {
      calls.push(["setSelectionDataValidation", input]);
      return true;
    },
    setSelectionNumberFormat: (format) => {
      calls.push(["setSelectionNumberFormat", format]);
      return true;
    },
    updateChart: (id, patch) => calls.push(["updateChart", id, patch]),
    updateConditionalFormatRule: (index, draft) => {
      calls.push(["updateConditionalFormatRule", index, draft]);
      return true;
    },
    ...overrides,
  };

  const host = document.createElement("div");
  document.body.append(host);
  const closes = [];
  const app = createApp({
    render: () => h(XlsxAuthoringDialog, {
      controller: base,
      kind,
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

test("the dialog renders on the shared modal surface with a titled header", async () => {
  const mounted = mountDialog("numberFormat");
  try {
    await flush(120);
    const dialog = byTestId("xlsx-authoring-dialog");
    assert.ok(dialog, "the dialog carries the public test id");
    assert.equal(dialog.getAttribute("role"), "dialog", "Reka renders a real dialog");
    assert.equal(dialog.getAttribute("data-kind"), "numberFormat");
    assert.ok(byTestId("xlsx-number-format-section"), "the form section still renders");
    assert.ok(byTestId("xlsx-authoring-range"), "the range chip stays inside the body");
    assert.ok(byTestId("xlsx-authoring-close"), "the shared close button keeps its test id");
    assert.ok(
      dialog.contains(document.activeElement),
      "focus moves into the dialog on open",
    );
    assert.equal(document.body.style.overflow, "hidden", "the modal locks the page scroll");
  } finally {
    mounted.cleanup();
  }
});

test("Escape, the close button and the backdrop all close the dialog", async () => {
  for (const close of [
    async () => {
      document.dispatchEvent(new window.KeyboardEvent("keydown", { bubbles: true, key: "Escape" }));
    },
    async () => {
      click("xlsx-authoring-close");
    },
    async () => {
      const overlay = document.querySelector('[data-slot="dialog-overlay"]');
      assert.ok(overlay, "the shared overlay is rendered");
      overlay.dispatchEvent(new window.MouseEvent("pointerdown", { bubbles: true }));
    },
  ]) {
    const mounted = mountDialog("numberFormat");
    try {
      await flush(120);
      await close();
      await flush(120);
      assert.equal(mounted.closes.length, 1, "the dialog emits close exactly once");
    } finally {
      mounted.cleanup();
    }
  }
});

test("the number-format form still builds and applies a format", async () => {
  const mounted = mountDialog("numberFormat");
  try {
    await flush(120);
    setSelect("xlsx-number-format-category", "percent");
    setInput("xlsx-number-format-decimals", "1");
    await flush();
    click("xlsx-number-format-apply");
    assert.equal(mounted.calls.at(-1)[0], "setSelectionNumberFormat");
    const format = mounted.calls.at(-1)[1];
    assert.equal(format.formatType, "custom");
    assert.match(format.formatString, /%/);
  } finally {
    mounted.cleanup();
  }
});

test("the comment dialog saves and deletes the note on the active cell", async () => {
  const existing = { author: "Dana", cell: { row: 2, col: 1 }, text: "Old note" };
  const mounted = mountDialog("comment", { getCellComment: () => existing });
  try {
    await flush(120);
    assert.equal(byTestId("xlsx-comment-text").value, "Old note");
    setInput("xlsx-comment-text", "New note");
    await flush();
    click("xlsx-comment-save");
    await flush();
    assert.deepEqual(mounted.calls[0], [
      "setCellComment",
      { row: 2, col: 1 },
      { author: "Dana", text: "New note" },
    ]);
    assert.match(byTestId("xlsx-authoring-status").textContent, /comment/i);
    click("xlsx-comment-delete");
    assert.deepEqual(mounted.calls[1], ["setCellComment", { row: 2, col: 1 }, null]);
  } finally {
    mounted.cleanup();
  }
});

test("the hyperlink dialog saves the link on the active cell", async () => {
  const mounted = mountDialog("hyperlink");
  try {
    await flush(120);
    setInput("xlsx-hyperlink-target", "https://example.com");
    await flush();
    click("xlsx-hyperlink-save");
    assert.deepEqual(mounted.calls[0], [
      "setCellHyperlink",
      { row: 2, col: 1 },
      { target: "https://example.com", tooltip: undefined },
    ]);
  } finally {
    mounted.cleanup();
  }
});

test("the table dialog creates and removes a ListObject", async () => {
  const table = { name: "ExistingTable", reference: "B3:D5" };
  const mounted = mountDialog("table", { getTableAtCell: () => table });
  try {
    await flush(120);
    setInput("xlsx-table-name", "NewTable");
    await flush();
    click("xlsx-table-create");
    assert.deepEqual(mounted.calls[0], [
      "createTable",
      {
        hasHeaders: true,
        name: "NewTable",
        range: { end: { col: 3, row: 4 }, start: { col: 1, row: 2 } },
      },
    ]);
    click("xlsx-table-remove");
    assert.deepEqual(mounted.calls[1], ["removeTable", "ExistingTable"]);
  } finally {
    mounted.cleanup();
  }
});

test("the conditional-format dialog adds, edits and deletes rules", async () => {
  const rule = {
    cfvos: [{ type: "num", value: 10 }, { type: "percentile", value: 30 }, { type: "num", value: 90 }],
    colors: [{ rgb: "#f8696b" }, { rgb: "#ffeb84" }, { rgb: "#63be7b" }],
    kind: "colorScale",
    priority: 1,
    ranges: [{ end: { col: 1, row: 4 }, start: { col: 1, row: 2 } }],
  };
  const mounted = mountDialog("conditional", { getConditionalFormatRules: () => [rule] });
  try {
    await flush(120);
    assert.ok(byTestId("xlsx-conditional-rule"), "existing rules are listed");
    setSelect("xlsx-conditional-kind", "dataBar");
    await flush();
    click("xlsx-conditional-add");
    const [, draft] = mounted.calls[0];
    assert.equal(draft.kind, "dataBar");
    assert.equal(draft.colors.length, 1);
    assert.deepEqual(draft.thresholds.map((threshold) => threshold.type), ["min", "max"]);

    click("xlsx-conditional-edit");
    await flush();
    click("xlsx-conditional-add");
    assert.equal(mounted.calls[1][0], "updateConditionalFormatRule");
    assert.equal(mounted.calls[1][1], 0);
    assert.deepEqual(
      mounted.calls[1][2].thresholds.map((threshold) => [threshold.type, threshold.value ?? null]),
      [["num", "10"], ["percentile", "30"], ["num", "90"]],
      "editing keeps the stored cfvos",
    );

    click("xlsx-conditional-delete");
    assert.deepEqual(mounted.calls[2], ["deleteConditionalFormatRule", 0]);
  } finally {
    mounted.cleanup();
  }
});

test("the validation dialog applies and removes data validation", async () => {
  const existing = { listSource: '"A,B"', validationType: "list" };
  const mounted = mountDialog("validation", { getCellDataValidation: () => existing });
  try {
    await flush(120);
    setInput("xlsx-validation-source", "Red,Green");
    await flush();
    click("xlsx-validation-apply");
    const [, input] = mounted.calls[0];
    assert.equal(input.validationType, "list");
    assert.equal(input.listSource, "Red,Green");
    assert.deepEqual(input.ranges, [{ end: { col: 3, row: 4 }, start: { col: 1, row: 2 } }]);
    click("xlsx-validation-remove");
    assert.deepEqual(mounted.calls[1], ["setSelectionDataValidation", null]);
  } finally {
    mounted.cleanup();
  }
});

test("the chart dialog inserts from the selection and selects the new chart", async () => {
  const mounted = mountDialog("chart");
  try {
    await flush(120);
    setSelect("xlsx-chart-family", "pie");
    setInput("xlsx-chart-title", "Share");
    await flush();
    click("xlsx-chart-insert");
    assert.deepEqual(mounted.calls[0], [
      "insertChart",
      {
        family: "pie",
        hasHeaderColumn: true,
        hasHeaderRow: true,
        range: { end: { col: 3, row: 4 }, start: { col: 1, row: 2 } },
        title: "Share",
      },
    ]);
    assert.deepEqual(mounted.calls[1], ["selectChart", "chart-9-0"]);
  } finally {
    mounted.cleanup();
  }
});

test("the chart dialog falls back to the used range for a single cell", async () => {
  const cell = { col: 1, row: 2 };
  const mounted = mountDialog("chart", {
    selectedRangeAddress: "B3",
    selection: { end: cell, start: cell },
  });
  try {
    await flush(120);
    click("xlsx-chart-insert");
    const [, input] = mounted.calls[0];
    assert.equal(input.range, undefined, "the controller falls back to the used range");
  } finally {
    mounted.cleanup();
  }
});

test("the chart dialog reports a rejected insertion instead of failing silently", async () => {
  const mounted = mountDialog("chart", { insertChart: () => null });
  try {
    await flush(120);
    click("xlsx-chart-insert");
    await flush();
    assert.match(byTestId("xlsx-authoring-status").textContent, /Nothing to change/);
  } finally {
    mounted.cleanup();
  }
});

test("the chart dialog edits and deletes the selected chart", async () => {
  const chart = {
    anchor: {
      from: { col: 0, colOffsetEmu: 0, row: 0, rowOffsetEmu: 0 },
      kind: "one-cell",
      sizeEmu: { cx: 1, cy: 1 },
    },
    chartType: "Line",
    dataLabels: { showValue: false },
    id: "chart-1-0",
    title: "Old title",
  };
  const mounted = mountDialog("chart", {
    getChartById: (id) => (id === chart.id ? chart : null),
    selectedChartId: chart.id,
  });
  try {
    await flush(120);
    assert.equal(byTestId("xlsx-chart-title").value, "Old title");
    setInput("xlsx-chart-title", "New title");
    setCheckbox("xlsx-chart-data-labels", true);
    await flush();
    click("xlsx-chart-update");
    assert.deepEqual(mounted.calls[0], [
      "updateChart",
      "chart-1-0",
      { dataLabels: { showValue: true }, title: "New title" },
    ]);
    click("xlsx-chart-delete");
    assert.deepEqual(mounted.calls[1], ["deleteChart", "chart-1-0"]);
  } finally {
    mounted.cleanup();
  }
});

test("the dialog keeps the shared-surface contract in source", async () => {
  const { readFileSync } = await import("node:fs");
  const source = readFileSync(
    new URL("../../packages/vue-xlsx/src/components/XlsxAuthoringDialog.vue", import.meta.url),
    "utf8",
  );
  assert.match(source, /<OfficeDialog/, "the dialog renders OfficeDialog");
  assert.match(source, /:close-test-id="'xlsx-authoring-close'"/, "the close button keeps its id");
  assert.match(source, /:content-test-id="'xlsx-authoring-dialog'"/, "the surface keeps its id");
  assert.equal(/<Teleport/.test(source), false, "no hand-rolled portal/backdrop remains");
  assert.equal(
    /office-overlay__backdrop xlsx-authoring-backdrop/.test(source),
    false,
    "the overlay comes from the shared dialog",
  );
});

test("the paste-special dialog applies the chosen mode and transpose flag", async () => {
  const mounts = [];
  const mounted = mountDialog("pasteSpecial", {
    pasteSpecialFromClipboard: async (options) => {
      mounts.push(options);
      return true;
    },
  });
  try {
    await flush(120);
    assert.ok(byTestId("xlsx-paste-special-section"));
    setSelect("xlsx-paste-special-mode", "values");
    setCheckbox("xlsx-paste-special-transpose", true);
    await flush();
    click("xlsx-paste-special-apply");
    await flush();
    assert.deepEqual(mounts, [{ mode: "values", transpose: true }]);
  } finally {
    mounted.cleanup();
  }
});

test("the number-format dialog builds decimal, separator and custom patterns", async () => {
  const formats = [];
  const mounted = mountDialog("numberFormat", {
    setSelectionNumberFormat: (format) => {
      formats.push(format);
      return true;
    },
  });
  try {
    await flush(120);
    setInput("xlsx-number-format-decimals", "0");
    setCheckbox("xlsx-number-format-thousands", false);
    await flush();
    assert.match(byTestId("xlsx-number-format-preview").textContent, /0/);
    click("xlsx-number-format-apply");
    assert.deepEqual(formats.at(-1), { formatString: "0", formatType: "custom" });

    setSelect("xlsx-number-format-category", "custom");
    await flush();
    setInput("xlsx-number-format-pattern", '0.0"%"');
    await flush();
    click("xlsx-number-format-apply");
    assert.deepEqual(formats.at(-1), { formatString: '0.0"%"', formatType: "custom" });
  } finally {
    mounted.cleanup();
  }
});

test("the name manager lists, adds, updates and deletes defined names", async () => {
  let names = [{ name: "Existing", refersTo: "'Sheet1'!$A$1:$B$2" }];
  const calls = [];
  const mounted = mountDialog("names", {
    getNamedRanges: () => names,
    setNamedRange: (name, refersTo) => {
      calls.push(["set", name, refersTo]);
      names = [{ name, refersTo }];
      return true;
    },
    deleteNamedRange: (name) => {
      calls.push(["delete", name]);
      names = names.filter((entry) => entry.name !== name);
      return true;
    },
  });
  try {
    await flush(120);
    assert.ok(byTestId("xlsx-name-item"), "existing names are listed");
    click("xlsx-name-use-selection");
    await flush();
    assert.equal(
      byTestId("xlsx-name-refers-to").value,
      "'Sheet1'!$B$3:$D$5",
      "a name can point at the current selection",
    );
    setInput("xlsx-name-input", "Revenue");
    setInput("xlsx-name-refers-to", "'Sheet1'!$C$1:$C$3");
    await flush();
    click("xlsx-name-save");
    await flush();
    assert.deepEqual(calls[0], ["set", "Revenue", "'Sheet1'!$C$1:$C$3"]);
    assert.ok(names.some((entry) => entry.name === "Revenue"));
    click("xlsx-name-delete");
    await flush();
    assert.deepEqual(calls[1], ["delete", "Revenue"]);
  } finally {
    mounted.cleanup();
  }
});
