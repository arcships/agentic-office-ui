// Issue #178 — CSV viewer behaviour.
//
// The upstream Extend block parses with papaparse and renders a Glide canvas
// grid. This Vue surface keeps the observable behaviour (header row, RFC 4180
// quoting, greedy blank-line skipping, case-insensitive cell search) while
// rendering a virtualized semantic table on the shared Office chrome. These
// tests pin the parser, the toolbar, the empty/error/loading states, upload,
// download, zoom and the search navigation contract.

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test, { after } from "node:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register({ url: "http://localhost/" });

const { importFromDemo, vue, waitFor } = await import(
  "./vue-test-renderer.mjs"
);

const {
  CsvViewer,
  csvCellAddress,
  decodeCsvBuffer,
  findCsvMatches,
  parseCsvText,
  resolveCsvVisibleRange,
  serializeCsv,
} = await importFromDemo("@arcships/vue-ui");

const FIXTURE = readFileSync(
  new URL("../../test-data/csv/sales-pipeline.csv", import.meta.url),
);

const SAMPLE = [
  "Name,Team,Amount,Notes",
  'Alice,Alpha,"1,200",First row',
  "Bob,Beta,900,Second row",
  "Carla,Alpha,1500,Third row",
  "Dmitri,Gamma,300,Fourth row",
  "Alice,Delta,450,Fifth row",
].join("\n");

after(() => {
  GlobalRegistrator.unregister();
});

function byTestId(root, testId) {
  return root.querySelector(`[data-testid="${testId}"]`);
}

function rowsOf(root) {
  return [...root.querySelectorAll("tr[data-csv-row]")];
}

function cellAt(root, row, col) {
  return root.querySelector(`[data-csv-cell="${row}:${col}"]`);
}

function changeFile(input, file) {
  Object.defineProperty(input, "files", {
    configurable: true,
    value: [file],
  });
  input.dispatchEvent(new Event("change", { bubbles: true }));
}

async function mountDom(component, props = {}) {
  const host = document.createElement("div");
  document.body.append(host);
  const warnings = [];
  const componentRef = vue.ref(null);
  const app = vue.createApp({
    render: () => vue.h(component, { ...props, ref: componentRef }),
  });
  app.config.warnHandler = (message) => warnings.push(message);
  app.mount(host);
  await vue.nextTick();
  return {
    app,
    host,
    warnings,
    component: componentRef,
    cleanup() {
      app.unmount();
      host.remove();
      document.body.innerHTML = "";
    },
  };
}

test("parseCsvText handles BOM, quotes, newlines, CRLF and uneven records", () => {
  const bom = "\ufeffName,Role,Note\r\n";
  const parsed = parseCsvText(
    `${bom}"Ada, Lovelace",Engineer,"Line one\nLine two"\r\nGrace,Admiral,"""quoted"""\r\nAlan,Turing`,
  );
  assert.equal(parsed.error, null);
  assert.equal(parsed.delimiter, ",");
  assert.deepEqual(parsed.headers, ["Name", "Role", "Note"]);
  assert.equal(parsed.rows.length, 3);
  assert.deepEqual(parsed.rows[0], ["Ada, Lovelace", "Engineer", "Line one\nLine two"]);
  assert.deepEqual(parsed.rows[1], ["Grace", "Admiral", '"quoted"']);
  // Missing trailing cells are padded so the table always has a rectangle.
  assert.deepEqual(parsed.rows[2], ["Alan", "Turing", ""]);
  assert.equal(parsed.columnCount, 3);
  assert.equal(parsed.sourceHeaderCount, 3);
});

test("parseCsvText detects TSV/semicolon/pipe and reports unclosed quotes", () => {
  const tsv = parseCsvText("A\tB\tC\n1\t2\t3");
  assert.equal(tsv.delimiter, "\t");
  assert.deepEqual(tsv.headers, ["A", "B", "C"]);

  const semicolon = parseCsvText("A;B\n1;2");
  assert.equal(semicolon.delimiter, ";");

  const pipe = parseCsvText("A|B\n1|2");
  assert.equal(pipe.delimiter, "|");

  const explicit = parseCsvText("A;B\n1;2", { delimiter: ";" });
  assert.equal(explicit.delimiter, ";");

  const invalid = parseCsvText("A;B", { delimiter: ";;" });
  assert.equal(invalid.error?.code, "INVALID_DELIMITER");

  const broken = parseCsvText('A,B\n"unclosed,2');
  assert.equal(broken.error?.code, "UNTERMINATED_QUOTE");
  assert.deepEqual(broken.rows, []);
});

test("decodeCsvBuffer reads UTF-8 BOM, UTF-16LE and UTF-16BE", () => {
  const utf8 = Uint8Array.from([0xef, 0xbb, 0xbf, 0x41, 0x2c, 0x42]);
  assert.equal(decodeCsvBuffer(utf8), "A,B");

  const utf16le = Uint8Array.from([0xff, 0xfe, 0x41, 0x00, 0x2c, 0x00, 0x42, 0x00]);
  assert.equal(decodeCsvBuffer(utf16le), "A,B");

  const utf16be = Uint8Array.from([0xfe, 0xff, 0x00, 0x41, 0x00, 0x2c, 0x00, 0x42]);
  assert.equal(decodeCsvBuffer(utf16be), "A,B");
});

test("serializeCsv quotes separators, quotes and newlines and round-trips", () => {
  const rows = [
    ["Ada, Lovelace", 'say "hi"', "line one\nline two"],
    ["plain", "", "semi;colon"],
  ];
  const text = serializeCsv(["Name", "Quote", "Note"], rows);
  assert.equal(
    text,
    'Name,Quote,Note\r\n"Ada, Lovelace","say ""hi""","line one\nline two"\r\nplain,,semi;colon',
  );
  const reparsed = parseCsvText(text);
  assert.deepEqual(reparsed.rows, rows);
});

test("search helpers find every case-insensitive cell and compute A1 addresses", async () => {
  const parsed = parseCsvText(SAMPLE);
  const matches = await findCsvMatches(parsed.headers, parsed.rows, "alice", {
    yieldBetweenBatches: false,
  });
  // The spreadsheet address is relative to the grid row index (header excluded),
  // the same coordinate the upstream Glide selection uses.
  assert.deepEqual(
    matches.map((match) => match.address),
    ["A1", "A5"],
  );
  assert.equal(csvCellAddress(0, 0), "A1");
  assert.equal(csvCellAddress(3, 27), "AB4");

  const batched = await findCsvMatches(parsed.headers, parsed.rows, "alpha", {
    batchSize: 1,
  });
  assert.deepEqual(
    batched.map((match) => match.address),
    ["B1", "B3"],
  );

  assert.deepEqual(await findCsvMatches(parsed.headers, parsed.rows, "  "), []);
});

test("resolveCsvVisibleRange keeps an overscanned window and spacer heights", () => {
  const range = resolveCsvVisibleRange({
    scrollTop: 340,
    viewportHeight: 340,
    rowCount: 100,
    rowHeight: 34,
    overscan: 2,
  });
  assert.deepEqual(range, {
    start: 8,
    end: 22,
    padTop: 272,
    padBottom: 2652,
  });

  assert.deepEqual(
    resolveCsvVisibleRange({
      scrollTop: 0,
      viewportHeight: 0,
      rowCount: 0,
      rowHeight: 34,
    }),
    { start: 0, end: 0, padTop: 0, padBottom: 0 },
  );
});

test("the fixture keeps the BOM, CJK copy and quoted cells the viewer must read", () => {
  const parsed = parseCsvText(decodeCsvBuffer(FIXTURE));
  assert.equal(parsed.error, null);
  assert.deepEqual(parsed.headers, [
    "Opportunity",
    "Account",
    "Amount",
    "Stage",
    "Owner",
    "Notes",
  ]);
  assert.ok(parsed.rows.length >= 8);
  assert.ok(FIXTURE.subarray(0, 3).equals(Buffer.from([0xef, 0xbb, 0xbf])));
  assert.equal(parsed.rows[0][1], "Acme, Inc.");
  assert.match(parsed.rows[0][5], /Follow-up in Q1/);
  assert.ok(parsed.rows.some((row) => row[1] === "上海远航"));
});

test("CsvViewer renders the 560px surface, toolbar and semantic grid", async () => {
  const mounted = await mountDom(CsvViewer, { data: SAMPLE });
  try {
    const root = byTestId(mounted.host, "csv-viewer");
    assert.ok(root, "the viewer exposes a stable test id");
    assert.equal(root.getAttribute("data-status"), "ready");
    assert.equal(root.getAttribute("data-row-count"), "5");
    assert.equal(root.getAttribute("data-column-count"), "4");
    assert.ok(root.classList.contains("office-ui"));
    assert.ok(root.classList.contains("office-csv-viewer"));
    assert.equal(root.style.getPropertyValue("--csv-viewer-height"), "");

    const toolbar = byTestId(mounted.host, "office-toolbar");
    assert.ok(toolbar, "the shared Office toolbar is reused");
    assert.ok(toolbar.classList.contains("office-toolbar"));
    assert.ok(byTestId(mounted.host, "office-zoom-control"));
    assert.ok(byTestId(mounted.host, "csv-actions-trigger"));
    assert.ok(byTestId(mounted.host, "csv-search-trigger"));

    const grid = byTestId(mounted.host, "csv-grid");
    assert.equal(grid.getAttribute("role"), "grid");
    assert.equal(grid.getAttribute("aria-rowcount"), "6");
    assert.equal(grid.getAttribute("aria-colcount"), "5");

    const headers = [...mounted.host.querySelectorAll('th[role="columnheader"]')];
    assert.deepEqual(
      headers.slice(1).map((node) => node.textContent.trim()),
      ["Name", "Team", "Amount", "Notes"],
    );

    assert.equal(rowsOf(mounted.host).length, 5);
    assert.equal(cellAt(mounted.host, 1, 0)?.textContent.trim(), "Bob");
    assert.equal(cellAt(mounted.host, 0, 2)?.textContent.trim(), "1,200");
    assert.equal(
      byTestId(mounted.host, "csv-meta")?.textContent.trim(),
      "5 rows × 4 columns",
    );
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.cleanup();
  }
});

test("CsvViewer honours the search prop, height override and host class", async () => {
  const mounted = await mountDom(CsvViewer, {
    data: SAMPLE,
    search: false,
    height: 420,
    class: "host-csv",
  });
  try {
    const root = byTestId(mounted.host, "csv-viewer");
    assert.ok(root.classList.contains("host-csv"));
    assert.equal(root.style.blockSize, "420px");
    assert.equal(byTestId(mounted.host, "csv-search-trigger"), null);
  } finally {
    mounted.cleanup();
  }
});

test("CsvViewer localizes chrome copy through the shared catalog", async () => {
  const mounted = await mountDom(CsvViewer, {
    data: SAMPLE,
    locale: "zh-CN",
  });
  try {
    assert.equal(
      byTestId(mounted.host, "csv-meta")?.textContent.trim(),
      "5 行 × 4 列",
    );
    assert.equal(
      byTestId(mounted.host, "csv-search-trigger")?.getAttribute("aria-label"),
      "搜索 CSV",
    );
    assert.equal(
      byTestId(mounted.host, "csv-actions-trigger")?.getAttribute("aria-label"),
      "打开 CSV 操作",
    );
  } finally {
    mounted.cleanup();
  }
});

test("selection, zoom and download stay on the public contract", async () => {
  const downloads = [];
  const mounted = await mountDom(CsvViewer, {
    data: SAMPLE,
    fileName: "pipeline",
    onDownload: (event) => downloads.push(event),
  });
  try {
    const viewer = mounted.component.value;
    viewer.selectCell(1, 2);
    await vue.nextTick();
    const cell = cellAt(mounted.host, 1, 2);
    assert.equal(cell.getAttribute("aria-selected"), "true");
    assert.ok(cell.classList.contains("office-csv-viewer__cell--selected"));

    viewer.setZoom(150);
    await vue.nextTick();
    assert.equal(viewer.getState().zoom, 150);

    viewer.setZoom(500);
    assert.equal(viewer.getState().zoom, 200);
    viewer.zoomOut();
    assert.equal(viewer.getState().zoom, 175);

    viewer.download();
    assert.equal(downloads.length, 1);
    assert.equal(downloads[0].fileName, "pipeline.csv");
    assert.match(downloads[0].text, /^Name,Team,Amount,Notes/);
    assert.match(downloads[0].text, /"1,200"/);
  } finally {
    mounted.cleanup();
  }
});

test("search returns matches, selects the active cell and wraps navigation", async () => {
  const mounted = await mountDom(CsvViewer, {
    data: SAMPLE,
    search: true,
    searchDebounce: 0,
  });
  try {
    const viewer = mounted.component.value;
    await viewer.runSearch("alice");
    await waitFor(
      () => viewer.getState().searchResults.length === 2,
      "two Alice matches",
    );
    assert.deepEqual(
      viewer.getState().searchResults.map((match) => match.address),
      ["A1", "A5"],
    );
    assert.equal(viewer.getState().activeResultIndex, 0);
    await vue.nextTick();
    assert.equal(cellAt(mounted.host, 0, 0).getAttribute("aria-selected"), "true");
    const marked = [...mounted.host.querySelectorAll(".office-csv-viewer__match")];
    assert.ok(marked.length >= 2, "every visible match is highlighted");
    assert.equal(marked[0].textContent, "Alice");

    viewer.nextResult();
    await vue.nextTick();
    assert.equal(viewer.getState().activeResultIndex, 1);
    assert.equal(cellAt(mounted.host, 4, 0).getAttribute("aria-selected"), "true");

    viewer.nextResult();
    assert.equal(viewer.getState().activeResultIndex, 0);
    viewer.previousResult();
    assert.equal(viewer.getState().activeResultIndex, 1);

    viewer.clearSearch();
    await vue.nextTick();
    assert.equal(viewer.getState().searchQuery, "");
    assert.deepEqual(viewer.getState().searchResults, []);
    assert.equal(viewer.getState().activeResultIndex, 0);
  } finally {
    mounted.cleanup();
  }
});

test("upload replaces the parsed table and emits update:data/load", async () => {
  const updates = [];
  const loads = [];
  const mounted = await mountDom(CsvViewer, {
    data: SAMPLE,
    "onUpdate:data": (value) => updates.push(value),
    onLoad: (event) => loads.push(event),
  });
  try {
    const input = byTestId(mounted.host, "csv-file-input");
    assert.ok(input);
    const file = new File(
      ['City,Country\n"Shanghai, CN",China\nOslo,Norway'],
      "cities.csv",
      { type: "text/csv" },
    );
    changeFile(input, file);
    await waitFor(
      () => byTestId(mounted.host, "csv-viewer")?.getAttribute("data-row-count") === "2",
      "uploaded rows",
    );
    const root = byTestId(mounted.host, "csv-viewer");
    assert.equal(root.getAttribute("data-file-name"), "cities.csv");
    assert.equal(updates.at(-1), 'City,Country\n"Shanghai, CN",China\nOslo,Norway');
    assert.equal(loads.at(-1).fileName, "cities.csv");
    assert.deepEqual(loads.at(-1).headers, ["City", "Country"]);
    assert.equal(cellAt(mounted.host, 0, 0)?.getAttribute("title"), "Shanghai, CN");
  } finally {
    mounted.cleanup();
  }
});

test("upload size limit surfaces the localized FILE_TOO_LARGE error", async () => {
  const errors = [];
  const mounted = await mountDom(CsvViewer, {
    data: SAMPLE,
    maxFileSize: 4,
    onError: (error) => errors.push(error),
  });
  try {
    const input = byTestId(mounted.host, "csv-file-input");
    changeFile(input, new File(["Name\nAlice"], "too-big.csv", { type: "text/csv" }));
    await waitFor(
      () => byTestId(mounted.host, "csv-viewer")?.getAttribute("data-status") === "error",
      "size error",
    );
    assert.equal(errors.length, 1);
    assert.equal(errors[0].code, "FILE_TOO_LARGE");
    const errorState = byTestId(mounted.host, "csv-error");
    assert.ok(errorState);
    assert.match(errorState.textContent, /exceeds the allowed maximum/);
  } finally {
    mounted.cleanup();
  }
});

test("empty, error and external loading states use the shared state components", async () => {
  const empty = await mountDom(CsvViewer, {});
  try {
    assert.ok(byTestId(empty.host, "csv-empty"));
    assert.equal(byTestId(empty.host, "csv-viewer")?.getAttribute("data-status"), "empty");
    assert.ok(byTestId(empty.host, "csv-empty-upload"));
  } finally {
    empty.cleanup();
  }

  const brokenErrors = [];
  const broken = await mountDom(CsvViewer, {
    data: 'A,B\n"unclosed,2',
    onError: (error) => brokenErrors.push(error),
  });
  try {
    assert.ok(byTestId(broken.host, "csv-error"));
    assert.match(byTestId(broken.host, "csv-error").textContent, /not closed/);
    assert.equal(broken.component.value.getState().error.code, "UNTERMINATED_QUOTE");
    assert.equal(brokenErrors.at(-1)?.code, "UNTERMINATED_QUOTE");
  } finally {
    broken.cleanup();
  }

  const loading = await mountDom(CsvViewer, { data: SAMPLE, loading: true });
  try {
    assert.ok(byTestId(loading.host, "csv-loading"));
    assert.equal(byTestId(loading.host, "csv-viewer")?.getAttribute("data-status"), "loading");
  } finally {
    loading.cleanup();
  }
});

test("a file prop decodes UTF-16 buffers before parsing", async () => {
  const text = "Name,City\nAda,London";
  const utf16 = new Uint8Array(2 + text.length * 2);
  utf16[0] = 0xff;
  utf16[1] = 0xfe;
  for (let index = 0; index < text.length; index += 1) {
    utf16[2 + index * 2] = text.charCodeAt(index) & 0xff;
    utf16[3 + index * 2] = text.charCodeAt(index) >> 8;
  }
  const file = new File([utf16], "utf16.csv", { type: "text/csv" });
  const mounted = await mountDom(CsvViewer, { file });
  try {
    await waitFor(
      () => byTestId(mounted.host, "csv-viewer")?.getAttribute("data-status") === "ready",
      "UTF-16 file",
    );
    assert.equal(cellAt(mounted.host, 0, 1)?.textContent.trim(), "London");
  } finally {
    mounted.cleanup();
  }
});

test("a src prop fetches and parses remote CSV text", async () => {
  const originalFetch = globalThis.fetch;
  const calls = [];
  globalThis.fetch = async (url) => {
    calls.push(String(url));
    return new Response(SAMPLE, {
      status: 200,
      headers: { "content-type": "text/csv" },
    });
  };
  const mounted = await mountDom(CsvViewer, {
    src: "https://example.test/data/sales.csv",
  });
  try {
    await waitFor(
      () => byTestId(mounted.host, "csv-viewer")?.getAttribute("data-status") === "ready",
      "remote CSV",
    );
    assert.deepEqual(calls, ["https://example.test/data/sales.csv"]);
    assert.equal(
      byTestId(mounted.host, "csv-viewer")?.getAttribute("data-file-name"),
      "sales.csv",
    );
    assert.equal(rowsOf(mounted.host).length, 5);
  } finally {
    mounted.cleanup();
    globalThis.fetch = originalFetch;
  }
});

test("large tables render an overscanned window and reveal searched rows", async () => {
  const rows = Array.from({ length: 2000 }, (_, index) => `Row ${index},${index * 2}`);
  const mounted = await mountDom(CsvViewer, {
    data: ["Name,Value", ...rows].join("\n"),
    search: true,
    searchDebounce: 0,
  });
  try {
    const viewer = mounted.component.value;
    const initialRows = rowsOf(mounted.host);
    assert.ok(initialRows.length < 60, `only the visible window is mounted (${initialRows.length})`);
    assert.ok(initialRows.length > 5);
    assert.equal(initialRows[0].getAttribute("data-csv-row"), "0");

    await viewer.runSearch("Row 1999");
    await waitFor(
      () => viewer.getState().searchResults.length === 1,
      "last-row search",
    );
    await waitFor(
      () => rowsOf(mounted.host).some((row) => row.getAttribute("data-csv-row") === "1999"),
      "last row in the window",
    );
    assert.equal(cellAt(mounted.host, 1999, 0)?.getAttribute("aria-selected"), "true");
    assert.ok(mounted.host.querySelectorAll(".office-csv-viewer__spacer").length >= 1);
  } finally {
    mounted.cleanup();
  }
});
