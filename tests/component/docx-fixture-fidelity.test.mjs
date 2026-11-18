// Issue #53 — rendered-level assertions for the DOCX feature fixtures.
//
// The unit test (`tests/unit/docx-fixture-fidelity.test.mjs`) proves the
// fixtures import with threads, merges, sections and notes. This file proves
// the viewer *renders* them where they belong:
//
//   * one comment card per comment, on exactly one page, numbered like its
//     inline anchor marker,
//   * comment cards track their anchor within one line height instead of being
//     an unpositioned list (the "comments look scrambled" defect),
//   * a deleted table row is visibly marked and listed once, and a row-height
//     revision is reported against its own row,
//   * merged/nested tables keep their column count and cell text,
//   * the header and footer of a section render on every page of that section,
//   * a footnote marker label equals the label in the notes section.

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";

import { importFromDemo, mount, walk } from "./vue-test-renderer.mjs";

const { DocxTableHost, DocxViewerRoot } = await importFromDemo("@arcships/vue-docx");
const core = await importFromDemo("@arcships/docx-core");

const SAMPLES = new URL("../../apps/demo/public/samples/", import.meta.url);
/** One line of body text at the fixture's 10.5pt font size. */
const LINE_HEIGHT_PX = 24;

function fixtureBuffer(name) {
  const bytes = readFileSync(new URL(name, SAMPLES));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

const fixtureCache = new Map();
async function loadFixture(name) {
  if (!fixtureCache.has(name)) {
    fixtureCache.set(
      name,
      core.importDocxBuffer(fixtureBuffer(name), { useWorker: false }),
    );
  }
  return fixtureCache.get(name);
}

function dataNodes(root, attribute) {
  return walk(root).filter((node) => node.props?.[attribute] !== undefined);
}

function nodesWithClass(root, className) {
  return walk(root).filter((node) =>
    String(node.props?.class || "")
      .split(/\s+/)
      .includes(className),
  );
}

function htmlText(root) {
  return walk(root)
    .map((node) => node.props?.innerHTML ?? node.text ?? "")
    .join("");
}

/** Visible text of the rendered HTML string held by `htmlText`. */
function plainHtmlText(root) {
  return htmlText(root).replace(/<[^>]*>/g, "").trim();
}

/** Extracts `<span data-docx-note-marker="true">label</span>` labels in order. */
function noteMarkerLabels(root) {
  const labels = [];
  for (const node of walk(root)) {
    const html = node.props?.innerHTML;
    if (typeof html !== "string" || !html.includes("data-docx-note-marker")) continue;
    for (const match of html.matchAll(
      /data-docx-note-marker="true"[^>]*>([^<]*)</g,
    )) {
      labels.push(match[1]);
    }
  }
  return labels;
}

function noteBodyLabels(root, kind) {
  const section = walk(root).find(
    (node) => node.props?.["data-testid"] === `docx-${kind}s`,
  );
  assert.ok(section, `the ${kind} section renders`);
  return walk(section)
    .filter((node) => node.props?.["data-docx-note-body-marker"] !== undefined)
    .map((node) => {
      const children = walk(node).filter((child) => child.type === "#text");
      return children.map((child) => child.text).join("");
    });
}

function pageWrappers(root) {
  return dataNodes(root, "data-docx-page-wrapper").sort(
    (left, right) =>
      Number(left.props["data-docx-page-index"]) -
      Number(right.props["data-docx-page-index"]),
  );
}

function cardsOnPage(page) {
  return nodesWithClass(page, "docx-gutter-card");
}

/** Rows of the outermost rendered table, ignoring nested tables. */
function tbodyRows(root) {
  const table = walk(root).find((node) => node.type === "table");
  if (!table) return [];
  const tbody = table.children.find((node) => node.type === "tbody");
  return tbody ? tbody.children.filter((node) => node.type === "tr") : [];
}

function withContainerWidth(t, widthPx) {
  const originalWidth = Object.getOwnPropertyDescriptor(
    Object.prototype,
    "clientWidth",
  );
  const originalHeight = Object.getOwnPropertyDescriptor(
    Object.prototype,
    "clientHeight",
  );
  Object.defineProperty(Object.prototype, "clientWidth", {
    configurable: true,
    get: () => widthPx,
  });
  // The viewer only renders the pages inside its scroll viewport. A tall fake
  // viewport keeps every page mounted so multi-page assertions can see them.
  Object.defineProperty(Object.prototype, "clientHeight", {
    configurable: true,
    get: () => 100_000,
  });
  t.after(() => {
    if (originalWidth) {
      Object.defineProperty(Object.prototype, "clientWidth", originalWidth);
    } else {
      delete Object.prototype.clientWidth;
    }
    if (originalHeight) {
      Object.defineProperty(Object.prototype, "clientHeight", originalHeight);
    } else {
      delete Object.prototype.clientHeight;
    }
  });
}

async function mountViewer(t, fixture, overrides = {}) {
  withContainerWidth(t, 1440);
  const { model } = await loadFixture(fixture);
  const mounted = await mount(DocxViewerRoot, {
    model,
    zoomScale: 100,
    showComments: true,
    showTrackedChanges: true,
    ...overrides,
  });
  t.after(() => mounted.app.unmount());
  return mounted;
}

// ── Comment placement ──────────────────────────────────────────────
test("every visible comment renders exactly one card across the document's pages", async (t) => {
  const mounted = await mountViewer(t, "review-threads.docx");
  const { model } = await loadFixture("review-threads.docx");
  const comments = core.collectCommentsFromModel(model);
  const commentsById = new Map(comments.map((comment) => [comment.commentId, comment]));
  const hiddenResolved = comments
    .filter((comment) => core.commentThreadIsResolved(comment, commentsById))
    .map((comment) => comment.commentId);
  assert.deepEqual(hiddenResolved.sort((a, b) => a - b), [8, 9], "the resolved thread is hidden");

  const pages = pageWrappers(mounted.root);
  assert.ok(pages.length >= 1, "the fixture renders at least one page");

  const cardsByPage = pages.map((page) => cardsOnPage(page));
  const commentCardIds = cardsByPage.flatMap((cards) =>
    cards
      .map((card) => card.props["data-docx-gutter-card-id"])
      .filter((id) => String(id).startsWith("comment:")),
  );
  const expected = comments
    .filter((comment) => !hiddenResolved.includes(comment.commentId))
    .map((comment) => `comment:${comment.id}`)
    .sort();
  assert.deepEqual(
    [...commentCardIds].sort(),
    expected,
    "one card per visible comment, no duplicate cards across pages",
  );
  assert.equal(
    new Set(commentCardIds).size,
    commentCardIds.length,
    "a comment must not be listed on two pages",
  );

  // Cards follow document order, not author/id order: the thread root comes
  // before its replies, and a later paragraph's comment comes after them.
  const bodyCommentOrder = commentCardIds.filter(
    (id) => !String(id).includes("header:"),
  );
  const expectedBodyOrder = comments
    .filter(
      (comment) =>
        !hiddenResolved.includes(comment.commentId) && comment.region === "body",
    )
    .map((comment) => `comment:${comment.id}`);
  assert.deepEqual(
    bodyCommentOrder,
    expectedBodyOrder,
    "body comment cards are listed in document order",
  );

  // A card's ordinal and its inline marker number come from the same
  // annotation list, so they identify the same comment.
  const markerById = new Map();
  for (const page of pages) {
    for (const marker of dataNodes(page, "data-docx-comment-marker")) {
      const id = marker.props["data-docx-gutter-marker-id"];
      assert.equal(markerById.has(id), false, `${id} has more than one marker`);
      markerById.set(id, marker.props["data-docx-comment-marker"]);
    }
  }
  for (const cards of cardsByPage) {
    for (const card of cards) {
      const id = card.props["data-docx-gutter-card-id"];
      if (!String(id).startsWith("comment:")) continue;
      assert.equal(
        markerById.get(id),
        card.props["data-docx-gutter-card-index"],
        `${id}: the marker number matches the card number`,
      );
    }
  }
  assert.deepEqual(mounted.warnings, []);
});

test("comment cards track their anchors within one line instead of a fixed list", async () => {
  const { model } = await loadFixture("review-threads.docx");
  const snapshot = core.buildLayoutSnapshot(model, {});
  const pageSegments = snapshot.pages.map((page) =>
    page.blocks
      .map((block) => ({ nodeIndex: block.source?.nodeIndex }))
      .filter((segment) => Number.isFinite(segment.nodeIndex)),
  );
  const comments = core.collectCommentsFromModel(model);
  const commentsById = new Map(comments.map((comment) => [comment.commentId, comment]));
  const visible = comments.filter(
    (comment) => !core.commentThreadIsResolved(comment, commentsById),
  );

  let checkedBodyAnchors = 0;
  for (const [pageIndex, page] of snapshot.pages.entries()) {
    const annotations = core.buildPageGutterAnnotations({
      pageSegments: pageSegments[pageIndex],
      comments: visible,
      headerPartName: "word/header1.xml",
    }).filter((annotation) => annotation.comment);
    if (annotations.length === 0) continue;

    const anchors = new Map();
    for (const annotation of annotations) {
      const anchorNodeIndex = annotation.location?.nodeIndex;
      const block = page.blocks.find(
        (candidate) => candidate.source?.nodeIndex === anchorNodeIndex,
      );
      if (!block) continue;
      anchors.set(annotation.id, {
        x: block.framePx.x,
        y: block.framePx.y,
      });
    }
    if (anchors.size === 0) continue;

    const positioned = core.layoutTrackedChangesForPage(
      annotations,
      anchors,
      undefined,
      page.pageSizePx.width,
      page.pageSizePx.height,
    );

    let previousBottom = 8;
    for (const entry of positioned) {
      const anchorY = anchors.get(entry.annotation.id)?.y;
      if (anchorY === undefined) continue;
      checkedBodyAnchors += 1;
      // A card starts at its anchor; only a collision with the card above may
      // push it down, and never above its anchor.
      const desiredTop = Math.max(8, anchorY - 8);
      assert.ok(
        entry.top >= desiredTop - 1,
        `${entry.annotation.id}: card top ${entry.top} sits above its anchor ${anchorY}`,
      );
      assert.ok(
        entry.top <= desiredTop + LINE_HEIGHT_PX ||
          entry.top >= previousBottom,
        `${entry.annotation.id}: card top ${entry.top} is not within one line of anchor ${anchorY}`,
      );
      previousBottom = entry.top + entry.heightPx;
    }

    const ordered = positioned
      .filter((entry) => anchors.has(entry.annotation.id))
      .map((entry) => entry.anchorY);
    assert.deepEqual(
      ordered,
      [...ordered].sort((left, right) => left - right),
      "cards are laid out in anchor order",
    );
  }
  assert.ok(
    checkedBodyAnchors >= 5,
    `expected several body comment anchors, checked ${checkedBodyAnchors}`,
  );

  // A header comment belongs to every page of its section, not to the body
  // paragraphs of one page. Its on-screen band is measured in a browser, so the
  // blackbox annotation suite asserts that geometry; here the annotation must
  // still be owned by the first page instead of being dropped.
  const headerComment = visible.find((comment) => comment.region === "header");
  assert.ok(headerComment, "the fixture has a header comment");
  const firstPageAnnotations = core.buildPageGutterAnnotations({
    pageSegments: pageSegments[0],
    comments: visible,
    headerPartName: "word/header1.xml",
  });
  assert.ok(
    firstPageAnnotations.some(
      (annotation) => annotation.comment?.commentId === headerComment.commentId,
    ),
    "the header comment is owned by the first page",
  );
  // A page without the header part must not claim it.
  const secondPageAnnotations = core.buildPageGutterAnnotations({
    pageSegments: pageSegments[0],
    comments: visible,
  });
  assert.equal(
    secondPageAnnotations.some(
      (annotation) => annotation.comment?.commentId === headerComment.commentId,
    ),
    false,
    "a page without the header part keeps the header comment out of its gutter",
  );
});

// ── Table revisions ────────────────────────────────────────────────
test("a deleted row is marked once and a row-height revision is listed against its own row", async (t) => {
  const mounted = await mountViewer(t, "review-threads.docx");
  const pages = pageWrappers(mounted.root);

  const deletedRows = pages.flatMap((page) =>
    walk(page).filter(
      (node) => node.props?.["data-docx-tracked-change"] === "deletion",
    ),
  );
  assert.equal(deletedRows.length, 1, "exactly one row is marked as deleted");
  assert.equal(
    Number(deletedRows[0].props["data-docx-table-row-index"]),
    2,
    "the deleted row is the third row, not the header or the first row",
  );
  assert.ok(
    String(deletedRows[0].props.class).includes("docx-table-row--deletion"),
    "the deleted row carries its visual marker class",
  );
  assert.ok(
    htmlText(deletedRows[0]).includes("DELETED_ROW_TARGET"),
    "the deleted row keeps its text for review instead of disappearing",
  );

  const changeCards = pages
    .flatMap((page) => cardsOnPage(page))
    .filter((card) => String(card.props["data-docx-gutter-card-id"]).startsWith("change:"));
  const deletionCards = changeCards.filter((card) =>
    String(card.props["data-docx-gutter-card-id"]).includes("row-deletion"),
  );
  assert.equal(deletionCards.length, 1, "the deleted row is listed exactly once");
  const heightCards = changeCards.filter((card) =>
    String(card.props["data-docx-gutter-card-id"]).includes("row-format-change"),
  );
  assert.equal(heightCards.length, 1, "the row-height revision is listed exactly once");
  assert.notEqual(
    deletionCards[0].props["data-docx-gutter-card-id"],
    heightCards[0].props["data-docx-gutter-card-id"],
  );

  // The revision records resolve to their own rows, so the cards cannot be
  // anchored to the first row by accident.
  const { model } = await loadFixture("review-threads.docx");
  const revisions = core
    .collectTrackedChangesFromModel(model)
    .filter((change) => change.kind === "deletion" || change.kind === "row-format-change");
  assert.deepEqual(
    revisions.map((change) => change.location.rowIndex).sort(),
    [2, 3],
  );
  assert.ok(
    heightCards[0].props["data-docx-gutter-card-id"].includes(
      revisions.find((change) => change.kind === "row-format-change").id.split(":").at(-2),
    ),
    "the row-height card points at the row-height revision",
  );
});

// ── Merged, nested and repeating tables ────────────────────────────
test("merged and nested tables keep their column count and cell text", async (t) => {
  const { model } = await loadFixture("table-structures.docx");
  const table = model.nodes.find(
    (node) => node.type === "table" && node.rows[1]?.cells[0]?.style?.rowSpan === 3,
  );
  assert.ok(table, "the fixture's merged table is imported");

  const mounted = await mount(DocxTableHost, {
    table,
    tableIndex: 0,
    model,
    showTrackedChanges: true,
  });
  t.after(() => mounted.app.unmount());

  const outerTable = walk(mounted.root).find((node) => node.type === "table");
  assert.ok(outerTable, "the table renders");
  const colgroup = walk(outerTable).find((node) => node.type === "colgroup");
  const columns = walk(colgroup).filter((node) => node.type === "col");
  assert.equal(columns.length, 4, "the merged table keeps four grid columns");

  const tbody = outerTable.children.find((node) => node.type === "tbody");
  const rows = tbody.children.filter((node) => node.type === "tr");
  const gridColumnCount = (row) =>
    row.children
      .filter((node) => node.type === "td")
      .reduce((total, cell) => total + Number(cell.props?.colspan ?? 1), 0);
  assert.equal(rows.length, 5, "the continuation rows still render their own cells");
  const continuationCells = (rowIndex) =>
    table.rows[rowIndex].cells.filter((cell) => cell.style?.vMergeContinuation === true)
      .length;
  for (const row of rows) {
    const rowIndex = Number(row.props?.["data-docx-table-row-index"]);
    assert.equal(
      gridColumnCount(row) + continuationCells(rowIndex),
      4,
      `row ${rowIndex} still covers four columns`,
    );
  }

  const headerCells = rows[0].children.filter((node) => node.type === "td");
  assert.equal(headerCells.length, 4, "the header row keeps four cells");
  const spannedRowCells = rows[1].children.filter((node) => node.type === "td");
  assert.equal(spannedRowCells.length, 3, "the gridSpan cell replaces two cells");
  assert.equal(spannedRowCells[1].props.colspan, 2, "the span survives rendering");
  assert.equal(
    spannedRowCells[0].props.rowspan,
    3,
    "the vertically merged cell spans three rows",
  );

  const rowTwoCells = rows[2].children.filter((node) => node.type === "td");
  assert.equal(
    rowTwoCells.length,
    3,
    "the vMerge continuation cell is dropped instead of leaving a hole",
  );
  assert.deepEqual(
    rowTwoCells.map((cell) => cell.props["data-docx-table-cell-index"]),
    [1, 2, 3],
    "the remaining cells keep their original cell indexes",
  );

  const text = htmlText(mounted.root);
  assert.equal(
    text.split("North region").length - 1,
    1,
    "the merged cell text is painted once",
  );
  for (const marker of ["Q1+Q2 combined", "1,200", "(audited)", "(signed off)", "5,300"]) {
    assert.ok(text.includes(marker), `cell text ${marker} survived rendering`);
  }

  const tableHosts = walk(mounted.root).filter(
    (node) => node.props?.["data-docx-table-host"] !== undefined,
  );
  assert.equal(tableHosts.length, 2, "the nested table renders as its own table host");
  const nestedHost = tableHosts.find(
    (host) =>
      String(host.props["data-docx-table-runtime-key"] ?? "").includes("nested-table"),
  );
  assert.ok(nestedHost, "the nested table host sits inside the cell");
  const nestedCell = walk(rows[3]).find(
    (node) => Number(node.props?.["data-docx-table-cell-index"]) === 3,
  );
  assert.ok(
    walk(nestedCell).includes(nestedHost),
    "the nested table is a descendant of its owning cell",
  );
  const nestedText = htmlText(nestedHost);
  for (const marker of ["N1", "N2", "N3", "N4"]) {
    assert.ok(nestedText.includes(marker), `nested cell ${marker} survived rendering`);
  }

  const multiParagraphCell = walk(rows[2]).find(
    (node) => Number(node.props?.["data-docx-table-cell-index"]) === 1,
  );
  assert.equal(
    dataNodes(multiParagraphCell, "data-docx-table-cell-paragraph-host").length,
    3,
    "the multi-paragraph cell renders all three paragraphs",
  );
  assert.deepEqual(mounted.warnings, []);
});

test("the long table repeats its header row on the continuation page", async () => {
  const { model } = await loadFixture("table-structures.docx");
  const snapshot = core.buildLayoutSnapshot(model, {});
  const longTableIndex = model.nodes.findIndex(
    (node) => node.type === "table" && node.rows.length === 46,
  );
  assert.ok(longTableIndex >= 0, "the long table is imported");

  const tableSlices = snapshot.pages.flatMap((page) =>
    page.blocks
      .filter((block) => block.source?.nodeIndex === longTableIndex)
      .map((block) => ({ pageIndex: page.index, range: block.tableRowRange })),
  );
  assert.ok(tableSlices.length >= 2, "the long table spans more than one page");
  assert.equal(tableSlices[0].range.startRowIndex, 0, "page one starts at the header row");
  const continuation = tableSlices.find((slice) => slice.range.startRowIndex > 0);
  assert.ok(continuation, "a continuation slice exists");
  assert.equal(
    continuation.range.repeatHeaderRowCount,
    1,
    "the continuation slice repeats exactly one header row",
  );

  // Render the continuation slice the way the page body does: the header row
  // appears above the slice's own rows.
  const table = model.nodes[longTableIndex];
  const mounted = await mount(DocxTableHost, {
    table,
    tableIndex: 0,
    model,
    rowRange: continuation.range,
  });
  try {
    const renderedRows = tbodyRows(mounted.root);
    assert.equal(
      Number(renderedRows[0].props["data-docx-table-row-index"]),
      0,
      "the repeated header row is rendered first",
    );
    const text = htmlText(renderedRows[0]);
    assert.ok(text.includes("Column 1"), "the repeated row is the real header row");
    assert.ok(
      renderedRows.length > 2,
      "the header row is repeated in addition to the slice rows",
    );
  } finally {
    mounted.app.unmount();
  }
});

test("a table inside the page header renders on every page through the table host", async (t) => {
  const mounted = await mountViewer(t, "table-structures.docx");
  const pages = pageWrappers(mounted.root);
  assert.equal(pages.length, 2, "the fixture renders two pages");

  for (const [pageIndex, page] of pages.entries()) {
    const header = dataNodes(page, "data-docx-page-header")[0];
    assert.ok(header, `page ${pageIndex} renders its header`);
    assert.ok(
      htmlText(header).includes("Table Structures Fixture"),
      `page ${pageIndex} keeps the running-head paragraph`,
    );

    const hosts = dataNodes(header, "data-docx-table-host");
    assert.equal(
      hosts.length,
      1,
      `page ${pageIndex} renders exactly one header table host`,
    );
    const table = walk(hosts[0]).find((node) => node.type === "table");
    assert.ok(table, `page ${pageIndex} renders a real table element`);
    const tbody = table.children.find((node) => node.type === "tbody");
    assert.ok(tbody, `page ${pageIndex} renders the header table body`);
    const rows = tbody.children.filter((node) => node.type === "tr");
    assert.deepEqual(
      rows.map((row) =>
        row.children
          .filter((node) => node.type === "td")
          .map((cell) => plainHtmlText(cell)),
      ),
      [
        ["Register", "Period", "Owner"],
        ["Table Register", "FY2026", "O. North"],
      ],
      `page ${pageIndex} keeps the header table's cells and text`,
    );

    // The header table is rendered inside the header band, not in the body
    // flow, so it cannot be mistaken for a body table.
    const bodyTables = dataNodes(page, "data-docx-table-host").filter(
      (node) => !walk(header).includes(node),
    );
    assert.ok(
      bodyTables.every((node) => !htmlText(node).includes("Table Register")),
      `page ${pageIndex} does not duplicate the header table in the body`,
    );
  }
  assert.deepEqual(mounted.warnings, []);
});

// ── Section headers, footers and columns ───────────────────────────
test("each section renders its own header and footer on every page of that section", async (t) => {
  const mounted = await mountViewer(t, "sections-columns.docx");
  const pages = pageWrappers(mounted.root);
  assert.equal(pages.length, 4, "the section fixture renders four pages");

  const headerText = (page) => {
    const header = dataNodes(page, "data-docx-page-header")[0];
    return header ? htmlText(header) : "";
  };
  const footerText = (page) => {
    const footer = dataNodes(page, "data-docx-page-footer")[0];
    return footer ? htmlText(footer) : "";
  };
  const headerHtml = (page) => {
    const header = dataNodes(page, "data-docx-page-header")[0];
    return header ? htmlText(header) : "";
  };

  assert.deepEqual(
    pages.map((page) => headerText(page).includes("Section One Header")),
    [true, true, false, false],
    "the first section's header repeats on every page of that section",
  );
  assert.deepEqual(
    pages.map((page) => headerText(page).includes("Section Two Landscape Header")),
    [false, false, true, false],
    "the landscape section carries its own header",
  );
  assert.deepEqual(
    pages.map((page) => headerText(page).includes("Section Three Two Column Header")),
    [false, false, false, true],
    "the two-column section carries its own header",
  );
  assert.deepEqual(
    pages.map((page) => footerText(page).includes("Section One Footer")),
    [true, true, false, false],
    "each section keeps its own footer",
  );
  assert.ok(footerText(pages[2]).includes("Section Two Footer"));
  assert.ok(footerText(pages[3]).includes("Section Three Footer"));

  // The landscape page is wider than the portrait pages.
  const surfaces = pages.map(
    (page) => dataNodes(page, "data-docx-page-surface")[0],
  );
  const widths = surfaces.map((surface) =>
    Number.parseFloat(surface.props.style.width),
  );
  assert.ok(widths[2] > widths[0], "the landscape section renders wider");
  assert.equal(widths[0], widths[1], "both portrait pages share one page size");
  assert.equal(widths[3], widths[0], "the two-column section is portrait again");

  // The watermark lives in the first section's header only, and it must paint
  // behind the text (a negative z-index) on every page of that section.
  const watermarkCounts = pages.map((page) => {
    const html = headerHtml(page);
    return [...html.matchAll(/<img\b[^>]*>/g)].filter(
      (match) =>
        match[0].includes('data-docx-image-layout="absolute"') &&
        /z-index:-\d+/.test(match[0]),
    ).length;
  });
  assert.equal(watermarkCounts[0], 1, "the first section header renders the watermark");
  assert.equal(watermarkCounts[1], watermarkCounts[0], "the watermark repeats on page two");
  assert.deepEqual(watermarkCounts.slice(2), [0, 0], "later sections do not inherit it");
  assert.deepEqual(mounted.warnings, []);
});

// ── Notes ──────────────────────────────────────────────────────────
test("footnote and endnote markers match the notes-section numbers", async (t) => {
  const mounted = await mountViewer(t, "review-threads.docx");
  const pages = pageWrappers(mounted.root);
  const markers = pages.flatMap((page) => noteMarkerLabels(page));

  // The fixture writes the footnote marker before the endnote marker in the
  // first annotated paragraph, and the reverse in the closing paragraph, so the
  // expected sequence is derived from the model's own numbering.
  const { model } = await loadFixture("review-threads.docx");
  const snapshot = core.buildLayoutSnapshot(model, {});
  const pageSegments = snapshot.pages.map((page) =>
    page.blocks
      .map((block) => ({ nodeIndex: block.source?.nodeIndex }))
      .filter((segment) => Number.isFinite(segment.nodeIndex)),
  );
  const [pageNotes] = core.buildDocumentPageNotes(model, pageSegments);
  const expectedMarkers = [
    pageNotes.footnoteLabels.get(1),
    pageNotes.endnoteLabels.get(2),
    pageNotes.footnoteLabels.get(2),
    pageNotes.endnoteLabels.get(1),
  ];
  assert.deepEqual(
    markers,
    expectedMarkers,
    "body markers are numbered by the same map the notes section uses",
  );

  const footnoteBodyLabels = pages.flatMap((page) => noteBodyLabels(page, "footnote"));
  const endnoteBodyLabels = pages.flatMap((page) => noteBodyLabels(page, "endnote"));
  assert.deepEqual(
    footnoteBodyLabels,
    [pageNotes.footnoteLabels.get(1), pageNotes.footnoteLabels.get(2)],
    "the footnote section prints the same numbers as the body markers",
  );
  assert.deepEqual(
    endnoteBodyLabels,
    [pageNotes.endnoteLabels.get(2), pageNotes.endnoteLabels.get(1)],
    "the endnote section prints the same numbers as the body markers",
  );
  assert.deepEqual(footnoteBodyLabels, ["1", "2"]);
  assert.deepEqual(endnoteBodyLabels, ["i", "ii"]);

  const text = htmlText(mounted.root);
  assert.ok(text.includes("Footnote one"), "the first footnote body renders");
  assert.ok(text.includes("Endnote one"), "the first endnote body renders");
});
