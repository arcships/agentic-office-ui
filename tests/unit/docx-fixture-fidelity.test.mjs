// Issue #53 — the DOCX fixtures must exercise the review/table/section
// features the user-reported defects live in, and the suite must assert the
// parsed result instead of only "a card exists".
//
// This file parses the generated fixtures in
// `apps/demo/public/samples/` with the real DOCX importer and asserts the
// document model: comment threads and replies, several comments on one
// paragraph, a range that crosses paragraphs, comments in a table cell and in a
// header, a resolved thread, vertical merges and nested tables, repeating header
// rows, portrait/landscape/two-column sections, a text box, a behind-text
// watermark, note numbering, and an import → edit → export → re-import field
// diff.
//
// The rendered page assertions (gutter placement, one card per comment per
// page, deleted-row styling, header repetition) live in
// `tests/component/docx-fixture-fidelity.test.mjs`.

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromVueDocx = createRequire(
  new URL("../../packages/vue-docx/package.json", import.meta.url),
);
const core = await import(
  pathToFileURL(requireFromVueDocx.resolve("@arcships/docx-core")).href
);

const SAMPLES = new URL("../../apps/demo/public/samples/", import.meta.url);

function fixtureBuffer(name) {
  const bytes = readFileSync(new URL(name, SAMPLES));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

const fixtureCache = new Map();

async function loadFixture(name) {
  const cached = fixtureCache.get(name);
  if (cached) return cached;
  const pending = core.importDocxBuffer(fixtureBuffer(name), { useWorker: false });
  fixtureCache.set(name, pending);
  return pending;
}

async function modelOf(name) {
  const { model } = await loadFixture(name);
  return model;
}

function paragraphNodes(model) {
  return model.nodes
    .map((node, index) => ({ node, index }))
    .filter((entry) => entry.node.type === "paragraph");
}

function tableNodes(model) {
  return model.nodes
    .map((node, index) => ({ node, index }))
    .filter((entry) => entry.node.type === "table");
}

function cellParagraphs(cell) {
  return cell.nodes.filter((node) => node.type === "paragraph");
}

function plainText(model) {
  const texts = [];
  const walk = (nodes) => {
    nodes.forEach((node) => {
      if (node.type === "paragraph") {
        texts.push(core.paragraphText(node));
        return;
      }
      node.rows.forEach((row) => row.cells.forEach((cell) => walk(cell.nodes)));
    });
  };
  walk(model.nodes);
  return texts.join("\n");
}

/** Field-level snapshot used for the import → export → re-import diff. */
function documentFields(model) {
  const paragraphs = [];
  const tables = [];
  const walk = (nodes, path) => {
    nodes.forEach((node) => {
      if (node.type === "paragraph") {
        paragraphs.push([path, core.paragraphText(node)]);
        return;
      }
      tables.push([
        path,
        node.rows.map((row) =>
          row.cells.map((cell) =>
            cellParagraphs(cell).map((paragraph) => core.paragraphText(paragraph)),
          ),
        ),
      ]);
      node.rows.forEach((row, rowIndex) =>
        row.cells.forEach((cell, cellIndex) =>
          walk(cell.nodes, `${path}.${rowIndex}.${cellIndex}`),
        ),
      );
    });
  };
  walk(model.nodes, "body");
  return {
    paragraphs,
    tables,
    comments: (model.metadata.comments ?? []).map((comment) => [
      comment.id,
      comment.author ?? null,
      comment.initials ?? null,
      comment.text,
      comment.parentId ?? null,
      comment.resolved ?? false,
    ]),
    footnotes: (model.metadata.footnotes ?? []).map((note) => [note.id, note.text]),
    endnotes: (model.metadata.endnotes ?? []).map((note) => [note.id, note.text]),
    headers: (model.metadata.sections ?? []).flatMap((section) =>
      section.headerSections.map((header) => [
        header.partName,
        header.nodes.map((node) =>
          node.type === "paragraph" ? core.paragraphText(node) : node.type,
        ),
      ]),
    ),
    footers: (model.metadata.sections ?? []).flatMap((section) =>
      section.footerSections.map((footer) => [
        footer.partName,
        footer.nodes.map((node) =>
          node.type === "paragraph" ? core.paragraphText(node) : node.type,
        ),
      ]),
    ),
  };
}

// ── Comment threads, replies and resolved state ────────────────────
test("the review fixture carries threads, replies and a resolved thread", async () => {
  const model = await modelOf("review-threads.docx");
  const comments = model.metadata.comments ?? [];

  assert.equal(comments.length, 11, "every comment in the fixture is imported");
  assert.deepEqual(
    comments.map((comment) => comment.id).sort((left, right) => left - right),
    [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
  );

  const byId = new Map(comments.map((comment) => [comment.id, comment]));
  assert.equal(byId.get(1).parentId, 0, "reply one belongs to the thread root");
  assert.equal(byId.get(2).parentId, 0, "reply two belongs to the thread root");
  assert.equal(byId.get(9).parentId, 8, "the resolved reply belongs to its root");
  assert.equal(byId.get(8).resolved, true, "the thread root is marked done");
  assert.equal(
    comments.filter((comment) => comment.resolved).length,
    1,
    "only one thread is resolved",
  );
  assert.equal(
    comments.filter((comment) => comment.parentId !== undefined).length,
    3,
    "three comments are replies",
  );
  assert.ok(
    comments.every((comment) => comment.text.trim().length > 0),
    "no comment body is lost",
  );
});

test("comment anchors cover the body, a table cell and the page header", async () => {
  const model = await modelOf("review-threads.docx");
  const comments = core.collectCommentsFromModel(model);
  const byId = new Map(comments.map((comment) => [comment.commentId, comment]));

  assert.equal(comments.length, 11, "every comment resolves to an anchor");

  const cellComment = byId.get(6);
  assert.equal(cellComment.region, "body");
  assert.equal(cellComment.location.kind, "table-cell");
  assert.equal(cellComment.location.rowIndex, 1);
  assert.equal(cellComment.location.cellIndex, 1);
  assert.equal(cellComment.anchorText, "1,200.00");

  const headerComment = byId.get(7);
  assert.equal(headerComment.region, "header");
  assert.equal(headerComment.partName, "word/header1.xml");
  assert.equal(headerComment.anchorText, "Review Threads Fixture");

  // Reply two has no inline markup of its own and must inherit the parent's
  // anchor instead of falling back to the first paragraph.
  const replyTwo = byId.get(2);
  const root = byId.get(0);
  assert.deepEqual(replyTwo.location, root.location);
  assert.equal(replyTwo.nodeIndex, root.nodeIndex);
  assert.notEqual(replyTwo.nodeIndex, 0, "the thread is not anchored to the heading");
});

test("two comments share one paragraph and one range crosses paragraphs", async () => {
  const model = await modelOf("review-threads.docx");
  const paragraphs = paragraphNodes(model);
  const rangesByIndex = new Map(
    paragraphs.map((entry) => [entry.index, entry.node.commentRanges ?? []]),
  );

  const twoComments = paragraphs.find(
    (entry) => core.paragraphText(entry.node) === "LEFT_SEGMENT and RIGHT_SEGMENT",
  );
  assert.ok(twoComments, "the two-comment paragraph is imported verbatim");
  const shared = rangesByIndex.get(twoComments.index);
  assert.equal(shared.length, 2, "both comments keep a range on the same paragraph");
  assert.deepEqual(
    shared.map((range) => range.commentId),
    [3, 4],
  );
  const [left, right] = shared;
  assert.equal(left.start, 0);
  assert.equal(left.end, "LEFT_SEGMENT".length);
  assert.ok(right.start >= left.end, "the second range starts after the first one");
  assert.equal(right.end - right.start, "RIGHT_SEGMENT".length);

  const startParagraph = paragraphs.find((entry) =>
    core.paragraphText(entry.node).startsWith("Cross paragraph range begins"),
  );
  const endParagraph = paragraphs.find(
    (entry) => core.paragraphText(entry.node) === "and closes at this point",
  );
  assert.ok(startParagraph && endParagraph, "both ends of the range are imported");

  const startRange = rangesByIndex
    .get(startParagraph.index)
    .find((range) => range.commentId === 5);
  const endRange = rangesByIndex
    .get(endParagraph.index)
    .find((range) => range.commentId === 5);
  assert.ok(startRange, "the range starts in the first paragraph");
  assert.ok(endRange, "the range ends three paragraphs later");
  assert.ok(
    endParagraph.index > startParagraph.index + 1,
    "at least one paragraph sits between the two ends",
  );
  assert.equal(startRange.anchorText, "here");
  assert.equal(endRange.anchorText, "and closes at this point");

  // The gutter anchors a cross-paragraph comment at the paragraph that starts
  // the range, never at the end paragraph.
  const anchor = core
    .collectCommentsFromModel(model)
    .find((comment) => comment.commentId === 5);
  assert.equal(anchor.nodeIndex, startParagraph.index);
});

// ── Table revisions ────────────────────────────────────────────────
test("table revisions anchor the deleted row and the row-height change to their own rows", async () => {
  const model = await modelOf("review-threads.docx");
  const changes = core.collectTrackedChangesFromModel(model);

  const rowChanges = changes.filter(
    (change) => change.kind === "deletion" || change.kind === "row-format-change",
  );
  assert.deepEqual(
    rowChanges.map((change) => change.kind).sort(),
    ["deletion", "row-format-change"],
    "the fixture carries exactly one deleted row and one row-height revision",
  );

  const deletion = rowChanges.find((change) => change.kind === "deletion");
  assert.equal(deletion.location.kind, "table-cell");
  assert.equal(deletion.location.rowIndex, 2, "the deleted row is the third row");
  assert.equal(deletion.text, "DELETED_ROW_TARGET");

  const rowHeight = rowChanges.find((change) => change.kind === "row-format-change");
  assert.equal(rowHeight.location.rowIndex, 3, "the height change belongs to the row below");
  assert.notEqual(
    rowHeight.location.rowIndex,
    deletion.location.rowIndex,
    "a row-height revision must not be reported against the deleted row",
  );

  const table = tableNodes(model).find(
    (entry) => entry.node.rows.some((row) => row.style?.heightTwips === 720),
  );
  assert.ok(table, "the row-height revision belongs to the review table");
  assert.equal(table.node.rows[3].style.heightTwips, 720);
  assert.equal(table.node.rows[3].style.heightRule, "atLeast");
});

// ── Merged, nested and repeating tables ────────────────────────────
test("the table fixture keeps vertical merges, grid spans, nested tables and header rows", async () => {
  const model = await modelOf("table-structures.docx");
  const tables = tableNodes(model);
  assert.equal(tables.length, 2, "the fixture has two top-level tables");

  const [summary, longTable] = tables.map((entry) => entry.node);
  assert.deepEqual(
    summary.style?.columnWidthsTwips?.map((width) => Math.round(width)),
    [2484, 2484, 2484, 2484],
    "the merged table keeps its four-column grid",
  );

  const restart = summary.rows[1].cells[0];
  assert.equal(restart.style?.rowSpan, 3, "the vertical merge spans three rows");
  assert.equal(restart.style?.vMergeContinuation, undefined);
  for (const rowIndex of [2, 3]) {
    assert.equal(
      summary.rows[rowIndex].cells[0].style?.vMergeContinuation,
      true,
      `row ${rowIndex} carries the vMerge continuation cell`,
    );
    assert.equal(
      cellParagraphs(summary.rows[rowIndex].cells[0])
        .map((paragraph) => core.paragraphText(paragraph))
        .join(""),
      "",
      "a continuation cell must not repeat the merged text",
    );
  }

  const spanned = summary.rows[1].cells[1];
  assert.equal(spanned.style?.gridSpan, 2, "the horizontal merge keeps its grid span");
  assert.equal(
    summary.rows[1].cells.length,
    3,
    "a gridSpan cell replaces two grid cells",
  );
  assert.equal(
    summary.rows[1].cells.reduce((total, cell) => total + (cell.style?.gridSpan ?? 1), 0),
    4,
    "every row still covers four grid columns",
  );

  const multiParagraph = summary.rows[2].cells[1];
  assert.equal(cellParagraphs(multiParagraph).length, 3, "the cell keeps all three paragraphs");
  assert.deepEqual(
    cellParagraphs(multiParagraph).map((paragraph) => core.paragraphText(paragraph)),
    ["1,200", "(audited)", "(signed off)"],
  );

  const nestedCell = summary.rows[3].cells[3];
  const nestedTables = nestedCell.nodes.filter((node) => node.type === "table");
  assert.equal(nestedTables.length, 1, "the nested table is imported as a table cell child");
  assert.deepEqual(
    nestedTables[0].rows.map((row) =>
      row.cells.map((cell) => core.paragraphText(cellParagraphs(cell)[0])),
    ),
    [
      ["N1", "N2"],
      ["N3", "N4"],
    ],
    "the nested table keeps its own grid and cell text",
  );

  assert.equal(longTable.rows.length, 46, "the long table is long enough to paginate");
  assert.equal(longTable.rows[0].style?.isHeader, true, "the header row is flagged");
  assert.ok(
    longTable.rows.slice(1).every((row) => row.style?.isHeader !== true),
    "only the first row repeats as a header",
  );
  assert.equal(
    longTable.rows[0].cells.length,
    8,
    "the long table keeps all eight columns",
  );
  assert.equal(
    longTable.style?.columnWidthsTwips?.length,
    8,
    "the long table declares eight grid columns",
  );
});

// ── Sections, columns, text box and watermark ──────────────────────
test("the section fixture keeps portrait, landscape and two-column sections", async () => {
  const model = await modelOf("sections-columns.docx");
  const sections = model.metadata.sections ?? [];
  assert.equal(sections.length, 3, "the fixture declares three sections");
  assert.deepEqual(
    sections.map((section) => section.headerSections.map((header) => header.partName)),
    [["word/header1.xml"], ["word/header2.xml"], ["word/header3.xml"]],
    "every section keeps its own header part",
  );
  assert.deepEqual(
    sections.map((section) => section.footerSections.map((footer) => footer.partName)),
    [["word/footer1.xml"], ["word/footer2.xml"], ["word/footer3.xml"]],
  );
  assert.match(sections[1].sectionPropertiesXml ?? "", /w:orient="landscape"/);
  assert.match(sections[2].sectionPropertiesXml ?? "", /<w:cols w:num="2"/);
  assert.doesNotMatch(sections[0].sectionPropertiesXml ?? "", /w:num="2"/);

  const snapshot = core.buildLayoutSnapshot(model, {});
  assert.equal(snapshot.pages.length, 4, "the fixture paginates into four pages");
  assert.deepEqual(
    snapshot.pages.map((page) => [page.pageSizePx.width, page.pageSizePx.height]),
    [
      [816, 1056],
      [816, 1056],
      [1056, 816],
      [816, 1056],
    ],
    "portrait, portrait, landscape, portrait",
  );
  assert.deepEqual(
    snapshot.pages.map((page) => page.columnCount),
    [1, 1, 1, 2],
    "only the last section is two-column",
  );
  assert.equal(snapshot.pages[3].columnWidthsPx.length, 2);
  assert.ok(
    snapshot.pages[3].columnWidthsPx[0] < snapshot.pages[2].columnWidthsPx[0],
    "two columns are narrower than the single landscape column",
  );
  // The section with more than one page is the one whose header must repeat.
  assert.equal(
    snapshot.pages.filter((page) => page.blocks.some((block) => block.source?.nodeIndex < 26))
      .length,
    2,
    "the first section spans two pages",
  );
});

test("the section fixture keeps a text box and a behind-text watermark", async () => {
  const model = await modelOf("sections-columns.docx");
  const paragraphs = paragraphNodes(model);
  const textBoxParagraph = paragraphs.find((entry) =>
    entry.node.children.some(
      (child) => child.type === "image" && child.syntheticTextBox === true,
    ),
  );
  assert.ok(textBoxParagraph, "the text box paragraph is imported");
  const textBoxImage = textBoxParagraph.node.children.find(
    (child) => child.type === "image" && child.syntheticTextBox === true,
  );
  assert.equal(
    core.syntheticTextBoxTextValue(textBoxImage),
    "TEXT BOX: quarterly review pack\nSecond text box paragraph",
  );
  assert.ok(textBoxImage.floating, "a text box keeps its floating anchor");

  const firstHeader = model.metadata.sections[0].headerSections[0];
  const watermark = firstHeader.nodes
    .flatMap((node) => node.children ?? [])
    .find((child) => child.type === "image");
  assert.ok(watermark, "the first section's header carries the watermark image");
  assert.equal(watermark.floating?.behindDocument, true, "the watermark paints behind the text");
  assert.equal(watermark.floating?.wrapType, "none");
  assert.equal(watermark.floating?.horizontalRelativeTo, "page");
  assert.equal(watermark.floating?.verticalRelativeTo, "page");
  assert.match(firstHeader.partName, /^word\/header\d+\.xml$/);
  assert.ok(
    Number(watermark.floating?.zIndex) < 0 || watermark.floating?.behindDocument === true,
    "the watermark never paints above the body text",
  );
  // Sections two and three must not inherit the watermark.
  for (const section of model.metadata.sections.slice(1)) {
    const images = section.headerSections.flatMap((header) =>
      header.nodes.flatMap((node) => node.children ?? []),
    );
    assert.equal(
      images.filter((child) => child.type === "image").length,
      0,
      `section ${section.startNodeIndex} header has no watermark`,
    );
  }
});

// ── Note numbering ─────────────────────────────────────────────────
test("footnote and endnote markers match the notes-section numbering", async () => {
  const model = await modelOf("review-threads.docx");
  const snapshot = core.buildLayoutSnapshot(model, {});
  const pageSegments = snapshot.pages.map((page) =>
    page.blocks
      .map((block) => ({ nodeIndex: block.source?.nodeIndex }))
      .filter((segment) => Number.isFinite(segment.nodeIndex)),
  );
  const [page] = core.buildDocumentPageNotes(model, pageSegments);

  assert.equal(page.footnotes.length, 2, "both footnotes belong to the first page");
  assert.deepEqual(
    page.footnotes.map((entry) => entry.label),
    ["1", "2"],
    "footnote markers are numbered in document order",
  );
  assert.equal(page.footnoteLabels.get(1), "1", "marker one matches note body one");
  assert.equal(page.footnoteLabels.get(2), "2", "marker two matches note body two");

  // Endnotes use lower roman numbering by default; the marker label and the
  // notes-section label come from one map, so they cannot drift apart.
  assert.deepEqual(
    page.endnotes.map((entry) => entry.label),
    ["i", "ii"],
    "endnote markers are numbered in document order",
  );
  for (const entry of page.endnotes) {
    assert.equal(
      page.endnoteLabels.get(entry.note.id),
      entry.label,
      `endnote ${entry.note.id} uses the same label in the body and the notes list`,
    );
  }

  const numbering = core.buildDocumentNoteNumbering(model);
  assert.deepEqual(
    [...numbering.footnote],
    [
      [1, 1],
      [2, 2],
    ],
  );
  assert.deepEqual(
    [...numbering.endnote],
    [
      [2, 1],
      [1, 2],
    ],
    "endnote numbers follow the reference order, not w:id",
  );
});

// ── Import → edit → export → re-import ─────────────────────────────
test("editing one paragraph survives export and re-import without changing other fields", async () => {
  const { model, package: pkg } = await loadFixture("review-threads.docx");
  const target = paragraphNodes(model).find((entry) =>
    core.paragraphText(entry.node).startsWith("the range continues"),
  );
  assert.ok(target, "the fixture has a plain paragraph to edit");

  const edited = core.updateParagraphText(model, target.index, "edited paragraph text");
  assert.equal(core.paragraphText(edited.nodes[target.index]), "edited paragraph text");
  assert.equal(
    core.paragraphText(model.nodes[target.index]),
    "the range continues through this middle paragraph",
    "the edit does not mutate the imported model",
  );

  const exported = await core.exportDocxBuffer(edited, pkg);
  assert.ok(exported.byteLength > 1000, "the export produced a real DOCX package");

  const { model: reimported } = await core.importDocxBuffer(exported, {
    useWorker: false,
  });
  const before = documentFields(model);
  const after = documentFields(reimported);
  assert.equal(
    before.paragraphs.length,
    after.paragraphs.length,
    "the export keeps the paragraph count",
  );

  const changed = before.paragraphs
    .map((entry, index) => ({
      path: entry[0],
      before: entry[1],
      after: after.paragraphs[index]?.[1],
    }))
    .filter((entry) => entry.before !== entry.after);
  assert.deepEqual(
    changed,
    [
      {
        path: "body",
        before: "the range continues through this middle paragraph",
        after: "edited paragraph text",
      },
    ],
    "only the edited paragraph changes",
  );

  assert.deepEqual(after.comments, before.comments, "comments survive the round trip");
  assert.deepEqual(after.footnotes, before.footnotes, "footnotes survive the round trip");
  assert.deepEqual(after.endnotes, before.endnotes, "endnotes survive the round trip");
  assert.deepEqual(after.tables, before.tables, "table cell text survives the round trip");
  assert.deepEqual(after.headers, before.headers, "headers survive the round trip");
  assert.deepEqual(after.footers, before.footers, "footers survive the round trip");

  // The edited paragraph is still the same node in the same section, and the
  // document still renders.
  assert.equal(core.buildLayoutSnapshot(reimported, {}).pages.length, 1);
  assert.deepEqual(reimported.metadata.warnings, []);
});

test("the review fixture renders as a single page with every annotation", async () => {
  const model = await modelOf("review-threads.docx");
  assert.deepEqual(model.metadata.warnings, []);
  const text = plainText(model);
  for (const marker of [
    "Threaded comment anchor",
    "LEFT_SEGMENT and RIGHT_SEGMENT",
    "Resolved thread anchor paragraph",
    "simple comment anchor paragraph",
    "1,200.00",
  ]) {
    assert.ok(text.includes(marker), `the body still contains ${marker}`);
  }
  assert.equal(core.buildLayoutSnapshot(model, {}).pages.length, 1);
});

// ── Documentation ↔ test contract ──────────────────────────────────
// `docs/visual-acceptance-handoff.md` lists the DOCX acceptance rows. Each row
// carries an id and must name either an existing, CI-registered test or a
// manual/browser check. A claim that loses its test fails this suite instead of
// quietly reading as accepted.
const REPO_ROOT = new URL("../../", import.meta.url);

function readRepoFile(relativePath) {
  return readFileSync(new URL(relativePath, REPO_ROOT), "utf8");
}

test("every acceptance row in the hand-off doc maps to a real, registered test", () => {
  const doc = readRepoFile("docs/visual-acceptance-handoff.md");
  const runner = readRepoFile("scripts/ci/run-suite.mjs");

  const gatesStart = doc.indexOf("## Feature acceptance gates");
  const gatesEnd = doc.indexOf("## Acceptance coverage map");
  assert.ok(gatesStart >= 0 && gatesEnd > gatesStart, "the gates section exists");
  const gates = doc.slice(gatesStart, gatesEnd);
  const gateIds = [...gates.matchAll(/`(G-[A-Z-]+-\d+)`/g)].map((match) => match[1]);
  assert.ok(gateIds.length >= 30, `expected every gate row to carry an id, found ${gateIds.length}`);
  assert.equal(new Set(gateIds).size, gateIds.length, "gate ids are unique");

  const mapStart = doc.indexOf("<!-- acceptance-test-map:start -->");
  const mapEnd = doc.indexOf("<!-- acceptance-test-map:end -->");
  assert.ok(mapStart >= 0 && mapEnd > mapStart, "the coverage map is delimited");
  const map = doc.slice(mapStart, mapEnd);
  const rows = [...map.matchAll(/^\| `(G-[A-Z-]+-\d+)` \| (.+) \|$/gm)].map((match) => ({
    id: match[1],
    coverage: match[2].trim(),
  }));
  assert.deepEqual(
    rows.map((row) => row.id).sort(),
    [...gateIds].sort(),
    "every acceptance row appears exactly once in the coverage map",
  );

  for (const row of rows) {
    if (row.coverage.startsWith("manual")) {
      const target = row.coverage.match(/`([^`]+)`/)?.[1];
      assert.ok(target, `${row.id}: a manual row names the file that still accepts it`);
      assert.doesNotThrow(
        () => readRepoFile(target),
        `${row.id}: ${target} exists`,
      );
      continue;
    }
    const automatedEntries = row.coverage.split("; ").map((entry) => entry.trim());
    for (const entry of automatedEntries) {
      const match = entry.match(/^`([^`]+?)::(.+)`$/);
      assert.ok(match, `${row.id}: automated rows use \`file::test name\``);
      const [, file, testName] = match;
      const source = readRepoFile(file);
      assert.ok(
        source.includes(`test("${testName}"`),
        `${row.id}: ${file} contains the test \`${testName}\``,
      );
      assert.ok(
        runner.includes(`"${file}"`),
        `${row.id}: ${file} is registered in scripts/ci/run-suite.mjs`,
      );
    }
  }
});
