import assert from "node:assert/strict";
import { test } from "node:test";

const vueDocx = await import(
  new URL("../../packages/vue-docx/dist/index.js", import.meta.url).href
);
const runtime = await import(
  new URL("../../packages/office-runtime/dist/index.js", import.meta.url).href
);
const docxCore = await import(
  new URL("../../packages/docx-core/dist/index.js", import.meta.url).href
);

const { findDocxSearchMatches, findDocxSearchTargetPage } = vueDocx;
const { createSurfaceSearchSession } = runtime;
const { createBlankDocumentModel } = docxCore;

function paragraph(text) {
  return { type: "paragraph", children: [{ type: "text", text }] };
}

function modelWith(nodes, metadata = {}) {
  const base = createBlankDocumentModel();
  return {
    ...base,
    nodes,
    metadata: { ...base.metadata, ...metadata },
  };
}

test("DOCX search finds text inside recursively nested tables", () => {
  const nestedTable = {
    type: "table",
    rows: [{
      cells: [{
        nodes: [
          paragraph("nested needle"),
          {
            type: "table",
            rows: [{
              cells: [{ nodes: [paragraph("deep needle")] }],
            }],
          },
        ],
      }],
    }],
  };
  const outerTable = {
    type: "table",
    rows: [{
      cells: [{
        nodes: [paragraph("outer cell"), nestedTable],
      }],
    }],
  };
  const model = modelWith([paragraph("body needle"), outerTable]);

  const matches = findDocxSearchMatches(model, "needle");
  assert.equal(matches.length, 3, "body, nested and deeply nested text");

  assert.deepEqual(matches[0].range.start.location, { kind: "paragraph", nodeIndex: 0 });
  assert.deepEqual(matches[1].range.start.location, {
    kind: "nested-table-cell",
    tableIndex: 1,
    path: [{ rowIndex: 0, cellIndex: 0, contentIndex: 1 }],
    rowIndex: 0,
    cellIndex: 0,
    paragraphIndex: 0,
    nestingPath: "0-0-1",
  });
  assert.deepEqual(matches[1].range, {
    start: {
      location: {
        kind: "nested-table-cell",
        tableIndex: 1,
        path: [{ rowIndex: 0, cellIndex: 0, contentIndex: 1 }],
        rowIndex: 0,
        cellIndex: 0,
        paragraphIndex: 0,
        nestingPath: "0-0-1",
      },
      offset: 7,
    },
    end: {
      location: {
        kind: "nested-table-cell",
        tableIndex: 1,
        path: [{ rowIndex: 0, cellIndex: 0, contentIndex: 1 }],
        rowIndex: 0,
        cellIndex: 0,
        paragraphIndex: 0,
        nestingPath: "0-0-1",
      },
      offset: 13,
    },
  });

  assert.deepEqual(matches[2].range.start.location.path, [
    { rowIndex: 0, cellIndex: 0, contentIndex: 1 },
    { rowIndex: 0, cellIndex: 0, contentIndex: 1 },
  ]);
  assert.equal(
    matches[2].range.start.location.nestingPath,
    "0-0-1/0-0-1",
  );
});

test("DOCX search covers headers, footers, footnotes and endnotes with region tags", () => {
  const headerTable = {
    type: "table",
    rows: [{
      cells: [{
        nodes: [
          {
            type: "table",
            rows: [{
              cells: [{ nodes: [paragraph("header nested needle")] }],
            }],
          },
        ],
      }],
    }],
  };
  const model = modelWith([paragraph("body text")], {
    headerSections: [{
      partName: "word/header1.xml",
      referenceType: "default",
      nodes: [paragraph("header needle"), headerTable],
    }],
    footerSections: [{
      partName: "word/footer1.xml",
      referenceType: "default",
      nodes: [paragraph("footer needle")],
    }],
    footnotes: [{ id: 2, text: "", nodes: [paragraph("footnote needle")] }],
    endnotes: [{ id: 7, text: "endnote needle" }],
  });

  const matches = findDocxSearchMatches(model, "needle");
  assert.deepEqual(
    matches.map((match) => [match.region, match.partName, match.noteId]),
    [
      ["header", "word/header1.xml", undefined],
      ["header", "word/header1.xml", undefined],
      ["footer", "word/footer1.xml", undefined],
      ["footnote", undefined, 2],
      ["endnote", undefined, 7],
    ],
  );

  assert.deepEqual(matches[0].range.start.location, { kind: "paragraph", nodeIndex: 0 });
  assert.equal(matches[0].nodeIndex, 0);

  const headerNested = matches[1];
  assert.equal(headerNested.nodeIndex, 1);
  assert.equal(headerNested.range.start.location.kind, "nested-table-cell");
  assert.equal(headerNested.range.start.location.tableIndex, 1);
  assert.equal(
    headerNested.range.start.location.nestingPath,
    "0-0-0",
  );

  assert.deepEqual(matches[3].range.start.location, { kind: "paragraph", nodeIndex: 0 });
  assert.deepEqual(matches[4].range.start.location, { kind: "note-text", noteId: 7 });
  assert.equal(matches[4].match, "needle");
});

test("header/footer/note parts are searched once even when sections share them", () => {
  const sharedHeader = {
    partName: "word/header1.xml",
    referenceType: "default",
    nodes: [paragraph("shared header needle")],
  };
  const model = modelWith([paragraph("body")], {
    headerSections: [sharedHeader, { ...sharedHeader }],
    footerSections: [
      { partName: "word/footer1.xml", nodes: [paragraph("first footer needle")] },
      { partName: "word/footer1.xml", nodes: [paragraph("duplicate footer needle")] },
    ],
  });

  const matches = findDocxSearchMatches(model, "needle");
  assert.deepEqual(
    matches.map((match) => [match.region, match.partName]),
    [
      ["header", "word/header1.xml"],
      ["footer", "word/footer1.xml"],
    ],
  );
});

test("DOCX search prefers the header/footer parts the viewer actually renders", () => {
  const model = modelWith([paragraph("body")], {
    headerSections: [{ partName: "word/header1.xml", nodes: [paragraph("stale header text")] }],
    footerSections: [{ partName: "word/footer1.xml", nodes: [paragraph("stale footer text")] }],
    sections: [{
      startNodeIndex: 0,
      headerSections: [{
        partName: "word/header1.xml",
        nodes: [paragraph("rendered header needle")],
      }],
      footerSections: [{
        partName: "word/footer1.xml",
        nodes: [paragraph("rendered footer needle")],
      }],
    }],
  });

  const matches = findDocxSearchMatches(model, "needle");
  assert.deepEqual(
    matches.map((match) => [match.region, match.match]),
    [
      ["header", "needle"],
      ["footer", "needle"],
    ],
  );
  assert.equal(findDocxSearchMatches(model, "stale").length, 0);
});

test("DOCX search options distinguish case, whole words and regular expressions", () => {
  const model = modelWith([
    paragraph("Word word WORD cat category concatenate."),
    paragraph("version 1.2.3"),
  ]);

  assert.deepEqual(
    findDocxSearchMatches(model, "word").map((match) => match.match),
    ["Word", "word", "WORD"],
  );
  assert.deepEqual(
    findDocxSearchMatches(model, "Word", undefined, { caseSensitive: true })
      .map((match) => [match.match, match.range.start.offset]),
    [["Word", 0]],
  );

  assert.deepEqual(
    findDocxSearchMatches(model, "cat", undefined, { wholeWord: true })
      .map((match) => [match.match, match.range.start.offset]),
    [['cat', 15]],
  );
  assert.equal(
    findDocxSearchMatches(model, "cat", undefined, { wholeWord: false }).length,
    3,
  );

  assert.deepEqual(
    findDocxSearchMatches(model, "\\d+\\.\\d+\\.\\d+", undefined, { regex: true })
      .map((match) => match.match),
    ["1.2.3"],
  );
  assert.deepEqual(
    findDocxSearchMatches(model, "WORD", undefined, { regex: true, caseSensitive: true })
      .map((match) => match.match),
    ["WORD"],
  );
  assert.throws(
    () => findDocxSearchMatches(model, "(", undefined, { regex: true }),
    /Invalid regular expression/u,
  );
});

test("search session forwards options and refreshes offsets after a model edit", async () => {
  let model = modelWith([paragraph("needle one needle two")]);
  const requests = [];
  const activations = [];
  const session = createSurfaceSearchSession({
    search(query, { signal }, options) {
      requests.push([query, options]);
      return findDocxSearchMatches(model, query, signal, options);
    },
    activate(match, index) {
      activations.push([match.match, index]);
    },
  });

  const first = await session.search("needle", { caseSensitive: true, wholeWord: true });
  assert.deepEqual(requests, [["needle", { caseSensitive: true, wholeWord: true }]]);
  assert.equal(first.matches.length, 2);
  assert.equal(first.activeIndex, 0);

  await session.activateSearchMatch(1);
  assert.equal(session.getSearchState().activeIndex, 1);

  // An edit replaces the model object (clone-on-write) and shifts match offsets.
  model = modelWith([paragraph("xx needle one needle two")]);
  const refreshed = await session.refreshSearch();
  assert.equal(refreshed.status, "ready");
  assert.equal(refreshed.query, "needle");
  assert.deepEqual(requests.at(-1), ["needle", { caseSensitive: true, wholeWord: true }]);
  assert.equal(refreshed.matches.length, 2, "the result list survives the edit");
  assert.deepEqual(
    refreshed.matches.map((match) => match.range.start.offset),
    [3, 14],
    "offsets are re-resolved against the new model",
  );
  assert.equal(refreshed.activeIndex, 1, "the active result is preserved");

  // A shorter document clamps the preserved index instead of failing.
  model = modelWith([paragraph("yy needle")]);
  const clamped = await session.refreshSearch();
  assert.equal(clamped.matches.length, 1);
  assert.equal(clamped.activeIndex, 0);

  session.clearSearch();
  const cleared = session.getSearchState();
  assert.equal(cleared.status, "idle");
  const afterClear = await session.refreshSearch();
  assert.equal(afterClear, cleared);
  session.dispose();
});

test("search page lookup resolves header/footer/note targets for navigation", () => {
  const presentations = [
    { headerSection: { partName: "word/header1.xml" }, footerSection: { partName: "word/footer1.xml" } },
    { headerSection: { partName: "word/header2.xml" }, footerSection: { partName: "word/footer1.xml" } },
    {},
  ];
  const notes = [
    { footnotes: [], endnotes: [] },
    { footnotes: [{ note: { id: 3 } }], endnotes: [] },
    { footnotes: [], endnotes: [{ note: { id: 9 } }] },
  ];

  assert.equal(
    findDocxSearchTargetPage({ region: "header", partName: "word/header2.xml" }, presentations, notes),
    1,
  );
  assert.equal(
    findDocxSearchTargetPage({ region: "footer", partName: "word/footer1.xml" }, presentations, notes),
    0,
  );
  assert.equal(
    findDocxSearchTargetPage({ region: "footnote", noteId: 3 }, presentations, notes),
    1,
  );
  assert.equal(
    findDocxSearchTargetPage({ region: "endnote", noteId: 9 }, presentations, notes),
    2,
  );
  assert.equal(findDocxSearchTargetPage({ region: "body" }, presentations, notes), undefined);
  assert.equal(
    findDocxSearchTargetPage({ region: "header", partName: "word/missing.xml" }, presentations, notes),
    undefined,
  );
});
