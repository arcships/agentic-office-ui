// Issue #134 — DOCX export pre-serialization.
//
// Every block names the bullet it covers so the mapping stays checkable:
//
//   1. prepareModelForSerialization recurses every story .. prepare-model
//   2. collapsed table grids project onto physical columns  projectCollapsedTableGrid
//   3. style/border-spacing invariants + typed error ...... DocxSerializationError
//   4. merged-table fixture round-trip with edits ......... serialize + rebuild
//   5. illegal models reject instead of writing a package . serializeDocx / exportDocxBuffer
//
// The acceptance case is the merged-table corpus fixture (#72) edited through
// the same table-cell command the editor uses, then exported and reopened.

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromDocxCore = createRequire(
  new URL("../../packages/docx-core/package.json", import.meta.url),
);
const core = await import(
  pathToFileURL(requireFromDocxCore.resolve("@arcships/docx-core")).href
);
const { strFromU8, unzipSync } = await import(
  pathToFileURL(requireFromDocxCore.resolve("fflate")).href
);

const WORD_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const SECTION_PROPS =
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>';

function sampleBytes(relativePath) {
  const bytes = readFileSync(new URL(relativePath, import.meta.url));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function documentXml(body) {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<w:document ${WORD_NS}><w:body>${body}<w:sectPr>${SECTION_PROPS}</w:sectPr></w:body></w:document>`
  );
}

function paragraph(text) {
  return {
    type: "paragraph",
    children: text === null ? [] : [{ type: "text", text }],
  };
}

function modelOf(nodes, metadata = {}) {
  return {
    nodes,
    metadata: {
      sourceParts: 1,
      warnings: [],
      headerSections: [],
      footerSections: [],
      paragraphStyles: [],
      ...metadata,
    },
  };
}

/**
 * Logical grid `[900, 0, 1800]`: column 1 was collapsed by a merge, so both
 * rows span it with `gridSpan: 2`.
 */
function collapsedTable(style = {}) {
  return {
    type: "table",
    style: {
      layout: "fixed",
      columnWidthsTwips: [900, 0, 1800],
      collapsedGridColumns: [1],
      ...style,
    },
    rows: [
      {
        type: "table-row",
        cells: [
          { type: "table-cell", style: { gridSpan: 2 }, nodes: [paragraph("Wide")] },
          { type: "table-cell", nodes: [paragraph("Right")] },
        ],
      },
      {
        type: "table-row",
        cells: [
          { type: "table-cell", style: { gridSpan: 2 }, nodes: [paragraph("Left")] },
          { type: "table-cell", nodes: [paragraph("Right2")] },
        ],
      },
    ],
  };
}

function mergedTable() {
  return {
    type: "table",
    style: { layout: "fixed", columnWidthsTwips: [2500, 2500] },
    rows: [
      {
        type: "table-row",
        cells: [
          { type: "table-cell", style: { rowSpan: 2 }, nodes: [paragraph("A1")] },
          { type: "table-cell", nodes: [paragraph("B1")] },
        ],
      },
      {
        type: "table-row",
        cells: [
          {
            type: "table-cell",
            style: { vMergeContinuation: true },
            nodes: [paragraph(null)],
          },
          { type: "table-cell", nodes: [paragraph("B2")] },
        ],
      },
    ],
  };
}

function documentXmlOf(bytes) {
  return strFromU8(unzipSync(new Uint8Array(bytes))["word/document.xml"]);
}

function tableXmls(xml) {
  return xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/g) ?? [];
}

function rowSignature(table) {
  return table.rows.map((row) =>
    row.cells.map((cell) => ({
      gridSpan: cell.style?.gridSpan ?? 1,
      rowSpan: cell.style?.rowSpan ?? 1,
      continuation: cell.style?.vMergeContinuation === true,
    })),
  );
}

function findTable(nodes) {
  for (const node of nodes) {
    if (node.type === "table") return node;
    if (node.type === "paragraph") continue;
  }
  return undefined;
}

// ── 2. collapsed grid projection ─────────────────────────────────────

test("collapsed grid projection derives physical widths, boundaries and spans", () => {
  const source = collapsedTable();
  const before = structuredClone(source);

  const result = core.projectCollapsedTableGrid(source, {
    resolvedBorderSpacingPx: 0,
  });

  assert.equal(result.kind, "projected");
  assert.deepEqual(result.table.style.columnWidthsTwips, [900, 1800]);
  assert.equal(result.table.style.collapsedGridColumns, undefined);
  assert.equal(result.table.rows[0].cells[0].style.gridSpan, undefined);
  assert.deepEqual(result.logicalToPhysicalBoundary, [0, 1, 1, 2]);
  assert.deepEqual(result.physicalToLogicalBoundaryRanges, [
    { start: 0, end: 0 },
    { start: 1, end: 2 },
    { start: 3, end: 3 },
  ]);
  assert.deepEqual(result.physicalToLogicalColumn, [0, 2]);
  assert.deepEqual(result.collapsedGridColumns, [1]);
  assert.deepEqual(source, before, "projection must not mutate the imported model");
});

test("projection rejects grids the marker cannot reconcile", () => {
  const changes = {
    "duplicate collapsed indices": (table) => {
      table.style.collapsedGridColumns = [1, 1];
    },
    "cell collapsed to an empty span": (table) => {
      table.rows[1].cells[0].style.gridSpan = 1;
    },
    "row grid gaps": (table) => {
      table.rows[0].style = { gridBefore: 1 };
    },
    "non-fixed layout": (table) => {
      table.style.layout = "autofit";
    },
    "nonzero cell spacing": (table) => {
      table.style.cellSpacingTwips = 20;
    },
    "invalid collapsed index": (table) => {
      table.style.collapsedGridColumns = [7];
    },
    "marked column with a positive width": (table) => {
      table.style.columnWidthsTwips = [900, 600, 1800];
    },
  };
  const reasons = new Map();
  for (const [label, change] of Object.entries(changes)) {
    const table = collapsedTable();
    change(table);
    const result = core.projectCollapsedTableGrid(table, {
      resolvedBorderSpacingPx: 0,
    });
    assert.equal(result.kind, "rejected", label);
    reasons.set(label, result.reason);
  }
  // An empty marker is "no collapsed grid", not a rejection.
  assert.equal(
    core.projectCollapsedTableGrid(collapsedTable({ collapsedGridColumns: [] }), {
      resolvedBorderSpacingPx: 0,
    }).kind,
    "identity",
  );
  assert.equal(reasons.get("duplicate collapsed indices"), "duplicate-collapsed-index");
  assert.equal(
    reasons.get("cell collapsed to an empty span"),
    "cell-collapses-to-empty-span",
  );
  assert.equal(reasons.get("row grid gaps"), "row-grid-gaps-not-supported");
  assert.equal(reasons.get("non-fixed layout"), "fixed-layout-required");
  assert.equal(reasons.get("nonzero cell spacing"), "nonzero-source-cell-spacing");
  assert.equal(reasons.get("invalid collapsed index"), "invalid-collapsed-index");
  assert.equal(
    reasons.get("marked column with a positive width"),
    "marked-column-has-positive-width",
  );
});

test("projection requires confirmed zero border spacing", () => {
  assert.equal(core.projectCollapsedTableGrid(collapsedTable()).kind, "rejected");
  assert.equal(
    core.projectCollapsedTableGrid(collapsedTable(), {
      resolvedBorderSpacingPx: 1,
    }).kind,
    "rejected",
  );
  assert.equal(
    core.projectCollapsedTableGrid(collapsedTable(), {
      resolvedBorderSpacingPx: 0,
    }).kind,
    "projected",
  );
  assert.equal(core.hasCollapsedTableGrid(collapsedTable()), true);
  assert.equal(core.hasCollapsedTableGrid(mergedTable()), false);
});

// ── 1. story recursion ───────────────────────────────────────────────

test("preparation projects every story and leaves the source model untouched", () => {
  const nested = collapsedTable();
  const body = collapsedTable();
  body.rows[0].cells[0].nodes.push(nested);
  const header = collapsedTable();
  const footer = collapsedTable();
  const sectionHeader = collapsedTable();
  const footnote = collapsedTable();
  const endnote = collapsedTable();

  const model = modelOf([body], {
    headerSections: [
      { partName: "word/header1.xml", referenceType: "default", nodes: [header] },
    ],
    footerSections: [
      { partName: "word/footer1.xml", referenceType: "default", nodes: [footer] },
    ],
    sections: [
      {
        startNodeIndex: 0,
        headerSections: [
          { partName: "word/header2.xml", referenceType: "first", nodes: [sectionHeader] },
        ],
        footerSections: [],
      },
    ],
    footnotes: [{ id: 1, text: "note", nodes: [footnote] }],
    endnotes: [{ id: 2, text: "endnote", nodes: [endnote] }],
  });

  const prepared = core.prepareModelForSerialization(model);
  assert.notEqual(prepared, model);

  const projected = (table) => {
    assert.deepEqual(table.style.columnWidthsTwips, [900, 1800]);
    assert.equal(table.style.collapsedGridColumns, undefined);
    return true;
  };

  assert.ok(projected(prepared.nodes[0]));
  assert.ok(projected(prepared.nodes[0].rows[0].cells[0].nodes[1]));
  assert.ok(projected(prepared.metadata.headerSections[0].nodes[0]));
  assert.ok(projected(prepared.metadata.footerSections[0].nodes[0]));
  assert.ok(projected(prepared.metadata.sections[0].headerSections[0].nodes[0]));
  assert.ok(projected(prepared.metadata.footnotes[0].nodes[0]));
  assert.ok(projected(prepared.metadata.endnotes[0].nodes[0]));

  // The imported model keeps the logical grid it was parsed with.
  assert.deepEqual(model.nodes[0].style.columnWidthsTwips, [900, 0, 1800]);
  assert.deepEqual(model.metadata.headerSections[0].nodes[0].style.columnWidthsTwips, [
    900, 0, 1800,
  ]);
});

test("preparation is an identity for models that need no projection", () => {
  const plain = modelOf([paragraph("Body")], {
    headerSections: [{ partName: "word/header1.xml", nodes: [paragraph("Header")] }],
  });
  assert.equal(core.prepareModelForSerialization(plain), plain);

  // A table that still carries its stored markup is written verbatim by the
  // engine, so its logical grid — and its borders — are deliberately left
  // alone; the export must not reject a document Word already wrote.
  const stored = collapsedTable({ borders: { top: { type: "single" } } });
  stored.sourceXml = "<w:tbl/>";
  const model = modelOf([stored]);
  assert.equal(core.prepareModelForSerialization(model), model);
});

// ── 3. invariants and typed errors ───────────────────────────────────

test("visible collapsed borders fail before an illegal grid is written", () => {
  const table = collapsedTable({
    borders: { top: { type: "single", sizeEighthPt: 4 } },
  });
  assert.throws(
    () => core.prepareModelForSerialization(modelOf([table])),
    (error) => {
      assert.ok(error instanceof core.DocxSerializationError);
      assert.equal(error.code, "COLLAPSED_TABLE_GRID_BORDER_SPACING");
      assert.match(error.message, /zero border spacing/);
      return true;
    },
  );
});

test("a table style changed after import rejects the collapsed grid", () => {
  const table = collapsedTable({ styleId: "TableGrid2", sourceStyleId: "TableGrid" });
  assert.throws(
    () => core.prepareModelForSerialization(modelOf([table])),
    (error) => {
      assert.equal(error.code, "COLLAPSED_TABLE_GRID_STYLE_CHANGED");
      assert.match(error.message, /"TableGrid" → "TableGrid2"/);
      return true;
    },
  );
});

test("an illegal grid rejects with the projection reason as a typed error", async () => {
  const table = collapsedTable();
  table.rows[1].cells[0].style.gridSpan = 1;
  const model = modelOf([table]);

  await assert.rejects(
    core.modelToDocumentXml(model),
    (error) => {
      assert.ok(error instanceof core.DocxSerializationError);
      assert.equal(error.code, "INVALID_COLLAPSED_TABLE_GRID");
      assert.equal(error.reason, "cell-collapses-to-empty-span");
      assert.match(error.message, /cell-collapses-to-empty-span/);
      return true;
    },
  );
  await assert.rejects(core.serializeDocModel(model), core.DocxSerializationError);
  await assert.rejects(core.serializeDocx(model), core.DocxSerializationError);
});

test("exportDocxBuffer surfaces an illegal model as INVALID_MODEL", async () => {
  const table = collapsedTable();
  table.style.collapsedGridColumns = [1, 1];
  const model = modelOf([table], { sourceParts: 1 });

  await assert.rejects(
    core.exportDocxBuffer(model),
    (error) => {
      assert.ok(error instanceof core.DocxExportError);
      assert.equal(error.code, "INVALID_MODEL");
      assert.match(error.message, /duplicate-collapsed-index/);
      return true;
    },
  );
});

// ── 4. acceptance: merged table round-trip with edits ────────────────

test("editing the merged-table fixture keeps its grid through export and reload", async () => {
  const bytes = sampleBytes("../../test-data/corpus/docx-merged-table.docx");
  const { model, package: pkg } = await core.buildDocModelFromBytes(bytes);

  // Node 1 is the vertical-merge table, node 3 the gridSpan table.
  assert.equal(model.nodes[1].type, "table");
  assert.equal(model.nodes[3].type, "table");

  let edited = core.updateTableCellParagraphText(model, 1, 0, 0, 0, "Merged A1 edited");
  edited = core.updateTableCellParagraphText(edited, 3, 0, 0, 0, "Wide header edited");
  assert.equal(edited.nodes[1].sourceXml, undefined);
  assert.equal(edited.nodes[3].sourceXml, undefined);

  const exported = await core.serializeDocx(edited, pkg);
  const xml = documentXmlOf(exported);
  const [verticalMergeTable, horizontalMergeTable] = tableXmls(xml);
  assert.match(verticalMergeTable, /<w:vMerge w:val="restart"\/>/);
  assert.match(verticalMergeTable, /<w:vMerge\/>/);
  assert.match(horizontalMergeTable, /<w:gridSpan w:val="2"\/>/);
  assert.match(verticalMergeTable, /Merged A1 edited/);

  const reimported = (await core.buildDocModelFromBytes(exported)).model;
  const vertical = reimported.nodes[1];
  const horizontal = reimported.nodes[3];
  assert.deepEqual(vertical.style.columnWidthsTwips, [2500, 2500, 2500]);
  assert.deepEqual(rowSignature(vertical), [
    [
      { gridSpan: 1, rowSpan: 2, continuation: false },
      { gridSpan: 1, rowSpan: 1, continuation: false },
      { gridSpan: 1, rowSpan: 1, continuation: false },
    ],
    [
      { gridSpan: 1, rowSpan: 1, continuation: true },
      { gridSpan: 1, rowSpan: 1, continuation: false },
      { gridSpan: 1, rowSpan: 1, continuation: false },
    ],
    [
      { gridSpan: 1, rowSpan: 1, continuation: false },
      { gridSpan: 1, rowSpan: 1, continuation: false },
      { gridSpan: 1, rowSpan: 1, continuation: false },
    ],
  ]);
  assert.deepEqual(rowSignature(horizontal), [
    [
      { gridSpan: 2, rowSpan: 1, continuation: false },
      { gridSpan: 1, rowSpan: 1, continuation: false },
    ],
    [
      { gridSpan: 1, rowSpan: 1, continuation: false },
      { gridSpan: 1, rowSpan: 1, continuation: false },
      { gridSpan: 1, rowSpan: 1, continuation: false },
    ],
  ]);
  assert.equal(
    core.paragraphText(vertical.rows[0].cells[0].nodes[0]),
    "Merged A1 edited",
  );
  assert.equal(
    core.paragraphText(horizontal.rows[0].cells[0].nodes[0]),
    "Wide header edited",
  );
});

test("the same merges survive the package-returning serializer", async () => {
  const bytes = sampleBytes("../../test-data/corpus/docx-merged-table.docx");
  const { model } = await core.buildDocModelFromBytes(bytes);
  const edited = core.updateTableCellParagraphText(model, 1, 0, 0, 0, "Edited");

  const pkg = await core.serializeDocModel(edited);
  const documentPart = pkg.parts.get("word/document.xml");
  assert.ok(documentPart);
  assert.match(documentPart.content, /<w:vMerge w:val="restart"\/>/);
  assert.match(documentPart.content, /<w:vMerge\/>/);

  const documentXml = await core.modelToDocumentXml(edited);
  assert.match(documentXml, /<w:vMerge w:val="restart"\/>/);
});

test("a merge created by the editor command survives export and reload", async () => {
  const { model } = await core.buildDocModelFromBytes(
    sampleBytes("../../test-data/corpus/docx-merged-table.docx"),
  );
  // Node 5 is the 2x2 table at the end of the fixture.
  assert.equal(model.nodes[5].type, "table");
  const next = core.cloneDocModel(model);
  const merged = core.mergeTableCells(next.nodes[5], {
    anchorRowIndex: 0,
    anchorCellIndex: 0,
    focusRowIndex: 0,
    focusCellIndex: 1,
  });
  assert.ok(merged, "the two cells of row 0 merge");
  assert.equal(next.nodes[5].sourceXml, undefined);

  const exported = await core.serializeDocx(next);
  const reimported = (await core.buildDocModelFromBytes(exported)).model;
  assert.deepEqual(rowSignature(reimported.nodes[5]), [
    [
      { gridSpan: 2, rowSpan: 1, continuation: false },
    ],
    [
      { gridSpan: 1, rowSpan: 1, continuation: false },
      { gridSpan: 1, rowSpan: 1, continuation: false },
    ],
  ]);
});

test("a nested table imported from stored markup keeps its merges after an edit", async () => {
  const innerTableXml =
    '<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
    '<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>' +
    '<w:p><w:r><w:t>Inner A</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:p><w:r><w:t>Inner B</w:t></w:r></w:p></w:tc></w:tr>' +
    '<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>' +
    '<w:tc><w:p><w:r><w:t>Inner B2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  const outerTableXml =
    '<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    `<w:tr><w:tc>${innerTableXml}<w:p><w:r><w:t>Outer A</w:t></w:r></w:p></w:tc>` +
    '<w:tc><w:p><w:r><w:t>Outer B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  const pkg = core.createMinimalDocxPackage(documentXml(outerTableXml));
  const model = await core.buildDocModel(pkg);

  // The recursive cell editor clears the inner table and every ancestor table's
  // stored markup, so the export has to rebuild both levels from the model.
  const edited = core.updateTableCellParagraphTextRecursive(
    model,
    0,
    0,
    0,
    0,
    "Inner A edited",
  );
  assert.equal(edited.nodes[0].sourceXml, undefined);
  const editedInner = edited.nodes[0].rows[0].cells[0].nodes.find(
    (node) => node.type === "table",
  );
  assert.ok(editedInner);
  assert.equal(editedInner.sourceXml, undefined);

  const exported = await core.serializeDocx(edited, pkg);
  const xml = documentXmlOf(exported);
  assert.equal((xml.match(/<w:vMerge w:val="restart"\/>/g) ?? []).length, 1);
  assert.equal((xml.match(/<w:vMerge\/>/g) ?? []).length, 1);
  assert.match(xml, /Inner A edited/);

  const reimported = (await core.buildDocModelFromBytes(exported)).model;
  const reloadedInner = reimported.nodes[0].rows[0].cells[0].nodes.find(
    (node) => node.type === "table",
  );
  assert.ok(reloadedInner);
  assert.deepEqual(rowSignature(reloadedInner), [
    [
      { gridSpan: 1, rowSpan: 2, continuation: false },
      { gridSpan: 1, rowSpan: 1, continuation: false },
    ],
    [
      { gridSpan: 1, rowSpan: 1, continuation: true },
      { gridSpan: 1, rowSpan: 1, continuation: false },
    ],
  ]);
  assert.equal(
    core.paragraphText(reloadedInner.rows[0].cells[0].nodes[0]),
    "Inner A edited",
  );
});

test("a nested merged table keeps its own merges through export", async () => {
  const inner = mergedTable();
  const outer = {
    type: "table",
    style: { layout: "fixed", columnWidthsTwips: [3000, 3000] },
    rows: [
      {
        type: "table-row",
        cells: [
          { type: "table-cell", style: { rowSpan: 2 }, nodes: [inner] },
          { type: "table-cell", nodes: [paragraph("Outer B")] },
        ],
      },
      {
        type: "table-row",
        cells: [
          {
            type: "table-cell",
            style: { vMergeContinuation: true },
            nodes: [paragraph(null)],
          },
          { type: "table-cell", nodes: [paragraph("Outer B2")] },
        ],
      },
    ],
  };
  const model = modelOf([outer]);

  const exported = await core.serializeDocx(model);
  const xml = documentXmlOf(exported);
  assert.equal((xml.match(/<w:vMerge w:val="restart"\/>/g) ?? []).length, 2);
  assert.equal((xml.match(/<w:vMerge\/>/g) ?? []).length, 2);

  const reimported = (await core.buildDocModelFromBytes(exported)).model;
  const reloadedOuter = reimported.nodes[0];
  const reloadedInner = reloadedOuter.rows[0].cells[0].nodes.find(
    (node) => node.type === "table",
  );
  assert.ok(reloadedInner);
  assert.deepEqual(rowSignature(reloadedOuter), [
    [
      { gridSpan: 1, rowSpan: 2, continuation: false },
      { gridSpan: 1, rowSpan: 1, continuation: false },
    ],
    [
      { gridSpan: 1, rowSpan: 1, continuation: true },
      { gridSpan: 1, rowSpan: 1, continuation: false },
    ],
  ]);
  assert.deepEqual(rowSignature(reloadedInner), [
    [
      { gridSpan: 1, rowSpan: 2, continuation: false },
      { gridSpan: 1, rowSpan: 1, continuation: false },
    ],
    [
      { gridSpan: 1, rowSpan: 1, continuation: true },
      { gridSpan: 1, rowSpan: 1, continuation: false },
    ],
  ]);
});

// ── 2b. collapsed grid export/reload ────────────────────────────────

test("a column insert keeps the collapsed marker on the zero-width column", () => {
  const table = collapsedTable();
  assert.equal(core.insertTableColumn(table, 0, "left"), true);
  assert.deepEqual(table.style.collapsedGridColumns, [2]);
  assert.deepEqual(table.style.columnWidthsTwips, [900, 900, 0, 1800]);

  const prepared = core.prepareModelForSerialization(modelOf([table]));
  assert.deepEqual(prepared.nodes[0].style.columnWidthsTwips, [900, 900, 1800]);
  assert.equal(prepared.nodes[0].style.collapsedGridColumns, undefined);
  assert.deepEqual(
    rowSignature(prepared.nodes[0]),
    [
      [1, 1, 1].map((gridSpan) => ({ gridSpan, rowSpan: 1, continuation: false })),
      [1, 1, 1].map((gridSpan) => ({ gridSpan, rowSpan: 1, continuation: false })),
    ],
  );
});

test("an edited collapsed table exports its physical grid and reloads cleanly", async () => {
  const model = modelOf([collapsedTable()]);

  const documentXml = await core.modelToDocumentXml(model);
  const tableXml = tableXmls(documentXml)[0];
  assert.match(tableXml, /<w:gridCol w:w="900"\/><w:gridCol w:w="1800"\/>/);
  assert.doesNotMatch(tableXml, /<w:gridCol w:w="0"\/>/);
  assert.doesNotMatch(tableXml, /<w:gridSpan/);

  const exported = await core.serializeDocx(model);
  const reimported = (await core.buildDocModelFromBytes(exported)).model;
  const table = reimported.nodes[0];
  assert.deepEqual(table.style.columnWidthsTwips, [900, 1800]);
  assert.equal(table.style.collapsedGridColumns, undefined);
  assert.deepEqual(rowSignature(table), [
    [
      { gridSpan: 1, rowSpan: 1, continuation: false },
      { gridSpan: 1, rowSpan: 1, continuation: false },
    ],
    [
      { gridSpan: 1, rowSpan: 1, continuation: false },
      { gridSpan: 1, rowSpan: 1, continuation: false },
    ],
  ]);
});

test("a collapsed grid imported from a document is projected after an edit", async () => {
  const tableXml =
    '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblLayout w:type="fixed"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="900"/><w:gridCol w:w="0"/><w:gridCol w:w="1800"/></w:tblGrid>' +
    '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Wide</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:p><w:r><w:t>R</w:t></w:r></w:p></w:tc></w:tr>' +
    '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Left</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:p><w:r><w:t>R2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  const pkg = core.createMinimalDocxPackage(documentXml(tableXml));
  const model = await core.buildDocModel(pkg);

  // Import reattaches the logical grid and its collapsed-column marker.
  const imported = model.nodes[0];
  assert.deepEqual(imported.style.columnWidthsTwips, [900, 0, 1800]);
  assert.deepEqual(imported.style.collapsedGridColumns, [1]);
  assert.equal(imported.style.styleId, "TableGrid");
  assert.equal(imported.style.sourceStyleId, "TableGrid");

  // The untouched table is written from its stored markup, zero column and all.
  const untouched = documentXmlOf(await core.serializeDocx(model, pkg));
  assert.match(tableXmls(untouched)[0], /<w:gridCol w:w="0"\/>/);

  // Editing the cell clears the stored markup; export now projects.
  const edited = core.updateTableCellParagraphText(model, 0, 1, 0, 0, "Left edited");
  const exported = await core.serializeDocx(edited, pkg);
  const tableAfterEdit = tableXmls(documentXmlOf(exported))[0];
  assert.match(tableAfterEdit, /<w:gridCol w:w="900"\/><w:gridCol w:w="1800"\/>/);
  assert.doesNotMatch(tableAfterEdit, /<w:gridCol w:w="0"\/>/);
  assert.match(tableAfterEdit, /Left edited/);

  const reimported = (await core.buildDocModelFromBytes(exported)).model;
  assert.deepEqual(reimported.nodes[0].style.columnWidthsTwips, [900, 1800]);
  assert.deepEqual(rowSignature(reimported.nodes[0]), [
    [
      { gridSpan: 1, rowSpan: 1, continuation: false },
      { gridSpan: 1, rowSpan: 1, continuation: false },
    ],
    [
      { gridSpan: 1, rowSpan: 1, continuation: false },
      { gridSpan: 1, rowSpan: 1, continuation: false },
    ],
  ]);
});

// ── 1b. header/footer stories survive export ────────────────────────

test("a merged table inside a header exports and reloads with its grid", async () => {
  const { model, package: pkg } = await core.buildDocModelFromBytes(
    sampleBytes("../../apps/demo/public/samples/demo.docx"),
  );
  let next = core.addHeaderFooter(model, 0, "header", { referenceType: "default" });
  const partName = next.metadata.headerSections.at(-1).partName;

  const headerTable = collapsedTable();
  headerTable.rows[0].cells[0].style.rowSpan = 2;
  headerTable.rows[0].cells[0].nodes = [paragraph("Header merged")];
  headerTable.rows[1].cells[0].style.vMergeContinuation = true;
  const nodes = [headerTable, paragraph("Header tail")];
  next = structuredClone(next);
  next.metadata.headerSections = next.metadata.headerSections.map((section) =>
    section.partName === partName ? { ...section, nodes } : section,
  );
  next.metadata.sections = next.metadata.sections.map((section) => ({
    ...section,
    headerSections: section.headerSections.map((entry) =>
      entry.partName === partName ? { ...entry, nodes } : entry,
    ),
  }));

  const exported = await core.serializeDocx(next, pkg);
  const archive = unzipSync(new Uint8Array(exported));
  const headerXml = strFromU8(archive[partName]);
  const tableXml = tableXmls(headerXml)[0];
  assert.match(tableXml, /<w:gridCol w:w="900"\/><w:gridCol w:w="1800"\/>/);
  assert.match(tableXml, /<w:vMerge w:val="restart"\/>/);
  assert.match(tableXml, /<w:vMerge\/>/);

  // The package-returning serializer patches the story parts too.
  const packageExport = await core.serializeDocModel(next, pkg);
  const packageHeader = packageExport.parts.get(partName);
  assert.ok(packageHeader);
  assert.match(packageHeader.content, /<w:vMerge w:val="restart"\/>/);
  assert.match(packageHeader.content, /<w:gridCol w:w="1800"\/>/);

  const reimported = (await core.buildDocModelFromBytes(exported)).model;
  const reloadedHeader = reimported.metadata.headerSections.find(
    (section) => section.partName === partName,
  );
  assert.ok(reloadedHeader);
  const reloadedTable = findTable(reloadedHeader.nodes);
  assert.ok(reloadedTable);
  assert.deepEqual(reloadedTable.style.columnWidthsTwips, [900, 1800]);
  assert.deepEqual(rowSignature(reloadedTable), [
    [
      { gridSpan: 1, rowSpan: 2, continuation: false },
      { gridSpan: 1, rowSpan: 1, continuation: false },
    ],
    [
      { gridSpan: 1, rowSpan: 1, continuation: true },
      { gridSpan: 1, rowSpan: 1, continuation: false },
    ],
  ]);
});
