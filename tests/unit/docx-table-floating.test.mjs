// Issue #150 — DOCX table move handle and floating-table drag.
//
// Every block names the bullet it covers so the mapping stays checkable:
//
//   1. drop → `tblpPr` patch (position + wrap) ........ resolveTableFloatingDropPatch
//   2. patch merge / refusal of bad values ............ mergeTableFloatingStyle
//   3. wrapper style renders the drop ................. tableWrapperStyle
//   4. `tblpPr` is written on export .................. modelToDocumentXml / serializeDocx
//   5. export → reopen keeps floating ................. buildDocModelFromBytes
//   6. an edited floating table keeps its `tblpPr` .... updateTableCellParagraphText
//   7. inline again writes no `tblpPr` ................ cleared TableStyle.floating
//   8. nested tables are aligned, never floated ....... depth-first table order
//   9. header/footer stories keep `tblpPr` ............ serializeDocx / serializeDocModel
//
// The acceptance case is a table dropped at a position, exported, reopened and
// then put back into the text flow.

import assert from "node:assert/strict";
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

// ── Fixtures ───────────────────────────────────────────────────────
function paragraph(text) {
  return { type: "paragraph", children: [{ type: "text", text }] };
}

function cell(text, style = {}) {
  return { type: "table-cell", style: { widthTwips: 1500, ...style }, nodes: [paragraph(text)] };
}

function tableNode({ floating, rows = 2, style = {} } = {}) {
  return {
    type: "table",
    style: {
      layout: "fixed",
      widthTwips: 3000,
      columnWidthsTwips: [1500, 1500],
      ...(floating ? { floating } : {}),
      ...style,
    },
    rows: Array.from({ length: rows }, (_, rowIndex) => ({
      type: "table-row",
      cells: [cell(`r${rowIndex}c0`), cell(`r${rowIndex}c1`)],
    })),
  };
}

function modelOf(nodes, metadata = {}) {
  return {
    ...core.createBlankDocumentModel(),
    nodes,
    metadata: { ...core.createBlankDocumentModel().metadata, ...metadata },
  };
}

/** A stored floating table, the way Word writes one, as an import fixture. */
function floatingTableXml(xmlAttributes = 'w:vertAnchor="text" w:horzAnchor="margin" w:tblpX="720" w:tblpY="360"') {
  return (
    `<w:tbl><w:tblPr><w:tblpPr ${xmlAttributes}/>` +
    '<w:tblW w:w="3000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>' +
    '<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>' +
    '<w:tr><w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  );
}

function documentXml(body) {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<w:document ${WORD_NS}><w:body>${body}<w:sectPr>${SECTION_PROPS}</w:sectPr></w:body></w:document>`
  );
}

function tblpPrOf(xml) {
  return /<w:tblpPr\b[^>]*\/>/.exec(xml)?.[0];
}

function tblpPrMatches(xml) {
  return xml.match(/<w:tblpPr\b[^>]*\/>/g) ?? [];
}

function documentXmlOf(bytes) {
  return strFromU8(unzipSync(new Uint8Array(bytes))["word/document.xml"]);
}

const DROP_FLOATING = {
  xTwips: 1440,
  yTwips: 720,
  horizontalAnchor: "margin",
  verticalAnchor: "text",
  leftFromTextTwips: 180,
  rightFromTextTwips: 180,
  topFromTextTwips: 0,
  bottomFromTextTwips: 0,
};

// ── 1. drop position → `tblpPr` patch ──────────────────────────────
test("a drop position becomes tblpX/tblpY twips anchored to the margin", () => {
  // 96px right and 24px down from the text anchor: 1440 / 360 twips.
  assert.deepEqual(
    core.resolveTableFloatingDropPatch({ deltaXPx: 96, deltaYPx: 24 }),
    {
      xTwips: 1440,
      yTwips: 360,
      horizontalAnchor: "margin",
      verticalAnchor: "text",
    },
  );

  // A second drag keeps the offsets the table already has.
  assert.deepEqual(
    core.resolveTableFloatingDropPatch({
      startXPx: 96,
      startYPx: 24,
      deltaXPx: -32,
      deltaYPx: 8,
    }),
    {
      xTwips: 960,
      yTwips: 480,
      horizontalAnchor: "margin",
      verticalAnchor: "text",
    },
  );

  // Travelling to the far right keeps a reachable sliver on the page.
  const clamped = core.resolveTableFloatingDropPatch({
    deltaXPx: 4000,
    deltaYPx: 0,
    pageWidthPx: 816,
    tableWidthPx: 300,
  });
  assert.equal(clamped.xTwips, (816 - 42) * 15);

  // Dragging past the left edge keeps most of the table on the page.
  const outdented = core.resolveTableFloatingDropPatch({
    deltaXPx: -4000,
    deltaYPx: 0,
    pageWidthPx: 816,
    tableWidthPx: 300,
  });
  assert.equal(outdented.xTwips, (42 - 300) * 15);
});

// ── 2. patch merge and refusal of unrepresentable values ───────────
test("a floating patch merges over the current wrap distances", () => {
  const merged = core.mergeTableFloatingStyle(
    {
      leftFromTextTwips: 180,
      rightFromTextTwips: 180,
      horizontalAnchor: "margin",
      xTwips: 720,
    },
    { xTwips: 1440.4, yTwips: 360 },
  );
  assert.deepEqual(merged, {
    leftFromTextTwips: 180,
    rightFromTextTwips: 180,
    horizontalAnchor: "margin",
    xTwips: 1440,
    yTwips: 360,
  });

  // Non-finite numbers and unknown anchor/alignment names never reach `tblpPr`.
  assert.deepEqual(
    core.mergeTableFloatingStyle(undefined, {
      xTwips: Number.NaN,
      yTwips: Number.POSITIVE_INFINITY,
      horizontalAnchor: "somewhere",
      verticalAnchor: "margin",
      horizontalAlign: "diagonal",
      verticalAlign: "top",
    }),
    { verticalAnchor: "margin", verticalAlign: "top" },
  );

  assert.equal(core.mergeTableFloatingStyle(undefined, {}), undefined);
});

// ── 3. the wrapper renders the drop ────────────────────────────────
test("the wrapper style shifts a floating table by its drop offsets", () => {
  const inline = core.tableWrapperStyle(tableNode(), 0);
  assert.equal(inline.float, undefined);
  assert.equal(inline.marginLeft, 0);

  // 1440 twips = 96px; the text distance stays on top of the offset.
  const dropped = core.tableWrapperStyle(
    tableNode({ floating: DROP_FLOATING }),
    0,
  );
  assert.equal(dropped.float, "left");
  assert.equal(dropped.marginLeft, 96 + 12);
  assert.equal(dropped.marginTop, 48);

  // A right-anchored table shifts from the right margin instead.
  const rightDropped = core.tableWrapperStyle(
    tableNode({
      floating: {
        horizontalAlign: "right",
        xTwips: 1440,
        yTwips: 0,
        leftFromTextTwips: 0,
        rightFromTextTwips: 0,
      },
    }),
    0,
  );
  assert.equal(rightDropped.float, "right");
  assert.equal(rightDropped.marginRight, 96);
  assert.equal(rightDropped.marginLeft, 0);

  // A table dropped far right stays a left-anchored float: `tblpX` is measured
  // from the margin, it does not mean the table is right-anchored.
  const farRightDrop = core.tableWrapperStyle(
    tableNode({
      floating: {
        horizontalAnchor: "margin",
        xTwips: 4800,
        yTwips: 0,
        leftFromTextTwips: 0,
      },
    }),
    0,
  );
  assert.equal(farRightDrop.float, "left");
  assert.equal(farRightDrop.marginLeft, 320);

  // Without an anchor the imported `tblpX` heuristic still picks a side.
  const legacyRight = core.tableWrapperStyle(
    tableNode({ floating: { xTwips: 4800, leftFromTextTwips: 0 } }),
    0,
  );
  assert.equal(legacyRight.float, "right");
});

// ── 4. export writes `w:tblpPr` ────────────────────────────────────
test("export writes w:tblpPr as the first tblPr child", async () => {
  const xml = await core.modelToDocumentXml(
    modelOf([tableNode({ floating: DROP_FLOATING })]),
  );
  const tableXml = /<w:tbl>[\s\S]*?<\/w:tbl>/.exec(xml)?.[0] ?? "";
  assert.match(
    tableXml,
    /^<w:tbl><w:tblPr><w:tblpPr w:leftFromText="180" w:rightFromText="180" w:topFromText="0" w:bottomFromText="0" w:vertAnchor="text" w:horzAnchor="margin" w:tblpX="1440" w:tblpY="720"\/><w:tblW/,
  );
});

test("a table with anchors only still writes its wrap distances", async () => {
  const xml = await core.modelToDocumentXml(
    modelOf([tableNode({ floating: { horizontalAlign: "right", verticalAlign: "bottom" } })]),
  );
  assert.equal(tblpPrOf(xml), '<w:tblpPr w:tblpXSpec="right" w:tblpYSpec="bottom"/>');
});

// ── 5. export → reopen keeps the floating table ────────────────────
test("a dropped table survives export and reopen", async () => {
  const model = modelOf([tableNode({ floating: DROP_FLOATING }), paragraph("after")]);
  const bytes = await core.serializeDocx(model);
  assert.equal(tblpPrMatches(documentXmlOf(bytes)).length, 1);

  const reimported = (await core.buildDocModelFromBytes(bytes)).model;
  assert.equal(reimported.nodes[0].type, "table");
  assert.deepEqual(reimported.nodes[0].style.floating, DROP_FLOATING);
});

// ── 6. an edited floating table keeps its position ─────────────────
test("editing a floating table keeps its tblpPr through export", async () => {
  const pkg = core.createMinimalDocxPackage(documentXml(floatingTableXml()));
  const model = await core.buildDocModel(pkg);
  assert.equal(model.nodes[0].type, "table");
  assert.equal(model.nodes[0].style.floating.xTwips, 720);
  assert.equal(model.nodes[0].style.floating.yTwips, 360);

  // A cell edit rebuilds the table from the model and drops `sourceXml`.
  const edited = core.updateTableCellParagraphText(model, 0, 1, 1, 0, "D edited");
  assert.equal(edited.nodes[0].sourceXml, undefined);

  const xml = await core.modelToDocumentXml(edited, pkg);
  assert.equal(
    tblpPrOf(xml),
    '<w:tblpPr w:vertAnchor="text" w:horzAnchor="margin" w:tblpX="720" w:tblpY="360"/>',
  );
  assert.match(xml, /D edited/);

  const reopened = (await core.buildDocModelFromBytes(await core.serializeDocx(edited, pkg))).model;
  assert.deepEqual(reopened.nodes[0].style.floating, model.nodes[0].style.floating);
  assert.equal(core.paragraphText(reopened.nodes[0].rows[1].cells[1].nodes[0]), "D edited");
});

// ── 7. back inline: no `tblpPr` anywhere ───────────────────────────
test("clearing the floating style exports the table inline again", async () => {
  const pkg = core.createMinimalDocxPackage(documentXml(floatingTableXml()));
  const model = await core.buildDocModel(pkg);
  const next = core.cloneDocModel(model);
  const table = next.nodes[0];
  const { floating: _floating, ...style } = table.style;
  table.style = style;
  table.sourceXml = undefined;

  const xml = await core.modelToDocumentXml(next, pkg);
  assert.equal(tblpPrOf(xml), undefined);
  assert.match(xml, /<w:tblPr><w:tblW/);

  const reopened = (await core.buildDocModelFromBytes(await core.serializeDocx(next, pkg))).model;
  assert.equal(reopened.nodes[0].style.floating, undefined);
});

// ── 8. nested tables are aligned, never floated ────────────────────
test("a nested table inside a floating table gets no tblpPr of its own", async () => {
  const outer = tableNode({ floating: DROP_FLOATING, rows: 1 });
  const inner = tableNode({ rows: 1, style: { widthTwips: 1000, columnWidthsTwips: [500, 500] } });
  outer.rows[0].cells[0].nodes.push(inner);

  const xml = await core.modelToDocumentXml(modelOf([outer]));
  const matches = tblpPrMatches(xml);
  assert.equal(matches.length, 1, "only the outer table floats");
  const innerXml = /<w:tbl><w:tblPr>(?:(?!<w:tblPr>)[\s\S])*?$/.exec(xml)?.[0] ?? "";
  assert.doesNotMatch(innerXml.slice(0, 200), /<w:tblpPr/);
});

// ── 9. stories: headers, footers and notes keep `tblpPr` ───────────
test("a floating table inside a header keeps its position on export", async () => {
  const { model, package: pkg } = await core.buildDocModelFromBytes(
    readSample("../../apps/demo/public/samples/demo.docx"),
  );
  let next = core.addHeaderFooter(model, 0, "header", { referenceType: "default" });
  const partName = next.metadata.headerSections.at(-1).partName;

  const headerTable = tableNode({ floating: DROP_FLOATING, rows: 1 });
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

  const archive = unzipSync(new Uint8Array(await core.serializeDocx(next, pkg)));
  const headerXml = strFromU8(archive[partName]);
  assert.match(headerXml, /<w:tblpPr\b[^>]*w:tblpX="1440"/);

  // The package-returning serializer patches the story parts too.
  const packageExport = await core.serializeDocModel(next, pkg);
  assert.match(packageExport.parts.get(partName).content, /<w:tblpPr\b[^>]*w:tblpX="1440"/);

  const reopened = (await core.buildDocModelFromBytes(await core.serializeDocx(next, pkg))).model;
  const header = reopened.metadata.headerSections.find(
    (section) => section.partName === partName,
  );
  assert.ok(header);
  const table = header.nodes.find((node) => node.type === "table");
  assert.ok(table);
  assert.equal(table.style.floating.xTwips, 1440);
  assert.equal(table.style.floating.horizontalAnchor, "margin");
});

test("a stored floating table is written verbatim, not patched twice", async () => {
  const stored = tableNode({ floating: { xTwips: 9999, yTwips: 0 }, rows: 1 });
  stored.sourceXml =
    '<w:tbl><w:tblPr><w:tblpPr w:vertAnchor="text" w:horzAnchor="margin" w:tblpX="480"/>' +
    '<w:tblW w:w="3000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>' +
    '<w:tr><w:tc><w:p><w:r><w:t>stored</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';

  const xml = await core.modelToDocumentXml(modelOf([stored]));
  const matches = tblpPrMatches(xml);
  assert.equal(matches.length, 1, "the stored markup is not patched a second time");
  assert.match(matches[0], /w:tblpX="480"/);
});

function readSample(relativePath) {
  const { readFileSync } = requireFromDocxCore("node:fs");
  const bytes = readFileSync(new URL(relativePath, import.meta.url));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}
