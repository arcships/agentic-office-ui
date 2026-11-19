import assert from "node:assert/strict";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

const requireFromVueDocx = createRequire(
  new URL("../../packages/vue-docx/package.json", import.meta.url),
);
const core = await import(
  pathToFileURL(requireFromVueDocx.resolve("@arcships/docx-core")).href
);

// Same acceptance shape the viewer passes: the Office URL policy decides which
// link targets survive a paste.
const sanitizeLink = (value) =>
  value.startsWith("https://") || value.startsWith("mailto:") || value.startsWith("#")
    ? value
    : undefined;

function starterModel(nodes) {
  return { ...core.createBlankDocumentModel(), nodes };
}

function paragraph(text, extra = {}) {
  return { type: "paragraph", children: [{ type: "text", text }], ...extra };
}

function caret(nodeIndex, offset) {
  return {
    start: { location: { kind: "paragraph", nodeIndex }, offset },
    end: { location: { kind: "paragraph", nodeIndex }, offset },
  };
}

function range(startOffset, endOffset, nodeIndex = 0) {
  return {
    start: { location: { kind: "paragraph", nodeIndex }, offset: startOffset },
    end: { location: { kind: "paragraph", nodeIndex }, offset: endOffset },
  };
}

function textOf(node) {
  return node.children
    .filter((child) => child.type === "text")
    .map((child) => child.text)
    .join("");
}

function runsOf(node) {
  return node.children.filter((child) => child.type === "text");
}

test("clipboard paragraphs keep run styles and links across a body selection", () => {
  const model = starterModel([
    {
      type: "paragraph",
      style: { headingLevel: 1 },
      children: [
        { type: "text", text: "Title ", style: { bold: true } },
        { type: "text", text: "link", link: "https://example.com/a" },
      ],
    },
    paragraph("middle"),
    {
      type: "paragraph",
      style: { align: "center", numbering: { numId: 4, ilvl: 1 } },
      children: [{ type: "text", text: "tail", style: { italic: true } }],
    },
  ]);

  const paragraphs = core.docxClipboardParagraphsForRange(model, {
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 6 },
    end: { location: { kind: "paragraph", nodeIndex: 2 }, offset: 2 },
  });
  assert.ok(paragraphs);
  assert.equal(paragraphs.length, 3);
  assert.equal(paragraphs[0].style.headingLevel, 1, "heading level travels with the selection");
  assert.deepEqual(paragraphs[0].runs, [
    { text: "link", link: "https://example.com/a" },
  ]);
  assert.equal(paragraphs[1].runs[0].text, "middle");
  assert.deepEqual(paragraphs[2].style, { align: "center" });
  assert.deepEqual(paragraphs[2].list, { type: "unordered", level: 1 });
  assert.deepEqual(paragraphs[2].runs, [{ text: "ta", style: { italic: true } }]);

  assert.equal(
    core.docxClipboardParagraphsForRange(model, range(1, 1)),
    undefined,
    "a collapsed range has nothing to copy",
  );
});

test("clipboard paragraphs read one table cell and reject cross-cell selections", () => {
  const model = starterModel([
    {
      type: "table",
      rows: [
        {
          cells: [
            {
              nodes: [
                {
                  type: "paragraph",
                  children: [{ type: "text", text: "cell one", style: { underline: true } }],
                },
                { type: "paragraph", children: [{ type: "text", text: "cell two" }] },
              ],
            },
            { nodes: [paragraph("other cell")] },
          ],
        },
      ],
    },
  ]);

  const paragraphs = core.docxClipboardParagraphsForRange(model, {
    start: {
      location: { kind: "table-cell", tableIndex: 0, rowIndex: 0, cellIndex: 0, paragraphIndex: 0 },
      offset: 5,
    },
    end: {
      location: { kind: "table-cell", tableIndex: 0, rowIndex: 0, cellIndex: 0, paragraphIndex: 1 },
      offset: 4,
    },
  });
  assert.ok(paragraphs);
  assert.deepEqual(paragraphs.map((entry) => entry.runs.map((run) => run.text)), [
    ["one"],
    ["cell"],
  ]);
  assert.deepEqual(paragraphs[0].runs[0].style, { underline: true });

  assert.equal(
    core.docxClipboardParagraphsForRange(model, {
      start: {
        location: { kind: "table-cell", tableIndex: 0, rowIndex: 0, cellIndex: 0, paragraphIndex: 0 },
        offset: 0,
      },
      end: {
        location: { kind: "table-cell", tableIndex: 0, rowIndex: 0, cellIndex: 1, paragraphIndex: 0 },
        offset: 2,
      },
    }),
    undefined,
    "a selection spanning two cells cannot be represented as one cell fragment",
  );
});

test("the internal payload round-trips range, runs and plain text", () => {
  const sourceRange = range(1, 4);
  const paragraphs = [
    {
      runs: [
        { text: "bo", style: { bold: true } },
        { text: "ld", link: "https://example.com/b" },
      ],
      style: { headingLevel: 2 },
    },
    { runs: [{ text: "next", style: { color: "#ff0000", fontSizePt: 14 } }] },
  ];
  const payload = core.createDocxClipboardPayload({ paragraphs, range: sourceRange });
  assert.equal(payload.version, core.DOCX_CLIPBOARD_VERSION);
  assert.equal(payload.kind, "docx-fragment");
  assert.equal(payload.text, "bold\nnext");
  assert.deepEqual(payload.range, sourceRange);

  const raw = core.serializeDocxClipboardPayload(payload);
  const parsed = core.parseDocxClipboardPayload(raw, { sanitizeLink });
  assert.ok(parsed);
  assert.equal(parsed.text, "bold\nnext");
  assert.deepEqual(parsed.range, sourceRange);
  assert.deepEqual(parsed.paragraphs[0].runs, [
    { text: "bo", style: { bold: true } },
    { text: "ld", link: "https://example.com/b" },
  ]);
  assert.deepEqual(parsed.paragraphs[0].style, { headingLevel: 2 });
  assert.deepEqual(parsed.paragraphs[1].runs, [
    { text: "next", style: { color: "#ff0000", fontSizePt: 14 } },
  ]);
});

test("the internal payload rejects unknown shapes, unsafe links and oversized data", () => {
  const unsafe = JSON.stringify({
    version: core.DOCX_CLIPBOARD_VERSION,
    kind: "docx-fragment",
    text: "danger",
    paragraphs: [
      {
        runs: [
          { text: "danger", link: "javascript:alert(1)", style: { color: "url(http://x)" } },
        ],
      },
    ],
  });
  const parsed = core.parseDocxClipboardPayload(unsafe, { sanitizeLink });
  assert.ok(parsed);
  assert.deepEqual(parsed.paragraphs[0].runs, [{ text: "danger" }], "unsafe link and CSS are dropped");

  assert.equal(core.parseDocxClipboardPayload("{", { sanitizeLink }), undefined);
  assert.equal(
    core.parseDocxClipboardPayload(JSON.stringify({ version: 99, kind: "docx-fragment", paragraphs: [] })),
    undefined,
    "an unknown schema version is ignored",
  );
  assert.equal(
    core.parseDocxClipboardPayload(unsafe, { sanitizeLink, maxPayloadChars: 8 }),
    undefined,
    "an oversized payload is ignored instead of parsed",
  );
  assert.equal(
    core.parseDocxClipboardPayload(
      JSON.stringify({ version: core.DOCX_CLIPBOARD_VERSION, kind: "docx-fragment", paragraphs: [] })
    ),
    undefined,
    "a payload without paragraphs cannot paste",
  );
});

test("text/html carries inline run styles, links, headings and list markup", () => {
  const html = core.renderDocxClipboardHtml([
    {
      runs: [{ text: "Heading", style: { bold: true, fontSizePt: 18 } }],
      style: { headingLevel: 1 },
    },
    {
      runs: [
        { text: "<b> & ", style: { italic: true, underline: true, strike: true } },
        { text: "link", link: "https://example.com/a", style: { color: "#0000ff" } },
        { text: "sup", style: { verticalAlign: "superscript" } },
      ],
    },
    { runs: [{ text: "one" }], list: { type: "ordered", level: 0 } },
    { runs: [{ text: "two" }], list: { type: "ordered", level: 0 } },
  ]);

  assert.match(html, /<h1 style="margin:0;font-size:18pt;font-weight:bold">/);
  assert.match(html, /&lt;b&gt; &amp;/, "text is escaped");
  assert.match(html, /font-style:italic/);
  assert.match(html, /text-decoration:underline line-through/);
  assert.match(html, /<a href="https:\/\/example.com\/a" style="color:#0000ff">link<\/a>/);
  assert.match(html, /vertical-align:super/);
  assert.match(html, /<ol><li style="margin:0">one<\/li><li style="margin:0">two<\/li><\/ol>/);
});

test("text/html paste maps Word-style markup back onto runs", () => {
  const html =
    '<p style="text-align:center"><b>Bold</b> <i>ital</i> <u>under</u> <s>strike</s> ' +
    '<a href="https://example.com/x">link</a> <a href="javascript:alert(1)">bad</a> ' +
    '<span style="color:#ff0000;background-color:yellow;font-size:14pt;font-family:\'Times New Roman\'">styled</span> ' +
    '<font color="00FF00" face="Courier" size="5">legacy</font><sup>2</sup><sub>n</sub></p>' +
    '<h2>Heading</h2><ul><li>one</li><li><p>two</p></li></ul>';

  const blocks = core.parseDocxClipboardHtml(html, { sanitizeLink });
  assert.equal(blocks.length, 4);

  const [first] = blocks;
  const run = (text) => first.runs.find((entry) => entry.text.trim() === text);
  assert.deepEqual(run("Bold").style, { bold: true });
  assert.deepEqual(run("ital").style, { italic: true });
  assert.deepEqual(run("under").style, { underline: true });
  assert.deepEqual(run("strike").style, { strike: true });
  assert.equal(run("link").link, "https://example.com/x");
  assert.equal(
    first.runs.find((entry) => entry.text.includes("bad"))?.link,
    undefined,
    "an unsafe href keeps its text but loses the link",
  );
  assert.deepEqual(run("styled").style, {
    backgroundColor: "yellow",
    color: "#ff0000",
    fontFamily: "Times New Roman",
    fontSizePt: 14,
  });
  assert.deepEqual(run("legacy").style, {
    color: "#00ff00",
    fontFamily: "Courier",
    fontSizePt: 18,
  });
  assert.deepEqual(run("2").style, { verticalAlign: "superscript" });
  assert.deepEqual(run("n").style, { verticalAlign: "subscript" });
  assert.deepEqual(first.style, { align: "center" });

  assert.equal(blocks[1].style.headingLevel, 2);
  assert.deepEqual(blocks[1].runs, [{ text: "Heading", style: { bold: true, fontSizePt: 16 } }]);
  assert.deepEqual(blocks[2].list, { type: "unordered", level: 0 });
  assert.equal(blocks[2].runs[0].text, "one");
  assert.deepEqual(blocks[3].list, { type: "unordered", level: 0 });
  assert.equal(blocks[3].runs[0].text, "two");
});

test("text/html parsing drops executable content and keeps entity text", () => {
  const blocks = core.parseDocxClipboardHtml(
    '<div><script>alert(1)</script><style>p{color:red}</style><iframe src="https://x"></iframe>' +
      '<svg><text>vector</text></svg><object data="x"></object>' +
      '<p>&lt;tag&gt; &amp; &#39;quote&#39; &nbsp;end</p>' +
      '<p><img src="x" onerror="alert(1)">visible</p></div>',
    { sanitizeLink },
  );

  const text = blocks.map((block) => block.runs.map((entry) => entry.text).join(""));
  assert.deepEqual(text, ["<tag> & 'quote' \u00a0end", "visible"]);
});

test("text/html parsing caps oversized markup instead of parsing all of it", () => {
  const longHtml = `<p>${"a".repeat(200)}</p><p>b</p>`;
  const blocks = core.parseDocxClipboardHtml(longHtml, { sanitizeLink, maxHtmlChars: 20 });
  assert.ok(blocks.length <= 1, "content past the cap is not parsed");

  const limited = core.parseDocxClipboardHtml("<p>one</p><p>two</p><p>three</p>", {
    sanitizeLink,
    maxParagraphs: 2,
  });
  assert.deepEqual(
    limited.map((block) => block.runs.map((run) => run.text).join("")),
    ["one", "two"],
  );
});

test("rendered HTML parses back into the same runs", () => {
  const blocks = [
    {
      runs: [
        { text: "bold ", style: { bold: true } },
        { text: "link", link: "https://example.com/c" },
        { text: " plain" },
      ],
    },
    { runs: [{ text: "quoted \"text\"" }] },
    { runs: [{ text: "bullet" }], list: { type: "unordered", level: 0 } },
  ];
  const html = core.renderDocxClipboardHtml(blocks);
  const parsed = core.parseDocxClipboardHtml(html, { sanitizeLink });
  assert.deepEqual(parsed, blocks);
  assert.equal(core.clipboardTextFromParagraphs(parsed), core.clipboardTextFromParagraphs(blocks));
});

test("plain text becomes one clipboard paragraph per line", () => {
  assert.deepEqual(core.clipboardParagraphsFromText("one\r\ntwo\n"), [
    { runs: [{ text: "one" }] },
    { runs: [{ text: "two" }] },
    { runs: [] },
  ]);
  assert.equal(core.clipboardTextFromParagraphs([{ runs: [{ text: "a" }] }, { runs: [] }]), "a\n");
});

test("paste inserts styled runs at a caret and keeps the surrounding text", () => {
  const model = starterModel([paragraph("before after")]);
  const blocks = [
    {
      runs: [
        { text: "bold", style: { bold: true } },
        { text: "+link", link: "https://example.com/d" },
      ],
    },
  ];

  const result = core.insertDocxClipboardParagraphs(model, caret(0, 7), blocks);
  assert.ok(result);
  const node = result.model.nodes[0];
  assert.equal(textOf(node), "before bold+linkafter");
  assert.deepEqual(runsOf(node).map((run) => run.text), ["before ", "bold", "+link", "after"]);
  assert.deepEqual(runsOf(node)[1].style, { bold: true });
  assert.equal(runsOf(node)[2].link, "https://example.com/d");
  assert.equal(model.nodes[0].children.length, 1, "the source model is untouched");
  assert.deepEqual(result.caret, caret(0, 16));
});

test("paste replaces an intra-paragraph range and merges comment offsets", () => {
  const model = starterModel([
    {
      type: "paragraph",
      commentRanges: [{ start: 1, end: 4, commentId: "c1" }],
      children: [{ type: "text", text: "abcdef" }],
    },
  ]);
  const result = core.insertDocxClipboardParagraphs(
    model,
    range(2, 4),
    [{ runs: [{ text: "XY", style: { italic: true } }] }],
  );
  assert.ok(result);
  const node = result.model.nodes[0];
  assert.equal(textOf(node), "abXYef");
  assert.equal(node.commentRanges.length, 1);
  assert.equal(node.commentRanges[0].commentId, "c1");
  assert.equal(node.commentRanges[0].start, 1);
  assert.equal(node.commentRanges[0].end, 4, "comment offsets follow the replaced text");
});

test("multi-block paste splits the paragraph, keeps headings and applies list numbering", () => {
  const model = starterModel([paragraph("tail")]);
  const blocks = [
    { runs: [{ text: "first", style: { bold: true } }] },
    {
      runs: [{ text: "heading" }],
      style: { headingLevel: 2, align: "center" },
    },
    { runs: [{ text: "item one" }], list: { type: "unordered", level: 0 } },
    { runs: [{ text: "item two" }], list: { type: "unordered", level: 0 } },
  ];

  const result = core.insertDocxClipboardParagraphs(model, caret(0, 0), blocks);
  assert.ok(result);
  assert.equal(result.model.nodes.length, 4);
  assert.equal(textOf(result.model.nodes[0]), "first");
  assert.deepEqual(runsOf(result.model.nodes[1])[0].style, undefined);

  const heading = result.model.nodes[1];
  assert.equal(textOf(heading), "heading");
  assert.equal(heading.style.headingLevel, 2);
  assert.equal(heading.style.align, "center");

  const firstItem = result.model.nodes[2];
  assert.equal(textOf(firstItem), "item one");
  assert.ok(firstItem.style.numbering.numId > 0);
  assert.equal(firstItem.style.numbering.ilvl, 0);
  assert.equal(
    result.model.nodes[3].style.numbering.numId,
    firstItem.style.numbering.numId,
    "sibling list items share the numbering definition",
  );
  assert.equal(textOf(result.model.nodes[3]), "item twotail");
  assert.equal(
    result.model.metadata.numberingDefinitions.instances.length,
    1,
    "a pasted list creates one numbering instance",
  );

  // The trailing text after a mid-paragraph paste stays in the last paragraph.
  const middle = core.insertDocxClipboardParagraphs(starterModel([paragraph("leftright")]), caret(0, 4), blocks);
  assert.ok(middle);
  assert.equal(textOf(middle.model.nodes[0]), "leftfirst");
  assert.equal(middle.model.nodes.length, 4);
  assert.equal(textOf(middle.model.nodes[3]), "item tworight");
  assert.deepEqual(middle.caret, caret(3, "item two".length));
});

test("paste into a table cell keeps one paragraph with newline-separated blocks", () => {
  const model = starterModel([
    {
      type: "table",
      rows: [
        {
          cells: [
            { nodes: [paragraph("cell")] },
            { nodes: [paragraph("other")] },
          ],
        },
      ],
    },
  ]);
  const cellRange = {
    start: {
      location: { kind: "table-cell", tableIndex: 0, rowIndex: 0, cellIndex: 0, paragraphIndex: 0 },
      offset: 0,
    },
    end: {
      location: { kind: "table-cell", tableIndex: 0, rowIndex: 0, cellIndex: 0, paragraphIndex: 0 },
      offset: 0,
    },
  };
  const result = core.insertDocxClipboardParagraphs(model, cellRange, [
    { runs: [{ text: "one", style: { bold: true } }] },
    { runs: [{ text: "two" }] },
  ]);
  assert.ok(result);
  const table = result.model.nodes[0];
  const cell = table.rows[0].cells[0];
  assert.equal(cell.nodes.length, 1, "a cell keeps its block structure");
  assert.equal(textOf(cell.nodes[0]), "one\ntwocell");
  assert.deepEqual(runsOf(cell.nodes[0])[0].style, { bold: true });
  assert.deepEqual(result.caret, {
    start: {
      location: { kind: "table-cell", tableIndex: 0, rowIndex: 0, cellIndex: 0, paragraphIndex: 0 },
      offset: 7,
    },
    end: {
      location: { kind: "table-cell", tableIndex: 0, rowIndex: 0, cellIndex: 0, paragraphIndex: 0 },
      offset: 7,
    },
  });
  assert.equal(textOf(model.nodes[0].rows[0].cells[0].nodes[0]), "cell", "source cell untouched");
});

test("paste into an empty paragraph adopts the pasted heading and alignment", () => {
  const model = starterModel([{ type: "paragraph", children: [{ type: "text", text: "" }] }]);
  const result = core.insertDocxClipboardParagraphs(model, caret(0, 0), [
    {
      runs: [{ text: "Title", style: { bold: true, fontSizePt: 18 } }],
      style: { headingLevel: 1, align: "center" },
    },
    { runs: [{ text: "body" }] },
  ]);
  assert.ok(result);
  assert.equal(result.model.nodes[0].style.headingLevel, 1);
  assert.equal(result.model.nodes[0].style.align, "center");
  assert.equal(textOf(result.model.nodes[0]), "Title");
  assert.equal(textOf(result.model.nodes[1]), "body");
  assert.equal(runsOf(result.model.nodes[1]).length, 1);
  assert.equal(runsOf(result.model.nodes[1])[0].text, "body");
  assert.deepEqual(result.caret, caret(1, 4));
});

test("text/html rendering can drop unsafe model links through the URL policy", () => {
  const html = core.renderDocxClipboardHtml(
    [
      {
        runs: [
          { text: "safe", link: "https://example.com/a" },
          { text: "unsafe", link: "javascript:alert(1)" },
        ],
      },
    ],
    { sanitizeLink },
  );
  assert.match(html, /<a href="https:\/\/example.com\/a">safe<\/a>/);
  assert.match(html, /unsafe/);
  assert.doesNotMatch(html, /javascript:/);
});

test("RTF rendering carries styles, links, heading sizes and list markers", () => {
  const rtf = core.renderDocxClipboardRtf(
    [
      { runs: [{ text: "Title" }], style: { headingLevel: 1, align: "center" } },
      {
        runs: [
          { text: "bold", style: { bold: true, color: "#ff0000" } },
          { text: " and " },
          { text: "link", link: "https://example.com/a", style: { underline: true } },
          { text: "unicode • ünïcode 中文" },
        ],
      },
      { runs: [{ text: "one" }], list: { type: "unordered", level: 0 } },
      { runs: [{ text: "two" }], list: { type: "unordered", level: 0 } },
      { runs: [{ text: "first" }], list: { type: "ordered", level: 0 } },
      { runs: [{ text: "second" }], list: { type: "ordered", level: 0 } },
    ],
    { sanitizeLink },
  );

  assert.match(rtf, /^\{\\rtf1\\ansi/);
  assert.match(rtf, /\\colortbl;\\red255\\green0\\blue0;/);
  assert.match(rtf, /\\pard\\qc \{\\b\\fs36 Title\}\\par/);
  assert.match(rtf, /\{\\b\\cf1 bold\}/);
  assert.match(rtf, /HYPERLINK "https:\/\/example.com\/a"/);
  assert.match(rtf, /\\u20013\?\\u25991\?/, "non-ASCII text is emitted as \\uN escapes");
  assert.match(rtf, /\\u8226\?\\tab one\\par/);
  assert.match(rtf, /1\.\\tab first\\par/);
  assert.match(rtf, /2\.\\tab second\\par/);
});

test("RTF round-trips runs, links and lists", () => {
  const blocks = [
    { runs: [{ text: "bold", style: { bold: true } }] },
    {
      runs: [
        { text: "link", link: "https://example.com/a", style: { underline: true } },
        { text: " tail" },
      ],
    },
    { runs: [{ text: "bullet" }], list: { type: "unordered", level: 0 } },
    { runs: [{ text: "numbered" }], list: { type: "ordered", level: 0 } },
  ];
  const parsed = core.parseDocxClipboardRtf(
    core.renderDocxClipboardRtf(blocks, { sanitizeLink }),
    { sanitizeLink },
  );
  assert.equal(parsed.length, 4);
  assert.deepEqual(parsed[0].runs, [{ text: "bold", style: { bold: true } }]);
  assert.equal(parsed[1].runs[0].link, "https://example.com/a");
  assert.equal(parsed[1].runs[0].style.underline, true);
  assert.deepEqual(parsed[2], {
    runs: [{ text: "bullet" }],
    list: { type: "unordered", level: 0 },
  });
  assert.deepEqual(parsed[3], {
    runs: [{ text: "numbered" }],
    list: { type: "ordered", level: 0 },
  });
});

test("RTF parsing reads Word markup, sanitizes links and caps the input", () => {
  const wordRtf = `{\\rtf1\\ansi\\ansicpg1252\\deff0
{\\fonttbl{\\f0\\fnil\\fcharset0 Calibri;}}
{\\colortbl;\\red0\\green0\\blue255;}
{\\stylesheet{\\s1 Heading 1;}}
{\\info{\\title ignored}}
{\\*\\generator Riched20 10.0.19041}\\viewkind4\\uc1
\\pard\\s1\\qc\\b\\fs36 Heading\\b0\\par
\\pard\\ql\\f0\\fs24 Normal {\\b bold} and {\\field{\\*\\fldinst{HYPERLINK "https://example.com/z"}}{\\fldrslt{\\ul\\cf1 link}}} text\\par
\\pard{\\pntext\\f0 \\'B7\\tab}{\\*\\pn\\pnlvlblt\\pnf0\\pnindent0{\\pntxtb\\'B7}}\\fi-360\\li720 bullet item\\par
\\pard{\\pntext\\f0 1.\\tab}\\fi-360\\li720 numbered item\\par
\\pard {\\field{\\*\\fldinst{HYPERLINK "javascript:alert(1)"}}{\\fldrslt bad}}}\\par
}`;

  const parsed = core.parseDocxClipboardRtf(wordRtf, { sanitizeLink });
  assert.equal(parsed.length, 5);
  assert.equal(parsed[0].style.align, "center");
  assert.deepEqual(parsed[0].runs[0].style, { bold: true, fontSizePt: 18 });
  assert.equal(parsed[0].runs[0].text, "Heading");

  const normal = parsed[1];
  assert.equal(normal.style.align, "left");
  assert.equal(normal.runs.find((run) => run.text === "bold").style.bold, true);
  const link = normal.runs.find((run) => run.text === "link");
  assert.equal(link.link, "https://example.com/z");
  assert.deepEqual(link.style, {
    bold: false,
    color: "#0000ff",
    fontSizePt: 12,
    underline: true,
  });

  assert.equal(parsed[2].runs[0].text, "bullet item", "Word's bullet marker is stripped");
  assert.deepEqual(parsed[2].list, { type: "unordered", level: 0 });
  assert.equal(parsed[3].runs[0].text, "numbered item");
  assert.deepEqual(parsed[3].list, { type: "ordered", level: 0 });

  const unsafe = parsed[4];
  assert.equal(unsafe.runs[0].text.trim(), "bad");
  assert.equal(unsafe.runs[0].link, undefined, "an unsafe RTF hyperlink is dropped");

  assert.deepEqual(
    core.parseDocxClipboardRtf("{\\rtf1\\ansi plain text\\par}", { maxRtfChars: 5 }),
    [],
    "content past the cap is not parsed",
  );
});
