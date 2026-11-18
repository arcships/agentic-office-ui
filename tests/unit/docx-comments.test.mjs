// DOCX comment authoring (#96): model CRUD, threading, anchors and the
// `word/comments.xml` / `word/commentsExtended.xml` round trip.
//
// Each block names the issue bullet it covers so the mapping stays checkable.

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromDocxCore = createRequire(
  new URL("../../packages/docx-core/package.json", import.meta.url),
);
const { strFromU8, unzipSync } = await import(
  pathToFileURL(requireFromDocxCore.resolve("fflate")).href
);
const core = await import(
  pathToFileURL(requireFromDocxCore.resolve("@arcships/docx-core")).href
);

const FIXTURE = new URL("../../test-data/corpus/docx-comments.docx", import.meta.url);

// ── Fixtures ───────────────────────────────────────────────────────
function paragraph(text) {
  return { type: "paragraph", children: [{ type: "text", text }] };
}

function modelWithParagraphs(...texts) {
  return {
    ...core.createBlankDocumentModel(),
    nodes: texts.map((text) => paragraph(text)),
  };
}

function textRange(nodeIndex, start, end, endNodeIndex = nodeIndex) {
  return {
    start: { location: { kind: "paragraph", nodeIndex }, offset: start },
    end: { location: { kind: "paragraph", nodeIndex: endNodeIndex }, offset: end },
  };
}

function paragraphTextOf(node) {
  return node.children
    .filter((child) => child.type === "text")
    .map((child) => child.text)
    .join("");
}

async function loadFixture() {
  const bytes = readFileSync(FIXTURE);
  const buffer = bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  );
  return core.importDocxBuffer(buffer, { useWorker: false });
}

function unzip(bytes) {
  return unzipSync(new Uint8Array(bytes));
}

function partText(archive, name) {
  return archive[name] ? strFromU8(archive[name]) : "";
}

// ── Bullet 1: comment model with ids, author, date, parentId, resolved ──
test("adding a comment stores a definition and its paragraph anchor offsets", () => {
  const model = modelWithParagraphs("Alpha beta gamma delta");
  const added = core.addDocumentComment(model, {
    range: textRange(0, 6, 16),
    text: "  Please confirm.  ",
    author: "Alex Reviewer",
    date: "2026-11-18T00:00:00Z",
  });

  assert.ok(added, "the comment is created");
  assert.equal(added.commentId, 0, "ids start at 0 like Word's");
  assert.deepEqual(added.model.metadata.comments, [
    {
      id: 0,
      text: "Please confirm.",
      author: "Alex Reviewer",
      initials: "AR",
      date: "2026-11-18T00:00:00Z",
      parentId: undefined,
      resolved: undefined,
    },
  ]);
  assert.deepEqual(added.model.nodes[0].commentRanges, [
    { commentId: 0, start: 6, end: 16, anchorText: "beta gamma" },
  ]);
  // The input model stays untouched (immutable revisions back undo/history).
  assert.equal(model.metadata.comments, undefined);
  assert.equal(model.nodes[0].commentRanges, undefined);
});

test("a selection spanning paragraphs anchors every paragraph it touches", () => {
  const model = modelWithParagraphs("First paragraph", "Middle paragraph", "Last paragraph");
  const added = core.addDocumentComment(model, {
    range: textRange(0, 6, 7, 2),
    text: "Cross-paragraph note",
    author: "Morgan Editor",
  });

  assert.ok(added);
  assert.deepEqual(
    added.model.nodes.map((node) => node.commentRanges),
    [
      [{ commentId: 0, start: 6, end: 15, anchorText: "paragraph" }],
      [{ commentId: 0, start: 0, end: 16, anchorText: "Middle paragraph" }],
      [{ commentId: 0, start: 0, end: 7, anchorText: "Last pa" }],
    ],
  );
});

test("comments anchor inside table cells without touching other cells", () => {
  const model = {
    ...core.createBlankDocumentModel(),
    nodes: [
      {
        type: "table",
        rows: [
          {
            type: "table-row",
            cells: [
              { type: "table-cell", nodes: [paragraph("Left cell text")] },
              { type: "table-cell", nodes: [paragraph("Right cell text")] },
            ],
          },
        ],
      },
    ],
  };
  const added = core.addDocumentComment(model, {
    range: {
      start: {
        location: {
          kind: "table-cell",
          tableIndex: 0,
          rowIndex: 0,
          cellIndex: 1,
          paragraphIndex: 0,
        },
        offset: 0,
      },
      end: {
        location: {
          kind: "table-cell",
          tableIndex: 0,
          rowIndex: 0,
          cellIndex: 1,
          paragraphIndex: 0,
        },
        offset: 5,
      },
    },
    text: "Cell note",
    author: "Priya Reviewer",
  });

  assert.ok(added);
  const cells = added.model.nodes[0].rows[0].cells;
  assert.equal(cells[0].nodes[0].commentRanges, undefined);
  assert.deepEqual(cells[1].nodes[0].commentRanges, [
    { commentId: 0, start: 0, end: 5, anchorText: "Right" },
  ]);
  // The original table is not mutated by the copy-on-write edit.
  assert.equal(model.nodes[0].rows[0].cells[1].nodes[0].commentRanges, undefined);
});

// ── Bullet 1: threads (reply / resolve / delete) ───────────────────
test("replies join the thread and inherit the parent's anchor in the read path", () => {
  const model = modelWithParagraphs("Threaded comment anchor");
  const root = core.addDocumentComment(model, {
    range: textRange(0, 0, 8),
    text: "Root",
    author: "Priya Reviewer",
  });
  const first = core.replyToDocumentComment(root.model, root.commentId, "First reply", {
    author: "Alex Reviewer",
  });
  assert.ok(first);
  assert.equal(first.commentId, 1);
  assert.deepEqual(first.model.metadata.comments[1], {
    id: 1,
    text: "First reply",
    author: "Alex Reviewer",
    initials: "AR",
    date: undefined,
    parentId: 0,
    resolved: undefined,
  });

  const second = core.replyToDocumentComment(first.model, root.commentId, "Second reply", {
    author: "Morgan Editor",
  });
  assert.equal(second.commentId, 2);
  assert.deepEqual(core.commentThreadMemberIds(second.model, 1), [0, 1, 2]);
  assert.equal(core.commentThreadRootId(second.model, 2), 0);

  const comments = core.collectCommentsFromModel(second.model);
  assert.deepEqual(
    comments.map((comment) => ({
      id: comment.commentId,
      parentId: comment.parentId,
      anchorText: comment.anchorText,
    })),
    [
      { id: 0, parentId: undefined, anchorText: "Threaded" },
      { id: 1, parentId: 0, anchorText: "Threaded" },
      { id: 2, parentId: 0, anchorText: "Threaded" },
    ],
  );
});

test("resolving marks the thread root and reopening clears the whole thread", () => {
  const model = modelWithParagraphs("Resolve this thread");
  const root = core.addDocumentComment(model, {
    range: textRange(0, 0, 7),
    text: "Root",
    author: "Priya Reviewer",
  });
  const reply = core.replyToDocumentComment(root.model, 0, "Reply");
  const resolved = core.resolveDocumentComment(reply.model, reply.commentId, true);

  assert.ok(resolved);
  assert.equal(resolved.metadata.comments[0].resolved, true, "the root carries w15:done");
  assert.equal(resolved.metadata.comments[1].resolved, undefined);

  const comments = core.collectCommentsFromModel(resolved);
  const byId = new Map(comments.map((comment) => [comment.commentId, comment]));
  assert.equal(core.commentThreadIsResolved(byId.get(0), byId), true);
  assert.equal(core.commentThreadIsResolved(byId.get(1), byId), true);

  const reopened = core.resolveDocumentComment(resolved, 1, false);
  assert.ok(reopened);
  assert.equal(reopened.metadata.comments[0].resolved, undefined);
  const reopenedComments = core.collectCommentsFromModel(reopened);
  const reopenedById = new Map(
    reopenedComments.map((comment) => [comment.commentId, comment]),
  );
  assert.equal(core.commentThreadIsResolved(reopenedById.get(0), reopenedById), false);
  // Reopening an unresolved thread is a no-op, not a new revision.
  assert.equal(core.resolveDocumentComment(reopened, 0, false), reopened);
});

test("deleting a root removes its replies and every anchor range", () => {
  const model = modelWithParagraphs("First anchor here", "Second anchor here");
  const root = core.addDocumentComment(model, {
    range: textRange(0, 0, 5),
    text: "Root",
    author: "Priya Reviewer",
  });
  const other = core.addDocumentComment(root.model, {
    range: textRange(1, 0, 6),
    text: "Unrelated",
    author: "Priya Reviewer",
  });
  const reply = core.replyToDocumentComment(other.model, root.commentId, "Reply");

  const deleted = core.deleteDocumentComment(reply.model, root.commentId);
  assert.ok(deleted);
  assert.deepEqual(
    deleted.metadata.comments.map((comment) => comment.id),
    [other.commentId],
  );
  assert.equal(deleted.nodes[0].commentRanges, undefined, "the root anchor is gone");
  assert.deepEqual(deleted.nodes[1].commentRanges, [
    { commentId: other.commentId, start: 0, end: 6, anchorText: "Second" },
  ]);
  assert.deepEqual(
    core.collectCommentsFromModel(deleted).map((comment) => comment.commentId),
    [other.commentId],
  );
  assert.equal(core.deleteDocumentComment(deleted, 99), undefined);
});

test("deleting one reply keeps the root and the other replies", () => {
  const model = modelWithParagraphs("Threaded comment anchor");
  const root = core.addDocumentComment(model, {
    range: textRange(0, 0, 8),
    text: "Root",
  });
  const first = core.replyToDocumentComment(root.model, 0, "First reply");
  const second = core.replyToDocumentComment(first.model, 0, "Second reply");
  const nested = core.replyToDocumentComment(second.model, first.commentId, "Nested reply");

  const deleted = core.deleteDocumentComment(nested.model, first.commentId);
  assert.ok(deleted);
  assert.deepEqual(
    deleted.metadata.comments.map((comment) => comment.id),
    [0, 2],
  );
});

test("empty comment text and unknown targets are rejected without a revision", () => {
  const model = modelWithParagraphs("Only paragraph");
  assert.equal(
    core.addDocumentComment(model, { range: textRange(0, 0, 4), text: "   " }),
    undefined,
  );
  assert.equal(core.replyToDocumentComment(model, 42, "Reply"), undefined);
  assert.equal(core.replyToDocumentComment(model, 0, "  "), undefined);
  assert.equal(core.resolveDocumentComment(model, 42), undefined);
});

// ── Bullet 2: serialization to comments.xml + commentsExtended.xml ──
test("export writes comment parts, threads, resolved state and package relationships", async () => {
  const model = modelWithParagraphs("Export this comment thread");
  const root = core.addDocumentComment(model, {
    range: textRange(0, 7, 15),
    text: "Root note\nwith a second line",
    author: "Priya Reviewer",
    date: "2026-11-18T00:00:00Z",
  });
  const reply = core.replyToDocumentComment(root.model, root.commentId, "Reply note", {
    author: "Alex Reviewer",
    date: "2026-11-18T01:00:00Z",
  });
  const resolved = core.resolveDocumentComment(reply.model, root.commentId, true);

  const bytes = await core.serializeDocx(resolved);
  const archive = unzip(bytes);
  const commentsXml = partText(archive, "word/comments.xml");
  const extendedXml = partText(archive, "word/commentsExtended.xml");
  const documentXml = partText(archive, "word/document.xml");

  assert.match(commentsXml, /<w:comment w:id="0" w:author="Priya Reviewer" w:initials="PR"/);
  assert.match(commentsXml, /Root note/);
  assert.match(commentsXml, /with a second line/);
  assert.match(commentsXml, /<w:comment w:id="1" w:author="Alex Reviewer"/);
  assert.match(extendedXml, /w15:paraIdParent=/, "replies are threaded");
  assert.match(extendedXml, /w15:done="1"/, "the resolved thread is marked done");

  // Anchors: markers outside runs plus a reference run in the paragraph.
  assert.match(documentXml, /<w:commentRangeStart w:id="0"\/>/);
  assert.match(documentXml, /<w:commentRangeEnd w:id="0"\/>/);
  assert.match(documentXml, /<w:commentReference w:id="0"\/>/);
  assert.doesNotMatch(documentXml, /<w:r>[^<]*<w:commentRangeStart/);

  const contentTypes = partText(archive, "[Content_Types].xml");
  assert.match(contentTypes, /PartName="\/word\/comments.xml"/);
  assert.match(contentTypes, /PartName="\/word\/commentsExtended.xml"/);
  const rels = partText(archive, "word/_rels/document.xml.rels");
  assert.match(rels, /Target="comments.xml"/);
  assert.match(rels, /Target="commentsExtended.xml"/);

  // Export -> reopen keeps the thread, the resolved flag and the anchor text.
  const reopened = await core.importDocxBuffer(bytes, { useWorker: false });
  const definitions = reopened.model.metadata.comments;
  assert.deepEqual(
    definitions.map((comment) => ({
      id: comment.id,
      parentId: comment.parentId,
      resolved: comment.resolved ?? false,
      text: comment.text,
    })),
    [
      { id: 0, parentId: undefined, resolved: true, text: "Root note\nwith a second line" },
      { id: 1, parentId: 0, resolved: false, text: "Reply note" },
    ],
  );
  assert.deepEqual(reopened.model.nodes[0].commentRanges, [
    { commentId: 0, start: 7, end: 15, anchorText: "this com" },
  ]);
});

test("exporting a document with comments keeps its other parts and strips deleted anchors", async () => {
  const { model } = await loadFixture();
  const rootId = model.metadata.comments[0].id;

  const deleted = core.deleteDocumentComment(model, rootId);
  assert.ok(deleted);
  const bytes = await core.serializeDocx(deleted, { parts: new Map(), binaryAssets: new Map() });
  const archive = unzip(bytes);
  const documentXml = partText(archive, "word/document.xml");
  const commentsXml = partText(archive, "word/comments.xml");

  assert.doesNotMatch(commentsXml, new RegExp(`w:comment w:id="${rootId}"`));
  assert.doesNotMatch(documentXml, new RegExp(`w:commentRangeStart w:id="${rootId}"`));
  assert.doesNotMatch(documentXml, new RegExp(`w:commentReference w:id="${rootId}"`));
  // The surviving resolved thread is still there.
  assert.match(commentsXml, /Resolved root: wording accepted\./);
  assert.match(partText(archive, "word/commentsExtended.xml"), /w15:done="1"/);
});

test("deleting every comment removes the comment parts, relationships and anchors", async () => {
  const { model } = await loadFixture();
  let current = model;
  for (const comment of [...model.metadata.comments].sort((a, b) => b.id - a.id)) {
    current = core.deleteDocumentComment(current, comment.id) ?? current;
  }
  assert.deepEqual(current.metadata.comments, []);

  const bytes = await core.serializeDocx(current);
  const archive = unzip(bytes);
  assert.equal(archive["word/comments.xml"], undefined);
  assert.equal(archive["word/commentsExtended.xml"], undefined);
  assert.doesNotMatch(partText(archive, "[Content_Types].xml"), /comments\.xml/);
  assert.doesNotMatch(partText(archive, "word/_rels/document.xml.rels"), /comments\.xml/);
  assert.doesNotMatch(partText(archive, "word/document.xml"), /commentRangeStart|commentReference/);
});

test("serializeDocModel patches the package with the same comment parts", async () => {
  const { model } = await loadFixture();
  const added = core.addDocumentComment(model, {
    range: textRange(2, 4, 11),
    text: "Package path note",
    author: "Priya Reviewer",
    date: "2026-11-18T02:00:00Z",
  });
  const pkg = await core.serializeDocModel(added.model);

  assert.ok(pkg.parts.has("word/comments.xml"));
  assert.ok(pkg.parts.has("word/commentsExtended.xml"));
  assert.match(
    pkg.parts.get("word/comments.xml").content,
    /Package path note/,
  );
  assert.match(
    pkg.parts.get("[Content_Types].xml").content,
    /PartName="\/word\/commentsExtended.xml"/,
  );
});

test("a document without comments exports no comment parts", async () => {
  const model = modelWithParagraphs("Plain document");
  const bytes = await core.serializeDocx(model);
  const archive = unzip(bytes);
  assert.equal(archive["word/comments.xml"], undefined);
  assert.equal(archive["word/commentsExtended.xml"], undefined);
});

// ── Bullet 3: fixtures carry a reply thread and a resolved thread ──
test("the comment fixture carries a plain thread, a reply thread and a resolved thread", async () => {
  const { model } = await loadFixture();
  const definitions = model.metadata.comments;
  assert.equal(definitions.length, 5);

  const root = definitions.find((comment) => comment.id === 0);
  const replies = definitions.filter((comment) => comment.parentId === 0);
  assert.equal(replies.length, 2, "the plain thread has two replies");
  assert.equal(root.text, "Root: confirm the review sentence.");

  const resolvedRoot = definitions.find((comment) => comment.id === 3);
  assert.equal(resolvedRoot.resolved, true, "the second thread is resolved");
  assert.ok(
    definitions.some((comment) => comment.parentId === resolvedRoot.id),
    "the resolved thread has a reply",
  );

  const comments = core.collectCommentsFromModel(model);
  const byId = new Map(comments.map((comment) => [comment.commentId, comment]));
  assert.equal(core.commentThreadIsResolved(byId.get(3), byId), true);
  assert.equal(core.commentThreadIsResolved(byId.get(4), byId), true);
  assert.equal(core.commentThreadIsResolved(byId.get(0), byId), false);
  assert.equal(byId.get(0).anchorText, "review thread");
  assert.equal(byId.get(3).anchorText, "resolved thread");
});

test("fixture comments survive an edit, export and reopen", async () => {
  const { model, package: basePackage } = await loadFixture();
  const edited = core.updateParagraphText(model, 0, "Edited: The review thread anchors on this sentence.");
  assert.equal(edited.nodes[0].sourceXml, undefined, "the edit clears sourceXml");

  // `exportDocxBuffer` is the controller's export entry point; without a Worker
  // it runs the same serializer inline.
  const bytes = await core.exportDocxBuffer(edited, basePackage);
  const reopened = await core.importDocxBuffer(bytes, { useWorker: false });
  const comments = core.collectCommentsFromModel(reopened.model);
  const root = comments.find((comment) => comment.commentId === 0);
  assert.equal(root.anchorText, "review thread");
  assert.equal(
    comments.filter((comment) => comment.parentId === 0).length,
    2,
    "the reply thread survives the edit and reload",
  );
  const resolvedReply = comments.find((comment) => comment.commentId === 4);
  assert.equal(resolvedReply.resolved, undefined);
  assert.equal(core.commentThreadIsResolved(resolvedReply, new Map(comments.map((comment) => [comment.commentId, comment]))), true);
});
