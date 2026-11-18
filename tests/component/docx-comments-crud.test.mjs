// DOCX comment authoring UI and controller (#96).
//
// Covers the controller API (`addComment` / `replyToComment` / `resolveComment`
// / `deleteComment`), the selection toolbar action, the comment card actions
// and the inline editor, plus export/reopen and undo.
//
// Each test names the issue bullet it covers so the mapping stays checkable.

import assert from "node:assert/strict";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

const { DocxPageSurface, DocxToolbar, useDocxEditor } = await importFromDemo(
  "@arcships/vue-docx",
);
const core = await importFromDemo("@arcships/docx-core");

// ── Fixtures ───────────────────────────────────────────────────────
function paragraph(text) {
  return { type: "paragraph", children: [{ type: "text", text }] };
}

function starterModel(...texts) {
  return {
    ...core.createBlankDocumentModel(),
    nodes: texts.map((text) => paragraph(text)),
  };
}

function textRange(nodeIndex, start, end) {
  return {
    start: { location: { kind: "paragraph", nodeIndex }, offset: start },
    end: { location: { kind: "paragraph", nodeIndex }, offset: end },
  };
}

function click(node) {
  node.props.onClick({ stopPropagation() {}, preventDefault() {} });
}

function nodesByTestId(root, testId) {
  return walk(root).filter((node) => node.props?.["data-testid"] === testId);
}

function cardIds(root) {
  return walk(root)
    .filter((node) => node.props?.["data-docx-gutter-card-id"] !== undefined)
    .map((node) => node.props["data-docx-gutter-card-id"]);
}

async function mountController(model, options = {}) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({
        starterModel: model,
        commentAuthor: "Review Bot",
        commentClock: () => new Date("2026-11-18T00:00:00Z"),
        ...options,
      });
      return () => vue.h("div");
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, get controller() { return controller; } };
}

const PAGE_LAYOUT = {
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginsPx: { top: 72, right: 72, bottom: 72, left: 72 },
};

async function mountSurface(model, options = {}) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({
        starterModel: model,
        initialShowComments: true,
        commentAuthor: "Review Bot",
        commentClock: () => new Date("2026-11-18T00:00:00Z"),
        ...options,
      });
      return () =>
        vue.h(DocxPageSurface, {
          pageIndex: 0,
          pageLayout: PAGE_LAYOUT,
          pageNodeSegments: [{ nodeIndex: 0 }, { nodeIndex: 1 }],
          model: controller.model,
          pageWidthPx: PAGE_LAYOUT.pageWidthPx,
          pageContentWidthPx: 672,
          pageNumber: 1,
          totalPages: 1,
          commentsEnabled: controller.showComments,
          trackedChangesEnabled: false,
          editable: true,
          controller,
        });
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, get controller() { return controller; } };
}

// ── Bullet 2: controller API with history ─────────────────────────
test("addComment creates a definition and its anchors in one undoable revision", async () => {
  const { app, controller } = await mountController(
    starterModel("Alpha beta gamma delta", "Second paragraph"),
  );
  try {
    const range = textRange(0, 6, 16);
    controller.setActiveTextRange(range);
    const commentId = controller.addComment(range, "  Please confirm.  ");

    assert.equal(commentId, 0);
    assert.equal(controller.showComments, true, "commenting reveals the gutter");
    assert.equal(controller.status, "Comment added");
    assert.deepEqual(
      controller.comments.map((comment) => ({
        commentId: comment.commentId,
        text: comment.text,
        author: comment.author,
        initials: comment.initials,
        anchorText: comment.anchorText,
      })),
      [
        {
          commentId: 0,
          text: "Please confirm.",
          author: "Review Bot",
          initials: "RB",
          anchorText: "beta gamma",
        },
      ],
    );
    assert.deepEqual(controller.model.metadata.comments, [
      {
        id: 0,
        text: "Please confirm.",
        author: "Review Bot",
        initials: "RB",
        date: "2026-11-18T00:00:00.000Z",
        parentId: undefined,
        resolved: undefined,
      },
    ]);
    assert.deepEqual(controller.model.nodes[0].commentRanges, [
      { commentId: 0, start: 6, end: 16, anchorText: "beta gamma" },
    ]);
    assert.equal(controller.isDirty, true);

    // Undo restores the pre-comment model (no definition, no anchor).
    controller.undo();
    await vue.nextTick();
    assert.deepEqual(controller.comments, []);
    assert.equal(controller.model.metadata.comments, undefined);
    assert.equal(controller.model.nodes[0].commentRanges, undefined);
    assert.equal(controller.isDirty, false);

    controller.redo();
    await vue.nextTick();
    assert.equal(controller.comments.length, 1);
    assert.deepEqual(controller.model.nodes[0].commentRanges, [
      { commentId: 0, start: 6, end: 16, anchorText: "beta gamma" },
    ]);
  } finally {
    app.unmount();
  }
});

test("reply twice, resolve the thread, then delete and undo restores it", async () => {
  const { app, controller } = await mountController(
    starterModel("Alpha beta gamma delta"),
  );
  try {
    const rootId = controller.addComment(textRange(0, 6, 16), "Root comment");
    const firstReply = controller.replyToComment(rootId, "First reply");
    const secondReply = controller.replyToComment(rootId, "Second reply");

    assert.equal(firstReply, 1);
    assert.equal(secondReply, 2);
    assert.equal(controller.comments.length, 3);
    assert.deepEqual(
      controller.comments.map((comment) => ({
        id: comment.commentId,
        parentId: comment.parentId,
      })),
      [
        { id: 0, parentId: undefined },
        { id: 1, parentId: 0 },
        { id: 2, parentId: 0 },
      ],
    );
    assert.equal(controller.status, "Reply added");

    // Resolve is a thread-level action, so a reply id resolves the whole thread.
    assert.equal(controller.resolveComment(firstReply, true), true);
    assert.equal(controller.status, "Comment thread resolved");
    const commentsById = new Map(
      controller.comments.map((comment) => [comment.commentId, comment]),
    );
    assert.equal(
      core.commentThreadIsResolved(commentsById.get(0), commentsById),
      true,
    );
    assert.equal(controller.model.metadata.comments[0].resolved, true);

    // Deleting the root removes the replies and their anchors.
    assert.equal(controller.deleteComment(rootId), true);
    assert.equal(controller.status, "Comment deleted");
    assert.deepEqual(controller.comments, []);
    assert.equal(controller.model.nodes[0].commentRanges, undefined);

    controller.undo();
    await vue.nextTick();
    assert.equal(controller.comments.length, 3, "undo restores the whole thread");
    assert.deepEqual(controller.model.nodes[0].commentRanges, [
      { commentId: 0, start: 6, end: 16, anchorText: "beta gamma" },
    ]);
    assert.equal(controller.model.metadata.comments[0].resolved, true);

    controller.undo();
    await vue.nextTick();
    assert.equal(controller.model.metadata.comments[0].resolved, undefined);
    assert.equal(controller.resolveComment(999, true), false);
    assert.equal(controller.deleteComment(999), false);
  } finally {
    app.unmount();
  }
});

// ── Bullet 2: export, reopen ───────────────────────────────────────
test("a created thread survives export and reopen with its resolved state", async () => {
  const { app, controller } = await mountController(
    starterModel("Export alpha beta gamma delta"),
  );
  try {
    const rootId = controller.addComment(textRange(0, 7, 17), "Root comment");
    controller.replyToComment(rootId, "First reply");
    controller.replyToComment(rootId, "Second reply");
    controller.resolveComment(rootId, true);

    const bytes = await core.serializeDocx(controller.model);
    const reopened = await core.importDocxBuffer(bytes, { useWorker: false });
    const definitions = reopened.model.metadata.comments;

    assert.deepEqual(
      definitions.map((comment) => ({
        id: comment.id,
        parentId: comment.parentId,
        resolved: comment.resolved ?? false,
        text: comment.text,
        author: comment.author,
      })),
      [
        {
          id: 0,
          parentId: undefined,
          resolved: true,
          text: "Root comment",
          author: "Review Bot",
        },
        { id: 1, parentId: 0, resolved: false, text: "First reply", author: "Review Bot" },
        { id: 2, parentId: 0, resolved: false, text: "Second reply", author: "Review Bot" },
      ],
    );
    assert.deepEqual(reopened.model.nodes[0].commentRanges, [
      { commentId: 0, start: 7, end: 17, anchorText: "alpha beta" },
    ]);

    // Reopen the exported document in a fresh controller: the gutter shows the
    // thread's cards and the resolved flag is inherited by the replies.
    const { app: secondApp, controller: second, root } = await mountSurface(
      reopened.model,
    );
    try {
      assert.equal(second.comments.length, 3);
      const commentsById = new Map(
        second.comments.map((comment) => [comment.commentId, comment]),
      );
      assert.equal(
        core.commentThreadIsResolved(commentsById.get(2), commentsById),
        true,
      );
      assert.equal(cardIds(root).length, 0, "resolved threads stay hidden by default");
    } finally {
      secondApp.unmount();
    }
  } finally {
    app.unmount();
  }
});

// ── Bullet 3: inline editor flow ───────────────────────────────────
test("the comment composer creates a comment and a reply, and rejects empty text", async () => {
  const { app, controller } = await mountController(starterModel("Composer target"));
  try {
    assert.equal(controller.beginComment(), false, "no selection, no composer");

    const range = textRange(0, 0, 8);
    controller.setActiveTextRange(range);
    assert.equal(controller.beginComment(range), true);
    assert.deepEqual(controller.commentComposer, {
      mode: "new",
      range,
      anchor: range.start.location,
    });
    assert.equal(controller.submitCommentComposer("   "), false);
    assert.ok(controller.commentComposer, "empty text keeps the editor open");

    assert.equal(controller.submitCommentComposer("  From the composer  "), true);
    assert.equal(controller.commentComposer, undefined);
    assert.equal(controller.comments.length, 1);
    assert.equal(controller.comments[0].text, "From the composer");

    assert.equal(controller.beginCommentReply(controller.comments[0].commentId), true);
    assert.deepEqual(controller.commentComposer, { mode: "reply", commentId: 0 });
    assert.equal(controller.submitCommentComposer("Reply from the composer"), true);
    assert.equal(controller.commentComposer, undefined);
    assert.equal(controller.comments.length, 2);
    assert.equal(controller.comments[1].parentId, 0);

    controller.beginComment(range);
    controller.cancelCommentComposer();
    assert.equal(controller.commentComposer, undefined);
    assert.equal(controller.comments.length, 2, "cancelling changes nothing");
  } finally {
    app.unmount();
  }
});

// ── Bullet 3: selection toolbar action ────────────────────────────
test("the toolbar exposes a new-comment action that follows the selection", async () => {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({
        starterModel: starterModel("Toolbar selection comment"),
        commentAuthor: "Review Bot",
      });
      return () =>
        vue.h(DocxToolbar, {
          controller,
          zoom: 100,
          currentPage: 1,
          totalPages: 1,
        });
    },
  });
  const { app, root, warnings } = await mount(Harness);
  try {
    const button = findByTestId(root, "editor-add-comment");
    assert.ok(button, "the toolbar renders the comment action");
    assert.equal(button.props.disabled, true, "no selection, no comment action");
    assert.equal(button.props["aria-label"], "New comment");

    controller.setActiveTextRange(textRange(0, 0, 7));
    await vue.nextTick();
    assert.equal(findByTestId(root, "editor-add-comment").props.disabled, false);

    click(findByTestId(root, "editor-add-comment"));
    await vue.nextTick();
    assert.equal(controller.commentComposer.mode, "new");
    assert.equal(controller.showComments, true);
    assert.deepEqual(warnings, []);
  } finally {
    app.unmount();
  }
});

// ── Bullet 3: comment card actions and inline reply editor ─────────
test("comment cards reply, resolve and delete through the page gutter", async () => {
  const baseModel = {
    ...core.createBlankDocumentModel(),
    nodes: [paragraph("Card actions target text"), paragraph("Second paragraph")],
  };
  const created = core.addDocumentComment(baseModel, {
    range: textRange(0, 7, 13),
    text: "Root comment",
    author: "Priya Reviewer",
  });
  const seeded = core.replyToDocumentComment(created.model, created.commentId, "Existing reply", {
    author: "Alex Reviewer",
  }).model;

  const { app, root, controller, warnings } = await mountSurface(seeded);
  try {
    assert.deepEqual(cardIds(root).sort(), [
      `comment:p:0:comment:${created.commentId}`,
      "comment:p:0:comment:1",
    ].sort());
    assert.equal(nodesByTestId(root, "docx-gutter-comment-reply").length, 2);
    assert.equal(
      nodesByTestId(root, "docx-gutter-comment-resolve").length,
      1,
      "only the thread root offers resolve",
    );
    assert.equal(nodesByTestId(root, "docx-gutter-comment-delete").length, 2);

    // Reply: the editor opens inside the card and its text is submitted.
    click(nodesByTestId(root, "docx-gutter-comment-reply")[0]);
    await vue.nextTick();
    assert.equal(controller.commentComposer.mode, "reply");
    assert.equal(controller.commentComposer.commentId, created.commentId);

    const input = findByTestId(root, "docx-gutter-comment-composer-input");
    assert.ok(input, "the inline editor renders inside the card");
    input.props.onInput({
      target: { value: "Reply from the page" },
      stopPropagation() {},
      preventDefault() {},
    });
    await vue.nextTick();
    click(findByTestId(root, "docx-gutter-comment-composer-submit"));
    await vue.nextTick();
    assert.equal(controller.commentComposer, undefined);
    assert.equal(controller.comments.length, 3);
    assert.equal(controller.comments[2].text, "Reply from the page");
    assert.equal(controller.comments[2].parentId, created.commentId);

    // Resolve hides the thread's cards (resolved threads are filtered out).
    click(nodesByTestId(root, "docx-gutter-comment-resolve")[0]);
    await vue.nextTick();
    assert.equal(controller.model.metadata.comments[0].resolved, true);
    assert.deepEqual(cardIds(root), [], "resolved thread cards are hidden");

    // Undo restores the cards, then delete removes the thread.
    controller.undo();
    await vue.nextTick();
    assert.equal(cardIds(root).length, 3);
    click(nodesByTestId(root, "docx-gutter-comment-delete")[0]);
    await vue.nextTick();
    assert.equal(controller.comments.length, 0);
    assert.equal(controller.model.nodes[0].commentRanges, undefined);
    assert.deepEqual(warnings, []);
  } finally {
    app.unmount();
  }
});

test("a new comment renders its composer beside the anchored page", async () => {
  const { app, root, controller } = await mountSurface(
    starterModel("Composer anchor paragraph", "Second paragraph"),
  );
  try {
    const range = textRange(0, 0, 8);
    controller.setActiveTextRange(range);
    controller.beginComment(range);
    await vue.nextTick();

    const composer = walk(root).find(
      (node) =>
        node.props?.["data-testid"] === "docx-gutter-comment-composer" &&
        node.props?.["data-docx-comment-composer-mode"] === "new",
    );
    assert.ok(composer, "the new-comment composer renders in the gutter");

    const input = findByTestId(root, "docx-gutter-comment-composer-input");
    input.props.onInput({
      target: { value: "Anchored comment" },
      stopPropagation() {},
      preventDefault() {},
    });
    await vue.nextTick();
    click(findByTestId(root, "docx-gutter-comment-composer-submit"));
    await vue.nextTick();

    assert.equal(controller.commentComposer, undefined);
    assert.equal(controller.comments.length, 1);
    assert.equal(controller.comments[0].anchorText, "Composer");
    assert.deepEqual(controller.model.nodes[0].commentRanges, [
      { commentId: 0, start: 0, end: 8, anchorText: "Composer" },
    ]);
  } finally {
    app.unmount();
  }
});
