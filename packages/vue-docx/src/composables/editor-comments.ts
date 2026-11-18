// editor-comments.ts — comment authoring commands (create/reply/resolve/delete)
//
// Comments live entirely in the model (`metadata.comments` plus each
// paragraph's `commentRanges`), so every command here is a single transaction:
// one undo entry that restores both the thread and its anchors.
//
// Extracted from useDocxEditor.ts.

import type {
  DocModel,
  DocxCommentComposer,
  DocxTextRange,
} from "@arcships/docx-core"
import {
  addDocumentComment,
  deleteDocumentComment,
  findDocumentComment,
  getParagraphAtLocation,
  normalizeTextRange,
  replyToDocumentComment,
  resolveDocumentComment,
} from "@arcships/docx-core"
import { watch } from "vue"
import type { EditorCore } from "./editor-shared"

type ApplyModelChange = (
  updater: (current: DocModel) => DocModel,
  successStatus?: string,
  options?: { clearSelectedFormField?: boolean; pushHistory?: boolean },
) => boolean

export interface EditorCommentAuthor {
  /** Display name stored on new comment definitions. */
  author: string
  /** Initials shown in Word; the model derives them from the author. */
  initials?: string
  /** Clock used for `w:date`; tests pin it through `UseDocxEditorOptions`. */
  clock: () => Date
}

function timestamp(clock: () => Date): string {
  const value = clock()
  return value instanceof Date && Number.isFinite(value.getTime())
    ? value.toISOString()
    : new Date().toISOString()
}

export function createEditorComments(
  ctx: EditorCore,
  applyChange: ApplyModelChange,
  authorOptions: EditorCommentAuthor,
) {
  const commentComposer = ctx.commentComposer

  function closeComposer(): void {
    commentComposer.value = undefined
  }

  function authorMetadata(): { author: string; initials?: string; date: string } {
    return {
      author: authorOptions.author,
      initials: authorOptions.initials,
      date: timestamp(authorOptions.clock),
    }
  }

  /**
   * Opens the inline editor for a new comment on `range` (defaults to the live
   * selection). Returns false when there is no selection to anchor to.
   */
  function beginComment(range?: DocxTextRange): boolean {
    const source = range ?? ctx.activeTextRangeSnapshot.value
    if (!source) {
      return false
    }
    const normalized = normalizeTextRange(source)
    ctx.showComments.value = true
    commentComposer.value = {
      mode: "new",
      range: normalized,
      anchor: normalized.start.location,
    }
    return true
  }

  /** Opens the inline editor for a reply to `commentId`. */
  function beginCommentReply(commentId: number): boolean {
    if (!findDocumentComment(ctx.modelSnapshot.value, commentId)) {
      return false
    }
    ctx.showComments.value = true
    commentComposer.value = { mode: "reply", commentId }
    return true
  }

  function cancelCommentComposer(): void {
    closeComposer()
  }

  // Undo, redo and imports can remove the paragraph or comment the open editor
  // is attached to; a composer pointing at a missing target cannot be submitted.
  watch(ctx.model, (current) => {
    const composer = commentComposer.value
    if (!composer) return
    if (composer.mode === "reply") {
      if (
        composer.commentId === undefined ||
        !findDocumentComment(current, composer.commentId)
      ) {
        closeComposer()
      }
      return
    }
    if (composer.range) {
      const { paragraph } = getParagraphAtLocation(
        current,
        composer.range.start.location,
      )
      if (!paragraph) closeComposer()
    }
  })

  function addComment(
    range: DocxTextRange | undefined,
    text: string,
  ): number | undefined {
    if (text.trim().length === 0) {
      ctx.status.value = ctx.statusMessage("commentTextRequired")
      return undefined
    }
    const target = range ?? ctx.activeTextRangeSnapshot.value
    if (!target) {
      ctx.status.value = ctx.statusMessage("commentTargetMissing")
      return undefined
    }

    let commentId: number | undefined
    applyChange(
      (current) => {
        const result = addDocumentComment(current, {
          range: target,
          text,
          ...authorMetadata(),
        })
        if (!result) {
          return current
        }
        commentId = result.commentId
        return result.model
      },
      ctx.statusMessage("commentAdded"),
    )

    if (commentId === undefined) {
      ctx.status.value = ctx.statusMessage("commentTargetMissing")
      return undefined
    }

    ctx.showComments.value = true
    closeComposer()
    return commentId
  }

  function replyToComment(commentId: number, text: string): number | undefined {
    if (text.trim().length === 0) {
      ctx.status.value = ctx.statusMessage("commentTextRequired")
      return undefined
    }
    if (!findDocumentComment(ctx.modelSnapshot.value, commentId)) {
      ctx.status.value = ctx.statusMessage("commentTargetMissing")
      return undefined
    }

    let createdId: number | undefined
    applyChange(
      (current) => {
        const result = replyToDocumentComment(current, commentId, text, authorMetadata())
        if (!result) {
          return current
        }
        createdId = result.commentId
        return result.model
      },
      ctx.statusMessage("commentReplied"),
    )

    if (createdId === undefined) {
      ctx.status.value = ctx.statusMessage("commentTargetMissing")
      return undefined
    }

    ctx.showComments.value = true
    closeComposer()
    return createdId
  }

  function resolveComment(commentId: number, resolved = true): boolean {
    let applied = false
    applyChange(
      (current) => {
        const next = resolveDocumentComment(current, commentId, resolved)
        if (!next || next === current) {
          return current
        }
        applied = true
        return next
      },
      ctx.statusMessage(resolved ? "commentResolved" : "commentReopened"),
    )
    if (!applied) {
      ctx.status.value = ctx.statusMessage("commentTargetMissing")
    }
    return applied
  }

  function deleteComment(commentId: number): boolean {
    let applied = false
    applyChange(
      (current) => {
        const next = deleteDocumentComment(current, commentId)
        if (!next) {
          return current
        }
        applied = true
        return next
      },
      ctx.statusMessage("commentDeleted"),
    )
    if (!applied) {
      ctx.status.value = ctx.statusMessage("commentTargetMissing")
      return false
    }

    if (
      commentComposer.value?.mode === "reply" &&
      commentComposer.value.commentId !== undefined &&
      !findDocumentComment(ctx.modelSnapshot.value, commentComposer.value.commentId)
    ) {
      closeComposer()
    }
    return true
  }

  /** Submits the open composer; the text is trimmed by the model op. */
  function submitCommentComposer(text: string): boolean {
    const composer: DocxCommentComposer | undefined = commentComposer.value
    if (!composer) {
      return false
    }
    if (text.trim().length === 0) {
      return false
    }

    if (composer.mode === "reply" && composer.commentId !== undefined) {
      return replyToComment(composer.commentId, text) !== undefined
    }
    if (composer.range) {
      return addComment(composer.range, text) !== undefined
    }
    return false
  }

  return {
    commentComposer,
    beginComment,
    beginCommentReply,
    cancelCommentComposer,
    submitCommentComposer,
    addComment,
    replyToComment,
    resolveComment,
    deleteComment,
  }
}
