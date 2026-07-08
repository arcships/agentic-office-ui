// Vue composable: useDocxComments
// Migrated from upstream @extend-ai/react-docx, editor.tsx lines 31589-31636
//
// Exposes document comments state and location-indexed comment lookups.

import {
  paragraphLocationKey,
  type DocxEditorController,
  type DocxTextRangeLocation,
  type DocxComment,
} from "@extend-ai/docx-core"

export interface UseDocxCommentsResult {
  comments: DocxComment[]
  showComments: boolean
  setShowComments: (showComments: boolean) => void
  toggleShowComments: () => void
  commentsByLocation: Map<string, DocxComment[]>
  getCommentsForLocation: (location: DocxTextRangeLocation) => DocxComment[]
}

export function useDocxComments(
  editor: Pick<
    DocxEditorController,
    "comments" | "showComments" | "setShowComments" | "toggleShowComments"
  >
): UseDocxCommentsResult {
  const commentsByLocation = new Map<string, DocxComment[]>()
  for (const comment of editor.comments) {
    const key = paragraphLocationKey(comment.location)
    const bucket = commentsByLocation.get(key) ?? []
    bucket.push(comment)
    commentsByLocation.set(key, bucket)
  }

  const getCommentsForLocation = (location: DocxTextRangeLocation): DocxComment[] => {
    return commentsByLocation.get(paragraphLocationKey(location)) ?? []
  }

  return {
    get comments() { return editor.comments },
    get showComments() { return editor.showComments },
    setShowComments: editor.setShowComments,
    toggleShowComments: editor.toggleShowComments,
    commentsByLocation,
    getCommentsForLocation,
  }
}
