// Vue composable: useDocxTrackChanges
// Migrated from upstream @extend-ai/react-docx, editor.tsx lines 31542-31587
//
// Exposes tracked changes state, location-indexed change lookups and the
// accept/reject decisions for the collected revisions.

import {
  paragraphLocationKey,
  type DocxEditorController,
  type DocxTextRangeLocation,
  type DocxTrackedChange,
} from "@arcships/docx-core"

export interface UseDocxTrackChangesResult {
  trackedChanges: DocxTrackedChange[]
  /** `null` while revisions are hidden, because the collection is skipped. */
  trackedChangeCount: number | null
  showTrackedChanges: boolean
  setShowTrackedChanges: (showTrackedChanges: boolean) => void
  toggleShowTrackedChanges: () => void
  changesByLocation: Map<string, DocxTrackedChange[]>
  getChangesForLocation: (location: DocxTextRangeLocation) => DocxTrackedChange[]
  acceptTrackedChange: (changeId: string) => boolean
  rejectTrackedChange: (changeId: string) => boolean
  acceptAllTrackedChanges: () => boolean
  rejectAllTrackedChanges: () => boolean
}

export function useDocxTrackChanges(
  editor: Pick<
    DocxEditorController,
    | "trackedChanges"
    | "trackedChangeCount"
    | "showTrackedChanges"
    | "setShowTrackedChanges"
    | "toggleShowTrackedChanges"
    | "acceptTrackedChange"
    | "rejectTrackedChange"
    | "acceptAllTrackedChanges"
    | "rejectAllTrackedChanges"
  >
): UseDocxTrackChangesResult {
  const changesByLocation = new Map<string, DocxTrackedChange[]>()
  for (const change of editor.trackedChanges) {
    const key = paragraphLocationKey(change.location)
    const bucket = changesByLocation.get(key) ?? []
    bucket.push(change)
    changesByLocation.set(key, bucket)
  }

  const getChangesForLocation = (location: DocxTextRangeLocation): DocxTrackedChange[] => {
    return changesByLocation.get(paragraphLocationKey(location)) ?? []
  }

  return {
    get trackedChanges() { return editor.trackedChanges },
    get trackedChangeCount() { return editor.trackedChangeCount },
    get showTrackedChanges() { return editor.showTrackedChanges },
    setShowTrackedChanges: editor.setShowTrackedChanges,
    toggleShowTrackedChanges: editor.toggleShowTrackedChanges,
    changesByLocation,
    getChangesForLocation,
    acceptTrackedChange: editor.acceptTrackedChange,
    rejectTrackedChange: editor.rejectTrackedChange,
    acceptAllTrackedChanges: editor.acceptAllTrackedChanges,
    rejectAllTrackedChanges: editor.rejectAllTrackedChanges,
  }
}
