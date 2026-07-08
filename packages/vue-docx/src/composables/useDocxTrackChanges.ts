// Vue composable: useDocxTrackChanges
// Migrated from upstream @extend-ai/react-docx, editor.tsx lines 31542-31587
//
// Exposes tracked changes state and location-indexed change lookups.

import {
  paragraphLocationKey,
  type DocxEditorController,
  type DocxTextRangeLocation,
  type DocxTrackedChange,
} from "@extend-ai/docx-core"

export interface UseDocxTrackChangesResult {
  trackedChanges: DocxTrackedChange[]
  showTrackedChanges: boolean
  setShowTrackedChanges: (showTrackedChanges: boolean) => void
  toggleShowTrackedChanges: () => void
  changesByLocation: Map<string, DocxTrackedChange[]>
  getChangesForLocation: (location: DocxTextRangeLocation) => DocxTrackedChange[]
}

export function useDocxTrackChanges(
  editor: Pick<
    DocxEditorController,
    "trackedChanges" | "showTrackedChanges" | "setShowTrackedChanges" | "toggleShowTrackedChanges"
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
    get showTrackedChanges() { return editor.showTrackedChanges },
    setShowTrackedChanges: editor.setShowTrackedChanges,
    toggleShowTrackedChanges: editor.toggleShowTrackedChanges,
    changesByLocation,
    getChangesForLocation,
  }
}
