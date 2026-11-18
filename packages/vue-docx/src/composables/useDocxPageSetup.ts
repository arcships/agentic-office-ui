// Vue composable: useDocxPageSetup
//
// Page-setup authoring state for the editor chrome. Resolves every section's
// page geometry from the model, exposes the section the caret is in, and
// forwards margin/paper-size/header-footer commands to the controller so each
// edit is one undoable step.

import { computed, ref, watch, type ComputedRef, type Ref } from "vue"
import {
  resolveDocumentSectionsFromMetadata,
  resolvePageSetup,
  resolveSectionIndexForNodeIndex,
  type DocxEditorController,
  type DocxPageMargins,
  type DocxPageMarginsUpdate,
  type DocxPageSize,
  type DocxPageSizeUpdate,
  type DocxSectionPageSetup,
  type DocxSectionRegion,
} from "@arcships/docx-core"

export interface DocxPageSetupSection extends DocxSectionPageSetup {
  /** One-based section number, as Word labels it ("Section 1"). */
  sectionNumber: number
}

export interface DocxPageSetupPatch {
  margins?: DocxPageMarginsUpdate
  pageSize?: DocxPageSizeUpdate
}

export interface UseDocxPageSetupResult {
  /** Every section in document order. */
  sections: ComputedRef<DocxPageSetupSection[]>
  /** Section the page-setup commands target: the caret's section. */
  activeSectionIndex: Ref<number>
  activeSection: ComputedRef<DocxPageSetupSection>
  pageSize: ComputedRef<DocxPageSize>
  margins: ComputedRef<DocxPageMargins>
  selectSection: (sectionIndex: number) => void
  applyMargins: (margins: DocxPageMarginsUpdate) => boolean
  applyPageSize: (pageSize: DocxPageSizeUpdate) => boolean
  /** Apply paper size and margins as one undoable step. */
  applyPageSetup: (patch: DocxPageSetupPatch) => boolean
  addHeaderFooter: (kind: DocxSectionRegion, referenceType?: string) => boolean
}

function clampSectionIndex(sectionIndex: number, sectionCount: number): number {
  if (!Number.isFinite(sectionIndex)) return 0
  return Math.max(0, Math.min(Math.round(sectionIndex), Math.max(0, sectionCount - 1)))
}

/**
 * Section that contains the current selection (paragraph or table cell).
 * Layout falls back to the first section when the caret sits outside any
 * section metadata.
 */
export function resolveCaretSectionIndex(
  controller: Pick<DocxEditorController, "model" | "selection">
): number {
  const selection = controller.selection
  if (!selection) return 0

  const nodeIndex =
    selection.kind === "paragraph"
      ? selection.nodeIndex
      : selection.kind === "table-cell"
        ? selection.tableIndex
        : undefined
  if (nodeIndex === undefined) return 0

  const sections = resolveDocumentSectionsFromMetadata(controller.model.metadata)
  if (sections.length === 0) return 0
  return resolveSectionIndexForNodeIndex(sections, nodeIndex, 0)
}

export function useDocxPageSetup(
  controller: Pick<
    DocxEditorController,
    "model" | "selection" | "setPageMargins" | "setPageSize" | "addHeaderFooter"
  >
): UseDocxPageSetupResult {
  const sections = computed<DocxPageSetupSection[]>(() =>
    resolvePageSetup(controller.model).map((section) => ({
      ...section,
      sectionNumber: section.index + 1,
    }))
  )

  const activeSectionIndex = ref(
    clampSectionIndex(resolveCaretSectionIndex(controller), sections.value.length)
  )

  // The caret decides the section until the user picks one explicitly; after
  // that the picker stays in control for the rest of the dialog session.
  let pinnedSectionIndex: number | undefined
  watch(
    () => [controller.selection, controller.model.metadata.sections, controller.model.metadata.sectionPropertiesXml],
    () => {
      if (pinnedSectionIndex !== undefined) {
        activeSectionIndex.value = clampSectionIndex(
          pinnedSectionIndex,
          sections.value.length
        )
        return
      }
      activeSectionIndex.value = clampSectionIndex(
        resolveCaretSectionIndex(controller),
        sections.value.length
      )
    }
  )

  watch(
    () => sections.value.length,
    (count) => {
      activeSectionIndex.value = clampSectionIndex(activeSectionIndex.value, count)
    }
  )

  const activeSection = computed<DocxPageSetupSection>(() => {
    const list = sections.value
    return list[activeSectionIndex.value] ?? list[0]
  })

  const pageSize = computed(() => activeSection.value.pageSize)
  const margins = computed(() => activeSection.value.margins)

  const selectSection = (sectionIndex: number): void => {
    pinnedSectionIndex = clampSectionIndex(sectionIndex, sections.value.length)
    activeSectionIndex.value = pinnedSectionIndex
  }

  const applyMargins = (update: DocxPageMarginsUpdate): boolean =>
    controller.setPageMargins(activeSectionIndex.value, update)

  const applyPageSize = (update: DocxPageSizeUpdate): boolean =>
    controller.setPageSize(activeSectionIndex.value, update)

  const applyPageSetup = (patch: DocxPageSetupPatch): boolean => {
    const sizeChanged = patch.pageSize
      ? controller.setPageSize(activeSectionIndex.value, patch.pageSize)
      : false
    // Fold the margin change into the paper-size entry so "Apply" is one undo.
    const marginsChanged = patch.margins
      ? controller.setPageMargins(activeSectionIndex.value, patch.margins, {
          pushHistory: !sizeChanged,
        })
      : false
    return sizeChanged || marginsChanged
  }

  const addHeaderFooter = (
    kind: DocxSectionRegion,
    referenceType?: string
  ): boolean =>
    controller.addHeaderFooter(activeSectionIndex.value, kind, { referenceType })

  return {
    sections,
    activeSectionIndex,
    activeSection,
    pageSize,
    margins,
    selectSection,
    applyMargins,
    applyPageSize,
    applyPageSetup,
    addHeaderFooter,
  }
}
