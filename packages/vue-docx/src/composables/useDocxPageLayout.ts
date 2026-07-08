// Vue composable: useDocxPageLayout
// Migrated from upstream @extend-ai/react-docx, editor.tsx lines 31638-31691
//
// Resolves page layout information from the document model's section properties.

import {
  parseSectionLayout,
  parseSectionColumns,
  parseSectionPageNumberStart,
  type DocxEditorController,
  type DocxPageLayoutInfo,
} from "@extend-ai/docx-core"

export interface UseDocxPageLayoutResult {
  layout: DocxPageLayoutInfo
}

export function useDocxPageLayout(
  editor: Pick<DocxEditorController, "model">
): UseDocxPageLayoutResult {
  const primarySectionPropertiesXml =
    editor.model.metadata.sections?.[0]?.sectionPropertiesXml ??
    editor.model.metadata.sectionPropertiesXml

  const sectionLayout = parseSectionLayout(primarySectionPropertiesXml)
  const sectionColumns = parseSectionColumns(primarySectionPropertiesXml)
  const pageNumberStart = parseSectionPageNumberStart(primarySectionPropertiesXml)

  const contentWidthPx = Math.max(
    120,
    sectionLayout.pageWidthPx -
      sectionLayout.marginsPx.left -
      sectionLayout.marginsPx.right
  )
  const contentHeightPx = Math.max(
    120,
    sectionLayout.pageHeightPx -
      sectionLayout.marginsPx.top -
      sectionLayout.marginsPx.bottom
  )

  const layout: DocxPageLayoutInfo = {
    pageWidthPx: sectionLayout.pageWidthPx,
    pageHeightPx: sectionLayout.pageHeightPx,
    contentWidthPx,
    contentHeightPx,
    marginsPx: { ...sectionLayout.marginsPx },
    headerDistancePx: sectionLayout.headerDistancePx,
    footerDistancePx: sectionLayout.footerDistancePx,
    pageNumberStart,
    columns: sectionColumns
      ? {
          count: sectionColumns.count,
          gapPx: sectionColumns.gapPx,
        }
      : undefined,
    viewportDefaults: {
      zoomPercent: 100,
      pageGapPx: 16,
    },
  }

  return { layout }
}
