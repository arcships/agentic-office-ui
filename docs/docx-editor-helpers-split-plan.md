# DOCX Editor Helpers 模块化拆分方案

## 源文件概况

- **文件**: `react-docx/packages/react-viewer/src/editor.tsx`
- **总行数**: 56,454
- **分析范围**: 第 1-24,953 行（纯函数/类型/常量区，零 React hook）
- **定义数量**: 881 个顶层定义（const / function / interface / type / class / let / var）
- **第一个 React hook**: `useDocxEditor` 在行 24,954

## 已有外部依赖（已拆分到独立文件）

以下功能已从 `editor.tsx` 拆分到独立模块并由 `editor.tsx` 引用：

| 文件 | 提供的能力 |
|------|-----------|
| `pagination-breaks.ts` | `collectTableExplicitPageBreakInfo`, `collectTopLevelExplicitPageBreakStartNodeIndexes` |
| `page-count-reconciliation.ts` | `reconcilePageCountCandidateToTargetCountByScalingHeight`, `resolveMeasuredBodyFooterOverlapLatchState`, `shouldAllowStoredPageCountReduction` |
| `section-layout.ts` | `DEFAULT_DOCUMENT_LAYOUT`, `pageMarginPaddingStyle`, `resolveDocumentLayout`, `parseSectionLayout`, `parseSectionPageBorders`, `TWIPS_PER_PIXEL`, `DocumentLayoutMetrics`, `twipsToPixels` |
| `image-render.ts` | `imageUsesPlaceholderFallback`, `resolveRenderableImageSource`, `subscribeRenderableImageSourceUpdates`, `unsupportedImageFallbackLabel` |
| `pretext-layout.ts` | `layoutItemsWithPretextAroundExclusions`, `layoutTextWithPretextAroundExclusions`, `measurePretextPlainTextLineCount`, `PretextLayoutItem`, `PretextExclusionRect`, `PretextLineFragment`, `PretextSelectionRect`, `PretextVariableWidthLayout`, `resolveCaretRectAtOffset`, `resolveOffsetAtPoint`, `resolveSelectionRects`, `sliceLayoutToLineRange` |
| `content-signature.ts` | `docModelThumbnailMetadataSignature`, `docNodeContentSignature` |
| `thumbnail-raster.ts` | `blitDocxThumbnailSurface`, `DOCX_THUMBNAIL_EXCLUDE_ATTRIBUTE`, `DocxPageThumbnailRenderSnapshot`, `DocxPageThumbnailSnapshotElement`, `DocxPageThumbnailTableCellSnapshot`, `DocxPageThumbnailTextRunSnapshot`, `DocxThumbnailSurfaceCache`, `rasterizeDocxThumbnailSurface`, `renderDocxThumbnailSnapshotSurface`, `SerialIdleTaskQueue` |

**约束**: 拆分模块时不得产生对这 7 个文件的循环依赖。它们都是 editor.tsx 的下游依赖。

---

## 拆分方案总览

共 **18 个模块文件**，覆盖 24,953 行。每个文件 500-2,500 行。

| # | 文件名 | 行数范围 | ~行数 |
|---|--------|----------|-------|
| 1 | `constants.ts` | 98-445, 892-980, 1383-1431 | ~480 |
| 2 | `editor-types.ts` | 109-110, 3055-4544 | ~1,520 |
| 3 | `cache-utils.ts` | 352-510 | ~160 |
| 4 | `ooxml-helpers.ts` | 1488-1691 | ~205 |
| 5 | `header-footer.ts` | 1692-2315 | ~625 |
| 6 | `page-measurement.ts` | 2315-3053 | ~740 |
| 7 | `default-model.ts` | 4546-4644 | ~100 |
| 8 | `paragraph-inspection.ts` | 4646-5811, 492-616, 7936-8640 | ~2,350 |
| 9 | `pretext-integration.ts` | 5811-7545, 6945-7210 | ~1,800 |
| 10 | `drop-cap.ts` | 543-616, 7625-7940 | ~450 |
| 11 | `line-height-estimate.ts` | 8923-10391 | ~1,470 |
| 12 | `table-height-estimate.ts` | 10541-11590 | ~1,050 |
| 13 | `pagination-plan.ts` | 8654-8923, 11481-13986 | ~2,500 |
| 14 | `style-to-css.ts` | 14888-16049 | ~1,160 |
| 15 | `xml-parsing.ts` | 16049-17170 | ~1,120 |
| 16 | `synthetic-textbox.ts` | 16414-17158 | ~750 |
| 17 | `tracked-changes.ts` | 382-418, 17170-17561, 22675-23442 | ~1,100 |
| 18 | `field-helpers.ts` | 15864-16049, 17653-18086 | ~700 |
| 19 | `numbering.ts` | 19581-20471 | ~600 |
| 20 | `paragraph-render.tsx` | 18126-19581 | ~1,455 |
| 21 | `table-utils.ts` | 20511-21915 | ~1,405 |
| 22 | `text-mutation.ts` | 21915-22670 | ~760 |
| 23 | `selection-helpers.ts` | 22534-24666, 24314-24364 | ~1,210 |
| 24 | `section-manipulation.ts` | 24641-24952 | ~310 |

> 注: 部分模块之间有共享行号范围（如 `xml-parsing.ts` 和 `synthetic-textbox.ts`），这是因为它们的功能在原始文件中交错排列，拆分时需要割裂连续行号。

---

## 逐模块详细清单

### 1. `constants.ts` (~480 行)

**功能领域**: 全局常量和编译期配置

**包含定义**:

| 行号 | 名称 | 类型 |
|------|------|------|
| 98-107 | `HIGHLIGHT_TO_CSS` | const |
| 109 | `DocxDocumentTheme` | export type |
| 110-112 | `DocxHeadingStyleMap` | export type |
| 114-151 | `DEFAULT_WORD_HEADING_STYLES` | const |
| 153-193 | `DEFAULT_WORD_HEADING_RUN_STYLES` | const |
| 195-225 | 排版常量 (DOC_PAGE_BREAK_GAP ~ WORD_TABLE_CELL_PARAGRAPH_AFTER_TWIPS) | const x15 |
| 266-269 | 表格切片常量 (TABLE_ROW_SLICE_VISUAL_BLEED_PX ~ DEFAULT_SPLIT_PARAGRAPH_AFTER_TWIPS) | const x4 |
| 270 | `PAGE_CONTENT_MEASUREMENT_IGNORE_ATTRIBUTE` | const |
| 271-276 | `WORD_TABLE_CELL_FALLBACK_PADDING_PX` | const |
| 277-281 | 字体/表格常量 (DEFAULT_UNSPECIFIED_DOCX_FONT_FAMILY ~ SPLITTABLE_TABLE_ROW_DEEP_CONTENT_NODE_THRESHOLD) | const x5 |
| 282-293 | XML 正则模式 (PAGE_BREAK_XML_PATTERN ~ ENDNOTE_REFERENCE_XML_PATTERN) | const x8 |
| 294-329 | 缓存/分页/测量常量 (XML_CACHE_MAX_ENTRIES ~ TABLE_ROW_HEIGHT_PAGINATION_ESTIMATE_PADDING_MIN_ROWS) | const x17 |
| 321-445 | 更多布局常量 (MIN_TABLE_ROW_SLICE_REMAINING_HEIGHT_PX ~ MODIFIER_ONLY_KEYS) | const x12 |
| 892-917 | 主题/滤镜常量 (DOC_SURFACE_STYLE_BY_THEME, NIGHT_READER_INVERSION_FILTER, appendCssFilters) | const x3 |
| 926-980 | 表格操作柄常量 (TABLE_RESIZE_HANDLE_SIZE ~ INITIAL_PAGINATION_STABILITY_IDLE_MS) | const x24 |
| 1284-1347 | 右键菜单默认配置 (DEFAULT_TABLE_CONTEXT_MENU_ACTIONS ~ DEFAULT_CONTEXT_MENU_IMAGE_LAYER_ACTIONS) | const x3 |
| 1383-1431 | 图片 wrap 模式常量 (DOCX_IMAGE_WRAP_MODE_ACTIONS, WORD_IMAGE_Z_INDEX_*) | const x2 |
| 1494-1508 | `DEFAULT_PAGE_NUMBER_START`, `DEFAULT_TOOLBAR_BORDER_SIZE_EIGHTH_PT`, `DEFAULT_TOOLBAR_BORDER_COLOR` | const x3 |
| 20078-20084 | 列表前缀正则 (UNORDERED_LIST_PREFIX_PATTERN ~ MAX_FALLBACK_LIST_LEVEL) | const x7 |

**依赖**: 无内部模块依赖。仅依赖 `@extend-ai/react-docx-doc-model` 的类型。

---

### 2. `editor-types.ts` (~1,520 行)

**功能领域**: 公共 API 类型定义（所有 `export type` / `export interface`）

**包含定义** (行 3055-4544):

- `DocxListType` (3055)
- `DocxEditorSelection` (3057-3067)
- `DocxTableCellLocation` (3069-3073)
- `DocxTableCellSelectionRange` (3075-3081)
- `ParagraphLocationInBody` (3083-3086)
- `ParagraphLocationInCell` (3088-3094)
- `DocxTextRangeLocation` (3096-3098)
- `DocxTextRangeBoundary` (3100-3103)
- `DocxTextRange` (3105-3108)
- `DocxHistorySnapshot` (3482-3489)
- `DocxHistoryRestoreRequest` (3488-3493)
- `DocxEditorTransactionContext` (3494-3500)
- `DocxEditorTransactionPatch` (3501-3509)
- `ParagraphLocation` (3511)
- `DocxImageLocation` (3513-3515)
- `DocxSectionRegion` (3517)
- `DocxSectionParagraphLocation` (3519-3526)
- `DocxSectionImageLocation` (3528-3530)
- `DocxFormFieldLocation` (3532-3534)
- `DocxSelectedFormField` (3536-3539)
- `DocxImageDropTarget` (3541-3543)
- `DocxTrackedChangeKind` (3545-3552)
- `DocxTrackedChange` (3553-3562)
- `DocxComment` (3571-3589)
- `DocxLineSpacingRule` (3591)
- `DocxSelectionSessionKind` (3592-3598)
- `DocxLineSpacingInfo` (3600-3604)
- `DocxBorderContext` (3606)
- `DocxBorderPreset` (3608-3621)
- `DocxBorderPresetState` (3623)
- `UseDocxEditorOptions` (3631-3686)
- `DocxEditorController` (3688-3879)
- `DocxTableContextMenuActionId` (3881-3888)
- `DocxContextMenuActionId` (3890-3911)
- `DocxTableContextMenuContext` (3913-3920)
- `DocxTableContextMenuAction` (3922-3929)
- `DocxContextMenuAction` (3931-3951)
- `DocxTableContextMenuRenderProps` (3953-3969)
- `DocxContextMenuContext` (3971-3991)
- `DocxContextMenuRenderProps` (3993-4016)
- `DocxPageVirtualizationOptions` (4018-4068)
- `DocxVisiblePageRange` (4070-4103)
- `DocxEditorViewerProps` (4105-4225)
- `DocxTrackedChangeCardRenderProps` (4227-4244)
- `DocxCommentCardRenderProps` (4246-4267)
- `DocxEditorViewerMode` (4269)
- `UseDocxDocumentThemeResult` (4271-4276)
- `UseDocxParagraphStylesResult` (4278-4282)
- `UseDocxLineSpacingResult` (4284-4287)
- `UseDocxBordersResult` (4289-4293)
- `UseDocxFormFieldsResult` (4295-4308)
- `UseDocxTrackChangesResult` (4310-4319)
- `UseDocxCommentsResult` (4321-4328)
- `DocxSectionColumnLayout` (4330-4333)
- `DocxPageLayoutInfo` (4335-4354)
- `UseDocxPageLayoutResult` (4356-4358)
- `DocxPaginationInfo` (4360-4363)
- `UseDocxPaginationResult` (4365-4367)
- `DocxPageThumbnailResolutionOptions` (4369-4412)
- `DocxPageThumbnailBounds` (4414-4421)
- `DocxPageThumbnailResolution` (4423-4434)
- `DocxPageThumbnailRenderWindow` (4436-4451)
- `UseDocxPageThumbnailsOptions` (4453-4489)
- `DocxPageThumbnailStatus` (4491-4496)
- `DocxPageThumbnailItem` (4498-4529)
- `UseDocxPageThumbnailsResult` (4531-4541)
- `UseDocxViewerThumbnailsOptions` (4543)
- `DocxViewerThumbnails` (4544)

**依赖**: `@extend-ai/react-docx-doc-model` 的类型, `React.CSSProperties`

---

### 3. `cache-utils.ts` (~160 行)

**功能领域**: 全局缓存 Map 声明和缓存工具函数

**包含定义**:

| 行号 | 名称 | 类型 |
|------|------|------|
| 352-361 | `paragraphBreakFlagsBySourceXml` | const (Map) |
| 362-365 | `paragraphEstimatedHeightBySourceXml` | const (Map) |
| 366 | `tableEstimatedHeightBySourceXml` | const (Map) |
| 367-370 | `tableEstimatedRowHeightsByNode` | const (WeakMap) |
| 371-374 | `paragraphExplicitIndentBySourceXml` | const (Map) |
| 375-381 | `paragraphDropCapBySourceXml` | const (Map) |
| 419 | `paragraphMeasureCanvasContext` | let |
| 420 | `textWidthByFontAndValue` | const (Map) |
| 421 | `estimatedTextAdvanceWidthByFontAndValue` | const (Map) |
| 422 | `pretextWordBreakModeByText` | const (Map) |
| 423 | `paragraphBaseFontSizePxByParagraph` | const (WeakMap) |
| 424-428 | `paragraphDominantFontFamilyByParagraph` | const (WeakMap) |
| 436-444 | `setCacheEntry` | function |
| 446-452 | `widthCacheKeyPx` | function |
| 454-468 | `heightEstimateCacheKeyPx` | function |

**依赖**: 无内部模块依赖。依赖 `@extend-ai/react-docx-doc-model` 类型。

---

### 4. `ooxml-helpers.ts` (~205 行)

**功能领域**: OOXML 包解析、内嵌字体、twips/points 转换

**包含定义**:

| 行号 | 名称 | 类型 |
|------|------|------|
| 1488-1492 | `SectionColumnLayout` | interface |
| 1496-1505 | `createDefaultEditorTableBorders` | function |
| 1510-1516 | `twipsToSignedPixels` | function |
| 1518-1524 | `pointsToPixels` | function |
| 1526-1531 | `EmbeddedFontFaceDescriptor` | type |
| 1533-1540 | `relationshipPartNameForOoxmlPart` | function |
| 1542-1573 | `resolveRelativeOoxmlPartName` | function |
| 1575-1604 | `parseOoxmlRelationships` | function |
| 1606-1632 | `deobfuscateEmbeddedFontData` | function |
| 1634-1691 | `collectEmbeddedFontFaceDescriptors` | function |

**依赖**: `cache-utils.ts` (TWIPS_PER_PIXEL 来自 section-layout.ts 外部模块), `@extend-ai/react-docx-ooxml-core` (OoxmlPackage), `@extend-ai/react-docx-doc-model` (TableBorderSet)

---

### 5. `header-footer.ts` (~625 行)

**功能领域**: 节(section)解析、页眉/页脚预留空间计算

**包含定义**:

| 行号 | 名称 | 类型 |
|------|------|------|
| 1692-1750 | `parseSectionColumns` | function |
| 1768-1784 | `resolveSectionPaginationContentWidthPx` | export function |
| 1786-1810 | `parseSectionPageNumberStart` | function |
| 1811-1837 | `parseSectionPageNumberStartOverride` | function |
| 1838-1851 | `parseSectionPageNumberFormat` | function |
| 1852-1866 | `parseSectionStartType` | function |
| 1867-1873 | `ResolvedDocumentSection` | interface |
| 1874-1878 | `normalizeSectionReferenceType` | function |
| 1879-1902 | `inheritSectionReferences` | function |
| 1903-1915 | `resolveInheritedSectionHeaderFooterReferences` | function |
| 1916-1955 | `resolveDocumentSectionsFromMetadata` | function |
| 1956-1963 | `PaginationSectionMetrics` | interface |
| 1964-1993 | `paragraphHasHeaderFooterReserveRelevantContent` | function |
| 1994-2006 | `sectionHasVisibleHeaderContent` | function |
| 2007-2019 | `sectionHasVisibleFooterContent` | function |
| 2020-2044 | `resolveHeaderFooterAbsoluteFloatingTopPx` | function |
| 2045-2054 | `shouldReserveHeaderFooterFloatingImageSpace` | function |
| 2055-2092 | `resolveFooterParagraphFloatingBoundaryTopPx` | function |
| 2093-2120 | `resolveFooterNodesFloatingBoundaryTopPx` | function |
| 2121-2216 | `estimateHeaderFooterParagraphFloatingReservePx` | function |
| 2217-2293 | `resolveHeaderPaginationReservePx` | export function |
| 2294-2315 | `resolveFooterPaginationReservePx` | export function |

**依赖**: `section-layout.ts` (DocumentLayoutMetrics, twipsToPixels, TWIPS_PER_PIXEL), `@extend-ai/react-docx-doc-model` (DocModel, HeaderSection, FooterSection)

---

### 6. `page-measurement.ts` (~740 行)

**功能领域**: 页面内容高度测量、脚注预留

**包含定义**:

| 行号 | 名称 | 类型 |
|------|------|------|
| 2433-2569 | `resolveMeasuredPageContentHeightDiagnostics` | function |
| 2571-2586 | `resolveMeasuredPageContentHeightPx` | export function |
| 2587-2616 | `resolveMeasuredBodyRenderedBottomPx` | export function |
| 2617-2636 | `stabilizeMeasuredPageContentHeights` | export function |
| 2643-2660 | `documentPageNodeSegmentIdentityKey` | function |
| 2661-2666 | `documentPageNodeSegmentsIdentityKey` | function |
| 2667-2720 | `buildPaginationSectionMetrics` | function |
| 2729-2756 | `resolvePaginationSectionMetricsIndexForNodeIndex` | function |
| 2757-2773 | `scalePaginationSectionMetricsHeights` | function |
| 2774-2793 | `scaleMeasuredPageContentHeights` | function |
| 2794-2833 | `resolvePageContentHeightPxForPageSegments` | export function |
| 2834-2892 | `resolveRenderPageContentHeightPxForPageSegments` | export function |
| 2893-2905 | `documentPageContainsOnlySplitParagraphSegments` | export function |
| 2906-2944 | `estimateDocumentNoteSectionHeightPx` | function |
| 2945-2977 | `collectReferencedFootnotesForPageSegments` | function |
| 2978-3053 | `applyEstimatedFootnoteReserveToPages` | function |

**依赖**: `header-footer.ts` (PaginationSectionMetrics), `section-layout.ts` (DocumentLayoutMetrics), `constants.ts`, `@extend-ai/react-docx-doc-model`

---

### 7. `default-model.ts` (~100 行)

**功能领域**: 默认空白文档模型

**包含定义**:

| 行号 | 名称 | 类型 |
|------|------|------|
| 4546-4641 | `defaultStarterModel` | export const |
| 4642-4645 | `createBlankDocumentModel` | function |

**依赖**: `@extend-ai/react-docx-doc-model` (DocModel)

---

### 8. `paragraph-inspection.ts` (~2,350 行)

**功能领域**: 段落结构分析、浮动图片 anchor 检测、分页断点检测、封面图检测

**包含定义**:

**A. 段落文本/工具函数** (4646-4904):
- `textRuns` (4646)
- `paragraphText` (4652)
- `nodeTreeContainsExplicitFontFamily` (4658)
- `firstExplicitFontFamilyInNodeTree` (4681)
- `resolveDocumentInheritedFontFamily` (4716)
- `replaceTabLayoutMarkersWithTabText` (4780)
- `editableTextFromElement` (4844)
- `editableTextFromTableCellElement` (4858)
- `editableTextFromDraftHtml` (4879)
- `editableTextFromTableCellDraftHtml` (4892)
- `renderStaticHtml` (4902)
- `formFieldDisplayValue` (4906)

**B. 段落/节点基本检查** (4923-5090):
- `firstRunStyle` (4923)
- `ensureTextRunNode` (4929)
- `isParagraphCellContentNode` (4946)
- `isTableCellTableContentNode` (4952)
- `tableCellParagraphs` (4958)
- `tableCellParagraphsRecursively` (4964)
- `tableCellHasImage` (4988)

**C. 段落样式解析** (478-541):
- `headingLevelFromStyleLabel` (478)
- `resolveParagraphStyleHeadingLevel` (492)
- `normalizeParagraphStyleDefinitionsForUi` (512)

**D. 浮动图片/封面检测** (5008-5811):
- `paragraphHasImage` (5008)
- `paragraphHasInFlowImage` (5012)
- `paragraphHasTextBearingAbsoluteFloatingTextBox` (5025)
- `paragraphIsFloatingImageAnchorOnly` (5038)
- `paragraphIsAbsoluteFloatingImageAnchorOnly` (5073)
- `paragraphContainsOnlyAbsoluteFloatingContent` (5109)
- `paragraphIsBehindTextAbsoluteFloatingImageAnchorOnly` (5132)
- `imageBehavesAsDecorativeBehindTextBackground` (5146)
- `paragraphActsAsDecorativeBehindTextBackgroundOverlay` (5157)
- `paragraphNeedsPageWidthAnchorHost` (5170)
- `paragraphHasAbsoluteFloatingImage` (5191)
- `collectHeadingTextColorByLevel` (5198)
- `sectionNodesNeedPageWideLayout` (5298)
- `sectionNodesNeedFullPageFooterOverlay` (5369)
- `paragraphHasFormField` (5393)
- `paragraphHasCheckboxFormField` (5397)
- `paragraphHasVisibleText` (5403)
- `normalizeFloatingTextBoxComparisonText` (5412)
- `floatingTextBoxVisibleTextFromImage` (5422)
- `paragraphVisibleTextIsOnlyAbsoluteFloatingTextBoxContent` (5447)
- `paragraphHasOnlyWhitespaceText` (5486)
- `paragraphHasActiveNumbering` (5500)
- `paragraphContainsSectionBreakProperties` (5509)
- `paragraphAbsoluteFloatingAnchorsDependOnParagraphFlow` (5515)
- `likelyFullPageCoverImageRelativeToContentBox` (5541)
- `isLikelyFullPageCoverFloatingImage` (5576, export)
- `resolveLinkedImageWrapperStyle` (5588, export)
- `paragraphIsLikelyFullPageCoverArtAnchor` (5606)
- `pageAnchoredImageLikelyStartsNearPageTop` (5625)
- `paragraphActsAsPageAnchoredCoverOverlayHost` (5653)
- `paragraphStartsNormalFlowContent` (5668)
- `paragraphParticipatesInLeadingCoverLayout` (5683)
- `fullPageCoverImageRenderKey` (5714)
- `paragraphActsAsLeadingCoverLayoutOverlay` (5721)
- `fullPageCoverAbsoluteFloatingImageStyle` (5754)
- `paragraphLooksLikeCheckboxChoiceRow` (5787)

**E. 段落结构分析** (7936-8640):
- `paragraphIsEffectivelyEmpty` (7936)
- `paragraphHasDeletedParagraphMark` (7946)
- `paragraphCollapsesIntoPreviousParagraph` (7955)
- `paragraphIsStructuralSectionBreakSpacer` (7966)
- `paragraphActsAsSectionBreakCarryoverSpacer` (7990)
- `paragraphActsAsTrailingRenderedPageBreakSpacer` (8012)
- `paragraphContextualSpacingStyleKey` (8055)
- `paragraphsSuppressInterParagraphSpacing` (8071)
- `effectiveParagraphAfterSpacingPx` (8092)
- `effectiveParagraphBeforeSpacingPx` (8108)
- `nodeHasSubstantiveContentForPagination` (8137)
- `paragraphBookmarkNames` (8159)
- `paragraphReferencedNoteIds` (8175)
- `nodeReferencedNoteIds` (8201)
- `eventTargetIsInteractiveControl` (8231)
- `eventTargetIsNestedTableParagraphEditor` (8241)
- `sectionBreakPropertiesStartNewPage` (8253)
- `paragraphHasExplicitPageBreak` (8279)
- `paragraphHasExplicitColumnBreak` (8310)
- `paragraphHasLastRenderedPageBreak` (8325)
- `paragraphStartsWithLastRenderedPageBreak` (8342)
- `shouldHonorParagraphStartLastRenderedPageBreak` (8371)
- `isOnOffTagEnabled` (8402)
- `sectionTitlePageEnabled` (8418)
- `selectSectionVariantForPage` (8429)
- `resolveSectionIndexForNodeIndex` (8476)
- `paragraphHasVisibleBorder` (8505)
- `paragraphIsSectionBreakAnchorCarryover` (8516)
- `resolveSectionIndexForPageSegments` (8537)
- `paragraphHasPageBreakBefore` (8596)
- `sectionBreakAfterParagraphStartsNewPage` (8615)
- `nodeAlreadyEndsAtExplicitPageBoundary` (8640)

**依赖**: `constants.ts`, `cache-utils.ts`, `section-layout.ts` (twipsToPixels), `@extend-ai/react-docx-doc-model`

---

### 9. `pretext-integration.ts` (~1,800 行)

**功能领域**: Pretext layout 集成、浮动图片 wrap 障碍物、双栏浮动图片几何

**包含定义**:

**A. 接口/类型** (5811-5883):
- `ParagraphPretextLayoutRun` (5813)
- `ParagraphPretextLayoutSource` (5825)
- `KEEP_ALL_SCRIPT_RE` (5830)
- `WrappedParagraphEditingSession` (5833)
- `WrappedParagraphSelectionDragState` (5844)
- `WrappedParagraphSurfaceRegistration` (5853)
- `DualWrappedFloatingImageGeometry` (5860)
- `MIN_DUAL_WRAPPED_INTERIOR_BAND_PX` (5871)
- `ParagraphDualWrappedTextLayout` (5873)
- `PageFlowFloatingWrapObstacle` (5880)

**B. Wrap 障碍物计算** (5884-6330):
- `resolveDualWrapParagraphRenderBlockHeightPx` (5884)
- `resolveForeignWrapExclusionsForFlowRange` (5905, export)
- `resolveWrappedImageGeometryForPageFlow` (5932)
- `collectPageFlowWrapObstaclesForParagraph` (5965, export)
- `resolveParagraphForeignOnlyWrappedTextLayout` (6097, export)
- `precomputePageSegmentForeignWrapExclusions` (6131)
- `applyWrappedFloatingInteractionPreviewToParagraph` (6232)

**C. Pretext 缓存 + Wrap 模式** (6331-7210):
- `paragraphPretextLayoutSourceCache` (6331)
- `paragraphPretextLayoutItemsBySource` (6335)
- `paragraphPretextUniformFontBySource` (6339)
- `imageWrapModeFromFloating` (6344)
- `floatingImageMovesWithText` (6368)
- `resolveDocxImageWrapState` (6379, export)
- `imageWrapModeActionId` (6390)
- `imageWrapModeFromActionId` (6413)
- `resolveFloatingForImageWrapMode` (6438)
- `buildParagraphPretextTabSpacerText` (6533)
- `buildParagraphPretextLayoutSource` (6586, export)
- `splitParagraphAtExplicitColumnBreaks` (6764)
- `estimateParagraphContentHeightPx` (6846)
- `projectParagraphConsumedHeightWithExplicitColumnBreaks` (6878)
- `buildSyntheticPretextLayoutSource` (6945)
- `pretextWordBreakModeForText` (6964)
- `sanitizeRenderedPretextFragmentText` (6984, export)
- `buildParagraphPretextLayoutItems` (6988)
- `resolveUniformPretextSourceFont` (7020)
- `buildMeasureSegmentsPretextLayoutItems` (7056)
- `layoutParagraphPretextSource` (7080)

**D. 文本扩展/编辑** (7120-7230):
- `wrappedParagraphSessionText` (7120)
- `buildInlineImagePlaceholderText` (7127)
- `paragraphChildAnchorOffset` (7147)
- `expandOffsetToWord` (7182)
- `pixelsToTwips` (7228)

**E. 双栏浮动图片几何** (7232-7545):
- `resolveDualWrappedFloatingImageGeometry` (7232, export)
- `resolveParagraphDualWrappedTextLayout` (7388, export)
- `resolveParagraphPretextExclusionLayout` (7545)

**依赖**: `pretext-layout.ts` (PretextExclusionRect, PretextLayoutItem, 各种 pretext 函数), `constants.ts`, `cache-utils.ts`, `section-layout.ts` (TWIPS_PER_PIXEL), `@extend-ai/react-docx-doc-model`

---

### 10. `drop-cap.ts` (~450 行)

**功能领域**: 首字下沉(drop cap)检测与视觉尺寸计算

**包含定义**:

| 行号 | 名称 | 类型 |
|------|------|------|
| 543-616 | `paragraphDropCap` | function |
| 7625-7645 | `checkboxChoiceRowTabWidthPx` | function |
| 7646-7666 | `attachTextToPreviousCheckbox` | function |
| 7667-7680 | `runFontSizePx` | function |
| 7681-7694 | `explicitRunFontSizePx` | function |
| 7695-7720 | `resolveDropCapFontSizePx` | export function |
| 7721-7745 | `resolveDropCapVisualHeightPx` | export function |
| 7747-7828 | `estimateTextAdvanceWidthPx` | function |
| 7829-7846 | `updateEstimatedLineWidthPxForText` | function |
| 7847-7879 | `resolveTabSpacerWidthPx` | function |
| 7880-7913 | `estimateInteractiveFieldWidthPx` | function |
| 7914-7935 | `resolveCheckboxFieldWidthPx` | export function |

**依赖**: `cache-utils.ts` (paragraphMeasureCanvasContext, textWidthByFontAndValue, estimatedTextAdvanceWidthByFontAndValue), `constants.ts`, `section-layout.ts` (TWIPS_PER_PIXEL), `@extend-ai/react-docx-doc-model`

---

### 11. `line-height-estimate.ts` (~1,470 行)

**功能领域**: 段落行高估算、字体分析、文本宽度测量

**包含定义**:

**A. 段落字体分析** (8923-9157):
- `paragraphDominantFontSizePt` (8923)
- `paragraphBaseFontSizePx` (8975)
- `paragraphMaxFontSizePx` (8995)
- `normalizeFontFamilyToken` (9012)
- `paragraphDominantFontFamily` (9025)
- `singleLineAutoScaleForFontFamily` (9077)
- `emptyParagraphLineScaleForFontFamily` (9103)
- `paragraphRendersTextFreeLine` (9132)
- `resolveParagraphSingleLineAutoScale` (9139)
- `paragraphLineCount` (9153)
- `paragraphContainsExplicitLineBreakText` (9157)

**B. 文本宽度测量** (9165-9632):
- `estimatedGlyphWidthPx` (9165)
- `fallbackMeasureTextWidthPx` (9185)
- `resolveMeasureFontSizePx` (9193)
- `resolveMeasureFont` (9209)
- `measureTextWidthPx` (9223)
- `resolveParagraphTabStopsPx` (9279)
- `resolveParagraphFirstLineOriginPx` (9292)
- `resolveParagraphFirstLineLeftTabStopsPx` (9310)
- `resolveNextTabStopPx` (9325)
- `ParagraphMeasureSegment` (9342)
- `ParagraphMeasureToken` (9347)
- `normalizeTextForTabLeaderMeasurement` (9351)
- `measureParagraphSegmentTextWidthPx` (9355)
- `estimateWrappedLineCountForSegments` (9371)
- `collectParagraphMeasureTokens` (9449)
- `collectParagraphTabLeaderMeasureSegments` (9495)
- `estimateTabLeaderWrappedLineCountForParagraph` (9548)

**C. 折行计数缓存** (9632-9985):
- `wrappedLineCountByParagraph` (9632)
- `cachedWrappedLineCountForParagraph` (9637)
- `rememberWrappedLineCountForParagraph` (9644)
- `estimateWrappedLineCountForParagraph` (9658)
- `paragraphAvailableTextWidthPx` (9936)
- `paragraphLineCountWithinWidth` (9982, export)

**D. Pretext 块高度** (10056-10164):
- `estimateWrappedFloatingImageFootprintPx` (10056)
- `pretextLayoutContentBottomPx` (10096)
- `topAndBottomExclusionCanOverflowParagraphBox` (10107)
- `wrappedPretextParagraphBlockHeightPx` (10125, export)
- `resolvePretextLineRangeContentHeightPx` (10136, export)
- `resolveMaxPretextLineRangeEndIndexThatFits` (10164, export)

**E. 行高计算** (10204-10391):
- `resolveAutoLineSpacingMultiple` (10204)
- `autoLineHeightScaleForMultiple` (10215)
- `calibrateAutoLineSpacingMultiple` (10244)
- `paragraphDocGridSnapState` (10263)
- `resolveParagraphDocGridLinePitchPx` (10298)
- `estimateParagraphLineHeightPx` (10316, export)
- `estimateParagraphHeightPx` (10391)

**依赖**: `cache-utils.ts`, `constants.ts`, `pretext-integration.ts` (resolveParagraphDualWrappedTextLayout), `section-layout.ts` (twipsToPixels, TWIPS_PER_PIXEL), `@extend-ai/react-docx-doc-model`

---

### 12. `table-height-estimate.ts` (~1,050 行)

**功能领域**: 表格单元格内容高度估算、行高度估算、表格切片

**包含定义**:

**A. 表格单元格内容** (10541-10908):
- `paragraphHasExplicitBeforeSpacing` (10541)
- `paragraphHasExplicitSpacing` (10555)
- `wordLikeTableCellParagraph` (10572)
- `suppressFirstTableCellParagraphTopSpacing` (10600)
- `estimateTableCellContentHeightPx` (10612)

**B. 表格行高度估算** (10729-11020):
- `rowAllowsPageSplit` (10729)
- `rowHasDeepFlowContent` (10733)
- `rowHasNestedTableContent` (10752)
- `capSplitFriendlyTableRowEstimatePx` (10758)
- `tableStyleIdFromSourceXml` (10802)
- `tableHasVisibleBorders` (10813)
- `tableContainsParagraphsWithoutExplicitSpacing` (10824)
- `tableUsesWordLikeParagraphDefaults` (10837)
- `estimateTableRowHeightsPx` (10850, export)
- `computeTableCellDerivedRowHeightsPx` (10909)
- `resolveTableRowHeightCss` (10997)

**C. 表格切片** (11021-11460):
- `TableCellSliceBoundaryLayout` (11021)
- `uniqueSortedPixelBoundaries` (11026)
- `estimateParagraphBoundaryOffsetsPx` (11044)
- `estimateNestedTableBoundaryOffsetsPx` (11148)
- `estimateTableCellSliceBoundaryLayoutPx` (11176)
- `tableCellSliceBoundaryIsSafe` (11254)
- `resolveTableRowSliceHeightOnSafeBoundaryPx` (11276)

**D. 表格/节点整体高度** (11416-11590):
- `estimateTableHeightPx` (11416)
- `estimateDocNodeHeightPx` (11460)

**依赖**: `line-height-estimate.ts`, `cache-utils.ts`, `constants.ts`, `paragraph-inspection.ts`, `section-layout.ts` (twipsToPixels, TWIPS_PER_PIXEL), `@extend-ai/react-docx-doc-model`

---

### 13. `pagination-plan.ts` (~2,500 行)

**功能领域**: 分页编排、页面断点计算、段落切片、列渲染

**包含定义**:

**A. 页面断点缓存** (8654-8923):
- `docxHardPageBreakStartNodeIndexesByModel` (8654)
- `collectDocxHardPageBreakStartNodeIndexes` (8659)
- `computeDocxHardPageBreakStartNodeIndexes` (8671)
- `docxSectionStartPageBreakNodeIndexesByModel` (8709)
- `collectDocxSectionStartPageBreakNodeIndexes` (8714)
- `computeDocxSectionStartPageBreakNodeIndexes` (8726)
- `buildNextHardBreakStartNodeIndexLookup` (8755)
- `paragraphIsSimpleTrailingSectionTailCandidate` (8772)
- `shouldKeepTrailingSectionTailOnCurrentPage` (8790)

**B. 段落间距/分页** (11481-11590):
- `paragraphBeforeSpacingPx` (11481)
- `paragraphAfterSpacingPx` (11485)
- `paragraphWidowControlEnabled` (11489)
- `paragraphIsOnlyExplicitPageBreak` (11493)
- `resolveParagraphBeforeSpacingPx` (11505)
- `paragraphCanSplitAcrossPages` (11531)
- `collectDocxEstimatedOverflowBreakStartNodeIndexes` (11563)

**C. 文档页面段** (11882-12758):
- `DocumentPageRange` (11882)
- `buildDocumentPageRanges` (11887)
- `collectDocxPageBreakStartNodeIndexes` (11922)
- `TableRowRange` (11945)
- `TableRowSlice` (11950)
- `ParagraphLineRange` (11957)
- `ParagraphSegmentIdentity` (11964)
- `DocumentPageNodeSegment` (11970)
- `paragraphSegmentHasPartialLineRange` (11977)
- `resolveLineRangeWithinVerticalSlice` (11990, export)
- `resolveTableCellParagraphVisualBottomPx` (12041, export)
- `tableCellParagraphFitsFullyWithinSlice` (12052, export)
- `resolveParagraphSegmentClipBleedPx` (12067, export)
- `resolveFallbackParagraphSegmentClipBleedPx` (12093, export)
- `resolveParagraphSegmentNonFlowReservePx` (12137, export)
- `resolveFallbackParagraphSegmentNonFlowReservePx` (12156)
- `paragraphSegmentIdentityMatches` (12179)
- `estimateRenderedPageSegmentHeightPx` (12195)
- `resolveParagraphColumnRenderLineRange` (12308)
- `splitParagraphSegmentForColumnRender` (12362)
- `buildRenderColumnSegmentsForPageSection` (12466, export)

**D. 表格行适配 + 页面构建** (12654-13986):
- `sumEstimatedTableRowHeightsPx` (12654)
- `fitTableRowsWithinHeightPx` (12674)
- `buildDocumentPageNodeSegmentsFromLastRenderedPageBreakHints` (12714, export)
- `buildDocumentPageNodeSegments` (12759, export)
- `mergeTrailingPagesToTargetCount` (13986)

**依赖**: `paragraph-inspection.ts`, `table-height-estimate.ts`, `line-height-estimate.ts`, `header-footer.ts`, `page-measurement.ts`, `constants.ts`, `cache-utils.ts`, `section-layout.ts`, `pagination-breaks.ts`, `page-count-reconciliation.ts`, `@extend-ai/react-docx-doc-model`

---

### 14. `style-to-css.ts` (~1,160 行)

**功能领域**: 样式对象到 CSS 的转换（run、paragraph、table cell、border、numbering marker）

**包含定义**:

| 行号 | 名称 | 类型 |
|------|------|------|
| 14888-14904 | `resolveHighlightColor` | function |
| 14905-15003 | `paragraphLineHeight` | function |
| 15004-15035 | `keepNextPaginationReservePx` | function |
| 15036-15041 | `paragraphBorderToCss` | function |
| 15042-15052 | `paragraphBorderPaddingPx` | function |
| 15053-15066 | `paragraphBorderStrokeWidthPx` | function |
| 15067-15080 | `paragraphBorderInsetPx` | function |
| 15081-15144 | `paragraphExplicitIndentTwips` | function |
| 15145-15324 | `resolveListParagraphIndent` | function |
| 15325-15394 | `resolveNumberingMarkerBoxWidthPx` | function |
| 15395-15432 | `numberingMarkerStyle` | function |
| 15433-15545 | `paragraphBlockStyle` | function |
| 15546-15635 | `tableCellParagraphBlockStyle` | function |
| 15636-15663 | `themedRunColor` | function |
| 15664-15694 | `cssFontFamily` | function |
| 15695-15783 | `runStyleToCss` | function |
| 15784-15800 | `linkStyleToCss` | function |
| 15801-15822 | `mergeTextDecorations` | function |
| 15823-15845 | `trackedInlineStyle` | function |
| 15846-15863 | `trackedDeletedStyle` | function |

**依赖**: `constants.ts` (HIGHLIGHT_TO_CSS), `section-layout.ts` (twipsToPixels, TWIPS_PER_PIXEL), `@extend-ai/react-docx-doc-model`

---

### 15. `xml-parsing.ts` (~1,120 行)

**功能领域**: DOCX XML 解析（run style、tracked change tokens、alignment、balanced tags）

**包含定义**:

**A. XML 文本/标签解析** (16061-16218):
- `decodeXmlText` (16061)
- `XmlBalancedTagRange` (16084)
- `extractBalancedTagRanges` (16091)
- `RevisionTagRange` (16147)
- `trackedChangeKindFromTagName` (16157)
- `normalizeTrackedChangeSnippet` (16177)
- `formatTrackedChangeDate` (16188)
- `stripTextBoxContentFromRunXml` (16211)
- `parseTrackedRunTokens` (16218)

**B. Run XML 样式解析** (16263-16364):
- `xmlBooleanFlag` (16263)
- `xmlColorValue` (16276)
- `parseRunStyleFromRunXml` (16297)
- `balancedTagXmlBlocks` (16365)
- `mergeTextRunStyles` (16371)
- `parseParagraphAlignmentFromXml` (16396)

**C. 图片变换** (17032-17158):
- `parseDrawingImageTransformFromSourceXml` (17032)
- `joinCssTransforms` (17071)
- `resolveImageRenderTransformStyle` (17080)

**依赖**: `constants.ts`, `@extend-ai/react-docx-doc-model`

> 注: 此模块与 `synthetic-textbox.ts` 的原始行号范围有重叠（16414-17158），实为交错排列。拆分时 XML 解析部分进入本模块，SVG/文本框生成部分进入 `synthetic-textbox.ts`。

---

### 16. `synthetic-textbox.ts` (~750 行)

**功能领域**: 合成文本框 SVG 渲染（Word 文本框/形状作为 SVG）

**包含定义**:

| 行号 | 名称 | 类型 |
|------|------|------|
| 16414-16421 | `escapeSvgText` | function |
| 16422-16425 | `svgDataUri` | function |
| 16426-16430 | `SyntheticTextBoxSegment` | interface |
| 16431-16436 | `SyntheticTextBoxParagraph` | interface |
| 16437-16446 | `SyntheticTextBoxFrameStyle` | interface |
| 16447-16454 | `emuToPixels` | function |
| 16455-16504 | `resolveSyntheticTextBoxFieldText` | function |
| 16505-16703 | `syntheticTextBoxParagraphsFromRunXml` | function |
| 16704-16753 | `syntheticTextBoxFrameStyleFromRunXml` | function |
| 16754-16773 | `syntheticTextBoxTextValue` | function |
| 16774-16810 | `resolveSyntheticTextBoxParagraphs` | function |
| 16811-16826 | `syntheticTextBoxForegroundZIndex` | function |
| 16827-17002 | `syntheticTextBoxSvg` | function |
| 17003-17031 | `syntheticTextBoxContainsPictureLayer` | function |
| 17158-17169 | `syntheticTextBoxActsAsTopAndBottomMasthead` | function |

**依赖**: `xml-parsing.ts` (parseRunStyleFromRunXml, xmlAttribute), `constants.ts`, `@extend-ai/react-docx-doc-model`

---

### 17. `tracked-changes.ts` (~1,100 行)

**功能领域**: 修订标记(tracked changes)和批注(comments)的解析、收集、排版

**包含定义**:

**A. 修订/批注标记接口** (382-418):
- `ParagraphTrackedInlineChange` (382)
- `ParagraphTrackedDeletionSegment` (390)
- `ParagraphTrackedMarkup` (396)
- `ParagraphCommentMarkup` (407)
- `paragraphTrackedMarkupBySourceXml` (411-414)
- `paragraphCommentMarkupBySourceXml` (415-418)

**B. 修订/批注解析** (17170-17561):
- `summarizeChangeFeatures` (17170)
- `summarizeRunFormattingChange` (17183)
- `summarizeParagraphFormattingChange` (17215)
- `summarizeTableFormattingChange` (17248)
- `resolveParagraphTrackedMarkup` (17290)
- `resolveParagraphCommentMarkup` (17561)

**C. 模型级修订收集** (22675-22770):
- `collectTablePropertyTrackedChanges` (22675)
- `collectTrackedChangesFromModel` (22770)

**D. 批注收集** (22866-22940):
- `decodeCommentRangeText` (22866)
- `resolveCommentAnchorText` (22880)
- `collectCommentsFromModel` (22908)

**E. 批注/修订卡片样式** (23005-23052):
- `hexColorWithAlpha` (23005)
- `commentAccentColor` (23017)
- `commentHighlightStyle` (23031)
- `estimateCommentCardHeight` (23048)

**F. 修订 gutter 排版** (23054-23442):
- `trackedChangeKindLabel` (23054)
- `trackedChangeAccentColor` (23073)
- `trackedChangeUsesGutterBalloon` (23111)
- `gutterAnnotationSortTuple` (23117)
- `trackedChangeBelongsToPageSegments` (23132)
- `resolveGutterAnnotationPageIndex` (23158)
- `findTrackedChangeAnchorElementInPage` (23180)
- `elementRectWithinContainer` (23204)
- `TrackedChangeAnchorPoint` (23244)
- `DocxGutterAnnotation` (23253)
- `findFirstElementWithSpaceSeparatedDataValue` (23260)
- `findGutterAnnotationScopeElementInPage` (23289)
- `findGutterAnnotationDataAnchorInPage` (23296)
- `findGutterAnnotationAnchorElementInPage` (23324)
- `PositionedGutterAnnotation` (23368)
- `estimateTrackedChangeCardHeight` (23377)
- `assignGutterConnectorLanes` (23385)
- `layoutTrackedChangesForPage` (23442)

**依赖**: `xml-parsing.ts`, `constants.ts`, `@extend-ai/react-docx-doc-model`

---

### 18. `field-helpers.ts` (~700 行)

**功能领域**: 域代码解析（PAGE/NUMPAGES/StyleRef）、目录检测、tab leader 渲染

**包含定义**:

**A. TOC/域字段检测** (15864-16049):
- `TABLE_OF_CONTENTS_STYLE_ID` (15864)
- `isTableOfContentsStyle` (15866)
- `tableOfContentsLevel` (15873)
- `isTableOfContentsParagraph` (15892)
- `paragraphUsesTabLeaders` (15899)
- `paragraphLeadingTabStop` (15924)
- `tableOfContentsLeadingLeftTabStopPx` (15947)
- `paragraphContainsTabCharacter` (15957)
- `paragraphTabCharacterCount` (15969)
- `ParagraphAnchoredTabLayout` (15985)
- `paragraphAnchoredTabLayout` (15987)
- `paragraphFirstTabStopPx` (16023)
- `paragraphUsesCenterTabLayout` (16037)
- `paragraphUsesRightTabLayout` (16041)
- `paragraphUsesCenterRightTabLayout` (16045)

**B. 域字段解析** (16049-16188):
- `PageFieldKind` (16049)
- `PageFieldValueToken` (16051)
- `StyleRefFieldValueToken` (16056)

**C. 域字段序列解析** (17653-18086):
- `instructionTextToPageFieldKind` (17653)
- `instructionTextToStyleRefTarget` (17674)
- `paragraphPageFieldSequence` (17687)
- `pageFieldValueSequenceBySourceXml` (17714)
- `paragraphPageFieldValueSequence` (17719)
- `styleRefFieldValueSequenceBySourceXml` (17845)
- `paragraphStyleRefFieldValueSequence` (17850)
- `normalizeStyleRefTarget` (17976)
- `paragraphStyledRunText` (17980)
- `tabLeaderStyle` (18025)
- `noteMarkerLabel` (18056)
- `normalizeDateInputValue` (18086)

**依赖**: `constants.ts`, `xml-parsing.ts` (decodeXmlText), `style-to-css.ts` (themedRunColor), `@extend-ai/react-docx-doc-model`

---

### 19. `numbering.ts` (~600 行)

**功能领域**: 编号/列表符号生成、列表类型检测、列表缩进

**包含定义**:

**A. 编号格式化** (19581-19907):
- `numberToRoman` (19581)
- `numberToLetters` (19615)
- `ParagraphNumberingLabel` (19631)
- `formatNumberingCounter` (19642)
- `formatPageFieldValue` (19674)
- `findNumberingLevelDefinition` (19682)
- `numberingLevelHasVisibleMarker` (19705)
- `numberingLevelIsBulletLike` (19713)
- `numberingAbstractLevelsForNumId` (19724)
- `effectiveNumberingNumIdForParagraph` (19741)
- `numberingStartValue` (19842)
- `numberingSuffix` (19861)
- `LEGACY_BULLET_GLYPH_FALLBACKS` (19871)
- `normalizeLegacyBulletGlyphs` (19878)
- `buildParagraphNumberingLabels` (19909, export)

**B. 列表模式检测** (20086-20188):
- `isUnorderedListText` (20086)
- `isOrderedListText` (20090)
- `isBulletLikeNumberingText` (20094)
- `paragraphHasNumbering` (20108)
- `paragraphListType` (20114)
- `paragraphIsList` (20155)
- `stripListPrefix` (20167)
- `listPrefixLength` (20171)
- `nextOrderedListItemText` (20175)
- `textWithListType` (20188)

**C. 模型克隆 + 列表缩进** (20193-20471):
- `cloneParagraphBorderStyle` (20193)
- `cloneParagraphBorderSet` (20199)
- `cloneParagraphStyle` (20216)
- `splitParagraphStyleWithDefaultSpacing` (20236)
- `cloneTableBoxSpacing` (20267)
- `cloneTableBorderStyle` (20282)
- `cloneTableBorderSet` (20296)
- `cloneTableCellStyle` (20315)
- `cloneTableRowStyle` (20329)
- `createEmptyParagraphFromTemplate` (20339)
- `createEmptyTableCellFromTemplate` (20357)
- `resolveMaxNumberingLevel` (20379)
- `shiftListIndent` (20415)
- `ensurePrefixListIndent` (20438)
- `clearAutoPrefixListIndent` (20471)

**依赖**: `section-layout.ts` (TWIPS_PER_PIXEL), `@extend-ai/react-docx-doc-model` (DocModel, NumberingDefinitionSet, etc.)

---

### 20. `paragraph-render.tsx` (~1,455 行)

**功能领域**: 段落 runs 渲染为 React JSX（**需重写为 Vue**）

**包含定义**:

| 行号 | 名称 | 类型 |
|------|------|------|
| 18109-18124 | `ParagraphRunRenderOptions` | interface |
| 18126-19581 | `renderParagraphRuns` | function |

**⚠️ 重写为 Vue 的注意事项**:
- 该函数生成大量 JSX（`<span>`, `<div>`, `<a>`, `<img>`, `<sup>` 等 React 元素）
- 包含 inline tab leader 布局、域字段渲染（PAGE/NUMPAGES）、StyleRef 字段、超链接、脚注/尾注标记、复选框、下拉框、图片渲染、浮动图片、合成文本框 SVG、修订标记和批注高亮
- React 特有模式: `useTabLeaderLayout` + `React.Fragment`, `key` prop 管理, `onClick`/`onMouseDown` 事件绑定
- Vue 迁移策略: 将 `renderParagraphRuns` 改为一个渲染函数，返回 VNode 数组；或将输出改为一个数据结构由 Vue 模板消费

**依赖**: `field-helpers.ts`, `numbering.ts`, `style-to-css.ts`, `synthetic-textbox.ts`, `tracked-changes.ts`, `paragraph-inspection.ts`, `pretext-integration.ts`, `constants.ts`, `@extend-ai/react-docx-doc-model`

---

### 21. `table-utils.ts` (~1,405 行)

**功能领域**: 表格列宽计算、表格边框 CSS、边框预设应用

**包含定义**:

**A. 表格基本工具** (20511-21086):
- `tableCellText` (20511)
- `tableColumnCount` (20515)
- `resolveFloatingTableSide` (20531)
- `tableWrapperStyle` (20554)
- `EmbeddedTableRuntimeKeySegment` (20586)
- `EmbeddedTableRuntimeKeyLocation` (20592)
- `parseEmbeddedTableRuntimeKey` (20600)
- `columnWidthsByTable` (20666)
- `columnWidthsFromTableDefinition` (20671)
- `computeColumnWidthsFromTableDefinition` (20687)
- `gridConflictsWithRowWidths` (20720)
- `deriveColumnWidthsFromTableRows` (20763)
- `normalizeColumnWidthsTwips` (20823)
- `normalizeColumnWidthsPx` (20843)
- `defaultColumnWidthsPx` (20864)
- `clampTableWidthPx` (20877)
- `fitColumnWidthsToWidth` (20892)
- `rowGridSpanCount` (20963)
- `resolveFittedTableColumnWidths` (20976)

**B. 表格间距/边框** (21086-21350):
- `resolveTableSpacingPaddingPx` (21086)
- `tableSpacingPaddingStyle` (21108)
- `mergeTableSpacing` (21121)
- `TableBorderSide` (21137)
- `normalizeBorderType` (21139)
- `tableBorderToCss` (21148)
- `tableBorderStrokeWidthPx` (21181)
- `borderTypeVisible` (21195)
- `paragraphBorderVisible` (21202)
- `tableBorderVisible` (21208)
- `tableBorderSetHasVisibleEdges` (21212)
- `tableUsesSeparateBorderModel` (21223)
- `resolveTableSeparateBorderSpacingPx` (21245)
- `tableElementBorderStyle` (21264)
- `resolvePreferredParagraphBorder` (21280)
- `resolvePreferredTableBorder` (21297)
- `toolbarParagraphBorderStyle` (21316)
- `toolbarTableBorderStyle` (21332)
- `nilParagraphBorderStyle` (21345)
- `nilTableBorderStyle` (21349)

**C. 边框预设** (21353-21870):
- `paragraphBorderPresetState` (21353)
- `paragraphRangePresetActive` (21381)
- `applyParagraphBorderPresetForRangeEntry` (21447)
- `tableBorderPresetState` (21510)
- `applyParagraphBorderPreset` (21553)
- `applyTableBorderPreset` (21616)

**D. 单元格边框渲染** (21695-21870):
- `resolveTableBorder` (21695)
- `resolveTableCellBorderCss` (21728)
- `resolveCollapsedTableHorizontalOuterBleedPx` (21794)
- `resolveTableCellDiagonalOverlayCss` (21870)

**依赖**: `section-layout.ts` (TWIPS_PER_PIXEL, twipsToPixels), `@extend-ai/react-docx-doc-model`

---

### 22. `text-mutation.ts` (~760 行)

**功能领域**: 文本样式/链接的克隆与范围修改、图片拖放目标解析

**包含定义**:

**A. 文本/表单字段克隆** (21915-22090):
- `cloneTextStyle` (21915)
- `cloneFormFieldWidget` (21921)
- `mergeFormFieldWidgetPatch` (21955)
- `cloneFormFieldRun` (22005)
- `textStylesEqual` (22029)
- `mergeAdjacentTextRuns` (22036)
- `paragraphHasOnlyTextRuns` (22083)
- `cloneTextRunWithMetadata` (22089)

**B. 文本运行拆分** (22099-22185):
- `splitTextRunsAtOffset` (22099)
- `firstTextStyleAtOffset` (22147)

**C. 链接操作** (22185-22440):
- `linkAtOffset` (22185)
- `uniformLinkInRange` (22223)
- `linkRangeAtOffset` (22263)
- `mutateParagraphTextStyleInRange` (22344)
- `mutateParagraphLinkInRange` (22440)

**D. 图片拖放目标解析** (22534-22670):
- `toBase64` (22534)
- `isParagraphSelected` (22553)
- `isCellSelected` (22560)
- `paragraphLocationKey` (22574)
- `imageLocationKey` (22582)
- `dropTargetKey` (22586)
- `parseImageDropTargetFromDataset` (22590)
- `firstTableCellAnchorLocation` (22646)

**依赖**: `@extend-ai/react-docx-doc-model`, `@extend-ai/react-docx-editor-ops`

---

### 23. `selection-helpers.ts` (~1,210 行)

**功能领域**: 选区/光标位置操作、文本范围标准化、模型位置定位

**包含定义**:

**A. 段落位置工具** (23538-23895):
- `sameParagraphLocation` (23538)
- `firstParagraphLocationInTable` (23558)
- `lastParagraphLocationInTable` (23597)
- `nodeIndexFromParagraphLocation` (23636)
- `adjustLocationAfterRemovedNodeIndexes` (23642)
- `tableCoverageBoundaries` (23683)
- `fullyCoveredTableNodeIndexesForRange` (23716)
- `firstParagraphLocationFromNode` (23755)
- `lastParagraphLocationInNode` (23774)
- `nextParagraphLocation` (23793)
- `firstParagraphLocationInDocument` (23882)
- `lastParagraphLocationInDocument` (23895)

**B. 范围变更** (23946-24145):
- `paragraphRangeForMutate` (23946)
- `normalizeRangeBoundaryParagraphOffset` (23970)
- `rangeCoversEntireDocument` (23978)
- `resolveRangeBoundaryOffsetsForParagraph` (24014)
- `cloneTextRangeLocation` (24045)
- `cloneTextRange` (24064)
- `cloneEditorSelection` (24081)
- `sameEditorSelection` (24099)
- `sameTextRangeBoundary` (24125)
- `sameTextRange` (24132)

**C. DOM 选区恢复** (24148-24261):
- `shouldReissueDomSelectionRestore` (24148, export)
- `shouldSyncActiveRangeOnKeyUp` (24179)
- `isCollapsedSelectionAtElementStart` (24208)
- `isSuspiciousCollapsedSelectionAtElementStart` (24247)

**D. 位置解析** (24261-24314):
- `imageLocationToParagraphLocation` (24261)
- `getParagraphAtLocation` (24280)
- `collectFormFieldsFromModel` (24314)

**E. 位置标准化** (24364-24666):
- `ParagraphTextUpdateOptions` (24364)
- `selectionFallbackParagraphLocation` (24366)
- `selectionFromTextRangeLocation` (24385)
- `paragraphLocationFromTextRangeLocation` (24403)
- `resolveSelectedParagraphLocation` (24422)
- `normalizeParagraphLocationForModel` (24434)
- `normalizeTextRangeForModel` (24532)
- `normalizeSelectionForModel` (24577)
- `normalizeEditorCursorStateForModel` (24595)
- `updateParagraphTextAtLocation` (24620)

**依赖**: `paragraph-inspection.ts`, `text-mutation.ts`, `@extend-ai/react-docx-doc-model`, `@extend-ai/react-docx-editor-ops`

---

### 24. `section-manipulation.ts` (~310 行)

**功能领域**: 节(section)中段落/图片的定位和修改

**包含定义**:

| 行号 | 名称 | 类型 |
|------|------|------|
| 24641-24653 | `sectionParagraphLocationKey` | function |
| 24654-24665 | `sectionImageLocationKey` | function |
| 24666-24674 | `tableCellParagraphDraftKey` | function |
| 24675-24723 | `parseSectionParagraphLocationKey` | function |
| 24724-24753 | `paragraphTextFromSectionLocation` | function |
| 24754-24793 | `sectionParagraphFromLocation` | function |
| 24794-24878 | `updateSectionParagraphTextAtLocation` | function |
| 24879-24952 | `updateSectionImageFloatingAtLocation` | function |

**依赖**: `selection-helpers.ts`, `paragraph-inspection.ts`, `text-mutation.ts`, `cache-utils.ts`, `@extend-ai/react-docx-doc-model`, `@extend-ai/react-docx-editor-ops`

---

## 依赖关系图

```
                           ┌─────────────────┐
                           │  editor-types.ts │  (纯类型, 无运行时依赖)
                           └─────────────────┘

  ┌──────────────┐     ┌──────────────────┐
  │ constants.ts │◄────│  cache-utils.ts   │
  └──────┬───────┘     └────────┬─────────┘
         │                      │
    ┌────┴──────────────────────┴─────────────┐
    │                                          │
    ▼                                          ▼
┌──────────────┐  ┌───────────────────┐  ┌──────────────┐
│ ooxml-helpers│  │ header-footer.ts  │  │ default-model│
└──────────────┘  └────────┬──────────┘  └──────────────┘
                            │
                     ┌──────┴──────┐
                     ▼             ▼
           ┌──────────────┐  ┌──────────────────┐
           │page-measure..│  │xml-parsing.ts    │
           └──────┬───────┘  └────────┬─────────┘
                  │                   │
                  │            ┌──────┴──────────────┐
                  │            ▼                     ▼
                  │   ┌──────────────────┐  ┌──────────────────┐
                  │   │synthetic-textbox │  │tracked-changes   │
                  │   └──────────────────┘  └────────┬─────────┘
                  │                                  │
    ┌─────────────┴──────────────────────────────────┤
    │                                                │
    ▼                                                ▼
┌──────────────────────────┐              ┌──────────────────┐
│ paragraph-inspection.ts  │              │ field-helpers.ts │
└────────────┬─────────────┘              └────────┬─────────┘
             │                                     │
    ┌────────┴──────────┐                          │
    ▼                   ▼                          │
┌───────────┐  ┌──────────────────┐               │
│ drop-cap  │  │pretext-integration│               │
└─────┬─────┘  └────────┬─────────┘               │
      │                 │                          │
      └────────┬────────┘                          │
               ▼                                   │
   ┌──────────────────────┐                        │
   │ line-height-estimate │                        │
   └──────────┬───────────┘                        │
              │                                    │
    ┌─────────┴────────────┐                       │
    ▼                      ▼                       │
┌──────────────────┐ ┌──────────────┐              │
│table-height-est..│ │ style-to-css │◄─────────────┘
└────────┬─────────┘ └──────┬───────┘
         │                  │
         └──────┬───────────┘
                ▼
     ┌──────────────────┐
     │ pagination-plan  │
     └──────────────────┘

  ┌──────────────────┐     ┌──────────────────┐     ┌──────────────────┐
  │ table-utils.ts   │     │ text-mutation.ts │     │ numbering.ts     │
  └──────────────────┘     └────────┬─────────┘     └──────────────────┘
                                    │
                             ┌──────┴──────┐
                             ▼             ▼
                   ┌──────────────────┐ ┌────────────────────┐
                   │selection-helpers │ │section-manipulation│
                   └──────────────────┘ └────────────────────┘

  ┌──────────────────────┐
  │ paragraph-render.tsx │ (JSX → Vue 重写, 依赖最广)
  └──────────────────────┘
    ↑ 依赖: field-helpers, numbering, style-to-css, synthetic-textbox,
    │        tracked-changes, paragraph-inspection, pretext-integration,
    │        drop-cap, constants
```

---

## 拆分执行建议

### 拆分顺序（按依赖关系，从底层到上层）

1. `constants.ts` — 零内部依赖
2. `cache-utils.ts` — 仅依赖 constants
3. `editor-types.ts` — 仅依赖外部 doc-model 类型
4. `ooxml-helpers.ts` — 仅依赖外部模块
5. `header-footer.ts` — 依赖 constants + section-layout(外部)
6. `page-measurement.ts` — 依赖 header-footer
7. `default-model.ts` — 仅依赖外部 doc-model
8. `xml-parsing.ts` — 依赖 constants
9. `synthetic-textbox.ts` — 依赖 xml-parsing
10. `tracked-changes.ts` — 依赖 xml-parsing + constants
11. `field-helpers.ts` — 依赖 xml-parsing + constants + style-to-css(需前置)
12. `paragraph-inspection.ts` — 依赖 constants + cache-utils
13. `drop-cap.ts` — 依赖 paragraph-inspection + cache-utils
14. `pretext-integration.ts` — 依赖 constants + cache-utils + pretext-layout(外部)
15. `line-height-estimate.ts` — 依赖 pretext-integration + cache-utils + constants
16. `style-to-css.ts` — 依赖 constants
17. `table-height-estimate.ts` — 依赖 line-height-estimate + paragraph-inspection
18. `pagination-plan.ts` — 依赖 paragraph-inspection + table-height-estimate + header-footer + page-measurement
19. `table-utils.ts` — 仅依赖外部 section-layout
20. `numbering.ts` — 仅依赖外部 doc-model
21. `text-mutation.ts` — 仅依赖外部 doc-model + editor-ops
22. `selection-helpers.ts` — 依赖 paragraph-inspection + text-mutation
23. `section-manipulation.ts` — 依赖 selection-helpers + paragraph-inspection + text-mutation
24. `paragraph-render.tsx` — 依赖几乎所有上层模块（最后拆分）

### 循环依赖风险点

- `style-to-css.ts` ↔ `field-helpers.ts`: field-helpers 中 `tabLeaderStyle` 使用 `themedRunColor`（来自 style-to-css）。需将 `themedRunColor` 提前到 field-helpers 可用的位置，或将 tabLeaderStyle 移入 style-to-css。
- `tracked-changes.ts` ↔ `paragraph-inspection.ts`: tracked-changes 中的 `trackedChangeBelongsToPageSegments` 引用段落分页逻辑。保持在 tracked-changes 中，通过运行时参数传递 pageSegments 避免循环。

### 每个模块的拆分步骤

1. 从 `editor.tsx` 中提取函数体到新文件
2. 在新文件中添加必要的 import
3. 从 `editor.tsx` 中移除已提取代码，替换为 `import { ... } from "./new-module"`
4. 运行 TypeScript 编译器检查类型错误
5. 运行现有测试确保无回归

### 风险标注

| 风险等级 | 模块 | 原因 |
|----------|------|------|
| 🔴 高 | `paragraph-render.tsx` | 1,455 行 JSX→Vue 重写，涉及所有渲染逻辑 |
| 🟡 中 | `pagination-plan.ts` | 2,500 行综合编排逻辑，依赖 6 个模块 |
| 🟡 中 | `paragraph-inspection.ts` | 2,350 行，覆盖 5 个子领域，拆得过散会降低可读性 |
| 🟡 中 | `pretext-integration.ts` | 1,800 行，与外部 pretext-layout.ts 紧耦合 |
| 🟢 低 | `constants.ts` | 纯常量，无运行时逻辑 |
| 🟢 低 | `editor-types.ts` | 纯类型，无运行时逻辑 |
| 🟢 低 | `default-model.ts` | 100 行，简单的模型工厂 |

---

## 文件行数统计

| 模块 | 行数 | 占总比 |
|------|------|--------|
| `pagination-plan.ts` | ~2,500 | 10.0% |
| `paragraph-inspection.ts` | ~2,350 | 9.4% |
| `pretext-integration.ts` | ~1,800 | 7.2% |
| `editor-types.ts` | ~1,520 | 6.1% |
| `line-height-estimate.ts` | ~1,470 | 5.9% |
| `paragraph-render.tsx` | ~1,455 | 5.8% |
| `table-utils.ts` | ~1,405 | 5.6% |
| `selection-helpers.ts` | ~1,210 | 4.9% |
| `style-to-css.ts` | ~1,160 | 4.7% |
| `xml-parsing.ts` | ~1,120 | 4.5% |
| `tracked-changes.ts` | ~1,100 | 4.4% |
| `table-height-estimate.ts` | ~1,050 | 4.2% |
| `text-mutation.ts` | ~760 | 3.0% |
| `synthetic-textbox.ts` | ~750 | 3.0% |
| `page-measurement.ts` | ~740 | 3.0% |
| `field-helpers.ts` | ~700 | 2.8% |
| `header-footer.ts` | ~625 | 2.5% |
| `numbering.ts` | ~600 | 2.4% |
| `constants.ts` | ~480 | 1.9% |
| `drop-cap.ts` | ~450 | 1.8% |
| `section-manipulation.ts` | ~310 | 1.2% |
| `ooxml-helpers.ts` | ~205 | 0.8% |
| `cache-utils.ts` | ~160 | 0.6% |
| `default-model.ts` | ~100 | 0.4% |
| **总计** | **~24,900** | **~100%** |

---

## `renderParagraphRuns` 和 `renderStaticHtml` 特别标注

### `renderParagraphRuns` (行 18126-19581, ~1,455 行)

- **现状**: 纯 React JSX 函数，接收 `ParagraphNode` 和渲染选项，返回 `React.ReactNode`
- **包含的子元素类型**: `<span>`, `<div>`, `<a>`, `<img>`, `<sup>`, `<sub>`, React Fragment
- **Vue 重写策略**:
  1. **方案 A（推荐）**: 保持为纯渲染函数，使用 Vue 3 的 `h()` 函数（createVNode）替代 JSX，输出 VNode 数组
  2. **方案 B**: 将函数拆分为小的 Vue 组件（`ParagraphRun`, `TabLeader`, `NoteMarker` 等），用模板语法渲染
  3. **方案 C**: 生成一个中间数据结构（树形 JSON），由统一的 Vue 模板消费
- **依赖的模块**: field-helpers, numbering, style-to-css, synthetic-textbox, tracked-changes, paragraph-inspection, pretext-integration, drop-cap
- **React 特有 API 需替换**:
  - `React.Fragment` → Vue 的 `Fragment`
  - `key` prop → Vue 的 `key`
  - `onClick` / `onMouseDown` → Vue 的 `@click` / `@mousedown`（或 h() 中的 `onClick` / `onMouseDown`）
  - `dangerouslySetInnerHTML` → Vue 的 `innerHTML` prop
  - `className` → Vue 的 `class`

### `renderStaticHtml` (行 4902-4904, ~3 行)

- **现状**: `renderToStaticMarkup(<>{node}</>)`
- **功能**: 将 React 节点渲染为静态 HTML 字符串
- **Vue 替代**: 使用 `@vue/server-renderer` 的 `renderToString`，或浏览器端使用 `innerHTML` 赋值后读取

---

## 残留不确定性

1. **内部交叉引用**: 部分函数在 24 个模块之间存在跨模块调用，实际提取时可能需要额外导出当前为局部（非 export）的函数
2. **性能测量函数** (226-265): `markDocxImportPerformance` / `measureDocxImportPerformance` / `createDocxImportPerformanceTraceName` 目前归入 `constants.ts`，但这些是函数而非常量。可单独拆为 `performance.ts`（~40 行）
3. **letterhead 函数** (617-832): 目前归入 `paragraph-inspection.ts`，但它们是一个独立领域（信头/letterhead 布局）。可单独拆为 `letterhead.ts`（~215 行）
4. **zoom/scroll 函数** (833-891): `clampNumber`, `hasVerticalScrollOverflow`, `nearestScrollableAncestor`, `resolveEffectiveZoomScale`, `normalizePageVirtualizationZoomScale` 是可复用的通用工具。可单拆为 `zoom-utils.ts`（~60 行）或保留在 `paragraph-inspection.ts`
5. **DOM 操作函数** (981-1230): `scheduleDomWrite`, `normalizeMeasuredTableRowHeightPx`, `reconcileMeasuredTableRowHeightsForImportPagination`, `resolveTableMeasuredRowHeightsForPagination`, `placeCaretInsideElementDom`, `selectionOffsetsWithinElementDom`, `setSelectionWithinElementByTextOffsetsDom` 是 DOM 操作。可单拆为 `dom-helpers.ts`（~250 行）或保留在 `table-height-estimate.ts`
