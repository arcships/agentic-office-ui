// Shared utilities and constants
export {
  REL_NS,
  PKG_REL_NS,
  SPREADSHEET_NS,
  CONTENT_TYPES_NS,
  DRAWING_REL_TYPE,
  VML_DRAWING_REL_TYPE,
  CTRL_PROP_REL_TYPE,
  IMAGE_REL_TYPE,
  HYPERLINK_REL_TYPE,
  DRAWING_CONTENT_TYPE,
  EMU_PER_PIXEL,
  type DukeWorksheet,
  type ArchiveEntries,
  type ContentTypesState,
  type RelationshipRecord,
  type DrawingRectEmu,
  type GroupTransform,
  type XlsxImageAttachment,
  type WorkbookImageOrigin,
  type WorkbookImageSheetOrigin,
  // Archive path utilities
  normalizeArchivePath,
  joinArchivePath,
  dirname,
  resolveArchiveTarget,
  relativeArchivePath,
  relsPathForDocument,
  cloneBytes,
  // XML utilities
  parseXml,
  serializeXml,
  readArchiveText,
  isElementNode,
  getLocalElements,
  getChildElements,
  getFirstChild,
  getFirstDescendant,
  getRelationshipId,
  getEmbeddedRelationshipId,
  setChildText,
  updateMarkerElement,
  // Content types & relationships
  parseContentTypes,
  resolveContentType,
  parseRelationships,
  // A1 reference parsing
  parseColumnReference,
  parseA1CellReference,
  parseA1RangeReference,
  // Color utilities
  normalizeHexColor,
  // Image parsing
  createImageAnchorNodes,
  createImageSource,
  parseMarker,
  parseAnchor,
  anchorToRect,
  parseTransformRect,
  applyGroupTransform,
  emuToPixels,
  pixelsToEmu,
  rectToAbsoluteAnchor,
  anchorFromNodeOrFallback,
} from "./image-parser";

// Drawing and shape parsing
export {
  getHyperlinkTarget,
  resolveFillColor,
  resolvePreferredTextTypeface,
  parseCustomGeometryPath,
  parseShapeGeometryAdjustments,
  parseGroupTransform,
  parseAnchorContents,
  parseDrawingObjects,
  normalizeControlLabel,
  flattenShapeText,
  parseSheetFormControls,
  enrichFormControlsWithHiddenShapes,
} from "./drawing-parser";

// Color and theme
export {
  type ThemeState,
  buildThemePalette,
  parseWorkbookTheme,
  parseSpreadsheetColor,
  parseSpreadsheetBorder,
  hasEnabledSpreadsheetFlag,
  parseSpreadsheetFont,
  parseSpreadsheetFill,
  parseSpreadsheetAlignment,
  parseWorkbookStyles,
} from "./theme-palette";

// Column / row sizing and coordinate conversion
export {
  MIN_COL_WIDTH_PX,
  MIN_ROW_HEIGHT_PX,
  resolveDeviceGridlineThicknessPx,
  measureColumnCharacterWidthPx,
  sheetColumnWidthToPixels,
  resolveWorksheetDefaultColumnWidthPixels,
  resolveWorksheetDefaultRowHeightPixels,
  resolveWorksheetHiddenRows,
  resolveWorksheetHiddenCols,
  resolveWorksheetMergeMetadata,
  resolveSheetColumnWidthPx,
  resolveSheetRowHeightPx,
  sumSheetColumnWidthsEmu,
  sumSheetRowHeightsEmu,
  anchorToAbsoluteRect,
  pxToSheetColumnWidth,
  resolveSheetColumnWidthPixels,
  resolveSheetRowHeightPixels,
  resolveRenderedSheetAxisPixels,
  resolveContentSheetAxisPixels,
  rectToImageAnchor,
  resizeImageRect,
} from "./column-width";

// Sheet state and workbook structure
export {
  type WorkbookSheetState,
  type WorkbookTableMetadata,
  type WorkbookImageAssets,
  type WorkbookStructureAssets,
  type WorkbookChartStyleAssets,
  parseSheetState,
  parseWorkbookSheets,
  parseWorkbookStructureAssetsFromArchive,
  normalizeWorkbookTableMetadata,
} from "./grid-render";

// Image export / merge
export {
  revokeWorkbookImageAssets,
  parseWorkbookStructureAssets,
  parseWorkbookChartStyleAssets,
  parseWorkbookImageAssets,
  updateWorkbookImageAnchor,
  mergeWorkbookImageAssets,
} from "./image-export";
export { validateXlsxImageAssets } from "./image-budget";
