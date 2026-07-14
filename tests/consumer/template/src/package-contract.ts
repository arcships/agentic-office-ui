import {
  bundledDocxWasmUrl,
  createDocxPageReferenceDraft,
  createDocxRegionReferenceDraft,
  createDocxTextReferenceDraft,
  createDocxRuntime,
  describeDocxReference,
  resolveDocxReference,
  setWasmSource as setLegacyDocxWasmSource,
  type DocModel,
  type DocxOfficeReferenceDraft,
  type DocxReferenceContext,
} from "@arcships/docx-core";
import {
  cloneDocModel,
  layoutDocument,
  type DocModel as PureDocModel,
} from "@arcships/docx-core/core";
import {
  DocxImportError,
  createDocxRuntime as createDocxRuntimeFromSubpath,
  type DocxFileLike,
  type DocxImportDiagnostic,
  type DocxImportDiagnosticType,
  type DocxImportErrorCode,
  type DocxImportOptions,
  type DocxImportResult,
  type DocxImportSource,
  type DocxImportWorkerTimings,
  type DocxRuntime,
  type DocxRuntimeConfig,
  type DocxRuntimeLimits,
  type DocxRuntimeLoader,
  type DocxSource,
  type DocxUrlPolicy,
} from "@arcships/docx-core/runtime";
import { bundledDocxWasmUrl as docxWasmFromSubpath } from "@arcships/docx-core/wasm-url";
import {
  bundledXlsxWasmUrl,
  canUseConfiguredWasmSourceInWorker,
  createXlsxCellReferenceDraft,
  createXlsxChartReferenceDraft,
  createXlsxRangeReferenceDraft,
  createXlsxRegionReferenceDraft,
  createXlsxWorksheetReferenceDraft,
  describeXlsxReference,
  getConfiguredWorkerWasmSource,
  getSheetsWasmModule,
  resolveXlsxReference,
  setWasmSource as setLegacyXlsxWasmSource,
  type XlsxOfficeReferenceDraft,
  type XlsxReferenceContext,
  type XlsxReferenceSheet,
  type XlsxViewerController,
} from "@arcships/xlsx-core";
import { columnLabel, rangeToA1, type XlsxChart } from "@arcships/xlsx-core/core";
import {
  XlsxRuntimeError,
  createXlsxRuntime,
  type XlsxRuntime,
  type XlsxRuntimeConfig,
  type XlsxRuntimeDiagnostic,
  type XlsxRuntimeErrorCode,
  type XlsxRuntimeParseOptions,
} from "@arcships/xlsx-core/runtime";
import { bundledXlsxWasmUrl as xlsxWasmFromSubpath } from "@arcships/xlsx-core/wasm-url";
import {
  createPptxObjectReferenceDraft,
  createPptxRegionReferenceDraft,
  createPptxSlideReferenceDraft,
  describePptxReference,
  pptxReferenceKindForObject,
  resolvePptxReference,
  type PptxObjectReferenceOptions,
  type PptxOfficeReferenceDraft,
  type PptxReferenceContext,
} from "@arcships/pptx-core";
import {
  DocxContextMenu,
  DocxDragOverlay,
  DocxEditorViewer,
  DocxFormFieldLayer,
  DocxImageLayer,
  DocxPageBody,
  DocxPageFooter,
  DocxPageHeader,
  DocxPageSurface,
  DocxPageWrapper,
  DocxParagraphHost,
  DocxTableHost,
  DocxThumbnailPanel,
  DocxToolbar,
  DocxTrackedChangeGutter,
  DocxViewer,
  DocxViewerRoot,
  createDocxViewerPageSurfaceRegistry,
  ensureDocxViewerPageSurfaceRegistry,
  notifyDocxViewerPageSurfaceListeners,
  renderParagraphRuns,
  renderStaticHtml,
  subscribeDocxViewerPageSurfaces,
  useDocxViewerThumbnails,
  type DocxViewerThumbnails,
  type UseDocxEditorOptions,
  type UseDocxViewerThumbnailsOptions,
} from "@arcships/vue-docx";
import {
  MemoChartSvg,
  MemoSurfaceChartComposite,
  XlsxChartOverlay,
  XlsxContextMenu,
  XlsxFormulaBar,
  XlsxGrid,
  XlsxImageLayer,
  XlsxRibbon,
  XlsxSelectionOverlay,
  XlsxSheetTabs,
  XlsxToolbar,
  XlsxViewer,
  useXlsxViewerController,
  type ChartLayout,
  type ChartRendererPalette,
  type ChartSvgProps,
  type LegendItem,
  type XlsxDiagnostic,
} from "@arcships/vue-xlsx";
import {
  DEFAULT_PDF_MAX_FILE_SIZE,
  PdfViewer,
  bundledPdfiumWasmUrl,
  createPdfPageReferenceDraft,
  createPdfRegionReferenceDraft,
  createPdfTextReferenceDraft,
  createPdfRenderRuntime,
  describePdfReference,
  normalizePdfReferenceRect,
  resolvePdfReference,
  type PdfDiagnostic,
  type PdfLoadOptions,
  type PdfOfficeReferenceDraft,
  type PdfReferenceContext,
  type PdfRenderDocument,
  type PdfRenderRuntime,
  type PdfRenderRuntimeConfig,
  type PdfSource,
} from "@arcships/vue-pdf";
import {
  FileUpload,
  OfficeObjectOutlineLayer,
  OfficeRegionSelector,
  type FileUploadRejection,
  type FileUploadRejectionCode,
  type OfficeObjectOutline,
  type OfficeObjectOutlineLayerProps,
  type OfficeRegionSelectorProps,
} from "@arcships/vue-ui";
import {
  OfficeInteractionValidationError,
  applyOfficeSelectionKeyboard,
  confirmOfficeCandidate,
  createOfficeCandidateNavigationState,
  createOfficeReferenceId,
  createOfficeSelectionSessionState,
  parseOfficeObjectReference,
  parseOfficeReferenceConfirmEvent,
  reduceOfficeCandidateNavigation,
  reduceOfficeSelectionSession,
  type OfficeCandidateNavigationState,
  type OfficeObjectReference,
  type OfficeReferenceConfirmEvent,
  type OfficeSelectionCancelEvent,
  type OfficeSelectionKeyboardResult,
  type OfficeSelectionSessionState,
} from "@arcships/office-interaction";

export const publicRuntimeExports = {
  OfficeInteractionValidationError,
  applyOfficeSelectionKeyboard,
  confirmOfficeCandidate,
  createOfficeCandidateNavigationState,
  createOfficeReferenceId,
  createOfficeSelectionSessionState,
  parseOfficeObjectReference,
  parseOfficeReferenceConfirmEvent,
  reduceOfficeCandidateNavigation,
  reduceOfficeSelectionSession,
  OfficeObjectOutlineLayer,
  OfficeRegionSelector,
  bundledDocxWasmUrl,
  createDocxPageReferenceDraft,
  createDocxRegionReferenceDraft,
  createDocxTextReferenceDraft,
  bundledXlsxWasmUrl,
  createXlsxCellReferenceDraft,
  createXlsxChartReferenceDraft,
  createXlsxRangeReferenceDraft,
  createXlsxRegionReferenceDraft,
  createXlsxWorksheetReferenceDraft,
  cloneDocModel,
  columnLabel,
  createDocxRuntime,
  createDocxRuntimeFromSubpath,
  describeDocxReference,
  describeXlsxReference,
  createXlsxRuntime,
  DocxImportError,
  docxWasmFromSubpath,
  layoutDocument,
  rangeToA1,
  xlsxWasmFromSubpath,
  XlsxRuntimeError,
  DocxViewer,
  XlsxViewer,
  useXlsxViewerController,
  PdfViewer,
  DEFAULT_PDF_MAX_FILE_SIZE,
  bundledPdfiumWasmUrl,
  createPdfPageReferenceDraft,
  createPdfRegionReferenceDraft,
  createPdfTextReferenceDraft,
  createPdfRenderRuntime,
  describePdfReference,
  normalizePdfReferenceRect,
  resolveDocxReference,
  resolvePdfReference,
  resolveXlsxReference,
  createPptxObjectReferenceDraft,
  createPptxRegionReferenceDraft,
  createPptxSlideReferenceDraft,
  describePptxReference,
  pptxReferenceKindForObject,
  resolvePptxReference,
  FileUpload,
};

// These names were already public in 0.1.x. They remain executable through
// 0.x while declarations steer new consumers to the high-level APIs.
export const legacyPublicExports = {
  setLegacyDocxWasmSource,
  setLegacyXlsxWasmSource,
  canUseConfiguredWasmSourceInWorker,
  getConfiguredWorkerWasmSource,
  getSheetsWasmModule,
  DocxEditorViewer,
  DocxViewerRoot,
  DocxPageWrapper,
  DocxPageSurface,
  DocxPageHeader,
  DocxPageFooter,
  DocxPageBody,
  DocxParagraphHost,
  DocxTableHost,
  DocxImageLayer,
  DocxFormFieldLayer,
  DocxTrackedChangeGutter,
  DocxContextMenu,
  DocxToolbar,
  DocxThumbnailPanel,
  DocxDragOverlay,
  renderParagraphRuns,
  renderStaticHtml,
  useDocxViewerThumbnails,
  createDocxViewerPageSurfaceRegistry,
  ensureDocxViewerPageSurfaceRegistry,
  subscribeDocxViewerPageSurfaces,
  notifyDocxViewerPageSurfaceListeners,
  MemoChartSvg,
  MemoSurfaceChartComposite,
  XlsxGrid,
  XlsxToolbar,
  XlsxRibbon,
  XlsxFormulaBar,
  XlsxSheetTabs,
  XlsxChartOverlay,
  XlsxImageLayer,
  XlsxSelectionOverlay,
  XlsxContextMenu,
};

export type PublicTypeExports = {
  officeCandidates: OfficeCandidateNavigationState;
  officeReference: OfficeObjectReference;
  officeReferenceConfirm: OfficeReferenceConfirmEvent;
  officeSelectionCancel: OfficeSelectionCancelEvent;
  officeSelectionKeyboard: OfficeSelectionKeyboardResult;
  officeSelectionSession: OfficeSelectionSessionState;
  officeObjectOutline: OfficeObjectOutline;
  officeObjectOutlineProps: OfficeObjectOutlineLayerProps;
  officeRegionSelectorProps: OfficeRegionSelectorProps;
  document: DocModel;
  docxOfficeReferenceDraft: DocxOfficeReferenceDraft;
  docxReferenceContext: DocxReferenceContext;
  workbook: XlsxViewerController;
  xlsxOfficeReferenceDraft: XlsxOfficeReferenceDraft;
  xlsxReferenceContext: XlsxReferenceContext;
  xlsxReferenceSheet: XlsxReferenceSheet;
  pptxOfficeReferenceDraft: PptxOfficeReferenceDraft;
  pptxObjectReferenceOptions: PptxObjectReferenceOptions;
  pptxReferenceContext: PptxReferenceContext;
  docxEditorOptions: UseDocxEditorOptions;
  xlsxDiagnostic: XlsxDiagnostic;
  pdfSource: PdfSource;
  pdfDiagnostic: PdfDiagnostic;
  pdfLoadOptions: PdfLoadOptions;
  pdfOfficeReferenceDraft: PdfOfficeReferenceDraft;
  pdfReferenceContext: PdfReferenceContext;
  pdfRenderDocument: PdfRenderDocument;
  pdfRenderRuntime: PdfRenderRuntime;
  pdfRenderRuntimeConfig: PdfRenderRuntimeConfig;
  pureDocument: PureDocModel;
  pureWorkbookChart: XlsxChart;
  docxFile: DocxFileLike;
  docxSource: DocxSource;
  docxUrlPolicy: DocxUrlPolicy;
  docxRuntimeLimits: DocxRuntimeLimits;
  docxRuntimeConfig: DocxRuntimeConfig;
  docxRuntimeLoader: DocxRuntimeLoader;
  docxRuntime: DocxRuntime;
  docxImportSource: DocxImportSource;
  docxImportErrorCode: DocxImportErrorCode;
  docxImportResult: DocxImportResult;
  docxImportDiagnosticType: DocxImportDiagnosticType;
  docxImportDiagnostic: DocxImportDiagnostic;
  docxImportOptions: DocxImportOptions;
  docxImportTimings: DocxImportWorkerTimings;
  xlsxRuntime: XlsxRuntime;
  xlsxRuntimeConfig: XlsxRuntimeConfig;
  xlsxRuntimeDiagnostic: XlsxRuntimeDiagnostic;
  xlsxRuntimeErrorCode: XlsxRuntimeErrorCode;
  xlsxRuntimeParseOptions: XlsxRuntimeParseOptions;
  legacyDocxViewerThumbnails: DocxViewerThumbnails;
  legacyDocxViewerThumbnailOptions: UseDocxViewerThumbnailsOptions;
  legacyChartPalette: ChartRendererPalette;
  legacyChartProps: ChartSvgProps;
  legacyChartLayout: ChartLayout;
  legacyLegendItem: LegendItem;
  fileUploadRejection: FileUploadRejection;
  fileUploadRejectionCode: FileUploadRejectionCode;
};
