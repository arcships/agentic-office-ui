import {
  bundledDocxWasmUrl,
  createDocxRuntime,
  setWasmSource as setLegacyDocxWasmSource,
  type DocModel,
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
  getConfiguredWorkerWasmSource,
  getSheetsWasmModule,
  setWasmSource as setLegacyXlsxWasmSource,
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
  createPdfRenderRuntime,
  type PdfDiagnostic,
  type PdfLoadOptions,
  type PdfRenderDocument,
  type PdfRenderRuntime,
  type PdfRenderRuntimeConfig,
  type PdfSource,
} from "@arcships/vue-pdf";
import {
  FileUpload,
  type FileUploadRejection,
  type FileUploadRejectionCode,
} from "@arcships/vue-ui";

export const publicRuntimeExports = {
  bundledDocxWasmUrl,
  bundledXlsxWasmUrl,
  cloneDocModel,
  columnLabel,
  createDocxRuntime,
  createDocxRuntimeFromSubpath,
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
  createPdfRenderRuntime,
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
  document: DocModel;
  workbook: XlsxViewerController;
  docxEditorOptions: UseDocxEditorOptions;
  xlsxDiagnostic: XlsxDiagnostic;
  pdfSource: PdfSource;
  pdfDiagnostic: PdfDiagnostic;
  pdfLoadOptions: PdfLoadOptions;
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
