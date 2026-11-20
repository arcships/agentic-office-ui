/**
 * Fixture controllers for the XLSX parity surfaces.
 *
 * The harness renders the production Vue components, not stand-ins, but it must
 * not boot the spreadsheet engine (the parity matrix only measures chrome). The
 * `XlsxViewerController` methods a chrome component touches are stubbed here and
 * every value comes from `tools/parity-fixtures/fixtures.json`, so the reference
 * page and the subject page render the same workbook.
 */
import type {
  XlsxCellAddress,
  XlsxSheetData,
  XlsxSelectionSummary,
  XlsxViewerController,
  XlsxWorkbookTab,
} from "@arcships/xlsx-core";

import { fixtures } from "../fixtures";

function sheet(name: string, workbookSheetIndex: number): XlsxSheetData {
  return {
    colStyleIds: [],
    maxUsedCol: 3,
    maxUsedRow: 12,
    name,
    rowStyleIds: [],
    styleById: {},
    visibleCols: [0, 1, 2, 3],
    visibleRows: [0, 1, 2, 3],
    workbookSheetIndex,
  } as unknown as XlsxSheetData;
}

function tabs(names: readonly string[]): XlsxWorkbookTab[] {
  return names.map((name, index) => ({
    id: `sheet-${index}`,
    kind: "sheet",
    name,
    sheetIndex: index,
    visibility: "visible",
    workbookSheetIndex: index,
  })) as unknown as XlsxWorkbookTab[];
}

export interface ChromeControllerOptions {
  disabled?: boolean;
  activeTabIndex?: number;
}

export function xlsxChromeController(
  options: ChromeControllerOptions = {},
): XlsxViewerController {
  const toolbar = fixtures.xlsxToolbar;
  const tabFixture = fixtures.xlsxSheetTabs;
  const { disabled = false } = options;
  const activeTabIndex = options.activeTabIndex ?? tabFixture.activeIndex;
  const sheets = disabled ? [] : tabFixture.sheets.map(sheet);
  const workbookTabs = disabled ? [] : tabs(tabFixture.sheets);
  const selection = {
    end: { col: 2, row: 9 } as XlsxCellAddress,
    start: { col: 0, row: 0 } as XlsxCellAddress,
  };
  const summary = {
    average: null,
    cellCount: 30,
    count: 3,
    max: null,
    min: null,
    numberFormat: null,
    numericCount: 3,
    sum: 42,
    truncated: false,
  } as unknown as XlsxSelectionSummary;
  const noop = () => undefined;

  return {
    activeCell: { col: 0, row: 0 },
    activeCellAddress: "A1",
    activeSheet: sheets[activeTabIndex] ?? sheets[0],
    activeSheetIndex: activeTabIndex,
    activeTabIndex,
    activeTab: workbookTabs[activeTabIndex] ?? workbookTabs[0],
    addSheet: noop,
    canDownload: toolbar.showDownload,
    canExport: toolbar.showExport,
    canRedo: false,
    canUndo: false,
    canZoomIn: !disabled,
    canZoomOut: !disabled,
    clearSelection: noop,
    displayFileName: toolbar.fileName,
    download: noop,
    exportXlsx: noop,
    freezePanes: null,
    getActiveWorksheet: () => null,
    getSelectionStatsAsync: async () => summary,
    getSheetNameError: () => null,
    hiddenAutoFilterRows: [],
    hiddenSheets: [],
    isEditingLocked: false,
    isLoading: false,
    isPageBreakPreview: false,
    isWorkerBacked: true,
    lastExportWarning: "",
    locale: "en",
    moveSheet: () => true,
    openPrintPreview: noop,
    readOnly: true,
    removeActiveSheet: noop,
    renameSheet: () => true,
    replaceInSheet: () => ({ replacedOccurrences: 0 }),
    revision: 1,
    selectedRangeAddress: "A1:C10",
    selection,
    setActiveSheetIndex: noop,
    setActiveTabIndex: noop,
    setPageBreakPreview: noop,
    setSheetVisibility: () => true,
    setZoomScale: noop,
    sheets,
    tabs: workbookTabs,
    zoomIn: noop,
    zoomOut: noop,
    zoomScale: toolbar.zoomPercent,
  } as unknown as XlsxViewerController;
}
