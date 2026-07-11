import type { CellValue, Workbook, Worksheet } from "@dukelib/sheets-wasm";

const WORKBOOK_LIFETIME = Symbol("@arcships/xlsx-core/workbook-lifetime");
const DISPOSE_SYMBOL = (Symbol as SymbolConstructor & { dispose?: symbol }).dispose;

interface WorkbookLifetimeState {
  cells: Set<CellValue>;
  disposed: boolean;
  disposing: boolean;
  worksheets: Map<Worksheet, Set<CellValue>>;
}

type TrackedWorkbook = Workbook & {
  [WORKBOOK_LIFETIME]?: WorkbookLifetimeState;
};

function defineInstanceMethod(
  target: object,
  key: PropertyKey,
  value: (...args: never[]) => unknown
): void {
  Object.defineProperty(target, key, {
    configurable: true,
    enumerable: false,
    value,
    writable: true
  });
}

function rememberFirstError(current: unknown, next: unknown): unknown {
  return current ?? next;
}

function trackCellValue(
  state: WorkbookLifetimeState,
  worksheetCells: Set<CellValue>,
  cellValue: CellValue
): CellValue {
  if (state.cells.has(cellValue)) return cellValue;

  const originalFree = cellValue.free.bind(cellValue);
  let released = false;
  const freeCellValue = (): void => {
    if (released) return;
    released = true;
    worksheetCells.delete(cellValue);
    state.cells.delete(cellValue);
    originalFree();
  };

  state.cells.add(cellValue);
  worksheetCells.add(cellValue);
  defineInstanceMethod(cellValue, "free", freeCellValue);
  if (DISPOSE_SYMBOL) defineInstanceMethod(cellValue, DISPOSE_SYMBOL, freeCellValue);
  return cellValue;
}

function trackWorksheet(
  state: WorkbookLifetimeState,
  worksheet: Worksheet
): Worksheet {
  if (state.worksheets.has(worksheet)) return worksheet;

  const worksheetCells = new Set<CellValue>();
  const originalFree = worksheet.free.bind(worksheet);
  const originalGetCell = worksheet.getCell.bind(worksheet);
  const originalGetCellAt = worksheet.getCellAt.bind(worksheet);
  const originalGetCalculatedValue = worksheet.getCalculatedValue.bind(worksheet);
  const originalGetCalculatedValueAt = worksheet.getCalculatedValueAt.bind(worksheet);
  let released = false;

  const freeWorksheet = (): void => {
    if (released) return;
    released = true;
    let firstError: unknown;
    for (const cellValue of [...worksheetCells]) {
      try {
        cellValue.free();
      } catch (error) {
        firstError = rememberFirstError(firstError, error);
      }
    }
    state.worksheets.delete(worksheet);
    try {
      originalFree();
    } catch (error) {
      firstError = rememberFirstError(firstError, error);
    }
    if (firstError) throw firstError;
  };

  state.worksheets.set(worksheet, worksheetCells);
  defineInstanceMethod(worksheet, "free", freeWorksheet);
  if (DISPOSE_SYMBOL) defineInstanceMethod(worksheet, DISPOSE_SYMBOL, freeWorksheet);
  defineInstanceMethod(worksheet, "getCell", ((address: string) =>
    trackCellValue(state, worksheetCells, originalGetCell(address))) as never);
  defineInstanceMethod(worksheet, "getCellAt", ((row: number, col: number) =>
    trackCellValue(state, worksheetCells, originalGetCellAt(row, col))) as never);
  defineInstanceMethod(worksheet, "getCalculatedValue", ((address: string) =>
    trackCellValue(state, worksheetCells, originalGetCalculatedValue(address))) as never);
  defineInstanceMethod(worksheet, "getCalculatedValueAt", ((row: number, col: number) =>
    trackCellValue(state, worksheetCells, originalGetCalculatedValueAt(row, col))) as never);
  return worksheet;
}

/**
 * Makes every WASM child handle created through one workbook explicitly owned
 * by that workbook. No shared module-level registry is used.
 */
export function trackXlsxWorkbookLifetime<T extends Workbook>(workbook: T): T {
  const trackedWorkbook = workbook as T & TrackedWorkbook;
  if (trackedWorkbook[WORKBOOK_LIFETIME]) return workbook;
  if (!Object.isExtensible(workbook)) {
    throw new TypeError("XLSX Workbook 实例不可扩展，无法登记子对象生命周期。");
  }

  const state: WorkbookLifetimeState = {
    cells: new Set(),
    disposed: false,
    disposing: false,
    worksheets: new Map()
  };
  const originalFree = workbook.free.bind(workbook);
  const originalGetSheet = workbook.getSheet.bind(workbook);
  const originalGetSheetByName = workbook.getSheetByName.bind(workbook);

  const assertAlive = (): void => {
    if (state.disposed || state.disposing) {
      throw new Error("XLSX Workbook 已释放，不能再创建 Worksheet。");
    }
  };
  const freeWorkbook = (): void => {
    if (state.disposed || state.disposing) return;
    state.disposing = true;
    let firstError: unknown;

    for (const cellValue of [...state.cells]) {
      try {
        cellValue.free();
      } catch (error) {
        firstError = rememberFirstError(firstError, error);
      }
    }
    for (const worksheet of [...state.worksheets.keys()]) {
      try {
        worksheet.free();
      } catch (error) {
        firstError = rememberFirstError(firstError, error);
      }
    }
    try {
      originalFree();
    } catch (error) {
      firstError = rememberFirstError(firstError, error);
    } finally {
      state.cells.clear();
      state.worksheets.clear();
      state.disposed = true;
      state.disposing = false;
    }

    if (firstError) throw firstError;
  };

  Object.defineProperty(trackedWorkbook, WORKBOOK_LIFETIME, {
    configurable: false,
    enumerable: false,
    value: state,
    writable: false
  });
  defineInstanceMethod(workbook, "free", freeWorkbook);
  if (DISPOSE_SYMBOL) defineInstanceMethod(workbook, DISPOSE_SYMBOL, freeWorkbook);
  defineInstanceMethod(workbook, "getSheet", ((index: number) => {
    assertAlive();
    return trackWorksheet(state, originalGetSheet(index));
  }) as never);
  defineInstanceMethod(workbook, "getSheetByName", ((name: string) => {
    assertAlive();
    return trackWorksheet(state, originalGetSheetByName(name));
  }) as never);
  return workbook;
}
