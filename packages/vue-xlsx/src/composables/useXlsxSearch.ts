import {
  getCurrentScope,
  nextTick,
  onScopeDispose,
  shallowReadonly,
  shallowRef,
  watch,
  type ShallowRef,
} from "vue"
import {
  createSurfaceSearchSession,
} from "@arcships/office-runtime"
import type { XlsxCellAddress, XlsxSheetData, XlsxViewerController } from "@arcships/xlsx-core"

export interface XlsxSearchMatch {
  kind: "xlsx-cell"
  address: string
  cell: XlsxCellAddress
  sheetIndex: number
  sheetName: string
  text: string
  start: number
  end: number
  workbookSheetIndex: number
}

export type XlsxSearchStatus = "idle" | "searching" | "ready" | "error"

export interface XlsxSearchState {
  status: XlsxSearchStatus
  query: string
  matches: readonly XlsxSearchMatch[]
  activeIndex: number
  error?: {
    code: "SEARCH_FAILED" | "ACTIVATION_FAILED"
    message: string
  }
}

export interface XlsxSurfaceSearch {
  readonly searchState: Readonly<ShallowRef<XlsxSearchState>>
  search(query: string): Promise<XlsxSearchState>
  activateSearchMatch(index: number): Promise<void>
  searchNext(): Promise<void>
  searchPrevious(): Promise<void>
  clearSearch(): void
  getSearchState(): XlsxSearchState
  dispose(): void
}

interface UseXlsxSearchOptions {
  scrollToCell?: (cell: XlsxCellAddress) => Promise<void> | void
}

function cellAddress(cell: XlsxCellAddress): string {
  let col = ""
  let index = cell.col
  while (index >= 0) {
    col = String.fromCharCode(65 + index % 26) + col
    index = Math.floor(index / 26) - 1
  }
  return `${col}${cell.row + 1}`
}

function abortError(): Error {
  if (typeof DOMException !== "undefined") return new DOMException("XLSX search aborted.", "AbortError")
  const error = new Error("XLSX search aborted.")
  error.name = "AbortError"
  return error
}

function normalizeMergedCell(sheet: XlsxSheetData, cell: XlsxCellAddress): XlsxCellAddress {
  const merged = sheet.mergedRegions?.find((range) =>
    cell.row >= range.start.row
    && cell.row <= range.end.row
    && cell.col >= range.start.col
    && cell.col <= range.end.col
  )
  return merged ? { row: merged.start.row, col: merged.start.col } : cell
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^\${}()|[\]\\]/g, "\\$&")
}

function matchedText(value: string, formula: string, expression: RegExp): {
  text: string
  start: number
  end: number
} | undefined {
  const valueMatch = expression.exec(value)
  if (valueMatch) return {
    text: value,
    start: valueMatch.index,
    end: valueMatch.index + valueMatch[0].length,
  }
  const formulaMatch = expression.exec(formula)
  if (formulaMatch) return {
    text: formula,
    start: formulaMatch.index,
    end: formulaMatch.index + formulaMatch[0].length,
  }
  return undefined
}

async function searchWorkbook(
  controller: XlsxViewerController,
  query: string,
  signal: AbortSignal,
): Promise<XlsxSearchMatch[]> {
  const expression = new RegExp(escapeRegExp(query), "iu")
  const matches: XlsxSearchMatch[] = []

  for (const [sheetIndex, sheet] of controller.sheets.entries()) {
    if (signal.aborted) throw abortError()
    if (controller.isWorkerBacked && controller.getRowsBatchAsync) {
      const rowLimit = Math.max(0, sheet.maxUsedRow + 1)
      for (let rowStart = 0; rowStart < rowLimit; rowStart += 128) {
        const rows = await controller.getRowsBatchAsync(
          sheet.workbookSheetIndex,
          rowStart,
          128,
          signal,
        )
        if (signal.aborted) throw abortError()
        for (const rowEntry of rows ?? []) {
          if (!rowEntry || typeof rowEntry !== "object") continue
          const row = Number(Reflect.get(rowEntry, "index"))
          const cells = Reflect.get(rowEntry, "cells")
          if (!Number.isInteger(row) || !Array.isArray(cells)) continue
          for (const entry of cells) {
            const col = Number(Reflect.get(entry, "col"))
            if (!Number.isInteger(col)) continue
            const value = String(Reflect.get(entry, "value") ?? "")
            const formula = String(Reflect.get(entry, "formula") ?? "")
            const matched = matchedText(value, formula, expression)
            if (!matched) continue
            const cell = normalizeMergedCell(sheet, { row, col })
            matches.push({
              kind: "xlsx-cell",
              address: cellAddress(cell),
              cell,
              sheetIndex,
              sheetName: sheet.name,
              ...matched,
              workbookSheetIndex: sheet.workbookSheetIndex,
            })
          }
        }
      }
      continue
    }

    const worksheet = controller.workbook?.getSheet(sheet.workbookSheetIndex)
    if (!worksheet) continue
    try {
      const rows = sheet.visibleRows.filter((row) => row >= sheet.minUsedRow && row <= sheet.maxUsedRow)
      const columns = sheet.visibleCols.filter((col) => col >= sheet.minUsedCol && col <= sheet.maxUsedCol)
      for (const row of rows) {
        if (signal.aborted) throw abortError()
        for (const col of columns) {
          const formatted = worksheet.getFormattedValueAt(row, col) ?? ""
          const formula = worksheet.getFormulaAt(row, col) ?? ""
          let value = formatted
          if (!value) {
            const calculated = worksheet.getCalculatedValueAt(row, col)
            try {
              value = calculated.is_empty ? "" : calculated.toString()
            } finally {
              calculated.free()
            }
          }
          const matched = matchedText(value, formula, expression)
          if (!matched) continue
          const cell = normalizeMergedCell(sheet, { row, col })
          matches.push({
            kind: "xlsx-cell",
            address: cellAddress(cell),
            cell,
            sheetIndex,
            sheetName: sheet.name,
            ...matched,
            workbookSheetIndex: sheet.workbookSheetIndex,
          })
        }
      }
    } finally {
      worksheet.free()
    }
  }

  return matches
}

export function useXlsxSearch(
  controller: () => XlsxViewerController,
  options: UseXlsxSearchOptions = {},
): XlsxSurfaceSearch {
  const searchState = shallowRef<XlsxSearchState>({
    status: "idle",
    query: "",
    matches: [],
    activeIndex: -1,
  })
  const session = createSurfaceSearchSession<XlsxSearchMatch>({
    search(query, { signal }) {
      return searchWorkbook(controller(), query, signal)
    },
    async activate(match, _index, { signal }) {
      const current = controller()
      if (current.activeSheetIndex !== match.sheetIndex) {
        current.setActiveSheetIndex(match.sheetIndex)
        await nextTick()
      }
      if (signal.aborted || controller() !== current) return
      await options.scrollToCell?.(match.cell)
      if (signal.aborted || controller() !== current) return
      current.selectCell(match.cell)
    },
  })
  session.subscribe((next) => {
    searchState.value = next
  })
  watch(controller, () => session.clearSearch())

  const result: XlsxSurfaceSearch = {
    searchState: shallowReadonly(searchState),
    search: (query) => session.search(query),
    activateSearchMatch: (index) => session.activateSearchMatch(index),
    searchNext: () => session.searchNext(),
    searchPrevious: () => session.searchPrevious(),
    clearSearch: () => session.clearSearch(),
    getSearchState: () => session.getSearchState(),
    dispose: () => session.dispose(),
  }
  if (getCurrentScope()) onScopeDispose(result.dispose)
  return result
}
