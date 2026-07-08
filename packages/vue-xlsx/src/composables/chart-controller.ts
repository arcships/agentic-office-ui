import { computed, watch } from "vue";
import type { Workbook } from "@dukelib/sheets-wasm";
import {
  applyChartSeriesFormula,
  buildChartSeriesFormula,
  rectToImageAnchor,
  resizeImageRect,
  resolveContentSheetAxisPixels,
  resolveRenderedSheetAxisPixels,
  resolveSheetColumnWidthPixels,
  resolveSheetRowHeightPixels,
  updateWorkbookChartAnchor,
  updateWorkbookChartDefinition,
  updateWorkbookImageAnchor,
  type XlsxChart,
  type XlsxChartElementSelection,
  type XlsxFormControl,
  type XlsxImage,
  type XlsxImageRect,
  type XlsxImageResizeHandlePosition
} from "@extend-ai/xlsx-core";
import {
  DEFAULT_COL_WIDTH,
  DEFAULT_ROW_HEIGHT,
  GRID_HEADER_HEIGHT,
  GRID_ROW_HEADER_WIDTH
} from "./internal";
import type { XlsxControllerContext } from "./navigation";

// Re-export image/chart asset parsing helpers so consumers can import the full
// surface from the chart-controller module.
export {
  loadWorkbookImageAssets,
  shouldSkipXmlParsingForWorkbook,
  isZipWorkbook,
  isLegacyXlsWorkbook,
  createBasicWorkbookAssets,
  collectWorksheetApiImages,
  collectWorksheetApiShapes,
  collectWorksheetBatchImages,
  mergeParsedAndApiImages,
  buildWorksheetApiImage,
  buildWorksheetDirectApiImage,
  buildWorksheetDirectApiShape,
  buildWorksheetDirectImageAnchor,
  normalizeWorksheetDirectShapeParagraphs,
  createWorksheetDirectImageSource,
  inferImageMimeType,
  inferWorksheetDirectImageMimeType
} from "./image-assets";

export function createChartImageDomain(ctx: XlsxControllerContext) {
  function mapPublicChart(chart: XlsxChart) {
    const visibleSheetIndex = ctx.visibleSheetIndexByWorkbookSheetIndex.value.get(chart.workbookSheetIndex);
    return {
      ...chart,
      sheetIndex: visibleSheetIndex ?? chart.workbookSheetIndex
    };
  }

  function mapPublicImage(image: XlsxImage) {
    const visibleSheetIndex = ctx.visibleSheetIndexByWorkbookSheetIndex.value.get(image.workbookSheetIndex);
    return {
      ...image,
      sheetIndex: visibleSheetIndex ?? image.workbookSheetIndex
    };
  }

  function mapPublicFormControl(control: XlsxFormControl) {
    const visibleSheetIndex = ctx.visibleSheetIndexByWorkbookSheetIndex.value.get(control.workbookSheetIndex);
    return {
      ...control,
      sheetIndex: visibleSheetIndex ?? control.workbookSheetIndex
    };
  }

  const publicChartsByWorkbookSheetIndex = computed(() => { return ctx.chartsByWorkbookSheetIndex.value.map((sheetCharts) => sheetCharts.map(mapPublicChart)); })
  const publicChartById = computed(() => {
    const lookup = new Map<string, XlsxChart>();
    for (const sheetCharts of publicChartsByWorkbookSheetIndex.value) {
      for (const chart of sheetCharts) {
        lookup.set(chart.id, chart);
      }
    }
    return lookup;
  })

  function getSheetCharts(sheetIndex = ctx.activeSheetIndex.value) {
    const targetSheet = ctx.sheets.value[sheetIndex];
    if (!targetSheet) {
      return [];
    }

    return publicChartsByWorkbookSheetIndex.value[targetSheet.workbookSheetIndex] ?? [];
  }

  function getChartById(id: string) {
    return publicChartById.value.get(id) ?? null;
  }

  function getChartsheetById(id: string) {
  return (
    ctx.chartsheets.value.find((chartsheet) => chartsheet.id === id) ?? null
  );
}

  const charts = computed(() => {
    if (ctx.activeTab.value?.kind === "chartsheet" && typeof ctx.activeTab.value.chartsheetIndex === "number") {
      const chartsheet = ctx.chartsheets.value[ctx.activeTab.value.chartsheetIndex];
      return (chartsheet?.chartIds ?? []).map((id) => getChartById(id)).filter((value): value is XlsxChart => Boolean(value));
    }

    return getSheetCharts(ctx.activeSheetIndex.value);
  })

  const selectedChart = computed(() => { return (ctx.selectedChartId.value ? getChartById(ctx.selectedChartId.value) : null); })

  watch(
    () => [selectedChart.value, ctx.selectedChartElement.value, ctx.selectedChartId.value],
    () => {
      if (!ctx.selectedChartId.value) {
        if (ctx.selectedChartElement.value) {
          ctx.selectedChartElement.value = null;
        }
        return;
      }

      if (!selectedChart.value) {
        ctx.selectedChartId.value = null;
        ctx.selectedChartElement.value = null;
        return;
      }

      if (!ctx.selectedChartElement.value) {
        ctx.selectedChartElement.value = { chartId: ctx.selectedChartId.value, kind: "chart" };
        return;
      }

      if (ctx.selectedChartElement.value.chartId !== ctx.selectedChartId.value) {
        ctx.selectedChartElement.value = { chartId: ctx.selectedChartId.value, kind: "chart" };
        return;
      }

      if (ctx.selectedChartElement.value.kind !== "chart") {
        const selectedSeries = selectedChart.value.series[ctx.selectedChartElement.value.seriesIndex];
        if (!selectedSeries || selectedSeries.id !== ctx.selectedChartElement.value.seriesId) {
          ctx.selectedChartElement.value = { chartId: ctx.selectedChartId.value, kind: "chart" };
        }
      }
    }
  );

  function selectChart(id: string | null) {
    ctx.selectedImageId.value = null;
    ctx.selectedChartId.value = id;
    ctx.selectedChartElement.value = id ? { chartId: id, kind: "chart" } : null;
  }

  function clearSelectedChart() {
    ctx.selectedChartId.value = null;
    ctx.selectedChartElement.value = null;
  }

  function clearSelectedChartElement() {
    ctx.selectedChartElement.value = ctx.selectedChartId.value ? { chartId: ctx.selectedChartId.value, kind: "chart" } : null;
  }

  function selectChartElement(selection: XlsxChartElementSelection | null) {
    ctx.selectedImageId.value = null;
    ctx.selectedChartId.value = selection?.chartId ?? null;
    ctx.selectedChartElement.value = selection;
  }

  function getSheetImages(sheetIndex = ctx.activeSheetIndex.value) {
    const targetSheet = ctx.sheets.value[sheetIndex];
    if (!targetSheet) {
      return [];
    }

    return (ctx.imagesByWorkbookSheetIndex.value[targetSheet.workbookSheetIndex] ?? []).map(mapPublicImage);
  }

  const images = computed(() => { return getSheetImages(ctx.activeSheetIndex.value); })

  function getSheetFormControls(sheetIndex = ctx.activeSheetIndex.value) {
    const targetSheet = ctx.sheets.value[sheetIndex];
    if (!targetSheet) {
      return [];
    }

    return (ctx.formControlsByWorkbookSheetIndex.value[targetSheet.workbookSheetIndex] ?? []).map(mapPublicFormControl);
  }

  const formControls = computed(() => { return getSheetFormControls(ctx.activeSheetIndex.value); })

  function getSheetShapes(sheetIndex = ctx.activeSheetIndex.value) {
    const targetSheet = ctx.sheets.value[sheetIndex];
    if (!targetSheet) {
      return [];
    }

    return (ctx.shapesByWorkbookSheetIndex.value[targetSheet.workbookSheetIndex] ?? []).map((shape) => {
      const visibleSheetIndex = ctx.visibleSheetIndexByWorkbookSheetIndex.value.get(shape.workbookSheetIndex);
      return {
        ...shape,
        sheetIndex: visibleSheetIndex ?? shape.workbookSheetIndex
      };
    });
  }

  const shapes = computed(() => { return getSheetShapes(ctx.activeSheetIndex.value); })

  function getImageById(id: string) {
    for (const sheetImages of ctx.imagesByWorkbookSheetIndex.value) {
      const match = sheetImages?.find((image) => image.id === id);
      if (match) {
        return mapPublicImage(match);
      }
    }

    return null;
  }

  const selectedImage = computed(() => { return (ctx.selectedImageId.value ? getImageById(ctx.selectedImageId.value) : null); })

  function selectImage(id: string | null) {
    ctx.selectedChartId.value = null;
    ctx.selectedChartElement.value = null;
    ctx.selectedImageId.value = id;
  }

  function clearSelectedImage() {
    ctx.selectedImageId.value = null;
  }

  function getColumnWidthPx(worksheet: ReturnType<Workbook["getSheet"]>, col: number) {
    const sheetState = ctx.imageAssetsRef.value?.sheetStatesByWorkbookSheetIndex[ctx.activeSheet.value?.workbookSheetIndex ?? -1] ?? null;
    const width = worksheet.getColumnWidth(col);
    const showGridLines = ctx.activeSheet.value?.showGridLines ?? true;
    if (width !== undefined && width !== null) {
      return resolveRenderedSheetAxisPixels(
        resolveSheetColumnWidthPixels(width, sheetState?.columnWidthCharacterWidthPx),
        showGridLines
      );
    }

    return resolveRenderedSheetAxisPixels(
      sheetState?.colWidthOverridesPx?.[col] ?? sheetState?.defaultColWidthPx ?? DEFAULT_COL_WIDTH,
      showGridLines
    );
  }

  function getRowHeightPx(worksheet: ReturnType<Workbook["getSheet"]>, row: number) {
    const sheetState = ctx.imageAssetsRef.value?.sheetStatesByWorkbookSheetIndex[ctx.activeSheet.value?.workbookSheetIndex ?? -1] ?? null;
    const height = worksheet.getRowHeight(row);
    const showGridLines = ctx.activeSheet.value?.showGridLines ?? true;
    if (height !== undefined && height !== null) {
      return resolveRenderedSheetAxisPixels(resolveSheetRowHeightPixels(height), showGridLines);
    }

    return resolveRenderedSheetAxisPixels(
      sheetState?.rowHeightOverridesPx?.[row] ?? sheetState?.defaultRowHeightPx ?? DEFAULT_ROW_HEIGHT,
      showGridLines
    );
  }

  function resolveAnchoredObjectRect(anchor: XlsxImage["anchor"], worksheet: ReturnType<Workbook["getSheet"]>) {
    const resolveAxisSum = (
      index: number,
      getSize: (target: number) => number
    ) => {
      let total = 0;
      for (let cursor = 0; cursor < index; cursor += 1) {
        total += getSize(cursor);
      }
      return total;
    };

    if (anchor.kind === "absolute") {
      return {
        height: anchor.sizeEmu.cy / 9525,
        left: GRID_ROW_HEADER_WIDTH + anchor.positionEmu.x / 9525,
        top: GRID_HEADER_HEIGHT + anchor.positionEmu.y / 9525,
        width: anchor.sizeEmu.cx / 9525
      };
    }

    const left = GRID_ROW_HEADER_WIDTH + resolveAxisSum(anchor.from.col, (col) => getColumnWidthPx(worksheet, col)) + anchor.from.colOffsetEmu / 9525;
    const top = GRID_HEADER_HEIGHT + resolveAxisSum(anchor.from.row, (row) => getRowHeightPx(worksheet, row)) + anchor.from.rowOffsetEmu / 9525;

    if (anchor.kind === "one-cell") {
      return {
        height: anchor.sizeEmu.cy / 9525,
        left,
        top,
        width: anchor.sizeEmu.cx / 9525
      };
    }

    const right = GRID_ROW_HEADER_WIDTH + resolveAxisSum(anchor.to.col, (col) => getColumnWidthPx(worksheet, col)) + anchor.to.colOffsetEmu / 9525;
    const bottom = GRID_HEADER_HEIGHT + resolveAxisSum(anchor.to.row, (row) => getRowHeightPx(worksheet, row)) + anchor.to.rowOffsetEmu / 9525;

    return {
      height: Math.max(1, bottom - top),
      left,
      top,
      width: Math.max(1, right - left)
    };
  }

  function setChartRect(id: string, rect: XlsxImageRect) {
    const hydratedChartAssets = ctx.ensureChartAssetsHydrated(ctx.workbook.value, ctx.sheets.value);
    console.info("[react-xlsx debug] setChartRect", {
      hasActiveSheet: Boolean(ctx.activeSheet.value),
      hasHydratedChartAssets: Boolean(hydratedChartAssets),
      hasImageAssets: Boolean(ctx.imageAssetsRef.value),
      hasWorkbook: Boolean(ctx.workbook.value),
      id,
      readOnly: ctx.readOnly.value,
      rect
    });
    if (ctx.readOnly.value || !ctx.workbook.value || !ctx.activeSheet.value || !ctx.imageAssetsRef.value || !hydratedChartAssets) {
      return;
    }

    const worksheet = ctx.workbook.value.getSheet(ctx.activeSheet.value.workbookSheetIndex);
    const currentChart = getChartById(id);
    console.info("[react-xlsx debug] currentChart", {
      activeWorkbookSheetIndex: ctx.activeSheet.value.workbookSheetIndex,
      editable: currentChart?.editable,
      found: Boolean(currentChart),
      originCount: hydratedChartAssets.chartOriginsById.size,
      workbookSheetIndex: currentChart?.workbookSheetIndex
    });
    if (!currentChart || currentChart.editable === false || currentChart.workbookSheetIndex !== ctx.activeSheet.value.workbookSheetIndex) {
      return;
    }

    const nextAnchor = rectToImageAnchor(rect, currentChart.anchor, {
      contentOffsetLeft: GRID_ROW_HEADER_WIDTH,
      contentOffsetTop: GRID_HEADER_HEIGHT,
      getColumnWidthPx: (col) => getColumnWidthPx(worksheet, col),
      getRowHeightPx: (row) => getRowHeightPx(worksheet, row)
    });

    ctx.recordHistoryBeforeMutation();
    const didUpdateAnchor = updateWorkbookChartAnchor(ctx.imageAssetsRef.value, hydratedChartAssets, id, nextAnchor);
    console.info("[react-xlsx debug] updateWorkbookChartAnchor", { didUpdateAnchor, nextAnchor });

    hydratedChartAssets.chartsByWorkbookSheetIndex = hydratedChartAssets.chartsByWorkbookSheetIndex.map((sheetCharts) => (
      sheetCharts.map((chart) => chart.id === id ? { ...chart, anchor: nextAnchor } : chart)
    ));

    ctx.chartsByWorkbookSheetIndex.value = ctx.chartsByWorkbookSheetIndex.value.map((sheetCharts) => (
      sheetCharts.map((chart) => chart.id === id ? { ...chart, anchor: nextAnchor } : chart)
    ));
    ctx.revision.value = ctx.revision.value + 1;
  }

  function setImageRect(id: string, rect: XlsxImageRect) {
    if (ctx.readOnly.value || !ctx.workbook.value || !ctx.activeSheet.value || !ctx.imageAssetsRef.value) {
      return;
    }

    const worksheet = ctx.workbook.value.getSheet(ctx.activeSheet.value.workbookSheetIndex);
    const currentImage = getImageById(id);
    if (!currentImage || currentImage.editable === false || currentImage.workbookSheetIndex !== ctx.activeSheet.value.workbookSheetIndex) {
      return;
    }

    const nextAnchor = rectToImageAnchor(rect, currentImage.anchor, {
      contentOffsetLeft: GRID_ROW_HEADER_WIDTH,
      contentOffsetTop: GRID_HEADER_HEIGHT,
      getColumnWidthPx: (col) => getColumnWidthPx(worksheet, col),
      getRowHeightPx: (row) => getRowHeightPx(worksheet, row)
    });

    ctx.recordHistoryBeforeMutation();
    if (!updateWorkbookImageAnchor(ctx.imageAssetsRef.value, id, nextAnchor)) {
      return;
    }

    ctx.imagesByWorkbookSheetIndex.value = [...ctx.imageAssetsRef.value.imagesByWorkbookSheetIndex];
    ctx.revision.value = ctx.revision.value + 1;
  }

  function moveChartBy(id: string, deltaX: number, deltaY: number) {
    const currentChart = getChartById(id);
    if (!currentChart || currentChart.editable === false) {
      return;
    }

    const worksheet = ctx.getActiveWorksheet();
    if (!worksheet) {
      return;
    }

    const currentRect = resolveAnchoredObjectRect(currentChart.anchor, worksheet);
    setChartRect(id, {
      ...currentRect,
      left: currentRect.left + deltaX,
      top: currentRect.top + deltaY
    });
  }

  function moveImageBy(id: string, deltaX: number, deltaY: number) {
    const currentImage = getImageById(id);
    if (!currentImage || currentImage.editable === false) {
      return;
    }

    const currentRect = (() => {
      const worksheet = ctx.getActiveWorksheet();
      if (!worksheet) {
        return null;
      }

      const resolveAxisSum = (
        index: number,
        getSize: (target: number) => number
      ) => {
        let total = 0;
        for (let cursor = 0; cursor < index; cursor += 1) {
          total += getSize(cursor);
        }
        return total;
      };

      if (currentImage.anchor.kind === "absolute") {
        return {
          height: currentImage.anchor.sizeEmu.cy / 9525,
          left: GRID_ROW_HEADER_WIDTH + currentImage.anchor.positionEmu.x / 9525,
          top: GRID_HEADER_HEIGHT + currentImage.anchor.positionEmu.y / 9525,
          width: currentImage.anchor.sizeEmu.cx / 9525
        };
      }

      const left = GRID_ROW_HEADER_WIDTH + resolveAxisSum(currentImage.anchor.from.col, (col) => getColumnWidthPx(worksheet, col)) + currentImage.anchor.from.colOffsetEmu / 9525;
      const top = GRID_HEADER_HEIGHT + resolveAxisSum(currentImage.anchor.from.row, (row) => getRowHeightPx(worksheet, row)) + currentImage.anchor.from.rowOffsetEmu / 9525;

      if (currentImage.anchor.kind === "one-cell") {
        return {
          height: currentImage.anchor.sizeEmu.cy / 9525,
          left,
          top,
          width: currentImage.anchor.sizeEmu.cx / 9525
        };
      }

      const right = GRID_ROW_HEADER_WIDTH + resolveAxisSum(currentImage.anchor.to.col, (col) => getColumnWidthPx(worksheet, col)) + currentImage.anchor.to.colOffsetEmu / 9525;
      const bottom = GRID_HEADER_HEIGHT + resolveAxisSum(currentImage.anchor.to.row, (row) => getRowHeightPx(worksheet, row)) + currentImage.anchor.to.rowOffsetEmu / 9525;

      return {
        height: Math.max(1, bottom - top),
        left,
        top,
        width: Math.max(1, right - left)
      };
    })();

    if (!currentRect) {
      return;
    }

    setImageRect(id, {
      ...currentRect,
      left: currentRect.left + deltaX,
      top: currentRect.top + deltaY
    });
  }

  function resizeChartBy(id: string, handle: XlsxImageResizeHandlePosition, deltaX: number, deltaY: number) {
    const currentChart = getChartById(id);
    if (!currentChart || currentChart.editable === false) {
      return;
    }

    const worksheet = ctx.getActiveWorksheet();
    if (!worksheet) {
      return;
    }

    const currentRect = resolveAnchoredObjectRect(currentChart.anchor, worksheet);
    setChartRect(id, resizeImageRect(currentRect, handle, deltaX, deltaY, 48));
  }

  function resizeImageBy(id: string, handle: XlsxImageResizeHandlePosition, deltaX: number, deltaY: number) {
    const currentImage = getImageById(id);
    if (!currentImage || currentImage.editable === false) {
      return;
    }

    const worksheet = ctx.getActiveWorksheet();
    if (!worksheet) {
      return;
    }

    const resolveAxisSum = (
      index: number,
      getSize: (target: number) => number
    ) => {
      let total = 0;
      for (let cursor = 0; cursor < index; cursor += 1) {
        total += getSize(cursor);
      }
      return total;
    };

    const left = currentImage.anchor.kind === "absolute"
      ? GRID_ROW_HEADER_WIDTH + currentImage.anchor.positionEmu.x / 9525
      : GRID_ROW_HEADER_WIDTH + resolveAxisSum(currentImage.anchor.from.col, (col) => getColumnWidthPx(worksheet, col)) + currentImage.anchor.from.colOffsetEmu / 9525;
    const top = currentImage.anchor.kind === "absolute"
      ? GRID_HEADER_HEIGHT + currentImage.anchor.positionEmu.y / 9525
      : GRID_HEADER_HEIGHT + resolveAxisSum(currentImage.anchor.from.row, (row) => getRowHeightPx(worksheet, row)) + currentImage.anchor.from.rowOffsetEmu / 9525;
    const width = currentImage.anchor.kind === "two-cell"
      ? Math.max(
          1,
          GRID_ROW_HEADER_WIDTH + resolveAxisSum(currentImage.anchor.to.col, (col) => getColumnWidthPx(worksheet, col)) + currentImage.anchor.to.colOffsetEmu / 9525 - left
        )
      : currentImage.anchor.sizeEmu.cx / 9525;
    const height = currentImage.anchor.kind === "two-cell"
      ? Math.max(
          1,
          GRID_HEADER_HEIGHT + resolveAxisSum(currentImage.anchor.to.row, (row) => getRowHeightPx(worksheet, row)) + currentImage.anchor.to.rowOffsetEmu / 9525 - top
        )
      : currentImage.anchor.sizeEmu.cy / 9525;

    const nextRect = resizeImageRect({ height, left, top, width }, handle, deltaX, deltaY);
    setImageRect(id, nextRect);
  }

  function updateChart(id: string, patch: Partial<XlsxChart>) {
    const currentChart = getChartById(id);
    const hydratedChartAssets = ctx.ensureChartAssetsHydrated(ctx.workbook.value, ctx.sheets.value);
    if (ctx.readOnly.value || !currentChart) {
      return;
    }

    ctx.recordHistoryBeforeMutation();
    if (patch.anchor && ctx.imageAssetsRef.value && hydratedChartAssets) {
      updateWorkbookChartAnchor(ctx.imageAssetsRef.value, hydratedChartAssets, id, patch.anchor);
    }
    if (ctx.imageAssetsRef.value && hydratedChartAssets) {
      updateWorkbookChartDefinition(ctx.imageAssetsRef.value, hydratedChartAssets, id, patch);
    }

    ctx.chartsByWorkbookSheetIndex.value = ctx.chartsByWorkbookSheetIndex.value.map((sheetCharts) => (
      sheetCharts.map((chart) => chart.id === id ? { ...chart, ...patch } : chart)
    ));
    ctx.revision.value = ctx.revision.value + 1;
  }

  function setChartSeriesFormula(chartId: string, seriesIndex: number, formula: string) {
    if (ctx.readOnly.value) {
      return false;
    }

    const chart = getChartById(chartId);
    if (!chart || chart.editable === false) {
      return false;
    }

    const nextChart = applyChartSeriesFormula(chart, seriesIndex, formula, ctx.workbook.value);
    if (!nextChart) {
      return false;
    }

    updateChart(chartId, { series: nextChart.series });
    const selectedSeries = nextChart.series[seriesIndex];
    if (selectedSeries) {
      ctx.selectedChartElement.value = (
        ctx.selectedChartElement.value && ctx.selectedChartElement.value.chartId === chartId && ctx.selectedChartElement.value.kind !== "chart"
          ? {
              ...ctx.selectedChartElement.value,
              seriesId: selectedSeries.id,
              seriesIndex
            }
          : ctx.selectedChartElement.value
      );
    }
    return true;
  }

  function getChartSeriesFormula(chartId: string, seriesIndex: number) {
  return (
    buildChartSeriesFormula(getChartById(chartId), seriesIndex)
  );
}

  return {
    charts,
    selectedChart,
    selectedImage,
    images,
    formControls,
    shapes,
    getSheetCharts,
    getChartById,
    getChartsheetById,
    selectChart,
    clearSelectedChart,
    clearSelectedChartElement,
    selectChartElement,
    getSheetImages,
    getSheetFormControls,
    getSheetShapes,
    getImageById,
    selectImage,
    clearSelectedImage,
    setChartRect,
    setImageRect,
    moveChartBy,
    moveImageBy,
    resizeChartBy,
    resizeImageBy,
    updateChart,
    setChartSeriesFormula,
    getChartSeriesFormula
  };
}
