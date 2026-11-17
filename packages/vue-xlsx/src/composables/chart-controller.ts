import { computed, watch } from "vue";
import type { Workbook } from "@dukelib/sheets-wasm";
import {
  addWorkbookFormControl,
  anchorToAbsoluteRect,
  applyChartSeriesFormula,
  buildChartSeriesFormula,
  deleteWorkbookChart,
  deleteWorkbookFormControl,
  deleteWorkbookImage,
  deleteWorkbookShape,
  emuToPixels,
  rectToImageAnchor,
  resizeImageRect,
  resolveRenderedSheetAxisPixels,
  resolveSheetColumnWidthPixels,
  resolveSheetColumnWidthPx,
  resolveSheetRowHeightPixels,
  resolveSheetRowHeightPx,
  updateWorkbookChartAnchor,
  updateWorkbookChartDefinition,
  updateWorkbookFormControl,
  updateWorkbookImageAnchor,
  updateWorkbookShapeAnchor,
  type XlsxChart,
  type XlsxChartElementSelection,
  type XlsxChartLegend,
  type XlsxFormControl,
  type XlsxFormControlInput,
  type XlsxImage,
  type XlsxImageRect,
  type XlsxImageResizeHandlePosition,
  type XlsxShape
} from "@arcships/xlsx-core";
import {
  DEFAULT_COL_WIDTH,
  DEFAULT_ROW_HEIGHT,
  GRID_HEADER_HEIGHT,
  GRID_ROW_HEADER_WIDTH
} from "./internal";
import type { XlsxControllerContext } from "./internal";

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
    for (const chartsheet of ctx.chartsheets.value) {
      for (const chart of chartsheet.charts ?? []) {
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
    ctx.selectedShapeId.value = null;
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
    ctx.selectedShapeId.value = null;
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

  function getFormControlById(id: string) {
    for (const sheetControls of ctx.formControlsByWorkbookSheetIndex.value) {
      const match = sheetControls?.find((control) => control.id === id);
      if (match) {
        return mapPublicFormControl(match);
      }
    }
    return null;
  }

  function getFormControlItems(sheetIndex = ctx.activeSheetIndex.value) {
    return getSheetFormControls(sheetIndex);
  }

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

  function getShapeById(id: string) {
    for (const sheetShapes of ctx.shapesByWorkbookSheetIndex.value) {
      const match = sheetShapes?.find((shape) => shape.id === id);
      if (match) {
        const visibleSheetIndex = ctx.visibleSheetIndexByWorkbookSheetIndex.value.get(match.workbookSheetIndex);
        return { ...match, sheetIndex: visibleSheetIndex ?? match.workbookSheetIndex };
      }
    }
    return null;
  }

  const selectedShape = computed(() => (ctx.selectedShapeId.value ? getShapeById(ctx.selectedShapeId.value) : null));

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
    ctx.selectedShapeId.value = null;
    ctx.selectedImageId.value = id;
  }

  function clearSelectedImage() {
    ctx.selectedImageId.value = null;
  }

  function selectShape(id: string | null) {
    ctx.selectedChartId.value = null;
    ctx.selectedChartElement.value = null;
    ctx.selectedImageId.value = null;
    ctx.selectedShapeId.value = id;
  }

  function clearSelectedShape() {
    ctx.selectedShapeId.value = null;
  }

  function zoomFactor() {
    const scale = ctx.zoomScale.value / 100;
    return Number.isFinite(scale) ? Math.max(0.1, scale) : 1;
  }

  function getColumnWidthPx(worksheet: ReturnType<Workbook["getSheet"]>, col: number) {
    const sheet = ctx.activeSheet.value;
    if (sheet) {
      // The grid renders with display-order widths, so write-back must use the
      // same model (`visibleCols` + `colWidths`) instead of the worksheet API.
      return resolveSheetColumnWidthPx(sheet, col);
    }

    const sheetState = ctx.imageAssetsRef.value?.sheetStatesByWorkbookSheetIndex[ctx.activeSheet.value?.workbookSheetIndex ?? -1] ?? null;
    if (typeof worksheet.isColumnHidden === "function" && worksheet.isColumnHidden(col)) {
      return 0;
    }
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
    const sheet = ctx.activeSheet.value;
    if (sheet) {
      return resolveSheetRowHeightPx(sheet, row);
    }

    const sheetState = ctx.imageAssetsRef.value?.sheetStatesByWorkbookSheetIndex[ctx.activeSheet.value?.workbookSheetIndex ?? -1] ?? null;
    if (typeof worksheet.isRowHidden === "function" && worksheet.isRowHidden(row)) {
      return 0;
    }
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

  function resolveAnchoredObjectRect(anchor: XlsxImage["anchor"]) {
    const rect = anchorToAbsoluteRect(anchor, ctx.activeSheet.value);
    const zoom = zoomFactor();
    return {
      height: Math.max(1, emuToPixels(rect.cy) * zoom),
      left: GRID_ROW_HEADER_WIDTH + emuToPixels(rect.x) * zoom,
      top: GRID_HEADER_HEIGHT + emuToPixels(rect.y) * zoom,
      width: Math.max(1, emuToPixels(rect.cx) * zoom)
    };
  }

  /**
   * Convert an overlay rect (viewer CSS pixels, including zoom) into the
   * unscaled sheet-pixel space that image anchors are written in.
   */
  function toSheetRect(rect: XlsxImageRect): XlsxImageRect {
    const zoom = zoomFactor();
    if (zoom === 1) {
      return rect;
    }
    return {
      height: rect.height / zoom,
      left: GRID_ROW_HEADER_WIDTH + (rect.left - GRID_ROW_HEADER_WIDTH) / zoom,
      top: GRID_HEADER_HEIGHT + (rect.top - GRID_HEADER_HEIGHT) / zoom,
      width: rect.width / zoom
    };
  }

  function setChartRect(id: string, rect: XlsxImageRect) {
    const hydratedChartAssets = ctx.ensureChartAssetsHydrated(ctx.workbook.value, ctx.sheets.value);
    if (ctx.isEditingLocked.value || !ctx.workbook.value || !ctx.activeSheet.value || !ctx.imageAssetsRef.value || !hydratedChartAssets) {
      return;
    }

    const worksheet = ctx.getActiveWorksheet();
    const currentChart = getChartById(id);
    if (!worksheet || !currentChart || currentChart.editable === false || currentChart.workbookSheetIndex !== ctx.activeSheet.value.workbookSheetIndex) {
      return;
    }
    // Validate before snapshotting so a rejected move does not leave an empty
    // undo entry (and dragging never serialises the workbook per frame).
    if (!hydratedChartAssets.chartOriginsById.has(id)) {
      return;
    }

    const nextAnchor = rectToImageAnchor(toSheetRect(rect), currentChart.anchor, {
      contentOffsetLeft: GRID_ROW_HEADER_WIDTH,
      contentOffsetTop: GRID_HEADER_HEIGHT,
      getColumnWidthPx: (col) => getColumnWidthPx(worksheet, col),
      getRowHeightPx: (row) => getRowHeightPx(worksheet, row)
    });

    ctx.recordHistoryBeforeMutation();
    if (!updateWorkbookChartAnchor(ctx.imageAssetsRef.value, hydratedChartAssets, id, nextAnchor)) {
      return;
    }

    hydratedChartAssets.chartsByWorkbookSheetIndex = hydratedChartAssets.chartsByWorkbookSheetIndex.map((sheetCharts) => (
      sheetCharts.map((chart) => chart.id === id ? { ...chart, anchor: nextAnchor } : chart)
    ));

    ctx.chartsByWorkbookSheetIndex.value = ctx.chartsByWorkbookSheetIndex.value.map((sheetCharts) => (
      sheetCharts.map((chart) => chart.id === id ? { ...chart, anchor: nextAnchor } : chart)
    ));
    ctx.revision.value = ctx.revision.value + 1;
  }

  function setImageRect(id: string, rect: XlsxImageRect) {
    if (ctx.isEditingLocked.value || !ctx.workbook.value || !ctx.activeSheet.value || !ctx.imageAssetsRef.value) {
      return;
    }

    const worksheet = ctx.getActiveWorksheet();
    const currentImage = getImageById(id);
    if (!worksheet || !currentImage || currentImage.editable === false || currentImage.workbookSheetIndex !== ctx.activeSheet.value.workbookSheetIndex) {
      return;
    }
    // Validate before snapshotting so a rejected move does not leave an empty
    // undo entry (and dragging never serialises the workbook per frame).
    if (!ctx.imageAssetsRef.value.imageOriginsById.has(id)) {
      return;
    }

    const nextAnchor = rectToImageAnchor(toSheetRect(rect), currentImage.anchor, {
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

    const currentRect = resolveAnchoredObjectRect(currentChart.anchor);
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

    if (!ctx.getActiveWorksheet()) {
      return;
    }

    const currentRect = resolveAnchoredObjectRect(currentImage.anchor);
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

    if (!ctx.getActiveWorksheet()) {
      return;
    }

    const currentRect = resolveAnchoredObjectRect(currentChart.anchor);
    setChartRect(id, resizeImageRect(currentRect, handle, deltaX, deltaY, 48));
  }

  function resizeImageBy(id: string, handle: XlsxImageResizeHandlePosition, deltaX: number, deltaY: number) {
    const currentImage = getImageById(id);
    if (!currentImage || currentImage.editable === false) {
      return;
    }

    if (!ctx.getActiveWorksheet()) {
      return;
    }

    const currentRect = resolveAnchoredObjectRect(currentImage.anchor);
    setImageRect(id, resizeImageRect(currentRect, handle, deltaX, deltaY));
  }

  function setShapeRect(id: string, rect: XlsxImageRect) {
    const assets = ctx.imageAssetsRef.value;
    if (ctx.isEditingLocked.value || !ctx.workbook.value || !ctx.activeSheet.value || !assets) {
      return;
    }
    const currentShape = getShapeById(id);
    if (!currentShape || currentShape.workbookSheetIndex !== ctx.activeSheet.value.workbookSheetIndex) {
      return;
    }
    if (!assets.shapeOriginsById.has(id)) {
      return;
    }
    const worksheet = ctx.getActiveWorksheet();
    if (!worksheet) {
      return;
    }
    const nextAnchor = rectToImageAnchor(toSheetRect(rect), currentShape.anchor, {
      contentOffsetLeft: GRID_ROW_HEADER_WIDTH,
      contentOffsetTop: GRID_HEADER_HEIGHT,
      getColumnWidthPx: (col) => getColumnWidthPx(worksheet, col),
      getRowHeightPx: (row) => getRowHeightPx(worksheet, row)
    });
    ctx.recordHistoryBeforeMutation();
    if (!updateWorkbookShapeAnchor(assets, id, nextAnchor)) {
      return;
    }
    ctx.shapesByWorkbookSheetIndex.value = [...assets.shapesByWorkbookSheetIndex];
    ctx.revision.value = ctx.revision.value + 1;
  }

  function moveShapeBy(id: string, deltaX: number, deltaY: number) {
    const currentShape = getShapeById(id);
    if (!currentShape || !ctx.getActiveWorksheet()) {
      return;
    }
    const currentRect = resolveAnchoredObjectRect(currentShape.anchor);
    setShapeRect(id, {
      ...currentRect,
      left: currentRect.left + deltaX,
      top: currentRect.top + deltaY
    });
  }

  function resizeShapeBy(id: string, handle: XlsxImageResizeHandlePosition, deltaX: number, deltaY: number) {
    const currentShape = getShapeById(id);
    if (!currentShape || !ctx.getActiveWorksheet()) {
      return;
    }
    const currentRect = resolveAnchoredObjectRect(currentShape.anchor);
    setShapeRect(id, resizeImageRect(currentRect, handle, deltaX, deltaY));
  }

  function deleteShape(id: string) {
    const assets = ctx.imageAssetsRef.value;
    if (ctx.isEditingLocked.value || !assets || !assets.shapeOriginsById.has(id)) {
      return false;
    }
    ctx.recordHistoryBeforeMutation();
    if (!deleteWorkbookShape(assets, id)) {
      return false;
    }
    ctx.shapesByWorkbookSheetIndex.value = [...assets.shapesByWorkbookSheetIndex];
    if (ctx.selectedShapeId.value === id) ctx.selectedShapeId.value = null;
    ctx.revision.value = ctx.revision.value + 1;
    return true;
  }

  function deleteImage(id: string) {
    const assets = ctx.imageAssetsRef.value;
    if (ctx.isEditingLocked.value || !assets || !assets.imageOriginsById.has(id)) {
      return false;
    }
    ctx.recordHistoryBeforeMutation();
    if (!deleteWorkbookImage(assets, id)) {
      return false;
    }
    ctx.imagesByWorkbookSheetIndex.value = [...assets.imagesByWorkbookSheetIndex];
    if (ctx.selectedImageId.value === id) ctx.selectedImageId.value = null;
    ctx.revision.value = ctx.revision.value + 1;
    return true;
  }

  function deleteChart(id: string) {
    const hydratedChartAssets = ctx.ensureChartAssetsHydrated(ctx.workbook.value, ctx.sheets.value);
    if (ctx.isEditingLocked.value || !ctx.imageAssetsRef.value || !hydratedChartAssets || !hydratedChartAssets.chartOriginsById.has(id)) {
      return false;
    }
    ctx.recordHistoryBeforeMutation();
    if (!deleteWorkbookChart(ctx.imageAssetsRef.value, hydratedChartAssets, id)) {
      return false;
    }
    ctx.chartsByWorkbookSheetIndex.value = hydratedChartAssets.chartsByWorkbookSheetIndex;
    if (ctx.selectedChartId.value === id) {
      ctx.selectedChartId.value = null;
      ctx.selectedChartElement.value = null;
    }
    ctx.revision.value = ctx.revision.value + 1;
    return true;
  }

  function addFormControl(input: XlsxFormControlInput) {
    const assets = ctx.imageAssetsRef.value;
    const sheet = ctx.activeSheet.value;
    if (
      ctx.isEditingLocked.value
      || !assets
      || !sheet
      || !assets.sheetPathsByWorkbookSheetIndex[sheet.workbookSheetIndex]
    ) {
      return null;
    }
    ctx.recordHistoryBeforeMutation();
    const control = addWorkbookFormControl(assets, sheet.workbookSheetIndex, input);
    if (!control) {
      return null;
    }
    ctx.formControlsByWorkbookSheetIndex.value = [...assets.formControlsByWorkbookSheetIndex];
    ctx.revision.value = ctx.revision.value + 1;
    return control;
  }

  function updateFormControl(id: string, patch: Partial<XlsxFormControlInput>) {
    const assets = ctx.imageAssetsRef.value;
    if (ctx.isEditingLocked.value || !assets || !assets.formControlOriginsById.has(id)) {
      return false;
    }
    ctx.recordHistoryBeforeMutation();
    if (!updateWorkbookFormControl(assets, id, patch)) {
      return false;
    }
    ctx.formControlsByWorkbookSheetIndex.value = [...assets.formControlsByWorkbookSheetIndex];
    ctx.revision.value = ctx.revision.value + 1;
    return true;
  }

  function removeFormControl(id: string) {
    const assets = ctx.imageAssetsRef.value;
    if (ctx.isEditingLocked.value || !assets || !assets.formControlOriginsById.has(id)) {
      return false;
    }
    ctx.recordHistoryBeforeMutation();
    if (!deleteWorkbookFormControl(assets, id)) {
      return false;
    }
    ctx.formControlsByWorkbookSheetIndex.value = [...assets.formControlsByWorkbookSheetIndex];
    ctx.revision.value = ctx.revision.value + 1;
    return true;
  }

  function updateChart(id: string, patch: Partial<XlsxChart>) {
    const currentChart = getChartById(id);
    const hydratedChartAssets = ctx.ensureChartAssetsHydrated(ctx.workbook.value, ctx.sheets.value);
    if (ctx.isEditingLocked.value || !currentChart || currentChart.editable === false) {
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
    if (ctx.isEditingLocked.value) {
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
    selectedShape,
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
    getFormControlById,
    getFormControlItems,
    getSheetShapes,
    getImageById,
    getShapeById,
    selectImage,
    selectShape,
    clearSelectedImage,
    clearSelectedShape,
    setChartRect,
    setImageRect,
    setShapeRect,
    moveChartBy,
    moveImageBy,
    moveShapeBy,
    resizeChartBy,
    resizeImageBy,
    resizeShapeBy,
    deleteChart,
    deleteImage,
    deleteShape,
    addFormControl,
    updateFormControl,
    removeFormControl,
    updateChart,
    setChartSeriesFormula,
    getChartSeriesFormula
  };
}
