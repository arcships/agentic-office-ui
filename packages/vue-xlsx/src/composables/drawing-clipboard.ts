import { watch, type ShallowRef } from "vue";
import {
  anchorToAbsoluteRect,
  cloneWorkbookDrawing,
  createDrawingClipboardGroup,
  emuToPixels,
  rectToImageAnchor,
  removeWorkbookDrawing,
  resolveSheetColumnWidthPx,
  resolveSheetRowHeightPx,
  type XlsxCellAddress,
  type XlsxChart,
  type XlsxDrawingClipboardItem,
  type XlsxDrawingClipboardKind,
  type XlsxImage,
  type XlsxImageAnchor,
  type XlsxSheetData,
  type XlsxShape,
} from "@arcships/xlsx-core";
import { officeMessagesEn } from "@arcships/office-interaction";
import type { ClipboardPayload, XlsxControllerContext } from "./internal";
import type { ClipboardWritePayload } from "./clipboard-text";

/** A selected drawing model plus the family it belongs to. */
export type XlsxSelectedDrawing = {
  kind: XlsxDrawingClipboardKind;
  model: XlsxChart | XlsxImage | XlsxShape;
};

type StoredDrawingClipboard = {
  cut: boolean;
  payload: ClipboardPayload;
  sheetIndex: number;
};

const DUPLICATE_OFFSET_PX = 20;

function drawingLabel(kind: XlsxDrawingClipboardKind, model: XlsxChart | XlsxImage | XlsxShape) {
  if (kind === "chart") {
    const chart = model as XlsxChart;
    return chart.title || chart.name || officeMessagesEn.xlsx.chart.ariaLabel;
  }
  if (kind === "image") {
    const image = model as XlsxImage;
    return image.name || image.description || officeMessagesEn.common.descriptors.image;
  }
  const shape = model as XlsxShape;
  const text = shape.paragraphs
    .flatMap((paragraph) => paragraph.runs.map((run) => run.text))
    .join(" ")
    .trim();
  return shape.name || text || officeMessagesEn.docx.drawing.shape;
}

function escapeHtml(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * Drawing clipboard for charts, images and shapes.
 *
 * Copy/cut stores a JSON payload in the internal clipboard flavour. Paste
 * clones the DrawingML anchor on the target sheet so the pasted object keeps
 * its geometry, its media/chart part and its relationships, and the whole
 * paste is one undo entry.
 */
export function createDrawingClipboardDomain(ctx: XlsxControllerContext) {
  let stored: StoredDrawingClipboard | null = null;

  watch(() => ctx.workbook.value, () => { stored = null; });

  function findSelectedDrawing(): XlsxSelectedDrawing | null {
    if (ctx.selectedChartId.value) {
      const chart = ctx.chartsByWorkbookSheetIndex.value
        .flat()
        .find((entry) => entry.id === ctx.selectedChartId.value);
      return chart ? { kind: "chart", model: chart } : null;
    }
    if (ctx.selectedImageId.value) {
      const image = ctx.imagesByWorkbookSheetIndex.value
        .flat()
        .find((entry) => entry.id === ctx.selectedImageId.value);
      return image ? { kind: "image", model: image } : null;
    }
    if (ctx.selectedShapeId.value) {
      const shape = ctx.shapesByWorkbookSheetIndex.value
        .flat()
        .find((entry) => entry.id === ctx.selectedShapeId.value);
      return shape ? { kind: "shape", model: shape } : null;
    }
    return null;
  }

  function buildClipboardPayload(drawing: XlsxSelectedDrawing): ClipboardPayload {
    const sheetState = ctx.imageAssetsRef.value?.sheetStatesByWorkbookSheetIndex[
      drawing.model.workbookSheetIndex
    ] ?? null;
    return {
      cells: [],
      cols: 0,
      drawings: createDrawingClipboardGroup([drawing], sheetState),
      merges: [],
      rows: 0,
    };
  }

  function buildClipboardWritePayload(payload: ClipboardPayload): ClipboardWritePayload {
    const item = payload.drawings?.[0];
    const label = item ? drawingLabel(item.kind, item.model) : "Drawing";
    const description = `${item?.kind ?? "drawing"}: ${label}`;
    return {
      html: `<div data-xlsx-clipboard-drawing="${escapeHtml(item?.kind ?? "drawing")}">${escapeHtml(label)}</div>`,
      structured: JSON.stringify(payload),
      text: description,
    };
  }

  function storeSelection(cut: boolean): ClipboardWritePayload | null {
    if (cut && ctx.isEditingLocked.value) return null;
    const drawing = findSelectedDrawing();
    if (!drawing) return null;
    const payload = buildClipboardPayload(drawing);
    stored = {
      cut,
      payload,
      sheetIndex: drawing.model.workbookSheetIndex,
    };
    return buildClipboardWritePayload(payload);
  }

  function hasStored() {
    return Boolean(stored?.payload.drawings?.length);
  }

  function getStored() {
    return stored;
  }

  function clear() {
    stored = null;
  }

  function cancelCut() {
    if (!stored?.cut) return false;
    stored = null;
    return true;
  }

  function matchesStored(payload: ClipboardPayload) {
    const candidate = payload.drawings?.[0];
    const current = stored?.payload.drawings?.[0];
    return Boolean(
      candidate
      && current
      && candidate.kind === current.kind
      && candidate.sourceId === current.sourceId
      && candidate.sourceSheetId === current.sourceSheetId
    );
  }

  function sheetContentPxBefore(sheet: XlsxSheetData, cell: XlsxCellAddress) {
    let left = 0;
    for (let col = 0; col < cell.col; col += 1) {
      left += resolveSheetColumnWidthPx(sheet, col);
    }
    let top = 0;
    for (let row = 0; row < cell.row; row += 1) {
      top += resolveSheetRowHeightPx(sheet, row);
    }
    return { left, top };
  }

  /** Converts an item's relative EMU offset into an anchor on the target sheet. */
  function targetAnchorForItem(
    item: XlsxDrawingClipboardItem,
    sheet: XlsxSheetData,
    cell: XlsxCellAddress,
    extraOffsetPx = { x: 0, y: 0 },
  ): XlsxImageAnchor {
    const origin = sheetContentPxBefore(sheet, cell);
    return rectToImageAnchor({
      left: origin.left + emuToPixels(item.relativeAnchor.x) + extraOffsetPx.x,
      top: origin.top + emuToPixels(item.relativeAnchor.y) + extraOffsetPx.y,
      width: Math.max(1, emuToPixels(item.size.cx)),
      height: Math.max(1, emuToPixels(item.size.cy)),
    }, item.anchor, {
      contentOffsetLeft: 0,
      contentOffsetTop: 0,
      getColumnWidthPx: (col) => resolveSheetColumnWidthPx(sheet, col),
      getRowHeightPx: (row) => resolveSheetRowHeightPx(sheet, row),
    });
  }

  function syncAssets(chartAssets: ReturnType<XlsxControllerContext["ensureChartAssetsHydrated"]>) {
    const imageAssets = ctx.imageAssetsRef.value;
    if (imageAssets) {
      ctx.imagesByWorkbookSheetIndex.value = [...imageAssets.imagesByWorkbookSheetIndex];
      ctx.shapesByWorkbookSheetIndex.value = [...imageAssets.shapesByWorkbookSheetIndex];
      ctx.sheetOriginsRef.value = [...imageAssets.sheetOrigins];
    }
    if (chartAssets) {
      ctx.chartsByWorkbookSheetIndex.value = [...chartAssets.chartsByWorkbookSheetIndex];
      ctx.chartAssetsRef.value = chartAssets;
    }
    ctx.revision.value = ctx.revision.value + 1;
  }

  function selectDrawing(kind: XlsxDrawingClipboardKind, id: string) {
    ctx.selectedChartId.value = kind === "chart" ? id : null;
    ctx.selectedChartElement.value = kind === "chart" ? { chartId: id, kind: "chart" } : null;
    ctx.selectedImageId.value = kind === "image" ? id : null;
    ctx.selectedShapeId.value = kind === "shape" ? id : null;
  }

  function validateSources(
    items: XlsxDrawingClipboardItem[],
    chartAssets: ReturnType<XlsxControllerContext["ensureChartAssetsHydrated"]>,
  ) {
    const imageAssets = ctx.imageAssetsRef.value;
    if (!imageAssets) return false;
    return items.every((item) => {
      if (item.kind === "chart") {
        return Boolean(chartAssets?.chartOriginsById.has(item.sourceId));
      }
      if (item.kind === "image") {
        return imageAssets.imageOriginsById.has(item.sourceId);
      }
      return imageAssets.shapeOriginsById.has(item.sourceId);
    });
  }

  function pasteItems(
    items: XlsxDrawingClipboardItem[],
    options: {
      cut?: boolean;
      extraOffsetPx?: { x: number; y: number };
    } = {},
  ) {
    if (ctx.isEditingLocked.value || !ctx.workbook.value) return false;
    const sheet = ctx.activeSheet.value;
    const targetCell = ctx.activeCell.value ?? ctx.selection.value?.start ?? null;
    const imageAssets = ctx.imageAssetsRef.value;
    if (!sheet || !targetCell || !imageAssets || items.length === 0) return false;

    const needsChartAssets = items.some((item) => item.kind === "chart");
    const chartAssets = needsChartAssets
      ? ctx.ensureChartAssetsHydrated(ctx.workbook.value, ctx.sheets.value)
      : ctx.chartAssetsRef.value;
    if (needsChartAssets && !chartAssets) return false;
    if (!validateSources(items, chartAssets)) return false;

    const anchors = items.map((item) => targetAnchorForItem(
      item,
      sheet,
      targetCell,
      options.extraOffsetPx,
    ));

    ctx.recordHistoryBeforeMutation();
    const pasted: Array<{ kind: XlsxDrawingClipboardKind; id: string }> = [];
    items.forEach((item, index) => {
      const result = cloneWorkbookDrawing(
        imageAssets,
        chartAssets,
        {
          id: item.sourceId,
          kind: item.kind,
          workbookSheetIndex: item.sourceSheetId,
        },
        sheet.workbookSheetIndex,
        anchors[index],
        { visibleSheetIndex: ctx.activeSheetIndex.value },
      );
      if (!result) return;
      pasted.push({ kind: result.kind, id: result.id });
      if (options.cut) {
        removeWorkbookDrawing(imageAssets, chartAssets, {
          id: item.sourceId,
          kind: item.kind,
          workbookSheetIndex: item.sourceSheetId,
        });
      }
    });

    if (pasted.length === 0) return false;
    syncAssets(chartAssets);
    const last = pasted[pasted.length - 1];
    selectDrawing(last.kind, last.id);
    if (options.cut) {
      stored = null;
    }
    return true;
  }

  function paste(payload: ClipboardPayload, options: { cut?: boolean } = {}) {
    const items = payload.drawings;
    if (!Array.isArray(items) || items.length === 0) return false;
    return pasteItems(items, { cut: options.cut });
  }

  function pasteStored() {
    if (!stored) return false;
    return paste(stored.payload, { cut: stored.cut });
  }

  function duplicate() {
    if (ctx.isEditingLocked.value) return false;
    const drawing = findSelectedDrawing();
    if (!drawing) return false;
    const sheetState = ctx.imageAssetsRef.value?.sheetStatesByWorkbookSheetIndex[
      drawing.model.workbookSheetIndex
    ] ?? null;
    const items = createDrawingClipboardGroup([drawing], sheetState);
    return pasteItems(items, {
      extraOffsetPx: { x: DUPLICATE_OFFSET_PX, y: DUPLICATE_OFFSET_PX },
    });
  }

  function getClipboardData() {
    const drawing = findSelectedDrawing();
    return drawing ? buildClipboardWritePayload(buildClipboardPayload(drawing)) : null;
  }

  return {
    cancelCut,
    clear,
    copyDrawing: () => storeSelection(false),
    cutDrawing: () => storeSelection(true),
    duplicate,
    findSelectedDrawing,
    getClipboardData,
    getStored,
    hasStored,
    matchesStored,
    paste,
    pasteStored,
  };
}
