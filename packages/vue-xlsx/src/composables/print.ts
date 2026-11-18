import { computed, nextTick, ref, watch, type ComputedRef, type Ref } from "vue";
import {
  buildWorksheetPrintLayout,
  type XlsxPrintLayout,
  type XlsxPrintPage,
  type XlsxSheetData,
} from "@arcships/xlsx-core";

export interface XlsxPrintDomain {
  closePrintPreview: () => void;
  getPrintLayout: () => XlsxPrintLayout | null;
  isPageBreakPreview: Ref<boolean>;
  isPrintPreviewOpen: Ref<boolean>;
  openPrintPreview: () => boolean;
  print: () => boolean;
  printLayout: ComputedRef<XlsxPrintLayout | null>;
  printPages: ComputedRef<XlsxPrintPage[]>;
  setPageBreakPreview: (enabled: boolean) => void;
}

export interface XlsxPrintDomainContext {
  activeSheet: ComputedRef<XlsxSheetData | null>;
  revision: Ref<number>;
}

function browserWindow(): (Window & typeof globalThis) | null {
  return typeof window === "undefined" ? null : window;
}

/**
 * Print state shared by the controller, the grid page-break overlay and the
 * print-preview surface.
 *
 * Pagination itself lives in `@arcships/xlsx-core`; this domain only exposes the
 * active sheet's layout and the preview/print actions.
 */
export function createPrintDomain(ctx: XlsxPrintDomainContext): XlsxPrintDomain {
  const printLayout = computed<XlsxPrintLayout | null>(() => {
    // A cell edit can replace the sheet object without replacing its range;
    // reading `revision` keeps the layout in sync with both cases.
    void ctx.revision.value;
    const sheet = ctx.activeSheet.value;
    return sheet ? buildWorksheetPrintLayout(sheet) : null;
  });
  const printPages = computed(() => printLayout.value?.pages ?? []);
  const isPageBreakPreview = ref(false);
  const isPrintPreviewOpen = ref(false);

  watch(ctx.activeSheet, (sheet) => {
    if (sheet) return;
    isPageBreakPreview.value = false;
    isPrintPreviewOpen.value = false;
  });

  function getPrintLayout() {
    return printLayout.value;
  }

  function setPageBreakPreview(enabled: boolean) {
    isPageBreakPreview.value = Boolean(enabled) && Boolean(ctx.activeSheet.value);
  }

  function openPrintPreview() {
    if (!ctx.activeSheet.value) return false;
    isPrintPreviewOpen.value = true;
    return true;
  }

  function closePrintPreview() {
    isPrintPreviewOpen.value = false;
  }

  function print() {
    if (!ctx.activeSheet.value) return false;
    const target = browserWindow();
    if (typeof target?.print !== "function") return false;
    isPrintPreviewOpen.value = true;
    void nextTick(() => target.print());
    return true;
  }

  return {
    closePrintPreview,
    getPrintLayout,
    isPageBreakPreview,
    isPrintPreviewOpen,
    openPrintPreview,
    print,
    printLayout,
    printPages,
    setPageBreakPreview,
  };
}
