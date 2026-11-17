import { defineAsyncComponent, defineComponent, h } from "vue";
import { useXlsxMessages } from "../i18n";

const RendererLoading = defineComponent({
  name: "XlsxOptionalRendererLoading",
  setup() {
    const m = useXlsxMessages();
    return () => h("div", {
      "aria-live": "polite",
      "data-testid": "xlsx-optional-renderer-loading",
      role: "status",
      style: {
        alignItems: "center",
        display: "flex",
        height: "100%",
        justifyContent: "center",
        width: "100%",
      },
    }, m.value.xlsx.chart.loading);
  },
});

const RendererError = defineComponent({
  name: "XlsxOptionalRendererError",
  setup() {
    const m = useXlsxMessages();
    return () => h("div", {
      "data-testid": "xlsx-optional-renderer-error",
      role: "alert",
    }, m.value.xlsx.chart.loadFailed);
  },
});

export const MemoChartSvg = defineAsyncComponent({
  loader: () => import("./chart").then((module) => module.MemoChartSvg),
  loadingComponent: RendererLoading,
  errorComponent: RendererError,
  delay: 0,
  suspensible: false,
  timeout: 30_000,
});

export const MemoSurfaceChartComposite = defineAsyncComponent({
  loader: () => import("./webgl").then((module) => module.MemoSurfaceChartComposite),
  loadingComponent: RendererLoading,
  errorComponent: RendererError,
  delay: 0,
  suspensible: false,
  timeout: 30_000,
});
