import { defineAsyncComponent, defineComponent, h } from "vue";

const RendererLoading = defineComponent({
  name: "XlsxOptionalRendererLoading",
  setup() {
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
    }, "正在加载图表…");
  },
});

const RendererError = defineComponent({
  name: "XlsxOptionalRendererError",
  setup() {
    return () => h("div", {
      "data-testid": "xlsx-optional-renderer-error",
      role: "alert",
    }, "图表加载失败，请重试。");
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
