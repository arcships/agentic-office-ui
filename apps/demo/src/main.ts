import { createApp } from "vue"
import { createRouter, createWebHashHistory } from "vue-router"
import App from "./App.vue"
import { applyDemoTheme } from "./lib/demo-visual-settings"
import "@arcships/vue-docx/style.css"
import "@arcships/vue-xlsx/style.css"
import "@arcships/vue-pptx/style.css"
import "@arcships/vue-pdf/style.css"
import "@arcships/vue-ui/style.css"
import "@arcships/vue-ui/theme.css"

// Apply `?theme=dark` before the app mounts so the first paint matches the
// captured state and no light flash can leak into a screenshot.
applyDemoTheme()

const router = createRouter({
  history: createWebHashHistory(),
  routes: [
    { path: "/", component: () => import("./pages/HomePage.vue") },
    { path: "/docx-surface", component: () => import("./pages/DocxSurfacePage.vue") },
    { path: "/docx-viewer", component: () => import("./pages/DocxViewerPage.vue") },
    { path: "/docx-editor", component: () => import("./pages/DocxEditorPage.vue") },
    { path: "/xlsx-surface", component: () => import("./pages/XlsxSurfacePage.vue") },
    { path: "/xlsx-viewer", component: () => import("./pages/XlsxViewerPage.vue") },
    { path: "/csv-viewer", component: () => import("./pages/CsvViewerPage.vue") },
    { path: "/pptx-surface", component: () => import("./pages/PptxSurfacePage.vue") },
    { path: "/pptx-viewer", component: () => import("./pages/PptxViewerPage.vue") },
    { path: "/pptx-headless", component: () => import("./pages/PptxHeadlessPage.vue") },
    { path: "/pdf-surface", component: () => import("./pages/PdfSurfacePage.vue") },
    { path: "/pdf-viewer", component: () => import("./pages/PdfViewerPage.vue") },
    { path: "/components", component: () => import("./pages/ComponentsPage.vue") },
    { path: "/file-components", component: () => import("./pages/FileComponentsPage.vue") },
    { path: "/file-system", component: () => import("./pages/FileSystemPage.vue") },
    { path: "/schema-builder", component: () => import("./pages/SchemaBuilderPage.vue") },
    { path: "/primitives", component: () => import("./pages/PrimitivesPage.vue") },
    { path: "/chrome", component: () => import("./pages/ChromePage.vue") },
    { path: "/runtime-isolation", component: () => import("./pages/RuntimeIsolationPage.vue") },
    { path: "/docx-parity", component: () => import("./pages/DocxParityPage.vue") },
    { path: "/runtime-limits", component: () => import("./pages/RuntimeLimitsPage.vue") },
    { path: "/runtime-worker", component: () => import("./pages/RuntimeWorkerPage.vue") },
  ],
})

const app = createApp(App)
app.use(router)
app.mount("#app")
