import { createApp } from "vue"
import { createRouter, createWebHashHistory } from "vue-router"
import App from "./App.vue"
import "@arcships/vue-docx/style.css"
import "@arcships/vue-xlsx/style.css"
import "@arcships/vue-extend/style.css"

const router = createRouter({
  history: createWebHashHistory(),
  routes: [
    { path: "/", component: () => import("./pages/HomePage.vue") },
    { path: "/docx-viewer", component: () => import("./pages/DocxViewerPage.vue") },
    { path: "/docx-editor", component: () => import("./pages/DocxEditorPage.vue") },
    { path: "/xlsx-viewer", component: () => import("./pages/XlsxViewerPage.vue") },
    { path: "/pdf-viewer", component: () => import("./pages/PdfViewerPage.vue") },
    { path: "/components", component: () => import("./pages/ComponentsPage.vue") },
    { path: "/runtime-isolation", component: () => import("./pages/RuntimeIsolationPage.vue") },
    { path: "/docx-parity", component: () => import("./pages/DocxParityPage.vue") },
    { path: "/runtime-limits", component: () => import("./pages/RuntimeLimitsPage.vue") },
    { path: "/runtime-worker", component: () => import("./pages/RuntimeWorkerPage.vue") },
  ],
})

const app = createApp(App)
app.use(router)
app.mount("#app")
