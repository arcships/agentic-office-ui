import { createApp } from "vue"
import { createRouter, createWebHashHistory } from "vue-router"
import App from "./App.vue"
import "@extend-ai/vue-docx/style.css"
import "@extend-ai/vue-xlsx/style.css"
import "@extend-ai/vue-extend/style.css"
import HomePage from "./pages/HomePage.vue"
import DocxViewerPage from "./pages/DocxViewerPage.vue"
import DocxEditorPage from "./pages/DocxEditorPage.vue"
import XlsxViewerPage from "./pages/XlsxViewerPage.vue"
import PdfViewerPage from "./pages/PdfViewerPage.vue"
import ComponentsPage from "./pages/ComponentsPage.vue"
import RuntimeIsolationPage from "./pages/RuntimeIsolationPage.vue"
import DocxParityPage from "./pages/DocxParityPage.vue"

const router = createRouter({
  history: createWebHashHistory(),
  routes: [
    { path: "/", component: HomePage },
    { path: "/docx-viewer", component: DocxViewerPage },
    { path: "/docx-editor", component: DocxEditorPage },
    { path: "/xlsx-viewer", component: XlsxViewerPage },
    { path: "/pdf-viewer", component: PdfViewerPage },
    { path: "/components", component: ComponentsPage },
    { path: "/runtime-isolation", component: RuntimeIsolationPage },
    { path: "/docx-parity", component: DocxParityPage },
  ],
})

const app = createApp(App)
app.use(router)
app.mount("#app")
