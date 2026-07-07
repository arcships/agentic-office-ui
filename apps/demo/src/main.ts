import { createApp } from "vue"
import { createRouter, createWebHashHistory } from "vue-router"
import App from "./App.vue"
import "../../../packages/vue-docx/dist/index.css"
import "../../../packages/vue-extend/dist/index.css"
import "../../../packages/vue-xlsx/dist/index.css"
import HomePage from "./pages/HomePage.vue"
import DocxViewerPage from "./pages/DocxViewerPage.vue"
import DocxEditorPage from "./pages/DocxEditorPage.vue"
import XlsxViewerPage from "./pages/XlsxViewerPage.vue"
import PdfViewerPage from "./pages/PdfViewerPage.vue"
import ComponentsPage from "./pages/ComponentsPage.vue"

const router = createRouter({
  history: createWebHashHistory(),
  routes: [
    { path: "/", component: HomePage },
    { path: "/docx-viewer", component: DocxViewerPage },
    { path: "/docx-editor", component: DocxEditorPage },
    { path: "/xlsx-viewer", component: XlsxViewerPage },
    { path: "/pdf-viewer", component: PdfViewerPage },
    { path: "/components", component: ComponentsPage },
  ],
})

const app = createApp(App)
app.use(router)
app.mount("#app")
