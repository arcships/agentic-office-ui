<template>
  <div class="home">
    <h2>Vue Extend UI</h2>
    <p class="subtitle">在 Vue 3 中查看和编辑 DOCX、XLSX 与 PDF 文件。</p>

    <div class="grid">
      <div class="card" v-for="item in navItems" :key="item.path">
        <router-link :to="item.path" class="card-link">
          <h3>{{ item.title }}</h3>
          <p>{{ item.desc }}</p>
          <span class="card-status" :class="item.statusClass">{{ item.status }}</span>
        </router-link>
      </div>
    </div>

    <details class="api-map" @toggle="onApiMapToggle">
      <summary>开发状态与接口对照</summary>
      <table v-if="showApiMap">
        <thead>
          <tr><th>React API</th><th>Vue API</th><th>Status</th></tr>
        </thead>
        <tbody>
          <tr v-for="row in apiMap" :key="row.react">
            <td><code>{{ row.react }}</code></td>
            <td><code>{{ row.vue }}</code></td>
            <td :class="row.status">{{ row.status }}</td>
          </tr>
        </tbody>
      </table>
    </details>
  </div>
</template>

<script setup lang="ts">
import { ref } from "vue"

const showApiMap = ref(false)

function onApiMapToggle(event: Event) {
  showApiMap.value = (event.currentTarget as HTMLDetailsElement).open
}

const navItems = [
  { path: "/docx-surface", title: "DOCX Surface", desc: "Minimal embeddable renderer — no toolbar, no thumbnails. Build your own controls around it.", status: "Ready", statusClass: "ready" },
  { path: "/docx-viewer", title: "DOCX Viewer", desc: "View .docx files with pagination and thumbnails", status: "Ready", statusClass: "ready" },
  { path: "/docx-editor", title: "DOCX Editor", desc: "Edit DOCX text, formatting, tables and positioned images with undo and redo", status: "Ready", statusClass: "ready" },
  { path: "/xlsx-surface", title: "XLSX Surface", desc: "Minimal embeddable grid — no toolbar, ribbon, or formula bar. Bring your own sheet tabs and controls.", status: "Ready", statusClass: "ready" },
  { path: "/xlsx-viewer", title: "XLSX Viewer", desc: "View and edit workbooks with formulas, formatting, search, charts, tables and sheet tabs", status: "Ready", statusClass: "ready" },
  { path: "/pptx-surface", title: "PPTX Surface", desc: "Minimal embeddable surface — all slides stacked vertically, with host-owned controls.", status: "Ready", statusClass: "ready" },
  { path: "/pptx-viewer", title: "PPTX Viewer", desc: "View PPTX with browse/present modes, playback and search", status: "Ready", statusClass: "ready" },
  { path: "/pptx-headless", title: "PPTX 最小组合", desc: "组合文档、播放和舞台三个公开接口", status: "Ready", statusClass: "ready" },
  { path: "/pdf-surface", title: "PDF Surface", desc: "Minimal embeddable PDF — all pages stacked vertically, host-owned zoom and scroll.", status: "Ready", statusClass: "ready" },
  { path: "/pdf-viewer", title: "PDF Viewer", desc: "View PDF with zoom, search, and thumbnails", status: "Ready", statusClass: "ready" },
  { path: "/components", title: "Components", desc: "Signature, FileUpload, BoundingBox, LayoutBlocks", status: "Ready", statusClass: "ready" },
  { path: "/runtime-limits", title: "Runtime Limits", desc: "Configure public DOCX/XLSX resource limits and inspect structured errors", status: "Ready", statusClass: "ready" },
  { path: "/runtime-worker", title: "Worker Lifecycle", desc: "Verify public XLSX Worker timeout, cancellation, termination, and recovery", status: "Ready", statusClass: "ready" },
]

const apiMap = [
  { react: "useDocxEditor(options)", vue: "useDocxEditor(options)", status: "⚠️ scoped" },
  { react: "useDocxComments(editor)", vue: "useDocxComments(editor)", status: "✅ implemented" },
  { react: "useDocxTrackChanges(editor)", vue: "useDocxTrackChanges(editor)", status: "✅ implemented" },
  { react: "useDocxDocumentTheme(editor)", vue: "useDocxDocumentTheme(editor)", status: "✅ implemented" },
  { react: "useDocxPageLayout(editor)", vue: "useDocxPageLayout(editor)", status: "✅ implemented" },
  { react: "useDocxPagination(editor)", vue: "useDocxPagination(editor)", status: "✅ implemented" },
  { react: "useDocxParagraphStyles(editor)", vue: "useDocxParagraphStyles(editor)", status: "✅ implemented" },
  { react: "useDocxLineSpacing(editor)", vue: "useDocxLineSpacing(editor)", status: "✅ implemented" },
  { react: "useDocxBorders(editor)", vue: "useDocxBorders(editor)", status: "✅ implemented" },
  { react: "useDocxViewerThumbnails(editor)", vue: "useDocxViewerThumbnails(editor)", status: "✅ implemented" },
  { react: "useDocxModel(file)", vue: "useDocxModel(file)", status: "✅ implemented" },
  { react: "DocxEditorViewer", vue: "DocxEditorViewer", status: "✅ compatible" },
  { react: "ReactDocxViewer", vue: "DocxViewer", status: "✅ implemented" },
  { react: "XlsxViewer", vue: "XlsxViewer", status: "✅ implemented" },
  { react: "XlsxViewerProvider", vue: "XlsxViewerProvider", status: "✅ implemented" },
  { react: "useXlsxViewer()", vue: "useXlsxViewer()", status: "✅ implemented" },
  { react: "useXlsxViewerController()", vue: "useXlsxViewerController()", status: "✅ implemented" },
]
</script>

<style scoped>
.home { padding: 32px 24px; max-width: 1200px; margin: 0 auto; width: 100%; min-width: 0; }
h2 { font-size: 28px; margin-bottom: 8px; }
.subtitle { color: var(--muted-foreground); margin-bottom: 32px; }
.grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 16px; margin-bottom: 48px; }
.card { border: 1px solid var(--border); border-radius: var(--radius); overflow: auto; transition: box-shadow 0.2s; }
.card:hover { box-shadow: 0 4px 12px rgba(0,0,0,0.1); }
.card-link { display: block; padding: 20px; text-decoration: none; color: inherit; }
.card h3 { margin: 0 0 4px; font-size: 16px; }
.card p { font-size: 13px; color: var(--muted-foreground); }
.card-status { display: inline-block; margin-top: 8px; font-size: 12px; padding: 2px 8px; border-radius: 12px; background: var(--muted); }
.card-status.ready { background: #dcfce7; color: #166534; }
.card-status.scoped { background: #fef3c7; color: #92400e; }
.api-map { margin-top: 24px; }
.api-map { border: 1px solid var(--border); border-radius: var(--radius); padding: 12px; }
.api-map summary { cursor: pointer; font-size: 14px; font-weight: 600; }
.api-map table { margin-top: 12px; }
table { width: 100%; border-collapse: collapse; font-size: 13px; }
th, td { padding: 8px 12px; text-align: left; border-bottom: 1px solid var(--border); }
th { background: var(--muted); font-weight: 600; }
code { background: var(--muted); padding: 2px 6px; border-radius: 4px; font-size: 12px; }
td:last-child { font-weight: 600; }
</style>
