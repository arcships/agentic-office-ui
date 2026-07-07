<template>
  <div class="home">
    <h2>Vue Extend UI — Verification Demo</h2>
    <p class="subtitle">Vue 3 port verification demo; parity is partial and explicitly tracked against upstream Extend UI</p>

    <div class="grid">
      <div class="card" v-for="item in navItems" :key="item.path">
        <router-link :to="item.path" class="card-link">
          <h3>{{ item.title }}</h3>
          <p>{{ item.desc }}</p>
          <span class="card-status" :class="item.statusClass">{{ item.status }}</span>
        </router-link>
      </div>
    </div>

    <div class="api-map">
      <h3>API Compatibility Map</h3>
      <table>
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
    </div>
  </div>
</template>

<script setup lang="ts">
const navItems = [
  { path: "/docx-viewer", title: "DOCX Viewer", desc: "View .docx files with pagination and thumbnails", status: "Ready", statusClass: "ready" },
  { path: "/docx-editor", title: "DOCX Editing Demo", desc: "Scoped DOCX text/paragraph editing demo; not full upstream editor parity", status: "Scoped", statusClass: "scoped" },
  { path: "/xlsx-viewer", title: "XLSX Viewer", desc: "View/edit .xlsx workbooks with ribbon, sheet tabs, selection and scoped editing", status: "Scoped", statusClass: "scoped" },
  { path: "/pdf-viewer", title: "PDF Viewer", desc: "View PDF with zoom, search, and thumbnails", status: "Ready", statusClass: "ready" },
  { path: "/components", title: "Components", desc: "Signature, FileUpload, BoundingBox, LayoutBlocks", status: "Ready", statusClass: "ready" },
]

const apiMap = [
  { react: "useDocxEditor(options)", vue: "useDocxEditor(options)", status: "⚠️ scoped" },
  { react: "useDocxComments(editor)", vue: "useDocxComments(editor)", status: "⚠️ shell/no comments" },
  { react: "useDocxTrackChanges(editor)", vue: "useDocxTrackChanges(editor)", status: "⚠️ shell/no tracked changes" },
  { react: "useDocxDocumentTheme(editor)", vue: "useDocxDocumentTheme(editor)", status: "✅ implemented" },
  { react: "useDocxPageLayout(editor)", vue: "useDocxPageLayout(editor)", status: "⚠️ fixed defaults" },
  { react: "useDocxPagination(editor)", vue: "useDocxPagination(editor)", status: "⚠️ basic" },
  { react: "useDocxParagraphStyles(editor)", vue: "useDocxParagraphStyles(editor)", status: "⚠️ API shell" },
  { react: "useDocxLineSpacing(editor)", vue: "useDocxLineSpacing(editor)", status: "⚠️ API shell" },
  { react: "useDocxBorders(editor)", vue: "useDocxBorders(editor)", status: "⚠️ API shell" },
  { react: "useDocxViewerThumbnails(editor)", vue: "useDocxViewerThumbnails(editor)", status: "⚠️ placeholder" },
  { react: "useDocxModel(file)", vue: "useDocxModel(file)", status: "✅ implemented" },
  { react: "DocxEditorViewer", vue: "DocxEditorViewer", status: "⚠️ scoped" },
  { react: "ReactDocxViewer", vue: "DocxViewer", status: "⚠️ scoped" },
  { react: "XlsxViewer", vue: "XlsxViewer", status: "⚠️ scoped" },
  { react: "XlsxViewerProvider", vue: "XlsxViewerProvider", status: "⚠️ partial" },
  { react: "useXlsxViewer()", vue: "useXlsxViewer()", status: "⚠️ partial" },
  { react: "useXlsxViewerController()", vue: "useXlsxViewerController()", status: "⚠️ partial" },
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
.api-map h3 { margin-bottom: 12px; }
table { width: 100%; border-collapse: collapse; font-size: 13px; }
th, td { padding: 8px 12px; text-align: left; border-bottom: 1px solid var(--border); }
th { background: var(--muted); font-weight: 600; }
code { background: var(--muted); padding: 2px 6px; border-radius: 4px; font-size: 12px; }
td:last-child { font-weight: 600; }
</style>
