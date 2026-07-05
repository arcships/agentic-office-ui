<template>
  <div class="home">
    <h2>Vue Extend UI — Verification Demo</h2>
    <p class="subtitle">1:1 API compatible Vue 3 port of Extend UI document components</p>

    <div class="grid">
      <div class="card" v-for="item in navItems" :key="item.path">
        <router-link :to="item.path" class="card-link">
          <span class="card-icon">{{ item.icon }}</span>
          <h3>{{ item.title }}</h3>
          <p>{{ item.desc }}</p>
          <span class="card-status" :class="item.status">{{ item.status }}</span>
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
  { path: "/docx-viewer", icon: "📄", title: "DOCX Viewer", desc: "View .docx files with pagination and thumbnails", status: "✅ Ready" },
  { path: "/docx-editor", icon: "✏️", title: "DOCX Editor", desc: "Edit .docx with formatting toolbar and track changes", status: "✅ Ready" },
  { path: "/xlsx-viewer", icon: "📊", title: "XLSX Viewer", desc: "View .xlsx workbooks with sheet tabs and search", status: "✅ Ready" },
  { path: "/pdf-viewer", icon: "📑", title: "PDF Viewer", desc: "View PDF with zoom, search, and thumbnails", status: "✅ Ready" },
  { path: "/components", icon: "🧩", title: "Components", desc: "Signature, FileUpload, BoundingBox, LayoutBlocks", status: "✅ Ready" },
]

const apiMap = [
  { react: "useDocxEditor(options)", vue: "useDocxEditor(options)", status: "✅ 1:1" },
  { react: "useDocxComments(editor)", vue: "useDocxComments(editor)", status: "✅ 1:1" },
  { react: "useDocxTrackChanges(editor)", vue: "useDocxTrackChanges(editor)", status: "✅ 1:1" },
  { react: "useDocxDocumentTheme(editor)", vue: "useDocxDocumentTheme(editor)", status: "✅ 1:1" },
  { react: "useDocxPageLayout(editor)", vue: "useDocxPageLayout(editor)", status: "✅ 1:1" },
  { react: "useDocxPagination(editor)", vue: "useDocxPagination(editor)", status: "✅ 1:1" },
  { react: "useDocxParagraphStyles(editor)", vue: "useDocxParagraphStyles(editor)", status: "✅ 1:1" },
  { react: "useDocxLineSpacing(editor)", vue: "useDocxLineSpacing(editor)", status: "✅ 1:1" },
  { react: "useDocxBorders(editor)", vue: "useDocxBorders(editor)", status: "✅ 1:1" },
  { react: "useDocxViewerThumbnails(editor)", vue: "useDocxViewerThumbnails(editor)", status: "✅ 1:1" },
  { react: "useDocxModel(file)", vue: "useDocxModel(file)", status: "✅ 1:1" },
  { react: "DocxEditorViewer", vue: "DocxEditorViewer", status: "✅ 1:1" },
  { react: "ReactDocxViewer", vue: "DocxViewer", status: "✅ 1:1" },
  { react: "XlsxViewer", vue: "XlsxViewer", status: "✅ 1:1" },
  { react: "XlsxViewerProvider", vue: "XlsxViewerProvider", status: "✅ 1:1" },
  { react: "useXlsxViewer()", vue: "useXlsxViewer()", status: "✅ 1:1" },
  { react: "useXlsxViewerController()", vue: "useXlsxViewerController()", status: "✅ 1:1" },
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
.card-icon { font-size: 32px; }
.card h3 { margin: 8px 0 4px; font-size: 16px; }
.card p { font-size: 13px; color: var(--muted-foreground); }
.card-status { display: inline-block; margin-top: 8px; font-size: 12px; padding: 2px 8px; border-radius: 12px; background: var(--muted); }
.card-status.Ready { background: #dcfce7; color: #166534; }
.api-map { margin-top: 24px; }
.api-map h3 { margin-bottom: 12px; }
table { width: 100%; border-collapse: collapse; font-size: 13px; }
th, td { padding: 8px 12px; text-align: left; border-bottom: 1px solid var(--border); }
th { background: var(--muted); font-weight: 600; }
code { background: var(--muted); padding: 2px 6px; border-radius: 4px; font-size: 12px; }
td:last-child { font-weight: 600; }
</style>
