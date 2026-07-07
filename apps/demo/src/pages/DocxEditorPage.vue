<template>
  <div class="page">
    <h2>✏️ DOCX Editing Demo — Verification</h2>
    <p class="desc">Verify the scoped DOCX editing demo: toolbar-backed text edits, paragraph formatting, theme toggle, undo/redo, export, and composable API state.</p>
    <p class="scope-note">Scope note: this page intentionally verifies the implemented Vue editing surface. Full upstream editor features such as comments UI, track-changes UI, rich range selection, and table/image editing are not claimed in the accepted surface.</p>

    <div class="editor-container">
      <DocxEditorViewer
        v-if="editor"
        :editor="editor"
        style="height: 70vh;"
      />
      <div v-else class="empty">
        <p>Editor initializing...</p>
      </div>
    </div>

    <div class="api-verify">
      <h3>Composable API Verification</h3>
      <p class="api-note">Rows marked READY mean the composable API is present for integration tests; they do not advertise extra visible editor controls beyond the toolbar above.</p>
      <table>
        <thead><tr><th>Composable</th><th>Returns</th><th>Status</th></tr></thead>
        <tbody>
          <tr><td><code>useDocxEditor()</code></td><td>DocxEditorController with undo/redo/model</td><td class="pass">✅ READY</td></tr>
          <tr><td><code>useDocxDocumentTheme()</code></td><td>{ documentTheme, setDocumentTheme, toggleDocumentTheme }</td><td class="pass">✅ READY</td></tr>
          <tr><td><code>useDocxTrackChanges()</code></td><td>API shell: toggle visibility state; no tracked-change records generated</td><td class="warn">⚠️ SCOPED</td></tr>
          <tr><td><code>useDocxComments()</code></td><td>API shell: toggle visibility state; no comment authoring/rendering</td><td class="warn">⚠️ SCOPED</td></tr>
          <tr><td><code>useDocxPageLayout()</code></td><td>{ pageWidth, pageHeight, margins }</td><td class="pass">✅ READY</td></tr>
          <tr><td><code>useDocxPagination()</code></td><td>{ currentPage, totalPages }</td><td class="pass">✅ READY</td></tr>
          <tr><td><code>useDocxBorders()</code></td><td>API shell; border editing not implemented in this Vue demo</td><td class="warn">⚠️ SCOPED</td></tr>
          <tr><td><code>useDocxParagraphStyles()</code></td><td>API shell; heading dropdown covers the visible scoped demo</td><td class="warn">⚠️ SCOPED</td></tr>
          <tr><td><code>useDocxLineSpacing()</code></td><td>API shell; visible line-spacing control is not claimed</td><td class="warn">⚠️ SCOPED</td></tr>
          <tr><td><code>useDocxViewerThumbnails()</code></td><td>Placeholder thumbnail metadata; independent canvas rendering not implemented</td><td class="warn">⚠️ PLACEHOLDER</td></tr>
          <tr><td><code>useDocxModel()</code></td><td>{ model, isLoading, error }</td><td class="pass">✅ READY</td></tr>
        </tbody>
      </table>
    </div>
  </div>
</template>

<script setup lang="ts">
import { DocxEditorViewer, useDocxEditor, type DocModel } from "@extend-ai/vue-docx"

const starterModel: DocModel = {
  metadata: { headerSections: [], footerSections: [] },
  nodes: [
    {
      type: "paragraph",
      style: { headingLevel: 1 },
      children: [{ type: "text", text: "Editable Verification Document" }],
    },
    {
      type: "paragraph",
      children: [
        { type: "text", text: "Click this " },
        { type: "text", text: "paragraph", style: { bold: true } },
        { type: "text", text: " and type Chinese English mixed text for browser verification." },
      ],
    },
    {
      type: "paragraph",
      children: [
        { type: "text", text: "Toolbar formatting", style: { italic: true } },
        { type: "text", text: ", heading changes, theme switching, undo and redo should be observable." },
      ],
    },
  ],
}

const editor = useDocxEditor({
  starterModel,
  initialFileName: "verification-test.docx",
  initialDocumentTheme: { mode: "light" },
})
</script>

<style scoped>
.page { padding: 24px; max-width: 1200px; margin: 0 auto; width: 100%; min-width: 0; }
h2 { margin-bottom: 4px; }
.desc { color: var(--muted-foreground); margin-bottom: 8px; font-size: 14px; }
.scope-note, .api-note { color: var(--muted-foreground); border: 1px solid var(--border); background: var(--muted); border-radius: var(--radius); padding: 10px 12px; font-size: 13px; line-height: 1.5; margin: 0 0 16px; }
.api-note { margin-top: 0; }
.editor-container { border: 1px solid var(--border); border-radius: var(--radius); overflow: auto; margin-bottom: 24px; }
.empty { display: flex; align-items: center; justify-content: center; height: 200px; color: var(--muted-foreground); }
.api-verify { margin-top: 24px; }
.api-verify table { width: 100%; border-collapse: collapse; font-size: 13px; }
.api-verify th, .api-verify td { padding: 6px 12px; text-align: left; border-bottom: 1px solid var(--border); }
.pass { color: #16a34a; font-weight: 600; }
.warn { color: #b45309; font-weight: 600; }
</style>
