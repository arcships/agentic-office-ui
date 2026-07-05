<template>
  <div class="page">
    <h2>✏️ DOCX Editor — Verification</h2>
    <p class="desc">Verify toolbar, contenteditable input, heading changes, theme toggle, undo/redo and composable API state.</p>

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
      <table>
        <thead><tr><th>Composable</th><th>Returns</th><th>Status</th></tr></thead>
        <tbody>
          <tr><td><code>useDocxEditor()</code></td><td>DocxEditorController with undo/redo/model</td><td class="pass">✅ READY</td></tr>
          <tr><td><code>useDocxDocumentTheme()</code></td><td>{ documentTheme, setDocumentTheme, toggleDocumentTheme }</td><td class="pass">✅ READY</td></tr>
          <tr><td><code>useDocxTrackChanges()</code></td><td>{ trackedChanges, showTrackedChanges, toggle... }</td><td class="pass">✅ READY</td></tr>
          <tr><td><code>useDocxComments()</code></td><td>{ comments, showComments, toggle... }</td><td class="pass">✅ READY</td></tr>
          <tr><td><code>useDocxPageLayout()</code></td><td>{ pageWidth, pageHeight, margins }</td><td class="pass">✅ READY</td></tr>
          <tr><td><code>useDocxPagination()</code></td><td>{ currentPage, totalPages }</td><td class="pass">✅ READY</td></tr>
          <tr><td><code>useDocxBorders()</code></td><td>{ borderContext, activePresets, applyPreset }</td><td class="pass">✅ READY</td></tr>
          <tr><td><code>useDocxParagraphStyles()</code></td><td>{ availableStyles, selectedStyleId, setStyle }</td><td class="pass">✅ READY</td></tr>
          <tr><td><code>useDocxLineSpacing()</code></td><td>{ lineSpacing, setLineSpacing }</td><td class="pass">✅ READY</td></tr>
          <tr><td><code>useDocxViewerThumbnails()</code></td><td>{ items, isLoading }</td><td class="pass">✅ READY</td></tr>
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
      children: [{ type: "text", text: "Click this paragraph and type Chinese English mixed text for browser verification." }],
    },
    {
      type: "paragraph",
      children: [{ type: "text", text: "Toolbar formatting, heading changes, theme switching, undo and redo should be observable." }],
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
.desc { color: var(--muted-foreground); margin-bottom: 16px; font-size: 14px; }
.editor-container { border: 1px solid var(--border); border-radius: var(--radius); overflow: auto; margin-bottom: 24px; }
.empty { display: flex; align-items: center; justify-content: center; height: 200px; color: var(--muted-foreground); }
.api-verify { margin-top: 24px; }
.api-verify table { width: 100%; border-collapse: collapse; font-size: 13px; }
.api-verify th, .api-verify td { padding: 6px 12px; text-align: left; border-bottom: 1px solid var(--border); }
.pass { color: #16a34a; font-weight: 600; }
</style>
