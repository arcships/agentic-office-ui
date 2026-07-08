<template>
  <div class="page">
    <h2>✏️ DOCX Editing Demo</h2>
    <p class="desc">
      Full-featured DOCX editor demo aligned with upstream playground (26 features):
      undo/redo, paragraph styles, font formatting, lists, alignment, theme toggle,
      import/export, and more. Keyboard shortcuts: ⌘B/I/U for bold/italic/underline,
      ⌘Z/⇧⌘Z for undo/redo, ⌘S for export.
    </p>

    <!-- Toolbar: external file import + status controls -->
    <div class="controls control-panel">
      <button @click="onImport" class="btn-primary">📂 Import .docx</button>
      <button @click="editor?.exportDocx()" class="btn-primary">💾 Export .docx</button>
      <button @click="editor?.newDocument()" class="btn-secondary">✨ New</button>

      <span class="control-separator" />

      <label class="toggle-label">
        <input type="checkbox" :checked="showThumbnails" @change="showThumbnails = ($event.target as HTMLInputElement).checked" />
        Thumbnails
      </label>
      <label class="toggle-label">
        <input type="checkbox" :checked="showTrackedChanges" @change="toggleShowTrackedChanges" />
        Track Changes
      </label>
      <label class="toggle-label">
        <input type="checkbox" :checked="showComments" @change="toggleShowComments" />
        Comments
      </label>
      <label class="toggle-label">
        <input type="checkbox" :checked="isReadOnly" @change="isReadOnly = ($event.target as HTMLInputElement).checked" />
        Read-only
      </label>
    </div>

    <!-- Editor -->
    <div class="editor-container">
      <DocxEditorViewer
        v-if="editor"
        :editor="editor"
        :editable="!isReadOnly"
        :show-toolbar="true"
        :show-thumbnails="showThumbnails"
        style="height: 70vh;"
      />
      <div v-else class="empty">
        <p>Editor initializing...</p>
      </div>
    </div>

    <!-- Feature verification status -->
    <div class="api-verify">
      <h3>Feature Verification — 26 Items</h3>
      <p class="api-note">
        All features are provided by the DocxEditorViewer toolbar and composables.
        Rows marked ✅ are implemented; ⚠️ are stub/placeholder.
      </p>
      <table>
        <thead><tr><th>#</th><th>Feature</th><th>Status</th></tr></thead>
        <tbody>
          <tr><td>1</td><td>Document theme (light/dark)</td><td class="pass">✅</td></tr>
          <tr><td>2</td><td>App theme (light/dark/system)</td><td class="pass">✅</td></tr>
          <tr><td>3</td><td>Undo / Redo</td><td class="pass">✅</td></tr>
          <tr><td>4</td><td>Paragraph style select</td><td class="pass">✅</td></tr>
          <tr><td>5</td><td>Font family (Calibri/Arial/TNR/Georgia/Helvetica/Courier)</td><td class="warn">⚠️</td></tr>
          <tr><td>6</td><td>Font size (8–48pt)</td><td class="warn">⚠️</td></tr>
          <tr><td>7</td><td>Line spacing (1/1.15/1.2/1.5/2/2.5/3)</td><td class="warn">⚠️</td></tr>
          <tr><td>8</td><td>Bold / Italic / Underline / Strikethrough</td><td class="pass">✅</td></tr>
          <tr><td>9</td><td>Superscript / Subscript</td><td class="pass">✅</td></tr>
          <tr><td>10</td><td>Text color + highlight color</td><td class="warn">⚠️</td></tr>
          <tr><td>11</td><td>Hyperlink edit / remove</td><td class="warn">⚠️</td></tr>
          <tr><td>12</td><td>Alignment (Left/Center/Right/Justify)</td><td class="pass">✅</td></tr>
          <tr><td>13</td><td>Bullet / Numbered lists</td><td class="pass">✅</td></tr>
          <tr><td>14</td><td>Columns display</td><td class="warn">⚠️</td></tr>
          <tr><td>15</td><td>Page thumbnails (virtual scroll)</td><td class="warn">⚠️</td></tr>
          <tr><td>16</td><td>Borders (13 presets)</td><td class="warn">⚠️</td></tr>
          <tr><td>17</td><td>Insert image</td><td class="warn">⚠️</td></tr>
          <tr><td>18</td><td>Insert table</td><td class="warn">⚠️</td></tr>
          <tr><td>19</td><td>Zoom (50–200%)</td><td class="warn">⚠️</td></tr>
          <tr><td>20</td><td>Import .docx / .doc</td><td class="pass">✅</td></tr>
          <tr><td>21</td><td>Export .docx</td><td class="pass">✅</td></tr>
          <tr><td>22</td><td>Track changes toggle</td><td class="warn">⚠️</td></tr>
          <tr><td>23</td><td>Comments toggle</td><td class="warn">⚠️</td></tr>
          <tr><td>24</td><td>Read-only mode toggle</td><td class="pass">✅</td></tr>
          <tr><td>25</td><td>Context menu (cut/copy/paste/table ops)</td><td class="warn">⚠️</td></tr>
          <tr><td>26</td><td>Form field config (double-click dialog)</td><td class="warn">⚠️</td></tr>
        </tbody>
      </table>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref } from "vue"
import { DocxEditorViewer, useDocxEditor, useDocxTrackChanges, useDocxComments } from "@extend-ai/vue-docx"
import type { DocModel } from "@extend-ai/docx-core"

// ── Starter model ────────────────────────────────────────────────────
const starterModel: DocModel = {
  metadata: { headerSections: [], footerSections: [] },
  nodes: [
    {
      type: "paragraph",
      style: { headingLevel: 1 },
      children: [{ type: "text", text: "DOCX Editor Verification Document" }],
    },
    {
      type: "paragraph",
      children: [
        { type: "text", text: "This document verifies the DOCX editor surface. " },
        { type: "text", text: "Bold text", style: { bold: true } },
        { type: "text", text: ", " },
        { type: "text", text: "italic text", style: { italic: true } },
        { type: "text", text: ", " },
        { type: "text", text: "underlined text", style: { underline: true } },
        { type: "text", text: ", and mixed Chinese 中文 + English content." },
      ],
    },
    {
      type: "paragraph",
      children: [
        { type: "text", text: "Try editing this paragraph: click and type. Use the toolbar to apply formatting, change headings, toggle lists, and switch themes. Keyboard shortcuts are enabled." },
      ],
    },
    {
      type: "paragraph",
      style: { headingLevel: 2 },
      children: [{ type: "text", text: "Features to Verify" }],
    },
    {
      type: "paragraph",
      children: [
        { type: "text", text: "1. Undo/Redo — press ⌘Z / ⇧⌘Z after editing" },
      ],
    },
    {
      type: "paragraph",
      children: [
        { type: "text", text: "2. Heading styles — use the dropdown in toolbar" },
      ],
    },
    {
      type: "paragraph",
      children: [
        { type: "text", text: "3. Bold/Italic/Underline/Strikethrough — use toolbar buttons" },
      ],
    },
    {
      type: "paragraph",
      children: [
        { type: "text", text: "4. Superscript/Subscript — e.g. E=mc² or H₂O" },
      ],
    },
    {
      type: "paragraph",
      children: [
        { type: "text", text: "5. Alignment — Left/Center/Right/Justify" },
      ],
    },
    {
      type: "paragraph",
      children: [
        { type: "text", text: "6. Lists — bullet and numbered lists" },
      ],
    },
    {
      type: "paragraph",
      children: [
        { type: "text", text: "7. Theme toggle — 🌙/☀️ button in toolbar" },
      ],
    },
    {
      type: "paragraph",
      children: [
        { type: "text", text: "8. Import/Export — use buttons above or toolbar" },
      ],
    },
  ],
}

// ── Editor ───────────────────────────────────────────────────────────
const editor = useDocxEditor({
  starterModel,
  initialFileName: "verification-test.docx",
  initialDocumentTheme: { mode: "light" },
})

// ── UI toggles ───────────────────────────────────────────────────────
const showThumbnails = ref(false)
const isReadOnly = ref(false)

const { showTrackedChanges, toggleShowTrackedChanges } = useDocxTrackChanges(editor)
const { showComments, toggleShowComments } = useDocxComments(editor)

// ── File import ──────────────────────────────────────────────────────
function onImport() {
  const input = document.createElement("input")
  input.type = "file"
  input.accept = ".docx"
  input.onchange = async () => {
    const file = input.files?.[0]
    if (file) {
      await editor.importDocxFile(file)
    }
  }
  input.click()
}
</script>

<style scoped>
.page { padding: 24px; max-width: 1200px; margin: 0 auto; width: 100%; min-width: 0; }
h2 { margin-bottom: 4px; }
.desc { color: var(--muted-foreground); margin-bottom: 16px; font-size: 14px; line-height: 1.5; }

.controls { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; margin-bottom: 16px; }
.control-separator { width: 1px; height: 24px; background: var(--border); margin: 0 4px; }
.btn-primary { padding: 8px 14px; border: none; background: var(--primary); color: var(--primary-foreground); border-radius: var(--radius-md); cursor: pointer; font-size: 13px; font-weight: 500; }
.btn-primary:hover { opacity: 0.9; }
.btn-secondary { padding: 8px 14px; border: 1px solid var(--border); background: var(--background); color: var(--foreground); border-radius: var(--radius-md); cursor: pointer; font-size: 13px; }
.btn-secondary:hover { background: var(--accent); }
.toggle-label { display: flex; align-items: center; gap: 6px; font-size: 13px; color: var(--muted-foreground); cursor: pointer; user-select: none; }

.editor-container { border: 1px solid var(--border); border-radius: var(--radius-lg); overflow: hidden; margin-bottom: 24px; background: #f3f4f6; }
.empty { display: flex; align-items: center; justify-content: center; height: 200px; color: var(--muted-foreground); }

.api-verify { margin-top: 24px; }
.api-verify h3 { margin-bottom: 4px; }
.api-verify table { width: 100%; border-collapse: collapse; font-size: 13px; }
.api-verify th, .api-verify td { padding: 6px 12px; text-align: left; border-bottom: 1px solid var(--border); }
.api-note { color: var(--muted-foreground); border: 1px solid var(--border); background: var(--muted); border-radius: var(--radius); padding: 10px 12px; font-size: 13px; line-height: 1.5; margin: 0 0 16px; }
.pass { color: #16a34a; font-weight: 600; }
.warn { color: #b45309; font-weight: 600; }
</style>