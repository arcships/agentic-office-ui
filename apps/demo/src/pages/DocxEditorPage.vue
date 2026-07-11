<template>
  <div class="page" data-testid="docx-editor-page">
    <header class="product-header">
      <div>
        <h2>DOCX 编辑器</h2>
        <p class="desc">直接选择文字并编辑格式；支持撤销、缩放、缩略图、导入和导出。</p>
      </div>
      <div class="controls control-panel product-actions">
        <label class="toggle-label">
          <input data-testid="editor-readonly" type="checkbox" :checked="isReadOnly" @change="isReadOnly = ($event.target as HTMLInputElement).checked" />
          只读
        </label>
      </div>
    </header>

    <div class="editor-container product-surface">
      <DocxEditorViewer
        v-if="editor"
        :editor="editor"
        :editable="!isReadOnly"
        :show-toolbar="true"
        style="height: 76vh;"
      />
      <div v-else class="empty"><p>编辑器正在初始化……</p></div>
    </div>

    <div class="editor-runtime-state verification-section">
      <div data-testid="page-status" :data-state="pageState">状态：{{ pageState }}</div>
      <div data-testid="loaded-file">文件：{{ editor.fileName }}</div>
      <div data-testid="editor-editable-state" :data-state="isReadOnly ? 'readonly' : 'editable'">
        编辑：{{ isReadOnly ? "关闭" : "开启" }}
      </div>
      <div data-testid="editor-history-state" :data-can-undo="String(editor.canUndo)" :data-can-redo="String(editor.canRedo)">
        历史：撤销={{ editor.canUndo }}，重做={{ editor.canRedo }}
      </div>
      <p v-if="editor.importError" class="error" data-testid="load-error" :data-error-code="editorErrorCode">
        {{ editorErrorCode }}: {{ editor.importError.message }}
      </p>
    </div>

    <details class="acceptance-panel" data-testid="editor-acceptance-panel">
      <summary>自动化验收入口</summary>
      <div class="acceptance-actions">
        <button data-testid="editor-test-format-range" class="btn-secondary" @click="formatAcceptanceRange">格式化中间文字</button>
        <button data-testid="editor-test-insert-table" class="btn-secondary" @click="insertAcceptanceTable">插入并扩展表格</button>
        <button data-testid="editor-test-insert-image" class="btn-secondary" @click="insertAcceptanceImage">插入并定位图片</button>
      </div>
      <pre data-testid="editor-selection-snapshot" class="acceptance-snapshot">{{ selectionSnapshot }}</pre>
      <pre data-testid="editor-model-snapshot" class="acceptance-snapshot">{{ modelSnapshot }}</pre>
    </details>
  </div>
</template>

<script setup lang="ts">
import { computed, nextTick, ref } from "vue"
import { DocxEditorViewer, useDocxEditor } from "@arcships/vue-docx"
import { createBlankDocumentModel, type DocModel } from "@arcships/docx-core"

// ── Starter model ────────────────────────────────────────────────────
const starterModel: DocModel = {
  ...createBlankDocumentModel(),
  nodes: [
    {
      type: "paragraph",
      style: { headingLevel: 1 },
      children: [{ type: "text", text: "Quarterly Planning Brief" }],
    },
    {
      type: "paragraph",
      children: [
        { type: "text", text: "The team is preparing the next quarter around " },
        { type: "text", text: "customer retention", style: { bold: true } },
        { type: "text", text: ", " },
        { type: "text", text: "faster delivery", style: { italic: true } },
        { type: "text", text: ", " },
        { type: "text", text: "clear ownership", style: { underline: true } },
        { type: "text", text: ", and a smaller set of measurable goals." },
      ],
    },
    {
      type: "paragraph",
      children: [
        { type: "text", text: "This paragraph is ready for editing. Select a phrase, apply a format, and keep writing without leaving the page." },
      ],
    },
    {
      type: "paragraph",
      style: { headingLevel: 2 },
      children: [{ type: "text", text: "Priorities" }],
    },
    {
      type: "paragraph",
      children: [
        { type: "text", text: "1. Finalize the launch plan and assign owners." },
      ],
    },
    {
      type: "paragraph",
      children: [
        { type: "text", text: "2. Review customer feedback from the last quarter." },
      ],
    },
    {
      type: "paragraph",
      children: [
        { type: "text", text: "3. Confirm the budget, key risks, and delivery milestones." },
      ],
    },
    {
      type: "paragraph",
      children: [
        { type: "text", text: "4. Share the approved brief with the project team." },
      ],
    },
    {
      type: "paragraph",
      children: [
        { type: "text", text: "团队备注：下周五前完成评审并同步风险。" },
      ],
    },
    {
      type: "paragraph",
      style: { headingLevel: 2 },
      children: [{ type: "text", text: "Ownership" }],
    },
    {
      type: "table",
      style: { layout: "fixed", widthTwips: 7200, columnWidthsTwips: [2400, 2400, 2400] },
      rows: [
        {
          type: "table-row",
          cells: [
            {
              type: "table-cell",
              style: { rowSpan: 2, backgroundColor: "#eff6ff", verticalAlign: "center" },
              nodes: [{ type: "paragraph", children: [{ type: "text", text: "Launch readiness", style: { bold: true } }] }],
            },
            {
              type: "table-cell",
              nodes: [{ type: "paragraph", children: [{ type: "text", text: "Product" }] }],
            },
            {
              type: "table-cell",
              nodes: [{ type: "paragraph", children: [{ type: "text", text: "On track" }] }],
            },
          ],
        },
        {
          type: "table-row",
          cells: [
            {
              type: "table-cell",
              style: { vMergeContinuation: true },
              nodes: [{ type: "paragraph", children: [{ type: "text", text: "" }] }],
            },
            {
              type: "table-cell",
              nodes: [{ type: "paragraph", children: [{ type: "text", text: "Operations" }] }],
            },
            {
              type: "table-cell",
              nodes: [{ type: "paragraph", children: [{ type: "text", text: "Reviewing risks" }] }],
            },
          ],
        },
      ],
    },
  ],
}

// ── Editor ───────────────────────────────────────────────────────────
const editor = useDocxEditor({
  starterModel,
  initialFileName: "quarterly-planning-brief.docx",
  initialDocumentTheme: "light",
})

type PageState = "idle" | "loading" | "ready" | "error"

const pageState = computed<PageState>(() => {
  if (editor.isImporting) return "loading"
  if (editor.importError) return "error"
  return "ready"
})

const editorErrorCode = computed(() => {
  const code = (editor.importError as Error & { code?: unknown } | undefined)?.code
  return typeof code === "string" ? code : "PARSE_FAILED"
})

function textOfParagraph(node: DocModel["nodes"][number]): string {
  if (node.type !== "paragraph") return ""
  return node.children
    .map((child) => child.type === "text" ? child.text : "")
    .join("")
}

function summarizeNode(node: DocModel["nodes"][number]) {
  if (node.type === "paragraph") {
    return {
      type: node.type,
      text: textOfParagraph(node),
      runs: node.children.map((child) => {
        if (child.type === "text") {
          return { type: child.type, text: child.text, style: child.style ?? null }
        }
        if (child.type === "image") {
          return {
            type: child.type,
            alt: child.alt ?? null,
            widthPx: child.widthPx ?? null,
            heightPx: child.heightPx ?? null,
            floating: child.floating ?? null,
          }
        }
        return { type: child.type }
      }),
    }
  }

  return {
    type: node.type,
    style: {
      widthTwips: node.style?.widthTwips ?? null,
      layout: node.style?.layout ?? null,
      columnWidthsTwips: node.style?.columnWidthsTwips ?? null,
    },
    rows: node.rows.map((row) => ({
      cells: row.cells.map((cell) => ({
        text: cell.nodes.map((cellNode) => textOfParagraph(cellNode)).join(""),
      })),
    })),
  }
}

const modelSnapshot = computed(() => JSON.stringify({
  nodes: editor.model.nodes.map(summarizeNode),
  tableCount: editor.model.nodes.filter((node) => node.type === "table").length,
  imageCount: editor.model.nodes.reduce((count, node) => {
    if (node.type !== "paragraph") return count
    return count + node.children.filter((child) => child.type === "image").length
  }, 0),
}))

const selectionSnapshot = computed(() => JSON.stringify({
  selection: editor.selection,
  activeTextRange: editor.activeTextRange ?? null,
  historyRestoreRequest: editor.historyRestoreRequest ?? null,
}))

function acceptanceParagraph(): { nodeIndex: number; text: string } | undefined {
  const nodeIndex = editor.model.nodes.findIndex((node) =>
    node.type === "paragraph" && textOfParagraph(node).includes("This paragraph is ready for editing"))
  if (nodeIndex < 0) return undefined
  return { nodeIndex, text: textOfParagraph(editor.model.nodes[nodeIndex]) }
}

async function formatAcceptanceRange(): Promise<void> {
  if (isReadOnly.value) return
  const target = acceptanceParagraph()
  if (!target) return
  const start = target.text.indexOf("editing")
  if (start < 0) return
  editor.selectParagraph(target.nodeIndex)
  await nextTick()
  editor.setActiveTextRange({
    start: { location: { kind: "paragraph", nodeIndex: target.nodeIndex }, offset: start },
    end: { location: { kind: "paragraph", nodeIndex: target.nodeIndex }, offset: start + "editing".length },
  })
  await nextTick()
  editor.toggleBold()
}

async function insertAcceptanceTable(): Promise<void> {
  if (isReadOnly.value) return
  editor.insertTable()
  await nextTick()
  let tableIndex = -1
  editor.model.nodes.forEach((node, nodeIndex) => {
    if (node.type === "table") tableIndex = nodeIndex
  })
  if (tableIndex < 0) return
  editor.insertTableRow(tableIndex, 0, "below")
  await nextTick()
  editor.insertTableColumn(tableIndex, 0, "right")
  await nextTick()
  editor.setTableColumnWidths(tableIndex, [96, 132, 144, 156])
}

async function insertAcceptanceImage(): Promise<void> {
  if (isReadOnly.value) return
  const target = acceptanceParagraph()
  if (!target) return
  editor.selectParagraph(target.nodeIndex)
  await nextTick()
  const imageOffset = 0
  editor.setActiveTextRange({
    start: {
      location: { kind: "paragraph", nodeIndex: target.nodeIndex },
      offset: imageOffset,
    },
    end: {
      location: { kind: "paragraph", nodeIndex: target.nodeIndex },
      offset: imageOffset,
    },
  })
  await nextTick()
  const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="32" height="18"><rect width="32" height="18" fill="#2563eb"/></svg>'
  await editor.insertImageFile(new File([svg], "acceptance.svg", { type: "image/svg+xml" }))
  await nextTick()
  const node = editor.model.nodes[target.nodeIndex]
  if (!node || node.type !== "paragraph") return
  const childIndex = node.children.findIndex((child) => child.type === "image")
  if (childIndex < 0) return
  const location = { kind: "paragraph" as const, nodeIndex: target.nodeIndex, childIndex }
  editor.setImageWrapMode(location, "square", {
    xPx: 12,
    yPx: 6,
    wrapType: "square",
    wrapText: "bothSides",
    distRPx: 12,
    distBPx: 8,
  })
  await nextTick()
  editor.moveFloatingImage(location, { xPx: 12, yPx: 6 })
  await nextTick()
  editor.resizeImage(location, 160, 90)
}

// ── UI toggles ───────────────────────────────────────────────────────
const isReadOnly = ref(false)
</script>

<style scoped>
.page { padding: 16px; max-width: 1440px; margin: 0 auto; width: 100%; min-width: 0; }
h2 { margin-bottom: 4px; }
.desc { color: var(--muted-foreground); margin-bottom: 16px; font-size: 14px; line-height: 1.5; }
.product-header { display: flex; align-items: end; justify-content: space-between; gap: 16px; flex-wrap: wrap; margin-bottom: 12px; }
.product-header .desc { margin-bottom: 0; }
.product-actions { margin-bottom: 0; }
.product-surface { box-shadow: 0 10px 30px rgb(15 23 42 / 8%); }
.verification-section { margin-top: 8px; }

.controls { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; margin-bottom: 16px; }
.control-separator { width: 1px; height: 24px; background: var(--border); margin: 0 4px; }
.btn-primary { padding: 8px 14px; border: none; background: var(--primary); color: var(--primary-foreground); border-radius: var(--radius-md); cursor: pointer; font-size: 13px; font-weight: 500; }
.btn-primary:hover { opacity: 0.9; }
.btn-secondary { padding: 8px 14px; border: 1px solid var(--border); background: var(--background); color: var(--foreground); border-radius: var(--radius-md); cursor: pointer; font-size: 13px; }
.btn-secondary:hover { background: var(--accent); }
.toggle-label { display: flex; align-items: center; gap: 6px; font-size: 13px; color: var(--muted-foreground); cursor: pointer; user-select: none; }
.editor-runtime-state { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; margin: -4px 0 16px; color: var(--muted-foreground); font-size: 13px; }
.error { color: #b91c1c; margin: 0; }
.acceptance-panel { margin: 0 0 16px; padding: 10px 12px; border: 1px solid var(--border); border-radius: var(--radius-md); background: var(--muted); }
.acceptance-actions { display: flex; gap: 8px; flex-wrap: wrap; }
.acceptance-snapshot { margin: 8px 0 0; max-height: 96px; overflow: auto; white-space: pre-wrap; overflow-wrap: anywhere; font-size: 11px; color: var(--muted-foreground); }

.editor-container { border: 1px solid var(--border); border-radius: var(--radius-lg); overflow: hidden; margin-bottom: 24px; background: #f3f4f6; }
.empty { display: flex; align-items: center; justify-content: center; height: 200px; color: var(--muted-foreground); }

.api-verify { margin-top: 24px; }
.api-verify h3 { margin-bottom: 4px; }
.api-verify table { width: 100%; border-collapse: collapse; font-size: 13px; }
.api-verify th, .api-verify td { padding: 6px 12px; text-align: left; border-bottom: 1px solid var(--border); }
.api-note { color: var(--muted-foreground); border: 1px solid var(--border); background: var(--muted); border-radius: var(--radius); padding: 10px 12px; font-size: 13px; line-height: 1.5; margin: 0 0 16px; }
.pass { color: #16a34a; font-weight: 600; }
.warn { color: #b45309; font-weight: 600; }
@media (max-width: 760px) {
  .page { padding: 10px; }
  .product-header { align-items: stretch; }
}
</style>
