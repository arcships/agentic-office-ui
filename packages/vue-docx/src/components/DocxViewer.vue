<template>
  <div class="docx-viewer" :class="className">
    <!-- Loading state -->
    <div v-if="isLoading" class="docx-viewer-loading">
      <div class="docx-viewer-loading-spinner" />
      <span>Loading DOCX...</span>
    </div>

    <!-- Error state -->
    <div v-else-if="error" class="docx-viewer-error">
      <div class="docx-viewer-error-icon">⚠</div>
      <div class="docx-viewer-error-text">Failed to parse DOCX</div>
      <div class="docx-viewer-error-detail">{{ error.message }}</div>
    </div>

    <!-- Empty state -->
    <div v-else-if="!resolvedModel" class="docx-viewer-empty">
      {{ emptyState ?? "No DOCX loaded." }}
    </div>

    <!-- Document pages -->
    <template v-else>
      <section
        v-for="(page, pageIndex) of pages"
        :key="page.number"
        :data-page="page.number"
        class="docx-page"
        :style="pageStyle"
      >
        <!-- Page header (if present) -->
        <div
          v-if="pageHeaderNodes(pageIndex).length > 0"
          class="docx-page-section docx-page-section--header"
        >
          <p
            v-for="(node, ni) of pageHeaderNodes(pageIndex)"
            :key="`header-${pageIndex}-${ni}`"
            class="docx-viewer-paragraph"
            :style="({
              color: '#6b7280',
              fontSize: '10pt',
              textAlign: node.style?.align,
            } as any)"
          >
            <template v-for="run of renderNodeRuns(node)" :key="run.id">
              <img
                v-if="run.kind === 'image' && run.src"
                :src="run.src"
                :alt="run.alt ?? 'image'"
                :style="{
                  maxWidth: run.widthPx ? `${run.widthPx}px` : '100%',
                  maxHeight: run.heightPx ? `${run.heightPx}px` : undefined,
                  verticalAlign: 'middle',
                  marginInline: '4px',
                }"
              />
              <a
                v-else-if="run.kind === 'text' && run.link"
                :key="run.id"
                :href="run.link"
                :target="run.link.startsWith('#') ? undefined : '_blank'"
                :rel="run.link.startsWith('#') ? undefined : 'noreferrer noopener'"
                :style="runTextStyle(run, true)"
              >{{ run.text }}</a>
              <span
                v-else-if="run.kind === 'text'"
                :key="run.id"
                :style="runTextStyle(run)"
              >{{ run.text }}</span>
            </template>
          </p>
        </div>

        <!-- Page body: blocks -->
        <template v-for="block of page.blocks" :key="block.id">
          <!-- Paragraph block -->
          <p
            v-if="block.kind === 'paragraph'"
            class="docx-viewer-paragraph"
            :style="paragraphBlockStyle(block)"
          >
            <!-- Numbering label -->
            <span
              v-if="(block as any).numberingLabel"
              class="docx-viewer-numbering-label"
              :style="{ marginRight: '8px', minWidth: '24px' }"
            >{{ (block as any).numberingLabel }}</span>

            <template v-for="run of block.runs" :key="run.id">
              <!-- Image run -->
              <template v-if="run.kind === 'image'">
                <span
                  v-if="!run.src"
                  class="docx-viewer-missing-image"
                  :style="{
                    display: 'inline-flex',
                    minWidth: '120px',
                    minHeight: '80px',
                    alignItems: 'center',
                    justifyContent: 'center',
                    border: '1px dashed #c4c4c4',
                    color: '#6b7280',
                    fontSize: '12px',
                    padding: '8px',
                    marginInline: '4px',
                  }"
                >Missing image</span>

                <img
                  v-else
                  :src="run.src"
                  :alt="run.alt ?? 'DOCX image'"
                  :style="{
                    maxWidth: run.widthPx ? `${run.widthPx}px` : '100%',
                    maxHeight: run.heightPx ? `${run.heightPx}px` : undefined,
                    verticalAlign: 'middle',
                    marginInline: '4px',
                  }"
                />
              </template>

              <!-- Link text run -->
              <a
                v-else-if="run.link"
                :href="run.link"
                :target="run.link.startsWith('#') ? undefined : '_blank'"
                :rel="run.link.startsWith('#') ? undefined : 'noreferrer noopener'"
                :style="runTextStyle(run, true)"
              >{{ run.text }}</a>

              <!-- Plain text run -->
              <span
                v-else
                :style="runTextStyle(run)"
              >{{ run.text }}</span>
            </template>
          </p>

          <!-- Table block -->
          <div v-else-if="block.kind === 'table'" class="docx-viewer-table">
            <table class="docx-viewer-table-element">
              <colgroup>
                <col
                  v-for="(width, ci) of resolveTableColumnWidths(block)"
                  :key="ci"
                  :style="{ width: width }"
                />
              </colgroup>
              <tbody>
                <tr v-for="(row, ri) of block.rows" :key="ri">
                  <td
                    v-for="(cell, ci) of row.cells"
                    :key="ci"
                    :colspan="cell.colSpan"
                    class="docx-viewer-table-cell"
                    :style="{ backgroundColor: (cell as any)?.backgroundColor as string | undefined }"
                  >
                    <p
                      v-for="(para, pi) of cell.paragraphs"
                      :key="pi"
                      class="docx-viewer-paragraph"
                      :style="({
                        margin: '0',
                        fontWeight: para.headingLevel ? '700' : undefined,
                        fontSize: para.headingLevel ? headingFontSize(para.headingLevel) : undefined,
                        textAlign: para.align,
                      } as any)"
                    >
                      <template v-for="run of para.runs" :key="run.id">
                        <img
                          v-if="run.kind === 'image' && run.src"
                          :src="run.src"
                          :alt="run.alt ?? 'image'"
                          :style="{ maxWidth: run.widthPx ? `${run.widthPx}px` : '100%' }"
                        />
                        <span v-else-if="run.kind === 'text'" :style="runTextStyle(run)">
                          {{ run.text }}
                        </span>
                      </template>
                    </p>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </template>

        <!-- Page footer (if present) -->
        <div
          v-if="pageFooterNodes(pageIndex).length > 0"
          class="docx-page-section docx-page-section--footer"
        >
          <p
            v-for="(node, ni) of pageFooterNodes(pageIndex)"
            :key="`footer-${pageIndex}-${ni}`"
            class="docx-viewer-paragraph"
            :style="({
              color: '#6b7280',
              fontSize: '10pt',
              textAlign: node.style?.align,
            } as any)"
          >
            <template v-for="run of renderNodeRuns(node)" :key="run.id">
              <span
                v-if="run.kind === 'text'"
                :style="runTextStyle(run)"
              >{{ run.text }}</span>
            </template>
          </p>
        </div>
      </section>
    </template>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, watch } from "vue"
import type { DocModel } from "@extend-ai/docx-core"
import {
  layoutDocument,
  importDocxBuffer,
  resolveDocumentLayout,
  DEFAULT_DOC_PAGE_WIDTH,
  DEFAULT_DOC_PAGE_HEIGHT,
  pageMarginPaddingStyle,
} from "@extend-ai/docx-core"
import type {
  LayoutBlock,
  LayoutParagraphBlock,
  LayoutRun,
  LayoutTableBlock,
  LayoutOptions,
} from "@extend-ai/docx-core"
import DocxPageBody from "./DocxPageBody"

// ── Constants ──────────────────────────────────────────────────────
const HIGHLIGHT_TO_CSS: Record<string, string> = {
  yellow: "#fff59d",
  green: "#bbf7d0",
  cyan: "#a5f3fc",
  magenta: "#f5d0fe",
  blue: "#bfdbfe",
  red: "#fecaca",
  black: "#111827",
  white: "#ffffff",
  darkgray: "#9ca3af",
  lightgray: "#e5e7eb",
}

const SCRIPT_FONT_SCALE = 0.65

// ── Props ──────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<{
    file?: ArrayBuffer
    model?: DocModel
    className?: string
    layoutOptions?: LayoutOptions
    emptyState?: string
  }>(),
  { emptyState: "No DOCX loaded." }
)

// ── State ──────────────────────────────────────────────────────────
const isLoading = ref(false)
const error = ref<Error | undefined>(undefined)
const parsedModel = ref<DocModel | undefined>(undefined)

const resolvedModel = computed(() => props.model ?? parsedModel.value)

// ── Import DOCX from file ──────────────────────────────────────────
watch(
  () => props.file,
  async (newFile) => {
    if (!newFile) {
      parsedModel.value = undefined
      error.value = undefined
      isLoading.value = false
      return
    }
    if (props.model) return

    const docxFile = newFile
    isLoading.value = true
    error.value = undefined
    try {
      const { model } = await importDocxBuffer(docxFile, { transferBuffer: false })
      parsedModel.value = model
    } catch (err) {
      error.value = err instanceof Error ? err : new Error("Unknown DOCX parse error")
    } finally {
      isLoading.value = false
    }
  },
  { immediate: true }
)

// ── Layout ─────────────────────────────────────────────────────────
const resolvedLayoutOptions = computed<LayoutOptions | undefined>(() => {
  if (!resolvedModel.value) return props.layoutOptions
  return props.layoutOptions ?? {}
})

const pages = computed(() => {
  if (!resolvedModel.value) return []
  return layoutDocument(resolvedModel.value, resolvedLayoutOptions.value)
})

const documentLayout = computed(() => {
  if (!resolvedModel.value) {
    return {
      pageWidthPx: DEFAULT_DOC_PAGE_WIDTH,
      pageHeightPx: DEFAULT_DOC_PAGE_HEIGHT,
      marginsPx: { top: 72, bottom: 72, left: 72, right: 72 },
    }
  }
  return resolveDocumentLayout(resolvedModel.value)
})

const pageWidth = computed(
  () => resolvedLayoutOptions.value?.pageWidth ?? documentLayout.value.pageWidthPx
)
const pageHeight = computed(
  () => resolvedLayoutOptions.value?.pageHeight ?? documentLayout.value.pageHeightPx
)
const margin = computed(() => resolvedLayoutOptions.value?.margin ?? 72)

const pageStyle = computed(() => ({
  width: `${pageWidth.value}px`,
  minHeight: `${pageHeight.value}px`,
  boxSizing: "border-box" as const,
  padding: `${margin.value}px`,
  background: "#fff",
  border: "1px solid #d4d4d4",
  boxShadow: "0 8px 24px rgba(0, 0, 0, 0.08)",
  display: "grid",
  gap: "8px",
  alignContent: "start",
}))

// ── Header / Footer ────────────────────────────────────────────────
function pageHeaderNodes(_pageIndex: number): any[] {
  // Return header nodes for the page — simplified
  return []
}

function pageFooterNodes(_pageIndex: number): any[] {
  // Return footer nodes for the page — simplified
  return []
}

// ── Run rendering ──────────────────────────────────────────────────
function resolveHighlightColor(value?: string): string | undefined {
  if (!value) return undefined
  const normalized = value.trim().toLowerCase()
  if (!normalized) return undefined
  if (normalized.startsWith("#")) return normalized
  return HIGHLIGHT_TO_CSS[normalized] ?? normalized
}

function headingFontSize(level?: 1 | 2 | 3 | 4 | 5 | 6): string | undefined {
  if (!level) return undefined
  const sizes: Record<number, string> = {
    1: "2rem", 2: "1.6rem", 3: "1.35rem",
    4: "1.2rem", 5: "1.05rem", 6: "0.95rem",
  }
  return sizes[level]
}

function runTextStyle(
  run: LayoutRun,
  isLink = false
): Record<string, any> {
  if (run.kind === "image") return {}

  const hasScriptVerticalAlign =
    run.style?.verticalAlign === "superscript" ||
    run.style?.verticalAlign === "subscript"

  const verticalAlign =
    run.style?.verticalAlign === "superscript"
      ? "super"
      : run.style?.verticalAlign === "subscript"
        ? "sub"
        : undefined

  const textDecorationTokens = [
    run.style?.underline ? "underline" : "",
    run.style?.strike ? "line-through" : "",
  ].filter(Boolean)
  const textDecoration =
    textDecorationTokens.length > 0 ? textDecorationTokens.join(" ") : "none"

  return {
    fontWeight: run.style?.bold ? "700" : undefined,
    fontStyle: run.style?.italic ? "italic" : undefined,
    textDecoration,
    color: isLink ? (run.style?.color ?? "inherit") : run.style?.color,
    backgroundColor: resolveHighlightColor(run.style?.highlight),
    fontSize: run.style?.fontSizePt
      ? `${Number(
          ((run.style.fontSizePt as number) * (hasScriptVerticalAlign ? SCRIPT_FONT_SCALE : 1)).toFixed(3)
        )}pt`
      : hasScriptVerticalAlign
        ? `${SCRIPT_FONT_SCALE}em`
        : undefined,
    fontFamily: run.style?.fontFamily,
    verticalAlign,
    whiteSpace: "pre-wrap",
  }
}

function renderNodeRuns(_node: any): LayoutRun[] {
  return []
}

function paragraphBlockStyle(block: LayoutParagraphBlock): Record<string, any> {
  return {
    margin: "0",
    minHeight: (block as any).height ?? undefined,
    textAlign: block.align as string | undefined,
    fontWeight: block.headingLevel ? "700" : undefined,
    fontSize: headingFontSize(block.headingLevel),
  }
}

// ── Table ──────────────────────────────────────────────────────────
function resolveTableColumnWidths(block: LayoutTableBlock): string[] {
  if (block.rows.length === 0) return ["100%"]
  const cellCount = block.rows[0].cells.length
  return Array.from({ length: cellCount }, () => `${100 / cellCount}%`)
}
</script>

<style scoped>
.docx-viewer {
  display: grid;
  justify-items: center;
  gap: 16px;
  padding: 24px 0;
  background: #f3f4f6;
}
.docx-page {
  box-sizing: border-box;
  display: grid;
  gap: 8px;
  align-content: start;
}
.docx-page-section--header {
  padding-bottom: 8px;
  margin-bottom: 4px;
  border-bottom: 1px solid #e5e7eb;
}
.docx-page-section--footer {
  padding-top: 8px;
  margin-top: auto;
  border-top: 1px solid #e5e7eb;
}
.docx-viewer-paragraph {
  margin: 0;
  min-height: 1em;
  white-space: pre-wrap;
  word-wrap: break-word;
  line-height: 1.35;
  font-family: "Calibri", sans-serif;
  font-size: 11pt;
}
.docx-viewer-numbering-label {
  display: inline-block;
  text-align: right;
  user-select: none;
}
.docx-viewer-missing-image {
  display: inline-flex;
  min-width: 120px;
  min-height: 80px;
  align-items: center;
  justify-content: center;
  border: 1px dashed #c4c4c4;
  color: #6b7280;
  font-size: 12px;
  padding: 8px;
  margin-inline: 4px;
}
.docx-viewer-table {
  margin-bottom: 8px;
}
.docx-viewer-table-element {
  width: 100%;
  border-collapse: collapse;
  table-layout: fixed;
}
.docx-viewer-table-cell {
  border: 1px solid #d1d5db;
  padding: 8px;
  vertical-align: top;
  min-width: 0;
  word-wrap: break-word;
  overflow-wrap: break-word;
  word-break: break-word;
}
.docx-viewer-loading {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 24px;
  color: #6b7280;
  font-size: 14px;
}
.docx-viewer-loading-spinner {
  width: 20px;
  height: 20px;
  border: 2px solid #e5e7eb;
  border-top-color: #3b82f6;
  border-radius: 50%;
  animation: docx-viewer-spin 0.6s linear infinite;
}
@keyframes docx-viewer-spin {
  to { transform: rotate(360deg); }
}
.docx-viewer-error {
  padding: 24px;
  text-align: center;
}
.docx-viewer-error-icon {
  font-size: 32px;
  margin-bottom: 8px;
}
.docx-viewer-error-text {
  color: #dc2626;
  font-weight: 600;
  margin-bottom: 4px;
}
.docx-viewer-error-detail {
  color: #6b7280;
  font-size: 13px;
}
.docx-viewer-empty {
  padding: 24px;
  color: #6b7280;
  font-size: 14px;
}
</style>
