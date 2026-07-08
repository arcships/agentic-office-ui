<template>
  <div class="docx-table-host" :style="tableStyle">
    <table class="docx-table-host-table">
      <colgroup>
        <col v-for="(width, ci) of columnWidths" :key="ci" :style="{ width: width }" />
      </colgroup>
      <tbody>
        <tr v-for="(row, ri) of table.rows" :key="ri">
          <td
            v-for="(cell, ci) of row.cells"
            :key="ci"
            :colspan="cell.colSpan"
            class="docx-table-cell"
          >
            <p
              v-for="(para, pi) of cell.paragraphs"
              :key="pi"
              :style="paraStyle(para)"
            >
              <template v-for="run of para.runs" :key="run.id">
                <img
                  v-if="run.kind === 'image' && run.src"
                  :src="run.src"
                  :alt="run.alt ?? 'image'"
                  :style="{ maxWidth: run.widthPx ? run.widthPx + 'px' : '100%' }"
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

<script setup lang="ts">
import { computed } from "vue"
import type {
  LayoutTableBlock,
  LayoutParagraphBlock,
  LayoutRun,
} from "@extend-ai/docx-core"

const props = withDefaults(
  defineProps<{
    table: LayoutTableBlock
    editable?: boolean
    resizable?: boolean
  }>(),
  { editable: false, resizable: false }
)

const columnCount = computed(() => {
  if (props.table.rows.length === 0) return 1
  return props.table.rows[0].cells.length
})

const columnWidths = computed(() =>
  Array.from({ length: columnCount.value }, () => `${100 / columnCount.value}%`)
)

const tableStyle = computed(() => ({
  width: "100%",
  marginBottom: "8px",
  position: "relative" as const,
}))

function paraStyle(para: LayoutParagraphBlock) {
  return {
    margin: "0",
    textAlign: para.align,
    fontWeight: para.headingLevel ? 700 : undefined,
  }
}

function runTextStyle(run: LayoutRun): Record<string, string | undefined> {
  if (run.kind !== "text") return {}
  return {
    fontWeight: run.style?.bold ? "700" : undefined,
    fontStyle: run.style?.italic ? "italic" : undefined,
    color: run.style?.color,
    fontFamily: run.style?.fontFamily,
    fontSize: run.style?.fontSizePt ? `${run.style.fontSizePt}pt` : undefined,
  }
}
</script>

<style scoped>
.docx-table-host-table {
  width: 100%;
  border-collapse: collapse;
  table-layout: fixed;
}
.docx-table-cell {
  border: 1px solid #d1d5db;
  padding: 8px;
  vertical-align: top;
  min-width: 0;
  word-wrap: break-word;
  overflow-wrap: break-word;
  word-break: break-word;
}
</style>
