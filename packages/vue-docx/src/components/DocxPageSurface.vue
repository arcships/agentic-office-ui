<template>
  <div
    ref="pageSurfaceRef"
    data-docx-page-surface="true"
    class="docx-page-surface"
    :style="surfaceStyle"
    @click="onSurfaceClick"
  >
    <!-- Cover layer (full-page background images) -->
    <div
      v-if="coverEntries.length > 0"
      data-docx-page-cover-layer="true"
      class="docx-page-cover-layer"
    >
      <img
        v-for="cover of coverEntries"
        :key="cover.key"
        :src="cover.src"
        alt=""
        aria-hidden="true"
        draggable="false"
        class="docx-page-cover-image"
        :style="cover.style"
      />
    </div>

    <!-- Border overlay -->
    <div
      v-if="borderStyle"
      data-docx-page-border-overlay="true"
      class="docx-page-border-overlay"
      :style="borderStyle"
    />

    <!-- Body -->
    <DocxPageBody
      :page-index="pageIndex"
      :page-node-segments="pageNodeSegments"
      :page-layout="pageLayout"
      :page-content-width-px="pageContentWidthPx"
      :model="model"
      :controller="controller"
      :editable="editable"
      :theme="theme"
      :page-number="pageNumber"
      :total-pages="totalPages"
      :page-number-format="pageNumberFormat"
      :tracked-changes-enabled="trackedChangesEnabled"
      :comments-enabled="commentsEnabled"
    >
      <template v-if="hasHeader" #header>
        <DocxPageHeader
          :page-index="pageIndex"
          :section="headerSection"
          :page-layout="pageLayout"
          :page-content-width-px="pageContentWidthPx"
          :theme="theme"
          :model="model"
          :controller="controller"
          :page-number="pageNumber"
          :total-pages="totalPages"
          :page-number-format="pageNumberFormat"
          :tracked-changes-enabled="trackedChangesEnabled"
          :comments-enabled="commentsEnabled"
        />
      </template>
      <template v-if="hasFooter" #footer>
        <DocxPageFooter
          :page-index="pageIndex"
          :section="footerSection"
          :page-layout="pageLayout"
          :page-content-width-px="pageContentWidthPx"
          :theme="theme"
          :model="model"
          :controller="controller"
          :page-number="pageNumber"
          :total-pages="totalPages"
          :page-number-format="pageNumberFormat"
          :tracked-changes-enabled="trackedChangesEnabled"
          :comments-enabled="commentsEnabled"
        />
      </template>
    </DocxPageBody>

    <!-- Floating image layer -->
    <DocxImageLayer
      v-if="editable && controller"
      :page-index="pageIndex"
      :page-width-px="pageWidthPx"
      :page-height-px="pageLayout.pageHeightPx"
      :controller="controller"
      :page-node-segments="pageNodeSegments"
      :editable="editable"
    />

    <!-- Form field layer -->
    <DocxFormFieldLayer
      v-if="editable && controller && hasFormFields"
      :page-index="pageIndex"
      :controller="controller"
      :page-node-segments="pageNodeSegments"
    />

    <!-- Tracked change / comment gutter -->
    <DocxTrackedChangeGutter
      v-if="controller && showGutter"
      :page-index="pageIndex"
      :page-layout="pageLayout"
      :controller="controller"
      :page-node-segments="pageNodeSegments"
    />
  </div>
</template>

<script setup lang="ts">
import { ref, computed, onMounted, onUnmounted, watch } from "vue"
import type {
  DocModel,
  DocxEditorController,
  DocumentPageNodeSegment,
  FooterSection,
  HeaderSection,
  TableCellContentNode,
} from "@extend-ai/docx-core"
import {
  docModelThumbnailMetadataSignature,
  docNodeContentSignature,
  resolveDocumentLayout,
  pageMarginPaddingStyle,
} from "@extend-ai/docx-core"
import DocxPageHeader from "./DocxPageHeader.vue"
import DocxPageFooter from "./DocxPageFooter.vue"
import DocxPageBody from "./DocxPageBody"
import DocxImageLayer from "./DocxImageLayer.vue"
import DocxFormFieldLayer from "./DocxFormFieldLayer.vue"
import DocxTrackedChangeGutter from "./DocxTrackedChangeGutter.vue"
import {
  ensureDocxViewerPageSurfaceRegistry,
  notifyDocxViewerPageSurfaceListeners,
} from "../composables/page-surface-registry"

// ── Types ──────────────────────────────────────────────────────────
interface DocxPageSurfaceProps {
  pageIndex: number
  pageLayout: ReturnType<typeof resolveDocumentLayout>
  pageNodeSegments: DocumentPageNodeSegment[]
  model: DocModel
  controller?: DocxEditorController
  editable?: boolean
  pageWidthPx: number
  pageContentWidthPx: number
  pageNumber: number
  totalPages: number
  pageNumberFormat?: string
  headerSection?: HeaderSection
  footerSection?: FooterSection
  trackedChangesEnabled?: boolean
  commentsEnabled?: boolean
  theme?: "light" | "dark"
}

interface CoverEntry {
  key: string
  src: string
  style: Record<string, string>
}

// ── Props ──────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<DocxPageSurfaceProps>(),
  {
    editable: true,
    trackedChangesEnabled: false,
    commentsEnabled: false,
    theme: "light" as const,
  }
)

const emit = defineEmits<{
  measure: [event: { heightPx: number }]
}>()

// ── Refs ───────────────────────────────────────────────────────────
const pageSurfaceRef = ref<HTMLDivElement | undefined>()

let publishedSurface:
  | {
      controller: DocxEditorController
      element: HTMLDivElement
      pageIndex: number
    }
  | undefined

function unpublishPageSurface(): void {
  if (!publishedSurface) return
  const { controller, element, pageIndex } = publishedSurface
  const registry = ensureDocxViewerPageSurfaceRegistry(controller)
  if (registry.pageElements.get(pageIndex) === element) {
    registry.pageElements.delete(pageIndex)
    registry.pageContentKeys.delete(pageIndex)
    registry.pageSizes.delete(pageIndex)
    registry.pageThumbnailSnapshots.delete(pageIndex)
    notifyDocxViewerPageSurfaceListeners(controller)
  }
  publishedSurface = undefined
}

function publishPageSurface(): void {
  unpublishPageSurface()
  const controller = props.controller
  const element = pageSurfaceRef.value
  if (!controller || !element) return

  const registry = ensureDocxViewerPageSurfaceRegistry(controller)
  registry.pageElements.set(props.pageIndex, element)
  registry.pageSizes.set(props.pageIndex, {
    widthPx: props.pageWidthPx,
    heightPx: props.pageLayout.pageHeightPx,
  })
  registry.pageContentKeys.set(
    props.pageIndex,
    [
      controller.documentLoadNonce,
      docModelThumbnailMetadataSignature(props.model.metadata),
      ...props.pageNodeSegments.map((segment) =>
        docNodeContentSignature(props.model.nodes[segment.nodeIndex])
      ),
    ].join(":")
  )
  publishedSurface = { controller, element, pageIndex: props.pageIndex }
  notifyDocxViewerPageSurfaceListeners(controller)
}

// ── Computed ───────────────────────────────────────────────────────
const hasHeader = computed(() => (props.headerSection?.nodes.length ?? 0) > 0)
const hasFooter = computed(() => (props.footerSection?.nodes.length ?? 0) > 0)

function nodesContainFormField(nodes: TableCellContentNode[]): boolean {
  return nodes.some((node) => {
    if (node.type === "paragraph") {
      return node.children.some((child) => child.type === "form-field")
    }
    return node.rows.some((row) =>
      row.cells.some((cell) => nodesContainFormField(cell.nodes))
    )
  })
}

const hasFormFields = computed(() =>
  props.pageNodeSegments.some((segment) => {
    const node = props.model.nodes[segment.nodeIndex]
    if (!node) return false
    if (node.type === "paragraph") {
      return node.children.some((child) => child.type === "form-field")
    }
    return node.rows.some((row) =>
      row.cells.some((cell) => nodesContainFormField(cell.nodes))
    )
  })
)

const showGutter = computed(
  () => props.trackedChangesEnabled || props.commentsEnabled
)

const coverEntries = computed<CoverEntry[]>(() => {
  // Collect full-page cover background images from paragraph segments
  // Simplified: empty for now — full implementation requires image run detection
  return []
})

const borderStyle = computed<Record<string, string> | undefined>(() => {
  // Page border overlay — simplified: no border by default
  return undefined
})

const surfaceStyle = computed(() => ({
  width: `${props.pageWidthPx}px`,
  minHeight: `${props.pageLayout.pageHeightPx}px`,
  backgroundColor: props.theme === "dark" ? "#1f2937" : "#ffffff",
  color: props.theme === "dark" ? "#f9fafb" : "#111827",
  position: "relative" as const,
  boxSizing: "border-box" as const,
  border: "1px solid #d4d4d4",
  boxShadow: "0 8px 24px rgba(0, 0, 0, 0.08)",
  ...pageMarginPaddingStyle(props.pageLayout.marginsPx ?? { top: 72, bottom: 72, left: 72, right: 72 }),
}))

// ── Surface click → select paragraph ───────────────────────────────
function onSurfaceClick(event: MouseEvent): void {
  // Only handle clicks directly on the surface (not on child interactive elements)
  if (event.target === pageSurfaceRef.value || (event.target as HTMLElement)?.dataset?.docxPageSurface === "true") {
    // Don't steal focus from form fields or editable elements
  }
}

// ── Measure page height ────────────────────────────────────────────
let resizeObserver: ResizeObserver | null = null

function measurePage(): void {
  if (!pageSurfaceRef.value) return
  const rect = pageSurfaceRef.value.getBoundingClientRect()
  emit("measure", { heightPx: rect.height })
}

onMounted(() => {
  publishPageSurface()
  const ResizeObserverConstructor = globalThis.ResizeObserver
  if (pageSurfaceRef.value && typeof ResizeObserverConstructor === "function") {
    resizeObserver = new ResizeObserverConstructor(() => {
      measurePage()
    })
    resizeObserver.observe(pageSurfaceRef.value)
    measurePage()
  }
})

onUnmounted(() => {
  unpublishPageSurface()
  if (resizeObserver) {
    resizeObserver.disconnect()
    resizeObserver = null
  }
})

watch(
  () => [
    props.controller,
    props.controller?.documentLoadNonce,
    props.model,
    props.pageIndex,
    props.pageWidthPx,
    props.pageLayout.pageHeightPx,
  ] as const,
  () => publishPageSurface(),
  { flush: "post" }
)
</script>

<style scoped>
.docx-page-surface {
  box-sizing: border-box;
  display: flex;
  flex-direction: column;
}
.docx-page-cover-layer {
  position: absolute;
  left: 0;
  top: 0;
  right: 0;
  bottom: 0;
  overflow: clip;
  pointer-events: none;
  z-index: 0;
}
.docx-page-cover-image {
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  display: block;
  object-fit: cover;
}
.docx-page-border-overlay {
  position: absolute;
  inset: 0;
  pointer-events: none;
  z-index: 5;
}
</style>
