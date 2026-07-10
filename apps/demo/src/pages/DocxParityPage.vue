<template>
  <main class="page" data-testid="docx-parity-page">
    <h2>DOCX Viewer / Editor 只读渲染一致性</h2>
    <p class="description">
      同一个公开 DocModel 同时交给 Viewer 和只读 Editor，并按实际渲染结果比较结构。
    </p>

    <div class="controls">
      <label>
        样例
        <select v-model="selectedSample" data-testid="docx-parity-sample">
          <option v-for="sample in samples" :key="sample.file" :value="sample.file">
            {{ sample.label }}
          </option>
        </select>
      </label>
      <button data-testid="docx-parity-load" type="button" @click="loadSelectedSample">
        加载并比较
      </button>
    </div>

    <div class="status-grid">
      <div data-testid="page-status" :data-state="pageState">状态：{{ pageState }}</div>
      <div data-testid="loaded-file">文件：{{ loadedFile || "未加载" }}</div>
      <div data-testid="parity-status" :data-state="parityState">
        一致性：{{ parityState }}
      </div>
    </div>
    <p v-if="loadError" class="error" data-testid="load-error">{{ loadError }}</p>

    <div v-if="model" class="parity-grid">
      <section>
        <h3>DocxViewer</h3>
        <div ref="viewerHost" class="parity-surface" data-testid="parity-viewer">
          <DocxViewer :model="model" />
        </div>
      </section>
      <section>
        <h3>DocxEditor（只读、无工具栏）</h3>
        <div ref="editorHost" class="parity-surface" data-testid="parity-editor">
          <DocxEditor
            :model="model"
            :editable="false"
            :show-toolbar="false"
            :show-thumbnails="false"
          />
        </div>
      </section>
    </div>

    <details open>
      <summary>结构差异</summary>
      <pre data-testid="parity-diff">{{ JSON.stringify(parityDiff, null, 2) }}</pre>
    </details>
    <details>
      <summary>Viewer 结构</summary>
      <pre data-testid="viewer-structure">{{ JSON.stringify(viewerStructure, null, 2) }}</pre>
    </details>
    <details>
      <summary>Editor 结构</summary>
      <pre data-testid="editor-structure">{{ JSON.stringify(editorStructure, null, 2) }}</pre>
    </details>
  </main>
</template>

<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref, shallowRef } from "vue"
import { useRoute } from "vue-router"
import { createDocxRuntime, type DocModel } from "@extend-ai/docx-core"
import { DocxEditor, DocxViewer } from "@extend-ai/vue-docx"

type PageState = "idle" | "loading" | "ready" | "error"
type ImageShape = { alt: string; sourceKind: string }
type CellShape = { text: string; images: ImageShape[] }
type NodeShape =
  | { type: "paragraph"; text: string; images: ImageShape[] }
  | { type: "table"; text: string; rows: CellShape[][]; images: ImageShape[] }
type SectionShape = { text: string; images: ImageShape[] }
type PageShape = {
  index: number
  nodes: NodeShape[]
  text: string
  tables: number
  cells: number
  images: number
  header: SectionShape
  footer: SectionShape
}
type InteractionShape = {
  contentEditableTrue: number
  inputs: number
  selects: number
  resizeHandles: number
}
type RenderStructure = {
  pageCount: number
  pages: PageShape[]
  totals: { nodes: number; paragraphs: number; tables: number; cells: number; images: number }
  interaction: InteractionShape
}
type ParityDiff = { equal: boolean; differences: string[] }

const samples = [
  { file: "legal-contract.docx", label: "多页合同" },
  { file: "invoice-table.docx", label: "发票表格" },
  { file: "report-with-image.docx", label: "图文报告" },
  { file: "chinese-mixed.docx", label: "中英文混排" },
] as const
type SampleFile = (typeof samples)[number]["file"]
const allowedSamples = new Set(samples.map((sample) => sample.file))
const emptyInteraction = (): InteractionShape => ({
  contentEditableTrue: 0,
  inputs: 0,
  selects: 0,
  resizeHandles: 0,
})
const emptyStructure = (): RenderStructure => ({
  pageCount: 0,
  pages: [],
  totals: { nodes: 0, paragraphs: 0, tables: 0, cells: 0, images: 0 },
  interaction: emptyInteraction(),
})

const route = useRoute()
const runtime = createDocxRuntime()
const loader = runtime.createLoader()
const selectedSample = ref<SampleFile>(samples[0].file)
const pageState = ref<PageState>("idle")
const loadedFile = ref("")
const loadError = ref("")
const model = shallowRef<DocModel>()
const viewerHost = ref<HTMLElement>()
const editorHost = ref<HTMLElement>()
const viewerStructure = ref<RenderStructure>(emptyStructure())
const editorStructure = ref<RenderStructure>(emptyStructure())
const parityDiff = ref<ParityDiff>({ equal: false, differences: ["尚未比较"] })
let loadSequence = 0
let activeFetch: AbortController | undefined

const parityState = computed(() => {
  if (pageState.value !== "ready") return "PENDING"
  return parityDiff.value.equal ? "PASS" : "FAIL"
})

function normalizeText(value: string | null | undefined): string {
  return (value ?? "").replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim()
}

function imageSourceKind(source: string | null): string {
  if (!source) return "missing"
  if (source.startsWith("blob:")) return "blob"
  if (source.startsWith("data:")) return "data"
  if (source.startsWith("http://") || source.startsWith("https://")) return "http"
  return "relative"
}

function normalizeImages(root: ParentNode | null): ImageShape[] {
  if (!root) return []
  return Array.from(root.querySelectorAll("img")).map((image) => ({
    alt: normalizeText(image.getAttribute("alt")),
    sourceKind: imageSourceKind(image.getAttribute("src")),
  }))
}

function normalizeTable(tableRoot: Element): NodeShape {
  const rows = Array.from(tableRoot.querySelectorAll("tr")).map((row) =>
    Array.from(row.querySelectorAll(":scope > td, :scope > th")).map((cell) => ({
      text: normalizeText(cell.textContent),
      images: normalizeImages(cell),
    })),
  )
  return {
    type: "table",
    text: normalizeText(rows.flat().map((cell) => cell.text).join(" ")),
    rows,
    images: normalizeImages(tableRoot),
  }
}

function normalizeParagraph(paragraph: Element): NodeShape {
  return {
    type: "paragraph",
    text: normalizeText(paragraph.textContent),
    images: normalizeImages(paragraph),
  }
}

function normalizeSection(section: Element | null): SectionShape {
  return {
    text: normalizeText(section?.textContent),
    images: normalizeImages(section),
  }
}

function normalizeViewerPage(page: Element, position: number): PageShape {
  const bodyNodes = Array.from(page.children)
    .filter((child) => !child.classList.contains("docx-page-section--header"))
    .filter((child) => !child.classList.contains("docx-page-section--footer"))
    .filter((child) =>
      child.classList.contains("docx-viewer-paragraph") ||
      child.classList.contains("docx-viewer-table"),
    )
    .map((child) =>
      child.classList.contains("docx-viewer-table")
        ? normalizeTable(child)
        : normalizeParagraph(child),
    )
  return pageShape(
    position,
    bodyNodes,
    normalizeSection(page.querySelector(".docx-page-section--header")),
    normalizeSection(page.querySelector(".docx-page-section--footer")),
    page,
  )
}

function normalizeEditorPage(wrapper: Element): PageShape {
  const surface = wrapper.querySelector('[data-docx-page-surface="true"]')
  const body = surface?.querySelector('[data-docx-page-body="true"]')
  const bodyNodes = body
    ? Array.from(body.children)
        .filter((child) =>
          child.matches('[data-docx-paragraph-host="true"], [data-docx-table-host="true"]'),
        )
        .map((child) =>
          child.matches('[data-docx-table-host="true"]')
            ? normalizeTable(child)
            : normalizeParagraph(child),
        )
    : []
  const rawIndex = Number(wrapper.getAttribute("data-docx-page-index"))
  return pageShape(
    Number.isFinite(rawIndex) ? rawIndex : 0,
    bodyNodes,
    normalizeSection(surface?.querySelector('[data-docx-page-header="true"]') ?? null),
    normalizeSection(surface?.querySelector('[data-docx-page-footer="true"]') ?? null),
    surface,
  )
}

function pageShape(
  index: number,
  nodes: NodeShape[],
  header: SectionShape,
  footer: SectionShape,
  pageRoot: ParentNode | null,
): PageShape {
  return {
    index,
    nodes,
    text: normalizeText(nodes.map((node) => node.text).join(" ")),
    tables: nodes.filter((node) => node.type === "table").length,
    cells: nodes.reduce(
      (total, node) => total + (node.type === "table" ? node.rows.flat().length : 0),
      0,
    ),
    images: normalizeImages(pageRoot).length,
    header,
    footer,
  }
}

function readInteraction(root: ParentNode | null): InteractionShape {
  if (!root) return emptyInteraction()
  return {
    contentEditableTrue: root.querySelectorAll('[contenteditable="true"]').length,
    inputs: root.querySelectorAll("input").length,
    selects: root.querySelectorAll("select").length,
    resizeHandles: root.querySelectorAll(
      '[class*="resize-handle"], [class*="column-handle"], [class*="move-handle"], [class*="add-row"]',
    ).length,
  }
}

function mergeInteraction(left: InteractionShape, right: InteractionShape): InteractionShape {
  return {
    contentEditableTrue: Math.max(left.contentEditableTrue, right.contentEditableTrue),
    inputs: Math.max(left.inputs, right.inputs),
    selects: Math.max(left.selects, right.selects),
    resizeHandles: Math.max(left.resizeHandles, right.resizeHandles),
  }
}

function buildStructure(pages: PageShape[], interaction: InteractionShape): RenderStructure {
  const orderedPages = [...pages].sort((left, right) => left.index - right.index)
  const nodes = orderedPages.flatMap((page) => page.nodes)
  return {
    pageCount: orderedPages.length,
    pages: orderedPages,
    totals: {
      nodes: nodes.length,
      paragraphs: nodes.filter((node) => node.type === "paragraph").length,
      tables: nodes.filter((node) => node.type === "table").length,
      cells: nodes.reduce(
        (total, node) => total + (node.type === "table" ? node.rows.flat().length : 0),
        0,
      ),
      images: orderedPages.reduce((total, page) => total + page.images, 0),
    },
    interaction,
  }
}

function collectLegacyViewerStructure(): RenderStructure | undefined {
  const root = viewerHost.value
  if (!root) return undefined
  const pageElements = Array.from(root.querySelectorAll('[data-testid="docx-page"]'))
  if (pageElements.length === 0) return undefined
  const pages = pageElements
    .map((page, index) => normalizeViewerPage(page, index))
  return buildStructure(pages, readInteraction(root))
}

function waitForPaint(delay = 0): Promise<void> {
  return new Promise((resolve) => {
    requestAnimationFrame(() => requestAnimationFrame(() => {
      if (delay > 0) window.setTimeout(resolve, delay)
      else resolve()
    }))
  })
}

async function collectVirtualStructure(
  host: HTMLElement | undefined,
  expectedPages: number,
  sequence: number,
): Promise<RenderStructure> {
  const scrollRoot = host?.querySelector<HTMLElement>(".docx-viewer-root")
  if (!host || !scrollRoot) return emptyStructure()

  const pageMap = new Map<number, PageShape>()
  let interaction = emptyInteraction()
  const capture = () => {
    for (const wrapper of host.querySelectorAll('[data-docx-page-wrapper="true"]')) {
      const page = normalizeEditorPage(wrapper)
      pageMap.set(page.index, page)
    }
    interaction = mergeInteraction(interaction, readInteraction(host))
  }

  const previousTop = scrollRoot.scrollTop
  capture()
  const maxScroll = Math.max(0, scrollRoot.scrollHeight - scrollRoot.clientHeight)
  const viewportStep = Math.max(540, Math.floor(scrollRoot.clientHeight * 0.75))
  const scanSteps = Math.min(
    64,
    Math.max(expectedPages * 2, Math.ceil(maxScroll / viewportStep) + 1),
  )
  for (let step = 0; step <= scanSteps; step += 1) {
    if (sequence !== loadSequence) return emptyStructure()
    scrollRoot.scrollTop = scanSteps === 0 ? 0 : Math.round(maxScroll * step / scanSteps)
    scrollRoot.dispatchEvent(new Event("scroll"))
    await waitForPaint(20)
    capture()
  }
  scrollRoot.scrollTop = previousTop
  scrollRoot.dispatchEvent(new Event("scroll"))
  await waitForPaint()
  return buildStructure([...pageMap.values()], interaction)
}

function compareStructures(left: unknown, right: unknown): ParityDiff {
  const differences: string[] = []
  const visit = (leftValue: unknown, rightValue: unknown, path: string): void => {
    if (differences.length >= 200 || Object.is(leftValue, rightValue)) return
    if (Array.isArray(leftValue) || Array.isArray(rightValue)) {
      if (!Array.isArray(leftValue) || !Array.isArray(rightValue)) {
        differences.push(`${path}: 类型不同`)
        return
      }
      if (leftValue.length !== rightValue.length) {
        differences.push(`${path}.length: ${leftValue.length} != ${rightValue.length}`)
      }
      const length = Math.max(leftValue.length, rightValue.length)
      for (let index = 0; index < length; index += 1) {
        visit(leftValue[index], rightValue[index], `${path}[${index}]`)
      }
      return
    }
    if (
      typeof leftValue === "object" && leftValue !== null &&
      typeof rightValue === "object" && rightValue !== null
    ) {
      const keys = new Set([
        ...Object.keys(leftValue as Record<string, unknown>),
        ...Object.keys(rightValue as Record<string, unknown>),
      ])
      for (const key of keys) {
        visit(
          (leftValue as Record<string, unknown>)[key],
          (rightValue as Record<string, unknown>)[key],
          `${path}.${key}`,
        )
      }
      return
    }
    differences.push(`${path}: ${JSON.stringify(leftValue)} != ${JSON.stringify(rightValue)}`)
  }
  visit(left, right, "structure")
  return { equal: differences.length === 0, differences }
}

async function collectAndCompare(sequence: number): Promise<void> {
  await nextTick()
  await waitForPaint(80)
  if (sequence !== loadSequence) return
  // Both public components now use the same virtualized document surface.
  // Scan the complete virtual scroll range instead of treating the currently
  // mounted page subset as the full document.
  const viewer = await collectVirtualStructure(viewerHost.value, 0, sequence)
  const editor = await collectVirtualStructure(editorHost.value, viewer.pageCount, sequence)
  if (sequence !== loadSequence) return
  viewerStructure.value = viewer
  editorStructure.value = editor
  parityDiff.value = compareStructures(viewer, editor)
  pageState.value = "ready"
}

async function loadSelectedSample(): Promise<void> {
  const sequence = ++loadSequence
  activeFetch?.abort()
  loader.cancel()
  const controller = new AbortController()
  activeFetch = controller
  pageState.value = "loading"
  loadError.value = ""
  loadedFile.value = ""
  model.value = undefined
  viewerStructure.value = emptyStructure()
  editorStructure.value = emptyStructure()
  parityDiff.value = { equal: false, differences: ["渲染尚未完成"] }
  try {
    const fileName = selectedSample.value
    const response = await fetch(`/samples/${encodeURIComponent(fileName)}`, {
      signal: controller.signal,
    })
    if (!response.ok) throw new Error(`样例返回 ${response.status}`)
    const bytes = await response.arrayBuffer()
    if (sequence !== loadSequence) return
    const result = await loader.load({ kind: "bytes", bytes, name: fileName }, {
      transferBuffer: false,
    })
    if (sequence !== loadSequence) return
    model.value = result.model
    loadedFile.value = fileName
    await collectAndCompare(sequence)
  } catch (error) {
    if (sequence !== loadSequence || controller.signal.aborted) return
    pageState.value = "error"
    loadError.value = error instanceof Error ? error.message : String(error)
  } finally {
    if (activeFetch === controller) activeFetch = undefined
  }
}

onMounted(() => {
  const requested = typeof route.query.sample === "string" ? route.query.sample : ""
  if (allowedSamples.has(requested as SampleFile)) {
    selectedSample.value = requested as SampleFile
  }
  void loadSelectedSample()
})

onBeforeUnmount(() => {
  loadSequence += 1
  activeFetch?.abort()
  loader.dispose()
  runtime.dispose()
})
</script>

<style scoped>
.page { max-width: 1720px; margin: 0 auto; padding: 24px; }
.description { color: var(--muted-foreground); }
.controls, .status-grid { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; margin: 14px 0; }
.controls label { display: flex; gap: 8px; align-items: center; }
.controls select, .controls button { padding: 7px 10px; }
.status-grid > div { border: 1px solid var(--border); border-radius: var(--radius); padding: 8px 10px; }
[data-testid="parity-status"][data-state="PASS"] { color: #166534; }
[data-testid="parity-status"][data-state="FAIL"], .error { color: #b91c1c; }
.parity-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; align-items: start; }
.parity-grid section { min-width: 0; }
.parity-surface { height: 720px; overflow: hidden; border: 1px solid var(--border); border-radius: var(--radius); background: #f3f4f6; }
.parity-surface :deep(.docx-viewer), .parity-surface :deep(.docx-editor-viewer), .parity-surface :deep(.docx-viewer-root) { height: 100%; }
details { margin-top: 14px; }
pre { max-height: 420px; overflow: auto; padding: 12px; background: var(--muted); white-space: pre-wrap; overflow-wrap: anywhere; }
@media (max-width: 1100px) { .parity-grid { grid-template-columns: 1fr; } }
</style>
