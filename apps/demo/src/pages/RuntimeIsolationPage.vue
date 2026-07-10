<template>
  <div class="page" data-testid="runtime-isolation-page">
    <h2>DOCX Runtime 实例隔离</h2>
    <p>两个公开 Runtime 在同页使用不同 WASM 地址、输入限制和任务序列。</p>

    <div data-testid="page-status" :data-state="pageState">{{ pageState }}</div>
    <p v-if="loadError" data-testid="runtime-isolation-error">{{ loadError }}</p>

    <div class="runtime-grid">
      <section data-testid="runtime-a">
        <h3>Runtime A</h3>
        <div data-testid="runtime-a-status" :data-state="stateA">{{ stateA }}</div>
        <code data-testid="runtime-a-wasm">{{ wasmUrlA }}</code>
        <DocxViewer
          v-if="bufferA"
          :file="bufferA"
          :runtime="runtimeA"
          @load-success="stateA = 'ready'"
          @load-error="onError('A', $event)"
        />
        <pre data-testid="runtime-a-diagnostics">{{ JSON.stringify(diagnosticsA, null, 2) }}</pre>
      </section>

      <section data-testid="runtime-b">
        <h3>Runtime B</h3>
        <div data-testid="runtime-b-status" :data-state="stateB">{{ stateB }}</div>
        <code data-testid="runtime-b-wasm">{{ wasmUrlB }}</code>
        <DocxViewer
          v-if="bufferB"
          :file="bufferB"
          :runtime="runtimeB"
          @load-success="stateB = 'ready'"
          @load-error="onError('B', $event)"
        />
        <pre data-testid="runtime-b-diagnostics">{{ JSON.stringify(diagnosticsB, null, 2) }}</pre>
      </section>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, onBeforeUnmount, onMounted, ref, shallowRef } from "vue"
import {
  bundledDocxWasmUrl,
  createDocxRuntime,
  type DocxImportDiagnostic,
} from "@extend-ai/docx-core"
import { DocxViewer } from "@extend-ai/vue-docx"

type State = "loading" | "ready" | "error"

const wasmUrlA = `${bundledDocxWasmUrl}?runtime=alpha`
const wasmUrlB = `${bundledDocxWasmUrl}?runtime=beta`
const diagnosticsA = ref<DocxImportDiagnostic[]>([])
const diagnosticsB = ref<DocxImportDiagnostic[]>([])
const stateA = ref<State>("loading")
const stateB = ref<State>("loading")
const loadError = ref("")
const bufferA = shallowRef<ArrayBuffer>()
const bufferB = shallowRef<ArrayBuffer>()

const configA = {
  wasmUrl: wasmUrlA,
  limits: { maxInputBytes: 2 * 1024 * 1024 },
  onDiagnostic: (event: DocxImportDiagnostic) => diagnosticsA.value.push(event),
}
const configB = {
  wasmUrl: wasmUrlB,
  limits: { maxInputBytes: 3 * 1024 * 1024 },
  onDiagnostic: (event: DocxImportDiagnostic) => diagnosticsB.value.push(event),
}
const runtimeA = createDocxRuntime(configA)
const runtimeB = createDocxRuntime(configB)

// Mutating the caller objects after construction must not affect either runtime.
configA.limits.maxInputBytes = 1
configB.limits.maxInputBytes = 1

const pageState = computed<State>(() => {
  if (stateA.value === "error" || stateB.value === "error") return "error"
  if (stateA.value === "ready" && stateB.value === "ready") return "ready"
  return "loading"
})

function onError(instance: "A" | "B", error: Error) {
  if (instance === "A") stateA.value = "error"
  else stateB.value = "error"
  loadError.value = `${instance}: ${error.message}`
}

onMounted(async () => {
  try {
    const response = await fetch("/samples/invoice-table.docx")
    if (!response.ok) throw new Error(`sample returned ${response.status}`)
    const bytes = await response.arrayBuffer()
    bufferA.value = bytes.slice(0)
    bufferB.value = bytes.slice(0)
  } catch (error) {
    stateA.value = "error"
    stateB.value = "error"
    loadError.value = error instanceof Error ? error.message : String(error)
  }
})

onBeforeUnmount(() => {
  runtimeA.dispose()
  runtimeB.dispose()
})
</script>

<style scoped>
.page { padding: 24px; max-width: 1440px; margin: 0 auto; }
.runtime-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; }
section { border: 1px solid var(--border); border-radius: var(--radius); padding: 12px; min-width: 0; }
code { display: block; overflow-wrap: anywhere; margin-bottom: 8px; }
pre { max-height: 220px; overflow: auto; white-space: pre-wrap; font-size: 11px; }
@media (max-width: 900px) { .runtime-grid { grid-template-columns: 1fr; } }
</style>
