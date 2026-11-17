<template>
  <div class="page" data-testid="runtime-isolation-page">
    <h2>DOCX Runtime 实例隔离</h2>
    <p>两个公开 Runtime 在同页使用不同 WASM 地址、输入限制和任务序列。</p>

    <div data-testid="page-status" :data-state="pageState">{{ pageState }}</div>
    <p v-if="loadError" data-testid="runtime-isolation-error">{{ loadError }}</p>

    <div class="runtime-grid">
      <section
        data-testid="runtime-a"
        :data-runtime-id="runtimeAId"
        :data-input-limit="ISOLATION_SUCCESS_LIMIT_BYTES"
      >
        <h3>Runtime A</h3>
        <div data-testid="runtime-a-status" :data-state="stateA">{{ stateA }}</div>
        <code data-testid="runtime-a-wasm">{{ wasmUrlA }}</code>
        <code data-testid="runtime-a-input-limit">maxInputBytes={{ ISOLATION_SUCCESS_LIMIT_BYTES }}</code>
        <DocxViewer
          v-if="bufferA"
          :file="bufferA"
          :runtime="runtimeA"
          @load-success="onSuccess('A')"
          @load-error="onError('A', $event)"
        />
        <div
          v-if="errorA"
          data-testid="runtime-a-error"
          :data-error-code="errorA.code"
          :data-error-phase="errorA.phase"
          :data-error-limit="errorA.limit"
          :data-error-actual="errorA.actual"
          :data-error-allowed="errorA.allowed"
        >{{ errorA.code }}: {{ errorA.message }}</div>
        <pre data-testid="runtime-a-diagnostics">{{ JSON.stringify(diagnosticsA, null, 2) }}</pre>
      </section>

      <section
        data-testid="runtime-b"
        :data-runtime-id="runtimeBId"
        :data-input-limit="ISOLATION_REJECT_LIMIT_BYTES"
        :data-demo-rejection="isExpectedLimitRejection ? 'input-size' : ''"
      >
        <h3>Runtime B</h3>
        <div data-testid="runtime-b-status" :data-state="stateB">{{ stateB }}</div>
        <code data-testid="runtime-b-wasm">{{ wasmUrlB }}</code>
        <code data-testid="runtime-b-input-limit">maxInputBytes={{ ISOLATION_REJECT_LIMIT_BYTES }}</code>
        <DocxViewer
          v-if="bufferB"
          :file="bufferB"
          :runtime="runtimeB"
          @load-success="onSuccess('B')"
          @load-error="onError('B', $event)"
        />
        <div
          v-if="errorB"
          data-testid="runtime-b-error"
          :data-error-code="errorB.code"
          :data-error-phase="errorB.phase"
          :data-error-limit="errorB.limit"
          :data-error-actual="errorB.actual"
          :data-error-allowed="errorB.allowed"
        >{{ errorB.code }}: {{ errorB.message }}</div>
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
} from "@arcships/docx-core"
import { DocxViewer } from "@arcships/vue-docx"
import {
  ISOLATION_REJECT_LIMIT_BYTES,
  ISOLATION_SUCCESS_LIMIT_BYTES,
  isInputLimitError,
  normalizeRuntimeError,
  type StructuredRuntimeError,
} from "../runtime-limits"

type State = "loading" | "ready" | "error"

const wasmUrlA = `${bundledDocxWasmUrl}?runtime=alpha`
const wasmUrlB = `${bundledDocxWasmUrl}?runtime=beta`
const diagnosticsA = ref<DocxImportDiagnostic[]>([])
const diagnosticsB = ref<DocxImportDiagnostic[]>([])
const stateA = ref<State>("loading")
const stateB = ref<State>("loading")
const errorA = shallowRef<StructuredRuntimeError | null>(null)
const errorB = shallowRef<StructuredRuntimeError | null>(null)
const loadError = ref("")
const bufferA = shallowRef<ArrayBuffer>()
const bufferB = shallowRef<ArrayBuffer>()

const configA = {
  wasmUrl: wasmUrlA,
  limits: { maxInputBytes: ISOLATION_SUCCESS_LIMIT_BYTES },
  onDiagnostic: (event: DocxImportDiagnostic) => diagnosticsA.value.push(event),
}
const configB = {
  wasmUrl: wasmUrlB,
  // Deliberately below the 36 KB sample: Runtime B must reject the input before
  // the Worker starts while Runtime A loads the same bytes successfully.
  limits: { maxInputBytes: ISOLATION_REJECT_LIMIT_BYTES },
  onDiagnostic: (event: DocxImportDiagnostic) => diagnosticsB.value.push(event),
}
const runtimeA = createDocxRuntime(configA)
const runtimeB = createDocxRuntime(configB)
const runtimeAId = runtimeA.id
const runtimeBId = runtimeB.id

// Mutating the caller objects after construction must not affect either runtime.
configA.limits.maxInputBytes = 1
configB.limits.maxInputBytes = 1

const isExpectedLimitRejection = computed(() =>
  stateA.value === "ready" && stateB.value === "error" && isInputLimitError(errorB.value)
)

const pageState = computed<State>(() => {
  if (stateA.value === "loading" || stateB.value === "loading") return "loading"
  if (stateA.value === "ready" && (stateB.value === "ready" || isExpectedLimitRejection.value)) {
    return "ready"
  }
  return "error"
})

function onSuccess(instance: "A" | "B") {
  if (instance === "A") {
    stateA.value = "ready"
    errorA.value = null
  } else {
    stateB.value = "ready"
    errorB.value = null
  }
}

function onError(instance: "A" | "B", error: Error) {
  const normalized = normalizeRuntimeError(error)
  if (instance === "A") {
    stateA.value = "error"
    errorA.value = normalized
  } else {
    stateB.value = "error"
    errorB.value = normalized
  }
  if (!(instance === "B" && isInputLimitError(normalized))) {
    loadError.value = `${instance}: ${error.message}`
  } else {
    loadError.value = ""
  }
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
