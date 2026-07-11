<template>
  <main class="page" data-testid="runtime-worker-page">
    <h2>XLSX Worker 超时、取消与恢复</h2>
    <p>本页只使用公开 Runtime、Worker 工厂和资源上限，用于验证故障时不会静默回到主线程解析。</p>

    <div data-testid="page-status" data-state="ready">ready</div>
    <section class="controls">
      <button data-testid="worker-timeout-run" type="button" :disabled="busy" @click="runTimeoutRecovery">
        运行超时与恢复
      </button>
      <button data-testid="worker-cancel-run" type="button" :disabled="busy" @click="runCancellation">
        运行取消
      </button>
      <button data-testid="worker-heartbeat" type="button" @click="heartbeat += 1">
        主线程交互
      </button>
    </section>

    <dl>
      <dt>状态</dt>
      <dd data-testid="worker-case-status" :data-state="state">{{ state }}</dd>
      <dt>错误码</dt>
      <dd data-testid="worker-error-code">{{ errorCode }}</dd>
      <dt>错误阶段</dt>
      <dd data-testid="worker-error-phase">{{ errorPhase }}</dd>
      <dt>已终止 Worker</dt>
      <dd data-testid="worker-terminated-count">{{ terminatedCount }}</dd>
      <dt>恢复后工作表</dt>
      <dd data-testid="worker-recovered-sheets">{{ recoveredSheets }}</dd>
      <dt>交互计数</dt>
      <dd data-testid="worker-heartbeat-count">{{ heartbeat }}</dd>
    </dl>

    <pre data-testid="worker-diagnostics">{{ JSON.stringify(diagnostics, null, 2) }}</pre>
  </main>
</template>

<script setup lang="ts">
import { onBeforeUnmount, ref } from "vue"
import {
  createXlsxRuntime,
  type XlsxRuntime,
  type XlsxRuntimeDiagnostic,
} from "@arcships/xlsx-core/runtime"

type CaseState = "idle" | "running" | "timeout" | "cancelled" | "recovered" | "error"

const state = ref<CaseState>("idle")
const busy = ref(false)
const errorCode = ref("")
const errorPhase = ref("")
const terminatedCount = ref(0)
const recoveredSheets = ref(0)
const heartbeat = ref(0)
const diagnostics = ref<unknown[]>([])
const runtimes = new Set<XlsxRuntime>()
let activeAbortController: AbortController | null = null

function recordDiagnostic(event: XlsxRuntimeDiagnostic | Record<string, unknown>): void {
  diagnostics.value = [...diagnostics.value.slice(-49), event]
}

function trackedHangingWorker(): Worker {
  const worker = new Worker("/fixtures/hanging-xlsx-worker.js", { type: "module" })
  const terminate = worker.terminate.bind(worker)
  let terminated = false
  worker.terminate = () => {
    if (!terminated) {
      terminated = true
      terminatedCount.value += 1
    }
    terminate()
  }
  return worker
}

async function sampleBuffer(): Promise<ArrayBuffer> {
  const response = await fetch("/samples/sales-table.xlsx", { cache: "no-store" })
  if (!response.ok) throw new Error(`样本返回 ${response.status}`)
  return response.arrayBuffer()
}

function normalizeFailure(error: unknown): void {
  const value = error as { code?: unknown; phase?: unknown; message?: unknown }
  errorCode.value = typeof value?.code === "string" ? value.code : "WORKER_FAILED"
  errorPhase.value = typeof value?.phase === "string" ? value.phase : "parse"
  recordDiagnostic({
    type: "worker-error",
    code: errorCode.value,
    phase: errorPhase.value,
    message: typeof value?.message === "string" ? value.message : String(error),
  })
}

function createTrackedRuntime(maxParseMs: number): XlsxRuntime {
  const runtime = createXlsxRuntime({
    createWorker: trackedHangingWorker,
    limits: { maxParseMs },
    onDiagnostic: recordDiagnostic,
  })
  runtimes.add(runtime)
  return runtime
}

async function recover(buffer: ArrayBuffer): Promise<void> {
  const runtime = createXlsxRuntime({ onDiagnostic: recordDiagnostic })
  runtimes.add(runtime)
  const client = runtime.createWorkerClient()
  const result = await client.loadWorkbook(buffer)
  recoveredSheets.value = result.sheets.length
  client.dispose()
  runtime.dispose()
  runtimes.delete(runtime)
}

function resetCase(): void {
  state.value = "running"
  errorCode.value = ""
  errorPhase.value = ""
  recoveredSheets.value = 0
  diagnostics.value = []
}

async function runTimeoutRecovery(): Promise<void> {
  if (busy.value) return
  busy.value = true
  resetCase()
  try {
    const buffer = await sampleBuffer()
    const runtime = createTrackedRuntime(80)
    const client = runtime.createWorkerClient()
    try {
      await client.loadWorkbook(buffer)
      throw new Error("Worker 应超时但返回了成功。")
    } catch (error) {
      normalizeFailure(error)
      if (errorCode.value !== "TIMEOUT") throw error
      state.value = "timeout"
    } finally {
      client.dispose()
      runtime.dispose()
      runtimes.delete(runtime)
    }
    await recover(buffer)
    state.value = "recovered"
  } catch (error) {
    normalizeFailure(error)
    state.value = "error"
  } finally {
    busy.value = false
  }
}

async function runCancellation(): Promise<void> {
  if (busy.value) return
  busy.value = true
  resetCase()
  try {
    const buffer = await sampleBuffer()
    const runtime = createTrackedRuntime(5_000)
    const client = runtime.createWorkerClient()
    const controller = new AbortController()
    activeAbortController = controller
    const pending = client.loadWorkbook(buffer, false, false, controller.signal)
    window.setTimeout(() => controller.abort(), 40)
    try {
      await pending
      throw new Error("Worker 应取消但返回了成功。")
    } catch (error) {
      const name = error instanceof Error ? error.name : ""
      if (name !== "AbortError") throw error
      errorCode.value = "ABORTED"
      errorPhase.value = "parse"
      recordDiagnostic({ type: "worker-cancelled", code: "ABORTED", phase: "parse" })
      state.value = "cancelled"
    } finally {
      activeAbortController = null
      client.dispose()
      runtime.dispose()
      runtimes.delete(runtime)
    }
  } catch (error) {
    normalizeFailure(error)
    state.value = "error"
  } finally {
    busy.value = false
  }
}

onBeforeUnmount(() => {
  activeAbortController?.abort()
  activeAbortController = null
  for (const runtime of runtimes) runtime.dispose()
  runtimes.clear()
})
</script>

<style scoped>
.page { max-width: 900px; margin: 0 auto; padding: 24px; }
.controls { display: flex; flex-wrap: wrap; gap: 8px; margin: 16px 0; }
button { font: inherit; padding: 8px 12px; }
dl { display: grid; grid-template-columns: 180px minmax(0, 1fr); gap: 6px 12px; }
dt { font-weight: 600; }
dd { margin: 0; overflow-wrap: anywhere; }
pre { max-height: 320px; overflow: auto; white-space: pre-wrap; }
</style>
