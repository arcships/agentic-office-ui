<template>
  <main class="consumer-shell">
    <h1>TGZ-only consumer verification</h1>
    <p data-testid="consumer-ready" :data-state="overallState">{{ overallState }}</p>
    <section>
      <h2>DOCX</h2>
      <p data-testid="consumer-docx-state" :data-state="docxState">{{ docxState }}</p>
      <p data-testid="consumer-docx-source">{{ docxSource }}</p>
      <DocxViewer
        v-if="docxFile"
        :file="docxFile"
        :runtime="docxRuntime"
        :layout-options="{ pageWidth: 640, pageHeight: 860 }"
        @load-success="onDocxSuccess"
        @load-error="onDocxError"
      />
    </section>
    <section>
      <h2>XLSX</h2>
      <p data-testid="consumer-xlsx-state" :data-state="xlsxController.sourceState">{{ xlsxController.sourceState }}</p>
      <p data-testid="consumer-xlsx-worker">{{ xlsxController.isWorkerBacked ? "worker" : "not-worker" }}</p>
      <p v-if="xlsxController.sourceError" data-testid="consumer-xlsx-error">{{ xlsxController.sourceError.code }}</p>
      <XlsxViewer :controller="xlsxController" height="480px" />
    </section>
    <section>
      <h2>PDF</h2>
      <p data-testid="consumer-pdf-state" :data-state="pdfState">{{ pdfState }}</p>
      <p v-if="pdfError" data-testid="consumer-pdf-error">{{ pdfError }}</p>
      <PdfViewer
        :source="pdfSource"
        :url-policy="pdfUrlPolicy"
        file-name="sample.pdf"
        :show-toolbar="false"
        @document-load-success="onPdfSuccess"
        @document-load-error="onPdfError"
      />
    </section>
    <details open>
      <summary>Public diagnostics</summary>
      <pre data-testid="consumer-diagnostics">{{ JSON.stringify(diagnostics, null, 2) }}</pre>
    </details>
  </main>
</template>

<script setup lang="ts">
import { computed, onMounted, ref, shallowRef } from "vue";
import { createDocxRuntime, type DocxImportDiagnostic, type DocxImportResult } from "@extend-ai/docx-core";
import docxWorkerUrl from "@extend-ai/docx-core/worker?worker&url";
import docxWasmUrl from "@extend-ai/docx-core/assets/docx_wasm_bg.wasm?url";
import { setWasmSource } from "@extend-ai/xlsx-core";
import xlsxWorkerUrl from "@extend-ai/xlsx-core/worker?worker&url";
import xlsxWasmUrl from "@extend-ai/xlsx-core/assets/duke_sheets_wasm_bg.wasm?url";
import { DocxViewer } from "@extend-ai/vue-docx";
import { XlsxViewer, useXlsxViewerController, type XlsxDiagnostic } from "@extend-ai/vue-xlsx";
import { PdfViewer, type PdfLoadError, type PdfSource, type PdfUrlPolicy } from "@extend-ai/vue-extend";

type LoadState = "loading" | "ready" | "error";
const diagnostics = ref<unknown[]>([]);
const docxFile = shallowRef<ArrayBuffer>();
const docxState = ref<LoadState>("loading");
const docxSource = ref("pending");
const pdfState = ref<LoadState>("loading");
const pdfError = ref("");
const origin = globalThis.location.origin;

function addDiagnostic(event: unknown) {
  diagnostics.value = [...diagnostics.value.slice(-19), event];
}

const docxRuntime = createDocxRuntime({
  allowMainThreadFallback: false,
  wasmUrl: docxWasmUrl,
  workerUrl: docxWorkerUrl,
  onDiagnostic: (event: DocxImportDiagnostic) => addDiagnostic({ package: "docx", ...event }),
});

setWasmSource(xlsxWasmUrl);
const xlsxController = useXlsxViewerController({
  src: "/fixtures/sales-table.xlsx",
  fileName: "sales-table.xlsx",
  readOnly: true,
  useWorker: true,
  workerUrl: xlsxWorkerUrl,
  urlPolicy: {
    baseUrl: globalThis.location.href,
    allowRelativeUrl: true,
    allowedProtocols: ["http:"],
    allowedOrigins: [origin],
    allowHttpOnLocalhost: true,
  },
  onDiagnostic: (event: XlsxDiagnostic) => addDiagnostic({ package: "xlsx", ...event }),
});

const pdfSource: PdfSource = { kind: "url", url: "/fixtures/sample.pdf" };
const pdfUrlPolicy: PdfUrlPolicy = {
  baseUrl: globalThis.location.href,
  allowRelativeUrl: true,
  allowedProtocols: ["http:"],
  allowedOrigins: [origin],
  allowHttpOnLocalhost: true,
};

const overallState = computed<LoadState>(() => {
  if (docxState.value === "error" || xlsxController.sourceState === "error" || pdfState.value === "error") return "error";
  if (docxState.value === "ready" && xlsxController.sourceState === "ready" && pdfState.value === "ready") return "ready";
  return "loading";
});

onMounted(async () => {
  try {
    const response = await fetch("/fixtures/invoice-table.docx");
    if (!response.ok) throw new Error(`fixture request failed: ${response.status}`);
    docxFile.value = await response.arrayBuffer();
  } catch (error) {
    docxState.value = "error";
    addDiagnostic({ package: "docx", type: "fixture-error", message: String(error) });
  }
});

function onDocxSuccess(result: DocxImportResult) {
  docxSource.value = result.source;
  docxState.value = "ready";
}

function onDocxError(error: Error) {
  docxState.value = "error";
  addDiagnostic({ package: "docx", type: "load-error", message: error.message });
}

function onPdfSuccess() {
  pdfState.value = "ready";
}

function onPdfError(error: PdfLoadError) {
  pdfState.value = "error";
  pdfError.value = `${error.code}: ${error.message}`;
}
</script>

<style>
* { box-sizing: border-box; }
body { margin: 0; font-family: system-ui, sans-serif; color: #172033; background: #f6f8fc; }
.consumer-shell { max-width: 1180px; margin: 0 auto; padding: 20px; }
section { margin: 18px 0; padding: 14px; border: 1px solid #cbd5e1; border-radius: 8px; background: white; }
[data-state="error"] { color: #b91c1c; font-weight: 700; }
[data-state="ready"] { color: #047857; font-weight: 700; }
pre { max-height: 240px; overflow: auto; background: #0f172a; color: #e2e8f0; padding: 12px; }
</style>
