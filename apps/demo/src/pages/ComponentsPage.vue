<template>
  <div class="page">
    <h2>🧩 Components — Verification</h2>
    <p class="desc">Verify SignaturePad, FileUpload, FileThumbnail, BoundingBoxCitations, LayoutBlocks, Spinner and Tooltip with real fixtures.</p>

    <div class="component-grid">
      <section>
        <h3>SignaturePad</h3>
        <SignaturePad
          ref="sigRef"
          :width="400"
          :height="200"
          @update:signature="onSignatureUpdate"
        />
        <div class="sig-actions">
          <button @click="sigRef?.clear()">Clear</button>
          <button @click="emptyStatus = sigRef?.isEmpty() ? 'Empty' : 'Has signature'">Check empty</button>
        </div>
        <p class="sig-preview">Status: {{ emptyStatus }}</p>
        <p v-if="sig" class="sig-preview">Signature captured ({{ sig.length }} chars)</p>
      </section>

      <section>
        <h3>FileUpload</h3>
        <div class="split">
          <div>
            <h4>Accepted PDF/DOCX, max 2MB</h4>
            <FileUpload
              accept=".pdf,.docx"
              :max-size="2 * 1024 * 1024"
              @files-accepted="uploaded = $event"
            />
            <p v-if="uploaded.length" class="result">Uploaded: {{ uploaded.map(f => f.name).join(', ') }}</p>
          </div>
          <div>
            <h4>Disabled state</h4>
            <FileUpload accept=".xlsx" disabled />
          </div>
        </div>
      </section>

      <section>
        <h3>FileThumbnail</h3>
        <div class="thumbnail-grid">
          <FileThumbnail :file="{ name: 'large-contract.pdf', type: 'application/pdf' }" />
          <FileThumbnail :file="{ name: 'legal-contract.docx', type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }" />
          <FileThumbnail :file="{ name: 'financial-model.xlsx', type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }" />
          <FileThumbnail :file="{ name: 'invoice.png', type: 'image/png', url: '/samples/invoice.png' }" />
          <FileThumbnail :file="{ name: 'very-long-unknown-document-name-without-known-type.bin', type: '' }" />
        </div>
      </section>

      <section>
        <h3>BoundingBoxCitations</h3>
        <div class="split">
          <BoundingBoxCitations
            file="/samples/invoice.png"
            :fields="fields"
            @field-click="selectedField = $event"
          />
          <div class="side-panel">
            <h4>Selected field</h4>
            <pre>{{ selectedField ? JSON.stringify(selectedField, null, 2) : 'Click a field box or list row' }}</pre>
            <h4>Empty state</h4>
            <BoundingBoxCitations file="/samples/invoice.png" :fields="[]" />
          </div>
        </div>
      </section>

      <section>
        <h3>LayoutBlocks</h3>
        <div class="split">
          <LayoutBlocks
            file="/samples/contract-page.png"
            :output="ocrOutput"
            @block-click="selectedBlock = $event"
          />
          <div class="side-panel">
            <h4>Selected block</h4>
            <pre>{{ selectedBlock ? JSON.stringify(selectedBlock, null, 2) : 'Click an OCR block or list row' }}</pre>
            <h4>Empty state</h4>
            <LayoutBlocks file="/samples/contract-page.png" :output="emptyOcrOutput" />
          </div>
        </div>
      </section>

      <section>
        <h3>Spinner</h3>
        <div class="spinner-row">
          <Spinner />
          <div class="loading-card"><Spinner /> Loading document...</div>
          <button class="button-like"><Spinner /> Processing</button>
        </div>
      </section>

      <section>
        <h3>Tooltip</h3>
        <div class="tooltip-row">
          <Tooltip content="Tooltip appears on hover after a short delay." :delay-ms="100">
            <button class="button-like">Hover top</button>
          </Tooltip>
          <Tooltip content="Focusable tooltip for keyboard verification." side="bottom" :delay-ms="100">
            <button class="button-like">Focus bottom</button>
          </Tooltip>
          <Tooltip content="A longer tooltip message that validates wrapping and readable visual styling near viewport edges." side="right" align="start" :delay-ms="100">
            <button class="button-like">Long right</button>
          </Tooltip>
        </div>
      </section>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref } from "vue"
import {
  BoundingBoxCitations,
  FileThumbnail,
  FileUpload,
  LayoutBlocks,
  SignaturePad,
  Spinner,
  Tooltip,
  type BoundingBoxField,
  type LayoutBlock,
  type ParsedOcrOutput,
} from "@extend-ai/vue-extend"

const sigRef = ref<InstanceType<typeof SignaturePad> | null>(null)
const sig = ref<string | null>(null)
const emptyStatus = ref("Empty")
const uploaded = ref<File[]>([])

const fields = ref<BoundingBoxField[]>([
  { id: "vendor", label: "Vendor", page: 1, rect: [0.07, 0.10, 0.48, 0.15], value: "Northwind Industrial Supplies", confidence: 0.96 },
  { id: "tax_id", label: "Tax ID", page: 1, rect: [0.07, 0.135, 0.33, 0.17], value: "US-98-7654321", confidence: 0.83 },
  { id: "invoice_no", label: "Invoice #", page: 1, rect: [0.64, 0.10, 0.91, 0.14], value: "INV-2026-0705", confidence: 0.74 },
  { id: "total", label: "Total Due", page: 1, rect: [0.64, 0.64, 0.92, 0.70], value: "$2,882.58", confidence: 0.42 },
  { id: "terms", label: "Payment Terms", page: 1, rect: [0.07, 0.88, 0.58, 0.91], value: "Net 30", confidence: 0.91 },
])

const ocrOutput = ref<ParsedOcrOutput>({
  width: 1000,
  height: 1330,
  blocks: [
    { id: "title", kind: "title", bbox: [90, 65, 640, 55], text: "MASTER SERVICES AGREEMENT", confidence: 0.98 },
    { id: "h1", kind: "header", bbox: [90, 148, 360, 42], text: "1. Scope of Services", confidence: 0.93 },
    { id: "body1", kind: "text", bbox: [90, 195, 820, 130], text: "Provider will deliver document intelligence and workflow automation.", confidence: 0.89 },
    { id: "h2", kind: "header", bbox: [90, 360, 310, 42], text: "2. Service Levels", confidence: 0.91 },
    { id: "body2", kind: "text", bbox: [90, 410, 820, 130], text: "Provider will maintain availability targets and escalation paths.", confidence: 0.86 },
    { id: "table", kind: "table", bbox: [90, 870, 820, 190], text: "Authorized Signatures", confidence: 0.77 },
    { id: "footer", kind: "footer", bbox: [90, 1195, 650, 30], text: "Confidential", confidence: 0.64 },
  ],
})

const emptyOcrOutput: ParsedOcrOutput = { width: 1000, height: 1330, blocks: [] }
const selectedField = ref<BoundingBoxField | null>(null)
const selectedBlock = ref<LayoutBlock | null>(null)

function onSignatureUpdate(value: string | null) {
  sig.value = value
  emptyStatus.value = value ? "Has signature" : "Empty"
}
</script>

<style scoped>
.page { padding: 24px; max-width: 1200px; margin: 0 auto; width: 100%; min-width: 0; }
h2 { margin-bottom: 4px; }
h3 { font-size: 16px; margin-bottom: 12px; }
h4 { font-size: 13px; margin-bottom: 8px; }
.desc { color: var(--muted-foreground); margin-bottom: 24px; font-size: 14px; }
.component-grid { display: flex; flex-direction: column; gap: 32px; }
section { border: 1px solid var(--border); border-radius: var(--radius); padding: 20px; background: var(--background); }
.sig-actions, .spinner-row, .tooltip-row { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; margin-top: 8px; }
.sig-actions button, .button-like { display: inline-flex; align-items: center; gap: 6px; padding: 6px 12px; border: 1px solid var(--border); border-radius: var(--radius); cursor: pointer; background: var(--background); color: var(--foreground); }
.sig-actions button:hover, .button-like:hover { background: var(--muted); }
.sig-preview, .result { font-size: 12px; color: var(--muted-foreground); margin-top: 8px; }
.thumbnail-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(min(240px, 100%), 1fr)); gap: 16px; min-width: 0; }
.split { display: grid; grid-template-columns: minmax(0, 1fr) minmax(260px, 380px); gap: 20px; align-items: start; min-width: 0; }
.side-panel { border: 1px solid var(--border); border-radius: var(--radius); padding: 12px; background: var(--muted); }
.side-panel pre { max-height: 260px; overflow: auto; white-space: pre-wrap; font-size: 11px; }
.loading-card { display: inline-flex; align-items: center; gap: 8px; padding: 12px; border: 1px solid var(--border); border-radius: var(--radius); color: var(--muted-foreground); }
@media (max-width: 820px) { .split { grid-template-columns: minmax(0, 1fr); } section { padding: 14px; } }
</style>
