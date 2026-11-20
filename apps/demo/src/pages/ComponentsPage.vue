<template>
  <div class="page">
    <h2>🧩 Components — Verification</h2>
    <p class="desc">Verify SignaturePad, ESignatureFields, FileUpload, FileThumbnail, BoundingBoxCitations, LayoutBlocks, DocumentSplits, Spinner and Tooltip with real fixtures.</p>

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
        <h4 class="demo-subhead">Typed signature</h4>
        <SignaturePad
          mode="typed"
          :model-value="typedSignature"
          show-confirm
          show-cancel
          @update:model-value="typedSignature = $event"
          @confirm="confirmedSignature = $event"
        />
        <p class="sig-preview">Typed: {{ typedSignature || "—" }}</p>
        <p v-if="confirmedSignature" class="sig-preview">Confirmed: {{ confirmedSignature }}</p>
      </section>

      <section>
        <h3>ESignatureFields</h3>
        <div class="split">
          <ESignatureFields
            :fields="signatureFields"
            :can-export="Boolean(signatureFields[0]?.imageDataUrl)"
            show-preview
            @sign="signFirstField"
            @edit="signFirstField"
            @clear="clearFirstSignature"
            @export="exportStatus = 'Signed PDF export requested'"
          />
          <div class="side-panel">
            <h4>State</h4>
            <pre>{{ signatureFields[0]?.imageDataUrl ? "Signed" : "Unsigned" }}</pre>
            <p v-if="exportStatus" class="sig-preview">{{ exportStatus }}</p>
          </div>
        </div>
      </section>

      <section>
        <h3>FileUpload</h3>
        <div class="split">
          <div>
            <h4>Accepted PDF/DOCX, max 2MB</h4>
            <FileUpload
              accept=".pdf,.docx"
              :max-size="2 * 1024 * 1024"
              show-file-list
              title="Click to upload or drop files"
              description="PDF or DOCX, up to 2 MB"
              @files-accepted="uploaded = $event"
            />
            <p v-if="uploaded.length" class="result">Uploaded: {{ uploaded.map(f => f.name).join(', ') }}</p>
          </div>
          <div>
            <h4>Disabled state</h4>
            <FileUpload
              accept=".xlsx"
              disabled
              title="Upload disabled"
              description="This drop zone is not interactive"
            />
          </div>
        </div>
      </section>

      <section>
        <h3>FileThumbnail</h3>
        <div class="thumbnail-grid">
          <FileThumbnail
            :file="{ name: 'large-contract.pdf', type: 'application/pdf' }"
            :size="96"
            show-meta
            show-type-glyph
          />
          <FileThumbnail
            :file="{ name: 'legal-contract.docx', type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }"
            :size="96"
            show-meta
            show-type-glyph
          />
          <FileThumbnail
            :file="{ name: 'financial-model.xlsx', type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }"
            :size="96"
            show-meta
            show-type-glyph
          />
          <FileThumbnail
            :file="{ name: 'contract-page.png', type: 'image/png' }"
            :size="96"
            show-meta
            preview-image-url="/samples/contract-page.png"
          />
          <FileThumbnail
            :file="{ name: 'very-long-unknown-document-name-without-known-type.bin', type: '' }"
            :size="96"
            show-meta
            show-type-glyph
          />
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
        <h4 class="demo-subhead">Review mode (Extend HumanReviewPanel)</h4>
        <div class="split">
          <BoundingBoxCitations
            view="split"
            page-image="/samples/contract-page.png"
            :review-fields="reviewFields"
            :active-field-key="activeReviewField"
            :expected-values="reviewExpected"
            @field-focus="activeReviewField = $event.key"
            @update:expected="onReviewExpected"
          />
          <div class="side-panel">
            <h4>Active field</h4>
            <pre>{{ activeReviewField }}</pre>
            <h4>Expected draft</h4>
            <pre>{{ JSON.stringify(reviewExpected, null, 2) }}</pre>
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
        <h4 class="demo-subhead">Extend block panel + markdown</h4>
        <LayoutBlocks
          view="panel"
          heading="Layout blocks"
          :blocks="extendBlocks"
          render-markdown
          :active-block-id="activeOcrBlock"
          @update:active-block-id="activeOcrBlock = $event"
          @block-click="selectedBlock = $event"
        />
      </section>

      <section>
        <h3>DocumentSplits</h3>
        <div :style="{ height: '640px', border: '1px solid var(--border)', borderRadius: '10px', overflow: 'hidden' }">
          <DocumentSplits
            :splits="splits"
            :thumbnail-images="splitThumbnails"
            :with-frame-divider="false"
            @update:splits="splits = $event"
            @select-page="selectedSplitPage = $event"
          />
        </div>
        <p class="sig-preview">Selected page: {{ selectedSplitPage ?? "—" }}</p>
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
  DocumentSplits,
  ESignatureFields,
  FileThumbnail,
  FileUpload,
  LayoutBlocks,
  SignaturePad,
  Spinner,
  Tooltip,
  type BoundingBoxField,
  type BoundingBoxReviewField,
  type DocumentSplit,
  type ESignatureField,
  type LayoutBlock,
  type OfficeJsonValue,
  type ParsedOcrOutput,
} from "@arcships/vue-ui"

const sigRef = ref<InstanceType<typeof SignaturePad> | null>(null)
const sig = ref<string | null>(null)
const emptyStatus = ref("Empty")
const typedSignature = ref<string | null>(null)
const confirmedSignature = ref<string | null>(null)
const exportStatus = ref("")
const activeReviewField = ref<string | undefined>("vendor_name")
const reviewExpected = ref<Record<string, OfficeJsonValue>>({})
const activeOcrBlock = ref<string | undefined>("block-1")
const selectedSplitPage = ref<number | null>(null)
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

const SIGNATURE_IMAGE =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAPoAAAA6CAYAAACZBESJAAACUklEQVR42u3dy47CMAxAUVqxZ8P/f+Fs+AJmMRoJjWCgbR5OfO4aShP71m764HQCAAAAAAAAAAAA8CFLjx+9XK73d5+53b4W4QHKsEaUfMvnAASs6L8C/1exHyVX2eN0XGJB9GKSv0o2SRan2xILon+URFsSZcvBAfVjo9sierXkIHv96r11bgk/Fk0X444mhAW6OjHZE5fH710u17vYJK/opaqxc/Z4MREfFf1UOuh/K48KEqsrEp/EotcItmQqE49aFVd8krXutZPqWRJpF/tK3vP30EH0lkF2c0dc6cieRPReSUX8WLK12gc50FD0SNe8s1f7aBW19P7sXQPIKPzUomduJSOP9ci+HT14Zz2dqCJ65AnMEugR7iY8sirvvoyOoo9yq+rsso82vp4P02Q58C+Sa64AW+k2Z89Ys06USo5nczXrTT5rZnkeH8qYNXFB9mJVZPSJmSGwniCTC9Ur+gwVZOQAE1w3VFX02RJs9PFo2fGKM8l/BBl1PKp57HhEOfguRwczSxUZcUxW2eOL3eopznfbXiTX9kkjOan3zH2pmO3ZjhcCDig7ycdvw/ds+0hunnsPGCTPOL+ftvil9meRWONUdZI72OzlLLFUGsSiRnzXT37Q21n67gPJgUaCZd8PjM1qCuJWdX9FBaimAMgOaN0BogOA9h1Q0QEQffDKbhYwIq7P7pA84186RcZ9BkSvkvi9E6uVgCO+gMMBJpjo2mAVDFBBgIJYjCM7iA6yg+jOkckOopMdIDrZgUa4TNPpXN0lMhA9kfCkRwu+ATiMGdq08K5VAAAAAElFTkSuQmCC"

const signatureFields = ref<ESignatureField[]>([
  { id: "signature-1", label: "Authorized signature", page: 1, bbox: { x: 300, y: 504, width: 250, height: 58 } },
])

const reviewFields = ref<BoundingBoxReviewField[]>([
  {
    key: "vendor_name",
    schema: { type: "string", title: "Vendor name" },
    actual: "Northwind Industrial Supplies",
    expected: "Northwind Industrial Supplies",
    location: { page: 1, area: { left: 8, top: 9, width: 42, height: 4.5 } },
  },
  {
    key: "invoice_total",
    schema: { type: "number", title: "Invoice total" },
    actual: 2882.58,
    expected: 2882.58,
    location: { page: 1, area: { left: 62, top: 64, width: 30, height: 4 } },
  },
  {
    key: "payment_terms",
    schema: { type: "string", title: "Payment terms", description: "Net terms from the contract." },
    actual: "Net 30",
    expected: "Net 30",
    location: { page: 1, area: { left: 8, top: 88, width: 50, height: 4 } },
  },
])

const extendBlocks = ref<LayoutBlock[]>([
  { id: "block-1", type: "heading", kind: "heading", bbox: [90, 65, 640, 55], page: 1, confidence: 0.98, text: "## MASTER SERVICES AGREEMENT" },
  { id: "block-2", type: "paragraph", kind: "paragraph", bbox: [90, 148, 820, 130], page: 1, confidence: 0.89, text: "Provider will deliver **document intelligence** and `workflow automation` for finance operations." },
  { id: "block-3", type: "table", kind: "table", bbox: [90, 870, 820, 190], page: 2, confidence: 0.77, text: "Service level targets and response windows." },
  { id: "block-4", type: "figure", kind: "figure", bbox: [90, 1195, 650, 60], page: 2, confidence: 0.91, text: "Figure 1 — Payment flow" },
])

const splits = ref<DocumentSplit[]>([
  { id: "split-1", title: "Abstract and intro", pages: ["page-1", "page-2", "page-3"] },
  { id: "split-2", title: "Model architecture", pages: ["page-4", "page-5", "page-6", "page-7"] },
  { id: "split-3", title: "Training and results", pages: ["page-8", "page-9", "page-10"] },
])

const splitThumbnails = Object.fromEntries(
  splits.value.flatMap((group) => group.pages.map((pageId) => [pageId, "/samples/contract-page.png"])),
)

function onReviewExpected(key: string, value: OfficeJsonValue) {
  reviewExpected.value = { ...reviewExpected.value, [key]: value }
}

function signFirstField() {
  signatureFields.value = signatureFields.value.map((field) => ({
    ...field,
    imageDataUrl: SIGNATURE_IMAGE,
  }))
}

function clearFirstSignature() {
  signatureFields.value = signatureFields.value.map((field) => ({
    ...field,
    imageDataUrl: undefined,
  }))
  exportStatus.value = "Signature cleared"
}

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
.demo-subhead { margin-top: 20px; }
.thumbnail-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(min(240px, 100%), 1fr)); gap: 16px; min-width: 0; }
.split { display: grid; grid-template-columns: minmax(0, 1fr) minmax(260px, 380px); gap: 20px; align-items: start; min-width: 0; }
.side-panel { border: 1px solid var(--border); border-radius: var(--radius); padding: 12px; background: var(--muted); }
.side-panel pre { max-height: 260px; overflow: auto; white-space: pre-wrap; font-size: 11px; }
.loading-card { display: inline-flex; align-items: center; gap: 8px; padding: 12px; border: 1px solid var(--border); border-radius: var(--radius); color: var(--muted-foreground); }
@media (max-width: 820px) { .split { grid-template-columns: minmax(0, 1fr); } section { padding: 14px; } }
</style>
