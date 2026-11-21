<script setup lang="ts">
import { computed, ref } from "vue"
import { CsvViewer, OfficeThemeProvider, type OfficeTheme } from "@arcships/vue-ui"

const theme = ref<OfficeTheme>("light")
const locale = ref<"en" | "zh-CN">("en")
const height = ref(560)
const status = ref("Choose a sample or upload a CSV file.")

const sampleText = [
  "Opportunity,Account,Amount,Stage,Owner,Notes",
  '"Enterprise renewal","Acme, Inc.",48000,Closing,Ada Lovelace,"Follow-up in Q1; legal review pending."',
  '"Data platform","上海远航",36000,Discovery,Grace Hopper,Initial call complete',
  '"Pilot expansion","Globex ""International""",12000,Proposal,Alan Turing,Needs security review',
  '"Support retainer",Initech,9000,Won,Ada Lovelace,',
  '"Regional rollout",Umbrella Health,27500,Negotiation,Katherine Johnson,Procurement wants a discount',
  '"Analytics add-on",Hooli,7200,Discovery,Grace Hopper,Budget confirmed',
  '"Mobile pilot",Vandelay Industries,5100,Closing,Ada Lovelace,Legal approved',
].join("\n")

const brokenText = 'Name,Team\n"unclosed,Alpha'

const stats = computed(() => {
  const rows = sampleText.trim().split("\n").length - 1
  return `${rows} sample rows · ${sampleText.length} characters`
})

function onLoad(event: { fileName: string | null; headers: string[]; rows: string[][] }) {
  status.value = `Loaded ${event.fileName ?? "CSV data"} — ${event.headers.length} columns × ${event.rows.length} rows.`
}

function onDownload(event: { fileName: string; text: string }) {
  status.value = `Download prepared: ${event.fileName} (${event.text.length} characters).`
}

function onError(error: { code: string }) {
  status.value = `Viewer reported ${error.code}.`
}
</script>

<template>
  <OfficeThemeProvider :theme="theme" :locale="locale">
    <section class="csv-page" data-testid="csv-viewer-page">
      <header class="csv-page__header">
        <div>
          <p class="csv-page__eyebrow">@arcships/vue-ui · issue #178</p>
          <h2>CSV Viewer</h2>
          <p class="csv-page__lead">
            A read-only delimited-text surface on the shared Office chrome:
            RFC 4180 parsing with BOM/UTF-16 detection, a virtualized semantic
            table, cell search, zoom, upload and download. The table replaces
            the upstream Glide canvas, so the visual delta is documented in
            <code>docs/design/csv-viewer-adr.md</code>.
          </p>
        </div>
        <div class="csv-page__controls">
          <label>
            Theme
            <select v-model="theme" data-testid="csv-page-theme">
              <option value="light">Light</option>
              <option value="dark">Dark</option>
              <option value="system">System</option>
            </select>
          </label>
          <label>
            Locale
            <select v-model="locale" data-testid="csv-page-locale">
              <option value="en">English</option>
              <option value="zh-CN">简体中文</option>
            </select>
          </label>
          <label>
            Height
            <select v-model.number="height" data-testid="csv-page-height">
              <option :value="420">420</option>
              <option :value="560">560</option>
              <option :value="700">700</option>
            </select>
          </label>
        </div>
      </header>

      <p class="csv-page__status" data-testid="csv-page-status">{{ status }}</p>

      <section class="csv-page__section">
        <div class="csv-page__section-heading">
          <h3>Delimited text with search</h3>
          <p>
            {{ stats }}. Search highlights every visible match and scrolls the
            active result into the virtual window.
          </p>
        </div>
        <CsvViewer
          :data="sampleText"
          :height="height"
          :locale="locale"
          file-name="sales-pipeline.csv"
          search
          class="csv-page__viewer"
          @load="onLoad"
          @download="onDownload"
          @error="onError"
        />
      </section>

      <section class="csv-page__section">
        <div class="csv-page__section-heading">
          <h3>Empty / error / loading</h3>
          <p>
            The viewer reuses <code>OfficeEmptyState</code>,
            <code>OfficeErrorState</code> and <code>OfficeLoadingState</code>.
            Upload a file from any state.
          </p>
        </div>
        <div class="csv-page__states">
          <article class="csv-page__state-card">
            <h4>Empty</h4>
            <CsvViewer :height="320" :locale="locale" class="csv-page__viewer" />
          </article>
          <article class="csv-page__state-card">
            <h4>Parse error</h4>
            <CsvViewer
              :data="brokenText"
              :height="320"
              :locale="locale"
              class="csv-page__viewer"
            />
          </article>
          <article class="csv-page__state-card">
            <h4>Loading</h4>
            <CsvViewer
              :data="sampleText"
              :height="320"
              :locale="locale"
              loading
              class="csv-page__viewer"
            />
          </article>
        </div>
      </section>
    </section>
  </OfficeThemeProvider>
</template>

<style scoped>
.csv-page {
  display: flex;
  flex-direction: column;
  gap: 28px;
  width: 100%;
  max-width: 1280px;
  margin: 0 auto;
  padding: 28px 24px 48px;
  min-width: 0;
}
.csv-page__header {
  display: flex;
  flex-wrap: wrap;
  justify-content: space-between;
  gap: 20px;
}
.csv-page__eyebrow {
  margin: 0 0 4px;
  color: var(--muted-foreground);
  font-size: 12px;
  letter-spacing: 0.04em;
  text-transform: uppercase;
}
.csv-page__header h2 {
  margin: 0 0 8px;
  font-size: 26px;
}
.csv-page__lead {
  max-width: 720px;
  margin: 0;
  color: var(--muted-foreground);
  line-height: 1.6;
}
.csv-page__lead code {
  font-size: 0.92em;
}
.csv-page__controls {
  display: flex;
  flex-wrap: wrap;
  align-content: flex-start;
  gap: 10px;
}
.csv-page__controls label {
  display: flex;
  flex-direction: column;
  gap: 4px;
  color: var(--muted-foreground);
  font-size: 12px;
}
.csv-page__controls select {
  min-width: 110px;
  padding: 6px 8px;
  border: 1px solid var(--border);
  border-radius: var(--radius-md);
  background: var(--background);
  color: var(--foreground);
}
.csv-page__status {
  margin: 0;
  padding: 8px 12px;
  border: 1px solid var(--border);
  border-radius: var(--radius-md);
  background: var(--muted);
  color: var(--muted-foreground);
  font-size: 13px;
}
.csv-page__section {
  display: flex;
  flex-direction: column;
  gap: 12px;
  min-width: 0;
}
.csv-page__section-heading h3 {
  margin: 0 0 4px;
  font-size: 17px;
}
.csv-page__section-heading p {
  margin: 0;
  color: var(--muted-foreground);
  font-size: 13px;
  line-height: 1.55;
}
.csv-page__viewer {
  border: 1px solid var(--border);
  border-radius: var(--radius-lg);
  overflow: hidden;
}
.csv-page__states {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
  gap: 16px;
}
.csv-page__state-card {
  padding: 12px;
  border: 1px solid var(--border);
  border-radius: var(--radius-lg);
  background: var(--card);
}
.csv-page__state-card h4 {
  margin: 0 0 10px;
  font-size: 13px;
  color: var(--muted-foreground);
}
@media (max-width: 720px) {
  .csv-page {
    padding: 20px 14px 36px;
  }
  .csv-page__controls {
    width: 100%;
  }
}
</style>
