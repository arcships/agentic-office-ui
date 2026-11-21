<script setup lang="ts">
import { computed, ref } from "vue"
import {
  OfficeThemeProvider,
  SchemaBuilder,
  parseJsonSchema,
  type OfficeTheme,
  type SchemaBuilderIssue,
  type SchemaBuilderSchema,
  type SerializedSchema,
} from "@arcships/vue-ui"
import { SCHEMA_BUILDER_FIXTURE, SCHEMA_BUILDER_IMPORT_SAMPLE } from "../lib/schema-builder-fixture"
import { readDemoTheme } from "../lib/demo-visual-settings"

const theme = ref<OfficeTheme>(readDemoTheme())
const readOnly = ref(false)
const schema = ref<SchemaBuilderSchema>(structuredClone(SCHEMA_BUILDER_FIXTURE))
const issues = ref<SchemaBuilderIssue[]>([])
const lastEvent = ref("—")
const exported = ref("—")

const fieldCount = computed(() => {
  let count = 0
  const visit = (fields: SchemaBuilderSchema["properties"]) => {
    for (const field of fields) {
      count += 1
      if (field.properties?.length) visit(field.properties)
    }
  }
  visit(schema.value.properties)
  return count
})

function onValidationChange(next: SchemaBuilderIssue[]): void {
  issues.value = next
  lastEvent.value = `validation-change → ${next.length} issue(s)`
}

function onSchemaChange(): void {
  lastEvent.value = "schema-change"
}

function onExport(next: SerializedSchema): void {
  const text = JSON.stringify(next)
  exported.value = text.length > 160 ? `${text.slice(0, 160)}…` : text
  lastEvent.value = "export"
}

function loadImportSample(): void {
  const result = parseJsonSchema(JSON.parse(SCHEMA_BUILDER_IMPORT_SAMPLE))
  schema.value = result.schema
  lastEvent.value = `load sample → ${result.warnings.length} warning(s)`
}

function reset(): void {
  schema.value = structuredClone(SCHEMA_BUILDER_FIXTURE)
  lastEvent.value = "reset"
}
</script>

<template>
  <OfficeThemeProvider :theme="theme">
    <section class="sb-page" data-testid="schema-builder-demo">
      <header class="sb-page__header">
        <div>
          <p class="sb-page__eyebrow">@arcships/vue-ui</p>
          <h2>SchemaBuilder</h2>
          <p class="sb-page__lead">
            Extraction schema editor: named fields with type, required flag and
            description, nested object/array fields, JSON Schema import/export and
            live validation. Resize to 390 / 768 / 1440 for the responsive states.
          </p>
        </div>
        <div class="sb-page__controls">
          <label>
            Theme
            <select v-model="theme" data-testid="schema-builder-theme">
              <option value="light">Light</option>
              <option value="dark">Dark</option>
              <option value="system">System</option>
            </select>
          </label>
          <label class="sb-page__toggle">
            <input v-model="readOnly" type="checkbox" data-testid="schema-builder-readonly">
            Read only
          </label>
          <button type="button" class="sb-page__button" data-testid="schema-builder-sample" @click="loadImportSample">
            Load sample JSON
          </button>
          <button type="button" class="sb-page__button" data-testid="schema-builder-reset" @click="reset">
            Reset
          </button>
        </div>
      </header>

      <div class="sb-page__layout">
        <SchemaBuilder
          v-model:schema="schema"
          :read-only="readOnly"
          description="Invoice extraction schema used by the review surfaces."
          data-testid="schema-builder-block"
          @schema-change="onSchemaChange"
          @validation-change="onValidationChange"
          @export="onExport"
        />

        <aside class="sb-page__panel">
          <h3>Events</h3>
          <dl>
            <dt>Last event</dt>
            <dd data-testid="schema-builder-last-event">{{ lastEvent }}</dd>
            <dt>Fields</dt>
            <dd data-testid="schema-builder-field-count">{{ fieldCount }}</dd>
            <dt>Validation issues</dt>
            <dd data-testid="schema-builder-issue-count">{{ issues.length }}</dd>
            <dt>Exported JSON</dt>
            <dd data-testid="schema-builder-exported">{{ exported }}</dd>
          </dl>
          <p class="sb-page__hint">
            Reorder with the arrow buttons (no drag and drop), add nested fields on
            object rows and on array rows whose item type is object. “Show JSON”
            renders the serialized JSON Schema; “Import JSON” accepts a JSON Schema
            object and reports unsupported keywords as warnings.
          </p>
        </aside>
      </div>
    </section>
  </OfficeThemeProvider>
</template>

<style scoped>
.sb-page {
  display: flex;
  flex-direction: column;
  gap: 20px;
  padding: 24px;
}
.sb-page__header {
  display: flex;
  flex-wrap: wrap;
  align-items: flex-start;
  justify-content: space-between;
  gap: 16px;
}
.sb-page__eyebrow {
  margin: 0 0 4px;
  color: var(--office-muted-foreground);
  font-size: 12px;
  letter-spacing: .08em;
  text-transform: uppercase;
}
.sb-page h2 { margin: 0 0 6px; font-size: 22px; letter-spacing: -.02em; }
.sb-page__lead { margin: 0; max-width: 720px; color: var(--office-muted-foreground); font-size: 13px; line-height: 1.5; }
.sb-page__controls { display: flex; flex-wrap: wrap; align-items: center; gap: 12px; font-size: 13px; }
.sb-page__controls label { display: inline-flex; align-items: center; gap: 6px; color: var(--office-muted-foreground); }
.sb-page__controls select {
  border: 1px solid var(--office-border);
  border-radius: var(--office-radius-md);
  background: var(--office-background);
  color: var(--office-foreground);
  font: inherit;
  padding: 5px 8px;
}
.sb-page__button {
  border: 1px solid var(--office-border);
  border-radius: var(--office-radius-md);
  background: var(--office-background);
  color: var(--office-foreground);
  font: inherit;
  padding: 5px 10px;
  cursor: pointer;
}
.sb-page__layout {
  display: grid;
  grid-template-columns: minmax(0, 1fr) 260px;
  align-items: start;
  gap: 20px;
}
.sb-page__panel {
  border: 1px solid var(--office-border);
  border-radius: var(--office-radius-lg);
  background: var(--office-background);
  padding: 16px;
}
.sb-page__panel h3 { margin: 0 0 12px; font-size: 14px; font-weight: 650; }
.sb-page__panel dl { display: grid; gap: 4px; margin: 0; }
.sb-page__panel dt { color: var(--office-muted-foreground); font-size: 11px; text-transform: uppercase; letter-spacing: .06em; }
.sb-page__panel dd { margin: 0 0 10px; font-size: 13px; overflow-wrap: anywhere; }
.sb-page__hint { margin: 0; color: var(--office-muted-foreground); font-size: 12px; line-height: 1.5; }
@media (max-width: 900px) {
  .sb-page__layout { grid-template-columns: minmax(0, 1fr); }
}
@media (max-width: 640px) {
  .sb-page { padding: 16px 12px; }
}
</style>
