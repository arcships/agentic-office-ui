<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed, nextTick, provide, ref, watch } from "vue"
import { formatOfficeMessage } from "@arcships/office-interaction"
import OfficeButton from "./ui/OfficeButton.vue"
import OfficeIcon from "./OfficeIcon.vue"
import OfficeTextarea from "./ui/OfficeTextarea.vue"
import SchemaBuilderFieldRow from "./SchemaBuilderFieldRow.vue"
import { SCHEMA_BUILDER_CONTEXT, type SchemaBuilderRowContext } from "./schema-builder-context"
import { useOfficeUiMessages } from "../i18n"
import {
  SCHEMA_BUILDER_ARRAY_ITEM_TYPES,
  SCHEMA_BUILDER_FIELD_TYPES,
  SchemaBuilderParseError,
  createSchemaBuilderField,
  findSchemaBuilderField,
  nextSchemaBuilderFieldName,
  parseSchemaBuilderJsonText,
  serializeSchemaBuilderSchema,
  validateSchemaBuilderSchema,
  type SchemaBuilderImportWarning,
  type SchemaBuilderIssue,
  type SerializedSchema,
} from "../schema-builder"
import type {
  SchemaBuilderArrayItemType,
  SchemaBuilderField,
  SchemaBuilderFieldType,
  SchemaBuilderProps,
  SchemaBuilderSchema,
} from "../types"
import type { OfficeSelectEntry } from "./ui/types"

const props = withDefaults(defineProps<SchemaBuilderProps>(), {
  title: undefined,
  description: undefined,
  readOnly: false,
  loading: false,
  error: undefined,
  emptyMessage: undefined,
  className: "",
})

const emit = defineEmits<{
  "update:schema": [schema: SchemaBuilderSchema]
  "schema-change": [schema: SchemaBuilderSchema]
  "validation-change": [issues: SchemaBuilderIssue[]]
  /** The serialized JSON Schema, emitted when a host copies/exports it. */
  export: [schema: SerializedSchema]
}>()

const m = useOfficeUiMessages()
const copy = computed(() => m.value.vueUi.schemaBuilder)
const rootRef = ref<HTMLElement | null>(null)

const showJson = ref(false)
const showImport = ref(false)
const importText = ref("")
const importError = ref<string | null>(null)
const importWarnings = ref<string[]>([])
const fieldCounter = ref(0)

/**
 * Optimistic local copy. The block edits immediately and emits
 * `update:schema`; a host that passes a new `schema` object replaces the draft,
 * which is how reset/cancel and "load a saved schema" work.
 */
const draft = ref<SchemaBuilderSchema>(props.schema)
watch(
  () => props.schema,
  (next) => {
    if (next !== draft.value) draft.value = next
  },
)

const schema = computed(() => draft.value)
const fields = computed(() => schema.value.properties ?? [])
const issues = computed(() => validateSchemaBuilderSchema(schema.value))
const issuesByField = computed(() => {
  const map = new Map<string, SchemaBuilderIssue[]>()
  for (const issue of issues.value) {
    const list = map.get(issue.fieldId) ?? []
    list.push(issue)
    map.set(issue.fieldId, list)
  }
  return map
})
const jsonText = computed(() =>
  JSON.stringify(serializeSchemaBuilderSchema(schema.value), null, 2),
)
const title = computed(() => props.title ?? copy.value.title)
const emptyCopy = computed(() => props.emptyMessage ?? copy.value.empty)
const showContent = computed(() => !props.error && !props.loading)
const addFieldLabel = computed(() => copy.value.addField)
const showJsonLabel = computed(() => (showJson.value ? copy.value.hideJson : copy.value.showJson))

function typeLabel(type: SchemaBuilderFieldType): string {
  return copy.value.types[type]
}

function typeOptions(): OfficeSelectEntry[] {
  return SCHEMA_BUILDER_FIELD_TYPES.map((type) => ({
    value: type,
    label: typeLabel(type),
  }))
}

function itemTypeOptions(): OfficeSelectEntry[] {
  return SCHEMA_BUILDER_ARRAY_ITEM_TYPES.map((type) => ({
    value: type,
    label: typeLabel(type),
  }))
}

function cloneSchema(value: SchemaBuilderSchema): SchemaBuilderSchema {
  return JSON.parse(JSON.stringify(value)) as SchemaBuilderSchema
}

function commit(next: SchemaBuilderSchema): void {
  draft.value = next
  emit("update:schema", next)
  emit("schema-change", next)
}

function nextFieldId(next: SchemaBuilderSchema): string {
  for (;;) {
    fieldCounter.value += 1
    const candidate = `schema-field-${fieldCounter.value}`
    if (!findSchemaBuilderField(next, candidate)) return candidate
  }
}

function isNestedContainer(field: SchemaBuilderField): boolean {
  if (field.type === "object") return true
  return field.type === "array" && (field.itemType ?? "string") === "object"
}

function focusFieldName(fieldId: string): void {
  void nextTick(() => {
    const row = rootRef.value?.querySelector<HTMLElement>(`[data-field-id="${fieldId}"]`)
    const input = row?.querySelector<HTMLInputElement>("input")
    input?.focus()
  })
}

function addField(parentId: string | null = null): void {
  if (props.readOnly) return
  const next = cloneSchema(schema.value)
  let siblings = next.properties
  if (parentId) {
    const parent = findSchemaBuilderField(next, parentId)?.field
    if (!parent || !isNestedContainer(parent)) return
    parent.properties = parent.properties ?? []
    siblings = parent.properties
  }
  const field = createSchemaBuilderField(nextFieldId(next), "string")
  field.name = nextSchemaBuilderFieldName(siblings)
  siblings.push(field)
  commit(next)
  focusFieldName(field.id)
}

function updateField(fieldId: string, patch: Partial<SchemaBuilderField>): void {
  if (props.readOnly) return
  const next = cloneSchema(schema.value)
  const location = findSchemaBuilderField(next, fieldId)
  if (!location) return
  Object.assign(location.field, patch)
  commit(next)
}

function removeField(fieldId: string): void {
  if (props.readOnly) return
  const next = cloneSchema(schema.value)
  const location = findSchemaBuilderField(next, fieldId)
  if (!location) return
  const index = location.siblings.findIndex((field) => field.id === fieldId)
  if (index < 0) return
  location.siblings.splice(index, 1)
  commit(next)
}

function moveField(fieldId: string, delta: -1 | 1): void {
  if (props.readOnly) return
  const next = cloneSchema(schema.value)
  const location = findSchemaBuilderField(next, fieldId)
  if (!location) return
  const index = location.siblings.findIndex((field) => field.id === fieldId)
  const target = index + delta
  if (index < 0 || target < 0 || target >= location.siblings.length) return
  const [moved] = location.siblings.splice(index, 1)
  if (!moved) return
  location.siblings.splice(target, 0, moved)
  commit(next)
}

const rowContext: SchemaBuilderRowContext = {
  readOnly: () => props.readOnly,
  issuesFor: (fieldId: string) => issuesByField.value.get(fieldId) ?? [],
  updateField,
  removeField,
  moveField,
  addChild: (fieldId: string) => addField(fieldId),
  typeLabel,
  typeOptions,
  itemTypeOptions,
}

provide(SCHEMA_BUILDER_CONTEXT, rowContext)

watch(
  issues,
  (next) => emit("validation-change", next),
  { immediate: true },
)

function exportJson(): void {
  emit("export", serializeSchemaBuilderSchema(schema.value))
}

function formatWarning(warning: SchemaBuilderImportWarning): string {
  switch (warning.code) {
    case "unsupported-item-type":
      return formatOfficeMessage(copy.value.importUnsupportedItemType, {
        path: warning.path,
        type: warning.type ?? "?",
      })
    case "unsupported-item-shape":
      return formatOfficeMessage(copy.value.importUnsupportedItemShape, {
        path: warning.path,
      })
    case "missing-name":
      return formatOfficeMessage(copy.value.importMissingName, { path: warning.path })
    default:
      return formatOfficeMessage(copy.value.importUnsupportedType, {
        path: warning.path,
        type: warning.type ?? "?",
      })
  }
}

function formatImportError(error: unknown): string {
  if (error instanceof SchemaBuilderParseError) {
    // A root-level shape problem has no path to point at; use the root copy.
    if (error.code === "invalid-properties" && error.path) {
      return formatOfficeMessage(copy.value.importInvalidProperties, { path: error.path })
    }
    if (error.code === "invalid-property") {
      return formatOfficeMessage(copy.value.importInvalidProperty, { path: error.path })
    }
    return copy.value.importInvalidRoot
  }
  const message = error instanceof Error ? error.message : String(error)
  return formatOfficeMessage(copy.value.importInvalidJson, { message })
}

function toggleImport(): void {
  if (props.readOnly) return
  showImport.value = !showImport.value
  if (!showImport.value) {
    importText.value = ""
    importError.value = null
  }
}

function applyImport(): void {
  if (props.readOnly) return
  importError.value = null
  importWarnings.value = []
  if (!importText.value.trim()) {
    importError.value = copy.value.importEmpty
    return
  }
  try {
    const result = parseSchemaBuilderJsonText(importText.value)
    commit(result.schema)
    importWarnings.value = result.warnings.map(formatWarning)
    if (result.warnings.length === 0) {
      showImport.value = false
      importText.value = ""
    }
  } catch (error) {
    importError.value = formatImportError(error)
  }
}

function cancelImport(): void {
  showImport.value = false
  importText.value = ""
  importError.value = null
  importWarnings.value = []
}

function fieldPath(issue: SchemaBuilderIssue): string {
  return issue.path
}

defineExpose({
  /** Serialized JSON Schema (draft 2020-12 subset) of the current draft. */
  toJSONSchema: (): SerializedSchema => serializeSchemaBuilderSchema(schema.value),
  /** Current validation issues, by field id. */
  issues: (): SchemaBuilderIssue[] => issues.value,
})
</script>

<template>
  <section
    ref="rootRef"
    :class="['schema-builder', 'office-ui', props.className]"
    :data-read-only="props.readOnly ? 'true' : undefined"
    data-slot="schema-builder"
  >
    <header class="schema-builder__header" data-slot="schema-builder-header">
      <div class="schema-builder__heading">
        <h2 class="schema-builder__title">{{ title }}</h2>
        <p v-if="props.description" class="schema-builder__description">{{ props.description }}</p>
      </div>
      <div class="schema-builder__toolbar">
        <OfficeButton
          v-if="!props.readOnly"
          variant="default"
          size="sm"
          :aria-label="addFieldLabel"
          data-slot="schema-builder-add"
          @click="addField(null)"
        >
          <OfficeIcon name="plus" />
          {{ copy.addField }}
        </OfficeButton>
        <OfficeButton
          variant="outline"
          size="sm"
          :aria-pressed="showJson"
          data-slot="schema-builder-json-toggle"
          @click="showJson = !showJson"
        >
          <OfficeIcon name="file-text" />
          {{ showJsonLabel }}
        </OfficeButton>
        <OfficeButton
          v-if="!props.readOnly"
          variant="outline"
          size="sm"
          :aria-pressed="showImport"
          data-slot="schema-builder-import-toggle"
          @click="toggleImport"
        >
          <OfficeIcon name="download" />
          {{ copy.importJson }}
        </OfficeButton>
      </div>
    </header>

    <div v-if="props.error" class="schema-builder__state schema-builder__state--error" role="alert" data-slot="schema-builder-error">
      {{ props.error }}
    </div>
    <div v-else-if="props.loading" class="schema-builder__state" role="status" data-slot="schema-builder-loading">
      {{ copy.loading }}
    </div>

    <div v-else-if="showContent" class="schema-builder__body" data-slot="schema-builder-body">
      <div
        v-if="showImport && !props.readOnly"
        class="schema-builder__import"
        data-slot="schema-builder-import"
      >
        <OfficeTextarea
          :model-value="importText"
          :aria-label="copy.importJson"
          :placeholder="copy.importPlaceholder"
          :invalid="Boolean(importError)"
          :rows="6"
          data-slot="schema-builder-import-input"
          @update:model-value="importText = $event"
        />
        <p v-if="importError" class="schema-builder__issue" role="alert" data-slot="schema-builder-import-error">
          {{ importError }}
        </p>
        <ul v-if="importWarnings.length" class="schema-builder__warnings" role="status" data-slot="schema-builder-import-warnings">
          <li v-for="(warning, index) in importWarnings" :key="index">{{ warning }}</li>
        </ul>
        <div class="schema-builder__import-actions">
          <OfficeButton variant="default" size="sm" data-slot="schema-builder-import-apply" @click="applyImport">
            {{ copy.importApply }}
          </OfficeButton>
          <OfficeButton variant="outline" size="sm" data-slot="schema-builder-import-cancel" @click="cancelImport">
            {{ copy.importCancel }}
          </OfficeButton>
        </div>
      </div>

      <p v-if="fields.length === 0" class="schema-builder__state" role="status" data-slot="schema-builder-empty">
        {{ emptyCopy }}
      </p>
      <div v-else class="schema-builder__fields" data-slot="schema-builder-fields">
        <SchemaBuilderFieldRow
          v-for="(field, index) in fields"
          :key="field.id"
          :field="field"
          :index="index"
          :sibling-count="fields.length"
        />
      </div>

      <ul v-if="issues.length" class="schema-builder__issues" role="alert" data-slot="schema-builder-issues">
        <li
          v-for="issue in issues"
          :key="`${issue.fieldId}-${issue.code}`"
          :data-issue-code="issue.code"
          :data-issue-path="fieldPath(issue)"
        >
          {{ fieldPath(issue) }}
        </li>
      </ul>

      <div v-if="showJson" class="schema-builder__json" data-slot="schema-builder-json-panel">
        <div class="schema-builder__json-head">
          <span class="schema-builder__json-title">{{ copy.jsonTitle }}</span>
          <OfficeButton variant="outline" size="xs" data-slot="schema-builder-export" @click="exportJson">
            {{ copy.exportJson }}
          </OfficeButton>
        </div>
        <pre class="schema-builder__json-body" data-slot="schema-builder-json">{{ jsonText }}</pre>
      </div>
    </div>
  </section>
</template>

<style scoped>
.schema-builder {
  display: flex;
  min-width: 0;
  flex-direction: column;
  gap: var(--office-ui-space-3);
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-lg);
  background: var(--office-background);
  color: var(--office-foreground);
  padding: var(--office-ui-space-4);
}
.schema-builder,
.schema-builder * {
  box-sizing: border-box;
}
.schema-builder__header {
  display: flex;
  flex-wrap: wrap;
  align-items: flex-start;
  gap: var(--office-ui-space-2);
  justify-content: space-between;
}
.schema-builder__heading {
  min-width: 0;
}
.schema-builder__title {
  margin: 0;
  font-size: var(--office-text-base);
  font-weight: 600;
  line-height: 1.3;
}
.schema-builder__description {
  margin: var(--office-ui-space-1) 0 0;
  color: var(--office-muted-foreground);
  font-size: var(--office-text-sm);
}
.schema-builder__toolbar {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--office-ui-space-2);
}
.schema-builder__body {
  display: flex;
  min-width: 0;
  flex-direction: column;
  gap: var(--office-ui-space-2);
}
.schema-builder__fields {
  display: flex;
  min-width: 0;
  flex-direction: column;
  gap: var(--office-ui-space-2);
}
.schema-builder__state {
  display: flex;
  min-height: 96px;
  align-items: center;
  justify-content: center;
  margin: 0;
  padding: var(--office-ui-space-6);
  color: var(--office-muted-foreground);
  font-size: var(--office-text-sm);
  text-align: center;
}
.schema-builder__state--error {
  color: var(--office-destructive);
}
.schema-builder__import {
  display: flex;
  flex-direction: column;
  gap: var(--office-ui-space-2);
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-md);
  background: color-mix(in oklab, var(--office-muted) 30%, var(--office-background));
  padding: var(--office-ui-space-3);
}
.schema-builder__import-actions {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--office-ui-space-2);
}
.schema-builder__issue {
  margin: 0;
  color: var(--office-destructive);
  font-size: var(--office-text-xs);
}
.schema-builder__warnings,
.schema-builder__issues {
  display: flex;
  flex-direction: column;
  gap: var(--office-ui-space-1);
  margin: 0;
  padding-inline-start: var(--office-ui-space-4);
  font-size: var(--office-text-xs);
}
.schema-builder__warnings {
  color: var(--office-muted-foreground);
}
.schema-builder__issues {
  color: var(--office-destructive);
}
.schema-builder__json {
  display: flex;
  flex-direction: column;
  gap: var(--office-ui-space-1);
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-md);
  background: color-mix(in oklab, var(--office-muted) 24%, var(--office-background));
  padding: var(--office-ui-space-2);
}
.schema-builder__json-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--office-ui-space-2);
}
.schema-builder__json-title {
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  font-weight: 600;
}
.schema-builder__json-body {
  margin: 0;
  max-height: 280px;
  overflow: auto;
  color: var(--office-foreground);
  font-family: var(--office-font-mono);
  font-size: var(--office-text-xs);
  line-height: 1.5;
  white-space: pre-wrap;
  word-break: break-word;
}
@media (max-width: 640px) {
  .schema-builder {
    padding: var(--office-ui-space-3);
  }
  .schema-builder__toolbar {
    width: 100%;
  }
}
</style>
