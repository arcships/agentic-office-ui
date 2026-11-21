<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed, inject } from "vue"
import { formatOfficeMessage } from "@arcships/office-interaction"
import OfficeButton from "./ui/OfficeButton.vue"
import OfficeIcon from "./OfficeIcon.vue"
import OfficeInput from "./ui/OfficeInput.vue"
import OfficeSelect from "./ui/OfficeSelect.vue"
import { SCHEMA_BUILDER_CONTEXT } from "./schema-builder-context"
import { useOfficeUiMessages } from "../i18n"
import type {
  SchemaBuilderArrayItemType,
  SchemaBuilderField,
  SchemaBuilderFieldType,
} from "../types"

const props = defineProps<{
  field: SchemaBuilderField
  /** Position inside the sibling list; drives the move buttons. */
  index: number
  siblingCount: number
}>()

const injected = inject(SCHEMA_BUILDER_CONTEXT)
if (!injected) throw new Error("SchemaBuilderFieldRow must be rendered inside SchemaBuilder")
const context = injected
const m = useOfficeUiMessages()
const copy = computed(() => m.value.vueUi.schemaBuilder)

const fieldName = computed(() => props.field.name.trim() || "?")
const issues = computed(() => context.issuesFor(props.field.id))
const issueText = computed(() =>
  issues.value
    .map((issue) => {
      if (issue.code === "empty-name") return copy.value.issueEmptyName
      if (issue.code === "duplicate-name") {
        return formatOfficeMessage(copy.value.issueDuplicateName, { name: issue.name })
      }
      return formatOfficeMessage(copy.value.issueInvalidType, { type: issue.type ?? "?" })
    })
    .join(" "),
)
const hasTypeIssue = computed(() => issues.value.some((issue) => issue.code === "invalid-type"))
const isObject = computed(() => props.field.type === "object")
const isArray = computed(() => props.field.type === "array")
const itemType = computed<SchemaBuilderArrayItemType>(() => props.field.itemType ?? "string")
const showsChildren = computed(
  () => isObject.value || (isArray.value && itemType.value === "object"),
)
const children = computed(() => props.field.properties ?? [])
const canMoveUp = computed(() => !context.readOnly() && props.index > 0)
const canMoveDown = computed(
  () => !context.readOnly() && props.index < props.siblingCount - 1,
)
const removeLabel = computed(() =>
  formatOfficeMessage(copy.value.removeField, { name: fieldName.value }),
)
const addChildLabel = computed(() =>
  formatOfficeMessage(copy.value.addNestedField, { name: fieldName.value }),
)
const moveUpLabel = computed(() =>
  formatOfficeMessage(copy.value.moveFieldUp, { name: fieldName.value }),
)
const moveDownLabel = computed(() =>
  formatOfficeMessage(copy.value.moveFieldDown, { name: fieldName.value }),
)

function onName(value: string): void {
  if (context.readOnly()) return
  context.updateField(props.field.id, { name: value })
}

function onDescription(value: string): void {
  if (context.readOnly()) return
  context.updateField(props.field.id, { description: value })
}

function onType(value: string): void {
  if (context.readOnly()) return
  const type = value as SchemaBuilderFieldType
  const patch: Partial<SchemaBuilderField> = { type }
  if (type === "object" || (type === "array" && itemType.value === "object")) {
    if (!props.field.properties) patch.properties = []
  }
  if (type === "array" && !props.field.itemType) patch.itemType = "string"
  context.updateField(props.field.id, patch)
}

function onItemType(value: string): void {
  if (context.readOnly()) return
  const next = value as SchemaBuilderArrayItemType
  const patch: Partial<SchemaBuilderField> = { itemType: next }
  if (next === "object" && !props.field.properties) patch.properties = []
  context.updateField(props.field.id, patch)
}

function onRequired(event: Event): void {
  if (context.readOnly()) return
  context.updateField(props.field.id, {
    required: (event.target as HTMLInputElement).checked,
  })
}

function onRemove(): void {
  if (context.readOnly()) return
  context.removeField(props.field.id)
}

function onMove(delta: -1 | 1): void {
  if (context.readOnly()) return
  context.moveField(props.field.id, delta)
}

function onAddChild(): void {
  if (context.readOnly()) return
  context.addChild(props.field.id)
}
</script>

<template>
  <div
    :class="['schema-builder__field', { 'is-invalid': issues.length > 0 }]"
    :data-field-id="props.field.id"
    :data-field-type="props.field.type"
    data-slot="schema-builder-field"
  >
    <div class="schema-builder__row">
      <OfficeInput
        class="schema-builder__input schema-builder__input--name"
        :model-value="props.field.name"
        :aria-label="copy.fieldName"
        :placeholder="copy.fieldNamePlaceholder"
        :invalid="issues.length > 0"
        :readonly="context.readOnly()"
        data-slot="schema-builder-field-name"
        @update:model-value="onName"
      />
      <OfficeSelect
        class="schema-builder__select schema-builder__select--type"
        :model-value="props.field.type"
        :options="context.typeOptions()"
        :aria-label="copy.fieldType"
        :invalid="hasTypeIssue"
        :disabled="context.readOnly()"
        size="sm"
        data-slot="schema-builder-field-type"
        @update:model-value="onType"
      />
      <label class="schema-builder__required">
        <input
          type="checkbox"
          :checked="Boolean(props.field.required)"
          :disabled="context.readOnly()"
          data-slot="schema-builder-field-required"
          @change="onRequired"
        />
        <span>{{ copy.required }}</span>
      </label>
      <div v-if="!context.readOnly()" class="schema-builder__actions">
        <button
          type="button"
          class="schema-builder__action"
          :aria-label="moveUpLabel"
          :disabled="!canMoveUp"
          data-slot="schema-builder-field-move-up"
          @click="onMove(-1)"
        >
          <OfficeIcon name="arrow-up" />
        </button>
        <button
          type="button"
          class="schema-builder__action"
          :aria-label="moveDownLabel"
          :disabled="!canMoveDown"
          data-slot="schema-builder-field-move-down"
          @click="onMove(1)"
        >
          <OfficeIcon name="arrow-down" />
        </button>
        <button
          type="button"
          class="schema-builder__action schema-builder__action--remove"
          :aria-label="removeLabel"
          data-slot="schema-builder-field-remove"
          @click="onRemove"
        >
          <OfficeIcon name="trash-2" />
        </button>
      </div>
    </div>

    <OfficeInput
      class="schema-builder__input schema-builder__input--description"
      :model-value="props.field.description ?? ''"
      :aria-label="copy.fieldDescription"
      :placeholder="copy.fieldDescriptionPlaceholder"
      :readonly="context.readOnly()"
      size="sm"
      data-slot="schema-builder-field-description"
      @update:model-value="onDescription"
    />

    <div v-if="isArray" class="schema-builder__array" data-slot="schema-builder-field-array">
      <span class="schema-builder__array-label">{{ copy.arrayItems }}</span>
      <OfficeSelect
        :model-value="itemType"
        :options="context.itemTypeOptions()"
        :aria-label="copy.arrayItemType"
        :disabled="context.readOnly()"
        size="sm"
        data-slot="schema-builder-field-item-type"
        @update:model-value="onItemType"
      />
    </div>

    <p
      v-if="issueText"
      class="schema-builder__issue"
      data-slot="schema-builder-field-issue"
    >
      {{ issueText }}
    </p>

    <div v-if="showsChildren" class="schema-builder__children" data-slot="schema-builder-field-children">
      <div class="schema-builder__children-head">
        <span class="schema-builder__children-title">{{ copy.nestedFields }}</span>
        <OfficeButton
          v-if="!context.readOnly()"
          variant="outline"
          size="xs"
          data-slot="schema-builder-field-add-child"
          @click="onAddChild"
        >
          <OfficeIcon name="plus" />
          {{ copy.addField }}
        </OfficeButton>
      </div>
      <SchemaBuilderFieldRow
        v-for="(child, childIndex) in children"
        :key="child.id"
        :field="child"
        :index="childIndex"
        :sibling-count="children.length"
      />
      <p v-if="children.length === 0" class="schema-builder__children-empty">
        {{ copy.empty }}
      </p>
    </div>
  </div>
</template>

<style scoped>
.schema-builder__field {
  display: flex;
  min-width: 0;
  flex-direction: column;
  gap: var(--office-ui-space-1);
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-md);
  background: var(--office-background);
  padding: var(--office-ui-space-2);
}
.schema-builder__field.is-invalid {
  border-color: var(--office-destructive);
}
.schema-builder__row {
  display: flex;
  min-width: 0;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--office-ui-space-1);
}
.schema-builder__input--name {
  min-width: 120px;
  flex: 1 1 140px;
}
.schema-builder__input--description {
  flex: 1 1 auto;
}
.schema-builder__select--type {
  flex: 0 0 auto;
  min-width: 104px;
}
.schema-builder__required {
  display: inline-flex;
  flex-shrink: 0;
  align-items: center;
  gap: var(--office-ui-space-1);
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
}
.schema-builder__required input {
  width: 14px;
  height: 14px;
  accent-color: var(--office-primary);
  margin: 0;
}
.schema-builder__actions {
  display: inline-flex;
  flex-shrink: 0;
  align-items: center;
  gap: 2px;
  margin-inline-start: auto;
}
.schema-builder__action {
  display: inline-flex;
  width: var(--office-ui-height-xs);
  height: var(--office-ui-height-xs);
  align-items: center;
  justify-content: center;
  margin: 0;
  border: 0;
  border-radius: var(--office-radius-sm);
  background: transparent;
  color: var(--office-muted-foreground);
  cursor: pointer;
  padding: 0;
}
.schema-builder__action:hover:not(:disabled) {
  background: var(--office-accent);
  color: var(--office-foreground);
}
.schema-builder__action--remove:hover:not(:disabled) {
  color: var(--office-destructive);
}
.schema-builder__action:disabled {
  cursor: not-allowed;
  opacity: 0.4;
}
.schema-builder__action:focus-visible {
  outline: 2px solid var(--office-ring);
  outline-offset: 1px;
}
.schema-builder__array {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--office-ui-space-1);
}
.schema-builder__array-label {
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
}
.schema-builder__issue {
  margin: 0;
  color: var(--office-destructive);
  font-size: var(--office-text-xs);
}
.schema-builder__children {
  display: flex;
  min-width: 0;
  flex-direction: column;
  gap: var(--office-ui-space-1);
  border-inline-start: var(--office-border-width) solid var(--office-border);
  margin-inline-start: var(--office-ui-space-2);
  padding-inline-start: var(--office-ui-space-3);
  padding-block: var(--office-ui-space-1);
}
.schema-builder__children-head {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--office-ui-space-2);
  justify-content: space-between;
}
.schema-builder__children-title {
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  font-weight: 600;
}
.schema-builder__children-empty {
  margin: 0;
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
}
@media (max-width: 640px) {
  .schema-builder__actions {
    margin-inline-start: 0;
  }
  .schema-builder__input--name {
    flex: 1 1 100%;
  }
}
</style>
