<template>
  <OfficeCard
    class="docx-link-card"
    :class="documentTheme === 'dark' ? 'docx-link-card--dark' : undefined"
    :data-office-theme="documentTheme"
    data-office-density="comfortable"
    role="dialog"
    :aria-label="m.docx.linkCard.label"
    data-testid="docx-link-card"
    :data-docx-link-card="true"
    :style="cardStyle"
    @pointerenter="emit('pointerenter')"
    @pointerleave="emit('pointerleave')"
    @keydown.esc.stop.prevent="close()"
  >
    <template v-if="editing">
      <label class="docx-link-card__field">
        <span class="docx-link-card__label">{{ m.docx.linkCard.url }}</span>
        <input
          ref="inputEl"
          v-model="draft"
          type="text"
          class="office-input office-input--sm docx-link-card__input"
          :placeholder="m.docx.linkCard.urlPlaceholder"
          :aria-label="m.docx.linkCard.url"
          data-testid="docx-link-card-input"
          @keydown.enter.stop.prevent="save"
        />
      </label>
      <div class="docx-link-card__actions">
        <OfficeButton
          type="button"
          size="xs"
          class="docx-link-card__action docx-link-card__action--primary"
          data-testid="docx-link-card-save"
          @click="save"
        >
          {{ m.docx.linkCard.save }}
        </OfficeButton>
        <OfficeButton
          type="button"
          variant="outline"
          size="xs"
          class="docx-link-card__action"
          data-testid="docx-link-card-cancel"
          @click="cancel"
        >
          {{ m.docx.linkCard.cancel }}
        </OfficeButton>
      </div>
    </template>
    <template v-else>
      <span class="docx-link-card__url" :title="href" data-testid="docx-link-card-url">
        {{ displayHref }}
      </span>
      <div class="docx-link-card__actions">
        <OfficeButton
          type="button"
          variant="outline"
          size="xs"
          class="docx-link-card__action"
          data-testid="docx-link-card-open"
          :title="m.docx.linkCard.open"
          @click="emit('open', href)"
        >
          {{ m.docx.linkCard.open }}
        </OfficeButton>
        <OfficeButton
          v-if="editable"
          type="button"
          variant="outline"
          size="xs"
          class="docx-link-card__action"
          data-testid="docx-link-card-edit"
          :title="m.docx.linkCard.edit"
          @click="startEdit"
        >
          {{ m.docx.linkCard.edit }}
        </OfficeButton>
        <OfficeButton
          v-if="editable"
          type="button"
          variant="ghost"
          size="xs"
          class="docx-link-card__action docx-link-card__action--danger"
          data-testid="docx-link-card-remove"
          :title="m.docx.linkCard.remove"
          @click="emit('remove', href)"
        >
          {{ m.docx.linkCard.remove }}
        </OfficeButton>
      </div>
    </template>
  </OfficeCard>
</template>

<script setup lang="ts">
import { computed, nextTick, ref, watch } from "vue"
import { OfficeButton, OfficeCard } from "@arcships/vue-ui"
import { useDocxMessages } from "../i18n"

const props = withDefaults(
  defineProps<{
    href: string
    /** Viewport coordinates of the hovered link. */
    x?: number
    y?: number
    /** Editing/removing needs a selection the controller can rewrite. */
    editable?: boolean
    documentTheme?: "light" | "dark"
  }>(),
  { x: 0, y: 0, editable: true, documentTheme: "light" }
)

const emit = defineEmits<{
  (event: "open", href: string): void
  (event: "edit", href: string): void
  (event: "remove", href: string): void
  (event: "close"): void
  (event: "pointerenter"): void
  (event: "pointerleave"): void
}>()

const m = useDocxMessages()

const editing = ref(false)
const draft = ref("")
const inputEl = ref<HTMLInputElement | undefined>()

const internal = computed(() => props.href.startsWith("#"))
const displayHref = computed(() =>
  internal.value ? m.value.docx.linkCard.internal : props.href
)

const cardStyle = computed(() => ({
  position: "fixed" as const,
  left: `${Math.max(0, Math.round(props.x))}px`,
  top: `${Math.max(0, Math.round(props.y))}px`,
  zIndex: 70,
}))

watch(
  () => props.href,
  () => {
    editing.value = false
    draft.value = ""
  }
)

function startEdit(): void {
  draft.value = props.href
  editing.value = true
  void nextTick(() => inputEl.value?.focus?.())
}

function cancel(): void {
  editing.value = false
  draft.value = ""
}

function save(): void {
  const value = draft.value.trim()
  if (value === "") {
    // An empty address is the remove action, not a silent no-op.
    emit("remove", props.href)
  } else {
    emit("edit", value)
  }
  cancel()
}

function close(): void {
  cancel()
  emit("close")
}

defineExpose({ startEdit, save, cancel, editing })
</script>

<style scoped>
.docx-link-card {
  display: flex;
  flex-direction: column;
  gap: 6px;
  max-width: 22rem;
  padding: 8px 10px;
}
.docx-link-card :deep(.office-card__content) {
  display: flex;
  flex-direction: column;
  gap: 6px;
  padding: 0;
}
.docx-link-card__url {
  color: #1d4ed8;
  font-size: 12px;
  max-width: 20rem;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.docx-link-card__actions {
  display: flex;
  gap: 6px;
}
.docx-link-card__action {
  min-block-size: var(--office-ui-height-xs, 28px);
}
.docx-link-card__action--primary {
  background: var(--office-primary, #2563eb);
  border-color: var(--office-primary, #2563eb);
  color: var(--office-primary-foreground, #ffffff);
}
.docx-link-card__action--danger {
  color: var(--office-destructive-foreground, #b91c1c);
}
.docx-link-card__field {
  display: flex;
  flex-direction: column;
  gap: 4px;
}
.docx-link-card__label {
  color: #6b7280;
  font-size: 11px;
}
.docx-link-card__input {
  border: 1px solid #d1d5db;
  border-radius: 4px;
  font: inherit;
  font-size: 12px;
  padding: 3px 6px;
  width: 100%;
}
.docx-link-card--dark {
  background: var(--office-card, #18181b);
  border-color: var(--office-border, #3f3f46);
  color: var(--office-card-foreground, #f4f4f5);
}
.docx-link-card--dark .docx-link-card__url {
  color: #93c5fd;
}
.docx-link-card--dark .docx-link-card__action {
  background: #27272a;
  border-color: #3f3f46;
  color: #e4e4e7;
}
.docx-link-card--dark .docx-link-card__label {
  color: #a1a1aa;
}
.docx-link-card--dark .docx-link-card__input {
  background: #09090b;
  border-color: #3f3f46;
  color: #f4f4f5;
}
</style>
