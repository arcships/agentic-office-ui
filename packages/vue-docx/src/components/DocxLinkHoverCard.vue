<template>
  <div
    class="docx-link-card"
    :class="{ 'docx-link-card--dark': documentTheme === 'dark' }"
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
          class="docx-link-card__input"
          :placeholder="m.docx.linkCard.urlPlaceholder"
          :aria-label="m.docx.linkCard.url"
          data-testid="docx-link-card-input"
          @keydown.enter.stop.prevent="save"
        />
      </label>
      <div class="docx-link-card__actions">
        <button
          type="button"
          class="docx-link-card__action docx-link-card__action--primary"
          data-testid="docx-link-card-save"
          @click="save"
        >
          {{ m.docx.linkCard.save }}
        </button>
        <button
          type="button"
          class="docx-link-card__action"
          data-testid="docx-link-card-cancel"
          @click="cancel"
        >
          {{ m.docx.linkCard.cancel }}
        </button>
      </div>
    </template>
    <template v-else>
      <span class="docx-link-card__url" :title="href" data-testid="docx-link-card-url">
        {{ displayHref }}
      </span>
      <div class="docx-link-card__actions">
        <button
          type="button"
          class="docx-link-card__action"
          data-testid="docx-link-card-open"
          :title="m.docx.linkCard.open"
          @click="emit('open', href)"
        >
          {{ m.docx.linkCard.open }}
        </button>
        <button
          v-if="editable"
          type="button"
          class="docx-link-card__action"
          data-testid="docx-link-card-edit"
          :title="m.docx.linkCard.edit"
          @click="startEdit"
        >
          {{ m.docx.linkCard.edit }}
        </button>
        <button
          v-if="editable"
          type="button"
          class="docx-link-card__action docx-link-card__action--danger"
          data-testid="docx-link-card-remove"
          :title="m.docx.linkCard.remove"
          @click="emit('remove', href)"
        >
          {{ m.docx.linkCard.remove }}
        </button>
      </div>
    </template>
  </div>
</template>

<script setup lang="ts">
import { computed, nextTick, ref, watch } from "vue"
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
  background: #ffffff;
  border: 1px solid #d1d5db;
  border-radius: 8px;
  box-shadow: 0 10px 30px rgba(15, 23, 42, 0.18);
  color: #111827;
  display: flex;
  flex-direction: column;
  gap: 6px;
  max-width: 22rem;
  padding: 8px 10px;
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
  background: #f3f4f6;
  border: 1px solid #d1d5db;
  border-radius: 4px;
  color: #374151;
  cursor: pointer;
  font-size: 12px;
  padding: 3px 8px;
}
.docx-link-card__action:hover {
  background: #e5e7eb;
}
.docx-link-card__action--primary {
  background: #2563eb;
  border-color: #2563eb;
  color: #ffffff;
}
.docx-link-card__action--danger {
  color: #b91c1c;
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
  background: #18181b;
  border-color: #3f3f46;
  color: #f4f4f5;
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
