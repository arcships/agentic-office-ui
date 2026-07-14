<template>
  <div
    :class="['office-object-outline-layer', props.className, { 'is-interactive': props.interactive }]"
    :role="props.interactive ? 'listbox' : undefined"
    :aria-label="props.interactive ? props.ariaLabel : undefined"
    data-testid="office-object-outline-layer"
    @keydown="onKeydown"
  >
    <component
      :is="props.interactive ? 'button' : 'span'"
      v-for="item in props.items"
      :id="optionId(item.id)"
      :key="item.id"
      :ref="(element: unknown) => setOutlineElement(item.id, element)"
      :type="props.interactive ? 'button' : undefined"
      :role="props.interactive ? 'option' : undefined"
      :tabindex="props.interactive ? (item.id === effectiveActiveId ? 0 : -1) : undefined"
      :aria-selected="props.interactive ? item.id === effectiveActiveId : undefined"
      :aria-disabled="item.disabled || undefined"
      :aria-label="outlineAriaLabel(item)"
      :class="[
        'office-object-outline',
        `is-${item.state ?? 'available'}`,
        `reliability-${item.reliability ?? 'unknown'}`,
        { 'is-active': item.id === effectiveActiveId, 'is-disabled': item.disabled },
      ]"
      :style="outlineStyle(item)"
      :data-outline-id="item.id"
      :data-reference-id="item.referenceId"
      :data-kind="item.kind"
      :data-state="item.state ?? 'available'"
      @focus="activate(item)"
      @pointerenter="activate(item)"
      @click.stop="confirm(item, $event)"
    >
      <span class="office-object-outline__label">{{ item.label }}</span>
    </component>
    <span class="office-object-outline-layer__announcement" role="status" aria-live="polite" aria-atomic="true">
      {{ activeAnnouncement }}
    </span>
  </div>
</template>

<script setup lang="ts">
import { computed, getCurrentInstance, nextTick } from "vue"
import type {
  OfficeObjectOutline,
  OfficeObjectOutlineLayerProps,
  OfficeOutlineConfirmOptions,
} from "../types"

const props = withDefaults(defineProps<OfficeObjectOutlineLayerProps>(), {
  activeId: undefined,
  interactive: true,
  ariaLabel: "可选择的文档对象",
  className: "",
})

const emit = defineEmits<{
  activate: [item: OfficeObjectOutline]
  confirm: [item: OfficeObjectOutline, options: OfficeOutlineConfirmOptions]
  dismiss: []
  "navigate-hierarchy": [direction: "parent" | "child"]
}>()

const outlineElements = new Map<string, HTMLElement>()
const optionIdPrefix = `office-object-outline-${getCurrentInstance()?.uid ?? 0}`
const enabledItems = computed(() => props.items.filter((item) => !item.disabled))
const effectiveActiveId = computed(() => {
  if (props.activeId && enabledItems.value.some((item) => item.id === props.activeId)) return props.activeId
  return enabledItems.value[0]?.id
})
const activeItem = computed(() => props.items.find((item) => item.id === effectiveActiveId.value))
const activeAnnouncement = computed(() => activeItem.value
  ? `${activeItem.value.label}，${activeItem.value.kind}`
  : props.items.length ? "没有可用对象" : "没有候选对象",
)

function optionId(id: string): string {
  return `${optionIdPrefix}-${encodeURIComponent(id)}`
}

function setOutlineElement(id: string, element: unknown): void {
  if (element && typeof (element as { focus?: unknown }).focus === "function") {
    outlineElements.set(id, element as HTMLElement)
  }
  else outlineElements.delete(id)
}

function percent(value: number): string {
  return `${Math.min(1, Math.max(0, value)) * 100}%`
}

function outlineStyle(item: OfficeObjectOutline): Record<string, string> {
  return {
    left: percent(item.rect.x),
    top: percent(item.rect.y),
    width: percent(item.rect.width),
    height: percent(item.rect.height),
  }
}

function outlineAriaLabel(item: OfficeObjectOutline): string {
  return `${item.label}，类型 ${item.kind}，识别可靠度 ${item.reliability ?? "unknown"}`
}

function activate(item: OfficeObjectOutline): void {
  if (props.interactive && !item.disabled) emit("activate", item)
}

function confirm(item: OfficeObjectOutline, event: Event): void {
  if (!props.interactive || item.disabled) return
  const pointer = event as MouseEvent
  emit("confirm", item, {
    additiveRequested: pointer.shiftKey === true,
    penetrateRequested: pointer.altKey === true,
  })
}

async function moveActive(direction: 1 | -1): Promise<void> {
  const items = enabledItems.value
  if (!items.length) return
  const current = items.findIndex((item) => item.id === effectiveActiveId.value)
  const index = current < 0
    ? direction === 1 ? 0 : items.length - 1
    : (current + direction + items.length) % items.length
  const item = items[index]
  if (!item) return
  emit("activate", item)
  await nextTick()
  outlineElements.get(item.id)?.focus()
}

function onKeydown(event: KeyboardEvent): void {
  if (!props.interactive) return
  if (event.key === "Tab" || event.key === "ArrowDown" || event.key === "ArrowUp") {
    event.preventDefault()
    const previous = event.key === "ArrowUp" || (event.key === "Tab" && event.shiftKey)
    void moveActive(previous ? -1 : 1)
  } else if (event.key === "Enter" && activeItem.value) {
    event.preventDefault()
    confirm(activeItem.value, event)
  } else if (event.key === "Escape") {
    event.preventDefault()
    emit("dismiss")
  } else if (event.key === "ArrowRight") {
    event.preventDefault()
    emit("navigate-hierarchy", "child")
  } else if (event.key === "ArrowLeft") {
    event.preventDefault()
    emit("navigate-hierarchy", "parent")
  }
}
</script>

<style scoped>
.office-object-outline-layer {
  inset: 0;
  pointer-events: none;
  position: absolute;
  z-index: 20;
}

.office-object-outline {
  appearance: none;
  background: rgb(37 99 235 / 0.06);
  border: 2px solid #2563eb;
  border-radius: 3px;
  box-sizing: border-box;
  color: inherit;
  margin: 0;
  min-height: 2px;
  min-width: 2px;
  padding: 0;
  pointer-events: none;
  position: absolute;
}

.is-interactive .office-object-outline {
  cursor: pointer;
  pointer-events: auto;
}

.office-object-outline.is-active,
.office-object-outline.is-selected {
  background: rgb(37 99 235 / 0.12);
  box-shadow: 0 0 0 2px rgb(255 255 255 / 0.88), 0 0 0 4px rgb(37 99 235 / 0.42);
}

.office-object-outline.is-invalid {
  background: rgb(220 38 38 / 0.08);
  border-color: #dc2626;
  border-style: dashed;
}

.office-object-outline.reliability-uncertain,
.office-object-outline.reliability-unknown {
  border-style: dashed;
}

.office-object-outline.is-disabled {
  cursor: default;
  opacity: .45;
}

.office-object-outline:focus-visible {
  outline: 3px solid #f59e0b;
  outline-offset: 3px;
}

.office-object-outline__label {
  background: #172554;
  border-radius: 3px 3px 0 0;
  bottom: 100%;
  color: #fff;
  font: 600 11px/18px ui-sans-serif, system-ui, sans-serif;
  left: -2px;
  max-width: 220px;
  overflow: hidden;
  padding: 0 6px;
  position: absolute;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.office-object-outline-layer__announcement {
  height: 1px;
  margin: -1px;
  overflow: hidden;
  padding: 0;
  position: absolute;
  width: 1px;
  clip: rect(0 0 0 0);
  clip-path: inset(50%);
  white-space: nowrap;
}
</style>
