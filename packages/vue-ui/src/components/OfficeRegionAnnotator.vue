<template>
  <div
    :class="['office-region-annotator', props.className, { 'is-disabled': props.disabled, 'is-read-only': props.readOnly }]"
    role="group"
    :aria-label="props.ariaLabel ?? m.vueUi.regionAnnotator.defaultAriaLabel"
    data-testid="office-region-annotator"
    @keydown="onKeydown"
  >
    <div
      class="office-region-annotator__toolbar"
      role="toolbar"
      :aria-label="m.vueUi.regionAnnotator.toolbarLabel"
      data-testid="office-region-annotator-toolbar"
    >
      <button
        v-for="tool in tools"
        :key="tool.kind"
        type="button"
        class="office-region-annotator__tool"
        :class="{ 'is-active': activeMode === tool.kind }"
        :aria-pressed="activeMode === tool.kind"
        :aria-label="tool.label"
        :title="tool.label"
        :disabled="props.disabled || props.readOnly"
        :data-testid="`office-region-annotator-tool-${tool.kind}`"
        @click="setMode(tool.kind)"
      >
        <span aria-hidden="true">{{ tool.glyph }}</span>
        <span>{{ tool.label }}</span>
      </button>
      <span class="office-region-annotator__separator" />
      <label class="office-region-annotator__color">
        <span>{{ m.vueUi.regionAnnotator.color }}</span>
        <input
          v-model="colorValue"
          type="color"
          :aria-label="m.vueUi.regionAnnotator.color"
          :disabled="props.disabled || props.readOnly"
          data-testid="office-region-annotator-color"
        />
      </label>
      <button
        type="button"
        class="office-region-annotator__action"
        :disabled="props.disabled || props.readOnly || !selectedId"
        :aria-label="m.vueUi.regionAnnotator.remove"
        data-testid="office-region-annotator-remove"
        @click="removeSelected"
      >
        {{ m.vueUi.regionAnnotator.remove }}
      </button>
      <button
        type="button"
        class="office-region-annotator__action"
        :disabled="props.disabled || props.readOnly || annotations.length === 0"
        :aria-label="m.vueUi.regionAnnotator.clear"
        data-testid="office-region-annotator-clear"
        @click="clearAnnotations"
      >
        {{ m.vueUi.regionAnnotator.clear }}
      </button>
    </div>

    <div
      ref="stageRef"
      class="office-region-annotator__stage"
      :class="{ 'is-drawing': dragging }"
      :style="stageStyle"
      data-testid="office-region-annotator-stage"
      @pointerdown="onPointerDown"
      @pointermove="onPointerMove"
      @pointerup="onPointerUp"
      @pointercancel="onPointerCancel"
    >
      <img
        v-if="imageSrc"
        class="office-region-annotator__image"
        :src="imageSrc"
        :alt="props.alt ?? m.vueUi.regionAnnotator.imageAlt"
        draggable="false"
        data-testid="office-region-annotator-image"
        @load="onImageLoad"
      />
      <div v-else class="office-region-annotator__empty" data-testid="office-region-annotator-empty">
        {{ m.vueUi.regionAnnotator.empty }}
      </div>
      <svg
        v-if="imageSrc"
        class="office-region-annotator__overlay"
        :viewBox="`0 0 ${imageWidth} ${imageHeight}`"
        preserveAspectRatio="none"
        aria-hidden="true"
        data-testid="office-region-annotator-overlay"
      >
        <g
          v-for="annotation in annotations"
          :key="annotation.id"
          :class="[
            'office-region-annotator__annotation',
            `is-${annotation.kind}`,
            { 'is-selected': annotation.id === selectedId, 'is-interactive': !props.disabled && !props.readOnly },
          ]"
          :data-testid="`office-region-annotation-${annotation.id}`"
          :data-kind="annotation.kind"
          @pointerdown.stop="selectAnnotation(annotation.id)"
        >
          <template v-if="annotation.kind === 'rectangle'">
            <rect
              :x="annotation.rect.x * imageWidth"
              :y="annotation.rect.y * imageHeight"
              :width="annotation.rect.width * imageWidth"
              :height="annotation.rect.height * imageHeight"
              fill="none"
              :stroke="annotationColor(annotation)"
              :stroke-width="annotationStroke(annotation)"
              :opacity="annotationOpacity(annotation)"
            />
          </template>
          <template v-else-if="annotation.kind === 'highlight'">
            <rect
              :x="annotation.rect.x * imageWidth"
              :y="annotation.rect.y * imageHeight"
              :width="annotation.rect.width * imageWidth"
              :height="annotation.rect.height * imageHeight"
              :fill="annotationColor(annotation)"
              :opacity="annotation.style?.opacity ?? 0.35"
            />
          </template>
          <template v-else-if="annotation.kind === 'arrow'">
            <line
              :x1="annotation.start.x * imageWidth"
              :y1="annotation.start.y * imageHeight"
              :x2="annotation.end.x * imageWidth"
              :y2="annotation.end.y * imageHeight"
              :stroke="annotationColor(annotation)"
              :stroke-width="annotationStroke(annotation)"
              stroke-linecap="round"
              :opacity="annotationOpacity(annotation)"
            />
            <polygon
              :points="arrowHeadPoints(annotation)"
              :fill="annotationColor(annotation)"
              :opacity="annotationOpacity(annotation)"
            />
          </template>
          <template v-else-if="annotation.kind === 'text'">
            <text
              :x="annotation.point.x * imageWidth"
              :y="annotation.point.y * imageHeight"
              :fill="annotationColor(annotation)"
              :font-size="annotationFontSize(annotation)"
              :opacity="annotationOpacity(annotation)"
              dominant-baseline="hanging"
              class="office-region-annotator__text"
            >
              <tspan
                v-for="(line, index) in annotation.text.split('\n')"
                :key="index"
                :x="annotation.point.x * imageWidth"
                :dy="index === 0 ? 0 : annotationFontSize(annotation) * 1.25"
              >{{ line }}</tspan>
            </text>
          </template>
        </g>
        <g
          v-if="selectionBounds"
          class="office-region-annotator__selection"
          :data-testid="'office-region-annotator-selection'"
        >
          <rect
            :x="selectionBounds.x * imageWidth"
            :y="selectionBounds.y * imageHeight"
            :width="Math.max(1, selectionBounds.width * imageWidth)"
            :height="Math.max(1, selectionBounds.height * imageHeight)"
            fill="none"
            stroke="#2563eb"
            :stroke-width="Math.max(1, imageMinSide * 0.004)"
            stroke-dasharray="6 4"
          />
        </g>
        <g v-if="draftRect" class="office-region-annotator__draft">
          <rect
            :x="draftRect.x * imageWidth"
            :y="draftRect.y * imageHeight"
            :width="draftRect.width * imageWidth"
            :height="draftRect.height * imageHeight"
            :fill="draftKind === 'highlight' ? colorValue : 'none'"
            :fill-opacity="draftKind === 'highlight' ? 0.35 : 0"
            :stroke="colorValue"
            :stroke-width="draftStroke"
            :stroke-dasharray="draftKind === 'highlight' ? undefined : '5 4'"
          />
        </g>
        <g v-if="draftArrow" class="office-region-annotator__draft">
          <line
            :x1="draftArrow.start.x * imageWidth"
            :y1="draftArrow.start.y * imageHeight"
            :x2="draftArrow.end.x * imageWidth"
            :y2="draftArrow.end.y * imageHeight"
            :stroke="colorValue"
            :stroke-width="draftStroke"
            stroke-linecap="round"
            stroke-dasharray="5 4"
          />
        </g>
      </svg>
      <div
        v-if="textDraft"
        class="office-region-annotator__text-editor"
        :style="textEditorStyle"
        data-testid="office-region-annotator-text-editor"
      >
        <input
          ref="textInputRef"
          v-model="textValue"
          type="text"
          :placeholder="m.vueUi.regionAnnotator.textPlaceholder"
          :aria-label="m.vueUi.regionAnnotator.textPlaceholder"
          data-testid="office-region-annotator-text-input"
          @keydown.enter.prevent="commitText"
          @keydown.esc.prevent="cancelText"
          @blur="commitText"
        />
        <button type="button" data-testid="office-region-annotator-text-save" @mousedown.prevent @click="commitText">
          {{ m.vueUi.regionAnnotator.textSave }}
        </button>
        <button type="button" data-testid="office-region-annotator-text-cancel" @mousedown.prevent @click="cancelText">
          {{ m.vueUi.regionAnnotator.textCancel }}
        </button>
      </div>
    </div>

    <div class="office-region-annotator__footer">
      <span data-testid="office-region-annotator-count">
        {{ formatOfficeMessage(m.vueUi.regionAnnotator.count, { count: annotations.length }) }}
      </span>
      <span class="office-region-annotator__hint">{{ m.vueUi.regionAnnotator.hint }}</span>
    </div>
    <span class="office-region-annotator__live" role="status" aria-live="polite">
      {{ liveMessage }}
    </span>
  </div>
</template>

<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, ref, watch } from "vue"
import {
  formatOfficeMessage,
  normalizedRectFromPoints,
  resolveOfficeMessages,
  type NormalizedPoint,
  type NormalizedRect,
} from "@arcships/office-interaction"
import type { OfficeRegionAnnotatorProps } from "../types"
import { useOfficeUiMessages } from "../i18n"
import {
  DEFAULT_REGION_ANNOTATION_MAX_COUNT,
  createOfficeRegionAnnotationDocument,
  createOfficeRegionAnnotationId,
  officeRegionAnnotationBounds,
  parseOfficeRegionAnnotationDocument,
  stringifyOfficeRegionAnnotationDocument,
  type OfficeRegionAnnotation,
  type OfficeRegionAnnotationDocument,
  type OfficeRegionAnnotationInput,
  type OfficeRegionAnnotationKind,
  type OfficeRegionAnnotationStyle,
  type OfficeRegionArrowAnnotation,
} from "../region-annotations"

const props = withDefaults(defineProps<OfficeRegionAnnotatorProps>(), {
  modelValue: null,
  image: null,
  width: undefined,
  height: undefined,
  mode: "rectangle",
  color: "#ef4444",
  strokeWidth: 0.006,
  fontSize: 0.05,
  disabled: false,
  readOnly: false,
  maxAnnotations: DEFAULT_REGION_ANNOTATION_MAX_COUNT,
  alt: undefined,
  ariaLabel: undefined,
  className: "",
  locale: undefined,
  messages: undefined,
})

const emit = defineEmits<{
  "update:modelValue": [document: OfficeRegionAnnotationDocument]
  "update:mode": [mode: OfficeRegionAnnotationKind]
  "update:color": [color: string]
  annotationAdd: [annotation: OfficeRegionAnnotation]
  annotationRemove: [annotation: OfficeRegionAnnotation]
  annotationSelect: [annotation: OfficeRegionAnnotation | null]
  annotationChange: [document: OfficeRegionAnnotationDocument]
  annotationError: [error: unknown]
  imageLoad: [size: { width: number; height: number }]
}>()

const inheritedMessages = useOfficeUiMessages()
const m = computed(() => (props.locale || props.messages)
  ? resolveOfficeMessages(props.locale, props.messages)
  : inheritedMessages.value)

const tools = computed(() => [
  { kind: "arrow" as const, glyph: "↗", label: m.value.vueUi.regionAnnotator.arrow },
  { kind: "rectangle" as const, glyph: "▭", label: m.value.vueUi.regionAnnotator.rectangle },
  { kind: "text" as const, glyph: "T", label: m.value.vueUi.regionAnnotator.text },
  { kind: "highlight" as const, glyph: "▮", label: m.value.vueUi.regionAnnotator.highlight },
])

const stageRef = ref<HTMLElement | null>(null)
const textInputRef = ref<HTMLInputElement | null>(null)
const activeMode = ref<OfficeRegionAnnotationKind>(props.mode)
const colorValue = ref(props.color)
const annotations = ref<OfficeRegionAnnotation[]>([])
const source = ref<OfficeRegionAnnotationDocument["source"]>()
const selectedId = ref<string | null>(null)
const draftStart = ref<NormalizedPoint | null>(null)
const draftKind = ref<"rectangle" | "arrow" | "highlight">("rectangle")
const draftPointerId = ref<number | null>(null)
const draftRect = ref<NormalizedRect | null>(null)
const draftArrow = ref<{ start: NormalizedPoint; end: NormalizedPoint } | null>(null)
const textDraft = ref<NormalizedPoint | null>(null)
const textValue = ref("")
const liveMessage = ref("")
const decodedWidth = ref(0)
const decodedHeight = ref(0)
const blobUrl = ref<string | null>(null)

const dragging = computed(() => draftPointerId.value !== null)
const imageWidth = computed(() => {
  if (Number.isFinite(props.width) && (props.width ?? 0) > 0) return props.width as number
  if (decodedWidth.value > 0) return decodedWidth.value
  return 1000
})
const imageHeight = computed(() => {
  if (Number.isFinite(props.height) && (props.height ?? 0) > 0) return props.height as number
  if (decodedHeight.value > 0) return decodedHeight.value
  return 1000
})
const imageMinSide = computed(() => Math.min(imageWidth.value, imageHeight.value))
const stageStyle = computed(() => ({
  aspectRatio: `${imageWidth.value} / ${imageHeight.value}`,
}))
const draftStroke = computed(() => Math.max(1, imageMinSide.value * props.strokeWidth))
const selectedAnnotation = computed(() =>
  annotations.value.find((annotation) => annotation.id === selectedId.value) ?? null,
)
const selectionBounds = computed(() =>
  selectedAnnotation.value ? officeRegionAnnotationBounds(selectedAnnotation.value) : null,
)
const textEditorStyle = computed(() => textDraft.value ? {
  left: `${textDraft.value.x * 100}%`,
  top: `${textDraft.value.y * 100}%`,
} : undefined)
const imageSrc = computed(() => {
  if (typeof props.image === "string") return props.image
  return blobUrl.value
})

function resetImageObjectUrl(): void {
  if (blobUrl.value && typeof URL !== "undefined" && typeof URL.revokeObjectURL === "function") URL.revokeObjectURL(blobUrl.value)
  blobUrl.value = null
}

watch(() => props.image, (image) => {
  resetImageObjectUrl()
  if (image && typeof image === "object" && typeof Blob !== "undefined" && image instanceof Blob) {
    if (typeof URL !== "undefined" && typeof URL.createObjectURL === "function") blobUrl.value = URL.createObjectURL(image)
  }
}, { immediate: true })

watch(() => props.mode, (mode) => {
  activeMode.value = mode
})

watch(() => props.color, (color) => {
  colorValue.value = color
})

function applyDocument(document: OfficeRegionAnnotationDocument | null | undefined): void {
  try {
    const parsed = document
      ? parseOfficeRegionAnnotationDocument(document, { maxAnnotations: props.maxAnnotations })
      : createOfficeRegionAnnotationDocument()
    annotations.value = parsed.annotations
    source.value = parsed.source
    selectedId.value = parsed.annotations.some((annotation) => annotation.id === selectedId.value)
      ? selectedId.value
      : null
  } catch (reason) {
    emit("annotationError", reason)
    annotations.value = []
    source.value = undefined
    selectedId.value = null
  }
}

watch(() => props.modelValue, (document) => applyDocument(document), { immediate: true, deep: true })

function currentDocument(): OfficeRegionAnnotationDocument {
  return createOfficeRegionAnnotationDocument({
    annotations: annotations.value,
    ...(source.value ? { source: source.value } : {}),
  })
}

function announce(message: string): void {
  liveMessage.value = message
}

function emitDocument(next: OfficeRegionAnnotation[], extra?: (document: OfficeRegionAnnotationDocument) => void): boolean {
  if (props.disabled || props.readOnly) return false
  try {
    const document = createOfficeRegionAnnotationDocument({
      annotations: next,
      ...(source.value ? { source: source.value } : {}),
    })
    annotations.value = document.annotations
    emit("update:modelValue", document)
    emit("annotationChange", document)
    extra?.(document)
    return true
  } catch (reason) {
    emit("annotationError", reason)
    return false
  }
}

function styleSnapshot(kind: OfficeRegionAnnotationKind): OfficeRegionAnnotationStyle {
  const style: OfficeRegionAnnotationStyle = {
    color: colorValue.value,
    strokeWidth: props.strokeWidth,
  }
  if (kind === "text") style.fontSize = props.fontSize
  return style
}

/** Add an annotation. Invalid input throws a validation error. */
function addAnnotation(input: OfficeRegionAnnotationInput & { id?: string }): OfficeRegionAnnotation | null {
  if (annotations.value.length >= props.maxAnnotations) {
    const error = new RangeError(`annotations must not exceed ${props.maxAnnotations} entries`)
    emit("annotationError", error)
    return null
  }
  const annotation = {
    ...input,
    id: input.id?.trim() || createOfficeRegionAnnotationId(),
    style: input.style ?? styleSnapshot(input.kind),
  } as OfficeRegionAnnotation
  if (!emitDocument([...annotations.value, annotation], () => {
    emit("annotationAdd", annotation)
    announce(formatOfficeMessage(m.value.vueUi.regionAnnotator.count, { count: annotations.value.length }))
  })) return null
  return annotation
}

function removeAnnotation(id: string): boolean {
  const annotation = annotations.value.find((candidate) => candidate.id === id)
  if (!annotation) return false
  return emitDocument(annotations.value.filter((candidate) => candidate.id !== id), () => {
    if (selectedId.value === id) {
      selectedId.value = null
      emit("annotationSelect", null)
    }
    emit("annotationRemove", annotation)
  })
}

function removeSelected(): void {
  if (selectedId.value) removeAnnotation(selectedId.value)
}

function clearAnnotations(): void {
  selectedId.value = null
  emit("annotationSelect", null)
  emitDocument([])
}

function selectAnnotation(id: string | null): void {
  if (props.disabled) return
  selectedId.value = id
  emit("annotationSelect", annotations.value.find((annotation) => annotation.id === id) ?? null)
}

function setMode(mode: OfficeRegionAnnotationKind): void {
  if (props.disabled || props.readOnly) return
  activeMode.value = mode
  if (mode !== "text") cancelText()
  emit("update:mode", mode)
}

function pointFromEvent(event: PointerEvent): NormalizedPoint | null {
  const stage = stageRef.value
  if (!stage) return null
  const bounds = stage.getBoundingClientRect()
  if (!(bounds.width > 0) || !(bounds.height > 0)) return null
  return {
    x: Math.min(1, Math.max(0, (event.clientX - bounds.left) / bounds.width)),
    y: Math.min(1, Math.max(0, (event.clientY - bounds.top) / bounds.height)),
  }
}

function updateDraft(point: NormalizedPoint): void {
  const start = draftStart.value
  if (!start) return
  if (draftKind.value === "arrow") {
    draftArrow.value = { start, end: point }
    draftRect.value = null
    return
  }
  try {
    draftRect.value = normalizedRectFromPoints(start, point)
  } catch {
    draftRect.value = null
  }
  draftArrow.value = null
}

function onPointerDown(event: PointerEvent): void {
  if (props.disabled || props.readOnly || event.button !== 0 || event.isPrimary === false || !imageSrc.value) return
  const point = pointFromEvent(event)
  if (!point) return
  if (activeMode.value === "text") {
    textDraft.value = point
    textValue.value = ""
    selectedId.value = null
    void nextTick(() => textInputRef.value?.focus())
    event.preventDefault()
    return
  }
  draftKind.value = activeMode.value
  draftStart.value = point
  draftRect.value = null
  draftArrow.value = null
  draftPointerId.value = event.pointerId
  ;(event.currentTarget as HTMLElement).setPointerCapture?.(event.pointerId)
  event.preventDefault()
}

function onPointerMove(event: PointerEvent): void {
  if (draftPointerId.value === null || event.pointerId !== draftPointerId.value) return
  const point = pointFromEvent(event)
  if (point) updateDraft(point)
}

function releasePointer(event: PointerEvent): void {
  const target = event.currentTarget as HTMLElement
  if (target.hasPointerCapture?.(event.pointerId)) target.releasePointerCapture?.(event.pointerId)
}

function commitDraft(): void {
  if (draftKind.value === "arrow" && draftArrow.value) {
    const { start, end } = draftArrow.value
    if (Math.hypot(end.x - start.x, end.y - start.y) >= 0.005) {
      addAnnotation({ kind: "arrow", start, end, style: styleSnapshot("arrow") })
    }
  } else if (draftRect.value && draftRect.value.width >= 0.005 && draftRect.value.height >= 0.005) {
    const rect = { ...draftRect.value }
    if (draftKind.value === "highlight") {
      addAnnotation({ kind: "highlight", rect, style: styleSnapshot("highlight") })
    } else {
      addAnnotation({ kind: "rectangle", rect, style: styleSnapshot("rectangle") })
    }
  }
}

function clearDraft(): void {
  draftStart.value = null
  draftRect.value = null
  draftArrow.value = null
  draftPointerId.value = null
}

function onPointerUp(event: PointerEvent): void {
  if (draftPointerId.value === null || event.pointerId !== draftPointerId.value) return
  releasePointer(event)
  const point = pointFromEvent(event)
  if (point) updateDraft(point)
  commitDraft()
  clearDraft()
}

function onPointerCancel(event: PointerEvent): void {
  if (draftPointerId.value === null || event.pointerId !== draftPointerId.value) return
  releasePointer(event)
  clearDraft()
}

function commitText(): void {
  const point = textDraft.value
  const text = textValue.value.trim()
  if (!point) return
  textDraft.value = null
  textValue.value = ""
  if (text) addAnnotation({ kind: "text", point, text, style: styleSnapshot("text") })
}

function cancelText(): void {
  textDraft.value = null
  textValue.value = ""
}

function onKeydown(event: KeyboardEvent): void {
  if (props.disabled) return
  const target = event.target as HTMLElement | null
  if (target?.tagName === "INPUT" || target?.tagName === "TEXTAREA") return
  if (event.key === "Escape") {
    event.preventDefault()
    if (textDraft.value) cancelText()
    else if (dragging.value) clearDraft()
    else selectAnnotation(null)
  } else if ((event.key === "Delete" || event.key === "Backspace") && selectedId.value) {
    event.preventDefault()
    removeSelected()
  }
}

function annotationColor(annotation: OfficeRegionAnnotation): string {
  return annotation.style?.color ?? colorValue.value
}

function annotationOpacity(annotation: OfficeRegionAnnotation): number {
  return annotation.style?.opacity ?? 1
}

function annotationStroke(annotation: OfficeRegionAnnotation): number {
  return Math.max(1, imageMinSide.value * (annotation.style?.strokeWidth ?? props.strokeWidth))
}

function annotationFontSize(annotation: OfficeRegionAnnotation): number {
  return Math.max(1, imageMinSide.value * (annotation.style?.fontSize ?? props.fontSize))
}

function arrowHeadPoints(annotation: OfficeRegionArrowAnnotation): string {
  const x1 = annotation.start.x * imageWidth.value
  const y1 = annotation.start.y * imageHeight.value
  const x2 = annotation.end.x * imageWidth.value
  const y2 = annotation.end.y * imageHeight.value
  const angle = Math.atan2(y2 - y1, x2 - x1)
  const head = Math.max(6, annotationStroke(annotation) * 3)
  const spread = Math.PI / 7
  const left = `${x2 - head * Math.cos(angle - spread)},${y2 - head * Math.sin(angle - spread)}`
  const right = `${x2 - head * Math.cos(angle + spread)},${y2 - head * Math.sin(angle + spread)}`
  return `${x2},${y2} ${left} ${right}`
}

function onImageLoad(event: Event): void {
  const image = event.target as HTMLImageElement
  if (image.naturalWidth > 0 && image.naturalHeight > 0) {
    decodedWidth.value = image.naturalWidth
    decodedHeight.value = image.naturalHeight
    emit("imageLoad", { width: image.naturalWidth, height: image.naturalHeight })
  }
}

function toJSON(): string {
  return stringifyOfficeRegionAnnotationDocument(currentDocument())
}

function loadJSON(value: string | OfficeRegionAnnotationDocument): void {
  try {
    const parsed = parseOfficeRegionAnnotationDocument(value, { maxAnnotations: props.maxAnnotations })
    applyDocument(parsed)
    emit("update:modelValue", parsed)
    emit("annotationChange", parsed)
  } catch (reason) {
    emit("annotationError", reason)
    throw reason
  }
}

defineExpose({
  /** Adds one annotation and emits the updated document. */
  addAnnotation,
  removeAnnotation,
  removeSelected,
  clearAnnotations,
  getDocument: currentDocument,
  toJSON,
  loadJSON,
  selectAnnotation,
  getSelectedId: (): string | null => selectedId.value,
  getImageSize: () => ({ width: imageWidth.value, height: imageHeight.value }),
})

onBeforeUnmount(() => {
  resetImageObjectUrl()
})
</script>

<style scoped>
.office-region-annotator {
  background: var(--office-surface-bg, #ffffff);
  border: 1px solid var(--office-border, #d4d4d8);
  border-radius: 8px;
  color: var(--office-foreground, #18181b);
  display: flex;
  flex-direction: column;
  gap: 10px;
  min-width: 0;
  padding: 10px;
}

.office-region-annotator__toolbar {
  align-items: center;
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
}

.office-region-annotator__tool,
.office-region-annotator__action {
  align-items: center;
  appearance: none;
  background: var(--office-surface-bg, #ffffff);
  border: 1px solid var(--office-border, #d4d4d8);
  border-radius: 5px;
  color: inherit;
  cursor: pointer;
  display: inline-flex;
  font: inherit;
  font-size: 12px;
  gap: 4px;
  padding: 4px 8px;
}

.office-region-annotator__tool[aria-pressed="true"] {
  background: #2563eb;
  border-color: #2563eb;
  color: #ffffff;
}

.office-region-annotator__tool:disabled,
.office-region-annotator__action:disabled {
  cursor: default;
  opacity: 0.45;
}

.office-region-annotator__separator {
  background: var(--office-border, #d4d4d8);
  height: 20px;
  width: 1px;
}

.office-region-annotator__color {
  align-items: center;
  display: inline-flex;
  font-size: 12px;
  gap: 4px;
}

.office-region-annotator__color input {
  background: none;
  border: 0;
  cursor: pointer;
  height: 24px;
  padding: 0;
  width: 30px;
}

.office-region-annotator__stage {
  background: repeating-conic-gradient(#f4f4f5 0 25%, #ffffff 0 50%) 50% / 16px 16px;
  cursor: crosshair;
  overflow: hidden;
  position: relative;
  touch-action: none;
  width: 100%;
}

.is-read-only .office-region-annotator__stage,
.is-disabled .office-region-annotator__stage {
  cursor: default;
  touch-action: auto;
}

.office-region-annotator__image {
  display: block;
  height: 100%;
  object-fit: fill;
  pointer-events: none;
  user-select: none;
  width: 100%;
}

.office-region-annotator__empty {
  align-items: center;
  color: var(--office-muted-foreground, #71717a);
  display: flex;
  font-size: 13px;
  inset: 0;
  justify-content: center;
  padding: 24px;
  position: absolute;
  text-align: center;
}

.office-region-annotator__overlay {
  height: 100%;
  inset: 0;
  pointer-events: none;
  position: absolute;
  width: 100%;
}

.office-region-annotator__annotation.is-interactive {
  cursor: pointer;
  pointer-events: auto;
}

.office-region-annotator__annotation.is-selected {
  filter: drop-shadow(0 0 1px rgb(37 99 235 / 0.9));
}

.office-region-annotator__text {
  font-family: ui-sans-serif, system-ui, sans-serif;
  font-weight: 600;
  paint-order: stroke;
  stroke: rgb(255 255 255 / 0.85);
  stroke-width: 2px;
}

.office-region-annotator__text-editor {
  background: var(--office-surface-bg, #ffffff);
  border: 1px solid #2563eb;
  border-radius: 5px;
  box-shadow: 0 4px 14px rgb(0 0 0 / 0.18);
  display: flex;
  gap: 4px;
  max-width: min(340px, 92vw);
  padding: 4px;
  position: absolute;
  transform: translate(-2px, -2px);
  z-index: 2;
}

.office-region-annotator__text-editor input {
  border: 1px solid var(--office-border, #d4d4d8);
  border-radius: 4px;
  font: inherit;
  font-size: 12px;
  min-width: 120px;
  padding: 3px 6px;
}

.office-region-annotator__text-editor button {
  background: var(--office-surface-bg, #ffffff);
  border: 1px solid var(--office-border, #d4d4d8);
  border-radius: 4px;
  cursor: pointer;
  font: inherit;
  font-size: 12px;
  padding: 3px 6px;
}

.office-region-annotator__footer {
  color: var(--office-muted-foreground, #71717a);
  display: flex;
  flex-wrap: wrap;
  font-size: 12px;
  gap: 6px 16px;
}

.office-region-annotator__hint {
  flex: 1 1 280px;
}

.office-region-annotator__live {
  clip: rect(0 0 0 0);
  clip-path: inset(50%);
  height: 1px;
  overflow: hidden;
  position: absolute;
  white-space: nowrap;
  width: 1px;
}
</style>
