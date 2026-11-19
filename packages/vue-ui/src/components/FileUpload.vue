// Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn.
<script setup lang="ts">
import { computed, onBeforeUnmount, ref, type Component } from "vue"
import {
  FileImageIcon,
  FileSpreadsheetIcon,
  FileUploadIcon,
  Upload01Icon,
} from "@hugeicons/core-free-icons"
import FileThumbnail from "./FileThumbnail.vue"
import { HugeIcon, type HugeIconData } from "./HugeIcon"
import { useOfficeUiMessages } from "../i18n"
import type {
  FileUploadAcceptedIcon,
  FileUploadAcceptedType,
  FileUploadProps,
  FileUploadRejection,
} from "../types"

const props = withDefaults(defineProps<FileUploadProps>(), {
  accept: "",
  multiple: false,
  maxSize: Infinity,
  disabled: false,
  className: "",
  title: undefined,
  description: undefined,
  browseLabel: undefined,
  draggingLabel: undefined,
  acceptedFileTypes: undefined,
  showFileList: false,
  showBorderBeam: false,
  borderBeamTheme: "light",
})

const emit = defineEmits<{
  "files-accepted": [files: File[]]
  "files-rejected": [rejections: FileUploadRejection[]]
}>()

const m = useOfficeUiMessages()

const isDragging = ref(false)
const error = ref("")
const inputRef = ref<HTMLInputElement | null>(null)
const dragDepth = ref(0)
const selectedFiles = ref<FileUploadListItem[]>([])

interface FileUploadListItem {
  id: string
  name: string
  type: string
  size: number
  url: string | null
}

interface ResolvedAcceptedType {
  label?: string
  component: Component
  iconData?: HugeIconData
}

const HUGE_ICONS: Record<string, HugeIconData> = {
  image: FileImageIcon,
  pdf: FileUploadIcon,
  sheet: FileSpreadsheetIcon,
}

const DEFAULT_ACCEPTED_TYPES: FileUploadAcceptedType[] = [
  { icon: "image" },
  { icon: "pdf" },
  { icon: "sheet" },
]

const ICON_TRANSFORMS = [
  {
    idle: "translate(-78%, -50%) rotate(-8deg)",
    active: "translate(-114%, -50%) rotate(-12deg) scale(1.08)",
  },
  {
    idle: "translate(-50%, -50%) rotate(0deg)",
    active: "translate(-50%, -50%) rotate(0deg) scale(1.18)",
  },
  {
    idle: "translate(-22%, -50%) rotate(8deg)",
    active: "translate(14%, -50%) rotate(12deg) scale(1.08)",
  },
]

const resolvedAcceptedTypes = computed<ResolvedAcceptedType[]>(() => {
  const types = props.acceptedFileTypes ?? DEFAULT_ACCEPTED_TYPES
  return types.map((entry, index) => resolveAcceptedType(entry, index))
})

const displayTitle = computed(() => props.title ?? m.value.vueUi.fileUpload.prompt)
const displayBrowseLabel = computed(
  () => props.browseLabel ?? m.value.vueUi.fileUpload.browse,
)
const displayDraggingLabel = computed(
  () => props.draggingLabel ?? m.value.vueUi.fileUpload.drop,
)
const showList = computed(() => props.showFileList && selectedFiles.value.length > 0)
const showBeam = computed(() => props.showBorderBeam && isDragging.value && !props.disabled)
const dropZoneClass = computed(() => ({
  "is-disabled": props.disabled,
  "is-dragging": isDragging.value && !props.disabled,
}))

function resolveAcceptedType(
  entry: FileUploadAcceptedType,
  index: number,
): ResolvedAcceptedType {
  const icon = entry.icon
  const fallback = DEFAULT_ACCEPTED_TYPES[index % DEFAULT_ACCEPTED_TYPES.length]?.icon
  const key = typeof icon === "string" ? icon : typeof fallback === "string" ? fallback : "image"
  if (icon && typeof icon !== "string") {
    return { label: entry.label, component: icon as Component }
  }
  return {
    label: entry.label,
    component: HugeIcon,
    iconData: HUGE_ICONS[key] ?? FileImageIcon,
  }
}

function typeStyle(index: number): Record<string, string> {
  const total = resolvedAcceptedTypes.value.length
  const transform =
    total === 1
      ? `translate(-50%, -50%) scale(${isDragging.value && !props.disabled ? 1.14 : 1})`
      : isDragging.value && !props.disabled
        ? ICON_TRANSFORMS[index]?.active ?? ICON_TRANSFORMS[1]!.active
        : ICON_TRANSFORMS[index]?.idle ?? ICON_TRANSFORMS[1]!.idle
  return { transform }
}

function formatBytes(bytes: number): string {
  if (bytes === 0) return "0 B"

  const units = ["B", "KB", "MB", "GB"]
  const index = Math.min(
    Math.floor(Math.log(bytes) / Math.log(1024)),
    units.length - 1,
  )
  return `${(bytes / 1024 ** index).toFixed(index === 0 ? 0 : 1)} ${units[index]}`
}

function matchesAccept(file: File, accept: string): boolean {
  if (!accept) return true

  return accept.split(",").some((rawToken) => {
    const token = rawToken.trim().toLowerCase()
    if (!token) return false
    if (token.startsWith(".")) return file.name.toLowerCase().endsWith(token)
    if (token.endsWith("/*")) {
      return file.type.toLowerCase().startsWith(token.slice(0, -1))
    }
    return file.type.toLowerCase() === token
  })
}

function validateFiles(files: FileList | File[]): {
  accepted: File[]
  rejected: FileUploadRejection[]
} {
  const accepted: File[] = []
  const rejected: FileUploadRejection[] = []

  for (const file of Array.from(files as ArrayLike<File>)) {
    if (!matchesAccept(file, props.accept)) {
      rejected.push({
        code: "FILE_TYPE_NOT_ACCEPTED",
        file,
        message: `"${file.name}" is not an accepted file type`,
      })
      continue
    }
    if (file.size > props.maxSize) {
      const maxMB = (props.maxSize / (1024 * 1024)).toFixed(1)
      rejected.push({
        code: "FILE_TOO_LARGE",
        file,
        message: `"${file.name}" exceeds the ${maxMB}MB size limit`,
      })
      continue
    }
    accepted.push(file)
  }

  return { accepted, rejected }
}

function createPreviewUrl(file: File): string | null {
  if (!file.type.startsWith("image/")) return null
  if (typeof Blob === "undefined" || !(file instanceof Blob)) return null
  if (typeof URL === "undefined" || typeof URL.createObjectURL !== "function") {
    return null
  }
  return URL.createObjectURL(file)
}

function revokePreviewUrls(items: FileUploadListItem[]): void {
  if (typeof URL === "undefined" || typeof URL.revokeObjectURL !== "function") {
    return
  }
  for (const item of items) {
    if (item.url) URL.revokeObjectURL(item.url)
  }
}

function replaceSelectedFiles(files: File[]): void {
  revokePreviewUrls(selectedFiles.value)
  selectedFiles.value = files.map((file) => ({
    id: `${file.name}-${file.size}-${file.lastModified}`,
    name: file.name,
    type: file.type || m.value.vueUi.fileUpload.unknownType,
    size: file.size,
    url: createPreviewUrl(file),
  }))
}

function handleFiles(files: FileList | File[]): void {
  error.value = ""
  const result = validateFiles(files)
  if (result.rejected.length) {
    emit("files-rejected", result.rejected)
  }
  if (result.accepted.length) {
    emit("files-accepted", result.accepted)
    replaceSelectedFiles(result.accepted)
  }

  if (!result.rejected.length) return
  const onlyUnsupportedTypes =
    result.accepted.length === 0 &&
    result.rejected.every((item) => item.code === "FILE_TYPE_NOT_ACCEPTED")
  error.value = onlyUnsupportedTypes
    ? m.value.vueUi.fileUpload.rejected
    : result.rejected.map((item) => item.message).join("; ")
}

function resetDrag(): void {
  dragDepth.value = 0
  isDragging.value = false
}

function onDragEnter(event: DragEvent): void {
  event.preventDefault()
  if (props.disabled) return
  dragDepth.value += 1
  isDragging.value = true
}

function onDragLeave(event: DragEvent): void {
  event.preventDefault()
  if (props.disabled) {
    resetDrag()
    return
  }
  dragDepth.value = Math.max(0, dragDepth.value - 1)
  if (dragDepth.value === 0) isDragging.value = false
}

function onDragOver(event: DragEvent): void {
  event.preventDefault()
  if (event.dataTransfer && !props.disabled) event.dataTransfer.dropEffect = "copy"
}

function onDrop(event: DragEvent): void {
  event.preventDefault()
  resetDrag()
  if (props.disabled) return
  const files = event.dataTransfer?.files
  if (files && files.length) handleFiles(files)
}

function onInputChange(event: Event): void {
  const input = event.target as HTMLInputElement
  if (input.files?.length) handleFiles(input.files)
  input.value = ""
}

function openPicker(): void {
  if (!props.disabled) inputRef.value?.click()
}

onBeforeUnmount(() => revokePreviewUrls(selectedFiles.value))
</script>

<template>
  <div :class="['file-upload', 'office-ui', props.className]" data-slot="file-upload">
    <div
      :class="['file-upload__dropzone', dropZoneClass]"
      role="button"
      :aria-disabled="props.disabled ? 'true' : undefined"
      tabindex="0"
      data-slot="file-upload-dropzone"
      @dragenter="onDragEnter"
      @dragleave="onDragLeave"
      @dragover="onDragOver"
      @drop="onDrop"
      @click="openPicker"
      @keydown.enter.prevent="openPicker"
      @keydown.space.prevent="openPicker"
    >
      <div class="file-upload__icons" aria-hidden="true">
        <div
          v-for="(type, index) in resolvedAcceptedTypes"
          :key="index"
          :class="[
            'file-upload__icon-card',
            { 'is-middle': index === 1, 'is-active': isDragging && !props.disabled },
          ]"
          :style="typeStyle(index)"
        >
          <component
            :is="type.component"
            v-bind="type.iconData ? { icon: type.iconData } : {}"
            class="file-upload__icon"
          />
        </div>
      </div>
      <div class="file-upload__copy">
        <div class="file-upload__title">{{ displayTitle }}</div>
        <div v-if="props.description" class="file-upload__description">
          {{ props.description }}
        </div>
        <div v-if="error" class="file-upload__rejection" role="alert">
          {{ error }}
        </div>
      </div>
      <div class="file-upload__chip">
        <HugeIcon :icon="Upload01Icon" class="file-upload__chip-icon" />
        <span>{{ isDragging && !props.disabled ? displayDraggingLabel : displayBrowseLabel }}</span>
      </div>
      <input
        ref="inputRef"
        type="file"
        :accept="props.accept"
        :multiple="props.multiple"
        :disabled="props.disabled"
        class="file-upload__input"
        tabindex="-1"
        @change="onInputChange"
      />
      <span
        v-if="showBeam"
        :class="['file-upload__beam', `is-${props.borderBeamTheme}`]"
        aria-hidden="true"
      />
    </div>
    <div
      v-if="showList"
      class="file-upload__list"
      data-slot="file-upload-list"
    >
      <div
        v-for="file in selectedFiles"
        :key="file.id"
        class="file-upload__row"
        data-slot="file-upload-row"
      >
        <FileThumbnail
          :file="{ name: file.name, type: file.type }"
          :size="40"
          :preview-image-url="file.url"
          class="file-upload__thumbnail"
        />
        <div class="file-upload__file">
          <p class="file-upload__file-name" :title="file.name">{{ file.name }}</p>
          <p class="file-upload__file-detail">{{ file.type }} - {{ formatBytes(file.size) }}</p>
        </div>
        <div class="file-upload__ready">{{ m.vueUi.fileUpload.ready }}</div>
      </div>
    </div>
  </div>
</template>

<style scoped>
.file-upload {
  --file-upload-border: color-mix(in oklab, var(--office-foreground) 20%, transparent);
  --file-upload-border-hover: color-mix(in oklab, var(--office-foreground) 35%, transparent);
  --file-upload-border-active: color-mix(in oklab, var(--office-foreground) 40%, transparent);
  --file-upload-icon-duration: 500ms;
  --file-upload-icon-ease: cubic-bezier(0.22, 1, 0.36, 1);
  --file-upload-min-height: 16rem;
  --file-upload-card-shadow: 0 1px 2px 0 oklab(0 0 0 / 0.05);
  --file-upload-card-inner-shadow: 0 1px 0 oklab(0 0 0 / 0.04);

  min-width: 0;
}
.file-upload,
.file-upload * {
  box-sizing: border-box;
}
.file-upload > * + * {
  margin-top: var(--office-ui-space-3);
}
.file-upload__dropzone {
  position: relative;
  display: flex;
  min-height: var(--file-upload-min-height);
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: var(--office-ui-space-5);
  overflow: hidden;
  padding: calc(var(--office-ui-space-5) * 2) var(--office-ui-space-6);
  border: var(--office-border-width) dashed var(--file-upload-border);
  border-radius: var(--office-radius-2xl);
  background: var(--office-background);
  color: var(--office-foreground);
  text-align: center;
  cursor: pointer;
  outline-color: color-mix(in oklab, var(--office-ring) 50%, transparent);
  transition:
    border-color 200ms ease-out,
    background-color 200ms ease-out;
}
.file-upload__dropzone:hover:not(.is-disabled) {
  border-color: var(--file-upload-border-hover);
  background: color-mix(in oklab, var(--office-muted) 35%, transparent);
}
.file-upload__dropzone.is-dragging {
  border-color: var(--file-upload-border-active);
  background: color-mix(in oklab, var(--office-accent) 35%, transparent);
}
.file-upload__dropzone.is-disabled {
  background: color-mix(in oklab, var(--office-muted) 70%, transparent);
  color: var(--office-muted-foreground);
  cursor: not-allowed;
  opacity: 0.65;
}
.file-upload__icons {
  position: relative;
  width: 9rem;
  height: 3.5rem;
}
.file-upload__icon-card {
  position: absolute;
  top: 50%;
  left: 50%;
  display: grid;
  width: 3rem;
  height: 3rem;
  place-items: center;
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-xl);
  background: var(--office-background);
  color: var(--office-muted-foreground);
  background-clip: padding-box;
  box-shadow: var(--file-upload-card-shadow);
  transition:
    transform var(--file-upload-icon-duration) var(--file-upload-icon-ease),
    color var(--file-upload-icon-duration) var(--file-upload-icon-ease),
    background-color var(--file-upload-icon-duration) var(--file-upload-icon-ease);
}
.file-upload__icon-card::before {
  content: "";
  position: absolute;
  inset: 0;
  border-radius: calc(var(--office-radius-xl) - 1px);
  box-shadow: var(--file-upload-card-inner-shadow);
  pointer-events: none;
}
.file-upload__icon-card.is-middle {
  z-index: 10;
}
.file-upload__icon-card.is-active {
  background: var(--office-popover);
  color: var(--office-foreground);
  background-clip: border-box;
  box-shadow: var(--office-shadow-md);
}
.file-upload__icon {
  width: 1.25rem;
  height: 1.25rem;
}
.file-upload__copy {
  margin: 0;
}
.file-upload__copy > * + * {
  margin-top: var(--office-ui-space-1);
}
.file-upload__title {
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
  font-weight: 500;
}
.file-upload__description {
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
}
.file-upload__rejection {
  color: var(--office-destructive);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
}
.file-upload__chip {
  display: inline-flex;
  align-items: center;
  gap: var(--office-ui-space-2);
  padding: var(--office-ui-space-1) var(--office-ui-space-3);
  border: var(--office-border-width) solid var(--office-border);
  border-radius: 9999px;
  background: var(--office-background);
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
}
.file-upload__chip-icon {
  width: 0.875rem;
  height: 0.875rem;
}
.file-upload__input {
  display: none;
}
.file-upload__beam {
  position: absolute;
  inset: 0;
  padding: 1px;
  border-radius: inherit;
  background: conic-gradient(
    from var(--file-upload-beam-angle, 0deg),
    transparent 0deg,
    var(--office-info) 70deg,
    var(--office-success) 140deg,
    transparent 210deg
  );
  mask:
    linear-gradient(#fff 0 0) content-box,
    linear-gradient(#fff 0 0);
  mask-composite: exclude;
  -webkit-mask:
    linear-gradient(#fff 0 0) content-box,
    linear-gradient(#fff 0 0);
  -webkit-mask-composite: xor;
  pointer-events: none;
  animation: office-file-upload-beam 2.4s linear infinite;
}
.file-upload__beam.is-dark {
  background: conic-gradient(
    from var(--file-upload-beam-angle, 0deg),
    transparent 0deg,
    var(--office-chart-2) 70deg,
    var(--office-success) 140deg,
    transparent 210deg
  );
}
.file-upload__list {
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-xl);
  background: var(--office-background);
}
.file-upload__row {
  display: flex;
  align-items: center;
  gap: var(--office-ui-space-3);
  padding: calc(var(--office-ui-space-2) + var(--office-ui-space-0-5))
    var(--office-ui-space-3);
  border-bottom: var(--office-border-width) solid var(--office-border);
}
.file-upload__row:last-child {
  border-bottom: 0;
}
.file-upload__thumbnail {
  width: 2.5rem;
}
.file-upload__file {
  flex: 1 1 auto;
  min-width: 0;
}
.file-upload__file-name,
.file-upload__file-detail {
  margin: 0;
  overflow: hidden;
  white-space: nowrap;
  text-overflow: ellipsis;
}
.file-upload__file-name {
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
  font-weight: 500;
}
.file-upload__file-detail {
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
}
.file-upload__ready {
  flex: 0 0 auto;
  padding: var(--office-ui-space-1) var(--office-ui-space-2);
  border-radius: 9999px;
  background: var(--office-muted);
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
  white-space: nowrap;
}
[data-office-theme="dark"] .file-upload,
[data-office-theme-resolved="dark"] .file-upload {
  --file-upload-border: color-mix(in oklab, var(--office-foreground) 25%, transparent);
  --file-upload-border-hover: color-mix(in oklab, var(--office-foreground) 40%, transparent);
  --file-upload-border-active: color-mix(in oklab, var(--office-foreground) 40%, transparent);
  --file-upload-card-inner-shadow: 0 -1px 0 oklab(1 0 0 / 0.06);
}
[data-office-theme="dark"] .file-upload__icon-card::before,
[data-office-theme-resolved="dark"] .file-upload__icon-card::before {
  box-shadow: var(--file-upload-card-inner-shadow);
}
[data-office-theme="dark"] .file-upload__icon-card,
[data-office-theme-resolved="dark"] .file-upload__icon-card {
  background-clip: border-box;
}
@media (prefers-color-scheme: dark) {
  [data-office-theme="system"] .file-upload {
    --file-upload-border: color-mix(in oklab, var(--office-foreground) 25%, transparent);
    --file-upload-border-hover: color-mix(in oklab, var(--office-foreground) 40%, transparent);
    --file-upload-border-active: color-mix(in oklab, var(--office-foreground) 40%, transparent);
    --file-upload-card-inner-shadow: 0 -1px 0 oklab(1 0 0 / 0.06);
  }
  [data-office-theme="system"] .file-upload__icon-card::before {
    box-shadow: var(--file-upload-card-inner-shadow);
  }
  [data-office-theme="system"] .file-upload__icon-card {
    background-clip: border-box;
  }
}
@media (pointer: coarse) {
  .file-upload__dropzone {
    min-height: max(11rem, var(--office-ui-touch-target));
  }
  .file-upload__chip,
  .file-upload__row {
    min-height: var(--office-ui-touch-target);
  }
}
@media (prefers-reduced-motion: reduce) {
  .file-upload__dropzone,
  .file-upload__icon-card {
    transition: none;
  }
  .file-upload__beam {
    animation: none;
  }
}
@property --file-upload-beam-angle {
  syntax: "<angle>";
  inherits: false;
  initial-value: 0deg;
}
@keyframes office-file-upload-beam {
  to {
    --file-upload-beam-angle: 360deg;
  }
}
</style>
