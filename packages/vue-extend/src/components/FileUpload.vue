<script setup lang="ts">
import { computed, ref } from "vue"

export interface FileUploadProps {
  accept?: string
  multiple?: boolean
  maxSize?: number // bytes
  disabled?: boolean
  className?: string
}

const props = withDefaults(defineProps<FileUploadProps>(), {
  accept: "",
  multiple: false,
  maxSize: Infinity,
  disabled: false,
  className: "",
})

const emit = defineEmits<{
  "files-accepted": [files: File[]]
}>()

const isDragging = ref(false)
const error = ref("")
const inputRef = ref<HTMLInputElement | null>(null)

const dropZoneClass = computed(() => ({
  "is-disabled": props.disabled,
  "is-dragging": isDragging.value && !props.disabled,
}))

function validateFiles(files: FileList): File[] {
  const accepted: File[] = []
  const errors: string[] = []

  for (let i = 0; i < files.length; i++) {
    const file = files[i]
    if (props.accept) {
      const types = props.accept.split(",").map((t) => t.trim())
      const matches = types.some((type) => {
        if (type.endsWith("/*")) {
          return file.type.startsWith(type.slice(0, -1))
        }
        if (type.startsWith(".")) {
          return file.name.toLowerCase().endsWith(type.toLowerCase())
        }
        return file.type === type
      })
      if (!matches) {
        errors.push(`"${file.name}" is not an accepted file type`)
        continue
      }
    }
    if (file.size > props.maxSize) {
      const maxMB = (props.maxSize / (1024 * 1024)).toFixed(1)
      errors.push(`"${file.name}" exceeds the ${maxMB}MB size limit`)
      continue
    }
    accepted.push(file)
  }

  if (errors.length) {
    error.value = errors.join("; ")
  }
  return accepted
}

function handleFiles(files: FileList) {
  error.value = ""
  const valid = validateFiles(files)
  if (valid.length) {
    emit("files-accepted", valid)
  }
}

function onDragOver(e: DragEvent) {
  e.preventDefault()
  if (!props.disabled) isDragging.value = true
}

function onDragLeave(e: DragEvent) {
  e.preventDefault()
  isDragging.value = false
}

function onDrop(e: DragEvent) {
  e.preventDefault()
  isDragging.value = false
  if (props.disabled || !e.dataTransfer?.files.length) return
  handleFiles(e.dataTransfer.files)
}

function onInputChange(e: Event) {
  const input = e.target as HTMLInputElement
  if (input.files?.length) {
    handleFiles(input.files)
    input.value = ""
  }
}

function openPicker() {
  if (!props.disabled) inputRef.value?.click()
}
</script>

<template>
  <div class="file-upload-root">
    <div
      :class="['file-upload-zone', dropZoneClass, props.className]"
      role="button"
      :aria-disabled="props.disabled"
      tabindex="0"
      @dragover="onDragOver"
      @dragleave="onDragLeave"
      @drop="onDrop"
      @click="openPicker"
      @keydown.enter.prevent="openPicker"
      @keydown.space.prevent="openPicker"
    >
      <input
        ref="inputRef"
        type="file"
        :accept="props.accept"
        :multiple="props.multiple"
        :disabled="props.disabled"
        class="sr-only"
        tabindex="-1"
        @change="onInputChange"
      />
      <slot name="icon">
        <svg
          class="file-upload-icon"
          xmlns="http://www.w3.org/2000/svg"
          fill="none"
          viewBox="0 0 24 24"
          stroke-width="1.5"
          stroke="currentColor"
          aria-hidden="true"
        >
          <path stroke-linecap="round" stroke-linejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5m-13.5-9L12 3m0 0l4.5 4.5M12 3v13.5" />
        </svg>
      </slot>
      <div class="file-upload-copy">
        <p class="file-upload-label">
          <slot name="label">Drag & drop files here, or click to select</slot>
        </p>
        <p v-if="props.accept" class="file-upload-meta">
          Accepted: {{ props.accept }}
        </p>
        <p v-if="props.maxSize !== Infinity" class="file-upload-meta">
          Max size: {{ (props.maxSize / (1024 * 1024)).toFixed(1) }}MB
        </p>
      </div>
    </div>
    <p v-if="error" class="file-upload-error">{{ error }}</p>
  </div>
</template>

<style scoped>
.file-upload-root { min-width: 0; }
.file-upload-zone {
  position: relative;
  display: flex;
  min-height: 180px;
  min-width: 0;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 8px;
  border: 2px dashed var(--border, #e5e7eb);
  border-radius: var(--radius, 10px);
  background: var(--background, #fff);
  color: var(--muted-foreground, #737373);
  cursor: pointer;
  padding: 32px;
  text-align: center;
  transition: border-color 0.15s ease, background-color 0.15s ease, color 0.15s ease, box-shadow 0.15s ease;
}
.file-upload-zone:hover:not(.is-disabled) {
  border-color: color-mix(in oklch, var(--muted-foreground, #737373) 50%, var(--border, #e5e7eb));
  background: color-mix(in oklch, var(--accent, #f5f5f5) 55%, transparent);
}
.file-upload-zone:focus-visible {
  outline: 2px solid var(--ring, #3b82f6);
  outline-offset: 2px;
}
.file-upload-zone.is-dragging {
  border-color: var(--primary, #111827);
  background: color-mix(in oklch, var(--primary, #111827) 7%, transparent);
  color: var(--primary, #111827);
  box-shadow: 0 0 0 3px color-mix(in oklch, var(--primary, #111827) 16%, transparent);
}
.file-upload-zone.is-disabled {
  background: color-mix(in oklch, var(--muted, #f5f5f5) 70%, transparent);
  color: var(--muted-foreground, #737373);
  cursor: not-allowed;
  opacity: 0.65;
}
.file-upload-icon { width: 32px; height: 32px; flex-shrink: 0; }
.file-upload-copy { font-size: 14px; line-height: 1.35; }
.file-upload-label { margin: 0; font-weight: 500; color: var(--foreground, #111827); }
.file-upload-meta { margin: 4px 0 0; color: var(--muted-foreground, #737373); font-size: 12px; }
.file-upload-error { margin: 8px 0 0; color: var(--destructive, #ef4444); font-size: 14px; }
.sr-only {
  position: absolute;
  width: 1px;
  height: 1px;
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
}
@media (max-width: 640px) {
  .file-upload-zone { min-height: 148px; padding: 20px; }
}
</style>
