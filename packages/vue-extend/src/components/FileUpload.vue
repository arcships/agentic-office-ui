<script setup lang="ts">
import { ref, computed } from "vue"

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

const dropZoneClass = computed(() => {
  const base =
    "relative flex flex-col items-center justify-center gap-2 rounded-lg border-2 border-dashed p-8 transition-colors"
  if (props.disabled) return `${base} border-muted bg-muted/50 cursor-not-allowed text-muted-foreground`
  if (isDragging.value) return `${base} border-primary bg-primary/5 text-primary cursor-copy`
  return `${base} border-border bg-background hover:border-muted-foreground/50 hover:bg-accent/50 cursor-pointer text-muted-foreground`
})

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
  <div>
    <div :class="[dropZoneClass, props.className]" @dragover="onDragOver" @dragleave="onDragLeave" @drop="onDrop" @click="openPicker">
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
          class="size-8"
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
      <div class="text-center text-sm">
        <p class="font-medium">
          <slot name="label">Drag & drop files here, or click to select</slot>
        </p>
        <p v-if="props.accept" class="text-xs text-muted-foreground">
          Accepted: {{ props.accept }}
        </p>
        <p v-if="props.maxSize !== Infinity" class="text-xs text-muted-foreground">
          Max size: {{ (props.maxSize / (1024 * 1024)).toFixed(1) }}MB
        </p>
      </div>
    </div>
    <p v-if="error" class="mt-2 text-sm text-destructive">{{ error }}</p>
  </div>
</template>
