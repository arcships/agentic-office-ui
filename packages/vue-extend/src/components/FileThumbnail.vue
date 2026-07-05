<script setup lang="ts">
import { computed } from "vue"

export interface FileThumbnailFile {
  name: string
  type: string
  url?: string
}

export interface FileThumbnailProps {
  file: FileThumbnailFile
  size?: "sm" | "md" | "lg"
  className?: string
}

const props = withDefaults(defineProps<FileThumbnailProps>(), {
  size: "md",
  className: "",
})

const sizePx: Record<"sm" | "md" | "lg", number> = {
  sm: 32,
  md: 48,
  lg: 64,
}

const extension = computed(() => {
  const parts = props.file.name.split(".")
  return parts.length > 1 ? parts[parts.length - 1]!.toLowerCase() : ""
})

const isImage = computed(() => props.file.type.startsWith("image/") && !!props.file.url)

const iconByExt: Record<string, string> = {
  pdf: "PDF",
  doc: "DOC",
  docx: "DOC",
  xls: "XLS",
  xlsx: "XLS",
  ppt: "PPT",
  pptx: "PPT",
  txt: "TXT",
  csv: "CSV",
  zip: "ZIP",
  rar: "RAR",
  mp4: "VID",
  mov: "VID",
  mp3: "AUD",
  wav: "AUD",
}

const label = computed(() => {
  const ext = extension.value
  return (iconByExt[ext] ?? (ext.toUpperCase() || "FILE")).slice(0, 3)
})

const colorByType = computed(() => {
  const type = props.file.type
  if (type.startsWith("image/")) return "bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-400"
  if (type === "application/pdf") return "bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400"
  if (type.includes("word") || type.includes("document")) return "bg-sky-100 text-sky-700 dark:bg-sky-900/30 dark:text-sky-400"
  if (type.includes("sheet") || type.includes("excel") || type.includes("csv")) return "bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-400"
  if (type.startsWith("video/")) return "bg-purple-100 text-purple-700 dark:bg-purple-900/30 dark:text-purple-400"
  if (type.startsWith("audio/")) return "bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-400"
  return "bg-muted text-muted-foreground"
})

const iconStyle = computed(() => {
  const size = sizePx[props.size]
  return { width: `${size}px`, height: `${size}px`, fontSize: props.size === "lg" ? "16px" : props.size === "sm" ? "12px" : "14px" }
})
</script>

<template>
  <div :class="['file-thumbnail', props.className]">
    <div :class="['file-thumbnail-icon', isImage ? '' : colorByType]" :style="iconStyle">
      <img v-if="isImage" :src="props.file.url" :alt="props.file.name" class="file-thumbnail-image" style="width: 100%; height: 100%; object-fit: cover; display: block;" />
      <span v-else class="font-semibold">{{ label }}</span>
    </div>
    <div class="file-thumbnail-meta">
      <p class="file-thumbnail-name">{{ props.file.name }}</p>
      <p class="file-thumbnail-type">{{ props.file.type || "Unknown type" }}</p>
    </div>
  </div>
</template>

<style scoped>
.file-thumbnail { display: flex; align-items: center; gap: 12px; min-width: 0; max-width: 100%; }
.file-thumbnail-icon { display: flex; flex: 0 0 auto; align-items: center; justify-content: center; overflow: hidden; border-radius: 6px; }
.file-thumbnail-image { width: 100%; height: 100%; object-fit: cover; display: block; }
.file-thumbnail-meta { min-width: 0; flex: 1 1 auto; }
.file-thumbnail-name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 14px; font-weight: 500; color: var(--foreground, #111); }
.file-thumbnail-type { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 12px; color: var(--muted-foreground, #737373); }
.font-semibold { font-weight: 600; }
</style>
