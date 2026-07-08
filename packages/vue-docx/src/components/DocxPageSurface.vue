<template>
  <DocxPageWrapper :page="page">
    <DocxPageHeader v-if="hasHeader" :page="page" />
    <DocxPageBody :blocks="page.blocks" />
    <DocxPageFooter v-if="hasFooter" :page="page" />
    <DocxImageLayer
      v-if="hasFloatingImages"
      :page="page"
      :controller="controller"
    />
    <DocxFormFieldLayer
      v-if="hasFormFields"
      :page="page"
      :controller="controller"
    />
    <DocxTrackedChangeGutter
      v-if="showGutter"
      :page="page"
      :controller="controller"
    />
  </DocxPageWrapper>
</template>

<script setup lang="ts">
import { computed } from "vue"
import type { DocxEditorController, LayoutPage } from "@extend-ai/docx-core"
import DocxPageWrapper from "./DocxPageWrapper.vue"
import DocxPageHeader from "./DocxPageHeader.vue"
import DocxPageFooter from "./DocxPageFooter.vue"
import DocxPageBody from "./DocxPageBody"
import DocxImageLayer from "./DocxImageLayer.vue"
import DocxFormFieldLayer from "./DocxFormFieldLayer.vue"
import DocxTrackedChangeGutter from "./DocxTrackedChangeGutter.vue"

const props = defineProps<{
  page: LayoutPage
  editable?: boolean
  controller?: DocxEditorController
}>()

const hasHeader = computed(() => false)
const hasFooter = computed(() => false)
const hasFloatingImages = computed(() => false)
const hasFormFields = computed(() => false)
const showGutter = computed(
  () => props.controller?.showTrackedChanges || props.controller?.showComments
)
</script>
