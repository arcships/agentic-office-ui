<template>
  <section class="xlsx-authoring__section xlsx-paste-special" data-testid="xlsx-paste-special-section">
    <label class="xlsx-authoring__field">
      <span>{{ m.xlsx.authoring.pasteSpecial.mode }}</span>
      <select v-model="mode" data-testid="xlsx-paste-special-mode">
        <option value="all">{{ m.xlsx.authoring.pasteSpecial.modeAll }}</option>
        <option value="values">{{ m.xlsx.authoring.pasteSpecial.modeValues }}</option>
        <option value="formats">{{ m.xlsx.authoring.pasteSpecial.modeFormats }}</option>
      </select>
    </label>
    <label class="xlsx-authoring__check">
      <input v-model="transpose" type="checkbox" data-testid="xlsx-paste-special-transpose" />
      <span>{{ m.xlsx.authoring.pasteSpecial.transpose }}</span>
    </label>
    <p v-if="status" class="xlsx-authoring__hint" data-testid="xlsx-paste-special-status">{{ status }}</p>
    <div class="xlsx-authoring__actions">
      <button
        type="button"
        class="xlsx-authoring__button xlsx-authoring__button--primary"
        data-testid="xlsx-paste-special-apply"
        @click="applyPaste"
      >{{ m.xlsx.authoring.pasteSpecial.paste }}</button>
    </div>
  </section>
</template>

<script setup lang="ts">
import { ref } from "vue";
import type { XlsxPasteSpecialOptions, XlsxViewerController } from "@arcships/xlsx-core";
import { useXlsxMessages } from "../i18n";

const props = defineProps<{ controller: XlsxViewerController }>();
const emit = defineEmits<{ applied: [] }>();

const m = useXlsxMessages();
const mode = ref<NonNullable<XlsxPasteSpecialOptions["mode"]>>("all");
const transpose = ref(false);
const status = ref("");

async function applyPaste() {
  const applied = await props.controller.pasteSpecialFromClipboard?.({
    mode: mode.value,
    transpose: transpose.value
  });
  if (applied) {
    status.value = "";
    emit("applied");
    return;
  }
  status.value = m.value.xlsx.authoring.pasteSpecial.empty;
}
</script>
