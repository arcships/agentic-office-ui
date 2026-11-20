<template>
  <section class="xlsx-authoring__section xlsx-names" data-testid="xlsx-name-manager-section">
    <ul class="xlsx-authoring__rules" data-testid="xlsx-name-list">
      <li
        v-for="entry in names"
        :key="entry.name"
        class="xlsx-authoring__rule"
        data-testid="xlsx-name-item"
      >
        <button
          type="button"
          class="xlsx-authoring__link xlsx-names__select"
          :aria-pressed="selectedName === entry.name"
          @click="selectEntry(entry)"
        >
          <strong>{{ entry.name }}</strong>
          <span>{{ entry.refersTo }}</span>
        </button>
        <button
          type="button"
          class="xlsx-authoring__link"
          data-testid="xlsx-name-delete"
          @click="removeEntry(entry.name)"
        >{{ m.xlsx.authoring.names.delete }}</button>
      </li>
      <li v-if="names.length === 0" class="xlsx-authoring__hint" data-testid="xlsx-name-empty">
        {{ m.xlsx.authoring.names.empty }}
      </li>
    </ul>
    <label class="xlsx-authoring__field">
      <span>{{ m.xlsx.authoring.names.name }}</span>
      <input
        v-model="name"
        type="text"
        data-testid="xlsx-name-input"
        :placeholder="m.xlsx.authoring.names.namePlaceholder"
      />
    </label>
    <label class="xlsx-authoring__field">
      <span>{{ m.xlsx.authoring.names.refersTo }}</span>
      <input
        v-model="refersTo"
        type="text"
        data-testid="xlsx-name-refers-to"
        :placeholder="m.xlsx.authoring.names.refersToPlaceholder"
      />
    </label>
    <div class="xlsx-authoring__actions">
      <button
        type="button"
        class="office-button office-button--outline office-button--size-sm xlsx-authoring__button"
        data-testid="xlsx-name-use-selection"
        @click="useSelection"
      >{{ m.xlsx.authoring.names.useSelection }}</button>
      <button
        type="button"
        class="office-button office-button--default office-button--size-sm xlsx-authoring__button xlsx-authoring__button--primary"
        data-testid="xlsx-name-save"
        :disabled="!name.trim() || !refersTo.trim()"
        @click="saveEntry"
      >{{ selectedName ? m.xlsx.authoring.names.update : m.xlsx.authoring.names.add }}</button>
    </div>
  </section>
</template>

<script setup lang="ts">
import { computed, ref } from "vue";
import type { XlsxNamedRange, XlsxViewerController } from "@arcships/xlsx-core";
import { formatDefinedNameRange } from "../composables/number-format-dialog";
import { useXlsxMessages } from "../i18n";

const props = defineProps<{ controller: XlsxViewerController }>();
const emit = defineEmits<{ applied: [] }>();

const m = useXlsxMessages();
const refreshKey = ref(0);
const revision = computed(() => props.controller.revision);
const names = computed<XlsxNamedRange[]>(() => {
  revision.value;
  refreshKey.value;
  return props.controller.getNamedRanges?.() ?? [];
});
const selectedName = ref("");
const name = ref("");
const refersTo = ref("");

function selectEntry(entry: XlsxNamedRange) {
  selectedName.value = entry.name;
  name.value = entry.name;
  refersTo.value = entry.refersTo;
}

function useSelection() {
  const sheet = props.controller.activeSheet;
  const address = props.controller.selectedRangeAddress;
  if (!sheet || !address) {
    return;
  }
  refersTo.value = formatDefinedNameRange(sheet.name, address);
}

function saveEntry() {
  const nextName = name.value.trim();
  const nextRefersTo = refersTo.value.trim();
  if (!nextName || !nextRefersTo) {
    return;
  }
  const previous = selectedName.value;
  const saved = props.controller.setNamedRange?.(nextName, nextRefersTo);
  if (!saved) {
    return;
  }
  if (previous && previous.toLowerCase() !== nextName.toLowerCase()) {
    props.controller.deleteNamedRange?.(previous);
  }
  selectedName.value = nextName;
  refreshKey.value += 1;
  emit("applied");
}

function removeEntry(entryName: string) {
  if (!props.controller.deleteNamedRange?.(entryName)) {
    return;
  }
  if (selectedName.value.toLowerCase() === entryName.toLowerCase()) {
    selectedName.value = "";
    name.value = "";
    refersTo.value = "";
  }
  refreshKey.value += 1;
  emit("applied");
}
</script>

<style scoped>
.xlsx-names__select {
  align-items: flex-start;
  display: inline-flex;
  flex: 1;
  flex-direction: column;
  gap: 2px;
  text-align: left;
}
.xlsx-names__select span {
  color: #71717a;
  font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
  font-size: 11px;
}
</style>
