<template>
  <section class="xlsx-authoring__section xlsx-number-format" data-testid="xlsx-number-format-section">
    <label class="xlsx-authoring__field">
      <span>{{ m.xlsx.authoring.numberFormat.category }}</span>
      <select v-model="category" data-testid="xlsx-number-format-category">
        <option value="number">{{ m.xlsx.authoring.numberFormat.number }}</option>
        <option value="currency">{{ m.xlsx.authoring.numberFormat.currency }}</option>
        <option value="percent">{{ m.xlsx.authoring.numberFormat.percent }}</option>
        <option value="date">{{ m.xlsx.authoring.numberFormat.date }}</option>
        <option value="text">{{ m.xlsx.authoring.numberFormat.text }}</option>
        <option value="custom">{{ m.xlsx.authoring.numberFormat.custom }}</option>
      </select>
    </label>
    <label v-if="showDecimals" class="xlsx-authoring__field">
      <span>{{ m.xlsx.authoring.numberFormat.decimals }}</span>
      <input
        v-model.number="decimals"
        type="number"
        min="0"
        max="10"
        step="1"
        data-testid="xlsx-number-format-decimals"
      />
    </label>
    <label v-if="showDecimals" class="xlsx-authoring__check">
      <input v-model="thousands" type="checkbox" data-testid="xlsx-number-format-thousands" />
      <span>{{ m.xlsx.authoring.numberFormat.thousands }}</span>
    </label>
    <label v-if="category === 'custom'" class="xlsx-authoring__field">
      <span>{{ m.xlsx.authoring.numberFormat.pattern }}</span>
      <input v-model="customPattern" type="text" data-testid="xlsx-number-format-pattern" />
    </label>
    <p class="xlsx-authoring__hint" data-testid="xlsx-number-format-preview">
      {{ m.xlsx.authoring.numberFormat.preview }}: <strong>{{ preview }}</strong>
      <code data-testid="xlsx-number-format-code">{{ pattern }}</code>
    </p>
    <div class="xlsx-authoring__actions">
      <button
        type="button"
        class="xlsx-authoring__button xlsx-authoring__button--primary"
        data-testid="xlsx-number-format-apply"
        @click="applyFormat"
      >{{ m.xlsx.authoring.numberFormat.apply }}</button>
    </div>
  </section>
</template>

<script setup lang="ts">
import { computed, ref } from "vue";
import type { XlsxViewerController } from "@arcships/xlsx-core";
import {
  buildNumberFormatPattern,
  buildNumberFormatPreview,
  clampNumberFormatDecimals,
  resolveNumberFormatInput,
  type XlsxNumberFormatCategory
} from "../composables/number-format-dialog";
import { useXlsxMessages } from "../i18n";

const props = defineProps<{ controller: XlsxViewerController }>();
const emit = defineEmits<{ applied: [] }>();

const m = useXlsxMessages();
const category = ref<XlsxNumberFormatCategory>("number");
const decimals = ref(2);
const thousands = ref(true);
const customPattern = ref("#,##0.00");

const pattern = computed(() => buildNumberFormatPattern({
  category: category.value,
  customPattern: customPattern.value,
  decimals: decimals.value,
  thousands: thousands.value
}));
const preview = computed(() => buildNumberFormatPreview({
  category: category.value,
  customPattern: customPattern.value,
  decimals: decimals.value,
  thousands: thousands.value
}));
const showDecimals = computed(() =>
  category.value === "number" || category.value === "currency" || category.value === "percent"
);

function applyFormat() {
  const format = resolveNumberFormatInput({
    category: category.value,
    customPattern: customPattern.value,
    decimals: clampNumberFormatDecimals(decimals.value),
    thousands: thousands.value
  });
  if (props.controller.setSelectionNumberFormat?.(format)) {
    emit("applied");
  }
}
</script>
