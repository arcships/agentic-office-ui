<template>
  <div class="xlsx-sheettabs" :style="tabsStyle">
    <button
      type="button"
      class="xlsx-sheettabs__nav-btn"
      :disabled="activeTabIndex <= 0"
      :title="m.xlsx.sheetTabs.previous"
      :aria-label="m.xlsx.sheetTabs.previous"
      @click="prevTab"
    >
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="m14.5 6-6 6 6 6" /></svg>
    </button>
    <div class="xlsx-sheettabs__list">
      <div
        v-for="(tab, index) in controller.tabs"
        :key="tab.id ?? index"
        class="xlsx-sheettabs__tab-wrapper"
        @focusin="onTabFocus(index)"
        @mouseenter="onTabHover(index, $event)"
        @mouseleave="onTabLeave"
        @contextmenu="onTabContextMenu(index, $event)"
      >
        <input
          v-if="renamingTabIndex === index"
          ref="renameInputRef"
          class="xlsx-sheettabs__rename-input"
          data-testid="xlsx-sheet-rename-input"
          type="text"
          :aria-invalid="renameError ? 'true' : 'false'"
          :aria-label="m.xlsx.sheetTabs.renameLabel"
          :maxlength="maxSheetNameLength"
          :style="renameInputStyle(index)"
          :value="renameDraft"
          @blur="onRenameBlur"
          @click.stop
          @dblclick.stop
          @input="onRenameInput"
          @keydown.enter.prevent.stop="commitRename"
          @keydown.esc.prevent.stop="cancelRename"
        />
        <button
          v-else
          class="xlsx-sheettabs__tab"
          data-testid="xlsx-sheet-tab"
          :data-sheet-name="tab.name"
          role="tab"
          :aria-selected="index === activeTabIndex"
          :class="{
            'xlsx-sheettabs__tab--active': index === activeTabIndex,
          }"
          :style="tabStyle(index)"
          @click="onTabClick(index)"
          @dblclick="onTabDblClick(tab, index)"
          @keydown="onTabKeydown(index, tab, $event)"
        >
          <span class="xlsx-sheettabs__tab-name">{{ tab.name }}</span>
        </button>
        <div
          v-if="hoveredTabIndex === index && renamingTabIndex !== index && thumbnailComputed[index]"
          class="xlsx-sheettabs__thumbnail"
          :style="thumbnailPopupStyle"
        >
          <canvas
            :ref="(el) => setThumbnailCanvas(index, el)"
            :height="thumbnailComputed[index]?.height"
            :width="thumbnailComputed[index]?.width"
          />
        </div>
      </div>
      <button
        v-if="!isReadOnly"
        type="button"
        class="xlsx-sheettabs__add-btn"
        data-testid="xlsx-add-sheet"
        :title="m.xlsx.sheetTabs.add"
        :aria-label="m.xlsx.sheetTabs.add"
        @click="controller.addSheet()"
      >
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 5v14M5 12h14" /></svg>
      </button>
      <button
        v-if="!isReadOnly"
        type="button"
        class="xlsx-sheettabs__remove-btn"
        data-testid="xlsx-remove-sheet"
        :disabled="!canRemoveActiveSheet"
        :title="m.xlsx.sheetTabs.remove"
        :aria-label="m.xlsx.sheetTabs.remove"
        @click="removeActiveSheet"
      >
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 7h14M9 7V4h6v3M8 10v8M12 10v8M16 10v8M7 7l1 14h8l1-14" /></svg>
      </button>
    </div>
    <button
      type="button"
      class="xlsx-sheettabs__nav-btn"
      :disabled="activeTabIndex >= controller.tabs.length - 1"
      :title="m.xlsx.sheetTabs.next"
      :aria-label="m.xlsx.sheetTabs.next"
      @click="nextTab"
    >
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="m9.5 6 6 6-6 6" /></svg>
    </button>
    <div
      v-if="renameError"
      class="xlsx-sheettabs__rename-error"
      data-testid="xlsx-sheet-rename-error"
      role="alert"
      :style="renameErrorStyle"
    >
      {{ renameError }}
    </div>
    <div
      v-if="tabMenu"
      class="xlsx-sheettabs__menu-backdrop"
      data-testid="xlsx-sheet-tab-menu-backdrop"
      @contextmenu.prevent="closeTabMenu"
      @pointerdown="closeTabMenu"
    />
    <div
      v-if="tabMenu"
      class="xlsx-sheettabs__menu"
      data-testid="xlsx-sheet-tab-menu"
      role="menu"
      :aria-label="m.xlsx.sheetTabs.rename"
      :style="tabMenuStyle"
      @keydown.esc="closeTabMenu"
    >
      <button
        type="button"
        class="xlsx-sheettabs__menu-item"
        data-testid="xlsx-rename-sheet"
        role="menuitem"
        @click="renameFromMenu"
      >
        {{ m.xlsx.sheetTabs.rename }}
      </button>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, ref, watch, nextTick, onUnmounted, type CSSProperties } from "vue";
import type {
  XlsxSheetNameIssue,
  XlsxViewerController,
  XlsxWorkbookTab,
  XlsxSheetThumbnail,
  UseXlsxViewerThumbnailsOptions,
} from "@arcships/xlsx-core";
import { XLSX_MAX_SHEET_NAME_LENGTH } from "@arcships/xlsx-core";
import { useXlsxViewerThumbnails } from "../composables/useXlsxViewerThumbnails";
import { useXlsxMessages } from "../i18n";

const props = defineProps<{
  controller: XlsxViewerController;
  isDark?: boolean;
  readOnly?: boolean;
  thumbnailOptions?: UseXlsxViewerThumbnailsOptions;
}>();

const activeTabIndex = computed(() => props.controller.activeTabIndex);
const isReadOnly = computed(() => props.controller.isEditingLocked === true || props.controller.readOnly || props.readOnly === true);
// Renaming a protected sheet is allowed, like Excel; only read-only mode blocks it.
const canRenameSheets = computed(() => props.controller.readOnly !== true && props.readOnly !== true);
const m = useXlsxMessages();
const canRemoveActiveSheet = computed(() => Boolean(props.controller.activeSheet) && props.controller.sheets.length > 1);
const maxSheetNameLength = XLSX_MAX_SHEET_NAME_LENGTH;

const { prefetchThumbnails, thumbnails: thumbnailComputedRaw } = useXlsxViewerThumbnails(
  () => props.controller,
  props.thumbnailOptions ?? { includeHeaders: true, resolution: { maxHeight: 132, maxWidth: 200 } },
);

const thumbnailComputed = computed(() => (thumbnailComputedRaw as unknown as XlsxSheetThumbnail[]));

const hoveredTabIndex = ref<number | null>(null);
const thumbnailCanvasRefs = ref<Map<number, HTMLCanvasElement | null>>(new Map());
const popupX = ref(0);
const popupY = ref(0);
let popupTimer: ReturnType<typeof setTimeout> | null = null;

// Inline rename state. The tab id, not the index, identifies the editor so a
// sheet added or removed underneath the input cannot retarget the commit.
const renamingTabId = ref<string | null>(null);
const renamingTabIndex = computed(() => {
  if (renamingTabId.value === null) return -1;
  return props.controller.tabs.findIndex((tab) => tab.id === renamingTabId.value);
});
const renameDraft = ref("");
const renameError = ref<string | null>(null);
// The message is positioned against the editor and fixed to the viewport, so
// the scrollable tab list cannot clip it.
const renameAnchor = ref<{ left: number; top: number } | null>(null);
// Vue collects element refs inside `v-for` into an array, so the single
// inline editor is read back through this helper.
const renameInputRef = ref<HTMLInputElement | HTMLInputElement[] | null>(null);
const tabMenu = ref<{ tabId: string; x: number; y: number } | null>(null);

// Prefetch tuning: one pending idle warm-up at most, for the hovered tab and
// its immediate neighbours. Sweeping the pointer across the strip cancels the
// previous task instead of queueing every tab, and no workspace is read until
// something actually points at it.
const THUMBNAIL_PREFETCH_RADIUS = 1;
type IdleHandle = number | ReturnType<typeof setTimeout>;
const idleHost = globalThis as {
  cancelIdleCallback?: (handle: number) => void;
  requestIdleCallback?: (callback: () => void, options?: { timeout: number }) => number;
};
let prefetchHandle: IdleHandle | null = null;

function cancelScheduledPrefetch(): void {
  if (prefetchHandle === null) return;
  if (typeof prefetchHandle === "number" && idleHost.cancelIdleCallback) {
    idleHost.cancelIdleCallback(prefetchHandle);
  } else {
    clearTimeout(prefetchHandle as ReturnType<typeof setTimeout>);
  }
  prefetchHandle = null;
}

function scheduleHoverPrefetch(index: number): void {
  cancelScheduledPrefetch();
  const tabCount = props.controller.tabs.length;
  const sheetIndices: number[] = [];
  for (let offset = -THUMBNAIL_PREFETCH_RADIUS; offset <= THUMBNAIL_PREFETCH_RADIUS; offset += 1) {
    const candidate = index + offset;
    if (candidate >= 0 && candidate < tabCount) sheetIndices.push(candidate);
  }
  const run = () => {
    prefetchHandle = null;
    prefetchThumbnails(sheetIndices);
  };
  prefetchHandle = typeof idleHost.requestIdleCallback === "function"
    ? idleHost.requestIdleCallback(run, { timeout: 200 })
    : setTimeout(run, 0);
}

const thumbnailPopupStyle = computed<CSSProperties>(() => ({
  left: `${popupX.value}px`,
  position: "fixed",
  top: `${popupY.value}px`,
  zIndex: 9999,
  transform: "translate(-50%, calc(-100% - 8px))",
  backgroundColor: props.isDark ? "#27272a" : "#fff",
  border: `1px solid ${props.isDark ? "#3f3f46" : "#e4e4e7"}`,
  borderRadius: "8px",
  boxShadow: "0 4px 16px rgba(0,0,0,0.12)",
  padding: "8px",
  pointerEvents: "none",
}));

function setThumbnailCanvas(index: number, el: unknown) {
  const canvas = el as HTMLCanvasElement | null;
  thumbnailCanvasRefs.value.set(index, canvas);
}

function onTabHover(index: number, event: MouseEvent) {
  if (popupTimer) clearTimeout(popupTimer);
  scheduleHoverPrefetch(index);
  const rect = (event.currentTarget as HTMLElement).getBoundingClientRect();
  popupX.value = rect.left + rect.width / 2;
  popupY.value = rect.top;
  hoveredTabIndex.value = index;

  nextTick(() => {
    const canvas = thumbnailCanvasRefs.value.get(index);
    const thumb = thumbnailComputed.value[index];
    if (canvas && thumb) {
      thumb.paint(canvas as HTMLCanvasElement);
    }
  });
}

function onTabFocus(index: number) {
  // Keyboard users reach the same preview; warm it without opening the popup.
  scheduleHoverPrefetch(index);
}

function onTabLeave() {
  cancelScheduledPrefetch();
  popupTimer = setTimeout(() => {
    hoveredTabIndex.value = null;
  }, 150);
}

function onTabClick(index: number) {
  props.controller.setActiveTabIndex(index);
}

function onTabDblClick(tab: XlsxWorkbookTab, index: number) {
  startRename(tab, index);
}

function renameErrorMessage(issue: XlsxSheetNameIssue) {
  const messages = m.value.xlsx;
  switch (issue) {
    case "empty":
      return messages.sheetTabs.renameEmpty;
    case "too-long":
      return messages.sheetTabs.renameTooLong;
    case "duplicate":
      return messages.sheetTabs.renameDuplicate;
    case "apostrophe":
      return messages.sheetTabs.renameApostrophe;
    case "invalid-characters":
    default:
      return messages.diagnostics.sheetNameInvalid;
  }
}

function closeRename() {
  renamingTabId.value = null;
  renameDraft.value = "";
  renameError.value = null;
  renameAnchor.value = null;
}

function refreshRenameAnchor() {
  const target = renameInputRef.value;
  const input = Array.isArray(target) ? target[0] : target;
  const rect = input?.getBoundingClientRect?.();
  renameAnchor.value = rect ? { left: rect.left, top: rect.top } : null;
}

function showRenameError(message: string) {
  renameError.value = message;
  refreshRenameAnchor();
}

function startRename(tab: XlsxWorkbookTab, index: number) {
  if (!canRenameSheets.value || tab.kind !== "sheet" || typeof tab.sheetIndex !== "number") {
    return;
  }
  closeTabMenu();
  props.controller.setActiveTabIndex(index);
  renamingTabId.value = tab.id;
  renameDraft.value = tab.name;
  renameError.value = null;
  void nextTick(() => {
    const target = renameInputRef.value;
    const input = Array.isArray(target) ? target[0] : target;
    input?.focus();
    input?.select();
    refreshRenameAnchor();
  });
}

function commitRename() {
  const index = renamingTabIndex.value;
  const tab = index >= 0 ? props.controller.tabs[index] : null;
  if (!tab || tab.kind !== "sheet" || typeof tab.sheetIndex !== "number") {
    closeRename();
    return;
  }
  const issue = props.controller.getSheetNameError(renameDraft.value, tab.sheetIndex);
  if (issue) {
    showRenameError(renameErrorMessage(issue));
    return;
  }
  if (!props.controller.renameSheet(tab.sheetIndex, renameDraft.value)) {
    showRenameError(m.value.xlsx.sheetTabs.renameFailed);
    return;
  }
  closeRename();
}

function cancelRename() {
  closeRename();
}

function onRenameInput(event: Event) {
  renameDraft.value = (event.target as HTMLInputElement | null)?.value ?? "";
  renameError.value = null;
}

function onRenameBlur() {
  // Blur commits, matching Excel; a rejected name keeps the editor open with
  // the reason visible instead of dropping the edit silently.
  if (renamingTabId.value === null) return;
  commitRename();
}

function onTabContextMenu(index: number, event: MouseEvent) {
  const tab = props.controller.tabs[index];
  if (!tab || tab.kind !== "sheet" || !canRenameSheets.value) {
    closeTabMenu();
    return;
  }
  event.preventDefault();
  props.controller.setActiveTabIndex(index);
  tabMenu.value = { tabId: tab.id, x: event.clientX, y: event.clientY };
}

function onTabKeydown(index: number, tab: XlsxWorkbookTab, event: KeyboardEvent) {
  // Keyboard users reach the same menu with the ContextMenu key or Shift+F10.
  if (event.key !== "ContextMenu" && !(event.shiftKey && event.key === "F10")) {
    return;
  }
  if (!canRenameSheets.value || tab.kind !== "sheet") {
    return;
  }
  event.preventDefault();
  const rect = (event.currentTarget as HTMLElement | null)?.getBoundingClientRect?.();
  props.controller.setActiveTabIndex(index);
  tabMenu.value = { tabId: tab.id, x: rect?.left ?? 0, y: rect?.bottom ?? 0 };
}

function closeTabMenu() {
  tabMenu.value = null;
}

function renameFromMenu() {
  const menu = tabMenu.value;
  closeTabMenu();
  if (!menu) return;
  const index = props.controller.tabs.findIndex((tab) => tab.id === menu.tabId);
  const tab = index >= 0 ? props.controller.tabs[index] : null;
  if (tab) startRename(tab, index);
}

function removeActiveSheet() {
  if (!canRemoveActiveSheet.value || isReadOnly.value) return;
  props.controller.removeActiveSheet();
}

function prevTab() {
  if (activeTabIndex.value > 0) {
    props.controller.setActiveTabIndex(activeTabIndex.value - 1);
  }
}

function nextTab() {
  if (activeTabIndex.value < props.controller.tabs.length - 1) {
    props.controller.setActiveTabIndex(activeTabIndex.value + 1);
  }
}

const tabsStyle = computed<CSSProperties>(() => ({
  alignItems: "center",
  backgroundColor: props.isDark ? "#27272a" : "#f4f4f5",
  borderTop: `1px solid ${props.isDark ? "#3f3f46" : "#e4e4e7"}`,
  display: "flex",
  flexShrink: "0",
  gap: "2px",
  overflow: "hidden",
  padding: "0 4px",
}));

function tabStyle(index: number): CSSProperties {
  const isActive = index === activeTabIndex.value;
  return {
    backgroundColor: isActive
      ? props.isDark ? "#3f3f46" : "#ffffff"
      : "transparent",
    border: isActive
      ? `1px solid ${props.isDark ? "#52525b" : "#d4d4d8"}`
      : "1px solid transparent",
    borderBottom: "none",
    borderRadius: "4px 4px 0 0",
    color: isActive
      ? "inherit"
      : props.isDark ? "#a1a1aa" : "#71717a",
    cursor: "pointer",
    fontSize: "12px",
    maxWidth: "120px",
    padding: "4px 10px",
    whiteSpace: "nowrap",
  };
}

function renameInputStyle(index: number): CSSProperties {
  return {
    ...tabStyle(index),
    cursor: "text",
    fontFamily: "inherit",
    minWidth: "72px",
    outline: `2px solid ${props.isDark ? "#60a5fa" : "#2563eb"}`,
    outlineOffset: "-1px",
  };
}

const renameErrorStyle = computed<CSSProperties>(() => ({
  backgroundColor: "#b91c1c",
  borderRadius: "4px",
  color: "#ffffff",
  fontSize: "11px",
  left: `${renameAnchor.value?.left ?? 0}px`,
  maxWidth: "260px",
  padding: "4px 8px",
  position: "fixed",
  top: `${renameAnchor.value?.top ?? 0}px`,
  transform: "translateY(calc(-100% - 8px))",
  whiteSpace: "normal",
  zIndex: 9998,
}));

const tabMenuStyle = computed<CSSProperties>(() => ({
  backgroundColor: props.isDark ? "#27272a" : "#ffffff",
  border: `1px solid ${props.isDark ? "#3f3f46" : "#e4e4e7"}`,
  borderRadius: "8px",
  boxShadow: "0 4px 16px rgba(0,0,0,0.16)",
  left: `${tabMenu.value?.x ?? 0}px`,
  padding: "4px",
  position: "fixed",
  top: `${tabMenu.value?.y ?? 0}px`,
  zIndex: 9999,
}));

// A sheet added or removed under the editor must not leave it pointing at a
// different tab, and a rebuilt tab strip closes a stale context menu.
watch(renamingTabIndex, (index) => {
  if (index < 0 && renamingTabId.value !== null) closeRename();
});
watch(() => props.controller.tabs, () => closeTabMenu());
// Losing edit rights (the host flips readOnly mid-edit) must not leave a live
// editor that can only fail on commit.
watch(canRenameSheets, (allowed) => {
  if (!allowed) {
    closeRename();
    closeTabMenu();
  }
});

onUnmounted(() => {
  cancelScheduledPrefetch();
  if (popupTimer) clearTimeout(popupTimer);
});
</script>

<style scoped>
.xlsx-sheettabs {
  user-select: none;
}

.xlsx-sheettabs__list {
  display: flex;
  flex: 1;
  gap: 2px;
  overflow-x: auto;
}

.xlsx-sheettabs__tab-wrapper {
  position: relative;
  flex-shrink: 0;
}

.xlsx-sheettabs__tab {
  align-items: center;
  background: transparent;
  border: none;
  display: flex;
  font-family: inherit;
  gap: 4px;
  outline: none;
  overflow: hidden;
}

.xlsx-sheettabs__tab:hover {
  background: rgba(128, 128, 128, 0.1);
}

.xlsx-sheettabs__tab-name {
  overflow: hidden;
  text-overflow: ellipsis;
}

.xlsx-sheettabs__rename-input {
  box-sizing: border-box;
}

.xlsx-sheettabs__rename-error {
  background: #b91c1c;
  border-radius: 4px;
  color: #ffffff;
  font-size: 11px;
  max-width: 260px;
  padding: 4px 8px;
  white-space: normal;
  z-index: 9998;
}

.xlsx-sheettabs__menu-backdrop {
  inset: 0;
  position: fixed;
  z-index: 9998;
}

.xlsx-sheettabs__menu {
  min-width: 140px;
}

.xlsx-sheettabs__menu-item {
  background: transparent;
  border: none;
  border-radius: 4px;
  color: inherit;
  cursor: pointer;
  display: block;
  font-family: inherit;
  font-size: 12px;
  padding: 6px 10px;
  text-align: left;
  width: 100%;
}

.xlsx-sheettabs__menu-item:hover {
  background: rgba(128, 128, 128, 0.15);
}

.xlsx-sheettabs__thumbnail {
  display: flex;
  flex-direction: column;
  gap: 4px;
  max-width: 212px;
}

.xlsx-sheettabs__thumbnail canvas {
  border-radius: 6px;
  display: block;
  height: auto;
  width: 100%;
}

.xlsx-sheettabs__add-btn,
.xlsx-sheettabs__remove-btn {
  align-items: center;
  background: transparent;
  border: 1px solid transparent;
  border-radius: 4px;
  color: inherit;
  cursor: pointer;
  display: flex;
  font-family: inherit;
  font-size: 16px;
  height: 28px;
  justify-content: center;
  width: 28px;
}

.xlsx-sheettabs__add-btn:hover,
.xlsx-sheettabs__remove-btn:hover:not(:disabled) {
  background: rgba(128, 128, 128, 0.15);
}

.xlsx-sheettabs__remove-btn:disabled {
  cursor: default;
  opacity: 0.3;
}

.xlsx-sheettabs__nav-btn {
  align-items: center;
  background: transparent;
  border: none;
  border-radius: 4px;
  color: inherit;
  cursor: pointer;
  display: flex;
  flex-shrink: 0;
  font-size: 16px;
  height: 28px;
  justify-content: center;
  width: 24px;
}

.xlsx-sheettabs__nav-btn:hover:not(:disabled) {
  background: rgba(128, 128, 128, 0.15);
}

.xlsx-sheettabs__nav-btn:disabled {
  cursor: default;
  opacity: 0.3;
}

.xlsx-sheettabs__nav-btn svg,
.xlsx-sheettabs__add-btn svg,
.xlsx-sheettabs__remove-btn svg {
  fill: none;
  height: 16px;
  stroke: currentColor;
  stroke-linecap: round;
  stroke-linejoin: round;
  stroke-width: 1.8;
  width: 16px;
}

@media (pointer: coarse) {
  .xlsx-sheettabs button {
    min-height: 44px;
    min-width: 44px;
  }
}
</style>
