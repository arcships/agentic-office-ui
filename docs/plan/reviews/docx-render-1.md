# docx-render 开发产出 Review #1

Date: 2026-07-08
Reviewer: automated review
Status: **blocked**

## 总评

docx-core（框架无关引擎 + 布局 + 编辑逻辑）基本对齐架构文档，typecheck 通过，import 路径正确。vue-docx 的 **components 层完全缺失**（架构文档定义 18 个组件文件，实际 0 个），index.ts 内用 stub 组件占位。这是阻塞性问题。

---

## 一、架构文档文件清单对齐

### 1.1 docx-core: 基本覆盖

| 架构层 | 架构预期文件数 | 实际文件数 | 状态 |
|---|---|---|---|
| engine/ | 8 | 8 | ✅ |
| layout/ | 5 | 5 | ✅ |
| viewer/ | 15 | 16 | ✅（多出 pretext-items-layout.ts，为 pretext-layout.ts 的合理拆分） |
| canvas/ | 2 | 2 | ✅ |
| editor/editor-ops.ts | 1 | 1 | ✅ |
| editor/helpers/ | 39 | 41 | ✅（多出 paragraph-toc.ts，为打破循环依赖引入的合理拆分） |
| **小计** | **~70** | **~74** | |

### 1.2 vue-docx: 组件层缺失

| 架构层 | 架构预期文件数 | 实际文件数 | 状态 |
|---|---|---|---|
| composables/ | 26 | 26 | ✅ |
| render/ | 6 | 7 | ✅（多出 paragraph-runs-field-resolve.ts，合理拆分） |
| **components/** | **18** | **0** | ❌ **完全缺失** |

**components/ 预期 18 个文件（架构文档 §3.2）：**

- DocxViewer.vue
- DocxEditor.vue
- DocxViewerRoot.vue
- DocxPageWrapper.vue
- DocxPageSurface.vue
- DocxPageHeader.vue
- DocxPageFooter.vue
- DocxPageBody.ts（render function）
- DocxParagraphHost.vue
- DocxTableHost.vue
- DocxImageLayer.vue
- DocxFormFieldLayer.vue
- DocxTrackedChangeGutter.vue
- DocxContextMenu.vue
- DocxToolbar.vue
- DocxThumbnailPanel.vue
- DocxDragOverlay.vue
- index.ts

**实际产物：0 个文件。** `packages/vue-docx/src/index.ts:4-17` 内是两个 stub 组件：

```typescript
// "-- Stub components (full implementations land in components/) --"
export const DocxViewer = defineComponent({
  setup() { return () => h("div", { class: "docx-viewer-stub" }, "DOCX Viewer (pending)") },
})
export const DocxEditorViewer = defineComponent({
  setup() { return () => h("div", { class: "docx-editor-stub" }, "DOCX Editor (pending)") },
})
```

demo 页面 `DocxEditorPage.vue` 导入 `DocxEditorViewer` 后实际渲染的是 "DOCX Editor (pending)" 文本，不可用于任何功能验证。

---

## 二、stub/mock/fake 残留检查

### 2.1 阻塞性 stub

**`packages/docx-core/src/editor/helpers/line-height-table.ts:63-100`**

三个运行时注入式 stub，缺少时会抛出 `"not injected"` 错误：

| stub 函数 | 实际来源 | 状态 |
|---|---|---|
| `estimateParagraphHeightPx` | 应在 line-height.ts，但 line-height.ts 未导出该函数 | ❌ stub |
| `paragraphAvailableTextWidthPx` | 应来自 paragraph-geometry.ts，但未导出 | ❌ stub |
| `paragraphLineCountWithinWidth` | 应来自 line-height.ts，但未导出 | ❌ stub |

```typescript
// line-height-table.ts:63-100
// -- Missing dependency stubs (replace with real imports when modules land) --
function makeInjectable<Args extends unknown[], R>(name: string) {
  let _fn: Fn<Args, R> | undefined;
  return {
    inject(fn: Fn<Args, R>): void { _fn = fn; },
    call(...args: Args): R {
      if (!_fn) throw new Error(`${name} not injected`);  // 运行时崩溃
      return _fn(...args);
    },
  };
}
```

这三个函数在 `line-height-table.ts` 的 `estimateTableCellContentHeightPx`（L162-L250）中被实际调用。无注入时所有表格单元格高度估算都会抛出异常，导致分页引擎不可用。架构文档未将 `line-height-table.ts` 标注为待实现。

### 2.2 Scoped/shell 状态（不在架构文档待实现清单内）

以下是 demo 页面 `DocxEditorPage.vue` 自行声明的 scoped/shell 状态，均不是架构文档明确标注的"待实现"项：

| composable | 自述状态 | 实际表现 |
|---|---|---|
| `useDocxTrackChanges` | "API shell" | 无可视化修订、无记录生成 |
| `useDocxComments` | "API shell" | 无可视化批注、无交互 |
| `useDocxBorders` | "API shell" | 无边框编辑功能 |
| `useDocxParagraphStyles` | "API shell" | 仅 heading dropdown |
| `useDocxLineSpacing` | "API shell" | 无可见行距控制 |
| `useDocxViewerThumbnails` | "⚠️ PLACEHOLDER" | 无 canvas 渲染 |

### 2.3 无 stub 残留（干净）

以下领域无额外 stub/mock/fake：
- engine/ 全部文件 ✅
- layout/ 全部文件 ✅
- viewer/ 全部文件 ✅
- canvas/ 全部文件 ✅
- editor/editor-ops.ts ✅
- editor/helpers/（除 line-height-table.ts 的 3 个 injectable stub 外）✅
- vue-docx render/ ✅
- vue-docx composables/（数据结构层，逻辑因无组件无法验证）⚠️

---

## 三、import 路径检查

### 3.1 docx-core: 全相对路径 ✅

- 文件内跨目录引用使用相对路径（`../../engine/types`、`../../viewer/section-layout`）
- 无 `React.` 类型引用残留（仅有注释提及替换历史）
- 无 `@extend-ai/react-docx-*` 引用残留

### 3.2 vue-docx: 跨包引用正确 ✅

- 跨包 import 使用 `@arcships/docx-core`（workspace 协议）
- composables/ 内文件间引用用相对路径

---

## 四、typecheck 结果

| 包 | 命令 | 结果 |
|---|---|---|
| `@arcships/docx-core` | `tsc --noEmit` | ✅ 通过 |
| `@arcships/vue-docx` | `tsc --noEmit` | ✅ 通过 |

---

## 五、上游功能对齐检查

### 5.1 引擎层（§2.1）：全覆盖 ✅

全部 8 个对齐要点均已实现。`types.ts` 438 行完整 DocModel 类型体系、手工深拷贝、JSON 归一化、无状态 wasm 封装、basePackage 机制。

### 5.2 布局/分页（§2.2）：基本覆盖，缺运行时注入

架构文档 15 个布局对齐要点中，分页引擎核心算法已实现（page-segmentation-core.ts 864行 + page-segmentation-table.ts 376行 + pagination.ts 690行），但 **line-height-table.ts 的 3 个 injectable stub 导致表格分页运行时不可用**。

### 5.3 编辑（§2.3）：composables 实现，无组件验证

dispatchEditorTransaction 事务分发器、快照式历史、pendingRunStyle 机制、contentEditable draft 缓存均在 composables/ 中实现。但因无 Vue 组件层，无法验证 contentEditable 交互和 DOM 选区恢复的正确性。

### 5.4 渲染（§2.4）：全部阻塞于组件缺失

架构文档 9 个渲染对齐要点全部依赖 Vue 组件层——虚拟化（@tanstack/vue-virtual）、每页 contain 隔离、DOM 渲染、浮动图片定位、表格 resize、缩略图三级策略。组件层为空，所有渲染功能不可用。

### 5.5 26 项功能：仅 composable API 存在，无可见 UI

上游 26 项功能的 UI 层（toolbar、右键菜单、缩略图面板、修订 gutter）全部缺失。

---

## 六、行数统计

| 层 | 架构预期 ~行数 | 实际行数 | 差异 |
|---|---|---|---|
| docx-core engine | 1700 | 1678 | 基本吻合 |
| docx-core layout | 2653 | 2608 | 基本吻合 |
| docx-core viewer | 4713 | 4588 | 基本吻合 |
| docx-core canvas | 200 | 200 | 吻合 |
| docx-core editor/helpers | 21200 | ~16600 | 行数偏低（部分文件行数不达预期） |
| docx-core editor(ops) | 1329+1556 | 1329+? | editor-ops 吻合 |
| vue-docx composables | 8700 | ~5430 | 偏低 |
| vue-docx components | 10900 | 0 | **完全缺失** |
| vue-docx render | 1505 | ~1505 | 基本吻合 |
| **合计** | **~53127** | **~33938** | **缺 ~19200 行（组件层 + helpers 缩减）** |

---

## 七、问题分级

### 阻塞（blocked）

| # | 问题 | 影响范围 |
|---|---|---|
| 1 | **components/ 目录完全缺失**（18 个文件，~10900 行） | 所有可视化功能不可用：viewer 渲染、编辑器交互、toolbar、缩略图、右键菜单 |
| 2 | **line-height-table.ts 3 个运行时 injectable stub** | 表格分页引擎运行时抛异常，影响分页和表格渲染 |

### 严重（需要修复但不阻塞基础验证）

| # | 问题 | 影响范围 |
|---|---|---|
| 3 | 5 个 composable 为 API shell 状态 | 修订/批注/边框/段落样式/行距仅有 toggle 开关，无实际功能 |
| 4 | `useDocxViewerThumbnails` 为 PLACEHOLDER | 缩略图功能不可用 |
| 5 | `selection-restore.ts` 135 行（预期 ~605 行） | DOM 选区恢复逻辑可能不完整 |
| 6 | `paragraph-inspect.ts` 452 行（预期 ~800 行） | 段落属性提取可能不完整 |

### 信息

| # | 问题 |
|---|---|
| 7 | 多出 3 个合理文件：paragraph-toc.ts（打破循环依赖）、paragraph-runs-field-resolve.ts（render 拆分）、editor-shared.ts（composable 类型共享） |
| 8 | 多出 composables.ts 为 barrel re-export 兼容文件，无功能影响 |

---

## 八、结论

**status: blocked**

阻塞原因：
1. vue-docx 组件层 18 个文件零产出，所有可视化交互不可验证
2. docx-core line-height-table.ts 的 3 个 injectable stub 导致表格分页运行时异常

下一步：
- 优先实现 components/ 目录的 18 个 Vue 组件（至少完成 DocxViewer.vue、DocxEditor.vue、DocxViewerRoot.vue、DocxPageSurface.vue、DocxParagraphHost.vue、DocxToolbar.vue 6 个核心组件）
- 补全 line-height-table.ts 的 3 个函数实现（estimateParagraphHeightPx、paragraphAvailableTextWidthPx、paragraphLineCountWithinWidth）
- 补全 selection-restore.ts 和 paragraph-inspect.ts 的缺失行数
