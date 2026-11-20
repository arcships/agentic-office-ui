# FileUpload / FileThumbnail：Extend 对齐合同（0.8 · #163）

> 参考基准：`extend-hq/ui@70fe70c6f1e1eee53954c6923f3d9f583557a57e`（MIT）。
> 视觉判定以 #155 harness 的 `pnpm ui:parity` 报告为准；本文记录 API、状态、测试和本轮实测结果。
>
> **0.8 视觉默认变更**：`FileUpload` 的 `showFileList` 默认改为 `true`，并同步 Extend 的
> 默认 `accept`、`multiple`、标题/描述文案；默认渲染即与 Extend 参考逐像素对齐。
> `FileThumbnail` 在 Extend 中本就没有 glyph/meta 层，因此 `showTypeGlyph` / `showMeta`
> 保持 `false`（显式开启是 Vue 侧扩展能力），否则默认态会多出参考不存在的像素。

## 1. 范围

- `FileUpload`：标题/描述、拖拽深度、accept 类型 glyph cluster、browse/drop chip、组件内文件列表、大小/类型、Ready 徽标、拒绝提示、可选 BorderBeam。
- `FileThumbnail`：方形或自定义 `aspect-ratio` 预览卡、图片 blur-in、loading overlay、error/empty fallback、可选文件类型 glyph、可选文件名/类型 meta。
- `DocumentViewerSidebar`：**不在本 issue 重复实现**。规范组件由 #158 在
  `packages/vue-ui/src/components/chrome/` 交付；本文第 7 节记录集成 seam。

## 2. FileUpload

### 2.1 Props

| 属性 | 类型 | 默认 | 说明 |
|---|---|---|---|
| `accept` | `string` | Extend 默认文档/图片列表 | 逗号分隔的 MIME / 扩展名 / `type/*`，大小写不敏感；默认值与参考 `DEFAULT_ACCEPT` 一致 |
| `multiple` | `boolean` | `true` | 传给原生 input；drop 仍会逐个校验 |
| `maxSize` | `number` | `Infinity` | 单文件字节上限 |
| `disabled` | `boolean` | `false` | 禁用点击、键盘、拖拽，并应用 disabled token |
| `className` | `string` | `""` | 追加到根节点 |
| `title` | `string` | 本地化 upload title | 主标题；默认与 Extend 的 “Click to upload or drop files” 一致 |
| `description` | `string` | 本地化 upload description | 副标题；默认与 Extend 的 “PDF, DOC/DOCX, XLSX, CSV, PNG, or JPG” 一致，传 `""` 可隐藏 |
| `browseLabel` | `string` | 本地化 “Browse files” | chip 空闲文案 |
| `draggingLabel` | `string` | 本地化 “Drop to add” | chip 拖拽文案 |
| `acceptedFileTypes` | `FileUploadAcceptedType[]` | Image / PDF / Sheet | cluster glyph；`icon` 支持内置 `"image"`/`"pdf"`/`"sheet"` 或任意 Vue 组件 |
| `showFileList` | `boolean` | `true` | 在 dropzone 下方渲染组件内文件列表（Extend 默认）；旧宿主需要自己渲染列表时传 `false` |
| `showBorderBeam` | `boolean` | `false` | 拖拽时显示可选 CSS beam；默认关闭以保持宿主旧视觉 |
| `borderBeamTheme` | `"light" \| "dark"` | `"light"` | beam 调色提示 |

### 2.2 事件与兼容

| 事件 | 参数 | 说明 |
|---|---|---|
| `files-accepted` | `File[]` | 校验通过的文件；混合批次只包含通过项 |
| `files-rejected` | `FileUploadRejection[]` | 每项含稳定 `code`（`FILE_TYPE_NOT_ACCEPTED` / `FILE_TOO_LARGE`）、原始 `file` 和 `message` |

- 旧 props（`accept`、`multiple`、`maxSize`、`disabled`、`className`）保留；没有重命名或删除。
- **0.8 默认变更**：`accept`、`multiple` 和 `showFileList` 的默认值改为 Extend 参考值；依赖旧默认的宿主可显式传 `accept=""`、`:multiple="false"`、`:show-file-list="false"`。
- 旧 `#label` slot 不再渲染；需要自定义标题时使用 `title` prop。
- 仅类型拒绝且没有通过项时，界面显示与 Extend 一致的
  “This file type is not supported here.”；`files-rejected[].message` 仍保留可诊断的详细文案。尺寸拒绝和混合批次继续显示详细原因。
- 文件列表默认渲染，列表项按 Extend 的顺序显示 `type - size`，右侧为 Ready；长文件名使用 `min-width:0` + `text-overflow: ellipsis`。
- 选中的文件被替换时，旧的图片 object URL 会 `revokeObjectURL`；组件卸载时同样释放。

### 2.3 视觉与交互

- dropzone：`16rem` 最小高度、`18px` 圆角、`1px dashed` 边框、`24px/40px` padding、`20px` gap；颜色全部来自 `--office-*`。
- 状态：默认 / hover / dragging / disabled 使用 token 组合；dark 下边框 alpha 从 20% 调到 25%，hover/drag 40%。
- focus 使用浏览器 `outline: auto`，`outline-color` 绑定
  `color-mix(in oklab, var(--office-ring) 50%, transparent)`，与 Extend 的全局 `outline-ring/50` 一致。
- `(pointer: coarse)` 下 dropzone、chip 和文件行最小命中高度使用 `--office-ui-touch-target`（44px）。
- icon cluster：三张 `48px` 卡片按 Extend 的 `translate/rotate` 变换；拖拽时抬升、切换 popover 背景和 `--office-shadow-md`。
- BorderBeam：纯 CSS conic-gradient + mask，不引入 `border-beam` 运行时依赖；只在拖拽时绘制。

## 3. FileThumbnail

### 3.1 Props

| 属性 | 类型 | 默认 | 说明 |
|---|---|---|---|
| `file` | `{ name, type, url?, size? }` | 必填 | 文件名、MIME、旧版图片 URL、可选大小 |
| `size` | `"sm" \| "md" \| "lg" \| number \| string` | `"md"` | `sm/md/lg` 保持 32/48/64px 旧语义；number 为像素，string 原样作为 CSS 长度 |
| `className` | `string` | `""` | 追加到根节点 |
| `previewImageUrl` | `string \| null` | — | 显式预览图；未传时对 `image/*` 回退到 `file.url`，传 `null` 可强制 fallback |
| `previewAspectRatio` | `number` | — | 例如 `16 / 9`；省略时为方形 |
| `previewClassName` | `string` | `""` | 追加到 preview 层 |
| `isLoading` | `boolean` | `false` | 强制 loading overlay |
| `hasError` | `boolean` | `false` | 强制 error fallback |
| `showTypeGlyph` | `boolean` | `false` | fallback 中渲染文件类型 glyph（Vue 扩展）；Extend 的默认 fallback 是纯 muted 色块，开启会破坏默认像素对齐 |
| `showMeta` | `boolean` | `false` | 在预览下渲染文件名/类型行（Vue 扩展，旧版 inline row 的等价能力） |
| `alt` | `string` | 文件名 | 预览图 alt |

### 3.2 渲染层

1. 根卡：`1px` 边框、`--office-radius-lg` 圆角、background 背景、`overflow: hidden`。
2. preview：`aspect-ratio` 容器、muted 背景、`contain: layout paint`。
3. 图片：`object-fit: cover`、160ms opacity/filter 过渡；首次加载完成后用两帧 rAF 移除 overlay；同一 URL 在本会话内再次挂载时直接显示。
4. loading overlay：muted 底 + `background/55` pulse；`prefers-reduced-motion` 下关闭动画。
5. fallback：muted 底；`showTypeGlyph` 开启时按 MIME/扩展名映射 Lucide glyph。
6. `showMeta`：文件名和 `type || unknownType`，均 `nowrap + ellipsis`。

## 4. i18n

`@arcships/office-interaction` 的 `vueUi.fileUpload` 新增：
`title`、`description`、`browse`、`drop`、`ready`、`rejected`、`unknownType`；`prompt` 保留。en / zh-CN 成对维护，
`tests/unit/i18n.test.mjs` 检查 key 完整。宿主仍可用 `messages` 覆盖。

## 5. Extend 像素对照

命令（本机使用 `/tmp/bb-venv` 的 Playwright/Pillow）：

```bash
PYTHON=/tmp/bb-venv/bin/python pnpm ui:parity -- \
  --only file-upload,file-thumbnail --skip-tokens
```

本轮结果（2026-11-20，Chromium 149，DPR 2，两侧均使用**默认 props**）：

| Surface | 状态数 × viewport × theme | 结果 |
|---|---|---|
| `file-upload` | 6 × 4 × 2 = 48 | 全部通过（含文本抗锯齿容差） |
| `file-thumbnail` | 4 × 4 × 2 = 32 | 全部通过 |
| 合计 | 80 | 61 `pass` + 19 `pass-with-antialiasing`，0 `fail` |

最大差异：`diffPixelRatio` 0.000008（0.0008%）、单通道 1/255、连续差异 1px、几何 delta 0px；
全部低于 manifest 的 0.1% / 2/255 / 2×2 门限。`FileUpload` 和 `FileThumbnail` 的 Vue/参考
surface 都不再传 `show-file-list` 等 opt-in prop，直接测量组件默认值。报告：
`output/ui-parity-defaults/report.json`（`output/` 为本地产物，不入库）。通过标准来自
`tools/parity-fixtures/manifest.json`：几何 0px；单通道 ≤2/255、差异像素 ≤0.1%、连续差异 ≤2×2。

反证：把 `FileThumbnail.showTypeGlyph` 临时改成默认 `true`，同一参考在 `default`/`error` 两个
状态立即失败（`diffRatio` 3.61%、单通道 132/255、连续 8px 且出现 3×3 块），因为 Extend 的
`file-thumbnail.tsx` 默认 fallback 只渲染 muted 色块。因此该 prop 必须保持 opt-in。

对照过程中还修正了 harness 的一处环境差异：Extend 参考页用 `.dark` 只切 token，而 #156 的
`theme.css` 会同时设置 `color-scheme`。`ParitySurface` 现在按主题设置 `color-scheme`，两侧在
相同的原生控件配色下测量 focus ring。

## 6. Story / 示例页

- demo 路由：`/#/file-components`（`apps/demo/src/pages/FileComponentsPage.vue`）。
- 状态：default（默认开启文件列表）、drag（可 Start/End drag）、file list + Ready、rejected、host-managed list opt-out、empty fallback、image blur-in、loading、error + glyph、16:9、长文件名。
- light/dark 由页面内 `OfficeThemeProvider` 切换；390 / 768 / 1440 由真实 viewport 截图验证（模板使用 `minmax(min(…, 100%), 1fr)` 自适应）。
- 组件验证页 `/#/components` 继续展示 FileUpload/FileThumbnail 的基本用法。
- 本地截图 helper：`python3 scripts/ci/capture-file-components-gallery.py`（先以 `timeout` 启动 demo 预览），
  输出 `output/parity/file-components/<viewport>-<theme>-*.png`；脚本在 console error 时失败。

## 7. DocumentViewerSidebar 集成 seam

规范实现属于 #158：`packages/vue-ui/src/components/chrome/DocumentViewerSidebar.vue` 与
`DocumentViewerSidebarSkeleton.vue`，导出 `DocumentViewerSidebar`（及 skeleton / `useElementWidth`）。
本 issue 不创建同名组件，也不在 manifest 中把它标成 available。

预期组合方式：

- viewer 负责 `open` / `mode` / `container` 与 `update:open`；sidebar 负责宽度测量、inline/overlay 切换和开合动画。
- sidebar 的默认 slot 放 viewer 自己的缩略图条；每个缩略图使用
  `FileThumbnail`：`file` 传页面名与 MIME，`previewImageUrl` 传页位图，`size` 用数字宽度（例如 80），
  `previewAspectRatio` 按页面宽高，长页面名交给 `showMeta` 的截断规则。
- loading 时使用 `DocumentViewerSidebarSkeleton`（#158）或把 `loading` 传给 sidebar；不要用 FileThumbnail
  的 loading overlay 冒充 skeleton。
- #163 不修改 sidebar manifest 的 `pendingIssue` 之外的语义；#158 合入后，由 viewer issue（#159–#162）
  接上 demo 和 `document-sidebar` parity surface。

## 8. 测试

```bash
pnpm --filter @arcships/vue-ui build
node --test --test-concurrency=1 tests/component/file-components.test.mjs
node --test --test-concurrency=1 tests/component/vue-components.test.mjs
node --test --test-concurrency=1 tests/component/file-drop-import.test.mjs
node --test --test-concurrency=1 tests/component/touch-interactions.test.mjs
node --test --test-concurrency=1 tests/component/i18n-source.test.mjs
node --test --test-concurrency=1 tests/unit/i18n.test.mjs
node --test --test-concurrency=1 tests/unit/extend-parity-harness.test.mjs
```

`file-components.test.mjs` 覆盖：默认 props（标题/描述/accept/multiple/file list）、`showFileList` opt-out、
拖拽深度、选择/拖放/多文件、拒绝、maxSize、disabled、长文件名、
coarse-pointer 命中区、loading/error/fallback/type glyph、图片 reveal 与失败恢复、可配置 aspect ratio。
`file-drop-import.test.mjs` 与 `touch-interactions.test.mjs` 仍覆盖四个 viewer 的共享拖放/触控合同。

## 9. 已知边界

- BorderBeam 是 CSS 近似，不是 `border-beam` 包的逐像素复制；默认关闭，开启后只在拖拽态绘制。
- FileUpload 的默认 glyph 来自 `@hugeicons/core-free-icons@4.4.0`（MIT，构建 external，不内联到 `dist/index.js`），以复用 Extend 参考的同一套 icon data；宿主无需额外配置，依赖随 `@arcships/vue-ui` 安装。
- `showFileList` 默认 `true`，与 Extend 对齐；旧宿主可传 `:show-file-list="false"` 保留自己的列表布局。
- `showTypeGlyph` 默认 `false` 不是兼容性遗留，而是参考事实：Extend 的 `FileThumbnail` 没有 glyph/meta 层，默认态是纯 muted 色块；需要类型 glyph 时显式开启（Vue 扩展）。
- `DocumentViewerSidebar` 的组件、demo 和 `document-sidebar` 像素对照仍由 #158 完成；本 issue 只提供 seam 文档。
- i18n source 测试在本分支上仍会报告 #157 遗留的
  `packages/vue-ui/src/components/ui/labels.ts` 中文字面量；该问题不属于 #163，合并时需由 #157 跟进。
