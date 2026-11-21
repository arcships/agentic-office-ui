# Extend UI ↔ `@arcships/vue-ui` 像素对照 harness

> 适用批次：0.8（#155）。固定基准：[`extend-hq/ui@70fe70c6`](https://github.com/extend-hq/ui/tree/70fe70c6f1e1eee53954c6923f3d9f583557a57e)（2026-11-17，MIT）。
>
> 本 harness 是所有 `[ui]` issue 的**唯一视觉判定依据**。不允许凭肉眼改样式：先跑 `pnpm ui:parity`，看报告里的几何/像素差值，再改代码。

## 1. 一条命令

```bash
pnpm ui:parity                 # 提取 token → 构建两个对照页 → 跑完整矩阵 → 超阈值时退出码 1
pnpm ui:parity:report          # 同上，但永远退出 0（PR artifact 用）
pnpm ui:parity:self-check      # 参考页自己跟自己比，验证采集/阈值链路是确定的（应全 PASS）
pnpm ui:tokens                 # 只重新生成 docs/design/extend-tokens.json
pnpm ui:tokens:check           # token 报告是否与 vendored Extend CSS 同步
```

常用筛选（都会传给 Playwright helper）：

```bash
pnpm ui:parity -- --only button,file-upload --viewports 1440x900 --themes light
pnpm ui:parity -- --no-build --only pdf-toolbar     # 复用上一次构建
UI_PARITY_OUT=output/ui-parity-local pnpm ui:parity
```

Python 解释器由 `scripts/ci/run-python.mjs` 解析（与其他浏览器套件一致）：设置 `PYTHON=/tmp/bb-venv/bin/python` 或把装了 `playwright` + `Pillow` 的 Python 放进 `PATH`。

## 2. 目录与职责

| 路径 | 作用 |
|---|---|
| `tools/parity-fixtures/fixtures.json` | 两侧共用：PDF/DOCX/XLSX/PPTX/PNG 文档、FileUpload 文案、PDF toolbar 数值、Citations / LayoutBlocks / Signature / Splits / FileSystem 的 block 数据 |
| `tools/parity-fixtures/manifest.json` | 矩阵：4 个 viewport × light/dark、每个 surface 的状态与交互、阈值、待办 issue |
| `tools/extend-reference/` | 最小 Vite + React 参考页；只 vendored 渲染所需的 Extend 组件子集（MIT 头 + `src/vendor/NOTICE.md`） |
| `tools/vue-reference/` | Vue 对照页；直接 import `packages/vue-ui` / `packages/vue-pdf` 源码，并加载 #156 的 `@arcships/vue-ui/theme.css` |
| `scripts/ci/pixel-parity.mjs` | 入口：token → 构建 → 调用 Python → 汇总退出码 |
| `scripts/ci/pixel-parity.py` | Playwright 采集 + 几何比较 + 像素 diff/热力图 + 报告 |
| `scripts/ci/extract-extend-tokens.mjs` | 从 vendored `globals.css` + 固定 Tailwind 默认值生成 token 报告 |
| `docs/design/extend-tokens.json` | 生成的 token 报告（见 [extend-tokens.md](extend-tokens.md)） |

参考页**不是** Extend 的 Next.js 站点。只 vendored 参与对照的组件；每个文件的改动都记在
`tools/extend-reference/src/vendor/NOTICE.md` 与 `docs/upstream/extend-ui.md`。固定 commit 与逐文件清单在
`tools/extend-reference/UPSTREAM.md`。

## 3. 两侧共用的 DOM 协议

```html
<div data-parity-surface="file-upload" data-parity-theme="light">
  <div data-parity-cell="default" data-parity-matrix-state="empty">
    <div data-parity-host><!-- 第一个元素子节点 = 被测目标 --></div>
  </div>
</div>
```

- 页面 URL：`/?surface=<id>&theme=light|dark`。
- 参考页在 surface 根节点加 `.dark`；Vue 页在 `<html>` 上设置 `data-office-theme="light|dark"`（#156 的作用域）。
- 目标选择：默认取 `[data-parity-host]` 的第一个元素子节点；`pdf-toolbar` 的 Vue 侧用
  `manifest.surfaces[].subjectTargetSelector = [data-testid='pdf-toolbar']` 指向真实 viewer 的 toolbar。
- 交互（`manifest.states[].interaction`）：`none` / `hover` / `focus` / `press`（按住鼠标）/ `dragenter`
  （派发 `DragEvent` + `DataTransfer`）/ `set-input-files`（Playwright 写入 fixture 文件）。
  选择器先匹配目标自身，再匹配后代，所以 `button` 既能命中 Button 的根节点，也能命中 toolbar 里的第一个按钮。
- 截图范围：目标 box 四周各 12px（`--capture-padding`），保证 focus ring / shadow 不被裁掉。

## 4. 判定规则

| 规则 | 值 | 说明 |
|---|---|---|
| 几何 | `0px` | box、padding、gap、border 宽度、四角 radius、font-size、line-height 逐项比较 |
| 像素（非文本） | 0 像素差异 | 完全相同 → `pass` |
| 文本抗锯齿 | 单通道 ≤ `2/255`，差异像素 ≤ `0.1%`，连续差异区域 ≤ `2×2` | 同时满足 → `pass-with-antialiasing` |
| 其他 | 任一条件不满足 → `fail` | 保留双方截图、diff、热力图 |
| 尺寸不一致 | 直接 `fail`（`sizeMismatch`） | 额外对左上公共区域做 diff，仅作诊断，不改变结论 |
| 参考/对照页 JS 报错 | `error` | 404 favicon 已过滤 |
| Expected diff | `waived` | 产品有意偏离上游时，manifest `expectedDiffs` 精确匹配 case，且实际像素差异 ≤ `maxDiffPixelRatio`；几何/尺寸/报错/基线失败不豁免 |

阈值来自 `tools/parity-fixtures/manifest.json` 的 `thresholds`，不要在脚本里硬编码另一套。
`expectedDiffs` 条目必须有 surface、state、themes、viewports、`maxDiffPixelRatio`、issue、reason 和
evidence；命中时 case 仍保留 reference/subject/diff 截图，并在报告顶部与 `SUMMARY.md` 的
`Expected diffs (waived)` 表里列出实际 ratio；未命中的声明会显示为 unused，避免豁免过期后继续“隐身”。

> 说明：当前实现把「文本抗锯齿容忍」统一应用于所有差异像素，因为 harness 还没有文本区域 mask；任何超过
> `2/255` 的差异都会 fail。更细的文本/非文本分区判定留给 #165。

## 5. 矩阵与状态

- viewport：`1440×900`、`1280×720`、`768×1024`、`390×844`（Playwright `device_scale_factor: 2`，Chromium headless）。
- 主题：`light`、`dark`。
- 状态（对应 issue 的 default/hover/focus/active/disabled/loading/empty/error）：

| Surface | 已定义状态 | Vue 对照 |
|---|---|---|
| `button` | default、hover、focus、active、disabled、loading + outline/secondary/ghost/destructive/link/size-sm/size-lg/icon-sm | ✅ `OfficeButton.vue`（104 pass + 8 `waived`，0 fail） |
| `input` | default、focus、disabled、invalid、size-sm | ✅ `OfficeInput.vue`（40/40 pass，#157b 已消掉 52px 残差） |
| `select` | default、focus、disabled、placeholder | ✅ `OfficeSelect.vue` |
| `dialog` | open（header/空 panel/footer/close） | ✅ `OfficeDialog.vue` |
| `tooltip` | open（hover 触发） | ✅ `OfficeTooltip.vue` |
| `file-upload` | default(empty)、hover、focus、dragging、rejected(error)、file-selected | ✅ `FileUpload.vue` |
| `file-thumbnail` | default(empty)、loading、error、image | ✅ `FileThumbnail.vue` |
| `pdf-toolbar` | default、hover、focus、active、disabled(empty document) | ✅ `PdfViewer.vue` 的真实 toolbar（40/40） |
| `pdf-sidebar` | open、closed、skeleton | ✅ `DocumentViewerSidebar` + `PdfThumbnailStrip`（24/24） |
| `docx-toolbar` | default、hover、focus、disabled(empty document) | ✅ `DocxViewerToolbar.vue`（32/32 pass） |
| `docx-sidebar` | open、closed、skeleton | ✅ `DocxThumbnailPanel.vue`（24/24 pass） |
| `document-sidebar` | open、closed、skeleton、overlay-open、overlay-closed | ✅ `DocumentViewerSidebar.vue` + `DocumentViewerSidebarSkeleton.vue`（40/40 pass） |
| `office-empty-state` | default、action | ✅ `OfficeEmptyState.vue`（16/16 pass） |
| `office-error-state` | default、message、wrapped | ✅ `OfficeErrorState.vue`（24/24 pass） |
| `office-loading-state` | default、hidden | ✅ `OfficeLoadingState.vue`（16/16 pass） |
| `office-status-bar` | default | ✅ `OfficeStatusBar.vue`（8/8 pass） |
| `xlsx-toolbar` | default、hover、focus、disabled(empty workbook) | ✅ `XlsxToolbar.vue`（32/32 pass） |
| `xlsx-sheet-tabs` | default、hover、focus | ✅ `XlsxSheetTabs.vue`（24/24 pass） |
| `xlsx-status-bar` | default | ✅ `XlsxStatusBar.vue`（8/8 pass） |
| `xlsx-ribbon` | default、hover、focus、disabled(no selection) | ✅ `XlsxRibbonSurface.vue`（OfficeToolbar/OfficeButton/OfficeToggle + HugeIcon；32/32 pass） |
| `pptx-toolbar` | default、hover、focus、active、disabled(empty presentation) | ✅ `PptxViewerToolbar.vue`（40/40 pass） |
| `pptx-sorter` | default、hover、selected | ✅ `PptxSlideCard.vue` + `PptxThumbnail.vue`（24/24 pass） |
| `pptx-presenter` | browse、present（声明） | — 上游无 presenter 布局；#183 定义 `office-reference`（规格 checklist + 自基线，见 [pptx-presenter-reference.md](pptx-presenter-reference.md)），manifest 切换在实现 issue |
| `bounding-box-citations` | default、hover、focus、selected、modified、empty、loading、error | ✅ `BoundingBoxCitations.vue`（Form 标量字段；232/232 pass） |
| `layout-blocks` | default、hover、focus、active、empty、loading、error | ✅ `LayoutBlocks.vue`（OCR 面板；232/232 pass） |
| `e-signature` | unsigned、signed、exporting、preview、loading、error | ✅ `ESignatureFields.vue`（字段面板；232/232 pass） |
| `document-splits` | default、hover、focus、single、empty-pages、empty、loading、error | ✅ `DocumentSplits.vue`（232/232 pass） |

截图前统一执行：`document.fonts.ready`、`prefers-reduced-motion: reduce`、注入
`animation:none / transition:none`、`caret-color:transparent`；每张截图前还会把所有 `<img>` 的
`loading` 改为 `eager`、等待 `decode()`（`settle_images`），避免视口外的 lazy 缩略图只在一侧加载造成
假差异。字体：Geist / Geist Mono（`@fontsource-variable/*`，与 `apps/v4/lib/fonts.ts` 同族）。

## 6. 产物

`output/ui-parity/`（已 gitignore）：

```
report.html                    # 人看：双方截图、diff、热力图、失败原因、top-N 差异区域、token diff
report.json                    # 机器读：cases[]、summary、gate 模式、baselines、tokens{...}
SUMMARY.md                     # PR 摘要：按 surface 的通过/失败计数、top-N 差异区域 + token 差异
cases/<surface>/<viewport>/<theme>/<state>/
  reference.png  subject.png  diff.png  heatmap.png  compare.png
  baseline-diff.png  baseline-heatmap.png     # 与提交基线不一致时
  overlap-diff.png  overlap-heatmap.png       # 仅尺寸不一致时
```

CI 与门禁（#165）：详见 [testing/ui-parity-ci.md](../testing/ui-parity-ci.md)。
样式迁移与 token 用法见 [guide/ui-style-migration.md](../guide/ui-style-migration.md)，
组件与 surface 的对应关系见 [components/ui-component-map.md](../components/ui-component-map.md)。

历史 artifact 名 `ui-parity-report` 在 #165 拆成 `ui-parity-core-<sha>`（PR 核心子集）与
`ui-parity-full-<sha>`（nightly 全矩阵），两者都包含同一份 `report.json` / `report.html` / `cases/**`。
`.github/workflows/ui-parity.yml` 现在有四个 job：`ui-contracts`（样式 guard + 基线审批）、
`ui-parity-core`（PR 核心子集 × 1440/390 + 基线比对）、`ui-parity-nightly`（全矩阵 × 4 viewport +
4 个 viewport 的基线比对）、
`ui-a11y`（axe-core + 键盘 smoke）。所有 job 上传 artifact（含 diff/heatmap/baseline diff）。
两个 parity job 现在都传 `--gate-mode all`：PR 子集与 nightly 全矩阵都严格阻断；
`gateStatus: pending` 只剩上游不可达的 `pptx-presenter`，它只进 report 的 unreachable 表。

## 6.1 提交基线（#165 / #174）

`tests/visual/extend/<surface>/<state>/<theme>-<viewport>.png` 是**对照侧**的版本化基线：
`ui-parity` 采集结束后额外与它比对（`--baseline`），因此即使参考页没变，组件改坏也会失败。
矩阵为全部 4 个 parity viewport —— 1440×900、1280×720、768×1024、390×844 × light + dark，
共 **864 张 / 14.0 MiB**（#174），只覆盖 `gateStatus: blocking` 的 surface。PR 的 `ui-parity-core`
只采集并比对 `coreSubset.viewports`（1440×900 + 390×844，快速门槛），nightly 全矩阵比对全部 4 个
viewport；缺基线或漂移都会在 `--gate-mode all` 下阻断。
更新必须走 `pnpm ui:baseline capture --approve --reviewer/--issue/--reason`，
`pnpm ui:baseline:verify` 在 CI 校验哈希与审批记录；细节见
[tests/visual/extend/README.md](../../tests/visual/extend/README.md)。

## 7. 实施现状（2026-11-21，0.8 wave 2）

- 参考侧全部非 fixtures-only surface 可渲染；Vue 侧 **25 个 surface** 完成端到端采集并提交基线
  （5 个 primitive + FileUpload / FileThumbnail / PdfToolbar+PdfSidebar / DOCX ×2 / PPTX ×2 /
  document-sidebar + 4 个 office-state / XLSX ×3 / document blocks ×4）。
- `pnpm ui:parity:self-check`（参考页对自身）在完整矩阵下 344/344 全部 `pass`
  （0 像素差、0 几何差），证明截图/阈值链路可重复；`pass` 不是永远无法到达的状态。
- #157 的 5 个 primitive（Button/Input/Select/Dialog/Tooltip）已接入；0.8 复跑
  `--only input` 在 4 viewport × light/dark 下 40/40 `pass`、0px、0 AA（#157b 修正了
  focus/light 的 52px 残差）。
- `file-upload` / `file-thumbnail` 已由 #163 收敛，两侧均使用**默认 props**——80 个 case
  全部通过（61 `pass` + 19 `pass-with-antialiasing`，0 `fail`；最大 diff ratio 0.000008、
  单通道 1/255、几何 0px）。这 19 个 AA 是 0.8 唯一保留的已知残差，低于 0.1% 阈值。
  实现细节见 [components/file-components.md](../components/file-components.md)。
- `pdf-toolbar` / `pdf-sidebar` 已由 #162 收敛为上游 48px toolbar 行与 160px rail：40/40、
  24/24 像素 0，并有提交基线。
- `docx-toolbar` / `docx-sidebar` 已由 #159 切到上游 `DocxToolbar` 与
  `DocumentViewerThumbnailSidebar` markup：32/32 + 24/24 像素 0；已提升为 `blocking` 并补基线。
- `xlsx-toolbar` / `xlsx-sheet-tabs` / `xlsx-status-bar` 已由 **#160（重开）** 建成：上游
  `xlsx-viewer.tsx` 里 `WorkbookToolbar` / `WorkbookSheetTabsInner` 的 markup 与 class 串是自包含的，
  vendored 子集（`tools/extend-reference/src/vendor/extend/xlsx-*.tsx`）只剔除私有引擎 hook 与弹层，
  class 串、图标、间距一字未改。`pnpm ui:parity -- --only xlsx-toolbar,xlsx-sheet-tabs,xlsx-status-bar`
  在 4 viewport × light/dark × 全部 state 上 **64/64 pass、0px、0 AA**；三条 surface 已补基线并切
  `blocking`。
- `xlsx-ribbon` 已由 **#170** 建成：`xlsx-editor.tsx` 的 `EditorToolbar` 格式行静态化后
  vendored 到 `tools/extend-reference/src/vendor/extend/xlsx-ribbon.tsx`（私有 hook → props，
  弹层不渲染），Vue 侧用 production primitives 复现；4 viewport × light/dark ×
  default/hover/focus/disabled = **32/32 pass、0px、0 AA**，已补基线并切 `blocking`。
  公式行仍是 product-gap；例外清单见
  [design/xlsx-office-chrome-migration.md §4.2](xlsx-office-chrome-migration.md)。
- `document-sidebar` 已由 #158 补齐对照侧：`DocumentViewerSidebar.vue` + `DocumentViewerSidebarSkeleton.vue`
  与 vendored `DocumentViewerThumbnailSidebar` 对齐，5 个状态 × 4 viewport × 2 theme = 40/40 `pass`，
  0 px、几何 0 px；已补基线并切 `blocking`。
- `office-empty-state` / `office-error-state` / `office-loading-state` / `office-status-bar` 是 #158
  新增的 state surface：对照分别是 `PdfEditorEmptyState`、viewer error card、`ViewerLoadingSurface`
  与上游 footer 类名组合，共 64/64 `pass`（0 px、几何 0 px）；已补基线并切 `blocking`。
- `bounding-box-citations` / `layout-blocks` / `e-signature` / `document-splits` 已由 #164 完成：
  232/232 case、0px、0 AA；四条 surface 已补基线并切 `blocking`。FileSystem / SchemaBuilder 按
  [components/document-blocks.md](../components/document-blocks.md) 记录为延期，不建 surface。
- **#174（2026-11-21）**：提交基线从 2 个 viewport 扩到全部 4 个（1440×900 / 1280×720 /
  768×1024 / 390×844 × light + dark）：432 → **864 张**（+7.0 MiB，共 14.0 MiB），一轮
  `pnpm ui:baseline capture` 写入 864 张、原有 432 张字节级不变（0 updated），本地 566 s。
  `pnpm ui:parity:full` 在 4 viewport 下 **864 case / 837 pass + 19 AA + 8 waived / 0 fail**，
  每个 viewport 216 次基线比对全部 `pass`，578 s；nightly job 不传 `--viewports`，因此
  1280/768 的基线确实会被比对（单元测试锁住）。PR job 仍只跑 1440/390
  （`coreSubset.viewports`），保持分钟级反馈。
- `pptx-toolbar` / `pptx-sorter` 已收敛到 vendored 上游标记：40 + 24 个 case
  全部 pass，0 像素差、0 几何差；#165 在合并后的 0.8 树上复跑 64/64 后补基线并切
  `blocking`，两条都进 PR 子集（摘要见 [pptx-chrome.md](../components/pptx-chrome.md#7-测试与像素证据)）。
  `pptx-presenter` 在固定 commit 没有上游布局（presenter 在私有
  `@extend-ai/react-pptx` 内）。#183 已把它的参考改成 `office-reference`：
  [pptx-presenter-reference.md](pptx-presenter-reference.md) 定义区域/尺寸/交互与
  checklist，验收靠规格 checklist + 自基线 + 组件测试；manifest 字段与 runner
  分支在实现 issue 中接入，在那之前仍按 `unreachable` 跳过采集。
- 环境对齐：Extend 参考页的 `.dark` 只切 token，而 #156 的 `theme.css` 同时设置 `color-scheme`。
  `ParitySurface` 现在按主题写入 `color-scheme`，避免 focus ring 因为原生控件配色不同而误报。

- token 层：27 个标量 token **27/27 一致**，46 个语义色 **46/46 一致**（0 missing）。
  `--office-control-padding-inline` 按“完整内缩”比较（Extend 11px padding + 1px border = 12px，
  Office token = 12px，消费方在自绘描边时减去 `--office-border-width`）；4 个 `--office-code*`
  token 已补齐。`--office-card` / `--office-popover` 浅色改为 literal `white`（上游 `var(--color-white)`），
  避开 `oklch(1 0 0)` 在 Chromium 中的 1/255 合成漂移。

## 8. 新增一个 surface

1. 在 `tools/extend-reference/src/surfaces/` 加参考渲染，并在 `App.tsx` 的 `renderers` 注册；需要新组件时先把最小上游文件 vendored 到 `src/vendor/`（MIT 头 + NOTICE + `UPSTREAM.md` + `attribution.json`）。
2. 在 `tools/vue-reference/src/surfaces/` 加对照渲染，并在 `App.vue` 里接上；如果组件还没实现，把 manifest 标成 `subjectStatus: "missing"` + `pendingIssue`。
3. 在 `tools/parity-fixtures/manifest.json` 增加 surface（`cellWidth`、`states`、交互、依赖 issue），共享文案/数据放 `fixtures.json`；同时声明 `core`（是否进 PR 子集）与 `gateStatus`（`pending` + `gatePendingIssue`，或 `blocking`）。
4. 在 [components/ui-component-map.md](../components/ui-component-map.md) 用实际的 surface id（不是组件名）补一行，`pnpm ui:docs` 会检查。
5. 运行 `pnpm ui:parity -- --only <id>` 和 `node --test tests/unit/ui-parity-gate.test.mjs`；surface 就绪后按
   [ui-parity-ci.md §5](../testing/ui-parity-ci.md) 切成 `blocking` 并用 `pnpm ui:baseline capture --only <id>` 补基线。

## 9. 后续 issue 分工

- #156：token/theme/OfficeIcon（已合入；本 harness 直接消费 `theme.css`）。
- #157：Button/Input/Select/Dialog/Tooltip primitive + reka-ui 层（已合入并 `blocking`，像素对照见 §7）。
- #158：共享 viewer chrome + 规范 `DocumentViewerSidebar`（`components/chrome/`）；
  `document-sidebar` 与 4 个 office-state surface 已合入、补基线并切 `blocking`。
- #159：DOCX toolbar/sidebar（已合入，32/32 + 24/24 像素 0）。
- #160：XLSX toolbar/sheet tabs/status（已合入，64/64 像素 0）；#170 补 `xlsx-ribbon`（32/32 像素 0）。
- #161b：PPTX toolbar/sorter/presenter（已合入；toolbar/sorter 64/64 像素 0，已补基线并切
  `blocking`；presenter 无上游布局，保持 `unreachable`）。
- #162：PDF toolbar/sidebar（已合入，40/40 + 24/24 像素 0）。
- #163：FileUpload / FileThumbnail 已收敛；DocumentViewerSidebar 的组件与 demo 由 #158 负责。
- #164：document blocks（已合入，232/232 像素 0；FileSystem/SchemaBuilder 延期）。
- #165：gate、基线、stylelint/axe、文档——25 个可渲染 surface 已全部 `blocking`，
  `prMode`/workflow 已切到 `all`（见 [testing/ui-parity-ci.md §5](../testing/ui-parity-ci.md)）。
