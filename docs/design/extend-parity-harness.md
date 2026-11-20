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

阈值来自 `tools/parity-fixtures/manifest.json` 的 `thresholds`，不要在脚本里硬编码另一套。

> 说明：当前实现把「文本抗锯齿容忍」统一应用于所有差异像素，因为 harness 还没有文本区域 mask；任何超过
> `2/255` 的差异都会 fail。更细的文本/非文本分区判定留给 #165。

## 5. 矩阵与状态

- viewport：`1440×900`、`1280×720`、`768×1024`、`390×844`（Playwright `device_scale_factor: 2`，Chromium headless）。
- 主题：`light`、`dark`。
- 状态（对应 issue 的 default/hover/focus/active/disabled/loading/empty/error）：

| Surface | 已定义状态 | Vue 对照 |
|---|---|---|
| `button` | default、hover、focus、active、disabled、loading + outline/secondary/ghost/destructive/link/size-sm/size-lg/icon-sm | ✅ `OfficeButton.vue`（196/200 矩阵中全部 pass） |
| `input` | default、focus、disabled、invalid、size-sm | ✅ `OfficeInput.vue`（focus/light 有 52 px / 1-255 抗锯齿残差） |
| `select` | default、focus、disabled、placeholder | ✅ `OfficeSelect.vue` |
| `dialog` | open（header/空 panel/footer/close） | ✅ `OfficeDialog.vue` |
| `tooltip` | open（hover 触发） | ✅ `OfficeTooltip.vue` |
| `file-upload` | default(empty)、hover、focus、dragging、rejected(error)、file-selected | ✅ `FileUpload.vue` |
| `file-thumbnail` | default(empty)、loading、error、image | ✅ `FileThumbnail.vue` |
| `pdf-toolbar` | default、hover、focus、active、disabled(empty document) | ✅ `PdfViewer.vue` 的真实 toolbar |
| `document-sidebar` | open、closed、skeleton(loading) | ❌ 无侧栏（#158）→ 仅参考基线；#163 记录 seam |
| `document-blocks` | 只有 fixtures（渲染对照由 #164 接管） | — |

截图前统一执行：`document.fonts.ready`、`prefers-reduced-motion: reduce`、注入
`animation:none / transition:none`、`caret-color:transparent`。字体：Geist / Geist Mono
（`@fontsource-variable/*`，与 `apps/v4/lib/fonts.ts` 同族）。

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
`ui-parity-core`（PR 核心子集 + 基线比对，`--gate-mode ready`）、`ui-parity-nightly`（全矩阵）、
`ui-a11y`（axe-core + 键盘 smoke）。所有 job 上传 artifact（含 diff/heatmap/baseline diff）。
`--gate-mode ready` 只判定 `gateStatus: blocking` 的 surface；pending surface（`input`/`docx-*`/
`document-sidebar`/`document-blocks`）仍采集、仍报告，但不阻断。最终切换用 `strict: true`
（`--gate-mode all`），正式切换步骤见门禁文档 §5。

## 6.1 提交基线（#165）

`tests/visual/extend/<surface>/<state>/<theme>-<viewport>.png` 是**对照侧**的版本化基线：
`ui-parity` 采集结束后额外与它比对（`--baseline`），因此即使参考页没变，组件改坏也会失败。
矩阵为 1440×900 + 390×844 × light + dark，只覆盖 `gateStatus: blocking` 的 surface。
更新必须走 `pnpm ui:baseline capture --approve --reviewer/--issue/--reason`，
`pnpm ui:baseline:verify` 在 CI 校验哈希与审批记录；细节见
[tests/visual/extend/README.md](../../tests/visual/extend/README.md)。

## 7. 实施现状（2026-11-20，0.8 wave 2）

- 参考侧全部非 fixtures-only surface 可渲染；Vue 侧 8 个 surface 完成端到端采集
  （5 个 primitive + FileUpload / FileThumbnail / PdfToolbar）。
- `pnpm ui:parity:self-check`（参考页对自身）在完整矩阵下 344/344 全部 `pass`
  （0 像素差、0 几何差），证明截图/阈值链路可重复；`pass` 不是永远无法到达的状态。
- #157 的 5 个 primitive（Button/Input/Select/Dialog/Tooltip）已接入，最终矩阵 200 case
  中 196 `pass`、4 个 `input/focus/light` 仅有 52 px、单通道 1/255（diff ratio 0.077%）
  的抗锯齿残差；其余状态 0 px。
- `file-upload` / `file-thumbnail` / `pdf-toolbar` 全部 `fail`，这是**真实 delta**：
  - FileUpload：参考是 icon-cluster + browse chip + 组件内文件列表；当前 Vue 组件是宿主渲染的简化 dropzone（180px、单图标、额外 “Accepted/Max size” 行）。
  - FileThumbnail：参考是方形预览卡（loading/error/fallback 图层）；当前 Vue 是带文件名/类型的行内缩略图。
  - PDF toolbar：参考 `min-h-12`、`px-3 py-2`、`gap-2`、ghost icon buttons；当前 Vue toolbar 在 936px 内换行成两行（88px vs 48px），控件仍含原生 `<select>`、搜索表单行和文件名。
- `file-upload` / `file-thumbnail` 已由 #163 收敛：
  `pnpm ui:parity -- --only file-upload,file-thumbnail --skip-tokens` 的 80 个 case 全部通过
  （61 `pass` + 19 `pass-with-antialiasing`，0 `fail`）。实现细节、状态和边界见
  [components/file-components.md](../components/file-components.md)。
- `pdf-toolbar` 仍有真实 delta：参考 `min-h-12`、`px-3 py-2`、`gap-2`、ghost icon buttons；
  当前 Vue toolbar 在 936px 内换行成两行（88px vs 48px），控件仍含原生 `<select>`、搜索表单行和文件名。
- `xlsx-toolbar` / `xlsx-ribbon` **没有**可跑的 surface：上游 `xlsx-viewer.tsx` / `xlsx-editor.tsx`
  依赖不可分发的私有包 `@extend-ai/react-xlsx`，参考侧无法渲染。结论与替代证据
  （geometry / structural / token 三层已对齐，像素 diff 不可达）见
  [design/xlsx-office-chrome-migration.md §4.2](xlsx-office-chrome-migration.md)。
- `document-sidebar` 只保留参考基线：规范组件由 #158 在
  `packages/vue-ui/src/components/chrome/` 实现；#163 记录 seam，不重复实现。
- 环境对齐：Extend 参考页的 `.dark` 只切 token，而 #156 的 `theme.css` 同时设置 `color-scheme`。
  `ParitySurface` 现在按主题写入 `color-scheme`，避免 focus ring 因为原生控件配色不同而误报。

- token 层：27 个标量 token 中 26 个与 #156 一致；唯一差异是 `--office-control-padding-inline`（12px）与
  Extend 的 `px-[calc(--spacing(3)-1px)]`（11px）语义不同，已在报告里注明。#156 未声明的 4 个 `--office-code*`
  token 由 token 报告标为 `missing-in-token-layer`。

## 8. 新增一个 surface

1. 在 `tools/extend-reference/src/surfaces/` 加参考渲染，并在 `App.tsx` 的 `renderers` 注册；需要新组件时先把最小上游文件 vendored 到 `src/vendor/`（MIT 头 + NOTICE + `UPSTREAM.md` + `attribution.json`）。
2. 在 `tools/vue-reference/src/surfaces/` 加对照渲染，并在 `App.vue` 里接上；如果组件还没实现，把 manifest 标成 `subjectStatus: "missing"` + `pendingIssue`。
3. 在 `tools/parity-fixtures/manifest.json` 增加 surface（`cellWidth`、`states`、交互、依赖 issue），共享文案/数据放 `fixtures.json`；同时声明 `core`（是否进 PR 子集）与 `gateStatus`（`pending` + `gatePendingIssue`，或 `blocking`）。
4. 在 [components/ui-component-map.md](../components/ui-component-map.md) 补一行，`pnpm ui:docs` 会检查。
5. 运行 `pnpm ui:parity -- --only <id>` 和 `node --test tests/unit/ui-parity-gate.test.mjs`；surface 就绪后按
   [ui-parity-ci.md §5](../testing/ui-parity-ci.md) 切成 `blocking` 并用 `pnpm ui:baseline capture --only <id>` 补基线。

## 9. 后续 issue 分工

- #156：token/theme/OfficeIcon（已合入；本 harness 直接消费 `theme.css`）。
- #157：Button/Input/Select/Dialog/Tooltip primitive + reka-ui 层（已合入，像素对照见 §7）。
- #158：共享 viewer chrome + 规范 `DocumentViewerSidebar`（`components/chrome/`）。

- #158/#162：PDF chrome → 收敛 `pdf-toolbar` 的 2 行布局、padding、gap 与控件样式。
- #163：FileUpload / FileThumbnail 已收敛；DocumentViewerSidebar 的组件与 demo 由 #158 负责。
- #164：document blocks 渲染对照。
- #165：gate、基线、stylelint/axe、文档——已完成基础设施；pending surface 逐个切 `blocking`
  （见 [testing/ui-parity-ci.md §5](../testing/ui-parity-ci.md)）。
