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
| `button` | default、hover、focus、active、disabled、loading + outline/secondary/ghost/destructive/link/size-sm/size-lg/icon-sm | ❌ 无 Button（#157）→ 仅参考基线 |
| `file-upload` | default(empty)、hover、focus、dragging、rejected(error)、file-selected | ✅ `FileUpload.vue` |
| `file-thumbnail` | default(empty)、loading、error、image | ✅ `FileThumbnail.vue` |
| `pdf-toolbar` | default、hover、focus、active、disabled(empty document) | ✅ `PdfViewer.vue` 的真实 toolbar |
| `document-sidebar` | open、closed、skeleton(loading) | ❌ 无侧栏（#163）→ 仅参考基线 |
| `document-blocks` | 只有 fixtures（渲染对照由 #164 接管） | — |

截图前统一执行：`document.fonts.ready`、`prefers-reduced-motion: reduce`、注入
`animation:none / transition:none`、`caret-color:transparent`。字体：Geist / Geist Mono
（`@fontsource-variable/*`，与 `apps/v4/lib/fonts.ts` 同族）。

## 6. 产物

`output/ui-parity/`（已 gitignore）：

```
report.html                    # 人看：双方截图、diff、热力图、失败原因、token diff
report.json                    # 机器读：cases[]、summary、tokens{...}
SUMMARY.md                     # PR 摘要：按 surface 的通过/失败计数 + token 差异
cases/<surface>/<viewport>/<theme>/<state>/
  reference.png  subject.png  diff.png  heatmap.png  compare.png
  overlap-diff.png  overlap-heatmap.png     # 仅尺寸不一致时
```

CI：`.github/workflows/ui-parity.yml` 上传 artifact `ui-parity-report`（含 `report.json`）。
该 job 现在只在 `pull_request` 路径命中或手动触发时运行，采集步骤用 `pnpm ui:parity:report`（不因
阈值失败而卡住），并在之后跑 `ui:parity:self-check`（1440×900 / light）确认 harness 自身可重复。
#165 把 `pnpm ui:parity`（退出码 gate）接成必需检查。

## 7. 实施现状（2026-11-20，0.8 wave 1）

- 参考侧 5 个 surface 全部可渲染；Vue 侧 3 个 surface 完成端到端采集。
- `pnpm ui:parity:self-check`（参考页对自身）在 1440×900 / light 下 32/32 全部 `pass`
  （0 像素差、0 几何差），证明截图/阈值链路可重复；`pass` 不是永远无法到达的状态。
- `button` 只有参考基线：`@arcships/vue-ui` 还没有 Button primitive（#157）。报告里状态是
  `baseline-only`，用于给 #157 提供逐状态目标图。
- `file-upload` / `file-thumbnail` / `pdf-toolbar` 全部 `fail`，这是**真实 delta**：
  - FileUpload：参考是 icon-cluster + browse chip + 组件内文件列表；当前 Vue 组件是宿主渲染的简化 dropzone（180px、单图标、额外 “Accepted/Max size” 行）。
  - FileThumbnail：参考是方形预览卡（loading/error/fallback 图层）；当前 Vue 是带文件名/类型的行内缩略图。
  - PDF toolbar：参考 `min-h-12`、`px-3 py-2`、`gap-2`、ghost icon buttons；当前 Vue toolbar 在 936px 内换行成两行（88px vs 48px），控件仍含原生 `<select>`、搜索表单行和文件名。
- token 层：27 个标量 token 中 26 个与 #156 一致；唯一差异是 `--office-control-padding-inline`（12px）与
  Extend 的 `px-[calc(--spacing(3)-1px)]`（11px）语义不同，已在报告里注明。#156 未声明的 4 个 `--office-code*`
  token 由 token 报告标为 `missing-in-token-layer`。

## 8. 新增一个 surface

1. 在 `tools/extend-reference/src/surfaces/` 加参考渲染，并在 `App.tsx` 的 `renderers` 注册；需要新组件时先把最小上游文件 vendored 到 `src/vendor/`（MIT 头 + NOTICE + `UPSTREAM.md` + `attribution.json`）。
2. 在 `tools/vue-reference/src/surfaces/` 加对照渲染，并在 `App.vue` 里接上；如果组件还没实现，把 manifest 标成 `subjectStatus: "missing"` + `pendingIssue`。
3. 在 `tools/parity-fixtures/manifest.json` 增加 surface（`cellWidth`、`states`、交互、依赖 issue），共享文案/数据放 `fixtures.json`。
4. 运行 `pnpm ui:parity -- --only <id>` 和 `node --test tests/unit/extend-parity-harness.test.mjs`。

## 9. 后续 issue 分工

- #156：token/theme/OfficeIcon（已合入；本 harness 直接消费 `theme.css`）。
- #157：Button primitive + reka-ui 层 → 让 `button` 从 `baseline-only` 变成可比。
- #158/#162：PDF chrome → 收敛 `pdf-toolbar` 的 2 行布局、padding、gap 与控件样式。
- #163：FileUpload / FileThumbnail / DocumentViewerSidebar → 收敛上传区与缩略图。
- #164：document blocks 渲染对照。
- #165：把 `pnpm ui:parity` 变成必需 CI gate、补文本/非文本分区与 stylelint/axe。
