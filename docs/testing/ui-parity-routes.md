# L2 产品路由 parity：demo 路由 vs 上游参考路由（#182）

> 批次：0.8.1 / issue [#182](https://github.com/arcships/agentic-office-ui/issues/182)。
> 上游固定基准：[`extend-hq/ui@70fe70c6`](https://github.com/extend-hq/ui/tree/70fe70c6f1e1eee53954c6923f3d9f583557a57e)（MIT）。
> 架构背景见 [design/viewer-editor-architecture.md](../design/viewer-editor-architecture.md) §11.2 与附录 C。

## 1. 这一层解决什么

L1（`pnpm ui:parity`）比较的是**我们自己拼装的组件格子**：一个 toolbar、一个 sidebar、一个 block。它能证明“组件本身的尺寸/颜色对齐”，但不能证明 `pnpm dev` 打开一个页面时看到的**整块产品区域**和上游一致。L1 为 0 像素仍然可能整页不一致，原因有三类：

1. demo 页面有上游 docs 页面没有的外壳（页面 header、样例选择器、props 面板、状态网格、诊断列表）；
2. 区域尺寸是 demo 自己定的（PDF 页原本 `80vh`，上游 docs 是固定 `560px`）；
3. 多个组件组合后的布局（toolbar 换行、侧栏占位、画布背景、页码位置）不在任何单个格子里。

L2 因此不再由 harness 拼装组件，而是**在同一状态下渲染两个完整路由**，只截取产品区域：

- reference：`tools/extend-reference/route.html`，原样 import 固定 commit 的上游 wrapper（R2）；
- subject：`apps/demo` 的真实 hash 路由（`#/pdf-viewer`）；
- 两侧用同一 fixture 字节、同一主题、同一 viewport、同一缩放与侧栏状态，区域高度固定为上游预览高度。

## 2. 路由矩阵（当前）

| route | reference route / 文件 | subject | 状态 | 归属 |
|---|---|---|---|---|
| `/pdf-viewer` | `/docs/components/pdf-viewer`，`apps/v4/components/pdf-viewer-docs.tsx:207` → `pdf-viewer-preview-client.tsx` | `apps/demo/src/pages/PdfViewerPage.vue` | **pilot，report-only**（R2 已跑通，8 个 case） | #182 |
| `/docx-viewer` | `/docs/components/docx-viewer`，`docx-viewer-docs.tsx:42` | `DocxViewerPage.vue` | 待建：R2 需要 `@extend-ai/react-docx@0.15.0`（MIT、公开） | #159 |
| `/xlsx-viewer` | `/docs/components/xlsx-viewer`，`xlsx-viewer-docs.tsx:40` | `XlsxViewerPage.vue` | 待建：R2 需要 `@extend-ai/react-xlsx@0.23.0`（MIT、公开） | #160 |
| `/pptx-viewer` | `/docs/components/pptx-viewer`，`pptx-viewer-demo.tsx:33` | `PptxViewerPage.vue` | 待建：R2 需要 `@extend-ai/react-pptx@0.5.0`（MIT、公开） | #161 |
| `/pdf-editor` | `/docs/components/pdf-editor`，`pdf-editor-docs.tsx:100` | 无页面 | 无对应页面（先建 demo 路由） | #172 |
| `/csv` | `/docs/components/csv-tsv-viewer` | 无页面 | 无对应页面（#178 决策后再建） | #178 |
| `/file-system` | `/docs/components/file-system` | 无页面 | 无对应页面（#179） | #179 |
| `/schema-builder` | `/docs/components/schema-builder` | 无页面 | 无对应页面（#179） | #179 |

为什么先做 `/pdf-viewer`：PDF 是唯一 **0 个 `@extend-ai/*` import** 的 viewer，依赖全部是公开 Apache-2.0 的 EmbedPDF 插件和 MIT 的 shadcn/base-ui 原语，因此可以真正做到“原样 import 上游 wrapper”，而不是先复制一份近似实现。DOCX/XLSX/PPTX 的 R2 需要安装公开的 MIT runtime 包，工作量与内存更大，留给各自 owner issue。

## 3. Product region 定义（可复现）

| 项 | reference | subject |
|---|---|---|
| app | `tools/extend-reference/dist`（`route.html`） | `apps/demo/dist`（`#/pdf-viewer`） |
| URL | `/route.html?route=pdf-viewer&theme={theme}&zoom=1&doc=/fixtures/sample.pdf` | `/?theme={theme}&sample=sample.pdf&zoom=1&statusbar=off#/pdf-viewer` |
| region selector | `[data-slot="pdf-viewer"]` | `[data-testid="pdf-viewer"]` |
| region 高度 | `560px`（`className="h-[560px]"`，`pdf-viewer-preview-client.tsx:18`） | `560px`（`.viewer-container :deep(.pdf-viewer-region)`；原为 `80vh`） |
| demo 区域标记 | — | `.viewer-container[data-route-region="viewer"][data-testid="pdf-viewer-region"]` |
| 宽度来源 | 复刻 demo 页面盒：`padding 24px` + `max-width 1200px` + 1px 外框 | `.page` 的 `padding 24px` + `max-width 1200px` + `.viewer-container` 1px 外框 |
| ready 信号 | `[data-pdf-viewer-page="1"] img[src^="blob:"]` 且 `naturalWidth > 0` | `[data-testid="pdf-status"][data-state="ready"]` 且存在 `[data-testid="pdf-page-render"]` |
| fixture | `/fixtures/sample.pdf`（harness 从 `apps/demo/public/samples/sample.pdf` 提供） | `/samples/sample.pdf`（demo 自身静态文件） |

**fixture 同一性**：两侧响应在运行前都会被抓取并计算 sha256，必须同时等于 `routes.json` 里钉住的 `76e533d3…`。任一侧字节不同，整个矩阵记为不可信并以非零退出。

**几何测量**：harness 以 2x DPR 截取**整页**，再按 region 的文档坐标裁成整数设备像素，因此 region 落在半 CSS 像素上（demo 页 viewer 上方有一个 0.5px 行盒）也不会造成两侧半像素相位差。region 的文档坐标、地标 box 与样式都会写进报告。

**对比的地标**（`routes.json` → `reference.landmarks` / `subject.landmarks`）：

| landmark | reference | subject | 容差 |
|---|---|---|---|
| `root` | `[data-slot="pdf-viewer"]` | `[data-testid="pdf-viewer"]` | 0px |
| `toolbar` | `[data-slot="pdf-viewer"] > div:first-child` | `[data-testid="pdf-toolbar"]` | 0px |
| `viewport` | `[data-slot="scroll-area-viewport"]` | `.pdf-frame-wrap` | 0px |
| `page1` | `[data-pdf-viewer-page="1"]` | `[data-testid="pdf-page-render"]` | 1px（页面尺寸亚像素取整） |
| `page1Raster` | 页面里的 `<img>` | `[data-testid="pdf-page-image"]` | 2px（栅格取整 + 1px 内边距） |

**chrome-only 对比**：额外隐藏两侧的页面框（`[data-pdf-viewer-page]` / `[data-testid="pdf-page"]`）再截一次，只比较 toolbar、画布背景、滚动条和留白。整页像素差里混着 PDF 栅格差异，chrome 差才是“产品外壳”差异；报告把两者分开。

**明确排除的 demo 脚手架**（`routes.json` → `exclusions`，报告里逐条列出）：

| 侧 | selector | 类别 | 说明 |
|---|---|---|---|
| reference | `[data-slot="code"]` | demo-scaffolding | docs 页的用法代码块 |
| reference | `[data-slot="component-preview"]` | demo-scaffolding | docs 预览卡片；只复刻其盒模型以保证两侧宽度一致 |
| subject | `[data-testid="pdf-viewer-page"] > h2/.desc/.controls/.policy-panel/.status-grid/.diagnostics` | demo-scaffolding | 页面标题、样例选择、运行时策略面板、状态网格、诊断列表 |
| subject | `[data-route-region="viewer"]` | demo-scaffolding | demo 的 1px 外框；只复刻盒模型 |

## 4. Reference 构建策略（R1/R2/R3）

| 方案 | 内容 | 本 issue 的取舍 |
|---|---|---|
| R1 | 手工裁剪 vendored 组件（现状 `src/vendor/`） | 只能覆盖 L1 组件格子；裁剪版无法证明整页一致，因此不作为 L2 reference |
| **R2（采用）** | 原样 import 上游 wrapper：`src/upstream/` 保存固定 commit 的文件正文（仅前置来源头），`@/` alias 指到该目录，安装固定版本公开依赖（`@embedpdf/*@2.15.0`、`@base-ui/react@1.4.1`、`pdf-lib@1.17.1`、`@hugeicons/*`），`route.html` 作为第二个 Vite entry | 已在纸面上和构建上验证：`vite build` 产出 6348 modules / `route-*.js` 546 kB + `worker-engine-*.js` 709 kB，浏览器里真实加载 PDF、渲染 4 页、toolbar 完整 |
| R3 | 跑固定 commit 的 Next 应用（`pnpm --filter v4 dev`）再截图 | 最真实但 CI 成本最高；R2 已经能给出同样的组件树，暂不引入 |

R2 的两个刻意偏离（都记录在 `tools/extend-reference/src/vendor/NOTICE.md`）：

1. `public/pdfium.wasm` 由 `scripts/copy-pdfium.mjs` 从 `@embedpdf/pdfium@2.15.0` 复制，`lib/pdf-thumbnail-utils.ts` 改指本地 URL；上游从 jsDelivr 拉取同一个文件。这样截图不依赖网络。
2. docs 页面壳（`padding 24px`、`max-width 1200px`、1px 外框）由 `route.css` 复刻，目的是让两侧 region 拿到同一可用宽度；docs 的侧栏/导航/代码块一律不渲染。

已知的环境差异（不是产品差异，报告里单列）：

- 上游 reference 跑 EmbedPDF **2.15.0**，产品当前 pin **2.14.4**（`packages/vue-pdf/package.json`）。页面栅格的抗锯齿/缩放可能因此不同，所以整页像素差不能直接归因于我们的 chrome。
- reference 是 React 19 + Tailwind v4，subject 是 Vue 3 + 手写 CSS；字体栈都是 Geist（reference 通过 `@fontsource-variable/geist`，demo 用同一份 CSS）。

## 5. 状态合同

`pnpm dev` 的默认状态**不因本 issue 改变**：查询参数只在 harness 传入时生效。

| 参数（demo） | 值 | 作用 |
|---|---|---|
| `theme` | `light` / `dark` | 已有：`readDemoTheme()` 写 `documentElement.dataset.theme` |
| `sample` | `/samples/` 下的文件名 | 新增：初始 fixture（非法值回退到页面默认） |
| `zoom` | `0.5`–`2` | 新增：初始缩放，越界夹取 |
| `statusbar` | `off` | 新增：关闭可选状态栏（上游没有该条） |

为什么需要这些开关：上游 docs demo 的默认状态是 `defaultZoom=0.5`、侧栏关闭、无状态栏；我们的 demo 默认 `zoom=1`、侧栏打开、状态栏打开。要比较**同一个产品状态**，捕获状态必须显式（`zoom=1` + 侧栏关闭 + 状态栏关闭），而默认差异本身作为“产品差异”记录在 §6，不被静默吞掉。侧栏没有公开 prop，harness 用 `prepare` 点击 `[data-testid="pdf-sidebar-toggle"]` 关闭。

## 6. 差异分类与当前结论

分类规则（只有 `product-component` 阻断；`demo-scaffolding` 必须显式排除；`absent-upstream` 不得伪造对照）：

| id | 类别 | 阻断 | 状态 | 证据 |
|---|---|---|---|---|
| `sidebar-default-open` | product-component | 是 | open | `packages/vue-pdf/src/components/PdfViewer.vue:120` `showThumbnails = ref(true)`；上游 `pdf-viewer.tsx:2280` `useState(false)` |
| `optional-status-bar` | product-component | 是 | open | `PdfViewer.vue:1952` 有 `[data-testid="pdf-status-bar"]`，上游 viewer 无状态栏 |
| `canvas-background` | product-component | 是 | open | `PdfViewer.vue:2008` `--pdf-canvas-bg: #e5e7eb`；上游画布 `bg-muted/30`（light ≈ `#fcfcfc`，dark ≈ `#0d0d0d`） |
| `dark-page-inversion` | product-component | 是 | open | `PdfViewer.vue:2024` `--pdf-page-filter: invert(1) hue-rotate(180deg)`；上游暗色不反转页面 |
| `page-raster-softening` | product-component | 是 | open | 上游页面 `<img>` 带 `blur-[0.35px]`；我们的 `.pdf-page-image` 不带 |
| `mobile-canvas-padding` | product-component | 是 | open | 768/390 视口下我们的 `.pdf-frame-wrap` padding 12px，上游恒为 24px；page1 地标偏移 -12px |
| `page-raster-size-rounding` | product-component | 否 | open | page1 尺寸 595.27×841.88 vs 595×841.28（zoom=1 亚像素取整），在 1px 地标容差内 |

Pilot 的定量结果（`output/ui-parity-routes/routes-matrix.json`，2x DPR）：

| viewport | theme | 整页像素差 | chrome 像素差 | 最差地标几何差 |
|---|---|---|---|---|
| 1440x900 | light | 56.484% | 92.063% | 1.73px |
| 1440x900 | dark | 92.388% | 92.436% | 1.73px |
| 1280x720 | light | 56.484% | 92.063% | 1.73px |
| 1280x720 | dark | 92.388% | 92.436% | 1.73px |
| 768x1024 | light | 43.319% | 92.050% | 13.00px |
| 768x1024 | dark | 91.287% | 92.647% | 13.00px |
| 390x844 | light | 38.586% | 79.567% | 13.00px |
| 390x844 | dark | 78.993% | 81.076% | 13.00px |

读法与已证明/未证明的边界：

- **已证明（0 像素）**：`root` / `toolbar` / `viewport` 三个地标在 8 个 case 里 Δ=0，说明区域尺寸、toolbar 高度（含 390 下的 125px 换行高度）、滚动容器高度与上游完全一致；harness 自检（reference 对 reference）0.000% 差异，说明截图/裁剪/相位对齐可重复。
- **chrome 差 92% 的主因是画布背景**：报告 `chrome.dominantColors` 显示 light 下 reference `#fcfcfc` 占 91.4%、subject `#e5e7eb` 占 91.3%；dark 下 `#0d0d0d` 91.4% vs `#020617` 91.3%。这不是测量误差，而是上表第 3 条产品差异。
- **整页差的主因是页面栅格**：上游对页面位图做 0.35px 模糊、我们不做；dark 下我们还整体反色。这两条决定暗色永远不可能达标。
- **未证明**：没有一条路由达到“几何 0px、像素 ≤0.1%”的验收线；因此 `gateStatus` 保持 `report-only`。要转 blocking，先按 §8 修完阻断项。

## 7. 运行与产物

```bash
pnpm install                                   # 新增 @embedpdf/*@2.15.0 等公开依赖
pnpm ui:parity:routes                          # 构建 reference + demo，跑完整矩阵（report-only）
pnpm ui:parity:routes -- --only pdf-viewer --viewports 1440x900 --themes light
pnpm ui:parity:routes -- --self-check --only pdf-viewer --viewports 1440x900 --themes light
pnpm ui:parity:routes -- --gate-mode ready      # 只让 gateStatus=blocking 的路由阻断
pnpm ui:parity:routes -- --no-build --out output/ui-parity-routes
```

需要 Python Playwright + Pillow（与 `tests/visual` 相同的环境）：

```bash
PYTHON=/tmp/bb-venv/bin/python RUN_PYTHON_MODULES=playwright,PIL \
  pnpm ui:parity:routes -- --no-build
```

产物（默认 `output/ui-parity-routes/`，不进版本库）：

```
report.html            人读报告：每 case 的 reference/subject/diff/heatmap
report.json            机读报告（与 routes-matrix.json 同内容）
SUMMARY.md             矩阵、地标 delta、分类差异、排除项
routes-matrix.json     route × viewport × theme 矩阵（CI artifact 入口）
cases/<route>/<viewport>/<theme>/
  reference.png subject.png diff.png heatmap.png compare.png
  reference-chrome.png subject-chrome.png diff-chrome.png heatmap-chrome.png
```

退出码：默认 `--gate-mode off`（report-only）恒为 0，但报告会写 `status: fail`；`ready` 只让 `gateStatus: blocking` 的路由失败；`all` 任何路由失败都失败。fixture 字节不一致直接非零退出，因为此时矩阵不可信。

本 issue 的 pilot 证据（1x 缩略图 + 生成矩阵）已提交在
[`docs/assets/182-route-parity/`](../assets/182-route-parity/)：`1440x900` light/dark
的 reference/subject/diff/heatmap、`1440x900/light` 的 chrome-only 三张，以及
`390x844/light` 的 reference/subject/diff，配 `SUMMARY.md` 与
`routes-matrix.json`。需要 2x 原图时重跑上面的命令。

`tools/parity-fixtures/routes.json` 是唯一配置源：

- `apps.{reference,subject}`：两个构建目录与 fixture 前缀；
- `fixtures`：path + sha256 + bytes；
- `routes[]`：URL 模板、ready 选择器、prepare 动作、region/landmark 选择器、chrome mask、排除项、分类差异、地标容差；
- `viewports` / `themes` / `thresholds`：与 L1 同义（几何 0px、像素 0.1%、channel delta 2、热区 Top-N）。

## 8. 后续路由与阻断项

1. **先修 proxy 级产品差异，再谈 blocking**（建议顺序）：画布背景 token → 页面栅格 softening → 暗色不反转页面 → 侧栏默认值 → 状态栏归属（保留但默认关闭，或从 viewer 移除）→ 移动端画布 padding。每修一项重跑 `pnpm ui:parity:routes`，把对应 `differences[].status` 从 `open` 改为 `fixed`。
2. **DOCX/XLSX/PPTX 复用同一 harness**：只需在 `tools/extend-reference/src/upstream/` 增加对应公开文件、在 `src/routes/` 加一个 route 组件、在 `routes.json` 加一条 route 记录。R2 的主要成本是安装公开 MIT runtime（分别约 67k/…/… 行）与构建内存，不是 harness 改造。三条路由的 region 高度、选择器与状态可以从 §3 的模板直接推导。
3. **无对应页面的能力**（PDF editor、CSV、FileSystem、SchemaBuilder、PPTX presenter）不在 L2 做像素对照，按 `office-reference` 或 `product-gap` 处理（见架构审计 §1.2）。
4. **CI**：`.github/workflows/ui-parity.yml` 已加 `ui-parity-routes` job（PR / nightly / dispatch），以 `--gate-mode off` 产出 `output/ui-parity-routes/` artifact；等到 `/pdf-viewer` 的阻断项清零，再把 `routes.json` 的 `gateStatus` 改为 `blocking`，并把命令改成 `--gate-mode ready`。

## 附录：证据清单

| 证据 | 位置 |
|---|---|
| 上游固定 commit | `tools/parity-fixtures/routes.json` → `referenceCommit`；`/tmp/extend-ui` 本地 checkout |
| reference app 构建 | `pnpm --filter @aou/extend-reference build`；`dist/route.html` + `route-*.js` + `worker-engine-*.js` |
| 上游 wrapper 原样性 | `tools/extend-reference/src/upstream/**` 与 `apps/v4/**` 逐文件 diff（仅来源头） |
| fixture 字节 | `routes.json` → `fixtures.sample-pdf.sha256`，运行时双侧抓取校验 |
| 截图/差异 | `output/ui-parity-routes/cases/**`；已提交的 1x 证据在 `docs/assets/182-route-parity/` |
| 矩阵与分类 | `output/ui-parity-routes/routes-matrix.json` |
| 自检确定性 | `pnpm ui:parity:routes -- --self-check` → 0.000% |
| attribution | `node scripts/ci/verify-attribution.mjs`（87 个 adapted/copied 文件） |
