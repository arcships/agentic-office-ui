# DOCX chrome screenshot stories（issue #159）

`scripts/fixtures/capture-docx-ui-stories.py` 为 DOCX viewer/editor 的 chrome
迁移提供确定性截图。它不复用 #74 的 12-scene 产品基线，而是补齐本 issue 需要但
产品矩阵没有组合到的状态：

| Story | Route | Viewport | Theme | 状态 |
|---|---|---|---|---|
| `viewer-continuous-light-1440` | `/docx-viewer` | 1440×900 | light | continuous |
| `viewer-single-dark-1440` | `/docx-viewer` | 1440×900 | dark | single |
| `viewer-spread-light-1440` | `/docx-viewer` | 1440×900 | light | spread |
| `viewer-continuous-light-390-overlay` | `/docx-viewer` | 390×844 | light | thumbnails overlay |
| `viewer-continuous-dark-768-overlay` | `/docx-viewer` | 768×1024 | dark | thumbnails overlay |
| `editor-light-1440` | `/docx-editor` | 1440×900 | light | editable + inline sidebar |
| `editor-readonly-dark-1440` | `/docx-editor` | 1440×900 | dark | read-only |
| `editor-light-390-overlay` | `/docx-editor` | 390×844 | light | thumbnails overlay |

## 运行

```bash
# 先构建 demo 和 workspace 依赖
pnpm --filter demo build

# 复用当前 dist，输出到 tests/baseline/visual/docx-stories/
python3 scripts/fixtures/capture-docx-ui-stories.py --skip-build

# 只重跑一个 story
python3 scripts/fixtures/capture-docx-ui-stories.py --skip-build \
  --only viewer-continuous-light-1440
```

需要 Playwright 和 Pillow（本机 `/tmp/bb-venv/bin/python`，CI 用
`pip install -r requirements-ci.txt`）。

## 确定性

每个 story 都新建 browser context：DPR 1、`en-US`、UTC、`reduced_motion=reduce`、
固定字体栈、隐藏 scrollbar、关闭动画/过渡与 caret；加载后等待 `document.fonts`、
两帧 `requestAnimationFrame` 和 300ms settle。`?theme=` 选择配色。截图以无损 WebP
写入 `tests/baseline/visual/docx-stories/<story>.webp`，manifest 记录 SHA-256、
viewport、toolbar 外高、最大控件高度、横向 `scrollWidth/clientWidth` 以及 sidebar
解析出的 `inline` / `overlay` 模式和宽度。

脚本对每个 story 做 blocking 断言：

- toolbar 无横向隐藏：`scrollWidth <= clientWidth + 1`；
- toolbar 高度 ≥ 47px（导航 chrome 合同是 48px；1px 为底边）；
- thumbnails 打开时：≤768px 必须 overlay，>768px 必须 inline 且宽度 160px。

## 当前结果

2026-11-20 在 `fix/159` 上采集（`page-sizes.docx`）：

- viewer 1440：toolbar 48px，最大控件 28px；
- viewer 390：toolbar 85px（compact 控件 + 搜索 trigger 在 356px 容器内换行），无横向溢出，sidebar overlay；
- viewer 768：toolbar 48px（compact overflow），无横向溢出，sidebar overlay；
- editor 1440：topbar + formatting bar 合计 117px，sidebar inline 160px；
- editor 390：toolbar 346px（多行 formatter），无横向溢出，sidebar overlay。

## 与 #155 parity harness 的关系

本 issue 已扩展 `tools/parity-fixtures/manifest.json`：

| Surface | Reference | Subject | States |
|---|---|---|---|
| `docx-toolbar` | vendored `DocxToolbar` / `DocxPageNumberControl` | `DocxViewerToolbar` | default / hover / focus / disabled |
| `docx-sidebar` | vendored `DocumentViewerThumbnailSidebar` | 真实 `DocxThumbnailPanel` + `DocumentViewerSidebar` + 真实 DOCX canvas 缩略图 | open / closed / skeleton |

运行：

```bash
PYTHON=/tmp/bb-venv/bin/python pnpm ui:parity -- \
  --only docx-toolbar,docx-sidebar
```

参考侧新增 `tools/extend-reference/src/vendor/extend/docx-toolbar.tsx`，带 pinned
commit / MIT / copyright 头并登记在 `vendor/NOTICE.md`；#155 harness 合同测试与
fixture 同步更新。该命令当前**真实 FAIL**（没有伪造 pass），精确 delta：

- `docx-toolbar`：32 cases，**几何全矩阵 PASS**（1440/1280/768 单行 48px；390
  响应式换行 89px，与 reference 高度一致）。像素差：1440 9.19%–16.08%（宽屏保留
  可见 search/edit/theme/upload/download）；1280 2.20%–4.73%；768 3.77%–8.10%；
  390 5.36%–12.81%；全矩阵平均 7.99%。宽度 <1360px 时 edit/theme/upload/download
  等非参考控件收进真实 overflow dropdown；搜索保留为共享 popover 的 compact
  trigger，因此 `Ctrl/Cmd+F`、touch target 和搜索能力不丢失。
- `docx-sidebar`：24 cases，几何全矩阵 PASS（160px；open 292px、skeleton 240px），
  像素差 open 64.0%–71.8%、closed 9.08%、skeleton 49.4%–51.3%，平均 42.52%。
  subject 现在是真实 `DocxThumbnailPanel` 和真实 canvas 缩略图，不再使用 fixture
  FileThumbnail 代替。closed 的边框已在共享 sidebar 中修为
  `border-inline-end-width: 0`。

报告：`output/ui-parity/report.json` / `report.html`（未提交，由命令生成）。
剩余 delta 属于有意的产品差异：1440 宽屏保留可见 search/edit/theme/upload/download
控件；窄于 1360px 时这些控件进入真实 overflow dropdown，所以 1280/768/390 的
几何与 reference 一致，但 1440 仍保留可发现的完整控制集；sidebar 的 reference
使用 fixture FileThumbnail 卡片，subject 是真实 DOCX canvas 缩略图，因此像素差
主要来自缩略图内容本身；skeleton 的卡片数量/文案也不同。
未接入 axe-core（#165）；键盘/焦点行为由 component/blackbox 测试覆盖。

## UX-DOCX-004 square-wrap 排查

`tests/blackbox/ux_parity.py` 的 `UX-DOCX-004` 在当前分支和迁移前基线都失败：

| 结果 | Commit | 失败 |
|---|---|---|
| `/tmp/b08-159-base` | `fa7b45a6`（#159 合并前） | `besideTextRects: []`，`top/bottom 555.5/733.5` |
| `fix/159` | `2142fea9` + 本 issue 收尾 | 同一断言，`top/bottom 534.5/712.5` |

两次失败信息完全同构（`collidingTextRects: []` / `besideTextRects: []`），且
`fa7b45a6..HEAD` 没有修改文档布局/渲染路径：

```text
packages/vue-docx/src/components/DocxDocumentSurface.vue
packages/vue-docx/src/components/DocxPageSurface.vue
packages/vue-docx/src/components/DocxParagraphHost.vue
packages/vue-docx/src/components/DocxImageLayer.vue
packages/docx-core/src/**
```

以上 diff 为空；本迁移只改 toolbar、thumbnail shell、gutter/link/dialog 外壳和
viewer/editor 状态包裹层。因此 `UX-DOCX-004` 是既有 square-wrap 布局几何问题，
不是本次 chrome 迁移引入；验收不把它作为 #159 回归。

## 验证结果（2026-11-20，fix/159）

| 命令 | 结果 |
|---|---|
| `pnpm --filter @arcships/vue-ui build` | PASS |
| `pnpm --filter @arcships/vue-docx build` | PASS（vite + vue-tsc declarations） |
| `pnpm --filter demo build` | PASS |
| `pnpm --filter @arcships/vue-ui typecheck` / `@arcships/vue-docx typecheck` | PASS |
| `node --test --test-concurrency=1 tests/component/docx-*.test.mjs …` | 53 files / 全部 PASS（分批运行，含 toolbar、view modes、thumbnail、page budget、editor affordances、gutter、revision、search、RTL、tables、print） |
| `tests/blackbox/docx_editing.py` | PASS |
| `tests/blackbox/docx_export_bytes.py` | PASS |
| `tests/blackbox/docx_annotations.py` | PASS |
| `tests/blackbox/fidelity_content.py` | PASS |
| `tests/blackbox/touch_mobile.py` | PASS（20/20；390px DOCX toolbar 搜索/缩放/图标命中区 ≥44px） |
| `tests/blackbox/ux_parity.py` | `UX-DOCX-001/002/003` PASS；`UX-DOCX-004` 在基线 `fa7b45a6` 与当前分支均失败（见上节）；upstream reference capture BLOCKED |
| `pnpm ui:parity -- --only docx-toolbar,docx-sidebar` | 56 cases FAIL（未伪造通过）；几何/像素精确 delta 见上节，报告 `output/ui-parity/report.json` |
| `node scripts/ci/verify-docs.mjs` | PASS（262 checks） |
| `node scripts/ci/verify-public-api.mjs` | PASS |
| `node scripts/ci/verify-attribution.mjs` | PASS |
| `node scripts/ci/verify-test-registration.mjs` | PASS（255 files） |
| `scripts/fixtures/capture-docx-ui-stories.py --skip-build` | PASS，8 stories |
