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
| `docx-toolbar` | Extend 风格 `DocxToolbarSurface`：与 subject 相同的文件/页/缩放/搜索/视图/修订/主题/上传下载控件，同一 `<1360px` overflow 断点 | `DocxViewerToolbar` | default / hover / focus / disabled |
| `docx-sidebar` | `DocumentViewerThumbnailSidebar` + 与 subject 相同的单页真实 DOCX 页图、卡片尺寸/标签、header/计数和 3 项 skeleton | 真实 `DocxThumbnailPanel` + `DocumentViewerSidebar` + 真实 DOCX canvas 缩略图 | open / closed / skeleton |

运行：

```bash
PYTHON=/tmp/bb-venv/bin/python pnpm ui:parity -- \
  --only docx-toolbar,docx-sidebar
```

参考侧扩展了 `tools/extend-reference/src/surfaces/DocxToolbarSurface.tsx` 与
`DocxSidebarSurface.tsx`：toolbar 补齐与 subject 相同的控制集、顺序和 `<1360px`
overflow 断点；sidebar 使用相同的一张真实 DOCX 页图（`public/docx-sidebar-page-1.png`，
由真实 `DocxThumbnailPanel` canvas 生成）、相同卡片 140×210、相同 label/active 样式、
相同 “Pages” header/计数和相同 3 项 skeleton 结构。vendored 组件仍保留 pinned
commit / MIT / copyright 头并登记在 `vendor/NOTICE.md`。

该命令当前**真实 FAIL**（没有伪造 pass），精确 delta：

- `docx-toolbar`：32 cases，**几何全矩阵 PASS**（1440/1280/768 单行 48px；390
  响应式换行 89px，与 reference 高度一致）。像素差：1440 11.82%–20.54%；1280
  2.90%–6.32%；768 4.96%–10.82%；390 7.03%–15.99%；全矩阵平均 9.99%。剩余差异
  集中在控件文字/图标字形与 1px 位置差（Extend Hugeicons vs Office Lucide、页控件
  输入宽度、搜索选项/计数 pill 尺寸）。
- `docx-sidebar`：24 cases，几何全矩阵 PASS；10 pass、2 pass-with-AA、12 fail，
  平均 3.17%。closed 全 PASS；open 仅 0.50%–0.53%（header “Pages”/计数文字与 border
  的亚像素抗锯齿）；skeleton 在 1440/768 light PASS，dark 18.8%–23.5%、390
  11.3%–18.3%，原因是 subject 的 `office-skeleton` shimmer 动画取帧相位与 reference
  静态块不同。

报告：`output/ui-parity/report.json` / `report.html`（未提交，由命令生成）。
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
