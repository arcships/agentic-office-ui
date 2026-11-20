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

### 重开补齐（fix/159b，2026-11-20）

#159 第一次关闭时只接受了「几何对齐 + 控件集合一致」，像素差仍有 7.99%。重开后
把参考面改成 pinned commit `70fe70c6` 的**真实上游 markup**，产品面同步：

| Surface | Reference（改后） | Subject | States |
|---|---|---|---|
| `docx-toolbar` | `vendor/extend/docx-toolbar.tsx` = `docx-viewer.tsx` 的 `DocxToolbar`（缩略图开关 + Page/页码/of N + 缩放减/档位/加 + 分隔线 + 关闭的 `⋯` 动作菜单） | `DocxViewerToolbar` 默认行 | default / hover / focus / disabled |
| `docx-sidebar` | `vendor/extend/docx-sidebar.tsx` = 上游 `DOCX pages` listbox：无 header、16px 轨迹留白、12px 行内缩进、92px `FileThumbnail` 页卡（同一张 2× 页图 canvas）、`text-xs` 页码、`bg-sidebar-accent` 选中态、`p-4` 骨架 | 真实 `DocxThumbnailPanel`（`DocumentViewerSidebar` + `FileThumbnail` + 真实 DOCX canvas） | open / closed / skeleton |

删除的旧参考面内容（它们不是上游默认行）：文件名字段、内联搜索条、版式/适配分段、
修订显示下拉、主题/上传/下载按钮，以及 sidebar 的 “Pages” header 和计数 pill。
产品侧对应命令收进 `⋯` 动作菜单（Find / Comments / Edits / 主题 / 上传 / 下载）或
改为默认关闭的可选 props（`showViewControls` / `showReviewControls` /
`showFileName` / `showSearch`）。

运行：

```bash
PYTHON=/tmp/bb-venv/bin/python pnpm ui:parity -- \
  --only docx-toolbar,docx-sidebar
```

结果（`output/ui-parity-docx/report.json`，未提交）：

| 指标 | 数值 |
|---|---|
| cases | 56（4 视口 × 2 主题 × 7 状态） |
| 通过 | **56 / 56 PASS**，0 pass-with-AA，0 fail |
| 最大 diff 像素 | **0** |
| 最大 diff ratio | **0**（阈值 0.1%） |
| 最大 channel delta | 0（阈值 2） |
| 几何 delta | 0px（宽/高/偏移全矩阵一致） |

剩余非像素类差异（`styleDeltas`，不影响栅格化）：sidebar 根节点的 `display`
（上游 block / 共享 shell flex）与 `boxShadow` 序列化文本（透明层数量），
像素比较为 0。

### 早前一轮（fix/159）的历史 delta

第一次实现把「Extend 风格完整控件集」写进参考面，于是参考面和产品面一起偏离
上游默认行。当时的报告是 toolbar 32 cases 几何 PASS / 像素平均 9.99%、sidebar
24 cases 平均 3.17%。这一轮已由上面的上游参考面取代。

报告文件（`output/ui-parity*/report.json` / `report.html`）由命令生成，不提交。
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
| `pnpm ui:parity -- --only docx-toolbar,docx-sidebar`（fix/159b） | **56/56 PASS，diff 像素 0，几何 delta 0**；报告 `output/ui-parity-docx/report.json` |
| `node scripts/ci/verify-docs.mjs` | PASS（262 checks） |
| `node scripts/ci/verify-public-api.mjs` | PASS |
| `node scripts/ci/verify-attribution.mjs` | PASS |
| `node scripts/ci/verify-test-registration.mjs` | PASS（255 files） |
| `scripts/fixtures/capture-docx-ui-stories.py --skip-build` | PASS，8 stories |
