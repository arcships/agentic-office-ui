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
- viewer 390：toolbar 85px（leading + 两行 actions），无横向溢出，sidebar overlay；
- viewer 768：toolbar 113px（leading + 多行 actions），无横向溢出，sidebar overlay；
- editor 1440：topbar + formatting bar 合计 117px，sidebar inline 160px；
- editor 390：toolbar 346px（多行 formatter），无横向溢出，sidebar overlay。

## 与 #155 parity harness 的关系

本轮没有扩展 `tools/parity-fixtures/manifest.json` 的 surface（`docx-toolbar` /
`docx-sidebar` 需要同时新增 React reference 与 Vue subject fixture）。因此没有可运行
的 `pnpm ui:parity -- --only docx-...` surface。作为替代，本脚本用同一固定 Extend
commit 的 48px toolbar / 160px sidebar / 28/32px icon / 768px overlay 合同做结构
断言，并保留截图供人工 diff。剩余 delta 记录在
[office-chrome.md](./office-chrome.md#10-docx-迁移issue-159)：

- 上游把文件动作收进 dropdown；DOCX 仍显示 theme/upload/download 独立按钮。
- 上游搜索是 popover；DOCX 仍是常驻内联搜索，以保留 `Ctrl/Cmd+F` 和现有测试契约。
- 未接入 axe-core（#165）；键盘/焦点行为由 component/blackbox 测试覆盖。

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
| `tests/blackbox/ux_parity.py` | `UX-DOCX-001/002/003` PASS；`UX-DOCX-004` 仍因本机 square-wrap 几何基线失败（非 chrome 改动）；upstream reference capture BLOCKED |
| `node scripts/ci/verify-docs.mjs` | PASS（262 checks） |
| `node scripts/ci/verify-public-api.mjs` | PASS |
| `node scripts/ci/verify-attribution.mjs` | PASS |
| `node scripts/ci/verify-test-registration.mjs` | PASS（255 files） |
| `scripts/fixtures/capture-docx-ui-stories.py --skip-build` | PASS，8 stories |
