# Migration Plan

## 概述

将上游 `@extend-ai/react-xlsx`（~41000 行）和 `@extend-ai/react-docx`（~68000 行）迁移到本地 Vue workspace。

设计文档：
- `docs/upstream-xlsx-feature-alignment.md` — XLSX 逐文件对齐清单 + 操作方案
- `docs/upstream-docx-feature-alignment.md` — DOCX 逐文件对齐清单 + 操作方案
- `docs/docx-migration-architecture.md` — DOCX 模块化重做架构设计（单文件 ≤1000 行）
- `docs/docx-editor-helpers-split-plan.md` — editor.tsx 24953 行拆分为 24 模块的详细方案
- `docs/visual-acceptance-handoff.md` — 最终验收标准
- `docs/INDEX.md` — 文档总索引

## 当前状态

### DOCX 迁移（模块化重做）

前一次迁移按旧方案机械复制上游巨型文件（editor-helpers.ts 23445 行），已回退（commit `ff149fb`）。当前采用模块化重做，硬约束：**单文件 ≤ 1000 行**。

已完成：
- ✅ 引擎层（engine/，~1700 行，8 文件）— wasm smoke + python-docx round-trip 验证通过
- ✅ 架构设计（docs/docx-migration-architecture.md，~124 文件规划）

进行中：按架构设计第六节执行顺序推进 layout → viewer → editor/helpers → composables → components。

### XLSX 迁移

master 上已有 xlsx-core 模块化迁移成果（charts/colors/images/wasm 等），vue-xlsx 为 stub。task/xlsx-005 分支有未合并的 XlsxViewer 重写工作（已回退）。

## 任务列表

### DOCX 模块化重做任务

| ID | 任务 | scope | depends-on | status |
|---|---|---|---|---|
| docx-engine | 引擎层（engine/，8 文件） | docx-core | — | ✅ done |
| docx-layout | 布局层（layout/，5 文件，~2653 行） | docx-core | docx-engine | pending |
| docx-canvas | canvas 层（canvas/，2 文件） | docx-core | docx-engine | pending |
| docx-viewer | viewer 辅助模块（viewer/，15 文件，~4713 行） | docx-core | docx-layout | pending |
| docx-helpers | editor/helpers 30+ 模块（~21200 行） | docx-core | docx-viewer | pending |
| docx-editor-ops | editor-ops + state 拆分（5 文件） | docx-core | docx-helpers | pending |
| docx-composables | vue-docx composables（26 文件，~8700 行） | vue-docx | docx-helpers | pending |
| docx-render | renderParagraphRuns 重写（6 文件，~1505 行） | vue-docx | docx-composables | pending |
| docx-components | Vue 组件重写（18 文件，~10900 行） | vue-docx | docx-composables, docx-render | pending |
| docx-demo | demo 页面对齐 26 项功能 + 构建配置 | demo | docx-components | pending |
| docx-verify | 全量验证（typecheck + build + python-docx + 视口） | — | docx-demo | pending |

### XLSX 迁移任务（旧任务体系，保留参考）

| ID | 任务 | depends-on | status |
|---|---|---|---|
| xlsx-001 | xlsx-core 引擎层复制 + types 清理 | — | done（master） |
| xlsx-003 | vue-xlsx controller 机械改写 | xlsx-001 | done（master） |
| xlsx-004 | vue-xlsx chart-renderer + surface-regl 改写 | xlsx-001 | done（master） |
| xlsx-005 | vue-xlsx XlsxViewer 重写 | xlsx-003, xlsx-004 | 待评估（task/xlsx-005 分支已回退） |
| xlsx-006 | demo XlsxViewerPage 接入 | xlsx-005 | pending |
| cross-001 | 构建配置 + wasm 部署 | xlsx-001, docx-engine | done（master） |

### 集成验证任务

| ID | 验证内容 | depends-on | status |
|---|---|---|---|
| docx-verify | DOCX 全量验证 | docx-demo | pending |
| xlsx-integ | composables → xlsx-core 真实调用 | xlsx-003 | done（master） |
| cross-integ | viewer → controller, demo → 全链路 | xlsx-005, docx-components | pending |

## DOCX 模块化重做依赖图

```
docx-engine ✅ ──→ docx-layout ──→ docx-viewer ──→ docx-helpers ──→ docx-composables ──→ docx-components ──→ docx-demo ──→ docx-verify
                                       ↑                                ↑       ↑
                                   docx-canvas                    docx-render ┘
```

## 旧任务文件说明

`docs/plan/tasks/docx-001.md` ~ `docx-006.md` + `docx-integ.md` 是按旧方案（机械复制）设计的，**已失效**。保留作为历史参考，实际执行以本文件的"DOCX 模块化重做任务"为准。

对应旧 review 文件（`docs/plan/reviews/docx-001-1.md`、`docx-002-1.md`）针对的是已回退的代码，**已失效**。

## 执行模式

### DOCX（模块化重做）

按架构设计 `docs/docx-migration-architecture.md` 第六节执行顺序，逐步推进。每步验证 typecheck + build。大块（docx-helpers、docx-composables、docx-components）可分批用 agent 并行。

### XLSX

旧任务体系保留。xlsx-005 需重新评估是否重做或继续。
