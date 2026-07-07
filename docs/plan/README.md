# Migration Plan

## 概述

将上游 `@extend-ai/react-xlsx`（~41000 行）和 `@extend-ai/react-docx`（~68000 行）迁移到本地 Vue workspace。

设计文档：
- `docs/upstream-xlsx-feature-alignment.md` — XLSX 逐文件对齐清单 + 操作方案
- `docs/upstream-docx-feature-alignment.md` — DOCX 逐文件对齐清单 + 操作方案
- `docs/visual-acceptance-handoff.md` — 最终验收标准
- `docs/INDEX.md` — 文档总索引

## 任务列表

### Module tasks（模块实现）

| ID | 任务 | depends-on |
|---|---|---|
| xlsx-001 | xlsx-core 引擎层复制 + types 清理 | — |
| xlsx-003 | vue-xlsx controller 机械改写 | xlsx-001 |
| xlsx-004 | vue-xlsx chart-renderer + surface-regl 改写 | xlsx-001 |
| xlsx-005 | vue-xlsx XlsxViewer 重写 | xlsx-003, xlsx-004 |
| xlsx-006 | demo XlsxViewerPage 接入 | xlsx-005 |
| docx-001 | docx-core 引擎+布局+辅助复制 | — |
| docx-002 | docx-core editor helpers 复制+清理 | docx-001 |
| docx-003 | vue-docx composables 改写 | docx-002 |
| docx-004 | vue-docx DocxEditor 重写 | docx-003 |
| docx-005 | vue-docx DocxViewer 重写 | docx-003 |
| docx-006 | demo DocxViewer/Editor 接入 | docx-004, docx-005 |
| cross-001 | 构建配置 + wasm 部署 | xlsx-001, docx-001 |

### Integration tasks（集成验证）

| ID | 验证内容 | depends-on |
|---|---|---|
| xlsx-integ | composables → xlsx-core 真实调用（X1-X3） | xlsx-003 |
| docx-integ | composables → docx-core 真实调用（D1-D5） | docx-003 |
| cross-integ | viewer → controller, demo → 全链路（X4-X7, D6-D9, C1-C2） | xlsx-005, docx-004, docx-005 |

## 依赖图

```
xlsx-001 ──→ xlsx-003 ──→ xlsx-005 ──→ xlsx-006
    │            │              ↑
    │            ↓              │
    │       xlsx-integ     xlsx-004
    │
    └──→ cross-001

docx-001 ──→ docx-002 ──→ docx-003 ──→ docx-004 ──→ docx-006
    │                          │            ↑
    │                          ↓            │
    │                     docx-integ   docx-005 ──→ docx-006
    │
    └──→ cross-001

cross-integ ← xlsx-005, docx-004, docx-005
```

## 执行模式

### 无 worktree（当前环境）

所有任务串行执行——一个工作目录无法维持并发分支。

串行顺序建议（按依赖 + 价值优先）：
1. xlsx-001 + docx-001（选一条先做，或串行做两条）
2. cross-001（构建配置，让后续任务能 build 验证）
3. 后续按依赖顺序

### 有 worktree

并行条件：两个任务的 `path`（含代码和文档路径）无 parent/child/same 重叠，且无 depends-on 链。

可并行的组合：
- xlsx-001 + docx-001（不同包）
- xlsx-003 + xlsx-004（同包不同文件：composables.ts vs chart-renderer.ts）
- xlsx-003/004 + docx-003（不同包）
- xlsx-005 + docx-004/005（不同包）
- cross-001 + xlsx-003/docx-002（构建配置 vs 源码）
