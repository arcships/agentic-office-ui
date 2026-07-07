# Documentation Index

## 设计文档

| 文档 | 用途 |
|---|---|
| [visual-acceptance-handoff.md](visual-acceptance-handoff.md) | 验收标准：路由/视口/交互 gate，上游 parity 要求 |
| [upstream-xlsx-feature-alignment.md](upstream-xlsx-feature-alignment.md) | XLSX 上游功能对齐清单 + 迁移操作方案（14 源文件，~41000 行） |
| [upstream-docx-feature-alignment.md](upstream-docx-feature-alignment.md) | DOCX 上游功能对齐清单 + 迁移操作方案（8 子包，~68000 行） |
| [docx-migration-architecture.md](docx-migration-architecture.md) | DOCX 模块化重做架构设计（单文件 ≤1000 行，~124 文件规划） |
| [docx-editor-helpers-split-plan.md](docx-editor-helpers-split-plan.md) | editor.tsx 24953 行拆分为 24 模块的详细方案（分析员产出） |
| [upstream-extend-ui.md](upstream-extend-ui.md) | 上游 Extend UI 归属与 sync 策略 |
| [shadcn-components-browser-acceptance.md](shadcn-components-browser-acceptance.md) | shadcn 组件验收标准（PDF/Components 路由用） |
| [component-browser-verification-plan.md](component-browser-verification-plan.md) | 组件浏览器验证计划 |

## 交付计划

| 文档 | 用途 |
|---|---|
| [plan/README.md](plan/README.md) | 迁移计划总览 + 任务依赖图 |
| [plan/analysis/migration-split.md](plan/analysis/migration-split.md) | 任务拆分分析：模块分解 + 集成关系枚举 |
| [plan/tasks/](plan/tasks/) | 任务文件（xlsx-001~006, docx-001~006, cross-001, xlsx-integ, docx-integ） |
| [plan/backlog.md](plan/backlog.md) | 非阻塞 findings 记录 |

## Scope 关系

```
xlsx-core（引擎层）
  ↓ 被调用
vue-xlsx（controller + viewer）
  ↓ 被调用
demo（XlsxViewerPage）

docx-core（引擎层 + 布局 + 辅助 + editor helpers）
  ↓ 被调用
vue-docx（composables + editor + viewer）
  ↓ 被调用
demo（DocxViewerPage + DocxEditorPage）

vue-extend（shadcn 组件，无上游，不迁移）
demo（路由 + fixture + wasm 部署）
```

## 包导出状态

| 包 | 状态 | 说明 |
|---|---|---|
| `@extend-ai/xlsx-core` | ✅ done | 引擎层：wasm, safe-calculate, worker-client, xlsx-worker, colors, images, charts, types（React 类型已清理） |
| `@extend-ai/vue-xlsx` | pending | controller + viewer |
| `@extend-ai/docx-core` | 进行中 | 引擎层（engine/）已完成 ✅，布局/viewer/editor 层待迁移 |
| `@extend-ai/vue-docx` | pending | composables + components + render（stub 状态） |
