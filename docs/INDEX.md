# Documentation Index

## 当前执行入口

> [!IMPORTANT]
> 当前整改、测试和发布验收以本节四份文档为准。旧迁移计划、旧视觉验收记录和局部测试结果只用于了解背景，不能作为当前版本已经可以发布的证明。

| 文档 | 用途 |
|---|---|
| [architecture-review-and-target-design.md](architecture-review-and-target-design.md) | 记录整体审查结论、主要风险、目标架构和设计取舍 |
| [plan/stabilization-roadmap.md](plan/stabilization-roadmap.md) | 当前稳定化执行计划、优先级、阶段目标和完成标准 |
| [end-to-end-blackbox-test-plan.md](end-to-end-blackbox-test-plan.md) | 端到端与黑盒验收范围、用例、环境、证据和发布门槛 |
| [testing/agent-execution-runbook.md](testing/agent-execution-runbook.md) | Agent 可直接执行的测试步骤、结果记录方式和异常处理规则 |

## 历史迁移与补充设计文档

> 本节主要保存迁移、拆分、上游对齐和视觉验收资料。它们是历史背景或专项补充；其中的 `done`、截图、局部验收和功能清单，不代表当前代码已经通过完整发布验收。

| 文档 | 用途 |
|---|---|
| [visual-acceptance-handoff.md](visual-acceptance-handoff.md) | 历史视觉验收口径：路由、视口、交互和上游对齐要求；仅作补充参考 |
| [upstream-xlsx-feature-alignment.md](upstream-xlsx-feature-alignment.md) | XLSX 上游功能对齐清单 + 迁移操作方案（14 源文件，~41000 行） |
| [xlsx-migration-architecture.md](xlsx-migration-architecture.md) | XLSX 模块化重做架构设计（单文件 ≤1000 行，大文件拆分 + vue-xlsx 重写） |
| [upstream-docx-feature-alignment.md](upstream-docx-feature-alignment.md) | DOCX 上游功能对齐清单 + 迁移操作方案（8 子包，~68000 行） |
| [docx-migration-architecture.md](docx-migration-architecture.md) | DOCX 模块化重做架构设计（单文件 ≤1000 行，~124 文件规划） |
| [docx-editor-helpers-split-plan.md](docx-editor-helpers-split-plan.md) | editor.tsx 24953 行拆分为 24 模块的详细方案（分析员产出） |
| [upstream-extend-ui.md](upstream-extend-ui.md) | 上游 Extend UI 归属与 sync 策略 |
| [shadcn-components-browser-acceptance.md](shadcn-components-browser-acceptance.md) | shadcn 组件验收标准（PDF/Components 路由用） |
| [component-browser-verification-plan.md](component-browser-verification-plan.md) | 组件浏览器验证计划 |

## 历史交付计划与补充记录

> [plan/README.md](plan/README.md) 保存的是 2026-07-07 的迁移计划。当前执行请使用 [稳定化路线图](plan/stabilization-roadmap.md)。

| 文档 | 用途 |
|---|---|
| [plan/README.md](plan/README.md) | 2026-07-07 历史迁移计划总览 + 当时的任务依赖图 |
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

## 包当前状态（截至 2026-07-10 审查）

| 包 | 状态 | 说明 |
|---|---|---|
| `@extend-ai/xlsx-core` | ⚠️ 代码存在，发布受阻 | 核心代码和部分测试已存在；发布包缺少可消费的完整构建产物，完整黑盒门槛尚未通过 |
| `@extend-ai/vue-xlsx` | ⚠️ 范围受限 / 部分实现 | 控制器和组件代码已存在；功能完整性、性能、真实浏览器验收和发布包验证仍未通过 |
| `@extend-ai/docx-core` | ⚠️ 代码存在，运行与发布受阻 | 引擎、布局、查看和编辑相关代码已存在；正式构建中的 Worker/WASM 加载链及发布包仍有阻断问题 |
| `@extend-ai/vue-docx` | ⚠️ 范围受限 / 部分实现 | 查看器和编辑器组件已存在；渲染能力仍有占位与分叉，正式编辑页面仍有崩溃阻断 |
| `@extend-ai/vue-extend` | ⚠️ 范围受限 / 部分实现 | 组件代码已存在；PDF 原始 iframe 的安全边界和发布包完整性尚未通过验收 |
