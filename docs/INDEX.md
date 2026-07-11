# 文档索引

## 当前执行入口

> [!IMPORTANT]
> 当前候选版本是 `0.2.0`，尚未发布。任务状态以稳定化路线图为准；发布结论必须来自同一批真实压缩包的 `BB-RELEASE` 和全新会话独立复核。

| 文档 | 用途 |
|---|---|
| [architecture-review-and-target-design.md](architecture-review-and-target-design.md) | 目标架构、包边界、状态所有权和兼容约束 |
| [plan/stabilization-roadmap.md](plan/stabilization-roadmap.md) | 唯一任务顺序、依赖、状态和执行证据来源 |
| [end-to-end-blackbox-test-plan.md](end-to-end-blackbox-test-plan.md) | 用户侧黑盒用例、通过标准和发布暂停条件 |
| [upstream-parity-gap-audit.md](upstream-parity-gap-audit.md) | DOCX/XLSX 与官方 Extend UI React 的固定基准、差异和收口状态 |
| [testing/agent-execution-runbook.md](testing/agent-execution-runbook.md) | 正式构建、浏览器、真实压缩包和证据的执行步骤 |
| [api/public-api-contract.md](api/public-api-contract.md) | 五个公开包的入口、类型、事件、错误、弃用和删除期限 |

## 格式扩展设计

| 文档 | 用途 |
|---|---|
| [pptx-preview-playback-design.md](pptx-preview-playback-design.md) | PPTX 静态渲染底座、播放层、包边界、实施顺序和验收门槛；尚未实施，不属于 `0.2.0` |

## 0.2.0 候选资料

| 文档 | 用途 |
|---|---|
| [../RELEASE_NOTES.md](../RELEASE_NOTES.md) | `0.2.0` 候选变更、限制、发布条件和回退办法 |
| [migration-0.2.md](migration-0.2.md) | 从 `0.1.x` 升级到 `0.2.0` 的代码与资源迁移 |
| [testing/compatibility-matrix.md](testing/compatibility-matrix.md) | Node、Vue、Vite、TypeScript 和三种浏览器的声明与待验证矩阵 |

## 当前包状态

以下状态更新于 2026-07-11，只描述代码和候选制品现状，不代表已经发布。

| 包 | 当前能力 | 发布状态 |
|---|---|---|
| `@arcships/docx-core` | 提供模型、布局、编辑命令、实例 Runtime，以及公开 Worker/WASM 入口 | `0.2.0` 候选；本会话 P4 自测通过，等待独立 `BB-RELEASE` |
| `@arcships/vue-docx` | 提供可用的 `DocxViewer`、`DocxEditor`、组合函数和公开样式入口 | `0.2.0` 候选；不承诺与上游全部功能完全一致 |
| `@arcships/xlsx-core` | 提供工作簿、公式、图表/图片数据、实例 Runtime，以及公开 Worker/WASM 入口 | `0.2.0` 候选；本会话 P4 自测通过，等待独立 `BB-RELEASE` |
| `@arcships/vue-xlsx` | 提供 `XlsxViewer`、控制器、公开样式，以及按需图表/地图/WebGL 入口 | `0.2.0` 候选；六组合矩阵自测通过，等待独立复核 |
| `@arcships/vue-extend` | 提供真实 PDF 查看、翻页、缩放、旋转、缩略图、搜索、下载和通用组件 | `0.2.0` 候选；PDF 默认仅有整份文件 `50 MiB` 上限 |

`@arcships/office-runtime` 仍是私有工作区包，不属于公开安装清单。

本会话已经用新命名空间的五个真实压缩包完成两次隔离构建、工作区外消费和六组合浏览器矩阵。实际发布仍暂停：本机 npm 凭证当前返回 `E401`，最终 GitHub 仓库、五个新包的首次发布和后续可信发布方式尚未确定；这些外部步骤不包含在本次改名自测中。

## 测试与发布资料

| 文档 | 用途 |
|---|---|
| [testing/performance-baseline.md](testing/performance-baseline.md) | 已批准的性能环境、数值和比较规则 |
| [visual-acceptance-handoff.md](visual-acceptance-handoff.md) | 历史视觉口径和专项补充，不替代当前黑盒方案 |
| [shadcn-components-browser-acceptance.md](shadcn-components-browser-acceptance.md) | 通用组件的专项浏览器检查 |
| [component-browser-verification-plan.md](component-browser-verification-plan.md) | 组件浏览器验证补充计划 |

## 历史迁移资料

> [!WARNING]
> 本节文件保存 2026-07-06 至 2026-07-07 的迁移分析。文件中的“当前状态”、`done`、`pending`、旧目录、旧截图和旧行数都只代表当时情况，不能覆盖本页“当前执行入口”中的文档。

| 文档 | 历史用途 |
|---|---|
| [upstream-xlsx-feature-alignment.md](upstream-xlsx-feature-alignment.md) | XLSX 上游功能清单与最初迁移方案 |
| [xlsx-migration-architecture.md](xlsx-migration-architecture.md) | XLSX 模块化迁移设计 |
| [upstream-docx-feature-alignment.md](upstream-docx-feature-alignment.md) | DOCX 上游功能清单与最初迁移方案 |
| [docx-migration-architecture.md](docx-migration-architecture.md) | DOCX 模块化迁移设计 |
| [docx-editor-helpers-split-plan.md](docx-editor-helpers-split-plan.md) | 旧编辑器辅助代码拆分计划 |
| [plan/README.md](plan/README.md) | 2026-07-07 的历史迁移任务表 |
| [plan/analysis/migration-split.md](plan/analysis/migration-split.md) | 历史任务拆分分析 |
| [plan/tasks/](plan/tasks/) | 历史任务文件 |
| [plan/reviews/](plan/reviews/) | 历史局部审查记录 |
| [plan/backlog.md](plan/backlog.md) | 非阻断历史记录 |

## 代码范围

```text
packages/docx-core/      DOCX 核心与 Runtime
packages/vue-docx/       DOCX Vue 组件
packages/xlsx-core/      XLSX 核心与 Runtime
packages/vue-xlsx/       XLSX Vue 组件与按需渲染
packages/vue-extend/     PDF 与通用 Vue 组件
apps/demo/               正式黑盒消费页
tests/                   单元、组件、黑盒和外部消费测试
scripts/ci/              固定检查入口和证据编排
```
