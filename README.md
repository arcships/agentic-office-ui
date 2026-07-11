<div align="center">

<img src="docs/assets/banner.png" alt="Agentic Office UI" width="100%">

# Agentic Office UI

Vue 3 组件库 · DOCX / XLSX / PPTX / PDF 浏览器内查看与编辑

[![CI](https://github.com/arcships/agentic-office-ui/actions/workflows/ci.yml/badge.svg)](https://github.com/arcships/agentic-office-ui/actions/workflows/ci.yml)
[![Release](https://github.com/arcships/agentic-office-ui/actions/workflows/release.yml/badge.svg)](https://github.com/arcships/agentic-office-ui/actions/workflows/release.yml)
![Vue](https://img.shields.io/badge/Vue-3.2%2B-42b883)
![License](https://img.shields.io/badge/license-Apache--2.0-blue)
![Status](https://img.shields.io/badge/status-candidate%200.2.0-orange)

</div>

## 概述

Agentic Office UI 提供一组面向 Vue 3 的 Office 文档组件，支持在浏览器中查看和编辑 DOCX、XLSX、PPTX 和 PDF。核心层使用 WASM + Worker 解析，Vue 层提供声明式组件和组合函数接口。

## 公开包

| 包 | 用途 |
|---|---|
| `@arcships/docx-core` | DOCX 模型、布局、命令、实例 Runtime、Worker/WASM |
| `@arcships/vue-docx` | `DocxViewer`、`DocxEditor` 和组合函数 |
| `@arcships/xlsx-core` | XLSX 模型、公式、图表、实例 Runtime、Worker/WASM |
| `@arcships/vue-xlsx` | `XlsxViewer`、按需图表/地图/WebGL |
| `@arcships/vue-pdf` | PDFium 引擎 PDF 查看器 |
| `@arcships/vue-ui` | 签名、上传、缩略图、引用框、版面组件 |

> `@arcships/pptx-core` 和 `@arcships/vue-pptx` 已提供完整的 PPTX 预览与播放实现（加载、浏览、搜索、动画、切换、媒体、全屏），尚未进入 `0.2.0` 公开合同。开发入口见 [PPTX 指南](docs/pptx-development-guide.md)。

## 快速开始

```bash
pnpm install --frozen-lockfile
pnpm typecheck
pnpm build
```

发布后按需安装：

```bash
pnpm add @arcships/docx-core@0.2.0 @arcships/vue-docx@0.2.0
pnpm add @arcships/xlsx-core@0.2.0 @arcships/vue-xlsx@0.2.0
pnpm add @arcships/vue-pdf@0.2.0 @arcships/vue-ui@0.2.0
```

Vue 包导入样式：

```ts
import "@arcships/vue-docx/style.css"
import "@arcships/vue-xlsx/style.css"
import "@arcships/vue-pdf/style.css"
import "@arcships/vue-ui/style.css"
```

Worker 和 WASM 通过公开入口交给构建工具处理：

```ts
import docxWorkerUrl from "@arcships/docx-core/worker?worker&url"
import docxWasmUrl from "@arcships/docx-core/assets/docx_wasm_bg.wasm?url"
import { createDocxRuntime } from "@arcships/docx-core/runtime"

const runtime = createDocxRuntime({ workerUrl: docxWorkerUrl, wasmUrl: docxWasmUrl })
```

## 当前状态

六个公开包候选版本为 `0.2.0`，**尚未正式发布**。仓库内自测不替代候选发布复核；只有同一批真实压缩包完成 `BB-RELEASE` 并由全新会话独立复核后才能给出发布结论。

核心文档：

- [目标架构](docs/architecture-review-and-target-design.md)
- [稳定化路线图](docs/plan/stabilization-roadmap.md)
- [黑盒验收方案](docs/end-to-end-blackbox-test-plan.md)
- [Agent 执行手册](docs/testing/agent-execution-runbook.md)
- [公开接口合同](docs/api/public-api-contract.md)
- [完整文档索引](docs/INDEX.md)

## PDF 行为

`PdfViewer` 使用受控 PDF 引擎显示真实页面，支持查看、翻页、缩放、旋转、缩略图、搜索和下载。唯一的公开文件拒绝上限是整份文件体积 `maxFileSize`，默认 `50 MiB`，超过返回 `PDF_TOO_LARGE` 且不启动引擎。

DOCX 和 XLSX 另有实例级输入、压缩包、XML、图片、模型、历史记录和运行时间限制。错误码与配置项以[公开接口合同](docs/api/public-api-contract.md)为准。

## 验证命令

```bash
pnpm typecheck          # 全仓类型检查
pnpm build              # 构建
pnpm test:unit          # 单元测试
pnpm test:component     # 组件测试
pnpm test:blackbox      # 浏览器黑盒
pnpm test:stress        # 压力测试
pnpm test:performance   # 性能基线
pnpm test:consumer      # 工作区外 .tgz 消费
pnpm test:docs          # 文档契约
pnpm check              # 仓库门禁
```

浏览器结论必须来自正式构建 preview。真实发布包结论必须来自工作区外安装的 `.tgz`。兼容范围见[兼容矩阵](docs/testing/compatibility-matrix.md)。

## 版本资料

- [0.2.0 候选发布说明](RELEASE_NOTES.md)
- [从 0.1.x 迁移到 0.2.0](docs/migration-0.2.md)
- [公开接口与弃用期限](docs/api/public-api-contract.md)

## 上游归属

本项目面向 Vue 实现，设计参考公开的 Extend UI / Extend AI React 包。原项目、包名、接口思路和设计方向属于 Extend 及其维护者。详细来源和同步规则见 [docs/upstream-extend-ui.md](docs/upstream-extend-ui.md)。
