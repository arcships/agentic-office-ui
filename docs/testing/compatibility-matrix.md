# 0.2.0 兼容矩阵

## 状态

`0.2.0` 候选版本已用同一批五个真实 `.tgz` 完成两套 Vue × 三种浏览器的本会话矩阵自测，六个组合均为 `PASS`，没有重试、`FLAKY` 或 `BLOCKED`。**候选尚未发布，本结果不是实现者之外的独立验收签字。** 候选发布前仍需全新会话按 `BB-RELEASE` 复核。

## 声明与验证范围

| 维度 | 公开声明或仓库基线 | `0.2.0` 候选验证范围 | 当前结果 |
|---|---|---|---|
| Node.js | 五个公开包暂未声明宽泛的 `engines` 范围 | `22.13.0` | PASS |
| pnpm | 工作区固定 `9.0.6` | `9.0.6`，仅用于构建和生成候选包 | PASS（候选构建） |
| Vue | 三个 Vue 包的 peer 范围为 `>=3.2.25 <4` | 下限 `3.2.25` 和仓库锁定的 `3.5.39` | PASS |
| Vite | 资源消费流程使用 Vite 6，公开包不声明 Vite peer | 仓库锁定的 `6.4.3` | PASS |
| TypeScript | 候选消费流程使用 TypeScript 5，公开包不声明 TypeScript peer | 仓库锁定的 `5.9.3` | PASS |
| Chromium | 候选发布必须运行核心用例 | Playwright `1.61.0` / Chromium `149.0.7827.55` | PASS |
| Firefox | 候选发布必须运行核心用例 | Playwright `1.61.0` / Firefox `151.0` | PASS |
| WebKit | 候选发布必须运行核心用例 | Playwright `1.61.0` / WebKit `26.5` | PASS |

上表只证明列出的候选材料和版本组合通过了本会话自测。没有列出的 Node、Vue、Vite、TypeScript 或浏览器版本仍为待验证，不从相邻版本自动推断兼容。

## 组合方式

为避免用无意义的全排列延长流水线，候选版本至少验证以下组合：

| 组合 | Node | Vue | Vite | TypeScript | 浏览器 |
|---|---:|---:|---:|---:|---|
| Vue 下限 + 当前工具链 | `22.13.0` | `3.2.25` | `6.4.3` | `5.9.3` | Chromium、Firefox、WebKit |
| 当前 Vue + 当前工具链 | `22.13.0` | `3.5.39` | `6.4.3` | `5.9.3` | Chromium、Firefox、WebKit |

正式矩阵固定为两套 Vue × 三种浏览器，共六个组合。局部调试必须显式设置 `COMPATIBILITY_ALLOW_PARTIAL=1`，其总结果只能是 `PARTIAL_PASS`，不能用于发布。如果 Vue 下限组合失败，要么修复产品，要么在公开清单和本文中提高下限；不能只删除失败行。若某个浏览器缺少固定二进制，结果记为 `BLOCKED`，候选版本不能据此发布。首次失败允许重试一次；重试通过记为 `FLAKY`，仍不能通过发布门禁。

## 每个组合的通过条件

每一行必须从五个候选 `.tgz` 开始，并完成：

1. 在工作区外安装，不使用 `workspace:`、符号链接或源码别名。
2. 不复制 `apps/demo/public` 中的 Worker/WASM。
3. TypeScript 与 Vue 类型检查通过。
4. Vite 正式构建通过，公开 JS、类型和三份 `style.css` 可解析。
5. 正式 preview 中 DOCX 有可见正文，XLSX 有可见网格并可选择已知单元格，PDF 可查看、翻页、缩放、旋转、使用缩略图、搜索和下载。
6. DOCX/XLSX Worker 与三份 WASM 来自候选包：安装文件哈希匹配 `.tgz` 清单，浏览器实际响应哈希匹配消费项目正式构建，无 404、MIME 错误或静默主线程回退。
7. 控制台、页面异常、失败请求、关键响应和下载使用统一规则采集。
8. 每个重要用例使用新的浏览器 context；首次失败最多重试一次，并将重试通过标为 `FLAKY`。

三种浏览器使用同一材料和用户侧断言。正常 PDF 必须能够查看、翻页、缩放、旋转、使用缩略图、搜索和下载；默认公开体积上限必须为 `50 MiB`，不能再叠加页数、像素或内存等隐藏上限，也不能把只下载验证字节当成 PDF 通过。

## 实际结果

| 组合 | 安装 | 类型 | 构建 | 核心黑盒 | 证据目录 | 结论 |
|---|---|---|---|---|---|---|
| Vue 3.2.25 / Chromium 149.0.7827.55 | PASS | PASS | PASS | PASS | `.../matrix/vue-minimum/browser-chromium/attempt-1/` | PASS |
| Vue 3.2.25 / Firefox 151.0 | PASS | PASS | PASS | PASS | `.../matrix/vue-minimum/browser-firefox/attempt-1/` | PASS |
| Vue 3.2.25 / WebKit 26.5 | PASS | PASS | PASS | PASS | `.../matrix/vue-minimum/browser-webkit/attempt-1/` | PASS |
| Vue 3.5.39 / Chromium 149.0.7827.55 | PASS | PASS | PASS | PASS | `.../matrix/current/browser-chromium/attempt-1/` | PASS |
| Vue 3.5.39 / Firefox 151.0 | PASS | PASS | PASS | PASS | `.../matrix/current/browser-firefox/attempt-1/` | PASS |
| Vue 3.5.39 / WebKit 26.5 | PASS | PASS | PASS | PASS | `.../matrix/current/browser-webkit/attempt-1/` | PASS |

证据目录使用 `output/acceptance/<commit>/<suite-id>/<run-id>/`，并记录操作系统、浏览器完整版本、npm 实际解析的 Vue/Vite/TypeScript 版本、语言、时区、视口、设备倍率、五个 `.tgz` 的 SHA-256、命令和退出码。浏览器首次运行与重试分别保存到 `attempt-1/` 和 `attempt-2/`。

本次完整证据位于：

`output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/BB-RELEASE/20260711T-arcships-scope-release-02/matrix/compatibility-matrix/`

总结果文件 `summary.json` 的 SHA-256 为 `ad0246a5989e3318ded2843bbc552d8ecfd1bd5259648ab3ea461bf8d82ed775`。同一发布自测生成的候选清单位于 `.../candidate/candidate-manifest.json`，SHA-256 为 `98cd784e1e05afb5206966ccc267ecb683361973511bb106c7720c28225c01c6`。五个候选包的 SHA-256 为：

- `@arcships/docx-core@0.2.0`：`8a3e77aa50a9e99421e9fc46dc41a13ba2da60071f2b035c57fb4f264a080b6b`
- `@arcships/xlsx-core@0.2.0`：`b9a55937cc382df105f63855c753aed69e3bd7abe7d6d51edc8646874f2baf97`
- `@arcships/vue-docx@0.2.0`：`3c21e915567753fdaf5857fe53313b19f85d1c6f64863ed932eece29f1dc4a46`
- `@arcships/vue-xlsx@0.2.0`：`bad6c8ac5570134b61a2d1d72826e55484646d26deb96ecb3a07aceda98fa91b`
- `@arcships/vue-extend@0.2.0`：`45eebffb31c8fefe7de852d91b45c1687c352423e5b910c3719391ff73d010b5`
