# 项目稳定化整改路线图

> 状态：`pending`
>
> 基线日期：2026-07-10
>
> 审查基线提交：`41913ca03666aa99302d13924e038db17e478480`
>
> 适用范围：`packages/docx-core`、`packages/vue-docx`、`packages/xlsx-core`、`packages/vue-xlsx`、`packages/vue-extend`、`apps/demo`、计划新增的工作区私有包 `@arcships/office-runtime`，以及它们的构建、测试和发布流程。

## 1. 文档目的

本计划把项目整体审查中发现的问题转成可执行任务，并规定每个任务和每个阶段必须提供的验证证据。计划的首要目标不是继续增加功能，而是先让源码、类型检查、正式构建、浏览器运行和发布包五个层面的结果保持一致。

本计划只记录整改顺序和完成标准，具体设计与黑盒用例分别见：

- [整体架构审查与目标设计](../architecture-review-and-target-design.md)
- [端到端与黑盒验收测试方案](../end-to-end-blackbox-test-plan.md)
- [Agent 黑盒验收执行手册](../testing/agent-execution-runbook.md)
- [现有 DOCX 迁移架构](../docx-migration-architecture.md)
- [现有 XLSX 迁移架构](../xlsx-migration-architecture.md)

若上述文档之间发生冲突，处理顺序如下：

1. 安全边界和本计划的发布暂停条件优先。
2. 目标架构负责规定长期边界。
3. 黑盒验收方案负责规定用户可见行为和执行证据。
4. 旧迁移文档只用于了解历史实现，不得覆盖新的验收要求。

## 2. 当前状态

当前项目已有工作区拆分、严格 TypeScript 配置、DOCX/XLSX 测试材料和部分集成验证，但尚未达到可稳定发布的状态。

| 方面 | 当前事实 | 影响 |
|---|---|---|
| DOCX 正式运行 | Worker 中未可靠获得 WASM 地址，正式构建会请求错误资源；Viewer 可解析失败 | 开发环境或局部测试通过不能代表正式页面可用 |
| DOCX Editor | 示例页构造的模型缺少必填字段，正式页面可能直接崩溃 | 核心示例链路不可验收 |
| 类型检查 | Vue 包和 demo 的 `typecheck` 仍使用普通 `tsc` | `.vue` 模板和脚本中的错误可漏过检查 |
| 发布包 | 五个包的入口指向 `dist`，但发布预检不能保证包含入口、类型、样式、Worker 和 WASM | 本仓库内可构建，不等于外部项目可安装使用 |
| 样式 | `vue-docx`、`vue-xlsx` 的样式没有稳定的公开入口 | 外部使用者可能得到无样式或样式路径不稳定的组件 |
| URL 安全 | PDF 原始地址和 XLSX 下载地址缺少统一的协议与来源限制 | 不可信地址可能触发脚本执行或加载非预期内容 |
| 文件安全 | 解压总量、压缩比例、图片像素、页数、工作表规模和内存缺少统一预算 | 恶意或超大文件可能耗尽主线程和内存 |
| 请求并发 | 文件快速切换、取消、Worker 回包先后顺序缺少统一约束 | 旧请求可能覆盖新请求，资源也可能泄漏 |
| DOCX 渲染 | Viewer 和 Editor 存在不同渲染路径，部分能力仍是占位实现 | 两条路径会继续漂移，修复成本成倍增加 |
| 自动测试 | 缺少根级单元、组件、正式构建黑盒和发布包消费测试命令 | 当前检查不能作为合并和发布依据 |
| 用户界面与交互 | 现有门禁主要证明 DOCX/XLSX 能加载、编辑和导出，没有把上游 Extend UI React 控件的可见外壳与真实交互作为硬标准 | 技术检查通过仍可能交付不能让真实用户使用的界面；此前 P4 发布自测结论不能作为候选发布结论 |
| XLSX 性能 | 地图和重型图表静态进入主包；可见区定位存在重复线性计算 | 首屏包体和大表滚动成本偏高 |
| 文档 | 迁移状态、旧任务和真实代码状态并不完全一致 | Agent 和开发者容易依据过期结论继续工作 |

本表只描述审查基线，不表示任何整改任务已经完成。下文所有任务初始状态均为 `pending`。

审查时工作区已有一处不属于本计划的未提交修改：`packages/docx-core/tsup.config.ts`。基线结论没有把这处修改视为已完成整改，后续 Agent 也不得覆盖它。

### 2.1 基线复现记录

以下记录来自本轮整体审查，用于说明为什么暂停发布。P0 实施者必须在基线提交或其可追溯后继提交上重新执行，并把完整输出保存到证据目录。

| 检查 | 基线结果 | P0 复现时必须保存的证据 |
|---|---|---|
| `pnpm typecheck` | 通过，但普通 `tsc` 漏掉 `.vue` 中的真实错误 | 普通 `tsc` 与 `vue-tsc --noEmit` 的退出码和完整日志 |
| `pnpm build` | 通过，但出现 WASM 路径和大包警告 | 构建日志、产物清单和资源 URL |
| 正式预览 `/docx-viewer` | Worker 请求错误 WASM 资源，解析失败 | 页面截图、控制台、Worker/WASM 请求和响应类型 |
| 正式预览 `/docx-editor` | 初始模型字段不完整，页面崩溃 | 页面截图、控制台堆栈和触发步骤 |
| 五个包打包预检 | 声明入口指向 `dist`，压缩包没有可消费的完整入口与类型 | 每个包的文件清单、入口解析结果和独立安装结果 |
| XLSX 结构验证 | 结构检查通过，但出现 Vue 生命周期警告 | 命令日志和控制台警告 |

### 2.2 统一验收编号与证据位置

黑盒文档必须为下列编号提供具体命令、用例、输入材料和断言。路线图中的“黑盒方案”均指这些编号，不接受只有自然语言描述的替代记录。

| 编号 | 内容 | 默认覆盖范围 |
|---|---|---|
| `BB-P0-ROUTES` | 正式构建路由冒烟 | `/`、`/docx-viewer`、`/docx-editor`、`/xlsx-viewer`、`/pdf-viewer`、`/components` |
| `BB-PDF-VIEWER` | PDF 真实查看、操作、体积上限和错误恢复 | P3-PDF-BUDGET-01、P3-PDF-01 |
| `BB-DOCX-WORKER` | DOCX Worker/WASM 开发、正式构建与故障路径 | P0-DOCX-01、P0-PACK-02 |
| `BB-SEC-URL` | PDF、XLSX 下载和外部资源 URL 安全 | P0-URL-01、P0-URL-02 |
| `BB-RACE` | 快速切换、取消、卸载与旧请求晚到 | P0-RACE-01、P1-RACE-TEST-01、P2-RACE-01 |
| `BB-PACK-CONSUMER` | 五个真实压缩包的独立消费项目 | P0-PACK-01、P0-STYLE-01、P0-PACK-02、P1-PACK-TEST-01 |
| `BB-PERF-XLSX` | XLSX 包体、按需加载和大表滚动 | P3-PERF-BASELINE-01、P3-XLSX-LOAD-01、P3-XLSX-SCROLL-01 |
| `BB-STRESS` | DOCX/XLSX 资源超限、内存、长时间切换和极端输入 | P3-BUDGET-01、P3-STRESS-01 |
| `BB-UX-PARITY` | DOCX/XLSX 与上游 Extend UI React 的视觉和交互一致性 | P4-UX-DOCX-01、P4-UX-XLSX-01 |
| `BB-RELEASE` | 候选制品全量验收 | P4-MATRIX-01、P4-RC-01 |

本地证据统一保存到 `output/acceptance/<commit>/<suite-id>/<run-id>/`；持续集成使用相同层级作为制品名称。每次运行至少写入环境说明、命令、开始与结束时间、退出码、测试材料校验值和结果摘要。`output/` 不提交仓库，任务状态记录必须填写可访问的持续集成链接或证据包校验值。

控制台允许清单默认是空。确需保留的 warning 必须登记精确文本、来源、负责人和到期日，不能使用正则吞掉一类错误。性能设备、数据集和数值预算由 P3-PERF-BASELINE-01 产出并写入黑盒文档；在该任务达到 `ready_for_gate` 前，相关性能任务不得开始。

## 3. 整改目标

整改完成后，项目必须同时满足以下结果：

1. 全新检出后，可以使用锁定依赖完成安装、类型检查、构建和全部自动测试。
2. 正式构建中的 DOCX、XLSX、PDF 主要页面能够完成加载、查看、编辑和错误处理，控制台没有未处理异常。
3. DOCX Worker 与 WASM 在开发和正式构建中使用同一套显式配置；失败时给出可判断的错误，不静默退回主线程掩盖故障。
4. 五个公开发布包可从实际生成的压缩包安装到独立 Vue 项目，并能导入 JavaScript、类型、样式及所需资源。`@arcships/office-runtime` 保持 `private: true`，不作为第六个发布包；所需实现必须打入公开包产物，公开 `.d.ts` 不能要求消费端安装它。
5. 不可信 URL 和文件在进入解析、iframe、下载或图片解码前经过统一检查。
6. 解析任务具有资源上限、取消能力和“只接受最新请求结果”的规则。
7. DOCX 只读与编辑共用同一渲染基础，差异仅位于交互和编辑命令层。
8. XLSX 的可选地图、图表和 WebGL 能力按需加载，大表滚动达到黑盒方案规定的性能基线。
9. 端到端和黑盒验收可以由 Agent 在无人手工点选的情况下重复执行，并保存日志、截图、网络记录和测试结果。
10. 文档只保留一套当前状态和一套执行入口，历史文档明确标记为历史资料。
11. DOCX/XLSX 的工具栏、分页或工作表导航、缩放、选择、编辑、状态反馈和主要文件操作达到上游 Extend UI React 控件的实际使用水平；不能用 DOM 存在、接口同名或能加载文件代替视觉与交互验收。

## 4. 发布暂停条件

从本计划建立起，默认暂停公开发布和版本升级。只要出现以下任一情况，就不得发布候选版本：

- 任一正式构建页面存在崩溃、空白页、未处理异常或资源 404。
- DOCX Worker 或 WASM 依赖静默回退主线程才能成功。
- `vue-tsc`、构建、单元测试、组件测试或正式构建黑盒测试未通过。
- 浏览器控制台出现未在允许清单中的 error 或 warning。
- 任一发布包缺少声明的入口、类型、样式、Worker、WASM 或许可证文件。
- 任一公开包在运行时或类型中要求消费端额外安装私有的 `@arcships/office-runtime`。
- 实际发布压缩包未在独立消费项目中安装验证。
- PDF、下载地址或外部资源仍能接受危险协议或绕过来源检查。
- DOCX/XLSX 文件没有资源预算，或超限后仍继续解压、解析、解码。
- 快速切换文件时，旧请求仍可能覆盖最新结果。
- 黑盒测试缺少可重复执行命令或关键证据。
- DOCX/XLSX 仍明显是原生表单控件、表情符号按钮或验证页拼接，或者缺少上游已有的核心导航、缩放、搜索、选择和编辑反馈。
- 已知阻断问题仅写入文档，但没有修复和回归测试。
- 工作区存在无法解释的依赖漏洞、构建产物差异或公开接口破坏。

解除暂停需要满足“阶段 P4 发布准备”的全部检查，并由实现者之外的验证者复核证据。任何单项“本地通过”都不能单独解除暂停。

## 5. 优先级与状态定义

### 5.1 优先级

| 优先级 | 含义 | 处理规则 |
|---|---|---|
| P0 阻断 | 已导致正式运行、类型检查、安全边界或发布包不可用 | 暂停新增功能，优先完成；未全部通过不得进入架构重构 |
| P1 可信检查 | 建立能持续发现回归的自动检查 | 与 P0 后半段可有限并行，但必须先覆盖已修复的问题 |
| P2 架构边界 | 消除全局配置、重复渲染和核心层越界 | 以 P1 测试作为保护后实施 |
| P3 安全与性能 | 完成资源限制、并发治理、按需加载和大文件性能 | 依赖稳定接口与可重复基准 |
| P4 发布准备 | 收敛公开接口、文档和真实发布包流程 | 前四阶段全部通过后执行 |

### 5.2 状态

| 状态 | 含义 |
|---|---|
| `pending` | 尚未开始，或没有足够证据证明完成 |
| `in_progress` | 已分配并正在修改或验证 |
| `ready_for_gate` | 任务自身的产出、验收和独立复核已通过，可以满足同阶段后续任务的依赖；仍未通过阶段全量检查 |
| `blocked` | 依赖、外部条件或明确决策阻止继续；必须记录原因和下一步 |
| `done` | 产出、任务验收、阶段全量检查和独立复核均已通过 |

同阶段依赖在前置任务达到 `ready_for_gate` 后即可继续，避免“必须全阶段完成才能开始下一个任务”的循环等待。跨阶段任务仍以之前阶段全量检查通过为前提；任务表明确写出的例外除外。

状态更新不得只提供部分证据。从 `in_progress` 进入 `ready_for_gate` 时，必须完整登记负责人、开始与结束时间、基线提交、结果提交、测试命令、运行环境、退出码、证据位置和独立验证者。从 `ready_for_gate` 进入 `done` 时，必须再登记阶段全量检查的提交、退出码和证据位置。仅提交代码不能改变状态。

每次分派或状态变化都在本文末尾“执行记录”追加一行；尚未分派的任务保持 `pending`，负责人留空。

## 6. 分阶段任务

### 6.1 P0：消除运行与发布阻断

| ID | 范围 | 依赖 | 产出 | 验收条件 | 状态 |
|---|---|---|---|---|---|
| P0-DOCX-01 | `docx-core`、`vue-docx`、demo 的 Worker/WASM 加载 | 无 | 实例级 DOCX 运行配置；稳定的 Worker/WASM 资源入口；结构化错误 | 开发与正式构建都能在真实 Worker 中解析测试 DOCX；WASM 请求成功且类型正确；错误地址能显示明确错误；测试禁止以主线程回退冒充通过 | `ready_for_gate` |
| P0-DOCX-02 | demo 的 DOCX Editor 初始模型和公共模型创建方式 | P0-DOCX-01 | 完整模型工厂或统一默认值；示例页改用公共入口 | `vue-tsc` 不再报告模型字段缺失；Editor 正式页面能打开、输入、撤销和保存；缺字段输入有明确校验 | `ready_for_gate` |
| P0-TYPE-01 | `vue-docx`、`vue-xlsx`、`vue-extend`、demo | 无 | 所有 Vue 项目的 `typecheck` 改为 `vue-tsc --noEmit`；根命令覆盖全部包 | 人为加入一个 `.vue` 类型错误时检查会失败，移除后全仓通过；类型检查在构建和持续集成之前执行 | `ready_for_gate` |
| P0-PACK-01 | 五个可发布包的清单、脚本与构建产物 | P0-TYPE-01 | `files`、`exports`、发布前构建脚本和产物清单；逐包压缩包预检 | 五个实际压缩包均包含并能解析声明的 JS 与类型入口；没有工作区源码路径和缺失文件；压缩包内容快照纳入测试 | `ready_for_gate` |
| P0-STYLE-01 | `vue-docx`、`vue-xlsx`、`vue-extend` 样式出口 | P0-PACK-01 | 稳定的 `./style.css` 或等价公开出口；使用说明 | 独立消费项目只能通过公开出口加载样式；构建后主要组件样式生效；不依赖仓库内部 `dist` 路径 | `ready_for_gate` |
| P0-PACK-02 | Worker、WASM 与浏览器资源的发布方式 | P0-DOCX-01、P0-PACK-01 | 资源复制或 URL 构造规则；消费端配置说明；资源存在性测试 | 从实际压缩包安装后，Vite 开发与正式预览均能加载 Worker/WASM；路径不依赖 demo 根目录；资源 404 会使测试失败 | `ready_for_gate` |
| P0-URL-01 | `vue-extend` PDF 地址与 iframe/渲染入口 | 无 | 协议与初始来源白名单；最终 `Response.url` 复检；PDF 内外部资源加载规则；浏览器直连默认用 `redirect: 'error'` 禁止重定向；原始文档不进入无隔离同源 iframe | `javascript:`、`data:text/html`、伪装 PDF、白名单外初始/最终来源、任何浏览器直连重定向和不允许的外部资源均按 `BB-SEC-URL` 失败；合法 Blob、File 和允许的 HTTPS 地址可加载；如业务必须跟随重定向，只能通过能逐跳检查 `Location` 的受控同源服务端代理 | `ready_for_gate` |
| P0-URL-02 | `vue-xlsx` 下载、复制和外部资源地址 | 无 | 协议与初始来源白名单；最终 `Response.url` 复检；下载 URL 校验；外部图片/图表资源规则；浏览器直连默认禁止重定向 | `javascript:`、非预期 `data:`、控制字符、白名单外初始/最终来源和任何浏览器直连重定向不会被点击、执行或加载；合法 Blob 下载保留文件名并可释放 URL；若允许重定向，只能由受控同源服务端代理逐跳检查；`BB-SEC-URL` 通过 | `ready_for_gate` |
| P0-RACE-01 | DOCX/XLSX 当前加载控制器的最低并发正确性 | P0-DOCX-01 | 每实例请求编号、取消入口和只接受最新结果的最小实现 | `BB-RACE` 中慢请求晚到、连续切换 20 次和组件卸载用例通过；旧结果不能覆盖最新状态；本任务只建立正确性底线，统一实现留给 P2-RACE-01 | `ready_for_gate` |
| P0-DEMO-01 | 六个 demo 正式构建路由 | P0-DOCX-01、P0-DOCX-02、P0-URL-01、P0-URL-02、P0-RACE-01 | `BB-P0-ROUTES` 固定路由清单与正式预览冒烟测试 | 六个明确路由都能直接访问和刷新；关键内容可见；网络无关键 404；控制台允许清单为空且无未处理 error/warning | `ready_for_gate` |

#### P0 全量检查

P0 只有在以下检查全部通过后才能结束：

1. 在全新检出环境执行锁定依赖安装，不依赖已有 `node_modules` 或旧 `dist`。
2. 运行根级 `vue-tsc` 类型检查和全仓构建。
3. 从正式构建启动预览，执行 `BB-P0-ROUTES`、`BB-DOCX-WORKER`、`BB-SEC-URL` 和 `BB-RACE`。
4. 对五个包执行真实打包，不只执行文件列表预览；将压缩包安装到独立 Vue + Vite + TypeScript 项目。
5. 验证独立项目能够导入组件、类型、样式和资源，并完成正式构建与浏览器冒烟测试。
6. 按统一证据位置保存测试日志、关键页面截图、浏览器控制台记录和 Worker/WASM 网络记录。
7. 运行 `git diff --check`，确认没有意外生成物和越界修改。

### 6.2 P1：建立可信的自动检查

| ID | 范围 | 依赖 | 产出 | 验收条件 | 状态 |
|---|---|---|---|---|---|
| P1-CI-01 | 根脚本与持续集成 | P0-TYPE-01 | `typecheck`、`test`、`test:component`、`test:e2e`、`test:pack` 等统一入口；持续集成配置 | 全新检出可用一组固定命令完成检查；任一子任务失败都会让持续集成失败；任务可定位到具体包和用例 | `ready_for_gate` |
| P1-UNIT-01 | DOCX/XLSX 核心解析、布局、编辑、序列化 | P1-CI-01 | 可重复的核心层单元与集成测试 | 正常、损坏、空文件、取消、并发、往返和关键格式用例有明确断言；测试不允许只检查函数存在 | `ready_for_gate` |
| P1-VUE-01 | 三个 Vue 包 | P1-CI-01 | 组件和组合函数测试 | 属性、事件、加载、错误、空状态、卸载清理、对象 URL 释放和 Worker 终止均有断言 | `ready_for_gate` |
| P1-FIXTURE-01 | `test-data` 或统一测试材料目录 | P1-UNIT-01 | 测试材料清单、来源说明、预期行为、校验值和生成脚本 | 每个材料可追溯；生成结果稳定；正常、边界、恶意和大文件分类完整；敏感数据不进入仓库 | `ready_for_gate` |
| P1-E2E-01 | demo 正式构建端到端流程 | P0-DEMO-01、P1-CI-01 | 可由 Agent 直接运行的浏览器测试 | 测试只访问正式构建预览；覆盖打开、切换、编辑、撤销、保存、下载、错误恢复；失败自动保留截图和跟踪文件 | `ready_for_gate` |
| P1-CONSOLE-01 | 浏览器控制台、页面错误、请求失败 | P1-E2E-01 | 统一采集器和极小允许清单 | 未捕获异常、关键请求失败、Vue 警告和未列入清单的控制台 warning/error 会直接使测试失败 | `ready_for_gate` |
| P1-RACE-TEST-01 | DOCX/XLSX 连续加载与卸载 | P0-RACE-01、P1-E2E-01 | 把 `BB-RACE` 固化为持续集成回归套件 | 覆盖慢请求后发先至、连续切换 20 次、组件中途卸载；页面最终只展示最新文件且无遗留 Worker/URL；测试会在移除 P0 最小保护后稳定失败 | `ready_for_gate` |
| P1-PACK-TEST-01 | 五个发布包的消费端矩阵 | P0-PACK-02、P1-CI-01 | 自动创建临时消费项目并安装本地压缩包的测试 | ESM、TypeScript、样式和资源路径均由公开出口验证；测试结束清理临时目录；不使用 workspace 链接 | `ready_for_gate` |
| P1-DOC-TEST-01 | 测试文档与实际命令 | P1-CI-01、P1-E2E-01 | 黑盒方案中的命令、目录和报告格式与仓库保持一致 | 新 Agent 只阅读文档即可运行；不存在“文档写了但脚本不存在”的命令；失败排查路径清楚 | `ready_for_gate` |

#### P1 全量检查

P1 只有在 P0 全量检查继续通过，并额外满足以下要求时才能结束：

1. 单元、组件、端到端、发布包消费测试全部接入同一持续集成流程。
2. 正式构建黑盒套件连续运行三次，结果一致，不依赖用例顺序。
3. 断网可执行的测试不访问外部服务；必须联网的检查明确隔离并记录原因。
4. 测试失败时，报告包含用例、输入材料、浏览器版本、截图、控制台和网络证据。
5. 验证者至少人为制造一次类型错误、资源 404、控制台异常和包入口缺失，确认检查能阻止合并。
6. 测试材料校验脚本和所有自动测试均通过，仓库没有未说明的快照更新。

### 6.3 P2：收敛架构边界

| ID | 范围 | 依赖 | 产出 | 验收条件 | 状态 |
|---|---|---|---|---|---|
| P2-RUNTIME-01 | DOCX/XLSX/PDF 共享加载流程与私有 `@arcships/office-runtime` | P1-UNIT-01、P1-VUE-01 | `private: true` 的薄共享包；明确的加载上下文：来源、资源地址、限制、取消信号、请求编号和错误映射 | 各组件不再各自拼接安全与并发规则；相同错误得到一致错误类型；共享包不依赖 Vue/DOM、不持有全局任务状态；不进入公开发布清单 | `ready_for_gate` |
| P2-CONFIG-01 | Worker/WASM 和解析选项 | P2-RUNTIME-01 | 实例级配置与生命周期；移除模块级可变全局状态 | 同页两个不同配置的实例互不污染；销毁后资源释放；并行测试不会因执行顺序改变结果 | `ready_for_gate` |
| P2-RACE-01 | 所有异步加载、Worker 消息与保存流程 | P2-RUNTIME-01、P1-RACE-TEST-01 | 取消、请求编号和只接受最新结果的统一实现 | 旧请求完成后不能改写新状态；取消能终止可终止的工作；所有完成、失败、取消路径都会清理监听器和对象 URL | `ready_for_gate` |
| P2-DOCX-RENDER-01 | DOCX Viewer 与 Editor | P1-E2E-01、P2-RUNTIME-01 | 唯一文档渲染面；只读与编辑模式共享页面、段落、表格、图片和页眉页脚渲染 | 同一文档在两种模式下的非交互视觉结果一致；表格正文等占位实现被补齐或明确降级；旧重复路径不再承载新功能 | `ready_for_gate` |
| P2-DOCX-EDIT-01 | DOCX 选区、命令、撤销重做和浮层 | P2-DOCX-RENDER-01 | 独立于基础渲染的编辑交互层 | 关闭编辑能力后没有编辑监听器；开启后编辑命令只修改模型；撤销重做和选区恢复有单元与黑盒覆盖 | `ready_for_gate` |
| P2-CORE-01 | `docx-core`、`xlsx-core` | P2-RUNTIME-01 | 纯数据模型、布局、解析适配和编辑命令边界；依赖方向检查 | 核心包不导入 Vue、浏览器组件或 demo；浏览器能力通过显式适配注入；公开入口不暴露内部文件路径 | `ready_for_gate` |
| P2-API-01 | 五个包的公开类型、事件和错误 | P2-CORE-01、P2-DOCX-EDIT-01 | 最小公开接口清单、弃用策略、错误说明和兼容分类登记 | 黑盒测试只使用公开入口；内部重排不要求消费端修改；已导出的 `setWasmSource`、内部组件等公开 API 在 `0.2.0` 标记弃用并保留整个 `0.x`，最早在 `1.0.0` 删除；从未公开的内部桥接另设删除日期 | `ready_for_gate` |
| P2-DEAD-01 | 占位实现、重复入口和未使用兼容代码 | P2-DOCX-RENDER-01、P2-CORE-01 | 删除清单或明确的受支持降级清单 | 不再用空字符串、固定 `false` 或空事件处理器表示“已支持”；所有保留降级均有用户可见提示和测试 | `ready_for_gate` |

#### P2 全量检查

P2 只有在 P0、P1 全量检查继续通过，并额外满足以下要求时才能结束：

1. 依赖方向检查确认核心层不反向依赖 Vue、DOM 或 demo。
2. 同一 DOCX 在 Viewer 与 Editor 只读状态的结构快照和关键截图对比通过。
3. 两个不同运行配置的实例并行加载，互不覆盖 Worker/WASM 地址和资源限制。
4. 快速切换、取消、卸载和错误恢复测试连续执行，不出现旧结果覆盖与资源泄漏。
5. 公开接口清单和破坏性变化记录经过独立复核；所有黑盒用例只使用公开接口。
6. 删除旧路径前，现有支持能力全部有等价测试；无法等价的能力必须保留清楚的风险说明，不能静默删除。

### 6.4 P3：安全与性能治理

| ID | 范围 | 依赖 | 产出 | 验收条件 | 状态 |
|---|---|---|---|---|---|
| P3-BUDGET-01 | DOCX/XLSX 压缩包、XML 与解析过程 | P2-RUNTIME-01 | 输入大小、文件数、解压总量、单项大小、压缩比例、页数、行列数和共享字符串等限制；归档路径、XML 结构和 MIME 校验 | 路径穿越名称、畸形 XML、错误 MIME、边界值、略超限和组合超限由 `BB-STRESS` 覆盖；每一项在超限前或尽可能早地停止；错误指出具体限制 | `ready_for_gate` |
| P3-IMAGE-01 | DOCX/XLSX 图片和预览 | P3-BUDGET-01 | 图片字节、宽高、总像素、总解码量和并发解码限制 | 畸形图片、超大尺寸和大量小图不会导致页面失去响应；拒绝后释放临时资源；正常高分辨率样本仍可用 | `ready_for_gate` |
| P3-PDF-BUDGET-01 | PDF 整份文件体积 | P0-URL-01、P2-RUNTIME-01 | 单一公开 `maxFileSize` 配置，默认 `50 MiB`，宿主可调整且不叠加第二层隐藏硬上限 | 默认值进入公开合同；使用较小测试值验证等于上限可打开、超过上限返回 `PDF_TOO_LARGE` 且不创建引擎或页面；不再把页数、单页像素、总内存、并发页或复杂边界列为 PDF 阻塞预算 | `ready_for_gate` |
| P3-PDF-01 | PDF 渲染与隔离 | P0-URL-01、P2-RUNTIME-01 | 使用受控 PDF 渲染能力；不再依赖直接加载原始文档的同源 iframe | 伪装内容不能执行脚本；合法 PDF 的翻页、缩放、搜索等已承诺能力通过；跨源和错误响应按策略处理；整份文件超过当前 `maxFileSize` 时不显示 | `ready_for_gate` |
| P3-CACHE-01 | 文档缓存、历史记录、缩略图和对象 URL | P2-RACE-01、P3-BUDGET-01 | 按字节和数量限制的实例缓存；统一释放规则 | 达到上限后按规则回收；组件销毁后无永久全局缓存；长时间切换文件时内存回落到基线允许范围 | `ready_for_gate` |
| P3-PERF-BASELINE-01 | 五个包体积、浏览器资源与交互性能 | P2-API-01、P1-FIXTURE-01 | 固定 Chromium 版本、1440×900 视口、设备倍率 1、测试数据校验值、运行轮数、包体/JS/WASM/PDF 资源字节预算、帧耗时、长任务、交互延迟和内存回落数值 | 黑盒文档登记明确数值和测量命令；在同一提交连续三轮的结果可重复；`BB-PERF-XLSX` 与 `BB-STRESS` 能读取这些预算并在越界时失败 | `ready_for_gate` |
| P3-XLSX-LOAD-01 | 地图、图表、WebGL 和可选依赖 | P2-API-01、P3-PERF-BASELINE-01 | 按功能动态加载的分包；主入口不静态引入地图数据 | 不使用图表/地图时网络不请求相关分包；启用功能后按需成功加载；主包和功能分包满足 `BB-PERF-XLSX` 的字节预算 | `ready_for_gate` |
| P3-XLSX-SCROLL-01 | `XlsxGrid` 可见区定位与绘制 | P1-E2E-01、P3-PERF-BASELINE-01 | 行列累计偏移索引、二分定位、帧内合并绘制和测量工具 | 大表滚动不再从第 0 行列重复线性查找；固定设备和数据集上的帧耗时、长任务和交互延迟满足 `BB-PERF-XLSX` 数值 | `ready_for_gate` |
| P3-WORKER-01 | 大文件解析和高成本计算 | P2-CONFIG-01、P3-BUDGET-01 | 明确的 Worker 执行范围、超时和终止策略 | 大文件解析期间主线程仍可交互；超时或取消可终止 Worker；不允许因为 Worker 故障静默在主线程重跑大任务 | `ready_for_gate` |
| P3-STRESS-01 | 长时间运行与极端输入 | P3-IMAGE-01、P3-PDF-01、P3-CACHE-01、P3-XLSX-SCROLL-01、P3-WORKER-01 | `BB-STRESS` 压力测试脚本和基准报告 | 连续加载、切换和编辑达到方案规定次数；Worker、监听器、对象 URL 不持续增长；内存回落满足 P3-PERF-BASELINE-01 数值；结果可重复比较 | `ready_for_gate` |

#### P3 全量检查

P3 只有在 P0 至 P2 全量检查继续通过，并额外满足以下要求时才能结束：

1. `BB-STRESS` 覆盖压缩炸弹、路径穿越名称、畸形 XML、超大图片、危险 URL、错误 MIME 和重定向。
2. 每种超限输入都在规定时间内失败，页面保持可操作，并显示不泄露内部路径的可理解错误。
3. `BB-PERF-XLSX` 使用 P3-PERF-BASELINE-01 固定的浏览器、视口、测试材料、运行方式和至少三轮数据，报告记录中位数和最差值。
4. XLSX 主包、功能分包和 DOCX 资源都有明确字节预算；PDF 只检查公开 `maxFileSize`，默认 `50 MiB`，宿主可调整。超过当前生效值会让对应持续集成用例失败。
5. 大表滚动、文档连续切换和长时间运行测试通过；浏览器进程、Worker、对象 URL 和监听器没有持续增长。
6. 安全修复不能通过禁用正常功能冒充通过；合法边界样本和可恢复错误路径必须同时验证。

### 6.5 P4：发布准备与文档收敛

| ID | 范围 | 依赖 | 产出 | 验收条件 | 状态 |
|---|---|---|---|---|---|
| P4-API-01 | 五个包公开入口和版本策略 | P2-API-01、P3-BUDGET-01、P3-IMAGE-01、P3-PDF-BUDGET-01、P3-PDF-01、P3-CACHE-01、P3-PERF-BASELINE-01、P3-XLSX-LOAD-01、P3-XLSX-SCROLL-01、P3-WORKER-01、P3-STRESS-01 | 最终导出清单、兼容说明、弃用项、兼容分类登记和版本选择 | 导出与压缩包内容一致；公开 API 在 `0.2.0` 给出弃用警告、迁移文档和替代入口，整个 `0.x` 保持可用且最早 `1.0.0` 删除；内部桥接按日期清理；不存在未文档化的深层导入 | `ready_for_gate` |
| P4-SCOPE-01 | 五个公开包和工作区私有运行包的组织命名空间 | P4-API-01 | 公开包统一迁移为 `@arcships/*`；内部依赖、源码导入、构建配置、测试、文档和发布脚本同步更新；上游来源和版权记录保持原样 | 五个真实 tgz 的包名、依赖、JS、类型、样式、Worker 和 WASM 均来自 `@arcships/*`；工作区产品代码与当前文档不再引用六个旧本地包名；真实工作区外消费和正式浏览器发布套件通过；不执行真实发布 | `ready_for_gate` |
| P4-PACK-01 | 五个公开包的实际发布压缩包 | P4-API-01、P4-SCOPE-01、P1-PACK-TEST-01 | 可复现的五个压缩包及内容报告；私有共享实现的内联检查 | 两次干净构建的有效内容一致；包内无源码密钥、绝对路径、测试材料或无关大文件；入口全部可用；运行时代码和 `.d.ts` 均不要求安装 `@arcships/office-runtime` | `ready_for_gate` |
| P4-MATRIX-01 | 消费端和浏览器组合 | P4-PACK-01 | Vue/Vite/TypeScript 支持范围与浏览器矩阵报告 | 每个声明支持的组合完成安装、类型检查、构建和核心黑盒流程；不支持组合在文档中明确说明 | `ready_for_gate` |
| P4-DOC-01 | README、文档索引、迁移文档和任务状态 | P4-API-01、P4-SCOPE-01 | 单一当前状态页；安装、资源、安全限制、错误与排查说明；历史标记 | 新读者能从 `docs/INDEX.md` 找到当前架构、路线图和黑盒方案；过期结论不再被标成“当前”；示例命令实际可执行 | `ready_for_gate` |
| P4-UX-DOCX-01 | `vue-docx` Viewer/Editor 与正式消费页 | P4-API-01、P4-SCOPE-01 | 对照上游官方源码提交和浏览器状态的 DOCX 工具栏、页码、缩放、缩略图、文件操作、编辑反馈与响应式外壳；可重复截图和交互断言 | `BB-UX-PARITY` 的 DOCX 用例在正式构建和真实 tgz 消费页通过；上游已有的核心用户路径不得缺失或以字符、表情符号和原生未统一控件冒充；首次失败和重试证据完整 | `ready_for_gate` |
| P4-UX-XLSX-01 | `vue-xlsx` Viewer/Editor 与正式消费页 | P4-API-01、P4-SCOPE-01 | 对照上游官方源码提交和浏览器状态的工作簿工具栏、搜索、缩放、公式栏、网格、选择、编辑、工作表标签与状态反馈；可重复截图和交互断言 | `BB-UX-PARITY` 的 XLSX 用例在正式构建和真实 tgz 消费页通过；网格主要交互达到上游控件实际使用水平；首次失败和重试证据完整 | `ready_for_gate` |
| P4-FIDELITY-DOCX-01 | `vue-docx` Viewer 批注与修订 | P2-DOCX-RENDER-01、P4-UX-DOCX-01 | 普通 `DocxViewer` 直接从公开模型展示修订标记、批注高亮和页边卡片；公开开关与 Demo 专用样例 | Viewer 不依赖编辑控制器即可展示；开关真实改变页面；正式构建 Demo 样例可重复验证；组件、类型、构建和浏览器证据完整 | `completed` |
| P4-FIDELITY-XLSX-WORKER-01 | `vue-xlsx` Worker 图片与合并单元格 | P3-WORKER-01、P4-UX-XLSX-01 | Worker 加载快照带回实例所属的图片资源和精确合并区域；网格与图片层在 Worker/主线程路径显示一致；Demo 可切换验证 | 只读 Worker 模式的图片、合并文字、尺寸和选择锚点与主线程一致；对象地址可释放；正式构建 Demo、组件、类型和构建证据完整 | `completed` |
| P4-FIDELITY-CONTENT-02 | DOCX/XLSX 剩余内容展示补齐 | P4-FIDELITY-DOCX-01、P4-FIDELITY-XLSX-WORKER-01 | DOCX 脚注、尾注正文和全文搜索；XLSX 单元格超链接、批注、条件格式、迷你图、形状和只读表单控件；正式 Demo 可直接查看 | 五类缺口均使用公开模型或实例数据；无测试后门；类型、构建、组件和正式 preview 黑盒通过，证据含截图、网络、控制台和页面异常 | `completed` |
| P4-RELEASE-01 | 发布流水线 | P4-PACK-01、P4-MATRIX-01、P4-DOC-01、P4-UX-DOCX-01、P4-UX-XLSX-01 | 带检查、制品留存、版本和发布说明的流程 | 只能发布已通过相同提交全部检查的制品；发布步骤不重新生成未测试内容；失败不会产生半发布状态 | `pending` |
| P4-RC-01 | 候选版本最终验收 | P4-RELEASE-01 | 候选版本、`BB-RELEASE` 完整证据包和签字记录 | 所有发布暂停条件已消除；P0 至 P4 检查全部通过；实现者之外的验证者复核；已知剩余问题都有非阻断依据 | `pending` |

#### P4 全量检查

P4 是发布前最终检查，必须完整执行而不是引用旧结果：

1. 从全新检出开始安装、类型检查、构建、单元、组件、端到端、安全、压力和发布包消费测试。
2. 使用将要发布的真实压缩包执行浏览器矩阵，不得改用工作区链接或源码别名。
3. 核对包清单、入口、类型、样式、Worker、WASM、许可证、版本和发布说明。
4. 正式构建六个路由和独立消费项目的控制台、网络和页面错误均达到黑盒方案标准。
5. 对比性能和包体预算，没有未经批准的退化。
6. 文档中的所有安装与运行命令由新环境复跑；所有内部链接有效。
7. 保存不可修改的制品校验值、测试报告、浏览器证据和对应提交号。
8. 两名角色完成复核：一名负责代码与架构，一名负责黑盒与发布包；两者不得都是该任务实现者。
9. 使用固定上游官方提交逐页对照 DOCX Viewer、DOCX Editor 和 XLSX，保存桌面与窄屏的初始、就绪、交互、错误、恢复截图；核心操作逐项断言，不能只比较像素或只检查元素存在。

## 7. Agent 并行分工

### 7.1 分工原则

- 同一时间并行的任务必须有清楚且尽量不重叠的文件范围。
- 实现 Agent 不得自行把自己的任务标为 `done`；必须交给独立验证 Agent。
- 验证 Agent 优先使用公开入口和正式构建，不读取实现细节来“帮助”测试通过。
- 每个 Agent 开始前先检查工作区已有修改，保留其他人的改动，不执行破坏性回退。
- 跨包接口由一个集成负责人确定；其他 Agent 不得各自增加不同的临时接口。
- 任一任务发现范围外阻断时，只记录证据并通知负责人，不顺手大改相邻模块。

### 7.2 建议并行批次

同一批次表示“在各自任务依赖已达到 `ready_for_gate` 且文件范围不重叠时可以并行”，不表示忽略任务表中的依赖。

| 批次 | 可并行任务 | 前置条件 | 批次结束检查 |
|---|---|---|---|
| A1：P0 基础 | P0-DOCX-01、P0-TYPE-01、P0-URL-01、P0-URL-02 | 各任务依赖为“无”；分派前排除重叠配置文件 | 各任务独立验收，达到 `ready_for_gate` |
| A2：P0 衔接 | P0-DOCX-02、P0-PACK-01、P0-RACE-01 | 对应 A1 前置任务达到 `ready_for_gate` | 各任务独立验收 |
| A3：P0 产物 | P0-STYLE-01、P0-PACK-02 | P0-PACK-01 及各自其他依赖达到 `ready_for_gate` | `BB-PACK-CONSUMER` |
| A4：P0 集成 | P0-DEMO-01 | 表内全部前置任务达到 `ready_for_gate` | P0 全量检查，通过后把 P0 任务改为 `done` |
| B1：检查入口 | P1-CI-01 | P0 全量检查通过 | 根命令能在干净环境启动 |
| B2：主体测试 | P1-UNIT-01、P1-VUE-01、P1-E2E-01、P1-PACK-TEST-01 | P1-CI-01 达到 `ready_for_gate`，其余按任务依赖 | 各套件独立验收 |
| B3：检查收口 | P1-FIXTURE-01、P1-CONSOLE-01、P1-RACE-TEST-01、P1-DOC-TEST-01 | 对应 B2 前置任务达到 `ready_for_gate` | P1 全量检查 |
| C1：加载边界 | P2-RUNTIME-01 | P1 全量检查通过 | 共享接口独立验收 |
| C2：并行收敛 | P2-CONFIG-01、P2-RACE-01、P2-DOCX-RENDER-01、P2-CORE-01 | P2-RUNTIME-01 达到 `ready_for_gate`，并由接口负责人锁定共享类型 | 各任务独立验收 |
| C3：架构收口 | P2-DOCX-EDIT-01、P2-API-01、P2-DEAD-01 | 对应 C2 前置任务达到 `ready_for_gate` | P2 全量检查 |
| D1：预算基线 | P3-BUDGET-01、P3-PDF-BUDGET-01、P3-PERF-BASELINE-01 | P2 全量检查通过，且各自表内依赖满足 | 预算写入黑盒文档并可自动读取 |
| D2：安全性能实现 | P3-IMAGE-01、P3-PDF-01、P3-CACHE-01、P3-XLSX-LOAD-01、P3-XLSX-SCROLL-01、P3-WORKER-01 | 对应表内前置任务达到 `ready_for_gate`；P3-PDF-01 可与 P3-PDF-BUDGET-01 并行，二者都要在 P3 全量检查前完成；重叠核心文件拆成顺序任务 | `BB-PERF-XLSX` 与相关安全用例 |
| D3：压力收口 | P3-STRESS-01 | 表内全部前置任务达到 `ready_for_gate` | P3 全量检查 |
| E1：公开接口 | P4-API-01 | P3 全量检查通过 | 公开接口复核 |
| E2：产物与文档 | P4-PACK-01、P4-DOC-01 | P4-API-01 达到 `ready_for_gate`，文件范围不重叠 | 压缩包清单和文档命令复核 |
| E3：消费矩阵 | P4-MATRIX-01 | P4-PACK-01 达到 `ready_for_gate` | `BB-PACK-CONSUMER`、浏览器矩阵 |
| E4：候选版本 | P4-RELEASE-01、P4-RC-01 | 按任务表顺序，不能互相并行 | P4 全量检查和 `BB-RELEASE` |

建议为每个批次保留一个集成 Agent，其职责仅是解决已经确认的接口冲突、运行全量检查和汇总证据，不承担大块功能实现。

### 7.3 分派范围与记录

下面的路径是分派时的默认上限。负责人必须再根据具体任务缩小到文件或子目录；如果确需越界，先在执行记录中写明原因和批准者。

| 任务类别 | 默认允许路径 | 默认禁止范围 |
|---|---|---|
| DOCX 运行、模型、渲染、编辑 | `packages/docx-core/`、`packages/vue-docx/`、对应 `apps/demo/src/pages/Docx*.vue` | XLSX、PDF 和无关 demo 页面 |
| XLSX URL、加载和滚动 | `packages/xlsx-core/`、`packages/vue-xlsx/`、对应 `apps/demo/src/pages/Xlsx*.vue` | DOCX、PDF 和无关 demo 页面 |
| PDF 安全、预算和渲染 | `packages/vue-extend/`、对应 `apps/demo/src/pages/Pdf*.vue` | DOCX、XLSX 和其他组件 |
| 类型、打包、样式与公开入口 | 根 `package.json`、对应包 `package.json`、构建配置、公开入口和样式文件 | 不相关的业务实现；已有未提交文件除非明确列入 |
| 测试与持续集成 | 根脚本、持续集成配置、测试目录、测试材料目录、必要的 demo 测试入口 | 为让测试通过而修改产品行为；产品缺陷另派实现任务 |
| 共享运行层和核心边界 | 目标架构明确列出的共享包、`docx-core`、`xlsx-core` 公开边界 | Vue 组件视觉与 demo 功能，除非任务明确列入 |
| 文档与发布 | `docs/`、README、包清单、发布配置 | 核心业务实现 |

提示模板中的占位符不得原样发送。分派人必须先填写下列记录，未填全的任务不得进入 `in_progress`：

| 字段 | 必填内容 |
|---|---|
| `TASK_ID` | 任务表中的唯一 ID |
| `OWNER` / `VERIFIER` | 实现者和独立验证者，不能是同一人 |
| `BASE_COMMIT` | 开始工作的完整提交号；这只是回退起点，不等同于安全版本 |
| `ALLOWED_PATHS` | 从上表缩小后的精确文件或目录 |
| `FORBIDDEN_PATHS` | 至少包含所有范围外包和已存在的用户修改 |
| `SUITE_NAME` | 第 2.2 节中的一个或多个 `BB-*` 编号；纯核心任务填写对应单元套件编号 |
| `COMMANDS` | 黑盒文档或根脚本中真实存在的命令，不得写计划中的虚构命令 |
| `EVIDENCE_DIR` | `output/acceptance/<commit>/<suite-id>/<run-id>/` 的实际值 |
| `ROLLBACK_OWNER` | 集成负责人；安全任务另加安全复核人 |
| `COMPAT_KIND` | `public-api`、`internal-bridge` 或“不适用”。五个包已经导出的入口、组件、属性、事件和资源入口属于 `public-api`；从未导出且消费端不可见的阶段性适配才属于 `internal-bridge` |
| `COMPAT_INTRODUCED` | 公开 API 填最初发布版本和开始弃用版本；内部桥接填引入日期与提交 |
| `COMPAT_REPLACEMENT` | 明确替代入口或删除后使用方式；无兼容代码填“不适用” |
| `REMOVE_IN_VERSION` | 仅 `public-api` 填写。现有公开 API 在 `0.2.0` 开始弃用、整个 `0.x` 保留，最早填 `1.0.0`；内部桥接和无兼容代码填“不适用” |
| `REMOVE_BY_DATE` | 仅 `internal-bridge` 填明确日期，最晚为合并后 30 天或下一个次版本发布日中的较早者；公开 API 和无兼容代码填“不适用” |

### 7.4 实现任务提示模板

```text
你的任务是实现 {TASK_ID}。

先完整阅读：
1. docs/plan/stabilization-roadmap.md 中该任务、依赖和阶段检查；
2. docs/architecture-review-and-target-design.md 中相关边界；
3. docs/end-to-end-blackbox-test-plan.md 中相关黑盒用例。

工作范围：{ALLOWED_PATHS}
禁止修改：{FORBIDDEN_PATHS}
基线提交：{BASE_COMMIT}
验收套件：{SUITE_NAME}
必须执行：{COMMANDS}
证据目录：{EVIDENCE_DIR}
兼容分类：{COMPAT_KIND}
兼容引入：{COMPAT_INTRODUCED}
替代方式：{COMPAT_REPLACEMENT}
公开 API 删除版本：{REMOVE_IN_VERSION}
内部桥接删除日期：{REMOVE_BY_DATE}
已有改动属于其他工作，必须保留。

要求：
- 先复现问题并记录最小证据；
- 只实现任务范围内的修复；
- 为失败场景和正常场景增加自动测试；
- 运行任务验收命令和当前已经可运行的阶段回归子集；阶段全量检查由集成负责人在全部任务达到 ready_for_gate 后执行；
- 不得用跳过测试、静默回退或放宽断言让检查通过；
- 不得自行把任务标为 done。

交付时报告：修改文件、设计选择、测试命令与结果、剩余风险、需要验证者重点检查的内容。
```

### 7.5 独立验证提示模板

```text
你是 {TASK_ID} 的独立验证者，不是实现者。

只依据以下材料验证：
- docs/plan/stabilization-roadmap.md 的任务验收与阶段检查；
- docs/end-to-end-blackbox-test-plan.md 的用例和证据要求；
- 实际提交差异与真实发布/构建产物。

验证步骤：
1. 检查修改是否越界、是否隐藏错误或弱化测试；
2. 在干净环境复跑任务测试；
3. 至少加入一个可恢复的故障，证明检查确实能失败；
4. 对用户可见行为使用正式构建和公开入口验证；
5. 核对日志、截图、网络、控制台和产物清单；
6. 再运行当前已经可运行的阶段回归子集；若本任务是阶段最后一个任务，再由集成负责人运行阶段全量检查。

输出结论只能是 PASS 或 BLOCKED。BLOCKED 必须给出复现命令、证据位置、影响和最小修复建议。没有完整证据时不得给 PASS。
```

### 7.6 黑盒执行提示模板

```text
你负责执行 {SUITE_NAME} 黑盒套件。

必须从干净构建开始，只操作公开页面、公开包入口和用户可见控件；不要调用内部函数伪造结果。严格按 docs/end-to-end-blackbox-test-plan.md 执行。

记录：提交号、操作系统、浏览器版本、服务地址、测试材料校验值、开始/结束时间、命令退出码。失败时保留截图、页面内容、控制台、网络请求和跟踪文件。遇到产品缺陷时停止修改产品代码，只报告可复现证据。
```

### 7.7 文档收敛提示模板

```text
你负责 P4-DOC-01。逐项核对源码、package.json、实际命令和发布压缩包，不得从旧文档复制未经验证的“已完成”结论。

将当前入口、架构、限制、测试和发布方法收敛到索引可发现的位置；旧文档保留时明确标注历史状态和替代文档。最后检查所有相对链接，并让一个没有项目背景的 Agent 根据文档完成一次安装与黑盒运行。
```

## 8. 每次合并的共同检查

阶段检查之外，每个任务合并前至少完成以下共同检查：

1. 同阶段任务依赖至少达到 `ready_for_gate`，跨阶段依赖所在阶段已完成全量检查；或有记录证明本任务可以独立进行。
2. 修改只覆盖任务声明范围，没有覆盖他人的未提交工作。
3. 类型检查、受影响包构建、相关单元/组件测试通过。
4. 涉及用户页面、Worker、WASM、URL、样式或打包时，必须运行对应正式构建黑盒用例。
5. 新增错误分支同时有失败用例和恢复用例。
6. 没有新增未说明的控制台 warning/error、资源 404、定时器、Worker、监听器或对象 URL 泄漏。
7. 文档、公开类型和示例与实现保持一致。
8. `git diff --check` 通过，生成文件策略清楚。
9. 独立验证者给出 PASS，证据可以由另一名 Agent 重跑；此时任务进入 `ready_for_gate`，并可满足同阶段后续依赖。
10. 当前阶段所有任务均达到 `ready_for_gate` 后，在合并后的同一提交执行阶段全量检查；通过后才批量改为 `done`。

## 9. 完成定义

单个任务满足以下前七项并通过独立复核后进入 `ready_for_gate`；再满足最后一项后才算 `done`：

- 任务表中的产出实际存在，且没有使用仓库内部路径冒充公开能力。
- 所有验收条件都有自动检查或可重复的黑盒步骤。
- 正常、边界、失败、取消和恢复路径按任务性质得到覆盖。
- 相关正式构建和真实发布产物已经验证。
- 测试证据记录提交号、命令、环境、退出码和产物位置。
- 实现者之外的 Agent 已复核，未发现越界修改、弱断言或静默降级。
- 相关文档已经更新，且没有把未完成能力描述为已支持。
- 当前阶段的全量检查在合并后的主线上通过。

整个稳定化计划满足以下全部条件才算完成：

- P0 至 P4 所有任务均为 `done`，不存在未解释的 `blocked`。
- 发布暂停条件全部解除。
- 黑盒全套、消费端矩阵、安全输入集和性能基准均通过。
- 候选版本制品与被测试制品具有相同校验值。
- 当前状态、架构、测试和发布说明已收敛到文档索引。
- 剩余问题均有明确的非阻断理由、负责人和后续时间，不影响安全、数据正确性、正式运行和安装使用。

## 10. 风险与回退

### 10.1 回退决策与步骤

当前审查基线含已知阻断问题，因此 `41913ca...` 只是复现起点，不是“安全版本”。第一个可称为阶段安全点的提交，必须是 P0 全量检查通过的提交；此后每个阶段检查通过时，在执行记录中更新 `LAST_SAFE_COMMIT`。

出现以下任一情况时，集成负责人必须暂停后续合并并发起回退判断：

- 合并后 P0 路由、类型、构建、真实压缩包消费或安全用例从通过变为失败。
- 新提交导致数据错误、脚本执行风险、无法取消的主线程阻塞或公开接口无说明破坏。
- 包体、帧耗时、长任务、内存回落等超过 P3-PERF-BASELINE-01 的硬预算。
- 同一失败在干净环境稳定复现两次，且不能在当前小提交内安全修复。

回退由该批次集成负责人提出，独立验证者批准；URL、文件解析、PDF 隔离等安全任务还需要安全复核人批准。实现者不能单独决定保留自己的失败提交。紧急安全风险可先停止发布和禁用受影响公开入口，但仍需补齐记录。

标准步骤如下：

1. 在执行记录中填写失败提交、`LAST_SAFE_COMMIT`、复现命令、证据目录、影响范围和决策人。
2. 停止依赖该提交的 Agent，不覆盖它们的工作区。
3. 优先为最小失败提交创建普通回退提交；不使用破坏性重置。
4. 回退后至少运行受影响任务测试、共同检查、`BB-P0-ROUTES`、相关 `BB-*` 套件和当前阶段全量检查。
5. 只有回退后检查恢复通过，才更新 `LAST_SAFE_COMMIT` 或恢复后续合并。
6. 保留失败证据和回退提交，建立新的修复任务；不得删除失败测试。

兼容代码必须先分类，不能使用同一删除期限：

- `public-api`：五个公开包已经导出的 `setWasmSource`、内部组件、属性、事件、错误和资源入口均属于公开契约。新实例 API 计划在 `0.2.0` 引入并开始弃用旧入口；旧入口在整个 `0.x` 保持可用并给出开发警告，最早在 `1.0.0` 删除。删除前必须有替代入口、迁移文档、发布说明，以及真实压缩包对旧、新入口的消费测试。记录 `COMPAT_INTRODUCED`、`COMPAT_REPLACEMENT` 和 `REMOVE_IN_VERSION`，不得用 30 天规则提前删除。
- `internal-bridge`：只为分阶段重构新增、从未导出且消费端不可见的内部桥接。记录引入日期、提交、负责人、替代方式和 `REMOVE_BY_DATE`；删除日期不得晚于合并后 30 天或下一个次版本发布日中的较早者。到期前未完成替换就暂停后续发布，不能无限延期。
- 无兼容代码：`COMPAT_KIND` 和其余兼容字段统一填“不适用”，不能留空造成误判。

### 10.2 风险表

| 风险 | 早期信号 | 控制措施 | 回退方式 |
|---|---|---|---|
| Worker/WASM 配置调整破坏现有调用方 | 开发可用但独立消费项目资源 404 | 先增加压缩包消费测试，再切换入口；已导出的 `setWasmSource` 按 `public-api` 登记并保留到最早 `1.0.0`，未导出的桥接才填写 `REMOVE_BY_DATE` | 使用普通回退提交恢复旧公开入口，同时保留新测试和故障证据；不得恢复静默主线程回退 |
| DOCX 统一渲染引起视觉或编辑回归 | Viewer/Editor 对比截图或结构快照出现差异 | 按段落、表格、图片、页眉页脚分批迁移，每批执行双路径对比 | 回退当前最小迁移提交，保留已通过的基础组件；不得一次回退其他已验证模块 |
| 安全限制误伤合法大文件 | 边界样本被拒绝或错误信息不可操作 | 默认值由真实材料和压力数据确定，限制可由实例配置但必须有硬上限 | 回退单项阈值配置，不移除整个检查；补充合法边界回归材料 |
| PDF 渲染替换造成能力缺失 | 搜索、选择、缩放或打印等已有行为失效 | 先列出承诺能力并建立黑盒基线，再替换加载方式 | 只回退到安全隔离的上一实现；禁止回退到可执行不可信同源内容的方式 |
| 按需加载造成首次使用失败 | 图表或地图首次打开出现空白、竞态或分包 404 | 为加载、错误和重试状态增加组件与正式构建测试 | 临时合并较小的安全分包，不恢复全部静态进入主包；记录包体影响 |
| 滚动优化改变坐标或选择逻辑 | 可见区、冻结行列、合并单元格或命中测试偏移 | 使用固定大表和坐标断言，逐步替换定位算法 | 回退单个算法提交，保留测量工具与失败样本 |
| 发布包清单过严漏掉必要资源 | 独立消费项目构建或运行失败 | 以实际压缩包为唯一测试对象，快照记录允许文件 | 恢复明确缺失文件并补测试，不使用宽泛 `files` 规则打入整个仓库 |
| 并行 Agent 修改相同入口产生冲突 | 同一文件出现相互覆盖的接口和临时适配 | 提前分配路径和接口负责人；小批次合并后立即跑共同检查 | 停止后续合并，由集成负责人基于已验证提交逐个重放；不使用破坏性重置 |
| 测试只在特定机器通过 | 多次运行结果不同、依赖缓存或外网 | 固定版本、材料校验值、端口与时区；清楚区分离线和联网测试 | 回退不稳定测试的合并资格，先修复确定性；不得直接标记为长期跳过 |
| 文档再次落后于代码 | 命令不存在、状态与包清单不一致 | P4 文档任务读取实际产物；持续集成检查链接和示例命令 | 恢复最后一个已验证文档版本，并在同一修复中更新实现差异 |

回退的基本原则是“小提交、可定位、保留证据”。优先使用新的回退提交，不使用破坏性 Git 命令覆盖工作区，也不以删除测试或放宽验收条件作为回退。

## 11. 计划维护规则

1. 本文是稳定化工作的任务状态来源；旧任务表若仍保留，必须链接到本文并注明用途。
2. 每次状态变化都记录日期、负责人、提交号和证据位置。
3. 新发现的问题先判断是否触发发布暂停条件，再归入最近的阶段；不得无期限放入“以后处理”。
4. 任务范围或公开接口发生实质变化时，先更新目标架构和黑盒验收，再修改实现。
5. 每个阶段结束后由新 Agent 做一次读者检查：只阅读四份核心文档，说明它将如何执行下一阶段；无法正确复述依赖和检查时，先修文档。
6. 所有任务在建立本计划时均为 `pending`。只有真实执行并提供完整证据后，才允许更新状态。

## 12. 执行记录

当前 `LAST_SAFE_COMMIT`：无。审查基线存在已知阻断问题，不能登记为安全点。

### 12.1 分派记录

任务进入 `in_progress` 前先在下表追加一行。路径可以用逗号分隔的精确列表；命令较长时可填写黑盒文档中的命令编号，但该编号必须已经存在。

| 任务 ID | 负责人 | 验证者 | `BASE_COMMIT` | `ALLOWED_PATHS` | `FORBIDDEN_PATHS` | `SUITE_NAME` | `COMMANDS` | `EVIDENCE_DIR` | `ROLLBACK_OWNER` | `COMPAT_KIND` | `COMPAT_INTRODUCED` | `COMPAT_REPLACEMENT` | `REMOVE_IN_VERSION` | `REMOVE_BY_DATE` |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| P0-DOCX-01 | Codex `/root` | Codex `/root/quick_p0_verifier` | `41913ca03666aa99302d13924e038db17e478480` | `packages/docx-core/src/engine/wasm.ts`、`packages/docx-core/src/viewer/`、`packages/docx-core/src/index.ts`、`packages/vue-docx/src/` 中 DOCX 加载公开入口、`apps/demo/src/main.ts`、`apps/demo/src/pages/Docx*.vue`、`apps/demo/index.html`（消除正式路由 favicon 404）、`packages/vue-docx/tests/verify-integration.mjs`、`docs/plan/stabilization-roadmap.md` | XLSX/PDF 相关包和页面；`README.md`、`docs/INDEX.md`、`docs/plan/README.md` 的现有用户改动；`packages/docx-core/tsup.config.ts` 的现有用户改动（已阅读差异，只做资源 loader 最小合并） | `BB-DOCX-WORKER`、`BB-P0-ROUTES`、`BB-SEC-URL`（回归） | `pnpm typecheck`；`pnpm build`；`node packages/vue-docx/tests/verify-integration.mjs`；`node packages/vue-xlsx/test/structure.mjs`；`git diff --check`；正式 preview + Python Playwright | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/{BB-DOCX-WORKER,BB-P0-ROUTES,BB-SEC-URL}/20260710T121023+0800/`、`.../BB-DOCX-WORKER/20260710T134000+0800-quick-verifier/` | Codex `/root` | `public-api` | 现有 `setWasmSource`，在 `0.2.0` 起弃用 | 新的显式 DOCX Runtime 配置与公开诊断入口 | `1.0.0`（最早） | 不适用 |
| P0-DOCX-02 | Codex `/root` | Codex `/root/quick_p0_verifier` | `41913ca03666aa99302d13924e038db17e478480` | `packages/docx-core/src/editor/helpers/model-validation.ts`、`packages/docx-core/src/editor/helpers/index.ts`、`packages/vue-docx/src/composables/useDocxEditor.ts`、`packages/vue-docx/src/components/{DocxEditor,DocxToolbar,DocxParagraphHost}.vue`、`packages/vue-docx/src/render/static-html.ts`、`apps/demo/src/pages/DocxEditorPage.vue`、`packages/vue-docx/tests/verify-integration.mjs`、`docs/plan/stabilization-roadmap.md` | DOCX Worker/WASM 实现、XLSX/PDF 包和页面；`README.md`、`docs/INDEX.md`、`docs/plan/README.md` 的现有用户改动；`packages/docx-core/tsup.config.ts` 的现有用户改动 | `BB-P0-ROUTES`、DOCX Editor 正式 preview | `vue-tsc -p apps/demo/tsconfig.json --noEmit`；`pnpm typecheck`；`pnpm build`；`node packages/vue-docx/tests/verify-integration.mjs`；`node packages/vue-xlsx/test/structure.mjs`；`git diff --check`；正式 preview + Python Playwright | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/BB-P0-ROUTES/20260710T132252+0800/`（修复前失败）、`.../20260710T133709+0800/`（实现者通过）、`.../20260710T140943+0800-final-verifier/`（独立通过） | Codex `/root` | `public-api` | 现有 `UseDocxEditorOptions.starterModel` | `createBlankDocumentModel` 与显式模型校验错误 | `1.0.0`（最早） | 不适用 |
| P0-TYPE-01 | Codex `/root` | Codex `/root/typecheck_audit` | `41913ca03666aa99302d13924e038db17e478480` | `package.json`、`apps/demo/{package.json,tsconfig.typecheck.json}`、`packages/vue-docx/{package.json,tsconfig.typecheck.json}`、`packages/vue-xlsx/{package.json,tsconfig.typecheck.json}`、`packages/vue-extend/package.json`、`pnpm-lock.yaml`（仅因新增 demo 的既有开发工具依赖而更新）、`docs/plan/stabilization-roadmap.md` | 所有业务源码、DOCX Worker/WASM、XLSX/PDF 实现；`README.md`、`docs/INDEX.md`、`docs/plan/README.md`、`packages/docx-core/tsup.config.ts` 的现有用户改动 | 根 `pnpm typecheck`、临时 Vue 类型错误探针 | `pnpm --filter demo exec vue-tsc --version`；`pnpm typecheck`；临时 `.vue` 错误的失败/恢复检查；`pnpm build`；`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/P0-TYPE-01/20260710T141925+0800/`（实现者）与 `.../20260710T142145+0800-independent/`（独立验证） | Codex `/root` | 不适用 | 不适用 | 不适用 | 不适用 | 不适用 |
| P0-URL-01 | Codex `/root` | Codex `/root/sec_url_harness` | `41913ca03666aa99302d13924e038db17e478480` | `packages/vue-extend/src/{components/PdfViewer.vue,types.ts,index.ts,pdf-url-policy.ts}`、`apps/demo/src/pages/PdfViewerPage.vue`、`docs/plan/stabilization-roadmap.md` | DOCX、XLSX、Worker/WASM、其他 demo 页面；`README.md`、`docs/INDEX.md`、`docs/plan/README.md`、`packages/docx-core/tsup.config.ts` 的现有用户改动 | `BB-SEC-URL`、`BB-P0-ROUTES` | `pnpm typecheck`；`pnpm build`；正式 preview + Python Playwright；`node packages/vue-docx/tests/verify-integration.mjs`；`node packages/vue-xlsx/test/structure.mjs`；`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/BB-SEC-URL/20260710T142847+0800/`（实现者）；`.../20260710T144645+0800-final-independent/`（独立） | Codex `/root` | `public-api` | 现有 `PdfViewer.src` | 新增 `source`、`urlPolicy`、验证字节、公开诊断和 `document-load-error`；`src` 保留 | `1.0.0`（最早） | 不适用 |
| P0-URL-02 | Codex `/root` | Codex `/root/sec_url_harness` | `41913ca03666aa99302d13924e038db17e478480` | `packages/xlsx-core/src/{runtime/xlsx-url-policy.ts,types/worksheet-types.ts,types/index.ts,index.ts}`、`packages/vue-xlsx/src/{index.ts,composables.ts,composables/useXlsxViewerController.ts,composables/workbook-state.ts,composables/clipboard.ts,composables/internal.ts,composables/navigation.ts}`、`apps/demo/src/pages/XlsxViewerPage.vue`、`docs/plan/stabilization-roadmap.md` | DOCX/PDF、Worker/WASM 实现、其他 demo 页面；`README.md`、`docs/INDEX.md`、`docs/plan/README.md`、`packages/docx-core/tsup.config.ts` 的现有用户改动 | `BB-SEC-URL`（SEC-002 及 XLSX URL 子集）、`BB-P0-ROUTES` | `pnpm typecheck`；`pnpm build`；正式 preview + Python Playwright；`node packages/vue-docx/tests/verify-integration.mjs`；`node packages/vue-xlsx/test/structure.mjs`；`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/BB-SEC-URL/20260710T145414+0800-xlsx-baseline/`；`.../20260710T150543+0800-xlsx/`（实现者）；`.../20260710T151315+0800-xlsx-final2-independent/`（独立） | Codex `/root` | `public-api` | 现有 `UseXlsxViewerControllerOptions.src` 与 `XlsxViewerController.download` | 新增显式 URL 策略、结构化来源错误、公开诊断和已验证源字节下载；`src` 保留 | `1.0.0`（最早） | 不适用 |
| P0-RACE-01 | Codex `/root` | Codex `/root/race_baseline` | `41913ca03666aa99302d13924e038db17e478480` | `packages/docx-core/src/runtime.ts`、`packages/vue-docx/src/composables/{useDocxModel.ts,editor-import-export.ts}`、`packages/vue-xlsx/src/composables/{internal.ts,useXlsxViewerController.ts,navigation.ts}`、`apps/demo/src/pages/{DocxViewerPage.vue,XlsxViewerPage.vue}`、`docs/plan/stabilization-roadmap.md` | PDF、其他 demo 页面；`README.md`、`docs/INDEX.md`、`docs/plan/README.md` 的现有用户改动；`packages/docx-core/tsup.config.ts` 的现有用户改动 | `BB-RACE`（RACE-001、RACE-002、RACE-004、RACE-005）和 `BB-P0-ROUTES` 回归 | `pnpm typecheck`；`pnpm build`；`node packages/vue-docx/tests/verify-integration.mjs`；`node packages/vue-xlsx/test/structure.mjs`；正式 preview + Python Playwright；`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/BB-RACE/20260710T152255+0800-docx-baseline/`；最终目录待生成 | Codex `/root` | `internal-bridge` | `2026-07-10`；工作区未提交 | 控制器各自维护的请求代次与取消；P2-RACE-01 统一加载协调器 | 不适用 | `2026-08-09` |
| P0-PACK-01 | Codex `/root` | Codex `/root/race_baseline` | `41913ca03666aa99302d13924e038db17e478480` | `packages/{docx-core,xlsx-core,vue-docx,vue-xlsx,vue-extend}/package.json`、`packages/docx-core/scripts/normalize-root-worker.mjs`（仅修复构建产物根 Worker 的相对 chunk 路径）、`.npmrc`（本地开发仍链接满足 semver 的工作区包）、`pnpm-lock.yaml`（仅把两个公开包内部依赖从 `workspace:*` 同步为已发布 semver）、`docs/plan/stabilization-roadmap.md` | 所有业务源码、Vite/tsup 配置、demo、DOCX/XLSX/PDF 实现；`README.md`、`docs/INDEX.md`、`docs/plan/README.md` 的现有用户改动；`packages/docx-core/tsup.config.ts` 的现有用户改动 | `BB-PACK-CONSUMER`（PACK-001、PACK-002 的包入口预检） | `pnpm typecheck`；`pnpm build`；五个 `npm pack --json --pack-destination`；tgz 入口/类型/资源清单检查；工作区外 `npm install`、`tsc --noEmit`、Vite 生产构建；`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/BB-PACK-CONSUMER/20260710T154710+0800-pack-baseline/`；`.../20260710T160437+0800/`（实现者）；`.../20260710T160437+0800/independent-20260710T160948+0800/`（独立） | Codex `/root` | `public-api` | 现有 `0.1.0` 包入口 | 同名根入口，改为压缩包内真实 `dist` 文件 | `1.0.0`（最早） | 不适用 |
| P0-PACK-02 | Codex `/root` | Codex `/root/pack02_independent`（同会话交叉复测，不替代候选发布的全新会话复核） | `41913ca03666aa99302d13924e038db17e478480` | `packages/docx-core/{package.json,src/engine/wasm-asset.ts,scripts/prepare-public-resources.mjs}`、`packages/xlsx-core/{package.json,src/{index.ts,wasm.ts,wasm-asset.ts,worker-client.ts,types/worksheet-types.ts},scripts/prepare-public-resources.mjs}`、`packages/vue-xlsx/src/composables/useXlsxViewerController.ts`（公开 Worker URL 透传）、`docs/{testing/agent-execution-runbook.md,plan/stabilization-roadmap.md}` | `packages/docx-core/tsup.config.ts` 的既有用户改动；DOCX/XLSX 业务解析、Vue 组件和 demo 页面；`README.md`、`docs/INDEX.md`、`docs/plan/README.md` 的既有用户改动 | `BB-PACK-CONSUMER`（PACK-003）、`BB-DOCX-WORKER` 回归 | `pnpm typecheck`；`pnpm build`；五个真实 `npm pack`；工作区外 tgz 消费项目的开发构建和正式 preview；Python Playwright 网络/控制台/截图；`git diff --check` | `.../BB-PACK-CONSUMER/20260710T160437+0800/independent-20260710T160948+0800/`（失败基线）；`.../BB-PACK-CONSUMER/20260710T164800+0800/`（五包 tgz、清单、开发与正式预览）；`.../independent-agent-20260710T170329+0800/`（同会话交叉复测） | Codex `/root` | `public-api` | 现有根入口、`setWasmSource`、DOCX Runtime 配置 | 新增稳定 `./runtime`、`./wasm-url`、`./assets/*`、`./worker` 入口；内置 Worker 保留静态 `new URL`；XLSX 控制器新增可选 `workerUrl`/Worker 工厂 | `1.0.0`（最早） | 不适用 |

| P0-STYLE-01 | Codex `/root` | 待同会话交叉复测 | `41913ca03666aa99302d13924e038db17e478480` | `packages/{vue-docx,vue-xlsx,vue-extend}/{package.json,vite.config.ts}`、`apps/demo/{vite.config.ts,src/main.ts}`、`docs/{testing/agent-execution-runbook.md,plan/stabilization-roadmap.md}` | `packages/docx-core/tsup.config.ts` 的既有用户改动；DOCX/XLSX/PDF 业务逻辑和页面；`README.md`、`docs/INDEX.md`、`docs/plan/README.md` 的既有用户改动；`apps/demo/src/main.ts` 中已阅读的 DOCX WASM 用户差异只做相邻样式 import 最小合并 | `BB-PACK-CONSUMER`（CSS 公开入口）、`BB-RACE`（RACE-002）、`BB-P0-ROUTES` 回归 | 先保存公开 CSS 入口解析失败和 RACE-002 高度为 0 的基线；`pnpm typecheck`；`pnpm build`；五包真实 `npm pack`；外部 tgz 消费项目正式 preview + Python Playwright；`BB-RACE`；`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/BB-RACE/20260710T153732+0800-race-self/`（布局失败基线）；本任务目录待生成 | Codex `/root` | `public-api` | 无稳定 CSS 子路径；仅 `vue-extend` 有内部 `./dist/index.css` | `./style.css`；保留 `vue-extend` 的旧内部路径兼容映射 | `1.0.0`（最早） | 不适用 |

| P0-DEMO-01 | Codex `/root` | 待同会话交叉复测 | `41913ca03666aa99302d13924e038db17e478480` | `docs/plan/stabilization-roadmap.md`；仅在黑盒验证暴露新问题时才使用对应 demo 页面最小修复路径 | 无关包、DOCX/XLSX/PDF 业务实现；`packages/docx-core/tsup.config.ts` 的既有用户改动；`README.md`、`docs/INDEX.md`、`docs/plan/README.md` 的既有用户改动 | `BB-P0-ROUTES` | `pnpm build`；正式 preview；Python Playwright 六路由直接访问与刷新；`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/P0-STYLE-01/20260710T171141+0800/BB-P0-ROUTES/attempt-1/summary.json`（首次当前构建通过；刷新复测待生成） | Codex `/root` | 不适用 | 不适用 | 不适用 | 不适用 | 不适用 |

| P1-CI-01 | Codex `/root` | 待同会话交叉复测；候选发布仍需全新会话 | `41913ca03666aa99302d13924e038db17e478480` | 根 `package.json`、`requirements-ci.txt`、`.github/workflows/ci.yml`、`scripts/ci/`、`tests/blackbox/routes_smoke.py`、`packages/vue-xlsx/test/structure.mjs`（只修复测试生命周期 warning）、`pnpm-lock.yaml`（仅在新增检查工具时）、`docs/{plan/stabilization-roadmap.md,testing/agent-execution-runbook.md,end-to-end-blackbox-test-plan.md}` | 所有产品业务源码；各包构建配置和公开接口；`packages/docx-core/tsup.config.ts` 的既有用户改动；`README.md`、`docs/INDEX.md`、`docs/plan/README.md` 的既有用户改动 | `P1-CI-01` 根门禁、自身失败传播探针；P0 根检查回归 | 保存缺失统一命令的失败基线；锁定依赖安装；`pnpm check`；`pnpm test:unit`；`pnpm test:component`；`pnpm test:blackbox`；`pnpm test:consumer`；人为制造可恢复的子任务失败并确认非零退出码；`pnpm typecheck`；`pnpm build`；`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/P1-CI-01/20260710T175228+0800/` | Codex `/root` | 不适用 | 不适用 | 不适用 | 不适用 | 不适用 |

| P1-UNIT-01 | Codex `/root` | 待同会话交叉复测；候选发布仍需全新会话 | `41913ca03666aa99302d13924e038db17e478480` | `tests/unit/`、`scripts/ci/run-suite.mjs`（只登记真实单元用例）、`docs/plan/stabilization-roadmap.md` | 所有产品源码、demo 页面、构建配置、包清单；`packages/docx-core/tsup.config.ts` 的既有用户改动；其他文档既有改动 | `P1-UNIT-01`、根 `pnpm test:unit`、P1 根门禁回归 | 保存只有旧 DOCX 集成和 XLSX 方法结构检查的覆盖缺口；新增正常、损坏、空输入、取消、并发、往返和格式断言；`pnpm test:unit`；`pnpm check`；`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/P1-UNIT-01/20260710T181440+0800/` | Codex `/root` | 不适用 | 不适用 | 不适用 | 不适用 | 不适用 |

| P1-FIXTURE-01 | Codex `/root` | 本会话复测；候选发布仍需全新会话 | `41913ca03666aa99302d13924e038db17e478480` | `test-data/manifest.json`、`scripts/{generate_test_materials.py,ci/verify-fixtures.mjs,ci/run-suite.mjs}`、`apps/demo/public/samples/`（仅由修复后的确定性生成器重建）、`.gitattributes`（只标记生成的 PDF 为二进制）、`docs/plan/stabilization-roadmap.md` | 产品源码、构建配置和包清单；`packages/docx-core/tsup.config.ts` 及其他受保护用户改动 | `P1-FIXTURE-01`、`pnpm test:unit`、材料漂移故障探针 | 保存旧 manifest 只有文件名/字节数的缺口；核对文件集合、SHA-256、类型签名、分类、合成来源、别名关系和预期行为；临时篡改副本证明校验非零；两次独立生成内容逐文件一致；`pnpm test:unit`；`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/P1-FIXTURE-01/20260710T181702+0800/` | Codex `/root` | 不适用 | 不适用 | 不适用 | 不适用 | 不适用 |

| P1-VUE-01 | Codex `/root` | 本会话复测；候选发布仍需全新会话 | `41913ca03666aa99302d13924e038db17e478480` | `tests/component/`、`scripts/ci/{component-smoke.mjs,run-suite.mjs}`（只登记组件测试）、必要时三个 Vue 包内由测试暴露的最小修复、`docs/plan/stabilization-roadmap.md` | demo 页面、核心解析实现、构建与发布配置；`packages/docx-core/tsup.config.ts` 及其他受保护用户改动 | `P1-VUE-01`、根 `pnpm test:component`、P1 根门禁回归 | 保存现有套件只覆盖两个空状态的失败基线；通过公开属性、事件和组合函数验证 DOCX/PDF 加载成功与失败、三个空状态、DOCX/PDF 卸载取消、PDF 对象 URL 回收、XLSX 快捷键以及控制器卸载终止 Worker；`pnpm test:component`；`pnpm typecheck`；`pnpm build`；`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/P1-VUE-01/20260710T182408+0800/` | Codex `/root` | 不适用 | 不适用 | 不适用 | 不适用 | 不适用 |

| P1-E2E-01 | Codex `/root` | 本会话复测；候选发布仍需全新会话 | `41913ca03666aa99302d13924e038db17e478480` | `tests/blackbox/`、`scripts/ci/run-suite.mjs`（只登记正式预览流程）、三个 Vue 包和 demo 页面中本流程需要的稳定 `data-testid`、`docs/plan/stabilization-roadmap.md` | 核心解析和 Worker/WASM 实现、构建与包清单；`packages/docx-core/tsup.config.ts` 及其他受保护用户改动 | `P1-E2E-01`、根 `pnpm test:blackbox`、`BB-P0-ROUTES` 回归 | 保存只验证四路由打开/刷新的覆盖缺口；只启动正式 preview；为每个重要流程创建新浏览器 context；导航前监听控制台、页面错误、请求失败、响应和下载；覆盖 DOCX 切换与错误恢复、DOCX 编辑/撤销/导出、XLSX 工作表切换/编辑/撤销/下载/错误恢复；首次失败最多重试一次并保留两次证据；每次保存截图、事件和 trace；`pnpm typecheck`；`pnpm build`；`pnpm test:blackbox`；`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/P1-E2E-01/20260710T182856+0800/` | Codex `/root` | 不适用 | 不适用 | 不适用 | 不适用 | 不适用 |

| P1-CONSOLE-01 | Codex `/root` | 本会话复测；候选发布仍需全新会话 | `41913ca03666aa99302d13924e038db17e478480` | `tests/blackbox/{browser_evidence.py,console_allowlist.json,routes_smoke.py,e2e_workflows.py}`、必要的 CI 故障探针编排、`docs/plan/stabilization-roadmap.md` | 产品源码、demo 页面、构建和发布配置；`packages/docx-core/tsup.config.ts` 及其他受保护用户改动 | `P1-CONSOLE-01`、根 `pnpm test:blackbox`、浏览器事件故障探针 | 保存两个浏览器脚本重复采集且无统一规则的覆盖缺口；建立唯一采集器和显式最小允许清单；路由与业务流程共同使用；人为注入控制台错误、页面异常、请求失败和资源 404，逐项确认父命令非零且证据包含违规事件；恢复后正式黑盒通过；`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/P1-CONSOLE-01/20260710T183348+0800/` | Codex `/root` | 不适用 | 不适用 | 不适用 | 不适用 | 不适用 |

| P1-RACE-TEST-01 | Codex `/root` | 本会话复测；候选发布仍需全新会话 | `41913ca03666aa99302d13924e038db17e478480` | `tests/blackbox/{fault_server.py,race_workflows.py,browser_evidence.py}`、`scripts/ci/run-suite.mjs`（只登记竞态套件）、`docs/plan/stabilization-roadmap.md` | 产品源码、demo 页面、构建和发布配置；`packages/docx-core/tsup.config.ts` 及其他受保护用户改动 | `BB-RACE`、根 `pnpm test:blackbox`、P1 根门禁回归 | 保存旧竞态脚本只存在 output 的覆盖缺口；从正式 `apps/demo/dist` 启动可控延迟服务器；DOCX/XLSX 各覆盖慢 A→快 B、20 次连续切换和中途卸载；通过页面公开状态/诊断确认最终只保留 B；在页面脚本前包装标准 Worker 和 URL API，仅统计实例拥有资源的创建/释放，不读取模块私有变量；每个用例新 context、失败最多重试一次、保存截图/事件/trace；`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/P1-RACE-TEST-01/20260710T183813+0800/` | Codex `/root` | 不适用 | 不适用 | 不适用 | 不适用 | 不适用 |

| P1-PACK-TEST-01 | Codex `/root` | 本会话复测；候选发布仍需全新会话 | `41913ca03666aa99302d13924e038db17e478480` | `scripts/ci/{pack-manifests.mjs,pack-consumer.mjs,run-suite.mjs}`、`tests/consumer/`、`tests/blackbox/pack_consumer.py`、`docs/plan/stabilization-roadmap.md` | 五包产品源码与包清单（测试若发现失败另开最小修复记录）、demo 页面；`packages/docx-core/tsup.config.ts` 及其他受保护用户改动 | `BB-PACK-CONSUMER`、根 `pnpm test:consumer`、P1 根门禁回归 | 保存当前只检查 tgz 清单的覆盖缺口；从当前构建重新打五包并记录逐包与资源 SHA-256；在系统临时目录生成消费项目、只用 `file:*.tgz` 安装且核对 node_modules 无符号链接；公共 ESM、类型、三份 CSS 和资源子路径均从包出口消费；公共目录只允许测试文档，不复制 Worker/WASM；TypeScript、Vue 类型和 Vite 正式构建通过；正式 preview 中 DOCX、XLSX Worker、PDF 均 ready，两 Worker/两 WASM 请求 200 且 MIME/文件 SHA 与 tgz 一致；结束清理临时目录；包入口缺失故障探针必须非零 | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/P1-PACK-TEST-01/20260710T184242+0800/` | Codex `/root` | 不适用 | 不适用 | 不适用 | 不适用 | 不适用 |

| P1-DOC-TEST-01 | Codex `/root` | 本会话复测；候选发布仍需全新会话 | `41913ca03666aa99302d13924e038db17e478480` | `docs/{end-to-end-blackbox-test-plan.md,testing/agent-execution-runbook.md,plan/stabilization-roadmap.md}`、根 `package.json`（仅新增文档检查入口）、`scripts/ci/{verify-docs.mjs,run-suite.mjs}` | 产品源码、demo、包清单、`docs/INDEX.md` 等已有用户改动；`packages/docx-core/tsup.config.ts` 的既有用户改动 | `P1-DOC-TEST-01`、根 `pnpm test:docs`、P1 根门禁回归 | 保存核心测试文档仍描述 P1 后续任务未完成、临时 fault server、旧材料和消费夹具路径的漂移基线；更新统一命令、实际套件内容、证据和排错说明；建立机器检查，验证根脚本、别名、测试文件、材料清单、固定依赖和文档文字均与仓库一致，并禁止已知过期表述；故意改成不存在脚本时必须非零；`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/P1-DOC-TEST-01/20260710T184948+0800/` | Codex `/root` | 不适用 | 不适用 | 不适用 | 不适用 | 不适用 |

| P2-RUNTIME-01 | Codex `/root` | 本会话复测；候选发布仍需全新会话 | `41913ca03666aa99302d13924e038db17e478480` | `packages/office-runtime/**`、`packages/{docx-core,xlsx-core}/src/` 中运行实例与加载适配、`packages/{vue-docx,vue-xlsx,vue-extend}/src/` 中来源适配与控制器、`apps/demo/src/pages/{PdfViewerPage.vue,XlsxViewerPage.vue}`（仅公开显示 taskId）、相关包清单和锁文件、`tests/{unit,component,blackbox}/`、`scripts/ci/`、`docs/plan/stabilization-roadmap.md` | DOCX 渲染和编辑面、XLSX 网格和可选图表实现、其他 demo 页面、`README.md`、`docs/INDEX.md`、`docs/plan/README.md` 的已有用户改动；`packages/docx-core/tsup.config.ts` 的既有用户改动，除非真实 tgz 证明必须做最小合并 | `P2-RUNTIME-01`、`BB-RACE`、`BB-SEC-URL`、`BB-PACK-CONSUMER` 回归 | 保存私有共享包不存在且三套来源/错误/取消规则分散的失败基线；新增先失败的共享运行边界测试；`pnpm typecheck`；`pnpm build`；`pnpm test:unit`；`pnpm test:component`；正式 preview `pnpm test:blackbox`；`pnpm test:consumer`；`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/P2-RUNTIME-01/20260710T190942+0800/` | Codex `/root` | 不适用 | 不适用 | 不适用 | 不适用 | 不适用 |

| P2-CONFIG-01 | Codex `/root` | 本会话复测；候选发布仍需全新会话 | `41913ca03666aa99302d13924e038db17e478480` | `packages/docx-core/src/{runtime.ts,engine/wasm.ts,viewer/wasm-source.ts}`、`packages/xlsx-core/src/{wasm.ts,worker-client.ts,index.ts,runtime/**,types/**}`、`packages/vue-xlsx/src/composables/useXlsxViewerController.ts`、相关单元/组件测试、`apps/demo/src/{main.ts,pages/RuntimeIsolationPage.vue}`（正式公开隔离消费页）、`tests/blackbox/`、`scripts/ci/`、`docs/plan/stabilization-roadmap.md` | DOCX 渲染/编辑实现、XLSX 网格/图表/编辑实现、PDF、其他 demo 页面、受保护的 `packages/docx-core/tsup.config.ts` 和其他已有用户改动 | `P2-CONFIG-01`、`BB-DOCX-WORKER`、`BB-RACE`、`BB-PACK-CONSUMER` 回归 | 保存配置对象修改会影响已创建实例、XLSX 没有 Runtime 及全局 WASM 默认值会渗入 Worker client 的失败基线；新增双实例/乱序/销毁测试；`pnpm typecheck`；`pnpm build`；`pnpm test:unit`；`pnpm test:component`；正式 preview 双实例隔离；`pnpm test:blackbox`；`pnpm test:consumer`；`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/P2-CONFIG-01/20260710T193226+0800/` | Codex `/root` | `public-api` | `2026-07-10` 新增实例 API；旧 `setWasmSource` 保留 | `createDocxRuntime` / `createXlsxRuntime` 实例配置 | `1.0.0`（旧入口最早删除） | 不适用 |

| P2-RACE-01 | Codex `/root` | 本会话复测；候选发布仍需全新会话 | `41913ca03666aa99302d13924e038db17e478480` | `packages/office-runtime/src/{load-task.ts,index.ts}`、`packages/xlsx-core/src/{worker-client.ts,runtime/xlsx-runtime.ts}`、`packages/vue-xlsx/src/composables/{useXlsxViewerController.ts,navigation.ts,internal.ts,history.ts,image-assets.ts}`、`packages/vue-docx/src/composables/editor-import-export.ts`、必要的 `packages/vue-extend/src/components/PdfViewer.vue` 最小清理、相关单元/组件/黑盒测试、正式 demo 公开状态与 `docs/plan/stabilization-roadmap.md` | DOCX 渲染/编辑命令实现、XLSX 网格/图表功能扩展、核心模型/API 整理、资源预算、其他 demo 页面、受保护的 `packages/docx-core/tsup.config.ts` 和其他已有用户改动 | `P2-RACE-01`、`BB-RACE`、`BB-P0-ROUTES`、`BB-DOCX-WORKER`、`BB-PACK-CONSUMER` 回归 | 保存 XLSX Worker 单请求不能取消、延迟加载丢失统一 task/signal、旧主线程解析不释放 Workbook、DOCX 异步导出会在新建/导入/卸载后继续下载且异常路径泄漏对象 URL 的失败基线；统一任务完成/取消，给 Worker 请求接入 AbortSignal，延迟加载复用同一任务上下文，所有过期结果释放一次；新增 Worker 取消、延迟加载、DOCX 导出版本/卸载、对象 URL 异常清理测试；`pnpm typecheck`；`pnpm build`；`pnpm test:unit`；`pnpm test:component`；正式 preview `pnpm test:blackbox`；`pnpm test:consumer`；`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/P2-RACE-01/20260710T194517+0800/` | Codex `/root` | 不适用 | 不适用 | 不适用 | 不适用 | 不适用 |

| P2-DOCX-RENDER-01 | Codex `/root` | 本会话复测；候选发布仍需全新会话 | `41913ca03666aa99302d13924e038db17e478480` | `packages/vue-docx/src/components/{DocxEditor.vue,DocxViewer.vue,DocxViewerRoot.vue,DocxPageSurface.vue,DocxPageBody.vue,DocxParagraphHost.vue,DocxTableHost.vue,DocxHeader.vue,DocxFooter.vue,DocxImageLayer.vue,DocxCommentsGutter.vue,DocxTrackChangesGutter.vue,DocxDocumentSurface.vue}`、必要的 `packages/vue-docx/src/render/` 与只读渲染组合函数、`packages/docx-core/src/layout/` 中稳定 snapshot/section 纯帮助函数、正式 demo 对比页、相关组件/黑盒/截图证据和路线图 | Runtime/Worker/WASM、DOCX 导入导出与编辑命令/选区/撤销重做、XLSX/PDF、包构建配置和公开出口、受保护的 `packages/docx-core/tsup.config.ts` 及其他已有用户改动 | `P2-DOCX-RENDER-01`、`BB-DOCX-PARITY`、`BB-DOCX-WORKER`、`BB-P0-ROUTES`、`BB-RACE` 回归 | 保存 Editor 分页把页内序号误作模型 nodeIndex、表格正文固定空、页眉页脚固定 false、Viewer/Editor 两套渲染树、Editor model 属性未消费的失败基线；建立唯一内部文档 surface，以 layout snapshot 的 source.nodeIndex 为定位；段落/表格/图片/字段/header/footer/批注修订复用只读渲染；只读不挂编辑控件；新增同 model 结构快照与正式对比页，保存首/中/末页截图及差异；`pnpm typecheck`；`pnpm build`；`pnpm test:unit`；`pnpm test:component`；正式 preview `pnpm test:blackbox`；`pnpm test:consumer`；`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/P2-DOCX-RENDER-01/20260710T200908+0800/` | Codex `/root` | 不适用 | 不适用 | 不适用 | 不适用 | 不适用 |

| P2-CORE-01 | Codex `/root` | 本会话复测；候选发布仍需全新会话 | `41913ca03666aa99302d13924e038db17e478480` | 新增 `packages/{docx-core,xlsx-core}/src/core.ts` 与精确平台适配边界、两包 `package.json` 的稳定 `./core` 出口、`packages/xlsx-core/tsup.config.ts` 最小入口、依赖图/纯函数/真实 tgz 测试、路线图；`packages/docx-core/tsup.config.ts` 仅由主负责人先读用户差异后做必要的单行入口合并 | Vue 包、demo、DOCX 渲染/编辑交互、Runtime/WASM 配置、性能预算、其他包、受保护 docx tsup 中除必要 entry 之外的用户改动 | `P2-CORE-01`、core 依赖方向、纯函数不变性/顺序隔离、深层入口拒绝、`BB-PACK-CONSUMER` 回归 | 保存两包没有 `./core`、根声明泄漏 DOM/Worker、纯转换直接读 DOMParser/Canvas/URL 且存在模块级可变缓存的失败基线；新增稳定纯入口和显式 XML/测量/二进制 URL 适配接口；纯入口在无 DOM/Worker 子进程可执行且输入不变；TypeScript AST 依赖图禁止 Vue/demo/浏览器全局进入纯图；真实 tgz 可导入 `./core` 并拒绝 src/viewer/editor/images 等深层路径；保持根入口 0.x 兼容；`pnpm typecheck`、`pnpm build`、单元/组件/consumer、`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/P2-CORE-01/20260710T201047+0800/` | Codex `/root` | `public-api` | `2026-07-10` 新增纯入口；旧根兼容保留 | `@arcships/{docx-core,xlsx-core}/core` | `1.0.0` 前不删除旧根入口 | 不适用 |

| P2-DOCX-EDIT-01 | Codex `/root` | 本会话复测；候选发布仍需全新会话 | `41913ca03666aa99302d13924e038db17e478480` | `packages/vue-docx/src/components/{DocxEditor.vue,DocxParagraphHost.vue,DocxTableHost.vue,DocxImageLayer.vue,DocxContextMenu.vue,DocxToolbar.vue}`、`packages/vue-docx/src/composables/{useDocxEditor.ts,editor-transaction.ts,editor-history.ts,editor-selection.ts,editor-text-input.ts,editor-image.ts,editor-format.ts,editor-table.ts}`、必要的 `packages/docx-core/src/editor/` 纯命令、`apps/demo/src/pages/DocxEditorPage.vue`、相关组件/黑盒测试、套件登记和路线图 | DOCX Runtime/Worker/WASM/导入导出、只读基础渲染树、XLSX/PDF、构建与包出口、受保护的 `packages/docx-core/tsup.config.ts` 及其他已有用户改动 | `P2-DOCX-EDIT-01`、DOCX-102～DOCX-109、`BB-RACE` 回归 | 保存只读仍注册全局快捷键、图片与表格拖缩不提交真实模型、多个公开编辑方法为空以及黑盒只覆盖整段替换/单次撤销的失败基线；把监听器、选区、输入、命令、历史和浮层作为可关闭交互层；开启时所有操作通过新模型与历史提交；覆盖中英文组合输入、部分选区格式、撤销重做与选区恢复、表格和图片操作、只读切换；`pnpm typecheck`、`pnpm build`、组件/正式 preview 黑盒/竞态/consumer、`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/P2-DOCX-EDIT-01/20260710T203805+0800/` | Codex `/root` | 不适用 | 不适用 | 不适用 | 不适用 | 不适用 |

| P2-API-01 | Codex `/root` | 本会话复测；候选发布仍需全新会话 | `41913ca03666aa99302d13924e038db17e478480` | 五个公开包的根 `src/index.ts`、必要的稳定子入口与公开错误类型、三个 Vue 包的公开组件事件类型、五包 `package.json`（只做出口兼容核对）、`docs/api/public-api-contract.md`、公开接口清单与机器门禁、真实 tgz 消费矩阵、`scripts/ci/run-suite.mjs` 和路线图 | 业务渲染与编辑行为、Runtime/Worker/WASM 内部实现、性能与资源预算、demo 私有状态、受保护的 `packages/docx-core/tsup.config.ts` 及其他已有用户改动 | `P2-API-01`、公开声明与弃用标记、严格深层导入拒绝矩阵、`BB-PACK-CONSUMER`、正式路由/编辑/安全/竞态回归 | 保存五包生成声明没有 `@deprecated`、DOCX runtime 子入口缺少公开错误诊断、XLSX 销毁后泄漏私有错误、FileUpload 缺少结构化拒绝事件、深层导入拒绝只覆盖少量路径的失败基线；发布最小公开接口和兼容分类；0.2.0 只标记不删除旧导出；公开事件/错误结构；构建后从真实 tgz 在工作区外正向导入全部稳定入口并拒绝所有未声明深层路径；`pnpm typecheck`、`pnpm build`、单元/组件/正式 preview 黑盒/consumer、`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/P2-API-01/20260710T205140+0800/` | Codex `/root` | `public-api` | `2026-07-10` 登记现有 0.1.0 根导出与旧别名 | 新实例入口、高层组件、公开错误/事件与 `./style.css` | `1.0.0`（已公开旧入口最早删除） | 从未公开桥接最晚 `2026-08-09` 或下个次版本（取更早） |

| P2-DEAD-01 | Codex `/root` | 本会话复测；候选发布仍需全新会话 | `41913ca03666aa99302d13924e038db17e478480` | `packages/vue-docx/src/components/DocxViewer.vue`、DOCX 缩略图与页面 surface registry、`packages/vue-xlsx/src/components/{XlsxSelectionOverlay.vue,XlsxChartOverlay.vue}`、`packages/xlsx-core/src/index.ts`、从未公开且无消费者的组件桶、错误注释、demo 支持状态、精确占位检查与路线图；图片/表格/工具栏空交互由 P2-DOCX-EDIT-01 负责后再复核 | P2-DOCX-EDIT-01 正在修改的图片/表格/工具栏/编辑组合函数、Runtime/Worker/WASM、PDF、构建与公开出口、受保护的 `packages/docx-core/tsup.config.ts` 及其他已有用户改动 | `P2-DEAD-01`、占位/重复路径检查、`BB-P0-ROUTES`、`BB-DOCX-PARITY` 回归 | 保存 Viewer 第二套死渲染、缩略图永远 unavailable/空 rerender、空 XLSX overlay/double-click、重复 A1 函数和错误深层导入说明；删除无消费者旧路或改为用户可见且有测试的降级，建立精确允许清单，不机械禁止合法空值；待编辑任务结束后复核其空处理器；`pnpm typecheck`、`pnpm build`、单元/组件/正式 preview 黑盒、`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/P2-DEAD-01/20260710T203805+0800/` | Codex `/root` | `internal-bridge` | `2026-07-10` 识别旧公开兼容组件 | 根入口保留并在 P2-API-01 标记弃用；从未公开桶可删除 | `1.0.0`（公开兼容项最早删除） | 不适用 |

| P3-PDF-BUDGET-01 | Codex `/root` | 本会话 PDF 黑盒子任务复测；候选发布仍需全新会话 | `41913ca03666aa99302d13924e038db17e478480` | `packages/vue-extend/src/{index.ts,types.ts,pdf-url-policy.ts,components/PdfViewer.vue}`、`apps/demo/src/pages/PdfViewerPage.vue`、`tests/{unit/pdf-source.test.mjs,component/vue-components.test.mjs,blackbox/pdf_workflows.py}`、`scripts/ci/{run-suite.mjs,public-api-contract.json}`、`docs/{architecture-review-and-target-design.md,end-to-end-blackbox-test-plan.md,api/public-api-contract.md,plan/stabilization-roadmap.md}` | DOCX/XLSX 产品源码与预算、PDF 页数/像素/总内存/并发页限制、其他 demo 页面、`README.md`、`docs/INDEX.md`、`docs/plan/README.md` 的既有用户改动；受保护的 `packages/docx-core/tsup.config.ts` | `BB-PDF-VIEWER` 的 PDF-008、PDF 正常加载与错误恢复回归 | 保存 PDF 没有整份文件体积上限的基线；公开默认 `50 MiB` 的 `maxFileSize`，宿主可调整且不叠加隐藏硬上限；验证本地声明大小、URL `Content-Length` 和实际响应字节；超限返回 `PDF_TOO_LARGE`，不得启动引擎或保留旧页面；`pnpm typecheck`、`pnpm build`、`pnpm test:unit`、`pnpm test:component`、正式 preview PDF-008、`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/BB-PDF-VIEWER/20260710T235500+0800-size-limit/` | Codex `/root` | `public-api` | `2026-07-10` 新增 PDF 整份文件体积配置与错误码 | `DEFAULT_PDF_MAX_FILE_SIZE`、`PdfLoadOptions.maxFileSize`、`PdfViewer.maxFileSize`、`PDF_TOO_LARGE` | 不适用（兼容新增） | 不适用 |
| P3-PDF-01 | Codex `/root` | 本会话 PDF 黑盒子任务复测；候选发布仍需全新会话 | `41913ca03666aa99302d13924e038db17e478480` | `packages/vue-extend/{package.json,vite.config.ts,src/index.ts,src/types.ts,src/pdf-url-policy.ts,src/pdf/**,src/components/PdfViewer.vue}`、`apps/demo/src/pages/PdfViewerPage.vue`、`tests/{component/vue-components.test.mjs,blackbox/pdf_workflows.py,blackbox/pack_consumer.py,blackbox/race_workflows.py,blackbox/security_workflows.py,consumer/template/src/package-contract.ts}`、`scripts/ci/{run-suite.mjs,pack-manifests.mjs,pack-consumer.mjs,public-api-contract.json}`、`docs/{architecture-review-and-target-design.md,end-to-end-blackbox-test-plan.md,api/public-api-contract.md,plan/stabilization-roadmap.md}`、`packages/{xlsx-core,vue-xlsx}/package.json` 与 `pnpm-lock.yaml`（仅修复真实 tgz 暴露的 WASM 依赖版本漂移） | DOCX 产品源码与构建配置、XLSX 业务逻辑、其他 demo 页面、DOCX/XLSX 的 P3 资源预算实现、PDF 页数/像素/总内存/并发页限制、`README.md`、`docs/INDEX.md`、`docs/plan/README.md` 的既有用户改动；受保护的 `packages/docx-core/tsup.config.ts` | `BB-P0-ROUTES` 的 PDF 用例、PDF-001～006、PDF-008、`BB-SEC-URL`、`BB-RACE`、`BB-PACK-CONSUMER` | 保存字节验证后伪 `ready`、正则猜页数、原始压缩字节假搜索、缩放/翻页不作用于文档和 tgz 缺 PDFium 的失败基线；建立实例 PDF Runtime 与专用 Worker，从已验证字节真实打开、渲染、缩略图和搜索；首张图片解码后才能 `ready`；运行期间由实例持有 Worker 与图片地址，切换和离页后全部终止、撤销；接入单一公开整份文件体积限制；`pnpm typecheck`、`pnpm build`、`pnpm test:unit`、`pnpm test:component`、正式 preview PDF-001～006 与 PDF-008、五包真实 tgz 外部消费、`git diff --check` | `output/acceptance/41913ca03666aa99302d13924e038db17e478480/PDF-FUNCTION-01/20260710T232300+0800/` | Codex `/root` | `public-api` | `2026-07-10` 新增 PDF Runtime、WASM 地址、实例注入、整份文件体积配置和渲染/搜索类型；旧 `src` 输入保留 | `createPdfRenderRuntime`、`PdfViewer.runtime`、`PdfViewer.pdfiumWasmUrl`、`PdfViewer.maxFileSize`、包内 `./assets/pdfium.wasm` | 不适用（兼容新增） | 不适用 |

| P3-BUDGET-01 | Codex `/root/pdf_acceptance_audit` | Codex `/root` 本会话集成复测；候选发布仍需全新会话 | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | `packages/{docx-core,xlsx-core}/src/` 中实例资源限制与解析入口、`packages/{vue-docx,vue-xlsx}/src/` 中公开配置和错误透传、直接相关单元/组件/压力测试、`scripts/ci/` 中套件登记、`docs/plan/stabilization-roadmap.md` | PDF 体积规则与 PDF 实现、可选图表加载和网格滚动优化、受保护的 `packages/docx-core/tsup.config.ts` | `BB-STRESS` 的压缩包、路径、XML、MIME 与边界组合用例；P0～P2 回归 | 保存缺少归档/XML/模型限制和稳定错误的失败基线；新增实例级公开限制、尽早停止、取消和诊断；`pnpm typecheck`、`pnpm build`、单元、组件、正式 preview 压力子集、`git diff --check` | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/P3-BUDGET-01/20260711T000727+0800/` | Codex `/root` | `public-api` | `2026-07-11` 兼容新增 | DOCX/XLSX 实例资源限制与结构化错误 | 不适用 | 不适用 |
| P3-PERF-BASELINE-01 | Codex `/root/pdf_blackbox_tests` | Codex `/root` 本会话集成复测；候选发布仍需全新会话 | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | `tests/blackbox/` 中 XLSX 性能套件、性能预算机器配置、必要测试脚本与 `docs/{end-to-end-blackbox-test-plan.md,testing/agent-execution-runbook.md,plan/stabilization-roadmap.md}` | 产品运行源码、包出口和构建配置、PDF 规则、受保护的 `packages/docx-core/tsup.config.ts` | `BB-PERF-XLSX`、包体清单与正式 preview 三轮基线 | 固定 Chromium、`1440×900`、设备倍率 1、材料校验值和同机连续三轮；保存旧套件缺失的失败基线；记录原始值、中位数和最差值；首份待批准基线不冒充性能 PASS；`git diff --check` | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/P3-PERF-BASELINE-01/20260711T000727+0800/` | Codex `/root` | 不适用 | 不适用 | 不适用 | 不适用 | 不适用 |
| P3-XLSX-SCROLL-01 | Codex `/root/xlsx_scroll_design` | Codex `/root` 本会话集成复测；候选发布仍需全新会话 | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | `packages/vue-xlsx/src/components/XlsxGrid.vue`、直接的网格坐标/滚动组件测试、`tests/baseline/p3-xlsx-performance-baseline.test.mjs`（只补正式断言，不改预算值） | Home/demo 路由与分包、Workbook/Worksheet 生命周期、PDF、DOCX、包出口、受保护的 `packages/docx-core/tsup.config.ts` | `BB-PERF-XLSX` 的 PERF-004/PERF-005、网格组件回归 | 三轮正式基线已稳定证明横纵滚动范围均为 0，现有帧数据只是空闲循环；增加真实内容尺寸层、累计偏移、二分定位、帧内合并绘制和仅尺寸变化时重设 Canvas；覆盖深滚动点击/编辑/Tab/撤销与隐藏行列映射；`pnpm typecheck`、`pnpm build`、定向组件、正式 preview 三轮、`git diff --check` | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/P3-PERF-BASELINE-01/20260711T011029+0800/formal-after-lifecycle/` | Codex `/root` | 不适用 | 不适用 | 不适用 | 不适用 | 不适用 |
| P3-XLSX-LOAD-01 | Codex `/root/xlsx_lazy_load_impl` | Codex `/root` 本会话集成复测；候选发布仍需全新会话 | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | `apps/demo/src/main.ts`、`packages/vue-xlsx/src/{index.ts,components/XlsxChartOverlay.vue,optional/**}`、`packages/vue-xlsx/{package.json,vite.config.ts}`、必要时仅为 WASM 去重的 `apps/demo/vite.config.ts`、直接按需加载测试 | `XlsxGrid.vue`、Workbook/Worksheet 生命周期、PDF、DOCX、性能预算数值、受保护的 `packages/docx-core/tsup.config.ts` | `BB-PERF-XLSX` 的 PERF-001/PERF-002/PERF-006、正式 XLSX 图表回归 | 三轮正式基线已稳定证明 Home 主包静态包含 XLSX、进入图表页没有新功能分包请求，demo 还包含重复 WASM；把 demo 路由和可选图表/地图/WebGL 能力改为动态加载，保持根公开导出兼容；无图表时不得请求功能包，图表样本必须按需加载并正常显示；`pnpm typecheck`、`pnpm build`、静态门禁、正式 preview 三轮、`git diff --check` | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/P3-PERF-BASELINE-01/20260711T011029+0800/formal-after-lifecycle/` | Codex `/root` | `public-api` | `2026-07-11` 保持根导出兼容并增加内部按需入口 | 无新增必须公开的符号 | 不适用 | 不适用 |
| P3-IMAGE-01 | Codex `/root/pdf_acceptance_audit` | Codex `/root` 本会话集成复测；候选发布仍需全新会话 | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | `packages/office-runtime/src/` 图片预检与实例预算、`packages/{docx-core,xlsx-core,vue-docx,vue-xlsx}/src/` 直接图片入口、相关单元/组件/压力测试 | PDF、可选图表分包、网格滚动、历史记录、受保护的 `packages/docx-core/tsup.config.ts` | `BB-STRESS` 的 STRESS-004/014、图片组件错误与恢复、P0～P2 图片回归 | 保存解码失败无稳定事件、尺寸检查太晚、并发与总量无上限的失败基线；在对象 URL、Canvas 和像素缓冲前预检；限制归属运行实例；失败和销毁释放资源；执行类型、构建、单元、组件、正式 preview 压力子集和差异检查 | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/BB-STRESS/20260710T161327Z-p3-image-baseline/` | Codex `/root` | `public-api` | `2026-07-11` 兼容新增 | 实例图片预算、图片错误与诊断 | 不适用 | 不适用 |
| P3-CACHE-01 | Codex `/root/pdf_component_tests`、Codex `/root/cache_cleanup_impl` | Codex `/root` 本会话集成复测；候选发布仍需全新会话 | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | DOCX/XLSX 实例历史、缩略图、分页/排版缓存、图片对象 URL、XLSX Workbook/Worksheet/CellValue 生命周期与直接单元/组件/压力测试 | PDF、URL/归档预算、Worker 降级策略、图表分包、网格滚动、受保护的 `packages/docx-core/tsup.config.ts` | `BB-STRESS` 的 STRESS-006/008/009/010/014、`BB-PERF-XLSX` PERF-007、历史/缓存/生命周期单元与组件回归 | 保存历史只限数量、缩略图只限条目、模块缓存串实例、对象 URL 重复/异常泄漏和父 Workbook 先释放导致 WASM 越界的失败基线；实现按字节+数量淘汰、WeakMap/无强缓存、幂等 URL 清理、实例登记与子对象先释放；执行类型、构建、单元、组件、真实 WASM、正式 preview 三轮、压力回归和差异检查 | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/BB-STRESS/20260711T001709+0800-p3-cache-baseline/` | Codex `/root` | `public-api` | `2026-07-11` 兼容新增 | 历史字节上限、实例缩略图、对象 URL 与 Workbook 子对象生命周期 | 不适用 | 不适用 |
| P3-WORKER-01 | Codex `/root` | 本会话开发后自测；候选发布仍需全新会话 | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | `packages/{docx-core,xlsx-core}/src/` Worker 客户端与运行实例、`packages/vue-xlsx/src/composables/useXlsxViewerController.ts` Worker 加载/图表分支、直接单元/组件/黑盒测试 | PDF、图片预算、历史/缓存、图表分包入口、网格滚动、受保护的 `packages/docx-core/tsup.config.ts` | Worker 生命周期基线、`BB-STRESS` STRESS-005/010、正式 XLSX 加载与错误恢复 | 保存的失败基线 6 项中有 2 项失败：XLSX 取消未终止 Worker，Worker 失败后静默主线程重跑并报 ready；修复后执行类型、构建、基线、组件、正式 preview 超时/恢复和差异检查 | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/BB-STRESS/20260711T002107+0800-p3-worker-baseline/` | Codex `/root` | `public-api` | `2026-07-11` 行为修复 | Worker 取消/超时终止和失败诊断 | 不适用 | 不适用 |
| P3-STRESS-01 | Codex `/root` | 本会话开发后自测；候选发布仍需全新会话 | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | `tests/blackbox/{stress_workflows.py,worker_workflows.py}`、`scripts/ci/run-suite.mjs` 的既有正式压力入口、证据和路线图 | 产品公开接口、PDF 复杂预算、受保护的 `packages/docx-core/tsup.config.ts` | `BB-STRESS` 全部正式用例 | 仅在图片、PDF、缓存、滚动和 Worker 前置项均达到 `ready_for_gate` 后启动；用同一正式构建执行极端输入、边界恢复、Worker 超时/取消/恢复和资源清理，首轮失败最多重试一次并保留证据；`git diff --check` | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/BB-STRESS/20260711T-p3-final-stress/` | Codex `/root` | 不适用 | 不适用 | 不适用 | 不适用 | 不适用 |
| P4-API-01 | Codex `/root`、Codex `/root/cache_cleanup_impl`、Codex `/root/xlsx_scroll_design`、Codex `/root/xlsx_lazy_load_impl` | Codex `/root` 本会话集成复测；候选发布仍需全新会话 | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 五个公开包 `package.json`、DOCX/XLSX 公开结构类型、`scripts/ci/public-api-contract.json`、`docs/api/public-api-contract.md`、锁文件和路线图 | 产品运行行为、PDF 复杂预算、受保护的 `packages/docx-core/tsup.config.ts`、P4 打包/矩阵/流水线文件 | `P4-API-01`、公开接口检查、运行边界、P3 全量门禁回归 | P3 全量检查只因 P3 新增接口尚未冻结和声明引用私有包失败，保留该失败证据后先执行接口收口；公共声明不得依赖私有 `office-runtime`；五包统一 `0.2.0`，新增入口、错误、事件和兼容规则写入机器合同与文档；人工核对导出名称后更新指纹；类型、正式构建、公开接口与差异检查通过后回跑 P3 全量门禁 | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/P4-API-01/20260711T-p4-api/` | Codex `/root` | `public-api` | `2026-07-11` P3 兼容新增与 0.2.0 冻结 | 实例资源限制、图片/Worker 诊断、XLSX 按需子入口 | `1.0.0` 前保留已有 0.x 入口 | 不适用 |
| P4-SCOPE-01 | Codex `/root` | 本会话开发后自测；候选发布仍需全新会话 | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 六个工作区包清单和内部依赖、引用六个本地包名的产品源码/构建配置/demo/测试/脚本/当前文档、根清单与锁文件、路线图；五个公开包增加公开访问发布配置 | `output/` 历史证据、`node_modules/`、五份上游 `LICENSE`、`@extend-ai/react-*` 与 `extend-hq/ui` 来源记录、受保护的 `packages/docx-core/tsup.config.ts`、GitHub/npm 外部状态 | `P4-SCOPE-01`、`BB-PACK-CONSUMER`、`BB-RELEASE` | 修改前保存 `pnpm typecheck`、`pnpm build`、DOCX 集成、XLSX 结构和差异基线；精确替换六个本地包名，不改上游出处；重建安装状态；执行类型、正式构建、DOCX/XLSX 集成、真实五包 tgz、工作区外消费、正式浏览器黑盒、完整 `pnpm test:release` 和 `git diff --check`；不发布 | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/P4-SCOPE-01/20260711T-arcships-scope-01/` | Codex `/root` | `public-api` | `2026-07-11` 首次发布前改名 | `@arcships/{docx-core,xlsx-core,vue-docx,vue-xlsx,vue-extend}`；私有包为 `@arcships/office-runtime` | 首发前完成，不保留未发布的旧 npm 包 | 不适用 |
| P4-PACK-01 | Codex `/root`、Codex `/root/cache_cleanup_impl` | Codex `/root` 本会话集成复测；候选发布仍需全新会话 | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | `scripts/ci/p4-reproducible-pack.mjs`、五个公开包的 `package.json` 与 `LICENSE`、P4 就绪基线、路线图 | 产品源码、受保护的 `packages/docx-core/tsup.config.ts`、矩阵和发布工作流 | `P4-PACK-01`、双隔离构建、五包清单、外部消费 | 从当前含未提交修改的统一源码快照建立两份无 `node_modules/dist/output` 的隔离目录；每份冻结安装、正式构建和五包 pack；逐文件比较有效内容；保存 tgz/文件 SHA、日志和差异；用经上游核实的 MIT 原文进入五包；再运行真实 tgz 外部消费与 P4 就绪检查 | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/P4-PACK-01/<run-id>/` | Codex `/root` | 不适用 | 不适用 | 不适用 | 不适用 | 不适用 |
| P4-DOC-01 | Codex `/root`、Codex `/root/xlsx_lazy_load_impl` | Codex `/root` 本会话集成复测；候选发布仍需全新会话 | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | `README.md`、`RELEASE_NOTES.md`、`docs/INDEX.md`、`docs/migration-0.2.md`、兼容矩阵、四份历史迁移资料、两份核心测试文档、文档检查器和路线图 | 产品源码、包构建配置、受保护的 `packages/docx-core/tsup.config.ts` | `P4-DOC-01`、`pnpm test:docs`、候选命令复跑 | 删除过期当前状态；明确 0.2.0 尚未发布；登记安装、样式、Worker/WASM、Runtime、错误、限制、弃用与回退；PDF 必须真实可看可操作且只有默认 50 MiB 单一体积上限；检查链接和示例命令 | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/P4-DOC-01/<run-id>/` | Codex `/root` | `public-api` | `2026-07-11` 0.2.0 候选说明 | Runtime 与公开资源入口 | `1.0.0` 前保留旧公开入口 | 不适用 |
| P4-MATRIX-01 | Codex `/root`、Codex `/root/matrix_review` | Codex `/root` 本会话集成复测；候选发布仍需全新会话 | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | `scripts/ci/{compatibility-matrix.mjs,pack-consumer.mjs}`、`tests/blackbox/pack_consumer.py`、`tests/consumer/template/`、兼容矩阵文档、三个 Vue 包 peer 与锁文件、路线图 | 产品业务源码、demo、受保护的 `packages/docx-core/tsup.config.ts` | `P4-MATRIX-01`、`BB-PACK-CONSUMER`、三浏览器两 Vue 组合 | 只消费 P4 双构建产出的同一批五个 tgz；在工作区外安装、类型和正式构建；Chromium/Firefox/WebKit 均验证 DOCX 内容、XLSX 单元格和 PDF 翻页/缩放/旋转/缩略图/搜索/下载；核对 Worker/WASM SHA；首次失败最多重试一次并保留状态；完整矩阵才可 PASS | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/P4-MATRIX-01/<run-id>/` | Codex `/root` | `public-api` | `2026-07-11` Vue peer 上界收口 | `vue >=3.2.0 <4` | 不适用 | 不适用 |
| P4-RELEASE-01 | Codex `/root`、Codex `/root/xlsx_scroll_design` | Codex `/root` 本会话集成复测；候选发布仍需全新会话 | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 根脚本、`scripts/ci/{run-suite.mjs,prepare-release-artifact.mjs}`、`.github/workflows/{ci,release}.yml`、两份核心测试文档和路线图 | 产品源码、公开接口、包构建配置、受保护的 `packages/docx-core/tsup.config.ts` | `P4-RELEASE-01`、`BB-RELEASE`、制品校验与失败传播 | 显式运行 P0～P4 全部门禁；消费、矩阵和发布严格复用同批 tgz；候选清单绑定提交、统一版本、套件与 SHA；发布任务只下载并校验，不重建；五包先用临时标签，完整后提升正式标签，提升失败恢复旧标签并非零退出 | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/BB-RELEASE/<run-id>/` | Codex `/root` | 不适用 | 不适用 | 不适用 | 不适用 | 不适用 |
| P4-UX-DOCX-01 | Codex `/root` | 同会话浏览器自测；候选发布仍需全新会话 | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | `packages/vue-docx/src/components/`、必要的 `packages/vue-docx/src/composables/`、`apps/demo/src/pages/DocxViewerPage.vue`、`apps/demo/src/pages/DocxEditorPage.vue`、`tests/blackbox/` 中 DOCX 上游一致性用例、`scripts/ci/run-suite.mjs`、`docs/{plan/stabilization-roadmap.md,end-to-end-blackbox-test-plan.md,testing/agent-execution-runbook.md,upstream-parity-gap-audit.md}` | DOCX 核心解析、Worker/WASM 配置、XLSX/PDF 产品源码、受保护的 `packages/docx-core/tsup.config.ts`、无关用户改动 | `BB-UX-PARITY`（DOCX）、`BB-DOCX-WORKER`、`BB-P0-ROUTES` | 固定上游官方 `extend-hq/ui@f2ff2f9`；保存正式 preview 对照基线；`pnpm typecheck`；`pnpm build`；DOCX 组件与集成测试；正式 preview 的页码、缩放、缩略图、文件操作、选择编辑、撤销恢复、桌面/窄屏截图；`git diff --check` | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/UX-PARITY-BASELINE/20260711T035624+0800/local/`、`.../BB-UX-PARITY/20260711T-ux-parity-formal-19/local/`、`.../P1-CI-01/20260711T-ux-p4-consumer-formal-02/consumer/` | Codex `/root` | `public-api` | `2026-07-11` 上游一致性从补充材料提升为发布硬门禁 | 保持现有公开组件和控制器接口，补齐默认外壳与交互；必要新增属性必须有默认兼容值 | 不适用 | 不适用 |
| P4-UX-XLSX-01 | Codex `/root` | 同会话浏览器自测；候选发布仍需全新会话 | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | `packages/vue-xlsx/src/components/`、必要的 `packages/vue-xlsx/src/composables/`、`apps/demo/src/pages/XlsxViewerPage.vue`、`tests/{component,blackbox}/` 中 XLSX 一致性用例、`scripts/ci/run-suite.mjs`、`docs/{plan/stabilization-roadmap.md,end-to-end-blackbox-test-plan.md,testing/agent-execution-runbook.md,upstream-parity-gap-audit.md}` | DOCX/PDF 产品源码、XLSX Worker/WASM 运行边界、受保护的 `packages/docx-core/tsup.config.ts`、无关用户改动 | `BB-UX-PARITY`（XLSX）、`BB-PERF-XLSX`、`BB-P0-ROUTES` | 固定上游官方 `extend-hq/ui@f2ff2f9`；保存正式 preview 对照基线；`pnpm typecheck`；`pnpm build`；XLSX 组件测试；正式 preview 的公式栏、样式、合并、选择、编辑、冻结、行列尺寸、图表/图片、桌面/笔记本截图；`git diff --check` | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/UX-PARITY-BASELINE/20260711T035624+0800/local/`、`.../BB-UX-PARITY/20260711T-ux-parity-formal-19/local/`、`.../P1-CI-01/20260711T-ux-p4-consumer-formal-02/consumer/` | Codex `/root` | `public-api` | `2026-07-11` 上游一致性从补充材料提升为发布硬门禁 | 保持现有公开组件和控制器接口，补齐默认外壳与交互 | 不适用 | 不适用 |
| P4-FIDELITY-DOCX-01 | Codex `/root` | 本会话开发后自测；候选发布仍需全新会话 | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | `packages/vue-docx/src/components/{DocxViewer.vue,DocxViewerToolbar.vue,DocxDocumentSurface.vue,DocxViewerRoot.vue,DocxPageSurface.vue,DocxTrackedChangeGutter.vue}`、`apps/demo/src/pages/DocxViewerPage.vue`、专用 DOCX 样例生成与组件/黑盒测试、路线图 | XLSX/PDF、DOCX Worker/WASM 与核心解析、受保护的 `packages/docx-core/tsup.config.ts`、无关用户改动 | DOCX Viewer 批注/修订组件回归、`BB-DOCX-WORKER` 对应 Viewer 用例、`BB-P0-ROUTES` DOCX 子集 | 修改前五项基础命令均为 0，但公开 Viewer/Demo 功能断言为 1；新增兼容公开开关、从公开模型只读派生批注和修订、Demo 专用样例；验证 `pnpm typecheck`、`pnpm build`、DOCX 集成与组件测试、正式 preview 截图/控制台/网络、`git diff --check` | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/P4-FIDELITY-DOCX-01/20260711T102803+0800/` | Codex `/root` | `public-api` | `2026-07-11` 兼容新增 | `DocxViewer.defaultShowTrackedChanges`、`DocxViewer.defaultShowComments` 与可见开关 | 不适用 | 不适用 |
| P4-FIDELITY-XLSX-WORKER-01 | Codex `/root` | 本会话开发后自测；候选发布仍需全新会话 | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | `packages/xlsx-core/src/{xlsx-worker.ts,worker-client.ts,types/worksheet-types.ts,images/grid-render.ts}`、`packages/vue-xlsx/src/{components/XlsxGrid.vue,composables/useXlsxViewerController.ts}`、XLSX Demo 样例生成、直接组件/黑盒测试和路线图 | DOCX/PDF、XLSX 公式与编辑命令、受保护的 `packages/docx-core/tsup.config.ts`、无关用户改动 | Worker 快照/资源组件回归、`BB-PERF-XLSX` 图片子集、`BB-P0-ROUTES` XLSX 子集 | 复用本轮开始时五项全绿基线，并保存 Worker 快照缺图片、网格不读精确合并区域的功能断言失败；新增 Worker 可传输视觉快照和实例清理、精确合并区域；验证类型、构建、XLSX 结构/组件、正式 preview Worker/主线程对照、截图/网络/控制台和差异检查 | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/P4-FIDELITY-XLSX-WORKER-01/20260711T104342+0800/` | Codex `/root` | 内部运行结构 | 不适用 | 不新增公开接入方式 | 不适用 | 不适用 |

### 12.2 状态与证据记录

任务首次进入 `in_progress` 时在下表追加一行；后续状态变化更新同一行。时间使用带时区的 ISO 8601 格式，提交号使用完整值，证据位置填写持续集成链接或证据包路径与校验值。

| 任务 ID | 负责人 | 验证者 | 开始时间 | 结束时间 | 基线提交 | 结果提交 | 命令与退出码 | 环境 | 证据位置 | 状态 |
|---|---|---|---|---|---|---|---|---|---|---|
| P4-FIDELITY-XLSX-WORKER-01 | Codex `/root` | 本会话开发后自测；不是独立验收 | `2026-07-11T10:43:42+0800` | `2026-07-11T10:56:07+0800` | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 工作区未提交；无结果提交 | 基线五项均为 0，Worker 图片/合并功能断言=1。实现后 `pnpm typecheck=0`、`pnpm build=0`、DOCX 集成=0（54 通过）、XLSX 结构=0、XLSX 单元=0（22 通过）、完整组件套件=0、`git diff --check=0`。首个浏览器运行两次均因 Playwright 参数写法错误而未进入产品断言；修正脚本后的首次产品运行暴露主线程快照缺精确合并数组并两次 FAIL；补齐后新运行首轮 PASS，Worker/主线程均为 1 个合并区域、次单元格选择回到 A7、图片 ready，四项几何差值均为 0，网络/控制台/页面异常为空。 | macOS arm64；Node `v22.13.0`；pnpm `9.0.6`；Python `3.12.2`；Playwright Chromium；工作区有大量既有未提交改动 | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/P4-FIDELITY-XLSX-WORKER-01/20260711T104342+0800/`；最终摘要 SHA-256 `33a6438c954dc72af1688894883839638369971790f546f4d4884f8bacd6db7b`；样例 SHA-256 `bbc0cd1e61a8ec9ee9620870dcbab95490b7d36f2d655aeb92beeebba10c6dc0` | `completed` |
| P4-FIDELITY-DOCX-01 | Codex `/root` | 本会话开发后自测；不是独立验收 | `2026-07-11T10:28:03+0800` | `2026-07-11T10:39:10+0800` | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 工作区未提交；无结果提交 | 基线五项均为 0，但 Viewer/Demo 功能断言=1；实现后 `pnpm typecheck=0`、`pnpm build=0`、DOCX 集成=0（54 通过）、XLSX 结构=0、定向组件=0（13 通过）、`git diff --check=0`。正式 preview 首轮及重试均因 Demo 受控属性未回写而 FAIL；修复双向绑定后新运行首轮 PASS，3 张卡片、隐藏/恢复、Worker/WASM、网络、控制台和页面异常断言全部通过。 | macOS arm64；Node `v22.13.0`；pnpm `9.0.6`；Python `3.12.2`；Playwright Chromium；工作区有大量既有未提交改动 | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/P4-FIDELITY-DOCX-01/20260711T102803+0800/`；最终摘要 SHA-256 `0288c7d29cd5d09507376cebc8dfa41e752b695a0a515b7bc019f9deb3ae506f`；样例 SHA-256 `cdda780c1aabcd01cf0fc7f9b4d17e71ce96880b1e9abdebc11a29e79f1849c0` | `completed` |
| P4-FIDELITY-CONTENT-02 | Codex `/root` | 本会话开发后自测；不是独立验收 | `2026-07-11T13:49:16+0800` | `2026-07-11T14:04:05+0800` | `4ca9acbe993419f595ea8d06616cdaeb2aa42b94` | 本次结果已提交；提交号见 Git 历史 | 前置两项已完成。允许路径未超出登记范围，受保护的 `packages/docx-core/tsup.config.ts` 未修改。基线 `pnpm typecheck=0`、`pnpm build=0`、DOCX 集成 54 项通过、XLSX 结构通过、`git diff --check=0`。实现 DOCX 脚注/尾注正文、全文搜索与高亮；新增 XLSX 单元格超链接/批注公开数据和网格交互、数据条/色阶/图标集、迷你图、形状与只读控件层；条件格式统计在每次绘制中按规则缓存，避免可见单元格重复扫描。Demo 使用真实 OOXML 样例。完成后全仓类型和正式构建通过，52 项组件行为、组件烟测、17 项 P3 回归、完整单元套件、DOCX 54 项集成、XLSX 结构、21 份样例清单、样例真实打开和差异检查通过。第一轮浏览器脚本因错误使用非哈希路由两次未进入产品；修正后的第一次产品运行发现搜索测试词不在正文、形状样例未写入默认命名空间的绘图根并两次失败；修复后新的正式运行两用例均首轮 PASS，无重试、FLAKY、控制台、页面异常、失败请求或 4xx/5xx；最终 XLSX 用例明确以只读 Worker 路径验证六类内容和内部超链接。首次 tgz 消费验证与单元套件并行构建导致 `dist` 清理竞争而失败；改为串行后五个真实 tgz 清单、哈希、工作区外安装、公开 JS/类型/样式和 Worker/WASM 浏览器消费全部 PASS。新增两个公开只读类型，按兼容规则更新公开接口指纹。 | macOS arm64；Node `v22.13.0`；pnpm `9.0.6`；Python `3.12.2`；Playwright Chromium | `output/acceptance/4ca9acbe993419f595ea8d06616cdaeb2aa42b94/P4-FIDELITY-CONTENT-02/20260711T134916+0800/`；最终浏览器 `browser-final-04/summary.json` SHA-256 `61c4b69d433dab0b5c37b3c016085a5c1d03ad2769a473fa03f5270a2ebf850d`；最终组件摘要 SHA-256 `4109be7313b4b81a1dfb9822962f9fd06f708ebfd6d9572b9549e66a292cb180`；最终单元摘要 SHA-256 `b4792411c2696c02c67754a0ac1c1fc7f6b114d8f2f34099aef0bd314872f88e`；最终真实 tgz 消费摘要 SHA-256 `a187e56b1bf1c4f5f8909a4a3284d658fc57a6d0864f1a141e16a28b56ef8899`；文档门禁 SHA-256 `f3fbb687d2b4a7496707343ce1a61ec77ea7d625eac4fda0ab21337186c444f8`；DOCX 样例 SHA-256 `9e8d1852f7cde12e83bd91d0452a2941fff5b4b55f901235d1ba5c22f1f60e09`；XLSX 样例 SHA-256 `d012d0bcaf13a461a21e676642403b76a6ec13799b3126f2f5d50203c400694b` | `completed` |
| P0-DOCX-01 | Codex `/root` | Codex `/root/quick_p0_verifier` | `2026-07-10T12:12:35+0800` | `2026-07-10T13:19:51+08:00` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | `pnpm typecheck=0`；`pnpm build=0`；`verify-integration=0（52 通过）`；`structure=0（保留既有 Vue warning）`；`BB-DOCX-WORKER=PASS`；`git diff --check` 待本轮尾检 | macOS；Node `v22.13.0`；pnpm `9.0.6`；Python `3.12.2`；系统 Chrome `150.0.7871.115` | `.../BB-DOCX-WORKER/20260710T121023+0800/`（失败基线）；`.../20260710T131750+0800/`（实现者正式自测 PASS）；`.../20260710T134000+0800-quick-verifier/`（独立 PASS） | `ready_for_gate` |
| P0-DOCX-02 | Codex `/root` | Codex `/root/quick_p0_verifier` | `2026-07-10T13:19:51+08:00` | `2026-07-10T14:11:25+08:00` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 修复前 `vue-tsc=2`：示例模型缺字段、主题类型不匹配；修复后 `vue-tsc=0`；`pnpm typecheck=0`；`pnpm build=0`；`verify-integration=0（54 通过）`；`structure=0`（保留既有 Vue 生命周期 warning）；`git diff --check=0`；`BB-P0-ROUTES=PASS`（实现者与独立正式 preview） | macOS；Node `v22.13.0`；pnpm `9.0.6`；Python `3.12`；系统 Chrome `150.0.7871.115`；工作区有用户改动 | `.../BB-P0-ROUTES/20260710T132252+0800/`（修复前编辑输入与键盘重试均失败）；`.../20260710T133709+0800/`（实现者 PASS，含检查完整日志）；`.../20260710T140943+0800-final-verifier/`（独立 PASS）。`.../20260710T134145+0800-independent-editor/` 与 `.../20260710T140745+0800-final-verifier/` 仅为验证脚本自身错误，均标为 BLOCKED 并保留。 | `ready_for_gate` |
| P0-TYPE-01 | Codex `/root` | Codex `/root/typecheck_audit` | `2026-07-10T14:13:19+08:00` | `2026-07-10T14:24:34+08:00` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 基线 `pnpm typecheck=0` 但四个 Vue 项目用普通 `tsc`，demo 无 `vue-tsc`；临时 `.vue` 探针使根 `pnpm typecheck=2`（TS2322），删除后 `pnpm install --frozen-lockfile=0`、`pnpm typecheck=0`、`pnpm build=0`（日志证明先类型检查）、`verify-integration=0（54 通过）`、`structure=0`（保留既有 Vue 生命周期 warning）、`git diff --check=0`；`--listFiles` 确认跨包检查读 `packages/*/src` 而非 `dist` | macOS；Node `v22.13.0`；pnpm `9.0.6`；TypeScript `5.9.3`；vue-tsc `3.3.6`；工作区有用户改动 | `.../P0-TYPE-01/20260710T141925+0800/`（实现者 PASS）；`.../20260710T142145+0800-independent/`（独立 PASS，探针与恢复证据） | `ready_for_gate` |
| P0-URL-01 | Codex `/root` | Codex `/root/sec_url_harness` | `2026-07-10T14:28:47+08:00` | `2026-07-10T14:51:53+08:00` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 基线 `BB-SEC-URL=BLOCKED`：PDF 页面没有公开策略和安全状态；修复后 `pnpm typecheck=0`；`pnpm build=0`；`verify-integration=0（54 通过）`；`structure=0`（保留既有 Vue 生命周期 warning）；`git diff --check=0`；实现者 PDF 范围 `SEC-001/003/004=PASS`；独立正式构建复核 `SEC-001/003/004=PASS`。整个 `BB-SEC-URL` 仍 BLOCKED，唯一原因是 P0-URL-02 的 XLSX `SEC-002` 尚未实现。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；Python `3.12.2`；系统 Chrome `150.0.7871.115`；工作区有用户改动 | `.../BB-SEC-URL/20260710T123539+0800/`（失败基线）；`.../20260710T143628+0800/`（实现者，attempt-2 发现 sandboxed Blob iframe 请求失败，attempt-3 修复后 PASS，含允许 HTTPS）；`.../20260710T144645+0800-final-independent/`（独立 PASS） | `ready_for_gate` |
| P0-URL-02 | Codex `/root` | Codex `/root/sec_url_harness` | `2026-07-10T14:54:14+08:00` | `2026-07-10T15:20:29+08:00` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 基线 `SEC-002=FAIL`：正式预览输入 `javascript:void(window.__blackboxMarker='changed')` 后，加载虽失败，原始下载按钮仍可点击并把 marker 从 `unchanged` 改为 `changed`；根因是裸 `fetch(src)`、`canDownload=Boolean(src)` 和 `downloadUrl(ctx.src)`。早期 `145414` 仅记录缺少稳定入口导致的 BLOCKED，不替代此漏洞证据。修复后 `pnpm typecheck=0`；`pnpm build=0`；`verify-integration=0（54 通过）`；`structure=0`（保留既有 Vue 生命周期 warning）；`git diff --check=0`；实现者正式 preview `BB-SEC-URL=PASS`（attempt-2 是验证脚本把安全的“不存在旧下载按钮”误判为失败，保留后以 attempt-3 通过）；独立正式构建 `P0-URL-02=PASS`，且与 P0-URL-01 的独立结果合并后 `BB-SEC-URL` P0 的 `SEC-001` 至 `SEC-004=PASS`。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；Python `3.12.2`；系统 Chrome `150.0.7871.115`；工作区有用户改动 | `.../BB-SEC-URL/20260710T145414+0800-xlsx-baseline/`（入口 BLOCKED）；`.../20260710T145442+0800-xlsx-url-baseline/`（真实 FAIL）；`.../20260710T150543+0800-xlsx/summary-attempt-3.json`（实现者 PASS）；`.../20260710T151315+0800-xlsx-final2-independent/attempt-1/summary.json`（独立 PASS）；`.../20260710T151315+0800-xlsx-final2-independent/summary.json`（P0 汇总 PASS） | `ready_for_gate` |
| P0-RACE-01 | Codex `/root` | Codex `/root/race_baseline` | `2026-07-10T15:22:55+08:00` | `2026-07-10T15:47:10+08:00` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 依赖 P0-DOCX-01=`ready_for_gate`。基线 `RACE-001=FAIL`：正式 preview 中慢 `demo.docx`（5 秒）→快 `invoice-table.docx`；B 先显示 1 页 `INVOICE`，A 晚到后覆盖成 6 页 `MASTER SERVICES AGREEMENT`，但文件名仍显示 B。首次与唯一重试均失败，且无 console/pageerror/requestfailed。实现后 `pnpm typecheck=0`、`pnpm build=0`；DOCX `RACE-001/RACE-004/RACE-005=PASS`，XLSX 本地 FileReader 与延迟加载路径也通过，但 `RACE-002` 的“切到 B 并编辑 B”首次与唯一重试均失败：正式 demo 未加载 `vue-xlsx` 组件样式，`.xlsx-viewer__body` 和 `.xlsx-grid` 高度均为 0，公式输入保持禁用。因此不能用只验证加载成功冒充竞态通过。阻断于 P0-PACK-01→P0-STYLE-01 的公开样式出口；完成后恢复本任务并以正式 preview 复测。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；Python `3.12.2`；系统 Chrome `150.0.7871.115`；工作区有用户改动 | `.../BB-RACE/20260710T152255+0800-docx-baseline/traces/RACE-001-demo-to-invoice/`（首次及重试 FAIL）；`.../BB-RACE/20260710T153732+0800-race-self/docx-summary.json`（RACE-001/004/005 PASS）；`.../xlsx-summary.json`（RACE-002 编辑两次 FAIL，FileReader/延迟路径 PASS）；`.../traces/xlsx-layout.json`（grid 高度 0）；`demo.docx` 与 `legal-contract.docx` SHA 相同的无效探测也已保留，不计产品失败 | `blocked` |
| P0-PACK-01 | Codex `/root` | Codex `/root/race_baseline` | `2026-07-10T15:48:22+08:00` | `2026-07-10T16:18:06+08:00` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 依赖 P0-TYPE-01=`ready_for_gate`。基线五个真实 `npm pack` 均退出 0，却只包含 `src/`、测试和配置，全部缺少声明的 `dist/index.js` 与 `dist/index.d.ts`；包清单还泄露源码和测试。修复后 `pnpm typecheck=0`；`pnpm build=0`；五个真实 tgz 均只包含 `dist` 和 package manifest，根 JS/类型入口完整、无 `workspace:*`、无 `src/`/测试目录；独立消费者 `npm install=0`、根出口解析/非符号链接=`PASS`、`tsc=0`、`vue-tsc=0`、Vite 生产构建=`0`。独立 `PACK-003` 的 Worker/WASM 失败已明确移交 P0-PACK-02，不能当作本任务的资源通过或发布通过。 | macOS；Node `v22.13.0`；npm `11.0.0`；pnpm `9.0.6`；Python `3.12.2`；系统 Chrome `150.0.7871.115`；工作区有用户改动 | `.../BB-PACK-CONSUMER/20260710T154710+0800-pack-baseline/tgz/`（失败基线）；`.../20260710T160437+0800/tgz-manifest.json`（实现者 PASS）；`.../20260710T160437+0800/independent-20260710T160948+0800/summary.json`（独立：PACK-001/002 PASS，PACK-003 失败移交） | `ready_for_gate` |
| P0-PACK-02 | Codex `/root` | Codex `/root/pack02_independent`（同会话交叉复测；不是候选发布独立验收） | `2026-07-10T16:24:19+08:00` | `2026-07-10T17:08:49+08:00` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 依赖 P0-DOCX-01/P0-PACK-01 均为 `ready_for_gate`。失败基线：真实 tgz 消费项目中 DOCX Worker 请求已失效哈希 WASM，收到 SPA `text/html` 后报 `WASM_LOAD_FAILED`；XLSX 夹具未设只读而显示 `not-worker`。修复后 `pnpm typecheck=0`、`pnpm build=0`、`verify-integration=0（54 通过）`、`structure=0`（保留既有 Vue 生命周期 warning）、`git diff --check=0`；五包 tgz 清单 `PACK-001=PASS`，当前构建的两个 Worker 与两份 WASM SHA-256 均与 tgz 内文件一致；真实 tgz 消费项目 `PACK-000=PASS`（开发模式）和 `PACK-003=PASS`（正式 preview），DOCX/XLSX 均 ready，两个 Worker 与两个 WASM 均 200，WASM 为 `application/wasm`，无 4xx、console/pageerror/requestfailed/download；`BB-DOCX-WORKER` 的 DOCX-001/002/003 正式 preview 首次均 PASS。两个早期浏览器结果仅因测试脚本把 Vite 的 `?import&url` 模块当作二进制、或只识别开发模式 `worker_file` URL 而失败，已保留并在同目录写明；不计产品失败。 | macOS；Node `v22.13.0`；npm `11.0.0`；pnpm `9.0.6`；Python `3.12.2`；系统 Chrome `150.0.7871.115`；工作区有用户改动 | `.../BB-PACK-CONSUMER/20260710T160437+0800/independent-20260710T160948+0800/summary.json`（失败基线）；`.../BB-PACK-CONSUMER/20260710T164800+0800/tgz-manifest.json`（五包及 SHA-256）；`.../PACK-000-dev-public-assets-corrected/attempt-1/result.json`（开发 PASS）；`.../PACK-003-final-corrected/attempt-1/result.json`（正式 preview PASS）；`.../independent-agent-20260710T170329+0800/result.json`（同会话交叉复测 PASS）；`.../BB-DOCX-WORKER/20260710T170732+0800-corrected/summary.json`（DOCX-001/002/003 PASS）。候选发布仍需全新会话跑 BB-RELEASE。 | `ready_for_gate` |

| P0-STYLE-01 | Codex `/root` | Codex `/root/style_crosscheck`（同会话交叉复测；不是候选发布独立验收） | `2026-07-10T17:10:09+08:00` | `2026-07-10T17:30:02+08:00` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 依赖 P0-PACK-01=`ready_for_gate`。基线真实 tgz 的三个 `./style.css` 均为 `ERR_PACKAGE_PATH_NOT_EXPORTED`，仅旧 `vue-extend/dist/index.css` 可解析；RACE-002 两次失败且 XLSX body/grid 高度为 0。修复后 `pnpm typecheck=0`、`pnpm build=0`、`verify-integration=0（54 通过）`、`structure=0`（保留既有 Vue 生命周期 warning）、`git diff --check=0`；五包真实 tgz 的 PACK-001 清单通过，STYLE-001 确认三个 `./style.css` 指向真实 `dist/style.css` 且声明为 side effect，旧 `vue-extend/dist/index.css` 兼容映射仍可用；正式 tgz consumer 的 CSS=200 text/css、Worker/WASM 均 200、XLSX body/grid 高度 209；六个正式 demo 路由均 PASS；RACE-002 首次 PASS，慢 A 被取消后 B 保持 `sales-table.xlsx` 和 `=RACE_B_VALUE`，body/grid 高度 391。早期两组 RACE 重测是测试脚本延迟计数和公式字符串断言错误，保留说明后用全新故障服务器重测，不计产品失败。 | macOS；Node `v22.13.0`；npm `11.0.0`；pnpm `9.0.6`；Python `3.12.2`；系统 Chrome `150.0.7871.115`；工作区有用户改动 | `.../P0-STYLE-01/20260710T171141+0800/css-resolve-baseline.json`（FAIL 基线）；`.../css-resolve-final.json`、`tgz-manifest.json`、`style-tgz-manifest.json`（PASS）；`.../STYLE-003-preview/attempt-1/result.json`（正式 tgz consumer PASS）；`.../BB-P0-ROUTES/attempt-1/summary.json`（六路由 PASS）；`.../crosscheck-agent-20260710T172559+0800/result.json`（同会话交叉复测 PASS）；`.../BB-RACE/20260710T172137+0800-style-retest-final/summary.json`（RACE-002 PASS）。候选发布仍需全新会话跑 BB-RELEASE。 | `ready_for_gate` |

| P0-RACE-01（恢复） | Codex `/root` | 同会话已有证据复核 | `2026-07-10T17:31:00+08:00` | `2026-07-10T17:31:42+08:00` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | P0-STYLE-01 已达 `ready_for_gate`，原样式阻断解除。DOCX 正式 preview 的 RACE-001（慢 A 到快 B）、RACE-004（卸载）和 RACE-005（20 次有效→损坏→有效）均 PASS；XLSX 的本地文件、延迟解析路径均 PASS，新增正式 preview RACE-002 首次 PASS：慢 A 被取消，B 的文件、公式和非零网格高度未被覆盖。`pnpm typecheck=0`、`pnpm build=0`、`verify-integration=0（54 通过）`、`structure=0`（保留既有 Vue 生命周期 warning）、`git diff --check=0`。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；Python `3.12.2`；系统 Chrome `150.0.7871.115` | `.../BB-RACE/20260710T153732+0800-race-self/docx-summary.json`；`.../xlsx-summary.json`；`.../BB-RACE/20260710T172137+0800-style-retest-final/summary.json` | `ready_for_gate` |

| P0-DEMO-01 | Codex `/root` | Codex `/root/routes_crosscheck`（同会话交叉复测；不是候选发布独立验收） | `2026-07-10T17:32:30+08:00` | `2026-07-10T17:35:01+08:00` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 依赖 P0-DOCX-01、P0-DOCX-02、P0-URL-01、P0-URL-02、P0-RACE-01 均为 `ready_for_gate`。正式 preview 自测中六个 hash 路由均直接访问、等待业务 ready 后刷新并再次 ready；无 4xx/5xx、console error/warning、pageerror、requestfailed 或下载。XLSX 直接访问与刷新后 grid 均为 391px 高。交叉复测使用另一端口和六个新 context，全部 PASS，XLSX grid 265px 高。`git diff --check=0`。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；Python `3.12.2`；系统 Chrome `150.0.7871.115`；工作区有用户改动 | `.../BB-P0-ROUTES/20260710T173253+0800-self/summary.json`（实现者 PASS，含直接访问和刷新）；`.../20260710T173324+0800-crosscheck-agent/summary.json`（同会话交叉复测 PASS）。候选发布仍需全新会话跑 BB-RELEASE。 | `ready_for_gate` |

| P1-CI-01 | Codex `/root` | Codex `/root/p1_ci_crosscheck`（同会话交叉复测；不是候选发布独立验收） | `2026-07-10T17:52:28+0800` | `2026-07-10T18:09:09+0800` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 直接依赖 P0-TYPE-01=`ready_for_gate`。批次 B1 原要求 P0 安全提交；因用户明确开始 P1，按共同检查例外仅修改测试编排与 CI，不能据此登记 `LAST_SAFE_COMMIT`。基线九个统一命令均因脚本不存在非零失败；修复后 `pnpm install --frozen-lockfile=0`、`pnpm check=0`，覆盖 typecheck、全仓 build、DOCX 54 项、无 Vue warning 的 XLSX 结构检查、DOCX/XLSX 公共组件渲染、正式 preview 路由、五包真实 tgz、失败传播和 `git diff --check`。故障探针让 `unit:docx-integration` 退出 97，父套件正确退出 1并将后续标为 `NOT_RUN`，自检命令退出 0。工作区外排除 `node_modules`/`dist` 的快照重新 frozen install 和 `pnpm check=0`。初始浏览器启动 BLOCKED 是本机缺少配套 Chromium，系统 Chrome 回退被系统杀掉；安装登记版本后最终完整运行首次 PASS。初始 tgz 失败是执行器把 prepack 日志当 JSON，修正解析后最终完整运行首次 PASS；两类首轮证据均保留，不计产品失败。交叉复测的 frozen install、`test`、`test:e2e`、`test:pack`、`ci:self-test`、`git diff --check` 全部首次退出 0。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；Python `3.12`；本地 Playwright `1.50.0` + Chromium `133.0.6943.16`；CI 固定 Python Playwright `1.61.0`；工作区有受保护的既有改动 | `.../P1-CI-01/20260710T175228+0800/baseline/summary.json`（缺失命令 FAIL 基线）；`.../20260710T175228+0800-final2/check/summary.json`（PASS，SHA-256 `764259f1aa81de3093e85879c82e166d19bff15e851ab1b491e139cb7866b647`）；`.../20260710T175228+0800-clean-snapshot/check/summary.json`（无旧依赖/产物 PASS，SHA-256 `51b16cc0d7ae5bf98aa50a1206310f7f7a5ca930162b8884a69fd1c27077d8f7`）；`.../20260710T175228+0800-crosscheck/result.json`（同会话交叉复测 PASS，SHA-256 `f6a07b894c2ad740bb435e290cd5830665311110138fd3da0c8a8308bd94fd28`）。候选发布仍需全新会话跑 BB-RELEASE。 | `ready_for_gate` |

| P1-UNIT-01 | Codex `/root` | 本会话复测；不是候选发布独立验收 | `2026-07-10T18:14:40+0800` | `2026-07-10T18:16:42+0800` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 依赖 P1-CI-01=`ready_for_gate`。基线 `test:unit=0`，但缺少十类要求用例。新增 Node 测试后 `test:unit=0`：9 个核心行为用例首次全部通过，覆盖 DOCX 正常解析与往返、空/损坏输入稳定错误、Abort 终止 Worker、旧加载取消与最新结果；覆盖 XLSX 正常解析、空/损坏输入、公式/计算值/样式往返、Worker 乱序响应映射和 dispose 取消。既有 DOCX 54 项和 XLSX 结构检查继续通过且无 Vue warning。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；工作区有受保护的既有改动 | `.../P1-UNIT-01/20260710T181440+0800/baseline/coverage-gap.json`（FAIL 覆盖基线）；`.../final/unit/summary.json`（PASS） | `ready_for_gate` |

| P1-FIXTURE-01 | Codex `/root` | 本会话复测；不是候选发布独立验收 | `2026-07-10T18:17:02+0800` | `2026-07-10T18:27:54+0800` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 依赖 P1-UNIT-01=`ready_for_gate`。基线旧 manifest 只有名称/字节数，不能检测内容漂移。新增 20 项材料清单，包含合成来源、类别、预期行为、字节数和 SHA-256；校验器核对完整文件集合、类型签名、显式别名、四类覆盖和固定生成器。首次复现发现 17 项与旧生成器输出不一致，进一步发现 XLSX/PDF 共 9 项跨目录生成不稳定；修复 OOXML 时间戳/ZIP 元数据、ReportLab 固定模式和图片内容命名后，两次独立临时目录生成的 21 个文件逐字节一致。重新生成正式材料后 `test:unit=0`，20 项校验 PASS；篡改一项 SHA 的副本使校验器退出 1。最终差异检查发现 Git 把 PDF 定长记录里的规范尾空格误当源码尾空格，新增精确目录规则将这些 PDF 标记为二进制后，根 `git diff --check=0`。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；Python `3.12` | `.../P1-FIXTURE-01/20260710T181702+0800/baseline/summary.json`（FAIL）；`.../generator-repro/summary.json`（生成不稳定 FAIL）；`.../generator-final/summary.json`（两轮一致 PASS）；`.../fault-probe/probe-summary.json`（故障拦截 PASS）；`.../final2/unit/summary.json`（PASS） | `ready_for_gate` |

| P1-VUE-01 | Codex `/root` | 本会话复测；不是候选发布独立验收 | `2026-07-10T18:24:08+0800` | `2026-07-10T18:27:54+0800` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 依赖 P1-CI-01=`ready_for_gate`。基线 `test:component=0`，但只有 DOCX/XLSX 空状态冒烟检查，按通过标准记为覆盖失败。新增 7 组公开接口测试：DOCX 空/加载/成功/错误及事件、卸载取消且不发过期错误；XLSX 三状态、属性、撤销/重做快捷键；组合函数 Worker 加载、公开诊断和卸载终止；PDF 空/成功/错误及事件、URL 加载卸载取消、下载临时对象 URL 回收。首次实现运行因自建无 DOM 渲染器缺少选择框与输入框所需节点方法而失败，补齐测试基础设施后连续两次通过。`pnpm typecheck=0`；正式重建后的 `pnpm test:component=0`，7 组测试和既有冒烟检查全部通过且无 Vue warning；`git diff --check=0`。未修改产品源码。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；工作区有受保护的既有改动 | `.../P1-VUE-01/20260710T182408+0800/baseline/coverage-gap.json`（FAIL 覆盖基线）；`.../first/component/summary.json`、`.../second/component/summary.json`（测试基础设施失败，保留）；`.../third/component/summary.json`（PASS）；`.../final/component/summary.json`（正式重建 PASS，SHA-256 `3203928abb364e97d44154b9d6f3e823af81104ae62fc6fa4ea0b8d4014b6b24`） | `ready_for_gate` |

| P1-E2E-01 | Codex `/root` | 本会话复测；不是候选发布独立验收 | `2026-07-10T18:28:56+0800` | `2026-07-10T18:32:53+0800` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 依赖 P0-DEMO-01、P1-CI-01 均为 `ready_for_gate`。基线正式预览测试只检查四个路由直接访问和刷新，缺少业务流程及 trace。新增三个正式构建流程，每个流程使用新 context、导航前监听、首次失败最多重试一次，并始终保存截图、控制台、网络、页面异常、下载和 trace：DOCX 切换→损坏错误→恢复；DOCX 编辑→撤销→再次编辑→导出；XLSX 工作表切换→公式编辑→撤销→损坏错误→恢复→源文件下载。首次运行 DOCX 两项一次通过；XLSX 两次稳定失败，证实公式已改成 `=41+1` 但撤销按钮仍禁用。根因是 `canUndo/canRedo` 计算值未依赖历史版本，浅引用数组原地更新后不会重算；最小修复后单独复测和最终全套均首次 PASS。新增稳定 `data-testid`，没有测试环境后门。`pnpm typecheck=0`、`pnpm build=0`、最终 `pnpm test:blackbox=0`，路由与三流程均 PASS；`git diff --check=0`。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；Python Playwright `1.50.0`；Chromium `133.0.6943.16`；工作区有受保护的既有改动 | `.../P1-E2E-01/20260710T182856+0800/baseline/coverage-gap.json`（FAIL 覆盖基线）；`.../first/blackbox/formal-preview-workflows/`（XLSX 首次及重试 FAIL，DOCX PASS）；`.../second-workflows/summary.json`（修复后 PASS）；`.../final/blackbox/summary.json`（正式全套 PASS，SHA-256 `88f13b8cc1d39a2a6bcd64bf8476ed5c99d978a7779b0870357c8ca942c3ff10`） | `ready_for_gate` |

| P1-CONSOLE-01 | Codex `/root` | 本会话复测；不是候选发布独立验收 | `2026-07-10T18:33:48+0800` | `2026-07-10T18:37:06+0800` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 依赖 P1-E2E-01=`ready_for_gate`。基线路由与业务流程各自重复监听浏览器事件，没有统一允许清单。新增唯一 `BrowserEvidence` 采集器，路由和业务流程都在导航前挂接并保存 `console.json`、`page-errors.json`、`network.json`、`downloads.json`、`violations.json`；允许清单四类均为空。控制台 error、Vue warning、页面异常、请求失败、404 响应五个探针各运行首次和唯一重试，父命令均退出 1，对应证据分别出现 1 项违规；请求失败与 404 同时触发的浏览器控制台错误也被保留。移除探针后，正式路由与三个业务流程全部首次 PASS，无违规项。`git diff --check=0`。故障注入只在黑盒脚本显式环境变量中存在，产品没有测试后门。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；Python Playwright `1.50.0`；Chromium `133.0.6943.16`；工作区有受保护的既有改动 | `.../P1-CONSOLE-01/20260710T183348+0800/baseline/coverage-gap.json`（FAIL 覆盖基线）；`.../fault-probes/{console,vue-warning,pageerror,requestfailed,response404}/`（均被拦截）；`.../final/blackbox/summary.json`（PASS，SHA-256 `e68a29471c61e7a9fb3ac52b50acd47545703d25d5113285c9e17ce9b65b60a0`） | `ready_for_gate` |

| P1-RACE-TEST-01 | Codex `/root` | 本会话复测；不是候选发布独立验收 | `2026-07-10T18:38:13+0800` | `2026-07-10T18:41:20+0800` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 依赖 P0-RACE-01、P1-E2E-01 均为 `ready_for_gate`。基线旧 `BB-RACE` 脚本只存在 output 证据目录，全新检出无法运行。新增正式构建延迟服务器和 6 个仓库内黑盒用例：DOCX/XLSX 各自覆盖慢 A→快 B、20 次切换和卸载；所有用例使用新 context、统一浏览器采集器、截图和 trace，首次失败才重试。通过页面前置脚本包装浏览器标准 Worker、`createObjectURL`/`revokeObjectURL`，只统计公开平台资源，不读取模块私有变量。首次单套与最终根黑盒均六项首次 PASS：DOCX 最终为 invoice 且含取消诊断；XLSX 最终为 sales；XLSX 20 次切换创建/回收对象 URL 均为 5，活动数为 0；Worker 卸载前活动数 1，卸载后终止数 1、活动数 0；所有用例无未预期控制台、页面、请求或响应异常。旧 P0 失败基线已证明同一慢 A→快 B 场景在移除代次/取消保护时两次稳定失败。`git diff --check=0`。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；Python Playwright `1.50.0`；Chromium `133.0.6943.16`；工作区有受保护的既有改动 | `.../P1-RACE-TEST-01/20260710T183813+0800/baseline/coverage-gap.json`（FAIL 覆盖基线）；`.../first/summary.json`（六项首次 PASS）；`.../final/blackbox/summary.json`（路由、业务、竞态全套 PASS，SHA-256 `9639c62314ed7be3741815d7082ed170879a507e3ef2fb994e9ae45ad64069bb`）；旧失败基线 `.../BB-RACE/20260710T152255+0800-docx-baseline/` | `ready_for_gate` |

| P1-PACK-TEST-01 | Codex `/root` | 本会话复测；不是候选发布独立验收 | `2026-07-10T18:42:42+0800` | `2026-07-10T18:49:03+0800` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 依赖 P0-PACK-02、P1-CI-01 均为 `ready_for_gate`。基线 consumer 只检查 tgz 清单。新增固定最小消费模板与自动执行器：当前构建重新打五包，记录 tgz 和两 Worker/两 WASM 源 SHA；系统临时目录只用五个 `file:*.tgz` 安装，五包均为普通目录且真实路径在临时项目内；公共目录只有三份测试文档，无 Worker/WASM；无源码别名；根 ESM、子路径、类型和三份公开 CSS 经 `tsc`、`vue-tsc`、Vite 正式构建验证；按权威手册配置 ES module Worker，并通过公开 `./worker?worker&url` 与 `./assets/*.wasm?url` 入口注入。正式 preview 中 DOCX/XLSX/PDF 全部 ready，DOCX/XLSX 都实际使用 Worker；两 Worker 和两 WASM 均 200，WASM MIME 正确且浏览器所取文件 SHA 与 tgz 相同，Worker 由 tgz 公开入口经 Vite 重打包；无控制台、页面、请求或响应异常。临时目录每次均清理。早期三次失败分别是 macOS `/var` 真实路径误判、模板缺 `.vue` 声明、遗漏手册要求的 Worker 格式；第四次是测试错误地要求重打包 Worker 字节不变，均保留且不计产品失败。删除临时安装包的 Vue DOCX JS 入口后 Vite 稳定退出 1，故障探针成立。最终根 consumer 套件首次 PASS；`git diff --check=0`。 | macOS；Node `v22.13.0`；npm `11.0.0`；pnpm `9.0.6`；Python Playwright `1.50.0`；Chromium `133.0.6943.16`；工作区有受保护的既有改动 | `.../P1-PACK-TEST-01/20260710T184242+0800/baseline/coverage-gap.json`（FAIL 覆盖基线）；`.../{first,second-consumer,third-consumer,fourth-consumer}/`（测试基础设施失败，保留）；`.../fifth-consumer/`（PASS）；`.../fault-probe/summary.json`（入口缺失被拦截）；`.../final/consumer/summary.json`（PASS，SHA-256 `b8ce92f8e5b968a27a6b563261cd2fc5b820837936abb327550bea5301bc46bb`） | `ready_for_gate` |

| P1-DOC-TEST-01 | Codex `/root` | 本会话复测；不是候选发布独立验收 | `2026-07-10T18:49:48+0800` | `2026-07-10T18:54:25+0800` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 依赖 P1-CI-01、P1-E2E-01 均为 `ready_for_gate`。基线核心文档仍称 P1 主体未完成并指向临时 fault server、旧材料清单和旧消费夹具。测试方案现登记 7 个根入口及各套件真实内容；材料权威路径改为 `test-data/manifest.json`；消费模板改为 `tests/consumer/template/`；执行手册登记仓库内延迟服务器、竞态脚本、自动 tgz 流程、证据位置和逐层排错方法，并纠正旧类型/warning 限制说明。新增 `pnpm test:docs`，55 项检查核对根脚本及别名、CI、实际测试文件、套件登记、固定 Playwright、空允许清单、20 项非敏感材料、无消费源码别名及已知过期文字；根 `check` 已接入。显式加入不存在脚本后检查退出 1，恢复后两次 PASS。额外补齐 XLSX 工作表公开 `role=tab/aria-selected` 以符合已登记稳定入口。`pnpm typecheck=0`、材料检查=0、`git diff --check=0`。未修改受保护的 `docs/INDEX.md` 用户改动。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；工作区有受保护的既有改动 | `.../P1-DOC-TEST-01/20260710T184948+0800/baseline/document-drift.json`（FAIL 漂移基线）；`.../first/docs/summary.json`（PASS）；`.../fault-probe/summary.json`（不存在脚本被拦截）；`.../final/docs/summary.json`（PASS，SHA-256 `855d156e97cf602a100a17f741eb81fb4fc800381f4cefa3bf4d11ea34628349`） | `ready_for_gate` |

| P2-RUNTIME-01 | Codex `/root` | 本会话复测；不是候选发布独立验收 | `2026-07-10T19:09:42+0800` | `2026-07-10T19:31:11+0800` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 依赖 P1-UNIT-01、P1-VUE-01 均为 `ready_for_gate`，P1 全量检查已连续三轮通过。真实失败基线为私有包不存在，新增共享边界测试退出 1；第一次基线仅因测试路径写法错误退出 1，单独保留并不冒充产品失败。新增 `private: true` 的 `@arcships/office-runtime`，只提供来源、显式 URL 规则、输入限制契约、不可变加载上下文、实例拥有的任务序列/最新任务协调器、统一错误和脱敏诊断；无 Vue、界面 DOM 或模块级任务状态。DOCX Runtime 新增 file/bytes/url 来源并让 Viewer、模型组合函数和 Editor 导入使用 loader；XLSX/PDF 保留公开错误与属性，由薄适配调用共享规则，控制器诊断新增完整 taskId。`pnpm check=0`：类型和正式构建通过；22 项核心行为、54 项 DOCX 集成、XLSX 结构、7 项组件和运行边界检查通过；正式 preview 路由/业务流程 PASS，`BB-SEC-URL` 四项 PASS，`BB-RACE` 原六项加 PDF 慢 A→快 B 共七项 PASS，均首次通过；五个真实 tgz 工作区外消费 PASS，公开 JS/声明/包清单无私有包引用；55 项文档检查、失败传播和 `git diff --check` 通过。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；Python Playwright `1.50.0`；Chromium `133.0.6943.16`；工作区有受保护的既有改动 | `.../baseline2/unit/summary.json`（共享包缺失 FAIL 基线）；`.../component1/component/summary.json`（旧测试替身不符合已有 Runtime 接口，测试基础设施失败）；`.../final/check/summary.json`（PASS，SHA-256 `033db0827e8a3f263a455956f9654777fafc99612fe539aa4d4605a75a1b3456`）；`.../final/blackbox/summary.json`（PASS，SHA-256 `e357e389201ec15856b1c61c775d1dc616311d3187af6d577521ec3ddcf0c04f`）；`.../final/consumer/summary.json`（PASS，SHA-256 `9b82ca7ca8293cc08e35fc347c83c10e17d2627c69c768b350e427f725e2cdd8`） | `ready_for_gate` |

| P2-CONFIG-01 | Codex `/root` | 本会话复测；不是候选发布独立验收 | `2026-07-10T19:32:26+0800` | `2026-07-10T19:44:29+0800` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 依赖 P2-RUNTIME-01=`ready_for_gate`。失败基线证明 DOCX Runtime 会读取创建后被调用方修改的限制对象，且 XLSX 没有实例 Runtime。DOCX Runtime 现快照限制、地址规则、回调和底层能力；新增 `createXlsxRuntime`，由实例拥有不可变 Worker/WASM、解析选项、任务序列、Worker client 清单和销毁过程。显式实例不读取旧全局默认值；旧 `setWasmSource` 仅保留为兼容默认实例入口，并在初始化后拒绝修改。正式 `/runtime-isolation` 消费页同时运行两个 DOCX 实例，公开验证不同 runtimeId、taskId、WASM 地址、限制快照以及 Worker 创建/终止。`pnpm check=0`：类型和正式构建通过；24 项核心行为、54 项 DOCX 集成、XLSX 结构、7 项组件测试通过；正式 preview 的路由、业务、安全、双实例隔离和 7 项竞态均首次 PASS；五个真实 tgz 工作区外消费 PASS，公开 JS/类型/样式/Worker/WASM 完整且无私有包泄漏；55 项文档检查、失败传播和 `git diff --check` 通过。未修改受保护的 `packages/docx-core/tsup.config.ts`。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；Python Playwright `1.50.0`；Chromium `133.0.6943.16`；工作区有受保护的既有改动 | `.../baseline/unit/summary.json`（DOCX 配置快照与 XLSX Runtime 缺失 FAIL）；`.../final/check/summary.json`（PASS，SHA-256 `366ac509a4e85c4914b340e20ad2bd40fc344d42e1148092c4a9b60222925cb2`）；`.../final/blackbox/formal-config-isolation/summary.json`（PASS，SHA-256 `89b131782ae920df5420f30d004f364d69d0f116081dc50928271d44df4abe61`）；`.../final/blackbox/summary.json`（PASS，SHA-256 `0724c085c25613f368ebf21e67c201fbc9fe15584687edd2311d91ccde21f7c4`）；`.../final/consumer/summary.json`（PASS，SHA-256 `ddcfa8bc8f105d6852809e4c04fd7776e836ee9b6fffcd03a4cc5f53cece01e6`） | `ready_for_gate` |

| P2-RACE-01 | Codex `/root` | 本会话复测；不是候选发布独立验收 | `2026-07-10T19:45:17+0800` | `2026-07-10T20:08:32+0800` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 依赖 P2-RUNTIME-01、P1-RACE-TEST-01 均为 `ready_for_gate`。失败基线实证 XLSX Worker 请求取消后仍为 pending、controller 从不调用 task.finish、延迟加载丢失 signal/taskId，且 DOCX 在 document 不可用和 click 异常时泄漏对象 URL。共享协调器清理改为幂等；XLSX Worker 四类请求支持 AbortSignal，故障后终止并拒绝复用；立即与延迟加载统一到同一任务，重复继续无效，Worker 高成本任务取消会终止实例，主线程过期结果同时释放 Workbook/图片 URL；Worker 和主线程替换 Workbook 均调用 free；异步历史恢复使用代次和来源代次，只在当前任务提交。DOCX 导出以 token、model 和 documentLoadNonce 判断最新任务，新建/导入/模型变化/卸载均使旧导出失效；下载成功、异常、取消均幂等清理 anchor、timer 和 URL，file 清空会取消导入。`pnpm check=0`：类型和正式构建通过；35 项核心行为、54 项 DOCX 集成、17 项组件测试通过；正式 preview 路由/业务/安全/配置全部 PASS，BB-RACE 原 7 项加 XLSX deferred 双继续/卸载和 DOCX 双导出/切路由共 11 项均首轮 PASS；五个真实 tgz 工作区外消费 PASS；55 项文档检查、失败传播和 `git diff --check` 通过。未修改受保护的 `packages/docx-core/tsup.config.ts`。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；Python Playwright `1.50.0`；Chromium `133.0.6943.16`；工作区有受保护的既有改动 | `.../baseline/summary.json`（3 项真实 FAIL）；`.../final/check/summary.json`（PASS，SHA-256 `d12813ed323c2a3128f2d71b0ccf6fff94744a8b215a4c43f07bf05370e0a900`）；`.../final/component/summary.json`（17 项 PASS，SHA-256 `d505a99591ff875d554c43ec48c0a14d104a736d82942d248c87e8914e514275`）；`.../final/blackbox/formal-race-regression/summary.json`（11 项首轮 PASS，SHA-256 `b48cb71927142c902a46eefd28e4c7a4a439b8ec999d70f2e16091f44d29b9d8`）；`.../final/blackbox/summary.json`（PASS，SHA-256 `e286c0caa1e8c13f6355a99b1a22b85fc0e9f1219b1de193b98e4ac8f0172c49`）；`.../final/consumer/summary.json`（PASS，SHA-256 `545c892f0e9ff929cffe67821dc25082c6698f0c9ad487bc2196b147e73105d0`） | `ready_for_gate` |

| P2-DOCX-RENDER-01 | Codex `/root` | 本会话复测；不是候选发布独立验收 | `2026-07-10T20:09:08+0800` | `2026-07-10T20:36:51+0800` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 依赖 P1-E2E-01、P2-RUNTIME-01 均为 `ready_for_gate`。失败基线实证多页节点误映射、表格正文固定空、页眉页脚固定关闭、Viewer/Editor 两套渲染树和只读交互泄漏。新增唯一 `DocxDocumentSurface`，两种模式都使用正式 layout snapshot 的 `source.nodeIndex`、同一页面/段落/表格/图片/页眉页脚实现；表格复用富文本段落渲染，section 首页/奇偶页正确选择，字段、批注与修订按页过滤，只读不挂输入、缩放、菜单和拖放。正式截图检查又发现 Viewer 外层布局把共享 surface 宽度压为 0，已修为全宽弹性容器，并加入宽高和模糊像素差门禁。全仓类型与正式构建通过；18 项组件回归 PASS；正式 preview 的四份对比材料均首轮 PASS，多页合同 126 节点、发票 34 单元格、图片 1 张、中英文 8 段在两种模式完全同构，截图差异率 0.11%～0.25%；路由、业务、安全、双实例和 11 项竞态全套 PASS；真实 tgz 消费随核心包验证通过；`git diff --check` 通过。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；Python Playwright `1.50.0`；Chromium `133.0.6943.16`；Pillow `11.1.0`；工作区有受保护的既有改动 | `.../baseline/summary.json`（真实 FAIL）；`.../final/component/summary.json`（18 项 PASS，SHA-256 `3e1935021958fda37cbea2bf7e0d60c88fbae330b81d1b90083e911f2717ccb7`）；`.../final4/blackbox/formal-docx-render-parity/summary.json`（4 项首轮 PASS，SHA-256 `8c05ca5763ba81c36d289c3b34bbe36ccdabb9b8f37d5f815b849150f501f3ed`）；`.../final4/blackbox/summary.json`（全套 PASS，SHA-256 `cdaee410e45e271d8077b0abbcc16c61925e328274416dc659d008ba51c139ce`） | `ready_for_gate` |

| P2-CORE-01 | Codex `/root` | 本会话复测；不是候选发布独立验收 | `2026-07-10T20:10:47+0800` | `2026-07-10T20:36:51+0800` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 依赖 P2-RUNTIME-01=`ready_for_gate`。失败基线证明两包没有纯 `./core`，根声明带入平台类型，且无发布声明边界检查。新增显式纯入口：DOCX 提供模型、布局和不可变段落/表格/样式命令，XLSX 提供纯类型、颜色、A1/公式和锚点数学；AST 门禁递归禁止 Vue/demo/平台标识、外部运行依赖、wildcard 和顶层可变缓存。新增真实 `./runtime`、`./wasm-url` 构建入口，替换曾转导整个根入口的代理；根入口保留 0.x 兼容。构建审计发现 XLSX 纯声明经类型桶泄漏 `@dukelib/sheets-wasm`/Worker，已拆出纯 theme 类型并加入生成声明图门禁。受保护 DOCX tsup 只做必要 entry 合并，原 Worker/WASM 配置未覆盖。全仓类型与正式构建通过；核心单元共 38 项 PASS，纯入口在禁用浏览器全局的子进程可执行且 deep-freeze 输入不变；五个真实 tgz 工作区外消费 PASS，`./core` JS/类型可导入，深层路径严格返回 `ERR_PACKAGE_PATH_NOT_EXPORTED`，Worker/WASM 无 404；`git diff --check` 通过。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；工作区有受保护的既有改动 | `.../baseline/summary.json`（真实 FAIL）；`.../final3/unit/summary.json`（38 项 PASS，SHA-256 `5e6b3264664c756bdda7aced0f3711813e17e2dca9a939d7508051e245938135`）；`.../final2/consumer/real-tgz-manifests/summary.json`（5 包 PASS，SHA-256 `0327849d24f8083226bed7bfc517d851b8a49dbbcc65699e4984902e6c465f79`）；`.../final3/consumer/external-tgz-consumer/summary.json`（PASS，SHA-256 `7a4f5f16ac61c09e8bea08b455dfd713eef04ccc42ecc429834541917e3e2b2e`） | `ready_for_gate` |

| P2-DOCX-EDIT-01 | Codex `/root` | 本会话复测；不是候选发布独立验收 | `2026-07-10T20:38:05+0800` | `2026-07-10T20:50:48+0800` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 依赖 P2-DOCX-RENDER-01=`ready_for_gate`。失败基线实证只读仍注册全局快捷键、图片/表格拖缩不提交模型、多个公开图片命令为空且正式编辑黑盒缺失。编辑监听器现随实例的 `editable` 动态挂卸，只读挂载和切换均为零；部分文本输入与格式、图片环绕/移动/缩放/跨段落移动、表格行列和列宽全部通过新模型事务进入统一历史，非法目标公开显示 `Unsupported:`；撤销重做同步恢复选区和文本范围。新增正式测试消费控件，只使用公开控制器和公开模型快照。全仓类型与正式构建通过；21 项组件回归 PASS，其中新增 3 项覆盖只读监听、模型不变性和图片/表格历史；正式 preview 的组合输入、局部格式与选区恢复、只读隔离、4×4 表格和图片事务 4 项均首轮 PASS；路由、业务、渲染一致性、安全、双实例配置及 11 项竞态全套 PASS；控制台、网络、页面异常和下载均无违规；`git diff --check` 通过。未改动 P2-RACE 的导入导出取消实现。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；Python Playwright `1.50.0`；Chromium `133.0.6943.16`；工作区有受保护的既有改动 | `.../baseline/summary.json`（真实 FAIL）；`.../final/component/summary.json`（21 项 PASS，SHA-256 `a04c169e74f9f884882cfe813f9f57e10eba49a109af6a314ab97307d2d1511c`）；`.../blackbox/formal-docx-editing/summary.json`（4 项首轮 PASS，SHA-256 `c2cb1005263f3482eacc702799528aff550dd44d55de6925d80edd9746d98c8c`）；`.../final/blackbox/summary.json`（全套 PASS，SHA-256 `bc681cff0f3d97f3102b9e7b4e8576838dd96e4c87bc19e57b7dfc479e7dc814`） | `ready_for_gate` |

| P2-API-01 | Codex `/root` | 本会话复测；不是候选发布独立验收 | `2026-07-10T20:51:40+0800` | `2026-07-10T21:04:51+0800` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 依赖 P2-CORE-01、P2-DOCX-EDIT-01 均为 `ready_for_gate`。失败基线实证五包正式声明没有计划中的弃用标记、DOCX runtime 子入口类型不自足、XLSX Runtime 泄漏私有错误、FileUpload 没有结构化拒绝事件、真实包深层导入矩阵不足。新增人读接口合同和机器清单，精确登记五包 21 个公开路径、稳定入口、事件、错误码、0.2.0 弃用与整个 0.x 保留规则，以及两个内部桥接的 2026-08-09 清理日期；`initWasm` 保持稳定不擅自弃用。DOCX runtime 重导出签名所需错误/诊断类型；XLSX 公开 `XlsxRuntimeError(RUNTIME_DISPOSED)`，不再抛私有共享包错误；FileUpload 新增兼容的 `files-rejected` 结构化事件。生成声明逐符号保留旧全局配置、DOCX 旧组件/渲染/registry/缩略图、XLSX 旧组件/渲染值和类型的 `@deprecated`。门禁冻结五包根声明完整导出指纹 1048/149/100/25/31，检查精确 exports、声明图、私有包/绝对路径泄漏和文档术语；缺失出口故障注入按预期 FAIL。全仓类型与正式构建通过；39 项单元行为、22 项组件和全套正式 preview PASS；五个真实 tgz 记录逐包 SHA-256，在工作区外无源码别名安装，21 个正向入口解析/读取、11 个 JS 入口执行、38 条深层路径严格返回 `ERR_PACKAGE_PATH_NOT_EXPORTED`，TypeScript、Vue 类型、Vite 正式构建和浏览器 Worker/WASM 全部 PASS，无公共目录资源复制；`git diff --check` 通过。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；Python Playwright `1.50.0`；Chromium `133.0.6943.16`；工作区有受保护的既有改动 | `.../baseline/summary.json`（真实 FAIL）；`.../public-api-final/summary.json`（PASS，SHA-256 `851f8f3fa5762f1c8fad26b48165c2d4e6b750687b4b20dddcd0e0fbddc0bd05`）；`.../public-api-probe-missing-export/summary.json`（预期 FAIL，SHA-256 `adbfee6cd25462d6fd160f41dac98c1f1d2f586bfbb1bb7ca0a6f19ec5b3b516`）；`.../pack-final/real-tgz-manifests/summary.json`（5 包 PASS，SHA-256 `9e7f63162cfa8bd9eb7367f68261b8aedadf10f992fdff2c31e84a2ab3a93138`）；`.../pack-final/external-tgz-consumer/summary.json`（PASS，SHA-256 `08d16272a53bcacd4996020e59675b7cc6799aa13001886f7f778ba80106ae99`）；`.../final/unit/summary.json`（PASS，SHA-256 `913db5fabd6e16344c373d6234f8c72e25afc68e023fe7b8e5dea17f44bbea09`）；`.../final/component/summary.json`（22 项 PASS，SHA-256 `ec55c0a56ea9fc26766c6683e65cb1d5663544d70e0052ba8a390bd7b2fc2cd2`）；`.../final/blackbox/summary.json`（全套 PASS，SHA-256 `e251e7289890733a851f9bf60c5607ebbed544b3fda453d7e1c5a639db5a927e`） | `ready_for_gate` |

| P2-DEAD-01 | Codex `/root` | 本会话复测；不是候选发布独立验收 | `2026-07-10T20:38:05+0800` | `2026-07-10T20:55:13+0800` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 依赖 P2-DOCX-RENDER-01、P2-CORE-01 均为 `ready_for_gate`。失败基线实证 Viewer 留有第二套死渲染、缩略图空成功、XLSX 空 overlay/双击处理、重复 A1 算法和错误深层导入说明。现删除 Viewer 重复分页/段落/表格/页眉页脚路径，只保留共享 surface；页面向实例 registry 发布真实 DOM/尺寸/内容签名，缩略图从真实 surface 渲染并缓存，旧无文档输入入口明确返回 `unavailable` 且绘制可见原因；XLSX 选区层提供 `aria-live` 状态，空双击拦截删除，根 A1 兼容函数委托纯 core；删除两个从未公开且无消费者的组件桶。新增精确 18 项门禁并纳入正式单元套件，不机械禁止合法空值；编辑任务结束后复核图片/表格/工具栏无空公开操作。全仓类型与正式构建通过；38 项单元行为、18 项死路径门禁、54 项 DOCX 集成、21 项组件回归全部 PASS；正式 preview 的路由、业务、渲染一致性、编辑、安全、双实例和竞态全套 PASS；`git diff --check` 通过。已公开旧组件未删除，交由 P2-API-01 标记弃用并保留整个 0.x。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；Python Playwright `1.50.0`；Chromium `133.0.6943.16`；工作区有受保护的既有改动 | `.../baseline/summary.json`（真实 FAIL）；`.../final/unit/summary.json`（PASS，SHA-256 `c2e10bcf3d4edcaf888fde7b4b3f323027d7f04bdd71f91423fad876ec815672`）；`.../final/component/summary.json`（21 项 PASS，SHA-256 `e57c130f9b0d8fe3ab7ea4297c8b0b8b17691f68102c16b74e6e171b9a574889`）；`.../final/blackbox/summary.json`（全套 PASS，SHA-256 `2642ddfb72c4b0372f479ade92542abd515c1d5aa66e95eb509f72edc7d45d91`） | `ready_for_gate` |

| P3-PDF-BUDGET-01 | Codex `/root` | 本会话 PDF 黑盒子任务；不是候选发布独立验收 | `2026-07-10T23:10:00+0800` | `2026-07-10T23:57:45+0800` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | 按用户决策只实现整份文件体积上限：默认 `50 MiB`，宿主可通过公开 `maxFileSize` 调整。Blob 在读取前检查，URL 优先检查 `Content-Length`，无长度头时在取得字节后、创建引擎前检查。超限返回含 `actual`、`allowed` 的 `PDF_TOO_LARGE`，不打开引擎、不创建页面或缩略图。PDF-008 使用 4651 字节上限拒绝 4652 字节样本，无新增 WASM、Worker 或渲染，恢复默认值后正常显示 4 页。页数、像素、内存、并发页和复杂度边界已从阻断条件移除。最终 `pnpm check=0`。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；Python Playwright `1.50.0`；Chromium `133.0.6943.16`；工作区有受保护的既有改动 | `.../P1-CI-01/20260710T155534Z/check/summary.json`（PASS，SHA-256 `a8e43093b07c8bc3442d1009eaedfebae7aeeaa6c02601004868c701176018cb`）；`.../blackbox/formal-pdf-workflows/summary.json`（PDF-001～006、PDF-008 首轮 PASS，SHA-256 `131a5b2776a86333f1d05327baf5cfadf88242b09c117856a25e4533ee63b9eb`） | `ready_for_gate` |
| P3-PDF-01 | Codex `/root` | 本会话 PDF 黑盒子任务；不是候选发布独立验收 | `2026-07-10T23:10:00+0800` | `2026-07-10T23:57:45+0800` | `41913ca03666aa99302d13924e038db17e478480` | 工作区未提交；无结果提交 | PDFium Worker 真实查看已完成：翻页、首页/末页、缩放、旋转、缩略图、全文搜索、损坏恢复和下载均操作真实文档；首张图片解码后才 `ready`，切换和离页会终止 Worker、撤销对象 URL。单一 `maxFileSize` 体积上限已经接入，复杂预算不再阻断。最终 `pnpm check=0`：类型、正式构建、43 项单元、25 项组件、八组正式浏览器、五包真实 tgz、55 项文档、失败传播和差异检查全部通过。 | macOS；Node `v22.13.0`；pnpm `9.0.6`；Python Playwright `1.50.0`；Chromium `133.0.6943.16`；工作区有受保护的既有改动 | `.../P1-CI-01/20260710T155534Z/check/summary.json`（PASS，SHA-256 `a8e43093b07c8bc3442d1009eaedfebae7aeeaa6c02601004868c701176018cb`）；`.../blackbox/formal-pdf-workflows/summary.json`（首轮 PASS）；`.../consumer/summary.json`（真实 tgz PASS，SHA-256 `bce175c5fbd289c310da5beffaf7e86af1569a402da5e7982615a4347f655d8d`） | `ready_for_gate` |

| P3-BUDGET-01 | Codex `/root/pdf_acceptance_audit` | Codex `/root` 本会话集成复测；不是候选发布独立验收 | `2026-07-11T00:07:27+0800` | `2026-07-11T00:39:11+0800` | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 工作区未提交；无结果提交 | 新增 DOCX/XLSX 实例级输入、归档、XML、模型和解析时间限制；中央目录与实际流式解压同时计数；路径穿越、重复规范化名称、XML 实体/嵌套/属性、MIME、行列、工作表、共享字符串和公式限制均返回结构化错误；修复命名空间前缀绕过和 XLSX URL 完整下载后才拒绝的缺口；DOCX Viewer/Editor 在布局页数超限时隐藏文档并可通过新 Runtime 恢复。全仓类型和正式构建 PASS；DOCX 集成 54/54、XLSX 结构、组件 28/28 PASS；合并审计修复后 `BB-STRESS` 18/18 首轮 PASS；`git diff --check` PASS。 | macOS arm64；Node `v22.13.0`；pnpm `9.0.6`；Python Playwright `1.50.0`；Chromium `133.0.6943.16`；视口 `1440×900`；`zh-CN`/`Asia/Shanghai` | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/P3-BUDGET-01/20260711T004100+0800/stress/formal-stress-workflows/summary.json`（PASS，SHA-256 `b9792c0c0d9c18c2a3e5ce6f38f0c1cc4deab67be313d2df8e21cefe15094086`）；`.../20260711T-runtime-limits-final3-component/component/summary.json`（PASS，SHA-256 `0e0fd3953fb78cf510d6542a3bf123c8a46eb90a6d832d5f3dcf761226e8d123`） | `ready_for_gate` |
| P3-PERF-BASELINE-01 | Codex `/root/pdf_blackbox_tests`、Codex `/root` | 本会话交叉复测；不是候选发布独立验收 | `2026-07-11T00:07:27+0800` | `2026-07-11T01:45:20+0800` | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 工作区未提交；无结果提交 | 修复 Workbook 子对象释放顺序、真实滚动和按需加载后，固定 Chromium/视口/材料连续三轮无异常；另一执行代理只读复核候选后批准 100 项预算。首次批准门禁如实失败于撤销耗时和监听器峰值的自然波动，依据该轮三次原始值更新这两项基线后重新执行三轮，最终 `PASS`，预算违规、阻断、控制台、页面和网络异常均为 0。大表纵向/横向范围 `10653/2618`，图表每轮首次新增 2 个 JS 请求，回 Home 后 Worker/Object URL 均为 0。 | macOS arm64；Node `v22.13.0`；pnpm `9.0.6`；Python Playwright `1.50.0`；Chromium `133.0.6943.16`；视口 `1440×900`；设备倍率 1；`zh-CN`/`Asia/Shanghai` | 最终 `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/P3-PERF-BASELINE-01/20260711T-formal-approved-gate-rerun/summary.json`（PASS，SHA-256 `9fa3aaa35d4505f18354026437f5f6bc8fcbbe502f0b662b0b98d5042b89bec6`）；批准配置 SHA-256 `737f5000ae80d275923acc84bb68d79f3b2564ada0f136babb2e30ecf8b702b5`；首次批准失败证据保留在相邻 `20260711T-formal-approved-gate/` | `ready_for_gate` |
| P3-XLSX-SCROLL-01 | Codex `/root/xlsx_scroll_design` | Codex `/root` 本会话集成复测；不是候选发布独立验收 | `2026-07-11T01:12:30+0800` | `2026-07-11T01:45:20+0800` | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 工作区未提交；无结果提交 | 增加真实滚动内容层、行列累计偏移、二分定位、按动画帧合并绘制和只在尺寸变化时重设 Canvas；深滚动点击、编辑、Tab、撤销和隐藏行列共用同一坐标模型。定向组件 4/4 和旧红门禁 2/2 PASS；正式三轮纵向 `10653`、横向 `2618`，约 601 帧，95 分位约 `16.7ms`，滚动长任务和 Canvas 尺寸写入均为 0。 | 同 P3 性能固定环境 | `.../P3-PERF-BASELINE-01/20260711T-formal-approved-gate-rerun/summary.json`（PASS，SHA-256 `9fa3aaa35d4505f18354026437f5f6bc8fcbbe502f0b662b0b98d5042b89bec6`） | `ready_for_gate` |
| P3-XLSX-LOAD-01 | Codex `/root/xlsx_lazy_load_impl` | Codex `/root` 本会话集成复测；不是候选发布独立验收 | `2026-07-11T01:12:30+0800` | `2026-07-11T01:45:20+0800` | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 工作区未提交；无结果提交 | demo 路由和 XLSX 图表、地图、WebGL 改为动态分包，公开根导出保持兼容并新增明确子入口；正式构建删除旧公共目录 WASM 副本。图表 XML 缺少 WASM 图表对象时建立兼容模型，Worker 内置 XML 解析能力且不静默伪装成功。正式预览图表数量 1、SVG 1，控制台/页面/请求失败/Worker 错误均为 0；普通页面不加载图表，图表样本每轮按需新增 2 个 JS 请求，重复 WASM 为 0。 | 同 P3 性能固定环境 | Worker 图表诊断 `.../P3-XLSX-LOAD-01/20260711T-worker-xml-fix/diagnostic.json`（SHA-256 `73493a394a7c11978d385f01870e62bb43d688ad719c9228c5f93d0135187393`）；最终三轮 `.../P3-PERF-BASELINE-01/20260711T-formal-approved-gate-rerun/summary.json` | `ready_for_gate` |
| P3-IMAGE-01 | Codex `/root/pdf_acceptance_audit` | Codex `/root` 本会话集成复测；不是候选发布独立验收 | `2026-07-11T00:40:10+0800` | `2026-07-11T00:52:10+0800` | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 工作区未提交；无结果提交 | 新增实例级图片字节、宽高、单图/总像素与并发解码预算；解析 PNG/JPEG/GIF/WebP/BMP/TIFF/ICO/SVG 文件头并在对象 URL、Canvas 和像素缓冲前拒绝；DOCX/XLSX 归档入口执行同一预检，失败和销毁释放临时资源；浏览器解码失败通过公开错误/诊断可观测。定向测试 10/10 PASS；DOCX 集成 54/54、XLSX 结构、全仓类型/正式构建、`git diff --check` PASS；统一构建后正式 preview 的 DOCX/XLSX 宽度、单图像素、畸形图片、6000×6000 正常高分辨率与总像素 8/8 首轮 PASS，每个错误后均恢复 ready。 | macOS arm64；Node `v22.13.0`；pnpm `9.0.6`；Python Playwright `1.50.0`；Chromium `133.0.6943.16`；视口 `1440×900`；`zh-CN`/`Asia/Shanghai` | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/P3-IMAGE-01/20260711T004527+0800/blackbox-image-final/summary.json`（PASS，SHA-256 `f5d0e7b916e01ef2116bcd1326ddc229ca0dbc4cabcde5141bcf7654df9b9912`）；定向测试日志 SHA-256 `3d4e7cd25771aedfb2221a971e7e6922ccaac602afdbf5d717adb194e90d4563` | `ready_for_gate` |
| P3-CACHE-01 | Codex `/root/pdf_component_tests`、Codex `/root/cache_cleanup_impl` | Codex `/root` 本会话集成复测；不是候选发布独立验收 | `2026-07-11T00:40:10+0800` | `2026-07-11T01:46:05+0800` | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 工作区未提交；无结果提交 | DOCX/XLSX 历史、缩略图和分页按实例、条目及字节回收；模块强缓存改为 WeakMap 或移除；对象 URL 异常和销毁路径统一撤销。XLSX Workbook 实例登记全部 Worksheet/CellValue，按子对象到父对象顺序幂等释放，连续材料切换不再出现 WASM 越界。缓存基线 5/5、组件 36/36、真实 WASM 50 轮、正式性能三轮和最终 BB-STRESS 均 PASS；回 Home 后 Worker/Object URL 为 0，内存回落满足批准预算。 | 同 P3 正式环境 | 组件 `.../P3-CACHE-01/20260711T-history-first-batch-final/component/`；性能 `.../P3-PERF-BASELINE-01/20260711T-formal-approved-gate-rerun/summary.json`；压力 `.../BB-STRESS/20260711T-p3-final-stress/stress/summary.json`（PASS，SHA-256 `ea840999eb57cc04ac54a9e2d4a5e6104edba2b4f5ea9802710dacf0dbc44ed8`） | `ready_for_gate` |
| P3-WORKER-01 | Codex `/root` | 本会话开发后自测；不是候选发布独立验收 | `2026-07-11T00:42:02+0800` | `2026-07-11T00:50:49+0800` | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 工作区未提交；无结果提交 | XLSX 取消现在终止客户端所属 Worker 并拒绝全部在途请求；超时、初始化故障和协议故障关闭不可信 Worker；只读加载不再因 DOM/图表错误静默转主线程；图表 Worker 失败保留快速结果但通过公开 `worker-error` 诊断明确报错，不重跑高成本任务；新增正式公开消费页和真实 Worker 故障材料。全仓类型/正式构建 PASS；Worker 基线 6/6、XLSX 相关单元 19/19（含子项 21 项）PASS；正式 preview 超时恢复和取消均首轮 PASS，主线程交互计数正常，Worker 创建/终止数一致；`git diff --check` PASS。 | macOS arm64；Node `v22.13.0`；pnpm `9.0.6`；Python Playwright `1.50.0`；Chromium `133.0.6943.16`；视口 `1440×900`；`zh-CN`/`Asia/Shanghai` | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/P3-WORKER-01/20260711T005000+0800/formal-worker-workflows/summary.json`（PASS，SHA-256 `4a055c115aa0e6b8e6465121942a2f851da455183c45428cdf4c1a16e98f9198`） | `ready_for_gate` |
| P3-STRESS-01 | Codex `/root` | 本会话开发后自测；不是候选发布独立验收 | `2026-07-11T01:45:37+0800` | `2026-07-11T01:46:05+0800` | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 工作区未提交；无结果提交 | 前置 P3-IMAGE/PDF/CACHE/XLSX-SCROLL/WORKER 均为 `ready_for_gate` 后启动。`CI_PREBUILT=1 CI_RUN_ID=20260711T-p3-final-stress pnpm test:stress` 退出 0；正式资源压力与 Worker 工作流均首轮 PASS，无重试。覆盖输入/归档/XML/路径/MIME/图片边界、合法边界恢复、Worker 超时、取消、重建和资源归零；每个拒绝用例后页面恢复 `ready`。 | macOS arm64；Node `v22.13.0`；pnpm `9.0.6`；Python Playwright `1.50.0`；正式 preview | `.../BB-STRESS/20260711T-p3-final-stress/stress/summary.json`（PASS，SHA-256 `ea840999eb57cc04ac54a9e2d4a5e6104edba2b4f5ea9802710dacf0dbc44ed8`）；资源压力 SHA-256 `5a22f8d4dc58ac94deffa54806db496ef5d53bb9acdf967b6fd38d62ce1ded33`；Worker SHA-256 `016465b2138ce1b40484fa23786c1447ffe5c4c6bc56850e6cb0136e487fae30` | `ready_for_gate` |
| P4-API-01 | Codex `/root`、三个分工执行代理 | 本会话交叉复测；不是候选发布独立验收 | `2026-07-11T01:49:13+0800` | `2026-07-11T02:26:01+0800` | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 工作区未提交；无结果提交 | 首次 P3 门禁如实发现 DOCX/XLSX 声明引用私有运行包、根导出指纹变化和 `vue-xlsx` 三个按需子入口未登记。随后把公开结构类型移回各自核心包，五个公开包统一 `0.2.0`，登记根入口与 `./chart`、`./map`、`./webgl`，公开合同、版本/弃用规则和运行边界同步收口；没有放宽检查器。图表锚点回归第一次正式工作流失败及重试失败均保留，按工作表真实行列尺寸修复后独立回归三项首轮 PASS。最终完整 `pnpm check=0`，公开声明不再要求安装私有 `office-runtime`，五包真实 tgz 外部消费通过；随后批准性能基线三轮通过。 | macOS arm64；Node `v22.13.0`；pnpm `9.0.6`；Python Playwright `1.50.0`；Chromium `133.0.6943.16` | 失败基线 `.../P1-CI-01/20260711T-p3-stage-gate/check/summary.json`；锚点正式回归 `.../P3-XLSX-LOAD-01/20260711T-anchor-regression-final/summary.json`（PASS，SHA-256 `331f5c6a36af749931cd65755071120f1adc491bd34d624937630449905bdb0a`）；最终门禁 `.../P1-CI-01/20260711T-p3-stage-gate-final/check/summary.json`（PASS，SHA-256 `56a5cf2e199c2991517fd016e9dc88dc7fcba8c2c270736087150fe461a35638`）；最终性能 `.../P3-PERF-BASELINE-01/20260711T-p3-performance-final/performance/summary.json`（PASS，SHA-256 `a1ac967d5772f20ccc3634bee23248a2abf1fd4d8e9f903346c88d6a8f015b42`） | `ready_for_gate` |
| P4-PACK-01 | Codex `/root`、Codex `/root/cache_cleanup_impl` | 本会话集成复测；不是候选发布独立验收 | `2026-07-11T02:13:00+0800` | `2026-07-11T03:15:15+0800` | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 工作区未提交；无结果提交 | 六红基线后实现双隔离构建和经上游核实的五包 MIT。首次双构建两轮都发现 demo 类型检查依赖旧 dist，已保留失败并给 `@arcships/xlsx-core/runtime` 增加明确工作区类型路径；修复后两轮冻结安装、类型、正式构建和五包 pack 全部 PASS，原始 tgz 与有效内容均逐包完全一致。首个加强版外部消费又发现 Worker 只读 XLSX 网格空白，修复可见行按 64 行从实例 Worker 批量加载并让选中值响应更新；随后测试脚本自身的 PDF 变量遮蔽失败也保留并修正。最终同批 tgz 在 Playwright 1.61 Chromium 首轮通过：DOCX 正文、XLSX `A1=Region`、PDF 四页/翻页/缩放/旋转/缩略图/搜索/下载、两 Worker/三 WASM/哈希/MIME/控制台全部 PASS。P4 就绪 14 项全绿。 | macOS arm64；Node `v22.13.0`；pnpm `9.0.6`；Python `3.12.2`；Playwright `1.61.0`；Chromium `149.0.7827.55` | 首次双构建 `.../20260711T-p4-pack-final/two-clean-builds/summary.json`（FAIL）；最终双构建 `.../20260711T-p4-pack-worker-grid-final/two-clean-builds/summary.json`（PASS，SHA-256 `88bdc006a270f99fa858ef5d3c7423db2b0ff37b9dbb37c96638b22ca81a18cd`）；最终外部消费 `.../external-consumer-test-fix/summary.json`（PASS，SHA-256 `50e0eda31409dfb4bf9f553a5564ae715c98997828d6d1e04912d320edd7ec6b`）；浏览器结果 SHA-256 `6bc752a44d2db3a7f4f4e6d57d85757c3449ca7fecf0b723e3ceca7a83525e24`；就绪检查 `.../20260711T-p4-pack-rerun/readiness/summary.json`（PASS，SHA-256 `62e5998eb7aef11cf0f1f88ad3577bfddaa4204f0aeadff1c668026ba95299a0`） | `ready_for_gate` |
| P4-DOC-01 | Codex `/root`、Codex `/root/xlsx_lazy_load_impl` | 本会话集成复测；不是候选发布独立验收 | `2026-07-11T02:14:00+0800` | `2026-07-11T03:15:15+0800` | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 工作区未提交；无结果提交 | README、索引、迁移、候选说明、历史标记、同批候选制品步骤和失败恢复均已收口；PDF 明确保持真实查看与全套操作，只设宿主可调的默认 50 MiB 整份体积上限。根文档检查实际核对脚本、文件、链接和禁止旧结论，共 127 项通过。兼容矩阵的实际版本和证据在 P4-MATRIX-01 结束时回填，不改变本文任务的结构与命令验收结果。 | 同上 | `.../P1-CI-01/20260711T-p4-doc-final/docs/summary.json`（PASS） | `ready_for_gate` |
| P4-MATRIX-01 | Codex `/root`、Codex `/root/matrix_review` | 本会话集成复测；不是候选发布独立验收 | `2026-07-11T03:15:15+0800` | `2026-07-11T03:20:40+0800` | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 工作区未提交；无结果提交 | 复用 P4 双构建同一批五个 tgz，在工作区外分别安装 Vue `3.2.25` 和 `3.5.39`。两套环境的 npm 安装、公开入口、深层导入拒绝、TypeScript、Vue 类型和 Vite 正式构建全部通过；Chromium `149.0.7827.55`、Firefox `151.0`、WebKit `26.5` 共六个组合均首轮 PASS。每个浏览器都验证 DOCX 正文、XLSX Worker 网格与 `A1=Region`、PDF 四页和翻页/缩放/旋转/缩略图/搜索/下载、默认 50 MiB 单一体积上限、两 Worker/三 WASM 的来源、哈希和 MIME；无重试、FLAKY、BLOCKED 或控制台异常。 | macOS arm64；Node `v22.13.0`；pnpm `9.0.6`；Python `3.12.2`；Playwright `1.61.0`；`zh-CN`/`Asia/Shanghai`；`1440×900` | `.../P4-MATRIX-01/20260711T-p4-matrix-final/matrix/summary.json`（PASS，SHA-256 `4af0776fd8c8b7ddc8c26c29fbea54aac6695b6ad2af5a335e7ff4c36f144407`）；当前 Vue 摘要 SHA-256 `52a9724bd853516a817441eb012bf13bd2f016fc3dd8875785f78f446110e8cc`；Vue 下限摘要 SHA-256 `842f3e31e02e10e923e61b057f6410dbd5e759151780058833c191594a23b6a3` | `ready_for_gate` |
| P4-RELEASE-01（前置准备） | Codex `/root`、Codex `/root/xlsx_scroll_design` | 本会话集成复测；不是候选发布独立验收 | `2026-07-11T02:18:00+0800` | 等待前置任务 | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 工作区未提交；无结果提交 | 基线显示 release 只是旧检查别名且没有受保护发布流程；已在不执行发布的前提下准备显式门禁、候选制品清单、同批 tgz 绑定、下载后校验、临时标签和失败恢复。P4-PACK、P4-MATRIX、P4-DOC 达到 `ready_for_gate` 前，P4-RELEASE-01 仍保持 `pending`，不运行完整 `BB-RELEASE`。 | GitHub Actions 目标：Ubuntu、Node `22.13.0`、pnpm `9.0.6`、Python `3.12.2`、Playwright `1.61.0` | `.../P4-READINESS-BASELINE/local/summary.json`（FAIL：P4-RELEASE-PIPELINE、P4-RELEASE-SUITE） | `pending` |
| P4-RELEASE-01 | Codex `/root` | 本会话开发后自测；不是候选发布独立验收 | `2026-07-11T03:20:40+0800` | 进行中 | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 工作区未提交；无结果提交 | P4-PACK-01、P4-MATRIX-01、P4-DOC-01 均已达到 `ready_for_gate`。开始执行显式 `BB-RELEASE`：从当前源码重新双构建五包，后续消费、六组合矩阵和候选清单必须复用这同一批 tgz；允许路径为发布脚本、工作流、候选文档和路线图，不执行真实 npm 发布。 | 同上 | `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/BB-RELEASE/20260711T-p4-release-final/` | `in_progress` |
| P4-RELEASE-01 | Codex `/root` | 本会话开发后自测；不是候选发布独立验收 | `2026-07-11T08:46:09+0800` | `2026-07-11T08:59:26+0800` | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 工作区未提交；无结果提交 | 先前界面整改前的候选已降为历史技术基线。第一次重跑在组件阶段发现第二段段首 Backspace 会被粗粒度段落选择清掉精确光标范围，门禁非零退出且后续步骤全部停止；修复为同段真实范围优先后，目标用例、51 项组件行为和 17 项 P3 回归通过。第二次 `CI_RUN_ID=20260711T-ux-p4-release-final-02 pnpm test:release` 的 15 个显式步骤全部 PASS：类型、全仓正式构建、单元、组件、正式黑盒、压力、三轮性能、双隔离打包、同批 tgz 材料化、工作区外消费、两 Vue×三浏览器矩阵、133 项文档、发布就绪、候选生成和差异检查。候选绑定提交、统一 `0.2.0`、五包 SHA-256/SHA-512；未执行真实 npm 发布。 | macOS arm64；Node `v22.13.0`；pnpm `9.0.6`；Python `3.12.2`；Playwright `1.61.0`；Chromium `149.0.7827.55`、Firefox `151.0`、WebKit `26.5` | 首次失败 `.../BB-RELEASE/20260711T-ux-p4-release-final-01/release/summary.json`（FAIL，SHA-256 `5a46f2f788023c7ca3fd38c82add5f71670e55af2cb5542c1f9013abb1c71b19`）；最终发布自测 `.../BB-RELEASE/20260711T-ux-p4-release-final-02/release/summary.json`（PASS，SHA-256 `e96107f39fe7f98db3be10b840e2ed62413045e2959ff05a4dab58f55823ff79`）；候选清单 `.../candidate/candidate-manifest.json`（SHA-256 `02ee4c3bec24a48eb4a08566252e1d42aef9797ce284d366e658eb053e33bbfb`）；黑盒 SHA-256 `6d1d34baf4681d9550ff7eec466f23e91ecf92d0eac8cda377fc750d8ba503c3`；同批 tgz 消费 SHA-256 `c15e8631c7aac351f75fc0cb0206d0de59ab97bad6a687384dd639b82a76cbdc`；矩阵 SHA-256 `65b1ec61801e0a231dd711a38158e463b9e0c06c11b42a448d188ffae5ccc6bb` | `ready_for_gate` |
| P4-UX-DOCX-01 | Codex `/root` | 同会话浏览器自测；不是候选发布独立验收 | `2026-07-11T03:58:46+0800` | `2026-07-11T08:46:09+0800` | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 工作区未提交；无结果提交 | 依赖 P4-API-01=`ready_for_gate`。固定官方 `extend-hq/ui@f2ff2f90954acfed1f60b7fd070a9491d4113047` 后，补齐 Viewer 页码、缩放、缩略图、下载、主题与响应式外壳；Editor 完成真实 DOM 选区、格式、撤销重做、回车拆段、段落合并、只读阻写、表格与方形环绕图片交互。`pnpm typecheck=0`、`pnpm build=0`、组件回归=0；正式 `BB-UX-PARITY` 的四项 DOCX 用例在 `1440×900` 和 `1280×720` 首轮 PASS；只读 Viewer/Editor 六页完整扫描与像素对比 PASS；真实 tgz 外部消费中的 DOCX 查看、编辑、撤销、只读和 Worker/WASM 均 PASS。失败基线、首次失败和重试证据全部保留。 | macOS arm64；Node `v22.13.0`；pnpm `9.0.6`；Python `3.12.2`；Playwright `1.61.0`；Chromium `149.0.7827.55` | 基线 `.../UX-PARITY-BASELINE/20260711T035624+0800/local/`；正式界面 `.../BB-UX-PARITY/20260711T-ux-parity-formal-19/local/summary.json`（PASS，SHA-256 `e94002c26c395a2e68d4591f13731d9d59652c8eab3a8bb79697aac49c005b4b`）；完整黑盒界面子套件 `.../P1-CI-01/20260711T-ux-p4-blackbox-formal-06/blackbox/formal-ux-parity/summary.json`（PASS，SHA-256 `19edf7462dace12538d038c6409b262d198bfb97785b1eb7cbb22e943ef5e6db`）；真实 tgz `.../P1-CI-01/20260711T-ux-p4-consumer-formal-02/consumer/summary.json`（PASS，SHA-256 `2e5c0787b42ad94eabd4219b4f406709c8bdf650387f7310a2875150181873bf`） | `ready_for_gate` |
| P4-UX-XLSX-01 | Codex `/root` | 同会话浏览器自测；不是候选发布独立验收 | `2026-07-11T04:10:00+0800` | `2026-07-11T08:46:09+0800` | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 工作区未提交；无结果提交 | 依赖 P4-API-01=`ready_for_gate`。工作区扩展到至少 200×50，补齐普通值/公式显示、工作簿样式、合并单元格、范围与行列选择、表头高亮、行列尺寸、冻结区域、图表/图片定位、搜索、右键命令、图表工作表、只读阻写和紧凑工具栏。正式 `BB-UX-PARITY` 四项 XLSX 用例在两个视口首轮 PASS；真实 tgz 外部消费验证 200×50、公式、标签、编辑、撤销、只读、Worker/WASM 和无 404；性能三轮门禁 PASS，加载中位数 `586.1ms`、帧耗时 95 分位 `10.2ms`、离页活动 Worker/对象 URL 为 0。 | 同上 | 正式界面与真实 tgz 证据同上一行；性能 `.../P3-PERF-BASELINE-01/20260711T-ux-p4-perf-gate-02/performance/summary.json`（PASS，SHA-256 `741be092aa7dd685a15d11c4f7b5ab90e10e759fad169e3c93adeca2dda14d6c`）；竞态 `.../BB-RACE/20260711T-ux-p4-race-formal-02/formal-race-regression/summary.json`（PASS，SHA-256 `5f299f9de563a9c0d0028ad8ae66598bb174fd1b5728589999730ea8fe28cf01`） | `ready_for_gate` |

| P4-SCOPE-01 | Codex `/root` | 本会话开发后自测与分工只读复核；不是候选发布独立验收 | `2026-07-11T09:42:45+08:00` | `2026-07-11T10:07:46+08:00` | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 工作区未提交；无结果提交 | 失败基线中类型、正式构建、DOCX 集成、XLSX 结构和差异检查均退出 0，但命名审计发现 772 行旧六包引用并退出 1。随后五个公开包统一改为 `@arcships/*`、私有运行包改为 `@arcships/office-runtime`，内部依赖、源码导入、构建配置、测试、发布脚本、锁文件和当前文档同步迁移；上游 `@extend-ai/react-*` 与版权记录保留。首次完整 `BB-RELEASE` 在性能阶段因不稳定的峰值堆比值失败，唯一重试通过但按规则记为 `FLAKY`；证据证明回落绝对堆仅比批准基线高 0.0087%，真实 DOM/生命周期/Worker/对象 URL 均正常，因此把该比值改为只诊断，其他泄漏门禁不变。最终 `CI_RUN_ID=20260711T-arcships-scope-release-02 pnpm test:release=0`，15 步全部 PASS；新五包公开访问检查、双隔离构建、工作区外消费和六组合矩阵通过；旧六包名审计为 0；未发布、未提交、未暂存、未改远端。 | macOS arm64；Node `v22.13.0`；pnpm `9.0.6`；Python `3.12.2`；Playwright `1.61.0`；Chromium `149.0.7827.55`、Firefox `151.0`、WebKit `26.5` | 基线 `.../P4-SCOPE-01/20260711T-arcships-scope-01/baseline/summary.txt`；首次发布自测 `.../BB-RELEASE/20260711T-arcships-scope-release-01/release/summary.json`（FAIL）；最终发布自测 `.../BB-RELEASE/20260711T-arcships-scope-release-02/release/summary.json`（PASS，SHA-256 `39b9923c4650f9146bfa0c38951037e4e1af2c8a758f555572a875b382251d8b`）；候选清单 SHA-256 `98cd784e1e05afb5206966ccc267ecb683361973511bb106c7720c28225c01c6`；外部消费 SHA-256 `0a008bf130784d61a838fcd2e2f1eb3114c6a9c979a908a650251015c7df9ef9`；矩阵 SHA-256 `ad0246a5989e3318ded2843bbc552d8ecfd1bd5259648ab3ea461bf8d82ed775` | `ready_for_gate` |
| P4-RELEASE-01（`@arcships` 首次外部发布前置） | Codex `/root` | 候选发布仍需全新会话与 npm 组织权限复核 | `2026-07-11T09:26:11+08:00` | 未结束 | `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | 无 | GitHub 已确认当前账号是 `arcships` 活跃管理员；五个目标 npm 包名均未公开占用。当前 npm 凭证返回 `E401`，尚未证明 `@arcships` 发布权限；五包还需在最终 GitHub 仓库确定后补准确 `repository`；可信发布运行环境和首次发布、候选标签晋级方式需按 npm 当前规则收口。以上都涉及外部身份或发布方向，本任务不擅自创建仓库、登录、发布或改权限。 | 同上；外部状态只读检查 | `.../P4-SCOPE-01/20260711T-arcships-scope-01/external-access-audit.json` | `pending` |

### 12.3 阶段安全点

阶段全量检查通过后在下表登记安全点，再把该阶段所有 `ready_for_gate` 任务更新为 `done`。

| 阶段 | 集成负责人 | 验证者 | 检查时间 | 检查提交 | 全量检查命令与退出码 | 证据位置 | `LAST_SAFE_COMMIT` |
|---|---|---|---|---|---|---|---|
| P0（工作区全量自测；不设安全点） | Codex `/root` | 本会话交叉复测；不是候选发布独立验收 | `2026-07-10T17:45:05+0800` | 工作区未提交；无结果提交 | `pnpm typecheck=0`；`pnpm build=0`；`node packages/vue-docx/tests/verify-integration.mjs=0`（54 通过）；`node packages/vue-xlsx/test/structure.mjs=0`（保留既有 Vue 生命周期 warning）；`git diff --check=0`；正式构建 `BB-P0-ROUTES=PASS`、`BB-DOCX-WORKER=PASS`、`BB-SEC-URL=PASS`、`BB-RACE=PASS`、`BB-PACK-CONSUMER=PASS` | `.../BB-P0-ROUTES/20260710T173253+0800-self/summary.json`；`.../BB-DOCX-WORKER/20260710T173642+0800-p0-gate/summary.json`；`.../BB-SEC-URL/20260710T173650+0800-p0-gate-agent/`；`.../BB-RACE/20260710T172137+0800-style-retest-final/summary.json`；`.../P0-STYLE-01/20260710T171141+0800/`；`.../BB-PACK-CONSUMER/20260710T164800+0800/` | 不设置：路线图要求同一结果提交和全新会话候选复核；当前工作区未提交，P0 任务保持 `ready_for_gate` |
| P1（工作区连续三轮全量自测；不设安全点） | Codex `/root` | 本会话复测；不是候选发布独立验收 | `2026-07-10T19:00:08+0800` | 工作区未提交；无结果提交 | 同一工作区连续三次 `pnpm check=0`；每轮均包含类型检查、全仓正式构建、DOCX/XLSX 单元与集成测试、Vue 组件测试、正式预览路由与业务流程、DOCX/XLSX 竞态回归、五个真实 tgz 的工作区外消费测试、55 项文档契约、真实类型错误和子任务失败传播探针、`git diff --check`；三轮均首次通过，无重试、无 `FLAKY` | `.../P1-GATE/20260710T185600+0800-run1/check/summary.json`（PASS，SHA-256 `bc0c388c0e1c97926b686920c3a1652d886060329e2f61ac5d64de6d4798f0f3`）；`.../20260710T190000+0800-run2/check/summary.json`（PASS，SHA-256 `8d98d93490cbec0c48940fea430b578a1ae0c13eca9b014bcd83eef30c263f06`）；`.../20260710T190200+0800-run3/check/summary.json`（PASS，SHA-256 `23ab08a9ed209f4d0ce28f491e377136ce3b967785dfd4ee93525f435012c661`） | 不设置：用户要求不提交、暂存、推送或创建分支；没有可登记的结果提交，也没有全新会话独立复核，因此九个 P1 任务保持 `ready_for_gate`。P1 开发与本会话自测已完成，进入 P2 前仍需由全新会话复跑阶段门禁并登记结果提交 |
| P2（工作区全量自测；不设安全点） | Codex `/root` | 本会话复测；不是候选发布独立验收 | `2026-07-10T21:07:48+0800` | 工作区未提交；无结果提交 | `pnpm check=0`：全仓类型与正式构建、39 项核心单元行为、纯核心依赖方向、18 项死路径、五包公开接口/声明指纹、54 项 DOCX 集成、22 项 Vue 组件、正式 preview 路由/业务/DOCX 一致性与编辑/安全/配置/11 项竞态、五个真实 tgz 工作区外消费、55 项文档合同、失败传播和 `git diff --check` 全部 PASS；黑盒用例均首次通过，无重试或 `FLAKY` | `.../P2-GATE/20260710T210500+0800/check/summary.json`（PASS，SHA-256 `a0de9d35b579adeb9e8da02c61f571c9248aae8f1040efdc0cc3a343ea8527c6`）；`.../unit/summary.json`（SHA-256 `67591f8200266ac398434db515fd03b0285f5c176aaab4cffdbf093a5d6ff590`）；`.../component/summary.json`（SHA-256 `fc433d2a9bf413e3808757d7be469b3bf3617b00a7dc448fd22442e4308c1415`）；`.../blackbox/summary.json`（SHA-256 `005038003c136350852a5d581715aed8ad93ce1ce513e783995f0c257398379a`）；`.../consumer/summary.json`（SHA-256 `1468a8a3bff7eee573b5ec390c44f97b6ce2d473304dcd4278cefaaa39f20f75`）；`.../docs/summary.json`（SHA-256 `ef195c3d9c208730a426bdf900d881f592ba1507c591861b52d2a5429c341043`） | 不设置：用户禁止提交、暂存、推送或创建分支，当前没有结果提交；本会话只能完成开发后自测，不能代替全新会话对公开接口和候选版本的独立复核。因此八个 P2 任务保持 `ready_for_gate`，P2 开发与本会话自测已完成，候选发布前需另一个全新会话执行 `BB-RELEASE` |
| P3（工作区全量自测；不设安全点） | Codex `/root` | 本会话交叉复测；不是候选发布独立验收 | `2026-07-11T02:26:01+0800` | 工作区未提交；无结果提交 | `CI_RUN_ID=20260711T-p3-stage-gate-final pnpm check=0`：类型、正式构建、单元、组件、正式 preview 八组黑盒、压力、五个真实 tgz 工作区外消费、文档、失败传播和差异检查全部 PASS；`CI_PREBUILT=1 CI_RUN_ID=20260711T-p3-performance-final pnpm test:performance=0`：批准预算连续三轮 PASS。PDF 只执行公开 `maxFileSize` 的单一体积上限，正常 PDF 的查看、翻页、缩放、搜索和下载保持可用。 | `.../P1-CI-01/20260711T-p3-stage-gate-final/check/summary.json`（PASS，SHA-256 `56a5cf2e199c2991517fd016e9dc88dc7fcba8c2c270736087150fe461a35638`）；`.../P3-PERF-BASELINE-01/20260711T-p3-performance-final/performance/summary.json`（PASS，SHA-256 `a1ac967d5772f20ccc3634bee23248a2abf1fd4d8e9f903346c88d6a8f015b42`） | 不设置：当前结果包含未提交的 P3/P4 接口收口，且没有全新会话独立复核；P3 各任务保持 `ready_for_gate`。开发与本会话全量自测已经完成，候选发布前仍由新会话执行 `BB-RELEASE` 并登记结果提交。 |
| P4（工作区完整发布自测；不设安全点） | Codex `/root` | 本会话交叉复测；不是候选发布独立验收 | `2026-07-11T08:59:26+0800` | 工作区未提交；HEAD 为 `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | `CI_RUN_ID=20260711T-ux-p4-release-final-02 pnpm test:release=0`；整改后的 15 个显式步骤全部 PASS。正式黑盒含固定上游八项 DOCX/XLSX 界面交互、安全、实例隔离和 11 项竞态；压力与三轮性能通过；两次隔离构建有效内容一致；外部消费和 Vue `3.2.25`/`3.5.39` × Chromium/Firefox/WebKit 六组合都严格复用最终候选同一批五个 tgz。 | `.../BB-RELEASE/20260711T-ux-p4-release-final-02/release/summary.json`（SHA-256 `e96107f39fe7f98db3be10b840e2ed62413045e2959ff05a4dab58f55823ff79`）；`.../candidate/candidate-manifest.json`（SHA-256 `02ee4c3bec24a48eb4a08566252e1d42aef9797ce284d366e658eb053e33bbfb`）；矩阵摘要 SHA-256 `65b1ec61801e0a231dd711a38158e463b9e0c06c11b42a448d188ffae5ccc6bb`；双构建摘要 SHA-256 `e49d3520a1bdb81ca3fbe9f3dc36eb8e9dc1529bc38d89efd520eafaea5698dc` | 不设置：当前仍是未提交工作区，且本会话是实现者自测。P4-API/PACK/MATRIX/DOC/UX/RELEASE 保持 `ready_for_gate`，P4-RC-01 保持 `pending`；候选发布前必须由全新会话按 `BB-RELEASE` 独立复核并登记结果提交。 |

| P4（`@arcships` 包名迁移后的完整自测；不设安全点） | Codex `/root` | 本会话开发后自测；不是候选发布独立验收 | `2026-07-11T10:05:00+08:00` | 工作区未提交；HEAD 为 `af0546b18bd2a23cf5bfb05c37a8e46aed605038` | `CI_RUN_ID=20260711T-arcships-scope-release-02 pnpm test:release=0`；类型、正式构建、单元、组件、正式黑盒、压力、修正后的三轮性能、双隔离打包、同批 tgz 材料化、工作区外消费、两 Vue×三浏览器矩阵、文档、就绪、候选和差异检查 15 步全部 PASS；正式黑盒、压力、消费和矩阵均没有第二次尝试 | `.../BB-RELEASE/20260711T-arcships-scope-release-02/release/summary.json`（SHA-256 `39b9923c4650f9146bfa0c38951037e4e1af2c8a758f555572a875b382251d8b`）；候选清单 SHA-256 `98cd784e1e05afb5206966ccc267ecb683361973511bb106c7720c28225c01c6`；黑盒 SHA-256 `5fd9f48ff6fa973fbe5ba5cfb4fe814198f7d7533d6ff962a61ffc453bc4d151`；双构建 SHA-256 `7b34536fba025a0a7b9cd89af5bdc06054a401b8af5819ea43afdd54f0ff5a0f` | 不设置：当前仍是未提交工作区且本会话是实现者自测。P4-SCOPE/PACK/MATRIX/DOC/UX 达到 `ready_for_gate`；P4-RELEASE 因 npm 登录、最终 GitHub 仓库和首次可信发布流程仍为 `pending`，P4-RC-01 继续 `pending`。 |

### 12.4 回退事件

触发第 10.1 节的回退判断时立即登记；即使最终决定不回退，也要保留决策和理由。

| 事件 ID | 任务 ID | 失败提交 | 当时的 `LAST_SAFE_COMMIT` | 复现命令 | 证据位置 | 影响范围 | 提出人 | 批准人 | 安全复核人 | 决定与理由 | 回退提交 | 回退后必跑命令与退出码 | 状态 |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
