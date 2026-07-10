# 项目架构审查与目标设计

> 文档状态：提议作为后续整改的架构基线，尚未表示代码已经按本文完成。
>
> 审查日期：2026-07-10。
>
> 代码基线：提交 `41913ca03666` 加当前工作区；审查时 `packages/docx-core/tsup.config.ts` 有未提交修改，因此涉及该文件的结论以当前工作区为准。
>
> 适用范围：`packages/docx-core`、`packages/vue-docx`、`packages/xlsx-core`、`packages/vue-xlsx`、`packages/vue-extend`、`apps/demo`，以及构建、测试、发布和文档流程。

## 1. 文档目的和配套文档

本文回答三件事：当前项目为什么还不能作为稳定组件库发布；长期应当采用什么包边界、加载流程和渲染结构；从现状迁移到目标结构时必须守住哪些约束。

本文不维护任务状态，也不重复列出全部黑盒用例。四份核心文档的职责如下：

| 文档 | 负责内容 | 发生冲突时的处理方式 |
|---|---|---|
| 本文 | 架构边界、数据流、公开接口方向和迁移约束 | 长期结构以本文为准 |
| [项目稳定化整改路线图](./plan/stabilization-roadmap.md) | 阶段、任务、依赖、状态和完成证据 | 实际执行顺序和状态以路线图为准 |
| [端到端与黑盒验收测试方案](./end-to-end-blackbox-test-plan.md) | 用户可见行为、测试材料、运行命令、证据和发布门禁 | 能否交付以黑盒验收结果为准 |
| [Agent 黑盒验收执行手册](./testing/agent-execution-runbook.md) | Agent 如何启动环境、执行用例、保存证据和清理现场 | 实际操作方法以执行手册为准 |

[DOCX 迁移架构](./docx-migration-architecture.md)、[XLSX 迁移架构](./xlsx-migration-architecture.md)、`docs/plan/tasks` 和 `docs/plan/reviews` 是历史迁移输入，仍可用于了解上游功能和拆分经过，但不再代表当前代码状态或目标架构。

特别说明：旧方案把“单文件不超过 1000 行”列为硬约束，见 [DOCX 文档“硬约束”](./docx-migration-architecture.md#硬约束) 和 [XLSX 文档“硬约束”](./xlsx-migration-architecture.md#硬约束)。本文把它降为代码维护提示，不把它当成主要架构目标。文件是否拆分，应由职责、状态所有权、依赖方向和可独立测试性决定；不能为了行数把同一职责机械切散，也不能因为文件未超 1000 行就容忍边界混乱。

## 2. 总体结论

项目已经具备较完整的 DOCX/XLSX 迁移骨架、真实样本、部分 Worker/WASM 能力和核心层验证，但目前仍是“可继续开发的迁移仓库”，不是“可稳定安装和发布的组件库”。

问题的共同根源不是某个大文件，而是同一件事在不同层有不同答案：

- 源码入口与发布压缩包内容不一致；
- 普通 TypeScript 检查与 Vue 单文件组件的真实类型结果不一致；
- 主线程与 Worker 的 WASM 配置不一致；
- DOCX Viewer 与 Editor 对同一模型使用不同渲染路径；
- demo 页面写着“已通过”，但自动检查没有验证对应用户行为；
- 文件大小已有局部限制，但 URL、全归档解压、图片、取消、缓存和请求先后顺序没有统一规则。

因此，建议接受以下架构决定：

1. 先暂停公开发布和大规模新增功能，优先消除正式运行、类型检查、URL 安全和发布包阻断。
2. 新增一个很薄的工作区私有包 `@extend-ai/office-runtime`，只提供输入、URL 规则、资源上限、取消、请求编号、缓存和统一错误的实现，不负责解析 DOCX/XLSX，也不作为第六个公开发布包。
3. DOCX 和 XLSX 都改成实例化运行配置；WASM、Worker、限制和缓存属于实例，不能继续依赖模块级可变全局变量。
4. Worker 默认是必须成功的执行路径。只有调用方明确允许、文件低于安全阈值并留下诊断记录时，才可以退回主线程。
5. DOCX Viewer 与 Editor 共用同一页面渲染树；编辑器只额外提供选区、命令、浮层和输入事件。
6. 发布接口只暴露稳定入口、类型、样式、Worker/WASM 资源和少量扩展点；内部组件和深层文件不再默认公开。
7. 是否完成不再由源码注释、demo 表格或结构测试判断，而由正式构建黑盒、真实 Worker/WASM、真实发布压缩包消费测试共同判断。

## 3. 审查方法和已验证事实

### 3.1 本轮执行结果

以下命令在 2026-07-10 的上述工作区执行：

| 检查 | 结果 | 能证明什么 | 不能证明什么 |
|---|---|---|---|
| `pnpm typecheck` | 退出码 0 | 当前各包的普通 `tsc --noEmit` 通过 | 不能证明 `.vue` 文件没有类型错误 |
| `packages/vue-docx/node_modules/.bin/vue-tsc -p apps/demo/tsconfig.json --noEmit` | 退出码 2；发现 DOCX Editor 模型缺 3 个必填字段，主题类型也不匹配 | demo 的 Vue 类型错误真实存在 | 不能替代运行时黑盒 |
| `node packages/vue-docx/tests/verify-integration.mjs` | 50 通过、0 失败、0 跳过 | DOCX WASM 解析/序列化、基础分页和纯编辑操作可工作 | 测试明确允许 Worker 退回主线程，不能证明浏览器 Worker 可用 |
| `node packages/vue-xlsx/test/structure.mjs` | 脚本退出码 0，但出现 Vue 生命周期警告 | 控制器对象和若干方法存在且空状态调用不抛错 | 没有加载工作簿，也没有验证渲染、编辑结果和资源清理 |
| 五个包分别执行 `npm pack --dry-run --json` | 全部只包含源码、配置或测试文件，不包含声明的 `dist` 入口 | 当前发布压缩包内容不可用 | 不能代替后续真实打包和独立项目安装 |

上述局限可以直接从测试实现看出：DOCX 测试接受 `worker` 或 `main-thread` 两种来源，[verify-integration.mjs 第 194～209 行](../packages/vue-docx/tests/verify-integration.mjs#L194-L209)；XLSX 结构测试只在空控制器上检查方法存在并用无害参数调用，[structure.mjs 第 34～92 行](../packages/vue-xlsx/test/structure.mjs#L34-L92)。

上一轮正式构建浏览器检查曾观察到 DOCX Viewer 的 Worker/WASM 解析失败和 DOCX Editor 崩溃。当前工作区尚未按新的黑盒方案重新执行完整浏览器套件，因此本文不会把“现已修复”写成事实；相反，下文按仍然存在的代码根因将其列为阻断项。

### 3.2 已有的良好基础

以下部分值得保留并继续建设：

- 工作区已经按 `docx-core`、`vue-docx`、`xlsx-core`、`vue-xlsx`、`vue-extend` 和 demo 拆分，[根 `package.json` 第 5～12 行](../package.json#L5-L12) 给出了统一构建和类型检查入口。
- 核心包和 Vue 包都开启了严格 TypeScript，例如 [docx-core 的 `tsconfig.json` 第 2～13 行](../packages/docx-core/tsconfig.json#L2-L13) 和 [vue-docx 的 `tsconfig.json` 第 2～14 行](../packages/vue-docx/tsconfig.json#L2-L14)。
- DOCX 导入会在完成、失败和取消后移除监听并终止 Worker，[docx-import.ts 第 160～184 行](../packages/docx-core/src/viewer/docx-import.ts#L160-L184)。
- XLSX 控制器在卸载时释放图片资源并终止 Worker，[useXlsxViewerController.ts 第 353～361 行](../packages/vue-xlsx/src/composables/useXlsxViewerController.ts#L353-L361)。
- XLSX 已有 25 MiB 输入限制和部分 ZIP 中央目录预检；当前阈值见 [internal.ts 第 38～42 行](../packages/vue-xlsx/src/composables/internal.ts#L38-L42)，预检逻辑见 [formatting.ts 第 119～150 行](../packages/vue-xlsx/src/composables/formatting.ts#L119-L150)。这可以作为统一资源限制的起点，而不是丢弃重写。
- DOCX 缩略图已经使用有数量上限的 LRU，[thumbnail-cache.ts 第 1～39 行](../packages/docx-core/src/viewer/thumbnail-cache.ts#L1-L39)。后续应补充按字节限制和统一销毁，不需要推翻现有队列设计。
- demo 已有正常、损坏、较大、含图片、中文和图表等多种材料，[samples/manifest.json 第 1～81 行](../apps/demo/public/samples/manifest.json#L1-L81)，可以迁移为正式测试材料库。

## 4. 问题分级

本节的“阻断”表示未修复前不应发布；“高”表示会造成安全、数据一致性或长期维护风险，应在结构重构前或同时处理；“中”表示不会单独阻止最小修复，但会明显增加包体、性能或后续变更成本。

### 4.1 阻断问题

#### B1. DOCX 主线程和 Worker 使用两套 WASM 配置状态

demo 从 `@extend-ai/docx-core` 根入口调用 `setWasmSource`，[main.ts 第 11～15 行](../apps/demo/src/main.ts#L11-L15)。根入口导出的同名函数来自 engine；它只更新 engine 模块中的 `overrideSource`，[engine/wasm.ts 第 14～29 行](../packages/docx-core/src/engine/wasm.ts#L14-L29)。

Worker 导入链路却从另一个 `viewer/wasm-source.ts` 读取 `configuredWorkerWasmSource`，[docx-import.ts 第 3～7 行](../packages/docx-core/src/viewer/docx-import.ts#L3-L7) 和 [第 223～231 行](../packages/docx-core/src/viewer/docx-import.ts#L223-L231)。只有调用 viewer 层自己的 `setWasmSource` 才会记录该值，[viewer/wasm-source.ts 第 44～67 行](../packages/docx-core/src/viewer/wasm-source.ts#L44-L67)。但 viewer barrel 又明确不导出这个模块，并建议使用一个包清单没有声明的深层路径，[viewer/index.ts 第 4～7 行](../packages/docx-core/src/viewer/index.ts#L4-L7)；实际包只导出根入口，[docx-core/package.json 第 5～13 行](../packages/docx-core/package.json#L5-L13)。

这会导致“主线程配置成功、Worker 收到 `undefined`”。Worker 随后按自身文件位置解析默认 WASM 地址。源码构建、正式构建和发布包的相对目录不同，因此局部测试通过不能证明正式资源路径正确。

此外，Worker 构造失败会静默回到主线程，[docx-import.ts 第 146～155 行](../packages/docx-core/src/viewer/docx-import.ts#L146-L155)。这会让测试把“Worker 不可用”误判成“导入成功”。

目标处理：用 `createDocxRuntime(config)` 建立实例；同一个实例把同一份 WASM 配置传给主线程和 Worker；默认 Worker 失败即返回结构化错误，不能静默掩盖。

#### B2. 当前类型门禁漏过真实 Vue 错误

根类型检查只递归调用各包的 `typecheck`，[根 `package.json` 第 9～13 行](../package.json#L9-L13)。Vue 包和 demo 的脚本仍然使用普通 `tsc --noEmit`，例如 [vue-docx/package.json 第 14～17 行](../packages/vue-docx/package.json#L14-L17) 和 [apps/demo/package.json 第 6～10 行](../apps/demo/package.json#L6-L10)。

`DocModel.metadata` 明确要求 `sourceParts`、`warnings`、`headerSections`、`footerSections` 和 `paragraphStyles`，[types.ts 第 418～437 行](../packages/docx-core/src/engine/types.ts#L418-L437)。demo 的 starter model 只提供页眉和页脚数组，[DocxEditorPage.vue 第 101～103 行](../apps/demo/src/pages/DocxEditorPage.vue#L101-L103)，并在 [第 184～188 行](../apps/demo/src/pages/DocxEditorPage.vue#L184-L188) 传入不匹配的主题对象。普通根检查显示通过，直接运行 `vue-tsc` 却准确报告这两个问题。

目标处理：所有含 `.vue` 的项目统一使用 `vue-tsc --noEmit`；类型检查必须先于构建；demo 必须使用公开的 `createBlankDocumentModel()` 或新的模型工厂，不再手写不完整元数据。

#### B3. 五个包声明了 `dist` 入口，但发布压缩包不包含入口

五个包都把 `main`、`module` 和 `types` 指向 `dist`；以 [docx-core/package.json 第 5～13 行](../packages/docx-core/package.json#L5-L13) 和 [vue-xlsx/package.json 第 5～13 行](../packages/vue-xlsx/package.json#L5-L13) 为例。根 `.gitignore` 又忽略所有 `dist`，[.gitignore 第 5～9 行](../.gitignore#L5-L9)。包清单没有 `files`，也没有发布前构建和压缩包消费测试。

本轮对五个包执行 `npm pack --dry-run --json`，结果均排除了 `dist`。`docx-core` 和 `xlsx-core` 的构建配置虽然复制 Worker/WASM，例如 [docx-core/tsup.config.ts 第 4～23 行](../packages/docx-core/tsup.config.ts#L4-L23) 和 [xlsx-core/tsup.config.ts 第 6～22 行](../packages/xlsx-core/tsup.config.ts#L6-L22)，但这些文件最终仍没有进入当前压缩包。

Vue 包还没有统一的样式出口。`vue-docx`、`vue-xlsx` 只导出根 JS，[vue-docx/package.json 第 8～13 行](../packages/vue-docx/package.json#L8-L13)；`vue-extend` 使用了带内部目录名的 `./dist/index.css`，[vue-extend/package.json 第 8～14 行](../packages/vue-extend/package.json#L8-L14)，不是稳定的 `./style.css`。

目标处理：每个包声明明确的 `files` 和 `exports`；发布前从干净目录构建；用真实压缩包安装到独立项目验证 JS、类型、CSS、Worker 和 WASM。不能只依赖工作区链接或 `npm pack --dry-run`。

### 4.2 高优先级问题

#### H1. URL、PDF iframe 和下载没有统一安全规则

PDF 组件先下载原地址，只检查前几个字节是否以 `%PDF-` 开头，[PdfViewer.vue 第 141～170 行](../packages/vue-extend/src/components/PdfViewer.vue#L141-L170)，随后又把原地址直接放入没有 `sandbox` 的 iframe，[第 199～214 行](../packages/vue-extend/src/components/PdfViewer.vue#L199-L214)。第一次检查通过不等于 iframe 的第二次请求一定返回相同内容；同源地址、重定向和服务端内容变化都可能绕过第一次检查。项目虽然已经声明 EmbedPDF 依赖，[vue-extend/package.json 第 30～44 行](../packages/vue-extend/package.json#L30-L44)，当前组件并未使用它渲染页面。

XLSX 的 URL 加载直接 `fetch(src)`，没有协议和来源规则，[workbook-state.ts 第 474～507 行](../packages/vue-xlsx/src/composables/workbook-state.ts#L474-L507)；下载时把同一个 `src` 原样写入 `<a href>` 并自动点击，[clipboard.ts 第 67～75 行](../packages/vue-xlsx/src/composables/clipboard.ts#L67-L75) 和 [workbook-state.ts 第 770～778 行](../packages/vue-xlsx/src/composables/workbook-state.ts#L770-L778)。

目标处理：URL 必须在网络请求、iframe、图片和下载之前通过同一规则；PDF 使用受控 PDF 引擎从已验证字节渲染；下载使用已取得的字节生成 Blob URL，不再自动点击未经验证的原始地址。

#### H2. 文件资源限制是局部检查，不是完整预算

XLSX 已检查输入字节、最大 worksheet XML、shared strings 和部分 XML 总量，但它只统计特定条目，[formatting.ts 第 125～145 行](../packages/vue-xlsx/src/composables/formatting.ts#L125-L145)。后续代码仍可能用 `unzipSync` 一次性展开整个归档，[xlsx-worker.ts 第 202～208 行](../packages/xlsx-core/src/xlsx-worker.ts#L202-L208)。其他图片、图表、关系文件、条目数量、单项压缩比例和图片解码像素没有纳入同一预算。

DOCX 的公开解析入口直接把整个 ArrayBuffer 交给 WASM，[ooxml-core.ts 第 20～23 行](../packages/docx-core/src/engine/ooxml-core.ts#L20-L23)，没有对应的公开资源限制。TIFF 转换还使用模块级、无数量和字节上限的 Map，[image-render.ts 第 38～40 行](../packages/docx-core/src/viewer/image-render.ts#L38-L40)。

目标处理：压缩包元数据预检和实际解压过程都必须计数；不能只相信可伪造的中央目录数值。图片在分配像素缓冲前检查宽高和总像素；解析超时或取消时终止 Worker。

#### H3. 异步加载没有统一的“只接受最新结果”规则

`DocxViewer.vue` 直接在 watcher 中等待解析并写回状态，没有 AbortController 或请求编号，[DocxViewer.vue 第 264～295 行](../packages/vue-docx/src/components/DocxViewer.vue#L264-L295)。快速切换文件时，较慢的旧请求可能最后完成并覆盖新文件。

`useDocxModel` 虽然有 AbortController，但使用一个布尔值表示当前迭代；每次加载先设为 `false` 又立即设为 `true`，随后捕获的仍是 `true`，[useDocxModel.ts 第 22～49 行](../packages/vue-docx/src/composables/useDocxModel.ts#L22-L49)。这不是可以区分多次请求的代次编号。

XLSX 的 watcher 做得更好：它同时使用 AbortController 和 `isCurrent`，并在回写前检查，[useXlsxViewerController.ts 第 384～454 行](../packages/vue-xlsx/src/composables/useXlsxViewerController.ts#L384-L454)。这套做法应上移为共享规则，并补上递增请求编号、保存/导出任务和 Worker 回包处理。

目标处理：每个 viewer/controller 只有一个加载协调器；新请求必须取消旧请求并递增编号；任何阶段回写前都校验编号；旧结果即使已算完也只能释放，不能更新页面。

#### H4. DOCX Viewer 和 Editor 的渲染路径已经分叉

只读 `DocxViewer.vue` 自己循环页面、段落、图片和表格，[DocxViewer.vue 第 22～211 行](../packages/vue-docx/src/components/DocxViewer.vue#L22-L211)。Editor 则通过 `DocxViewerRoot` 和 `DocxPageSurface` 渲染，[DocxEditor.vue 第 9～27 行](../packages/vue-docx/src/components/DocxEditor.vue#L9-L27)。这不是“同一渲染面加不同交互”，而是两套页面实现。

Editor 路径中仍有明确占位：表格单元格正文始终返回空字符串，[DocxTableHost.vue 第 228～235 行](../packages/vue-docx/src/components/DocxTableHost.vue#L228-L235)；页眉、页脚和表单域固定为 `false`，页面封面固定为空，[DocxPageSurface.vue 第 152～170 行](../packages/vue-docx/src/components/DocxPageSurface.vue#L152-L170)；图片拖动只计算位移但不更新模型，缩放结束时仍提交原尺寸，[DocxImageLayer.vue 第 165～213 行](../packages/vue-docx/src/components/DocxImageLayer.vue#L165-L213)。

目标处理：建立唯一的 `DocxDocumentSurface`；Viewer 和 Editor 都使用同一页面、段落、表格、图片、页眉页脚组件。未支持能力必须返回明确警告或隐藏入口，不能用固定空值同时在 demo 中声称“已实现”。

#### H5. Worker 的承诺和实际使用范围不一致

XLSX 公共选项说明 `useWorker` 会在 Worker 中解析工作簿，[worksheet-types.ts 第 536～549 行](../packages/xlsx-core/src/types/worksheet-types.ts#L536-L549)。实际实现只在请求只读或因大文件被强制只读时使用 Worker，[useXlsxViewerController.ts 第 135～168 行](../packages/vue-xlsx/src/composables/useXlsxViewerController.ts#L135-L168)；普通可编辑文件最终走主线程加载，[第 416～443 行](../packages/vue-xlsx/src/composables/useXlsxViewerController.ts#L416-L443)。Worker 的部分错误还会触发主线程回退，[第 207～210 行](../packages/vue-xlsx/src/composables/useXlsxViewerController.ts#L207-L210)。

目标处理：文档和实现必须一致。解析、结构提取和高成本计算默认在 Worker；可编辑状态通过命令或增量数据同步，不应成为必须主线程解析的理由。若某阶段暂不支持，公开选项要如实命名并给出诊断结果。

#### H6. 自动检查尚不能作为合并和发布门禁

根脚本只有 `dev`、`build` 和普通 `typecheck`，[根 `package.json` 第 9～13 行](../package.json#L9-L13)，没有根级单元、组件、黑盒、发布包消费和持续集成入口。demo 页面还把若干项目硬编码为 `READY` 或 `PASSED`，例如 [DocxViewerPage.vue 第 40～49 行](../apps/demo/src/pages/DocxViewerPage.vue#L40-L49) 和 [XlsxViewerPage.vue 第 74～83 行](../apps/demo/src/pages/XlsxViewerPage.vue#L74-L83)。这些文字不是测试证据。

目标处理：测试分为核心、Vue 组件、正式构建黑盒和真实压缩包消费四层；控制台异常、Vue 警告、关键网络失败和资源 404 默认让测试失败。详细执行方法见[端到端与黑盒验收测试方案](./end-to-end-blackbox-test-plan.md)。

### 4.3 中优先级问题

#### M1. XLSX 大表滚动存在重复线性扫描

可见行列范围每次从索引 0 累加到当前滚动位置，[XlsxGrid.vue 第 117～145 行](../packages/vue-xlsx/src/components/XlsxGrid.vue#L117-L145)；获取行列偏移和命中测试也重复线性累加，[第 147～160 行](../packages/vue-xlsx/src/components/XlsxGrid.vue#L147-L160) 和 [第 188～203 行](../packages/vue-xlsx/src/components/XlsxGrid.vue#L188-L203)。越接近大表底部，单次滚动和点击成本越高。

目标处理：预计算累计偏移数组，用二分查找定位可见行列；尺寸变化只增量更新受影响部分；滚动绘制使用 `requestAnimationFrame` 合并到一帧。

#### M2. 地图数据和重型渲染能力进入默认模块图

区域图模块静态导入美国和世界地图，[chart-region-map.ts 第 1～18 行](../packages/vue-xlsx/src/render/chart-region-map.ts#L1-L18)；surface 模块又静态导入同一批数据，[chart-surface.tsx 第 1～5 行](../packages/vue-xlsx/src/render/chart-surface.tsx#L1-L5)。`vue-xlsx` 根入口还直接导出完整 render 模块，[src/index.ts 第 7～14 行](../packages/vue-xlsx/src/index.ts#L7-L14)。本地现有构建中的 `vue-xlsx/dist/index.js` 约 1.9 MB，但该文件是被忽略的本地生成物，正式预算仍应由干净构建和持续集成重新测量。

目标处理：地图、复杂图表和 WebGL 作为可选功能动态加载；默认表格查看不应下载这些数据；对主入口、功能分包和 WASM 分别设置预算。

#### M3. 公开接口过宽，内部重排会变成破坏性变更

`vue-docx` 根入口直接导出 ViewerRoot、PageSurface、ParagraphHost、TableHost、ImageLayer 等内部组件，[vue-docx/src/index.ts 第 1～20 行](../packages/vue-docx/src/index.ts#L1-L20)，还导出大量内部注册和组合函数。这样会让消费端依赖当前目录结构，后续统一渲染面很难安全迁移。

目标处理：稳定入口只保留高层组件、运行实例、控制器、模型和明确承诺的扩展点；内部组件转到未公开目录。确有二次开发需要的能力，应以插槽、渲染回调或独立 `advanced` 入口提供，并单独说明兼容级别。

#### M4. 当前文档状态与代码事实不一致

文档索引仍在[“包导出状态”](./INDEX.md#包导出状态)中把 `vue-xlsx` 和 `vue-docx` 写成 pending 或 stub；旧计划的[“任务列表”](./plan/README.md#任务列表)也仍列出大量已经存在的模块为 pending。这些文档记录了迁移时点，不适合继续作为当前状态源。

目标处理：本文负责目标结构，稳定化路线图负责当前任务状态，黑盒方案负责验收；旧迁移文档在标题处标记“历史资料”和替代入口。Agent 开工前不得从旧文件中的 `done`、`pending` 或“完整”注释推断真实状态。

## 5. 设计原则

后续设计和代码审查统一使用以下原则：

1. **一个状态只有一个所有者。** WASM 地址、Worker 工厂、资源限制、当前加载任务和缓存分别只有一个实例负责，不能在 engine、viewer 和 demo 各留一份可变状态。
2. **公开行为先于内部实现。** 先定义输入、输出、错误、取消、清理和发布资源，再决定文件如何拆分。
3. **框架层不解析文件。** Vue 层负责生命周期、交互和展示；格式解析、模型、命令和序列化放在核心层。
4. **渲染只读模型，命令修改模型。** 渲染组件不直接改文档对象；编辑事件转换为命令，由控制器产生新模型和历史记录。
5. **默认安全失败。** Worker、WASM、URL 或资源检查失败时给出明确错误；不得为了“能打开”而静默扩大权限、无限解压或在主线程重跑大任务。
6. **取消和销毁是正常路径。** 文件切换、路由离开、组件卸载和超时都必须有可测试的资源释放行为。
7. **发布压缩包才是产品。** 工作区源码别名、已有 `dist` 和 demo 根目录资源只能帮助开发，不能作为可发布性的证据。
8. **测试读取用户可见结果。** 黑盒测试不调用内部函数制造状态，不把源码注释或页面上的“PASSED”当结果。
9. **按职责拆分，不按行数拆分。** 一个模块应当有清楚输入、输出和测试；行数只用于提醒复核，不决定边界。
10. **先兼容迁移，再删除旧入口。** 旧 API 需要明确弃用期、适配层和迁移说明，不能在大重构中顺手移除。

## 6. 目标数据流

```text
File / ArrayBuffer / URL
        │
        ▼
DOCX Runtime / XLSX Runtime（实例配置）
        │
        ├── 调用 office-runtime 的无状态工具
        │   （来源标准化、URL 规则、下载、资源计数、取消和统一错误）
        │
        ├── createLoader()：每个视图一个加载器和一个专用 Worker
        │                     │
        │                     └──► Worker 内初始化同一 WASM 来源
        │                     │
        │                     ▼
        │                 解析与结构提取
        │
        ▼
DocumentSession / WorkbookSession
（模型、原始包、警告、耗时、资源句柄、dispose）
        │
        ├── 布局与分页 / 行列索引
        │          │
        │          ▼
        │     唯一 Vue 渲染面
        │          │
        │          ├── 只读：无输入监听和编辑浮层
        │          └── 编辑：选区 + 命令 + 历史 + 浮层
        │
        └── 序列化 / 导出 ──► 已验证字节 ──► Blob URL 下载

销毁时：controller 释放 session；loader 取消任务并终止 Worker；Runtime 最后清空无引用缓存并拒绝新任务。
```

`office-runtime` 是被格式 Runtime 调用的私有实现库，不是排在格式 Runtime 前面的第二个有状态服务。唯一所有权如下：

| 对象 | 唯一负责的状态 |
|---|---|
| DOCX/XLSX Runtime | 不可变配置、全实例任务序号、Worker 工厂、共享只读缓存及其引用计数、已创建 loader 清单 |
| loader | 当前请求、内部 AbortController、最新任务 id、一个专用 Worker 及其监听器 |
| session | 当前文档/工作簿的可编辑模型、版本号、对象 URL、图片位图和对共享缓存项的引用 |
| Vue controller | 当前 loader 和 session；来源切换或卸载时按 session → loader 的顺序释放 |
| PDF Runtime | URL 策略、全实例任务序号、PDF 引擎/Worker 工厂、可调整的整份文件体积上限和已验证字节 |
| PDF controller | 一个 loader、当前 PDF session、内部 AbortController、最新任务 id、页面渲染队列和临时页面图片；来源切换或卸载时全部释放 |
| `office-runtime` 包 | 只提供上述对象使用的类型和无状态实现，不持有模块全局配置、任务或缓存 |

格式 Runtime 只处理对应格式，不操作 Vue 状态；Vue 组件只接收 session、模型、布局和命令接口。默认每个 loader 使用一个专用 Worker，避免多个视图互相取消；未来若改成 Worker 池，仍必须沿用 Runtime 全局唯一任务 id 和相同的资源隔离规则。

## 7. 目标包和目录边界

### 7.1 包职责

| 包 | 目标职责 | 不应继续承担 |
|---|---|---|
| `@extend-ai/office-runtime`（新增） | 输入类型、URL 规则、下载、资源预算、AbortSignal、请求编号、按字节缓存工具、统一错误和诊断字段 | 持有运行状态、DOCX/XLSX 解析、Vue 状态、页面渲染 |
| `@extend-ai/docx-core` | DOCX Runtime、Worker、WASM 适配、OOXML 模型、布局、编辑命令、序列化 | Vue 组件、demo 状态、隐式模块全局配置；纯入口不直接依赖 DOM |
| `@extend-ai/xlsx-core` | XLSX Runtime、Worker、WASM 适配、工作簿模型、公式/图表/图片数据、编辑命令、导出 | Vue 组件、静默主线程回退；纯入口不假设 Worker 一定有 DOMParser |
| `@extend-ai/vue-docx` | session/controller 与 Vue 生命周期连接、唯一文档渲染面、只读和编辑交互 | 重复解析、第二套布局、直接修改核心模型 |
| `@extend-ai/vue-xlsx` | controller、网格、选区、编辑交互、按需图表/地图/WebGL 渲染 | URL 安全和资源限制的私有副本、默认静态加载所有可选能力 |
| `@extend-ai/vue-extend` | PDF、上传、签名、缩略图等通用 Vue 组件 | 用原始同源 iframe 代替受控 PDF 渲染、各组件自行制定 URL 规则 |
| `apps/demo` | 使用公开入口演示和承载黑盒页面 | 源码深层导入、硬编码“已通过”、替发布包提供缺失资源 |

`@extend-ai/office-runtime` 在本轮确定为工作区私有包，`package.json` 使用 `"private": true`，不作为第六个发布包。任何公开包只要引用它，都必须把所需实现打入自己的 `dist`，且生成的 `.d.ts` 不能要求消费端安装这个私有包。五个现有公开包的发布和消费矩阵保持不变。将来若确有外部直接使用需求，应另开架构决定，不能在本轮中顺带公开。

### 7.2 建议目录

目录名可以在实现时微调，但依赖方向不能改变：

```text
packages/
├── office-runtime/src/
│   ├── source.ts          # File/bytes/URL 标准化
│   ├── url-policy.ts      # 协议、来源、重定向、凭据规则
│   ├── limits.ts          # 资源预算与计数器
│   ├── load-task.ts       # AbortSignal、请求编号、超时
│   ├── cache.ts           # 按字节和数量限制的实例缓存
│   ├── errors.ts          # 稳定错误码与用户消息映射
│   └── diagnostics.ts     # 字节、耗时、执行路径和警告
│
├── docx-core/src/
│   ├── runtime/           # createDocxRuntime、session、加载编排
│   ├── worker/            # Worker 入口与消息协议
│   ├── engine/            # WASM/OOXML 适配；逐步消除全局状态
│   ├── model/             # 模型、校验、标准化、克隆
│   ├── layout/            # 分页和页面布局
│   ├── commands/          # 文本、段落、表格、图片编辑命令
│   └── serialize/         # 保存和导出
│
├── vue-docx/src/
│   ├── controller/        # useDocxSession、useDocxEditorController
│   ├── surfaces/          # 唯一页面/段落/表格/图片渲染面
│   ├── interactions/      # 选区、输入、拖放、缩放、快捷键
│   ├── overlays/          # 工具栏、菜单、批注、修订、缩略图
│   └── public.ts          # 最小公开入口
│
├── xlsx-core/src/         # runtime/worker/model/commands/serialize + 现有 charts/images
├── vue-xlsx/src/
│   ├── controller/
│   ├── grid/
│   ├── interactions/
│   └── optional/          # chart、map、webgl 动态入口
└── vue-extend/src/
    ├── pdf/               # 受控 PDF 引擎适配
    └── components/
```

不要求一次性移动所有现有文件。先建立新边界和适配入口，再按功能迁移；旧目录在没有测试保护时不得大规模改名。

### 7.3 依赖方向

允许的依赖方向是：

```text
office-runtime
      ↑
docx-core / xlsx-core
      ↑
vue-docx / vue-xlsx
      ↑
demo 和外部消费项目
```

`vue-extend` 可以依赖 `office-runtime` 的 URL 和资源规则，但不应反向成为 DOCX/XLSX 核心依赖。核心包的纯数据入口不得导入 Vue；确实需要 Canvas、字体测量或 Worker 的浏览器能力时，放在明确的 `browser` 或 `runtime` 子入口，并通过工厂函数注入，避免导入核心模型就访问 `window`、`document` 或 `Worker`。

## 8. 实例化 Runtime 设计

### 8.1 为什么必须实例化

当前 DOCX 和 XLSX 都有模块级 WASM Promise、配置或缓存，例如 [docx engine/wasm.ts 第 14～16 行](../packages/docx-core/src/engine/wasm.ts#L14-L16) 和 [xlsx-core/wasm.ts 第 1～15 行](../packages/xlsx-core/src/wasm.ts#L1-L15)。模块全局状态会造成：

- 同页两个 viewer 不能使用不同 CDN、限制或 Worker；
- 测试先后顺序影响结果；
- 首次初始化后难以更换来源；
- 组件销毁不能真正释放缓存和资源；
- 主线程和 Worker 很容易再次出现两套状态。

Runtime 实例必须拥有：WASM 来源、Worker 工厂、Worker 策略、资源限制、URL 规则、缓存、诊断回调和销毁方法。模型纯函数仍可直接导出，不要求所有函数都变成类方法。

### 8.2 共享配置草案

以下代码是目标接口草案，不是当前已经存在的 API。命名可在实现阶段小幅调整，但能力和所有权不能删减。

```ts
const MiB = 1024 * 1024

const officeLimits = {
  maxInputBytes: 25 * MiB,
  maxArchiveEntries: 10_000,
  maxUncompressedBytes: 256 * MiB,
  maxSingleEntryBytes: 200 * MiB,
  maxCompressionRatio: 100,
  maxSharedStringsBytes: 50 * MiB,
  maxSingleImagePixels: 40_000_000,
  maxTotalImagePixels: 100_000_000,
  maxParseMs: 30_000,
}

const pdfLimits = {
  maxFileSize: 50 * MiB,
}

const officeUrlPolicy = {
  enabled: true,
  baseUrl: hostBaseUrl,
  allowRelativeUrl: true,
  allowedProtocols: ["https:"],
  allowedOrigins: hostAllowedOrigins,
  allowHttpOnLocalhost: false,
  credentials: "omit" as const,
  redirect: "error" as const,
}
```

`hostBaseUrl` 和 `hostAllowedOrigins` 由宿主在浏览器挂载时传入，库代码不直接读取 `window` 或 `document`。SSR/Node 若只处理字节，应传 `enabled: false`，此时 URL 来源直接返回 `SOURCE_NOT_ALLOWED`；需要在 Node 下载时也必须显式传入基准地址、允许来源和 fetch 实现。

`officeLimits` 用于 DOCX/XLSX 的压缩包、XML、图片和解析过程。`pdfLimits` 独立且只限制整份 PDF 的输入字节：默认 `50 MiB`，宿主可通过公开 `maxFileSize` 配置调整，不另设第二层隐藏硬上限。超过当前配置值时返回 `PDF_TOO_LARGE`，不得创建 PDF 引擎、页面图片或缩略图。

其余数值是 DOCX/XLSX 的应用默认基线，不是永远不变的产品承诺。其中 25 MiB、50 MiB、200 MiB 和 256 MiB 与当前 XLSX 局部限制接近，便于平滑迁移；条目数、压缩比例、图片像素和超时需要用真实材料和目标设备压测后确认。DOCX 页数、工作表数和历史记录上限应由具体产品配置，不能埋在组件内部。PDF 不再按页数、单页像素、画布总内存或同时渲染页数拒绝正常文件；这些指标可用于性能诊断，但不是显示前置条件。

DOCX/XLSX 的应用配置之上还要有调用方不能提高的库内硬上限。初始建议为：输入 100 MiB、ZIP 条目 50,000 个、展开总量 512 MiB、单条目 256 MiB、压缩比例 200、单图 1 亿像素、全部图片 2.5 亿像素、单次解析 120 秒。实际生效值取“应用配置与硬上限中较小者”。提高硬上限必须修改版本化常量、补压力与安全报告并经过代码审查，不能通过 Runtime 参数完成；普通大文件若超过硬上限，应转到受控服务端处理。此处第二层硬上限不适用于 PDF；PDF 只有公开 `maxFileSize` 这一层配置。

### 8.3 DOCX Runtime 示例

```ts
import {
  bundledDocxWasmUrl,
  createBundledDocxWorker,
  createDocxRuntime,
} from "@extend-ai/docx-core/runtime"

const docxRuntime = createDocxRuntime({
  wasmSource: bundledDocxWasmUrl,
  worker: {
    create: createBundledDocxWorker,
    mode: "required",
    allowMainThreadBelowBytes: 0,
  },
  limits: {
    ...officeLimits,
    maxPages: 500,
  },
  urlPolicy: officeUrlPolicy,
  cache: {
    maxEntries: 8,
    maxBytes: 64 * MiB,
  },
  onDiagnostic(event) {
    // 只记录大小、耗时、执行路径和错误码，不记录文档正文或带令牌的 URL。
    reportOfficeDiagnostic(event)
  },
})

const loader = docxRuntime.createLoader()
const session = await loader.load(
  { kind: "url", url: "/documents/contract.docx" },
  { signal: routeAbortController.signal },
)

// session.model / session.package / session.warnings / session.timings
// 组件卸载时：
session.dispose()
loader.dispose()
docxRuntime.dispose()
```

`bundledDocxWasmUrl` 和 `createBundledDocxWorker` 由已发布包提供，二者引用同一包内资源。需要使用私有 CDN 时，调用方可以显式替换 `wasmSource` 和 Worker 工厂，但不能再调用全局 `setWasmSource` 改变其他实例。

### 8.4 XLSX Runtime 示例

```ts
import {
  bundledXlsxWasmUrl,
  createBundledXlsxWorker,
  createXlsxRuntime,
} from "@extend-ai/xlsx-core/runtime"

const xlsxRuntime = createXlsxRuntime({
  wasmSource: bundledXlsxWasmUrl,
  worker: {
    create: createBundledXlsxWorker,
    mode: "required",
    allowMainThreadBelowBytes: 0,
  },
  limits: {
    ...officeLimits,
    maxWorksheets: 200,
    maxWorksheetXmlBytes: 200 * MiB,
  },
  urlPolicy: officeUrlPolicy,
  cache: {
    maxEntries: 4,
    maxBytes: 64 * MiB,
  },
})

const loader = xlsxRuntime.createLoader()
const session = await loader.load(
  { kind: "bytes", bytes: uploadedBuffer, name: uploadedFile.name },
  { signal: viewAbortController.signal },
)
```

可编辑工作簿也应先在 Worker 中完成解析。编辑命令可以在主线程模型上执行，或者通过消息发送给 Worker；具体方式由性能测试决定，但不得因为 `readOnly: false` 就把整个解析过程默认放回主线程。

图表、地图和 WebGL 的动态导入由 `vue-xlsx` controller 或组件负责，不进入 `xlsx-core` Runtime 配置。这样核心包只产生图表和图片数据，不反向依赖 Vue 渲染包。

### 8.5 兼容入口

迁移期可以暂时保留 `setWasmSource`，但它只能配置一个明确的“默认 Runtime”，并标记弃用。它已经属于公开包入口，因此按公开 API 版本规则处理：计划在 `0.2.0` 引入实例 API 并开始警告，在整个 `0.x` 系列保留兼容入口，最早在 `1.0.0` 删除：

- 调用一次必须同时影响默认实例的主线程和 Worker；
- 首次使用后再次调用要给出明确错误；
- 不得影响已经显式创建的实例；
- 文档和示例全部改用 `createDocxRuntime` / `createXlsxRuntime`；
- `1.0.0` 删除前必须在发布说明和迁移文档中列出替代代码，并由真实压缩包消费测试覆盖旧、新两种入口。

## 9. 统一加载、安全和资源管理

### 9.1 输入类型

所有格式统一接受以下三类来源：

```ts
type OfficeSource =
  | { kind: "file"; file: File }
  | { kind: "bytes"; bytes: ArrayBuffer; name?: string }
  | { kind: "url"; url: string }
```

组件不再同时维护互相覆盖的 `file`、`src` 和临时内部状态。若为兼容保留旧属性，适配层必须在入口处转换成 `OfficeSource`，同时传入两种来源时直接报参数错误。

### 9.2 加载阶段

每次加载都按固定阶段执行，并记录同一个全局唯一 `taskId`：

1. 标准化输入，检查文件名只用于展示，不能用于判断真实格式。
2. 若输入是 URL，先用宿主提供的 `baseUrl` 解析相对地址，同时检查协议和来源；不通过就不发请求。
3. 使用 `redirect: "error"` 的受控 fetch 下载；浏览器直连遇到任何重定向都失败。业务确需重定向时，改走受控同源服务端代理，由服务端逐跳校验后返回最终字节。
4. 下载或读取字节；有 `Content-Length` 时提前拒绝，实际读取仍再次计数。
5. 检查文件签名和格式；MIME 只作为辅助信息。
6. 读取 ZIP 元数据并建立预算；实际解压时再次累计条目、字节和压缩比例。
7. 在 Worker 中初始化同一 Runtime 配置并解析；超时或取消就终止 Worker。
8. 校验并标准化模型，产生警告、诊断和 session。
9. 只有 `taskId` 仍是 loader 的最新任务时，才把 session 交给 controller；否则立即 `dispose()`。
10. 组件卸载或来源切换时，释放上一 session 的缓存、对象 URL 和监听器。

### 9.3 URL 规则

默认规则如下：

- 相对 URL 先基于宿主注入的 `baseUrl` 解析。解析后的地址必须**同时**满足协议允许清单和来源允许清单；“是 HTTPS”不能绕过来源检查。
- 默认只允许当前站点的 HTTPS 来源。跨来源 HTTPS 必须逐项加入 `allowedOrigins`；开发环境可单独允许 localhost HTTP。
- 默认拒绝 `javascript:`、`data:`、`file:` 和未知协议。调用方产生的 Blob 应以 `File` 或字节传入；若确需 `blob:`，必须由受信任的宿主显式开启。
- 默认 `credentials: "omit"`。确需 cookie 的同源业务必须显式配置，不能让第三方 URL 携带凭据。
- 浏览器 fetch 固定使用 `redirect: "error"`，不提供“允许跟随”的 Runtime 开关。业务确需重定向时，必须由受控同源服务端代理逐跳检查协议、来源、目标网段和次数上限，再把最终字节返回浏览器；浏览器仍只访问这个同源代理。
- 下载和 PDF 渲染使用已经取得并验证的字节；不再二次导航到原始 URL。
- 外部超链接和文档内链接使用独立规则，只允许明确协议，并加 `noopener noreferrer`。
- 错误消息不回显带查询令牌的完整 URL；诊断可记录去掉敏感参数后的来源标识。

### 9.4 资源预算

DOCX/XLSX 的资源预算必须在 Runtime 配置中显式出现。PDF 只保留整份文件体积限制，具体如下：

| 类别 | 必须限制的项目 | 检查时机 |
|---|---|---|
| 输入 | 下载字节、本地文件字节、下载时间 | 下载前和读取过程中 |
| ZIP | 条目数、展开总字节、单条目字节、压缩比例、路径长度 | 中央目录预检和实际解压过程 |
| XML/文本 | 总 XML、单 worksheet、shared strings、单文本节点或关系数量 | 解析前和解析过程中 |
| 图片 | 压缩字节、宽高、单图像素、总像素、并发解码数 | 创建像素缓冲和对象 URL 之前 |
| DOCX | 节点数、页面数、页眉页脚/批注/图片数量 | 模型建立和布局过程中 |
| XLSX | 工作表数、有效行列、公式数、图表/图片/形状数 | 模型建立和可选能力加载前 |
| PDF | 整份文件字节；默认 `50 MiB`，公开配置可调整 | 本地文件读取前；URL 有 `Content-Length` 时在读取正文前，否则在取得字节后、创建引擎前 |
| 运行时间 | 下载、解析、布局、缩略图和导出超时 | 各阶段独立计时 |
| 内存近似值 | 原始字节、展开字节、模型估算、缓存和历史记录 | session 生命周期内持续累计 |

超限错误必须指出“哪一项、实际值、允许值”，但不泄露内部路径和堆栈。DOCX/XLSX 对可信的大文件场景只能把应用配置提高到库内硬上限；再大的文件改用服务端转换。PDF 的宿主可显式调整 `maxFileSize`，调整后的值就是唯一生效上限；组件不得静默取消它。

`office-runtime` 负责预算类型、计数器和统一错误；DOCX/XLSX Runtime 必须在实际解压、图片解码和模型构建时持续消耗同一个计数器，并把数值传入 Worker。解压适配器必须按条目或数据块工作，在分配输出缓冲前调用 `budget.reserve(kind, bytes)`；禁止先用 `unzipSync` 全量展开再统计。中央目录预检是第一道检查，实际输出计数是最终检查。

若现有 WASM 只能一次性解析、无法在过程中检查上限，就只能先做字节和归档预检并放入可终止 Worker；在 WASM 支持内部计数前，不得声称已经完整防住解压炸弹。PDF 由 `vue-extend` 的 PDF controller 持有 session；它在引擎打开前执行唯一的整份文件体积检查，默认 `50 MiB`，并使用宿主公开传入的 `maxFileSize` 覆盖默认值。页数、单页像素、总内存和并发页不再是 PDF 功能的阻塞预算。渲染仍须按需执行，并在来源切换或卸载时取消任务、关闭文档、终止 Worker、撤销对象 URL，但这些属于生命周期正确性，不是额外的文件拒绝规则。

### 9.5 请求编号与取消

每个 Runtime 创建时生成稳定 `runtimeId`，并维护全实例递增的 `taskSequence`。`taskId` 使用 `${runtimeId}:${taskSequence}`；每个 loader 只记录自己的最新 `taskId`。这样即使以后多个 loader 共享 Worker，消息 id 也不会碰撞。不能让每个 loader 都从 1 开始编号。

外部 AbortSignal 与 loader 的内部 AbortController 通过 `AbortSignal.any` 或等价兼容函数合并；任一信号取消都终止当前任务，但 loader 不得反向调用或改变调用方传入的 signal：

```ts
let latestTaskId: string | undefined
let activeAbort: AbortController | undefined

async function load(source: OfficeSource, externalSignal?: AbortSignal) {
  const taskId = runtime.allocateTaskId()
  latestTaskId = taskId
  activeAbort?.abort()
  activeAbort = new AbortController()
  const signal = combineAbortSignals(activeAbort.signal, externalSignal)

  const session = await runtime.load(source, {
    taskId,
    signal,
  })

  if (taskId !== latestTaskId) {
    session.dispose()
    throw new OfficeLoadError("STALE_RESULT")
  }

  return session
}
```

所有 Worker 消息都带完整 `taskId`。session 另有稳定 `sessionId` 和从 0 开始的 `version`；每次编辑命令成功后版本递增。保存、导出和缩略图任务必须捕获 `{ sessionId, version }`，完成时确认当前 session 和版本仍一致，再更新状态或触发下载；若产品允许导出旧快照，必须由调用方显式传入该快照，不能误当成当前版本。AbortSignal 只能说明“调用方不再需要结果”，无法中断的同步步骤完成后仍要检查 `taskId` 和版本并丢弃结果。

### 9.6 缓存和释放

缓存遵守以下规则：

- 共享缓存只属于 Runtime；它只保存已验证的原始字节和不可变解析快照，不保存可编辑模型、DOM、对象 URL、Canvas、ImageBitmap 或 Vue 响应式对象。
- session 从不可变快照建立自己的工作模型。编辑 session 必须克隆或使用写时复制，两个 session 绝不能共享可变模型；对象 URL、位图、选区和历史记录只属于 session。
- Runtime 缓存同时按条目数和估算字节限制；key 至少包含输入内容摘要、解析器版本和影响结果的配置，不能只使用文件名或 URL。
- session 获取缓存项时增加引用计数，`session.dispose()` 时减少。LRU 只能淘汰引用计数为 0 的项；`runtime.dispose()` 先取消并释放其 loader，再拒绝新任务，最后清空无引用项并使尚未释放的 session 失效。
- session 资源在替换、取消和销毁时撤销对象 URL、关闭位图并移除监听器；这些资源不进入 Runtime 共享缓存。
- 编辑历史按步骤数和字节双重限制；大对象快照优先使用结构共享或命令日志。
- 测试要连续切换文件并观察 Worker、对象 URL、监听器和内存是否回落，不能只断言 `Map.size`。

### 9.7 统一错误

建议稳定错误码至少包括：

| 错误码 | 含义 |
|---|---|
| `SOURCE_NOT_ALLOWED` | URL 协议、来源或重定向不允许 |
| `FETCH_FAILED` | 网络失败、状态码错误或读取中断 |
| `INPUT_TOO_LARGE` | 输入字节超过限制 |
| `ARCHIVE_LIMIT_EXCEEDED` | ZIP 条目、展开量或压缩比例超限 |
| `IMAGE_LIMIT_EXCEEDED` | 图片字节或像素超限 |
| `WORKER_UNAVAILABLE` | Worker 无法创建或启动 |
| `WASM_LOAD_FAILED` | WASM 地址、内容类型、编译或初始化失败 |
| `PARSE_FAILED` | 文件损坏或解析器拒绝 |
| `TIMED_OUT` | 阶段超过时间限制 |
| `ABORTED` | 调用方主动取消 |
| `STALE_RESULT` | 旧请求完成但已不是最新请求 |
| `UNSUPPORTED_FEATURE` | 文件包含当前明确不支持的能力 |

面向用户显示简短中文消息；完整错误码、阶段、耗时和执行路径进入测试报告。原始堆栈只在开发诊断中保留。

## 10. 统一 DOCX 渲染面

### 10.1 目标结构

```text
DocxViewer / DocxEditor（薄包装）
              │
              ▼
        useDocxSession
              │
              ▼
      DocxDocumentSurface
              │
      ┌───────┼────────┐
      ▼       ▼        ▼
  PageList  Thumbnail  Gutter
      │
      ▼
  DocxPageSurface
      ├── PageBackground / PageBorder
      ├── Header
      ├── Body
      │   ├── Paragraph
      │   ├── Table
      │   └── InlineImage
      ├── Footer
      └── FloatingImage

编辑模式另外挂载：Selection、Input、ContextMenu、DragResize、Comments。
```

### 10.2 只读和编辑的边界

基础渲染面输入是 `model + layout + pageRange + theme`，输出 DOM/Canvas 和命中区域，不直接修改模型。两种模式的差别如下：

| 能力 | 只读 | 编辑 |
|---|---|---|
| 页面、段落、表格、图片、页眉页脚渲染 | 同一实现 | 同一实现 |
| 文字选择和复制 | 可按产品需要开启 | 开启 |
| `contenteditable`、输入法和键盘编辑 | 不挂载 | 由交互层挂载 |
| 右键菜单、拖动、缩放和表单编辑 | 不挂载 | 产生命令，不直接改模型 |
| 批注和修订 | 同一显示组件 | 额外提供接受、拒绝、编辑命令 |
| 工具栏和历史记录 | 不需要 | controller 提供 |

### 10.3 布局和渲染契约

- 分页只在 `docx-core/layout` 中完成。Vue 层不得再根据 DOM 高度生成另一套页边界，只能回传测量结果供明确的二次布局使用。
- 每个渲染块有稳定 id，来自模型位置或生成的节点 id；不能用循环序号替代编辑定位。
- 表格单元格使用与正文相同的段落 renderer，不再单独返回 HTML 空字符串。
- 页眉、页脚、背景、边框、图片和表单域要么完整进入 layout/result，要么产生 `UNSUPPORTED_FEATURE` 警告；不能固定 `false`。
- 图片拖动和缩放先显示临时预览，结束时提交包含新位置/尺寸的单一命令；取消时恢复，不写入历史。
- Viewer 与 Editor 只读模式对同一 fixture 的结构快照和关键截图必须一致。允许的差异仅是编辑装饰层。

### 10.4 迁移方法

1. 先把当前 Editor 的 `DocxPageSurface` 补到能完整展示已承诺能力，并为每种块建立结构测试。
2. 新建薄的 `DocxDocumentSurface`，内部复用补齐后的 PageSurface。
3. 让新的 Viewer 在功能开关后使用该渲染面，保持旧 Viewer 用于对比。
4. 对同一组 DOCX 做 Viewer、新 Viewer、Editor 只读三方截图和结构对比。
5. 达到等价后删除旧 Viewer 内部的独立段落/表格循环；保留高层 `DocxViewer` 组件名作为兼容包装。
6. 再把编辑输入、选区和浮层逐项从渲染组件中分离，避免在统一过程中同时重写所有编辑逻辑。

## 11. 发布接口和资源设计

### 11.1 核心包

目标 `exports` 至少包含：

```json
{
  "files": ["dist", "README.md", "LICENSE"],
  "exports": {
    ".": {
      "types": "./dist/index.d.ts",
      "import": "./dist/index.js"
    },
    "./runtime": {
      "types": "./dist/runtime.d.ts",
      "import": "./dist/runtime.js"
    },
    "./wasm-url": {
      "types": "./dist/wasm-url.d.ts",
      "import": "./dist/wasm-url.js"
    },
    "./assets/docx_wasm_bg.wasm": "./dist/assets/docx_wasm_bg.wasm",
    "./worker": "./dist/docx-import-worker.js",
    "./package.json": "./package.json"
  }
}
```

XLSX 提供对应入口。`wasm-url` 模块返回相对已发布文件计算出的 URL；`createBundledDocxWorker` / `createBundledXlsxWorker` 内部使用静态 `new URL(..., import.meta.url)`，让构建工具能发现资源。自定义 CDN 仍通过 Runtime 配置覆盖。

### 11.2 Vue 包

Vue 包统一提供：

- 根入口：高层组件、controller/composable 和稳定类型；
- `./style.css`：唯一公开样式入口；
- 可选能力入口：例如 `./chart`、`./map`、`./webgl`；
- 必要时提供 `./advanced`，但不从根入口导出所有内部组件；
- `sideEffects` 明确保留 CSS，不依赖消费端碰巧没有摇掉样式。

文档和 demo 不再导入 `@extend-ai/vue-extend/dist/index.css` 这类内部路径，而统一导入 `@extend-ai/vue-extend/style.css`。

### 11.3 构建和发布流程

五个公开包分别执行以下流程；私有 `office-runtime` 只参加工作区类型检查和单元测试，不单独打包发布：

1. 清理自己的 `dist`，从源码构建 JS、声明、CSS、Worker 和 WASM。
2. 验证 `exports` 指向的每个文件存在，且 Worker 引用的 WASM 路径存在。
3. 执行真实 `npm pack`，保存压缩包和文件清单。
4. 在临时目录建立独立 Vue + Vite + TypeScript 项目，只安装压缩包，不使用 workspace 链接和源码 alias。
5. 导入 JS、类型和 `./style.css`，开发构建和正式构建都运行。
6. 启动正式预览，确认 Worker/WASM 请求成功、Content-Type 合理、无 404，加载真实 DOCX/XLSX。
7. 对压缩包内容设置允许清单：不得包含测试材料、源代码绝对路径、临时文件或无关大文件。
8. 检查生成的 `.d.ts` 和运行时代码都不引用私有 `@extend-ai/office-runtime`；需要的实现已打包，公开类型已由对应公开包重新导出。

是否提交 `dist` 到 Git 与是否发布 `dist` 是两件事。可以继续在 Git 中忽略 `dist`，但必须用 `files`、发布前构建和真实压缩包测试保证制品完整。

### 11.4 公开 API 兼容

- 建立公开导出快照，持续集成对新增、删除和类型变化给出报告。
- 现有高层组件名尽量保留；内部实现替换不要求消费端改代码。
- `0.2.0` 引入新 Runtime、稳定样式入口和最小公开入口，同时把已导出的内部组件标记弃用；整个 `0.x` 系列继续可用并给出开发警告，最早在 `1.0.0` 移除。若实际发布计划跳过这些版本，必须先更新本文和迁移说明，不能提前删除。
- 错误码、事件名、属性默认值和资源入口都属于公开契约，不能只比较函数名。
- demo 必须像外部项目一样只使用公开入口；发布验收再用真实压缩包覆盖一次。

这里要区分两类兼容代码：已经从五个包导出的公开 API 按上面的版本周期弃用，不能用“30 天清理”提前删除；只为分阶段重构新增、从未导出且消费端不可见的内部桥接代码，才必须在 30 天内或下一个次版本中删除。兼容记录至少包含 `kind: public-api | internal-bridge`、引入版本/日期、负责人、替代入口和 `removeInVersion` 或 `removeByDate`，路线图据此判断期限。

## 12. 迁移约束

1. **禁止一次性全仓重写。** 先修 P0 阻断和测试，再迁移 Runtime，再统一渲染，最后收敛 API。详细顺序见[稳定化整改路线图](./plan/stabilization-roadmap.md)。
2. **保留已有用户改动。** 当前 `packages/docx-core/tsup.config.ts` 有未提交修改；后续 Agent 必须先检查差异，不得用回退命令覆盖。
3. **每个阶段都要有双向测试。** 既验证正常文件，也验证损坏、超限、取消、危险 URL 和资源缺失；不能靠放宽断言完成迁移。
4. **旧接口通过适配层迁移。** `setWasmSource`、现有组件属性和控制器方法不能在统一渲染面时突然消失；先提供新入口和警告，再按版本计划删除。
5. **模型变化必须集中标准化。** 新增必填字段时，通过模型工厂和 `normalize` 补默认值；demo、测试和消费端不应各自拼出元数据。
6. **Worker 回退必须显式。** 测试环境没有 Worker 时应标记跳过或失败，不能把主线程结果计为 Worker 通过。生产大文件不允许自动主线程重跑。
7. **安全限制不可在组件中被悄悄关闭。** 宿主提高上限要显式配置并进入诊断；demo 不提供“无限制”默认值。
8. **删除旧渲染路径前必须证明等价。** 使用相同 fixture 的结构、截图、编辑往返和无障碍结果；不能仅比较页面是否非空。
9. **可选功能必须可缺省。** 不使用地图、复杂图表、WebGL 或 PDF 编辑时，默认入口不能加载相关大资源。
10. **文档状态由证据驱动。** 任务只有在对应命令、报告和独立验证通过后才能标记完成；源码中的 `complete`、`READY`、`PASSED` 注释不算证据。
11. **历史文档不再派生当前任务。** 旧迁移文档只用于功能清单和实现背景；新的 Agent 任务必须引用本文、路线图和黑盒方案。
12. **单文件 1000 行不是验收门禁。** 超过时复核职责是否过多；拆分后若产生循环依赖、重复状态或无法单测，宁可先保留完整职责再设计正确边界。

## 13. 非目标

本轮架构整改不承诺以下事项：

- 一次性实现 Microsoft Word/Excel/PDF 的所有高级功能或像素级完全一致；
- 更换 DOCX/XLSX 文件格式、重写现有 Rust/WASM 引擎或引入服务端转换平台；
- 新增多人协作、云端存储、权限系统和业务工作流；
- 为了目录整齐而移动所有文件、重命名所有类型或强制每个文件达到相似行数；
- 把 demo 变成正式产品应用；demo 只承担公开 API 示例和自动黑盒入口；
- 保证任何恶意文件都不会触发浏览器或第三方解析器缺陷。本文要求的是最小权限、资源上限、隔离、及时更新依赖和可重复安全测试；它不能替代浏览器与依赖自身的安全维护；
- 在缺少产品数据时永久锁定本文示例中的资源数值。数值应通过安全要求和真实压力测试调整，但限制项目本身不能被删除。

## 14. 架构完成标准

只有同时满足以下结果，才能认为目标架构已经落地：

1. DOCX/XLSX 都由显式 Runtime 实例管理 WASM、Worker、URL、限制、缓存和销毁；模块级兼容状态不再影响显式实例。
2. 同页两个不同配置的实例可以并行工作，互不覆盖资源地址、限制和缓存。
3. 所有正式构建的解析测试都证明真实 Worker 被使用；Worker/WASM 缺失会以明确错误失败。
4. URL、下载、PDF 和文档内外部资源使用同一规则，危险协议和不允许的来源在导航或执行前被拒绝。
5. DOCX Viewer 与 Editor 只读模式共用同一渲染树，关键结构和截图一致；旧重复路径已移除或只剩兼容包装。
6. 资源预算覆盖输入、ZIP、XML、图片、模型规模、运行时间、缓存和历史记录，边界及超限用例自动通过。
7. Vue 类型检查、核心测试、组件测试、正式构建黑盒和发布包消费测试成为根级固定命令和持续集成门禁。
8. 五个真实发布压缩包可在独立项目中导入 JS、类型、`./style.css`、Worker 和 WASM，并完成开发与正式构建。
9. 默认 XLSX 入口不静态加载地图、复杂图表和 WebGL；大表滚动达到黑盒方案的性能预算。
10. `docs/INDEX.md` 能明确找到本文、路线图和黑盒方案；旧迁移文档已标记历史状态，不再出现相互冲突的“当前进度”。

## 15. 与黑盒套件的对应关系

具体步骤和数值以[端到端与黑盒验收测试方案](./end-to-end-blackbox-test-plan.md)为准；架构决定至少由以下套件证明：

| 架构决定 | 黑盒套件 |
|---|---|
| 正式构建路由和控制台可用 | `BB-P0-ROUTES` |
| DOCX Worker/WASM 使用同一实例配置且不伪装回退 | `BB-DOCX-WORKER` |
| URL、PDF、下载和重定向规则 | `BB-SEC-URL` |
| 请求编号、取消、快速切换和卸载清理 | `BB-RACE` |
| 真实压缩包的 JS、类型、CSS、Worker/WASM 消费 | `BB-PACK-CONSUMER` |
| XLSX 按需加载、包体和大表滚动 | `BB-PERF-XLSX` |
| 缓存、对象 URL、Worker 和长时间运行资源回落 | `BB-STRESS` |
| 全量发布门禁和证据包 | `BB-RELEASE` |

单元或组件测试可以更早定位问题，但不能替代上表对应的用户侧结果。

## 16. 执行入口

实施者不要把本文直接当任务清单。按以下顺序开始：

1. 在[稳定化整改路线图](./plan/stabilization-roadmap.md)中选择无前置阻断的任务，记录负责人和状态。
2. 阅读[端到端与黑盒验收测试方案](./end-to-end-blackbox-test-plan.md)中对应套件，先复现并保存失败证据。
3. 只在任务允许的路径内修改，保留工作区其他改动。
4. 运行任务验收和所在阶段的全量检查；实现者不得只凭自己的本地结果把任务标为完成。
5. 交给独立 Agent 使用正式构建和公开入口复核；发布相关任务必须再用真实压缩包验证。

本文后续只在架构决定变化时更新；日常进度、阻断和测试结果分别写入路线图和黑盒报告，避免四份核心文档各自形成不同事实。
