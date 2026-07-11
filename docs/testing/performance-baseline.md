# 性能基线执行说明

> 状态：Chromium `149.0.7827.55` 性能预算已批准并通过新的三轮正式门禁；候选发布仍需全新会话复核。

本页对应 `P3-PERF-BASELINE-01` 和 `BB-PERF-XLSX`。首次连续三轮只建立待批准基线，不得把结果写成候选发布性能通过。

## 固定条件

- 正式页面：先执行 `pnpm build`，再由脚本启动 `pnpm --filter demo preview`；不得使用开发服务器。
- 浏览器：Chromium `149.0.7827.55`。
- 视口：`1440×900`。
- 设备像素比：`1`。
- 语言与时区：`zh-CN`、`Asia/Shanghai`。
- 每轮使用新的浏览器 context，同一浏览器进程连续执行三轮。
- 每轮大表帧采集时长：`10` 秒，先垂直再水平滚动。
- 材料：`financial-model.xlsx`、`sales-table.xlsx`、`charts-images.xlsx`、`large-grid.xlsx`、`corrupted.xlsx` 和 PDF 资源样本 `sample.pdf`。字节数与 SHA-256 由 [`performance-baseline.json`](../../tests/blackbox/performance-baseline.json) 固定，并与 `test-data/manifest.json` 双向核对。

当前机器为 macOS arm64、Apple M4 Max、128 GiB 内存。不同机器或浏览器版本的结果不得直接覆盖这份基线；必须建立新的可比较系列并重新批准。

## 执行命令

单轮短烟测用于验证脚本能完成全部步骤，不能用于形成基线：

```bash
pnpm build
CI_PERF_PREVIEW_PORT=4181 \
BLACKBOX_EVIDENCE_DIR=output/acceptance/<commit>/P3-PERF-BASELINE-01/<run-id>/smoke \
python tests/blackbox/performance_baseline.py --rounds 1 --scroll-seconds 2
```

正式三轮：

```bash
pnpm test:performance
```

也可以在已经完成同一工作树正式构建时跳过重复构建：

```bash
CI_PREBUILT=1 pnpm test:performance
```

脚本自行启动和停止正式 preview，保存每轮 trace、截图、控制台、页面异常、网络、资源和内存记录。端口被占用、浏览器版本不符或材料校验值变化时结果为 `BLOCKED`，不会混入旧基线。

## 指标口径

构建侧记录五个公开包和 demo 的原始字节、固定 gzip 字节、JS、Worker、CSS、WASM、类型文件、重复 WASM 内容和 PDFium 资源。浏览器侧记录：

- Home 首屏与进入 XLSX 后的传输字节；
- `large-grid.xlsx` 从选择到首个可操作网格的时间；
- 加载和滚动阶段的长任务数量与最差耗时；
- 10 秒滚动的帧耗时中位数、95 分位数和最差值；
- Canvas 尺寸写入次数和真实纵横滚动范围；
- 深滚动位置选择、Tab 提交和撤销延迟；
- 图表材料首次打开新增的 JavaScript 请求；
- 连续切换五个工作簿前后的 Worker、对象 URL、DOM、监听器和 JS 堆；
- 回到 Home 并等待 2 秒后的资源与内存回落。

每项机器指标保存三次原始值、中位数、最差值、最小值和轮间差值百分比。帧指标只有在纵向与横向滚动范围均大于零时才代表真实滚动；否则保留数值并登记 `LARGE_GRID_HAS_NO_SCROLL_RANGE`，不能把空转帧率描述成大表性能通过。

## 预算状态和判定

机器配置的 `status` 有三个阶段：

- `capture`：尚未形成三轮数据；
- `pending_approval`：已形成三轮中位数和最差值，但还不能给候选版本性能 PASS；
- `approved`：后续运行按每项已批准中位数加 `10%` 检查。非零的毫秒指标至少允许 `10ms` 的绝对调度波动；零长任务等零基线仍必须保持为零。超过预算即失败并要求人工分析。

脚本在 `capture` 或 `pending_approval` 阶段成功采集时输出 `BASELINE_CAPTURED_PENDING_APPROVAL`。只有配置明确改为 `approved` 且全部预算满足时才输出 `PASS`。

## 已批准结果

迁移到 Chromium `149.0.7827.55` 时，第一轮采集发现新浏览器会把尚未回收的页面对象计入保留内存，三轮节点数明显漂移，因此没有批准。测量流程随后改为回到首页、等待两秒并回收不可达页面对象，再读取保留内存。

DOCX/XLSX 上游界面整改后，CSS、脚本和真实工作区规模发生了预期变化，因此旧包体数值不再能作为当前候选预算。重新采集时同时修正了一个测量问题：调试 trace 会保留大量历史节点和监听器，不能当作页面仍存活的对象。当前门禁使用浏览器实时 DOM 树、实时元素、连续三轮页面切换增长、Worker、对象 URL 和 JS 堆作为泄漏判断；trace 中的历史节点仍保存在证据里，但不再作为发布阻断数值。新的批准配置已在相同机器和浏览器上重新执行三轮，结果为 `PASS`，预算违规、普通发现和阻断发现均为 0。本次仍是实现者会话内自测，不是候选发布独立验收。

包名迁移后的第一次发布自测只在 `browser.memory.retainedRatio` 失败：回到首页后的堆内存三轮都稳定在约 `16.37 MB`，实时 DOM、连续切换增长、Worker 和对象 URL 也没有变化，但不同运行的峰值堆内存受垃圾回收时机影响约 `13%`，使“回落值 ÷ 峰值”在产品没有新增保留对象时仍会越线；立即重试又恢复通过。因此该比值继续写入结果供诊断，但不再作为发布阻断预算。真实泄漏仍由回到首页后的绝对堆内存、实时 DOM/元素增量、连续切换的节点/监听器增长、活动 Worker 和对象 URL 共同阻断，这些检查没有删除或放宽。

关键结果：

- 大表真实纵向滚动范围每轮都大于 0，正式滚动回归通过；滚动期间没有长任务，也没有重复改写 Canvas 尺寸；
- 图表材料每轮首次新增 2 个 JavaScript 请求，Home 和普通工作簿不加载图表分包；
- 大表可操作时间中位数为约 `586.1ms`，三轮最差约 `591.9ms`；滚动帧耗时 95 分位中位数约 `10.2ms`；
- 离开 XLSX 页面后，活动 Worker 和对象 URL 每轮都回到 0；
- demo 中重复 WASM 数量为 0，Home 主入口不再包含 XLSX 网格实现；
- 回到 Home 后，实时 DOM 树和实时元素相对首页基线都只增加 `8`，三轮完全一致；连续页面切换的节点和监听器增长都为 `0`；回落后的绝对堆内存中位数约为 `16.37 MB`。

当前批准配置：`tests/blackbox/performance-baseline.json`，SHA-256 为 `fcd5787bb331b8247322ce28a1bf7c8bb0f088524fba3fe0fee7ec306426591d`。界面整改后的首次候选运行保存在 `output/acceptance/af0546b18bd2a23cf5bfb05c37a8e46aed605038/P3-PERF-BASELINE-01/20260711T-ux-p4-perf-gate-01/performance/`，它如实记录旧预算失败，不能作为通过证据。批准后的正式三轮门禁位于 `.../20260711T-ux-p4-perf-gate-02/performance/`，总摘要结果 `PASS`，SHA-256 为 `741be092aa7dd685a15d11c4f7b5ab90e10e759fad169e3c93adeca2dda14d6c`。包名迁移的首次失败证据位于 `.../BB-RELEASE/20260711T-arcships-scope-release-01/performance/`；只允许的一次重试位于 `.../P3-PERF-BASELINE-01/20260711T-arcships-scope-perf-retry-01/performance/`，结果按执行规则记为 `FLAKY`，不能作为稳定通过结论。修正指标职责后，定向三轮和随后从头执行的完整发布自测均通过；最终稳定结果位于 `.../BB-RELEASE/20260711T-arcships-scope-release-02/performance/formal-performance-baseline/summary.json`，SHA-256 为 `8f17bd990302c9ba35c0ecfdfe11b9ca49b92e34738248564306a29f940171cd`。
