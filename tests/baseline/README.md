# P3/P4 回归检查

这里的文件最初用于保存 P3/P4 开始时的失败断言。产品修复完成后，源码中的断言都必须通过；历史失败只保存在被 Git 忽略的 `output/acceptance/` 证据目录中，不能让测试源码继续预期失败。

| 文件 | 当前用途 | 自动入口 |
|---|---|---|
| `p3-cache-baseline.test.mjs` | 历史、缩略图、实例缓存和对象地址清理回归 | `pnpm test:component` 的 `p3-regression-baselines` |
| `p3-image-decode-baseline.test.mjs` | DOCX/XLSX 图片解码失败状态回归 | `pnpm test:component` 的 `p3-regression-baselines` |
| `p3-worker-lifecycle-baseline.test.mjs` | Worker 超时、取消、恢复和禁止静默回退 | `pnpm test:component` 的 `p3-regression-baselines` |
| `p3-xlsx-performance-baseline.test.mjs` | XLSX 按需分包和滚动算法静态回归 | `pnpm test:component` 的 `p3-regression-baselines` |
| `p3-xlsx-scroll-baseline.py` | 正式 preview 的大表横纵滚动范围回归 | `pnpm test:performance` 的 `formal-scroll-range-regression` |
| `p4-release-readiness-baseline.test.mjs` | P4 包、矩阵、文档和发布流程就绪检查 | `pnpm test:release` 的 `p4-release-readiness` |

Node 回归需要先完成根构建，并会随 `pnpm check` 和 `pnpm test:release` 执行。浏览器滚动回归只使用正式 preview，随独立的 `pnpm test:performance` 执行；GitHub 共享机器上的时间抖动不再阻断 npm 发布。直接运行叶子套件时会自动构建，父流程调用叶子套件时会复用已经生成的正式构建。
