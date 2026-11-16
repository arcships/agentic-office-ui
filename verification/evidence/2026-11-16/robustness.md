# 鲁棒性（V5，2026-11-16）

Chromium 155，本地 demo；临时构造 30 MB / 64 MB 压缩包等输入。

| 用例 | 结果 | 错误码 |
|---|---|---|
| `corrupted.docx` | error，6.6s | `PARSE_FAILED`（找不到目录结束记录） |
| 截断的 `demo.docx`（4 KB） | error，6.5s | `PARSE_FAILED` |
| 30 MB 随机 `.docx` | error，6.5s | `LIMIT_EXCEEDED`（31457284 > 26214400） |
| 30 MB 随机 `.xlsx` | error，6.5s | `INVALID_WORKBOOK`（先按 zip 解析失败，未先报大小） |
| 64 MB 全零 `sheet1.xml` 的 zip 炸弹 | error，6.5s | `LIMIT_EXCEEDED: maxCompressionRatio 1028.76 > 100` |
| 连续上传两个 XLSX | ready | 第二个文件获胜（仅 “Validation” tab） |
| 加载中卸载 | 无 pageerror | — |

观察：

- DOCX 大小限制在解析前生效；XLSX 对“损坏的大文件”先报 `INVALID_WORKBOOK`，没有 fail-fast 到 `LIMIT_EXCEEDED`（与 #54 的“读完才判”相关，但不影响安全）。
- zip 炸弹被压缩比限制拦住，错误码与限值清晰。
- 并发加载 latest-wins 正确；卸载无泄漏错误。
