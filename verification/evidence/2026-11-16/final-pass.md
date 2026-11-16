# 验证框架最终验收（第二轮全量，2026-11-16）

框架计划（L0–L6 / V1–V6）已完成，并连续跑了两轮全量矩阵。第二轮失败项与第一轮**完全相同**，没有新的 issue 类别。

## 第二轮结果

| 套件 | 结果 | 说明 |
|---|---|---|
| `office_oracle.py` | **PASS** | 18 个夹具页数与基线一致；`demo.docx` 仍为 9 页 |
| `real_files.py --viewer` | FAIL（已知） | 2 个 POI XLSX 失败，均为 **#152**（图表系列缺 `<c:tx>`） |
| `roundtrip.py` | FAIL（已知） | XLSX 导出丢 chart part，openpyxl 无法重开，均为 **#66** |
| `robustness.py` | **PASS** | 畸形/超限/zip 炸弹/并发/卸载 |
| `cross_browser.py`（Chromium+Firefox） | **PASS** | 4 查看器 × 2 引擎 |
| `webkit_webdriver.py`（系统 WebKitGTK 2.54.1） | **PASS** | 4 查看器 |
| `pixel_baseline.py` | **PASS** | DOCX 18.6%、PPTX 2.9%、PDF 18.1% |
| `scenario_matrix.py` | **PASS** | 16 视口 + 暗色/缩放/打印/键盘/IME/触摸 |

失败项映射：`#152`（解析崩溃）、`#66`（导出损坏）。第二轮没有新增类别，满足框架收口标准第 4 条。

## 覆盖矩阵

| 格式 | L0 | L1 | L2 | L3 | V2 | L5 | L6 |
|---|---|---|---|---|---|---|---|
| DOCX | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| XLSX | ✅ | ✅ | ✅ | ✅ | ✅（失败 #66） | ✅ | — |
| PPTX | ✅ | ✅ | — | ✅ | 只读 | ✅ | ✅ |
| PDF | ✅ | ✅ | — | ✅ | 只读 | ✅ | ✅ |

## 复现

```bash
uv run python scripts/fixtures/verify_review_corpus.py
uv run python tests/blackbox/office_oracle.py
uv run --with playwright python tests/blackbox/real_files.py --viewer
uv run --with playwright --with python-docx --with openpyxl python tests/blackbox/roundtrip.py
uv run --with playwright python tests/blackbox/robustness.py
uv run --with playwright python tests/blackbox/cross_browser.py --browsers chromium,firefox
WebKitWebDriver --port=4444   # 另一个终端
python tests/blackbox/webkit_webdriver.py
uv run --with playwright --with pillow python tests/blackbox/pixel_baseline.py
uv run --with playwright python tests/blackbox/scenario_matrix.py
```

## 结论

验证框架计划全部完成：L0–L6 每层都有可重复脚本、基线/阈值与证据；V1–V6 的两轮结果稳定；已知失败均可归入既有 issue（#152、#66），没有未归类的新类别。后续进入修复阶段。
