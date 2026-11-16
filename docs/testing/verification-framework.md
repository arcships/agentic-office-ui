# Office UI 验证框架

> 状态：v1（2026-11-16），运行在 `verification` 分支。
> 目标：把“发现问题”从一次性人工检查，变成**可重复、可追溯、可对比**的证据生产流程；在产品修复阶段开始前，尽可能把问题类别挖干净。

## 1. 原则

1. **证据优先**：每个结论必须能指向一份可重跑的输出（JSON、截图、页数、哈希），不接受“看起来不对”。
2. **验证代码与产品代码分离**：本分支只允许改 `tests/`、`scripts/fixtures/`、`test-data/`、`docs/`、`verification/`；`packages/`、`apps/demo/src` 不动。
3. **上游只是对照，不是真理**：同时保留 LibreOffice/pdftoppm 这类独立 oracle；三者不一致时记录三方结果，不预设谁对。
4. **固定版本**：上游仓库记录 commit，npm 包记录版本，LibreOffice/Chromium 记录版本；基线随版本更新。
5. **非目标显式化**：宏执行、协同编辑、PDF 注释创作等不进入验证矩阵，避免“验证范围”无限扩张。

## 2. 分支与目录

- 分支：`verification`（从 `master` 切出）。
- 验证脚本：`tests/blackbox/`、`scripts/fixtures/`。
- 夹具：`test-data/corpus/`（合成，可提交）、`test-data/real/`（公开真实文件，后续）。
- 证据：`verification/evidence/<YYYY-MM-DD>/<suite>/`。
- 文档：`docs/testing/verification-framework.md`、`docs/testing/upstream-parity.md`。

提交约定：

```text
verify: <suite> <date> — <one-line result>
```

每个证据批次一个 commit；产品修复分支引用证据 commit，不直接改本分支的脚本。

## 3. 分层框架（L0–L6）

| 层 | 名称 | 工具 | 产出 | 状态 |
|---|---|---|---|---|
| **L0** | 静态/源码/API 对齐 | 文档清单、上游仓库、npm `d.ts`、模块映射 | 功能/API 差距清单（已并入 #76） | ✅ 完成 |
| **L1** | 夹具合法性与 oracle | `soffice` → PDF、`pdfinfo`、`verify_review_corpus.py` | 页数基线、夹具有效性、损坏夹具修复 | ✅ v1（5 个夹具） |
| **L2** | 上游双跑 | React playground + 本地 demo + Playwright | 同文件页数/字数/卡片/图表/截图 | ✅ 3 个用例 |
| **L3** | 真实文件语料 | 公开 DOCX/XLSX/PPTX/PDF + oracle + 双跑 | 未知组合、异常结构、真实字体 | ✅ v1（18 个文件，发现 #152） |
| **L4** | 编辑→导出→重开回环 | 控制器脚本 + 文本/模型比较 | 数据丢失、schema 破坏、宏/格式丢失 | ✅ v1（DOCX PASS；XLSX 图表导出损坏 → #66） |
| **L5** | 横向矩阵 | Chromium/Firefox/WebKit、4 视口、明暗、缩放、IME/触摸/打印 | 浏览器/场景差异 | 🟡 V3：Chromium+Firefox 通过；WebKit 环境阻塞 |
| **L6** | 像素基线 | 截图 vs LibreOffice/pdftoppm，SSIM/容差 | 版式漂移、字体度量差异 | ✅ v1（DOCX 18.6%、PPTX 2.9%、PDF 18.1%） |

鲁棒性（畸形包、zip 炸弹、超大文件、并发加载）作为 L4 的旁支，单独一套用例（V5）。

## 4. 证据规范

```text
verification/evidence/2026-11-16/
├── corpus-audit.md                # 夹具特性/缺陷审计
├── render-observations.md         # 双跑、视口、评论对齐等人工可读记录
├── screenshots/                   # 关键截图（上游/本地成对）
├── office-oracle/summary.json     # L1 结果
└── upstream-parity/summary.json   # L2 结果
```

`summary.json` 至少包含：

```json
{
  "suite": "upstream-parity",
  "date": "2026-11-16",
  "commit": "<repo sha>",
  "environment": { "chromium": "155", "soffice": "26.8.1.1", "upstream": "react-docx@28a3bd9" },
  "cases": [
    { "id": "docx_demo", "upstream": {"pages": 9}, "local": {"pages": 6}, "verdict": "gap", "issues": [25, 99, 127] }
  ],
  "verdict": "PASS|FAIL|GAP",
  "artifacts": ["screenshots/up-docx.png", "screenshots/local-docx.png"]
}
```

证据要求：

- 每个 case 的夹具记录 `sha256` 与字节数；
- 截图命名 `<side>-<fixture>-<page>.png`；
- 不提交含敏感信息的真实文件，只记录哈希、页数与结果；
- 证据体积控制在单批次 ≤10 MB。

## 5. 执行入口

```bash
# L1 夹具校验 + oracle
uv run python scripts/fixtures/verify_review_corpus.py
uv run python tests/blackbox/office_oracle.py --output verification/evidence/2026-11-16/office-oracle/summary.json

# L2 双跑（需上游 playground 与本地 demo）
UPSTREAM_DOCX_URL=http://localhost:5181 \
UPSTREAM_XLSX_URL=http://localhost:5182 \
uv run --with playwright python tests/blackbox/upstream_parity.py --strict
```

上游 playground 的启动步骤见 `docs/testing/upstream-parity.md`。

## 6. 覆盖矩阵与收口标准

| 格式 | L0 | L1 | L2 | L3 | L4 | L5 | L6 |
|---|---|---|---|---|---|---|---|
| DOCX | ✅ | ✅ | ✅ | 🔜 | 🔜 | 🔜 | 🔜 |
| XLSX | ✅ | ✅ | ✅ | 🔜 | 🔜 | 🔜 | 🔜 |
| PPTX | ✅ | ✅ | —（无上游 UI 对照） | 🔜 | 🔜 | 🔜 | 🔜 |
| PDF | ✅ | ✅（pdftoppm） | —（EmbedPDF demo 待接） | 🔜 | —（只读） | 🔜 | 🔜 |

“issue 收集阶段”结束的门槛：

1. 每格式 ≥3 个真实文件 + ≥3 个合成夹具跑过 oracle；
2. 每格式有编辑→导出→重开的回环用例（PDF 除外）；
3. Chromium + Firefox + WebKit 各跑首屏/翻页/搜索；
4. 连续两轮完整矩阵没有**新的 issue 类别**；
5. 新发现要么归入现有 issue，要么明确标注“新类别”并说明扩展理由。

## 7. 路线图

| 版本 | 内容 | 主要产出 |
|---|---|---|
| **V1** | 真实文件语料 | `test-data/real/` + manifest + `real_files.py`；oracle/双跑差异 |
| **V2** | 编辑回环 | `roundtrip.py`：DOCX 编辑/导出/重开、XLSX 编辑/导出/重开 |
| **V3** | 跨浏览器 | Firefox/WebKit 首屏、翻页、搜索、剪贴板 |
| **V4** | 像素基线 | 截图 vs LibreOffice/pdftoppm，SSIM 阈值与容差表 |
| **V5** | 鲁棒性 | 畸形包、超大文件、并发加载、取消/超时 |
| **V6** | 场景矩阵 | 明暗、50/100/200% 缩放、IME、触摸、打印 |

## 8. 风险与边界

- 上游未 vendored：`/tmp` 安装的 playground 不随仓库保存；用 commit + 启动文档固化。
- LibreOffice 版本漂移会改变页数/分页；基线需随版本更新。
- 字体差异是像素对比的最大噪声源；L6 必须固定字体集。
- 截图与真实文件可能含敏感信息：只提交合成夹具的截图；真实文件只存哈希。
- 本框架不评价性能（另有 performance-baseline），不替代单元/组件测试。

## 9. 当前状态（2026-11-16）

- 142 个 open issue（65 bug、18 enhancement、2 product-gap、3 review、54 upstream-gap）。
- L1：16 个合成夹具，oracle PASS；`demo.docx` 记录 **9 页 vs 本地 6 页（gap 3）**。
- L2：3 个用例 PASS（`demo.docx`、`docx-comments.docx`、`xlsx-charts.xlsx`）。
- 本轮修复：`docx-comments.docx` 的页眉评论导致 LibreOffice 拒载；已改为正文/表格评论并补全 footnotes 注册。
- V1 真实文件：18 个 POI/PDFBox 文件，oracle PASS；viewer 16/18，两个 XLSX 因图表系列缺 `<c:tx>` 崩溃（#152）。
- V2 回环：DOCX PASS；XLSX 导出丢 chart part，openpyxl 无法重开（#66，P0）。
- V3 跨浏览器：Chromium+Firefox 通过；WebKit 因宿主库缺失 SKIP。
- V4 像素基线：3 个用例 PASS（阈值 20–30%）。
- 下一步：V5 鲁棒性。
