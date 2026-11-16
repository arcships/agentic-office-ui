# 上游双跑对比（opt-in）

对比 **上游 Extend playground** 与本地 demo 在**同一份文件**上的行为差异。它不是 CI 的默认步骤，因为需要上游仓库的 playground 在本地运行。

## 启动

```bash
# 1. 上游 DOCX（只做一次：pnpm install，Node 24 + pnpm 9.15）
cd /tmp/react-docx
corepack pnpm@9.15.0 install --frozen-lockfile --config.engine-strict=false
corepack pnpm@9.15.0 --config.engine-strict=false --filter @extend-ai/react-docx build
cd apps/playground && ./node_modules/.bin/vite --port 5181 --strictPort

# 2. 上游 XLSX（注意先复制 duke wasm，否则会误报 Could not load workbook）
cd /tmp/react-xlsx
corepack pnpm@9.15.0 install --frozen-lockfile --config.engine-strict=false
corepack pnpm@9.15.0 --config.engine-strict=false --filter @extend-ai/react-xlsx build
cd apps/playground && node scripts/copy-duke-wasm.mjs && ./node_modules/.bin/vite --port 5182 --strictPort

# 3. 本地 demo
cd apps/demo/dist && python3 -m http.server 4173
```

## 运行

```bash
UPSTREAM_DOCX_URL=http://localhost:5181 \
UPSTREAM_XLSX_URL=http://localhost:5182 \
LOCAL_DEMO_URL=http://127.0.0.1:4173 \
uv run --with playwright python tests/blackbox/upstream_parity.py --strict
```

- 不设置 `UPSTREAM_*` 时打印 `SKIP` 并退出 0。
- `--strict` 会把本地指标与 `tests/blackbox/upstream_parity_baseline.json` 比较，漂移即失败（回归守卫）。
- 结果写入 `output/acceptance/upstream-parity/summary.json`。

浏览器默认用系统 `/usr/bin/chromium`（可用 `UPSTREAM_PARITY_CHROMIUM` 覆盖），无需 `playwright install`。

## 当前基线（2026-11-16）

| 用例 | 上游 | 本地 | 说明 |
|---|---|---|---|
| `demo.docx` | 9 页 / 486 words | 6 页 | 差距与 #127 行内度量、#129 段间距/制表位、#47 行高相关 |
| `docx-comments.docx` | 1 页 / 32 words / 0 卡片 | 1 页 / 3 卡片 | 上游默认“最终视图”；本地 inline + gutter 并存（#149） |
| `xlsx-charts.xlsx` | 6 图表 | 6 图表 | 上游缺分类轴刻度，本地轴完整 |

截图由临时脚本保存在 `/tmp/compare/`（不进入仓库）。

## 维护

- 上游仓库更新后，按固定 commit 重新记录基线，并在 `docs/upstream-extend-ui.md` 更新引用。
- 本地修复导致指标变化时，更新 baseline，并在 PR/issue 里注明“向 9 页靠拢”或“卡片数变化”的原因。

## 独立 oracle（LibreOffice）

上游对照之外，另有一条不依赖上游的基准：

```bash
uv run python tests/blackbox/office_oracle.py
```

- 用 headless LibreOffice 把 DOCX/XLSX/PPTX 夹具转成 PDF，再用 `pdfinfo` 断言页数；
- 基线在 `tests/blackbox/office_oracle_baseline.json`，同时记录本地 viewer 的页数，差距会直接打印；
- 缺少 `soffice`/`pdfinfo` 时自动 SKIP。

当前结果（2026-11-16）：

| 夹具 | LibreOffice | 本地 viewer | 差距 |
|---|---:|---:|---:|
| `demo.docx` | 9 | 6 | **3 页** |
| `docx-comments.docx` | 1 | 1 | 0 |
| `docx-merged-table.docx` | 1 | 1 | 0 |
| `pptx-deck-basic.pptx` | 5 | 5 | 0 |
| `xlsx-charts.xlsx` | 4（打印页） | — | — |

说明：`demo.docx` 与 `legal-contract.docx` 是同一文件（md5 相同）。
