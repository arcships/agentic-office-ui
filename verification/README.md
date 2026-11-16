# Verification branch

本分支（`verification`）保存“发现问题”的全部验证资产与证据，不包含产品修复。

## 目录

```text
verification/
├── README.md
└── evidence/
    └── 2026-11-16/
        ├── corpus-audit.md
        ├── render-observations.md
        ├── screenshots/
        ├── office-oracle/summary.json
        └── upstream-parity/summary.json
```

## 资产位置

| 资产 | 路径 |
|---|---|
| 合成夹具生成器 | `scripts/fixtures/generate_review_corpus.py` |
| 夹具校验器 | `scripts/fixtures/verify_review_corpus.py` |
| 夹具文件 | `test-data/corpus/` |
| LibreOffice oracle | `tests/blackbox/office_oracle.py` + `office_oracle_baseline.json` |
| 上游双跑 | `tests/blackbox/upstream_parity.py` + `upstream_parity_baseline.json` |
| 框架说明 | `docs/testing/verification-framework.md` |
| 双跑启动步骤 | `docs/testing/upstream-parity.md` |

## 复现

```bash
# 夹具 + oracle
uv run python scripts/fixtures/verify_review_corpus.py
uv run python tests/blackbox/office_oracle.py --output verification/evidence/<date>/office-oracle/summary.json

# 双跑（需上游 playground 与本地 demo）
UPSTREAM_DOCX_URL=http://localhost:5181 UPSTREAM_XLSX_URL=http://localhost:5182 \
uv run --with playwright python tests/blackbox/upstream_parity.py --strict
```

## 规则

- 只改 `tests/`、`scripts/fixtures/`、`test-data/`、`docs/`、`verification/`；不改 `packages/` 或 `apps/demo/src`。
- 证据命名 `<YYYY-MM-DD>/<suite>/`，每个 case 记录夹具哈希、环境版本、上游/本地指标与 issue 归属。
- 真实文件只提交哈希与结果，不提交可能含敏感信息的原文。

## 进度（V1–V6）

| 阶段 | 状态 | 证据 |
|---|---|---|
| L0 静态/源码/API | ✅ | #76 / 源码仓核对 |
| L1 夹具 + oracle | ✅ | `evidence/2026-11-16/office-oracle/` |
| L2 上游双跑 | ✅ | `evidence/2026-11-16/upstream-parity/` |
| V1 真实文件 | ✅ 18 个 | `evidence/2026-11-16/real-files/` |
| V2 编辑回环 | ✅ DOCX；XLSX 导出损坏 #66 | `evidence/2026-11-16/roundtrip/` |
| V3 跨浏览器 | ✅ Chromium+Firefox+系统 WebKitGTK | `evidence/2026-11-16/cross-browser/` |
| V4 像素基线 | ✅ 3 用例 | `evidence/2026-11-16/pixel/` |
| V5 鲁棒性 | ✅ 7 用例 | `evidence/2026-11-16/robustness/` |
| V6 场景矩阵 | ✅ 16 视口组合 + 场景 | `evidence/2026-11-16/scenario/` |
| 第二轮全量 | ✅ 稳定，无新类别 | `evidence/2026-11-16/final-pass.md` |

本轮新发现（已提 issue）：#152（图表系列缺 `<c:tx>` 导致整包无法加载）；#66 补充导出丢 chart part 的实测证据。
