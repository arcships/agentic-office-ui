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
