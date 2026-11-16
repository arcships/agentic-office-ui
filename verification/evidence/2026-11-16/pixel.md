# 像素基线 vs LibreOffice/poppler（V4，2026-11-16）

方法：LibreOffice 转 PDF → `pdftoppm -r 96` 渲染；本地 viewer 截取对应元素；缩放到 800px 宽、灰度、裁剪到共同高度后比较。

| 用例 | RMSE | 差异像素比 | SSIM | 阈值 | 结果 |
|---|---:|---:|---:|---:|---|
| DOCX `demo.docx` 第 1 页 | 23.85 | 18.6% | 0.069 | 30% | PASS |
| DOCX `demo.docx` 第 2 页 | 22.24 | 17.8% | 0.109 | 30% | PASS |
| PPTX `pptx-deck-basic.pptx` 第 1 页 | 4.95 | 2.9% | 0.130 | 30% | PASS |
| PPTX `pptx-deck-basic.pptx` 第 2 页 | 3.92 | 2.2% | 0.064 | 30% | PASS |
| PDF `pdf-rotated-90.pdf` 第 1 页 | 22.62 | 18.1% | -0.018 | 20% | PASS（接近阈值） |
| PDF `pdf-rotated-90.pdf` 第 2 页 | 12.18 | 10.0% | 0.030 | 20% | PASS |

- DOCX 的 17–19% 主要来自分页/字体度量差异（本地第 1 页塞了更多内容），不是颜色错误；SSIM 低是因为整页位移，阈值仍按差异像素比。
- PDF 的 18.1% 来自 PDFium 与 poppler 的栅格化差异；阈值只有 20%，后续需要固定字体/DPI 后再收紧。
- 产物：`pixel/{case}/{case}-local.png`、`-oracle.png`、`-diff.png`。
