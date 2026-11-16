# 场景矩阵（V6，2026-11-16）

Chromium 155，本地 demo dist。

## 视口矩阵

4 视口（1440×900 / 1280×720 / 768×1024 / 390×844）× 4 查看器（DOCX/XLSX/PDF/PPTX）：

- 16/16 无页面级横向溢出；
- 无 pageerror / console error；
- 各查看器均到达可交互状态。

## DOCX 场景

| 场景 | 观测 | 关联 |
|---|---|---|
| 100% 基线 | 6 页，首页 1525px | #25/#99/#95 |
| 200% | 首页 3051px（≈2×） | 缩放正确 |
| 50% | 首页 763px（≈0.5×） | 缩放正确 |
| 暗色 | 纸张 `rgb(255,255,255)` → `rgb(31,41,55)`，文字转浅色 | 正常 |
| 打印媒体 | 工具栏仍 `display:flex`，仅挂载 1 页 | **#92 无打印样式** |
| PageDown | `scrollTop` 0 → 0 | **#116 无翻页快捷键** |

## XLSX / PPTX 缩放

| 场景 | 观测 | 关联 |
|---|---|---|
| XLSX 200% / 50% | select 值正确变化；canvas 内部分辨率保持 1358×476（仅显示缩放，未重绘像素） | 视觉缩放未量化 |
| PPTX 100% → 150% | 两次“＋”后按钮显示 150%，舞台 1392×710 | 正常 |

## IME

- **DOCX**：编辑器内 `compositionstart` → 输入 → `Enter`，段落数 15 → 15，**守卫生效**（`isComposing` 拦截）。
- **XLSX**：单元格编辑中 `compositionstart` → 输入 → `Enter`，**编辑器关闭（提交）**，即输入法选字回车会提交半成品；与 #89 一致（无 `composition` 守卫）。

## 触摸

390×844、`has_touch` 上下文：tap 后 6 页正常，无 pageerror；未验证双指捏合（#106）。

## 跨引擎场景（Chromium / Firefox / WebKit）

同一组场景在三个引擎上复跑：

| 场景 | Chromium | Firefox | WebKitGTK |
|---|---|---|---|
| DOCX 缩放 200/50 | 3051 / 763 | 3051 / 763 | 3049 / 762 |
| DOCX 暗色 | rgb(31,41,55) | 同 | 同 |
| DOCX IME 守卫 | 15→15 | 15→15 | —（WebDriver 合成事件不可靠，未测） |
| XLSX 缩放 200/50 | select 200/50，canvas 1358×476 | 同 | select 200/50，canvas 1350×945 |
| PPTX 缩放 100→150 | ✅ | ✅ | ✅ |
| PDF 页数 | 4 | 4 | 4 |

打印媒体：Chromium/Firefox 已测（工具栏仍显示、仅挂载 1 页）；WebKitWebDriver 不支持媒体模拟，记录为工具限制。
PDF/XLSX/PPTX 的暗色开关在 demo 中未暴露，记为产品缺口（#92）。
