# XLSX 页面设置参考（Excel）与 0.8 采用子集（#171）

> 批次：0.8.1 / issue [#171](https://github.com/arcships/agentic-office-ui/issues/171)。
> 关联：[#160](https://github.com/arcships/agentic-office-ui/issues/160) chrome 迁移、
> [design/viewer-editor-architecture.md](viewer-editor-architecture.md) §6.4–6.5。
> 上游 pinned `xlsx-editor.tsx` / `xlsx-viewer.tsx` 内**没有**可复用的页面设置对话框，
> 因此本文是 `office-reference` 规格：参考对象是 Excel 自己的页面设置，而不是上游像素。

## 1. 参考来源

| 来源 | 用来确认 |
|---|---|
| [Page Setup](https://support.microsoft.com/en-us/excel/page-setup) | 对话框四个选项卡（Page / Margins / Header-Footer / Sheet）里每一项的含义 |
| [Scale a worksheet](https://support.microsoft.com/en-us/excel/scale-a-worksheet) | `Scale to Fit` 组（Width/Height）、`Adjust to` 与 `Fit to` 的差别、`Fit to` 忽略手动分页线 |
| [Set or clear a print area](https://support.microsoft.com/en-gb/excel/set-or-clear-a-print-area-on-a-worksheet) | 打印区域是一组区域、每个区域单独成页、清空会移除全部区域 |
| [Keyboard shortcuts in Excel](https://support.microsoft.com/en-us/accessibility/excel/keyboard-shortcuts-in-excel) | `Page Layout` 选项卡承载 themes / page setup / scale / alignment |

## 2. Excel 的实际 UX（调查结论）
**入口**

- `Page Layout` → `Page Setup` 组的对话框启动器，打开**一个**多选项卡对话框；
  键盘为 `Alt+P, S, P`。
- `File` > `Print` 的打印预览里也有 `Page Setup` 链接，打开同一个对话框。
- 常用项在功能区直接可点：`Margins`、`Orientation`、`Size`、`Print Area`、`Breaks`、
  `Print Titles`、`Scale to Fit`（Width/Height）、`Sheet Options`（Gridlines / Headings
  各分 View 与 Print）。

**对话框四个选项卡**

| 选项卡 | 字段 | 备注 |
|---|---|---|
| Page | Orientation（Portrait/Landscape）、Scaling（`Adjust to` % / `Fit to` N 页宽 × M 页高）、Paper size、Print quality、First page number | `Fit to` 里宽度填 1、高度留空表示“按纸张宽度铺满，页数不限”；`Fit to` 会忽略手动分页线，且不会拉伸数据去填满纸张 |
| Margins | Top/Bottom/Left/Right、Header/Footer 距离、Center on page（Horizontally / Vertically） | 页眉页脚距离应小于对应页边距，否则会压到数据 |
| Header/Footer | 内置或自定义页眉页脚、奇偶页不同、首页不同、Scale with document、Align with page margins | 与 `headerFooter` 部件绑定 |
| Sheet | Print area、Print titles（Rows to repeat at top / Columns to repeat at left）、Print（Gridlines、Black and white、Draft quality、Row and column headings、Comments and notes、Cell errors as）、Page order（Down, then over / Over, then down） | 打印区域用区域选择器 + “折叠对话框”按钮取当前选区 |

**对话框的交互约定**（对本仓库成立的约束）

- 单一模态对话框，确认/取消成对出现；取消不写工作簿。Excel 的 OK 是**一次**提交，
  撤销一次就回到打开前的状态。
- 字段是从工作簿**当前值**初始化，不是从默认值初始化。
- 无工作表 / 只读时入口不可用。

## 3. 0.8 采用子集

实现 `XlsxPageSetupDialog.vue`：一个 `OfficeDialog`，字段是上表的**可写子集**。
每一项都对应 `XlsxSheetData` 上已解析的字段，读与写都走同一份 `xlsx-core` 类型。

| 字段 | 写回 OOXML | 采用理由 |
|---|---|---|
| Orientation（portrait/landscape） | `pageSetup/@orientation` | 打印预览的纸张轴交换直接依赖它 |
| Paper size（Letter / Legal / Executive / Tabloid / A3 / A4 / A5 / B4 / B5 + 当前值原样保留） | `pageSetup/@paperSize` | `PAPER_SIZES` 已经解析了完整 ST_PaperSize，子集只影响下拉长度 |
| Scaling：`Adjust to %`（10–400）或 `Fit to` 宽 × 高（0 = 自动） | `pageSetup/@scale`、`@fitToWidth`、`@fitToHeight`、`sheetPr/pageSetUpPr/@fitToPage` | 两类缩放互斥，是分页结果的主开关 |
| Margins：top/bottom/left/right/header/footer（英寸，0–10） | `pageMargins/@*` | 预览与 `@page` 尺寸都用它 |
| Center on page：horizontal/vertical | `printOptions/@horizontalCentered`、`@verticalCentered` | Margins 选项卡里的同名项 |
| Print：gridlines / row and column headings | `printOptions/@gridLines`、`@headings` | 打印预览已渲染这两项；Sheet 选项卡里的同名项 |
| Page order：Down, then over / Over, then down | `pageSetup/@pageOrder` | 多页输出顺序，分页器已实现 |
| Print area（可选）：区域输入框 + “使用当前选区” + “清除” | `xl/workbook.xml` 的 `_xlnm.Print_Area`（`localSheetId` 作用域） | Sheet 选项卡的核心项；空值表示整表打印 |

### 明确不做（附产品理由）

| 未做项 | 产品理由 |
|---|---|
| Header/Footer 选项卡（自定义页眉页脚、奇偶页/首页不同、Scale with document、Align with page margins） | 需要 `headerFooter` 部件与富文本编辑（Excel 的页眉页脚是带 `&L/&C/&R` 代码的编辑器）；0.8 的打印预览本身不渲染页眉页脚，先做写回等于“看不见也改不了”的字段 |
| Print quality（horizontalDpi / verticalDpi）、Copies、First page number（`useFirstPageNumber`） | 目标机器/打印机属性相关，浏览器 `window.print()` 不接受这些参数；写回只会在 Excel 里生效，但用户在我们这里改看不到任何结果 |
| Black and white、Draft quality、Cell comments、Cell errors as | 纯打印驱动/输出渲染项；我们的打印预览没有黑白/草稿渲染，也不打印批注与错误占位 |
| Print titles（Rows to repeat at top / Columns to repeat at left） | 已有解析与分页支持，但写入需要第二套区域选择器 UI；与打印区域分开评估，见 §5 |
| 纸张方向/尺寸的实时拖动预览图（Excel 的 Preview box） | 我们的预览就是 `XlsxPrintPreview` 全页渲染；对话框内再画一个缩略预览会与真实分页结果重复 |

## 4. 交互与验收（office-reference）

- 打开：`Page Layout` 等价入口（本仓库 `XlsxRibbon` 的 authoring 组“页面设置”按钮）→
  `XlsxPageSetupDialog`；字段由当前工作表已解析值初始化
  （`tests/component/xlsx-page-setup-dialog.test.mjs` 断言方向/纸张/缩放/边距/打印区域均来自工作簿）。
- Esc / 关闭按钮 / backdrop 关闭等同取消：不写工作簿、不改撤销栈，焦点回到触发按钮；
  三个关闭路径各有断言。
- 确认（OK）：**一次** `controller.setPageSetup(input)`，写工作簿、记一条撤销记录、
  随 `exportXlsx()` 导出；`undo()` 一次回到打开前
  （`tests/component/xlsx-page-setup-authoring.test.mjs` 断言模型、导出 XML 与撤销）。
- 只读、受保护工作表、Worker 只读大文件：表单控件与 OK 禁用并显示原因；控制器侧返回 `false` 且不写。
- 390px 视口：对话框宽度沿用 `.xlsx-authoring` 的 `min(92vw, 520px)`，字段网格为
  `repeat(auto-fit, minmax(120px, 1fr))`，源码级断言拒绝固定 `px` 宽度与 `min-width`，
  因此不会横向溢出。**浏览器截图证据（`--evidence-only` 捕获）仍需在有 Playwright 的机器上补，
  与本 issue 的组件验收分开记录。**
- 键盘：Tab 顺序按字段在 DOM 里的顺序，所有控件有可见 label；数值输入有 `min`/`max`，
  超界值在提交时收敛而不是抛错。
- 摘要行：对话框内每次修改都跑一次核心分页，显示实际页数/纸张/方向/缩放；
  打印预览 (`XlsxPrintPreview`) 在应用后渲染同一份结果（组件测试断言 `A4 landscape` 与纸张轴交换）。

## 5. 后续（不在 #171）

- Print titles 编辑器（`_xlnm.Print_Titles`）：复用同一套区域输入与撤销路径，单独 issue。
- Header/Footer 编辑器：取决于打印预览是否先渲染页眉页脚。
- 页面设置的 L1 截图对照：参考为 Excel 对话框截图（`office-reference`），不是上游组件
  像素；在引入浏览器证据前，由组件测试 + 390px 断言覆盖。
