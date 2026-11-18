# XLSX 使用指南

`XlsxViewer` 接收一个 `XlsxViewerController`。控制器负责文件、工作表、选区、编辑命令和历史记录；组件负责界面。

支持 `.xlsx`、`.xls`、`.xlsb`、`.xlsm`、`.xltx`、`.xltm` 和 `.csv`。XLSB 会按需转换为内存 XLSX 后进入现有解析链路；宏工作簿和宏模板只读取表格内容，不执行 VBA；模板按普通工作簿打开。

```vue
<script setup lang="ts">
import { useXlsxViewerController, XlsxViewer } from "@arcships/vue-xlsx"
import "@arcships/vue-xlsx/style.css"

const props = defineProps<{ bytes: ArrayBuffer }>()
const controller = useXlsxViewerController({
  file: props.bytes,
  fileName: "workbook.xlsx",
})
</script>

<template>
  <XlsxViewer
    :controller="controller"
    height="720px"
    :read-only="false"
    @cell-double-click="onCellDoubleClick"
  />
</template>
```

常用属性：

| 属性 | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `controller` | `XlsxViewerController` | 必填 | 查看器控制器 |
| `height` | `string` | `100%` | 查看器高度 |
| `minHeight` | `string \| number` | 无 | 可选最小高度（数字按 px）；不传时组件不设置最小高度 |
| `readOnly` | `boolean` | `false` | 禁止编辑 |
| `showDefaultToolbar` | `boolean` | `true` | 默认工具栏 |
| `showRibbon` | `boolean` | `true` | 功能区 |
| `showFormulaBar` | `boolean` | `true` | 公式栏 |
| `showImages` | `boolean` | `true` | 图片和绘图对象 |
| `showSheetTabs` | `boolean` | `true` | 底部工作表标签条 |
| `activeTabIndex` | `number` | 无 | 受控的活动标签下标；不传时由控制器持有 |
| `showUpload` | `boolean` | `false` | 上传入口 |
| `getCellStyle` | `function` | 无 | 按单元格补充样式 |
| `locale` | `string` | `"en"` | 界面文案与数字/日期默认区域；内置 `"en"`、`"zh-CN"` |
| `documentLocale` | `string` | 跟随 `locale` | 单独覆盖数字/日期区域 |

事件包括 `cellDoubleClick`、`upload`、`update:readOnly`、`tabChange`、`update:activeTabIndex`、`load-deferred`、`deferred-resumed` 和 `deferred-cancel`。打印时 `@page` 尺寸取自活动工作表的 `pageSetup` 纸张与方向；图表工作表使用 16:9 的自适应高度。

## 单元格编辑

双击单元格、按 `Enter` 或 `F2` 打开内联编辑器。编辑器是 `textarea`，可以输入多行：

- `Enter` 提交；`Alt+Enter`（macOS 为 `Option+Enter`）在光标处插入硬换行，编辑器保持打开。
- `Escape` 放弃草稿，单元格保持原值；`Tab` 提交并移到右侧单元格继续编辑。
- 草稿没有硬换行时，`↑`/`↓` 提交并上下移动选区；有硬换行时方向键在草稿的行间移动光标。
- 粘贴多行文本会保留换行（`CRLF` 统一为 `LF`）；单行粘贴仍走浏览器原生路径。
- 输入法组字期间的 `Enter`/`Escape` 由输入法处理，组字结束后只提交一次。

编辑器用一个隐藏的镜像层按当前宽度测量草稿的换行高度：硬换行或自动换行后编辑器向下增高（裁剪在网格可视区内），不会把后续行藏在单行滚动里。提交后的换行会写进单元格值；要让换行在单元格里显示为多行，请启用“自动换行”并给行足够的高度——未启用时与 Excel 一致，只显示第一行。

## 填充柄与自动列宽

选区右下角会显示 Excel 的填充柄（`data-testid="xlsx-fill-handle"`）：

- 按住填充柄拖动时，网格用灰色虚线预览将要覆盖的区域（根节点的 `data-fill-target` 同步给出该区域的 A1 地址），松开后调用控制器的 `fillSelection(targetRange)`：源区域按 `offset % sourceSize` 平铺到目标区域，公式按“目标 − 来源”的距离改写相对引用（`$` 绝对部分不变），样式一并复制。整次填充只写一条撤销记录，一次撤销即可恢复，选区随后扩展到填充后的范围。
- 拖动过程只预览、不修改单元格；指针在源区域内松开，或收到 `pointercancel`（系统手势打断等），都不会写入任何内容。只读或受保护的工作表不显示填充柄，拖动也不会生效。
- 填充柄跟随选区、滚动、冻结窗格和缩放定位：目标单元格右下角滚出可视区时直接隐藏，而不是贴在视口边缘。拖动语义由 `packages/vue-xlsx/src/components/grid-fill-handle.ts` 的 `resolveFillTargetRange()` / `normalizeFillRange()` / `fillRangesEqual()` 实现，都是无副作用的纯函数，自建渲染器可沿用同一套规则。

双击列头或行头的边界会把该列/该行调整到内容大小（拖动边界仍是手工调整）：

- 列宽用绘制时同一套字体测量该列每个单元格的格式化显示值（数字和日期按当前区域格式化），取最宽的一行文本加上 5px 单元格内边距和 2px 余量；富文本按每段自己的字体累加，旋转文字按旋转后的包围盒计算，所以 90° 竖排的列头大约只需要一行行高那么宽。空列保持原宽度，合并区域内的单元格不参与测量。
- 行高用该行每个单元格在当前列宽下的换行行数计算（与绘制时相同的贪心换行），取最高的一行文本加上 4px 垂直留白，因此自动换行的长文本不会被截断。
- 结果通过 `controller.resizeColumn()` / `controller.resizeRow()` 写入，与拖动边界完全同一条路径：只写一条撤销记录、遵守只读与工作表保护规则，并限制在列 32–1800px、行 18–546px 之内。测量本身有每轴 20 000 格扫描上限，格式化到很远的单元格不会让双击卡住。无 DOM 的宿主（SSR、测试渲染器）没有 Canvas 可测量时改用 `estimateAutofitTextWidth()` 的字宽估算；列头双击命中判定仍复用既有的 `hitTestAxis()`，所以边界 5px 内才是自动调整，点在列头中间不会误触发。

## 自动筛选与列头菜单

工作表自己的 `<autoFilter ref="…">` 和每个表部件（`xl/tables/tableN.xml`）的 `<autoFilter>` 会被解析成 `controller.autoFilterRanges`；同一区域的表筛选会被优先保留（带 `tableName`），所以表和工作表不会画出两个菜单。范围里的每个表头单元格都会显示一个常驻的 `▾`，不依赖鼠标悬停；滚动、冻结和缩放时会跟随单元格位置，滚出可视区就隐藏。

- 菜单提供“按此列升序/降序”“按值筛选”和“清除筛选”。值列表读取该列的格式化显示值（数字、日期按当前区域显示），空白单元格单独显示为“(空白)”，`全选` 控制全部值；`应用` 写入筛选项，`清除筛选` 恢复该范围隐藏的所有行。来自文件的筛选条件会在打开工作簿时直接生效，和 Excel 一致。
- 排序是一次普通的范围编辑：整行一起移动，只写一条撤销记录，`撤销` 恢复原顺序；只读或受保护的工作表不会排序。排序前会检查合并区域，数据区内有合并时拒绝排序（返回 `false`），避免把行移走而合并框留在原地。
- 筛选是视图状态：不修改单元格、不删除合并区域、不改冻结窗格，也不影响打印分页和导出内容。`controller.hiddenAutoFilterRows` 给出被隐藏的绝对行号，`controller.autoFilterStates` 给出每个范围的列条件；`applyAutoFilterValues(reference, columnIndex, values, { includeBlank })`、`clearAutoFilter(reference?)` 和 `sortAutoFilter(reference, columnIndex, direction)` 是同一套操作的程序入口。
- 导出时会把当前“应用/清除”的条件写回工作表或表部件的 `autoFilter`，重新打开后筛选状态保持一致；没有筛选操作时导出字节不变。只读 Worker 视图由 Worker 返回范围显示值，菜单和筛选同样可用（排序仍是只读禁止）。
- 宿主可以替换内置 `▾` 触发器：`XlsxViewer`/`XlsxSheetSurface`/`XlsxGrid` 的 `#table-header-menu` 插槽接收 `XlsxTableHeaderMenuRenderProps`（`cell`、`column`、`table`、`autoFilter`、`direction`、`sortAscending`、`sortDescending`、`getValueOptions`、`applyFilter`、`clearFilter`、`hasFilter`、`triggerIcon` 和必须绑定的 `triggerProps`）。菜单本体仍由查看器提供，排序、筛选、清除和撤销走同一条路径。

## 公式引用高亮

编辑公式时，内联编辑器和公式栏都会解析光标处的公式文本，并用 Excel 的引用颜色标出目标：

- 支持单元格（`B2`、`$B$2`）、区域（`A1:A3`）、整列（`A:A`）、整行（`3:3`）、跨表引用（`Sheet2!A1`、`'My Sheet'!$A$1:$B$2`）以及名称管理器中已定义的名称（`=SUM(Total)`）。字符串字面量、函数名、外部工作簿引用和无法解析的名称保持普通文本。
- 每个不同引用按固定调色板着色，重复引用沿用第一次的颜色；光标所在的引用在公式栏加下划线，网格里同时用同色边框描出该区域，活动引用的边框更粗。
- 跨表引用只在公式栏着色；网格只描当前工作表上的目标。
- 移动光标会立即切换活动引用；`Enter` 提交、`Escape` 取消、失焦或关闭编辑器后颜色和边框一起清除。整个高亮过程只读，不会修改工作簿、撤销记录或导出内容。
- `XlsxSheetSurface` 的内联编辑器同样生效；使用 `XlsxViewer` 时网格与公式栏共享同一个编辑会话。

自建公式栏可以复用导出的规则：`collectFormulaReferences(formula, caret, { activeSheetName, resolveName })` 返回引用、目标区域和调色板颜色，`formulaReferenceSegments()` 把公式切成普通/引用文本段，`FORMULA_REFERENCE_COLORS` 是 Excel 调色板；`useFormulaReferenceEditing(controller)` 让自定义输入框驱动内置网格的边框。

## 图表数据源与范围高亮

选中图表、系列、数据点或图例项时，查看器会解析图表 XML 中的分类轴、数值轴和气泡大小引用，并在网格上框出对应的源单元格：

- 分类（category）、数值（value）和气泡大小（bubbleSize）各用一种固定颜色，图例项与它的数值区域颜色一致；同一分类区域被多个系列共用时只画一次。
- 选中单个系列时只高亮该系列的分类和数值区域；选中数据点时收窄到该点自己的分类、数值单元格；选中整个图表时高亮所有系列。
- 支持 `Sheet1!$B$2:$B$5`、带引号的工作表名、`$` 绝对引用、单格、区域、整行/整列以及名称管理器中定义的名称。外部工作簿引用、联合区域和无法解析的名称不会画到错误的位置。
- 高亮跟随滚动和缩放；冻结窗格按行/列拆成多个片段并裁剪在各自窗格内；隐藏行/列会跳过，整行/整列引用只绘制屏幕可见部分。
- 取消选择图表后自动清除；切换工作表时只显示当前工作表上的区域。

宿主也可以用控制器主动高亮任意区域：

```ts
controller.highlightRanges(
  [{ start: { row: 4, col: 0 }, end: { row: 5, col: 1 } }],
  { color: "#f59e0b", label: "搜索结果" },
)
controller.highlightedRanges // 响应式：图表数据源 + 宿主区域
controller.clearHighlightedRanges()
```

`highlightRanges()` 会替换上一次宿主传入的区域，默认落在活动工作表、使用琥珀色；它只读，不修改工作簿、不写撤销记录、不影响导出。`clearHighlightedRanges()` 清除当前所有高亮；图表仍保持选中，下一次选择变化后会重新显示其数据源。`XlsxRangeHighlightLayer` 由 `XlsxViewer` 和 `XlsxSheetSurface` 自动挂载在画布之上、图表/图片之下；自建界面可以用 `createRangeHighlightLayout()` 复用同一套冻结窗格几何，或用 `collectChartRangeHighlights()` 和 `CHART_SOURCE_HIGHLIGHT_COLOR_BY_ROLE` 自行绘制。

## 图表内嵌图片与叠加形状

图表部件可以自带图片填充和 `chartUserShapes` 叠加层，查看器会把两者都画出来：

- 图表的 `a:blipFill` 图片填充（整张图表区、绘图区、系列、单个数据点）按图片原样拉伸铺满对应图形，柱形、扇区、面积和数据点标记都适用；图例色块显示同一张图片。
- `chartUserShapes` 里的文本框、标注和图片按 Excel 的 z 序叠加在绘图区之上：预设几何（`roundRect`、`ellipse`、`triangle`、`chevron`、箭头、星形等，未实现的预设回退为矩形）、`avLst` 调整值、纯色/图片填充、描边宽度与虚线、段落对齐、`anchor` 垂直对齐、`lIns`/`tIns` 边距以及逐段字体（字号、粗体、斜体、下划线、颜色、字体族）都会保留，长文本按文本框宽度换行并裁剪在形状内。
- 锚点坐标是相对图表区域的，所以缩放、调整图表大小、冻结窗格滚动时叠加层都跟随图表一起缩放；形状不参与命中测试，选择与右键菜单仍由图表本身处理。
- 嵌入的图片是从包内字节按签名（PNG/JPEG/GIF/WebP/BMP/TIFF）生成 data URL 的，外部链接、SVG、伪装成图片的 HTML 一律不画；单张 72 MB、单图表 100 MB、`chartUserShapes` 8 MB 的上限可防止异常包把内存拉满。宿主也可以读取 `chart.images`（关系 id → data URL）、`chart.drawings`（含 `shapes`）以及 `chart.chartAreaFillImage` / `chart.plotAreaFillImage` / `series.fillImage` / `dataPointStyles[].fillImage` 自行绘制。
- 导出/重新打开时图表部件、`xl/charts/chartUserShapes*.xml` 及其媒体通过绘图关系图一起保留，重新打开后图片填充和叠加层仍在。

## 内置图表样式

经典图表 XML 的 `<c:style>`（`chartStyle` 1–48，或 Excel 2010+ 的 101–148）现在会解析为 Excel 自带的系列调色板，而不是一律回退到主题的 accent1–6：

- `style 1/9` 使用主题深色的 6 级灰度，`3–8` 使用 accent1–6 的单色明暗阶梯，`41` 使用 6 级深灰，`2/42/102/142` 等彩色变体按系列循环 accent1–6。系列较多时按每 6 个一组追加亮度/阴影，计算在 DrawingML 的线性 RGB 空间完成。
- 显式 `chartColorStyle` 部件或系列 `spPr` 填充仍然优先；只有没有自定义颜色样式且图表带经典 `style` 时才使用内置调色板。带 `fillImage` 的系列/数据点仍保留图片填充，不会被内置颜色覆盖。
- 三维曲面图继续使用专用曲面调色板，同时从同一套内置样式取得墙面/底面填充和主网格线颜色；`chartStyleId` 为 140–148 的深色变体保留深色图表区与浅色文字。
- `XlsxChartAxis.majorGridlineColor` 会记录网格线颜色（作者 XML 中显式设置时优先），自定义渲染层可以直接使用。

## 形状预设几何与连接线

工作表绘图里的 `a:prstGeom` 不再只区分矩形和椭圆：

- `xlsx-core` 内置了 ECMA-376 的预设几何表（箭头、V 形、星形、标注、流程图、括号/花括号、月牙、齿轮、云、弧形等 150+ 个 `prst`）。`buildPresetShapePath(geometry, adjustments, { width, height })` 返回 `{ path, viewBox }`，路径坐标与形状的像素尺寸一致；`a:avLst` 的 `gd` 调整值按 DrawingML 的 1/100000 单位参与计算。`rect`/`ellipse` 等原生图元仍由 SVG 基本元素绘制。
- 解析后的 `XlsxShape` 会带上 `svgPath` 与 `svgViewBox`；`XlsxDrawingLayer` 在渲染时按当前矩形重新计算预设路径，所以调整形状大小时箭头、星形和流程图形状会像 Excel 一样重新展开，而不是把路径拉伸成矩形。未识别的 `prst` 才回退到矩形。
- 连接线（`cxnSp`，`bentConnector3`、`curvedConnector4`、`straightConnector1`、`line` 等）是开放路径，始终不填充；`a:ln` 的 `w`、`a:prstDash`（`dash`、`dot`、`lgDash`、`sysDash` 等）和 `a:headEnd` / `a:tailEnd`（`triangle`、`stealth`、`arrow`、`diamond`、`oval`）会渲染成 SVG 虚线、线宽和端点箭头。
- 标注（`wedgeRectCallout`、`callout1`/`callout2`/`callout3` 及其 `accent*` 变体）保留矩形主体和引线；引线可以落在形状矩形之外，绘制层不会把它裁掉。

```ts
import { buildPresetShapePath, isStrokeOnlyShapeGeometry } from "@arcships/xlsx-core"

const arrow = buildPresetShapePath("rightArrow", { adj1: 60000 }, { width: 160, height: 80 })
// arrow.path / arrow.viewBox 可直接用于 <path d={...} viewBox={...} />
isStrokeOnlyShapeGeometry("bentConnector3") // true：连接线不填充
```

预设几何表同样适用于自建渲染器：`XlsxShape.geometry` 是原始 `prst` 名称，`geometryAdjustments`、`svgPath` 和 `svgViewBox` 保留在形状模型里。

## 数字、日期与语言区域

查看器的 `locale` 同时决定界面文案和数字/日期的默认区域。数字格式里用 `[$-407]` 这类段显式指定的区域始终优先；没有指定区域的格式跟随查看器区域。需要单独控制数字/日期时，用 `documentLocale` 覆盖 `locale`；使用 `XlsxSheetSurface` 时同样可用这两个属性，或直接在 `useXlsxViewerController({ locale })` 中配置。

```vue
<XlsxViewer :controller="controller" locale="de-DE" />
```

- `locale="de-DE"`：键入 `1.234,5` 得到 1234.5（General 单元格会建议 `#,##0.00`），显示为 `1.234,50`；`31.12.2024` 解析为日期并按 `dd.mm.yyyy` 显示。`en-US` 下 `1,234.5` 保持 `1,234.50`。
- 显式 `@` 文本格式不参与数字转换：`1.234,5` 仍保存为文本。带引号/转义的格式字面量（例如 `"kg"#,##0.00`）保持原样，日期和时间格式的点、冒号不会被替换成数字分隔符。
- 控制器通过 `controller.locale` 报告当前区域，`controller.setDocumentLocale(tag)` 可在运行时切换并重绘；非法标签回退为 `en-US`。

## 富文本单元格

单元格内部可以有多段各自带格式的文字（例如加粗前缀、红色后半句、上标脚注）。查看器保留这些段落，而不是把它们压成一段纯文本：

- 网格画布按段绘制，每段使用自己的字体、字号、颜色、下划线/删除线和上标/下标偏移；自动换行和截断按各段实际宽度累加，段里的小号上标不会让文字提前折行或被错误截断。
- 打印预览（DOM 路径）为每段输出一个 `<span>`，行内样式与画布使用同一套规则，所以 DOM 与缩略图的断行位置一致。
- 工作表标签的缩略图同样按段绘制（粗体、颜色、上下标在缩略图里仍然可见）。
- 单元格的纯文本值始终是各段拼接，公式栏、搜索（`findCells`）、复制和导出继续使用它。

读取段落：

```ts
const runs = controller.getCellRuns({ row: 0, col: 0 })
// [{ text: "Quarter", font: { bold: true } }, { text: "ly", font: { color: { rgb: "#ff0000" } } }, ...]
```

`controller.activeSheet.richTextCells` 以 A1 地址保存整张工作表的段落；`getCellRuns` 会在单元格的值被改写（编辑、粘贴、重新计算）后自动不再返回旧段落，因此宿主不需要自己比对。需要自己渲染时，用 `@arcships/xlsx-core` 的 `resolveRichTextRuns(runs, value)`、`richTextRunsToPlainText(runs)` 做同样的“文本仍匹配”校验。导出的 `.xlsx` 会保留段落格式，重新打开后仍按段渲染。

## 文字旋转与缩小适应

单元格的 `textRotation` 和 `shrinkToFit` 会参与实际绘制，而不只是保留在样式里：

- Excel 的 0–90° 表示逆时针旋转，91–180° 表示顺时针旋转；查看器会把 45° 画成 −45°、90° 画成 −90°、180° 画成 +90°（自上而下）。旋转后的文字按旋转后的包围盒定位，90° 时占用的宽高互换；文字被裁剪在单元格矩形内，不会覆盖相邻单元格，也不会改变点击命中的单元格。
- 启用“缩小以适应”时，画布会从当前字号向下二分查找，直到文字宽度放进可用宽度；最小字号是基础字号的 0.4 倍。富文本单元格会按同一比例缩放每段字体，普通单元格保完整文本，不会先用省略号截断。
- 网格、工作表标签缩略图和打印预览共用同一套旋转/缩放规则；缩放比例（例如 100% 与 200%）只影响基础字号，0.4 倍下限和旋转角度保持 Excel 语义。
- `controller.activeSheet.cellTextLayouts` 以 A1 地址列出显式样式里需要旋转或缩小适应的单元格；行/列样式仍按原有继承规则生效。自建渲染器可复用 `packages/vue-xlsx/src/components/grid-paint.ts` 里的同一套纯函数（`resolveSpreadsheetTextRotation()`、`resolveCellTextLayout()`、`resolveRotatedTextBounds()`、`resolveRotatedTextCenter()`、`resolveRotatedTextFitWidth()`、`resolveShrinkToFitFontSize()`）；它们只接收尺寸和测量回调，不依赖 Canvas。

## 工作表标签

`showSheetTabs` 设为 `false` 只移除底部标签条：网格、功能区、公式栏和控制器状态都不变，`Ctrl/Cmd+PageUp/PageDown` 仍然可以切换工作表。此时切换只能通过宿主自己的界面调用 `controller.setActiveTabIndex(index)`。需要完全自定义的导航界面时，直接用 `XlsxSheetSurface` + `useXlsxViewerController`（见[自定义 XLSX 界面](../custom-components/xlsx.md)）。

需要宿主记录或恢复活动工作表时，用 `v-model:active-tab-index` 加 `tabChange`：

```vue
<script setup lang="ts">
import { ref } from "vue"
import { XlsxViewer, type XlsxWorkbookTab } from "@arcships/vue-xlsx"

const activeTab = ref(0)

function onTabChange(tab: XlsxWorkbookTab) {
  // kind 区分工作表与图表工作表；index 是 tabs 下标
  console.log(tab.kind, tab.name, tab.index, tab.sheetIndex, tab.chartsheetIndex)
}
</script>

<template>
  <XlsxViewer
    v-model:active-tab-index="activeTab"
    :controller="controller"
    @tab-change="onTabChange"
  />
</template>
```

- `activeTabIndex` 是 `controller.tabs` 的下标。工作表和图表工作表共用同一条标签条，所以它**不是** `controller.sheets` 的下标；越界值会被忽略，与 `setActiveTabIndex()` 一致。
- 不传（或传 `undefined`）时活动标签由控制器持有，行为与旧版本完全一致。传数字时，值变化会同步到控制器；控制器内部切换后发出 `update:activeTabIndex` 供 `v-model` 更新。受控值可以在工作簿解析完成前传入，首个标签出现时会应用一次。
- `tabChange(tab)` 在每次真正提交的切换后触发一次，参数是完整的 `XlsxWorkbookTab`（含 `kind`、`id`、`index`、`sheetIndex`/`chartsheetIndex`）。打开工作簿时的首个标签、重命名/排序/隐藏导致的标签重建、普通重渲染都不会触发。
- 宿主自己写入 `activeTabIndex` 引起的切换不会回发 `tabChange`，避免路由或日志出现回环。
- 拖拽排序活动标签时它的 `index` 会变化但没有切换工作表：只发 `update:activeTabIndex`，不发 `tabChange`。

## 查找与替换

工具栏的放大镜按钮打开查找面板。搜索覆盖工作簿所有工作表（值和公式），命中后切到对应工作表并选中单元格；大小写/全字/正则开关与结果列表都在同一个面板里。

面板同时提供替换：输入替换文字后可点“替换”（仅改写当前选中结果）或“全部替换”（改写当前工作表范围内的所有命中）。下方两个复选项控制替换范围：

- “匹配整个单元格”：只有单元格文本与查询完全相等才替换。
- “包含公式”：默认开启，公式单元格按公式文本替换（例如 `=B2*22%` → `=B2*25%`）；公式单元格的缓存值不会被当成普通文字替换。关闭后只替换普通值。

替换命令在控制器上，宿主可以绕开内置面板直接调用：

```ts
const result = controller.replaceInSheet("PDF Viewer", "PDF Reader", {
  matchCase: false,   // 默认 false，与 Excel 一致
  wholeCell: false,   // 默认 false：替换单元格内的子串
  formula: true,      // 默认 true：同时替换公式文本
  sheetIndex: 0,      // 默认活动工作表；非活动表会先被激活
  range: { start: { row: 1, col: 0 }, end: { row: 40, col: 7 } }, // 默认整张已用区域
})
// result: { sheetIndex, workbookSheetIndex, replacedCells, replacedOccurrences, changed }
```

每次调用只写一条撤销记录：一次“全部替换”无论改了多少单元格，撤销一次即可全部还原，重做重新应用。替换后的值按键入规则转换（数字写回数字），并随 `exportXlsx()` 写入导出文件；`readOnly`、受保护工作表和 Worker 只读大文件模式不改动工作簿，直接返回 `changed: false`。`occurrence` 选项可精确到搜索结果报告的那一处：

```ts
controller.replaceInSheet("PDF Viewer", "PDF Reader", {
  occurrence: { cell: { row: 4, col: 2 }, start: 0, end: 10 },
})
```

## 打印与分页

工作表的 `pageSetup`、`printOptions`、`pageMargins`、`rowBreaks`/`colBreaks` 以及 `_xlnm.Print_Area`/`_xlnm.Print_Titles` 会解析进 `XlsxSheetData`。分页在 `@arcships/xlsx-core` 内完成，Vue 层只消费页面结果，不从 DOM 高度推导分页：

```ts
import { buildWorksheetPrintLayout } from "@arcships/xlsx-core"

const layout = buildWorksheetPrintLayout(controller.activeSheet)
// layout.pages: { range, paperSize, widthPx, heightPx, scale, rowPage, colPage, ... }[]
```

控制器提供同一份结果和打印动作：

| 命令 | 说明 |
|---|---|
| `printLayout` / `printPages` | 活动工作表的扁平页面列表；每页带单元格范围和纸张尺寸 |
| `getPrintLayout()` | 同 `printLayout`，适合在命令式代码里取一次快照 |
| `setPageBreakPreview(enabled)` | 在网格中显示 Excel 风格的蓝色虚线分页线和页码 |
| `openPrintPreview()` / `closePrintPreview()` | 打开/关闭分页打印预览；`XlsxPrintPreview` 也可由宿主单独挂载 |
| `print()` | 打开打印预览并调用 `window.print()`；无工作表或无 `window.print` 时返回 `false` |

`XlsxViewer` 的 `@media print` 只输出预览页面，每页一个区块；纸张方向、页边距、手动分页线、重复标题行/列和合并单元格跨页都由核心分页结果决定。导出 XLSX 时会保留 `pageMargins`、`printOptions`、`pageSetUpPr`、分页线和打印区域定义，重新打开后分页结果一致。

图表、地图和 WebGL 是按需入口：

```ts
import "@arcships/vue-xlsx/chart"
import "@arcships/vue-xlsx/map"
import "@arcships/vue-xlsx/webgl"
```

不使用这些功能时，不需要加载相应的大型实现。旧的网格、工具栏和浮层组件只为兼容保留，新代码从 `XlsxViewer` 和控制器开始。

## 作者入口

功能区新增“图表 / 批注 / 超链接 / 表格 / 条件格式 / 数据验证 / 冻结窗格”一组按钮，右键菜单提供同名入口。对话框由 `XlsxViewer`（以及最小化的 `XlsxSheetSurface` 右键菜单）自动渲染，宿主也可以直接挂载 `XlsxAuthoringDialog`：

```vue
<script setup lang="ts">
import { XlsxAuthoringDialog, XlsxViewer } from "@arcships/vue-xlsx"
</script>

<template>
  <XlsxViewer :controller="controller" />
  <!-- 或自行控制 dialog：kind = "chart" | "comment" | "hyperlink" | "table" | "conditional" | "validation" -->
  <XlsxAuthoringDialog :controller="controller" kind="comment" @close="close" />
</template>
```

控制器 API 与对话框一一对应，宿主可以完全不用内置 UI：

| 命令 | 说明 |
|---|---|
| `insertChart({ family, range, title })` | 从区域插入柱形/折线/饼图，返回新图表 |
| `getCellComment(cell?)` / `setCellComment(cell, input \| null)` | 读取、写入或删除单元格批注 |
| `getCellHyperlink(cell?)` / `setCellHyperlink(cell, input \| null)` | 读取、写入或移除超链接 |
| `getConditionalFormatRules()` / `addConditionalFormatRule(draft)` / `updateConditionalFormatRule(index, draft)` / `deleteConditionalFormatRule(index)` | 条件格式（色阶、数据条、图标集，含 cfvo 阈值） |
| `createTable({ range, name, hasHeaders })` / `removeTable(name)` / `getTableAtCell(cell?)` | 从区域创建/删除 ListObject |
| `setSelectionDataValidation(input \| null)` | 为选区写入或清除数据验证 |
| `freezePanes` / `setFreezePanes(freeze)` / `toggleFreezePanesAtActiveCell()` | 冻结窗格状态与切换 |

每个命令都会改动工作簿模型、刷新视图、写一条撤销记录，并随导出的 XLSX 保存；导出后再打开仍然保留。命令只在可写的主线程工作簿上生效：`readOnly`、受保护工作表或 Worker 只读大文件模式下返回 `false` / `null`，不做任何修改。详细实现见 [XLSX 作者接口设计](../xlsx-authoring-surface.md)。

## 剪切与移动单元格

选区按 `Ctrl/Cmd+X`、右键菜单“剪切”或功能区“剪切”进入剪切状态：源单元格暂时保留内容，网格在源区域上画虚线“marching ants”，辅助技术会播报“已剪切 K1:K2，按 Esc 取消”。真正的移动发生在下一次粘贴：

| 操作 | 行为 |
|---|---|
| `Ctrl/Cmd+V` / `pasteFromClipboard()` | 值、公式和样式写到目标位置，公式按位移换算；源区域在**同一个撤销条目**里清空 |
| 一次 `Ctrl/Cmd+Z` | 同时恢复源区域并撤掉目标内容；一次重做又把整段移动放回去 |
| `Ctrl/Cmd+C` 或再次剪切 | 结束旧的剪切状态，按普通复制或新的剪切处理 |
| 仅格式粘贴（`pasteSpecialFromClipboard({ mode: "formats" })`） | 只写样式，源内容保留，剪切状态结束 |
| `Escape` / `cancelCut()` | 取消剪切状态，工作簿不变；之后粘贴按系统剪贴板内容当普通复制处理 |

控制器接口：`cutSelection()` 返回是否进入剪切状态，`cutState` 给出 `{ range, sheetIndex }`（源区域和它所在的工作簿工作表），`cancelCut()` 返回是否真的取消了剪切。跨工作表剪切是允许的：虚线只画在源工作表上，粘贴到其他工作表后仍然清空源区域。重新加载或重建工作簿（撤销快照、重命名、换文件）会自动清除剪切状态，不会把旧文件的数据粘到新文件里。

## 复制、剪切与创建绘图副本

选中图表、图片或形状后，图层接收 `Ctrl/Cmd+C`、`Ctrl/Cmd+X`、`Ctrl/Cmd+V` 和 `Ctrl/Cmd+D`，右键菜单也提供“复制”“剪切”“粘贴”“创建副本”“删除”。绘图与单元格共用 `copySelectionToClipboard()`、`cutSelection()`、`pasteFromClipboard()`；另有 `copyDrawing()`、`cutDrawing()`、`pasteDrawing()`、`duplicateDrawing()` 和 `deleteSelectedDrawing()` 供宿主直接调用。

内部载荷的 `drawings` 数组为每个对象记录 `kind`、`model`、`anchor`、`relativeAnchor`、`size`、`sourceSheetId` 和 `sourceId`。粘贴会在当前活动工作表的活动单元格处克隆 DrawingML 锚点：图片/形状获得新的模型 id 和关系 id，图表获得新的 `xl/charts/chartClipboardN.xml` 图表部件和系列 id；目标工作表没有绘图部件时会自动创建。跨工作表粘贴、剪切移动（源对象在同一个撤销条目里删除）以及导出后重新加载都会保留对象；`Escape`/`cancelCut()` 同样会取消待定的绘图剪切。

## 单元格格式

功能区的合并与边框组对应 Excel 的同一组操作：

| 控件 | 命令 | 说明 |
|---|---|---|
| 边框预设下拉 | `setSelectionBorders({ preset, style, color })` | `preset` 为 `all` / `outside` / `left` / `right` / `top` / `bottom` / `none`；`outside` 只画选区外围，内部格子保持原样 |
| 边框颜色 | 同上 | 颜色写入预设涉及的每一面；默认细线灰色 `A1A1AA` |
| 所有边框 | `setSelectionBorders({ preset: "all" })` | 保留的快捷按钮，同样是上面那条命令 |
| 合并并居中 | `mergeAndCenterSelection()` | 合并选区并写入 `horizontal: center` + `vertical: center`，一条撤销记录 |
| 取消合并 | `unmergeSelection()` | 只拆合并区 |

只想合并、不改对齐的宿主用 `mergeSelection()`。边框命令都走现有的样式命令，因此跨越导出/重新打开后保留：预设会写成 `styles.xml` 里对应边的 `<border>`，颜色是 8 位 ARGB。`XlsxCellBorderStyleInput` 仍然允许逐边不同的样式与颜色，需要时直接调 `setRangeStyle(range, { border })`。

## 延迟解析

控制器配置了 `deferLoadingAboveBytes` 且文件超过阈值时，不会立即解析工作簿。`XlsxViewer` 会直接渲染一个内置提示卡片，显示文件名和大小，并提供“继续解析”和“取消”两个动作；宿主不需要再额外实现加载入口。

行为：

- “继续解析”调用 `controller.continueDeferredLoad()`，事件 `deferred-resumed` 在调用前触发；
- “取消”调用 `controller.cancelDeferredLoad()`，释放缓存的源文件字节并回到空查看器，事件 `deferred-cancel` 在调用前触发；
- 进入延迟状态时组件触发 `load-deferred`，事件参数为 `{ fileName, bytes }`；
- 根元素暴露 `data-state="deferred"`，便于宿主和端到端测试识别。

如需替换默认卡片，使用 `deferredLoad` 插槽：

```vue
<XlsxViewer :controller="controller">
  <template #deferredLoad="{ fileName, bytes, sizeLabel, resume, cancel }">
    <div>
      <p>{{ fileName }}（{{ sizeLabel }}）较大，是否继续解析？</p>
      <button type="button" @click="resume">继续解析</button>
      <button type="button" @click="cancel">取消</button>
    </div>
  </template>
</XlsxViewer>
```

插槽属性：`fileName`、`bytes`、`sizeLabel`，以及 `resume`、`cancel` 两个动作函数。

## CSV 输入

控制器可以直接打开 CSV。接入方仍然传入 `file`/`src`；本地 `ArrayBuffer` 需要同时传入以 `.csv` 结尾的 `fileName`：

```ts
const controller = useXlsxViewerController({
  file: await csvFile.arrayBuffer(),
  fileName: csvFile.name,
})
```

远程 CSV 根据最终 URL 的 `.csv` 扩展名或响应的 `Content-Type: text/csv` 识别。CSV 会作为单工作表打开，支持 UTF-8（含 BOM）以及带 BOM 的 UTF-16；格式、图表和多工作表等 XLSX 专属信息不属于 CSV 文件内容。
