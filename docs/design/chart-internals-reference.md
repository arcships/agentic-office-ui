# Chart internals 参考：OOXML/Excel 图表几何与开源渲染器（#184）

> 批次：0.8.1 / issue [#184](https://github.com/arcships/agentic-office-ui/issues/184)。
> 关联：[#176](https://github.com/arcships/agentic-office-ui/issues/176) 对象矩阵、[#185](https://github.com/arcships/agentic-office-ui/issues/185) viewer/editor 架构审计。
> 引用核对日期：2026-11-21。本文的“精确”一律指**我们的渲染器在同一 frame 上实际绘制的坐标**，不等于 Microsoft Excel 的像素对齐；Excel/PowerPoint 的自动布局没有公开公式，不能承诺 parity。

## 0. 结论摘要

1. **OOXML 只规定“手动布局”，不规定自动布局。** `c:plotArea`/`c:legend`/`c:title`/`c:dLbl` 都可以带 `c:layout/c:manualLayout`，其中 `layoutTarget`、`xMode`/`yMode`/`wMode`/`hMode`、`x`/`y`/`w`/`h` 的语义是标准化的；但未写 `manualLayout` 时，plot area 的默认矩形、legend 的尺寸、axis label 的碰撞/换行/旋转全部由应用自行决定（§2.2、§2.3）。
2. **Excel 的行为可以观察到，但没有可移植公式。** Excel 会在窄图上倾斜或隐藏 category label，用户可以再用“interval between labels”强制显示；官方帮助只描述控件，不描述算法。LibreOffice、ONLYOFFICE 各自用“先量文本、再分配剩余空间”的流程实现，且互相结果不同（§2.3、§3）。
3. **我们缺的不是 XLSX 布局算法，而是公开几何。** `vue-xlsx` 的 SVG renderer 已有 `buildLayout`、axis/legend/bar/line/combo 的确定性布局；`xlsx-core` 已解析 chart 模型（series/axes/typeGroups/legend/dataLabels）。但公开 wrapper 只暴露 SVG 组件，私有 DOM 的 `data-xlsx-chart-*` 不是合同（[viewer-editor-architecture.md](viewer-editor-architecture.md) §6.4）。
4. **本轮落地 provider 合同 + XLSX 具体实现。** `@arcships/office-interaction` 新增框架无关的 `ChartGeometryProvider`/`ChartGeometrySnapshot`；`@arcships/vue-xlsx/chart` 新增 `computeXlsxChartGeometry()` 与 `createXlsxChartGeometryProvider()`，覆盖 2-D column/bar、line/area、column+line/area combo 的 plot/legend/axis/series 几何，并对 pie/scatter/radar/3-D/surface 等明确返回 `limitations` + `provider`，不伪造坐标。
5. **PPTX 仍是 `provider-dependent`。** `pptx-core` 的 `PptxChartSnapshot` 原来只有 series/point 标签；本轮增加可选的 `geometry?: ChartGeometrySnapshot` 接缝，由使用 `@aiden0z/pptx-renderer` + ECharts 的宿主把 renderer 内几何填入。具体 ECharts 适配是 follow-up（§8）。
6. **不承诺 provider 字段的像素 parity。** 字段级 `confidence` 区分 `exact`（渲染器实际绘制值）、`approximate`（文本度量/碰撞估计）、`provider`（引擎无规则，只能由 renderer 提供）；见 §5、§7。

## 1. 坐标模型与术语

### 1.1 chart-local CSS 像素

provider 的坐标空间固定为 `chart-local-px`：

- 原点：chart area（`c:chartSpace` 的 chart 画布）左上角；
- 单位：CSS px，与我们的 SVG viewport 相同；
- frame：provider 实际布局的 `{ width, height }`。`vue-xlsx` 的 `buildLayout` 会把 frame 取整并下限到 `80 × 60`，所以 snapshot 返回的是**取整后的 frame**，不是调用方传入的未取整值；
- Surface 把 snapshot 映射到页面时，加上 chart 自身的 page/slide/sheet 偏移即可；图表缩放/窗口变化时重新请求 snapshot，不缓存旧 frame 的坐标。

不采用 `0..1` 归一化作为主坐标，因为：

1. hit-test 需要与 SVG 的 `<rect>`/`<path>` 在同一坐标系，归一化会引入二次舍入；
2. 文本盒是估算值，归一化会把“近似”伪装成“精确比例”；
3. Surface 只在 chart 矩形内消费，chart-local px 已经足够，归一化可由宿主按需做。

### 1.2 confidence 三级

| 值 | 含义 | 例子 | 能做什么 |
|---|---|---|---|
| `exact` | 用与渲染器相同的公式、在同一 frame 上算出的**绘制坐标** | plot area、bar rect、line vertex、axis tick anchor | 命中、轮廓、选区手柄 |
| `approximate` | 依赖字体度量、换行/碰撞启发式或估算文本盒 | tick label 的 glyph box、legend item 的 text box、旋转 label 的 bbox | 近似命中、无障碍标签定位；不承诺像素 |
| `provider` | 引擎没有规则，必须由 renderer/provider 提供 | 3-D 投影 hit region、ECharts 内部 layout、Excel 自动 plot area | 只能经 provider 接入；没有 provider 时明确缺失 |

`exact` **不等于** Excel parity。它只说明“这个数字是我们在该 frame 上画出来的数字”。跨产品像素对齐需要另立 office-reference 验收，不能拿 provider 合同冒充。

## 2. OOXML/Excel 图表几何：哪些是标准，哪些是应用行为

### 2.1 结构树

```
c:chartSpace
└─ c:chart
   ├─ c:title / c:autoTitleDeleted
   ├─ c:plotArea
   │  ├─ c:layout?
   │  ├─ c:barChart | c:lineChart | c:pieChart | …（一个或多个 chart group）
   │  │  ├─ c:ser*
   │  │  │  ├─ c:idx / c:order / c:tx
   │  │  │  ├─ c:spPr / c:cat / c:val / c:bubbleSize
   │  │  │  ├─ c:dPt*（点级覆盖）
   │  │  │  └─ c:dLbls*（标签）
   │  │  ├─ c:dLbls?
   │  │  ├─ c:gapWidth? / c:overlap?（bar/column）
   │  │  └─ c:axId*
   │  ├─ c:catAx | c:dateAx | c:valAx | c:serAx（坐标轴）
   │  ├─ c:dTable?
   │  └─ c:spPr?
   ├─ c:legend?
   │  ├─ c:legendPos / c:legendEntry*
   │  ├─ c:layout?
   │  └─ c:overlay
   ├─ c:plotVisOnly / c:dispBlanksAs
   └─ c:spPr / c:txPr
```

`c:ser`、`c:dPt`、`c:dLbls` 本身**不带位置**；它们只描述数据、样式和标签内容。位置来自 chart group 的坐标轴、gap/overlap、应用布局以及可选的 manual layout。

### 2.2 manualLayout：唯一被标准化的几何

`c:layout` 的唯一子元素是 `c:manualLayout`（[Layout class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.layout?view=openxml-3.0.1)）。`manualLayout` 允许：

| 子元素 | 语义 |
|---|---|
| `c:layoutTarget` | `inner` = plot area 不含 tick mark 和 axis labels；`outer` = 含 tick mark 和 axis labels（[ST_LayoutTarget](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_LayoutTarget_topic_ID0EXWSRB.html)） |
| `c:xMode` / `c:yMode` | `edge`：`x`/`y` 是元素到 chart 左/上边缘的距离，按 chart 宽/高百分比；`factor`：`x`/`y` 是相对**默认位置**的偏移，按 chart 宽/高百分比（[MS-OI29500 §21.2.3.20](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/bb8565e7-dfea-453f-868b-ae680ab8a823)） |
| `c:wMode` / `c:hMode` | `edge`：`w`/`h` 是到 chart 右/下边缘的距离；`factor`：`w`/`h` 是元素宽/高占 chart 宽/高的百分比（同上） |
| `c:x` / `c:y` / `c:w` / `c:h` | 0..1 的比例值；应用应 clamp 到 chart frame（ONLYOFFICE 的 `calculatePosByLayout` 就是 `edge` 直接乘 chart size、`factor` 加默认位置，再 clamp 到 `[0, chartSize]`，[ChartSpace.js](https://github.com/ONLYOFFICE/sdkjs/blob/master/common/Drawings/Format/ChartSpace.js)） |

关键限制：

- **没有默认值公式。** 未出现 `manualLayout` 时，标准不定义 plot rect、legend 尺寸、title 高度、label 间距；`factor` 模式甚至以“应用默认位置”为基准，标准没有规定默认位置是什么。
- **manual layout 只覆盖元素整体**。它不描述 series 内 bar 的宽度、point 的坐标、legend item 的行列、axis label 的碰撞；这些仍由数据/坐标轴/应用布局推导。
- **我们的 parser 目前只读 dLbl 的 manual x/y**：`parseChartPointDataLabelsFromXml` 读取 `c:dLbl/c:layout/c:manualLayout/c:x|y`（[chart-parser.ts:200](../../packages/xlsx-core/src/charts/chart-parser.ts#L200)）；`c:plotArea`/`c:legend`/`c:title` 的 manualLayout 没有进入 `XlsxChart` 模型。因此 XLSX provider 明确带 `xlsx.manual-layout-not-parsed` limitation，见 §6.3。

### 2.3 自动布局：标准留白，各产品自行测量

证据：

- `c:legend` 只有 `legendPos`、`legendEntry`、`layout`、`overlay`（[Legend class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.legend?view=openxml-3.0.1)），没有 legend 宽高。
- `c:catAx`/`c:dateAx`/`c:serAx` 的 `tickLblSkip` 只是“画多少个 tick 后跳过一个 label”的**用户设置**（[TickLabelSkip](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.ticklabelskip?view=openxml-3.0.1)）；没有“当 label 太宽时自动跳过/旋转”的算法。
- `c:valAx` 的 `majorUnit`、`logBase`、`crosses`、`crossBetween` 只描述数值轴刻度与交叉方式（[ValueAxis](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.valueaxis?view=openxml-3.0.1)、[CrossBetween](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.crossbetween?view=openxml-3.0.1)）；不描述 axis label 的像素盒。
- `c:dLbls` 有 `dLblPos`（如 `ctr`/`inEnd`/`outEnd`/`bestFit`）、`showLeaderLines`、`layout`（[DataLabels](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.datalabels?view=openxml-3.0.1)），但 `bestFit` 如何避让、何时画 leader line 未定义。

#### Excel

Excel 的实际行为（可观察、不可公式化）：

- 窄图上优先倾斜 category label，仍放不下就隐藏一部分；用户可在 Format Axis → Labels → Interval between labels 强制间隔。Peltier Tech 的文章把这个过程描述为“Excel tries very hard … incline the labels … may also decide to only show some labels”（[Stagger Axis Labels](https://peltiertech.com/stagger-axis-labels-prevent-overlap/)）。
- 官方帮助只列控件（tick marks、label position、interval），不给布局公式（[Change the display of chart axes](https://support.microsoft.com/en-us/office/change-the-display-of-chart-axes)）。
- 因而 **Excel 自动 plot area / legend / label 位置属于 `provider-dependent`**：要拿到 Excel 的“精确”值，只能由 Excel/Office 渲染服务提供，不能从 OOXML 推。

#### LibreOffice

LibreOffice 的 chart2 是“顺序占位 + 文本测量”的典型实现：

- `ChartView::createShapes2D` 先 `formatPage`，再依次创建 title、legend、axis title，最后 `getAvailablePosAndSizeForDiagram` → `impl_createDiagramAndContent`；每一步都从上一步的 `maRemainingSpace` 扣减（[ChartView.cxx](https://github.com/LibreOffice/core/blob/master/chart2/source/view/main/ChartView.cxx)）。
- `impl_createDiagramAndContent` 在画完 axes 后，如果实际消费的 outer rect 与预期不同，会调用 `VDiagram::adjustInnerSize` 重新调整 inner rect；这说明 plot area 与 axis label 是**迭代测量**关系，不是一次闭式公式。
- `VLegend.cxx` 先为每个 entry 创建 text shape 并读取真实 `getSize()`，再算 column width/row height；`lcl_placeLegendEntries` 对 `ChartLegendExpansion_HIGH` 把文本最大宽度限制为可用宽度的 30%，行高/列宽按最大 entry 计算，空间不足时删 entry 或截断为 `...`（[VLegend.cxx](https://github.com/LibreOffice/core/blob/master/chart2/source/view/main/VLegend.cxx)）。
- 字体也按 `ReferencePageSize` 做相对缩放（`RelativeSizeHelper::calculate`），所以同一个 chart 在不同画布尺寸下不是简单等比缩放。

可借鉴的是**流程**（测量文本 → 分配 box → 剩余空间给 plot → 迭代校正），不是具体数值；LibreOffice 的数字不是 Excel 的数字。

#### ONLYOFFICE

ONLYOFFICE 的 `ChartsDrawer` 把 margin 写成显式模型：

- `recalculate()` 先 `_calculateProperties` → `_calculateMarginsChart` → `plotAreaChart.recalculate`；2-D 标准 margin 为 14（px 语义），3-D 为 16（[ChartsDrawer.js](https://github.com/ONLYOFFICE/sdkjs/blob/master/common/Charts/ChartsDrawer.js)）。
- `_calculateMarginsChartOrdinaryDiagramsOld` 把 margin 拆成“outward things: title and legend labels”与“inward things: axis titles, axes themselves and standard margin”，并规定优先级 `axisTitle > axis > normalMargin`；`_calculateMarginsChartOrdinaryDiagrams` 的新实现还区分 legend 四边、data table、title。
- `calculateSizePlotArea` 返回 `{w, h, startX, startY}`，由 `chartGutter` 四边扣减；`_checkMargins` 在负值/超出画布时回退到标准 margin。
- manualLayout 通过 `CChartSpace.calculatePosByLayout`/`calculateSizeByLayout` 实现 `edge`/`factor`，并 clamp 到 chart frame（[ChartSpace.js](https://github.com/ONLYOFFICE/sdkjs/blob/master/common/Drawings/Format/ChartSpace.js)）。

可借鉴的是**“outward/inward + 标准 margin + 兜底”**的模型；它同样不是 Excel 公式。

#### Univer

Univer 的图表文档明确使用 ECharts 作为渲染/导出路径（`registerTheme`、`chart.exportImage()`；[Univer Charts](https://docs.univer.ai/guides/sheets/features/charts)）。因此它的 chart internals 几何来自 ECharts，不是 OOXML/Excel 语义；只能作为“ECharts 布局参考”，不能作为 Office parity 依据。

#### Apache POI / ExcelJS

- Apache POI 官方限制写明 “XSSF has only limited chart support including making some simple changes and adding at least some line and scatter charts”（[POI limitations](https://poi.apache.org/components/spreadsheet/limitations.html)）；POI 是模型/读写库，不做图表布局/渲染。
- ExcelJS 的 chart support 仍是开放 issue（[exceljs#141](https://github.com/exceljs/exceljs/issues/141)）；没有可参考的 chart layout。

结论：**没有开源实现提供“Excel 等价”的自动图表布局**；LibreOffice/ONLYOFFICE 是可借鉴的测量式实现，ECharts/Chart.js 是可借鉴的通用 label/legend 布局，但不是 Office 语义。

### 2.4 坐标轴：category / date / value / series

| 轴 | OOXML | 语义 | 我们的处理 |
|---|---|---|---|
| category | `c:catAx` | 离散 category；`orientation` 可反序；`tickLblSkip` 控制 label 间隔 | `scaleBand`（bar）或 `scalePoint`（line/area）；label 碰撞见 §2.5 |
| date | `c:dateAx` | 连续日期轴；`baseTimeUnit`、`majorUnit`+`majorTimeUnit`、`minorUnit`+`minorTimeUnit` | `excelSerialToDate` + `scaleTime` + `buildDateAxisTicks`；按 span 选 days/months/years |
| value | `c:valAx` | 连续数值；`min`/`max`/`majorUnit`/`logBase`/`orientation`/`crosses`/`crossesAt`/`crossBetween` | `scaleLinear`/`scaleLog`；`resolveNumericAxisDomain` + `majorUnit` |
| series | `c:serAx` | 3-D 深度轴；2-D 不使用 | 3-D 不在本轮 provider 范围 |

两个常见误解：

1. **“百分比轴”不是独立的 OOXML 轴类型。** 100% stacked chart 的百分比语义来自 chart group 的 `grouping="percentStacked"`，渲染时把每个 category 的绝对值归一化到 100；轴的 number format 只是显示 `%`。我们的实现是在 bar/line 的 percent-stacked 分支里做归一化（[chart-bar.tsx](../../packages/vue-xlsx/src/render/chart-bar.tsx)、[chart-line.tsx](../../packages/vue-xlsx/src/render/chart-line.tsx)），provider 的 value 也是归一化后的值。
2. **日期轴不是“把 category 文本按日期排序”。** `c:dateAx` 用 Excel serial（1900 或 1904 基准）作为连续 domain；`majorUnit` 与 `majorTimeUnit` 决定 tick 间隔。我们的 `buildDateAxisTicks` 在 span ≤45 天用 days、≤900 天用 months、否则 years，并优先使用 `majorUnit`。

### 2.5 我们当前的布局规则（可作为 provider 的“精确”定义）

以下公式是 XLSX provider 与 SVG renderer 的共同来源；provider 不是另写一套“Office 近似”，而是复用/镜像这些规则。

#### (a) chart frame 与 plot area

`buildLayout(chart, rect, legendItems)`（[chart-legend.tsx:288](../../packages/vue-xlsx/src/render/chart-legend.tsx#L288)）：

```
width  = max(80, round(rect.width))
height = max(60, round(rect.height))
titleHeight = chart.title ? (surface ? 30 : 24) : 8
legendVertical = legend && position ∈ {left, right}
finalLegendPos = (width <= 280 || height <= 200) && legendVertical ? "bottom" : position
baseLeft = max(42, horizontalBar ? categoryLabelWidth + 12 : valueLabelWidth + 12)
plotLeft = baseLeft + (legendLeft ? legendWidth : 0)
plotRight = 16 + secondaryAxisWidth + (legendRight ? legendWidth : 0)
plotTop = titleHeight + (legendTop ? legendHeight + 4 : 0)
plotBottom = baseBottom + (legendBottom ? legendHeight + 4 : 0)
plot = {
  left: plotLeft,
  top: plotTop,
  width: max(40, width - plotLeft - plotRight),
  height: max(40, height - plotTop - plotBottom)
}
```

value label 预留宽度：

```
candidates = {0, axis.min?, axis.max?, 每个 series 的前 200 个 value}
maxLabelWidth = max(estimateSvgTextWidth(formatTick(v), 10))
valueLabelWidth = min(maxLabelWidth + 12, max(40, width * 0.36))
```

secondary axis（combo 第二 value axis）：

```
secondaryAxisWidth = min(maxLabelWidth + 18 + (axis.title ? 14 : 0), max(40, width * 0.3))
```

category label 底部预留（非 horizontal bar）：

```
maxLabelWidth = max(estimateSvgTextWidth(label, 10))
slotWidth = (width - 60) / ceil(categoryCount / tickLabelSkip)
if maxLabelWidth <= slotWidth * 0.92: baseBottom = 28
else: baseBottom = max(28, min(104, 18 + maxLabelWidth * sqrt(1/2)))
```

这些数字是**我们的 renderer 规则**，不是 Excel 的默认值。

#### (b) legend

`buildLegendMetrics`（[chart-legend.tsx:218](../../packages/vue-xlsx/src/render/chart-legend.tsx#L218)）：

- horizontal（top/bottom）：
  - `maxTextWidth = max(40, min(140, width * 0.45))`；
  - 每个 entry 先 `wrapSvgText(label, maxTextWidth, 10)`；
  - `entryWidth = 12 + maxLineWidth + 8`；
  - 行宽超过 `max(60, width - 16)` 时换行；
  - `rowHeight = max(lines in row) * 12 + 2`；
  - `legendHeight = min(height, sum(rowHeights) + 8)`。
- vertical（left/right）：
  - `maxTextWidth = max(48, min(160, width * 0.4))`；
  - `rowHeight = lines * 12 + 6`；
  - `legendWidth = clamp(contentWidth + 16, 60, width * 0.5)`。

provider 的 legend item `rect` 用同一 `cursorY`/`rowTop` 算法；`swatchRect` 是 8×8，精确；`labelRect`/item 盒因文本宽度估算而 `approximate`。

#### (c) bar/column

`renderBarChart`（[chart-bar.tsx:12](../../packages/vue-xlsx/src/render/chart-bar.tsx#L12)）：

```
padding = gapWidth
  ? { inner: clamp(gap/(100+gap), 0.05, 0.88),
      outer: clamp(inner * 0.5, 0, 0.45) }
  : { inner: 0.6, outer: 0.3 }   // gapWidth 默认 150
categoryScale = scaleBand(categories, plot.range, padding)
seriesScale   = scaleBand(seriesIndexes, [0, categoryScale.bandwidth()], {inner: 0.16, outer: 0.08})
barThickness  = stacked ? categoryScale.bandwidth() : seriesScale.bandwidth()
barOffset     = stacked ? 0 : seriesScale(String(seriesIndex))
```

`scaleBand` 的带宽：

```
step = range / (n - paddingInner + 2 * paddingOuter)
bandwidth = step * (1 - paddingInner)
```

vertical column 的 rect：

```
y1 = scale(start); y2 = scale(end)
rect = { x: categoryStart + barOffset, y: min(y1,y2),
         width: max(1, barThickness), height: max(1, |y2-y1|) }
```

horizontal bar 对称交换 x/y。`gapWidth` 的 OOXML 语义是“cluster 间距占 bar 宽的百分比”，范围 0..500（[GapWidth](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.gapwidth?view=openxml-3.0.1)）；`overlap` 是 2-D bar 的重叠百分比，范围 -100..100（[Overlap](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.overlap?view=openxml-3.0.1)）。我们的 renderer 目前只使用 `gapWidth`，不消费 `overlap`。

#### (d) value domain 与 tick

`resolveNumericAxisDomain` / `buildNiceStep`（[chart-axis.tsx:28](../../packages/vue-xlsx/src/render/chart-axis.tsx#L28)）：

```
span = max(1e-6, max - min)
rough = span / 5
magnitude = 10^floor(log10(rough))
normalized = rough / magnitude
nice = normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10
step = nice * magnitude
```

若 `axis.majorUnit > 0`，直接使用 `majorUnit`，不取整 domain 边界；否则：

```
roundedMin = floor(min / step) * step
roundedMax = ceil(max / step) * step
ticks = 从 floor(min/step)*step 起，每 step 一个，落在 [min, max] 内
```

`includeZero` 由 chart 类型决定：bar/column 总是 `true`；line 在 `crosses="autoZero"` 时为 `true`；area 总是把 min clamp 到 0。log 轴用 `resolveLogAxisDomain`，tick 为 base 的幂。

#### (e) line/area 与 combo

- line/area 的 category 位置是 `scalePoint`，即 `x_i = left + i * width/(n-1)`，没有 band padding（[chart-line.tsx:10](../../packages/vue-xlsx/src/render/chart-line.tsx#L10)）。
- date axis 用 `scaleTime`，category label 来自 `buildDateAxisTicks`。
- combo 的 category 位置改用 `scaleBand` 的 band center；column group 与 line group 共享 category 位置，line 使用自己的 secondary value scale（[chart-line.tsx:443](../../packages/vue-xlsx/src/render/chart-line.tsx#L443)）。
- **已知 renderer 行为**：combo 的 column 值在 `Math.max(logBaseline, value)` 时把负值 clamp 到 0（`logBaseline` 非 log 轴时为 0），所以负 column 只画 1px 高；provider 镜像这个“已绘制”结果并给出 `xlsx.combo-negative-zero-clamped` limitation。普通 bar 不受影响（负值正常向下画）。

#### (f) category label 碰撞

`resolveCategoryLabelLayout`（[chart-axis.tsx:385](../../packages/vue-xlsx/src/render/chart-axis.tsx#L385)）：

```
baseStep = tickLabelSkip > 0 ? floor(tickLabelSkip) : 1
visibleCount = ceil(count / baseStep)
slotWidth = max(1, availableWidth / visibleCount)
maxLabelWidth = max(estimateSvgTextWidth(label, 10))
if maxLabelWidth <= slotWidth * 0.92:
  rotate = false; step = baseStep
else:
  projectedWidth = maxLabelWidth * sqrt(1/2)
  overlapStep = projectedWidth > slotWidth * 0.92
    ? ceil(projectedWidth / (slotWidth * 0.92)) : 1
  rotate = true; step = baseStep * overlapStep
```

horizontal bar 的 category label 不旋转，使用 `step = max(tickLabelSkip, ceil(12 / (plot.height / categoryCount)))`；value label 的 skip 由 `valueAxis.tickLabelSkip` 控制。该规则是我们的确定性近似，Excel 的具体倾斜/隐藏点没有公开公式。

#### (g) date tick

`buildDateAxisTicks`（[chart-axis.tsx:247](../../packages/vue-xlsx/src/render/chart-axis.tsx#L247)）：

```
spanDays = maxSerial - minSerial
unit = axis.majorTimeUnit if days/months/years
     : spanDays <= 45 ? days
     : spanDays <= 900 ? months
     : years
step = axis.majorUnit if > 0
     : first candidate in (days: 1,2,7,14,30 | months: 1,2,3,6 | years: 1,2,5,10)
       with spanUnits / candidate <= 8
```

这是对 Excel 日期轴的可运行近似；Excel 在月末/年界、1904 基准和 `baseTimeUnit` 上的细节仍需 fixture 对照。

## 3. 开源渲染器对照

| 实现 | 有 chart 布局？ | 碰撞/换行 | manualLayout | 对 #184 的价值 | 不能做什么 |
|---|---|---|---|---|---|
| Microsoft Excel | 有，闭源 | 倾斜/隐藏/用户 interval；无公开公式 | 支持 | office-reference 的唯一真值 | 不能从 OOXML 推导自动布局 |
| LibreOffice chart2 | 有，测量式 | `VLegend` 量文本、限宽 30%、删 entry/省略号；`VDiagram::adjustInnerSize` 迭代 plot | 支持 | 借鉴“测量 → 分配 → 剩余空间 → 迭代” | 数值不是 Excel |
| ONLYOFFICE sdkjs | 有，显式 margin | title/legend/axisTitle/axis 优先级，标准 margin 14/16，越界回退 | 支持（`calculatePosByLayout`） | 借鉴 outward/inward margin 模型 | 数值不是 Excel |
| Univer | 有，走 ECharts | 继承 ECharts | 由 ECharts 决定 | ECharts 适配参考 | 不是 Office 语义 |
| Apache POI | 无渲染 | — | 仅模型 | 解析参考 | 不能当布局参考 |
| ExcelJS | 无 chart | — | — | 无 | 不能当布局参考 |
| ECharts | 有 | `calculateCategoryInterval` 以“avoid overlapping”为目标；`axisLabel.hideOverlap`；`grid.containLabel` 按 label bbox 缩小 grid（[axisTickLabelBuilder.ts](https://github.com/apache/echarts/blob/master/src/coord/axisTickLabelBuilder.ts)、[Grid.ts](https://github.com/apache/echarts/blob/master/src/coord/cartesian/Grid.ts)） | 自定义 | 借鉴 interval/overlap 策略 | 不是 Office 语义 |
| Chart.js | 有 | layout box 的 `fit`/`shrink`/`weight`/`fullSize`，`autoPadding`（[core.layouts.js](https://github.com/chartjs/Chart.js/blob/master/src/core/core.layouts.js)、[layout docs](https://www.chartjs.org/docs/latest/configuration/layout.html)） | 自定义 | 借鉴 box 模型 | 不是 Office 语义 |
| 我们的 SVG renderer | 有，确定性 | 自有 §2.5 规则 | 未解析 | provider 的 exact 来源 | 不是 Excel parity |

结论：**不引入第二套布局引擎**。provider 复用我们的 SVG 布局，把“哪些字段 exact/approximate/provider”写进合同；LibreOffice/ONLYOFFICE 只作为后续改进自动布局和 manual layout 的参考。

## 4. 我们的 parser/renderer 现状

### 4.1 `xlsx-core`

已解析（[chart-parser.ts](../../packages/xlsx-core/src/charts/chart-parser.ts)、[chart-types.ts](../../packages/xlsx-core/src/charts/chart-types.ts)）：

- chart 类型：`column`/`bar`/`line`/`area`/`pie`/`doughnut`/`scatter`/`bubble`/`radar`/`stock`/`surface`/`combo`（`typeGroups`）；
- series：`categories`、`values`、`dataPoints`、`dataPointStyles`、`marker*`、`valuesRef`/`categoriesRef`；
- axes：`axisType`、`min`/`max`/`majorUnit`/`logBase`/`orientation`/`crosses`/`crossBetween`/`tickLabelSkip`/`labelPosition`/`numberFormat`/`title`；
- legend、dataLabels、anchor、chart/plot area 的填充/图片、user shapes、embedded media。

未解析/未建模：

- `c:plotArea/c:layout/c:manualLayout`（`layoutTarget`/`xMode`/`yMode`/`wMode`/`hMode`/`x`/`y`/`w`/`h`）；
- `c:legend/c:layout/c:manualLayout`、`c:title/c:layout`；
- `c:barChart/c:overlap`（模型有 `overlap` 字段，但 renderer 未消费）；
- axis label 的实际字体度量（渲染器用估算）。

### 4.2 `vue-xlsx`

- `buildLayout`、`getLegendItems`、`renderCartesianAxes`、`renderBarChart`、`renderLineOrAreaChart`、`renderComboChart` 等已有确定性规则（§2.5）。
- 私有 DOM 写 `data-xlsx-chart-element-kind`、`data-xlsx-chart-series-index`、`data-xlsx-chart-point-index`，但这是 renderer 内部标记，不是公开几何合同。
- 本轮新增 `packages/vue-xlsx/src/render/chart-geometry.ts`，从 `@arcships/vue-xlsx/chart` 导出 `computeXlsxChartGeometry()` / `createXlsxChartGeometryProvider()`，不修改现有 renderer 行为。

### 4.3 `pptx-core`

- `PptxChartSnapshot` 原有 `series[].name` 与 `series[].points`（标签），`getObjectChart?` 由宿主提供（[reference-adapter.ts:54](../../packages/pptx-core/src/reference-adapter.ts#L54)）。
- PPTX 图表实际由 `@aiden0z/pptx-renderer`（ECharts）绘制；ECharts 的 plot/legend/axis 几何只在 renderer 内。
- 本轮在 `PptxChartSnapshot` 增加可选 `geometry?: ChartGeometrySnapshot`。这是**接缝**，不是实现；未提供时行为与原来一致。`pptx-core` 的对象索引（#176）可通过同一 snapshot 的 `geometryViewport` 把 chart-local 几何投影成 slide 归一化 sub-object 轮廓；Surface 渲染和 ECharts 适配仍是 follow-up。

## 5. Provider 合同

### 5.1 公共类型（`@arcships/office-interaction`）

```ts
type ChartGeometryConfidence = "exact" | "approximate" | "provider"

interface ChartGeometryPoint { x: number; y: number }
interface ChartGeometryRect extends ChartGeometryPoint { width: number; height: number }
interface ChartGeometryFrame { width: number; height: number }

interface ChartGeometrySnapshot {
  chartId: string
  provider: string
  format: OfficeFormat
  space: "chart-local-px"
  frame: ChartGeometryFrame
  layout: {
    chartArea: ChartGeometryRect
    plotArea: ChartGeometryRect
    title?: ChartGeometryRect
    titleLabel?: ChartGeometryLabel
    legend?: ChartLegendGeometry
    confidence: ChartGeometryConfidence
  }
  axes: readonly ChartAxisGeometry[]
  series: readonly ChartSeriesGeometry[]
  limitations?: readonly ChartGeometryLimitation[]
  exact: boolean
}

interface ChartGeometryProvider<TChart = unknown> {
  readonly id: string
  readonly format: OfficeFormat
  getChartGeometry(
    chart: TChart,
    request: ChartGeometryRequest,
  ): ChartGeometrySnapshot | undefined
}
```

`ChartGeometryRequest` 包含 `chartId?`、`frame`、`detail?: "layout" | "full"`、`documentId?`、`revision?`。`chartId` 不匹配时 provider 返回 `undefined`，避免 Surface 缓存把 A 图的几何用到 B 图。

配套纯函数：

- `validateChartGeometrySnapshot(snapshot)`：返回字段级 issue 字符串；测试用它保证 provider 输出结构合法。
- `assertChartGeometrySnapshot(snapshot)`：有 issue 时抛 `TypeError`。
- `chartGeometryRectContains` / `chartGeometryRectIntersects`。
- `hitTestChartGeometry(snapshot, point, options?)`：命中优先级 series mark → legend entry → axis tick label → plot area → chart area。

### 5.2 XLSX provider API

```ts
import {
  computeXlsxChartGeometry,
  createXlsxChartGeometryProvider,
} from "@arcships/vue-xlsx/chart"

const snapshot = computeXlsxChartGeometry(chart, {
  frame: { width: 567, height: 283 },
  palette,          // 仅影响 legend 颜色；几何不需要
  labels,           // 可选的本地化 series 名
})

const provider = createXlsxChartGeometryProvider({ palette, labels })
const snapshot2 = provider.getChartGeometry(chart, {
  chartId: chart.id,
  frame: { width: 567, height: 283 },
})
```

`provider.id = "xlsx-svg-renderer"`，`format = "xlsx"`。snapshot 的 `provider`、`space`、`frame`、`chartId` 都由 provider 填充。

### 5.3 兼容性

- 只新增类型/函数；没有修改现有函数签名。
- `@arcships/office-interaction` 根入口新增导出（public API contract 的 fingerprint 需要同步）。
- `@arcships/vue-xlsx/chart` 是可选入口，不改变 `@arcships/vue-xlsx` 根入口的历史导出；根入口的 deprecated chart 类型保持不变。
- `PptxChartSnapshot` 只增加可选字段，旧宿主不传 `geometry` 时行为不变。

## 6. 实现：XLSX provider

### 6.1 覆盖范围

| chartType | 本轮 provider | 说明 |
|---|---|---|
| `ColumnClustered` / `ColumnStacked` / `ColumnPercentStacked` | ✅ | bar rect、category/value axis、legend |
| `BarClustered` / `BarStacked` / `BarPercentStacked` | ✅ | horizontal；category 反序与 renderer 一致 |
| `Line` / `LineStacked` / `LinePercentStacked` | ✅ | point anchor（marker box approximate） |
| `Area` / `AreaStacked` / `AreaPercentStacked` | ✅ | point anchor（面积路径本身不返回） |
| combo `Column*` + `Line*`/`Area*` | ✅ | 双 value axis（若 axId 不同）、bar + line points |
| `Pie` / `Doughnut` / `Pie3D` / `PieExploded` / `BarOfPie` | ❌ | 只返回 layout + legend；limitation `xlsx.pie-provider` |
| `Scatter` / `Bubble` / `Radar` / `Stock` / `Surface` / `RegionMap` / `Treemap` / `Sunburst` / `Waterfall` / `Funnel` / `BoxWhisker` | ❌ | 只返回 layout + legend；limitation `<type>-provider` |
| 3-D（`chart.is3d` 或 combo group 3-D） | ❌ | 只返回 layout + legend；limitation `xlsx.3d-projection-provider` |

“只返回 layout + legend”不是失败：Surface 仍能命中 chart 边框、plot 背景、legend entry；只是不能命中 series/point/axis tick。`exact` 为 `false`，调用方必须检查 `limitations`。

### 6.2 fixture 证据

fixture：`apps/demo/public/samples/xlsx-chart-families.xlsx`（12 个 chart，包含 combo one-axis/two-axes、pie、line、3-D fallback）。

测试：

- `tests/unit/xlsx-chart-geometry.test.mjs`：解析 fixture，断言 `Combo Two Axes` 在 `567 × 283` 下的精确几何：
  - `plotArea = {42, 24, 398.6, 231}`；
  - category ticks `Q1..Q4` 的 label anchor `x = 91.825 / 191.475 / 291.125 / 390.775`，`y = 268`；
  - primary value ticks `[-10, 0, 10, 20, 30]`，secondary `[-5, 0, 5, 10, 15]`；
  - series 0 bars：`{75.0838, 139.5, 33.4824, 57.75}` 等；
  - series 1 line points：`(91.825, 151.05)`、`(191.475, 231.9)`、`(291.125, 104.85)`、`(390.775, 58.65)`；
  - legend item rects、seriesIndex 0/1；
  - `validateChartGeometrySnapshot` 无 issue；hit-test 命中 bar/legend/axis tick/plot area。
- `tests/component/xlsx-chart-geometry-parity.test.mjs`：用同一个 fixture 渲染 `MemoChartSvg`，逐项比较 provider 与 SVG：
  - 4 个 bar `<rect>` 的 `x/y/width/height` 与 provider 完全一致；
  - 4 个 line marker `translate(x, y)` 与 provider point 完全一致；
  - category / primary value / secondary value label anchor 与 SVG `<text>` 完全一致；
  - title anchor 一致。

这两个测试文件是“provider 合同 + 绘制几何”的可运行证据；**没有**生成 Excel/PowerPoint 截图矩阵，截图对照列为 §8 follow-up（本轮不伪造 office-reference 截图）。

### 6.3 limitations

snapshot 的 `limitations` 是本轮明确写出的“不承诺”：

| code | confidence | 含义 |
|---|---|---|
| `xlsx.text-metrics-estimated` | approximate | label/legend 的文本盒用估算 SVG 字宽；字体不同会差 |
| `xlsx.manual-layout-not-parsed` | approximate | `c:plotArea`/`c:legend`/`c:title` 的 manualLayout 未进入模型；provider 只报自动布局 |
| `xlsx.combo-negative-zero-clamped` | approximate | combo renderer 把负 column clamp 到 0；provider 报绘制结果 |
| `xlsx.3d-projection-provider` | provider | 3-D 投影/命中区域没有 2-D rect |
| `xlsx.<type>-provider` | provider | 该 chart family 的 series/point 几何未实现 |

`exact` 的定义：所有**结构性** part（layout、axes、series 顶层）都是 `exact`，且没有 `provider` 级 limitation。文本盒仍是 `approximate`，逐字段可见；不要用 `exact: true` 推断“没有近似字段”。

## 7. Exact / approximate / provider 矩阵

| 几何 part | 本轮 XLSX provider | confidence | 备注 |
|---|---|---|---|
| chartArea | ✅ | exact | frame 取整后的 `{0,0,w,h}` |
| plotArea | ✅ 2-D | exact | `buildLayout` 的 `layout.plot` |
| title band | ✅ | exact | 仅 title 预留带；`titleLabel` box 为 approximate |
| legend position / swatch | ✅ | exact / approximate | position + swatch 按 renderer 算法；item 盒含文本宽度，approximate |
| legend item rect / label box | ✅ | approximate | 文本宽度估算 |
| axis line / zero line | ✅ | exact | 与 renderer 相同 `plot` 和 `scale` |
| axis tick anchor | ✅ | exact | category band center / scalePoint / scaleLinear |
| axis tick label box | ✅ | approximate | 字体度量 |
| bar/column rect | ✅ | exact | 含 stacked/percent/horizontal |
| line/area point anchor | ✅ | exact | 含 stacked/percent/date axis |
| line marker box | ✅ | approximate | marker 形状/大小不同 |
| combo bar / line | ✅ | exact | 负 column clamp 有 limitation |
| pie/doughnut arc | ❌ | provider | follow-up |
| scatter/bubble/radar point | ❌ | provider | follow-up |
| 3-D / surface / region map / treemap / sunburst | ❌ | provider | follow-up |
| data label text/box | ❌ | provider/approximate | renderer 有 `dLbls`，provider 暂不暴露 |
| Excel 自动 plot/legend/label | ❌ | provider | 无公开公式；需 Office 参考服务 |
| manualLayout plot/legend/title | ❌ | approximate | 未解析；只报自动布局 |
| PPTX chart internals | 接缝 | provider | `PptxChartSnapshot.geometry?`，具体 ECharts 适配待做 |

## 8. Follow-up（按依赖排序）

1. **抽取共享几何内核。** 目前 provider 镜像 `renderBarChart`/`renderLineOrAreaChart`/`renderComboChart` 的 bar/point 公式，靠 parity test 防止漂移。下一步把这些公式抽成 renderer 与 provider 共用的纯函数，避免两处维护；再为 stacked/percent/horizontal/date 各加一个 parity case。
2. **解析并应用 manualLayout。** 在 `xlsx-core` 的 `XlsxChart` 增加 plot/legend/title 的 `manualLayout`（`layoutTarget`/`xMode`/`yMode`/`wMode`/`hMode`/`x`/`y`/`w`/`h`），在 `buildLayout` 前置应用；provider 去掉 `manual-layout-not-parsed`。这是把 `provider` 变成 `exact` 的最高价值项。
3. **Pie/doughnut arc provider。** 返回 `center`、`innerRadius`/`outerRadius`、`startAngle`/`endAngle`，并处理 explosion、3-D 前后面；Surface 可命中扇区。
4. **Scatter/bubble/radar provider。** scatter 需要 `categoryAxis`/`valueAxis` 双数值 domain；radar 需要极坐标；都要单独的 provider 合同测试。
5. **Surface 消费。** `XlsxSheetSurface` 用 `chartRect` + provider snapshot 做 hover/选中/轮廓；`PptxStage` 用 `PptxChartSnapshot.geometry` 消费宿主 ECharts 几何。先做 XLSX（有 exact provider），再做 PPTX（provider 依赖）。
6. **PPTX ECharts 适配。** 在 `@aiden0z/pptx-renderer` 的 chart 实例上读取 `getModel()`/`convertToPixel()`/`getZr()`，把 series/point/legend/axis 转成 `ChartGeometrySnapshot`；需要处理 zoom/pan 和 ECharts 版本升级。
7. **Date/percent axis 对照。** 增加 dateAxis、100% stacked、log axis 的 Excel/LibreOffice fixture 与截图；验证 `buildDateAxisTicks` 与 manual `majorUnit`/`baseTimeUnit`。
8. **Data label 几何。** `dLbls` 的 text/box、`bestFit`、leader line 目前只影响视觉，不进 provider；后续按 `resolveDataLabelsForPoint` + 文本度量输出 `ChartGeometryLabel`，并标注 approximate。
9. **Excel/PowerPoint 截图矩阵。** 同一 fixture 在 Excel、PowerPoint、LibreOffice 的 4 viewport/缩放输出；用于 office-reference 验收，不作为 provider 单元测试。
10. **更新 #176 矩阵。** `viewer-editor-architecture.md` 的 PPTX chart internals 行改为“provider 接缝已加，Surface 未接入”；XLSX chart internals 行改为“provider 合同 + 2-D bar/line/combo 已实现，pie/scatter/3-D provider-dependent”。

## 9. 引用

核对日期均为 2026-11-21。

1. ISO/IEC 29500-1 / ECMA-376（DrawingML Charts §21.2）：<https://ecma-international.org/publications-and-standards/standards/ecma-376/>
2. `c:manualLayout`（ISO 29500-1 §21.2.2.104 摘录）：<https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.manuallayout?view=openxml-3.0.1>
3. `c:layout`：<https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.layout?view=openxml-3.0.1>
4. `c:plotArea`：<https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.plotarea?view=openxml-3.0.1>
5. `ST_LayoutTarget`（inner/outer）：<https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_LayoutTarget_topic_ID0EXWSRB.html>
6. MS-OI29500 §21.2.3.20 `ST_LayoutMode`（edge/factor 的 Office 行为）：<https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/bb8565e7-dfea-453f-868b-ae680ab8a823>
7. `c:gapWidth`：<https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.gapwidth?view=openxml-3.0.1>
8. `c:overlap`：<https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.overlap?view=openxml-3.0.1>
9. `c:tickLblSkip`：<https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.ticklabelskip?view=openxml-3.0.1>
10. `c:crossBetween`：<https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.crossbetween?view=openxml-3.0.1>
11. `c:legend`：<https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.legend?view=openxml-3.0.1>
12. `c:legendPos`：<https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.legendposition?view=openxml-3.0.1>
13. `c:dateAx`：<https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.dateaxis?view=openxml-3.0.1>
14. `c:valAx`：<https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.valueaxis?view=openxml-3.0.1>
15. `c:dLbls`：<https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.datalabels?view=openxml-3.0.1>
16. `c:barChart`：<https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.barchart?view=openxml-3.0.1>
17. LibreOffice `ChartView.cxx`（顺序占位、remaining space、`adjustInnerSize`）：<https://github.com/LibreOffice/core/blob/master/chart2/source/view/main/ChartView.cxx>
18. LibreOffice `VLegend.cxx`（量文本、列宽/行高、30% 限宽、删 entry/省略号）：<https://github.com/LibreOffice/core/blob/master/chart2/source/view/main/VLegend.cxx>
19. ONLYOFFICE `ChartsDrawer.js`（margin 模型、标准 margin 14/16、`calculateSizePlotArea`）：<https://github.com/ONLYOFFICE/sdkjs/blob/master/common/Charts/ChartsDrawer.js>
20. ONLYOFFICE `ChartSpace.js`（manualLayout `calculatePosByLayout`/`calculateSizeByLayout`）：<https://github.com/ONLYOFFICE/sdkjs/blob/master/common/Drawings/Format/ChartSpace.js>
21. Univer Charts（ECharts 渲染/导出）：<https://docs.univer.ai/guides/sheets/features/charts>
22. Apache POI chart 限制：<https://poi.apache.org/components/spreadsheet/limitations.html>
23. ExcelJS chart support issue：<https://github.com/exceljs/exceljs/issues/141>
24. ECharts category tick label interval（“avoid overlapping”）：<https://github.com/apache/echarts/blob/master/src/coord/axisTickLabelBuilder.ts>
25. ECharts `grid.containLabel` 实现：<https://github.com/apache/echarts/blob/master/src/coord/cartesian/Grid.ts>
26. Chart.js layout boxes（`fitBoxes`/`fullSize`/`weight`）：<https://github.com/chartjs/Chart.js/blob/master/src/core/core.layouts.js>
27. Chart.js layout 文档：<https://www.chartjs.org/docs/latest/configuration/layout.html>
28. Peltier Tech，Excel axis label 倾斜/隐藏行为：<https://peltiertech.com/stagger-axis-labels-prevent-overlap/>
29. Microsoft Support，chart axes 显示/interval 控件：<https://support.microsoft.com/en-us/office/change-the-display-of-chart-axes>
30. 我们自己的实现：`packages/vue-xlsx/src/render/chart-geometry.ts`、`packages/vue-xlsx/src/render/chart-legend.tsx`、`packages/vue-xlsx/src/render/chart-axis.tsx`、`packages/vue-xlsx/src/render/chart-bar.tsx`、`packages/vue-xlsx/src/render/chart-line.tsx`。
