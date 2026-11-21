# XLSX 作者接口（图表/批注/超链接/表格/条件格式/验证/冻结窗格）

> [!NOTE]
> 本文记录 issue #102 的实现边界与验收口径。当前状态以代码和测试为准：`tests/unit/xlsx-authoring.test.mjs`、`tests/component/xlsx-authoring-ui.test.mjs`。

## 1. 目标

在 `@arcships/vue-xlsx` 上补一层非单元格对象的作者入口：功能区分组 + 右键菜单 + 独立对话框，同时把这些能力作为控制器命令公开，宿主可以完全替换内置 UI。每条命令必须满足：

1. 改动工作簿模型（工作表 XML / 图表资产），不是只改渲染；
2. 触发重新渲染；
3. 写入撤销历史；
4. 导出的 XLSX 中包含改动，重新打开仍然存在。

## 2. 为什么不能只依赖 WASM 模型

`@dukelib/sheets-wasm` 只提供单元格、样式、行列尺寸、合并和名称的写接口；批注、超链接、条件格式、数据表（ListObject）、冻结窗格和图表只有读接口。`saveXlsxBytes()` 的行为也不统一：

- 条件格式、超链接、冻结窗格、修复后的空引用会随保存保留；
- 批注、VML、表格部件通常被丢弃；
- 图表/绘图部件在部分文件里会被保留，并改名为 `drawing1.xml`，在另一些文件里被丢弃。

因此作者命令直接在**源压缩包**（`WorkbookImageAssets.archive`）上修改 OOXML：

- 工作表 XML：`<sheetViews><pane>`、`<hyperlinks>`、`<conditionalFormatting>`、`<dataValidations>`、`<legacyDrawing>`、`<tableParts>`；
- 包级部件：`xl/comments/*.xml`、`xl/drawings/vmlDrawing*.vml`、`xl/tables/*.xml`、`xl/charts/*.xml`、`xl/drawings/drawing*.xml`；
- 关系与内容类型：工作表 rels、绘图 rels、`[Content_Types].xml`。

`packages/xlsx-core/src/authoring/sheet-authoring.ts` 负责这些改动，并把解析后的新状态返回给控制器，用来即时刷新 `XlsxSheetData`（`comments`、`hyperlinks`、`conditionalFormatRules`、`dataValidations`、`freezePanes`）与 `tables`。

## 3. 导出与撤销

`createSavedWorkbookBytes()` 的链路现在是：

```
saveXlsxBytes()                 # 引擎提供权威单元格数据
→ sanitizeSavedWorkbookBytes()  # 修复二次转义的 XML 实体
→ mergeWorkbookSheetAuthoringAssets()  # 写入作者对象
→ mergeWorkbookImageAssets()           # 重建绘图 / 图表 / 图片部件
```

`mergeWorkbookSheetAuthoringAssets(savedBytes, sourceArchive)` 对每个工作表：

- 用引擎保存的单元格数据作为主体；
- 从源压缩包复制作者对象的“片段”：冻结 `<pane>`、`<hyperlinks>`、`<conditionalFormatting>`、`<dataValidations>`、`<legacyDrawing>`、`<tableParts>`，以及对应的 comments/VML/table 部件与关系；
- 按 OOXML 子元素顺序重新插入，避免 Excel 报“需要修复”；
- 不丢弃无法建模的条件格式规则（`cellIs`、`top10`、`expression` 等）；
- 没有任何作者内容时原样返回（不重新压缩）。

撤销历史本身就是这套字节的快照，所以撤销/重做天然覆盖作者对象：撤销会回到上一份字节，重做恢复作者对象。图表的 id/路径在恢复时按拆分后的绘图重新解析。

## 4. 逐项行为

### 4.1 插入图表

`insertWorkbookChart()`（`packages/xlsx-core/src/charts/chart-insert.ts`）生成：

- `xl/charts/chartAuthoring{n}.xml`：柱形（clustered column）、折线或饼图，包含分类/数值引用、标题、图例和轴；
- `xl/drawings/drawingAuthoring{n}.xml` + rels：`oneCellAnchor` 的 `graphicFrame`，默认放在区域下方两行；
- 工作表 rels 的 drawing 关系与 `<drawing r:id>`；
- 注册到 `sheetOrigins` 的绘图附件，导出时由图片合并逻辑统一重建绘图。

控制器随后调用 `loadWorkbookChartAssets()` 重新解析图表资产，因此插入后立即渲染；图表类型从图表 XML 读取（`resolveChartTypeFromTypeNode`），不再固定为柱形图。`updateChart` / `deleteChart` 继续服务“编辑图表”对话框（标题、删除）。

### 4.1.1 更改图表类型

`changeWorkbookChartType()`（`packages/xlsx-core/src/charts/chart-type-edit.ts`）在同一个图表部件上重写类型：

- 经典类型重建 `<c:plotArea>`，按目标族写入 `barChart`/`lineChart`/`pie3DChart`/`scatterChart`/`stockChart`/`surface3DChart` 等节点、轴（含隐藏的成交量轴）和 `view3D`；系列的名称公式、`spPr`、`dPt`、`dLbls`、标记与平滑设置从原节点克隆，因此颜色、数据标签和点样式不丢；
- chartEx 类型（treemap、sunburst、histogram、pareto、box & whisker、waterfall、funnel、filled map）生成 `cx:chartSpace`：`cx:chartData` 的 `strDim`/`numDim`、`cx:series@layoutId`、直方图 `cx:binning`、Pareto 的 `paretoLine` 系列；同时补 `chartStyle`/`chartColorStyle` 部件、关系和内容类型，且重复切换不会重复添加；
- 经典 ↔ chartEx 迁移时同步 drawing 关系类型、`a:graphicData@uri`、`cx:chart`/`c:chart` 引用、`mc:AlternateContent`（含 `mc:Fallback`）和图表内容类型；
- `loadWorkbookChartAssets()` 现在也会从压缩包解析 chartEx 部件（`parseChartExChartXml()`），所以导出的现代图表重新加载后仍可见；直方图分箱不会二次应用。

控制器入口为 `XlsxViewerController.changeChartType(id, type)`；右键图表时 `XlsxContextMenu` 用 `chartTypeOptions` 的族/类型子菜单打开“更改图表类型”，选择后重写部件并按 `chartPath` 重新选中图表。`chartTypeUnavailable()` 会禁用系列数量不匹配（如股票图）或只读的选项。

### 4.2 批注

- 单元格一条批注；写入/编辑/删除后重写 `xl/comments/commentAuthoring{n}.xml`；
- 同步生成 `xl/drawings/vmlDrawingCommentAuthoring{n}.vml`，并在工作表写入 `<legacyDrawing>`，Excel 才会显示批注标记；
- 删除最后一条批注时同时移除批注部件、VML、关系和 `<legacyDrawing>`；
- 作者名会写入 `<authors>`；没有作者时省略 `authorId`。

### 4.3 超链接

- `http(s)`、`mailto:` 等外部目标写成外部关系（`TargetMode="External"`）；
- `#Sheet!A1` 写成内部 `location`；
- 同一单元格只有一个超链接；删除后清理不再被引用的超链接关系。

### 4.4 条件格式

- 控制器把 `XlsxConditionalFormatDraft`（kind、`colors`、`iconSet`、`thresholds`）转换成持久规则；
- 阈值即 cfvo：`min` / `max` / `num` / `percent` / `percentile` / `formula`；
- 三色/两色色阶、数据条、3–5 个图标集都有默认值，列表顺序即优先级（1 = 最高）；
- 颜色在模型里保持与解析结果一致的 CSS `#rrggbb`，写入 OOXML 时转换为 `ST_HexColorRGB` 要求的 8 位 ARGB（`FFRRGGBB`）。主题色、索引色和自动色分别写成 `theme`/`indexed`/`auto`，不再把 CSS 十六进制直接塞进 `rgb`（LibreOffice/Excel 会把它读成黑色）；
- 编辑/删除只重写“可建模”的规则，其他规则原样保留。

### 4.5 表格（ListObject）

- `createTable()` 用选区（或显式 `range`）建表：生成 `xl/tables/tableAuthoring{n}.xml`、table 关系、`<tableParts>`、内容类型；
- 列名取首行显示值（`hasHeaders: false` 时使用 `Column N`），名称必须唯一且符合 Excel 标识符规则；
- 删除表格会移除部件、关系与 `<tableParts>`。

### 4.6 数据验证

- `setSelectionDataValidation()` 覆盖选区上的旧规则并写入 `<dataValidations>`；
- `list` 支持内联列表（自动加引号）、区域引用和公式；`showDropdown` 按 OOXML 反转写入；
- 传 `null` 删除选区上的规则。

### 4.7 冻结窗格

- `setFreezePanes({ row, col })` 在活动工作表写入 `<pane xSplit ySplit topLeftCell activePane state="frozen">`；
- `toggleFreezePanesAtActiveCell()` 在当前单元格上冻结；再点一次取消；
- 两个方向都为 0 等价于取消冻结。

### 4.8 页面设置（#171）

- `setPageSetup(input)` 把 `pageSetup`、`pageMargins`、`printOptions` 与可选的打印区域写成一次改动；
- 缩放由 `pageSetup.fitToPage` 选模式：`true` 写 `fitToWidth`/`fitToHeight` 与 `sheetPr/pageSetUpPr/@fitToPage`，`false` 写 `scale`；`fitTo* = 0` 是 Excel 的“自动页数”，解析器原样保留；
- 打印区域写入 `xl/workbook.xml` 的 `_xlnm.Print_Area`（`localSheetId` 作用域、`$A$1:$D$20` 绝对引用），`printAreas: []` 删除该定义；
- 只改传入的字段；未传入的页边距/纸张/DPI 保持工作簿原值，因此对话框打开再确定不会重置未展示的项；
- 与其它 authoring 命令一样：一条撤销记录、随导出保留、只读/受保护/Worker 只读下返回 `false`。

## 5. UI

- 功能区分组 `xlsx-ribbon-authoring`：图表、批注、超链接、表格、条件格式、数据验证、冻结开关、页面设置（文字按钮，打开 `XlsxPageSetupDialog`）；
- 右键菜单：插入/编辑批注、插入/编辑超链接、插入图表、条件格式、数据验证、创建/删除表格、冻结/取消冻结；
- `XlsxAuthoringDialog`：一个对话框按 `kind` 渲染各表单；`XlsxViewer` 与最小化的 `XlsxSheetSurface` 都会从右键菜单打开它；宿主也可以自行挂载；外壳自 0.8 起复用共享 `office-dialog` 表面（见 [design/xlsx-office-chrome-migration.md](design/xlsx-office-chrome-migration.md)），表单状态、校验与焦点行为不变；
- `XlsxPageSetupDialog`（#171）：页面设置的独立对话框，`kind="pageSetup"` 时由 `XlsxAuthoringDialog` 转交；字段取工作表当前值，摘要行调用核心分页显示页数/纸张/缩放，390px 单列不溢出；规格见 [design/xlsx-page-setup-reference.md](design/xlsx-page-setup-reference.md)；
- 所有文案走 `office-interaction` 消息目录（`xlsx.authoring.*`，en + zh-CN），不硬编码。

## 6. 限制

- 需要可修改的主线程工作簿：`readOnly`、受保护工作表（未启用 `allowProtectedSheetEditing`）或 Worker 只读大文件模式下命令返回 `false`/`null`；
- `skipXmlParsing` 模式没有解压后的源压缩包，作者命令不可用（与图表/条件格式渲染的限制一致）；
- 图表类型切换、图表类型子菜单（#139）与条件格式之外的高级规则编辑器（#21）不在本实现范围；
- 插入的图表使用固定 6" × 3.5" 尺寸，之后可拖动/缩放（沿用绘图对象交互）；
- 选区只有一个单元格时，“插入图表”回退到工作表已用区域；“创建表格”仍要求至少两行（含标题行），失败时对话框给出提示。
