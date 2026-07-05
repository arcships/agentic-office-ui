# shadcn/components 浏览器验收目标与方法

## 目标

按 shadcn/ui + Vue Extend components 的设计与交互标准，在共享浏览器中对 Vue Extend UI demo 全页面进行实际视觉与交互验收；逐页覆盖真实物料、全部组件状态、键盘/鼠标/上传/拖拽/响应式行为；发现视觉或交互问题立即修复并复测，直到用户逐页确认验收通过。

## 验收原则

1. 以共享浏览器中的实际界面为准。
2. 以真实用户交互为准，覆盖鼠标、键盘、上传、拖拽、滚动、hover、focus。
3. 以 shadcn/ui 的视觉系统为界面标准：token、间距、圆角、边框、阴影、层级、状态反馈、可访问性。
4. 以 Vue Extend components 的组件行为为功能标准：真实物料、真实状态、真实错误、真实交互。
5. 每页完成后进入用户确认点，用户确认后再进入下一页。
6. 工程门禁继续保留：物料校验、typecheck、build 作为最终底线。

## 验收范围

页面范围：

- Home：`/#/`
- DOCX Viewer：`/#/docx-viewer`
- DOCX Editor：`/#/docx-editor`
- XLSX Viewer：`/#/xlsx-viewer`
- PDF Viewer：`/#/pdf-viewer`
- Components：`/#/components`

组件范围：

- DocxViewer
- DocxEditor
- XlsxViewer
- PdfViewer
- SignaturePad
- FileUpload
- FileThumbnail
- BoundingBoxCitations
- LayoutBlocks
- Spinner
- Tooltip
- Home navigation cards
- API compatibility table

物料范围：

- DOCX：正常文档、表格文档、图片文档、中文混排、损坏文件
- XLSX：多 sheet、销售表、大表、图表/图片、损坏文件
- PDF：文本 PDF、扫描 PDF、旋转 PDF、大页数 PDF、损坏文件
- PNG：OCR/引用组件图片
- JSON：field citations、OCR layout、manifest

## shadcn/ui 视觉验收标准

每页都按以下标准检查：

| 项目 | 标准 |
|---|---|
| Layout | 页面主体宽度合理，内容分组清晰，卡片与 viewer shell 对齐 |
| Spacing | 标题、副标题、controls、status、content 区间距一致 |
| Radius | card、button、input、select、table、upload zone 圆角统一 |
| Border | 边框颜色与强度符合 shadcn muted/border 风格 |
| Background | 页面、card、muted 区域、viewer surface 层级清晰 |
| Typography | 标题、说明、表格、状态文本大小和权重合理 |
| Button | default/secondary/ghost/icon/disabled/hover/active 状态清晰 |
| Input/Select | 高度、边框、focus-visible、disabled 状态一致 |
| Table/Grid | header、cell、selected、hover、overflow 呈现合理 |
| Focus | Tab 可见焦点明确，键盘操作路径连续 |
| Error | 错误态清楚、可读、位置稳定，语气具体 |
| Empty/Loading | 空态和加载态占位合理，布局稳定 |

## 实际交互验收方法

### 通用步骤

1. 打开共享浏览器目标页面。
2. 观察首屏布局与视觉层级。
3. 按页面 checklist 执行真实交互。
4. 使用键盘检查 Tab、Enter、Escape、Arrow、focus-visible。
5. 切换到 4 个视口检查响应式。
6. 记录发现的问题。
7. 修复问题。
8. 回到共享浏览器复测。
9. 等待用户确认该页验收通过。

### 视口矩阵

| 视口 | 用途 |
|---|---|
| 1440×900 | 桌面默认体验 |
| 1280×720 | 笔记本低高度体验 |
| 768×1024 | 平板体验 |
| 390×844 | 手机体验 |

响应式通过标准：

- 页面级无横向滚动。
- 宽内容使用内部滚动容器。
- toolbar/control 自动换行且可点击。
- 移动端触控目标足够大。
- 卡片、表格、viewer shell 保持清晰层级。
- 关键状态文本可读。

## 页面验收 checklist

### Home

视觉：

- 页面标题、副标题清晰。
- 导航卡片间距、圆角、hover、状态 badge 统一。
- API compatibility table 与 shadcn table 风格一致。

交互：

- hover 每个卡片。
- Tab 遍历卡片链接。
- 点击进入每个 route。
- 返回 Home 后状态保持正确。

用户确认点：Home 页面视觉与导航体验通过。

### DOCX Viewer

视觉：

- controls、status、viewer shell、错误态符合 shadcn 风格。
- 文档页面在 viewer 中居中或可滚动，边界清楚。
- mobile 下宽文档通过内部滚动呈现。

交互：

- 切换所有 DOCX 样本。
- 上传 DOCX 文件。
- 加载 corrupted DOCX 并检查错误态。
- 滚动文档内容。
- 检查 toolbar/status 文案与文件名。
- Tab 到上传控件和按钮。

用户确认点：DOCX Viewer 的浏览、上传、错误态、滚动体验通过。

### DOCX Editor

视觉：

- 编辑器 shell、toolbar、正文 surface 层级清晰。
- contenteditable focus-visible 可见。
- toolbar button active/disabled 状态明确。

交互：

- 点击正文段落。
- 输入中文、英文、数字混排。
- 使用 Bold。
- 使用 Undo / Redo。
- 切换 Theme。
- Tab 遍历 toolbar 与正文。
- Escape 后焦点表现稳定。

用户确认点：DOCX Editor 的编辑、格式、撤销/重做、主题体验通过。

### XLSX Viewer

视觉：

- toolbar、sheet tabs、grid header、cell、selection outline 统一。
- 表格宽内容在内部滚动区域中呈现。
- readonly/error 状态清楚。

交互：

- 切换所有 workbook 样本。
- 切换 sheet tabs。
- 单击 cell 选择。
- 双击 cell 编辑。
- Enter 提交。
- Escape 取消。
- Tab 提交并移动。
- Undo / Redo。
- Zoom。
- Read only 阻止编辑。
- Upload XLSX。
- Download。
- 加载 corrupted XLSX 并检查错误态。

用户确认点：XLSX Viewer 的表格浏览、编辑、键盘、状态体验通过。

### PDF Viewer

视觉：

- toolbar、thumbnail rail、PDF frame、status/error 呈现一致。
- 当前页、缩略图、搜索反馈清楚。
- mobile 下 PDF frame 可读且可滚动。

交互：

- 切换所有 PDF 样本。
- 上传 PDF 文件。
- Next / Previous 翻页。
- 点击 thumbnail。
- Zoom in / out。
- Rotate。
- Search。
- Download。
- 加载 corrupted PDF 并检查错误态。
- Tab 遍历 toolbar。

用户确认点：PDF Viewer 的浏览、翻页、缩放、搜索、错误态体验通过。

### Components

视觉：

- 每个组件 demo 卡片风格统一。
- controls、状态文本、错误文本、空态对齐一致。
- 图片、overlay、tooltip、spinner 与 shadcn 风格协调。

交互：

- SignaturePad：画签名、检查 empty、清空。
- FileUpload：点击选择、拖拽、非法类型、超大文件、disabled。
- FileThumbnail：检查 PDF/DOCX/XLSX/PNG/unknown 类型缩略图。
- BoundingBoxCitations：点击字段，检查 selected field。
- LayoutBlocks：点击 OCR block，检查 selected block，检查空态。
- Spinner：检查尺寸和颜色。
- Tooltip：hover/focus 显示，移出消失。
- Tab 遍历可交互元素。

用户确认点：Components 页面所有组件视觉与交互体验通过。

## 问题记录格式

每个问题使用以下格式记录：

```text
页面/组件：
视口：
交互步骤：
实际表现：
期望表现：
问题类型：视觉 / 交互 / 响应式 / a11y / 状态 / 数据
修复文件：
复测结果：
用户确认：待确认 / 已确认
```

## 验收状态表

| 页面 | 浏览器实操 | shadcn 视觉 | 组件交互 | 键盘/a11y | 响应式 | 用户确认 |
|---|---|---|---|---|---|---|
| Home | 待验收 | 待验收 | 待验收 | 待验收 | 待验收 | 待确认 |
| DOCX Viewer | 待验收 | 待验收 | 待验收 | 待验收 | 待验收 | 待确认 |
| DOCX Editor | 待验收 | 待验收 | 待验收 | 待验收 | 待验收 | 待确认 |
| XLSX Viewer | 待验收 | 待验收 | 待验收 | 待验收 | 待验收 | 待确认 |
| PDF Viewer | 待验收 | 待验收 | 待验收 | 待验收 | 待验收 | 待确认 |
| Components | 待验收 | 待验收 | 待验收 | 待验收 | 待验收 | 待确认 |

## 最终完成条件

1. 所有页面在共享浏览器中完成实操验收。
2. 所有组件完成真实交互验收。
3. 所有 shadcn 视觉问题修复并复测。
4. 所有响应式断点通过。
5. 浏览器 page errors 与业务 console errors 清零。
6. 用户逐页确认验收通过。
7. 物料校验通过。
8. `pnpm typecheck` 通过。
9. `pnpm build` 通过。
