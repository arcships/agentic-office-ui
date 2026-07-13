# 文件格式支持范围与接入方案

> 状态：P0 已实现；精简 P1 已于 2026-07-13 完成
> 适用范围：XLSX、DOCX、PPTX、PDF 四类 View
> 目标：在不增加新的文档或图片渲染器的前提下，扩展现有 Office 文件家族的可预览格式，并统一文件识别、校验和错误处理

## 1. 结论

本阶段只扩展现有四类 View 能直接复用当前解析和渲染链路的文件格式：

- XLSX View：增加 Excel 二进制、宏工作簿和模板格式，并保留 CSV 支持。
- DOCX View：增加 Word 宏文档和模板格式。
- PPTX View：增加 PowerPoint 宏演示、放映和模板格式。
- PDF View：继续只支持 PDF。

本阶段明确不接入 Markdown、纯文本和任何图片格式，也不新增通用 Image View。

P0 完成后可以进入内部试用。正式对外声明新增格式支持前，只补充少量真实 Office 文件验证；不要求为每个扩展名建立完整测试矩阵。

## 2. 支持矩阵

### 2.1 XLSX View

| 文件格式 | 状态 | 接入方式 | 说明 |
|---|---|---|---|
| `.xlsx` | 已支持 | 现有工作簿解析链路 | 标准 Excel 工作簿 |
| `.xls` | 已支持 | 现有 WASM 工作簿解析链路 | 旧版 Excel 二进制工作簿 |
| `.csv` | 已支持 | CSV 文本转工作簿 | 需要继续处理编码、空文件和解析错误 |
| `.xlsb` | 已支持 | 按需转换为内存 XLSX 后复用现有解析链路 | 转换器只在 XLSB 输入时加载；不执行或保留宏，源文件下载仍使用原始字节 |
| `.xlsm` | 已支持 | 复用 XLSX OOXML 解析链路 | 只读取工作簿内容，不执行 VBA 宏 |
| `.xltx` | 已支持 | 复用 XLSX OOXML 解析链路 | 按普通工作簿预览模板内容 |
| `.xltm` | 已支持 | 复用 XLSX OOXML 解析链路 | 按普通工作簿预览，不执行 VBA 宏 |
| `.tsv`、`.tab` | 暂不接入 | — | 当前范围只保留 CSV，不扩展其他分隔文本格式 |
| `.ods`、Numbers | 不直接支持 | 接入方转换为 `.xlsx` 或 `.csv` | 需要独立解析器或服务端转换 |

### 2.2 DOCX View

| 文件格式 | 状态 | 接入方式 | 说明 |
|---|---|---|---|
| `.docx` | 已支持 | 现有 DOCX OOXML 解析链路 | 标准 Word 文档 |
| `.docm` | 已支持 | 复用 DOCX OOXML 解析链路 | 只读预览，不执行 VBA 宏 |
| `.dotx` | 已支持 | 复用 DOCX OOXML 解析链路 | 按普通文档预览模板内容 |
| `.dotm` | 已支持 | 复用 DOCX OOXML 解析链路 | 按普通文档预览，不执行 VBA 宏 |
| `.md` | 明确不接入 | — | 不提供 Markdown 到 `DocModel` 的转换 |
| `.txt` | 明确不接入 | — | 不提供纯文本到 `DocModel` 的转换 |
| `.doc`、`.rtf`、`.odt`、Pages | 不直接支持 | 接入方转换为 `.docx` 或 `.pdf` | 旧格式和第三方格式不进入浏览器端解析链路 |

### 2.3 PPTX View

| 文件格式 | 状态 | 接入方式 | 说明 |
|---|---|---|---|
| `.pptx` | 已支持 | 现有 PPTX OOXML 解析链路 | 标准 PowerPoint 演示文稿 |
| `.pptm` | 已支持 | 复用 PPTX OOXML 解析链路 | 不执行 VBA 宏，并报告宏已忽略 |
| `.ppsx` | 已支持 | 复用 PPTX OOXML 解析链路 | 在现有浏览或播放模式中打开 |
| `.ppsm` | 已支持 | 复用 PPTX OOXML 解析链路 | 在现有浏览或播放模式中打开，不执行 VBA 宏 |
| `.potx` | 已支持 | 复用 PPTX OOXML 解析链路 | 按普通演示文稿预览模板内容 |
| `.potm` | 已支持 | 复用 PPTX OOXML 解析链路 | 按普通演示文稿预览，不执行 VBA 宏 |
| `.ppt`、`.odp`、Keynote | 不直接支持 | 接入方转换为 `.pptx` 或 `.pdf` | 需要独立解析器或服务端转换 |

### 2.4 PDF View

| 文件格式 | 状态 | 接入方式 | 说明 |
|---|---|---|---|
| `.pdf` | 已支持 | 现有 PDFium 渲染链路 | 保持严格的 PDF 文件头和来源校验 |
| `.png`、`.jpg`、`.jpeg`、`.webp`、`.gif`、`.bmp` | 明确不接入 | — | 不把图片包装成单页 PDF，也不新增 Image View |
| `.tif`、`.tiff` | 明确不接入 | — | 不新增多页图片文档能力 |
| SVG、HEIC、HEIF | 明确不接入 | — | 不进入 PDF View |
| XPS、OXPS、DjVu、EPUB | 不直接支持 | 接入方转换为 `.pdf` | PDF View 不承担其他分页文档格式的解析 |

## 3. 接入边界

格式扩展由接入层和各格式 Controller/Runtime 完成，Surface 继续只负责渲染，不感知文件扩展名、MIME 或来源类型。

已在 `@arcships/office-runtime` 增加统一来源识别能力：

```ts
type OfficeSourceFormat =
  | "xlsx"
  | "xls"
  | "xlsb"
  | "xlsm"
  | "xltx"
  | "xltm"
  | "csv"
  | "docx"
  | "docm"
  | "dotx"
  | "dotm"
  | "pptx"
  | "pptm"
  | "ppsx"
  | "ppsm"
  | "potx"
  | "potm"
  | "pdf"

type SourceFormatDetection = {
  family: "xlsx" | "docx" | "pptx" | "pdf"
  format: OfficeSourceFormat
  confidence: "high" | "medium" | "low"
  evidence: string[]
}
```

识别时依次使用以下证据：

1. 文件头和容器类型，例如 PDF `%PDF-`、ZIP/OOXML、OLE 二进制容器。
2. OOXML 包中的 `[Content_Types].xml` 和核心部件路径。
3. 文件名扩展名。
4. MIME，仅作为辅助信息，不单独作为可信依据。

不能只根据扩展名放行文件。OOXML 文件以包内容确定所属 View；不属于目标文件家族时直接拒绝。扩展名、MIME 和实际内容的完整冲突策略不作为本阶段发布阻断项。

## 4. 宏与模板策略

宏格式只提供安全的只读预览能力：

- 永不执行 VBA、ActiveX、外部脚本或嵌入程序。
- 宏内容统一忽略；PPTX 当前通过 `warnings` 提示，XLSX/DOCX 是否增加同类提示不阻断发布。
- 第一阶段不承诺对宏文件进行原格式编辑和保存。
- 如果未来允许编辑后导出，必须明确选择“保留宏”或“移除宏”，不能静默改扩展名或丢失宏。

模板和放映格式统一复用现有 View：

- `.xltx`、`.xltm` 按普通工作簿预览。
- `.dotx`、`.dotm` 按普通文档预览。
- `.potx`、`.potm` 按普通演示文稿预览。
- `.ppsx`、`.ppsm` 进入现有 PPTX 浏览或播放模式，不额外实现 Office 桌面端的自动放映启动语义。

## 5. 实施顺序

### P0：统一识别和放开已具备解析能力的格式（已完成）

1. 增加统一 `detectSourceFormat`。
2. 扩展 OOXML 主内容类型白名单和压缩包校验。
3. 接入 `.xlsb`、`.xlsm`、`.xltx`、`.xltm`。
4. 接入 `.docm`、`.dotx`、`.dotm`。
5. 接入 `.pptm`、`.ppsx`、`.ppsm`、`.potx`、`.potm`。
6. 更新各 View 的文件选择器、错误提示、README 和使用指南。

### P1：最小发布验证

P1 只作为正式发布前的一次性轻量验收，不进入日常开发或每次改动的固定测试流程，也不为内部试用增加阻塞：

1. 使用 Office 实际生成的代表文件验证 XLSB/XLSM、DOCM、PPTM/PPSX。
2. 在对应 View 中各完成一次浏览器打开检查。
3. 确认真实宏文件能够预览，且没有宏、ActiveX 或脚本执行路径。
4. 使用代表性的损坏文件和加密文件确认加载会失败且界面可恢复。

本次验收结果（2026-07-13）：

| View | 代表文件 | 结果 |
|---|---|---|
| XLSX View | XLSB、XLSM | 均进入 `ready`；XLSB 通过按需转换进入现有工作簿管线 |
| DOCX View | DOCM | 进入 `ready` |
| PPTX View | PPTM、PPSX | 均成功加载 12 页 |
| XLSX View | 含 `vbaProject.bin` 的真实 XLSM | 成功预览；浏览器无弹窗、无宏执行，宏部件未进入解析或执行链路 |
| XLSX View | 损坏 XLSX、加密 XLSX | 均返回 `INVALID_WORKBOOK`；随后重新加载正常 XLSM 可恢复到 `ready` |

真实宏样本使用 [SheetJS 官方 VBA 示例](https://docs.sheetjs.com/vba/SheetJSVBAFormula.xlsm)，验收时确认包内存在 `xl/vbaProject.bin`，SHA-256 为 `12ab43fff0d40d977952916c51aafca5a0fcb170036570a2a1f588b11cedcadc`。其他代表文件仅作为本次临时验收输入，不纳入仓库 fixture。

以下项目不纳入当前 P1，可按真实接入问题再补：

- 为所有模板和放映扩展名分别维护真实 fixture。
- 扩展名、MIME、最终 URL 和实际内容的全组合冲突矩阵。
- 所有浏览器、所有来源类型的重复回归。
- 统一全部格式的细粒度错误码和“宏已忽略”诊断。
- 与本次格式接入无关的性能、压力和完整发布测试。

## 6. 验收标准

### 6.1 P0 内部可用

- OOXML 包内容能够被识别为正确的文件家族和具体格式。
- 新格式能够复用现有解析、渲染和资源限制链路。
- 宏文件不会执行任何宏或活动内容。
- View 和 Surface 的公开 API 不因文件扩展名增加而分叉。
- README、使用指南、文件选择器 `accept` 和支持矩阵保持一致。

### 6.2 正式发布门槛（已满足）

- 完成第 5 节的精简 P1。
- 代表性的真实 Office 文件可以在对应 View 中打开。
- 代表性的损坏或加密文件会失败，且不会导致 View 无法继续加载其他文件。
- 不存在宏、ActiveX、外部脚本或嵌入程序的执行入口。

精简 P1 在正式验收时执行并通过后即可对外声明支持；后续普通改动不要求重复执行，除非格式加载链路再次发生实质变化。

## 7. 明确不在本阶段范围内

- Markdown 和纯文本预览。
- 所有图片文件预览，包括单页图片和多页 TIFF。
- 通用 Image View 或“图片转 PDF”适配层。
- 浏览器端解析 `.doc`、`.ppt`、`.ods`、`.odt` 等旧格式或第三方格式。
- 执行宏、编辑宏或承诺保留宏后的原格式导出。
- 自动调用第三方转换服务；如需要格式转换，由接入方在文件进入本项目之前完成。
