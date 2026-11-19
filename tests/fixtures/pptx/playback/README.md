# PPTX 播放验收素材

这里存放正式发布前需要逐项复核的 PPTX 素材，清单见 `manifest.json`。13 份 `ready` 素材已经登记文件哈希和浏览器事件基准；另有 4 份 `controlled` 素材覆盖第二版新增分支，尚未送入桌面 PowerPoint 重新保存。仓库上一级另有可重复生成的组合回归素材 `playback-controlled.pptx`，用于自动测试。

每个 `ready` 素材必须由桌面版 PowerPoint 保存，并具有可确认的来源、许可、SHA-256 和预期结果。临时修改压缩包得到的派生文件可以用于调试解析分支，但不能单独证明与 PowerPoint 的播放结果一致。

`controlled` 素材由 `scripts/fixtures/generate-pptx-v2-fixtures.mjs` 从 `text-by-paragraph.pptx` 确定性派生，登记 SHA-256、焦点页、预期能力码和 XML 标记：

```text
text-by-word.pptx     按词 charRg、bldP
chart-series.pptx     oleChartEl 整图/按系列、bldGraphic
table-cells.pptx      表格 pRg 逐单元格、bldGraphic
transition-sound.pptx 切换声音 sndAc、三维 cube 切换
```

它们只证明解析、目标编译和能力报告分支；升级为 `ready` 前必须由桌面版 PowerPoint 另存，并补齐事件顺序、截图与浏览器基准。重新生成命令：

```bash
node scripts/fixtures/generate-pptx-v2-fixtures.mjs
```

不要把敏感业务文件放入本目录。业务文件只在受控环境中使用，并在探索或验收记录中保存哈希、页数、大小和结果。

每个正式素材都有同名目录，保存不受版权限制的事件清单和说明：

```text
timing-click-with-after.pptx
timing-click-with-after/
├── expected-events.json
└── README.md
```

重新生成浏览器事件基准使用 `pnpm capture:pptx-baselines`。网页实现以浏览器中的对象、时间、事件、画面状态和资源释放检查为准；PowerPoint 画面对照只在独立兼容复核时另行保存，不作为日常开发前置步骤。
