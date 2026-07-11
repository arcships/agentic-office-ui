# PPTX 播放验收素材

这里存放正式发布前需要逐项复核的 PPTX 素材，清单见 `manifest.json`。13 份素材已经登记文件哈希和浏览器事件基准。仓库上一级另有可重复生成的组合回归素材 `playback-controlled.pptx`，用于自动测试。

每个 `ready` 素材必须由桌面版 PowerPoint 保存，并具有可确认的来源、许可、SHA-256 和预期结果。临时修改压缩包得到的派生文件可以用于调试解析分支，但不能单独证明与 PowerPoint 的播放结果一致。

不要把敏感业务文件放入本目录。业务文件只在受控环境中使用，并在探索或验收记录中保存哈希、页数、大小和结果。

每个正式素材都有同名目录，保存不受版权限制的事件清单和说明：

```text
timing-click-with-after.pptx
timing-click-with-after/
├── expected-events.json
└── README.md
```

重新生成浏览器事件基准使用 `pnpm capture:pptx-baselines`。网页实现以浏览器中的对象、时间、事件、画面状态和资源释放检查为准；PowerPoint 画面对照只在独立兼容复核时另行保存，不作为日常开发前置步骤。
