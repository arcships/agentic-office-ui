# PPTX 正式播放验收记录

日期：2026-07-11

## 1. 环境

- macOS 26.5.1，Apple Silicon；
- Node.js 22.13.0，pnpm 9.0.6；
- Python Playwright 1.50.0；
- Chromium、Firefox、WebKit 均使用本机 Playwright 安装版本；
- Microsoft PowerPoint 16.110.3（16.110.26070318）。

## 2. 受控播放素材

`tests/fixtures/pptx/playback-controlled.pptx` 的 SHA-256 为 `0e98ea7bd47b8c5150a25d11836d2da1431a894b7634897c750d405f7a99f19a`。素材由 `scripts/fixtures/generate-pptx-playback-fixtures.py` 生成，覆盖：

- 普通单击、对象单击、前后步和重建；
- 有限重复、无限重复、自动往返、加减速和透明度强调；
- 直线运动、旋转、对象擦除、淡化、页面推进和擦除，以及强身份 Morph；
- 自动换页、隐藏页、内部跳页和外部链接请求；
- 音视频占位替换、播放、暂停、继续、裁剪、循环、音量、媒体书签和对象地址释放。
- 同一对象同一属性同时写入时的稳定裁决，以及三段文字逐次单击显示。

`pnpm test:pptx:blackbox` 在 Chromium、Firefox 和 WebKit 中全部通过。三种浏览器均满足：切换中只有一个旧页层，结束后旧页层归零；暂停期间中央时钟和媒体时间不推进；有限和无限重复能停止；书签只在正向越过时触发；自动换页到达目标页；冲突轨道结果唯一；三段文字依次显示；隐藏页按顺序跳过且显式跳转可进入；组件不直接打开外部链接。

正式 demo 的 Vue 演示页也在三种浏览器中完成了真实文件上传、搜索、缩放、浏览/演示切换、按钮下一步、键盘下一步、用户操作进入和退出全屏、回到浏览模式的完整操作。效果开始和结束事件按各自时间点检查，误差不超过 100 毫秒；上一步、重置和再次播放能重建相同状态；快速连续跳页只执行首个已开始操作和最后一个待执行操作；重复对象键会产生 `TARGET_AMBIGUOUS`，两个元素都保持静态不被动画修改；模拟浏览器拒绝媒体自动播放后，控制器进入 `blocked`，一次恢复操作可继续播放。

组合素材连续生成两次的文件哈希一致，并已在本机桌面版 Microsoft PowerPoint 中直接打开。12 页均被识别，页面切换和动画标记存在，未出现“修复演示文稿”提示。该检查证明组合回归文件结构可被 PowerPoint 接受；它仍不替代 `tests/fixtures/pptx/playback/manifest.json` 中 13 份逐项画面对照素材。

## 3. PowerPoint 正式素材

`tests/fixtures/pptx/playback/manifest.json` 中 13 份素材状态为 `ready`。每份素材都已登记来源、许可、SHA-256、PowerPoint 版本、焦点页和浏览器事件基准。

`pnpm capture:pptx-baselines` 已对 13 份素材逐一运行 Chromium 播放模型，生成稳定的对象、时间安排、事件顺序、断言和能力报告。开发准备检查会验证：素材确实是 PPTX 压缩包、文件哈希与清单一致、事件统计自洽、说明文件存在。PowerPoint 截图不参与网页实现的自动验收；正式公开发布前，再由未参与实现的人按运行手册做一次独立兼容复核。

PowerPoint 另存后会把部分嵌入视频的旧关系改为外部 `NULL`，真实媒体关系保存在 `p14:media`。浏览器解析器现在优先使用非外部的 `p14:media` 关系；正式媒体素材已加入 Chromium 黑盒回归，防止再次退回错误关系。

## 4. 性能

测试文件为不入库的真实业务操作手册 `02.组织员工用户操作手册.pptx`：

- 页数：79；
- 文件大小：43,370,462 字节；
- SHA-256：`4602219999368ce0620e53f5ff08a12e55b1ec58b3d3fa6364578cc1e6a23fd4`；
- 三轮全新 Chromium 进程的来源读取、解包、建模和首屏合计：173.3、170.4、170.3 毫秒；
- 受控文件的播放控制器编译：不超过 14 毫秒；
- 下一步启动页面切换：不超过 6 毫秒；
- 暂停响应：测试采样误差内为 0 毫秒；
- 当前页重置最大值：0.4 毫秒。

这些数值低于兼容与验收文档的 2 秒、100 毫秒和 200 毫秒门槛。页面切换的可见持续时间仍服从 PPTX 中的原始时长，6 毫秒指控制器响应和双层舞台准备时间。

## 5. 内存和释放

Chromium 通过开发者协议强制回收后记录：

- 首次播放前 JavaScript 堆：5,267,568 字节；
- 连续重播 10 次后：5,542,688 字节；
- 增长：275,120 字节，低于 20 MiB 门槛；
- 会话销毁后：5,251,040 字节；
- 舞台残留子元素：0；
- 媒体对象地址在销毁后不可再次读取。

本轮当前页重置最大值为 0.3 毫秒。

## 6. 包消费

`scripts/ci/verify-pptx-consumer.mjs` 会对 `@arcships/pptx-core` 和 `@arcships/vue-pptx` 执行真实打包，在工作区外安装压缩包，并完成 TypeScript 检查和 Vite 正式构建。当前结果通过。

## 7. 固定命令

```bash
pnpm typecheck:pptx
pnpm test:pptx
pnpm test:pptx:blackbox
pnpm capture:pptx-baselines
```

大文件不进入仓库。正式复测时按记录的哈希确认同一份业务文件，再运行 `tests/blackbox/pptx_playback_model.py --performance-check`。`scripts/fixtures/expand-pptx-performance.py` 只在缺少 79 页非敏感材料时用于构造压力测试文件，不替代上述真实业务基线。
