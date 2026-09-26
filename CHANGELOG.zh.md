## 1.0.0 - 2026-09-26

### 新功能

- **图形界面**：根目录的 `OneNote导出工具.exe` 双击即用，选笔记本、分区、导出目录后点一下就跑。窗口图标与 `.exe` 统一
- **启动器**：极小的原生 C 程序（41 KB，无运行时依赖），负责用系统自带的 Windows PowerShell 拉起界面脚本；源码在 `scripts/launcher/`，可在 Windows 上重新编译
- **导出为 Markdown 文件夹树**：笔记本 / 分区组 / 分区 / 页面分别映射为目录层级，子页面嵌套为父页面同名的子文件夹
- **手写笔迹提取**：读取 `InkWord@recognizedText`，让手写内容能被文本检索
- **手绘图渲染**：`InkDrawing` 的 ISF 笔画数据经 WPF `StrokeCollection` 渲染成 PNG；标注与其底图合成为同一张图，保留笔画间的相对位置
- **图片内联为 data URI**：不旁挂 `Assets/` 目录，避免知识库按单文件入库时相对路径断链
- **增量导出**：跳过未变化的页面，`-Force` 可强制全量重导
- **命令行入口**：`scripts/sync-onenote.cmd`，支持列出笔记本、预览页面清单、按笔记本或分区筛选

### 文档

- README 说明输出结构、与上游版本的差异、已知限制
- 补充 SmartScreen 首次运行提示的处理方式

### 说明

本仓库 fork 自 [QR4X/obsidian-onenote-sync](https://github.com/QR4X/obsidian-onenote-sync)（MIT），
在保留其 OneNote COM 读取能力的基础上，为「导出目录供知识库导入」这一用途做了实质改造：
移除 Obsidian 插件壳、改为纯 Markdown 输出、新增手写与手绘处理、新增图形界面。
原始项目版权归 Finn / QR4X 所有。
