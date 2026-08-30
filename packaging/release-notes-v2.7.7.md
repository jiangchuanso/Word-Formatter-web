# v2.7.7

* **新增行距单位选择**：题目、副标题、正文和表格行距支持磅值与 Word 倍数两种模式；倍数模式默认 1.0，并允许手动输入其他倍数。
* **新增段落缩进单位选择**：左右缩进支持厘米或字符，首行缩进继续使用 Word 字符缩进语义，默认 2 字符。
* **新增标题加粗选项**：支持对文章题目、一级标题和二级标题选择强制加粗。
* **重排配置界面**：配置项整合到单页可滚动区域，标题设置排在正文之前；加载、保存、保存为默认和恢复内置默认按钮固定在配置区下方。
* **保持默认输出兼容**：默认仍使用旧版磅值行距和厘米缩进配置；默认样例输出与 v2.7.4 基准文档的关键 Word XML 保持一致。
* **保持跨平台兼容**：继续使用 Tkinter/ttk，不引入 PySide6，兼顾 Windows、macOS 和 Kylin/Linux 的 DPI、系统字体及低分辨率环境。
* **完善 Kylin 打包命令**：补充 x86_64 环境依赖、`appimagetool` 下载、AppImage 构建、普通 Linux 可执行文件备用命令，以及通过 Miniforge/conda-forge 固定 Python 3.12 的 Kylin V10 SP1 Docker 构建镜像。
* **新增 UOS V20 打包支持**：新增统一的 ARM64/x86_64 Dockerfile、Linux/macOS Bash 构建脚本和 Windows PowerShell 构建脚本；Release 同时提供 UOS ARM64 AppImage 与不依赖 FUSE 的普通可执行文件。
* **修复 Windows 打包的 Tcl/Tk 冲突**：构建时按 `_tkinter` 实际版本选择匹配的 Tcl/Tk 数据，避免 Anaconda 环境中 8.6.9 与 8.6.15 混装导致 EXE 无法启动。
* **同步 doc-format Skill**：配置字段、格式化核心、CLI 配置说明、测试和文档与主程序保持一致。
