# PDF2Image (v1.5.11)

一个基于 Python 的现代化 PDF 转图片工具，支持交互式裁剪、多进程加速、文件拖拽等功能。

## 功能特点
- **交互式裁剪**：可视化调整裁剪区域，支持拖拽调整和步进器微调。
- **多进程处理**：基于 `ProcessPoolExecutor` 实现，充分利用多核 CPU 性能，极速转换。
- **文件拖拽**：支持将 PDF 文件直接拖入窗口进行处理。
- **实时预览**：支持选择任意页码进行预览，并在预览图上直观查看裁剪效果。
- **高清晰度**：预设 72, 150, 300, 600 DPI，满足不同场景需求。
- **现代化 UI**：基于 `customtkinter` 打造，支持系统主题跟随。
- **停止机制**：支持在转换过程中随时停止任务。

## 快速开始
1. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```
2. 运行程序：
   ```bash
   python pdf_to_img.py
   ```

## 打包说明
项目包含多个 `.spec` 文件，推荐使用最新版本：
```bash
python -m PyInstaller PDFToImage_v15.spec
```
