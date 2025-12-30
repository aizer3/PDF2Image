# PDF2Image

一个基于 Python 的 PDF 转图片工具，支持交互式裁剪、多进程加速。

## 功能特点
- **交互式裁剪**：可视化调整裁剪区域。
- **多进程处理**：充分利用多核 CPU 性能，极速转换。
- **高清晰度**：支持自定义 DPI 设置。
- **现代化 UI**：基于 `customtkinter` 打造。

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
