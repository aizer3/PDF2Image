# PDF2Image (Modern UI) 深度技术文档

本项目是一个基于 Python 开发的现代化 PDF 转图片工具。本文档旨在深入解析项目的核心代码实现、架构设计及关键技术点，帮助开发者理解其内部工作原理。

## 1. 核心架构设计

项目采用 **事件驱动 (Event-Driven)** 和 **多线程 (Multi-threading)** 架构。UI 线程负责响应用户操作和渲染界面，后台线程负责耗时的 PDF 解析与图像处理，确保在高负载任务下界面依然保持流畅。

## 2. 关键代码详解

### 2.1 资源路径兼容性解决方案
在开发环境和打包后的运行环境中，资源文件（如图标）的存放路径不同。
```python
def resource_path(relative_path):
    """ 获取资源的绝对路径，兼容开发环境和 PyInstaller 打包环境 """
    try:
        # PyInstaller 打包后会将资源释放到 sys._MEIPASS 临时目录
        base_path = sys._MEIPASS
    except Exception:
        # 开发环境下使用当前脚本所在目录
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
```
**解读**：此函数是打包 GUI 程序的必备技巧。`sys._MEIPASS` 是 PyInstaller 在单文件模式下运行时的特有属性。

### 2.2 PDF 渲染与 DPI 控制
项目使用 `PyMuPDF` (fitz) 作为底层引擎。核心在于 `fitz.Matrix` 的运用。
```python
# 计算缩放比例
# 72 是 PDF 的标准点数(point)，dpi_val 是用户选择的目标分辨率
zoom = dpi_val / 72 
mat = fitz.Matrix(zoom, zoom) # 创建缩放矩阵

# 加载页面并渲染为像素图
page = doc.load_page(i)
pix = page.get_pixmap(matrix=mat)
```
**解读**：
- `zoom` 决定了输出图片的尺寸和清晰度。例如选择 300 DPI 时，图片像素量将是 72 DPI 的约 17 倍。
- `get_pixmap` 直接在内存中生成位图，避免了频繁的磁盘 I/O。

### 2.3 动态像素级裁剪逻辑
裁剪逻辑需要同时考虑 DPI 缩放和边界越界保护。
```python
# 像素转换与边界限制
left = min(c_left, width - 1)
top = min(c_top, height - 1)
right = max(left + 1, width - c_right)
bottom = max(top + 1, height - c_bottom)

# 使用 PIL 进行物理裁剪
cropped_img = img.crop((left, top, right, bottom))
```
**解读**：
- 用户输入的裁剪值（像素）是直接作用于渲染后的位图上的。
- 代码通过 `min/max` 逻辑确保裁剪区域始终有效，防止因输入数值过大导致程序崩溃。

### 2.4 多进程并行转换与停止机制
为了彻底突破 Python GIL 的限制，v1.4 版本引入了多进程架构：
```python
# 使用 ProcessPoolExecutor 调度多核 CPU
max_workers = min(os.cpu_count() or 4, 8)
executor = ProcessPoolExecutor(max_workers=max_workers)

# 将每一页的处理任务分发到独立进程
future = executor.submit(process_page_task, pdf_path, i, zoom, crop_params, output_path)
```
**解读**：
- **性能飞跃**：转换速度提升了约 3-5 倍（取决于 CPU 核心数）。
- **内存安全**：每个子进程在处理完单页后会自动释放内存，彻底解决了主进程内存累积的问题。
- **稳定性**：子进程崩溃不会导致主程序挂掉。

### 2.5 现代化交互：文件拖拽支持
引入 `windnd` 库实现原生 Windows 文件拖拽：
```python
import windnd
windnd.hook_dropfiles(self, self.on_file_drop)
```
**解读**：用户只需将 PDF 拖入窗口即可自动填充路径，极大优化了操作链路。

### 2.6 预览窗口的自适应缩放
预览窗口通过计算缩放系数（Scale）来适配不同尺寸的屏幕。
```python
# 限制预览图大小，适应屏幕 80% 的尺寸
screen_w = self.winfo_screenwidth() * 0.8
screen_h = self.winfo_screenheight() * 0.8

scale = min(screen_w/img_w, screen_h/img_h, 1.0)
if scale < 1.0:
    new_size = (int(img_w * scale), int(img_h * scale))
    pil_img = pil_img.resize(new_size, Image.Resampling.LANCZOS)
```
**解读**：使用 `LANCZOS` 重采样滤镜是为了在缩小图片时尽可能保留细节，方便用户确认裁剪范围。

## 3. 打包技术内幕 (`.spec` 配置)

打包过程中，最复杂的环节是 `customtkinter` 的资源收集。
```python
from PyInstaller.utils.hooks import collect_all

# 自动收集 customtkinter 的所有依赖（包括 json 主题文件和图标）
tmp_ret = collect_all('customtkinter')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
```
**解读**：`customtkinter` 并非标准的 tkinter 库，它依赖大量的外部数据文件。手动添加极其繁琐，`collect_all` 钩子函数极大简化了这一过程。

---
*文档版本: v1.2 (高性能拖拽版)*
*更新日期: 2025-12-22*