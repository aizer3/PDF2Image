import customtkinter as ctk
from tkinter import filedialog, messagebox, Canvas
import fitz  # PyMuPDF
import os
import threading
from PIL import Image, ImageTk
import io
import sys
import windnd
from concurrent.futures import ProcessPoolExecutor
import multiprocessing

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def process_page_task(pdf_path, page_index, zoom, crop_params, output_path):
    """
    独立进程执行的单页处理函数
    """
    try:
        import fitz
        from PIL import Image
        import io
        
        # 显式打开文档（每个进程独立打开）
        doc = fitz.open(pdf_path)
        page = doc.load_page(page_index)
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        
        # 转换为 PIL Image 进行裁剪
        img_data = pix.tobytes("png")
        img = Image.open(io.BytesIO(img_data))
        width, height = img.size
        
        c_left, c_top, c_right, c_bottom = crop_params
        left = min(c_left, width - 1)
        top = min(c_top, height - 1)
        right = max(left + 1, width - c_right)
        bottom = max(top + 1, height - c_bottom)
        
        cropped_img = img.crop((left, top, right, bottom))
        cropped_img.save(output_path)
        
        # 显式内存释放
        pix = None
        img = None
        cropped_img = None
        doc.close()
        return True
    except Exception as e:
        return str(e)

# 设置外观
ctk.set_appearance_mode("System")  # 模式: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # 主题: "blue" (standard), "green", "dark-blue"

class PDFToImageConverter(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("PDF 转图片工具 (v1.5.11稳定版)")
        self.geometry("700x580")

        # 设置窗口图标
        try:
            icon_path = resource_path("app.ico")
            if os.path.exists(icon_path):
                self.after(200, lambda: self.iconbitmap(icon_path))
        except Exception:
            pass

        # 变量
        self.pdf_path = ctk.StringVar()
        self.output_dir = ctk.StringVar()
        self.quality_var = ctk.StringVar(value="普通 (150 DPI)")
        self.is_converting = False
        self.stop_requested = False
        
        self.quality_map = {
            "普通 (150 DPI)": 150,
            "高清 (300 DPI)": 300,
            "超清 (600 DPI)": 600,
            "原稿 (72 DPI)": 72
        }
        self.crop_left = ctk.StringVar(value="0")
        self.crop_top = ctk.StringVar(value="0")
        self.crop_right = ctk.StringVar(value="0")
        self.crop_bottom = ctk.StringVar(value="0")
        self.preview_page = ctk.StringVar(value="1")
        self.preview_window_obj = None  # 记录预览窗口对象
        self.preview_canvas = None      # 预览画布
        self.preview_image_id = None    # 画布上的图片ID
        self.preview_rect_id = None     # 画布上的裁剪框ID
        self.shade_ids = []             # 阴影遮罩ID列表
        self.full_preview_img = None    # 完整的预览图（PIL）
        self.preview_scale = 1.0        # 预览图缩放比例
        self.is_dragging = False        # 是否正在拖拽裁剪框
        self.drag_edge = None           # 正在拖拽哪个边缘
        self.drag_start_pos = (0, 0)    # 拖拽起始坐标
        self.initial_crops = (0, 0, 0, 0) # 拖拽起始裁剪值
        self.canvas_offset = 5          # 画布边缘留白，防止线条被切断
        
        # 预览窗口预留空间常量 (必须与布局组件占用的空间一致)
        # 包含：主容器边距、Canvas 边距、导航栏高度、画布 Offset
        self.PREVIEW_PAD_X = 120 
        self.PREVIEW_PAD_Y = 200
        
        # 绑定变量追踪，实现实时更新
        for var in [self.crop_left, self.crop_top, self.crop_right, self.crop_bottom]:
            var.trace_add("write", self.on_crop_var_change)

        self.setup_ui()
        
        # 注册拖拽事件 (增加异常保护)
        try:
            windnd.hook_dropfiles(self, self.on_file_drop)
        except Exception as e:
            print(f"拖拽功能注册失败: {e}")

    def on_file_drop(self, files):
        """
        拖拽回调函数：仅负责接收数据，立即交由主线程处理
        避免在系统钩子线程中直接操作 UI 导致闪退
        """
        self.after(10, lambda: self._process_dropped_files(files))

    def _process_dropped_files(self, files):
        try:
            if not files:
                return
            
            # 获取原始路径数据
            raw_path = files[0]
            
            # 健壮的解码逻辑
            if isinstance(raw_path, bytes):
                try:
                    file_path = raw_path.decode('utf-8')
                except UnicodeDecodeError:
                    try:
                        file_path = raw_path.decode('gbk')
                    except UnicodeDecodeError:
                        file_path = raw_path.decode('gbk', errors='ignore')
            else:
                file_path = raw_path

            # 标准化路径
            file_path = os.path.normpath(file_path.strip())
            
            if os.path.isfile(file_path) and file_path.lower().endswith('.pdf'):
                self.pdf_path.set(file_path)
                if not self.output_dir.get():
                    self.output_dir.set(os.path.dirname(file_path))
                self.status_label.configure(text=f"已加载: {os.path.basename(file_path)}")
            else:
                messagebox.showwarning("格式错误", "请拖拽有效的 PDF 文件！")
        except Exception as e:
            messagebox.showerror("拖拽失败", f"处理拖拽文件时出错: {str(e)}")

    def setup_ui(self):
        # 配置网格
        self.grid_columnconfigure(0, weight=1)
        
        # 标题
        self.label_title = ctk.CTkLabel(self, text="PDF 转图片工具 (支持拖拽)", font=ctk.CTkFont(size=24, weight="bold"))
        self.label_title.grid(row=0, column=0, padx=20, pady=(20, 10))

        # 拖拽提示
        self.drop_label = ctk.CTkLabel(self, text="💡 提示：支持直接将 PDF 文件拖拽到此处", font=ctk.CTkFont(size=12), text_color="gray")
        self.drop_label.grid(row=1, column=0, padx=20, pady=(0, 10))

        # 文件选择区域
        file_frame = ctk.CTkFrame(self)
        file_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        file_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(file_frame, text="PDF 文件:").grid(row=0, column=0, padx=10, pady=10)
        self.entry_pdf = ctk.CTkEntry(file_frame, textvariable=self.pdf_path)
        self.entry_pdf.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(file_frame, text="选择文件", width=100, command=self.browse_pdf).grid(row=0, column=2, padx=10, pady=10)

        ctk.CTkLabel(file_frame, text="保存路径:").grid(row=1, column=0, padx=10, pady=10)
        self.entry_out = ctk.CTkEntry(file_frame, textvariable=self.output_dir)
        self.entry_out.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(file_frame, text="选择目录", width=100, command=self.browse_output).grid(row=1, column=2, padx=10, pady=10)

        # 设置区域
        settings_frame = ctk.CTkFrame(self)
        settings_frame.grid(row=3, column=0, padx=20, pady=10, sticky="ew")
        settings_frame.grid_columnconfigure((1, 3), weight=1)

        ctk.CTkLabel(settings_frame, text="图片清晰度:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, pady=10)
        self.quality_combo = ctk.CTkComboBox(settings_frame, values=list(self.quality_map.keys()), variable=self.quality_var, width=200)
        self.quality_combo.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        # 裁剪区域
        crop_frame = ctk.CTkFrame(self)
        crop_frame.grid(row=4, column=0, padx=20, pady=10, sticky="ew")
        for i in range(4): crop_frame.grid_columnconfigure(i*2+1, weight=1)

        ctk.CTkLabel(crop_frame, text="裁剪设置 (像素):", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, columnspan=2, padx=10, pady=5, sticky="w")
        
        # 封装一个带步进器的输入框
        def create_stepper(label_text, var, row, col):
            ctk.CTkLabel(crop_frame, text=label_text).grid(row=row, column=col, padx=(10, 2), pady=10)
            f = ctk.CTkFrame(crop_frame, fg_color="transparent")
            f.grid(row=row, column=col+1, padx=2, pady=10)
            ctk.CTkButton(f, text="-", width=28, command=lambda: self.adjust_val(var, -10)).pack(side="left")
            ctk.CTkEntry(f, textvariable=var, width=50).pack(side="left", padx=2)
            ctk.CTkButton(f, text="+", width=28, command=lambda: self.adjust_val(var, 10)).pack(side="left")

        create_stepper("左:", self.crop_left, 1, 0)
        create_stepper("上:", self.crop_top, 1, 2)
        create_stepper("右:", self.crop_right, 1, 4)
        create_stepper("下:", self.crop_bottom, 1, 6)

        # 预览控制
        preview_ctrl_frame = ctk.CTkFrame(self)
        preview_ctrl_frame.grid(row=5, column=0, padx=20, pady=5, sticky="ew")
        
        ctk.CTkLabel(preview_ctrl_frame, text="预览页码:").grid(row=0, column=0, padx=10, pady=10)
        ctk.CTkEntry(preview_ctrl_frame, textvariable=self.preview_page, width=60).grid(row=0, column=1, padx=5, pady=10)
        ctk.CTkButton(preview_ctrl_frame, text="交互式裁剪预览", width=140, command=self.show_preview).grid(row=0, column=2, padx=10, pady=10)

        # 进度条
        self.progress_bar = ctk.CTkProgressBar(self)
        self.progress_bar.grid(row=6, column=0, padx=20, pady=10, sticky="ew")
        self.progress_bar.set(0)

        # 按钮和状态
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=7, column=0, padx=20, pady=10)

        self.convert_btn = ctk.CTkButton(button_frame, text="开始转换", height=40, width=120, font=ctk.CTkFont(size=16, weight="bold"), command=self.start_conversion)
        self.convert_btn.grid(row=0, column=0, padx=10)

        self.stop_btn = ctk.CTkButton(button_frame, text="停止", height=40, width=100, fg_color="#E74C3C", hover_color="#C0392B", font=ctk.CTkFont(size=16, weight="bold"), command=self.request_stop, state="disabled")
        self.stop_btn.grid(row=0, column=1, padx=10)

        self.status_label = ctk.CTkLabel(self, text="准备就绪", font=ctk.CTkFont(size=12))
        self.status_label.grid(row=8, column=0, padx=20, pady=(0, 20))

    def browse_pdf(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if filename:
            self.pdf_path.set(filename)
            if not self.output_dir.get():
                self.output_dir.set(os.path.dirname(filename))

    def browse_output(self):
        directory = filedialog.askdirectory()
        if directory:
            self.output_dir.set(directory)

    def adjust_val(self, var, delta):
        try:
            val = int(var.get() or 0)
            new_val = max(0, val + delta)
            var.set(str(new_val))
        except ValueError:
            var.set("0")

    def on_crop_var_change(self, *args):
        """当裁剪数值变化时，更新预览框"""
        if self.preview_window_obj and self.preview_window_obj.winfo_exists():
            self.update_preview_rect()

    def request_stop(self):
        if self.is_converting:
            self.stop_requested = True
            self.status_label.configure(text="正在停止...")
            self.stop_btn.configure(state="disabled")

    def show_preview(self):
        if not self.pdf_path.get():
            messagebox.showwarning("警告", "请先选择 PDF 文件！")
            return
        
        try:
            page_num = int(self.preview_page.get()) - 1
        except ValueError:
            messagebox.showerror("错误", "请输入有效的页码！")
            return

        try:
            doc = fitz.open(self.pdf_path.get())
            if page_num < 0 or page_num >= len(doc):
                messagebox.showerror("错误", f"页码超出范围 (1-{len(doc)})")
                return
            
            # 使用 150 DPI 进行预览（全图）
            zoom = 150 / 72
            mat = fitz.Matrix(zoom, zoom)
            page = doc.load_page(page_num)
            pix = page.get_pixmap(matrix=mat)
            
            img_data = pix.tobytes("png")
            self.full_preview_img = Image.open(io.BytesIO(img_data))
            
            # 弹出/更新预览窗口
            self.open_preview_window()
            doc.close()
        except Exception as e:
            messagebox.showerror("错误", f"预览生成失败: {str(e)}")

    def open_preview_window(self):
        img_w, img_h = self.full_preview_img.size
        
        # 1. 决定目标可用空间
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        
        if self.preview_window_obj and self.preview_window_obj.winfo_exists():
            current_win_w = self.preview_window_obj.winfo_width()
            current_win_h = self.preview_window_obj.winfo_height()
            if current_win_w > 100 and current_win_h > 100:
                available_w = current_win_w - self.PREVIEW_PAD_X
                available_h = current_win_h - self.PREVIEW_PAD_Y
            else:
                available_w = screen_w * 0.85 - self.PREVIEW_PAD_X
                available_h = screen_h * 0.8 - self.PREVIEW_PAD_Y
        else:
            available_w = screen_w * 0.85 - self.PREVIEW_PAD_X
            available_h = screen_h * 0.8 - self.PREVIEW_PAD_Y
        
        # 2. 计算缩放比例并更新图片
        self.preview_scale = min(available_w / img_w, available_h / img_h)
        display_w = int(round(img_w * self.preview_scale))
        display_h = int(round(img_h * self.preview_scale))
        
        pil_img_resized = self.full_preview_img.resize((display_w, display_h), Image.Resampling.LANCZOS)
        self.tk_img = ImageTk.PhotoImage(pil_img_resized)

        if self.preview_window_obj and self.preview_window_obj.winfo_exists():
            self.preview_window_obj.lift()
            self.preview_canvas.config(width=display_w + self.canvas_offset * 2, height=display_h + self.canvas_offset * 2)
            self.preview_canvas.itemconfig(self.preview_image_id, image=self.tk_img)
            self.preview_canvas.coords(self.preview_image_id, self.canvas_offset, self.canvas_offset)
            self.update_page_label()
            self.update_preview_rect()
            return

        self.preview_window_obj = ctk.CTkToplevel(self)
        self.preview_window_obj.title("裁剪区域预览 (拖拽边框或四个角进行调整)")
        self.preview_window_obj.attributes("-topmost", True)
        
        # 重置缩放状态变量，防止二次打开时受旧数据干扰
        self._last_resize_size = None
        self._last_resize_time = 0
        
        # 设置初始窗口几何尺寸并居中
        win_w = display_w + self.PREVIEW_PAD_X
        win_h = display_h + self.PREVIEW_PAD_Y
        x = (screen_w - win_w) // 2
        y = (screen_h - win_h) // 2
        self.preview_window_obj.geometry(f"{win_w}x{win_h}+{x}+{y}")
        
        # 绑定 Resize 事件
        self.preview_window_obj.bind("<Configure>", self.on_preview_resize)
        
        # 绑定关闭事件，清理状态
        self.preview_window_obj.protocol("WM_DELETE_WINDOW", self.close_preview_window)
        
        # 主容器
        main_container = ctk.CTkFrame(self.preview_window_obj)
        main_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 顶部工具栏 (翻页和信息)
        nav_frame = ctk.CTkFrame(main_container, fg_color="transparent")
        nav_frame.pack(fill="x", pady=(0, 5))
        
        ctk.CTkButton(nav_frame, text="上一页", width=80, command=self.prev_preview_page).pack(side="left", padx=5)
        self.page_info_label = ctk.CTkLabel(nav_frame, text="", font=ctk.CTkFont(weight="bold"))
        self.page_info_label.pack(side="left", expand=True)
        ctk.CTkButton(nav_frame, text="下一页", width=80, command=self.next_preview_page).pack(side="right", padx=5)
        
        # 画布容器 (居中)
        canvas_container = ctk.CTkFrame(main_container, fg_color="transparent")
        canvas_container.pack(fill="both", expand=True)
        
        self.preview_canvas = Canvas(
            canvas_container, 
            highlightthickness=0, 
            bg="#2b2b2b",
            width=display_w + self.canvas_offset * 2,
            height=display_h + self.canvas_offset * 2
        )
        self.preview_canvas.place(relx=0.5, rely=0.5, anchor="center")
        
        # 渲染图片
        self.preview_image_id = self.preview_canvas.create_image(
            self.canvas_offset, self.canvas_offset, 
            anchor="nw", 
            image=self.tk_img
        )
        
        # 创建阴影遮罩 (上, 下, 左, 右)
        self.shade_ids = []
        for _ in range(4):
            sid = self.preview_canvas.create_rectangle(0, 0, 0, 0, fill="black", stipple="gray50", outline="")
            self.shade_ids.append(sid)
        
        # 裁剪框 (红色虚线，加粗)
        self.preview_rect_id = self.preview_canvas.create_rectangle(
            0, 0, 0, 0, 
            outline="red", 
            width=3, 
            dash=(4, 4)
        )
        
        # 底部尺寸信息
        self.size_info_label = ctk.CTkLabel(main_container, text="裁剪尺寸: 0 x 0", text_color="gray")
        self.size_info_label.pack(side="bottom", pady=5)
        
        # 确保裁剪框在阴影之上
        self.preview_canvas.tag_raise(self.preview_rect_id)
        
        # 事件绑定
        self.preview_canvas.bind("<Button-1>", self.on_canvas_click)
        self.preview_canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.preview_canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        self.preview_canvas.bind("<Motion>", self.on_canvas_hover)
        
        self.update_page_label()
        self.update_preview_rect()
        
        # 强制更新一次布局，确保渲染完成
        self.preview_window_obj.update_idletasks()
        
        # 二次确认同步：延迟一小段时间强制校准比例，解决部分系统下二次打开尺寸不准的问题
        self.preview_window_obj.after(200, lambda: self.on_preview_resize(None, force=True))

    def close_preview_window(self):
        """关闭预览窗口并清理状态"""
        if self.preview_window_obj:
            self._last_resize_size = None
            self.preview_window_obj.destroy()
            self.preview_window_obj = None

    def on_preview_resize(self, event, force=False):
        """处理预览窗口缩放事件"""
        if event and event.widget != self.preview_window_obj:
            return
            
        import time
        curr_time = time.time()
        
        if event:
            new_w, new_h = event.width, event.height
            # 增加保护：忽略窗口初始化时可能出现的极小尺寸事件
            if new_w < 200 or new_h < 200:
                return
                
            if hasattr(self, '_last_resize_size'):
                if self._last_resize_size == (new_w, new_h):
                    return
            self._last_resize_size = (new_w, new_h)
        else:
            new_w = self.preview_window_obj.winfo_width()
            new_h = self.preview_window_obj.winfo_height()
            if new_w <= 200 or new_h <= 200: # 同样增加最小尺寸保护
                return

        available_w = max(100, new_w - self.PREVIEW_PAD_X)
        available_h = max(100, new_h - self.PREVIEW_PAD_Y)
        
        img_w, img_h = self.full_preview_img.size
        new_scale = min(available_w / img_w, available_h / img_h)
        
        if force or abs(new_scale - self.preview_scale) > 0.001:
            self.preview_scale = new_scale
            display_w = int(round(img_w * self.preview_scale))
            display_h = int(round(img_h * self.preview_scale))
            
            pil_img_resized = self.full_preview_img.resize((display_w, display_h), Image.Resampling.LANCZOS)
            self.tk_img = ImageTk.PhotoImage(pil_img_resized)
            
            self.preview_canvas.config(width=display_w + self.canvas_offset * 2, height=display_h + self.canvas_offset * 2)
            self.preview_canvas.itemconfig(self.preview_image_id, image=self.tk_img)
            self.preview_canvas.coords(self.preview_image_id, self.canvas_offset, self.canvas_offset)
            self.update_preview_rect()

    def update_page_label(self):
        if hasattr(self, 'page_info_label') and self.page_info_label.winfo_exists():
            current = self.preview_page.get()
            try:
                doc = fitz.open(self.pdf_path.get())
                total = len(doc)
                doc.close()
                self.page_info_label.configure(text=f"第 {current} / {total} 页")
            except:
                self.page_info_label.configure(text=f"第 {current} 页")

    def prev_preview_page(self):
        try:
            current = int(self.preview_page.get())
            if current > 1:
                self.preview_page.set(str(current - 1))
                self.show_preview()
        except ValueError:
            pass

    def next_preview_page(self):
        try:
            current = int(self.preview_page.get())
            doc = fitz.open(self.pdf_path.get())
            total = len(doc)
            doc.close()
            if current < total:
                self.preview_page.set(str(current + 1))
                self.show_preview()
        except Exception:
            pass

    def update_preview_rect(self):
        if not self.preview_canvas or not self.preview_window_obj.winfo_exists():
            return
            
        try:
            img_w, img_h = self.full_preview_img.size
            # 使用 float 先转再转 int，防止字符串带小数点导致报错
            l_val = int(float(self.crop_left.get() or 0))
            t_val = int(float(self.crop_top.get() or 0))
            r_val = int(float(self.crop_right.get() or 0))
            b_val = int(float(self.crop_bottom.get() or 0))
            
            l = int(round(l_val * self.preview_scale)) + self.canvas_offset
            t = int(round(t_val * self.preview_scale)) + self.canvas_offset
            r = int(round((img_w - r_val) * self.preview_scale)) + self.canvas_offset
            b = int(round((img_h - b_val) * self.preview_scale)) + self.canvas_offset
            
            # 更新主裁剪框
            self.preview_canvas.coords(self.preview_rect_id, l, t, r, b)
            
            # 更新阴影遮罩 (上, 下, 左, 右)
            canvas_w = int(round(img_w * self.preview_scale))
            canvas_h = int(round(img_h * self.preview_scale))
            off = self.canvas_offset
            self.preview_canvas.coords(self.shade_ids[0], off, off, canvas_w + off, t) # Top
            self.preview_canvas.coords(self.shade_ids[1], off, b, canvas_w + off, canvas_h + off) # Bottom
            self.preview_canvas.coords(self.shade_ids[2], off, t, l, b) # Left
            self.preview_canvas.coords(self.shade_ids[3], r, t, canvas_w + off, b) # Right
            
            # 确保裁剪框在阴影之上
            self.preview_canvas.tag_raise(self.preview_rect_id)
            
            # 更新尺寸信息
            cw = max(0, img_w - l_val - r_val)
            ch = max(0, img_h - t_val - b_val)
            if hasattr(self, 'size_info_label'):
                self.size_info_label.configure(text=f"裁剪尺寸: {cw} x {ch} 像素 (宽x高)")
        except Exception as e:
            print(f"Update rect error: {e}")

    def on_canvas_hover(self, event):
        if self.is_dragging: return
        
        edge = self.get_edge_at(event.x, event.y)
        cursor_map = {
            "nw": "size_nw_se", "se": "size_nw_se",
            "ne": "size_ne_sw", "sw": "size_ne_sw",
            "left": "size_we", "right": "size_we",
            "top": "size_ns", "bottom": "size_ns",
            "move": "fleur"
        }
        if edge in cursor_map:
            self.preview_canvas.config(cursor=cursor_map[edge])
        else:
            self.preview_canvas.config(cursor="")

    def get_edge_at(self, x, y):
        try:
            img_w, img_h = self.full_preview_img.size
            off = self.canvas_offset
            l = int(round(int(self.crop_left.get() or 0) * self.preview_scale)) + off
            t = int(round(int(self.crop_top.get() or 0) * self.preview_scale)) + off
            r = int(round((img_w - int(self.crop_right.get() or 0)) * self.preview_scale)) + off
            b = int(round((img_h - int(self.crop_bottom.get() or 0)) * self.preview_scale)) + off
            
            margin = 20
            # 优先检测角落
            if abs(x - l) < margin and abs(y - t) < margin: return "nw"
            if abs(x - r) < margin and abs(y - t) < margin: return "ne"
            if abs(x - l) < margin and abs(y - b) < margin: return "sw"
            if abs(x - r) < margin and abs(y - b) < margin: return "se"
            
            # 检测边缘
            if abs(x - l) < margin and t < y < b: return "left"
            if abs(x - r) < margin and t < y < b: return "right"
            if abs(y - t) < margin and l < x < r: return "top"
            if abs(y - b) < margin and l < x < r: return "bottom"
            
            # 检测中心移动
            if l < x < r and t < y < b: return "move"
        except Exception:
            pass
        return None

    def on_canvas_click(self, event):
        self.drag_edge = self.get_edge_at(event.x, event.y)
        if self.drag_edge:
            self.is_dragging = True
            self.drag_start_pos = (event.x, event.y)
            self.initial_crops = (
                int(self.crop_left.get() or 0),
                int(self.crop_top.get() or 0),
                int(self.crop_right.get() or 0),
                int(self.crop_bottom.get() or 0)
            )

    def on_canvas_drag(self, event):
        if not self.is_dragging or not self.drag_edge: return
        
        img_w, img_h = self.full_preview_img.size
        dx = (event.x - self.drag_start_pos[0]) / self.preview_scale
        dy = (event.y - self.drag_start_pos[1]) / self.preview_scale
        
        l, t, r, b = self.initial_crops
        
        if self.drag_edge == "move":
            # 限制移动范围，保持宽高不变
            curr_w = img_w - l - r
            curr_h = img_h - t - b
            
            new_l = max(0, min(int(round(l + dx)), img_w - curr_w))
            new_t = max(0, min(int(round(t + dy)), img_h - curr_h))
            
            self.crop_left.set(str(new_l))
            self.crop_right.set(str(img_w - new_l - curr_w))
            self.crop_top.set(str(new_t))
            self.crop_bottom.set(str(img_h - new_t - curr_h))
            return

        # 边缘与角落拖拽
        min_size = 10
        if "left" in self.drag_edge or "nw" in self.drag_edge or "sw" in self.drag_edge:
            new_l = max(0, int(round(l + dx)))
            # 确保不越过右边界 (保留最小宽度)
            new_l = min(new_l, img_w - r - min_size)
            self.crop_left.set(str(new_l))
            
        if "right" in self.drag_edge or "ne" in self.drag_edge or "se" in self.drag_edge:
            new_r = max(0, int(round(r - dx)))
            # 确保不越过左边界 (保留最小宽度)
            new_r = min(new_r, img_w - l - min_size)
            self.crop_right.set(str(new_r))
            
        if "top" in self.drag_edge or "nw" in self.drag_edge or "ne" in self.drag_edge:
            new_t = max(0, int(round(t + dy)))
            # 确保不越过下边界 (保留最小高度)
            new_t = min(new_t, img_h - b - min_size)
            self.crop_top.set(str(new_t))
            
        if "bottom" in self.drag_edge or "sw" in self.drag_edge or "se" in self.drag_edge:
            new_b = max(0, int(round(b - dy)))
            # 确保不越过上边界 (保留最小高度)
            new_b = min(new_b, img_h - t - min_size)
            self.crop_bottom.set(str(new_b))

    def on_canvas_release(self, event):
        self.is_dragging = False
        self.drag_edge = None

    def start_conversion(self):
        if not self.pdf_path.get():
            messagebox.showwarning("警告", "请先选择 PDF 文件！")
            return
        
        if not self.output_dir.get():
            messagebox.showwarning("警告", "请选择保存路径！")
            return
        
        self.is_converting = True
        self.stop_requested = False
        self.convert_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.status_label.configure(text="正在转换...")
        threading.Thread(target=self.convert, daemon=True).start()

    def convert(self):
        executor = None
        try:
            pdf_path = self.pdf_path.get()
            base_output_dir = self.output_dir.get()
            
            try:
                dpi_val = self.quality_map.get(self.quality_var.get(), 150)
                zoom = dpi_val / 72
                crop_params = (
                    int(self.crop_left.get()),
                    int(self.crop_top.get()),
                    int(self.crop_right.get()),
                    int(self.crop_bottom.get())
                )
            except ValueError:
                self.after(0, lambda: messagebox.showerror("错误", "请输入有效的裁剪像素数字！"))
                return

            pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
            final_output_dir = os.path.join(base_output_dir, pdf_name)
            if not os.path.exists(final_output_dir):
                os.makedirs(final_output_dir)

            doc = fitz.open(pdf_path)
            total_pages = len(doc)
            doc.close() # 进程中会重新打开
            
            # 使用进程池
            max_workers = min(os.cpu_count() or 4, 8)
            executor = ProcessPoolExecutor(max_workers=max_workers)
            
            futures = []
            for i in range(total_pages):
                output_path = os.path.join(final_output_dir, f"page{i+1}.png")
                future = executor.submit(process_page_task, pdf_path, i, zoom, crop_params, output_path)
                futures.append(future)
            
            completed = 0
            for future in futures:
                if self.stop_requested:
                    executor.shutdown(wait=False, cancel_futures=True)
                    break
                
                result = future.result() # 等待结果
                if result is not True:
                    print(f"Error in process: {result}")
                
                completed += 1
                progress = completed / total_pages
                self.after(0, lambda p=progress, c=completed, t=total_pages: self.update_progress(p, c, t))
            
            if self.stop_requested:
                self.after(0, lambda: messagebox.showinfo("提示", "转换已停止。"))
            else:
                self.after(0, lambda: self.show_success_dialog(pdf_name, total_pages, final_output_dir))
            
        except Exception as e:
            self.after(0, lambda msg=str(e): messagebox.showerror("错误", f"转换过程中发生错误: {msg}"))
        finally:
            if executor:
                executor.shutdown(wait=True)
            self.after(0, self.reset_ui_state)

    def update_progress(self, progress, completed, total):
        self.progress_bar.set(progress)
        self.status_label.configure(text=f"正在处理第 {completed}/{total} 页...")

    def show_success_dialog(self, pdf_name, total_pages, final_output_dir):
        if messagebox.askyesno("成功", f"转换完成！\n文件夹：{pdf_name}\n共生成 {total_pages} 张图片。\n是否打开文件夹？"):
            os.startfile(final_output_dir)

    def reset_ui_state(self):
        self.is_converting = False
        self.stop_requested = False
        self.convert_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.status_label.configure(text="准备就绪")
        self.progress_bar.set(0)

if __name__ == "__main__":
    # 多进程打包必须调用 freeze_support
    multiprocessing.freeze_support()
    app = PDFToImageConverter()
    app.mainloop()
