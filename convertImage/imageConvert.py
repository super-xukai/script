import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
from PIL import Image, ImageTk
import os

class ImageConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("图片格式转换器")
        self.root.geometry("600x400")
        self.root.resizable(False, False)
        
        # 支持的格式
        self.supported_formats = [
            "PNG", "JPEG", "JPG", "BMP", "GIF", "TIFF", "ICO", "WEBP"
        ]
        
        self.selected_file = None
        self.preview_image = None
        
        self.create_widgets()
        
        # 设置拖拽功能
        self.setup_drag_and_drop()
    
    def setup_drag_and_drop(self):
        """设置拖拽功能"""
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.on_drop)
    
    def on_drop(self, event):
        """处理文件拖拽释放事件"""
        file_path = event.data
        # 移除花括号（如果存在）
        file_path = file_path.strip('{}')
        
        if self.is_supported_image(file_path):
            self.selected_file = file_path
            self.file_path_var.set(file_path)
            self.show_preview()
        else:
            messagebox.showerror("错误", "不支持的文件格式")
    
    def is_supported_image(self, file_path):
        """检查文件是否为支持的图片格式"""
        supported_extensions = ['.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff', '.ico', '.webp']
        _, ext = os.path.splitext(file_path.lower())
        return ext in supported_extensions

    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="选择文件", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.file_path_var = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, width=50, state="readonly")
        file_entry.grid(row=0, column=0, padx=(0, 10))
        
        browse_btn = ttk.Button(file_frame, text="浏览", command=self.browse_file)
        browse_btn.grid(row=0, column=1)
        
        # 预览区域
        preview_frame = ttk.LabelFrame(main_frame, text="预览", padding="10")
        preview_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        
        self.preview_label = ttk.Label(preview_frame)
        self.preview_label.grid(row=0, column=0)
        
        # 转换设置区域
        settings_frame = ttk.LabelFrame(main_frame, text="转换设置", padding="10")
        settings_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 输出格式选择
        ttk.Label(settings_frame, text="输出格式:").grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        self.format_var = tk.StringVar(value="PNG")
        format_combo = ttk.Combobox(settings_frame, textvariable=self.format_var, 
                                   values=self.supported_formats, state="readonly", width=15)
        format_combo.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 质量设置（仅对JPEG有效）
        ttk.Label(settings_frame, text="质量 (JPEG):").grid(row=2, column=0, sticky=tk.W, pady=(0, 10))
        
        self.quality_var = tk.IntVar(value=95)
        quality_scale = ttk.Scale(settings_frame, from_=1, to=100, orient=tk.HORIZONTAL, 
                                 variable=self.quality_var)
        quality_scale.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        quality_label = ttk.Label(settings_frame, textvariable=self.quality_var)
        quality_label.grid(row=4, column=0)
        
        # 保存按钮
        convert_btn = ttk.Button(settings_frame, text="转换并保存", command=self.convert_image)
        convert_btn.grid(row=5, column=0, pady=(20, 0))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        settings_frame.columnconfigure(0, weight=1)
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="选择图片文件",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.bmp *.gif *.tiff *.ico *.webp"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.selected_file = file_path
            self.file_path_var.set(file_path)
            self.show_preview()
    
    def show_preview(self):
        if not self.selected_file:
            return
            
        try:
            # 打开并调整图片大小用于预览
            image = Image.open(self.selected_file)
            image.thumbnail((200, 200))
            
            # 保存预览图片引用防止被垃圾回收
            self.preview_image = ImageTk.PhotoImage(image)
            self.preview_label.configure(image=self.preview_image)
        except Exception as e:
            messagebox.showerror("错误", f"无法加载图片预览: {str(e)}")
    
    def convert_image(self):
        if not self.selected_file:
            messagebox.showwarning("警告", "请先选择一个图片文件")
            return
            
        # 获取输出格式
        output_format = self.format_var.get()
        
        # 确定文件扩展名
        extension = output_format.lower()
        if output_format == "JPEG":
            extension = "jpg"
        
        # 生成输出文件名
        base_name = os.path.splitext(os.path.basename(self.selected_file))[0]
        output_filename = f"{base_name}.{extension}"
        
        # 选择保存位置
        output_path = filedialog.asksaveasfilename(
            title="保存转换后的图片",
            initialfile=output_filename,
            defaultextension=f".{extension}",
            filetypes=[(f"{output_format} files", f"*.{extension}"), ("All files", "*.*")]
        )
        
        if not output_path:
            return
            
        try:
            # 打开原图
            image = Image.open(self.selected_file)
            
            # 处理RGBA到RGB的转换（JPEG不支持透明度）
            if output_format in ["JPEG", "JPG"] and image.mode in ('RGBA', 'LA', 'P'):
                # 创建白色背景
                background = Image.new('RGB', image.size, (255, 255, 255))
                if image.mode == 'P':
                    image = image.convert('RGBA')
                background.paste(image, mask=image.split()[-1] if image.mode == 'RGBA' else None)
                image = background
            
            # 保存图片
            save_kwargs = {}
            if output_format in ["JPEG", "JPG"]:
                save_kwargs['quality'] = self.quality_var.get()
                save_kwargs['optimize'] = True
            
            # 特殊处理ICO格式
            if output_format == "ICO":
                image.save(output_path, format="ICO", sizes=[(16, 16), (32, 32), (48, 48), (64, 64)])
            else:
                image.save(output_path, format=output_format, **save_kwargs)
            
            messagebox.showinfo("成功", f"图片已成功转换并保存到:\n{output_path}")
            
        except Exception as e:
            messagebox.showerror("错误", f"转换失败: {str(e)}")

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = ImageConverter(root)
    root.mainloop()