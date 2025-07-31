# -*- coding: utf-8 -*-
"""
GUI界面模块。
包含 App 类，负责构建和管理所有Tkinter界面元素和交互。
"""

import os
import threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

from ..config import CONFIG
from ..core.logic import DocumentGenerator

class App:
    def __init__(self, root):
        self.root = root
        self.surgery_query_files = []
        self.generation_thread = None
        self.generator_instance = None
        
        self.scaling_factor = self._get_scaling_factor()
        
        self.setup_fonts()
        self.setup_window()
        self.create_widgets()

    def _get_scaling_factor(self):
        """获取屏幕缩放比例"""
        try:
            # 使用 Tk 8.6+ 的方法来获取DPI
            dpi = self.root.winfo_fpixels('1i')
            scaling = dpi / 96.0
            # 对缩放比例进行合理性检查
            if scaling < 0.75: return 1.0
            return scaling
        except Exception:
            return 1.0

    def setup_fonts(self):
        """设置字体"""
        self.font_normal = ("微软雅黑", 9)
        self.font_bold = ("微软雅黑", 10, "bold")
        self.font_title = ("微软雅黑", 20, "bold")
        self.font_subtitle = ("微软雅黑", 16, "bold")
        self.font_button = ("微软雅黑", 12, "bold")
        self.font_disclaimer = ("微软雅黑", 10)

    def setup_window(self):
        """设置窗口大小并居中"""
        self.root.title(CONFIG['app_title'])
        s = self.scaling_factor
        width = int(685 * s)
        height = int(535 * s)
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        pos_x = (screen_width // 2) - (width // 2)
        pos_y = (screen_height // 2) - (height // 2)
        
        self.root.geometry(f"{width}x{height}+{pos_x}+{pos_y}") 
        self.root.resizable(False, False)

    def create_widgets(self):
        """创建所有界面组件"""
        style = ttk.Style(self.root)
        if "clam" in style.theme_names(): style.theme_use("clam")
        default_bg = style.lookup('TFrame', 'background')
        self.root.configure(bg=default_bg)
        self.log_text_tags = {"warning": {"foreground": "orange"}, "error": {"foreground": "red"}}

        bottom_frame = tk.Frame(self.root, bg=default_bg)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=2)
        tk.Label(bottom_frame, text=f"版本日期：{datetime.now().strftime('%Y年%m月%d日')}", font=self.font_normal, fg="#666666", bg=default_bg).pack(side=tk.LEFT)
        tk.Label(bottom_frame, text="作者：顾江江", font=self.font_normal, fg="#666666", bg=default_bg).pack(side=tk.RIGHT)
        
        disclaimer_frame = tk.Frame(self.root, pady=2, bg=default_bg)
        disclaimer_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10)
        tk.Label(disclaimer_frame, text="本工具仅供骨科内部测试，请勿外传", font=self.font_disclaimer, fg="red", bg=default_bg).pack()

        top_title_frame = tk.Frame(self.root, bg=default_bg)
        top_title_frame.pack(side=tk.TOP, fill=tk.X, pady=(10, 5))
        tk.Label(top_title_frame, text="丹阳市人民医院", font=self.font_subtitle, fg="#0066cc", bg=default_bg).pack()
        tk.Label(top_title_frame, text="日间手术随访表生成系统", font=self.font_title, bg=default_bg).pack()

        main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 5))

        left_panel = ttk.Frame(main_pane, padding=(5, 5, 5, 5))
        main_pane.add(left_panel)

        right_panel = ttk.Frame(main_pane, padding=(5, 5, 5, 5))
        main_pane.add(right_panel)

        s = self.scaling_factor
        sash_default = int(420 * s)
        sash_min = int(320 * s)
        sash_max = int(520 * s)

        def set_initial_sash(event):
            main_pane.sashpos(0, sash_default)
            main_pane.unbind("<Configure>")
        
        def limit_sash_movement(event):
            if event.x < sash_min:
                main_pane.sashpos(0, sash_min)
                return "break"
            if event.x > sash_max:
                main_pane.sashpos(0, sash_max)
                return "break"

        main_pane.bind("<Configure>", set_initial_sash)
        main_pane.bind("<B1-Motion>", limit_sash_movement)

        file_frame = ttk.LabelFrame(left_panel, text="步骤1: 选择文件和路径", padding=5)
        file_frame.pack(fill=tk.BOTH, expand=True) 
        
        self.excel_path_var = tk.StringVar()
        self.template_path_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        
        self.create_file_selector(file_frame, "出院患者列表:", self.excel_path_var, self.select_excel_file)
        
        surgery_frame = ttk.LabelFrame(file_frame, text="手术查询文件 (可多选, 用于补充床号)", padding=5)
        surgery_frame.pack(fill=tk.X, expand=True, pady=3)
        
        self.surgery_listbox = tk.Listbox(surgery_frame, height=5, font=self.font_normal)
        self.surgery_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        
        surgery_buttons_frame = ttk.Frame(surgery_frame)
        surgery_buttons_frame.pack(side=tk.LEFT, fill=tk.Y, anchor='n')
        ttk.Button(surgery_buttons_frame, text="添加文件", command=self.select_surgery_query_files, width=8).pack(fill=tk.X, pady=1)
        ttk.Button(surgery_buttons_frame, text="清空列表", command=self.clear_surgery_query_files, width=8).pack(fill=tk.X, pady=1)
        
        self.create_file_selector(file_frame, "Word模板:", self.template_path_var, self.select_template_file)
        self.create_file_selector(file_frame, "输出文件夹:", self.output_dir_var, self.select_output_dir)
        
        left_bottom_container = ttk.Frame(left_panel)
        left_bottom_container.pack(fill=tk.X, pady=(5,0))

        control_frame = ttk.LabelFrame(left_bottom_container, text="步骤2: 开始生成", padding=10)
        control_frame.pack(fill=tk.X)
        
        # 为不同状态的按钮定义样式
        style.configure("Accent.TButton", foreground="white", background="#0078D7", font=self.font_button)
        style.configure("Stop.TButton", foreground="white", background="#E81123", font=self.font_button)
        
        self.start_button = ttk.Button(control_frame, text="开始生成", command=self.toggle_generation, style="Accent.TButton")
        self.start_button.pack(pady=5, ipady=5, ipadx=20)

        progress_frame = ttk.LabelFrame(right_panel, text="处理进度与日志", padding=10)
        progress_frame.pack(fill=tk.BOTH, expand=True)
        
        self.progress_bar = ttk.Progressbar(progress_frame, orient='horizontal', mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))
        
        self.log_text = scrolledtext.ScrolledText(progress_frame, height=5, state='disabled', font=self.font_normal, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        for tag, config in self.log_text_tags.items():
            self.log_text.tag_config(tag, **config)

    def create_file_selector(self, parent, label_text, string_var, command):
        row_frame = ttk.Frame(parent)
        row_frame.pack(fill=tk.X, expand=True, pady=1)
        ttk.Label(row_frame, text=label_text, width=12, font=self.font_normal).pack(side=tk.LEFT)
        ttk.Entry(row_frame, textvariable=string_var, state='readonly', font=self.font_normal).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(row_frame, text="浏览...", command=command, width=8).pack(side=tk.RIGHT)

    def select_excel_file(self):
        path = filedialog.askopenfilename(title="选择出院患者记录单", filetypes=[("Excel文件", "*.xlsx *.xls")])
        if path: self.excel_path_var.set(path)

    def select_surgery_query_files(self):
        paths = filedialog.askopenfilenames(title="选择一个或多个手术查询文件", filetypes=[("Excel文件", "*.xlsx *.xls")])
        if paths:
            for path in paths:
                if path not in self.surgery_query_files:
                    self.surgery_query_files.append(path)
                    self.surgery_listbox.insert(tk.END, os.path.basename(path))

    def clear_surgery_query_files(self):
        self.surgery_query_files.clear()
        self.surgery_listbox.delete(0, tk.END)

    def select_template_file(self):
        path = filedialog.askopenfilename(title="选择随访表模板", filetypes=[("Word模板", "*.docx")])
        if path: self.template_path_var.set(path)

    def select_output_dir(self):
        path = filedialog.askdirectory(title="选择保存位置")
        if path: self.output_dir_var.set(path)

    def log_message(self, msg, level="info"):
        """向日志文本框中添加带时间戳的消息。"""
        if not msg or not str(msg).strip():
            return

        def append():
            self.log_text.config(state='normal')
            lines = str(msg).split('\n')
            timestamp = datetime.now().strftime('%H:%M:%S')
            
            for line in lines:
                if line.strip():
                    full_log_line = f"{timestamp} - {line}\n"
                    if level in self.log_text_tags:
                        self.log_text.insert(tk.END, full_log_line, (level,))
                    else:
                        self.log_text.insert(tk.END, full_log_line)

            self.log_text.config(state='disabled')
            self.log_text.see(tk.END)
            
        self.root.after(0, append)

    def log_raw(self, msg):
        """向日志框中添加不带时间戳的原始文本。"""
        def append():
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, str(msg) + '\n')
            self.log_text.config(state='disabled')
            self.log_text.see(tk.END)
        self.root.after(0, append)

    def update_progress(self, value):
        self.root.after(0, lambda: self.progress_bar.config(value=value))

    def toggle_generation(self):
        """根据当前状态，开始或停止文档生成过程。"""
        if self.generation_thread and self.generation_thread.is_alive():
            # 如果线程正在运行，则发送停止信号
            if self.generator_instance:
                self.generator_instance.stop()
            self.start_button.config(state='disabled', text="正在停止...")
        else:
            # 如果没有线程在运行，则开始新的生成过程
            if not all([self.excel_path_var.get(), self.template_path_var.get(), self.output_dir_var.get()]):
                messagebox.showwarning("信息不全", "请先选择好“出院患者列表”、“Word模板”和“输出文件夹”。")
                return
            if not self.surgery_query_files:
                if not messagebox.askyesno("确认操作", "您没有选择任何“手术查询文件”。\n程序将无法补充床号，是否继续？"):
                    return
            
            self.start_button.config(text="停止生成", style="Stop.TButton")
            self.progress_bar['value'] = 0
            self.log_text.config(state='normal'); self.log_text.delete('1.0', tk.END); self.log_text.config(state='disabled')
            
            self.generator_instance = DocumentGenerator(
                excel_path=self.excel_path_var.get(), 
                surgery_query_paths=self.surgery_query_files,
                template_path=self.template_path_var.get(), 
                output_dir=self.output_dir_var.get(), 
                app_instance=self
            )
            self.generation_thread = threading.Thread(target=self.generator_instance.run, daemon=True)
            self.generation_thread.start()

    def generation_finished(self):
        """当生成线程结束（无论是正常完成还是被停止）时，由线程本身调用此方法来更新UI。"""
        def _update_ui():
            self.start_button.config(state='normal', text="开始生成", style="Accent.TButton")
            self.generation_thread = None
            self.generator_instance = None
        
        self.root.after(0, _update_ui)
