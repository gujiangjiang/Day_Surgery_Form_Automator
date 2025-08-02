# -*- coding: utf-8 -*-
"""
GUI应用主逻辑模块。
此类负责窗口管理、状态维护和用户交互的响应。
"""
import os
import sys
import threading
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

from ..config import CONFIG
from ..core.logic import DocumentGenerator
from . import ui_builder # 导入新的UI构建模块

class MainApp:
    def __init__(self, root, base_path): # 增加 base_path 参数
        self.root = root
        self.base_path = base_path # 保存根目录路径
        
        # --- 状态和变量 ---
        self.surgery_query_files = []
        self.generation_thread = None
        self.generator_instance = None
        self.excel_path_var = tk.StringVar()
        self.template_path_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        
        # --- 初始化设置 ---
        self.scaling_factor = self._get_scaling_factor()
        self.setup_fonts()
        self.setup_window()
        
        # --- 构建UI ---
        # 将UI构建委托给ui_builder模块
        ui_builder.create_ui(self)

    def _get_scaling_factor(self):
        """获取屏幕缩放比例"""
        try:
            dpi = self.root.winfo_fpixels('1i')
            scaling = dpi / 96.0
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

    def open_template(self, template_type):
        """打开指定类型的模板文件。"""
        template_map = {
            'discharge': CONFIG['discharge_template_name'],
            'surgery': CONFIG['surgery_template_name'],
            'follow_up': CONFIG['follow_up_template_name']
        }
        template_name = template_map.get(template_type)
        if not template_name:
            messagebox.showerror("错误", "未知的模板类型。")
            return

        try:
            # 使用正确的根目录来构建路径
            path = os.path.join(self.base_path, 'templates', template_name)
            if os.path.exists(path):
                os.startfile(path)
            else:
                messagebox.showerror("错误", f"模板文件未找到！\n请确保 '{template_name}' 文件存在于 'templates' 文件夹中。")
        except Exception as e:
            messagebox.showerror("打开失败", f"无法打开模板文件，请确保您已安装对应的办公软件。\n错误: {e}")

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

    def use_builtin_word_template(self):
        """设置UI以表明正在使用内置模板。"""
        self.template_path_var.set("[使用内置模板]")

    def select_output_dir(self):
        path = filedialog.askdirectory(title="选择保存位置")
        if path: self.output_dir_var.set(path)

    def toggle_generation(self):
        """根据当前状态，开始或停止文档生成过程。"""
        if self.generation_thread and self.generation_thread.is_alive():
            if self.generator_instance:
                self.generator_instance.stop()
            self.start_button.config(state='disabled', text="正在停止...")
        else:
            if not all([self.excel_path_var.get(), self.output_dir_var.get()]):
                messagebox.showwarning("信息不全", "请先选择好“出院患者列表”和“输出文件夹”。")
                return
            
            template_path = self.template_path_var.get()
            if not template_path or template_path == "[使用内置模板]":
                self.log_message("未选择外部Word模板或已指定使用内置模板，将加载内置模板。", "info")
                template_path = os.path.join(self.base_path, 'templates', CONFIG['follow_up_template_name'])
                if not os.path.exists(template_path):
                    messagebox.showerror("错误", f"内置Word模板未找到！\n请确保 '{CONFIG['follow_up_template_name']}' 文件存在于 'templates' 文件夹中。")
                    return
            
            self.start_button.config(text="停止生成", style="Stop.TButton")
            self.progress_bar['value'] = 0
            self.log_text.config(state='normal'); self.log_text.delete('1.0', tk.END); self.log_text.config(state='disabled')
            
            self.generator_instance = DocumentGenerator(
                excel_path=self.excel_path_var.get(), 
                surgery_query_paths=self.surgery_query_files,
                template_path=template_path, 
                output_dir=self.output_dir_var.get(), 
                app_instance=self
            )
            self.generation_thread = threading.Thread(target=self.generator_instance.run, daemon=True)
            self.generation_thread.start()

    # --- 与后台线程通信的方法 ---

    def generation_finished(self):
        """当生成线程结束时，由线程本身调用此方法来更新UI。"""
        def _update_ui():
            self.start_button.config(state='normal', text="开始生成", style="Accent.TButton")
            self.generation_thread = None
            self.generator_instance = None
        
        self.root.after(0, _update_ui)

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
