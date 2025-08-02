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
from .. import temp_manager # 导入新的临时文件管理器

class MainApp:
    def __init__(self, root, base_path): # 增加 base_path 参数
        self.root = root
        self.base_path = base_path # 保存根目录路径
        
        # --- 状态和变量 ---
        self.surgery_query_files = []
        self.generation_thread = None
        self.generator_instance = None
        
        # 用于UI显示的StringVar
        self.excel_display_var = tk.StringVar()
        self.template_display_var = tk.StringVar()
        self.output_dir_display_var = tk.StringVar()

        # 用于存储完整路径的实例变量
        self.excel_full_path = ""
        self.template_full_path = ""
        self.output_dir_full_path = ""
        
        # --- 初始化设置 ---
        self.scaling_factor = self._get_scaling_factor()
        self.setup_fonts()
        self.setup_window()
        
        # --- 构建UI ---
        # 将UI构建委托给ui_builder模块
        ui_builder.create_ui(self)
        self._display_welcome_message()

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

    def _display_welcome_message(self):
        """在日志区显示欢迎和提示信息。"""
        separator = CONFIG.get("log_separator", "---")
        self.log(CONFIG.get("welcome_message", ""), add_timestamp=False)
        self.log(separator, add_timestamp=False)
        self.log(CONFIG.get("welcome_tips", ""), level="info", add_timestamp=False)
        self.log(CONFIG.get("welcome_warning", ""), level="error", add_timestamp=False) # 添加警告语
        self.log(separator, add_timestamp=False)

    def open_template(self, template_type):
        """将模板复制到临时的只读文件并打开它。"""
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
            original_path = os.path.join(self.base_path, 'templates', template_name)
            if not os.path.exists(original_path):
                messagebox.showerror("错误", f"模板文件未找到！\n请确保 '{template_name}' 文件存在于 'templates' 文件夹中。")
                return

            # 创建一个临时的、只读的副本
            temp_path = temp_manager.create_temp_read_only_copy(original_path)

            if temp_path:
                os.startfile(temp_path)
                self.log(f"已打开模板: {os.path.basename(temp_path)}", level="info")
            else:
                messagebox.showerror("错误", "创建临时模板文件失败。")

        except Exception as e:
            messagebox.showerror("打开失败", f"无法打开模板文件。\n错误: {e}")

    def _select_path(self, selection_type, title, filetypes=None):
        """通用路径选择函数，用于文件或文件夹。"""
        if selection_type == 'file':
            path = filedialog.askopenfilename(title=title, filetypes=filetypes)
        elif selection_type == 'directory':
            path = filedialog.askdirectory(title=title)
        else:
            return None # 不应发生此情况
        return path

    def select_excel_file(self):
        """选择出院患者记录单Excel文件。"""
        path = self._select_path('file', "选择出院患者记录单", [("Excel文件", "*.xlsx *.xls")])
        if path:
            self.excel_full_path = path
            self.excel_display_var.set(os.path.basename(path))
            self.log(f"已选择出院患者列表: {path}", level="info")

    def clear_excel_selection(self):
        self.excel_full_path = ""
        self.excel_display_var.set("")
        self.log("已清空出院患者列表选择。", level="info")

    def select_surgery_query_files(self):
        paths = filedialog.askopenfilenames(title="选择一个或多个手术查询文件", filetypes=[("Excel文件", "*.xlsx *.xls")])
        if paths:
            for path in paths:
                if path not in self.surgery_query_files:
                    self.surgery_query_files.append(path)
                    self.surgery_listbox.insert(tk.END, os.path.basename(path))
                    self.log(f"已添加手术查询文件: {path}", level="info")

    def clear_surgery_query_files(self):
        self.surgery_query_files.clear()
        self.surgery_listbox.delete(0, tk.END)
        self.log("已清空手术查询文件列表。", level="info")

    def select_template_file(self):
        """选择自定义的Word模板文件。"""
        path = self._select_path('file', "选择随访表模板", [("Word模板", "*.docx *.doc")])
        if path:
            self.template_full_path = path
            self.template_display_var.set(os.path.basename(path))
            self.log(f"已选择Word模板: {path}", level="info")

    def use_builtin_word_template(self):
        """设置UI以表明正在使用内置模板。"""
        self.template_full_path = "" # 清空外部路径
        self.template_display_var.set("[使用内置模板]")
        self.log("已选择使用内置Word模板。", level="info")
        
    def clear_template_selection(self):
        self.template_full_path = ""
        self.template_display_var.set("")
        self.log("已清空Word模板选择。", level="info")

    def select_output_dir(self):
        """选择用于保存生成文档的输出文件夹。"""
        path = self._select_path('directory', "选择保存位置")
        if path:
            self.output_dir_full_path = path
            self.output_dir_display_var.set(os.path.basename(path))
            self.log(f"已选择输出文件夹: {path}", level="info")
            
    def clear_output_dir_selection(self):
        self.output_dir_full_path = ""
        self.output_dir_display_var.set("")
        self.log("已清空输出文件夹选择。", level="info")

    def toggle_generation(self):
        """根据当前状态，开始或停止文档生成过程。"""
        if self.generation_thread and self.generation_thread.is_alive():
            if self.generator_instance:
                self.generator_instance.stop()
            self.start_button.config(state='disabled', text="正在停止...")
        else:
            # --- 优化：检查所有必填项 ---
            if not self.excel_full_path:
                messagebox.showwarning("信息不全", "请选择“出院患者列表”。")
                return
            if not self.output_dir_full_path:
                messagebox.showwarning("信息不全", "请选择“输出文件夹”。")
                return
            if not self.template_full_path and self.template_display_var.get() != "[使用内置模板]":
                messagebox.showwarning("信息不全", "请选择一个Word模板或点击“使用内置模板”。")
                return
            # ---------------------------

            template_path = ""
            if self.template_display_var.get() == "[使用内置模板]":
                template_path = os.path.join(self.base_path, 'templates', CONFIG['follow_up_template_name'])
                if not os.path.exists(template_path):
                    messagebox.showerror("错误", f"内置Word模板未找到！\n请确保 '{CONFIG['follow_up_template_name']}' 文件存在于 'templates' 文件夹中。")
                    return
            else:
                template_path = self.template_full_path
            
            # --- 优化：不清空日志，只重置进度条并添加分隔符 ---
            self.progress_bar['value'] = 0
            self.log(CONFIG.get("log_separator", "---"), add_timestamp=False)
            # ---------------------------------------------
            
            self.start_button.config(text="停止生成", style="Stop.TButton")
            
            self.generator_instance = DocumentGenerator(
                excel_path=self.excel_full_path, 
                surgery_query_paths=self.surgery_query_files,
                template_path=template_path, 
                output_dir=self.output_dir_full_path, 
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

    def clear_log(self):
        """清空日志区域的内容并恢复欢迎语。"""
        if messagebox.askyesno("确认", "确定要清空所有日志内容吗？"):
            self.log_text.config(state='normal')
            self.log_text.delete('1.0', tk.END)
            self.log_text.config(state='disabled')
            self._display_welcome_message()

    def export_log(self):
        """将日志内容导出到纯文本文件。"""
        log_content = self.log_text.get('1.0', tk.END)
        if not log_content.strip():
            messagebox.showinfo("提示", "日志内容为空，无需导出。")
            return

        default_filename = f"随访表生成日志_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        filepath = filedialog.asksaveasfilename(
            title="导出日志文件",
            initialfile=default_filename,
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )

        if not filepath:
            self.log("用户取消了日志导出。", level="info")
            return

        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(log_content)
            self.log(f"日志已成功导出到: {filepath}", level="info")
            messagebox.showinfo("成功", f"日志已成功导出到:\n{filepath}")
        except Exception as e:
            self.log(f"导出日志失败: {e}", level="error")
            messagebox.showerror("导出失败", f"无法将日志保存到指定位置。\n错误: {e}")

    def log(self, msg, level=None, add_timestamp=True):
        """
        统一的日志记录方法。
        :param msg: 要记录的消息。
        :param level: 日志级别 ('info', 'error', 'success', None)，用于文本着色。None为默认颜色。
        :param add_timestamp: 是否在消息前添加时间戳。
        """
        if not msg or not str(msg).strip():
            return

        def append():
            self.log_text.config(state='normal')
            
            log_line = str(msg)
            if add_timestamp:
                timestamp = datetime.now().strftime('%H:%M:%S')
                log_line = f"{timestamp} - {log_line}"
            
            full_log_line = f"{log_line}\n"

            tag_to_use = ()
            if level and level in self.log_text_tags:
                tag_to_use = (level,)

            self.log_text.insert(tk.END, full_log_line, tag_to_use)
            self.log_text.config(state='disabled')
            self.log_text.see(tk.END)

        self.root.after(0, append)

    def update_progress(self, value):
        self.root.after(0, lambda: self.progress_bar.config(value=value))