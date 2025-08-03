# -*- coding: utf-8 -*-
"""
GUI应用主逻辑模块。
此类负责窗口管理、状态维护和用户交互的响应。
"""
import os
import sys
import threading
import subprocess # 导入subprocess用于跨平台打开文件
from datetime import datetime
from pathlib import Path # 导入Path类
import tkinter as tk
from tkinter import filedialog, messagebox

from ..config import CONFIG, UI_CONFIG # 分别导入业务逻辑和UI配置
from ..core.logic import DocumentGenerator
from . import ui_builder # 导入新的UI构建模块
from .. import temp_manager # 导入新的临时文件管理器

class MainApp:
    def __init__(self, root, base_path): # base_path 是一个Path对象
        self.root = root
        self.base_path = base_path # 保存根目录路径 (Path对象)
        
        # --- 状态和变量 ---
        self.surgery_query_files = []
        self.generation_thread = None
        self.generator_instance = None
        
        # 用于UI显示的StringVar
        self.excel_display_var = tk.StringVar()
        self.template_display_var = tk.StringVar()
        self.output_dir_display_var = tk.StringVar()

        # 用于存储完整路径的实例变量 (仍然是字符串，因为它们来自filedialog)
        self.excel_full_path = ""
        self.template_full_path = ""
        self.output_dir_full_path = ""
        
        # --- 初始化设置 ---
        self.scaling_factor = self._get_scaling_factor()
        self._load_fonts_from_config() # 从配置文件加载字体
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

    def _load_fonts_from_config(self):
        """从config.py动态加载所有字体设置。"""
        # 直接从UI_CONFIG中获取字体
        font_config = UI_CONFIG['fonts']
        for name, font_tuple in font_config.items():
            # 使用setattr动态创建实例的字体属性
            # e.g., self.font_normal = ("微软雅黑", 9)
            setattr(self, f"font_{name}", font_tuple)

    def setup_window(self):
        """设置窗口大小并居中"""
        app_info = UI_CONFIG['app_info']
        layout_config = UI_CONFIG['layout']['window']

        window_title = f"{app_info['title']} V{app_info['version']}"
        self.root.title(window_title)
        
        s = self.scaling_factor
        # 从配置中计算窗口大小
        width = int(layout_config['base_width'] * s)
        height = int(layout_config['base_height'] * s)
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        pos_x = (screen_width // 2) - (width // 2)
        pos_y = (screen_height // 2) - (height // 2)
        
        self.root.geometry(f"{width}x{height}+{pos_x}+{pos_y}")
        
        # 从配置中设置窗口是否可调整大小
        resizable = layout_config['resizable']
        self.root.resizable(resizable, resizable)

    def _display_welcome_message(self):
        """在日志区显示欢迎和提示信息。"""
        ui_texts = UI_CONFIG['texts']
        separator = ui_texts.get("log_separator", "---")
        self.log(ui_texts.get("welcome_message", ""), add_timestamp=False)
        self.log(separator, add_timestamp=False)
        self.log(ui_texts.get("welcome_tips", ""), level="info", add_timestamp=False)
        self.log(ui_texts.get("welcome_warning", ""), level="error", add_timestamp=False)
        self.log(separator, add_timestamp=False)

    def _open_file_cross_platform(self, file_path):
        """跨平台安全地打开文件或文件夹。"""
        try:
            path_str = str(file_path) # 确保是字符串
            if sys.platform == "win32":
                os.startfile(path_str)
            elif sys.platform == "darwin": # macOS
                subprocess.run(["open", path_str], check=True)
            else: # Linux and other Unix-like
                subprocess.run(["xdg-open", path_str], check=True)
        except (FileNotFoundError, subprocess.CalledProcessError) as e:
             messagebox.showerror("打开失败", f"无法打开文件或目录：\n{file_path}\n\n错误: {e}")
        except Exception as e:
            messagebox.showerror("打开失败", f"发生未知错误：\n{e}")

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
            # 使用 pathlib 构建路径
            original_path = self.base_path / 'templates' / template_name
            if not original_path.exists():
                messagebox.showerror("错误", f"模板文件未找到！\n请确保 '{template_name}' 文件存在于 'templates' 文件夹中。")
                return

            # 创建一个临时的、只读的副本 (返回Path对象)
            temp_path = temp_manager.create_temp_read_only_copy(str(original_path))

            if temp_path:
                self._open_file_cross_platform(temp_path)
                self.log(f"已打开模板: {temp_path.name}", level="info")
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
            self.excel_display_var.set(Path(path).name) # 使用Path().name获取文件名
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
                    self.surgery_listbox.insert(tk.END, Path(path).name) # 使用Path().name获取文件名
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
            self.template_display_var.set(Path(path).name) # 使用Path().name获取文件名
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
            self.output_dir_display_var.set(Path(path).name) # 使用Path().name获取文件夹名
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
            if not self.excel_full_path:
                self.show_message("warning", "信息不全", "请选择“出院患者列表”。")
                return
            if not self.output_dir_full_path:
                self.show_message("warning", "信息不全", "请选择“输出文件夹”。")
                return
            if not self.template_full_path and self.template_display_var.get() != "[使用内置模板]":
                self.show_message("warning", "信息不全", "请选择一个Word模板或点击“使用内置模板”。")
                return

            template_path = ""
            if self.template_display_var.get() == "[使用内置模板]":
                # 使用 pathlib 构建路径
                template_path_obj = self.base_path / 'templates' / CONFIG['follow_up_template_name']
                if not template_path_obj.exists():
                    self.show_message("error", "错误", f"内置Word模板未找到！\n请确保 '{CONFIG['follow_up_template_name']}' 文件存在于 'templates' 文件夹中。")
                    return
                template_path = str(template_path_obj) # 传递字符串路径给核心逻辑
            else:
                template_path = self.template_full_path
            
            self.progress_bar['value'] = 0
            self.log(UI_CONFIG['texts'].get("log_separator", "---"), add_timestamp=False)
            
            self.start_button.config(text="停止生成", style="Stop.TButton")
            
            # --- 解耦：使用回调函数替代传递整个app实例 ---
            self.generator_instance = DocumentGenerator(
                excel_path=self.excel_full_path, 
                surgery_query_paths=self.surgery_query_files,
                template_path=template_path, 
                output_dir=self.output_dir_full_path,
                log_callback=self.log,
                progress_callback=self.update_progress,
                completion_callback=self.generation_finished,
                message_callback=self.show_message
            )
            self.generation_thread = threading.Thread(target=self.generator_instance.run, daemon=True)
            self.generation_thread.start()

    def show_message(self, level, title, message):
        """根据级别显示不同类型的消息框，确保线程安全。"""
        def _show():
            if level == "error":
                messagebox.showerror(title, message)
            elif level == "warning":
                messagebox.showwarning(title, message)
            elif level == "info":
                messagebox.showinfo(title, message)
            else: # 默认为 info
                messagebox.showinfo(title, message)
        self.root.after(0, _show)
        
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
            self.show_message("info", "提示", "日志内容为空，无需导出。")
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
            # 使用 with Path(filepath).open(...) 确保正确处理
            with Path(filepath).open('w', encoding='utf-8') as f:
                f.write(log_content)
            self.log(f"日志已成功导出到: {filepath}", level="info")
            self.show_message("info", "成功", f"日志已成功导出到:\n{filepath}")
        except Exception as e:
            self.log(f"导出日志失败: {e}", level="error")
            self.show_message("error", "导出失败", f"无法将日志保存到指定位置。\n错误: {e}")

    def log(self, msg, level=None, add_timestamp=True):
        """
        统一的日志记录方法。
        :param msg: 要记录的消息。
        :param level: 日志级别 ('info', 'error', 'warning', None)，用于文本着色。None为默认颜色。
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
