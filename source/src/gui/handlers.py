# -*- coding: utf-8 -*-
"""
UI事件处理器模块。
此类包含所有响应用户交互（如按钮点击、菜单选择）的方法。
"""
import os
import sys
import threading
import subprocess
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

from ..config import CONFIG, UI_CONFIG
from ..core.logic import DocumentGenerator
from .. import temp_manager

class Handlers:
    def __init__(self, app):
        """
        初始化事件处理器。
        :param app: AppController 的实例，用于访问和修改主应用的状态和UI。
        """
        self.app = app

    def _open_file_cross_platform(self, file_path):
        """跨平台安全地打开文件或文件夹。"""
        try:
            path_str = str(file_path)
            if sys.platform == "win32":
                os.startfile(path_str)
            elif sys.platform == "darwin":
                subprocess.run(["open", path_str], check=True)
            else:
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
            original_path = self.app.base_path / 'templates' / template_name
            if not original_path.exists():
                messagebox.showerror("错误", f"模板文件未找到！\n请确保 '{template_name}' 文件存在于 'templates' 文件夹中。")
                return

            temp_path = temp_manager.create_temp_read_only_copy(str(original_path))

            if temp_path:
                self._open_file_cross_platform(temp_path)
                self.app.log(f"已打开模板: {temp_path.name}", level="info")
            else:
                messagebox.showerror("错误", "创建临时模板文件失败。")

        except Exception as e:
            messagebox.showerror("打开失败", f"无法打开模板文件。\n错误: {e}")

    def _select_path(self, selection_type, title, filetypes=None):
        """通用路径选择函数。"""
        if selection_type == 'file':
            path = filedialog.askopenfilename(title=title, filetypes=filetypes)
        elif selection_type == 'directory':
            path = filedialog.askdirectory(title=title)
        else:
            return None
        return path

    def select_excel_file(self):
        """选择出院患者记录单Excel文件。"""
        path = self._select_path('file', "选择出院患者记录单", [("Excel文件", "*.xlsx *.xls")])
        if path:
            self.app.excel_full_path = path
            self.app.excel_display_var.set(Path(path).name)
            self.app.log(f"已选择出院患者列表: {path}", level="info")

    def clear_excel_selection(self):
        """清空出院患者列表选择。"""
        self.app.excel_full_path = ""
        self.app.excel_display_var.set("")
        self.app.log("已清空出院患者列表选择。", level="info")

    def select_surgery_query_files(self):
        """选择手术查询文件。"""
        paths = filedialog.askopenfilenames(title="选择一个或多个手术查询文件", filetypes=[("Excel文件", "*.xlsx *.xls")])
        if paths:
            for path in paths:
                if path not in self.app.surgery_query_files:
                    self.app.surgery_query_files.append(path)
                    self.app.surgery_listbox.insert(tk.END, Path(path).name)
                    self.app.log(f"已添加手术查询文件: {path}", level="info")

    def clear_surgery_query_files(self):
        """清空手术查询文件列表。"""
        self.app.surgery_query_files.clear()
        self.app.surgery_listbox.delete(0, tk.END)
        self.app.log("已清空手术查询文件列表。", level="info")

    def select_template_file(self):
        """选择自定义的Word模板文件。"""
        path = self._select_path('file', "选择随访表模板", [("Word模板", "*.docx *.doc")])
        if path:
            self.app.template_full_path = path
            self.app.template_display_var.set(Path(path).name)
            self.app.log(f"已选择Word模板: {path}", level="info")

    def use_builtin_word_template(self):
        """设置UI以表明正在使用内置模板。"""
        self.app.template_full_path = ""
        self.app.template_display_var.set("[使用内置模板]")
        self.app.log("已选择使用内置Word模板。", level="info")
        
    def clear_template_selection(self):
        """清空Word模板选择。"""
        self.app.template_full_path = ""
        self.app.template_display_var.set("")
        self.app.log("已清空Word模板选择。", level="info")

    def select_output_dir(self):
        """选择输出文件夹。"""
        path = self._select_path('directory', "选择保存位置")
        if path:
            self.app.output_dir_full_path = path
            self.app.output_dir_display_var.set(Path(path).name)
            self.app.log(f"已选择输出文件夹: {path}", level="info")
            
    def clear_output_dir_selection(self):
        """清空输出文件夹选择。"""
        self.app.output_dir_full_path = ""
        self.app.output_dir_display_var.set("")
        self.app.log("已清空输出文件夹选择。", level="info")

    def _validate_inputs(self):
        """验证所有必需的输入项。"""
        if not self.app.excel_full_path:
            return False, "请选择“出院患者列表”。"
        if not self.app.output_dir_full_path:
            return False, "请选择“输出文件夹”。"
        if not self.app.template_full_path and self.app.template_display_var.get() != "[使用内置模板]":
            return False, "请选择一个Word模板或点击“使用内置模板”。"
        return True, ""

    def _get_template_path(self):
        """获取有效的模板路径。"""
        if self.app.template_full_path:
            return self.app.template_full_path, ""
        
        if self.app.template_display_var.get() == "[使用内置模板]":
            template_path_obj = self.app.base_path / 'templates' / CONFIG['follow_up_template_name']
            if not template_path_obj.exists():
                error_msg = f"内置Word模板未找到！\n请确保 '{CONFIG['follow_up_template_name']}' 文件存在于 'templates' 文件夹中。"
                return None, error_msg
            return str(template_path_obj), ""
        
        return None, "未知的模板配置错误。"

    def toggle_generation(self):
        """开始或停止文档生成过程。"""
        app = self.app
        if app.generation_thread and app.generation_thread.is_alive():
            if app.generator_instance:
                app.generator_instance.stop()
            app.start_button.config(state='disabled', text="正在停止...")
            return

        is_valid, error_message = self._validate_inputs()
        if not is_valid:
            app.show_message("warning", "信息不全", error_message)
            return

        template_path, error_message = self._get_template_path()
        if not template_path:
            app.show_message("error", "错误", error_message)
            return
        
        app.progress_bar['value'] = 0
        app.log(UI_CONFIG['texts'].get("log_separator", "---"), add_timestamp=False)
        app.start_button.config(text="停止生成", style="Stop.TButton")
        app._set_ui_busy(True)

        app.generator_instance = DocumentGenerator(
            excel_path=app.excel_full_path, 
            surgery_query_paths=app.surgery_query_files,
            template_path=template_path, 
            output_dir=app.output_dir_full_path,
            log_callback=app.log,
            progress_callback=app.update_progress,
            completion_callback=app.generation_finished,
            message_callback=app.show_message
        )
        app.generation_thread = threading.Thread(target=app.generator_instance.run, daemon=True)
        app.generation_thread.start()

    def clear_log(self):
        """清空日志区域的内容。"""
        app = self.app
        if messagebox.askyesno("确认", "确定要清空所有日志内容吗？"):
            app.log_text.config(state='normal')
            app.log_text.delete('1.0', tk.END)
            app.log_text.config(state='disabled')
            app._display_welcome_message()

    def export_log(self):
        """将日志内容导出到文件。"""
        app = self.app
        log_content = app.log_text.get('1.0', tk.END)
        if not log_content.strip():
            app.show_message("info", "提示", "日志内容为空，无需导出。")
            return

        default_filename = f"随访表生成日志_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        filepath = filedialog.asksaveasfilename(
            title="导出日志文件",
            initialfile=default_filename,
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )

        if not filepath:
            app.log("用户取消了日志导出。", level="info")
            return

        try:
            with Path(filepath).open('w', encoding='utf-8') as f:
                f.write(log_content)
            app.log(f"日志已成功导出到: {filepath}", level="info")
            app.show_message("info", "成功", f"日志已成功导出到:\n{filepath}")
        except Exception as e:
            app.log(f"导出日志失败: {e}", level="error")
            app.show_message("error", "导出失败", f"无法将日志保存到指定位置。\n错误: {e}")

    def show_about_dialog(self):
        """显示“关于”对话框。"""
        app = self.app
        app_info = UI_CONFIG['app_info']
        ui_texts = UI_CONFIG['texts']
        
        title = ui_texts['about_title']
        content = ui_texts['about_content'].format(
            title=app_info['title'],
            version=app_info['version'],
            internal_version=app_info['internal_version'],
            author=app_info['author'],
            build_date=app_info['build_date']
        )
        app.show_message("info", title, content)
