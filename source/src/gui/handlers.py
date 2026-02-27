# -*- coding: utf-8 -*-
"""
UI事件处理器模块。
此类包含所有响应用户交互（如按钮点击、菜单选择）的方法。
"""
import os
import sys
import threading
import subprocess
import logging
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

from ..config import CONFIG, UI_CONFIG
from ..core.logic import DocumentGenerator
from .. import temp_manager

# 获取该模块的logger实例
logger = logging.getLogger(__name__)

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
                # Windows环境：os.startfile 在找不到关联程序时会抛出 OSError
                os.startfile(path_str)
            elif sys.platform == "darwin":
                # macOS环境：抑制终端错误输出，失败时会抛出 CalledProcessError
                subprocess.run(["open", path_str], check=True, stderr=subprocess.PIPE)
            else:
                # Linux环境：抑制终端错误输出，失败时会抛出 CalledProcessError
                subprocess.run(["xdg-open", path_str], check=True, stderr=subprocess.PIPE)
        except (FileNotFoundError, subprocess.CalledProcessError, OSError) as e:
            # 【修复日志显示】：在 GUI (Warning级别) 仅显示最干净、精简的文件名提示，不再显示冗长的临时路径和底层报错
            logger.warning(f"无法打开文件，系统可能未安装关联程序: {Path(file_path).name}")
            
            # 【调试日志隔离】：将底层的具体报错对象和完整临时路径通过 debug 级别记录
            # 这样它只会写入到后台的 app_runtime.log 文件中（用于开发者排查），绝对不会出现在用户的 GUI 界面上
            logger.debug(f"打开文件失败的完整路径: {file_path}, 底层详细信息: {e}")
            
            # 根据文件后缀名动态提供更友好的软件安装提示
            ext = Path(file_path).suffix.lower()
            app_hint = ""
            if ext in ['.xls', '.xlsx']:
                app_hint = "（建议安装 Microsoft Office Excel 或 WPS Office 等表格软件）"
            elif ext in ['.doc', '.docx']:
                app_hint = "（建议安装 Microsoft Office Word 或 WPS Office 等文档软件）"
                
            messagebox.showerror(
                "打开失败", 
                f"无法打开该文件：\n{Path(file_path).name}\n\n"
                f"您的系统中似乎没有关联能够打开此类文件的默认程序。\n"
                f"请确保已安装相应的办公软件{app_hint}后重试。"
            )
        except Exception as e:
            # 对于真正的未知错误，记录在 ERROR 级别
            logger.error(f"打开文件时发生未知错误", exc_info=True)
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

            # 调用更健壮的 create_temp_read_only_copy 函数
            temp_path = temp_manager.create_temp_read_only_copy(str(original_path))

            if temp_path:
                logger.info(f"尝试打开模板: {template_name}")
                self._open_file_cross_platform(temp_path)
            else:
                # 如果 temp_path 为 None，说明创建或清理失败
                logger.error(f"创建或清理临时文件失败: {template_name}")
                messagebox.showerror("操作失败", f"无法创建临时模板文件 '{template_name}'。\n\n这可能是因为旧的临时文件仍被其他程序（如Excel）占用。\n请关闭相关程序后重试。")

        except Exception as e:
            logger.error(f"无法打开模板文件", exc_info=True)
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
            logger.info(f"已选择出院患者列表: {path}")

    def clear_excel_selection(self):
        """清空出院患者列表选择。"""
        self.app.excel_full_path = ""
        self.app.excel_display_var.set("")
        logger.info("已清空出院患者列表选择。")

    def select_surgery_query_files(self):
        """选择手术查询文件。"""
        paths = filedialog.askopenfilenames(title="选择一个或多个手术查询文件", filetypes=[("Excel文件", "*.xlsx *.xls")])
        if paths:
            for path in paths:
                if path not in self.app.surgery_query_files:
                    self.app.surgery_query_files.append(path)
                    self.app.surgery_listbox.insert(tk.END, Path(path).name)
                    logger.info(f"已添加手术查询文件: {path}")

    def clear_surgery_query_files(self):
        """清空手术查询文件列表。"""
        self.app.surgery_query_files.clear()
        self.app.surgery_listbox.delete(0, tk.END)
        logger.info("已清空手术查询文件列表。")

    def select_template_file(self):
        """选择自定义的Word模板文件。"""
        path = self._select_path('file', "选择随访表模板", [("Word模板", "*.docx *.doc")])
        if path:
            self.app.template_full_path = path
            self.app.template_display_var.set(Path(path).name)
            logger.info(f"已选择Word模板: {path}")

    def use_builtin_word_template(self):
        """设置UI以表明正在使用内置模板。"""
        self.app.template_full_path = ""
        self.app.template_display_var.set("[使用内置模板]")
        logger.info("已选择使用内置Word模板。")
        
    def clear_template_selection(self):
        """清空Word模板选择。"""
        self.app.template_full_path = ""
        self.app.template_display_var.set("")
        logger.info("已清空Word模板选择。")

    def select_output_dir(self):
        """选择输出文件夹。"""
        path = self._select_path('directory', "选择保存位置")
        if path:
            self.app.output_dir_full_path = path
            self.app.output_dir_display_var.set(Path(path).name)
            logger.info(f"已选择输出文件夹: {path}")
            
    def clear_output_dir_selection(self):
        """清空输出文件夹选择。"""
        self.app.output_dir_full_path = ""
        self.app.output_dir_display_var.set("")
        logger.info("已清空输出文件夹选择。")

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
        logger.notice(UI_CONFIG['texts'].get("log_separator", "---"), extra={'simple': True})
        app.start_button.config(text="停止生成", style="Stop.TButton")
        app._set_ui_busy(True)

        app.generator_instance = DocumentGenerator(
            excel_path=app.excel_full_path, 
            surgery_query_paths=app.surgery_query_files,
            template_path=template_path, 
            output_dir=app.output_dir_full_path,
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
            logger.info("日志区域已手动清空。")

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
            logger.info("用户取消了日志导出。")
            return

        try:
            with Path(filepath).open('w', encoding='utf-8') as f:
                f.write(log_content)
            logger.info(f"日志已成功导出到: {filepath}")
            app.show_message("info", "成功", f"日志已成功导出到:\n{filepath}")
        except Exception as e:
            logger.error(f"导出日志失败", exc_info=True)
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
