# -*- coding: utf-8 -*-
"""
GUI应用控制器模块。
此类负责窗口管理、状态维护，并协调UI和业务逻辑。
"""
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from datetime import datetime

from ..config import UI_CONFIG
from . import ui_builder
from .handlers import Handlers

class AppController:
    """应用程序的主控制器类"""
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

        # 用于存储完整路径的实例变量
        self.excel_full_path = ""
        self.template_full_path = ""
        self.output_dir_full_path = ""
        
        # --- 初始化设置 ---
        self.scaling_factor = self._get_scaling_factor()
        self._load_fonts_from_config() # 从配置文件加载字体
        self.setup_window()
        
        # --- 实例化事件处理器 ---
        # 将自身实例(self)传递给处理器，以便处理器能访问和修改AppController的状态
        self.handlers = Handlers(self)
        
        # --- 构建UI ---
        # 将UI构建委托给ui_builder模块，并传入事件处理器
        ui_builder.create_ui(self, self.handlers)

        self.listbox_original_bg = self.surgery_listbox.cget("background")

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
        font_config = UI_CONFIG['fonts']
        for name, font_tuple in font_config.items():
            setattr(self, f"font_{name}", font_tuple)

    def setup_window(self):
        """设置窗口大小并居中"""
        app_info = UI_CONFIG['app_info']
        layout_config = UI_CONFIG['layout']['window']

        window_title = f"{app_info['title']} V{app_info['version']}"
        self.root.title(window_title)
        
        s = self.scaling_factor
        width = int(layout_config['base_width'] * s)
        height = int(layout_config['base_height'] * s)
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        pos_x = (screen_width // 2) - (width // 2)
        pos_y = (screen_height // 2) - (height // 2)
        
        self.root.geometry(f"{width}x{height}+{pos_x}+{pos_y}")
        
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

    def _set_ui_busy(self, is_busy):
        """
        设置界面的忙碌状态，通过禁用/启用控件实现。
        :param is_busy: 布尔值，True表示忙碌，False表示空闲。
        """
        cursor_type = "watch" if is_busy else ""
        disabled_bg = UI_CONFIG['colors']['disabled_bg']

        self.root.config(cursor=cursor_type)
        
        for widget in self.interactive_widgets:
            # --- 修复：当开始生成(is_busy=True)时，不禁用“停止”按钮，以便用户可以点击它 ---
            if is_busy and widget == self.start_button:
                continue

            try:
                if isinstance(widget, tk.Listbox):
                    new_state = tk.DISABLED if is_busy else tk.NORMAL
                    new_bg = disabled_bg if is_busy else self.listbox_original_bg
                    widget.config(state=new_state, bg=new_bg)
                elif isinstance(widget, ttk.Entry):
                    new_state = tk.DISABLED if is_busy else 'readonly'
                    widget.config(state=new_state)
                else:
                    new_state = tk.DISABLED if is_busy else tk.NORMAL
                    widget.config(state=new_state)
            except tk.TclError:
                pass

    # --- 核心回调方法 (由其他模块调用) ---
    def show_message(self, level, title, message):
        """根据级别显示不同类型的消息框，确保线程安全。"""
        def _show():
            if level == "error":
                messagebox.showerror(title, message)
            elif level == "warning":
                messagebox.showwarning(title, message)
            else: # 默认为 info
                messagebox.showinfo(title, message)
        self.root.after(0, _show)
        
    def generation_finished(self):
        """当生成线程结束时，由线程本身调用此方法来更新UI。"""
        def _update_ui():
            self._set_ui_busy(False)
            self.start_button.config(state='normal', text="开始生成", style="Accent.TButton")
            self.generation_thread = None
            self.generator_instance = None
        
        self.root.after(0, _update_ui)

    def log(self, msg, level=None, add_timestamp=True):
        """
        统一的日志记录方法。
        :param msg: 要记录的消息。
        :param level: 日志级别 ('info', 'error', 'warning', None)，用于文本着色。
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
        """线程安全地更新进度条。"""
        self.root.after(0, lambda: self.progress_bar.config(value=value))
