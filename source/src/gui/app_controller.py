# -*- coding: utf-8 -*-
"""
GUI应用控制器模块。
此类负责窗口管理、状态维护，并协调UI和业务逻辑。
"""
import sys # 导入sys模块，用于跨平台判断
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import logging # 导入logging模块

from ..config import UI_CONFIG
from . import ui_builder
from .handlers import Handlers
from .logger_handler import TkinterLogHandler

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
        self._load_fonts_from_config()
        self.setup_window()
        
        # 实例化事件处理器
        # 将自身实例(self)传递给处理器，以便处理器能访问和修改AppController的状态
        self.handlers = Handlers(self)
        
        # --- 构建UI ---
        # 将UI构建委托给ui_builder模块，并传入事件处理器
        ui_builder.create_ui(self, self.handlers)

        # 配置UI日志处理器
        self._setup_ui_logging()

        self.listbox_original_bg = self.surgery_listbox.cget("background")

        self._display_welcome_message()

    def _setup_ui_logging(self):
        """配置并添加Tkinter日志处理器到根logger。"""
        ui_log_handler = TkinterLogHandler(self.log_text)
        # 从配置中获取颜色标签并设置给处理器
        color_tags = UI_CONFIG['colors']['log_tags']
        ui_log_handler.set_tags(color_tags)
        
        # 将UI处理器添加到根logger
        logging.getLogger().addHandler(ui_log_handler)
        
        # 为不同级别的日志配置颜色
        for level, config in color_tags.items():
            self.log_text.tag_config(level, **config)

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
        
        # 使用 'extra' 参数来控制UI日志的格式
        # 1. 普通欢迎语: 使用自定义的NOTICE级别，无颜色，无时间戳
        logging.getLogger().notice(ui_texts.get("welcome_message", ""), extra={'simple': True})
        logging.getLogger().notice(separator, extra={'simple': True})
        
        # 2. 提示信息: 使用INFO级别，有蓝色，无时间戳
        logging.info(ui_texts.get("welcome_tips", ""), extra={'simple': True})
        
        # 3. 警告信息: 使用CRITICAL级别，有红色，无时间戳
        logging.critical(ui_texts.get("welcome_warning", ""), extra={'simple': True})
        
        logging.getLogger().notice(separator, extra={'simple': True})

    def _set_ui_busy(self, is_busy):
        """
        设置界面的忙碌状态，通过禁用/启用控件实现。
        :param is_busy: 布尔值，True表示忙碌，False表示空闲。
        """
        cursor_type = "watch" if is_busy else ""
        disabled_bg = UI_CONFIG['colors']['disabled_bg']

        self.root.config(cursor=cursor_type)
        
        for widget in self.interactive_widgets:
            # 当开始生成(is_busy=True)时，不禁用“停止”按钮，以便用户可以点击它
            if is_busy and widget == self.start_button:
                continue
            try:
                if isinstance(widget, tk.Listbox):
                    new_state = tk.DISABLED if is_busy else tk.NORMAL
                    # 【修复深色模式白屏Bug】：在 macOS 下，为了兼容深色模式，不要硬编码将背景色设为亮灰色。
                    # 仅修改状态让系统自动置灰文字，以保持深色背景不突变。
                    if sys.platform == "darwin":
                        widget.config(state=new_state)
                    else:
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
        """线程安全地显示消息框。"""
        show_func = {
            "error": messagebox.showerror,
            "warning": messagebox.showwarning,
            "info": messagebox.showinfo
        }.get(level, messagebox.showinfo)
        self.root.after(0, lambda: show_func(title, message))
        
    def generation_finished(self):
        """当生成线程结束时，由线程本身调用此方法来更新UI。"""
        def _update_ui():
            self._set_ui_busy(False)
            self.start_button.config(state='normal', text="开始生成", style="Accent.TButton")
            self.generation_thread = None
            self.generator_instance = None
        
        self.root.after(0, _update_ui)

    def update_progress(self, value):
        """线程安全地更新进度条。"""
        self.root.after(0, lambda: self.progress_bar.config(value=value))
