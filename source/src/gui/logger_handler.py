# -*- coding: utf-8 -*-
"""
自定义Tkinter日志处理器。
"""
import logging
import tkinter as tk

class TkinterLogHandler(logging.Handler):
    """
    一个将日志记录发送到Tkinter ScrolledText小部件的处理器。
    它支持通过 'extra' 字典动态控制日志格式。
    """
    def __init__(self, widget):
        super().__init__()
        self.widget = widget
        self.tags = {}
        self.default_formatter = logging.Formatter('%(asctime)s - %(message)s', datefmt='%H:%M:%S')
        self.simple_formatter = logging.Formatter('%(message)s')

    def set_tags(self, tags_config):
        """设置用于不同日志级别的颜色标签。"""
        self.tags = tags_config

    def emit(self, record):
        """
        线程安全地将日志记录发送到UI小部件。
        """
        # --- 优化：根据 'simple' 标志选择格式化器 ---
        is_simple = getattr(record, 'simple', False)
        msg = self.simple_formatter.format(record) if is_simple else self.default_formatter.format(record)
        
        level = record.levelname

        def append():
            self.widget.config(state='normal')
            
            # --- 优化：直接根据级别名称查找颜色标签 ---
            # 如果级别名称在tags配置中，则使用它，否则不使用任何标签（默认为黑色）
            tag_to_use = (level,) if level in self.tags else ()
            
            self.widget.insert(tk.END, f"{msg}\n", tag_to_use)
            
            self.widget.config(state='disabled')
            self.widget.see(tk.END)
        
        self.widget.after(0, append)
