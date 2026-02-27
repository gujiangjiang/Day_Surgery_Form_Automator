# -*- coding: utf-8 -*-
"""
模块功能：提供一个可重用的Tooltip类，用于为Tkinter控件添加悬停提示。
"""
import tkinter as tk
import sys # 引入 sys 用于操作系统判定

class Tooltip:
    """
    创建一个当鼠标悬停在控件上时显示的提示框。
    """
    def __init__(self, widget, text, wraplength=200):
        self.widget = widget
        self.text = text
        self.wraplength = wraplength # 提示文本的最大换行宽度
        self.tooltip_window = None
        self.id = None
        self.x = self.y = 0
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)

    def enter(self, event=None):
        """当鼠标进入控件时触发。"""
        self.schedule()

    def leave(self, event=None):
        """当鼠标离开控件时触发。"""
        self.unschedule()
        self.hidetip()

    def schedule(self):
        """安排在一小段时间后显示提示。"""
        self.unschedule()
        self.id = self.widget.after(500, self.showtip) # 500毫秒延迟

    def unschedule(self):
        """取消已安排的提示显示。"""
        id = self.id
        self.id = None
        if id:
            self.widget.after_cancel(id)

    def showtip(self, event=None):
        """
        显示提示窗口。
        此方法经过修改，以通用方式定位提示框，避免特定控件的兼容性问题。
        """
        # --- 错误修复 ---
        # 移除 self.widget.bbox("insert")，因为它不适用于所有控件（如Listbox）。
        # 改为使用控件的winfo_rootx/y和winfo_height来通用地定位提示框。
        # 这会将提示框放置在控件正下方，并带有一些偏移量。
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5

        self.tooltip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True) # 无边框窗口
        tw.wm_geometry(f"+{int(x)}+{int(y)}") # 确保坐标是整数

        # 【跨平台 UI 优化】：根据不同操作系统动态调整提示框的颜色、边框和字体风格
        if sys.platform == "darwin":
            # macOS 现代风格：深灰底色、纯白文字、无边框、苹方字体
            bg_color = "#323232"
            fg_color = "#ffffff"
            relief_style = "flat"
            bd_width = 0
            font_style = ("PingFang SC", 12, "normal")
        elif sys.platform == "win32":
            # Windows 经典风格：浅黄底色、黑色文字、实线边框、微软雅黑
            bg_color = "#ffffe0"
            fg_color = "#000000"
            relief_style = "solid"
            bd_width = 1
            font_style = ("微软雅黑", 9, "normal")
        else:
            # Linux/其他平台的安全回退风格
            bg_color = "#f0f0f0"
            fg_color = "#000000"
            relief_style = "solid"
            bd_width = 1
            font_style = ("sans-serif", 10, "normal")

        # 使用动态生成的样式变量替换原有的硬编码
        label = tk.Label(tw, text=self.text, justify='left',
                         background=bg_color, foreground=fg_color, 
                         relief=relief_style, borderwidth=bd_width,
                         wraplength=self.wraplength, # 设置自动换行
                         font=font_style)
                         
        # 稍微增加 ipadx 和 ipady 让文本看起来不那么拥挤，提升视觉效果
        label.pack(ipadx=8, ipady=4)

    def hidetip(self):
        """隐藏提示窗口。"""
        tw = self.tooltip_window
        self.tooltip_window = None
        if tw:
            tw.destroy()
