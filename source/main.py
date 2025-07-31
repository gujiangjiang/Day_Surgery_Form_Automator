# -*- coding: utf-8 -*-
"""
日间手术随访表生成系统 (重构版) - 主程序入口
"""

import os
import sys
import time
import tkinter as tk
from tkinter import messagebox
import ctypes

# 检查并安装必要的库
try:
    from docx import Document
    import openpyxl
    import xlrd
except ImportError:
    # 在主程序启动前，如果缺少库，则弹出错误提示
    root_err = tk.Tk()
    root_err.withdraw()
    messagebox.showerror(
        "依赖缺失",
        "缺少必要的库 (openpyxl, xlrd, python-docx)。\n"
        "请在命令行运行 'pip install openpyxl xlrd python-docx' 来安装。"
    )
    sys.exit(1)

# 尝试导入Nuitka启动画面模块
try:
    import nuitka_splashscreen_python
    splash_active = True
except ImportError:
    splash_active = False

# 从src目录导入App主类
from src.gui.app import App

def get_asset_path(relative_path):
    """
    获取资源的绝对路径。这可以确保无论从哪里运行脚本，
    都能正确找到 assets 文件夹中的文件。
    """
    # 获取脚本所在的目录
    base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

class SplashScreen(tk.Toplevel):
    """
    纯Tkinter实现的、无边框的启动画面类
    """
    def __init__(self, parent, image_path):
        tk.Toplevel.__init__(self, parent)

        # 定义一种颜色，后续会将其设置为透明
        transparent_color = '#abcdef'

        # 创建一个无边框的顶层窗口
        self.overrideredirect(True)
        # 将窗口背景和透明色绑定
        self.config(bg=transparent_color)
        self.wm_attributes('-transparentcolor', transparent_color)

        # 加载启动图片
        self.image = tk.PhotoImage(file=image_path)
        width = self.image.width()
        height = self.image.height()

        # 使用标签显示图片，背景也设置为透明色
        tk.Label(self, image=self.image, bg=transparent_color).pack()

        # 计算位置使其在屏幕居中
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        pos_x = (screen_width // 2) - (width // 2)
        pos_y = (screen_height // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{pos_x}+{pos_y}")

        # 强制立即绘制窗口
        self.update()

def main():
    """
    主函数，负责初始化和运行应用。
    """
    SPLASH_MIN_DURATION = 1500  # 1.5秒
    start_time = time.time()
    
    try:
        # 适配高DPI屏幕
        if sys.getwindowsversion().major >= 6 and sys.getwindowsversion().minor >= 3:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        else:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception as e:
        print(f"设置DPI感知失败: {e}")

    # 获取资源的绝对路径
    icon_path = get_asset_path("assets/app.ico")
    splash_path = get_asset_path("assets/splash.png")

    root = tk.Tk()
    root.withdraw() # 保持主窗口在初始化时隐藏

    # 1. 创建并显示我们自己的启动画面
    splash = SplashScreen(root, splash_path)

    try:
        root.iconbitmap(icon_path)
    except tk.TclError:
        print(f"警告：未找到图标文件。尝试的路径为: {icon_path}")

    # 2. 初始化主程序GUI
    app_instance = App(root)

    def show_main_window():
        splash.destroy()
        root.deiconify()
        root.lift()
        root.focus_force()

    init_duration_ms = (time.time() - start_time) * 1000
    delay = int(SPLASH_MIN_DURATION - init_duration_ms)
    if delay < 0:
        delay = 0

    # 3. 在指定的延迟后执行 `show_main_window` 函数
    root.after(delay, show_main_window)
    
    root.mainloop()

if __name__ == "__main__":
    main()