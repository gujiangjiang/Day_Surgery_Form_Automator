# -*- coding: utf-8 -*-
"""
日间手术随访表生成系统 (重构版) - 主程序入口
"""

import os
import sys
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

def main():
    """
    主函数，负责初始化和运行应用。
    """
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

    # 如果使用Nuitka打包并包含启动画面
    if splash_active:
        nuitka_splashscreen_python.mark_as_deployed(
            image_path=splash_path
        )

    root = tk.Tk()
    try:
        root.iconbitmap(icon_path)
    except tk.TclError:
        # 打印完整的、程序正在尝试访问的路径，方便调试
        print(f"警告：未找到图标文件。尝试的路径为: {icon_path}")

    # 在 App 初始化时隐藏主窗口，防止闪烁
    root.withdraw()

    app_instance = App(root)

    def finalize_startup():
        """在所有组件加载完毕后，关闭启动画面并显示主窗口"""
        if splash_active:
            nuitka_splashscreen_python.close()
        root.deiconify()
        root.focus_force()

    # 延迟执行，给GUI一点时间来渲染
    root.after(100, finalize_startup)
    root.mainloop()

if __name__ == "__main__":
    main()
