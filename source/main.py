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
from pathlib import Path # 导入Path类

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

# 从src目录导入App主类
# 假设main.py在项目根目录，而其他代码在src/下
# 为了让这个导入生效，需要将src目录的父目录（即项目根目录）加入sys.path
# get_base_path() 函数会返回这个根目录
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
# --- 修改：更新导入的类名 ---
from src.gui.app_controller import AppController


def get_base_path():
    """
    获取应用的根目录，兼容源码运行和PyInstaller打包后的情况。
    返回一个Path对象。
    """
    if getattr(sys, 'frozen', False):
        # 如果程序被打包，base_path是可执行文件所在的目录
        return Path(sys.executable).parent
    else:
        # 如果从源码运行，base_path是main.py所在的目录
        return Path(__file__).resolve().parent

# 在程序启动时就确定好根目录
BASE_PATH = get_base_path()

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
        # 适配高DPI屏幕 (Modernized approach)
        # Only perform this on Windows
        if sys.platform == "win32":
            # 优先尝试最新的API (适用于Windows 10 v1703+)，提供最佳的跨显示器缩放效果
            try:
                # Per-Monitor V2 DPI Awareness. Requires Windows 10 Creators Update (1703)+
                ctypes.windll.user32.SetProcessDpiAwarenessContext(-4)
            except (AttributeError, OSError):
                # 如果最新API不可用，则回退到旧版API (适用于Windows 8.1+)
                try:
                    # Per-Monitor DPI Awareness. Requires Windows 8.1+
                    ctypes.windll.shcore.SetProcessDpiAwareness(2)
                except (AttributeError, OSError):
                    # 如果shcore也不可用，则使用最旧的API (适用于Windows Vista+)
                    try:
                        # System DPI Awareness. Requires Windows Vista+
                        ctypes.windll.user32.SetProcessDPIAware()
                    except (AttributeError, OSError):
                        pass # Gracefully fail on very old Windows versions
    except Exception as e:
        print(f"设置DPI感知失败: {e}")

    # 使用 pathlib 构建资源路径
    icon_path = BASE_PATH / 'assets' / 'app.ico'
    splash_path = BASE_PATH / 'assets' / 'splash.png'

    root = tk.Tk()
    root.withdraw()

    # 检查资源文件是否存在
    if not splash_path.exists():
        messagebox.showwarning("资源缺失", f"启动画面文件 'splash.png' 未找到！\n请确保它位于 'assets' 文件夹中。")
        splash = None # 确保splash变量存在
    else:
        splash = SplashScreen(root, splash_path)

    if not icon_path.exists():
         print(f"警告：图标文件 'app.ico' 未找到。")
    else:
        try:
            # iconbitmap 在某些系统下可能需要字符串路径
            root.iconbitmap(str(icon_path))
        except tk.TclError:
            print(f"警告：无法加载图标文件。")

    # --- 修改：实例化新的类名 ---
    # 将根目录路径(Path对象)传递给App实例，以便其他模块也能正确找到文件
    app_instance = AppController(root, base_path=BASE_PATH)

    def show_main_window():
        if splash:
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
