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

# 将项目根目录添加到sys.path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.gui.app_controller import AppController
from src.logger_setup import setup_logging


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
        
        # 【修复跨平台 Bug】：针对不同平台设置透明属性，解决 macOS 和 Linux 报错问题
        try:
            if sys.platform == "win32":
                # Windows 环境：支持 -transparentcolor 属性
                self.wm_attributes('-transparentcolor', transparent_color)
            elif sys.platform == "darwin":
                # macOS 环境：使用 -transparent 属性，并使用系统透明色
                self.wm_attributes('-transparent', True)
                self.config(bg='systemTransparent')
            else:
                # Linux/X11 环境：设置为 splash 类型
                self.wm_attributes('-type', 'splash')
        except tk.TclError:
            # 忽略当前系统不支持的窗口属性，防止程序因此崩溃
            pass

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
    # 在程序最开始初始化日志系统
    setup_logging(BASE_PATH)

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
    icon_path_ico = BASE_PATH / 'assets' / 'app.ico'
    icon_path_png = BASE_PATH / 'assets' / 'app.png' # 新增：macOS/Linux 专用图标路径
    splash_path = BASE_PATH / 'assets' / 'splash.png'

    root = tk.Tk()
    root.withdraw()

    # 检查资源文件是否存在
    if not splash_path.exists():
        messagebox.showwarning("资源缺失", f"启动画面文件 'splash.png' 未找到！\n请确保它位于 'assets' 文件夹中。")
        splash = None # 确保splash变量存在
    else:
        splash = SplashScreen(root, splash_path)

    # 图标加载逻辑优化：分别检查对应平台的图标文件是否存在
    if sys.platform == "win32" and not icon_path_ico.exists():
         print(f"警告：Windows 图标文件 'app.ico' 未找到。")
    elif sys.platform != "win32" and not icon_path_png.exists():
         print(f"警告：macOS/Linux 图标文件 'app.png' 未找到。")
    else:
        try:
            # 【修复跨平台 Bug】：处理 macOS/Linux 环境下图加载 .ico 可能出错的问题
            if sys.platform == "win32":
                # iconbitmap 在某些系统下可能需要字符串路径 (Windows 首选)
                root.iconbitmap(str(icon_path_ico))
            else:
                # macOS 和 Linux 环境下，改用 iconphoto 加载 app.png 作为窗口图标
                if icon_path_png.exists():
                    img = tk.PhotoImage(file=str(icon_path_png))
                    root.iconphoto(True, img)
        except Exception as e:
            print(f"警告：无法加载图标文件。{e}")

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

    # 在指定的延迟后执行 `show_main_window` 函数
    root.after(delay, show_main_window)
    
    root.mainloop()

if __name__ == "__main__":
    main()
