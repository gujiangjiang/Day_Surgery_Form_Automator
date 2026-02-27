# -*- coding: utf-8 -*-
"""
GUI构建模块。
负责创建和布局所有Tkinter界面元素。
"""
import tkinter as tk
from tkinter import ttk, scrolledtext
import sys # 引入 sys 用于操作系统判定
from .tooltip import Tooltip
from ..config import UI_CONFIG
# --- 新增：导入新的UI组件模块 ---
from . import ui_components

def create_ui(app, handlers):
    """
    创建并布局所有界面组件。
    :param app: AppController的实例，用于访问状态和变量。
    :param handlers: Handlers的实例，用于绑定UI事件。
    """
    ui_texts = UI_CONFIG['texts']
    app_info = UI_CONFIG['app_info']

    style = ttk.Style(app.root)
    
    # 【修复跨平台 Bug】：macOS 下强制使用 clam主题会破坏原生 aqua 引擎，导致透明失效并出现大面积黑屏
    # 因此，仅在非 macOS 系统（如 Windows/Linux）上才使用 clam 主题。
    if sys.platform != "darwin" and "clam" in style.theme_names():
        style.theme_use("clam")

    # --- 新增：为ttk.Entry定义禁用时的背景色 ---
    # macOS 的原生 aqua 主题能够自动处理禁用状态，在此处强制覆盖会导致输入框背景变黑。
    if sys.platform != "darwin":
        disabled_bg = UI_CONFIG['colors']['disabled_bg']
        # 'map' 允许我们根据控件的状态来定义其外观
        # fieldbackground 是输入框内部的背景
        style.map('TEntry', fieldbackground=[('disabled', disabled_bg)])
    
    # 【修复跨平台 Bug】：更安全的背景色提取策略，防止 macOS 找不到 ttk 默认背景时反馈黑色
    if sys.platform == "darwin":
        # 在 Mac 下直接读取系统根窗口的背景色（通常是 systemWindowBackgroundColor）
        default_bg = app.root.cget('bg')
    else:
        default_bg = style.lookup('TFrame', 'background')
        
    # 如果系统出现异常无法拿到背景色，或者返回的是透明占位符，统一使用安全的系统灰底色
    if not default_bg or default_bg == "systemTransparent":
        default_bg = "#ECECEC"
        
    app.root.configure(bg=default_bg)

    log_colors = UI_CONFIG['colors']['log_tags']
    app.log_text_tags = log_colors
    # --- 新增：创建一个列表，用于存放所有在处理期间需要被禁用的控件 ---
    app.interactive_widgets = []

    # --- 底部和顶部UI元素 ---
    bottom_frame = tk.Frame(app.root, bg=default_bg)
    bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=2)
    
    # --- 修改：从配置中读取固定的版本日期 ---
    date_text = f"{ui_texts['date_prefix']}{app_info['build_date']}"
    author_text = f"{ui_texts['author_prefix']}{app_info['author']}"
    tk.Label(bottom_frame, text=date_text, font=app.font_normal, fg="#666666", bg=default_bg).pack(side=tk.LEFT)
    tk.Label(bottom_frame, text=author_text, font=app.font_normal, fg="#666666", bg=default_bg).pack(side=tk.RIGHT)
    
    disclaimer_frame = tk.Frame(app.root, pady=2, bg=default_bg)
    disclaimer_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10)
    tk.Label(disclaimer_frame, text=ui_texts['disclaimer'], font=app.font_disclaimer, fg="red", bg=default_bg).pack()

    # --- 顶部标题 ---
    top_title_frame = tk.Frame(app.root, bg=default_bg)
    top_title_frame.pack(side=tk.TOP, fill=tk.X, pady=(10, 5), padx=10)

    # --- 新增：关于按钮 ---
    # 定义一个简洁的按钮样式
    style.configure("About.TButton", padding=0, relief="flat", background=default_bg)
    style.map("About.TButton", background=[('active', '#e5f1fb')])
    
    # --- 修改：命令绑定到handlers实例 ---
    about_button = ttk.Button(
        top_title_frame, text="?", command=handlers.show_about_dialog, style="About.TButton", width=2
    )
    about_button.pack(side=tk.RIGHT, anchor='n', pady=(5,0))
    Tooltip(about_button, "关于本程序")
    app.interactive_widgets.append(about_button)

    # --- 标题容器，使其在剩余空间内居中 ---
    title_container = tk.Frame(top_title_frame, bg=default_bg)
    title_container.pack(side=tk.TOP, fill=tk.X, expand=True)

    # 从配置中读取标题文本
    tk.Label(title_container, text=ui_texts['subtitle'], font=app.font_subtitle, fg="#0066cc", bg=default_bg).pack()
    tk.Label(title_container, text=ui_texts['main_title'], font=app.font_title, bg=default_bg).pack()

    # --- 主内容区 ---
    main_pane = ttk.PanedWindow(app.root, orient=tk.HORIZONTAL)
    main_pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 5))

    left_panel = ttk.Frame(main_pane, padding=(5, 5, 5, 5))
    main_pane.add(left_panel)

    right_panel = ttk.Frame(main_pane, padding=(5, 5, 5, 5))
    main_pane.add(right_panel)
    
    pane_config = UI_CONFIG['layout']['paned_window']
    s = app.scaling_factor
    sash_default, sash_min, sash_max = int(pane_config['sash_default'] * s), int(pane_config['sash_min'] * s), int(pane_config['sash_max'] * s)

    def set_sash(event): main_pane.sashpos(0, sash_default); main_pane.unbind("<Configure>")
    def limit_sash(event):
        if event.x < sash_min: main_pane.sashpos(0, sash_min); return "break"
        if event.x > sash_max: main_pane.sashpos(0, sash_max); return "break"
    main_pane.bind("<Configure>", set_sash)
    main_pane.bind("<B1-Motion>", limit_sash)
    
    # --- 左侧面板 ---
    file_frame = ttk.LabelFrame(left_panel, text="步骤1: 选择文件和路径", padding=5)
    file_frame.pack(fill=tk.BOTH, expand=True) 
    
    # --- 修改：命令绑定到handlers实例 ---
    discharge_menu_items = [
        ("选择文件", handlers.select_excel_file),
        ("查看模板", lambda: handlers.open_template('discharge')),
        ("---", None),
        ("清空选择", handlers.clear_excel_selection)
    ]
    # --- 修改：使用新的ui_components模块创建控件 ---
    discharge_entry, _, discharge_menubutton = ui_components.create_dropdown_selector(file_frame, "出院患者列表:", app.excel_display_var, discharge_menu_items, app.font_normal)
    app.interactive_widgets.extend([discharge_entry, discharge_menubutton])

    # --- 手术查询文件 ---
    surgery_frame = ttk.LabelFrame(file_frame, text="手术查询文件 (可选, 用于补充床号)", padding=5)
    surgery_frame.pack(fill=tk.X, expand=True, pady=3)
    # 【修复跨平台 Bug】：设置 highlightthickness=0 防止在 macOS 深色模式下点击时出现丑陋的白色焦点边框
    app.surgery_listbox = tk.Listbox(surgery_frame, height=5, font=app.font_normal, highlightthickness=0)
    app.surgery_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
    app.interactive_widgets.append(app.surgery_listbox) # 添加列表框到禁用列表
    
    surgery_buttons_frame = ttk.Frame(surgery_frame)
    surgery_buttons_frame.pack(side=tk.LEFT, fill=tk.Y, anchor='n')
    
    app.surgery_menubutton = ttk.Menubutton(surgery_buttons_frame, text="选项...", width=8)
    surgery_menu = tk.Menu(app.surgery_menubutton, tearoff=False)
    # --- 修改：命令绑定到handlers实例 ---
    surgery_menu.add_command(label="添加文件", command=handlers.select_surgery_query_files)
    surgery_menu.add_command(label="查看模板", command=lambda: handlers.open_template('surgery'))
    surgery_menu.add_separator()
    surgery_menu.add_command(label="清空列表", command=handlers.clear_surgery_query_files)
    app.surgery_menubutton.config(menu=surgery_menu)
    app.surgery_menubutton.pack(fill=tk.X, pady=1)
    app.interactive_widgets.append(app.surgery_menubutton) # 添加按钮到禁用列表
    Tooltip(app.surgery_listbox, "（可选）添加一个或多个手术记录文件，用于自动匹配和补充主列表中缺失的床号信息。")
    Tooltip(app.surgery_menubutton, "管理手术查询文件列表。")

    word_menu_items = [
        ("选择文件", handlers.select_template_file),
        ("查看模板", lambda: handlers.open_template('follow_up')),
        ("使用内置模板", handlers.use_builtin_word_template),
        ("---", None),
        ("清空选择", handlers.clear_template_selection)
    ]
    word_entry, _, word_menubutton = ui_components.create_dropdown_selector(file_frame, "Word模板:", app.template_display_var, word_menu_items, app.font_normal)
    app.interactive_widgets.extend([word_entry, word_menubutton])

    output_dir_items = [
        ("选择文件夹", handlers.select_output_dir),
        ("---", None),
        ("清空选择", handlers.clear_output_dir_selection)
    ]
    output_dir_entry, _, output_dir_menubutton = ui_components.create_dropdown_selector(file_frame, "输出文件夹:", app.output_dir_display_var, output_dir_items, app.font_normal)
    app.interactive_widgets.extend([output_dir_entry, output_dir_menubutton])

    
    # --- 控制和进度条 ---
    control_frame = ttk.LabelFrame(left_panel, text="步骤2: 开始生成", padding=10)
    control_frame.pack(fill=tk.X, pady=(5,0))
    
    style.configure("Accent.TButton", foreground="white", background="#0078D7", font=app.font_button)
    style.configure("Stop.TButton", foreground="white", background="#E81123", font=app.font_button)
    # --- 修改：命令绑定到handlers实例 ---
    app.start_button = ttk.Button(control_frame, text="开始生成", command=handlers.toggle_generation, style="Accent.TButton")
    app.start_button.pack(pady=5, ipady=5, ipadx=20)
    app.interactive_widgets.append(app.start_button)
    Tooltip(app.start_button, "点击开始处理数据并生成Word文档。\n处理过程中，此按钮会变为“停止生成”。")


    # --- 右侧面板 ---
    progress_frame = ttk.LabelFrame(right_panel, text="处理进度与日志", padding=10)
    progress_frame.pack(fill=tk.BOTH, expand=True)
    
    app.progress_bar = ttk.Progressbar(progress_frame, orient='horizontal', mode='determinate')
    app.progress_bar.pack(fill=tk.X, pady=(0, 5))
    
    # 【修复跨平台 Bug】：设置 highlightthickness=0 防止在 macOS 深色模式下点击日志框时出现丑陋的白色焦点边框
    app.log_text = scrolledtext.ScrolledText(progress_frame, height=5, state='disabled', font=app.font_normal, wrap=tk.WORD, highlightthickness=0)
    app.log_text.pack(fill=tk.BOTH, expand=True)
    # 从app实例上获取已经配置好的颜色标签
    for tag, config in app.log_text_tags.items():
        app.log_text.tag_config(tag, **config)

    # --- 新增：为日志区域添加右键菜单 ---
    log_context_menu = tk.Menu(app.root, tearoff=False)
    # --- 修改：命令绑定到handlers实例 ---
    log_context_menu.add_command(label="清空日志", command=handlers.clear_log)
    log_context_menu.add_separator()
    log_context_menu.add_command(label="导出日志...", command=handlers.export_log)

    def show_log_context_menu(event):
        # 【修复跨平台 Bug】：使用 tk_popup 代替 post，防止菜单在 macOS/Linux 上弹出后无法自动收回
        try:
            log_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            log_context_menu.grab_release()

    # 【修复跨平台 Bug】：全面兼容各种系统下不同的右键单击事件绑定
    if sys.platform == "darwin":
        # macOS 上的右键可能是 Button-2, Button-3, 或者按住 Control 点击左键 (Control-Button-1)
        app.log_text.bind("<Button-2>", show_log_context_menu)
        app.log_text.bind("<Button-3>", show_log_context_menu)
        app.log_text.bind("<Control-Button-1>", show_log_context_menu)
    else:
        # Windows / Linux 标准右键
        app.log_text.bind("<Button-3>", show_log_context_menu)