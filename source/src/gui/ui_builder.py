# -*- coding: utf-8 -*-
"""
GUI构建模块。
负责创建和布局所有Tkinter界面元素。
"""
import tkinter as tk
from tkinter import ttk, scrolledtext
from datetime import datetime
from .tooltip import Tooltip
from ..config import CONFIG # 导入全局配置

def create_ui(app):
    """
    创建并布局所有界面组件。
    :param app: MainApp的实例，用于绑定命令和访问变量。
    """
    # 从配置中获取UI配置
    ui_config = CONFIG['UI_CONFIG']
    ui_texts = ui_config['texts']
    app_info = ui_config['app_info']

    style = ttk.Style(app.root)
    if "clam" in style.theme_names():
        style.theme_use("clam")
    default_bg = style.lookup('TFrame', 'background')
    app.root.configure(bg=default_bg)

    # 从配置中获取日志颜色
    log_colors = ui_config['colors']['log_tags']
    app.log_text_tags = log_colors

    # --- 底部和顶部UI元素 ---
    bottom_frame = tk.Frame(app.root, bg=default_bg)
    bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=2)
    
    # 从配置中读取底部标签文本
    date_text = f"{ui_texts['date_prefix']}{datetime.now().strftime('%Y年%m月%d日')}"
    author_text = f"{ui_texts['author_prefix']}{app_info['author']}"
    tk.Label(bottom_frame, text=date_text, font=app.font_normal, fg="#666666", bg=default_bg).pack(side=tk.LEFT)
    tk.Label(bottom_frame, text=author_text, font=app.font_normal, fg="#666666", bg=default_bg).pack(side=tk.RIGHT)
    
    disclaimer_frame = tk.Frame(app.root, pady=2, bg=default_bg)
    disclaimer_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10)
    tk.Label(disclaimer_frame, text=ui_texts['disclaimer'], font=app.font_disclaimer, fg="red", bg=default_bg).pack()

    # --- 顶部标题 ---
    top_title_frame = tk.Frame(app.root, bg=default_bg)
    top_title_frame.pack(side=tk.TOP, fill=tk.X, pady=(10, 5))
    # 从配置中读取标题文本
    tk.Label(top_title_frame, text=ui_texts['subtitle'], font=app.font_subtitle, fg="#0066cc", bg=default_bg).pack()
    tk.Label(top_title_frame, text=ui_texts['main_title'], font=app.font_title, bg=default_bg).pack()

    # --- 主内容区 (左右分割) ---
    main_pane = ttk.PanedWindow(app.root, orient=tk.HORIZONTAL)
    main_pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 5))

    left_panel = ttk.Frame(main_pane, padding=(5, 5, 5, 5))
    main_pane.add(left_panel)

    right_panel = ttk.Frame(main_pane, padding=(5, 5, 5, 5))
    main_pane.add(right_panel)
    
    # --- 分割线位置控制 (已恢复) ---
    s = app.scaling_factor
    sash_default = int(410 * s)
    sash_min = int(320 * s)
    sash_max = int(520 * s)
    def set_initial_sash(event): main_pane.sashpos(0, sash_default); main_pane.unbind("<Configure>")
    def limit_sash_movement(event):
        if event.x < sash_min: main_pane.sashpos(0, sash_min); return "break"
        if event.x > sash_max: main_pane.sashpos(0, sash_max); return "break"
    main_pane.bind("<Configure>", set_initial_sash)
    main_pane.bind("<B1-Motion>", limit_sash_movement)
    # ---------------------------------
    
    # --- 左侧面板 ---
    file_frame = ttk.LabelFrame(left_panel, text="步骤1: 选择文件和路径", padding=5)
    file_frame.pack(fill=tk.BOTH, expand=True) 
    
    discharge_menu_items = [
        ("选择文件", app.select_excel_file),
        ("查看模板", lambda: app.open_template('discharge')),
        ("---", None),
        ("清空选择", app.clear_excel_selection)
    ]
    discharge_entry, discharge_label, _ = _create_dropdown_selector(file_frame, "出院患者列表:", app.excel_display_var, discharge_menu_items, app.font_normal)
    Tooltip(discharge_entry, "必须项。选择包含所有患者出院信息的主Excel文件。")
    Tooltip(discharge_label, "必须项。选择包含所有患者出院信息的主Excel文件。")


    # --- 手术查询文件 ---
    surgery_frame = ttk.LabelFrame(file_frame, text="手术查询文件 (可选, 用于补充床号)", padding=5)
    surgery_frame.pack(fill=tk.X, expand=True, pady=3)
    app.surgery_listbox = tk.Listbox(surgery_frame, height=5, font=app.font_normal)
    app.surgery_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
    surgery_buttons_frame = ttk.Frame(surgery_frame)
    surgery_buttons_frame.pack(side=tk.LEFT, fill=tk.Y, anchor='n')
    
    # 手术查询的下拉菜单按钮
    surgery_menubutton = ttk.Menubutton(surgery_buttons_frame, text="选项...", width=8)
    surgery_menu = tk.Menu(surgery_menubutton, tearoff=False)
    surgery_menu.add_command(label="添加文件", command=app.select_surgery_query_files)
    surgery_menu.add_command(label="查看模板", command=lambda: app.open_template('surgery'))
    surgery_menu.add_separator()
    surgery_menu.add_command(label="清空列表", command=app.clear_surgery_query_files)
    surgery_menubutton.config(menu=surgery_menu)
    surgery_menubutton.pack(fill=tk.X, pady=1)
    # 为手术查询部分添加提示
    Tooltip(app.surgery_listbox, "（可选）添加一个或多个手术记录文件，用于自动匹配和补充主列表中缺失的床号信息。")
    Tooltip(surgery_menubutton, "管理手术查询文件列表。")

    
    word_menu_items = [
        ("选择文件", app.select_template_file),
        ("查看模板", lambda: app.open_template('follow_up')),
        ("使用内置模板", app.use_builtin_word_template),
        ("---", None),
        ("清空选择", app.clear_template_selection)
    ]
    word_entry, word_label, _ = _create_dropdown_selector(file_frame, "Word模板:", app.template_display_var, word_menu_items, app.font_normal)
    Tooltip(word_entry, "必须项。选择用于生成随访表的Word模板文件，或选择使用软件内置的默认模板。")
    Tooltip(word_label, "必须项。选择用于生成随访表的Word模板文件，或选择使用软件内置的默认模板。")

    
    output_dir_items = [
        ("选择文件夹", app.select_output_dir),
        ("---", None),
        ("清空选择", app.clear_output_dir_selection)
    ]
    output_entry, output_label, _ = _create_dropdown_selector(file_frame, "输出文件夹:", app.output_dir_display_var, output_dir_items, app.font_normal)
    Tooltip(output_entry, "必须项。选择一个文件夹用于保存所有生成的Word文档。")
    Tooltip(output_label, "必须项。选择一个文件夹用于保存所有生成的Word文档。")

    
    # --- 控制和进度条 ---
    left_bottom_container = ttk.Frame(left_panel)
    left_bottom_container.pack(fill=tk.X, pady=(5,0))
    control_frame = ttk.LabelFrame(left_bottom_container, text="步骤2: 开始生成", padding=10)
    control_frame.pack(fill=tk.X)
    
    style.configure("Accent.TButton", foreground="white", background="#0078D7", font=app.font_button)
    style.configure("Stop.TButton", foreground="white", background="#E81123", font=app.font_button)
    app.start_button = ttk.Button(control_frame, text="开始生成", command=app.toggle_generation, style="Accent.TButton")
    app.start_button.pack(pady=5, ipady=5, ipadx=20)
    Tooltip(app.start_button, "点击开始处理数据并生成Word文档。\n处理过程中，此按钮会变为“停止生成”。")


    # --- 右侧面板 ---
    progress_frame = ttk.LabelFrame(right_panel, text="处理进度与日志", padding=10)
    progress_frame.pack(fill=tk.BOTH, expand=True)
    
    app.progress_bar = ttk.Progressbar(progress_frame, orient='horizontal', mode='determinate')
    app.progress_bar.pack(fill=tk.X, pady=(0, 5))
    
    app.log_text = scrolledtext.ScrolledText(progress_frame, height=5, state='disabled', font=app.font_normal, wrap=tk.WORD)
    app.log_text.pack(fill=tk.BOTH, expand=True)
    # 从app实例上获取已经配置好的颜色标签
    for tag, config in app.log_text_tags.items():
        app.log_text.tag_config(tag, **config)

    # --- 新增：为日志区域添加右键菜单 ---
    log_context_menu = tk.Menu(app.root, tearoff=False)
    log_context_menu.add_command(label="清空日志", command=app.clear_log)
    log_context_menu.add_separator()
    log_context_menu.add_command(label="导出日志...", command=app.export_log)

    def show_log_context_menu(event):
        log_context_menu.post(event.x_root, event.y_root)

    app.log_text.bind("<Button-3>", show_log_context_menu)
    # ------------------------------------

def _create_dropdown_selector(parent, label_text, string_var, menu_items, font):
    """
    创建带有下拉菜单按钮的文件/目录选择器行。
    使用grid布局管理器来确保输入框可以正确缩放。
    返回创建的 (Entry, Label, Menubutton) 控件元组，以便可以为其绑定Tooltip。
    """
    row_frame = ttk.Frame(parent)
    row_frame.pack(fill=tk.X, expand=True, pady=1)

    # --- 布局修复 ---
    # 配置grid的列权重，让第1列（输入框）可以伸缩
    row_frame.columnconfigure(1, weight=1)

    label_widget = ttk.Label(row_frame, text=label_text, width=12, font=font)
    label_widget.grid(row=0, column=0, sticky='w', padx=(0, 5))
    
    entry = ttk.Entry(row_frame, textvariable=string_var, state='readonly', font=font)
    # 使用 sticky='ew' 让输入框水平填充其单元格
    entry.grid(row=0, column=1, sticky='ew')
    
    menubutton = ttk.Menubutton(row_frame, text="选项...", width=8)
    menubutton.grid(row=0, column=2, sticky='e', padx=(5, 0))
    # -----------------
    
    menu = tk.Menu(menubutton, tearoff=False)
    # --- 错误修复 ---
    # 在循环中使用不同的变量名(item_text)，以避免覆盖外部的 'label' 控件变量
    for item_text, command in menu_items:
        if item_text == "---":
            menu.add_separator()
        else:
            menu.add_command(label=item_text, command=command)
    menubutton.config(menu=menu)
    
    # 为 "选项..." 按钮添加提示
    Tooltip(menubutton, "点击展开操作菜单")
    
    # 返回所有创建的控件，以便外部可以为它们分别添加提示
    return entry, label_widget, menubutton
