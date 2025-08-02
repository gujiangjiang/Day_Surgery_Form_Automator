# -*- coding: utf-8 -*-
"""
GUI构建模块。
负责创建和布局所有Tkinter界面元素。
"""
import tkinter as tk
from tkinter import ttk, scrolledtext
from datetime import datetime

def create_ui(app):
    """
    创建并布局所有界面组件。
    :param app: MainApp的实例，用于绑定命令和访问变量。
    """
    style = ttk.Style(app.root)
    if "clam" in style.theme_names():
        style.theme_use("clam")
    default_bg = style.lookup('TFrame', 'background')
    app.root.configure(bg=default_bg)
    app.log_text_tags = {"warning": {"foreground": "orange"}, "error": {"foreground": "red"}}

    # --- 底部信息栏 ---
    bottom_frame = tk.Frame(app.root, bg=default_bg)
    bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=2)
    tk.Label(bottom_frame, text=f"版本日期：{datetime.now().strftime('%Y年%m月%d日')}", font=app.font_normal, fg="#666666", bg=default_bg).pack(side=tk.LEFT)
    tk.Label(bottom_frame, text="作者：顾江江", font=app.font_normal, fg="#666666", bg=default_bg).pack(side=tk.RIGHT)
    
    disclaimer_frame = tk.Frame(app.root, pady=2, bg=default_bg)
    disclaimer_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10)
    tk.Label(disclaimer_frame, text="本工具仅供骨科内部测试，请勿外传", font=app.font_disclaimer, fg="red", bg=default_bg).pack()

    # --- 顶部标题 ---
    top_title_frame = tk.Frame(app.root, bg=default_bg)
    top_title_frame.pack(side=tk.TOP, fill=tk.X, pady=(10, 5))
    tk.Label(top_title_frame, text="丹阳市人民医院", font=app.font_subtitle, fg="#0066cc", bg=default_bg).pack()
    tk.Label(top_title_frame, text="日间手术随访表生成系统", font=app.font_title, bg=default_bg).pack()

    # --- 主内容区 (左右分割) ---
    main_pane = ttk.PanedWindow(app.root, orient=tk.HORIZONTAL)
    main_pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 5))

    left_panel = ttk.Frame(main_pane, padding=(5, 5, 5, 5))
    main_pane.add(left_panel)

    right_panel = ttk.Frame(main_pane, padding=(5, 5, 5, 5))
    main_pane.add(right_panel)

    s = app.scaling_factor
    sash_default = int(420 * s)
    def set_initial_sash(event): main_pane.sashpos(0, sash_default)
    main_pane.bind("<Configure>", set_initial_sash)
    
    # --- 左侧面板内容 ---
    file_frame = ttk.LabelFrame(left_panel, text="步骤1: 选择文件和路径", padding=5)
    file_frame.pack(fill=tk.BOTH, expand=True) 
    
    _create_file_selector(file_frame, "出院患者列表:", app.excel_path_var, app.select_excel_file, lambda: app.open_template('discharge'), app.font_normal)
    
    surgery_frame = ttk.LabelFrame(file_frame, text="手术查询文件 (可选, 用于补充床号)", padding=5)
    surgery_frame.pack(fill=tk.X, expand=True, pady=3)
    
    listbox_frame = ttk.Frame(surgery_frame)
    listbox_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    app.surgery_listbox = tk.Listbox(listbox_frame, height=4, font=app.font_normal)
    app.surgery_listbox.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
    
    style.configure("Link.TButton", foreground="blue", relief="flat", borderwidth=0, padding=0)
    template_btn_surgery = ttk.Button(listbox_frame, text="查看模板", command=lambda: app.open_template('surgery'), style="Link.TButton")
    template_btn_surgery.pack(side=tk.BOTTOM, anchor='w', pady=(2,0))
    
    surgery_buttons_frame = ttk.Frame(surgery_frame)
    surgery_buttons_frame.pack(side=tk.LEFT, fill=tk.Y, anchor='n', padx=(5,0))
    ttk.Button(surgery_buttons_frame, text="添加文件", command=app.select_surgery_query_files, width=8).pack(fill=tk.X, pady=1)
    ttk.Button(surgery_buttons_frame, text="清空列表", command=app.clear_surgery_query_files, width=8).pack(fill=tk.X, pady=1)
    
    _create_file_selector(file_frame, "Word模板:", app.template_path_var, app.select_template_file, lambda: app.open_template('follow_up'), app.font_normal)
    _create_file_selector(file_frame, "输出文件夹:", app.output_dir_var, app.select_output_dir, None, app.font_normal)
    
    left_bottom_container = ttk.Frame(left_panel)
    left_bottom_container.pack(fill=tk.X, pady=(5,0))
    control_frame = ttk.LabelFrame(left_bottom_container, text="步骤2: 开始生成", padding=10)
    control_frame.pack(fill=tk.X)
    
    style.configure("Accent.TButton", foreground="white", background="#0078D7", font=app.font_button)
    style.configure("Stop.TButton", foreground="white", background="#E81123", font=app.font_button)
    app.start_button = ttk.Button(control_frame, text="开始生成", command=app.toggle_generation, style="Accent.TButton")
    app.start_button.pack(pady=5, ipady=5, ipadx=20)

    # --- 右侧面板内容 ---
    progress_frame = ttk.LabelFrame(right_panel, text="处理进度与日志", padding=10)
    progress_frame.pack(fill=tk.BOTH, expand=True)
    
    app.progress_bar = ttk.Progressbar(progress_frame, orient='horizontal', mode='determinate')
    app.progress_bar.pack(fill=tk.X, pady=(0, 5))
    
    app.log_text = scrolledtext.ScrolledText(progress_frame, height=5, state='disabled', font=app.font_normal, wrap=tk.WORD)
    app.log_text.pack(fill=tk.BOTH, expand=True)
    for tag, config in app.log_text_tags.items():
        app.log_text.tag_config(tag, **config)

def _create_file_selector(parent, label_text, string_var, browse_command, template_command, font):
    """
    一个私有辅助函数，用于创建文件/目录选择器行。
    此版本接收6个参数。
    """
    container = ttk.Frame(parent)
    container.pack(fill=tk.X, expand=True, pady=1)
    
    top_frame = ttk.Frame(container)
    top_frame.pack(fill=tk.X, expand=True)
    
    ttk.Label(top_frame, text=label_text, width=12, font=font).pack(side=tk.LEFT)
    ttk.Entry(top_frame, textvariable=string_var, state='readonly', font=font).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    ttk.Button(top_frame, text="浏览...", command=browse_command, width=8).pack(side=tk.RIGHT)
    
    if template_command:
        bottom_frame = ttk.Frame(container)
        bottom_frame.pack(fill=tk.X, expand=True)
        ttk.Label(bottom_frame, text="", width=12).pack(side=tk.LEFT) # 占位符标签，用于对齐
        template_btn = ttk.Button(bottom_frame, text="查看模板", command=template_command, style="Link.TButton")
        template_btn.pack(side=tk.LEFT, anchor='w', padx=5)
