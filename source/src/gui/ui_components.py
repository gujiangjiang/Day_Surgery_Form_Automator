# -*- coding: utf-8 -*-
"""
可复用的UI组件模块。
此模块包含用于创建复杂或标准化UI控件的工厂函数。
"""
import tkinter as tk
from tkinter import ttk
from .tooltip import Tooltip

def create_dropdown_selector(parent, label_text, string_var, menu_items, font):
    """
    创建带有下拉菜单按钮的文件/目录选择器行。
    :return: (Entry, Label, Menubutton) 控件元组，以便可以为其绑定Tooltip。
    """
    row_frame = ttk.Frame(parent)
    row_frame.pack(fill=tk.X, expand=True, pady=1)

    row_frame.columnconfigure(1, weight=1)

    label_widget = ttk.Label(row_frame, text=label_text, width=12, font=font)
    label_widget.grid(row=0, column=0, sticky='w', padx=(0, 5))
    
    entry = ttk.Entry(row_frame, textvariable=string_var, state='readonly', font=font)
    entry.grid(row=0, column=1, sticky='ew')
    
    menubutton = ttk.Menubutton(row_frame, text="选项...", width=8)
    menubutton.grid(row=0, column=2, sticky='e', padx=(5, 0))
    
    menu = tk.Menu(menubutton, tearoff=False)
    for item_text, command in menu_items:
        if item_text == "---":
            menu.add_separator()
        else:
            menu.add_command(label=item_text, command=command)
    menubutton.config(menu=menu)
    
    Tooltip(entry, f"已选择的{label_text.replace(':', '')}路径。")
    Tooltip(label_widget, f"点击右侧“选项...”按钮来选择{label_text.replace(':', '')}。")
    Tooltip(menubutton, "点击展开操作菜单")
    
    return entry, label_widget, menubutton
