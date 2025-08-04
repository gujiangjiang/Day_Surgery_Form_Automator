# -*- coding: utf-8 -*-
"""
存放所有全局配置信息。

此文件分为两部分：
1. UI_CONFIG: 专门存放所有与图形用户界面 (GUI) 相关的配置，如窗口标题、字体、颜色、文本和布局。
2. CONFIG: 存放所有与核心业务逻辑相关的配置，如文件处理规则、数据列映射等。
3. LOGGING_CONFIG: 日志系统配置

将两者分开有助于更好地组织和维护代码。
"""

# =========================================================================
# 1. UI界面配置 (所有界面相关的文字、字体、颜色、布局都在这里修改)
# =========================================================================
UI_CONFIG = {
    # --- 应用程序基本信息 ---
    "app_info": {
        "title": "日间手术随访表生成系统", # 主标题，会显示在窗口顶部
        "version": "8.3",              # 版本号，会显示在窗口标题栏
        "internal_version": "8.3.23.1",# 内部版本号，用于更精细的跟踪
        "author": "顾江江",             # 作者名，会显示在窗口底部
        "build_date": "2025-08-04"     # 新增：版本发布日期，每次发布新版时修改这里
    },

    # --- 字体配置 ---
    # 定义了程序中使用的所有字体样式
    # 格式: (字体名称, 字号, *样式)，样式可以是 "bold", "italic", "underline" 等
    "fonts": {
        "normal": ("微软雅黑", 9),
        "bold": ("微软雅黑", 10, "bold"),
        "title": ("微软雅黑", 20, "bold"),
        "subtitle": ("微软雅黑", 16, "bold"),
        "button": ("微软雅黑", 12, "bold"),
        "disclaimer": ("微软雅黑", 10)
    },

    # --- 界面文本配置 ---
    # 定义了界面上所有固定的文本标签内容
    "texts": {
        "main_title": "日间手术随访表生成系统",
        "subtitle": "丹阳市人民医院",
        "disclaimer": "本工具仅供骨科内部测试，请勿外传",
        "author_prefix": "作者：",
        "date_prefix": "版本日期：",
        "welcome_message": "欢迎使用日间手术随访表生成系统！",
        "welcome_tips": "提示：\n1、点击各项右侧的“选项”按钮来选择文件或查看模板。\n2、如果未选择Word模板，程序将使用内置模板。\n3、处理过程中可随时点击“停止生成”按钮中断任务。",
        "welcome_warning": "警告：本工具仅供骨科内部测试使用，请勿外传，否则后果自负！",
        "log_separator": "----------------",
        # 新增：关于对话框的文本
        "about_title": "关于本程序",
        "about_content": (
            "{title}\n\n"
            "版本: V{version} (内部版本: {internal_version})\n"
            "作者: {author}\n"
            "发布日期: {build_date}\n\n"
            "这是一个用于自动生成日间手术随访表的工具。\n"
            "如有任何问题，请联系作者。"
        ),
    },
    
    # --- 颜色配置 ---
    "colors": {
        # 定义了日志区域不同级别信息的显示颜色
        "log_tags": {
            "WARNING": {"foreground": "#FF8C00"}, # 警告（暗橙色）
            "ERROR": {"foreground": "red"},       # 错误（红色）
            "INFO": {"foreground": "#008B8B"},     # 信息（深青色）
            "CRITICAL": {"foreground": "red", "font": ("微软雅黑", 9, "bold")}
        },
        # 新增：定义了控件在禁用状态下的背景色
        "disabled_bg": "#f0f0f0" # 标准的灰色
    },

    # --- 布局配置 ---
    "layout": {
        # 主窗口布局
        "window": {
            "base_width": 685,      # 窗口基础宽度（会根据屏幕DPI缩放）
            "base_height": 535,     # 窗口基础高度（会根据屏幕DPI缩放）
            "resizable": False      # 窗口是否可调整大小 (True/False)
        },
        # 主界面左右分割窗格的布局
        "paned_window": {
            "sash_default": 410, # 分割线的默认位置
            "sash_min": 320,     # 分割线可向左移动的最小位置
            "sash_max": 520      # 分割线可向右移动的最大位置
        }
    }
}


# =========================================================================
# 2. 核心业务逻辑配置
# =========================================================================
CONFIG = {
    # --- 业务规则 ---
    "day_surgery_max_days": 2,      # 定义日间手术的最大住院天数
    "follow_up_days": 7,            # 定义随访发生于出院后的天数

    # --- 文件与占位符 ---
    # !!! 新增：可自定义的输出文件名格式 !!!
    # 您可以在下面的字符串中使用花括号 {} 来插入动态数据。
    # 可用占位符: {出院年月}, {科室}, {随访日期}, {姓名}, {住院号}
    # 示例: "{出院年月}_{科室}_{姓名}"  -> "2023年08月_骨科_张三.docx"
    "output_filename_format": "{出院年月}_{科室}_日间手术随访_{随访日期}_{姓名}",
    "unknown_bed_placeholder": "（手动填写）", # 当Word文档中床号未知时，使用的占位符
    "unknown_bed_filename_suffix": "（床号未知）",   # 当生成的文件名床号未知时，添加的后缀

    # --- 内置模板文件名 ---
    # 这些文件应存放在项目根目录下的 'templates' 文件夹中
    "discharge_template_name": "出院患者列表模板.xls",
    "surgery_template_name": "手术查询模板.xlsx",
    "follow_up_template_name": "日间手术随访登记表模板.docx",

    # --- Excel列名与内部键的映射 ---
    # 程序内部使用左侧的英文键，对应Excel表格中右侧的中文列名
    "column_mapping": {
        "name": "姓名", "department": "出院科室", "hospital_id": "住院号",
        "discharge_date": "出院日期", "hospital_days": "住院天数", "gender": "性别",
        "age": "年龄", "bed_number": "床号", "admission_date": "入院日期",
        "surgery_date": "手术日期", "surgery_name": "手术名称",
        "diagnosis": "最后诊断1",
        "phone": "联系电话", "doctor": "经治医生"
    },

    # --- 数据校验规则 ---
    # 定义了在不同Excel文件中必须存在的列，用于程序自动查找标题行
    "required_patient_cols": [
        "name", "department", "hospital_id", "discharge_date", "hospital_days"
    ],
    "required_surgery_cols": ["hospital_id", "name", "bed_number"],
    
    # --- Word模板占位符与内部键的映射 ---
    # 定义了Word模板中的占位符（左侧）与程序内部数据键（右侧）的对应关系
    "template_placeholders": {
        "{{科室}}": "department", "{{姓名}}": "name", "{{性别}}": "gender",
        "{{年龄}}": "age", "{{住院号}}": "hospital_id", "{{床号}}": "bed_number",
        "{{入院日期}}": "admission_date", "{{出院日期}}": "discharge_date",
        "{{手术日期}}": "surgery_date", "{{手术名称}}": "surgery_name",
        "{{出院诊断}}": "diagnosis", "{{联系电话}}": "phone", "{{经治医生}}": "doctor",
    }
}

# =========================================================================
# 3. 日志系统配置
# =========================================================================
LOGGING_CONFIG = {
    "enable_file_logging": False, # True 启用本地日志文件生成，False 禁用本地日志文件生成
    "log_filename": "app_runtime.log", # 日志文件名
    "log_level": "INFO", # 日志级别: DEBUG, INFO, WARNING, ERROR, CRITICAL
    "log_max_bytes": 10 * 1024 * 1024, # 单个日志文件最大大小 (10MB)
    "log_backup_count": 5 # 保留的旧日志文件数量
}
