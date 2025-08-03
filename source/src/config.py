# -*- coding: utf-8 -*-
"""
存放所有全局配置信息。
"""

CONFIG = {
    # --- 核心业务逻辑配置 ---
    "day_surgery_max_days": 2,
    "follow_up_days": 7,  # 随访发生于出院后的天数
    "unknown_bed_placeholder": "（手动填写）", # Word内容中的床号未知占位符
    "unknown_bed_filename_suffix": "（床号未知）",   # 文件名中的床号未知后缀

    # --- 内置模板文件名 ---
    "discharge_template_name": "出院患者列表模板.xls",
    "surgery_template_name": "手术查询模板.xlsx",
    "follow_up_template_name": "日间手术随访登记表模板.docx",

    # --- Excel列映射 ---
    "column_mapping": {
        "name": "姓名", "department": "出院科室", "hospital_id": "住院号",
        "discharge_date": "出院日期", "hospital_days": "住院天数", "gender": "性别",
        "age": "年龄", "bed_number": "床号", "admission_date": "入院日期",
        "surgery_date": "手术日期", "surgery_name": "手术名称",
        "diagnosis": "最后诊断1",
        "phone": "联系电话", "doctor": "经治医生"
    },
    "required_patient_cols": [
        "name", "department", "hospital_id", "discharge_date", "hospital_days"
    ],
    "required_surgery_cols": ["hospital_id", "name", "bed_number"],
    
    # --- Word模板占位符 ---
    "template_placeholders": {
        "{{科室}}": "department", "{{姓名}}": "name", "{{性别}}": "gender",
        "{{年龄}}": "age", "{{住院号}}": "hospital_id", "{{床号}}": "bed_number",
        "{{入院日期}}": "admission_date", "{{出院日期}}": "discharge_date",
        "{{手术日期}}": "surgery_date", "{{手术名称}}": "surgery_name",
        "{{出院诊断}}": "diagnosis", "{{联系电话}}": "phone", "{{经治医生}}": "doctor",
    },

    # =========================================================================
    # --- UI界面配置 (所有界面相关的文字、字体、颜色都在这里修改) ---
    # =========================================================================
    "UI_CONFIG": {
        "app_info": {
            "title": "日间手术随访表生成系统",
            "version": "8.3", # 版本号
            "author": "顾江江"  # 作者名
        },

        "fonts": {
            # 格式: (字体名称, 字号, *样式) e.g., ("微软雅黑", 9, "bold")
            "normal": ("微软雅黑", 9),
            "bold": ("微软雅黑", 10, "bold"),
            "title": ("微软雅黑", 20, "bold"),
            "subtitle": ("微软雅黑", 16, "bold"),
            "button": ("微软雅黑", 12, "bold"),
            "disclaimer": ("微软雅黑", 10)
        },

        "texts": {
            "main_title": "日间手术随访表生成系统",
            "subtitle": "丹阳市人民医院",
            "disclaimer": "本工具仅供骨科内部测试，请勿外传",
            "author_prefix": "作者：",
            "date_prefix": "版本日期：",
            "welcome_message": "欢迎使用日间手术随访表生成系统！",
            "welcome_tips": "提示：\n1、点击各项右侧的“选项”按钮来选择文件或查看模板。\n2、如果未选择Word模板，程序将使用内置模板。\n3、处理过程中可随时点击“停止生成”按钮中断任务。",
            "welcome_warning": "警告：本工具仅供骨科内部测试使用，请勿外传，否则后果自负！",
            "log_separator": "----------------"
        },
        
        "colors": {
            "log_tags": {
                "warning": {"foreground": "#FF8C00"}, # 暗橙色
                "error": {"foreground": "red"},
                "info": {"foreground": "#008B8B"} # 深青色
            }
        },

        "layout": {
            "window": {
                "base_width": 685,
                "base_height": 535,
                "resizable": False # 窗口是否可调整大小
            },
            "paned_window": {
                "sash_default": 410, # 左右分割线的默认位置
                "sash_min": 320,     # 分割线可移动的最小位置
                "sash_max": 520      # 分割线可移动的最大位置
            }
        }
    }
}
