# -*- coding: utf-8 -*-
"""
存放所有全局配置信息。
"""

CONFIG = {
    "app_title": "日间手术随访表生成系统 V8.0",
    "day_surgery_max_days": 2,
    "follow_up_days": 7,  # 随访发生于出院后的天数
    "unknown_bed_placeholder": "（手动填写）", # Word内容中的床号未知占位符
    "unknown_bed_filename_suffix": "（床号未知）",   # 文件名中的床号未知后缀
    "column_mapping": {
        # 内部键: Excel列名
        "name": "姓名", "department": "出院科室", "hospital_id": "住院号",
        "discharge_date": "出院日期", "hospital_days": "住院天数", "gender": "性别",
        "age": "年龄", "bed_number": "床号", "admission_date": "入院日期",
        "surgery_date": "手术日期", "surgery_name": "手术名称",
        "diagnosis": "最后诊断1",
        "phone": "联系电话", "doctor": "经治医生"
    },
    # 这些是必须在“出院患者列表”中找到的列
    "required_patient_cols": [
        "name", "department", "hospital_id", "discharge_date", "hospital_days"
    ],
    # 这些是必须在“手术查询”文件中找到的列
    "required_surgery_cols": ["hospital_id", "name", "bed_number"],
    "template_placeholders": {
        "{{科室}}": "department", "{{姓名}}": "name", "{{性别}}": "gender",
        "{{年龄}}": "age", "{{住院号}}": "hospital_id", "{{床号}}": "bed_number",
        "{{入院日期}}": "admission_date", "{{出院日期}}": "discharge_date",
        "{{手术日期}}": "surgery_date", "{{手术名称}}": "surgery_name",
        "{{出院诊断}}": "diagnosis", "{{联系电话}}": "phone", "{{经治医生}}": "doctor",
    }
}
