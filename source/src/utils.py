# -*- coding: utf-8 -*-
"""
存放通用的工具函数。
"""

from datetime import datetime, timedelta, date
import xlrd

def format_text(value):
    """通用文本格式化函数，处理None并去除首尾空格"""
    if value is None:
        return ""
    return str(value).strip()

def excel_date_to_str(val, book_datemode=0):
    """
    将来自 xlrd 或 openpyxl 的各种日期类型统一转换为 YYYY-MM-DD 格式字符串。
    此函数经过优化，调整了类型检查顺序并增加了注释。

    :param val: 单元格原始值
    :param book_datemode: 仅用于 xlrd，0 for 1900-based, 1 for 1904-based.
    """
    if val is None or val == '':
        return ""
    
    # 优先处理最明确的类型：datetime 对象 (通常来自 openpyxl)
    if isinstance(val, (datetime, date)):
        return val.strftime('%Y-%m-%d')
    
    # 其次处理数字类型 (通常来自 xlrd 或被 openpyxl 错误读取为数字的日期)
    if isinstance(val, (int, float)):
        try:
            # 首先尝试使用 xlrd 的标准库函数进行转换，这是最可靠的方法
            return xlrd.xldate_as_datetime(val, book_datemode).strftime('%Y-%m-%d')
        except (ValueError, TypeError, xlrd.xldate.XLDateError):
            # 如果 xlrd 转换失败 (例如，它不是一个有效的xlrd日期数字)，
            # 则尝试一个通用的Excel数字日期转换方法。
            # Excel的日期序列从1开始，对应1900-01-01，但它错误地认为1900是闰年。
            # Python的 datetime(1899, 12, 30) + timedelta(days=val) 是处理此问题的标准技巧。
            try:
                return (datetime(1899, 12, 30) + timedelta(days=val)).strftime('%Y-%m-%d')
            except (ValueError, TypeError):
                # 如果所有日期转换都失败，则将其作为普通数字字符串返回
                return str(val)

    # 最后处理字符串类型
    if isinstance(val, str):
        val = val.strip()
        try:
            # 尝试直接解析 "YYYY-MM-DD HH:MM:SS" 或 "YYYY-MM-DD" 这样的标准格式
            return datetime.strptime(val.split()[0], '%Y-%m-%d').strftime('%Y-%m-%d')
        except ValueError:
            # 如果字符串不是标准日期格式，它可能是一个数字字符串 (如 "44562.0")
            # 递归调用自身，让处理数字的逻辑部分来解决
            try:
                return excel_date_to_str(float(val), book_datemode)
            except (ValueError, TypeError):
                # 如果所有尝试都失败，返回原始的、清理过的字符串
                return val
    
    # 对于所有其他未预料到的类型，进行基本的文本格式化
    return format_text(val)


def get_day_after_discharge(discharge_date_str, days=7):
    """计算出院后N天的日期"""
    if not discharge_date_str:
        return ""
    try:
        base_date = datetime.strptime(discharge_date_str, "%Y-%m-%d")
        return (base_date + timedelta(days=days)).strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        return ""
