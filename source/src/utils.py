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
    :param val: 单元格原始值
    :param book_datemode: 仅用于 xlrd，0 for 1900-based, 1 for 1904-based.
    """
    if val is None or val == '':
        return ""
    # 如果已经是 datetime 对象 (来自 openpyxl)
    if isinstance(val, (datetime, date)):
        return val.strftime('%Y-%m-%d')
    # 如果是字符串
    if isinstance(val, str):
        val = val.strip()
        try:
            # 尝试直接解析 "YYYY-MM-DD HH:MM:SS" 或 "YYYY-MM-DD"
            return datetime.strptime(val.split()[0], '%Y-%m-%d').strftime('%Y-%m-%d')
        except ValueError:
            # 可能是数字字符串 "44562.0"
            try:
                return excel_date_to_str(float(val), book_datemode)
            except (ValueError, TypeError):
                return val # 无法解析，返回原样
    # 如果是数字 (来自 xlrd 或某些 .xlsx 文件)
    if isinstance(val, (int, float)):
        try:
            # xlrd 的 xldate_as_datetime 处理
            return xlrd.xldate_as_datetime(val, book_datemode).strftime('%Y-%m-%d')
        except (ValueError, TypeError, xlrd.xldate.XLDateError):
            # 如果失败，尝试 openpyxl 的数字转日期逻辑 (1899-12-30)
            try:
                # Excel 的序列日期从1开始，并且错误地认为1900是闰年，所以要小心处理
                # timedelta(days=val-1) 对于从1900-01-01开始的系统
                # timedelta(days=val) 对于从1899-12-31开始的系统
                # Python 的 datetime(1899, 12, 30) + timedelta(days=val) 是最常见的转换方式
                return (datetime(1899, 12, 30) + timedelta(days=val)).strftime('%Y-%m-%d')
            except (ValueError, TypeError):
                return str(val) # 转换失败
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
