# -*- coding: utf-8 -*-
"""
模块功能：负责读取和解析Excel文件。
"""
import os
import openpyxl
import xlrd
from ...config import CONFIG
from ...utils import format_text, excel_date_to_str

def _get_excel_rows(file_path, log_func):
    """根据文件扩展名选择合适的库来读取并迭代行"""
    log_func(f"正在读取文件: {os.path.basename(file_path)}")
    if file_path.lower().endswith('.xls'):
        book = xlrd.open_workbook(file_path)
        sheet = book.sheet_by_index(0)
        for i in range(sheet.nrows):
            # 返回一个生成器，包含行值和书籍的日期模式
            yield [sheet.cell_value(i, j) for j in range(sheet.ncols)], book.datemode
    elif file_path.lower().endswith('.xlsx'):
        # 注意：不使用 read_only=True 以兼容非标准文件
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        # 返回一个生成器，行值和固定的日期模式0
        for row in sheet.iter_rows(values_only=True):
            yield row, 0
    else:
        log_func(f"不支持的文件格式: {file_path}", "error")
        return

def find_header_and_map_cols(file_path, required_keys, log_func):
    """
    自动查找标题行并返回列索引映射。
    :param file_path: Excel文件路径。
    :param required_keys: 内部列名（如 'name', 'hospital_id'）。
    :param log_func: 日志记录函数。
    :return: (列映射字典, 标题行索引, 日期模式) 或 (None, -1, 0)
    """
    required_cols_text = {CONFIG['column_mapping'][key] for key in required_keys}
    
    for i, (row_values, datemode) in enumerate(_get_excel_rows(file_path, log_func)):
        row_values_cleaned = {format_text(v) for v in row_values}
        
        if required_cols_text.issubset(row_values_cleaned):
            log_func(f"在文件 '{os.path.basename(file_path)}' 中自动检测到标题行位于第 {i + 1} 行。")
            header_map = {format_text(col_name): col_idx for col_idx, col_name in enumerate(row_values)}
            col_map = {}
            for internal_key, excel_name in CONFIG['column_mapping'].items():
                if excel_name in header_map:
                    col_map[internal_key] = header_map[excel_name]
            return col_map, i, datemode
    
    log_func(f"在文件 '{os.path.basename(file_path)}' 中未能找到包含所有必需列的标题行: {', '.join(required_cols_text)}", "error")
    return None, -1, 0

def read_surgery_data(file_path, col_map, header_row_idx, log_func):
    """从手术查询文件中读取并生成(yield)记录"""
    h_id_idx = col_map.get('hospital_id')
    name_idx = col_map.get('name')
    bed_idx = col_map.get('bed_number')

    for i, (row_values, _) in enumerate(_get_excel_rows(file_path, log_func)):
        if i <= header_row_idx:
            continue
        
        h_id = format_text(row_values[h_id_idx]) if h_id_idx is not None and len(row_values) > h_id_idx else ""
        name = format_text(row_values[name_idx]) if name_idx is not None and len(row_values) > name_idx else ""
        bed = format_text(row_values[bed_idx]) if bed_idx is not None and len(row_values) > bed_idx else ""

        if h_id and name and bed:
            h_id_cleaned = h_id.lstrip('0') if h_id != '0' else '0'
            yield (h_id_cleaned, name, bed)

def read_patient_data(file_path, col_map, header_row_idx, log_func):
    """从主患者文件中读取并生成(yield)记录"""
    internal_keys = list(CONFIG['column_mapping'].keys())
    
    for i, (row_values, file_datemode) in enumerate(_get_excel_rows(file_path, log_func)):
        if i <= header_row_idx:
            continue

        record = {}
        for key in internal_keys:
            idx = col_map.get(key)
            if idx is not None and len(row_values) > idx:
                raw_val = row_values[idx]
                if "date" in key:
                    record[key] = excel_date_to_str(raw_val, file_datemode)
                else:
                    record[key] = format_text(raw_val)
            else:
                record[key] = ""
        
        if any(not record.get(req_key) for req_key in CONFIG['required_patient_cols']):
            log_func(f"跳过第 {i+1} 行，因为缺少必要信息。", "warning")
            continue
        
        yield [record.get(key, "") for key in internal_keys]
