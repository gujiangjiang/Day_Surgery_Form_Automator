# -*- coding: utf-8 -*-
"""
模块功能：负责高效读取和解析Excel文件。
采用单次遍历方法，避免重复读取文件。
"""
import os
import openpyxl
import xlrd
from ...config import CONFIG
from ...utils import format_text, excel_date_to_str

def _get_rows_generator(file_path, log_func, read_only=False):
    """根据文件扩展名，创建一个行的生成器。"""
    if file_path.lower().endswith('.xls'):
        try:
            book = xlrd.open_workbook(file_path)
            sheet = book.sheet_by_index(0)
            for i in range(sheet.nrows):
                yield [sheet.cell_value(i, j) for j in range(sheet.ncols)], book.datemode
        except Exception as e:
            log_func(f"读取 .xls 文件 '{os.path.basename(file_path)}' 时出错: {e}", "error")
            return
    elif file_path.lower().endswith('.xlsx'):
        try:
            workbook = openpyxl.load_workbook(file_path, read_only=read_only)
            sheet = workbook.active
            for row in sheet.iter_rows(values_only=True):
                yield row, 0
        except Exception as e:
            log_func(f"使用 read_only={read_only} 模式读取 '{os.path.basename(file_path)}' 时出错: {e}", "error")
            return
    else:
        log_func(f"不支持的文件格式: {file_path}", "error")
        return

def _process_surgery_row(row_values, col_map, datemode):
    """处理单行手术数据。"""
    h_id_idx = col_map.get('hospital_id')
    name_idx = col_map.get('name')
    bed_idx = col_map.get('bed_number')

    h_id = format_text(row_values[h_id_idx]) if h_id_idx is not None and len(row_values) > h_id_idx else ""
    name = format_text(row_values[name_idx]) if name_idx is not None and len(row_values) > name_idx else ""
    bed = format_text(row_values[bed_idx]) if bed_idx is not None and len(row_values) > bed_idx else ""

    if h_id and name and bed:
        h_id_cleaned = h_id.lstrip('0') if h_id != '0' else '0'
        return (h_id_cleaned, name, bed)
    return None

def _process_patient_row(row_values, col_map, datemode):
    """处理单行患者数据。"""
    internal_keys = list(CONFIG['column_mapping'].keys())
    record = {}
    for key in internal_keys:
        idx = col_map.get(key)
        if idx is not None and len(row_values) > idx:
            raw_val = row_values[idx]
            if "date" in key:
                record[key] = excel_date_to_str(raw_val, datemode)
            else:
                record[key] = format_text(raw_val)
        else:
            record[key] = ""
    
    if any(not record.get(req_key) for req_key in CONFIG['required_patient_cols']):
        return None
    
    return [record.get(key, "") for key in internal_keys]

def process_file(file_path, required_keys, processor_type, log_func):
    """
    一次性读取并处理整个Excel文件，自动处理读取模式和标题行查找。
    :return: (list, dict) 包含所有已处理记录的列表和列映射字典。
    """
    processor_map = {'surgery': _process_surgery_row, 'patient': _process_patient_row}
    row_processor = processor_map.get(processor_type)
    if not row_processor:
        log_func(f"未知的处理器类型: {processor_type}", "error")
        return [], None

    def _process_stream(rows_generator):
        col_map, datemode = None, 0
        processed_records = []
        header_found = False
        skipped_count = 0

        for i, (row_values, file_datemode) in enumerate(rows_generator):
            if row_values is None: continue

            if not header_found:
                required_cols_text = {CONFIG['column_mapping'][key] for key in required_keys}
                row_values_cleaned = {format_text(v) for v in row_values if v is not None}
                
                if required_cols_text.issubset(row_values_cleaned):
                    log_func(f"在文件 '{os.path.basename(file_path)}' 中自动检测到标题行位于第 {i + 1} 行。")
                    header_map = {format_text(col_name): idx for idx, col_name in enumerate(row_values)}
                    col_map = {}
                    for internal_key, excel_name in CONFIG['column_mapping'].items():
                        if excel_name in header_map:
                            col_map[internal_key] = header_map[excel_name]
                    
                    header_found = True
                    datemode = file_datemode
            else:
                record = row_processor(row_values, col_map, datemode)
                if record:
                    processed_records.append(record)
                elif processor_type == 'patient':
                    skipped_count += 1
        
        if skipped_count > 0:
            log_func(f"因缺少必要信息，共跳过了 {skipped_count} 行。", "warning")
        
        return (processed_records, col_map) if header_found else (None, None)

    log_func(f"正在分析文件: {os.path.basename(file_path)}")
    
    if file_path.lower().endswith('.xlsx'):
        log_func("...尝试使用快速只读模式。")
        records, col_map = _process_stream(_get_rows_generator(file_path, log_func, read_only=True))
        if records is not None:
            return records, col_map

        log_func("...快速模式未能找到标题行或处理失败，自动切换到标准模式。", "warning")
        records, col_map = _process_stream(_get_rows_generator(file_path, log_func, read_only=False))
        if records is not None:
            return records, col_map
    else:
        records, col_map = _process_stream(_get_rows_generator(file_path, log_func, read_only=False))
        if records is not None:
            return records, col_map

    log_func(f"在文件 '{os.path.basename(file_path)}' 中未能找到包含所有必需列的标题行: {', '.join({CONFIG['column_mapping'][key] for key in required_keys})}", "error")
    return [], None
