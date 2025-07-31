# -*- coding: utf-8 -*-
"""
模块功能：负责读取和解析Excel文件。
"""
import os
import openpyxl
import xlrd
from ...config import CONFIG
from ...utils import format_text, excel_date_to_str

def _get_excel_rows(file_path, log_func, read_only=False):
    """根据文件扩展名选择合适的库来读取并迭代行"""
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

def find_header_and_map_cols(file_path, required_keys, log_func):
    """
    自动查找标题行并返回列索引映射和成功的读取模式。
    对于.xlsx文件，优先使用快速只读模式，失败则回退到标准模式。
    """
    required_cols_text = {CONFIG['column_mapping'][key] for key in required_keys}

    def _search_for_header(read_only_mode):
        """一个辅助函数，用于在给定模式下搜索标题行。"""
        # 将标题行搜索限制在文件的前20行，以提高性能
        for i, (row_values, datemode) in enumerate(_get_excel_rows(file_path, log_func, read_only=read_only_mode)):
            if i > 20: return None, -1, 0
            if row_values is None: continue
            
            row_values_cleaned = {format_text(v) for v in row_values if v is not None}
            
            if required_cols_text.issubset(row_values_cleaned):
                log_func(f"在文件 '{os.path.basename(file_path)}' 中自动检测到标题行位于第 {i + 1} 行。")
                header_map = {format_text(col_name): col_idx for col_idx, col_name in enumerate(row_values)}
                col_map = {}
                for internal_key, excel_name in CONFIG['column_mapping'].items():
                    if excel_name in header_map:
                        col_map[internal_key] = header_map[excel_name]
                return col_map, i, datemode
        return None, -1, 0

    log_func(f"正在分析文件: {os.path.basename(file_path)}")

    # 对于 .xlsx 文件，优先尝试快速模式
    if file_path.lower().endswith('.xlsx'):
        log_func("...尝试使用快速只读模式。")
        col_map, header_row_idx, datemode = _search_for_header(read_only_mode=True)
        if col_map:
            return col_map, header_row_idx, datemode, True # 返回成功，并告知使用的是只读模式

        log_func("...快速模式未能找到标题行，自动切换到标准模式。", "warning")
        col_map, header_row_idx, datemode = _search_for_header(read_only_mode=False)
        if col_map:
            return col_map, header_row_idx, datemode, False # 返回成功，并告知使用的是标准模式
    else: # 对于 .xls 文件，只有一种模式
        col_map, header_row_idx, datemode = _search_for_header(read_only_mode=False)
        if col_map:
            return col_map, header_row_idx, datemode, False

    # 如果所有尝试都失败
    log_func(f"在文件 '{os.path.basename(file_path)}' 中未能找到包含所有必需列的标题行: {', '.join(required_cols_text)}", "error")
    return None, -1, 0, False

def read_surgery_data(file_path, col_map, header_row_idx, log_func, read_only_mode):
    """从手术查询文件中读取并生成(yield)记录"""
    h_id_idx = col_map.get('hospital_id')
    name_idx = col_map.get('name')
    bed_idx = col_map.get('bed_number')

    for i, (row_values, _) in enumerate(_get_excel_rows(file_path, log_func, read_only=read_only_mode)):
        if i <= header_row_idx:
            continue
        
        h_id = format_text(row_values[h_id_idx]) if h_id_idx is not None and len(row_values) > h_id_idx else ""
        name = format_text(row_values[name_idx]) if name_idx is not None and len(row_values) > name_idx else ""
        bed = format_text(row_values[bed_idx]) if bed_idx is not None and len(row_values) > bed_idx else ""

        if h_id and name and bed:
            h_id_cleaned = h_id.lstrip('0') if h_id != '0' else '0'
            yield (h_id_cleaned, name, bed)

def read_patient_data(file_path, col_map, header_row_idx, log_func, read_only_mode):
    """从主患者文件中读取并生成(yield)记录"""
    internal_keys = list(CONFIG['column_mapping'].keys())
    
    for i, (row_values, file_datemode) in enumerate(_get_excel_rows(file_path, log_func, read_only=read_only_mode)):
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
