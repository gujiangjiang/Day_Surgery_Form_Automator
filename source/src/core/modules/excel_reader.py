# -*- coding: utf-8 -*-
"""
模块功能：负责高效读取和解析Excel文件。
采用单次遍历方法，避免重复读取文件。
"""
import logging # 导入logging模块
from pathlib import Path
import openpyxl
import xlrd
from ...config import CONFIG
from ...utils import format_text, excel_date_to_str

# 获取该模块的logger实例
logger = logging.getLogger(__name__)

def _get_rows_generator(file_path, read_only=False):
    """根据文件扩展名，创建一个行的生成器。"""
    if not Path(file_path).exists():
        logger.error(f"文件不存在: {file_path}")
        return
        
    if file_path.lower().endswith('.xls'):
        try:
            # 增加 formatting_info=False 以提高兼容性
            book = xlrd.open_workbook(file_path)
            if book.nsheets == 0:
                logger.error(f"Excel文件为空: {file_name}")
                return
            sheet = book.sheet_by_index(0)
            for i in range(sheet.nrows):
                yield [sheet.cell_value(i, j) for j in range(sheet.ncols)], book.datemode
        except Exception as e:
            logger.error(f"读取 .xls 文件 '{Path(file_path).name}' 时出错: {e}", exc_info=True)
            return
    elif file_path.lower().endswith('.xlsx'):
        try:
            # data_only=True 确保读取的是计算后的值而不是公式
            workbook = openpyxl.load_workbook(file_path, read_only=read_only, data_only=True)
            sheet = workbook.active
            if sheet is None:
                logger.error(f"无法获取有效的工作表: {Path(file_path).name}")
                return
            for row in sheet.iter_rows(values_only=True):
                yield row, 0
        except Exception as e:
            logger.error(f"读取 .xlsx 文件 '{Path(file_path).name}' 时出错 (模式:{read_only}): {e}", exc_info=True)
            return
    else:
        logger.error(f"不支持的文件格式: {file_path}")
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
    
    # 检查核心必要列是否为空
    if any(not record.get(req_key) for req_key in CONFIG['required_patient_cols']):
        return None
    
    return [record.get(key, "") for key in internal_keys]

def process_file(file_path, processor_type='surgery'):
    """
    一次性读取并处理整个Excel文件，自动处理读取模式和标题行查找。
    :return: (list, dict) 包含所有已处理记录的列表和列映射字典。
    """
    processor_map = {'surgery': _process_surgery_row, 'patient': _process_patient_row}
    row_processor = processor_map.get(processor_type)
    
    required_keys_map = {
        'surgery': CONFIG['required_surgery_cols'],
        'patient': CONFIG['required_patient_cols']
    }
    required_keys = required_keys_map.get(processor_type)

    if not row_processor or not required_keys:
        logger.error(f"未知的处理器类型或未配置必须列: {processor_type}")
        return [], None

    def _process_stream(rows_generator):
        col_map, datemode = None, 0
        processed_records = []
        header_found = False
        skipped_count = 0
        
        file_name = Path(file_path).name

        for i, (row_values, file_datemode) in enumerate(rows_generator):
            if row_values is None: continue

            if not header_found:
                # 预处理 row_values 以便进行包含检查
                row_values_cleaned = [format_text(v) for v in row_values if v is not None]
                required_cols_text = {CONFIG['column_mapping'][key] for key in required_keys}
                
                if required_cols_text.issubset(set(row_values_cleaned)):
                    logger.notice(f"在文件 '{file_name}' 中自动检测到标题行位于第 {i + 1} 行。")
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
            logger.warning(f"文件 '{file_name}' 中由于缺失必要信息，共跳过了 {skipped_count} 行记录。")
        
        return (processed_records, col_map) if header_found else (None, None)

    logger.notice(f"正在分析文件: {Path(file_path).name}")
    
    # 获取生成器
    rows_gen = _get_rows_generator(file_path, read_only=True)
    if rows_gen:
        if file_path.lower().endswith('.xlsx'):
            logger.notice("...尝试使用快速只读模式。")
            records, col_map = _process_stream(rows_gen)
            if records is not None:
                return records, col_map

            logger.warning("...快速模式未能找到标题行，自动切换到标准加载模式。")
            rows_gen_standard = _get_rows_generator(file_path, read_only=False)
            if rows_gen_standard:
                records, col_map = _process_stream(rows_gen_standard)
                if records is not None:
                    return records, col_map
        else:
            records, col_map = _process_stream(rows_gen)
            if records is not None:
                return records, col_map

    return [], None
    logger.error(f"在文件 '{Path(file_path).name}' 中未能找到包含所有必需列的标题行。")
