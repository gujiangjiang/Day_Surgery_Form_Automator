# -*- coding: utf-8 -*-
"""
核心业务逻辑编排模块。
此类协调数据读取、数据库处理和文档生成。
"""
import traceback
import threading
import logging
from collections import deque
from pathlib import Path

from .modules import excel_reader
from .modules import doc_writer
from .modules.db_manager import DatabaseManager
from ..config import CONFIG

# 获取该模块的logger实例
logger = logging.getLogger(__name__)

class DocumentGenerator:
    def __init__(self, excel_path, surgery_query_paths, template_path, output_dir, 
                 progress_callback, completion_callback, message_callback):
        self.excel_path = excel_path
        self.surgery_query_paths = surgery_query_paths
        self.template_path = template_path
        # 将输出目录确保为Path对象
        self.output_dir = Path(output_dir)
        
        # --- 回调函数 ---
        self.update_progress = progress_callback
        self.on_completion = completion_callback
        self.show_message = message_callback
        
        self.stop_event = threading.Event()

    def stop(self):
        """设置停止事件，中断生成过程。"""
        self.stop_event.set()
        logger.warning("正在发送停止信号...")

    def run(self):
        """主执行函数，负责编排整个流程"""
        db_manager = None
        try:
            logger.notice("后台处理任务已启动，正在准备环境...")
            
            # 使用Path对象创建目录
            self.output_dir.mkdir(parents=True, exist_ok=True)
            db_manager = DatabaseManager()

            # --- 手术文件处理 ---
            logger.notice("开始处理手术查询文件...")
            if not self.surgery_query_paths:
                logger.notice("未选择任何手术查询文件，跳过床号补充步骤。")
            else:
                total_records_added = 0
                for file_path in self.surgery_query_paths:
                    if self.stop_event.is_set(): return
                    records, _ = excel_reader.process_file(file_path)
                    if records:
                        count = db_manager.load_surgery_data(records)
                        total_records_added += count
                logger.info(f"所有手术查询文件处理完毕，共加载了 {total_records_added} 条有效的床号记录。")

            if self.stop_event.is_set(): return

            # --- 主患者文件处理 ---
            logger.notice("开始处理主患者列表文件...")
            patient_records, col_map = excel_reader.process_file(self.excel_path, 'patient')

            if not patient_records:
                required_cols_str = ', '.join([CONFIG['column_mapping'][k] for k in CONFIG['required_patient_cols']])
                self.show_message("error", "读取失败", f"在 '出院患者列表' 文件中无法自动定位标题行或未找到任何有效数据。\n请确保文件包含以下列: {required_cols_str}")
                return

            if col_map and 'bed_number' not in col_map:
                msg = "主Excel文件中未找到“床号”列。 "
                msg += "将尝试从手术查询文件补充。" if self.surgery_query_paths else "床号信息可能为空。"
                logger.warning(msg)

            patient_count = db_manager.load_patient_data(patient_records)
            if patient_count <= 0:
                logger.error("未从主文件中加载任何有效的患者记录。")
                return
            logger.info(f"成功从主文件加载了 {patient_count} 条患者记录。")

            if self.stop_event.is_set(): return

            # --- 后续处理流程 ---
            final_patient_rows = db_manager.query_final_data()
            if self.stop_event.is_set(): return

            if not final_patient_rows:
                msg = f"未找到住院天数 <= {CONFIG['day_surgery_max_days']} 天的记录。"
                logger.error(msg)
                self.show_message("error", "无数据", msg)
                return

            total_rows = len(final_patient_rows)
            logger.info(f"共找到 {total_rows} 条符合条件的记录，开始生成文档...")
            success_count = 0
            unmatched_patients = deque()
            
            for index, row in enumerate(final_patient_rows):
                if self.stop_event.is_set():
                    logger.warning("生成过程已由用户手动停止。")
                    break
                try:
                    is_unmatched, filename = doc_writer.generate_single_document(row, self.template_path, self.output_dir)
                    logger.notice(f"已生成: {filename}")
                    if is_unmatched:
                        unmatched_patients.append(f"{row['name']} (住院号: {row['hospital_id']})")
                    success_count += 1
                except Exception as e:
                    logger.error(f"处理行 {index + 1} (姓名: {row['name']}) 时发生错误: {e}", exc_info=True)
                self.update_progress((index + 1) / total_rows * 100)
            
            if not self.stop_event.is_set():
                logger.notice("="*30, extra={'simple': True})
                # --- 修复：在完成日志中显示输出路径 ---
                logger.info(f"处理完成！共生成 {success_count} 份文档，已保存至: {self.output_dir}")
                
                if unmatched_patients:
                    summary_message = f"注意：有 {len(unmatched_patients)} 位符合条件的日间手术患者未能匹配到床号：\n\n" + "\n".join(unmatched_patients)
                    logger.warning("="*30, extra={'simple': True})
                    logger.warning("以下日间手术患者未能匹配到床号:", extra={'simple': True})
                    for patient_info in unmatched_patients:
                        logger.warning(f"- {patient_info}", extra={'simple': True})
                    self.show_message("warning", "匹配提醒", summary_message)
                
                # --- 修复：在最终弹窗中也显示输出路径 ---
                final_message = (
                    f"成功生成 {success_count} 份随访表。\n\n"
                    f"文件保存在: {self.output_dir}"
                )
                self.show_message("info", "完成", final_message)

        except Exception as e:
            if not self.stop_event.is_set():
                logger.critical(f"发生严重错误: {e}", exc_info=True)
                self.show_message("error", "严重错误", f"处理过程中发生严重错误：\n{e}\n\n详情请查看日志文件。")
        finally:
            if db_manager:
                db_manager.close()
            if self.on_completion:
                self.on_completion()
