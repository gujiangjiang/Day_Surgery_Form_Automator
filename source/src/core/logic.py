# -*- coding: utf-8 -*-
"""
核心业务逻辑编排模块。
此类协调数据读取、数据库处理和文档生成。
"""
import os
import traceback
import threading
from tkinter import messagebox
from collections import deque

# 导入路径已更新以反映 'modules' 子文件夹
from .modules import excel_reader
from .modules import doc_writer
from .modules.db_manager import DatabaseManager
from ..config import CONFIG

class DocumentGenerator:
    def __init__(self, excel_path, surgery_query_paths, template_path, output_dir, app_instance):
        self.excel_path = excel_path
        self.surgery_query_paths = surgery_query_paths
        self.template_path = template_path
        self.output_dir = output_dir
        self.app = app_instance
        self.stop_event = threading.Event()

    def stop(self):
        """设置停止事件，中断生成过程。"""
        self.stop_event.set()
        self.log("正在发送停止信号...", "warning")

    def log(self, message, level="info"):
        self.app.log_message(message, level)

    def update_progress(self, value):
        self.app.update_progress(value)

    def _load_surgery_data(self, db_manager):
        """加载所有手术查询文件数据到数据库"""
        self.log("开始处理手术查询文件...")
        if not self.surgery_query_paths:
            self.log("未选择任何手术查询文件，跳过床号补充步骤。", "info")
            return

        total_records_added = 0
        for file_path in self.surgery_query_paths:
            if self.stop_event.is_set(): return
            col_map, header_row_idx, _ = excel_reader.find_header_and_map_cols(file_path, CONFIG['required_surgery_cols'], self.log)
            if col_map:
                records = list(excel_reader.read_surgery_data(file_path, col_map, header_row_idx, self.log))
                count = db_manager.load_surgery_data(records)
                total_records_added += count
        
        self.log(f"所有手术查询文件处理完毕，共加载了 {total_records_added} 条有效的床号记录。")

    def _load_patient_data(self, db_manager):
        """加载主患者列表文件数据到数据库"""
        self.log("开始读取出院患者列表Excel文件...")
        col_map, header_row_idx, _ = excel_reader.find_header_and_map_cols(self.excel_path, CONFIG['required_patient_cols'], self.log)

        if not col_map:
            messagebox.showerror("读取失败", f"在 '出院患者列表' 文件中无法自动定位标题行。\n请确保文件包含以下列: {', '.join([CONFIG['column_mapping'][k] for k in CONFIG['required_patient_cols']])}")
            return False
            
        if 'bed_number' not in col_map:
            if self.surgery_query_paths:
                self.log("警告：主Excel文件中未找到“床号”列。将尝试从手术查询文件补充。", "warning")
            else:
                self.log("警告：主Excel文件中未找到“床号”列，床号信息可能为空。", "warning")

        records = list(excel_reader.read_patient_data(self.excel_path, col_map, header_row_idx, self.log))
        count = db_manager.load_patient_data(records)

        if count > 0:
            self.log(f"成功从主文件加载了 {count} 条患者记录。")
            return True
        else:
            self.log("未从主文件中加载任何有效的患者记录。", "error")
            return False

    def run(self):
        """主执行函数，负责编排整个流程"""
        db_manager = None
        try:
            os.makedirs(self.output_dir, exist_ok=True)
            db_manager = DatabaseManager(self.log)

            self._load_surgery_data(db_manager)
            if self.stop_event.is_set(): return

            if not self._load_patient_data(db_manager):
                messagebox.showerror("错误", "无法从'出院患者列表'加载任何有效数据，程序终止。")
                return
            if self.stop_event.is_set(): return

            final_patient_rows = db_manager.query_final_data()
            if self.stop_event.is_set(): return

            if not final_patient_rows:
                msg = f"错误：未找到住院天数 <= {CONFIG['day_surgery_max_days']} 天的记录。"
                self.log(msg, "error")
                messagebox.showerror("无数据", msg)
                return

            total_rows = len(final_patient_rows)
            self.log(f"共找到 {total_rows} 条符合条件的记录，开始生成文档...")
            success_count = 0
            unmatched_patients = deque()
            
            for index, row in enumerate(final_patient_rows):
                if self.stop_event.is_set():
                    self.log("生成过程已由用户手动停止。", "warning")
                    break
                try:
                    is_unmatched, filename = doc_writer.generate_single_document(row, self.template_path, self.output_dir)
                    self.log(f"已生成: {filename}")
                    if is_unmatched:
                        unmatched_patients.append(f"{row['name']} (住院号: {row['hospital_id']})")
                    success_count += 1
                except Exception as e:
                    self.log(f"处理行 {index + 1} (姓名: {row['name']}) 时发生错误: {e}", "error")
                self.update_progress((index + 1) / total_rows * 100)
            
            if not self.stop_event.is_set():
                self.app.log_raw("="*30)
                self.log(f"处理完成！成功生成 {success_count} 份文档。")
                
                if unmatched_patients:
                    summary_message = f"注意：有 {len(unmatched_patients)} 位符合条件的日间手术患者未能匹配到床号：\n\n" + "\n".join(unmatched_patients)
                    self.app.log_raw("="*30)
                    self.log("以下日间手术患者未能匹配到床号:", "warning")
                    for patient_info in unmatched_patients:
                        self.log(f"- {patient_info}", "warning")
                    messagebox.showwarning("匹配提醒", summary_message)
                
                final_message = f"成功生成 {success_count} 份随访表。\n" \
                              f"文件保存在: {self.output_dir}"
                messagebox.showinfo("完成", final_message)

        except Exception as e:
            if not self.stop_event.is_set():
                self.log(f"发生严重错误: {e}", "error")
                self.log(traceback.format_exc(), "error")
                messagebox.showerror("严重错误", f"处理过程中发生严重错误：\n{e}")
        finally:
            if db_manager:
                db_manager.close()
            self.app.generation_finished()
