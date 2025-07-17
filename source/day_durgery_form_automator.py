# -*- coding: utf-8 -*-
"""
日间手术随访表生成系统 (优化修复版)

功能：
- 自动从Excel批量生成Word随访表
- GUI界面，操作直观
- 自动检测Excel标题行
- 支持进度条和实时日志
- 可配置关键参数

作者：顾江江 (由AI优化和修复)
"""

import os
import sys
import threading
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, scrolledtext

# 检查并安装必要的库
try:
    from docx import Document
    import pandas as pd
except ImportError:
    messagebox.showerror(
        "依赖缺失",
        "缺少必要的库 (pandas, python-docx)。\n"
        "请在命令行运行 'pip install pandas python-docx' 来安装。"
    )
    sys.exit(1)

# ======================== 全局配置 ========================
# 将关键参数放在这里，方便未来修改
CONFIG = {
    "app_title": "日间手术随访表生成系统 V2.0",
    "day_surgery_max_days": 2,  # 定义日间手术的最大住院天数

    # 核心：定义程序内部字段名与Excel列名的映射关系
    "column_mapping": {
        # 内部字段名: Excel中的列名
        "name": "姓名",
        "department": "出院科室",
        "hospital_id": "住院号",
        "discharge_date": "出院日期",
        "hospital_days": "住院天数",
        "gender": "性别",
        "age": "年龄",
        "bed_number": "床号",
        "admission_date": "入院日期",
        "surgery_date": "手术日期",
        "surgery_name": "手术名称",
        "diagnosis": "出院诊断",
        "phone": "联系电话",
        "doctor": "经治医生"
    },

    # 定义哪些内部字段是必须存在的
    "required_internal_keys": [
        "name", "department", "hospital_id", "discharge_date", "hospital_days"
    ],

    # 定义Word模板占位符与内部字段名的映射关系
    "template_placeholders": {
        "{{科室}}": "department",
        "{{姓名}}": "name",
        "{{性别}}": "gender",
        "{{年龄}}": "age",
        "{{住院号}}": "hospital_id",
        "{{床号}}": "bed_number",
        "{{入院日期}}": "admission_date",
        "{{出院日期}}": "discharge_date",
        "{{手术日期}}": "surgery_date",
        "{{手术名称}}": "surgery_name",
        "{{出院诊断}}": "diagnosis",
        "{{联系电话}}": "phone",
        "{{经治医生}}": "doctor",
    }
}


# ======================== 工具函数 ========================
def excel_date_to_str(excel_date):
    """智能转换Excel中的各种日期格式为 'YYYY-MM-DD' 字符串"""
    if pd.isna(excel_date) or excel_date in ("", None):
        return ""
    try:
        # 优先使用pandas进行转换，兼容性更强
        return pd.to_datetime(excel_date).strftime('%Y-%m-%d')
    except (ValueError, TypeError):
        # 如果pandas转换失败，使用原始的回退方案
        return str(excel_date).split()[0]

def get_day_after_discharge(discharge_date_str, days=7):
    """根据出院日期字符串，计算N天后的日期"""
    if not discharge_date_str:
        return ""
    try:
        base_date = datetime.strptime(discharge_date_str, "%Y-%m-%d")
        return (base_date + timedelta(days=days)).strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        return ""

# ======================== 核心逻辑类 ========================
class DocumentGenerator:
    def __init__(self, excel_path, template_path, output_dir, app_instance):
        self.excel_path = excel_path
        self.template_path = template_path
        self.output_dir = output_dir
        self.app = app_instance

    def log(self, message):
        self.app.log_message(message)

    def update_progress(self, value):
        self.app.update_progress(value)

    def run(self):
        """执行生成过程的主方法"""
        try:
            self.log("开始读取Excel文件...")
            df = self.read_and_prepare_excel()
            if df is None: return

            self.log("分析出院日期分布...")
            selected_month = self.determine_dominant_month(df)
            if not selected_month:
                self.log("用户取消或无有效数据，操作中止。")
                return
            patient_year_month = f"{selected_month[:4]}年{selected_month[5:]}月"
            self.log(f"已确定主导月份为: {patient_year_month}")

            day_surgery_df = self.filter_day_surgery_patients(df)
            if day_surgery_df.empty:
                self.log(f"错误：未找到住院天数 <= {CONFIG['day_surgery_max_days']} 天的记录。")
                messagebox.showerror("无数据", f"未找到住院天数 <= {CONFIG['day_surgery_max_days']} 天的记录。")
                return

            total_rows = len(day_surgery_df)
            self.log(f"共找到 {total_rows} 条符合条件的记录，开始生成文档...")
            success_count = 0
            for index, row in enumerate(day_surgery_df.itertuples()):
                try:
                    self.generate_single_document(row, patient_year_month)
                    success_count += 1
                except Exception as e:
                    self.log(f"处理行 {getattr(row, 'Index', 'N/A')} 时发生错误: {e}")
                self.update_progress((index + 1) / total_rows * 100)

            self.log("="*30)
            self.log(f"处理完成！成功生成 {success_count} 份文档。")
            messagebox.showinfo("完成", f"成功生成 {success_count} 份随访表。\n文件保存在: {self.output_dir}")

        except Exception as e:
            self.log(f"发生严重错误: {e}")
            messagebox.showerror("严重错误", f"处理过程中发生严重错误：\n{e}")
        finally:
            self.app.generation_finished()

    def read_and_prepare_excel(self):
        """读取Excel，找到标题行，并将列名转换为内部标准字段名"""
        try:
            df_full = pd.read_excel(self.excel_path, sheet_name=0, header=None, dtype=str)
            
            required_excel_cols = {CONFIG['column_mapping'][key] for key in CONFIG['required_internal_keys']}

            header_row_index = -1
            for i, row in df_full.iterrows():
                row_values = set(str(v).strip() for v in row.dropna())
                if required_excel_cols.issubset(row_values):
                    header_row_index = i
                    break
            
            if header_row_index == -1:
                self.log("自动检测标题行失败，请求用户手动输入...")
                header_row_num = simpledialog.askinteger("设置标题行", "自动检测标题行失败，请手动输入Excel中列标题所在行号（从1开始）：", minvalue=1, maxvalue=100)
                if not header_row_num: return None
                header_row_index = header_row_num - 1

            df = pd.read_excel(self.excel_path, sheet_name=0, header=header_row_index)
            df.columns = [str(col).strip() for col in df.columns]

            missing_cols = required_excel_cols - set(df.columns)
            if missing_cols:
                messagebox.showerror("列名缺失", f"Excel中缺少以下必要列: {', '.join(missing_cols)}")
                return None
            
            reverse_mapping = {v: k for k, v in CONFIG['column_mapping'].items()}
            df.rename(columns=reverse_mapping, inplace=True)
            
            return df
        except Exception as e:
            messagebox.showerror("Excel读取失败", f"无法读取或解析Excel文件：\n{e}")
            return None

    def determine_dominant_month(self, df):
        """分析出院日期，让用户选择主导月份"""
        # 使用内部标准字段名 'discharge_date'
        df['discharge_month'] = pd.to_datetime(df['discharge_date'], errors='coerce').dt.strftime('%Y-%m')
        
        month_counts = df['discharge_month'].value_counts().to_dict()
        month_counts.pop(None, None)

        if not month_counts:
            date_col_excel_name = CONFIG['column_mapping']['discharge_date']
            messagebox.showerror("无有效日期", f"在“{date_col_excel_name}”列中未找到任何有效的日期。")
            return None
        
        if len(month_counts) == 1:
            return list(month_counts.keys())[0]

        options = [f"{month} ({count}例)" for month, count in month_counts.items()]
        choice = simpledialog.askstring("选择主导月份", "发现多个出院月份，请选择一个作为文件命名和标题的主导月份：\n\n" + "\n".join(options))
        return choice.split(" ")[0] if choice else None

    def filter_day_surgery_patients(self, df):
        """根据住院天数筛选患者"""
        # 使用内部标准字段名 'hospital_days'
        df['hospital_days'] = pd.to_numeric(df['hospital_days'], errors='coerce')
        return df[df['hospital_days'] <= CONFIG['day_surgery_max_days']].copy()

    def generate_single_document(self, row_data, patient_year_month):
        """为单条记录生成Word文档"""
        replacements = {}
        
        # 1. 处理特殊生成值
        replacements["{{患者出院年月}}"] = patient_year_month
        discharge_date_str = excel_date_to_str(getattr(row_data, 'discharge_date', ''))
        replacements["{{随访日期}}"] = get_day_after_discharge(discharge_date_str)

        # 2. 循环处理模板中定义的占位符
        for placeholder, internal_key in CONFIG['template_placeholders'].items():
            raw_value = getattr(row_data, internal_key, "")
            
            if "date" in internal_key:
                replacements[placeholder] = excel_date_to_str(raw_value)
            elif internal_key == "bed_number":
                replacements[placeholder] = str(raw_value) if pd.notna(raw_value) and str(raw_value).strip() else "（手动填写）"
            else:
                replacements[placeholder] = str(raw_value) if pd.notna(raw_value) else ""

        # 3. 生成文件名
        patient_name = replacements.get("{{姓名}}", "未知姓名")
        department = replacements.get("{{科室}}", "未知科室")
        follow_up_date_for_filename = replacements.get("{{随访日期}}", "未知日期")
        
        filename = f"{patient_year_month}_{department}_日间手术随访_{follow_up_date_for_filename}_{patient_name}.docx"
        filename = "".join(c for c in filename if c not in r'\/:*?"<>|')
        output_path = os.path.join(self.output_dir, filename)

        # 4. 执行替换并保存
        doc = Document(self.template_path)
        self.perform_replacements_in_doc(doc, replacements)
        doc.save(output_path)
        self.log(f"已生成: {filename}")

    def perform_replacements_in_doc(self, element, replacements):
        """在文档元素（文档、单元格）中递归执行文本替换"""
        for p in element.paragraphs:
            for old, new in replacements.items():
                if old in p.text:
                    inline = p.runs
                    for i in range(len(inline)):
                        if old in inline[i].text:
                            text = inline[i].text.replace(old, new)
                            inline[i].text = text
        
        if hasattr(element, 'tables'):
            for table in element.tables:
                for row in table.rows:
                    for cell in row.cells:
                        self.perform_replacements_in_doc(cell, replacements)

# ======================== GUI界面类 ========================
class App:
    def __init__(self, root):
        self.root = root
        self.setup_window()
        self.create_widgets()

    def setup_window(self):
        self.root.title(CONFIG['app_title'])
        self.root.geometry("700x550")
        self.root.minsize(600, 500)
        if sys.platform == 'win32':
            from ctypes import windll
            try: windll.shcore.SetProcessDpiAwareness(1)
            except Exception: pass

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        file_frame = ttk.LabelFrame(main_frame, text="步骤1: 选择文件和路径", padding="10")
        file_frame.pack(fill=tk.X, expand=True)
        
        self.excel_path_var = tk.StringVar()
        self.template_path_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()

        self.create_file_selector(file_frame, "Excel源文件:", self.excel_path_var, self.select_excel_file)
        self.create_file_selector(file_frame, "Word模板:", self.template_path_var, self.select_template_file)
        self.create_file_selector(file_frame, "输出文件夹:", self.output_dir_var, self.select_output_dir)

        control_frame = ttk.LabelFrame(main_frame, text="步骤2: 开始生成", padding="10")
        control_frame.pack(fill=tk.X, expand=True, pady=10)
        
        self.start_button = ttk.Button(control_frame, text="开始生成", command=self.start_generation, style="Accent.TButton")
        self.start_button.pack(pady=5)
        
        progress_frame = ttk.LabelFrame(main_frame, text="处理进度与日志", padding="10")
        progress_frame.pack(fill=tk.BOTH, expand=True)

        self.progress_bar = ttk.Progressbar(progress_frame, orient='horizontal', mode='determinate')
        self.progress_bar.pack(fill=tk.X, expand=True, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(progress_frame, height=10, state='disabled', font=("微软雅黑", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def create_file_selector(self, parent, label_text, string_var, command):
        row_frame = ttk.Frame(parent)
        row_frame.pack(fill=tk.X, expand=True, pady=2)
        ttk.Label(row_frame, text=label_text, width=12).pack(side=tk.LEFT)
        ttk.Entry(row_frame, textvariable=string_var, state='readonly').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(row_frame, text="浏览...", command=command).pack(side=tk.RIGHT)

    def select_excel_file(self):
        path = filedialog.askopenfilename(title="选择出院患者记录单", filetypes=[("Excel文件", "*.xlsx *.xls")])
        if path: self.excel_path_var.set(path)

    def select_template_file(self):
        path = filedialog.askopenfilename(title="选择随访表模板", filetypes=[("Word模板", "*.docx")])
        if path: self.template_path_var.set(path)

    def select_output_dir(self):
        path = filedialog.askdirectory(title="选择保存位置")
        if path: self.output_dir_var.set(path)

    def log_message(self, msg):
        def append():
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {msg}\n")
            self.log_text.config(state='disabled')
            self.log_text.see(tk.END)
        self.root.after(0, append)

    def update_progress(self, value):
        self.root.after(0, lambda: self.progress_bar.config(value=value))

    def start_generation(self):
        if not all([self.excel_path_var.get(), self.template_path_var.get(), self.output_dir_var.get()]):
            messagebox.showwarning("信息不全", "请先选择好Excel源文件、Word模板和输出文件夹。")
            return

        self.start_button.config(state='disabled')
        self.progress_bar['value'] = 0
        self.log_text.config(state='normal'); self.log_text.delete('1.0', tk.END); self.log_text.config(state='disabled')

        generator = DocumentGenerator(self.excel_path_var.get(), self.template_path_var.get(), self.output_dir_var.get(), self)
        threading.Thread(target=generator.run, daemon=True).start()

    def generation_finished(self):
        self.root.after(0, lambda: self.start_button.config(state='normal'))


# ======================== 主程序入口 ========================
if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style(root)
    if "clam" in style.theme_names():
        style.theme_use("clam")
        style.configure("Accent.TButton", foreground="white", background="#0078D7")
    
    app = App(root)
    root.mainloop()
