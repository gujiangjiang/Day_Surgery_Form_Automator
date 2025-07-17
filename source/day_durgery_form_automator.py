# -*- coding: utf-8 -*-
"""
日间手术随访表生成系统 (最终修复版)

功能：
- 自动从Excel批量生成Word随访表
- GUI界面，操作直观
- 自动检测Excel标题行
- 支持进度条和实时日志
- 可配置关键参数
- 修正高DPI显示，自动适应系统缩放
- 修正占位符替换逻辑，支持页眉替换
- 增加床号缺失检测和提醒功能

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
    # 在GUI启动前显示错误
    import tkinter as tk_error
    root_err = tk_error.Tk()
    root_err.withdraw()
    messagebox.showerror(
        "依赖缺失",
        "缺少必要的库 (pandas, python-docx)。\n"
        "请在命令行运行 'pip install pandas python-docx' 来安装。"
    )
    sys.exit(1)

# ======================== 全局配置 ========================
CONFIG = {
    "app_title": "日间手术随访表生成系统 V2.5",
    "day_surgery_max_days": 2,
    "column_mapping": {
        "name": "姓名", "department": "出院科室", "hospital_id": "住院号",
        "discharge_date": "出院日期", "hospital_days": "住院天数", "gender": "性别",
        "age": "年龄", "bed_number": "床号", "admission_date": "入院日期",
        "surgery_date": "手术日期", "surgery_name": "手术名称", 
        "diagnosis": "最后诊断1", # 关键修复：根据用户反馈修正列名
        "phone": "联系电话", "doctor": "经治医生"
    },
    "required_internal_keys": [
        "name", "department", "hospital_id", "discharge_date", "hospital_days"
    ],
    "template_placeholders": {
        "{{科室}}": "department", "{{姓名}}": "name", "{{性别}}": "gender",
        "{{年龄}}": "age", "{{住院号}}": "hospital_id", "{{床号}}": "bed_number",
        "{{入院日期}}": "admission_date", "{{出院日期}}": "discharge_date",
        "{{手术日期}}": "surgery_date", "{{手术名称}}": "surgery_name",
        "{{出院诊断}}": "diagnosis", "{{联系电话}}": "phone", "{{经治医生}}": "doctor",
    }
}

# ======================== 工具函数 ========================
def excel_date_to_str(excel_date):
    if pd.isna(excel_date) or excel_date in ("", None): return ""
    try: return pd.to_datetime(excel_date).strftime('%Y-%m-%d')
    except (ValueError, TypeError): return str(excel_date).split()[0]

def get_day_after_discharge(discharge_date_str, days=7):
    if not discharge_date_str: return ""
    try:
        base_date = datetime.strptime(discharge_date_str, "%Y-%m-%d")
        return (base_date + timedelta(days=days)).strftime("%Y-%m-%d")
    except (ValueError, TypeError): return ""

# ======================== 核心逻辑类 ========================
class DocumentGenerator:
    def __init__(self, excel_path, template_path, output_dir, app_instance):
        self.excel_path, self.template_path, self.output_dir, self.app = excel_path, template_path, output_dir, app_instance
        self.bed_number_missing = False # 新增：用于标记床号是否缺失

    def log(self, message, level="info"):
        self.app.log_message(message, level)
    def update_progress(self, value): self.app.update_progress(value)

    def run(self):
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
                msg = f"错误：未找到住院天数 <= {CONFIG['day_surgery_max_days']} 天的记录。"
                self.log(msg, "error")
                messagebox.showerror("无数据", msg)
                return

            total_rows = len(day_surgery_df)
            self.log(f"共找到 {total_rows} 条符合条件的记录，开始生成文档...")
            success_count = 0
            for index, row in enumerate(day_surgery_df.itertuples()):
                try:
                    self.generate_single_document(row, patient_year_month)
                    success_count += 1
                except Exception as e:
                    self.log(f"处理行 {getattr(row, 'Index', 'N/A')} 时发生错误: {e}", "error")
                self.update_progress((index + 1) / total_rows * 100)

            self.log("="*30)
            self.log(f"处理完成！成功生成 {success_count} 份文档。")
            
            # 新增：根据床号是否缺失，构建不同的成功消息
            final_message = f"成功生成 {success_count} 份随访表。\n" \
                          f"统一出院年月为: {patient_year_month}\n" \
                          f"文件保存在: {self.output_dir}"
            if self.bed_number_missing:
                final_message += "\n\n重要提醒：\n未在Excel中找到“床号”信息，请手动填写生成的文档！"
            
            messagebox.showinfo("完成", final_message)

        except Exception as e:
            self.log(f"发生严重错误: {e}", "error")
            messagebox.showerror("严重错误", f"处理过程中发生严重错误：\n{e}")
        finally:
            self.app.generation_finished()

    def read_and_prepare_excel(self):
        try:
            df_full = pd.read_excel(self.excel_path, sheet_name=0, header=None, dtype=str)
            required_excel_cols = {CONFIG['column_mapping'][key] for key in CONFIG['required_internal_keys']}
            header_row_index = -1
            for i, row in df_full.iterrows():
                if required_excel_cols.issubset(set(str(v).strip() for v in row.dropna())):
                    header_row_index = i
                    break
            if header_row_index == -1:
                self.log("自动检测标题行失败，请求用户手动输入...", "error")
                header_row_num = simpledialog.askinteger("设置标题行", "自动检测标题行失败，请手动输入Excel中列标题所在行号（从1开始）：", minvalue=1, maxvalue=100)
                if not header_row_num: return None
                header_row_index = header_row_num - 1
            df = pd.read_excel(self.excel_path, sheet_name=0, header=header_row_index)
            df.columns = [str(col).strip() for col in df.columns]
            if required_excel_cols - set(df.columns):
                messagebox.showerror("列名缺失", f"Excel中缺少以下必要列: {', '.join(required_excel_cols - set(df.columns))}")
                return None
            df.rename(columns={v: k for k, v in CONFIG['column_mapping'].items()}, inplace=True)
            
            # 新增：检测床号列是否存在，并记录状态
            if 'bed_number' not in df.columns:
                self.bed_number_missing = True
                self.log("警告：Excel文件中未找到“床号”列。生成文档中的床号需手动填写。", "warning")

            return df
        except Exception as e:
            messagebox.showerror("Excel读取失败", f"无法读取或解析Excel文件：\n{e}")
            return None

    def determine_dominant_month(self, df):
        df['discharge_month'] = pd.to_datetime(df['discharge_date'], errors='coerce').dt.strftime('%Y-%m')
        month_counts = df['discharge_month'].value_counts().to_dict()
        month_counts.pop(None, None)
        if not month_counts:
            messagebox.showerror("无有效日期", f"在“{CONFIG['column_mapping']['discharge_date']}”列中未找到任何有效的日期。")
            return None
        if len(month_counts) == 1: return list(month_counts.keys())[0]
        options = [f"{month} ({count}例)" for month, count in month_counts.items()]
        choice = simpledialog.askstring("选择主导月份", "发现多个出院月份，请选择一个作为文件命名和标题的主导月份：\n\n" + "\n".join(options))
        return choice.split(" ")[0] if choice else None

    def filter_day_surgery_patients(self, df):
        df['hospital_days'] = pd.to_numeric(df['hospital_days'], errors='coerce')
        return df[df['hospital_days'] <= CONFIG['day_surgery_max_days']].copy()

    def generate_single_document(self, row_data, patient_year_month):
        replacements = {}
        replacements["{{患者出院年月}}"] = patient_year_month
        discharge_date_str = excel_date_to_str(getattr(row_data, 'discharge_date', ''))
        replacements["{{随访日期}}"] = get_day_after_discharge(discharge_date_str)
        for placeholder, key in CONFIG['template_placeholders'].items():
            raw_value = getattr(row_data, key, "")
            if "date" in key:
                replacements[placeholder] = excel_date_to_str(raw_value)
            elif key == "bed_number":
                # 如果床号列本身就不存在，则强制为手动填写
                if self.bed_number_missing:
                    replacements[placeholder] = "（手动填写）"
                else:
                    replacements[placeholder] = str(raw_value) if pd.notna(raw_value) and str(raw_value).strip() else "（手动填写）"
            else:
                replacements[placeholder] = str(raw_value) if pd.notna(raw_value) else ""
        patient_name = replacements.get("{{姓名}}", "未知姓名")
        department = replacements.get("{{科室}}", "未知科室")
        filename = f"{patient_year_month}_{department}_日间手术随访_{replacements['{{随访日期}}']}_{patient_name}.docx"
        filename = "".join(c for c in filename if c not in r'\/:*?"<>|')
        doc = Document(self.template_path)
        self.perform_replacements(doc, replacements)
        doc.save(os.path.join(self.output_dir, filename))
        self.log(f"已生成: {filename}")

    def perform_replacements(self, doc, replacements):
        for p in doc.paragraphs:
            for old, new in replacements.items():
                if old in p.text:
                    p.text = p.text.replace(old, new)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for old, new in replacements.items():
                            if old in p.text:
                                p.text = p.text.replace(old, new)
        for section in doc.sections:
            header = section.header
            for p in header.paragraphs:
                for old, new in replacements.items():
                    if old in p.text:
                        p.text = p.text.replace(old, new)
            for table in header.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            for old, new in replacements.items():
                                if old in p.text:
                                    p.text = p.text.replace(old, new)

# ======================== GUI界面类 ========================
class App:
    def __init__(self, root):
        self.root = root
        self.setup_fonts()
        self.setup_window()
        self.create_widgets()

    def setup_fonts(self):
        self.font_normal = ("微软雅黑", 9)
        self.font_bold = ("微软雅黑", 10, "bold")
        self.font_title = ("微软雅黑", 20, "bold")
        self.font_subtitle = ("微软雅黑", 16, "bold")
        self.font_button = ("微软雅黑", 12, "bold")
        self.font_disclaimer = ("微软雅黑", 10)

    def setup_window(self):
        self.root.title(CONFIG['app_title'])
        self.root.geometry("1000x950")
        self.root.resizable(False, False)

    def create_widgets(self):
        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        default_bg = style.lookup('TFrame', 'background')
        self.root.configure(bg=default_bg)

        # 配置日志颜色
        self.log_text_tags = {
            "warning": {"foreground": "orange"},
            "error": {"foreground": "red"}
        }

        bottom_frame = tk.Frame(self.root, bg=default_bg)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)
        tk.Label(bottom_frame, text=f"编程日期：{datetime.now().strftime('%Y年%m月%d日')}", font=self.font_normal, fg="#666666", bg=default_bg).pack(side=tk.LEFT)
        tk.Label(bottom_frame, text="作者：顾江江", font=self.font_normal, fg="#666666", bg=default_bg).pack(side=tk.RIGHT)

        disclaimer_frame = tk.Frame(self.root, pady=5, bg=default_bg)
        disclaimer_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10)
        tk.Label(disclaimer_frame, text="本工具仅供骨科内部测试，请勿外传", font=self.font_disclaimer, fg="red", bg=default_bg).pack()
        
        title_frame = tk.Frame(self.root, bg=default_bg)
        title_frame.pack(pady=(15, 10))
        tk.Label(title_frame, text="丹阳市人民医院", font=self.font_subtitle, fg="#0066cc", bg=default_bg).pack()
        tk.Label(title_frame, text="日间手术随访表生成系统", font=self.font_title, bg=default_bg).pack(pady=(5, 0))

        content_frame = ttk.Frame(self.root, padding="10")
        content_frame.pack(fill=tk.BOTH, expand=True)

        file_frame = ttk.LabelFrame(content_frame, text="步骤1: 选择文件和路径", padding="10")
        file_frame.pack(fill=tk.X, expand=True, pady=5)
        self.excel_path_var, self.template_path_var, self.output_dir_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
        self.create_file_selector(file_frame, "Excel源文件:", self.excel_path_var, self.select_excel_file)
        self.create_file_selector(file_frame, "Word模板:", self.template_path_var, self.select_template_file)
        self.create_file_selector(file_frame, "输出文件夹:", self.output_dir_var, self.select_output_dir)

        control_frame = ttk.LabelFrame(content_frame, text="步骤2: 开始生成", padding="10")
        control_frame.pack(fill=tk.X, expand=True, pady=10)
        style.configure("Accent.TButton", foreground="white", background="#0078D7", font=self.font_button)
        self.start_button = ttk.Button(control_frame, text="开始生成", command=self.start_generation, style="Accent.TButton")
        self.start_button.pack(pady=5, ipady=5, ipadx=20)

        progress_frame = ttk.LabelFrame(content_frame, text="处理进度与日志", padding="10")
        progress_frame.pack(fill=tk.BOTH, expand=True)
        self.progress_bar = ttk.Progressbar(progress_frame, orient='horizontal', mode='determinate')
        self.progress_bar.pack(fill=tk.X, expand=True, pady=5)
        self.log_text = scrolledtext.ScrolledText(progress_frame, height=10, state='disabled', font=self.font_normal)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        # 为不同级别的日志设置tag
        for tag, config in self.log_text_tags.items():
            self.log_text.tag_config(tag, **config)

    def create_file_selector(self, parent, label_text, string_var, command):
        row_frame = ttk.Frame(parent)
        row_frame.pack(fill=tk.X, expand=True, pady=2)
        ttk.Label(row_frame, text=label_text, width=12, font=self.font_normal).pack(side=tk.LEFT)
        ttk.Entry(row_frame, textvariable=string_var, state='readonly', font=self.font_normal).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
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

    def log_message(self, msg, level="info"):
        def append():
            self.log_text.config(state='normal')
            # 根据日志级别使用不同的tag
            tag = self.log_text_tags.get(level)
            if tag:
                self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {msg}\n", level)
            else:
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
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = tk.Tk()
    app = App(root)
    root.mainloop()
