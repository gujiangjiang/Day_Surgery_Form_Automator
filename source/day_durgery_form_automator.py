# -*- coding: utf-8 -*-
"""
日间手术随访表生成系统 (功能增强版)

版本 V5.1 (最终优化版) 更新内容:
- [启动画面修正] 修复了使用Nuitka打包后，启动画面在某些情况下无法自动关闭的问题。通过after_idle确保关闭时机正确。

版本 V5.0 更新内容:
- [UI优化] 程序启动时窗口会自动居中显示，提升用户体验。

版本 V4.9 更新内容:
- [逻辑修正] 优化了“床号匹配失败”的提醒机制，现在只提醒符合日间手术条件的患者，过滤了无关人员的信息，提醒更精准。

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
    import tkinter as tk_error
    root_err = tk_error.Tk()
    root_err.withdraw()
    messagebox.showerror(
        "依赖缺失",
        "缺少必要的库 (pandas, python-docx)。\n"
        "请在命令行运行 'pip install pandas python-docx' 来安装。"
    )
    sys.exit(1)

# 尝试导入Nuitka启动画面模块
try:
    import nuitka_splashscreen_python
    splash_active = True
except ImportError:
    splash_active = False

# ======================== 全局配置 ========================
CONFIG = {
    "app_title": "日间手术随访表生成系统 V5.1",
    "day_surgery_max_days": 2,
    "follow_up_days": 7,  # 随访发生于出院后的天数
    "unknown_bed_placeholder": "（手动填写）", # Word内容中的床号未知占位符
    "unknown_bed_filename_suffix": "（床号未知）",   # 文件名中的床号未知后缀
    "column_mapping": {
        "name": "姓名", "department": "出院科室", "hospital_id": "住院号",
        "discharge_date": "出院日期", "hospital_days": "住院天数", "gender": "性别",
        "age": "年龄", "bed_number": "床号", "admission_date": "入院日期",
        "surgery_date": "手术日期", "surgery_name": "手术名称", 
        "diagnosis": "最后诊断1",
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

def normalize_text(text):
    if pd.isna(text):
        return ""
    return str(text).replace('\xa0', ' ').strip()

# ======================== 核心逻辑类 ========================
class DocumentGenerator:
    def __init__(self, excel_path, surgery_query_paths, template_path, output_dir, app_instance):
        self.excel_path = excel_path
        self.surgery_query_paths = surgery_query_paths
        self.template_path = template_path
        self.output_dir = output_dir
        self.app = app_instance
        self.bed_number_lookup = {}

    def log(self, message, level="info"):
        self.app.log_message(message, level)
    def update_progress(self, value): self.app.update_progress(value)

    def run(self):
        try:
            self.prepare_bed_number_lookup()
            self.log("开始读取出院患者列表Excel文件...")
            df = self.read_and_prepare_patient_excel()
            if df is None: return
            
            df = self.merge_bed_numbers(df)
            day_surgery_df = self.filter_day_surgery_patients(df)
            
            if day_surgery_df.empty:
                msg = f"错误：未找到住院天数 <= {CONFIG['day_surgery_max_days']} 天的记录。"
                self.log(msg, "error")
                messagebox.showerror("无数据", msg)
                return

            unmatched_day_surgery_patients = []
            for row in day_surgery_df.itertuples():
                final_bed_number = getattr(row, 'final_bed_number', None)
                is_unknown = not (pd.notna(final_bed_number) and str(final_bed_number).strip())
                if is_unknown:
                    name = getattr(row, 'name', '未知姓名')
                    hospital_id = getattr(row, 'hospital_id', '未知住院号')
                    unmatched_day_surgery_patients.append(f"{name} (住院号: {hospital_id})")

            total_rows = len(day_surgery_df)
            self.log(f"共找到 {total_rows} 条符合条件的记录，开始生成文档...")
            success_count = 0
            for index, row in enumerate(day_surgery_df.itertuples()):
                try:
                    self.generate_single_document(row)
                    success_count += 1
                except Exception as e:
                    self.log(f"处理行 {getattr(row, 'Index', 'N/A')} 时发生错误: {e}", "error")
                self.update_progress((index + 1) / total_rows * 100)
            
            self.log("="*30)
            self.log(f"处理完成！成功生成 {success_count} 份文档。")
            
            if unmatched_day_surgery_patients:
                summary_message = f"注意：有 {len(unmatched_day_surgery_patients)} 位符合条件的日间手术患者未能匹配到床号：\n\n" + "\n".join(unmatched_day_surgery_patients)
                self.log("="*30, "warning")
                self.log("以下日间手术患者未能匹配到床号:", "warning")
                for patient_info in unmatched_day_surgery_patients:
                    self.log(f"- {patient_info}", "warning")
                messagebox.showwarning("匹配提醒", summary_message)
            
            final_message = f"成功生成 {success_count} 份随访表。\n" \
                          f"文件保存在: {self.output_dir}"
            messagebox.showinfo("完成", final_message)
        except Exception as e:
            self.log(f"发生严重错误: {e}", "error")
            messagebox.showerror("严重错误", f"处理过程中发生严重错误：\n{e}")
        finally:
            self.app.generation_finished()

    def prepare_bed_number_lookup(self):
        self.log("开始处理手术查询文件...")
        if not self.surgery_query_paths:
            self.log("未选择任何手术查询文件，跳过床号补充步骤。", "warning")
            return
        for file_path in self.surgery_query_paths:
            self.log(f"正在读取文件: {os.path.basename(file_path)}", "info")
            try:
                required_cols = ["住院号", "姓名", "床号"]
                df_surgery_full = pd.read_excel(file_path, header=None, dtype=str)
                header_row_index = -1
                for i, row in df_surgery_full.iterrows():
                    row_values = {str(v).strip() for v in row.dropna()}
                    if set(required_cols).issubset(row_values):
                        header_row_index = i
                        self.log(f"在文件 '{os.path.basename(file_path)}' 中自动检测到标题行位于第 {i+1} 行。")
                        break
                if header_row_index == -1:
                    self.log(f"警告：在文件 '{os.path.basename(file_path)}' 中未能找到必需列({', '.join(required_cols)})。已跳过此文件。", "warning")
                    continue
                df_surgery = pd.read_excel(file_path, header=header_row_index, dtype=str)
                df_surgery.columns = [str(col).strip() for col in df_surgery.columns]
                df_surgery.dropna(subset=required_cols, inplace=True)
                for _, row in df_surgery.iterrows():
                    name = normalize_text(row.get("姓名"))
                    h_id_text = normalize_text(row.get("住院号"))
                    h_id = h_id_text.lstrip('0') if h_id_text != '0' else '0'
                    bed_number = normalize_text(row.get("床号"))
                    if h_id and name:
                        key = (h_id, name)
                        self.bed_number_lookup[key] = bed_number
            except Exception as e:
                self.log(f"读取文件 '{os.path.basename(file_path)}' 时出错: {e}。已跳过此文件。", "error")
        self.log(f"所有手术查询文件处理完毕，共加载了 {len(self.bed_number_lookup)} 条有效的床号记录。")

    def read_and_prepare_patient_excel(self):
        try:
            df_full = pd.read_excel(self.excel_path, sheet_name=0, header=None, dtype=str)
            required_excel_cols = {CONFIG['column_mapping'][key] for key in CONFIG['required_internal_keys']}
            header_row_index = -1
            for i, row in df_full.iterrows():
                if required_excel_cols.issubset(set(str(v).strip() for v in row.dropna())):
                    header_row_index = i
                    self.log(f"在出院患者列表中自动检测到标题行位于第 {i+1} 行。")
                    break
            if header_row_index == -1:
                self.log("自动检测标题行失败，请求用户手动输入...", "error")
                header_row_num = simpledialog.askinteger("设置标题行", "自动检测标题行失败，请手动输入Excel中列标题所在行号（从1开始）：", minvalue=1, maxvalue=100)
                if not header_row_num: return None
                header_row_index = header_row_num - 1
            df = pd.read_excel(self.excel_path, sheet_name=0, header=header_row_index, dtype=str)
            df.columns = [str(col).strip() for col in df.columns]
            if required_excel_cols - set(df.columns):
                messagebox.showerror("列名缺失", f"Excel中缺少以下必要列: {', '.join(required_excel_cols - set(df.columns))}")
                return None
            df.rename(columns={v: k for k, v in CONFIG['column_mapping'].items()}, inplace=True)
            if 'bed_number' not in df.columns:
                self.log("警告：主Excel文件中未找到“床号”列。将尝试从手术查询文件补充。", "warning")
                df['bed_number'] = None
            return df
        except Exception as e:
            messagebox.showerror("Excel读取失败", f"无法读取或解析出院患者列表文件：\n{e}")
            return None

    def merge_bed_numbers(self, df):
        if not self.bed_number_lookup:
            self.log("床号查找表为空，跳过合并步骤。", "warning")
            df['final_bed_number'] = df.get('bed_number')
            return df
        self.log("正在为患者匹配床号...")
        def find_bed_number(row):
            original_bed_number = normalize_text(row.get('bed_number'))
            if original_bed_number:
                return original_bed_number
            name_raw = row.get('name', "")
            h_id_raw = row.get('hospital_id', "")
            name = normalize_text(name_raw)
            h_id_text = normalize_text(h_id_raw)
            h_id = h_id_text.lstrip('0') if h_id_text != '0' else '0'
            lookup_key = (h_id, name)
            found_bed_number = self.bed_number_lookup.get(lookup_key)
            if found_bed_number:
                self.log(f"为患者 '{name_raw}' (住院号: {h_id_raw}) 成功匹配到床号: {found_bed_number}", "info")
            else:
                self.log(f"患者 '{name_raw}' (住院号: {h_id_raw}) 匹配失败。程序尝试使用的标准化键为: ('{h_id}', '{name}')", "warning")
            return found_bed_number
        df['final_bed_number'] = df.apply(find_bed_number, axis=1)
        self.log("床号匹配完成。")
        return df

    def filter_day_surgery_patients(self, df):
        df['hospital_days'] = pd.to_numeric(df['hospital_days'], errors='coerce')
        return df[df['hospital_days'] <= CONFIG['day_surgery_max_days']].copy()

    def generate_single_document(self, row_data):
        replacements = {}
        
        discharge_date_str = excel_date_to_str(getattr(row_data, 'discharge_date', ''))
        if discharge_date_str:
            try:
                dt_discharge = datetime.strptime(discharge_date_str, "%Y-%m-%d")
                patient_year_month = dt_discharge.strftime("%Y年%m月")
            except ValueError:
                patient_year_month = "未知年月"
        else:
            patient_year_month = "未知年月"

        replacements["{{患者出院年月}}"] = patient_year_month
        replacements["{{随访日期}}"] = get_day_after_discharge(discharge_date_str, days=CONFIG["follow_up_days"])
        
        final_bed_number = getattr(row_data, 'final_bed_number', None)
        bed_number_is_unknown = not (pd.notna(final_bed_number) and str(final_bed_number).strip())
        for placeholder, key in CONFIG['template_placeholders'].items():
            if key == "bed_number":
                replacements[placeholder] = str(final_bed_number) if not bed_number_is_unknown else CONFIG["unknown_bed_placeholder"]
                continue
            raw_value = getattr(row_data, key, "")
            if "date" in key:
                replacements[placeholder] = excel_date_to_str(raw_value)
            else:
                replacements[placeholder] = str(raw_value) if pd.notna(raw_value) else ""
        patient_name = replacements.get("{{姓名}}", "未知姓名")
        department = replacements.get("{{科室}}", "未知科室")
        
        base_filename = f"{patient_year_month}_{department}_日间手术随访_{replacements['{{随访日期}}']}_{patient_name}"
        if bed_number_is_unknown:
            filename = f"{base_filename}{CONFIG['unknown_bed_filename_suffix']}.docx"
        else:
            filename = f"{base_filename}.docx"
        filename = "".join(c for c in filename if c not in r'\/:*?"<>|')
        doc = Document(self.template_path)
        self.perform_replacements(doc, replacements)
        doc.save(os.path.join(self.output_dir, filename))
        self.log(f"已生成: {filename}")

    def _replace_in_element(self, element, replacements):
        """辅助函数：在给定的元素（如文档、页眉、单元格）中递归执行替换"""
        for p in element.paragraphs:
            for old, new in replacements.items():
                if old in p.text:
                    p.text = p.text.replace(old, str(new)) # 确保new是字符串
        for table in element.tables:
            for row in table.rows:
                for cell in row.cells:
                    self._replace_in_element(cell, replacements)

    def perform_replacements(self, doc, replacements):
        """在文档的各个部分（正文、页眉）执行文本替换"""
        self._replace_in_element(doc, replacements)
        for section in doc.sections:
            self._replace_in_element(section.header, replacements)

# ======================== GUI界面类 ========================
class App:
    def __init__(self, root):
        self.root = root
        self.surgery_query_files = []
        
        self.scaling_factor = self._get_scaling_factor()
        
        self.setup_fonts()
        self.setup_window()
        self.create_widgets()

    def _get_scaling_factor(self):
        """获取屏幕缩放比例"""
        try:
            dpi = self.root.winfo_fpixels('1i')
            scaling = dpi / 96.0
            if scaling < 0.75: return 1.0
            return scaling
        except Exception:
            return 1.0

    def setup_fonts(self):
        """使用固定的字体大小，不再手动缩放"""
        self.font_normal = ("微软雅黑", 9)
        self.font_bold = ("微软雅黑", 10, "bold")
        self.font_title = ("微软雅黑", 20, "bold")
        self.font_subtitle = ("微软雅黑", 16, "bold")
        self.font_button = ("微软雅黑", 12, "bold")
        self.font_disclaimer = ("微软雅黑", 10)

    def setup_window(self):
        """根据缩放比例设置窗口大小并居中"""
        self.root.title(CONFIG['app_title'])
        s = self.scaling_factor
        width = int(685 * s)
        height = int(535 * s)
        
        # 获取屏幕尺寸
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # 计算居中位置
        pos_x = (screen_width // 2) - (width // 2)
        pos_y = (screen_height // 2) - (height // 2)
        
        # 设置窗口大小和位置
        self.root.geometry(f"{width}x{height}+{pos_x}+{pos_y}") 
        self.root.resizable(False, False)

    def create_widgets(self):
        style = ttk.Style(self.root)
        if "clam" in style.theme_names(): style.theme_use("clam")
        default_bg = style.lookup('TFrame', 'background')
        self.root.configure(bg=default_bg)
        self.log_text_tags = {"warning": {"foreground": "orange"}, "error": {"foreground": "red"}}

        # 底部信息栏
        bottom_frame = tk.Frame(self.root, bg=default_bg)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=2)
        tk.Label(bottom_frame, text=f"编程日期：{datetime.now().strftime('%Y年%m月%d日')}", font=self.font_normal, fg="#666666", bg=default_bg).pack(side=tk.LEFT)
        tk.Label(bottom_frame, text="作者：顾江江", font=self.font_normal, fg="#666666", bg=default_bg).pack(side=tk.RIGHT)
        
        disclaimer_frame = tk.Frame(self.root, pady=2, bg=default_bg)
        disclaimer_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10)
        tk.Label(disclaimer_frame, text="本工具仅供骨科内部测试，请勿外传", font=self.font_disclaimer, fg="red", bg=default_bg).pack()

        # 顶部全局标题栏
        top_title_frame = tk.Frame(self.root, bg=default_bg)
        top_title_frame.pack(side=tk.TOP, fill=tk.X, pady=(10, 5))
        tk.Label(top_title_frame, text="丹阳市人民医院", font=self.font_subtitle, fg="#0066cc", bg=default_bg).pack()
        tk.Label(top_title_frame, text="日间手术随访表生成系统", font=self.font_title, bg=default_bg).pack()

        # 使用PanedWindow实现可拖拽的左右布局
        main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 5))

        left_panel = ttk.Frame(main_pane, padding=(5, 5, 5, 5))
        main_pane.add(left_panel)

        right_panel = ttk.Frame(main_pane, padding=(5, 5, 5, 5))
        main_pane.add(right_panel)

        s = self.scaling_factor
        sash_default = int(420 * s)
        sash_min = int(320 * s)
        sash_max = int(520 * s)

        def set_initial_sash(event):
            main_pane.sashpos(0, sash_default)
            main_pane.unbind("<Configure>")
        
        def limit_sash_movement(event):
            if event.x < sash_min:
                main_pane.sashpos(0, sash_min)
                return "break"
            if event.x > sash_max:
                main_pane.sashpos(0, sash_max)
                return "break"

        main_pane.bind("<Configure>", set_initial_sash)
        main_pane.bind("<B1-Motion>", limit_sash_movement)

        # --- 将控件填充到左侧面板 ---
        file_frame = ttk.LabelFrame(left_panel, text="步骤1: 选择文件和路径", padding=5)
        file_frame.pack(fill=tk.BOTH, expand=True) 
        
        self.excel_path_var = tk.StringVar()
        self.template_path_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        
        self.create_file_selector(file_frame, "出院患者列表:", self.excel_path_var, self.select_excel_file)
        
        surgery_frame = ttk.LabelFrame(file_frame, text="手术查询文件 (可多选)", padding=5)
        surgery_frame.pack(fill=tk.X, expand=True, pady=3)
        
        self.surgery_listbox = tk.Listbox(surgery_frame, height=5, font=self.font_normal)
        self.surgery_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        
        surgery_buttons_frame = ttk.Frame(surgery_frame)
        surgery_buttons_frame.pack(side=tk.LEFT, fill=tk.Y, anchor='n')
        ttk.Button(surgery_buttons_frame, text="添加文件", command=self.select_surgery_query_files, width=8).pack(fill=tk.X, pady=1)
        ttk.Button(surgery_buttons_frame, text="清空列表", command=self.clear_surgery_query_files, width=8).pack(fill=tk.X, pady=1)
        
        self.create_file_selector(file_frame, "Word模板:", self.template_path_var, self.select_template_file)
        self.create_file_selector(file_frame, "输出文件夹:", self.output_dir_var, self.select_output_dir)
        
        left_bottom_container = ttk.Frame(left_panel)
        left_bottom_container.pack(fill=tk.X, pady=(5,0))

        control_frame = ttk.LabelFrame(left_bottom_container, text="步骤2: 开始生成", padding=10)
        control_frame.pack(fill=tk.X)
        style.configure("Accent.TButton", foreground="white", background="#0078D7", font=self.font_button)
        self.start_button = ttk.Button(control_frame, text="开始生成", command=self.start_generation, style="Accent.TButton")
        self.start_button.pack(pady=5, ipady=5, ipadx=20)

        # --- 将控件填充到右侧面板 ---
        progress_frame = ttk.LabelFrame(right_panel, text="处理进度与日志", padding=10)
        progress_frame.pack(fill=tk.BOTH, expand=True)
        
        self.progress_bar = ttk.Progressbar(progress_frame, orient='horizontal', mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))
        
        self.log_text = scrolledtext.ScrolledText(progress_frame, height=5, state='disabled', font=self.font_normal)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        for tag, config in self.log_text_tags.items():
            self.log_text.tag_config(tag, **config)

    def create_file_selector(self, parent, label_text, string_var, command):
        row_frame = ttk.Frame(parent)
        row_frame.pack(fill=tk.X, expand=True, pady=1)
        ttk.Label(row_frame, text=label_text, width=12, font=self.font_normal).pack(side=tk.LEFT)
        ttk.Entry(row_frame, textvariable=string_var, state='readonly', font=self.font_normal).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(row_frame, text="浏览...", command=command, width=8).pack(side=tk.RIGHT)

    def select_excel_file(self):
        path = filedialog.askopenfilename(title="选择出院患者记录单", filetypes=[("Excel文件", "*.xlsx *.xls")])
        if path: self.excel_path_var.set(path)

    def select_surgery_query_files(self):
        paths = filedialog.askopenfilenames(title="选择一个或多个手术查询文件", filetypes=[("Excel文件", "*.xlsx *.xls")])
        if paths:
            for path in paths:
                if path not in self.surgery_query_files:
                    self.surgery_query_files.append(path)
                    self.surgery_listbox.insert(tk.END, os.path.basename(path))

    def clear_surgery_query_files(self):
        self.surgery_query_files.clear()
        self.surgery_listbox.delete(0, tk.END)

    def select_template_file(self):
        path = filedialog.askopenfilename(title="选择随访表模板", filetypes=[("Word模板", "*.docx")])
        if path: self.template_path_var.set(path)

    def select_output_dir(self):
        path = filedialog.askdirectory(title="选择保存位置")
        if path: self.output_dir_var.set(path)

    def log_message(self, msg, level="info"):
        def append():
            self.log_text.config(state='normal')
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
            messagebox.showwarning("信息不全", "请先选择好“出院患者列表”、“Word模板”和“输出文件夹”。")
            return
        if not self.surgery_query_files:
            if not messagebox.askyesno("确认操作", "您没有选择任何“手术查询文件”。\n程序将无法补充床号，是否继续？"):
                return
        self.start_button.config(state='disabled')
        self.progress_bar['value'] = 0
        self.log_text.config(state='normal'); self.log_text.delete('1.0', tk.END); self.log_text.config(state='disabled')
        generator = DocumentGenerator(
            excel_path=self.excel_path_var.get(), 
            surgery_query_paths=self.surgery_query_files,
            template_path=self.template_path_var.get(), 
            output_dir=self.output_dir_var.get(), 
            app_instance=self
        )
        threading.Thread(target=generator.run, daemon=True).start()

    def generation_finished(self):
        self.root.after(0, lambda: self.start_button.config(state='normal'))

# ======================== 主程序入口 ========================
if __name__ == "__main__":
    try:
        # 适配高DPI屏幕
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    
    # 如果启动画面模块已加载，标记Python已开始部署
    if splash_active:
        nuitka_splashscreen_python.mark_as_deployed()

    root = tk.Tk()
    app = App(root)

    # V5.1修正：使用 after_idle 确保在主循环开始处理事件后，再关闭启动画面
    # 这是最稳妥的方式，可以避免时序问题
    if splash_active:
        root.after_idle(nuitka_splashscreen_python.close)
        
    root.mainloop()

