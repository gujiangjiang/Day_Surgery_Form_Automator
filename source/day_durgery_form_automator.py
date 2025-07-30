# -*- coding: utf-8 -*-
"""
日间手术随访表生成系统 (Polars 重构版)

版本 V6.4 (功能修复版) 更新内容:
- [占位符修复] 重写了 Word 文档的文本替换函数 `_replace_in_element`，采用了更健壮的逻辑，确保即使占位符的文本格式不统一（如 `{{科室}}`）也能被成功替换。
- [日志颜色修复] 修正了 `log_message` 函数中应用颜色标签的逻辑，确保警告（黄色）和错误（红色）日志能正确显示颜色。

版本 V6.3 (过滤逻辑修正版) 更新内容:
- [错误修复] 修复了在 prepare_bed_number_lookup 函数中因过滤语法问题导致的崩溃。
- [代码优化] 将过滤逻辑修改为使用标准的 `&` 操作符。

作者：顾江江 (由AI使用 Polars 重构)
"""

import os
import sys
import threading
from datetime import datetime, timedelta, date
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, scrolledtext

# 检查并安装必要的库
try:
    from docx import Document
    import polars as pl
    # xlrd 和 fastexcel 会被 polars[excel] 依赖自动安装
except ImportError:
    import tkinter as tk_error
    root_err = tk_error.Tk()
    root_err.withdraw()
    messagebox.showerror(
        "依赖缺失",
        "缺少必要的库 (polars, python-docx, fastexcel)。\n"
        "请在命令行运行 'pip install polars python-docx fastexcel' 来安装。"
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
    "app_title": "日间手术随访表生成系统 V6.4",
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
def excel_date_to_str(val):
    """
    将来自Polars的各种可能日期类型（日期对象、字符串、数字）统一转换为 YYYY-MM-DD 格式。
    """
    if val is None or val == '': return ""
    # Polars可能已将其解析为date对象
    if isinstance(val, date):
        return val.strftime('%Y-%m-%d')
    # 处理字符串形式的日期或数字
    if isinstance(val, str):
        val = val.strip()
        try:
            # 尝试直接解析标准日期格式
            return datetime.strptime(val.split()[0], '%Y-%m-%d').strftime('%Y-%m-%d')
        except ValueError:
            # 如果解析失败，可能是 "44562.0" 这样的字符串
            try:
                return excel_date_to_str(float(val))
            except (ValueError, TypeError):
                return val # 无法解析，返回原样
    # 处理数字形式的Excel序列日期
    if isinstance(val, (int, float)):
        try:
            # Excel的日期原点是 1899-12-30
            return (datetime(1899, 12, 30) + timedelta(days=val)).strftime('%Y-%m-%d')
        except (TypeError, ValueError):
            return str(val)
    return str(val)


def get_day_after_discharge(discharge_date_str, days=7):
    if not discharge_date_str: return ""
    try:
        base_date = datetime.strptime(discharge_date_str, "%Y-%m-%d")
        return (base_date + timedelta(days=days)).strftime("%Y-%m-%d")
    except (ValueError, TypeError): return ""

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

    def _read_excel_with_header_detection(self, file_path, required_cols):
        """
        使用 Polars 读取 Excel，自动检测标题行，并以字符串形式安全加载。
        新版逻辑：一次性读取，然后在内存中处理，避免使用 skip_rows。
        """
        try:
            # 一次性读取整个工作表，不带标题
            df_full = pl.read_excel(file_path, sheet_id=1, has_header=False)
            # 为避免类型推断错误，立即将所有列转换为字符串
            df_full = df_full.select([pl.all().cast(pl.Utf8, strict=False)])
        except Exception as e:
            self.log(f"读取文件 '{os.path.basename(file_path)}' 失败: {e}", "error")
            return None

        header_row_index = -1
        for i, row in enumerate(df_full.iter_rows()):
            row_values = {str(v).strip() for v in row if v is not None}
            if required_cols.issubset(row_values):
                header_row_index = i
                break
        
        if header_row_index != -1:
            self.log(f"在文件 '{os.path.basename(file_path)}' 中自动检测到标题行位于第 {header_row_index + 1} 行。")
            
            # 提取标题行作为新的列名
            new_columns = [str(col).strip() for col in df_full.row(header_row_index)]
            
            # 提取数据行（标题行之后的所有行）
            df_data = df_full.slice(header_row_index + 1)
            
            # 将数据与新列名结合成最终的 DataFrame
            df_data.columns = new_columns
            
            return df_data
        else:
            return None
            
    def run(self):
        try:
            self.prepare_bed_number_lookup()
            self.log("开始读取出院患者列表Excel文件...")
            df = self.read_and_prepare_patient_excel()
            if df is None: return
            
            df = self.merge_bed_numbers(df)
            day_surgery_df = self.filter_day_surgery_patients(df)
            
            if day_surgery_df.is_empty():
                msg = f"错误：未找到住院天数 <= {CONFIG['day_surgery_max_days']} 天的记录。"
                self.log(msg, "error")
                messagebox.showerror("无数据", msg)
                return

            unmatched_day_surgery_patients = day_surgery_df.filter(
                pl.col('final_bed_number').is_null() | (pl.col('final_bed_number') == "")
            )

            total_rows = len(day_surgery_df)
            self.log(f"共找到 {total_rows} 条符合条件的记录，开始生成文档...")
            success_count = 0
            
            # 使用 iter_rows(named=True) 高效迭代
            for index, row_dict in enumerate(day_surgery_df.iter_rows(named=True)):
                try:
                    self.generate_single_document(row_dict)
                    success_count += 1
                except Exception as e:
                    self.log(f"处理行 {index + 1} (姓名: {row_dict.get('name', 'N/A')}) 时发生错误: {e}", "error")
                self.update_progress((index + 1) / total_rows * 100)
            
            self.log("="*30)
            self.log(f"处理完成！成功生成 {success_count} 份文档。")
            
            if not unmatched_day_surgery_patients.is_empty():
                unmatched_list = [
                    f"{row['name']} (住院号: {row['hospital_id']})" 
                    for row in unmatched_day_surgery_patients.select(['name', 'hospital_id']).to_dicts()
                ]
                summary_message = f"注意：有 {len(unmatched_list)} 位符合条件的日间手术患者未能匹配到床号：\n\n" + "\n".join(unmatched_list)
                self.log("="*30, "warning")
                self.log("以下日间手术患者未能匹配到床号:", "warning")
                for patient_info in unmatched_list:
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
            
        required_cols = {"住院号", "姓名", "床号"}
        for file_path in self.surgery_query_paths:
            self.log(f"正在读取文件: {os.path.basename(file_path)}", "info")
            df_surgery = self._read_excel_with_header_detection(file_path, required_cols)
            
            if df_surgery is None:
                self.log(f"警告：在文件 '{os.path.basename(file_path)}' 中未能找到必需列({', '.join(required_cols)})。已跳过此文件。", "warning")
                continue

            # 确保必需列存在
            if not required_cols.issubset(df_surgery.columns):
                self.log(f"警告：文件 '{os.path.basename(file_path)}' 标题行检测后仍缺少必需列。已跳过。", "warning")
                continue

            # 筛选、清洗并填充查找字典
            df_surgery = df_surgery.select(["住院号", "姓名", "床号"]) \
                                   .drop_nulls() \
                                   .filter(
                                       (pl.col("住院号").str.strip_chars() != "") &
                                       (pl.col("姓名").str.strip_chars() != "") &
                                       (pl.col("床号").str.strip_chars() != "")
                                   )

            for row in df_surgery.iter_rows(named=True):
                name = row["姓名"].strip()
                h_id_text = row["住院号"].strip()
                h_id = h_id_text.lstrip('0') if h_id_text != '0' else '0'
                bed_number = row["床号"].strip()
                if h_id and name:
                    key = (h_id, name)
                    self.bed_number_lookup[key] = bed_number

        self.log(f"所有手术查询文件处理完毕，共加载了 {len(self.bed_number_lookup)} 条有效的床号记录。")

    def read_and_prepare_patient_excel(self):
        required_excel_cols = {CONFIG['column_mapping'][key] for key in CONFIG['required_internal_keys']}
        df = self._read_excel_with_header_detection(self.excel_path, required_excel_cols)

        if df is None:
            self.log("自动检测标题行失败，请求用户手动输入...", "error")
            header_row_num = simpledialog.askinteger("设置标题行", "自动检测标题行失败，请手动输入Excel中列标题所在行号（从1开始）：", minvalue=1, maxvalue=100)
            if not header_row_num: return None
            header_row_index = header_row_num - 1
            
            # 手动指定标题行时的读取逻辑
            try:
                df_full = pl.read_excel(self.excel_path, sheet_id=1, has_header=False)
                new_columns = [str(col).strip() for col in df_full.row(header_row_index)]
                df = df_full.slice(header_row_index + 1)
                df.columns = new_columns
            except Exception as e:
                self.log(f"根据手动指定的行号 {header_row_num} 读取Excel失败: {e}", "error")
                messagebox.showerror("读取失败", f"无法根据指定的行号 {header_row_num} 读取文件。")
                return None

        if required_excel_cols - set(df.columns):
            messagebox.showerror("列名缺失", f"Excel中缺少以下必要列: {', '.join(required_excel_cols - set(df.columns))}")
            return None
        
        # 重命名列
        df = df.rename({v: k for k, v in CONFIG['column_mapping'].items() if v in df.columns})
        
        # 如果缺少可选的 'bed_number' 列，则添加一个空列
        if 'bed_number' not in df.columns:
            self.log("警告：主Excel文件中未找到“床号”列。将尝试从手术查询文件补充。", "warning")
            df = df.with_columns(pl.lit(None, dtype=pl.Utf8).alias("bed_number"))
            
        return df

    def merge_bed_numbers(self, df):
        if not self.bed_number_lookup:
            self.log("床号查找表为空，跳过合并步骤。", "warning")
            return df.with_columns(pl.col('bed_number').alias('final_bed_number'))

        self.log("正在为患者匹配床号...")
        
        final_bed_numbers = []
        # 为了保留详细的日志，此处采用迭代方式。对于GUI应用，性能影响可忽略。
        for row in df.select(['bed_number', 'name', 'hospital_id']).iter_rows(named=True):
            original_bed_number = str(row.get('bed_number') or "").strip()
            
            if original_bed_number:
                final_bed_numbers.append(original_bed_number)
                continue

            name_raw = str(row.get('name', "") or "")
            h_id_raw = str(row.get('hospital_id', "") or "")
            
            name = name_raw.strip()
            h_id_text = h_id_raw.strip()
            h_id = h_id_text.lstrip('0') if h_id_text != '0' else '0'

            lookup_key = (h_id, name)
            found_bed_number = self.bed_number_lookup.get(lookup_key)
            
            if found_bed_number:
                self.log(f"为患者 '{name_raw}' (住院号: {h_id_raw}) 成功匹配到床号: {found_bed_number}", "info")
            else:
                self.log(f"患者 '{name_raw}' (住院号: {h_id_raw}) 匹配失败。程序尝试使用的标准化键为: ('{h_id}', '{name}')", "warning")
            
            final_bed_numbers.append(found_bed_number)
            
        df = df.with_columns(pl.Series("final_bed_number", final_bed_numbers))
        self.log("床号匹配完成。")
        return df

    def filter_day_surgery_patients(self, df):
        # 使用 strict=False 将转换失败的值设为 null，类似 pandas 的 errors='coerce'
        df_with_numeric_days = df.with_columns(
            pl.col("hospital_days").cast(pl.Float64, strict=False)
        )
        return df_with_numeric_days.filter(
            pl.col("hospital_days") <= CONFIG['day_surgery_max_days']
        )

    def generate_single_document(self, row_data): # row_data is now a dict
        replacements = {}
        
        # 日期处理
        discharge_date_str = excel_date_to_str(row_data.get('discharge_date'))
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
        
        # 占位符替换
        final_bed_number = row_data.get('final_bed_number')
        bed_number_is_unknown = not final_bed_number or not str(final_bed_number).strip()

        for placeholder, key in CONFIG['template_placeholders'].items():
            raw_value = row_data.get(key)
            
            if key == "bed_number":
                replacements[placeholder] = str(final_bed_number) if not bed_number_is_unknown else CONFIG["unknown_bed_placeholder"]
                continue

            if "date" in key:
                replacements[placeholder] = excel_date_to_str(raw_value)
            else:
                replacements[placeholder] = str(raw_value) if raw_value is not None else ""

        patient_name = replacements.get("{{姓名}}", "未知姓名")
        department = replacements.get("{{科室}}", "未知科室")
        
        # 文件名生成
        base_filename = f"{patient_year_month}_{department}_日间手术随访_{replacements['{{随访日期}}']}_{patient_name}"
        if bed_number_is_unknown:
            filename = f"{base_filename}{CONFIG['unknown_bed_filename_suffix']}.docx"
        else:
            filename = f"{base_filename}.docx"
        filename = "".join(c for c in filename if c not in r'\/:*?"<>|')
        
        # Word文档生成
        doc = Document(self.template_path)
        self.perform_replacements(doc, replacements)
        doc.save(os.path.join(self.output_dir, filename))
        self.log(f"已生成: {filename}")

    def _replace_in_element(self, element, replacements):
        """
        辅助函数：在给定的元素（如文档、页眉、单元格）中递归执行替换。
        这个版本可以正确处理跨越不同文本格式的占位符。
        """
        for p in element.paragraphs:
            # 将段落内所有部分的文本连接起来，以便查找完整的占位符
            full_text = "".join(run.text for run in p.runs)
            
            found_placeholder = False
            for old in replacements.keys():
                if old in full_text:
                    found_placeholder = True
                    break
            
            if found_placeholder:
                # 在连接后的完整文本上执行所有替换
                for old, new in replacements.items():
                    full_text = full_text.replace(old, str(new))
                
                # 清空段落内原有的所有部分，然后添加一个包含新文本的新部分。
                # 这会保留段落的整体样式，但可能会丢失段落内部的局部格式（如单个词的粗体）。
                # 对于占位符替换场景，这是一个可靠的折中方案。
                style = p.runs[0].style if p.runs else None
                p.clear()
                new_run = p.add_run(full_text)
                if style:
                    new_run.style = style

        for table in element.tables:
            for row in table.rows:
                for cell in row.cells:
                    self._replace_in_element(cell, replacements)

    def perform_replacements(self, doc, replacements):
        """在文档的各个部分（正文、页眉、页脚）执行文本替换"""
        self._replace_in_element(doc, replacements)
        for section in doc.sections:
            self._replace_in_element(section.header, replacements)
            self._replace_in_element(section.footer, replacements)


# ======================== GUI界面类 (部分修改) ========================
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
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        pos_x = (screen_width // 2) - (width // 2)
        pos_y = (screen_height // 2) - (height // 2)
        
        self.root.geometry(f"{width}x{height}+{pos_x}+{pos_y}") 
        self.root.resizable(False, False)

    def create_widgets(self):
        style = ttk.Style(self.root)
        if "clam" in style.theme_names(): style.theme_use("clam")
        default_bg = style.lookup('TFrame', 'background')
        self.root.configure(bg=default_bg)
        self.log_text_tags = {"warning": {"foreground": "orange"}, "error": {"foreground": "red"}}

        bottom_frame = tk.Frame(self.root, bg=default_bg)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=2)
        tk.Label(bottom_frame, text=f"版本日期：{datetime.now().strftime('%Y年%m月%d日')}", font=self.font_normal, fg="#666666", bg=default_bg).pack(side=tk.LEFT)
        tk.Label(bottom_frame, text="作者：顾江江", font=self.font_normal, fg="#666666", bg=default_bg).pack(side=tk.RIGHT)
        
        disclaimer_frame = tk.Frame(self.root, pady=2, bg=default_bg)
        disclaimer_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10)
        tk.Label(disclaimer_frame, text="本工具仅供骨科内部测试，请勿外传", font=self.font_disclaimer, fg="red", bg=default_bg).pack()

        top_title_frame = tk.Frame(self.root, bg=default_bg)
        top_title_frame.pack(side=tk.TOP, fill=tk.X, pady=(10, 5))
        tk.Label(top_title_frame, text="丹阳市人民医院", font=self.font_subtitle, fg="#0066cc", bg=default_bg).pack()
        tk.Label(top_title_frame, text="日间手术随访表生成系统", font=self.font_title, bg=default_bg).pack()

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

        file_frame = ttk.LabelFrame(left_panel, text="步骤1: 选择文件和路径", padding=5)
        file_frame.pack(fill=tk.BOTH, expand=True) 
        
        self.excel_path_var = tk.StringVar()
        self.template_path_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        
        self.create_file_selector(file_frame, "出院患者列表:", self.excel_path_var, self.select_excel_file)
        
        surgery_frame = ttk.LabelFrame(file_frame, text="手术查询文件 (可多选, 用于补充床号)", padding=5)
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

        progress_frame = ttk.LabelFrame(right_panel, text="处理进度与日志", padding=10)
        progress_frame.pack(fill=tk.BOTH, expand=True)
        
        self.progress_bar = ttk.Progressbar(progress_frame, orient='horizontal', mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))
        
        self.log_text = scrolledtext.ScrolledText(progress_frame, height=5, state='disabled', font=self.font_normal, wrap=tk.WORD)
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
            # 修正后的逻辑：将标签名作为元组传递给 insert 方法
            if level in self.log_text_tags:
                self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {msg}\n", (level,))
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
    
    if splash_active:
        nuitka_splashscreen_python.mark_as_deployed()

    root = tk.Tk()
    root.withdraw()

    app = App(root)

    def finalize_startup():
        if splash_active:
            nuitka_splashscreen_python.close()
        root.deiconify()
        root.focus_force()

    root.after(100, finalize_startup)
    root.mainloop()
