import os
import sys
from datetime import datetime, timedelta
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from docx import Document
import pandas as pd
from openpyxl import load_workbook

# ======================== 工具函数 ========================
def clean_date_str(date_str):
    """清洗日期字符串（去除时分秒）"""
    if not date_str:
        return ""
    return str(date_str).split()[0]  # 取空格前的部分

def excel_date_to_str(excel_date):
    """智能日期转换（兼容Excel序列号、文本日期和空值）"""
    if excel_date in ("", None):
        return ""
    try:
        if isinstance(excel_date, (int, float)):
            date = datetime(1899, 12, 30) + timedelta(days=float(excel_date))
            return date.strftime("%Y-%m-%d")  # 仅日期部分
        else:
            date_str = str(excel_date)
            # 尝试多种日期格式
            for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"):
                try:
                    date = datetime.strptime(date_str, fmt)
                    return date.strftime("%Y-%m-%d")
                except:
                    continue
            return clean_date_str(date_str)  # 最终回退方案
    except:
        return clean_date_str(str(excel_date))

def get_day_after_discharge(discharge_date, days=7):
    """计算出院后第N天的日期（自动处理空值）"""
    if discharge_date in ("", None):
        return ""
    try:
        if isinstance(discharge_date, (int, float)):
            base_date = datetime(1899, 12, 30) + timedelta(days=discharge_date)
        else:
            base_date = datetime.strptime(clean_date_str(discharge_date), "%Y-%m-%d")
        return (base_date + timedelta(days=days)).strftime("%Y-%m-%d")
    except:
        return ""

def read_excel_smart(path):
    """智能读取Excel（兼容.xls和.xlsx）"""
    try:
        if path.endswith('.xlsx'):
            wb = load_workbook(path)
            return wb.active
        else:
            df = pd.read_excel(path, engine='xlrd')
            return df.where(pd.notnull(df), None).values.tolist()
    except Exception as e:
        raise ValueError(f"读取Excel失败：{str(e)}")

# ======================== 核心功能 ========================
def analyze_date_distribution(data, date_col_idx, header_row_idx):
    """分析出院日期分布情况（改为使用出院日期）"""
    month_counts = defaultdict(int)
    month_samples = defaultdict(list)
    
    for row in data[header_row_idx:]:
        date_str = excel_date_to_str(row[date_col_idx])
        if date_str:
            month_key = date_str[:7]  # YYYY-MM
            month_counts[month_key] += 1
            name = str(row[0]) if row[0] else "未知患者"
            month_samples[month_key].append(name)
    return month_counts, month_samples

def select_date_mode_gui(month_counts, month_samples):
    """GUI界面选择主导月份"""
    if len(month_counts) == 1:
        return list(month_counts.keys())[0]
    
    root = tk.Tk()
    root.withdraw()
    
    options = []
    for month, count in month_counts.items():
        sample_names = ", ".join(month_samples[month][:3])
        if len(month_samples[month]) > 3:
            sample_names += "等"
        options.append(f"{month}月（共{count}例，如：{sample_names}）")
    
    choice = simpledialog.askstring(
        "选择主导月份",
        "发现多个出院月份，请选择：\n" + "\n".join(options),
        parent=root
    )
    return choice.split("月")[0] if choice else None

def get_header_row_index_gui():
    """GUI界面获取标题行号"""
    root = tk.Tk()
    root.withdraw()
    return simpledialog.askinteger(
        "设置标题行",
        "请输入Excel中列标题所在行号（从1开始）：\n\n"
        "例如：\n"
        "• 如果有3行表头，输入4\n"
        "• 如果首行就是列标题，输入1",
        parent=root,
        minvalue=1,
        maxvalue=100
    )

def build_column_mapping(header_row):
    """动态建立列名到索引的映射"""
    return {str(cell).strip(): idx for idx, cell in enumerate(header_row) if cell}

def generate_documents():
    """主生成函数"""
    try:
        # ========== 1. 文件选择 ==========
        root = tk.Tk()
        root.withdraw()
        
        excel_path = filedialog.askopenfilename(
            title="选择出院患者记录单",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if not excel_path:
            return

        template_path = filedialog.askopenfilename(
            title="选择随访表模板",
            filetypes=[("Word模板", "*.docx"), ("所有文件", "*.*")]
        )
        if not template_path:
            return

        output_dir = filedialog.askdirectory(title="选择保存位置")
        if not output_dir:
            return

        # ========== 2. 数据准备 ==========
        # 智能读取Excel
        try:
            excel_data = read_excel_smart(excel_path)
            if isinstance(excel_data, list):  # .xls格式
                data = excel_data
            else:  # .xlsx格式
                data = list(excel_data.iter_rows(values_only=True))
        except Exception as e:
            messagebox.showerror("错误", f"读取Excel失败：{str(e)}")
            return

        # 获取标题行
        header_row_idx = get_header_row_index_gui()
        if not header_row_idx:
            return

        # 动态列映射
        col_map = build_column_mapping(data[header_row_idx - 1])
        
        # 检查必要字段
        required_fields = {
            "姓名": "患者姓名",
            "出院科室": "科室信息",
            "住院号": "住院编号",
            "出院日期": "出院时间",
            "住院天数": "住院时长"
        }
        missing_fields = [name for field, name in required_fields.items() if field not in col_map]
        if missing_fields:
            messagebox.showerror(
                "错误",
                f"Excel中缺少以下必要列:\n{', '.join(missing_fields)}\n\n"
                f"当前识别的列标题：\n{', '.join(col_map.keys())}"
            )
            return

        # 检查床号
        bed_warning = "床号" not in col_map
        if bed_warning:
            if not messagebox.askyesno("提示", "未找到床号信息，将在模板中显示为'（请手动填写）'。是否继续？"):
                return

        # ========== 3. 智能日期处理（改为使用出院日期） ==========
        month_counts, month_samples = analyze_date_distribution(
            data, col_map["出院日期"], header_row_idx
        )
        if not month_counts:
            messagebox.showerror("错误", "未找到有效的出院日期数据！")
            return

        selected_month = select_date_mode_gui(month_counts, month_samples)
        if not selected_month:
            return

        patient_year_month = f"{selected_month[:4]}年{selected_month[5:]}月"

        # ========== 4. 文档生成 ==========
        success_count = 0
        for row in data[header_row_idx:]:
            # 检查是否为日间手术（住院天数≤2天）
            try:
                if int(row[col_map["住院天数"]]) <= 2:  # 关键修改：确保转换为整数比较
                    # 关键字段提取
                    出院科室 = str(row[col_map["出院科室"]]) if row[col_map["出院科室"]] else "未知科室"
                    姓名 = str(row[col_map["姓名"]]) if row[col_map["姓名"]] else "未知姓名"
                    出院日期 = row[col_map["出院日期"]]

                    # 生成文件名
                    filename = (
                        f"{patient_year_month}"
                        f"{出院科室}日间手术随访登记表_"
                        f"{get_day_after_discharge(出院日期)}_{姓名}.docx"
                    )
                    output_path = os.path.join(output_dir, filename)

                    # 替换模板内容
                    doc = Document(template_path)
                    replacements = {
                        "{{患者出院年月}}": patient_year_month,  # 改为出院年月
                        "{{科室}}": 出院科室,
                        "{{姓名}}": 姓名,
                        "{{性别}}": str(row[col_map.get("性别", "")] or ""),
                        "{{年龄}}": str(row[col_map.get("年龄", "")] or ""),
                        "{{住院号}}": str(row[col_map["住院号"]] or ""),
                        "{{床号}}": str(row[col_map["床号"]] if not bed_warning and "床号" in col_map else "（请手动填写）"),
                        "{{入院日期}}": excel_date_to_str(row[col_map.get("入院日期", "")]),
                        "{{出院日期}}": excel_date_to_str(出院日期),
                        "{{手术日期}}": excel_date_to_str(row[col_map.get("手术日期", "")]),
                        "{{手术名称}}": str(row[col_map.get("手术名称", "")] or ""),
                        "{{出院诊断}}": str(row[col_map.get("最后诊断1", "")] or ""),
                        "{{联系电话}}": str(row[col_map.get("联系电话", "")] or ""),
                        "{{经治医生}}": str(row[col_map.get("经治医生", "")] or ""),
                        "{{随访日期}}": get_day_after_discharge(出院日期)
                    }

                    # 全文档替换
                    for paragraph in doc.paragraphs:
                        for old, new in replacements.items():
                            if old in paragraph.text:
                                paragraph.text = paragraph.text.replace(old, new)

                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for old, new in replacements.items():
                                    if old in cell.text:
                                        cell.text = cell.text.replace(old, new)

                    doc.save(output_path)
                    success_count += 1
                    print(f"已生成：{filename}")  # 调试用

            except Exception as e:
                print(f"处理第{data.index(row)+1}行时出错：{str(e)}")

        # ========== 5. 完成提示 ==========
        if success_count > 0:
            messagebox.showinfo(
                "完成",
                f"成功生成 {success_count} 份随访表\n"
                f"保存位置：{output_dir}\n"
                f"统一出院年月：{patient_year_month}"
            )
        else:
            messagebox.showerror(
                "错误",
                "未生成任何文档！可能原因：\n"
                "1. 没有住院天数=1的记录\n"
                "2. 日期格式不正确\n"
                "3. 必要列数据缺失"
            )

    except Exception as e:
        messagebox.showerror("错误", f"处理过程中出错：{str(e)}")

# ======================== 主程序 ========================
if __name__ == "__main__":
    # Windows高DPI适配
    if sys.platform == 'win32':
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    
    # 创建主界面
    root = tk.Tk()
    root.title("日间手术随访表生成系统v1.0 By顾江江")
    root.geometry("800x200")
    root.resizable(False, False)

    # 样式设置
    font_style = ("微软雅黑", 12)
    btn_style = {
        "font": font_style,
        "bg": "#4CAF50",
        "fg": "white",
        "activebackground": "#45a049",
        "padx": 20,
        "pady": 10
    }

    # 开始按钮
    tk.Label(root, text="日间手术随访表生成系统", font=("微软雅黑", 16)).pack(pady=20)
    tk.Button(
        root, 
        text="开始生成", 
        command=lambda: [root.destroy(), generate_documents()],
        **btn_style
    ).pack()

    root.mainloop()