# Day Surgery Form Automator (日间手术随访表生成系统)

[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/)
[![Version](https://img.shields.io/badge/Version-2.6-brightgreen.svg)]()
[![License](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)

这是一个功能强大且用户友好的医疗文档自动化工具，专为简化“日间手术随访表”的生成流程而设计。作为原始版本的重大升级，此版本在**代码健壮性、用户体验、可维护性和界面美观度**上都进行了全面的优化和重构。

![截图](./docs/images/screenshots/screenshot.png)
*<p align="center">软件截图</p>*

---

## ✨ 核心优势与优化 (V2.0)

- **现代化的UI/UX**:
    - **高DPI自适应**: 界面在高分辨率屏幕上显示清晰，字体大小恰当，解决了模糊和错位问题。
    - **统一的视觉风格**: 所有控件背景色统一，界面更加专业、美观。
    - **实时进度反馈**: 使用进度条和带颜色区分（警告/错误）的滚动日志框，实时显示处理进度，程序运行状态一目了然。
    - **智能预警**: 新增**床号缺失提醒**功能，若Excel缺少关键信息，会在日志和最终弹窗中明确提示用户，避免疏漏。

- **高度健壮的内核**:
    - **全文档替换**: 修复了原先无法替换**页眉**占位符的Bug，确保 `{{科室}}`、`{{患者出院年月}}` 等位于页眉的信息能被正确填充。
    - **强大的错误处理**: 对文件读取、数据转换等关键步骤增加了全面的 `try-except` 保护，能有效防止因数据格式错误导致的程序崩溃。
    - **智能标题行检测**: 能够更可靠地自动检测Excel中的标题行，并在失败时引导用户手动输入。

- **灵活的配置与维护**:
    - **中央配置系统**: 将所有关键参数（如日间手术天数定义、Excel列名映射等）提取到脚本顶部的 `CONFIG` 字典中，方便未来快速调整，无需修改核心代码。
    - **面向对象重构 (OOP)**: 代码被重构为 `App` (界面) 和 `DocumentGenerator` (逻辑) 两个独立的类，结构清晰，极大地提高了代码的可读性和可维护性。

- **多线程处理**:
    - 将耗时的文件生成任务放在独立的线程中执行，确保了GUI在处理过程中始终保持响应，提升了用户体验。

## 🚀 安装与运行

**1. 准备环境**

- 确保您的电脑上安装了 [Python 3.7](https://www.python.org/downloads/) 或更高版本。
- (推荐) 创建并激活一个虚拟环境，以保持项目依赖的纯净：
  ```bash
  python -m venv venv
  # Windows
  .\venv\Scripts\activate
  # macOS / Linux
  source venv/bin/activate
  ```

**2. 安装依赖**

项目依赖已记录在 `requirements.txt` 文件中。请在项目根目录下运行以下命令进行安装：

```bash
pip install -r requirements.txt
```
*(如果项目中还未创建 `requirements.txt` 文件，请创建一个并写入以下内容):*
```
pandas
python-docx
```

**3. 运行程序**

直接运行主脚本文件：

```bash
python day_durgery_form_automator.py
```

**4. 打包为EXE (可选)**

如果您希望将程序分享给没有安装Python环境的同事，可以使用 `PyInstaller` 将其打包为单个可执行文件：

```bash
# 安装 PyInstaller
pip install pyinstaller

# 执行打包命令 (使用新的英文名)
python -m PyInstaller --onefile --windowed --icon=app.ico --name="day_durgery_form_automator" day_durgery_form_automator.py
```
打包成功后，在生成的 `dist` 文件夹中即可找到 `.exe` 文件。

## 🔧 配置说明

本程序的所有关键配置项均位于脚本顶部的 `CONFIG` 字典中，您可以根据需要进行修改：

- `app_title`: 窗口标题。
- `day_surgery_max_days`: 定义日间手术的最大住院天数。
- `column_mapping`: **核心配置**。定义了程序内部字段名与您Excel文件中列名的对应关系。如果您的Excel列名有变，只需修改此处的中文部分即可。
- `required_internal_keys`: 定义了哪些字段是生成文档所必需的。
- `template_placeholders`: 定义了Word模板中的占位符与程序内部字段的对应关系。

## 👨‍💻 贡献者

- **原始作者**: 顾江江
- **重构与优化**: Google AI

## 📜 许可协议

本项目采用 [MIT License](https://opensource.org/licenses/MIT) 开源协议。
