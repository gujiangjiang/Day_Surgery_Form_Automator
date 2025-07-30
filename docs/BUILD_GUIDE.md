# 自动化程序编译与打包指南

本指南旨在为 "日间手术随访表生成系统 (Polars版)" 提供清晰的安装、运行及打包说明。

---

## 🚀 安装与运行

在开始之前，请确保您的电脑已正确配置好运行环境。

### 1. 环境准备

* **安装 Python**: 确保您的电脑上安装了 [Python 3.7](https://www.python.org/downloads/) 或更高版本。在安装时，建议勾选 "Add Python to PATH" 选项。

* **创建虚拟环境 (推荐)**: 为了保持项目依赖的纯净，避免与其他项目的库产生冲突，强烈建议创建一个独立的虚拟环境。

    ```bash
    # 在项目根目录下，创建名为 venv 的虚拟环境
    python -m venv venv
    ```

* **激活虚拟环境**:
    * **Windows (PowerShell)**:
        ```powershell
        # 激活虚拟环境
        .\venv\Scripts\activate
        
        # 如果上方命令因执行策略报错，请运行此命令后再试 (仅需一次)
        Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
        ```
    * **macOS / Linux**:
        ```bash
        source venv/bin/activate
        ```
    激活成功后，您会看到命令行提示符前面出现 `(venv)` 字样。

### 2. 安装依赖

项目依赖已写入 `requirements.txt` 文件。请在**已激活虚拟环境**的命令行中运行以下命令进行安装：

```bash
pip install -r requirements.txt
```

> **提示**: `requirements.txt` 文件应包含以下内容，这已是运行程序的最小依赖集合:
> ```text
> polars
> python-docx
> fastexcel
> ```

### 3. 运行程序

一切准备就绪后，直接运行主脚本即可启动程序：

```bash
python day_durgery_form_automator.py
```

---

## 📦 打包为可执行文件 (.EXE)

如果您希望将程序分享给没有安装 Python 环境的同事，可以将其打包为单个 `.exe` 文件。这里提供两种主流的打包方式：`PyInstaller` 和 `Nuitka`。

### 方式一：使用 PyInstaller (简单快捷)

PyInstaller 是一个非常流行的打包工具，使用简单，适合快速分发。

**1. 安装 PyInstaller**
```bash
pip install pyinstaller
```

**2. 执行打包命令**
```bash
pyinstaller --onefile --windowed --icon="app.ico" --name="day_durgery_form_automator" day_durgery_form_automator.py
```
* `--onefile`: 将所有依赖和脚本打包成一个独立的 `.exe` 文件。
* `--windowed`: 运行程序时不显示黑色的命令行窗口 (适用于图形界面程序)。
* `--icon`: 为生成的 `.exe` 文件指定一个图标 (请将 `app.ico` 文件放在同目录下)。
* `--name`: 指定生成的可执行文件的名称。

打包成功后，在生成的 `dist` 文件夹中即可找到 `Day Surgery Form Automator.exe` 文件。

### 方式二：使用 Nuitka (性能更优)

Nuitka 会将 Python 代码编译成 C++ 代码，然后再编译成可执行文件，性能和体积通常优于 PyInstaller。

**1. 安装 Nuitka**
```bash
pip install nuitka
```
*Nuitka* 还需要一个 C++ 编译器。在 Windows 上，它会自动提示并帮助您下载安装 *MinGW64*，按照提示操作即可。

**2. 执行打包命令**

为了命令的清晰性，建议使用多行格式。请根据您使用的命令行工具（PowerShell 或 CMD）选择对应的版本。

**如果您正在使用 Windows PowerShell (推荐):**

请复制并执行以下命令。注意，换行符是反引号 (`` ` ``)。

```powershell
python -m nuitka `
  --onefile `
  --windows-console-mode=disable `
  --enable-plugin=tk-inter `
  --include-module=polars,docx,fastexcel `
  --onefile-windows-splash-screen-image=splash.png `
  --windows-icon-from-ico=app.ico `
  --mingw64 `
  --output-dir=build `
  --output-filename="day_durgery_form_automator.exe" `
  --file-version=6.4.1.0 `
  --product-version=6.4 `
  --company-name="Danyang People's Hospital" `
  --product-name="Day Surgery Form Automator" `
  --copyright="© 2025 gujiangjiang" `
  day_durgery_form_automator.py
```

**命令提示符 (cmd.exe):**

```batch
python -m nuitka ^
  --onefile ^
  --windows-console-mode=disable ^
  --enable-plugin=tk-inter ^
  --include-module=polars,docx,fastexcel ^
  --onefile-windows-splash-screen-image=splash.png ^
  --windows-icon-from-ico=app.ico ^
  --mingw64 ^
  --output-dir=build ^
  --output-filename="day_durgery_form_automator.exe" ^
  --file-version=6.4.1.0 ^
  --product-version=6.4 ^
  --company-name="Danyang People's Hospital" ^
  --product-name="Day Surgery Form Automator" ^
  --copyright="© 2025 gujiangjiang" ^
  day_durgery_form_automator.py
```

**命令参数详解**:

* `` ` `` (反引号): **PowerShell** 中的换行符。
* `^`: 传统**命令提示符 (CMD)** 中的换行符。
* `--onefile`: 打包为单文件。
* `--windows-disable-console`: 禁用控制台窗口 (同 `pyinstaller --windowed`)。
* `--enable-plugin=tk-inter`: 如果您的程序使用了 `tkinter` 图形库，需启用此插件。
* `--include-module=...`: 强制包含某些 Nuitka 可能检测不到的库。
* `--onefile-windows-splash-screen-image`: 为程序启动时添加一个闪屏/加载图片。
* `--windows-icon-from-ico`: 指定程序图标。
* `--mingw64`: (可选) 明确指定使用 MinGW64 编译器。
* `--output-dir`: 指定输出文件夹的名称。
* `--output-filename`: 指定输出文件的名称。
* `--file-version`, `--product-version`, `--company-name`, `--product-name`, `--copyright`: 为 `.exe` 文件添加详细的元数据，可在文件属性中查看。

打包成功后，在生成的 `build` 文件夹中即可找到最终的 `.exe` 文件。
* `--include-module=...`: 强制包含 Nuitka 可能检测不到的库。对于 Polars 版本，我们明确加入 `polars`, `docx`, 和 `fastexcel`。
* 其他参数如 `--file-version`, `--product-name` 等已更新至最新版本信息。
