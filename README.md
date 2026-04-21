# 稽查总结会PPT生成器 V8.0（Streamlit Cloud 通用版）

## 版本说明

V8.0 是云端部署兼容版，已从 Windows PowerPoint COM 方案调整为 `python-pptx` 方案。

本版本可以部署在 Streamlit Cloud，也可以在 Windows、Mac 本地运行，生成标准 `.pptx` 文件，生成后可用 Microsoft PowerPoint、WPS、Mac PowerPoint 打开。

## 这版解决什么问题

V7.7 依赖 Windows 本机 Microsoft PowerPoint 和 `pywin32`，部署到 Streamlit Cloud 后会出现：

```text
ModuleNotFoundError: No module named 'win32com'
```

原因是 Streamlit Cloud 是 Linux 环境，不能安装或调用 Windows PowerPoint COM。

V8.0 已取消对 `win32com`、`pywin32`、本机 PowerPoint 的强依赖，改为使用 `python-pptx` 直接生成 PPT，因此可以在云端正常运行。

## 主文件

Streamlit 部署时主文件填写：

```text
src/app.py
```

## 运行要求

- Python 3.10+
- Streamlit
- openpyxl
- python-pptx
- Pillow

依赖已写入：

```text
requirements.txt
```

## Streamlit Cloud 部署方式

在 Streamlit Cloud 中选择仓库：

```text
1450341742-ops/audit-plan-ppt-engine
```

主文件填写：

```text
src/app.py
```

部署后页面标题应显示：

```text
稽查总结会PPT生成器 V8.0
```

## 本地运行方式

```powershell
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
streamlit run src/app.py
```

也可以在 Windows 上双击：

```text
run.bat
```

## 批量命令行运行

```powershell
python src/batch_generate.py --excel_dir "D:\你的Excel文件夹" --output_dir "D:\输出文件夹"
```

## 模板说明

V8.0 不强制依赖 `assets/template.pptx`。

- 没有 `assets/template.pptx`：系统自动生成蓝色商务风 PPT。
- 有 `assets/template.pptx`：系统会尝试读取模板尺寸和基础设置，但不会再调用 Windows PowerPoint。

## 文件结构

```text
README.md
requirements.txt
run.bat
src/app.py
src/batch_generate.py
src/parser.py
src/renderer.py
assets/README.md
```

## 当前版本重点

- 支持 Streamlit Cloud 部署；
- 不再出现 `No module named win32com` 报错；
- 支持 Excel 上传后直接生成 PPT；
- 支持单个下载和 ZIP 批量下载；
- 生成文件可用 PowerPoint、WPS、Mac PowerPoint 打开。
