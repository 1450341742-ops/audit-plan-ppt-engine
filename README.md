# 稽查总结会PPT生成器 V7.7（PowerPoint原生模板概述表格精准版）

## 这版解决什么问题

之前版本使用 `python-pptx`、背景图片或 XML 复制模板页，容易破坏 PPT 内部关系，导致：

- “单击此处添加标题”残留；
- 图片/Logo无法显示；
- 字体和版式漂移；
- 表格、边框、母版元素丢失；
- 结束页英文竖排等乱码问题。

V7.7 改为调用 Windows 本机 Microsoft PowerPoint 原生复制模板页，模板页由 PowerPoint 自己复制，因此最大程度保留原始 PPT 的 Logo、背景、图标、表格、母版、蓝色边框和页面比例。

## 运行要求

必须在 Windows 电脑本机运行，并安装：

1. Microsoft PowerPoint；
2. Python 3.10+；
3. pywin32。

## 安装与运行

```powershell
cd D:\final_template_engine_v7_7_overview_latest_table_flat
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
streamlit run src/app.py
```

打开网页后，标题应显示：

```text
稽查总结会PPT生成器 V7.7（PowerPoint原生模板概述表格精准版）
```

## 批量命令行运行

```powershell
cd D:\final_template_engine_v7_7_overview_latest_table_flat
.\.venv\Scripts\activate
python src/batch_generate.py --excel_dir "D:\你的Excel文件夹" --output_dir "D:\输出文件夹"
```

## 注意事项

- 不要删除 `assets/template.pptx`，这是原生模板来源。
- 运行时电脑可能会短暂打开 PowerPoint，这是正常现象。
- 如果提示 `No module named win32com`，执行：

```powershell
pip install pywin32
```

- 如果提示无法调用 PowerPoint，请确认本机已安装 Microsoft PowerPoint，而不是只安装 WPS。
