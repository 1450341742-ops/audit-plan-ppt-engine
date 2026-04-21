# 稽查总结会PPT生成器 V8.1（严格模板驱动云端版）

## 版本定位

V8.1 同时满足三个要求：

1. 严格使用 `assets/template.pptx` 作为模板来源；
2. 支持 Streamlit Cloud / Windows / Mac 运行；
3. 生成标准 `.pptx` 文件，可用 Microsoft PowerPoint、WPS、Mac PowerPoint 打开。

## 核心实现

本版本不再调用 Windows PowerPoint COM，也不再用代码重新画一套蓝色商务风页面。

系统采用：

```text
读取 assets/template.pptx
复制模板页 OOXML 结构
保留模板中的 Logo、背景、图片、表格、边框、母版和页面比例
仅向模板中的表格/文本框填充 Excel 解析结果
输出标准 pptx
```

## 主文件

Streamlit Cloud 部署时主文件填写：

```text
src/app.py
```

## 依赖

```text
streamlit
openpyxl
python-pptx
lxml
Pillow
```

依赖已写入：

```text
requirements.txt
```

## 模板要求

必须存在有效模板文件：

```text
assets/template.pptx
```

模板页序号需要保持当前结构：

```text
第2页：封面
第3页：感谢页
第4页：目录页
第5页：第一部分页
第6页：中心稽查概述
第7页：中心稽查范围
第8页：第二部分页
第9页：中心稽查分类和数量
第10-21页：各问题分类页
第22页：建议项
第23页：结束页
```

## 运行方式

```powershell
pip install -r requirements.txt
streamlit run src/app.py
```

## Streamlit Cloud 部署方式

选择仓库：

```text
1450341742-ops/audit-plan-ppt-engine
```

主文件：

```text
src/app.py
```

部署后页面标题应显示：

```text
稽查总结会PPT生成器 V8.1（模板驱动云端版）
```

## 当前版本重点

- 不依赖 pywin32；
- 不依赖 Microsoft PowerPoint；
- 不依赖 WPS；
- 支持 Streamlit Cloud；
- 支持 Windows / Mac 本地运行；
- 严格读取并复制 PPT 模板页；
- 生成的 PPT 可用 PowerPoint、WPS 打开。
