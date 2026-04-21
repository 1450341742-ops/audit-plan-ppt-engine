from __future__ import annotations

"""
严格模板版渲染器

本文件按用户上传的 auditpptv78 方向恢复：
1. 不再使用 python-pptx 重新画蓝色商务风页面；
2. 必须调用 Windows 本机 Microsoft PowerPoint 原生复制模板页；
3. 通过复制 assets/template.pptx 中的页面，保留 Logo、背景、母版、表格、边框和页面比例；
4. 仅清空/覆盖动态文字，把 Excel 解析结果写入模板对应页面。

注意：该方案无法在 Streamlit Cloud Linux 环境运行；必须部署在 Windows 主机，且安装 Microsoft PowerPoint + pywin32。
"""

from pathlib import Path
from typing import Any
import re
import time

BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = BASE_DIR / "assets" / "template.pptx"

MSO_TRUE = -1
MSO_FALSE = 0
MSO_TEXT_ORIENTATION_HORIZONTAL = 1
PP_ALIGN_LEFT = 1
PP_ALIGN_CENTER = 2
PPSAVEAS_OPENXML_PRESENTATION = 24

SLIDE_COVER = 2
SLIDE_THANKS = 3
SLIDE_TOC = 4
SLIDE_PART1 = 5
SLIDE_OVERVIEW = 6
SLIDE_SCOPE = 7
SLIDE_PART2 = 8
SLIDE_COUNTS = 9
SLIDE_SUGGESTION = 22
SLIDE_ENDING = 23

LEFT_CATS = [
    "国家药物临床试验政策法规的遵循",
    "伦理委员会审核要求的遵循",
    "知情同意书（ICF）的签署和记录",
    "原始文件的建立、内容和记录",
    "门诊/住院HIS、LIS、PACS等系统数据溯源",
    "方案依从性",
    "药物疗效/研究评价指标的评估",
]

RIGHT_CATS = [
    "安全性信息评估，记录与报告",
    "CRF填写（时效性、一致性、溯源性、完整性）",
    "试验用药品管理",
    "生物样本管理",
    "临床研究必须文件",
    "申办方/CRO职责",
    "其他",
]

CAT_TO_TEMPLATE_SLIDE = {
    "国家药物临床试验政策法规的遵循": 10,
    "伦理委员会审核要求的遵循": 11,
    "知情同意书（ICF）的签署和记录": 12,
    "原始文件的建立、内容和记录": 13,
    "门诊/住院HIS、LIS、PACS等系统数据溯源": 14,
    "方案依从性": 15,
    "药物疗效/研究评价指标的评估": 16,
    "安全性信息评估，记录与报告": 17,
    "CRF填写（时效性、一致性、溯源性、完整性）": 18,
    "试验用药品管理": 19,
    "生物样本管理": 20,
    "临床研究必须文件": 21,
    "申办方/CRO职责": 21,
    "其他": 21,
}


def rgb(r: int, g: int, b: int) -> int:
    return int(r) + int(g) * 256 + int(b) * 65536

BLACK = rgb(0, 0, 0)
WHITE = rgb(255, 255, 255)


def pt(inches: float) -> float:
    return inches * 72.0


def _clean(v: Any) -> str:
    s = str(v or "").replace("\r", "\n")
    s = re.sub(r"[ \t\xa0]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def _get_powerpoint_app():
    try:
        import pythoncom  # type: ignore
        import win32com.client as win32  # type: ignore
    except Exception as exc:
        raise RuntimeError(
            "严格模板版需要 Windows + Microsoft PowerPoint + pywin32。"
            "Streamlit Cloud 是 Linux 环境，不能调用 PowerPoint 原生模板。"
        ) from exc
    pythoncom.CoInitialize()
    app = win32.DispatchEx("PowerPoint.Application")
    app.Visible = MSO_TRUE
    try:
        app.DisplayAlerts = 0
    except Exception:
        pass
    return app


def _set_font(text_range, size: float, bold: bool = False, color: int = BLACK):
    try:
        text_range.Font.Name = "Microsoft YaHei"
        text_range.Font.NameFarEast = "Microsoft YaHei"
        text_range.Font.Size = float(size)
        text_range.Font.Bold = MSO_TRUE if bold else MSO_FALSE
        text_range.Font.Color.RGB = color
    except Exception:
        pass


def _write_shape_text(shape, text: str, size: float, bold: bool = False, color: int = BLACK,
                      align: int = PP_ALIGN_LEFT, margin: float = 3.0):
    text = str(text or "—")
    try:
        tf2 = shape.TextFrame2
        tf2.TextRange.Text = text
        tf2.MarginLeft = margin
        tf2.MarginRight = margin
        tf2.MarginTop = margin
        tf2.MarginBottom = margin
        tf2.WordWrap = MSO_TRUE
        tf2.TextRange.Font.Name = "Microsoft YaHei"
        tf2.TextRange.Font.NameFarEast = "Microsoft YaHei"
        tf2.TextRange.Font.Size = float(size)
        tf2.TextRange.Font.Bold = MSO_TRUE if bold else MSO_FALSE
        tf2.TextRange.Font.Fill.ForeColor.RGB = color
        tf2.TextRange.ParagraphFormat.Alignment = align
        return
    except Exception:
        pass
    try:
        tr = shape.TextFrame.TextRange
        tr.Text = text
        _set_font(tr, size, bold, color)
        tr.ParagraphFormat.Alignment = align
        shape.TextFrame.MarginLeft = margin
        shape.TextFrame.MarginRight = margin
        shape.TextFrame.MarginTop = margin
        shape.TextFrame.MarginBottom = margin
        shape.TextFrame.WordWrap = MSO_TRUE
    except Exception:
        pass


def _add_textbox(slide, x: float, y: float, w: float, h: float, text: str,
                 font_size: float = 12, bold: bool = False, color_rgb: int = BLACK,
                 align: int = PP_ALIGN_LEFT, margin: float = 3.0):
    last_err = None
    for i in range(4):
        shp = None
        try:
            shp = slide.Shapes.AddTextbox(MSO_TEXT_ORIENTATION_HORIZONTAL, pt(x), pt(y), pt(w), pt(h))
            shp.Fill.Visible = MSO_FALSE
            shp.Line.Visible = MSO_FALSE
            _write_shape_text(shp, text, font_size, bold, color_rgb, align, margin)
            return shp
        except Exception as e:
            last_err = e
            try:
                if shp is not None:
                    shp.Delete()
            except Exception:
                pass
            time.sleep(0.1 * (i + 1))
    raise last_err if last_err else RuntimeError("AddTextbox failed")


def _iter_shapes_reverse(slide):
    try:
        for i in range(slide.Shapes.Count, 0, -1):
            yield slide.Shapes.Item(i)
    except Exception:
        return


def _has_table(shp) -> bool:
    try:
        return bool(getattr(shp, "HasTable", 0))
    except Exception:
        return False


def _has_text(shp) -> bool:
    try:
        return bool(getattr(shp, "HasTextFrame", 0))
    except Exception:
        return False


def _delete_text_layer(slide, clear_tables: bool = False):
    for shp in _iter_shapes_reverse(slide):
        try:
            if _has_table(shp):
                if clear_tables:
                    tbl = shp.Table
                    for r in range(1, tbl.Rows.Count + 1):
                        for c in range(1, tbl.Columns.Count + 1):
                            try:
                                tbl.Cell(r, c).Shape.TextFrame.TextRange.Text = ""
                            except Exception:
                                pass
                continue
            if _has_text(shp):
                txt = ""
                try:
                    txt = shp.TextFrame.TextRange.Text
                except Exception:
                    pass
                if txt or "Title" in str(getattr(shp, "Name", "")) or "Placeholder" in str(getattr(shp, "Name", "")):
                    shp.Delete()
        except Exception:
            pass


def _find_largest_table(slide):
    best = None
    area = -1
    for shp in _iter_shapes_reverse(slide):
        try:
            if _has_table(shp):
                a = float(shp.Width) * float(shp.Height)
                if a > area:
                    best = shp
                    area = a
        except Exception:
            pass
    return best


def _set_cell(cell, text: str, size: float = 12, bold: bool = False,
              color: int = BLACK, align: int = PP_ALIGN_LEFT):
    text = str(text or "—")
    try:
        tf2 = cell.Shape.TextFrame2
        tf2.TextRange.Text = text
        tf2.MarginLeft = 6
        tf2.MarginRight = 6
        tf2.MarginTop = 3
        tf2.MarginBottom = 3
        tf2.WordWrap = MSO_TRUE
        tf2.VerticalAnchor = 3
        tf2.TextRange.Font.Name = "Microsoft YaHei"
        tf2.TextRange.Font.NameFarEast = "Microsoft YaHei"
        tf2.TextRange.Font.Size = float(size)
        tf2.TextRange.Font.Bold = MSO_TRUE if bold else MSO_FALSE
        tf2.TextRange.Font.Fill.ForeColor.RGB = color
        tf2.TextRange.ParagraphFormat.Alignment = align
        return
    except Exception:
        pass
    try:
        tr = cell.Shape.TextFrame.TextRange
        tr.Text = text
        _set_font(tr, size, bold, color)
        tr.ParagraphFormat.Alignment = align
        cell.Shape.TextFrame.MarginLeft = 6
        cell.Shape.TextFrame.MarginRight = 6
        cell.Shape.TextFrame.MarginTop = 3
        cell.Shape.TextFrame.MarginBottom = 3
        cell.Shape.TextFrame.WordWrap = MSO_TRUE
    except Exception:
        pass


def _duplicate_slide_to_end(prs, idx: int):
    prs.Slides(idx).Copy()
    rng = prs.Slides.Paste(prs.Slides.Count + 1)
    try:
        return rng.Item(1)
    except Exception:
        return prs.Slides(prs.Slides.Count)


def _delete_original_template_slides(prs, count: int):
    for _ in range(count):
        try:
            prs.Slides(1).Delete()
        except Exception:
            break


def _split_text(text: str, limit: int) -> list[str]:
    text = _clean(text)
    if not text:
        return ["—"]
    parts = []
    buf = ""
    for para in re.split(r"\n\s*\n", text):
        para = para.strip()
        if not para:
            continue
        if len(buf) + len(para) + 2 <= limit:
            buf = (buf + "\n\n" + para).strip()
        else:
            if buf:
                parts.append(buf)
            while len(para) > limit:
                parts.append(para[:limit])
                para = para[limit:]
            buf = para
    if buf:
        parts.append(buf)
    return parts or ["—"]


def _paginate_issue(issue: dict) -> list[dict]:
    basis_parts = _split_text(issue.get("basis", "—"), 520)
    desc_parts = _split_text(issue.get("description", "—"), 900)
    total = max(len(basis_parts), len(desc_parts))
    out = []
    for i in range(total):
        x = dict(issue)
        x["basis"] = basis_parts[i] if i < len(basis_parts) else "—"
        x["description"] = desc_parts[i] if i < len(desc_parts) else "—"
        x["_sub_page"] = i + 1
        x["_sub_total"] = total
        out.append(x)
    return out


def _render_cover(slide, context: dict):
    _delete_text_layer(slide, clear_tables=False)
    meta = context.get("meta", {})
    project = _clean(meta.get("project_name", "—"))
    center = _clean(meta.get("center_name", "—"))
    center_no = _clean(meta.get("center_no", ""))
    audit_date = _clean(meta.get("audit_date", "—"))
    title = f"{project}-{center}" + (f"（中心编号{center_no}）" if center_no else "")
    _add_textbox(slide, 0.85, 1.40, 11.65, 0.50, title, font_size=22, bold=True, color_rgb=BLACK, align=PP_ALIGN_CENTER, margin=0)
    _add_textbox(slide, 0.85, 1.92, 11.65, 0.48, "中心稽查末次会议", font_size=28, bold=True, color_rgb=BLACK, align=PP_ALIGN_CENTER, margin=0)
    _add_textbox(slide, 0.85, 4.82, 11.65, 0.35, f"时间：{audit_date}", font_size=16, color_rgb=BLACK, align=PP_ALIGN_CENTER, margin=0)
    _add_textbox(slide, 0.85, 5.28, 11.65, 0.35, "北京万宁睿和医药科技有限公司", font_size=16, bold=True, color_rgb=BLACK, align=PP_ALIGN_CENTER, margin=0)


def _render_overview(slide, context: dict):
    _delete_text_layer(slide, clear_tables=False)
    meta = context.get("meta", {})
    _add_textbox(slide, 0.72, 0.48, 6.30, 0.70, "一、中心稽查概述", font_size=28, bold=True, color_rgb=BLACK, margin=0)
    table_shape = _find_largest_table(slide)
    project = _clean(meta.get("project_name", "—"))
    sponsor = _clean(meta.get("sponsor", "—"))
    pi = _clean(meta.get("pi", "—"))
    if pi and pi != "—" and not pi.endswith("教授"):
        pi = f"{pi}教授"
    center = _clean(meta.get("center_name", "—"))
    enrollment = _clean(meta.get("enrollment", "—"))
    audit_date = _clean(meta.get("audit_date", "—"))
    auditor = _clean(meta.get("auditor", meta.get("audit_company", "北京万宁睿和医药科技有限公司")))
    subjects = "、".join(context.get("audited_subjects") or ["—"])
    if table_shape is not None:
        try:
            tbl = table_shape.Table
            for r in range(1, tbl.Rows.Count + 1):
                for c in range(1, tbl.Columns.Count + 1):
                    _set_cell(tbl.Cell(r, c), "", 16)
            fs = 16.0
            _set_cell(tbl.Cell(1, 1), "方案名称", fs, True, WHITE)
            _set_cell(tbl.Cell(1, 2), project, fs, True, BLACK)
            _set_cell(tbl.Cell(2, 1), "申办者", fs, True)
            _set_cell(tbl.Cell(2, 2), sponsor, fs)
            _set_cell(tbl.Cell(2, 3), "PI", fs, True, align=PP_ALIGN_CENTER)
            _set_cell(tbl.Cell(2, 4), pi, fs, True)
            _set_cell(tbl.Cell(3, 1), "中心名称", fs, True)
            _set_cell(tbl.Cell(3, 2), center, fs)
            _set_cell(tbl.Cell(3, 3), "中心入组\n情况", fs, True, align=PP_ALIGN_CENTER)
            _set_cell(tbl.Cell(3, 4), enrollment, fs)
            _set_cell(tbl.Cell(4, 1), "稽查时间", fs, True)
            _set_cell(tbl.Cell(4, 2), audit_date, fs)
            _set_cell(tbl.Cell(4, 3), "稽查员", fs, True, align=PP_ALIGN_CENTER)
            _set_cell(tbl.Cell(4, 4), auditor, fs)
            label = f"本次稽查{len(context.get('audited_subjects') or []) or 'x'}\n例受试者"
            _set_cell(tbl.Cell(5, 1), label, fs, True)
            _set_cell(tbl.Cell(5, 2), subjects, fs)
            return
        except Exception:
            pass
    _add_textbox(slide, 1.20, 1.45, 11.0, 4.5, f"方案名称：{project}\n申办者：{sponsor}\nPI：{pi}\n中心名称：{center}\n中心入组情况：{enrollment}\n稽查时间：{audit_date}\n稽查员：{auditor}\n本次稽查受试者：{subjects}", font_size=16)


def _render_counts(slide, context: dict):
    _delete_text_layer(slide, clear_tables=True)
    counts = context.get("summary", {})
    _add_textbox(slide, 0.72, 0.60, 6.60, 0.52, "二、中心稽查分类和数量", font_size=20, bold=True, color_rgb=BLACK, margin=0)
    y0, row_h = 1.72, 0.515
    for x0, cats in [(0.86, LEFT_CATS), (6.84, RIGHT_CATS)]:
        _add_textbox(slide, x0 + 0.12, y0 + 0.09, 3.60, 0.30, "分类", font_size=12, bold=True, color_rgb=WHITE, margin=0)
        _add_textbox(slide, x0 + 4.12, y0 + 0.09, 1.16, 0.30, "数量", font_size=12, bold=True, color_rgb=WHITE, align=PP_ALIGN_CENTER, margin=0)
        for i, cat in enumerate(cats):
            yy = y0 + 0.48 + i * row_h
            val = counts.get(cat, 0)
            _add_textbox(slide, x0 + 0.12, yy + 0.08, 3.78, 0.31, cat, font_size=12, color_rgb=BLACK, margin=0)
            _add_textbox(slide, x0 + 4.12, yy + 0.07, 1.16, 0.32, "—" if not val else str(val), font_size=12, bold=True, color_rgb=BLACK, align=PP_ALIGN_CENTER, margin=0)


def _render_issue(slide, cat: str, issue: dict, idx: int, total: int):
    _delete_text_layer(slide, clear_tables=True)
    sub_page = int(issue.get("_sub_page", 1))
    sub_total = int(issue.get("_sub_total", 1))
    page_tag = f"（{idx}/{total}）" if total > 1 else ""
    cont_tag = f" 续{sub_page}/{sub_total}" if sub_total > 1 else ""
    title = f"问题分类：{cat}{page_tag}{cont_tag}"
    basis = _clean(issue.get("basis") or "—")
    desc = _clean(issue.get("description") or "—")
    _add_textbox(slide, 0.64, 0.48, 11.30, 0.58, title, font_size=20, bold=True, color_rgb=BLACK, margin=0)
    _add_textbox(slide, 0.75, 1.76, 0.62, 0.38, "依据", font_size=12, bold=True, color_rgb=BLACK, align=PP_ALIGN_CENTER, margin=0)
    _add_textbox(slide, 1.58, 1.60, 10.48, 1.42, basis, font_size=12, color_rgb=BLACK, margin=4)
    _add_textbox(slide, 0.75, 4.00, 0.62, 0.38, "描述", font_size=12, bold=True, color_rgb=BLACK, align=PP_ALIGN_CENTER, margin=0)
    _add_textbox(slide, 1.58, 3.74, 10.48, 2.80, desc, font_size=12, color_rgb=BLACK, margin=4)


def _render_suggestion(slide, text: str):
    _delete_text_layer(slide, clear_tables=False)
    _add_textbox(slide, 0.58, 0.46, 5.20, 0.70, "建议项：", font_size=28, bold=True, color_rgb=BLACK, margin=0)
    table_shape = _find_largest_table(slide)
    if table_shape is not None:
        try:
            tbl = table_shape.Table
            for r in range(1, tbl.Rows.Count + 1):
                for c in range(1, tbl.Columns.Count + 1):
                    _set_cell(tbl.Cell(r, c), "", 14)
            _set_cell(tbl.Cell(1, 1), "描述", 14, True, BLACK, PP_ALIGN_CENTER)
            _set_cell(tbl.Cell(1, 2), text or "—", 14, False, BLACK, PP_ALIGN_LEFT)
            return
        except Exception:
            pass
    _add_textbox(slide, 1.24, 1.43, 10.88, 5.05, text or "—", font_size=14, color_rgb=BLACK, margin=6)


def render_ppt(context: dict, output_path: str | Path):
    template_path = Path(TEMPLATE_PATH).resolve()
    output_path = Path(output_path).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if not template_path.exists() or template_path.stat().st_size < 1024:
        raise FileNotFoundError(f"未找到有效PPT模板：{template_path}。请上传正式 assets/template.pptx。")

    app = _get_powerpoint_app()
    prs = None
    try:
        prs = app.Presentations.Open(str(template_path), WithWindow=MSO_FALSE)
        original_count = prs.Slides.Count

        slide = _duplicate_slide_to_end(prs, SLIDE_COVER)
        _render_cover(slide, context)
        _duplicate_slide_to_end(prs, SLIDE_THANKS)
        _duplicate_slide_to_end(prs, SLIDE_TOC)
        _duplicate_slide_to_end(prs, SLIDE_PART1)
        slide = _duplicate_slide_to_end(prs, SLIDE_OVERVIEW)
        _render_overview(slide, context)
        _duplicate_slide_to_end(prs, SLIDE_SCOPE)
        _duplicate_slide_to_end(prs, SLIDE_PART2)
        slide = _duplicate_slide_to_end(prs, SLIDE_COUNTS)
        _render_counts(slide, context)

        for cat in context.get("standard_categories", []):
            cat_issues = [x for x in context.get("issues", []) if x.get("category") == cat]
            if not cat_issues:
                continue
            template_slide = CAT_TO_TEMPLATE_SLIDE.get(cat, SLIDE_SUGGESTION)
            for i, issue in enumerate(cat_issues, start=1):
                for page_issue in _paginate_issue(issue):
                    slide = _duplicate_slide_to_end(prs, template_slide)
                    _render_issue(slide, cat, page_issue, i, len(cat_issues))

        sug_items = context.get("suggestions", [])
        sug_text = "\n\n".join([f"【{s.get('category', '其他')}】{s.get('text', '')}" for s in sug_items if s.get("text")]) or "—"
        for page in _split_text(sug_text, 650):
            slide = _duplicate_slide_to_end(prs, SLIDE_SUGGESTION)
            _render_suggestion(slide, page)

        _duplicate_slide_to_end(prs, SLIDE_ENDING)
        _delete_original_template_slides(prs, original_count)
        prs.SaveAs(str(output_path), PPSAVEAS_OPENXML_PRESENTATION)
        return output_path
    finally:
        try:
            if prs is not None:
                prs.Close()
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass
