from __future__ import annotations
from pathlib import Path
from typing import Any
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = BASE_DIR / "assets" / "template.pptx"


def safe(v: Any) -> str:
    return "" if v is None else str(v).strip()


def join_subjects(items: list[str]) -> str:
    return "、".join([x for x in items if x]) or "—"


def blank_layout(prs: Presentation):
    return prs.slide_layouts[6]


def delete_all_slides(prs: Presentation):
    while len(prs.slides) > 0:
        r_id = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(r_id)
        del prs.slides._sldIdLst[0]


def new_presentation() -> Presentation:
    if TEMPLATE_PATH.exists():
        try:
            prs = Presentation(str(TEMPLATE_PATH))
            delete_all_slides(prs)
            return prs
        except Exception:
            pass
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs


NAVY = RGBColor(25, 55, 109)
BLUE = RGBColor(235, 243, 252)
WHITE = RGBColor(255, 255, 255)
TEXT = RGBColor(38, 48, 65)
GRAY = RGBColor(120, 130, 150)


def set_run_font(run, size=12, bold=False, color=TEXT):
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def set_paragraph_text(p, text: str, size=12, bold=False, color=TEXT, align=None):
    p.clear()
    if align is not None:
        p.alignment = align
    run = p.add_run()
    run.text = safe(text)
    set_run_font(run, size, bold, color)
    return run


def add_bg(slide, title: str, page_no: int | None = None, title_size: int = 28):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = WHITE
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.16))
    bar.fill.solid(); bar.fill.fore_color.rgb = NAVY
    bar.line.fill.background()
    tx = slide.shapes.add_textbox(Inches(0.42), Inches(0.33), Inches(12.3), Inches(0.55))
    p = tx.text_frame.paragraphs[0]
    set_paragraph_text(p, title, title_size, True, NAVY)
    foot = slide.shapes.add_textbox(Inches(0.45), Inches(7.1), Inches(12.4), Inches(0.2))
    fp = foot.text_frame.paragraphs[0]
    fp_text = "北京万宁睿和医药科技有限公司" + (f"  |  {page_no}" if page_no else "")
    set_paragraph_text(fp, fp_text, 8, False, GRAY, PP_ALIGN.RIGHT)


def set_cell(cell, text: str, font_size=12, bold=False, fill=None, align=PP_ALIGN.LEFT):
    cell.text = safe(text)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    if fill is not None:
        cell.fill.solid(); cell.fill.fore_color.rgb = fill
    for p in cell.text_frame.paragraphs:
        p.alignment = align
        if not p.runs:
            run = p.add_run()
            run.text = safe(text)
            set_run_font(run, font_size, bold, WHITE if fill == NAVY else TEXT)
        else:
            for r in p.runs:
                set_run_font(r, font_size, bold, WHITE if fill == NAVY else TEXT)


def add_table(slide, data, x, y, w, h, font_size=12, first_col_w=2.0, header=False):
    if not data:
        data = [["—"]]
    rows = len(data)
    cols = max(len(r) for r in data)
    shape = slide.shapes.add_table(rows, cols, Inches(x), Inches(y), Inches(w), Inches(h))
    tbl = shape.table
    if cols >= 2:
        tbl.columns[0].width = Inches(first_col_w)
        tbl.columns[1].width = Inches(max(0.5, w - first_col_w))
    for i, row in enumerate(data):
        for j in range(cols):
            val = row[j] if j < len(row) else ""
            is_header = header and i == 0
            fill = NAVY if is_header else (BLUE if j == 0 else None)
            align = PP_ALIGN.CENTER if j == 0 or is_header else PP_ALIGN.LEFT
            set_cell(tbl.cell(i, j), val, font_size, bold=(j == 0 or is_header), fill=fill, align=align)
            tbl.cell(i, j).margin_left = Inches(0.06)
            tbl.cell(i, j).margin_right = Inches(0.06)
            tbl.cell(i, j).margin_top = Inches(0.03)
            tbl.cell(i, j).margin_bottom = Inches(0.03)
    return tbl


def add_text(slide, x, y, w, h, text, size=14, bold=False, color=TEXT, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    set_paragraph_text(p, text, size, bold, color, align)
    return box


def cover(prs, data, page):
    slide = prs.slides.add_slide(blank_layout(prs))
    add_bg(slide, "", page)
    meta = data.get("meta", {})
    add_text(slide, 0.75, 1.65, 12, 0.65, "中心稽查总结会", 34, True, NAVY, PP_ALIGN.CENTER)
    add_text(slide, 1.0, 2.55, 11.3, 0.45, meta.get("project_name", "—"), 16, False, TEXT, PP_ALIGN.CENTER)
    add_text(slide, 1.0, 3.05, 11.3, 0.35, f"方案编号：{meta.get('protocol_no','—')}", 14, False, TEXT, PP_ALIGN.CENTER)
    add_text(slide, 1.0, 3.5, 11.3, 0.35, f"中心：{meta.get('center_name','—')}", 14, False, TEXT, PP_ALIGN.CENTER)
    add_text(slide, 1.0, 4.15, 11.3, 0.35, "北京万宁睿和医药科技有限公司", 15, True, NAVY, PP_ALIGN.CENTER)
    return page + 1


def overview(prs, data, page):
    slide = prs.slides.add_slide(blank_layout(prs))
    add_bg(slide, "中心稽查概述", page, 28)
    meta = data.get("meta", {})
    rows = [
        ["项目名称", meta.get("project_name", "—")],
        ["方案编号", meta.get("protocol_no", "—")],
        ["研究中心", meta.get("center_name", "—")],
        ["中心编号", meta.get("center_no", "—") or "—"],
        ["主要研究者", meta.get("pi", "—")],
        ["稽查日期", meta.get("audit_date", "—")],
        ["受试者", join_subjects(data.get("audited_subjects", []))],
    ]
    add_table(slide, rows, 0.62, 1.15, 12.1, 5.15, 16, 2.1)
    return page + 1


def counts(prs, data, page):
    slide = prs.slides.add_slide(blank_layout(prs))
    add_bg(slide, "中心稽查分类和数量", page, 20)
    rows = [["问题分类", "数量"]] + [[c, str(data.get("summary", {}).get(c, 0))] for c in data.get("standard_categories", [])]
    add_table(slide, rows, 0.62, 1.0, 12.1, 5.9, 12, 10.5, header=True)
    return page + 1


def issue_slide(prs, data, page, issue, idx, total):
    slide = prs.slides.add_slide(blank_layout(prs))
    add_bg(slide, f"问题明细 {idx}/{total}", page, 20)
    rows = [
        ["问题分类", issue.get("category", "—")],
        ["问题标题", issue.get("title", "—")],
        ["问题级别", issue.get("severity", "—")],
        ["涉及受试者", join_subjects(issue.get("subject_ids", []))],
        ["依据", issue.get("basis", "—")],
        ["描述", issue.get("description", "—")],
    ]
    add_table(slide, rows, 0.45, 1.0, 12.45, 5.9, 12, 1.55)
    return page + 1


def suggestions(prs, data, page):
    slide = prs.slides.add_slide(blank_layout(prs))
    add_bg(slide, "建议项", page, 28)
    rows = [["分类", "建议内容"]] + [[s.get("category", "—"), s.get("text", "—")] for s in data.get("suggestions", [])[:18]]
    add_table(slide, rows, 0.62, 1.0, 12.1, 5.9, 14, 3.1, header=True)
    return page + 1


def ending(prs, data, page):
    slide = prs.slides.add_slide(blank_layout(prs))
    add_bg(slide, "", page)
    add_text(slide, 1.0, 2.7, 11.3, 0.65, "THANKS", 38, True, NAVY, PP_ALIGN.CENTER)
    add_text(slide, 1.0, 3.45, 11.3, 0.45, "提升质量，赋能上市", 20, True, TEXT, PP_ALIGN.CENTER)
    return page + 1


def render_ppt(context: dict, output_path: str | Path):
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs = new_presentation()
    page = 1
    page = cover(prs, context, page)
    page = overview(prs, context, page)
    page = counts(prs, context, page)
    issues = context.get("issues", [])
    for i, issue in enumerate(issues, start=1):
        page = issue_slide(prs, context, page, issue, i, len(issues))
    page = suggestions(prs, context, page)
    ending(prs, context, page)
    prs.save(str(output_path))
    return output_path
