from __future__ import annotations
from copy import deepcopy
from pathlib import Path
from typing import Any
import re

from pptx import Presentation
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Pt

BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = BASE_DIR / "assets" / "template.pptx"

# 模板页序号，严格对应用户上传的“稽查总结会模板”
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


def _clean(v: Any) -> str:
    s = str(v or "").replace("\r", "\n")
    s = re.sub(r"[ \t\xa0]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def _delete_slide(prs: Presentation, index: int = 0) -> None:
    r_id = prs.slides._sldIdLst[index].rId
    prs.part.drop_rel(r_id)
    del prs.slides._sldIdLst[index]


def _copy_slide(prs: Presentation, src_no: int):
    """复制模板页，保留图片、Logo、背景、表格、形状和版式。

    python-pptx 没有官方 duplicate slide API；这里使用 OOXML 深拷贝形状和关系。
    不调用 PowerPoint/pywin32，因此可在 Streamlit Cloud、Windows、Mac 运行。
    """
    source = prs.slides[src_no - 1]
    dest = prs.slides.add_slide(source.slide_layout)
    for shp in list(dest.shapes):
        shp.element.getparent().remove(shp.element)

    rel_map = {}
    for r_id, rel in source.part.rels.items():
        if "slideLayout" in rel.reltype:
            continue
        try:
            if rel.is_external:
                new_rid = dest.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
            else:
                new_rid = dest.part.rels.get_or_add(rel.reltype, rel._target)
            rel_map[r_id] = new_rid
        except Exception:
            pass

    for shape in source.shapes:
        new_el = deepcopy(shape.element)
        for el in new_el.iter():
            for attr, val in list(el.attrib.items()):
                if val in rel_map:
                    el.attrib[attr] = rel_map[val]
        dest.shapes._spTree.insert_element_before(new_el, "p:extLst")
    return dest


def _remove_original_template_slides(prs: Presentation, original_count: int) -> None:
    for _ in range(original_count):
        _delete_slide(prs, 0)


def _text(shape) -> str:
    if getattr(shape, "has_text_frame", False):
        return _clean(shape.text)
    return ""


def _set_shape_text(shape, text: str, font_size: int | None = None, bold: bool | None = None,
                    align=PP_ALIGN.LEFT) -> None:
    if not getattr(shape, "has_text_frame", False):
        return
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = _clean(text) or "—"
    if font_size:
        run.font.size = Pt(font_size)
    if bold is not None:
        run.font.bold = bold
    run.font.name = "Microsoft YaHei"


def _set_cell(cell, text: str, font_size: int = 12, bold: bool | None = None,
              align=PP_ALIGN.LEFT) -> None:
    cell.text = _clean(text) or "—"
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = Pt(4)
    cell.margin_right = Pt(4)
    cell.margin_top = Pt(2)
    cell.margin_bottom = Pt(2)
    for p in cell.text_frame.paragraphs:
        p.alignment = align
        for r in p.runs:
            r.font.name = "Microsoft YaHei"
            r.font.size = Pt(font_size)
            if bold is not None:
                r.font.bold = bold


def _tables(slide):
    return [shp.table for shp in slide.shapes if getattr(shp, "has_table", False)]


def _replace_first_shape_containing(slide, keywords: list[str], text: str,
                                    font_size: int | None = None, bold: bool | None = None,
                                    align=PP_ALIGN.LEFT) -> bool:
    for shp in slide.shapes:
        t = _text(shp)
        if t and any(k in t for k in keywords):
            _set_shape_text(shp, text, font_size=font_size, bold=bold, align=align)
            return True
    return False


def _split_text(text: str, limit: int) -> list[str]:
    text = _clean(text)
    if not text:
        return ["—"]
    out, buf = [], ""
    for part in re.split(r"\n\s*\n", text):
        part = part.strip()
        if not part:
            continue
        if len(buf) + len(part) + 2 <= limit:
            buf = (buf + "\n\n" + part).strip()
        else:
            if buf:
                out.append(buf)
            while len(part) > limit:
                out.append(part[:limit])
                part = part[limit:]
            buf = part
    if buf:
        out.append(buf)
    return out or ["—"]


def _paginate_issue(issue: dict) -> list[dict]:
    basis_parts = _split_text(issue.get("basis", "—"), 520)
    desc_parts = _split_text(issue.get("description", "—"), 900)
    total = max(len(basis_parts), len(desc_parts))
    pages = []
    for i in range(total):
        x = dict(issue)
        x["basis"] = basis_parts[i] if i < len(basis_parts) else "—"
        x["description"] = desc_parts[i] if i < len(desc_parts) else "—"
        x["_sub_page"] = i + 1
        x["_sub_total"] = total
        pages.append(x)
    return pages


def _render_cover(slide, context: dict) -> None:
    meta = context.get("meta", {})
    project = _clean(meta.get("project_name", "—"))
    center = _clean(meta.get("center_name", "—"))
    center_no = _clean(meta.get("center_no", ""))
    audit_date = _clean(meta.get("audit_date", "—"))
    title = f"{project}-{center}" + (f"（中心编号{center_no}）" if center_no else "")
    _replace_first_shape_containing(slide, ["**项目", "中心稽查末次会议"], f"{title}\n中心稽查末次会议", 24, True, PP_ALIGN.CENTER)
    _replace_first_shape_containing(slide, ["时间：", "北京万宁睿和医药科技有限公司"], f"时间：{audit_date}\n北京万宁睿和医药科技有限公司", 16, None, PP_ALIGN.CENTER)


def _render_overview(slide, context: dict) -> None:
    meta = context.get("meta", {})
    tables = _tables(slide)
    if not tables:
        return
    tbl = tables[0]
    project = _clean(meta.get("project_name", "—"))
    sponsor = _clean(meta.get("sponsor", "—"))
    pi = _clean(meta.get("pi", "—"))
    center = _clean(meta.get("center_name", "—"))
    enrollment = _clean(meta.get("enrollment", "—"))
    audit_date = _clean(meta.get("audit_date", "—"))
    auditor = _clean(meta.get("auditor", meta.get("audit_company", "北京万宁睿和医药科技有限公司")))
    subjects = "、".join(context.get("audited_subjects") or ["—"])
    rows = [
        [(0, 0, "方案名称", 16, True), (0, 1, project, 16, False)],
        [(1, 0, "申办者", 16, True), (1, 1, sponsor, 16, False), (1, 2, "PI", 16, True), (1, 3, pi, 16, False)],
        [(2, 0, "中心名称", 16, True), (2, 1, center, 16, False), (2, 2, "中心入组\n情况", 16, True), (2, 3, enrollment, 16, False)],
        [(3, 0, "稽查时间", 16, True), (3, 1, audit_date, 16, False), (3, 2, "稽查员", 16, True), (3, 3, auditor, 16, False)],
        [(4, 0, f"本次稽查{len(context.get('audited_subjects') or []) or 'x'}\n例受试者", 16, True), (4, 1, subjects, 16, False)],
    ]
    for group in rows:
        for r, c, text, size, bold in group:
            if r < len(tbl.rows) and c < len(tbl.columns):
                _set_cell(tbl.cell(r, c), text, size, bold, PP_ALIGN.CENTER if c in (0, 2) else PP_ALIGN.LEFT)


def _render_counts(slide, context: dict) -> None:
    counts = context.get("summary", {})
    tables = _tables(slide)
    for tbl, cats in zip(tables[:2], [LEFT_CATS, RIGHT_CATS]):
        _set_cell(tbl.cell(0, 0), "分类", 12, True, PP_ALIGN.CENTER)
        _set_cell(tbl.cell(0, 1), "数量", 12, True, PP_ALIGN.CENTER)
        for i, cat in enumerate(cats, start=1):
            if i >= len(tbl.rows):
                break
            _set_cell(tbl.cell(i, 0), cat, 12, False, PP_ALIGN.LEFT)
            _set_cell(tbl.cell(i, 1), str(counts.get(cat, 0) or "—"), 12, True, PP_ALIGN.CENTER)


def _render_issue(slide, category: str, issue: dict, idx: int, total: int) -> None:
    sub_page = int(issue.get("_sub_page", 1))
    sub_total = int(issue.get("_sub_total", 1))
    tag = f"（{idx}/{total}）" if total > 1 else ""
    cont = f" 续{sub_page}/{sub_total}" if sub_total > 1 else ""
    _replace_first_shape_containing(slide, ["问题分类"], f"问题分类:{category}{tag}{cont}", 20, True, PP_ALIGN.LEFT)
    tables = _tables(slide)
    if not tables:
        return
    tbl = tables[0]
    if len(tbl.rows) >= 2 and len(tbl.columns) >= 2:
        _set_cell(tbl.cell(0, 0), "依据", 12, True, PP_ALIGN.CENTER)
        _set_cell(tbl.cell(0, 1), issue.get("basis", "—"), 12, False, PP_ALIGN.LEFT)
        _set_cell(tbl.cell(1, 0), "描述", 12, True, PP_ALIGN.CENTER)
        _set_cell(tbl.cell(1, 1), issue.get("description", "—"), 12, False, PP_ALIGN.LEFT)


def _render_suggestion(slide, text: str) -> None:
    _replace_first_shape_containing(slide, ["建议项"], "建议项：", 28, True, PP_ALIGN.LEFT)
    tables = _tables(slide)
    if tables:
        tbl = tables[0]
        if len(tbl.rows) >= 1 and len(tbl.columns) >= 2:
            _set_cell(tbl.cell(0, 0), "描述", 14, True, PP_ALIGN.CENTER)
            _set_cell(tbl.cell(0, 1), text or "—", 14, False, PP_ALIGN.LEFT)


def render_ppt(context: dict, output_path: str | Path):
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if not TEMPLATE_PATH.exists() or TEMPLATE_PATH.stat().st_size < 1024:
        raise FileNotFoundError(f"未找到有效PPT模板：{TEMPLATE_PATH}，请上传 assets/template.pptx。")

    prs = Presentation(str(TEMPLATE_PATH))
    original_count = len(prs.slides)
    if original_count < 23:
        raise RuntimeError(f"模板页数不足：当前 {original_count} 页，至少需要 23 页。")

    slide = _copy_slide(prs, SLIDE_COVER); _render_cover(slide, context)
    _copy_slide(prs, SLIDE_THANKS)
    _copy_slide(prs, SLIDE_TOC)
    _copy_slide(prs, SLIDE_PART1)
    slide = _copy_slide(prs, SLIDE_OVERVIEW); _render_overview(slide, context)
    _copy_slide(prs, SLIDE_SCOPE)
    _copy_slide(prs, SLIDE_PART2)
    slide = _copy_slide(prs, SLIDE_COUNTS); _render_counts(slide, context)

    for cat in context.get("standard_categories", []):
        cat_issues = [x for x in context.get("issues", []) if x.get("category") == cat]
        if not cat_issues:
            continue
        template_no = CAT_TO_TEMPLATE_SLIDE.get(cat, SLIDE_SUGGESTION)
        for i, issue in enumerate(cat_issues, start=1):
            for page_issue in _paginate_issue(issue):
                slide = _copy_slide(prs, template_no)
                _render_issue(slide, cat, page_issue, i, len(cat_issues))

    sug_items = context.get("suggestions", [])
    sug_text = "\n\n".join([f"【{s.get('category', '其他')}】{s.get('text', '')}" for s in sug_items if s.get("text")]) or "—"
    for text_page in _split_text(sug_text, 650):
        slide = _copy_slide(prs, SLIDE_SUGGESTION)
        _render_suggestion(slide, text_page)

    _copy_slide(prs, SLIDE_ENDING)
    _remove_original_template_slides(prs, original_count)
    prs.save(str(output_path))
    return output_path
