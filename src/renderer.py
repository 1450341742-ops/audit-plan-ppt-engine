from __future__ import annotations
from copy import deepcopy
from pathlib import Path
from typing import Any
import re

from pptx import Presentation
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor

BASE_DIR = Path(__file__).resolve().parent.parent
DEFAULT_TEMPLATE_PATH = BASE_DIR / "assets" / "template.pptx"

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

WHITE = RGBColor(255, 255, 255)
YELLOW = RGBColor(255, 242, 0)
BLACK = RGBColor(0, 0, 0)

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

PLACEHOLDER_TEXTS = [
    "单击此处编辑标题",
    "单击此处编辑副标题",
    "Click to edit Master title style",
    "Click to edit Master subtitle style",
]


def _clean(v: Any) -> str:
    s = str(v or "").replace("\r", "\n")
    s = re.sub(r"[ \t\xa0]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def _delete_slide(prs: Presentation, index: int = 0) -> None:
    r_id = prs.slides._sldIdLst[index].rId
    prs.part.drop_rel(r_id)
    del prs.slides._sldIdLst[index]


def _remove_shape_xml(shape) -> None:
    try:
        shape.element.getparent().remove(shape.element)
    except Exception:
        pass


def _text(shape) -> str:
    if getattr(shape, "has_text_frame", False):
        return _clean(shape.text)
    return ""


def _shape_is_placeholder(shape) -> bool:
    try:
        if getattr(shape, "is_placeholder", False):
            return True
    except Exception:
        pass
    try:
        if shape.element.xpath(".//p:ph"):
            return True
    except Exception:
        pass
    return False


def _is_placeholder_text(text: str) -> bool:
    return any(p in _clean(text) for p in PLACEHOLDER_TEXTS)


def _shape_is_auto_placeholder(shape) -> bool:
    return getattr(shape, "has_text_frame", False) and _shape_is_placeholder(shape) and _is_placeholder_text(_text(shape))


def _copy_rels(source, dest) -> dict[str, str]:
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
    return rel_map


def _remap_relationship_ids(xml_el, rel_map: dict[str, str]) -> None:
    for el in xml_el.iter():
        for attr, val in list(el.attrib.items()):
            if val in rel_map:
                el.attrib[attr] = rel_map[val]


def _copy_slide(prs: Presentation, src_no: int):
    """复制模板页并严格保留该页版式背景。

    关键修复：
    1. 使用 source.slide_layout，保留模板母版/版式里的背景和Logo；
    2. 删除 add_slide 自动生成的标题/副标题占位符；
    3. 只复制源页自身 shapes，跳过“单击此处编辑标题/副标题”；
    4. 不再把 10-21 的空模板页全部生成，只使用实际有表格的模板页生成问题页。
    """
    source = prs.slides[src_no - 1]
    dest = prs.slides.add_slide(source.slide_layout)
    for shp in list(dest.shapes):
        _remove_shape_xml(shp)
    rel_map = _copy_rels(source, dest)
    for shape in source.shapes:
        try:
            if _shape_is_auto_placeholder(shape):
                continue
            new_el = deepcopy(shape.element)
            _remap_relationship_ids(new_el, rel_map)
            dest.shapes._spTree.insert_element_before(new_el, "p:extLst")
        except Exception:
            pass
    return dest


def _remove_original_template_slides(prs: Presentation, original_count: int) -> None:
    for _ in range(original_count):
        _delete_slide(prs, 0)


def _tables(slide):
    return [shp.table for shp in slide.shapes if getattr(shp, "has_table", False)]


def _slide_text(slide) -> str:
    return "\n".join(_text(shp) for shp in slide.shapes if getattr(shp, "has_text_frame", False))


def _find_issue_template_slide(prs: Presentation) -> int:
    """自动寻找真正的问题页模板，避免复制空白占位页。

    之前正文大量空白页，是因为按 10、11、12... 逐分类复制时，部分模板页只是“单击此处编辑标题”的占位页。
    现在从 10-21 页中优先找带表格或含“依据/描述/问题分类”的页面，统一作为问题明细页模板。
    """
    candidates = []
    for no in range(10, min(21, len(prs.slides)) + 1):
        slide = prs.slides[no - 1]
        text = _slide_text(slide)
        has_issue_words = any(k in text for k in ["依据", "描述", "问题分类"])
        has_table = bool(_tables(slide))
        if has_table or has_issue_words:
            score = (10 if has_table else 0) + (5 if has_issue_words else 0) - (3 if _is_placeholder_text(text) else 0)
            candidates.append((score, no))
    if candidates:
        return sorted(candidates, reverse=True)[0][1]
    return min(21, len(prs.slides))


def _remove_template_empty_slides(prs: Presentation) -> None:
    """兜底删除生成后仍残留的空占位页。"""
    for idx in range(len(prs.slides) - 1, -1, -1):
        slide = prs.slides[idx]
        text = _slide_text(slide)
        if _is_placeholder_text(text):
            _delete_slide(prs, idx)
            continue
        # 没有表格、没有业务文字、只有背景/图片的页面，也删除；封面/目录/章节页都有文字，不受影响。
        meaningful = re.sub(r"[\s\-—_：:|]+", "", text)
        if not meaningful and not _tables(slide):
            _delete_slide(prs, idx)


def _remove_placeholder_text_shapes(slide) -> None:
    for shp in list(slide.shapes):
        try:
            if getattr(shp, "has_table", False):
                continue
            if _shape_is_auto_placeholder(shp) or _is_placeholder_text(_text(shp)):
                _remove_shape_xml(shp)
        except Exception:
            pass


def _remove_text_shapes(slide, keep_keywords: tuple[str, ...] = ()) -> None:
    for shp in list(slide.shapes):
        try:
            if getattr(shp, "has_table", False):
                continue
            if getattr(shp, "has_text_frame", False):
                t = _text(shp)
                if keep_keywords and any(k in t for k in keep_keywords):
                    continue
                _remove_shape_xml(shp)
        except Exception:
            pass


def _add_textbox(slide, x: float, y: float, w: float, h: float, text: str,
                 font_size: int, bold: bool, color: RGBColor,
                 align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(0.03)
    tf.margin_right = Inches(0.03)
    tf.margin_top = Inches(0.01)
    tf.margin_bottom = Inches(0.01)
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = _clean(text) or "—"
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    return box


def _set_shape_text(shape, text: str, font_size: int | None = None, bold: bool | None = None,
                    align=PP_ALIGN.LEFT, color: RGBColor | None = None) -> None:
    if not getattr(shape, "has_text_frame", False):
        return
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = _clean(text) or "—"
    run.font.name = "Microsoft YaHei"
    if font_size:
        run.font.size = Pt(font_size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


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


def _replace_first_shape_containing(slide, keywords: list[str], text: str,
                                    font_size: int | None = None, bold: bool | None = None,
                                    align=PP_ALIGN.LEFT, color: RGBColor | None = None) -> bool:
    for shp in slide.shapes:
        t = _text(shp)
        if t and any(k in t for k in keywords):
            _set_shape_text(shp, text, font_size=font_size, bold=bold, align=align, color=color)
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


def _meaningful_text(value: Any) -> bool:
    t = _clean(value)
    if not t or t in {"—", "-", "无", "NA", "N/A"}:
        return False
    if any(p in t for p in PLACEHOLDER_TEXTS):
        return False
    return len(re.sub(r"[\s\-—_：:|]+", "", t)) > 0


def _paginate_issue(issue: dict) -> list[dict]:
    basis = issue.get("basis", "—")
    desc = issue.get("description", "—")
    if not _meaningful_text(basis) and not _meaningful_text(desc):
        return []
    basis_parts = _split_text(basis, 520) if _meaningful_text(basis) else ["—"]
    desc_parts = _split_text(desc, 900) if _meaningful_text(desc) else ["—"]
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


def _has_issue_content(issue: dict) -> bool:
    return _meaningful_text(issue.get("basis", "")) or _meaningful_text(issue.get("description", ""))


def _render_cover(slide, context: dict) -> None:
    _remove_text_shapes(slide)
    meta = context.get("meta", {})
    project = _clean(meta.get("project_name", "—"))
    center = _clean(meta.get("center_name", "—"))
    center_no = _clean(meta.get("center_no", ""))
    audit_date = _clean(meta.get("audit_date", "—"))
    title = f"{project}-{center}" + (f"（中心编号{center_no}）" if center_no else "")
    _add_textbox(slide, 0.25, 3.25, 12.70, 0.75, title, 28, True, WHITE, PP_ALIGN.LEFT)
    _add_textbox(slide, 0.25, 4.12, 12.70, 0.55, "中心稽查末次会议", 28, True, YELLOW, PP_ALIGN.LEFT)
    _add_textbox(slide, 0.35, 5.13, 12.30, 0.35, f"时间：{audit_date}", 14, True, WHITE, PP_ALIGN.LEFT)
    _add_textbox(slide, 0.35, 5.72, 12.30, 0.35, "北京万宁睿和医药科技有限公司", 14, True, WHITE, PP_ALIGN.LEFT)


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
    title = f"问题分类:{category}{tag}{cont}"
    basis = _clean(issue.get("basis", "—"))
    desc = _clean(issue.get("description", "—"))

    _remove_placeholder_text_shapes(slide)
    ok = _replace_first_shape_containing(slide, ["问题分类"], title, 20, True, PP_ALIGN.LEFT, BLACK)
    tables = _tables(slide)
    wrote_table = False
    if tables:
        tbl = tables[0]
        if len(tbl.rows) >= 2 and len(tbl.columns) >= 2:
            _set_cell(tbl.cell(0, 0), "依据", 12, True, PP_ALIGN.CENTER)
            _set_cell(tbl.cell(0, 1), basis, 12, False, PP_ALIGN.LEFT)
            _set_cell(tbl.cell(1, 0), "描述", 12, True, PP_ALIGN.CENTER)
            _set_cell(tbl.cell(1, 1), desc, 12, False, PP_ALIGN.LEFT)
            wrote_table = True
    if not ok:
        _add_textbox(slide, 0.65, 0.48, 11.4, 0.50, title, 20, True, BLACK, PP_ALIGN.LEFT)
    if not wrote_table:
        _add_textbox(slide, 0.75, 1.70, 0.70, 0.30, "依据", 12, True, BLACK, PP_ALIGN.CENTER)
        _add_textbox(slide, 1.58, 1.48, 10.55, 1.45, basis, 12, False, BLACK, PP_ALIGN.LEFT)
        _add_textbox(slide, 0.75, 4.00, 0.70, 0.30, "描述", 12, True, BLACK, PP_ALIGN.CENTER)
        _add_textbox(slide, 1.58, 3.70, 10.55, 2.75, desc, 12, False, BLACK, PP_ALIGN.LEFT)


def _render_suggestion(slide, text: str) -> None:
    _remove_placeholder_text_shapes(slide)
    _replace_first_shape_containing(slide, ["建议项"], "建议项：", 28, True, PP_ALIGN.LEFT, BLACK)
    tables = _tables(slide)
    wrote = False
    if tables:
        tbl = tables[0]
        if len(tbl.rows) >= 1 and len(tbl.columns) >= 2:
            _set_cell(tbl.cell(0, 0), "描述", 14, True, PP_ALIGN.CENTER)
            _set_cell(tbl.cell(0, 1), text or "—", 14, False, PP_ALIGN.LEFT)
            wrote = True
    if not wrote:
        _add_textbox(slide, 1.20, 1.45, 10.9, 5.0, text or "—", 14, False, BLACK, PP_ALIGN.LEFT)


def render_ppt(context: dict, output_path: str | Path, template_path: str | Path | None = None):
    template = Path(template_path or DEFAULT_TEMPLATE_PATH)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if not template.exists() or template.stat().st_size < 1024 * 100:
        raise FileNotFoundError(f"未找到有效PPT模板：{template}。请在页面先上传你的稽查总结会模板。")

    prs = Presentation(str(template))
    original_count = len(prs.slides)
    if original_count < 23:
        raise RuntimeError(f"模板页数不足：当前 {original_count} 页，至少需要 23 页。请上传完整的稽查总结会模板。")

    issue_template_no = _find_issue_template_slide(prs)

    slide = _copy_slide(prs, SLIDE_COVER); _render_cover(slide, context)
    _copy_slide(prs, SLIDE_THANKS)
    _copy_slide(prs, SLIDE_TOC)
    _copy_slide(prs, SLIDE_PART1)
    slide = _copy_slide(prs, SLIDE_OVERVIEW); _render_overview(slide, context)
    _copy_slide(prs, SLIDE_SCOPE)
    _copy_slide(prs, SLIDE_PART2)
    slide = _copy_slide(prs, SLIDE_COUNTS); _render_counts(slide, context)

    for cat in context.get("standard_categories", []):
        cat_issues = [x for x in context.get("issues", []) if x.get("category") == cat and _has_issue_content(x)]
        if not cat_issues:
            continue
        for i, issue in enumerate(cat_issues, start=1):
            for page_issue in _paginate_issue(issue):
                if not _has_issue_content(page_issue):
                    continue
                slide = _copy_slide(prs, issue_template_no)
                _render_issue(slide, cat, page_issue, i, len(cat_issues))

    sug_items = context.get("suggestions", [])
    sug_text = "\n\n".join([f"【{s.get('category', '其他')}】{s.get('text', '')}" for s in sug_items if s.get("text")]) or "—"
    for text_page in _split_text(sug_text, 650):
        if _clean(text_page) and _clean(text_page) != "—":
            slide = _copy_slide(prs, SLIDE_SUGGESTION)
            _render_suggestion(slide, text_page)

    _copy_slide(prs, SLIDE_ENDING)
    _remove_original_template_slides(prs, original_count)
    _remove_template_empty_slides(prs)
    prs.save(str(output_path))
    return output_path
