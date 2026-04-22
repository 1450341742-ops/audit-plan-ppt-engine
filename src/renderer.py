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

# 新版模板：共性问题、个性问题、Q&A 只复制模板页，不写入内容。
# Q&A 页不再固定写死页码，优先按页面文字自动识别，避免模板页数调整后漏加。
SLIDE_COVER = 2
SLIDE_THANKS = 3
SLIDE_TOC = 4
SLIDE_PART1 = 5
SLIDE_OVERVIEW = 6
SLIDE_SCOPE = 7
SLIDE_COMMON = 8
SLIDE_INDIVIDUAL = 9
SLIDE_PART2 = 10
SLIDE_COUNTS = 11
SLIDE_ISSUE = 12

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


def _is_placeholder_text(text: str) -> bool:
    return any(p in _clean(text) for p in PLACEHOLDER_TEXTS)


def _copy_rels(src_part, dest_part) -> dict[str, str]:
    rel_map = {}
    for r_id, rel in src_part.rels.items():
        if "slideLayout" in rel.reltype:
            continue
        try:
            if rel.is_external:
                new_rid = dest_part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
            else:
                new_rid = dest_part.rels.get_or_add(rel.reltype, rel._target)
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
    """按模板页复制：保留源页版式背景，删除自动占位符，只复制源页自身内容。"""
    source = prs.slides[src_no - 1]
    dest = prs.slides.add_slide(source.slide_layout)
    for shp in list(dest.shapes):
        _remove_shape_xml(shp)
    rel_map = _copy_rels(source.part, dest.part)
    for shape in source.shapes:
        try:
            if _is_placeholder_text(_text(shape)):
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


def _normalized_slide_text(slide) -> str:
    text = _slide_text(slide)
    return re.sub(r"\s+", "", text).upper()


def _find_slide_no(prs: Presentation, keywords: list[str], start: int = 1, end: int | None = None) -> int | None:
    """按页面文字查找模板页码，返回 1-based slide no。"""
    end = end or len(prs.slides)
    normalized_keywords = [re.sub(r"\s+", "", k).upper() for k in keywords]
    for no in range(max(1, start), min(end, len(prs.slides)) + 1):
        text = _normalized_slide_text(prs.slides[no - 1])
        if any(k in text for k in normalized_keywords):
            return no
    return None


def _find_qa_slide_no(prs: Presentation) -> int | None:
    # 优先识别含 Q&A / Q＆A / Q & A 的页面；若模板文字被拆分，也兼容 “Q” “A” 同页且接近末尾的情况。
    no = _find_slide_no(prs, ["Q&A", "Q＆A", "Q&A：", "Q&A页", "问答", "答疑"], start=1)
    if no:
        return no
    for i in range(len(prs.slides), 0, -1):
        text = _normalized_slide_text(prs.slides[i - 1])
        if "Q" in text and "A" in text and "&" in text:
            return i
    return None


def _remove_template_empty_slides(prs: Presentation) -> None:
    for idx in range(len(prs.slides) - 1, -1, -1):
        slide = prs.slides[idx]
        text = _slide_text(slide)
        if _is_placeholder_text(text):
            _delete_slide(prs, idx)
            continue
        meaningful = re.sub(r"[\s\-—_：:|]+", "", text)
        if not meaningful and not _tables(slide):
            _delete_slide(prs, idx)


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


def _clear_issue_content(slide) -> None:
    """清空问题页模板中的旧文字和旧表格，只保留背景、边框等视觉元素。"""
    for shp in list(slide.shapes):
        try:
            if getattr(shp, "has_table", False) or getattr(shp, "has_text_frame", False):
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
    title = issue.get("title", "")
    basis = issue.get("basis", "—")
    desc = issue.get("description", "—")
    if not _meaningful_text(title) and not _meaningful_text(basis) and not _meaningful_text(desc):
        return []
    basis_parts = _split_text(basis, 360) if _meaningful_text(basis) else ["—"]
    desc_parts = _split_text(desc, 680) if _meaningful_text(desc) else ["—"]
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
    return _meaningful_text(issue.get("title", "")) or _meaningful_text(issue.get("basis", "")) or _meaningful_text(issue.get("description", ""))


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
    # 新模板已去除公司名称和 Logo，封面不再自动写入公司名称。


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
    auditor = _clean(meta.get("auditor", "—"))
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
    _clear_issue_content(slide)
    sub_page = int(issue.get("_sub_page", 1))
    sub_total = int(issue.get("_sub_total", 1))
    tag = f"（{idx}/{total}）" if total > 1 else ""
    cont = f"  续{sub_page}/{sub_total}" if sub_total > 1 else ""
    page_title = f"问题分类：{category}{tag}{cont}"
    overview = _clean(issue.get("title", "")) or category
    basis = _clean(issue.get("basis", "—"))
    desc = _clean(issue.get("description", "—"))

    _add_textbox(slide, 0.70, 0.52, 12.0, 0.45, page_title, 20, True, BLACK, PP_ALIGN.LEFT)
    rows = [("问题概述", overview)]
    if _meaningful_text(basis) and basis != "—":
        rows.append(("依据", basis))
    rows.append(("问题描述", desc))

    shape = slide.shapes.add_table(len(rows), 2, Inches(0.65), Inches(1.22), Inches(12.05), Inches(5.45))
    tbl = shape.table
    tbl.columns[0].width = Inches(1.05)
    tbl.columns[1].width = Inches(11.00)
    heights = [0.75, 4.70] if len(rows) == 2 else [0.65, 1.55, 3.25]
    for i, h in enumerate(heights[:len(rows)]):
        tbl.rows[i].height = Inches(h)
    for r, (label, value) in enumerate(rows):
        _set_cell(tbl.cell(r, 0), label, 14, True, PP_ALIGN.CENTER)
        size = 11 if label == "问题描述" else 10
        _set_cell(tbl.cell(r, 1), value, size, False, PP_ALIGN.LEFT)


def _copy_if_exists(prs: Presentation, src_no: int | None):
    if src_no and 1 <= src_no <= len(prs.slides):
        return _copy_slide(prs, src_no)
    return None


def render_ppt(context: dict, output_path: str | Path, template_path: str | Path | None = None):
    template = Path(template_path or DEFAULT_TEMPLATE_PATH)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if not template.exists() or template.stat().st_size < 1024 * 100:
        raise FileNotFoundError(f"未找到有效PPT模板：{template}。请在页面先上传你的稽查总结会模板。")

    prs = Presentation(str(template))
    original_count = len(prs.slides)
    if original_count < SLIDE_COUNTS:
        raise RuntimeError(f"模板页数不足：当前 {original_count} 页，至少需要 {SLIDE_COUNTS} 页。请上传完整的稽查总结会模板。")

    qa_slide_no = _find_qa_slide_no(prs)
    ending_slide_no = original_count
    if qa_slide_no == ending_slide_no and original_count > 1:
        ending_slide_no = original_count - 1
    issue_template_no = SLIDE_ISSUE if original_count >= SLIDE_ISSUE else SLIDE_COUNTS

    slide = _copy_slide(prs, SLIDE_COVER); _render_cover(slide, context)
    _copy_slide(prs, SLIDE_THANKS)
    _copy_slide(prs, SLIDE_TOC)
    _copy_slide(prs, SLIDE_PART1)
    slide = _copy_slide(prs, SLIDE_OVERVIEW); _render_overview(slide, context)
    _copy_slide(prs, SLIDE_SCOPE)
    _copy_if_exists(prs, SLIDE_COMMON)      # 共性问题页：仅保留模板，不写内容
    _copy_if_exists(prs, SLIDE_INDIVIDUAL)  # 个性问题页：仅保留模板，不写内容
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

    # 新模板已删除“建议项”页，生成程序不再读取或写入 suggestions。
    _copy_if_exists(prs, qa_slide_no)       # Q&A 页：自动识别并复制，不写内容
    _copy_if_exists(prs, ending_slide_no)   # 结束页：默认复制模板最后一页
    _remove_original_template_slides(prs, original_count)
    _remove_template_empty_slides(prs)
    prs.save(str(output_path))
    return output_path
