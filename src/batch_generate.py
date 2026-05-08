from __future__ import annotations
import argparse
import re
from pathlib import Path

from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.util import Inches

from parser import parse_excel
import renderer

BASE_DIR = Path(__file__).resolve().parent.parent
DEFAULT_OUTPUT = BASE_DIR / "output"
DEFAULT_TEMPLATE = BASE_DIR / "assets" / "template.pptx"
DEFAULT_OUTPUT.mkdir(exist_ok=True, parents=True)

DARK_BLUE = RGBColor(31, 78, 121)
HEADER_BLUE = RGBColor(91, 155, 213)
HEADER_DARK_BLUE = RGBColor(31, 78, 121)
ROW_LIGHT_BLUE = RGBColor(222, 230, 242)
ROW_LIGHT_YELLOW = RGBColor(255, 242, 204)
FIXED_COVER_COMPANY = "北京xxxx医药科技有限公司"


def safe_stem(name: str) -> str:
    invalid = '<>:"/\\|?*'
    return "".join("_" if c in invalid else c for c in name).strip().strip(".") or "output"


def _set_cell_font_color(cell, color: RGBColor):
    try:
        for p in cell.text_frame.paragraphs:
            for r in p.runs:
                r.font.color.rgb = color
    except Exception:
        pass


def _patched_render_cover(slide, context):
    renderer._remove_text_shapes(slide)
    meta = context.get("meta", {})
    project = renderer._clean(meta.get("project_name", "—"))
    center = renderer._clean(meta.get("center_name", "—"))
    center_no = renderer._clean(meta.get("center_no", ""))
    audit_date = renderer._clean(meta.get("audit_date", "—"))

    title = f"{project}-{center}" + (f"（中心编号{center_no}）" if center_no else "")

    # 蓝色横框约位于 3.52-5.12 英寸；标题与黄字均固定在框内。
    renderer._add_textbox(slide, 0.25, 3.58, 12.70, 0.82, title, 22, True, renderer.WHITE, PP_ALIGN.LEFT)
    renderer._add_textbox(slide, 0.25, 4.62, 12.70, 0.40, "中心稽查末次会议", 22, True, renderer.YELLOW, PP_ALIGN.LEFT)

    # 时间与公司名放在蓝框下方，深蓝色 14 号字；公司名保留模板占位，不替换全称。
    renderer._add_textbox(slide, 0.35, 5.33, 12.30, 0.28, f"时间：{audit_date}", 14, True, DARK_BLUE, PP_ALIGN.LEFT)
    renderer._add_textbox(slide, 0.35, 5.68, 12.30, 0.28, FIXED_COVER_COMPANY, 14, True, DARK_BLUE, PP_ALIGN.LEFT)


renderer._render_cover = _patched_render_cover


def _contains(text: str, *words: str) -> bool:
    low = (text or "").lower()
    return any(w.lower() in low for w in words)


def _risk_advice_by_issue(category: str, text: str) -> str:
    text = renderer._clean(text)
    if _contains(text, "sae", "严重不良", "死亡", "住院", "转归"):
        return "逐例核对病历、AE/SAE表、EDC与安全上报系统，确认严重性、相关性、转归、上报时限、随访记录及研究者医学判断证据。"
    if _contains(text, "ae", "不良事件", "安全性评估", "合并用药"):
        return "以受试者时间轴复核AE识别、记录、分级、相关性、处理措施、转归和合并用药，确保原始记录、EDC与医学判断一致。"
    if _contains(text, "筛选", "入排", "入组", "排除标准", "入选标准"):
        return "建立入排标准逐项核查表，准备筛选期检查、医学判断、研究者确认、偏离判定及不影响受试者安全/数据可靠性的说明。"
    if _contains(text, "疗效", "影像", "recist", "肿瘤评估", "靶病灶", "非靶病灶"):
        return "复核影像检查日期、评估时间窗、靶/非靶病灶记录、疗效判定依据及EDC录入，提前准备评估差异说明和原始影像索引。"
    if _contains(text, "edc", "crf", "query", "迟录", "漏录", "不一致"):
        return "导出EDC关键字段和Query清单，逐项核对源数据、录入时限、逻辑一致性及Query关闭证据，形成差异说明和更正记录。"
    if _contains(text, "原始", "源文件", "病历", "溯源", "his", "lis", "pacs"):
        return "按受试者建立源数据追溯包，逐项核对病历、HIS/LIS/PACS、源文件与EDC，提前标注差异原因和研究者确认说明。"
    if _contains(text, "知情", "icf", "签署", "授权", "受试者权益"):
        return "逐份复核ICF版本、签署日期/时间、签署人、授权分工和告知过程记录，准备签署过程说明及必要的补充/更正证据。"
    if _contains(text, "药品", "试验用药", "发放", "回收", "温度", "超温", "清点"):
        return "复核药品接收、储存温度、发放、回收、清点、销毁和授权人员记录，确保账物卡一致并能解释异常处理。"
    if _contains(text, "样本", "采血", "离心", "运输", "中心实验室", "温控"):
        return "核对样本采集、处理、保存、运输、交接和检测结果回传链条，重点准备时间窗、标签、温控及偏差处理证据。"
    if _contains(text, "伦理", "批件", "递交", "持续审查", "版本"):
        return "核对伦理批件、递交材料、方案/ICF版本、生效日期、持续审查和安全性信息递交记录，确保版本执行无倒挂。"
    return "针对该问题准备原始证据、研究者说明、整改记录和CAPA闭环材料，并对同类受试者/同类流程开展横向复核。"


def _risk_title(issue: dict) -> str:
    desc = renderer._clean(issue.get("description", "")) or renderer._clean(issue.get("summary", ""))
    desc = re.sub(r"依据/风险逻辑[:：]?.*", "", desc, flags=re.S)
    desc = desc.replace("\n\n", "\n")
    return desc[:180].rstrip("；，,。 ")


def _patched_extract_top5_risks(context: dict) -> list[dict]:
    issues = [x for x in context.get("issues", []) if renderer._has_issue_content(x)]
    enriched = []
    seen = set()
    for issue in issues:
        category = issue.get("category", "其他")
        desc = _risk_title(issue)
        key = (category, desc[:70])
        if key in seen or not desc:
            continue
        seen.add(key)
        full_text = f"{issue.get('summary', '')}\n{issue.get('description', '')}\n{issue.get('basis', '')}"
        enriched.append({
            "category": category,
            "risk": desc,
            "advice": _risk_advice_by_issue(category, full_text),
            "score": renderer._risk_score(issue),
        })
    enriched.sort(key=lambda x: x["score"], reverse=True)
    return enriched[:5]


renderer._extract_top5_risks = _patched_extract_top5_risks


def _patched_render_risk_summary(slide, context):
    renderer._clear_issue_content(slide)
    risks = _patched_extract_top5_risks(context)
    renderer._add_textbox(slide, 0.55, 0.33, 12.20, 0.44, "核查准备重点关注问题及迎检建议", 24, True, renderer.BLACK, PP_ALIGN.LEFT)

    if not risks:
        renderer._add_textbox(slide, 0.75, 1.35, 11.80, 0.80, "本次上传文件中未识别到可用于提炼TOP5的问题内容，请复核Excel问题分类、问题描述和依据列是否完整。", 16, False, renderer.BLACK, PP_ALIGN.LEFT)
        return

    shape = slide.shapes.add_table(6, 4, Inches(0.45), Inches(0.95), Inches(12.45), Inches(6.05))
    tbl = shape.table
    widths = [0.55, 2.10, 5.05, 4.75]
    for i, w in enumerate(widths):
        tbl.columns[i].width = Inches(w)
    tbl.rows[0].height = Inches(0.42)
    for r in range(1, 6):
        tbl.rows[r].height = Inches(1.10)

    headers = ["序号", "风险类别", "TOP高风险问题", "迎检建议"]
    for c, h in enumerate(headers):
        renderer._set_cell(tbl.cell(0, c), h, 12, True, PP_ALIGN.CENTER)
        renderer._set_cell_fill(tbl.cell(0, c), HEADER_BLUE)
        _set_cell_font_color(tbl.cell(0, c), renderer.WHITE)

    for r in range(1, 6):
        if r <= len(risks):
            item = risks[r - 1]
            values = [str(r), item["category"], item["risk"], item["advice"]]
        else:
            values = [str(r), "—", "—", "—"]
        for c, v in enumerate(values):
            size = 8 if c >= 2 else 9
            align = PP_ALIGN.CENTER if c == 0 else PP_ALIGN.LEFT
            renderer._set_cell(tbl.cell(r, c), v, size, c == 0, align)
            if c == 0:
                renderer._set_cell_fill(tbl.cell(r, c), ROW_LIGHT_YELLOW)
            else:
                renderer._set_cell_fill(tbl.cell(r, c), ROW_LIGHT_BLUE)

    renderer._add_textbox(slide, 0.55, 7.03, 12.25, 0.18, "注：本页基于本次稽查发现自动提炼，用于核查准备优先级排序；正式迎检材料需结合项目医学判断及原始证据人工复核。", 8, False, renderer.GRAY, PP_ALIGN.LEFT)


renderer._render_risk_summary = _patched_render_risk_summary


def render_one(excel_path: str | Path, output_dir: str | Path | None = None, template_path: str | Path | None = None) -> Path:
    excel_path = Path(excel_path)
    output_dir = Path(output_dir or DEFAULT_OUTPUT)
    output_dir.mkdir(parents=True, exist_ok=True)
    context = parse_excel(excel_path)
    out = output_dir / f"{safe_stem(excel_path.stem)}-V8.5封面及核查准备页修正版.pptx"
    renderer.render_ppt(context, out, template_path=Path(template_path or DEFAULT_TEMPLATE))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", type=str)
    ap.add_argument("--excel_dir", type=str)
    ap.add_argument("--output_dir", type=str, default=str(DEFAULT_OUTPUT))
    ap.add_argument("--template", type=str, default=str(DEFAULT_TEMPLATE))
    args = ap.parse_args()
    files = []
    if args.excel:
        files.append(Path(args.excel))
    if args.excel_dir:
        d = Path(args.excel_dir)
        for ext in ("*.xlsx", "*.xlsm", "*.xls"):
            files += sorted(d.glob(ext))
    if not files:
        raise SystemExit("请提供 --excel 或 --excel_dir")
    for f in files:
        out = render_one(f, args.output_dir, template_path=args.template)
        print(out)


if __name__ == "__main__":
    main()
