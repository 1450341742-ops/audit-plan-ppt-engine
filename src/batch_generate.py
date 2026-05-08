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

HEADER_BLUE = RGBColor(91, 155, 213)
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
    renderer._add_textbox(slide, 0.25, 3.58, 12.70, 0.82, title, 22, True, renderer.WHITE, PP_ALIGN.LEFT)
    renderer._add_textbox(slide, 0.25, 4.62, 12.70, 0.40, "中心稽查末次会议", 22, True, renderer.YELLOW, PP_ALIGN.LEFT)
    renderer._add_textbox(slide, 0.35, 5.33, 12.30, 0.28, f"时间：{audit_date}", 14, True, renderer.WHITE, PP_ALIGN.LEFT)
    renderer._add_textbox(slide, 0.35, 5.68, 12.30, 0.28, FIXED_COVER_COMPANY, 14, True, renderer.WHITE, PP_ALIGN.LEFT)


renderer._render_cover = _patched_render_cover


def _contains(text: str, *words: str) -> bool:
    low = (text or "").lower()
    return any(w.lower() in low for w in words)


def _brief_issue(issue: dict) -> str:
    text = renderer._clean(issue.get("description", "")) or renderer._clean(issue.get("summary", ""))
    text = re.sub(r"依据[:：].*", "", text, flags=re.S)
    text = re.sub(r"问题[:：]", "", text)
    text = text.replace("\n", "；")
    text = re.sub(r"；{2,}", "；", text)
    return text[:70].rstrip("；，,。 ") or "高风险问题待复核"


def _risk_dimension_analysis(category: str, text: str) -> str:
    text = renderer._clean(text)
    lines = []
    if _contains(text, "edc", "crf", "源数据", "原始", "病历", "溯源", "不一致", "漏录", "迟录"):
        lines.append("数据可靠性（高）：可能影响源数据真实性、完整性、准确性和可追溯性。")
    if _contains(text, "sae", "susar", "严重不良", "ae", "不良事件", "死亡", "住院", "转归", "安全性"):
        lines.append("受试者安全（高）：可能影响AE/SAE识别、医学判断、及时上报和持续随访。")
    if _contains(text, "入排", "筛选", "入组", "排除标准", "方案偏离", "访视窗口", "给药", "检查"):
        lines.append("方案依从性（高）：可能影响受试者合规入组、关键流程执行和主要终点评价。")
    if _contains(text, "知情", "icf", "签署", "授权", "受试者权益"):
        lines.append("受试者权益（高）：可能影响知情同意有效性和伦理合规判断。")
    if _contains(text, "伦理", "批件", "递交", "持续审查", "版本"):
        lines.append("伦理合规（中高）：可能出现版本执行、递交审批或持续审查证据不完整。")
    if _contains(text, "药品", "试验用药", "发放", "回收", "温度", "超温", "清点", "销毁"):
        lines.append("试验用药品管理（中高）：可能影响药品可追溯性、用药安全和盲态/依从性判断。")
    if _contains(text, "样本", "采血", "离心", "运输", "中心实验室", "温控"):
        lines.append("样本链条（中高）：可能影响样本有效性、检测结果可信度和证据链完整性。")
    if _contains(text, "疗效", "影像", "recist", "肿瘤评估", "靶病灶"):
        lines.append("疗效评价（高）：可能影响疗效终点判定、影像证据一致性和统计分析可信度。")
    if not lines:
        lines.append(f"合规与质量风险（中高）：该问题归属于{category}，可能影响现场核查对流程执行和整改闭环的判断。")
    return "\n".join(lines[:3])


def _inspection_advice(category: str, text: str) -> str:
    text = renderer._clean(text)
    actions = []
    if _contains(text, "edc", "crf", "源数据", "原始", "病历", "溯源", "不一致", "漏录", "迟录"):
        actions.append("立即行动：逐例比对EDC与HIS/LIS/PACS、病历、实验室/影像报告，形成差异清单和研究者确认说明。")
        actions.append("系统改进：建立关键字段SDV清单、录入时限检查和Query关闭复核机制。")
    if _contains(text, "sae", "susar", "严重不良", "ae", "不良事件", "死亡", "住院", "转归", "安全性"):
        actions.append("立即行动：核查AE/SAE完整链条，确认严重性、相关性、转归、上报时限和随访闭环。")
        actions.append("应急演练：组织研究团队进行AE识别、记录、报告流程模拟演练。")
    if _contains(text, "入排", "筛选", "入组", "排除标准", "方案偏离", "访视窗口"):
        actions.append("立即行动：建立入排/访视窗口逐项核查表，准备医学判断、偏离判定和CAPA证据。")
        actions.append("根因排查：复核筛选评估、研究者确认和项目组审核流程是否落实。")
    if _contains(text, "知情", "icf", "签署", "授权", "受试者权益"):
        actions.append("立即行动：逐份复核ICF版本、签署日期/时间、签署人资质、授权分工和告知记录。")
        actions.append("证据准备：整理知情过程说明、授权表、培训记录及必要的更正/补充说明。")
    if _contains(text, "药品", "试验用药", "发放", "回收", "温度", "超温", "清点"):
        actions.append("立即行动：复核药品接收、储存温度、发放、回收、清点、销毁及异常处理记录。")
        actions.append("台账修正：确保账物卡一致，补齐批号、数量、日期和授权人员证据。")
    if _contains(text, "样本", "采血", "离心", "运输", "中心实验室", "温控"):
        actions.append("立即行动：核对样本采集、处理、保存、运输、交接和检测结果回传全链条。")
        actions.append("证据准备：补齐时间窗、标签、温控、交接单和偏差处理记录。")
    if _contains(text, "疗效", "影像", "recist", "肿瘤评估", "靶病灶"):
        actions.append("立即行动：复核影像检查日期、评估时间窗、靶/非靶病灶记录和疗效判定依据。")
        actions.append("证据准备：建立影像索引、评估表、研究者判断说明和EDC一致性复核记录。")
    if not actions:
        actions.append("立即行动：围绕该问题准备原始证据、研究者说明、整改记录和CAPA闭环材料。")
        actions.append("横向复核：对同类受试者、同类流程和同类记录开展全面排查。")
    return "\n".join(actions[:3])


def _patched_extract_top5_risks(context: dict) -> list[dict]:
    issues = [x for x in context.get("issues", []) if renderer._has_issue_content(x)]
    enriched = []
    seen = set()
    for issue in issues:
        category = issue.get("category", "其他")
        full_text = f"{issue.get('summary', '')}\n{issue.get('description', '')}\n{issue.get('basis', '')}"
        risk = _brief_issue(issue)
        key = (category, risk[:60])
        if key in seen:
            continue
        seen.add(key)
        enriched.append({
            "risk": risk,
            "analysis": _risk_dimension_analysis(category, full_text),
            "advice": _inspection_advice(category, full_text),
            "score": renderer._risk_score(issue),
        })
    enriched.sort(key=lambda x: x["score"], reverse=True)
    return enriched[:5]


renderer._extract_top5_risks = _patched_extract_top5_risks


def _patched_render_risk_summary(slide, context):
    renderer._clear_issue_content(slide)
    risks = _patched_extract_top5_risks(context)
    renderer._add_textbox(slide, 0.45, 0.28, 12.40, 0.42, "TOP5高风险问题及核查应对建议", 22, True, renderer.BLACK, PP_ALIGN.LEFT)

    if not risks:
        renderer._add_textbox(slide, 0.75, 1.35, 11.80, 0.80, "本次上传文件中未识别到可用于提炼TOP5的问题内容，请复核Excel问题分类、问题描述和依据列是否完整。", 16, False, renderer.BLACK, PP_ALIGN.LEFT)
        return

    shape = slide.shapes.add_table(6, 4, Inches(0.45), Inches(0.90), Inches(12.45), Inches(6.10))
    tbl = shape.table
    widths = [0.60, 2.55, 4.15, 5.15]
    for i, w in enumerate(widths):
        tbl.columns[i].width = Inches(w)
    tbl.rows[0].height = Inches(0.40)
    for r in range(1, 6):
        tbl.rows[r].height = Inches(1.14)

    headers = ["排名", "高风险问题", "风险维度分析", "核查应对建议"]
    for c, h in enumerate(headers):
        renderer._set_cell(tbl.cell(0, c), h, 11, True, PP_ALIGN.CENTER)
        renderer._set_cell_fill(tbl.cell(0, c), HEADER_BLUE)
        _set_cell_font_color(tbl.cell(0, c), renderer.WHITE)

    for r in range(1, 6):
        if r <= len(risks):
            item = risks[r - 1]
            values = [str(r), item["risk"], item["analysis"], item["advice"]]
        else:
            values = [str(r), "—", "—", "—"]
        for c, v in enumerate(values):
            size = 8 if c >= 2 else 9
            align = PP_ALIGN.CENTER if c == 0 else PP_ALIGN.LEFT
            renderer._set_cell(tbl.cell(r, c), v, size, c == 0, align)
            renderer._set_cell_fill(tbl.cell(r, c), ROW_LIGHT_YELLOW if c == 0 else ROW_LIGHT_BLUE)

    renderer._add_textbox(slide, 0.55, 7.05, 12.25, 0.16, "注：本页基于本次稽查发现自动提炼，用于核查准备优先级排序；正式迎检材料需结合项目医学判断及原始证据人工复核。", 8, False, renderer.GRAY, PP_ALIGN.LEFT)


renderer._render_risk_summary = _patched_render_risk_summary


def render_one(excel_path: str | Path, output_dir: str | Path | None = None, template_path: str | Path | None = None) -> Path:
    excel_path = Path(excel_path)
    output_dir = Path(output_dir or DEFAULT_OUTPUT)
    output_dir.mkdir(parents=True, exist_ok=True)
    context = parse_excel(excel_path)
    out = output_dir / f"{safe_stem(excel_path.stem)}-V8.6TOP5核查建议版.pptx"
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
