from __future__ import annotations
import argparse
import re
from pathlib import Path

from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.util import Inches

from parser import parse_excel
import renderer
from ai_summary import generate_ai_top5

BASE_DIR = Path(__file__).resolve().parent.parent
DEFAULT_OUTPUT = BASE_DIR / "output"
DEFAULT_TEMPLATE = BASE_DIR / "assets" / "template.pptx"
DEFAULT_OUTPUT.mkdir(exist_ok=True, parents=True)

HEADER_BLUE = RGBColor(91, 155, 213)
ROW_LIGHT_BLUE = RGBColor(222, 230, 242)
ROW_LIGHT_YELLOW = RGBColor(255, 242, 204)
FIXED_COVER_COMPANY = "北京xxxx医药科技有限公司"

TOPIC_RULES = [
    {"id":"source_edc","risk":"EDC与源数据不一致/录入证据不足","keywords":["edc","crf","源数据","原始","病历","溯源","不一致","漏录","迟录","query"],"weight":95,"analysis":"数据可靠性（高）：直接影响临床试验数据真实性、完整性、准确性和可追溯性。\n合规风险（高）：不符合ALCOA+原则，现场核查中易被判定为关键数据证据链不足。\n注册核查风险（高）：可能导致关键疗效/安全性数据被质疑。","advice":"立即行动：按受试者逐字段比对EDC与HIS/LIS/PACS、病历、实验室/影像报告，形成差异清单。\n根因排查：复核EDC录入SOP执行、录入时限、Query关闭和质控复核机制。\n证据准备：整理更正记录、研究者确认说明、Query关闭截图和SDV复核记录。"},
    {"id":"safety_reporting","risk":"AE/SAE识别、记录、评估或上报链条不完整","keywords":["sae","susar","严重不良","ae","不良事件","死亡","住院","转归","安全性","上报"],"weight":92,"analysis":"受试者安全（高）：可能影响AE/SAE及时识别、医学判断、处理和随访。\n合规风险（高）：可能涉及安全性信息报告时限、严重性/相关性判断和研究者职责落实。\n核查关注风险（高）：监管现场通常重点追溯病历、AE表、SAE报告和EDC一致性。","advice":"立即行动：逐例核查病历、AE/SAE表、EDC及安全上报系统的完整链条。\n医学复核：确认严重性、相关性、转归、处理措施、随访记录和上报时限。\n现场准备：准备研究者医学判断说明、补充随访证据和安全性事件流程演练记录。"},
    {"id":"eligibility_protocol","risk":"入排标准/方案依从性证据不足","keywords":["入排","筛选","入组","排除标准","入选标准","方案偏离","访视窗口","方案依从","给药","检查"],"weight":88,"analysis":"方案依从性（高）：可能影响受试者是否合规入组及关键流程是否按方案执行。\n数据解释风险（高）：入排或访视偏差可能影响疗效、安全性和统计分析人群判断。\n核查风险（高）：现场会追溯筛选记录、医学判断、检查报告和偏离闭环。","advice":"立即行动：建立入排标准和访视窗口逐项核查表，逐例确认支持性证据。\n证据准备：整理筛选检查、研究者医学判断、偏离记录、CAPA和影响评估。\n系统改进：前置设置入排复核、访视窗口提醒和关键方案流程质控节点。"},
    {"id":"efficacy_imaging","risk":"疗效/影像评价依据不充分或不一致","keywords":["疗效","影像","recist","肿瘤评估","靶病灶","非靶病灶","主要终点","评估"],"weight":84,"analysis":"疗效评价风险（高）：可能影响主要/关键终点判定的准确性和可重复性。\n数据可靠性风险（高）：影像报告、评估表和EDC不一致会影响疗效数据可信度。\n注册核查风险（高）：监管核查可能重点追溯原始影像、评估标准和研究者判断。","advice":"立即行动：复核影像检查日期、评估窗口、靶/非靶病灶记录、疗效判定和EDC录入。\n证据准备：建立影像索引、评估表、研究者判断说明和差异解释。\n专项复核：对所有疗效评价相关受试者开展横向一致性核查。"},
    {"id":"icf","risk":"知情同意过程或版本控制证据不足","keywords":["知情","icf","签署","受试者权益","授权","告知","版本"],"weight":82,"analysis":"受试者权益风险（高）：可能影响知情同意有效性和受试者权益保护。\n伦理合规风险（高）：版本、签署时间、签署人员或告知过程证据不足，易形成现场核查重点。\n证据链风险（高）：无法证明受试者在试验相关操作前完成充分知情。","advice":"立即行动：逐份复核ICF版本、签署日期/时间、签署人、授权分工和告知过程记录。\n证据准备：整理授权表、培训记录、签署过程说明和必要的更正/补充说明。\n横向复核：检查所有受试者是否存在同类版本或签署流程问题。"},
    {"id":"drug","risk":"试验用药品管理链条不完整","keywords":["药品","试验用药","发放","回收","温度","超温","清点","销毁","批号"],"weight":78,"analysis":"用药安全风险（中高）：可能影响受试者实际用药、依从性和安全性判断。\n可追溯性风险（中高）：药品接收、储存、发放、回收和销毁链条不完整会影响账物一致。\n核查风险（中高）：现场可能重点核对药品台账、温度记录和授权人员操作记录。","advice":"立即行动：复核药品接收、储存温度、发放、回收、清点、销毁和异常处理记录。\n台账修正：补齐批号、数量、日期、受试者编号、授权人员和温控证据。\n预防措施：建立药品管理月度自查和异常温度即时升级机制。"},
    {"id":"sample","risk":"生物样本采集/处理/运输证据链不足","keywords":["样本","采血","离心","运输","中心实验室","温控","交接","检测"],"weight":72,"analysis":"样本有效性风险（中高）：采集、处理、保存或运输偏差可能影响检测结果可信度。\n证据链风险（中高）：时间窗、标签、温控和交接记录缺失会影响样本可追溯性。\n核查风险（中高）：监管可能追溯样本从采集到检测结果回传的全过程。","advice":"立即行动：核对样本采集、处理、保存、运输、交接和检测结果回传全链条。\n证据准备：补齐时间窗、标签、温控、交接单和偏差处理记录。\n系统改进：建立样本链条关键节点清单和运输/温控异常升级机制。"},
    {"id":"ethics","risk":"伦理递交/批件/版本执行证据不足","keywords":["伦理","批件","递交","持续审查","修正案","备案"],"weight":70,"analysis":"伦理合规风险（中高）：伦理批件、递交材料或持续审查证据不足会影响试验合规基础。\n版本控制风险（中高）：方案/ICF版本执行时间与伦理批准时间不一致，可能形成重要发现。\n现场核查风险（中高）：需证明所有关键文件均在批准后执行。","advice":"立即行动：核对伦理批件、递交材料、方案/ICF版本、生效日期和持续审查记录。\n证据准备：形成伦理递交与版本执行时间轴，标注批准前后关键操作。\n预防措施：建立版本发布、培训、生效和执行的联动确认机制。"},
]


def _set_cell_font_color(cell, color: RGBColor):
    try:
        for p in cell.text_frame.paragraphs:
            for r in p.runs:
                r.font.color.rgb = color
    except Exception:
        pass


def safe_stem(name: str) -> str:
    invalid = '<>:"/\\|?*'
    return "".join("_" if c in invalid else c for c in name).strip().strip(".") or "output"


def safe_text(value: str, limit: int = 80) -> str:
    text = renderer._clean(value).replace("\n", "；")
    text = re.sub(r"；{2,}", "；", text)
    return text[:limit].rstrip("；，,。 ")


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


def topic_match_score(rule: dict, text: str) -> int:
    low = text.lower()
    hits = sum(1 for kw in rule["keywords"] if kw.lower() in low)
    if hits == 0:
        return 0
    return rule["weight"] + hits * 8


def issue_display_example(issue: dict) -> str:
    desc = renderer._clean(issue.get("description", "")) or renderer._clean(issue.get("summary", ""))
    desc = re.sub(r"依据[:：].*", "", desc, flags=re.S)
    desc = re.sub(r"问题[:：]", "", desc)
    return safe_text(desc, 48)


def _rule_top5(context: dict) -> list[dict]:
    issues = [x for x in context.get("issues", []) if renderer._has_issue_content(x)]
    topics = {}
    for issue in issues:
        text = f"{issue.get('category','')}\n{issue.get('summary','')}\n{issue.get('description','')}\n{issue.get('basis','')}"
        best_rule = None
        best_score = 0
        for rule in TOPIC_RULES:
            score = topic_match_score(rule, text)
            if score > best_score:
                best_rule = rule
                best_score = score
        if not best_rule:
            continue
        tid = best_rule["id"]
        if tid not in topics:
            topics[tid] = {"rule": best_rule, "score": 0, "examples": [], "count": 0}
        topics[tid]["score"] += best_score + renderer._risk_score(issue) // 4
        topics[tid]["count"] += 1
        ex = issue_display_example(issue)
        if ex and ex not in topics[tid]["examples"] and len(topics[tid]["examples"]) < 2:
            topics[tid]["examples"].append(ex)
    ranked = sorted(topics.values(), key=lambda x: x["score"], reverse=True)[:5]
    result = []
    for item in ranked:
        rule = item["rule"]
        examples = "；".join(item["examples"])
        risk = rule["risk"]
        if examples:
            risk += f"（涉及{item['count']}项发现，例：{examples}）"
        result.append({"risk": risk, "analysis": rule["analysis"], "advice": rule["advice"], "score": item["score"], "source": "规则聚类兜底"})
    return result


def _patched_extract_top5_risks(context: dict) -> list[dict]:
    ai_rows = generate_ai_top5(context)
    if ai_rows:
        for row in ai_rows:
            row["source"] = row.get("source") or "AI智能总结"
        return ai_rows[:5]
    return _rule_top5(context)


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
    widths = [0.60, 2.75, 4.15, 4.95]
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
            size = 7 if c in (2, 3) else 8
            align = PP_ALIGN.CENTER if c == 0 else PP_ALIGN.LEFT
            renderer._set_cell(tbl.cell(r, c), v, size, c == 0, align)
            renderer._set_cell_fill(tbl.cell(r, c), ROW_LIGHT_YELLOW if c == 0 else ROW_LIGHT_BLUE)
    source = risks[0].get("source", "未知") if risks else "未知"
    renderer._add_textbox(slide, 0.55, 7.05, 12.25, 0.16, f"注：生成来源：{source}；正式材料需结合项目医学判断及原始证据人工复核。", 8, False, renderer.GRAY, PP_ALIGN.LEFT)


renderer._render_risk_summary = _patched_render_risk_summary


def render_one(excel_path: str | Path, output_dir: str | Path | None = None, template_path: str | Path | None = None) -> Path:
    excel_path = Path(excel_path)
    output_dir = Path(output_dir or DEFAULT_OUTPUT)
    output_dir.mkdir(parents=True, exist_ok=True)
    context = parse_excel(excel_path)
    out = output_dir / f"{safe_stem(excel_path.stem)}-V8.9来源标识版.pptx"
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
