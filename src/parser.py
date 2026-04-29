from __future__ import annotations
import re
from pathlib import Path
from typing import Any
import openpyxl

STANDARD_CATEGORIES = [
    "国家药物临床试验政策法规的遵循",
    "伦理委员会审核要求的遵循",
    "知情同意书（ICF）的签署和记录",
    "原始文件的建立、内容和记录",
    "门诊/住院HIS、LIS、PACS等系统数据溯源",
    "方案依从性",
    "药物疗效/研究评价指标的评估",
    "安全性信息评估，记录与报告",
    "CRF填写（时效性、一致性、溯源性、完整性）",
    "试验用药品管理",
    "生物样本管理",
    "临床研究必须文件",
    "申办方/CRO职责",
    "其他",
]

ALIASES = {
    "授权与分工":"国家药物临床试验政策法规的遵循",
    "法规":"国家药物临床试验政策法规的遵循",
    "伦理":"伦理委员会审核要求的遵循",
    "知情同意":"知情同意书（ICF）的签署和记录",
    "知情同意书":"知情同意书（ICF）的签署和记录",
    "icf":"知情同意书（ICF）的签署和记录",
    "源文件":"原始文件的建立、内容和记录",
    "原始文件":"原始文件的建立、内容和记录",
    "原始记录":"原始文件的建立、内容和记录",
    "his":"门诊/住院HIS、LIS、PACS等系统数据溯源",
    "lis":"门诊/住院HIS、LIS、PACS等系统数据溯源",
    "pacs":"门诊/住院HIS、LIS、PACS等系统数据溯源",
    "方案及其他文件依从性":"方案依从性",
    "方案依从":"方案依从性",
    "方案偏离":"方案依从性",
    "疗效":"药物疗效/研究评价指标的评估",
    "终点":"药物疗效/研究评价指标的评估",
    "安全性":"安全性信息评估，记录与报告",
    "ae":"安全性信息评估，记录与报告",
    "sae":"安全性信息评估，记录与报告",
    "crf":"CRF填写（时效性、一致性、溯源性、完整性）",
    "edc":"CRF填写（时效性、一致性、溯源性、完整性）",
    "药品":"试验用药品管理",
    "样本":"生物样本管理",
    "必须文件":"临床研究必须文件",
    "研究者文件夹":"临床研究必须文件",
    "cro":"申办方/CRO职责",
    "申办方":"申办方/CRO职责",
    "其他":"其他",
}

STOP_WORDS = ["审核", "回复", "签字", "日期", "备注", "说明", "填表", "批准", "确认"]


def clean_text(v: Any) -> str:
    if v is None:
        return ""
    s = str(v).replace("\r", "\n")
    s = re.sub(r"[ \t\xa0]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def norm(s: str) -> str:
    s = clean_text(s).lower()
    s = s.replace("（", "(").replace("）", ")").replace("，", "、").replace(",", "、")
    return re.sub(r"\s+", "", s)


def normalize_category(raw: str) -> str:
    raw = clean_text(raw)
    if not raw:
        return "其他"
    nr = norm(raw)
    for c in STANDARD_CATEGORIES:
        if norm(c) == nr or nr in norm(c) or norm(c) in nr:
            return c
    for k, v in ALIASES.items():
        if norm(k) in nr:
            return v
    return "其他"


def rows_from_ws(ws) -> list[list[str]]:
    return [[clean_text(c) for c in row] for row in ws.iter_rows(values_only=True)]


def row_text(row: list[str]) -> str:
    return " ".join(x for x in row if x)


def find_label_value(rows: list[list[str]], labels: list[str]) -> str:
    labs = [norm(x) for x in labels]
    for row in rows:
        for i, cell in enumerate(row):
            if not cell:
                continue
            nc = norm(cell)
            if any(lab == nc or lab in nc for lab in labs):
                m = re.match(r"^.*?[：:]\s*(.+)$", cell)
                if m and clean_text(m.group(1)):
                    return clean_text(m.group(1))
                for j in range(i + 1, len(row)):
                    if row[j]:
                        return row[j]
    return ""


def extract_protocol(text: str) -> str:
    text = clean_text(text)
    m = re.search(r"(?<![A-Z0-9])([A-Z]{2,}(?:-[A-Z0-9]+)+)(?![A-Z0-9-])", text, flags=re.I)
    if m:
        return m.group(1).upper().rstrip("-_/)")
    m = re.search(r"\b([A-Z]{2,}[A-Z0-9\-_/]*\d[A-Z0-9\-_/]*)\b", text, flags=re.I)
    return m.group(1).upper().rstrip("-_/)") if m else ""


def extract_subject_ids(text: str) -> list[str]:
    text = clean_text(text)
    hits: list[str] = []
    patterns = [
        r"筛选号[:：]?\s*([ST]?\d{3,6})",
        r"受试者(?:编号|文件夹)?[:：]?\s*([ST]?\d{3,6})",
        r"\b([ST]\d{3,6})\b",
        r"\b(\d{5})\b",
    ]
    for pat in patterns:
        for m in re.findall(pat, text, flags=re.I):
            raw = clean_text(m).upper()
            if raw and raw not in hits and not re.fullmatch(r"20\d{2}", raw):
                hits.append(raw)
    return hits


def split_basis_desc(text: str) -> tuple[str, str]:
    text = clean_text(text)
    if not text:
        return "—", "—"
    m = re.search(r"依据[:：]\s*(.+?)(?=\n(?:问题|描述|筛选号|受试者|备注)[:：]?|$)", text, flags=re.S)
    if m:
        basis = clean_text(m.group(1))
        desc = clean_text(text.replace(m.group(0), "", 1))
        return basis or "—", desc or "—"
    markers = ["方案", "研究者手册", "药物临床试验质量管理规范", "GCP", "ICH", "SOP"]
    for mk in markers:
        pos = text.find(mk)
        if pos != -1 and pos < max(220, len(text) // 2):
            return clean_text(text[pos:]) or "—", clean_text(text[:pos]) or text
    return "—", text


def extract_title(desc: str, category: str) -> str:
    desc = clean_text(desc)
    for pat in [r"标题[:：]\s*(.+?)(?:\n|$)", r"问题[:：]\s*(.+?)(?:\n|$)"]:
        m = re.search(pat, desc, flags=re.S)
        if m:
            return clean_text(m.group(1))[:80]
    lines = [x.strip("—- ").strip() for x in desc.splitlines() if x.strip()]
    for line in lines[:5]:
        if len(line) <= 80 and not any(w in line for w in STOP_WORDS):
            return line
    return category


def is_header_row(row: list[str]) -> bool:
    txt = norm(row_text(row))
    has_cat = any(k in txt for k in ["问题分类", "分类", "问题类别", "检查内容"])
    has_desc = any(k in txt for k in ["问题描述", "描述", "总结描述", "问题概述", "稽查发现"])
    has_basis = any(k in txt for k in ["依据", "法规依据"])
    return (has_cat and has_desc) or (has_desc and has_basis)


def find_col(header: list[str], cands: list[str]) -> int:
    for idx, h in enumerate(header):
        nh = norm(h)
        for c in cands:
            nc = norm(c)
            if nc and (nc in nh or nh in nc):
                return idx
    return -1


def parse_table_after_header(rows: list[list[str]], header_idx: int) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    header = rows[header_idx]
    ci = find_col(header, ["问题分类", "问题类别", "分类", "检查内容"])
    di = find_col(header, ["问题描述", "描述", "总结描述", "问题概述", "稽查发现"])
    ti = find_col(header, ["问题概述", "标题", "总结描述"])
    bi = find_col(header, ["依据", "法规依据", "参考依据"])
    si = find_col(header, ["级别", "问题级别", "严重程度", "风险等级"])
    subi = find_col(header, ["受试者", "筛选号", "受试者编号"])

    current_cat = ""
    for row in rows[header_idx + 1:]:
        if not any(row):
            continue
        merged = row_text(row)
        nmerged = norm(merged)
        if any(x in nmerged for x in ["建议项", "capa回复", "纠正预防", "审核人", "批准人"]):
            if "建议项" in nmerged:
                break
        raw_cat = row[ci] if 0 <= ci < len(row) else ""
        if raw_cat:
            current_cat = raw_cat
        raw_desc = row[di] if 0 <= di < len(row) else merged
        raw_title = row[ti] if 0 <= ti < len(row) else ""
        if not clean_text(raw_desc) and not clean_text(raw_title):
            continue
        if any(w in clean_text(raw_desc) for w in STOP_WORDS) and len(clean_text(raw_desc)) < 20:
            continue
        category = normalize_category(raw_cat or current_cat or raw_title or raw_desc)
        full = clean_text(raw_desc if raw_desc else merged)
        basis, desc = split_basis_desc(full)
        if bi >= 0 and clean_text(row[bi]):
            basis = clean_text(row[bi])
        sev_raw = clean_text(row[si] if si >= 0 else "")
        sev_norm = norm(sev_raw)
        sev = "高" if sev_norm in ["高", "major", "high"] else ("一般" if sev_norm in ["一般", "低", "low", "minor"] else "中")
        subs = extract_subject_ids((row[subi] if subi >= 0 else "") + "\n" + full)
        title = clean_text(raw_title) or extract_title(full, category)
        issues.append({"category": category, "title": title, "severity": sev, "subject_ids": subs, "basis": basis or "—", "description": desc or full or "—", "full_text": full or "—"})
    return issues


def parse_issue_blocks(rows: list[list[str]]) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    current_cat = ""
    for row in rows:
        cells = [x for x in row if x]
        if not cells:
            continue
        merged = row_text(row)
        if len(cells) == 1:
            maybe = normalize_category(cells[0])
            if maybe != "其他" and len(cells[0]) < 80:
                current_cat = maybe
                continue
        cat = normalize_category(cells[0]) if cells else "其他"
        if cat == "其他" and current_cat:
            cat = current_cat
        if cat == "其他" and not any(k in merged.lower() for k in ["问题", "依据", "icf", "edc", "ae", "sae", "方案", "原始", "药品", "样本"]):
            continue
        if len(merged) < 20 or any(w in merged for w in STOP_WORDS):
            continue
        basis, desc = split_basis_desc(merged)
        issues.append({"category": cat, "title": extract_title(desc, cat), "severity": "中", "subject_ids": extract_subject_ids(merged), "basis": basis, "description": desc, "full_text": merged})
    return issues


def parse_issue_table(rows: list[list[str]]) -> list[dict[str, Any]]:
    for i, row in enumerate(rows):
        if is_header_row(row):
            found = parse_table_after_header(rows, i)
            if found:
                return found
    return parse_issue_blocks(rows)


def extract_meta(rows: list[list[str]], file_name: str) -> dict[str, str]:
    all_text = "\n".join(row_text(row) for row in rows if any(row))
    composite_project = find_label_value(rows, ["项目名称/方案编号", "项目名称", "方案名称", "试验名称"])
    file_protocol = extract_protocol(file_name)
    protocol = file_protocol if file_protocol.startswith("YHNK-") else (extract_protocol(composite_project) or find_label_value(rows, ["方案编号", "方案号", "项目编号"]) or file_protocol or extract_protocol(all_text))
    project_name = clean_text(composite_project.replace(protocol, "").strip("/ -_：:")) if protocol else composite_project
    sponsor = find_label_value(rows, ["申办者", "申办方", "委托方"]) or "—"
    composite_center = find_label_value(rows, ["研究中心名称/中心编号/研究者姓名", "研究中心名称", "中心名称", "机构名称"])
    center = ""; center_no = ""; pi = ""
    if composite_center:
        parts = [clean_text(x) for x in re.split(r"[/／]", composite_center) if clean_text(x)]
        center = parts[0] if parts else composite_center
        if len(parts) >= 2:
            center_no = re.sub(r"\D", "", parts[1]) or parts[1]
        if len(parts) >= 3:
            pi = parts[2]
    pi = pi or find_label_value(rows, ["研究者姓名", "PI", "主要研究者", "主要研究者姓名"]) or "—"
    audit_date = find_label_value(rows, ["稽查实施日期", "稽查日期", "稽查时间"]) or "—"
    audit_company = find_label_value(rows, ["稽查公司", "稽查方"]) or "北京万宁睿和医药科技有限公司"
    enrollment = find_label_value(rows, ["中心入组情况", "入组情况", "筛选/入组情况"]) or "—"
    audit_note = find_label_value(rows, ["稽查实施情况", "稽查情况", "实施情况"]) or ""
    if not center:
        m = re.search(r"项目[-_](.+?)(?:（|\(|-中心稽查|_中心稽查)", file_name)
        if m:
            center = clean_text(m.group(1))
    if not center_no:
        m = re.search(r"[（(](\d{1,3})(?:中心)?[）)]", file_name)
        if m:
            center_no = m.group(1)
    return {"protocol_no": protocol or "—", "project_name": project_name or "—", "sponsor": sponsor or "—", "center_name": center or "—", "center_no": center_no or "", "pi": pi or "—", "audit_date": audit_date or "—", "audit_company": audit_company or "—", "enrollment": enrollment or "—", "audit_note": audit_note or ""}


def find_suggestions(rows: list[list[str]]) -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    start = -1
    for i, row in enumerate(rows):
        if "建议项" in row_text(row):
            start = i
            break
    if start == -1:
        return out
    current_cat = "其他"
    for row in rows[start + 1:]:
        if not any(row):
            continue
        merged = row_text(row)
        if is_header_row(row):
            break
        cat = normalize_category(merged)
        if cat != "其他" and len(merged) < 100:
            current_cat = cat
            continue
        text = clean_text(" ".join(x for x in row if x and norm(x) not in ["na", "n/a"]))
        if text and len(text) > 8:
            out.append({"category": current_cat, "text": text})
    return out


def parse_excel(excel_path: str | Path) -> dict[str, Any]:
    excel_path = Path(excel_path)
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    all_rows: list[list[str]] = []
    issues: list[dict[str, Any]] = []
    suggestions: list[dict[str, str]] = []
    for ws in wb.worksheets:
        rows = rows_from_ws(ws)
        all_rows.extend(rows)
        issues.extend(parse_issue_table(rows))
        suggestions.extend(find_suggestions(rows))
    uniq: list[dict[str, Any]] = []
    seen = set()
    for it in issues:
        key = (it["category"], it["title"], it["basis"], it["description"])
        if key not in seen:
            seen.add(key)
            uniq.append(it)
    issues = uniq
    meta = extract_meta(all_rows, excel_path.name)
    counts = {c: 0 for c in STANDARD_CATEGORIES}
    for it in issues:
        counts[it["category"]] = counts.get(it["category"], 0) + 1
    audited: list[str] = []
    center_prefix = meta.get("center_no", "")
    for sid in extract_subject_ids(meta.get("audit_note", "")):
        sid = (center_prefix + sid) if center_prefix and sid.isdigit() and len(sid) == 3 else sid
        if sid not in audited:
            audited.append(sid)
    for it in issues:
        for sid in it["subject_ids"]:
            sid = (center_prefix + sid) if center_prefix and sid.isdigit() and len(sid) == 3 else sid
            if sid not in audited:
                audited.append(sid)
    if not suggestions:
        suggestions = [{"category": cat, "text": f"{cat}：建议针对相关问题开展复核、整改和闭环跟踪。"} for cat, cnt in counts.items() if cnt > 0]
    return {"source_excel": excel_path.name, "meta": meta, "issues": issues, "summary": counts, "audited_subjects": audited, "suggestions": suggestions, "standard_categories": STANDARD_CATEGORIES}
