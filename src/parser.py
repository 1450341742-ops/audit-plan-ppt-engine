from __future__ import annotations
import re
from pathlib import Path
from typing import Any
import openpyxl

CATEGORIES = [
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
    "伦理":"伦理委员会审核要求的遵循", "知情":"知情同意书（ICF）的签署和记录", "ICF":"知情同意书（ICF）的签署和记录",
    "原始":"原始文件的建立、内容和记录", "源文件":"原始文件的建立、内容和记录", "HIS":"门诊/住院HIS、LIS、PACS等系统数据溯源",
    "LIS":"门诊/住院HIS、LIS、PACS等系统数据溯源", "PACS":"门诊/住院HIS、LIS、PACS等系统数据溯源",
    "方案":"方案依从性", "偏离":"方案依从性", "疗效":"药物疗效/研究评价指标的评估", "终点":"药物疗效/研究评价指标的评估",
    "安全":"安全性信息评估，记录与报告", "AE":"安全性信息评估，记录与报告", "SAE":"安全性信息评估，记录与报告",
    "CRF":"CRF填写（时效性、一致性、溯源性、完整性）", "EDC":"CRF填写（时效性、一致性、溯源性、完整性）",
    "药品":"试验用药品管理", "样本":"生物样本管理", "必须文件":"临床研究必须文件", "研究者文件夹":"临床研究必须文件",
    "申办方":"申办方/CRO职责", "CRO":"申办方/CRO职责", "授权":"国家药物临床试验政策法规的遵循", "法规":"国家药物临床试验政策法规的遵循",
}


def clean(v: Any) -> str:
    return "" if v is None else re.sub(r"\s+", " ", str(v).replace("\r", "\n")).strip()


def norm(s: str) -> str:
    return clean(s).lower().replace(" ", "").replace("，", "、").replace(",", "、")


def cat(raw: str) -> str:
    t = clean(raw)
    if not t:
        return "其他"
    nt = norm(t)
    for c in CATEGORIES:
        if norm(c) == nt or norm(c) in nt or nt in norm(c):
            return c
    for k, v in ALIASES.items():
        if k.lower() in t.lower():
            return v
    return "其他"


def all_rows(wb):
    rows = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            rows.append([clean(c) for c in row])
    return rows


def find_value(rows, keys):
    for row in rows:
        for i, cell in enumerate(row):
            if not cell:
                continue
            for k in keys:
                if norm(k) in norm(cell):
                    m = re.search(r"[:：]\s*(.+)$", cell)
                    if m and clean(m.group(1)):
                        return clean(m.group(1))
                    for j in range(i + 1, len(row)):
                        if row[j]:
                            return row[j]
    return ""


def protocol(text: str) -> str:
    m = re.search(r"([A-Z]{2,}[A-Z0-9]*(?:-[A-Z0-9]+)+)", text, re.I)
    return m.group(1).upper() if m else ""


def split_basis_desc(text: str):
    text = clean(text)
    m = re.search(r"依据[:：]\s*(.+?)(?:\n|问题[:：]|描述[:：]|$)", text, re.S)
    if m:
        basis = clean(m.group(1))
        desc = clean(text.replace(m.group(0), " ", 1)) or text
        return basis or "—", desc or "—"
    return "—", text or "—"


def subject_ids(text: str):
    ids = []
    for m in re.findall(r"\b(?:S|T)?\d{3,6}\b", text or "", re.I):
        if m not in ids and not re.fullmatch(r"20\d{2}", m):
            ids.append(m.upper())
    return ids


def parse_issues(rows):
    issues = []
    header_i = None
    for i, row in enumerate(rows):
        joined = " ".join(row)
        if ("问题分类" in joined or "分类" in joined) and ("描述" in joined or "问题" in joined):
            header_i = i
            break
    if header_i is None:
        return issues
    header = rows[header_i]

    def col(names, default=-1):
        for idx, h in enumerate(header):
            nh = norm(h)
            if any(norm(n) in nh or nh in norm(n) for n in names):
                return idx
        return default

    ci = col(["问题分类", "分类"], 0)
    ti = col(["问题概述", "标题", "问题标题"], -1)
    di = col(["问题描述", "描述", "总结描述"], -1)
    bi = col(["依据", "法规依据"], -1)
    si = col(["级别", "严重程度"], -1)

    for row in rows[header_i + 1:]:
        if not any(row):
            continue
        joined = " ".join(row)
        if "建议项" in joined:
            break
        raw_desc = row[di] if 0 <= di < len(row) else joined
        raw_title = row[ti] if 0 <= ti < len(row) else ""
        raw_cat = row[ci] if 0 <= ci < len(row) else ""
        if not raw_desc and not raw_title:
            continue
        basis, desc = split_basis_desc(raw_desc or joined)
        if 0 <= bi < len(row) and row[bi]:
            basis = row[bi]
        category = cat(raw_cat or raw_title or raw_desc)
        title = raw_title or (desc[:45] + "..." if len(desc) > 45 else desc) or category
        sev = row[si] if 0 <= si < len(row) and row[si] else "中"
        issues.append({"category": category, "title": title, "severity": sev, "basis": basis, "description": desc, "subject_ids": subject_ids(joined)})
    return issues


def parse_suggestions(rows, counts):
    start = -1
    for i, row in enumerate(rows):
        if "建议项" in " ".join(row):
            start = i
            break
    out = []
    if start >= 0:
        for row in rows[start + 1:]:
            if not any(row):
                continue
            text = " ".join([x for x in row if x and x.upper() != "NA"])
            if len(text) > 4:
                out.append({"category": cat(text), "text": text})
            if len(out) >= 18:
                break
    if not out:
        out = [{"category": k, "text": f"建议针对{k}相关问题进行原因分析、整改落实和闭环跟踪。"} for k, v in counts.items() if v]
    return out


def parse_excel(excel_path: str | Path) -> dict[str, Any]:
    excel_path = Path(excel_path)
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    rows = all_rows(wb)
    full_text = "\n".join(" ".join(r) for r in rows)
    pno = protocol(excel_path.name) or protocol(full_text) or "—"
    center = find_value(rows, ["研究中心", "中心名称", "机构名称"]) or "—"
    if center == "—":
        m = re.search(r"项目-(.+?)-中心稽查", excel_path.name)
        if m:
            center = m.group(1)
    meta = {
        "project_name": find_value(rows, ["项目名称", "方案名称", "试验名称"]) or "—",
        "protocol_no": pno,
        "sponsor": find_value(rows, ["申办方", "申办者", "委托方"]) or "—",
        "center_name": center,
        "center_no": find_value(rows, ["中心编号", "中心号"]) or "",
        "pi": find_value(rows, ["主要研究者", "研究者姓名", "PI"]) or "—",
        "audit_date": find_value(rows, ["稽查日期", "稽查实施日期", "稽查时间"]) or "—",
        "audit_company": "北京万宁睿和医药科技有限公司",
    }
    issues = parse_issues(rows)
    counts = {c: 0 for c in CATEGORIES}
    for it in issues:
        counts[it["category"]] = counts.get(it["category"], 0) + 1
    ids = []
    for it in issues:
        for sid in it.get("subject_ids", []):
            if sid not in ids:
                ids.append(sid)
    return {"source_excel": excel_path.name, "meta": meta, "issues": issues, "summary": counts, "audited_subjects": ids, "suggestions": parse_suggestions(rows, counts), "standard_categories": CATEGORIES}
