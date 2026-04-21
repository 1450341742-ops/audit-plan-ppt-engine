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
    "知情同意书":"知情同意书（ICF）的签署和记录",
    "icf":"知情同意书（ICF）的签署和记录",
    "源文件":"原始文件的建立、内容和记录",
    "原始文件":"原始文件的建立、内容和记录",
    "原始文件的建立，内容和记录":"原始文件的建立、内容和记录",
    "his":"门诊/住院HIS、LIS、PACS等系统数据溯源",
    "lis":"门诊/住院HIS、LIS、PACS等系统数据溯源",
    "pacs":"门诊/住院HIS、LIS、PACS等系统数据溯源",
    "方案及其他文件依从性":"方案依从性",
    "方案依从":"方案依从性",
    "方案偏离":"方案依从性",
    "疗效":"药物疗效/研究评价指标的评估",
    "终点":"药物疗效/研究评价指标的评估",
    "安全性":"安全性信息评估，记录与报告",
    "安全性信息评估、记录与报告":"安全性信息评估，记录与报告",
    "ae":"安全性信息评估，记录与报告",
    "sae":"安全性信息评估，记录与报告",
    "crf":"CRF填写（时效性、一致性、溯源性、完整性）",
    "edc":"CRF填写（时效性、一致性、溯源性、完整性）",
    "药品":"试验用药品管理",
    "样本":"生物样本管理",
    "必须文件":"临床研究必须文件",
    "年度报告":"临床研究必须文件",
    "研究者文件夹":"临床研究必须文件",
    "cro":"申办方/CRO职责",
    "申办方":"申办方/CRO职责",
    "其他":"其他",
}

def clean_text(v:Any)->str:
    if v is None:
        return ""
    s=str(v).replace("\r","\n")
    s=re.sub(r"[ \t\xa0]+"," ",s)
    s=re.sub(r"\n{3,}","\n\n",s)
    return s.strip()

def norm(s:str)->str:
    s=clean_text(s).lower()
    s=s.replace("（","(").replace("）",")").replace("，","、").replace(",","、").replace(" ","")
    return s

def normalize_category(raw:str)->str:
    raw=clean_text(raw)
    if not raw:
        return "其他"
    nr = norm(raw)
    for c in STANDARD_CATEGORIES:
        if norm(c)==nr:
            return c
    for k,v in ALIASES.items():
        if norm(k)==nr:
            return v
    for key, target in ALIASES.items():
        if key.lower() in raw.lower():
            return target
    return "其他"

def iter_rows(ws):
    return [[clean_text(c) for c in row] for row in ws.iter_rows(values_only=True)]

def find_label_value(rows, labels:list[str])->str:
    labs=[norm(x) for x in labels]
    for row in rows:
        for i,cell in enumerate(row):
            if not cell:
                continue
            nc=norm(cell)
            # 支持“项目名称/方案编号：”“研究中心名称/中心编号/研究者姓名：”这类复合标签
            if any(lab == nc or lab in nc for lab in labs):
                m=re.match(r"^.*?[：:]\s*(.+)$", cell)
                if m and clean_text(m.group(1)):
                    return clean_text(m.group(1))
                for j in range(i+1, len(row)):
                    if row[j]:
                        return row[j]
    return ""

def extract_protocol(text:str)->str:
    text=clean_text(text)
    m=re.search(r"(?<![A-Z0-9])([A-Z]{2,}(?:-[A-Z0-9]+)+)(?![A-Z0-9-])", text, flags=re.I)
    if m:
        return m.group(1).upper().rstrip("-_/)")
    m=re.search(r"\b([A-Z]{2,}[A-Z0-9\-_/]*\d[A-Z0-9\-_/]*)\b", text, flags=re.I)
    return m.group(1).upper().rstrip("-_/)") if m else ""

def extract_subject_ids(text:str)->list[str]:
    text=clean_text(text)
    hits=[]
    for pat in [r"筛选号[:：]?\s*([ST]?\d{3,6})", r"受试者文件夹[:：]?\s*([0-9、,，/ ]{3,})", r"\b([ST]\d{3,6})\b", r"\b(T\d{3,6})\b", r"\b(\d{5})\b"]:
        for m in re.findall(pat, text, flags=re.I):
            raw=clean_text(m).upper()
            parts = re.split(r"[、,，/ ]+", raw) if re.fullmatch(r"[0-9、,，/ ]{3,}", raw) else [raw]
            for s in parts:
                s=clean_text(s)
                if not s or re.fullmatch(r"20\d{2}", s):
                    continue
                if s not in hits:
                    hits.append(s)
    return hits

def split_basis_desc(text:str)->tuple[str,str]:
    text=clean_text(text)
    if not text:
        return "—","—"
    m = re.search(r"(依据[:：].+?)(?=\n(?:问题[:：]|筛选号|受试者|备注[:：]|$)|$)", text, flags=re.S)
    if m:
        basis = clean_text(re.sub(r"^依据[:：]\s*", "", m.group(1)))
        desc = clean_text(text.replace(m.group(1), "", 1))
        return basis or "—", desc or "—"
    markers=["方案（","方案(","研究者手册","药物临床试验质量管理规范","GCP","中国酒精使用障碍防治指南"]
    for mk in markers:
        pos=text.find(mk)
        if pos!=-1:
            part=text[pos:]
            stop=re.search(r"(?=\n筛选号|\n受试者|\n问题[:：]|$)", part, flags=re.S)
            if stop:
                basis=clean_text(part[:stop.start()])
                desc=clean_text(text.replace(basis,"",1))
                return basis or "—", desc or "—"
    return "—", text

def extract_title(desc:str, category:str)->str:
    desc=clean_text(desc)
    for pat in [r"标题[:：]\s*(.+?)(?:\n|$)", r"问题[:：]\s*(.+?)(?:\n|$)"]:
        m=re.search(pat, desc, flags=re.S)
        if m:
            return clean_text(m.group(1))[:60]
    lines=[x.strip("—- ").strip() for x in desc.splitlines() if x.strip()]
    for line in lines[:4]:
        if len(line)<=60 and "筛选号" not in line:
            return line
    return category

def parse_issue_table(rows):
    issues=[]
    header_idx=None
    for i,row in enumerate(rows):
        merged=" ".join(x for x in row if x)
        if ("问题分类" in merged or "分类" in merged) and ("描述" in merged or "问题描述" in merged or "总结描述" in merged):
            header_idx=i;break
    if header_idx is None:
        return issues
    header=rows[header_idx]

    def find_col(cands):
        for idx,h in enumerate(header):
            nh=norm(h)
            for c in cands:
                if norm(c) in nh or nh in norm(c):
                    return idx
        return -1
    ci=find_col(["问题分类","分类"])
    di=find_col(["问题描述","描述","总结描述"])
    ti=find_col(["问题概述","标题","总结描述"])
    bi=find_col(["依据","法规依据"])
    si=find_col(["级别","问题级别","严重程度"])
    subi=find_col(["受试者","筛选号","受试者编号"])

    for row in rows[header_idx+1:]:
        if not any(row): 
            continue
        merged=" ".join(x for x in row if x)
        if "建议项" in merged:
            break
        raw_cat = row[ci] if ci>=0 else ""
        raw_desc = row[di] if di>=0 else merged
        raw_title = row[ti] if ti>=0 else ""
        if not clean_text(raw_desc) and not clean_text(raw_title):
            continue
        category = normalize_category(raw_cat or raw_title or raw_desc)
        full = clean_text(raw_desc if raw_desc else merged)
        basis, desc = split_basis_desc(full)
        if bi>=0 and clean_text(row[bi]):
            basis = clean_text(row[bi])
        sev = clean_text(row[si] if si>=0 else "")
        sev = "高" if sev in ["高","major","high"] else ("中" if sev in ["中","medium","moderate"] else ("一般" if sev in ["一般","低","low"] else "中"))
        subs = extract_subject_ids((row[subi] if subi>=0 else "") + "\n" + full)
        title = clean_text(raw_title) or extract_title(full, category)
        issues.append({
            "category": category,
            "title": title,
            "severity": sev,
            "subject_ids": subs,
            "basis": basis or "—",
            "description": desc or full or "—",
            "full_text": full or "—",
        })
    return issues

def extract_meta(rows, file_name:str):
    all_text="\n".join(" ".join(x for x in row if x) for row in rows if any(row))
    composite_project = find_label_value(rows, ["项目名称/方案编号","项目名称","方案名称","试验名称"]) or ""
    file_protocol = extract_protocol(file_name)
    protocol = file_protocol if file_protocol.startswith("YHNK-") else (extract_protocol(composite_project) or find_label_value(rows, ["方案编号","方案号","项目编号"]) or file_protocol or extract_protocol(all_text))
    project_name = clean_text(composite_project.replace(protocol, "").strip("/ -_：:")) if protocol else composite_project
    if (not project_name or project_name == "—") and protocol == "YHNK-XY-2-2021-01":
        project_name = "天麻苄醇酯苷片治疗轻中度血管性痴呆的有效性和安全性的随机、双盲、安慰剂、平行对照、多中心临床试验"
    sponsor = find_label_value(rows, ["申办者","申办方","委托方"]) or "—"
    if "审核" in sponsor or "回复" in sponsor:
        sponsor = "—"
    composite_center = find_label_value(rows, ["研究中心名称/中心编号/研究者姓名","研究中心名称","中心名称","机构名称"]) or ""
    center = ""; center_no = ""; pi = ""
    if composite_center:
        parts=[clean_text(x) for x in re.split(r"[/／]", composite_center) if clean_text(x)]
        center = parts[0] if parts else composite_center
        if len(parts) >= 2:
            center_no = re.sub(r"\D", "", parts[1]) or parts[1]
        if len(parts) >= 3:
            pi = parts[2]
    pi = pi or find_label_value(rows, ["研究者姓名","PI","主要研究者","主要研究者姓名"]) or "—"
    audit_date = find_label_value(rows, ["稽查实施日期","稽查日期","稽查时间"]) or "—"
    audit_company = find_label_value(rows, ["稽查公司","稽查方"]) or "北京万宁睿和医药科技有限公司"
    enrollment = find_label_value(rows, ["中心入组情况","入组情况","筛选/入组情况"]) or "—"
    audit_note = find_label_value(rows, ["稽查实施情况","稽查情况","实施情况"]) or ""
    if not center:
        m=re.search(r"项目[-_](.+?)(?:（|\(|-中心稽查|_中心稽查)", file_name)
        if m:
            center=clean_text(m.group(1))
    if not center_no:
        m=re.search(r"[（(](\d{1,3})(?:中心)?[）)]", file_name)
        if m: center_no=m.group(1)
    return {
        "protocol_no": protocol or "—",
        "project_name": project_name or "—",
        "sponsor": sponsor or "—",
        "center_name": center or "—",
        "center_no": center_no or "",
        "pi": pi or "—",
        "audit_date": audit_date or "—",
        "audit_company": audit_company or "—",
        "enrollment": enrollment or "—",
        "audit_note": audit_note or "",
    }

def find_suggestions(rows):
    start=-1
    header_idx=-1
    for i,row in enumerate(rows):
        merged=" ".join(x for x in row if x)
        if "建议项" in merged:
            start=i
            for k in range(i-1, max(-1, i-30), -1):
                if any("问题分类" in x for x in rows[k]) and any(("问题描述" in x or "问题概述" in x) for x in rows[k]):
                    header_idx=k
                    break
            break
    out=[]
    if start==-1:
        return out
    header=rows[header_idx] if header_idx>=0 else []
    def find_col(cands):
        for idx,h in enumerate(header):
            nh=norm(h)
            for c in cands:
                nc=norm(c)
                if nc in nh or nh in nc:
                    return idx
        return -1
    ci=find_col(["问题分类","分类"])
    di=find_col(["问题描述","描述","建议","问题概述"])
    if di < 0:
        di = 4
    for row in rows[start+1:]:
        if not any(row):
            continue
        merged=" ".join(x for x in row if x)
        if "问题分类" in merged and ("问题描述" in merged or "问题概述" in merged):
            break
        category=normalize_category(row[ci] if 0 <= ci < len(row) else "")
        if category=="其他":
            for item in row:
                t=clean_text(item)
                if t and normalize_category(t)!="其他" and len(t)<=60:
                    category=normalize_category(t); break
        text=clean_text(row[di] if 0 <= di < len(row) else "")
        if not text:
            text=" ".join(row[:5])
        text=re.sub(r"(^|\s)NA(\s|$)", " ", text)
        text=re.sub(r"\s{2,}", " ", text).strip()
        if text:
            out.append({"category":category, "text":text})
    return out

def parse_excel(excel_path:str|Path)->dict[str,Any]:
    excel_path=Path(excel_path)
    wb=openpyxl.load_workbook(excel_path, data_only=True)
    all_rows=[]; issues=[]; suggestions=[]
    for ws in wb.worksheets:
        rows=iter_rows(ws)
        all_rows.extend(rows)
        issues.extend(parse_issue_table(rows))
        suggestions.extend(find_suggestions(rows))
    # dedupe
    uniq=[]; seen=set()
    for it in issues:
        key=(it["category"],it["title"],it["basis"],it["description"])
        if key not in seen:
            seen.add(key); uniq.append(it)
    issues=uniq
    meta=extract_meta(all_rows, excel_path.name)
    counts={c:0 for c in STANDARD_CATEGORIES}
    for it in issues:
        counts[it["category"]] = counts.get(it["category"],0)+1
    aud=[]
    center_prefix = meta.get("center_no", "")
    for sid in extract_subject_ids(meta.get("audit_note","")):
        sid = (center_prefix + sid) if center_prefix and sid.isdigit() and len(sid)==3 else sid
        if sid not in aud:
            aud.append(sid)
    for it in issues:
        for sid in it["subject_ids"]:
            sid = (center_prefix + sid) if center_prefix and sid.isdigit() and len(sid)==3 else sid
            if sid not in aud:
                aud.append(sid)
    if not suggestions:
        suggestions=[{"category":cat, "text":f"{cat}：建议针对相关问题开展复核、整改和闭环跟踪。"} for cat,cnt in counts.items() if cnt>0]
    return {
        "source_excel": excel_path.name,
        "meta": meta,
        "issues": issues,
        "summary": counts,
        "audited_subjects": aud,
        "suggestions": suggestions,
        "standard_categories": STANDARD_CATEGORIES,
    }
