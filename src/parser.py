from __future__ import annotations
import re
from pathlib import Path
from typing import Any
import openpyxl

STANDARD_CATEGORIES = ["国家药物临床试验政策法规的遵循","伦理委员会审核要求的遵循","知情同意书（ICF）的签署和记录","原始文件的建立、内容和记录","门诊/住院HIS、LIS、PACS等系统数据溯源","方案依从性","药物疗效/研究评价指标的评估","安全性信息评估，记录与报告","CRF填写（时效性、一致性、溯源性、完整性）","试验用药品管理","生物样本管理","临床研究必须文件","申办方/CRO职责","其他"]
ALIASES={"法规":"国家药物临床试验政策法规的遵循","伦理":"伦理委员会审核要求的遵循","知情同意":"知情同意书（ICF）的签署和记录","icf":"知情同意书（ICF）的签署和记录","源文件":"原始文件的建立、内容和记录","原始文件":"原始文件的建立、内容和记录","原始记录":"原始文件的建立、内容和记录","his":"门诊/住院HIS、LIS、PACS等系统数据溯源","lis":"门诊/住院HIS、LIS、PACS等系统数据溯源","pacs":"门诊/住院HIS、LIS、PACS等系统数据溯源","方案依从":"方案依从性","方案偏离":"方案依从性","疗效":"药物疗效/研究评价指标的评估","recist":"药物疗效/研究评价指标的评估","肿瘤评估":"药物疗效/研究评价指标的评估","安全性":"安全性信息评估，记录与报告","ae":"安全性信息评估，记录与报告","sae":"安全性信息评估，记录与报告","crf":"CRF填写（时效性、一致性、溯源性、完整性）","edc":"CRF填写（时效性、一致性、溯源性、完整性）","药品":"试验用药品管理","药物":"试验用药品管理","样本":"生物样本管理","必须文件":"临床研究必须文件","研究者文件夹":"临床研究必须文件","cro":"申办方/CRO职责","申办方":"申办方/CRO职责","其他":"其他"}
STOP_WORDS=["审核","回复","签字","日期","备注","说明","填表","批准","确认"]

def clean_text(v:Any)->str:
    if v is None: return ""
    s=str(v).replace("\r","\n")
    s=re.sub(r"[ \t\xa0]+"," ",s); s=re.sub(r"\n{3,}","\n\n",s)
    return s.strip()
def norm(s:str)->str:
    s=clean_text(s).lower().replace("（","(").replace("）",")").replace("，","、").replace(",","、")
    return re.sub(r"\s+","",s)
def normalize_category(raw:str)->str:
    nr=norm(raw)
    if not nr: return "其他"
    for c in STANDARD_CATEGORIES:
        if norm(c)==nr or norm(c) in nr or nr in norm(c): return c
    for k,v in ALIASES.items():
        if norm(k) in nr: return v
    return "其他"
def rows_from_ws(ws): return [[clean_text(c) for c in row] for row in ws.iter_rows(values_only=True)]
def row_text(row): return " ".join(x for x in row if x)
def find_label_value(rows, labels, allow_long=False):
    labs=[norm(x) for x in labels]
    for row in rows:
        for i,cell in enumerate(row):
            if not cell: continue
            nc=norm(cell)
            if any(lab==nc or lab in nc for lab in labs):
                m=re.match(r"^.*?[：:]\s*(.+)$",cell)
                if m and clean_text(m.group(1)): return clean_text(m.group(1))
                for j in range(i+1,min(len(row),i+8)):
                    if row[j] and norm(row[j]) not in labs: return row[j]
    return ""
def extract_protocol(text):
    m=re.search(r"(?<![A-Z0-9])([A-Z]{2,}(?:-[A-Z0-9]+)+)(?![A-Z0-9-])",clean_text(text),flags=re.I)
    return m.group(1).upper().rstrip("-_/)") if m else ""
def extract_subject_ids(text):
    hits=[]
    for pat in [r"\b(S\d{5})\b",r"\b(T\d{3,6})\b",r"筛选号[:：]?\s*([ST]?\d{3,6})"]:
        for m in re.findall(pat,clean_text(text),flags=re.I):
            v=clean_text(m).upper()
            if v and v not in hits: hits.append(v)
    return hits
def find_col(header,cands):
    for idx,h in enumerate(header):
        nh=norm(h)
        for c in cands:
            nc=norm(c)
            if nc and (nc==nh or nc in nh or nh in nc): return idx
    return -1
def is_header_row(row):
    txt=norm(row_text(row))
    return ("问题分类" in txt or "问题类别" in txt or "分类" in txt) and ("问题描述" in txt or "总结描述" in txt or "问题概述" in txt or "稽查依据" in txt or "依据" in txt)
def basis_like(s):
    s=clean_text(s)
    return any(k in s for k in ["药物临床试验质量管理规范","第二十五条","RECIST","管理手册","方案","ICH","GCP","核查要点","依据"])
def desc_like(s): return len(clean_text(s))>=4 and not basis_like(s)
def merge_summary_desc(summary, desc):
    summary=clean_text(summary); desc=clean_text(desc)
    if summary and desc:
        if summary in desc: return desc
        return summary + "\n\n" + desc
    return desc or summary

def pick_fields_from_row(row,cat_idx,header=None):
    cells=[clean_text(x) for x in row]
    summary=""; desc=""; basis="—"
    if header:
        summary_col=find_col(header,["总结描述","问题概述","概述"])
        desc_col=find_col(header,["问题描述"])
        basis_col=find_col(header,["稽查依据","依据","法规依据","参考依据"])
        if 0<=summary_col<len(cells): summary=clean_text(cells[summary_col])
        if 0<=desc_col<len(cells): desc=clean_text(cells[desc_col])
        if 0<=basis_col<len(cells): basis=clean_text(cells[basis_col]) or "—"
    candidates=[cells[i] for i in range(cat_idx+1,min(len(cells),cat_idx+12)) if cells[i]]
    if not summary:
        for x in candidates:
            if desc_like(x) and len(x)<=220: summary=x; break
    if not desc:
        non_basis=[x for x in candidates if desc_like(x)]
        if non_basis: desc=max(non_basis,key=len)
    if basis in ["","—"]:
        for x in candidates:
            if x not in [summary,desc] and basis_like(x): basis=re.sub(r"^依据[:：]?","",x).strip() or "—"; break
    if not desc: desc=summary
    if not summary: summary=desc
    merged_desc=merge_summary_desc(summary, desc)
    return summary,merged_desc,basis

def parse_table_after_header(rows, header_idx):
    header=rows[header_idx]
    ci=find_col(header,["问题分类","问题类别","分类","检查内容"])
    si=find_col(header,["级别","问题级别","严重程度","风险等级"])
    issues=[]; current_cat=""
    for row in rows[header_idx+1:]:
        if not any(row): continue
        merged=row_text(row); nmerged=norm(merged)
        if any(x in nmerged for x in ["建议项","capa回复","纠正预防","审核人","批准人"]):
            if "建议项" in nmerged: break
        raw_cat=row[ci] if 0<=ci<len(row) else ""
        if raw_cat: current_cat=raw_cat
        category=normalize_category(raw_cat or current_cat or merged)
        if category=="其他" and norm(raw_cat)!="其他": continue
        summary,desc,basis=pick_fields_from_row(row,ci if ci>=0 else 0,header)
        if not desc or len(desc)<4: continue
        sev_raw=norm(row[si] if 0<=si<len(row) else "")
        sev="高" if sev_raw in ["高","major","high"] else ("一般" if sev_raw in ["一般","低","low","minor"] else "中")
        issues.append({"category":category,"title":"","summary":summary,"severity":sev,"subject_ids":extract_subject_ids(merged),"basis":basis or "—","description":desc,"full_text":desc})
    return issues

def parse_summary_rows(rows):
    issues=[]
    for row in rows:
        cells=[clean_text(x) for x in row]; merged=row_text(cells)
        if not merged or len(merged)<10: continue
        if any(x in merged for x in ["中心稽查概述","中心稽查范围","问题分类","建议项","CAPA","审核人","批准人"]): continue
        cat_idx=-1; cat="其他"
        for idx,cell in enumerate(cells):
            c=normalize_category(cell)
            if c!="其他" or norm(cell)=="其他": cat_idx=idx; cat=c; break
        if cat_idx<0: continue
        summary,desc,basis=pick_fields_from_row(cells,cat_idx,None)
        if not desc or len(desc)<4: continue
        issues.append({"category":cat,"title":"","summary":summary,"severity":"中","subject_ids":extract_subject_ids(merged),"basis":basis or "—","description":desc,"full_text":desc})
    return issues

def parse_issue_table(rows):
    for i,row in enumerate(rows):
        if is_header_row(row):
            found=parse_table_after_header(rows,i)
            if found: return found
    return parse_summary_rows(rows)

def extract_meta(rows,file_name):
    composite_project=find_label_value(rows,["项目名称/方案编号","项目名称","方案名称","试验名称"],allow_long=True)
    protocol=extract_protocol(file_name) or extract_protocol(composite_project)
    project_name=clean_text(composite_project.replace(protocol,"").strip("/ -_：:")) if protocol else composite_project
    center_raw=find_label_value(rows,["研究中心名称/中心编号/研究者姓名","研究中心名称","中心名称","机构名称"],allow_long=False)
    center=""; center_no=""; pi=""
    if center_raw:
        parts=[clean_text(x) for x in re.split(r"[/／]",center_raw) if clean_text(x)]
        center=parts[0] if parts else center_raw
        if len(parts)>=2: center_no=re.sub(r"\D","",parts[1]) or parts[1]
        if len(parts)>=3: pi=parts[2]
    pi=pi or find_label_value(rows,["研究者姓名","PI","主要研究者","主要研究者姓名"],allow_long=False) or "—"
    audit_date=find_label_value(rows,["稽查实施日期","稽查日期","稽查时间"],allow_long=False) or "—"
    audit_company=find_label_value(rows,["稽查公司","稽查方"],allow_long=False) or "北京万宁睿和医药科技有限公司"
    enrollment=find_label_value(rows,["中心入组情况","入组情况","筛选/入组情况"],allow_long=False) or "—"
    audit_note=find_label_value(rows,["稽查实施情况","稽查情况","实施情况"],allow_long=True) or ""
    if not center:
        m=re.search(r"项目[-_](.+?)(?:（|\(|-中心稽查|_中心稽查)",file_name)
        if m: center=clean_text(m.group(1))
    if not center_no:
        m=re.search(r"[（(](\d{1,3})(?:中心)?[）)]",file_name)
        if m: center_no=m.group(1)
    return {"protocol_no":protocol or "—","project_name":project_name or "—","sponsor":"","center_name":center or "—","center_no":center_no or "","pi":pi or "—","audit_date":audit_date or "—","audit_company":audit_company or "—","enrollment":enrollment or "—","audit_note":audit_note or ""}

def parse_excel(excel_path:str|Path)->dict[str,Any]:
    excel_path=Path(excel_path); wb=openpyxl.load_workbook(excel_path,data_only=True)
    all_rows=[]; issues=[]
    for ws in wb.worksheets:
        rows=rows_from_ws(ws); all_rows.extend(rows); issues.extend(parse_issue_table(rows))
    uniq=[]; seen=set()
    for it in issues:
        key=(it["category"],it["basis"],it["description"])
        if key not in seen: seen.add(key); uniq.append(it)
    issues=uniq; meta=extract_meta(all_rows,excel_path.name)
    counts={c:0 for c in STANDARD_CATEGORIES}
    for it in issues: counts[it["category"]]=counts.get(it["category"],0)+1
    audited=[]; center_prefix=meta.get("center_no","")
    for sid in extract_subject_ids(meta.get("audit_note","")+"\n"+"\n".join(i.get("full_text","") for i in issues)):
        sid=(center_prefix+sid) if center_prefix and sid.isdigit() and len(sid)==3 else sid
        if sid not in audited: audited.append(sid)
    suggestions=[{"category":cat,"text":f"{cat}：建议针对相关问题开展复核、整改和闭环跟踪。"} for cat,cnt in counts.items() if cnt>0]
    return {"source_excel":excel_path.name,"meta":meta,"issues":issues,"summary":counts,"audited_subjects":audited,"suggestions":suggestions,"standard_categories":STANDARD_CATEGORIES}
