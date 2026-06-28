# -*- coding: utf-8 -*-
"""Asset-management notes reader. For SEC fund-manager workbooks that carry their own
Notes narrative + Note_XX figure sheets + Capital Adequacy schedule."""
import re
import afs_notes

def _num(v):
    return v if isinstance(v,(int,float)) else None

def _read_note_table(ws):
    """Note_XX figure sheet: description in col B, CY in C, PY in D. Skips CHECK rows."""
    out=[]; started=False
    for r in ws.iter_rows(min_row=1,max_row=70,max_col=4,values_only=True):
        b=r[1]; cy=r[2]; py=r[3]
        bs=str(b).strip() if b is not None else ""
        if not started:
            head=" ".join(str(x) for x in r if x is not None).upper()
            if "CY" in head and ("PY" in head or "₦" in head or "NGN" in head): started=True
            continue
        if not bs or bs.upper().startswith("CHECK"): continue
        cyv=_num(cy); pyv=_num(py)
        if cyv is None and pyv is None: continue          # drop sub-headers / blank
        tot = bs.upper().startswith("TOTAL") or bs.lower().startswith("at end")
        out.append([bs, cyv or 0, pyv or 0] + (["total"] if tot else []))
    return out

def _read_captable(ws):
    """Capital Adequacy: label A, CY B, PY C."""
    out=[]; started=False
    for r in ws.iter_rows(min_row=1,max_row=40,max_col=3,values_only=True):
        a=r[0]; cy=r[1]; py=r[2]
        ay=str(a).strip() if a is not None else ""
        if not started:
            if ay.lower().startswith("component"): started=True
            continue
        if not ay: continue
        cyv=_num(cy); pyv=_num(py)
        if cyv is None and pyv is None: continue
        tot = ay.upper().startswith("TOTAL") or "ratio" in ay.lower() or "excess" in ay.lower() or "minimum required" in ay.lower()
        out.append([ay, cyv or 0, pyv or 0] + (["total"] if ay.upper().startswith("TOTAL") else []))
    return out

def build_am_notes(wb):
    if "Notes" not in wb.sheetnames: return []
    ns=wb["Notes"]; raw=[]
    cur=None
    for r in ns.iter_rows(min_row=4,max_row=110,max_col=3,values_only=True):
        a=r[0]; b=r[1]
        a_s=str(a).strip() if a is not None else ""
        b_s=str(b).strip() if b is not None else ""
        if a_s and re.match(r'^\d+\.', a_s):
            if cur: raw.append(cur)
            ref=re.search(r'Note_\d+\w*', a_s)
            title=re.split(r'\s*[—-]\s*see', a_s)[0].strip()
            cur={"title":title,"paras":[],"ref":(ref.group(0) if ref else None)}
        elif b_s and cur is not None and cur["ref"] is None:
            cur["paras"].append(afs_notes.clean_para(b_s))
    if cur: raw.append(cur)
    # build generator-ready notes
    out=[]
    for n in raw:
        nd={"title":n["title"]}
        if n["ref"] and n["ref"] in wb.sheetnames:
            tbl=_read_note_table(wb[n["ref"]])
            if tbl: nd["table"]=tbl
        if n["paras"]:
            nd["paras"]=n["paras"][:14]
        # attach capital adequacy schedule to the capital-management note
        if "capital" in n["title"].lower() and ("regulat" in n["title"].lower() or "management" in n["title"].lower()):
            if "Capital_Adequacy" in wb.sheetnames:
                cap=_read_captable(wb["Capital_Adequacy"])
                if cap: nd["table"]=cap
        if nd.get("paras") or nd.get("table"):
            out.append(nd)
    return out
