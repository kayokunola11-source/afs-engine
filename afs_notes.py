# -*- coding: utf-8 -*-
"""Detailed notes for the SME / full template: read the workbook's own Note_XX figure
sheets + PPE_Schedule, number the notes sequentially, and remap the statement note
references so they tally."""

def _f(v):
    if isinstance(v,(int,float)): return float(v)
    if isinstance(v,str):
        t=v.replace(",","").strip()
        try: return float(t)
        except ValueError: return None
    return None


def build_tb_index(wb):
    """Map account code -> (CY, PY) magnitude, read directly from the trial balance. This bypasses
    the note sheets' SUMIFS formulas, which LibreOffice evaluates to 0 on number/text criteria."""
    tb=None
    for sn in wb.sheetnames:
        if sn.strip().lower()=="tb": tb=wb[sn]; break
    if tb is None: return {}
    hdr=None; col={}
    for i,r in enumerate(tb.iter_rows(min_row=1,max_row=10,max_col=16,values_only=True),1):
        labs=[str(c).strip().lower() if c is not None else "" for c in r]
        if "code" in labs and ("cy dr" in labs or "cy cr" in labs):
            hdr=i
            for j,l in enumerate(labs):
                if l in ("code","cy dr","cy cr","py dr","py cr"): col[l]=j
            break
    if hdr is None or "code" not in col: return {}
    def g(r,key):
        j=col.get(key); v=r[j] if (j is not None and j<len(r)) else None
        return v if isinstance(v,(int,float)) else 0.0
    idx={}
    for r in tb.iter_rows(min_row=hdr+1,max_row=hdr+300,max_col=16,values_only=True):
        c=r[col["code"]]
        if c is None or str(c).strip()=="": continue
        code=str(c).strip()
        if code.endswith(".0"): code=code[:-2]
        cy=abs(g(r,"cy dr")-g(r,"cy cr")); py=abs(g(r,"py dr")-g(r,"py cr"))
        idx[code]=(cy,py)
    return idx

def read_fig_note(ws, tb=None):
    """Read a Note_XX figure note. Values come from the TB (by the code in col A) when available,
    otherwise from the sheet's own columns. Skips headers / CHECK / reconciliation rows."""
    items=[]; total_label=None
    for r in ws.iter_rows(min_row=1,max_row=70,max_col=4,values_only=True):
        a=r[0]; b=r[1]; cy=r[2]; py=r[3]
        code=str(a).strip() if a is not None else ""
        if code.endswith(".0"): code=code[:-2]
        bs=str(b).strip() if b is not None else ""
        if not bs: continue
        bl=bs.lower()
        if bl=="description" or code.lower()=="code": continue
        if (bl.startswith("the ") or bl.startswith("tip") or "auto-pull" in bl or "must = 0" in bl
                or bs.upper().startswith("CHECK") or "difference" in bl or "reconcil" in bl
                or " per note" in bl or "per ls" in bl or "per tb" in bl or "per sofp" in bl
                or "per soci" in bl or "row 9" in bl or "row 30" in bl):
            continue
        if bl.startswith("total"):
            total_label=bs; continue
        if tb and code and code in tb:
            cyv,pyv=tb[code]
        else:
            cyv=_f(cy) or 0; pyv=_f(py) or 0
        items.append([bs, cyv, pyv])
    line_items=[r for r in items if abs(r[1])>=1 or abs(r[2])>=1]
    if not line_items:
        return [[total_label or "Total", 0, 0, "total"]]   # uncoded -> fill_uncoded_totals handles
    tcy=sum(r[1] for r in line_items); tpy=sum(r[2] for r in line_items)
    return line_items + [[total_label or "Total", tcy, tpy, "total"]]

def read_ppe_schedule(ws):
    classes=[]; total=None; header=False
    for r in ws.iter_rows(min_row=4,max_row=22,max_col=12,values_only=True):
        a=r[0]; an=str(a).strip() if a is not None else ""
        if an=="Asset class": header=True; continue
        if not header or not an: continue
        cost=_f(r[5]) or 0; dep=_f(r[10]) or 0; charge=_f(r[7]) or 0; nbv=_f(r[11]) or 0
        if an.upper()=="TOTAL": total=[cost,dep,charge,nbv]; break
        if cost or dep or nbv or charge: classes.append([an,cost,dep,charge,nbv])
    return {"classes":classes,"total":total or [0,0,0,0]}

# canonical order: (key, title, source-sheet, match-keywords, special)
CANON=[
 ("revenue","Revenue","Note_06_Revenue",["revenue"],None),
 ("otherinc","Other Income","Note_06b_OtherIncome",["other income"],None),
 ("cos","Cost of Sales","Note_07_CostOfSales",["cost of sale"],None),
 ("admin","Administrative Expenses","Note_08_AdminExpenses",["administrative"],None),
 ("selling","Selling and Distribution Expenses","Note_19_SellingExpenses",["selling and distribution"],None),
 ("finance","Finance Cost","Note_09_FinanceCost",["finance cost"],None),
 ("tax","Taxation","Note_17_TaxPayable",["taxation","provision for tax","current tax payable"],None),
 ("ppe","Property, Plant and Equipment","PPE_Schedule",["property, plant","property plant"],"ppe"),
 ("recv","Trade and Other Receivables","Note_11_TradeReceivables",["receivable"],None),
 ("cash","Cash and Cash Equivalents","Note_12_Cash",["cash and cash","cash & cash"],None),
 ("sharecap","Share Capital","Note_13_ShareCapital",["share capital"],None),
 ("retearn","Retained Earnings","Note_14_RetainedEarnings",["retained earning"],None),
 ("ltloan","Long-term Loans","Note_15_LongTermLoan",["long-term loan","long term loan"],None),
 ("payables","Trade and Other Payables","Note_16_TradePayables",["payable","accrual"],None),
 ("dca","Directors' Current Account","Note_18_DirCurrentAccount",["directors' current","director's current","directors current"],None),
]

def build_figure_notes(wb, start_num, stmt_rows=None):
    """Return (notes, ref_map). notes are generator-ready dicts numbered from start_num;
    ref_map maps each note's match keywords to its number for statement remapping."""
    tb=build_tb_index(wb)
    notes=[]; ref=[]; num=start_num
    for key,title,sheet,kws,special in CANON:
        present = sheet in wb.sheetnames
        if not present: continue
        nd={"title":f"{num}. {title}"}
        if special=="ppe":
            nd["ppe"]=read_ppe_schedule(wb[sheet])
            if not (nd["ppe"]["classes"] or any(nd["ppe"]["total"])): continue
        else:
            tbl=read_fig_note(wb[sheet], tb)
            if not tbl: continue
            nd["table"]=tbl
        nd["_num"]=num; nd["_title"]=title
        notes.append(nd); ref.append((kws,num)); num+=1
    return notes, ref

def remap_statement_refs(rows, ref):
    for r in rows:
        if r.get("kind") not in (None,"normal"): continue
        lab=r.get("label","").lower()
        if not lab: continue
        for kws,num in ref:
            if any(k in lab for k in kws):
                r["note"]=str(num); break
    return rows


def fill_uncoded_totals(notes, stmt_rows):
    """For figure notes left uncoded in the workbook (only a ~0 total), total the matching
    statement lines (by their remapped note number) so the note ties to the face of the accounts."""
    for nd in notes:
        if "_num" not in nd or "table" not in nd: continue
        tbl=nd["table"]
        line_items=[r for r in tbl if not (len(r)>3 and r[3]=="total")]
        tot_zero=all(abs(r[1])<1 and abs(r[2])<1 for r in tbl)
        if line_items or not tot_zero: continue
        num=str(nd["_num"]); cy=py=0.0
        for sr in stmt_rows:
            if sr.get("kind") in (None,"normal") and str(sr.get("note",""))==num:
                cy+=sr.get("cy") or 0; py+=sr.get("py") or 0
        if abs(cy)>=1 or abs(py)>=1:
            nd["table"]=[[f"Total {nd['_title'].lower()}", abs(cy), abs(py), "total"]]
