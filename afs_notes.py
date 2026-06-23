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

def read_fig_note(ws):
    """Note_XX sheet: Description in col B, CY in col C, PY in col D. Keep line items that have a
    value (either year) plus the Total row; skip headers / CHECK / guidance text."""
    rows=[]; started=False
    for r in ws.iter_rows(min_row=1,max_row=70,max_col=4,values_only=True):
        a=r[0]; b=r[1]; cy=r[2]; py=r[3]
        bs=str(b).strip() if b is not None else ""
        if not started:
            if bs.lower()=="description" or (a and str(a).strip().lower()=="code"): started=True
            continue
        if not bs: continue
        bl=bs.lower()
        if (bl.startswith("the ") or bl.startswith("tip") or "auto-pull" in bl or "must = 0" in bl
                or bs.upper().startswith("CHECK") or "difference" in bl or "reconcil" in bl
                or " per note" in bl or "per ls" in bl or "per tb" in bl or "per sofp" in bl
                or "per soci" in bl or "row 9" in bl or "row 30" in bl):
            continue
        cyv=_f(cy); pyv=_f(py)
        tot = bl.startswith("total")
        rows.append([bs, cyv or 0, pyv or 0, ("total" if tot else "")])
    line_items=[r for r in rows if r[3]!="total" and (r[1] or r[2])]
    totals=[r for r in rows if r[3]=="total"]
    if line_items: return line_items+totals
    return totals or rows[-1:]   # nothing coded -> just the total line

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
    notes=[]; ref=[]; num=start_num
    for key,title,sheet,kws,special in CANON:
        present = sheet in wb.sheetnames
        if not present: continue
        nd={"title":f"{num}. {title}"}
        if special=="ppe":
            nd["ppe"]=read_ppe_schedule(wb[sheet])
            if not (nd["ppe"]["classes"] or any(nd["ppe"]["total"])): continue
        else:
            tbl=read_fig_note(wb[sheet])
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
        line_items=[r for r in tbl if r[3]!="total"]
        tot_zero=all(abs(r[1])<1 and abs(r[2])<1 for r in tbl)
        if line_items or not tot_zero: continue
        num=str(nd["_num"]); cy=py=0.0
        for sr in stmt_rows:
            if sr.get("kind") in (None,"normal") and str(sr.get("note",""))==num:
                cy+=sr.get("cy") or 0; py+=sr.get("py") or 0
        if abs(cy)>=1 or abs(py)>=1:
            nd["table"]=[[f"Total {nd['_title'].lower()}", abs(cy), abs(py), "total"]]
