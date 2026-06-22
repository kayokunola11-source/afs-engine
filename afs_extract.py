# -*- coding: utf-8 -*-
"""Generic extractor: read a recalculated firm-template workbook into the data dict
that afs_generator.build expects. Works on any client using the standard template
(sheets: Entity, Cover, SOCI, SOFP, SCF, SOCE, Note_10_PPE, ...)."""
import datetime, openpyxl

TEMPLATES = {
    "SME":          ("International Financial Reporting Standard for Small and Medium-sized Entities (IFRS for SMEs)","IFRS for SMEs"),
    "NGO":          ("applicable financial reporting framework for Not-for-Profit organisations in Nigeria","the NGO reporting framework"),
    "Manufacturing":("International Financial Reporting Standard for Small and Medium-sized Entities (IFRS for SMEs)","IFRS for SMEs"),
    "PFA":          ("the Pension Reform Act and applicable PenCom reporting guidelines","the PenCom framework"),
}
SIG_WORDS={1:"one",2:"two",3:"three",4:"four"}
SECTION_A={"ASSETS","EQUITY AND LIABILITIES"}
SECTION_B={"Non-current assets","Current assets","Equity","Non-current liabilities","Current liabilities"}
SCF_SECTION={"Cash flows from operating activities","Cash flows from investing activities","Cash flows from financing activities"}
SCF_SUBHEAD={"Adjustments for non-cash items:","Working capital changes:"}

def _num(v):
    if isinstance(v,(int,float)): return float(v)
    return None

def _fmtdate(v):
    if isinstance(v,datetime.datetime): return v.strftime("%-d %B %Y")
    return str(v) if v else ""

def _sheet_map(ws, key_col=2, val_col=3, maxr=40):
    out={}
    for r in ws.iter_rows(min_row=1,max_row=maxr,max_col=max(key_col,val_col),values_only=True):
        k=r[key_col-1]; v=r[val_col-1]
        if k and str(k).strip(): out[str(k).strip()]=v
    return out

def _is_total(label):
    L=label.upper()
    return (label.isupper() or label.startswith("Total") or "GROSS PROFIT" in L
            or "PROFIT/(LOSS)" in L or "OPERATING PROFIT" in L or "Net cash" in label
            or "Cash and cash equivalents at end" in label or label.startswith("Balance "))

def _kind_for(label, is_grand=False):
    L=label.upper(); ll=label.lower()
    if "net increase" in ll and "cash" in ll: return "total"
    if "cash" in ll and "end of" in ll: return "grandtotal"
    if is_grand or L in {"TOTAL ASSETS","TOTAL EQUITY AND LIABILITIES","TOTAL COMPREHENSIVE INCOME FOR THE YEAR","TOTAL COMPREHENSIVE INCOME"}:
        return "grandtotal"
    if "Cash and cash equivalents at end" in label: return "grandtotal"
    if label.startswith("Total") or L.startswith("OPERATING PROFIT") or L.startswith("PROFIT/(LOSS)") or L=="GROSS PROFIT" or label.startswith("Net cash"):
        return "total" if (label.startswith("Total") or "GROSS" in L or "FOR THE YEAR" in L) else "subtotal"
    return "normal"

def _read_statement(ws, drop_zero=True, maxr=70):
    rows=[]
    for r in ws.iter_rows(min_row=4,max_row=maxr,max_col=5,values_only=True):
        a=(r[0] or "").strip() if isinstance(r[0],str) else (r[0] or "")
        b=(r[1] or "").strip() if isinstance(r[1],str) else (r[1] or "")
        note=r[2]; cy=_num(r[3]); py=_num(r[4])
        def _add_section(lbl):
            if rows and rows[-1].get("kind") not in ("section","blank"):
                rows.append({"kind":"blank"})
            rows.append({"label":lbl,"kind":"section"})
        au=a.upper() if isinstance(a,str) else ""
        if au in SECTION_A: _add_section(a); continue
        if au.startswith("CASH FLOWS"): _add_section(a.title()); continue
        if not b:
            if a and not cy and not py: continue
            continue
        if b in SECTION_B or b in SCF_SECTION:
            _add_section(b); continue
        if b in SCF_SUBHEAD:
            rows.append({"label":b,"kind":"section"}); continue
        if b.upper().startswith("BS CHECK") or b.upper().startswith("CHECK") or b.startswith("Approved by") or b.startswith("___") or b in {"Director / Proprietor","Mode applied"} or b.startswith("Mode applied") or b.startswith("Basic earnings"):
            continue
        k=_kind_for(b)
        is_money = (cy is not None) or (py is not None)
        if drop_zero and k=="normal" and is_money and (cy or 0)==0 and (py or 0)==0:
            continue
        if k=="grandtotal" and rows and rows[-1].get("kind") not in ("blank","section"):
            rows.append({"kind":"blank"})
        row={"label":b,"note":(str(note) if note not in (None,"") else ""),"cy":cy or 0,"py":py or 0,"kind":k}
        if k=="normal": row["indent"]=True
        rows.append(row)
    # collapse consecutive blanks / trailing sections
    return rows

def _read_soce(ws, maxr=20):
    rows=[]
    for r in ws.iter_rows(min_row=4,max_row=maxr,max_col=8,values_only=True):
        b=r[1]
        if not b or not str(b).strip(): continue
        b=str(b).strip()
        sc=_num(r[2]); re=_num(r[5]); tot=_num(r[7])
        if sc is None and re is None and tot is None: continue
        k="total" if b.startswith("Balance at end") or b.startswith("Balance at 31") else "normal"
        rows.append({"label":b.replace("Balance at end of current year","Balance at 31 December (current year)"),
                     "sc":sc or 0,"re":re or 0,"tot":tot or 0,"kind":k})
    return rows

def _g(rows,label_contains,k="cy"):
    for r in rows:
        if label_contains.lower() in r.get("label","").lower(): return r.get(k)
    return None

def _find(rows, *contains, k="cy"):
    for r in rows:
        lab=r.get("label","").upper()
        if all(c.upper() in lab for c in contains): return r.get(k) or 0
    return 0

def tie_outs(soci,sofp,scf,soce):
    def exact(rows,lbl,k="cy"):
        for r in rows:
            if r.get("label","").upper()==lbl.upper(): return r.get(k) or 0
        return 0
    ta=exact(sofp,"TOTAL ASSETS"); tel=exact(sofp,"TOTAL EQUITY AND LIABILITIES")
    rev=exact(soci,"Revenue"); cos=exact(soci,"Cost of sales"); gp=exact(soci,"GROSS PROFIT")
    pat=exact(soci,"PROFIT/(LOSS) FOR THE YEAR")
    soce_pat=0; soce_re=0
    for r in soce:
        ll=r["label"].lower()
        if "for the year" in ll or "for the period" in ll: soce_pat=r["tot"]
        if r.get("kind")=="total" and ("31 december" in ll or "end of current" in ll or "balance at 31" in ll):
            soce_re=r["re"]
    sofp_re=_find(sofp,"RETAINED EARNINGS")
    scf_end=_find(scf,"CASH","END OF"); sofp_cash=_find(sofp,"CASH","EQUIVALENTS")
    return [
        ("SOFP balances", abs(ta-tel)<1),
        ("Gross profit", abs(rev+cos-gp)<1),
        ("Profit = equity movement", abs(pat-soce_pat)<1),
        ("SOCE RE = SOFP RE", abs(soce_re-sofp_re)<1),
        ("Cash flow = SOFP cash", abs(scf_end-sofp_cash)<1),
    ]

def get_data(xlsx_path, mode="draft", first_year=None, n_sig=2, template="SME",
             auditor="Kayode Okunola & Co (Chartered Accountants)", auditor_name="Kayode Okunola & Co",
             frc_no="0968263", ican_stamp_no="", stamp_image=None, signature_image=None):
    wb=openpyxl.load_workbook(xlsx_path, data_only=True)
    E=_sheet_map(wb["Entity"]); C=_sheet_map(wb["Cover"])
    name=str(E.get("Registered name") or "Company").strip()
    rc=str(E.get("RC / Business name no.") or "").strip()
    directors=[d.strip() for d in str(E.get("Directors / Partners / Proprietor") or "").split("\n") if d.strip()]
    office=[l.strip() for l in str(E.get("Registered office address") or "").replace("\n",", ").split(",") if l.strip()]
    office=[", ".join(office[:3]), ", ".join(office[3:])] if len(office)>3 else [", ".join(office)]
    activity=str(E.get("Principal activity") or "").strip() or "[Principal activity to be confirmed]"
    period_end=_fmtdate(C.get("Period ended (dd/mm/yyyy)"))
    sign_date=_fmtdate(C.get("Date of signing"))
    fy=period_end.split()[-1] if period_end else ""
    if first_year is None:
        first_year=str(C.get("First year of operations?") or "No").strip().lower().startswith("y")
    fw,fws=TEMPLATES.get(template,TEMPLATES["SME"])
    n_sig=max(1,min(n_sig,max(1,len(directors))))

    soci=_read_statement(wb["SOCI"]); sofp=_read_statement(wb["SOFP"])
    scf=_read_statement(wb["SCF"], drop_zero=True); soce=_read_soce(wb["SOCE"])

    def gv(rows,lbl,k="cy"):
        for r in rows:
            if r.get("label","").upper()==lbl.upper(): return r.get(k) or 0
        return 0
    rev_cy=gv(soci,"Revenue"); rev_py=gv(soci,"Revenue","py")
    pat_cy=gv(soci,"PROFIT/(LOSS) FOR THE YEAR"); pat_py=gv(soci,"PROFIT/(LOSS) FOR THE YEAR","py")
    dep=_g(scf,"Depreciation") or 0

    def naira(v): return "₦{:,.0f}".format(round(v))
    pl_word=lambda v:"loss" if v<0 else "profit"
    if first_year:
        results_para=(f"The operating results of the Company for the period are set out in the Statement of "
                      f"Profit or Loss and Other Comprehensive Income. The Company reported revenue of {naira(rev_cy)} "
                      f"and a {pl_word(pat_cy)} for the period of {naira(abs(pat_cy))}. As this is the Company's first "
                      "reporting period, no comparative figures are presented.")
    else:
        try: pct=abs((rev_cy-rev_py)/rev_py*100); dirn="increase" if rev_cy>=rev_py else "decrease"
        except ZeroDivisionError: pct=0; dirn="change"
        results_para=(f"The operating results of the Company for the year ended {period_end} are set out in the "
                      f"Statement of Profit or Loss and Other Comprehensive Income. The Company reported revenue of "
                      f"{naira(rev_cy)} (prior year: {naira(rev_py)}), representing a {dirn} of approximately "
                      f"{pct:.0f} per cent on the prior year, and a {pl_word(pat_cy)} for the year of {naira(abs(pat_cy))} "
                      f"(prior year: {pl_word(pat_py)} of {naira(abs(pat_py))}).")
    ppe_para=(f"Movements in property, plant and equipment during the year are set out in the notes. The "
              f"depreciation charge for the year amounted to {naira(dep)}.")

    flag=("The detailed line-item breakdown for this note was not coded in the trial balance provided; the "
          "total shown agrees to the face of the financial statements.")
    # auto figure-notes from the statements
    def srow(rows,lbl):
        for r in rows:
            if lbl.lower() in r.get("label","").lower(): return (r.get("cy") or 0, r.get("py") or 0)
        return (0,0)
    notes=[
        {"title":"1. General Information","paras":[f"{name.upper()} (the “Company”) is a limited liability company incorporated in Nigeria under the Companies and Allied Matters Act with Registration Number {rc}. The principal activity of the Company is {activity if not activity.startswith('[') else '[to be confirmed]'}. The Company is domiciled in {C.get('City / State') or 'Nigeria'}."]},
        {"title":"2. Basis of Preparation","paras":[f"The financial statements have been prepared in accordance with the {fw} and the applicable provisions of the Companies and Allied Matters Act, 2020. They are prepared under the historical-cost convention and presented in Nigerian Naira (₦), the functional and presentation currency of the Company."]},
        {"title":"3. Operating Environment","paras":[
            "The Nigerian business environment in "+(fy or "the year")+" continued to be characterised by macro-economic conditions including elevated inflation, foreign-exchange volatility following liberalisation of the exchange-rate regime, rising energy and input costs, and evolving fiscal and monetary policy. The Directors continue to monitor developments in the Company's industry and adapt the operating model accordingly.",
            "[App note: auto-tailored to the client's industry once principal activity is confirmed.]"]},
        {"title":"4. Significant Accounting Policies","paras":[
            "4.1 Revenue recognition — Revenue is recognised when control of goods or services transfers to the customer, measured at fair value of consideration received or receivable, net of discounts and VAT.",
            "4.2 Property, plant and equipment — Stated at cost less accumulated depreciation and impairment. Depreciation is on a straight-line basis over estimated useful lives.",
            "4.3 Trade receivables and payables — Initially at fair value, subsequently at amortised cost less impairment.",
            "4.4 Cash and cash equivalents — Cash in hand and deposits held at call with banks.",
            "4.5 Taxation — Current income tax, tertiary education tax and applicable levies under prevailing Nigerian tax legislation."]},
        {"title":"5. Financial Risk Management","paras":["The Company is exposed to financial, operational, regulatory and market risks. Management has established procedures to identify, evaluate, monitor and mitigate these risks, with overall oversight by the Board."]},
    ]
    rev=srow(soci,"Revenue"); cos=srow(soci,"Cost of sales"); adm=srow(soci,"Administrative expenses")
    fin=srow(soci,"Finance cost"); tax=srow(soci,"Taxation")
    notes.append({"title":"6. Revenue","table":[["Sales / turnover",rev[0],rev[1]],["Total revenue",rev[0],rev[1],"total"]]})
    if cos[0] or cos[1]:
        notes.append({"title":"7. Cost of Sales","paras":[flag],"table":[["Cost of sales",abs(cos[0]),abs(cos[1])],["Total cost of sales",abs(cos[0]),abs(cos[1]),"total"]]})
    notes.append({"title":"8. Administrative Expenses","paras":[flag],"table":[["Administrative expenses (per lead schedule)",abs(adm[0]),abs(adm[1])],["Total administrative expenses",abs(adm[0]),abs(adm[1]),"total"]]})
    # PPE movement from Note_10_PPE if present
    try:
        p=_sheet_map(wb["Note_10_PPE"],key_col=2,val_col=3,maxr=20)
        pc=_sheet_map(wb["Note_10_PPE"],key_col=2,val_col=4,maxr=20)
        def pv(k): 
            v=p.get(k); return v if isinstance(v,(int,float)) else 0
        def pvp(k):
            v=pc.get(k); return v if isinstance(v,(int,float)) else 0
        notes.append({"title":"10. Property, Plant and Equipment","paras":["Fixed-asset schedule for the year."],
            "table":[["Cost — balance brought forward",pv("Cost - opening"),pvp("Cost - opening")],
                     ["Additions during the year",pv("Additions during the year"),pvp("Additions during the year")],
                     ["Cost — balance carried forward",pv("Cost - closing"),pvp("Cost - closing"),"total"],
                     ["Depreciation — balance brought forward",pv("Accumulated depreciation - opening"),pvp("Accumulated depreciation - opening")],
                     ["Charge for the year",pv("Charge for the year"),pvp("Charge for the year")],
                     ["Depreciation — balance carried forward",pv("Accumulated depreciation - closing"),pvp("Accumulated depreciation - closing"),"total"],
                     ["Net book value",pv("Net book value"),pvp("Net book value"),"total"]]})
    except Exception:
        pass
    notes.append({"title":"19. Going Concern","paras":["The Directors have assessed the Company's ability to continue as a going concern and have a reasonable expectation that it has adequate resources to continue in operational existence for the foreseeable future. Accordingly, the going-concern basis has been adopted."]})

    entity={"name":name,"short_name":name.split()[0],"name_line2":" ".join(name.split()[1:]) or "Limited",
            "rc":rc,"activity":activity,"activity_short":activity if not activity.startswith("[") else "[Principal activity to be confirmed]",
            "activity_para":f"The principal activity of the Company during the year is {activity if not activity.startswith('[') else '[to be confirmed]'}. The Company continues to pursue its operations within its sector in Nigeria.",
            "directors":directors or ["Director"],"office":office,"bankers":str(E.get("Bankers") or "Banker details to be confirmed"),
            "auditor":auditor,"auditor_name":auditor_name,"city":str(C.get("City / State") or "Lagos, Nigeria")}
    meta={"mode":mode,"template":template,"entity_name":name,"short_name":entity["short_name"],"name_line2":entity["name_line2"],
          "activity_short":entity["activity_short"],"rc":rc,"auditor":auditor,"auditor_name":auditor_name,
          "fy":fy,"period_end":period_end,"sign_date":sign_date,"framework":fw,"framework_short":fws,
          "first_year":first_year,"signatories":(directors or ["Director"])[:n_sig],"sig_words":SIG_WORDS.get(n_sig,str(n_sig)),
          "results_para":results_para,"ppe_para":ppe_para,"frc_no":frc_no,"ican_stamp_no":ican_stamp_no,
          "stamp_image":stamp_image,"signature_image":signature_image,"total_pages":19}
    flags=[]
    if activity.startswith("["): flags.append("Principal activity not set in the workbook.")
    bk=str(E.get("Bankers") or "").strip().lower()
    if not bk or "to be confirmed" in bk: flags.append("Bankers not provided.")
    scale=str(C.get("Presentation scale") or "")
    if "000" in scale: flags.append("Presentation scale is labelled \u20a6'000 but figures appear to be full Naira \u2014 verify scale.")
    try:
        adm_detail=sum(v for v in _sheet_map(wb["Note_08_AdminExpenses"],3,3,60).values() if isinstance(v,(int,float)))
        if abs(adm_detail)<1 and abs(srow(soci,"Administrative expenses")[0])>0:
            flags.append("Note line-item detail (e.g. admin expenses) not coded in the trial balance; rolled-up totals shown.")
    except Exception: pass
    data={"meta":meta,"entity":entity,"soci":soci,"sofp":sofp,"scf":scf,"soce":soce,"notes":notes,"flags":flags}
    data["tie_outs"]=[{"name":n,"pass":bool(ok)} for n,ok in tie_outs(soci,sofp,scf,soce)]
    return data
