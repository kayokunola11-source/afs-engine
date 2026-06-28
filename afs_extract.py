# -*- coding: utf-8 -*-
"""Generic extractor: read a recalculated firm-template workbook into the data dict
that afs_generator.build expects. Works on any client using the standard template
(sheets: Entity, Cover, SOCI, SOFP, SCF, SOCE, Note_10_PPE, ...)."""
import datetime, re, openpyxl

TEMPLATES = {
    "SME":          ("International Financial Reporting Standard for Small and Medium-sized Entities (IFRS for SMEs)","IFRS for SMEs"),
    "NGO":          ("applicable financial reporting framework for Not-for-Profit organisations in Nigeria","the NGO reporting framework"),
    "Manufacturing":("International Financial Reporting Standard for Small and Medium-sized Entities (IFRS for SMEs)","IFRS for SMEs"),
    "PFA":          ("the Pension Reform Act and applicable PenCom reporting guidelines","the PenCom framework"),
    "Asset Management": ("International Financial Reporting Standards (IFRS)","IFRS"),
    "Fund Manager":     ("International Financial Reporting Standards (IFRS)","IFRS"),
}
SIG_WORDS={1:"one",2:"two",3:"three",4:"four"}
CURRENT_TEMPLATE_VERSION = 1   # bump when the master template changes
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
    # current-year movement = first "for the year/period"; current closing = first balance total AFTER it.
    # (Works whichever order the dialect lists the current vs prior blocks.)
    pat_idx=None
    for i,r in enumerate(soce):
        ll=r["label"].lower()
        if "for the year" in ll or "for the period" in ll: pat_idx=i; break
    soce_pat = soce[pat_idx]["tot"] if pat_idx is not None else 0
    soce_re=None
    start = (pat_idx+1) if pat_idx is not None else 0
    for r in soce[start:]:
        ll=r["label"].lower()
        if r.get("kind")=="total" and ll.startswith("balance"):
            soce_re=r["re"]; break
    soce_pat=soce_pat or 0; soce_re=soce_re or 0
    sofp_re=_find(sofp,"RETAINED EARNINGS")
    # closing cash: template uses "...end of year"; JUKES uses "...as at 31/12"; else last grandtotal row
    scf_end=_find(scf,"CASH","END OF") or _find(scf,"CASH","AS AT 31") or _find(scf,"CASH","31/12")
    if not scf_end:
        for r in reversed(scf):
            if r.get("kind")=="grandtotal": scf_end=r.get("cy") or 0; break
    sofp_cash=_find(sofp,"CASH","EQUIVALENTS")
    return [
        ("SOFP balances", abs(ta-tel)<1),
        ("Gross profit", abs(rev+cos-gp)<1),
        ("Profit = equity movement", abs(pat-soce_pat)<1),
        ("SOCE RE = SOFP RE", abs(soce_re-sofp_re)<1),
        ("Cash flow = SOFP cash", abs(scf_end-sofp_cash)<1),
    ]

def _build_fin_summary(wb):
    if "FinSum" not in wb.sheetnames: return None
    fs=wb["FinSum"]
    yrs=[fs.cell(4,c).value for c in range(3,8)]
    headers=["\u20a6'000"]+[ (str(int(y)) if isinstance(y,(int,float)) else (str(y) if y else "")) for y in yrs]
    rows=[]
    for r in range(6,20):
        lbl=fs.cell(r,2).value
        if not lbl or not str(lbl).strip(): continue
        rows.append([str(lbl).strip()]+[fs.cell(r,c).value for c in range(3,8)])
    return {"headers":headers,"rows":rows,"money_from":1} if rows else None

def _tax_schedules(wb):
    """Read the CIT tax computation and capital-allowances schedules (supplementary)."""
    out={}
    _JUNK=("enter ","pull ","prior-year","column","must =","tip","target","i13","h column",
           "per asset","(auto","see ","investigate","reconcil","from tb","cap at","capped at if")
    def _is_junk(lab):
        ll=lab.lower()
        return (len(lab)>46 or any(j in ll for j in _JUNK)) and not ll.startswith(("total","adjusted",
                "assessable","taxable","companies income","tertiary","development","police"))
    names={sn.lower():sn for sn in wb.sheetnames}
    cit=names.get("cit")
    if cit:
        ws=wb[cit]; rows=[]
        for r in range(6,46):
            lab=ws.cell(r,2).value
            if lab is None or not str(lab).strip(): continue
            raw=str(lab); lab=raw.strip(); low=lab.lower()
            if low.startswith("(capped") or _is_junk(lab): continue
            cyn=_num(ws.cell(r,3).value); pyn=_num(ws.cell(r,4).value)
            indent=raw[:1]==" "
            if cyn is None and pyn is None:
                rows.append({"label":lab,"cy":0,"py":0,"kind":"section"}); continue
            kind="normal"
            if low.startswith(("adjusted profit","assessable profit","taxable income")): kind="subtotal"
            if low.startswith("total tax payable"): kind="total"
            rows.append({"label":lab,"cy":cyn or 0,"py":pyn or 0,"kind":kind,"indent":indent})
            if kind=="total": break
        if any(r["kind"]!="section" and (abs(r["cy"])>0 or abs(r["py"])>0) for r in rows):
            out["cit"]={"title":str(ws.cell(1,1).value or "Companies Income Tax Computation").strip(),
                        "subtitle":(str(ws.cell(3,1).value).strip() if ws.cell(3,1).value else None),"rows":rows}
    ca=names.get("capallow") or names.get("cap all") or names.get("capital allowance")
    if ca:
        ws=wb[ca]; rows=[]
        for r in range(6,26):
            cls=ws.cell(r,2).value
            if cls is None or not str(cls).strip(): continue
            label=str(cls).strip()
            if _is_junk(label): continue
            vals=[ws.cell(r,c).value for c in range(3,9)]
            row=[label]+[(v if isinstance(v,str) else (_num(v) or 0)) for v in vals]
            if label.upper()=="TOTAL": row.append("total")
            rows.append(row)
        if rows:
            out["capallow"]={"title":str(ws.cell(1,1).value or "Capital Allowance Computation").strip(),
                             "headers":["Asset class","Rate","TWDV b/f","Additions","Initial","Annual","Total (CY)"],
                             "rows":rows}
    return out or None

def _renumber_continuous(notes, statement_lists):
    mp={}
    for i,nd in enumerate(notes,1):
        m=re.match(r"(\d+[a-z]?)\.\s*(.*)", nd.get("title",""))
        if not m: continue
        old=m.group(1); new=str(i); nd["title"]=f"{new}. {m.group(2)}"; mp[old]=new
    for rows in statement_lists:
        for r in rows:
            nt=str(r.get("note","")).strip()
            if nt in mp: r["note"]=mp[nt]

def _sub_placeholders(notes, name, rc):
    rc_clean = rc if (rc and "[" not in rc) else "[RC to be confirmed]"
    for nd in notes:
        out=[]
        for para in nd.get("paras",[]):
            para=re.sub(r"\[[^\]]*(?:entity|company|gaming)\s*name[^\]]*\]", name or "the Company", para, flags=re.I)
            para=re.sub(r"RC\s*\[[^\]]*\]", rc_clean, para)
            para=re.sub(r"\[[^\]]*to\s*(?:be\s*)?confirm[^\]]*\]", "[to be confirmed]", para, flags=re.I)
            out.append(para)
        if out: nd["paras"]=out

def _blank_dangling_refs(rows_lists, notes):
    have=set()
    for nd in notes:
        mm=re.match(r"(\d+[a-z]?)", nd.get("title",""))
        if mm: have.add(mm.group(1))
    for rows in rows_lists:
        for r in rows:
            nt=str(r.get("note","")).strip()
            if nt and nt not in have:
                r["note"]=""

def get_data(xlsx_path, mode="draft", first_year=None, n_sig=2, template="SME",
             auditor="[Audit Firm Name]", auditor_name="[Audit Firm Name]",
             frc_no="", ican_stamp_no="", stamp_image=None, signature_image=None, entity_overrides=None):
    import afs_jukes
    if afs_jukes.detect_dialect(xlsx_path)=="jukes":
        return afs_jukes.get_data_jukes(xlsx_path, mode=mode, first_year=first_year, n_sig=n_sig,
            template=template, auditor=auditor, auditor_name=auditor_name, frc_no=frc_no,
            ican_stamp_no=ican_stamp_no, stamp_image=stamp_image, signature_image=signature_image,
            entity_overrides=entity_overrides)
    wb=openpyxl.load_workbook(xlsx_path, data_only=True)
    E=_sheet_map(wb["Entity"]); C=_sheet_map(wb["Cover"], maxr=60)
    name=str(E.get("Registered name") or "Company").strip()
    rc=str(E.get("RC / Business name no.") or "").strip()
    directors=[d.strip() for d in str(E.get("Directors / Partners / Proprietor") or "").split("\n") if d.strip()]
    office=[l.strip() for l in str(E.get("Registered office address") or "").replace("\n",", ").split(",") if l.strip()]
    office=[", ".join(office[:3]), ", ".join(office[3:])] if len(office)>3 else [", ".join(office)]
    activity=str(E.get("Principal activity") or "").strip() or "[Principal activity to be confirmed]"
    period_end=_fmtdate(C.get("Period ended (dd/mm/yyyy)") or C.get("Period ended"))
    sign_date=_fmtdate(C.get("Date of signing"))
    # engagement details (from the app) take precedence over the workbook's own fields
    _eo=entity_overrides or {}
    if _eo.get("name"): name=_eo["name"]
    if _eo.get("rc"): rc=_eo["rc"]
    if _eo.get("period_end"): period_end=_eo["period_end"]
    if _eo.get("sign_date"): sign_date=_eo["sign_date"]
    if _eo.get("city"): city_override=_eo["city"]
    if _eo.get("directors"): directors=_eo["directors"]
    if _eo.get("activity"): activity=_eo["activity"]
    fy=period_end.split()[-1] if period_end else ""
    if first_year is None:
        first_year=str(C.get("First year of operations?") or "No").strip().lower().startswith("y")
    fw,fws=TEMPLATES.get(template,TEMPLATES["SME"])
    _mode=str(C.get("Mode") or C.get("Reporting mode") or "").lower()
    _full_ifrs = ("full ifrs" in _mode) or ("full ifrs" in str(template).lower())
    if _full_ifrs:                                  # full IFRS reporters (Cover mode OR app selection)
        fw,fws="International Financial Reporting Standards (IFRS)","IFRS"
    fin_summary=None
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

    # detailed notes: narrative (1-5) + figure notes read from the workbook (6..N), numbered
    def srow(rows,lbl):
        for r in rows:
            if lbl.lower() in r.get("label","").lower(): return (r.get("cy") or 0, r.get("py") or 0)
        return (0,0)
    narr=[
        {"title":"1. General Information","paras":[f"{name.upper()} (the \u201cCompany\u201d) is a limited liability company incorporated in Nigeria under the Companies and Allied Matters Act with Registration Number {rc}. The Company is domiciled in {C.get('City / State') or 'Nigeria'}."]},
        {"title":"2. Basis of Preparation","paras":[f"The financial statements have been prepared in accordance with the {fw} and the applicable provisions of the Companies and Allied Matters Act, 2020. They are prepared under the historical-cost convention and presented in Nigerian Naira (\u20a6), the functional and presentation currency of the Company."]},
        {"title":"3. Operating Environment","paras":["The Nigerian business environment in "+(fy or "the year")+" continued to be characterised by macro-economic conditions including elevated inflation, foreign-exchange volatility, rising input costs and evolving fiscal and monetary policy. The Directors continue to monitor developments in the Company's industry and adapt the operating model accordingly."]},
        {"title":"4. Significant Accounting Policies","paras":[
            "4.1 Statement of compliance and basis of preparation \u2014 The financial statements have been prepared in accordance with the "+fws+" as issued by the International Accounting Standards Board, on the historical-cost basis, and are presented in Nigerian Naira (\u20a6), the functional and presentation currency of the Company.",
            "4.2 Revenue recognition \u2014 Revenue is measured at the fair value of the consideration received or receivable, net of discounts, returns and value-added tax. Revenue is recognised when the amount can be reliably measured, it is probable that the economic benefits will flow to the Company, and control of the goods or services has been transferred to the customer.",
            "4.3 Property, plant and equipment \u2014 Stated at cost less accumulated depreciation and any accumulated impairment losses. Depreciation is recognised on a straight-line basis to write off the cost of each asset, less residual value, over its estimated useful life. The depreciation rates applied are set out in the property, plant and equipment note.",
            "4.4 Trade and other receivables \u2014 Recognised initially at fair value and subsequently measured at amortised cost using the effective interest method, less any provision for impairment (expected credit losses).",
            "4.5 Trade and other payables \u2014 Obligations for goods and services acquired in the ordinary course of business, recognised initially at fair value and subsequently measured at amortised cost.",
            "4.6 Cash and cash equivalents \u2014 Comprise cash in hand, deposits held at call with banks and other short-term highly liquid instruments with original maturities of three months or less, net of bank overdrafts where applicable.",
            "4.7 Provisions \u2014 Recognised when the Company has a present legal or constructive obligation as a result of a past event, it is probable that an outflow of resources will be required, and the amount can be reliably estimated.",
            "4.8 Taxation \u2014 The income tax charge comprises current income tax, tertiary education tax and applicable levies, computed on the basis of the tax laws enacted or substantively enacted at the reporting date in Nigeria.",
            "4.9 Events after the reporting date \u2014 New information on conditions existing at the reporting date is reflected in the financial statements. Events that do not affect the position at the reporting date but are material to users are disclosed."]},
        {"title":"5. Financial Risk Management","paras":[
            "The Company recognises that taking risk is inherent to its business activities and that effective risk management is fundamental to sustained performance. The Board of Directors retains overall responsibility for the establishment and oversight of the Company's risk-management framework, which is designed to identify, evaluate, monitor, manage and report the risks to which the Company is exposed.",
            "Risk-management objectives include: minimising surprises and protecting against unexpected losses; aligning business strategy with the risk appetite set by the Board; sustaining a strong, risk-aware culture across the workforce; and ensuring the prudent use of capital and resources.",
            "Financial risk \u2014 The Company is exposed principally to credit risk (on receivables and bank balances), liquidity risk (meeting obligations as they fall due) and market risk (including interest-rate and foreign-exchange risk). Management monitors these exposures on an ongoing basis and maintains adequate resources and controls to mitigate them."]},
    ]
    import afs_notes
    _fignotes,_ref=afs_notes.build_figure_notes(wb, 6)
    afs_notes.remap_statement_refs(soci,_ref); afs_notes.remap_statement_refs(sofp,_ref)
    afs_notes.fill_uncoded_totals(_fignotes, (soci or [])+(sofp or []))
    # warn (preparer-only) about notes coded for the prior year but not the current year
    _uncoded=[]
    for _nd in _fignotes:
        _t=_nd.get("table") or []
        _li=[r for r in _t if not (len(r)>3 and r[3]=="total")]
        if _li and all(abs(r[1])<1 for r in _li) and any(abs(r[2])>1 for r in _li):
            _uncoded.append(_nd["title"].split(". ",1)[-1])
    _gc=6+len(_fignotes)
    notes=narr+_fignotes+[{"title":f"{_gc}. Going Concern","paras":["The Directors have assessed the Company's ability to continue as a going concern and have a reasonable expectation that it has adequate resources to continue in operational existence for the foreseeable future. Accordingly, the going-concern basis has been adopted."]}]
    if _full_ifrs:
        import afs_ifrs
        _tb=afs_notes.build_tb_index(wb)
        notes,_iref=afs_ifrs.build_ifrs_notes(wb, _tb)
        afs_notes.remap_statement_refs(soci,_iref); afs_notes.remap_statement_refs(sofp,_iref)
        # --- omit immaterial nil notes; keep required narrative notes; renumber continuously ---
        _refd={str(r.get("note","")).strip() for r in (soci+sofp) if str(r.get("note","")).strip()}
        _ALWAYS={"1","2","3","4","23","24","25"}        # always-disclosed (policies, related parties, risk, events)
        _NILTXT={"23":"There were no related party transactions requiring disclosure during the year (prior year: nil)."}
        _kept=[]
        for _nd in notes:
            _mm=re.match(r"(\d+[a-z]?)\.", _nd["title"]); _num=_mm.group(1) if _mm else ""
            _empty=not (_nd.get("table") or _nd.get("grids") or _nd.get("ppe") or _nd.get("paras"))
            if _empty and _num not in _ALWAYS and _num not in _refd:
                continue                                # drop nil note that nothing references
            if _empty:
                _nd["paras"]=[_NILTXT.get(_num,"There were no balances under this heading at the reporting date (prior year: nil).")]
            _kept.append(_nd)
        _newmap={}; _i=0
        for _nd in _kept:
            _mm=re.match(r"(\d+[a-z]?)\.", _nd["title"]); _old=_mm.group(1) if _mm else ""
            _i+=1; _new=str(_i)
            _nd["title"]=re.sub(r"^\d+[a-z]?\.", _new+".", _nd["title"], count=1)
            if _old: _newmap[_old]=_new
        _newref=[(kws, str(_newmap.get(str(old), old))) for kws,old in _iref]
        afs_notes.remap_statement_refs(soci,_newref); afs_notes.remap_statement_refs(sofp,_newref)
        notes=_kept
        try:
            for _r in wb["SOCI"].iter_rows(min_row=4,max_row=30,max_col=5,values_only=True):
                _lab=str(_r[1]).strip() if _r[1] else ""
                if _lab.lower().startswith("basic earnings per share"):
                    soci.append({"label":_lab,"note":"","cy":_r[3],"py":_r[4],"kind":"normal"}); break
        except Exception: pass
        if not first_year:                              # 5-yr summary only when there is history
            fin_summary=_build_fin_summary(wb)
    # Asset-management workbooks carry their own full notes -> use them
    if ("Capital_Adequacy" in wb.sheetnames or "AUM_Schedule" in wb.sheetnames
            or "asset manager" in str(C.get("Entity type") or "").lower()):
        import afs_am
        _am=afs_am.build_am_notes(wb)
        if _am: notes=_am; _renumber_continuous(notes,[soci,sofp])

    _ph = (not activity) or activity.startswith("[") or any(w in activity.lower() for w in ("to be completed","to be confirmed","tbc"))
    _act_para = ("The principal activity of the Company during the year is to be confirmed by the Directors."
                 if _ph else f"The principal activity of the Company during the year is {activity}. "
                 "The Company continues to pursue its operations within its sector in Nigeria.")
    # bankers: app override (str newline-list or list) > workbook > placeholder
    _bk_raw = _eo.get("bankers") if _eo.get("bankers") else E.get("Bankers")
    if isinstance(_bk_raw,(list,tuple)):
        bankers=[str(b).strip() for b in _bk_raw if str(b).strip()]
    else:
        bankers=[b.strip() for b in str(_bk_raw or "").replace("\r","").split("\n") if b.strip()]
    if not bankers: bankers=["Banker details to be confirmed"]
    entity={"name":name,"short_name":name.split()[0],"name_line2":" ".join(name.split()[1:]) or "Limited",
            "rc":rc,"activity":activity,"activity_short":activity if not activity.startswith("[") else "[Principal activity to be confirmed]",
            "activity_para":_act_para,
            "directors":directors or ["Director"],"office":office,"bankers":bankers,
            "auditor":auditor,"auditor_name":auditor_name,"city":str(C.get("City / State") or "Lagos, Nigeria")}
    meta={"mode":mode,"template":template,"entity_name":name,"short_name":entity["short_name"],"name_line2":entity["name_line2"],
          "activity_short":entity["activity_short"],"rc":rc,"auditor":auditor,"auditor_name":auditor_name,
          "fy":fy,"period_end":period_end,"sign_date":sign_date,"framework":fw,"framework_short":fws,
          "first_year":first_year,"signatories":(directors or ["Director"])[:n_sig],"sig_words":SIG_WORDS.get(n_sig,str(n_sig)),
          "results_para":results_para,"ppe_para":ppe_para,"frc_no":frc_no,"ican_stamp_no":ican_stamp_no,
          "stamp_image":stamp_image,"signature_image":signature_image,"total_pages":19}
    flags=[]
    _tver=str(C.get("Template version") or C.get("Template Version") or "").strip()
    _m=re.search(r"\d+", _tver); _tv=int(_m.group()) if _m else 0
    meta["template_version"]=_tver or None
    if _tv==0:
        flags.append("This workbook has no template version on the Cover \u2014 it may be an outdated or non-standard template. Download the current template from the app.")
    # (stale-vs-current comparison is done app-side, against the published current version)
    if activity.startswith("["): flags.append("Principal activity not set in the workbook.")
    if any("to be confirmed" in b.lower() for b in bankers): flags.append("Bankers not provided.")
    scale=str(C.get("Presentation scale") or "")
    if "000" in scale: flags.append("Presentation scale is labelled \u20a6'000 but figures appear to be full Naira \u2014 verify scale.")
    try:
        adm_detail=sum(v for v in _sheet_map(wb["Note_08_AdminExpenses"],3,3,60).values() if isinstance(v,(int,float)))
        if abs(adm_detail)<1 and abs(srow(soci,"Administrative expenses")[0])>0:
            flags.append("Some note breakdowns are not coded at line-item level in the trial balance; the affected notes show totals only. Code the TB for full line-item disclosures. (This note is for the preparer and does not appear in the financial statements.)")
    except Exception: pass
    if _uncoded:
        flags.append("Current-year line-item detail is not coded in the trial balance for: "
                     + ", ".join(_uncoded[:10])
                     + ". These notes show the prior year only until the current-year trial balance is coded. (This note is for the preparer and does not appear in the financial statements.)")
    _sub_placeholders(notes, name, rc)
    _blank_dangling_refs([soci,sofp], notes)
    return {"entity":entity,"meta":meta,"soci":soci,"sofp":sofp,"scf":scf,"soce":soce,
            "notes":notes,"fin_summary":fin_summary,"tax_schedules":_tax_schedules(wb),
            "tie_outs":[{"name":n,"pass":bool(p)} for n,p in tie_outs(soci,sofp,scf,soce)],
            "flags":flags}
