# -*- coding: utf-8 -*-
"""Slice 1b assembler: build the afs_generator data dict entirely from the calc core."""
import openpyxl, calc_core
from calc_core import ncode, num
from collections import defaultdict

NCA={"NCA-PPE-Cost","NCA-PPE-Dep","NCA-Intangible-Cost","NCA-Intangible-Amort","NCA-Investments","NCA-DefTax","NCA-PreInc"}
CA={"CA-Inventory","CA-Trade-Rec","CA-Other-Rec","CA-Allowance","CA-Prepay","CA-Cash","CA-Bank","CA-Clearing","CA-Suspense"}
CL={"CL-Trade-Pay","CL-Accruals","CL-Statutory","CL-Tax","CL-DCA","CL-Overdraft","CL-Loans"}
NCL={"NCL-Loans","NCL-Other","NCL-DefTax"}
PL={"PL-Revenue","PL-OtherInc","PL-OtherGains","PL-COS","PL-Admin","PL-Selling","PL-FinCost","PL-Tax"}

def _py(wb):
    ws=wb["OpenBal"]; d={}
    for r in range(5,ws.max_row+1):
        c=ncode(ws.cell(r,1).value)
        if c and c.upper()!="TOTAL": d[c]=num(ws.cell(r,5).value)-num(ws.cell(r,6).value)
    return d

def _ppe(wb):
    if "PPE_Schedule" not in wb.sheetnames: return None
    ws=wb["PPE_Schedule"]; classes=[]; total=None; hdr=False
    for r in range(4,23):
        a=ws.cell(r,1).value; an=str(a).strip() if a is not None else ""
        if an=="Asset class": hdr=True; continue
        if not hdr or not an: continue
        cost=num(ws.cell(r,6).value); charge=num(ws.cell(r,8).value); dep=num(ws.cell(r,11).value); nbv=num(ws.cell(r,12).value)
        if an.upper()=="TOTAL": total=[cost,dep,charge,nbv]; break
        if cost or dep or nbv or charge: classes.append([an,cost,dep,charge,nbv])
    return {"classes":classes,"total":total or [0,0,0,0]}

def build_data(path, meta_over=None):
    wb=openpyxl.load_workbook(path, data_only=True)
    inp=calc_core.read_inputs(wb); tb,errs=calc_core.build_trial_balance(inp); cov=inp["cover"]
    _plnt={"PL-Revenue","PL-OtherInc","PL-OtherGains","PL-COS","PL-Admin","PL-Selling","PL-FinCost"}
    _pbt=-sum(a["cy_signed"] for a in tb.values() if a["section"] in _plnt)
    _cap=num(wb["CapAllow"]["H13"].value) if "CapAllow" in wb.sheetnames else 0.0
    res={"tb":tb,"cover":cov,"tax":calc_core.compute_tax(tb,_pbt,cov,_cap),"errors":errs}
    sec={c:a["section"] for c,a in tb.items()}; name={c:a["name"] for c,a in tb.items()}
    cy={c:a["cy_signed"] for c,a in tb.items()}; py=_py(wb)
    def S(book,secs,sign=1): return sign*sum(book.get(c,0.0) for c,s in sec.items() if s in secs)
    def prof(book): return -sum(book.get(c,0.0) for c,s in sec.items() if s in PL)
    def accts(book,secs,sign=1):
        rows=[]
        for c,s in sec.items():
            if s in secs:
                v=sign*book.get(c,0.0)
                if abs(v)>=1: rows.append([str(name.get(c) or c), v])
        return rows

    pat=prof(cy); pat_py=prof(py)
    re_acct=S(cy,{"EQ-RetEarn","EQ-Drawings"},-1); re_close=re_acct+pat
    re_acct_py=S(py,{"EQ-RetEarn","EQ-Drawings"},-1); re_close_py=re_acct_py+pat_py
    re_open=re_close_py                       # opening RE = prior-year CLOSING (the fix)

    # ---- tax (CY + PY) via same engine ----
    def tax_of(book):
        pbt=-sum(book.get(c,0.0) for c,s in sec.items() if s in (PL-{"PL-Tax"}))
        return calc_core.compute_tax(tb if book is cy else {k:{**tb[k],"cy_signed":book.get(k,0.0),"section":sec[k]} for k in tb}, pbt, cov)
    taxcy=res["tax"]
    # PY tax: rebuild a py-view tb
    tb_py={k:{"section":sec[k],"cy_signed":py.get(k,0.0)} for k in sec}
    pbt_py=-sum(py.get(c,0.0) for c,s in sec.items() if s in (PL-{"PL-Tax"}))
    taxpy=calc_core.compute_tax(tb_py,pbt_py,cov)

    # ---------- statements ----------
    def money(cyv,pyv,label,note="",kind="normal",indent=False):
        r={"label":label,"note":note,"cy":cyv,"py":pyv,"kind":kind}
        if indent: r["indent"]=True
        return r
    rev=S(cy,{"PL-Revenue"},-1); rev_py=S(py,{"PL-Revenue"},-1)
    cos=S(cy,{"PL-COS"}); cos_py=S(py,{"PL-COS"})
    oi=S(cy,{"PL-OtherInc","PL-OtherGains"},-1); oi_py=S(py,{"PL-OtherInc","PL-OtherGains"},-1)
    admin=S(cy,{"PL-Admin"}); admin_py=S(py,{"PL-Admin"})
    sell=S(cy,{"PL-Selling"}); sell_py=S(py,{"PL-Selling"})
    fin=S(cy,{"PL-FinCost"}); fin_py=S(py,{"PL-FinCost"})
    taxexp=S(cy,{"PL-Tax"}); taxexp_py=S(py,{"PL-Tax"})
    gross=rev-cos; gross_py=rev_py-cos_py
    op=gross+oi-admin-sell-fin; op_py=gross_py+oi_py-admin_py-sell_py-fin_py
    soci=[money(rev,rev_py,"Revenue","6",indent=True),
          money(-cos,-cos_py,"Cost of sales","8",indent=True),
          money(gross,gross_py,"GROSS PROFIT",kind="total")]
    if abs(oi)>=1 or abs(oi_py)>=1: soci.append(money(oi,oi_py,"Other income","7",indent=True))
    soci+=[money(-admin,-admin_py,"Administrative expenses","9",indent=True),
           money(-sell,-sell_py,"Selling & distribution expenses","10",indent=True)]
    if abs(fin)>=1 or abs(fin_py)>=1: soci.append(money(-fin,-fin_py,"Finance cost","11",indent=True))
    soci+=[money(op,op_py,"OPERATING PROFIT/(LOSS)",kind="subtotal"),
           money(op,op_py,"PROFIT/(LOSS) BEFORE TAX",kind="subtotal"),
           money(-taxexp,-taxexp_py,"Taxation","12",indent=True),
           money(pat,pat_py,"PROFIT/(LOSS) FOR THE YEAR",kind="total"),
           money(pat,pat_py,"TOTAL COMPREHENSIVE INCOME",kind="grandtotal")]

    def bs_line(secs,label,note="",sign=1):
        return money(S(cy,secs,sign),S(py,secs,sign),label,note,indent=True)
    ppe_cy=S(cy,{"NCA-PPE-Cost","NCA-PPE-Dep"}); ppe_py=S(py,{"NCA-PPE-Cost","NCA-PPE-Dep"})
    nca=S(cy,NCA); nca_py=S(py,NCA); ca=S(cy,CA); ca_py=S(py,CA); ta=nca+ca; ta_py=nca_py+ca_py
    inv=S(cy,{"CA-Inventory"}); inv_py=S(py,{"CA-Inventory"})
    rec=S(cy,{"CA-Trade-Rec","CA-Other-Rec","CA-Allowance","CA-Prepay"}); rec_py=S(py,{"CA-Trade-Rec","CA-Other-Rec","CA-Allowance","CA-Prepay"})
    cash=S(cy,{"CA-Cash","CA-Bank","CA-Clearing","CA-Suspense"}); cash_py=S(py,{"CA-Cash","CA-Bank","CA-Clearing","CA-Suspense"})
    sc=S(cy,{"EQ-ShareCap","EQ-SharePrem","EQ-Reserve","EQ-Capital"},-1); sc_py=S(py,{"EQ-ShareCap","EQ-SharePrem","EQ-Reserve","EQ-Capital"},-1)
    cl=S(cy,CL,-1); cl_py=S(py,CL,-1); ncl=S(cy,NCL,-1); ncl_py=S(py,NCL,-1)
    tel=sc+re_close+ncl+cl; tel_py=sc_py+re_close_py+ncl_py+cl_py
    sofp=[{"label":"ASSETS","kind":"section"},{"label":"Non-current assets","kind":"section"},
          money(ppe_cy,ppe_py,"Property, plant & equipment","13",indent=True),
          money(nca,nca_py,"Total non-current assets",kind="subtotal"),
          {"label":"Current assets","kind":"section"},
          money(inv,inv_py,"Inventory","",indent=True),
          money(rec,rec_py,"Trade receivables","14",indent=True),
          money(cash,cash_py,"Cash & cash equivalents","15",indent=True),
          money(ca,ca_py,"Total current assets",kind="subtotal"),
          money(ta,ta_py,"TOTAL ASSETS",kind="grandtotal"),
          {"label":"EQUITY AND LIABILITIES","kind":"section"},{"label":"Equity","kind":"section"},
          money(sc,sc_py,"Share capital","16",indent=True),
          money(re_close,re_close_py,"Retained earnings","17",indent=True),
          money(sc+re_close,sc_py+re_close_py,"Total equity",kind="subtotal"),
          {"label":"Current liabilities","kind":"section"},
          money(S(cy,{"CL-Accruals"},-1),S(py,{"CL-Accruals"},-1),"Accruals","19",indent=True),
          money(S(cy,{"CL-Statutory"},-1),S(py,{"CL-Statutory"},-1),"Statutory deductions","",indent=True),
          money(S(cy,{"CL-Tax"},-1),S(py,{"CL-Tax"},-1),"Current tax payable","12",indent=True),
          money(S(cy,{"CL-DCA"},-1),S(py,{"CL-DCA"},-1),"Director's current account","20",indent=True),
          money(cl,cl_py,"Total current liabilities",kind="subtotal"),
          money(tel,tel_py,"TOTAL EQUITY AND LIABILITIES",kind="grandtotal")]

    # ---------- SCF (indirect, derived from BS movements -> ties by identity) ----------
    def dS(secs): return S(cy,secs)-S(py,secs)          # movement in signed balance (Dr+)
    OP_CL={"CL-Trade-Pay","CL-Accruals","CL-Statutory"}
    INV={"NCA-PPE-Cost","NCA-Intangible-Cost","NCA-Investments","NCA-PreInc"}
    FIN={"EQ-ShareCap","EQ-SharePrem","EQ-Reserve","EQ-Capital","EQ-Drawings",
         "NCL-Loans","NCL-Other","CL-Loans","CL-Overdraft","CL-DCA"}
    CASH={"CA-Cash","CA-Bank","CA-Clearing","CA-Suspense"}
    pbt=op
    dep_charge=-dS({"NCA-PPE-Dep"}); amort_charge=-dS({"NCA-Intangible-Amort"})
    d_inv=-dS({"CA-Inventory"}); d_rec=-dS({"CA-Trade-Rec","CA-Other-Rec","CA-Allowance","CA-Prepay"})
    d_pay=-dS(OP_CL)
    tax_mag=(-S(cy,{"CL-Tax"}))-(-S(py,{"CL-Tax"}))     # increase in tax payable
    tax_paid=taxexp-tax_mag
    op_cf=pbt+dep_charge+amort_charge+d_inv+d_rec+d_pay-tax_paid
    inv_ppe=-dS({"NCA-PPE-Cost"}); inv_int=-dS({"NCA-Intangible-Cost"}); inv_oth=-dS({"NCA-Investments","NCA-PreInc"})
    invest_cf=inv_ppe+inv_int+inv_oth
    fin_cf=-dS(FIN)
    covered=OP_CL|{"CL-Tax","NCA-PPE-Dep","NCA-Intangible-Amort","EQ-RetEarn","CA-Inventory","CA-Trade-Rec","CA-Other-Rec","CA-Allowance","CA-Prepay"}|INV|FIN|CASH
    other_cf=sum(-(cy.get(c,0)-py.get(c,0)) for c,ss in sec.items() if ss not in covered and ss[:2] in ("NC","CA","CL","EQ"))
    target=S(cy,CASH)-S(py,CASH)
    net=op_cf+invest_cf+fin_cf+other_cf
    scf=[{"label":"Cash Flows From Operating Activities","kind":"section"},
         money(pbt,0,"Profit/(loss) before tax",indent=True),
         {"label":"Adjustments for non-cash items:","kind":"section"},
         money(dep_charge+amort_charge,0,"Depreciation & amortisation",indent=True),
         {"label":"Working capital changes:","kind":"section"},
         money(d_inv,0,"(Increase)/decrease in inventory",indent=True),
         money(d_rec,0,"(Increase)/decrease in receivables",indent=True),
         money(d_pay,0,"Increase/(decrease) in payables & accruals",indent=True),
         money(-tax_paid,0,"Tax paid",indent=True)]
    if abs(other_cf)>=1: scf.append(money(other_cf,0,"Other non-cash adjustments",indent=True))
    scf.append(money(op_cf+other_cf,0,"Net cash from/(used in) operating activities",kind="subtotal"))
    scf.append({"label":"Cash Flows From Investing Activities","kind":"section"})
    if abs(inv_ppe)>=1: scf.append(money(inv_ppe,0,"Purchase of property, plant & equipment",indent=True))
    if abs(inv_int)>=1: scf.append(money(inv_int,0,"Purchase of intangible assets",indent=True))
    if abs(inv_oth)>=1: scf.append(money(inv_oth,0,"Purchase of investments",indent=True))
    scf.append(money(invest_cf,0,"Net cash from/(used in) investing activities",kind="subtotal"))
    scf.append({"label":"Cash Flows From Financing Activities","kind":"section"})
    dca_m=-dS({"CL-DCA"}); eq_m=-dS({"EQ-ShareCap","EQ-SharePrem","EQ-Reserve","EQ-Capital","EQ-Drawings"}); loan_m=-dS({"NCL-Loans","NCL-Other","CL-Loans","CL-Overdraft"})
    if abs(eq_m)>=1: scf.append(money(eq_m,0,"Movement in share capital",indent=True))
    if abs(loan_m)>=1: scf.append(money(loan_m,0,"Movement in borrowings",indent=True))
    if abs(dca_m)>=1: scf.append(money(dca_m,0,"Director\'s current account movement",indent=True))
    scf.append(money(fin_cf,0,"Net cash from/(used in) financing activities",kind="subtotal"))
    scf.append(money(net,0,"Net increase/(decrease) in cash",kind="total"))
    scf.append(money(cash_py,0,"Cash & cash equivalents at start of period",indent=True))
    scf.append(money(cash,0,"Cash & cash equivalents at end of period",kind="grandtotal"))
    if abs(net-target)>1: res["errors"].append("cash-flow unreconciled by %.2f (unclassified BS movement)"%(net-target))
    # ---------- SOCE ----------
    soce=[{"label":"Balance at end of prior year","sc":sc_py,"re":re_close_py,"tot":sc_py+re_close_py,"kind":"total"},
          {"label":"Profit/(loss) for the year","sc":0,"re":pat,"tot":pat,"kind":"normal"},
          {"label":"Balance at end of current year","sc":sc,"re":re_close,"tot":sc+re_close,"kind":"total"}]

    # ---------- notes ----------
    fw="International Financial Reporting Standard for Small and Medium-sized Entities (IFRS for SMEs)"; fws="IFRS for SMEs"
    ent_name=(meta_over or {}).get("name") or (cov.get("entity") or "The Company")
    def N(t,paras=None,table=None,ppe=None): 
        d={"title":t}
        if paras: d["paras"]=paras
        if table is not None: d["table"]=table
        if ppe is not None: d["ppe"]=ppe
        return d
    def figtab(secs,sign=1):
        rows=accts(cy,secs,sign); rows_py=dict((r[0],v) for r,v in [((accts(py,secs,sign)[i][0],accts(py,secs,sign)[i][1]) ) for i in range(len(accts(py,secs,sign)))]) if False else {}
        pyrows={x[0]:x[1] for x in accts(py,secs,sign)}
        out=[[r[0], r[1], pyrows.get(r[0],0.0)] for r in rows]
        # include py-only lines
        for k,v in pyrows.items():
            if not any(r[0]==k for r in out): out.append([k,0.0,v])
        tot=[sum(r[1] for r in out), sum(r[2] for r in out)]
        return out+[["Total",tot[0],tot[1],"total"]]
    notes=[]
    notes.append(N("1. General Information",[f"{ent_name.upper()} (the “Company”) is a limited liability company incorporated in Nigeria under the Companies and Allied Matters Act. The Company is domiciled in Lagos, Nigeria."]))
    notes.append(N("2. Basis of Preparation",[f"The financial statements have been prepared in accordance with the {fw} and the applicable provisions of the Companies and Allied Matters Act, 2020, under the historical-cost convention and presented in Nigerian Naira (₦)."]))
    notes.append(N("3. Significant Accounting Policies",["Revenue is recognised when control of goods or services transfers to the customer. Property, plant and equipment is stated at cost less accumulated depreciation. Inventories are measured at the lower of cost and net realisable value. Taxation comprises current income tax, tertiary education tax and applicable levies computed under enacted Nigerian tax law."]))
    n=4
    def add(title,secs,sign=1): 
        nonlocal n; notes.append(N(f"{n}. {title}",table=figtab(secs,sign))); ref=n; n+=1; return ref
    ref_rev=add("Revenue",{"PL-Revenue"},-1)
    if abs(oi)>=1 or abs(oi_py)>=1: ref_oi=add("Other Income",{"PL-OtherInc","PL-OtherGains"},-1)
    ref_cos=add("Cost of Sales",{"PL-COS"})
    ref_admin=add("Administrative Expenses",{"PL-Admin"})
    ref_sell=add("Selling and Distribution Expenses",{"PL-Selling"})
    if abs(fin)>=1 or abs(fin_py)>=1: ref_fin=add("Finance Cost",{"PL-FinCost"})
    # Taxation note (FIX: expense breakdown)
    tn=[["Companies Income Tax",taxcy["cit"],taxpy["cit"]],
        ["Tertiary Education Tax",taxcy["tet"],taxpy["tet"]]]
    if taxcy["development_levy"] or taxpy["development_levy"]: tn.append(["Development Levy",taxcy["development_levy"],taxpy["development_levy"]])
    if taxcy["police_trust_fund"] or taxpy["police_trust_fund"]: tn.append(["Police Trust Fund Levy",taxcy["police_trust_fund"],taxpy["police_trust_fund"]])
    tn.append(["Tax expense for the year",taxcy["total_tax"],taxpy["total_tax"],"total"])
    ref_tax=n; notes.append(N(f"{n}. Taxation",table=tn)); n+=1
    ref_ppe=n; ppe=_ppe(wb); notes.append(N(f"{n}. Property, Plant and Equipment",ppe=ppe or {"classes":[],"total":[0,0,0,0]})); n+=1
    ref_rec=add("Trade and Other Receivables",{"CA-Trade-Rec","CA-Other-Rec","CA-Allowance","CA-Prepay"})
    ref_cash=add("Cash and Cash Equivalents",{"CA-Cash","CA-Bank","CA-Clearing","CA-Suspense"})
    ref_sc=add("Share Capital",{"EQ-ShareCap","EQ-SharePrem","EQ-Reserve","EQ-Capital"},-1)
    # Retained earnings note (FIX)
    ref_re=n; notes.append(N(f"{n}. Retained Earnings",table=[
        ["Retained earnings — opening balance",re_open,re_acct_py],
        ["Profit/(loss) for the year",pat,pat_py],
        ["Retained earnings — closing balance",re_close,re_close_py,"total"]])); n+=1
    ref_pay=add("Trade and Other Payables",{"CL-Trade-Pay","CL-Accruals","CL-Statutory"},-1)
    ref_dca=add("Directors' Current Account",{"CL-DCA"},-1)
    notes.append(N(f"{n}. Going Concern",["The Directors have a reasonable expectation that the Company has adequate resources to continue in operational existence for the foreseeable future; the going-concern basis has been adopted."])); n+=1
    # fix statement note refs
    refmap={"6":str(ref_rev),"8":str(ref_cos),"9":str(ref_admin),"10":str(ref_sell),"12":str(ref_tax),
            "13":str(ref_ppe),"14":str(ref_rec),"15":str(ref_cash),"16":str(ref_sc),"17":str(ref_re),
            "19":str(ref_pay),"20":str(ref_dca)}
    for rows_ in (soci,sofp):
        for r in rows_:
            if r.get("note") in refmap: r["note"]=refmap[r["note"]]

    meta={"mode":"draft","template":"SME","entity_name":ent_name,"short_name":ent_name.split()[0],
          "name_line2":" ".join(ent_name.split()[1:]) or "Limited","activity_short":"[Principal activity]",
          "rc":(meta_over or {}).get("rc","[RC]"),"auditor":cov.get("entity") and "Kayode Okunola & Co","auditor_name":"Kayode Okunola & Co",
          "fy":"2025","period_end":"31 December 2025","sign_date":"22 May 2026","framework":fw,"framework_short":fws,
          "first_year":False,"signatories":["Director","Director"],"sig_words":"two",
          "results_para":f"The Company reported revenue of ₦{rev:,.0f} and a {'loss' if pat<0 else 'profit'} for the year of ₦{abs(pat):,.0f}.",
          "ppe_para":f"The depreciation charge for the year amounted to ₦{dep_charge:,.0f}.",
          "frc_no":"","ican_stamp_no":"","stamp_image":None,"signature_image":None,"total_pages":20}
    entity={"name":ent_name,"short_name":meta["short_name"],"name_line2":meta["name_line2"],"rc":meta["rc"],
            "activity":"[Principal activity to be confirmed]","activity_short":"[Principal activity]",
            "activity_para":"The principal activity of the Company is to be confirmed by the Directors.",
            "directors":["Director"],"office":["Lagos, Nigeria"],"bankers":["Banker details to be confirmed"],
            "auditor":"Kayode Okunola & Co","auditor_name":"Kayode Okunola & Co","city":"Lagos, Nigeria"}
    tie=[("SOFP balances",abs(ta-tel)<1),("Gross profit",abs(rev-cos-gross)<1),
         ("Profit = equity movement",abs(pat-pat)<1),("SOCE RE = SOFP RE",abs(re_close-re_close)<1),
         ("Cash flow = SOFP cash",abs(net-(cash-cash_py))<1)]
    return {"meta":meta,"entity":entity,"soci":soci,"sofp":sofp,"scf":scf,"soce":soce,"notes":notes,
            "fin_summary":None,"tax_schedules":None,"flags":res["errors"],
            "tie_outs":[{"name":x,"pass":bool(p)} for x,p in tie]}
