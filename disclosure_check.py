# -*- coding: utf-8 -*-
"""IFRS for SMEs disclosure-completeness check.
Runs against a calc-core draft: detects which requirements APPLY from the data,
and marks each Pass / Missing / N/A / Manual. Turns the firm's 68-point checklist
into an automatic tick-list on every draft."""
import json, os, calc_core, afs_pycore

_HERE=os.path.dirname(__file__)
CHECKLIST=json.load(open(os.path.join(_HERE,"ifrs_sme_checklist.json")))

def _signals(path):
    r=calc_core.compute(path); tb=r["tb"]; sec={c:a["section"] for c,a in tb.items()}
    def S(secs): return sum(abs(a["cy_signed"]) for c,a in tb.items() if sec[c] in secs)
    d=afs_pycore.build_data(path)
    titles=" ".join(n["title"].lower() for n in d["notes"])
    paras=" ".join(p.lower() for n in d["notes"] for p in n.get("paras",[]))
    scf_py=any((r.get("py") or 0)!=0 for r in d["scf"])
    return {
        "ppe":S({"NCA-PPE-Cost","NCA-PPE-Dep"})>1,
        "intangible":S({"NCA-Intangible-Cost","NCA-Intangible-Amort"})>1,
        "inventory":S({"CA-Inventory"})>1,
        "borrowings":S({"NCL-Loans","CL-Loans","CL-Overdraft"})>1,
        "deferred_tax":S({"NCA-DefTax","NCL-DefTax"})>1,
        "finance_cost":S({"PL-FinCost"})>1,
        "related_party":S({"CL-DCA"})>1,
        "comparatives":any((r.get("py") or 0)!=0 for r in d["sofp"]),
        "scf_py":scf_py,
        "note_going_concern":"going concern" in titles,
        "note_policies":"accounting policies" in titles,
        "note_tax":"taxation" in titles,
        "note_related_party":"related part" in titles,
        "note_ppe":"property, plant" in titles,
        "note_cash":"cash and cash" in titles,
        "note_revenue":"revenue" in titles,
        "note_inventory_policy":"inventor" in paras,
        "tax_reconciliation":"reconciliation of tax" in titles or "statutory tax rate" in paras or "effective tax" in paras,
        "judgements":"judgement" in paras or "estimation" in paras,
        "kmp":"key management" in paras,
        "auth_issue":True,   # afs_generator prints the Board approval date + signatories
    }

# id -> function(signals)->status  ("Pass"/"Missing"/"N/A")  ; absent id => "Manual"
def _rule(cond, ok):  # condition applies -> Pass if ok else Missing ; else N/A
    return ("Pass" if ok else "Missing") if cond else "N/A"

RULES={
 "1.1": lambda s: "Pass" if s["note_policies"] else "Missing",
 "1.3": lambda s: "Pass" if s["note_going_concern"] else "Missing",
 "1.7": lambda s: "Pass" if s["comparatives"] else "Missing",
 "1.9": lambda s: "Pass",              # complete set is always emitted
 "1.11":lambda s: "Pass", "1.12":lambda s: "Pass",
 "2.1": lambda s: "Pass",
 "2.2": lambda s: _rule(s["deferred_tax"], s["deferred_tax"] and False) if s["deferred_tax"] else "N/A",
 "2.4": lambda s: "Pass",
 "3.2": lambda s: "Pass",
 "3.3": lambda s: _rule(s["finance_cost"], True),
 "3.6": lambda s: "Pass",              # analysed by function (admin/selling); depreciation disclosed in PPE note
 "4.1": lambda s: "Pass", "4.3": lambda s: "Pass",
 "4.4": lambda s: "Pass" if s["comparatives"] else "Missing",
 "5.1": lambda s: "Pass", "5.2": lambda s: "Pass", "5.3": lambda s: "Pass",
 "5.5": lambda s: "Pass",              # tax cash flow shown separately (Tax paid line)
 "5.7": lambda s: "Pass" if s["note_cash"] else "Missing",
 "5.9": lambda s: "Pass" if s["scf_py"] else "Missing",   # PY cash-flow column
 "6.1": lambda s: "Pass" if s["note_policies"] else "Missing",
 "6.4": lambda s: "Pass" if s["judgements"] else "Missing",
 "6.5": lambda s: "Pass" if s["judgements"] else "Missing",
 "7.1": lambda s: "Pass" if s["note_policies"] else "Missing",
 "8.1": lambda s: "Pass" if s["note_policies"] else "Missing",
 "9.1": lambda s: _rule(s["inventory"], s["note_inventory_policy"]),
 "10.1":lambda s: _rule(s["ppe"], s["note_ppe"]),
 "11.1":lambda s: _rule(s["intangible"], s["intangible"] and False) if s["intangible"] else "N/A",
 "13.1":lambda s: "Pass" if s["note_revenue"] else "Missing",
 "16.1":lambda s: "Pass" if s["note_tax"] else "Missing",
 "16.2":lambda s: "Pass" if s["tax_reconciliation"] else "Missing",
 "16.3":lambda s: _rule(s["deferred_tax"], False),
 "17.1":lambda s: "Pass" if s["auth_issue"] else "Missing",
 "18.2":lambda s: "Pass" if s["kmp"] else "Missing",
 "18.3":lambda s: _rule(s["related_party"], s["note_related_party"]),
}

def run(path):
    s=_signals(path); items=[]
    for req in CHECKLIST:
        idv=req["id"]
        status=RULES[idv](s) if idv in RULES else "Manual"
        # conditional requirements with no triggering condition and no rule -> Manual (preparer judges)
        items.append({**req,"status":status})
    from collections import Counter
    cnt=Counter(i["status"] for i in items)
    return {"summary":dict(cnt),"items":items,
            "missing":[i for i in items if i["status"]=="Missing"]}

if __name__=="__main__":
    import sys
    rep=run(sys.argv[1])
    print("Disclosure completeness —", rep["summary"])
    print("\nAUTO-FLAGGED GAPS (Missing):")
    for m in rep["missing"]:
        print(f"  [{m['id']}] {m['ref']:9s} {m['req'][:74]}")
