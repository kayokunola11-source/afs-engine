# -*- coding: utf-8 -*-
"""Full IFRS notes builder.

Used when the Cover sheet's Reporting mode = 'Full IFRS'. Follows the workbook's master
'Notes' sheet as the authoritative structure (notes 1-25): narrative + accounting policies
(3.1-3.9) and critical judgements are read VERBATIM from that sheet, figure notes pull their
breakdowns from the Note_XX sheets (by content), and the IFRS schedules - deferred tax (IAS 12),
revenue disaggregation (IFRS 15) and financial instruments / ECL / financial risk (IFRS 9 & 7) -
are rendered from their dedicated sheets.

Returns (notes, ref): notes is the ordered list the generator renders; ref maps keywords -> note
number so the statement note references can be aligned to this numbering.
"""
import re
import afs_notes


def _f(v):
    try:
        if v is None or v == "": return 0.0
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _sheet(wb, *name_contains):
    for sn in wb.sheetnames:
        low = sn.lower()
        if all(c.lower() in low for c in name_contains):
            return wb[sn]
    return None


# ---------------------------------------------------------------- master notes
def _clean_title(t):
    t=re.sub(r"\s*\((?:see|refer)[^)]*\)", "", t, flags=re.I)   # drop internal cross-refs
    t=re.sub(r"\s*\(Full IFRS only\)", "", t, flags=re.I)
    return t.strip()


def read_master_notes(wb):
    """Parse the 'Notes' sheet into {num: {'title':..., 'paras':[...]}}.
    Top-level headings look like '1. General information'; policy sub-items like '3.1 ...'
    are kept as paragraphs under note 3."""
    ws = _sheet(wb, "notes")
    out = {}
    cur = None
    if ws is None:
        return out
    for r in ws.iter_rows(min_row=4, max_row=120, max_col=2, values_only=True):
        a = (str(r[0]).strip() if r[0] is not None else "")     # headings live in col A
        bcol = (str(r[1]).strip() if r[1] is not None else "")  # narrative / policy text in col B
        m = re.match(r'^(\d+)\.\s+(.*)$', a) if a else None    # '1. General information'
        if m:
            num = int(m.group(1)); cur = num
            out[num] = {"title": _clean_title(f"{num}. {m.group(2).strip()}"), "paras": []}
            continue
        if cur is not None and bcol:
            out[cur]["paras"].append(bcol)              # policy sub-items + narrative
    return out


# ---------------------------------------------------------------- figure notes
def _fig_table(wb, tb, *name_contains):
    ws = _sheet(wb, *name_contains)
    if ws is None:
        return None
    tbl = afs_notes.read_fig_note(ws, tb)
    # keep only if it carries real figures
    if tbl and any(abs(_f(x[1])) >= 1 or abs(_f(x[2])) >= 1 for x in tbl):
        return tbl
    return None


def _movement_note(wb, tb, sheet_kw, close_kw):
    """Render a movement schedule (PPE, retained earnings) WITHOUT summing the lines.
    The closing/NBV line (matched by close_kw) is bolded; internal recon lines are skipped."""
    ws = _sheet(wb, sheet_kw)
    if ws is None:
        return None
    rows = []
    for r in ws.iter_rows(min_row=4, max_row=60, max_col=4, values_only=True):
        code = str(r[0]).strip() if r[0] is not None else ""
        bs = str(r[1]).strip() if r[1] is not None else ""
        if not bs:
            continue
        bl = bs.lower()
        if bl == "description" or code.lower() == "code":
            continue
        if bs.upper().startswith("RECONCILIATION"):
            break
        if (bl.startswith("total") or "from tb" in bl or "(auto" in bl or "h column" in bl
                or "per " in bl or "difference" in bl or "reconcil" in bl or "investigate" in bl
                or "target =" in bl or bl.startswith("tip") or "must = 0" in bl
                or bl.startswith("refer ") or "see ppe_schedule" in bl or "asset-class movement" in bl):
            continue
        cy = tb[code][0] if (tb and code in tb) else _f(r[2])
        py = tb[code][1] if (tb and code in tb) else _f(r[3])
        rows.append([bs, cy, py])
    if not rows:
        return None
    out = []
    for row in rows:
        out.append([row[0], row[1], row[2], "total"] if close_kw in row[0].lower() else row)
    return out


# ---------------------------------------------------------------- IFRS grids
def _grid(headers, rows, money_from=1, bold_last=False, subhead=None):
    return {"subhead": subhead, "headers": headers, "rows": rows,
            "money_from": money_from, "bold_last": bold_last}


def _deferred_tax_grids(wb):
    ws = _sheet(wb, "deferredtax")
    if ws is None:
        return []
    V = lambda r, c: ws.cell(r, c).value
    grids = []
    # (a) temporary differences
    rows = []
    for r in range(8, 16):
        lbl = V(r, 1)
        if not lbl: continue
        rows.append([str(lbl).strip(), _f(V(r, 2)), _f(V(r, 3))])
    if rows:
        rows[-1].append("total")
        grids.append(_grid(["Source of temporary difference", "2025", "2024"], rows,
                           bold_last=True, subhead="(a) Recognised deferred tax assets / (liabilities)"))
    # (b) reconciliation of accounting profit to tax expense
    rec = []
    for r in range(20, 35):
        lbl = V(r, 1)
        if not lbl: continue
        ls = str(lbl).strip()
        cy = V(r, 2)
        if ls.lower().startswith("effective tax rate"):
            rec.append([ls, f"{_f(cy)*100:.1f}%", ""]); continue
        rec.append([ls, _f(cy), _f(V(r, 3))])
    if rec:
        grids.append(_grid(["Reconciliation of accounting profit to tax", "2025", "2024"], rec,
                           subhead="(b) Reconciliation of accounting profit to tax expense"))
    return grids


def _revenue_disagg_grids(wb):
    ws = _sheet(wb, "revenue_disagg")
    if ws is None:
        return []
    V = lambda r, c: ws.cell(r, c).value
    grids = []
    rows = []
    for r in range(8, 17):
        lbl = V(r, 1)
        if not lbl: continue
        cy = _f(V(r, 2)); py = _f(V(r, 3))
        if abs(cy) < 1 and abs(py) < 1 and not str(lbl).lower().startswith("total"):
            continue
        rows.append([str(lbl).strip(), cy, py])
    if rows:
        if not str(rows[-1][0]).lower().startswith("total"):
            rows.append(["Total revenue from contracts with customers",
                         sum(x[1] for x in rows), sum(x[2] for x in rows)])
        rows[-1].append("total")
        grids.append(_grid(["Revenue stream", "2025", "2024"], rows, bold_last=True,
                           subhead="(a) Disaggregation of revenue by type of service"))
    # contract balances
    cb = []
    for r in range(29, 32):
        lbl = V(r, 1)
        if not lbl: continue
        cb.append([str(lbl).strip(), _f(V(r, 2)), _f(V(r, 3))])
    if cb:
        grids.append(_grid(["Contract balances at period end", "2025", "2024"], cb,
                           subhead="(c) Contract balances"))
    return grids


def _finrisk_grids(wb):
    grids = []
    # financial instruments by category (Note_26)
    fi = _sheet(wb, "fininst")
    if fi is not None:
        V = lambda r, c: fi.cell(r, c).value
        a = []
        for r in range(8, 15):
            lbl = V(r, 1)
            if not lbl: continue
            a.append([str(lbl).strip(), _f(V(r, 2)), _f(V(r, 3)), _f(V(r, 4)), _f(V(r, 5))])
        if a:
            a[-1].append("total")
            grids.append(_grid(["Financial assets", "Amortised cost", "FVOCI", "FVTPL", "Total"],
                               a, bold_last=True,
                               subhead="(a) Financial assets by IFRS 9 category"))
        liab = []
        for r in range(20, 26):
            lbl = V(r, 1)
            if not lbl: continue
            liab.append([str(lbl).strip(), _f(V(r, 2)), _f(V(r, 3)), _f(V(r, 4)), _f(V(r, 5))])
        if liab:
            liab[-1].append("total")
            grids.append(_grid(["Financial liabilities", "Amortised cost", "FVTPL", "Other", "Total"],
                               liab, bold_last=True, subhead="(b) Financial liabilities by category"))
        # ECL provision matrix
        ecl = []
        for r in range(32, 40):
            lbl = V(r, 1)
            if not lbl: continue
            ecl.append([str(lbl).strip(), _f(V(r, 2)), f"{_f(V(r,3))*100:.1f}%", _f(V(r, 4)), _f(V(r, 5))])
        if ecl:
            ecl[-1].append("total")
            grids.append(_grid(["Ageing bucket", "Gross", "ECL rate", "ECL allowance", "Net"],
                               ecl, bold_last=True,
                               subhead="(c) Expected credit loss - trade receivables (provision matrix)"))
    # credit risk / liquidity / capital (Note_28)
    fr = _sheet(wb, "finrisk")
    if fr is not None:
        V = lambda r, c: fr.cell(r, c).value
        cr = []
        for r in range(9, 15):
            lbl = V(r, 1)
            if not lbl: continue
            cr.append([str(lbl).strip(), _f(V(r, 2)), _f(V(r, 3))])
        if cr:
            cr[-1].append("total")
            grids.append(_grid(["Maximum exposure to credit risk", "2025", "2024"], cr,
                               bold_last=True, subhead="(d) Credit risk exposure"))
        liq = []
        for r in range(19, 26):
            lbl = V(r, 1)
            if not lbl: continue
            liq.append([str(lbl).strip(), _f(V(r, 2)), _f(V(r, 3)), _f(V(r, 4)), _f(V(r, 5)), _f(V(r, 6))])
        if liq:
            liq[-1].append("total")
            grids.append(_grid(["Liquidity - contractual maturities", "<3 mo", "3-12 mo", "1-5 yr", ">5 yr", "Total"],
                               liq, bold_last=True, subhead="(e) Liquidity risk"))
        cap = []
        for r in range(33, 38):
            lbl = V(r, 1)
            if not lbl: continue
            v = V(r, 2)
            if str(lbl).lower().startswith("net debt-to-equity"):
                cap.append([str(lbl).strip(), f"{_f(v):.2f}", ""]); continue
            cap.append([str(lbl).strip(), _f(v), ""])
        if cap:
            grids.append(_grid(["Capital management", "2025", ""], cap,
                               subhead="(f) Capital management"))
    return grids


# ---------------------------------------------------------------- assemble
# master note number -> Note_XX sheet keyword(s) for the figure breakdown
_FIG_MAP = {
    5: ("note_06_revenue",), 6: ("costofsales",), 7: ("otherincome",),
    8: ("adminexpenses",), 9: ("financecost",),
    14: ("tradereceivables",), 15: ("note_12_cash",), 16: ("sharecapital",),
    17: ("retainedearnings",), 18: ("longtermloan",), 19: ("tradepayables",),
    21: ("taxpayable",), 22: ("dircurrentaccount",),
}
# keyword -> note number, so statement note references align to this numbering
_REF = [
    (["revenue", "turnover"], 5), (["cost of sale"], 6), (["other income"], 7),
    (["administ"], 8), (["selling"], 8), (["finance cost"], 9),
    (["property", "plant"], 10), (["intangible"], "10a"), (["investment"], 11),
    (["deferred tax"], 12), (["inventory"], 13), (["receivable", "prepay"], 14),
    (["cash"], 15), (["share capital", "share premium"], 16),
    (["retained", "capital reserve"], 17), (["loan", "borrow"], 18),
    (["payable", "accrual"], 19), (["statutory"], 20), (["tax"], 21),
    (["director"], 22),
]


def _grid_has_data(g):
    """True if any money cell in the grid is non-trivial (>= 1 in magnitude)."""
    mf = g.get("money_from", 1)
    for row in g.get("rows", []):
        for c in row[mf:]:
            if isinstance(c, (int, float)) and abs(c) >= 1:
                return True
    return False


def _polish(notes):
    """Top-1% polish: drop all-zero schedules; give a short nil narrative to empty notes."""
    for nd in notes:
        if nd.get("grids"):
            kept = [g for g in nd["grids"] if _grid_has_data(g)]
            if kept:
                nd["grids"] = kept
            else:
                nd.pop("grids", None)
        # re-letter surviving sub-schedules sequentially: (a), (b), (c) ...
        i = 0
        for g in (nd.get("grids") or []):
            sh = g.get("subhead")
            if sh and re.match(r"^\([a-z]\)\s", sh):
                g["subhead"] = re.sub(r"^\([a-z]\)", f"({chr(97 + i)})", sh, count=1)
                i += 1
    return notes


def build_ifrs_notes(wb, tb=None):
    master = read_master_notes(wb)

    def note(num, default_title, paras=None, table=None, grids=None):
        m = master.get(num, {})
        return {"title": m.get("title", f"{num}. {default_title}"),
                "paras": paras if paras is not None else m.get("paras", []),
                **({"table": table} if table else {}),
                **({"grids": grids} if grids else {})}

    notes = []
    # 1-4 narrative + policies + judgements (verbatim from master)
    for n, dt in [(1, "General information"), (2, "Basis of preparation"),
                  (3, "Summary of significant accounting policies"),
                  (4, "Critical judgements and estimates")]:
        notes.append(note(n, dt))

    # 5 Revenue (+ IFRS 15 disaggregation)
    notes.append(note(5, "Revenue", paras=master.get(5, {}).get("paras", []),
                      table=_fig_table(wb, tb, "note_06_revenue"),
                      grids=_revenue_disagg_grids(wb)))
    # 6-9 P&L figure notes
    notes.append(note(6, "Cost of sales", table=_fig_table(wb, tb, "costofsales")))
    notes.append(note(7, "Other income", table=_fig_table(wb, tb, "otherincome")))
    # 8 admin + selling combined
    adm = _fig_table(wb, tb, "adminexpenses"); sell = _fig_table(wb, tb, "sellingexpenses")
    g8 = []
    if adm:  g8.append(_grid(["Administrative expenses", "2025", "2024"], _as_rows(adm), bold_last=True))
    if sell: g8.append(_grid(["Selling & distribution expenses", "2025", "2024"], _as_rows(sell), bold_last=True))
    notes.append(note(8, "Administrative, selling & distribution expenses", grids=g8 or None))
    notes.append(note(9, "Finance cost", table=_fig_table(wb, tb, "financecost")))

    # 10 PPE (summary + schedule)
    ppe_tbl = _movement_note(wb, tb, "note_10_ppe", "net book value")
    notes.append(note(10, "Property, plant and equipment", table=ppe_tbl))
    # 10a Intangible assets (IAS 38) - movement schedule, when present
    intang = _movement_note(wb, tb, "intangibles", "net book value")
    if intang:
        notes.append({"title": "10a. Intangible assets",
                      "paras": ["Internally developed software capitalised under IAS 38 and amortised on a straight-line basis over its estimated useful life of five years (20% per annum). Development costs meeting the recognition criteria are capitalised; research and maintenance costs are expensed as incurred."],
                      "table": intang})
    # 11-13
    notes.append(note(11, "Investments"))
    notes.append(note(12, "Deferred taxation", grids=_deferred_tax_grids(wb)))
    notes.append(note(13, "Inventory"))
    # 14-22 SOFP figure notes
    for n, dt, kw in [(14, "Trade and other receivables", "tradereceivables"),
                      (15, "Cash and cash equivalents", "note_12_cash"),
                      (16, "Share capital", "sharecapital"),
                      (17, "Retained earnings", "retainedearnings"),
                      (18, "Borrowings", "longtermloan"),
                      (19, "Trade and other payables", "tradepayables"),
                      (20, "Statutory deductions", None),
                      (21, "Taxation", "taxpayable"),
                      (22, "Director's current account", "dircurrentaccount")]:
        if n == 17:
            notes.append(note(n, dt, table=_movement_note(wb, tb, "retainedearnings", "closing balance")))
        else:
            notes.append(note(n, dt, table=(_fig_table(wb, tb, kw) if kw else None)))
    # 23 related parties (narrative)
    notes.append(note(23, "Related party transactions"))
    # 24 financial risk management + financial instruments + ECL (IFRS 7 & 9)
    notes.append(note(24, "Financial risk management", grids=_finrisk_grids(wb)))
    # 25 events after the reporting period (narrative)
    notes.append(note(25, "Events after the reporting period"))

    return _polish(notes), _REF


def _as_rows(tbl):
    """Convert a read_fig_note table ([lbl,cy,py,(kind)]) to grid rows, marking the total."""
    rows = []
    for x in tbl:
        r = [x[0], _f(x[1]), _f(x[2])]
        if len(x) > 3 and x[3] == "total":
            r.append("total")
        rows.append(r)
    return rows
