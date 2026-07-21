#!/usr/bin/env python3
"""AFS Generator - Nigerian SME Audited Financial Statements.
House style: Kayode Okunola & Co (navy/gold), TRIPZAPP layout.
mode="draft" -> diagonal DRAFT watermark, no stamp.
mode="final" -> no watermark, FRC + ICAN stamp number + scanned stamp image.
This single file is the generation engine the Lovable app runs server-side."""
import sys
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph,
                                Spacer, Table, TableStyle, NextPageTemplate,
                                PageBreak, KeepTogether, Image)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY

DEJ = "/usr/share/fonts/truetype/dejavu/"
pdfmetrics.registerFont(TTFont("DV",  DEJ+"DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DVB", DEJ+"DejaVuSans-Bold.ttf"))
pdfmetrics.registerFont(TTFont("DVI", DEJ+"DejaVuSans-Oblique.ttf"))

NAVY = colors.HexColor("#13315C"); NAVY2= colors.HexColor("#1F3864")
GOLD = colors.HexColor("#C49A2C"); GREY = colors.HexColor("#6B6B6B")
LGREY= colors.HexColor("#9AA0A6"); RULE = colors.HexColor("#1F3864")
PAGE_W, PAGE_H = A4
LM = RM = 22*mm; TM = 26*mm; BM = 20*mm


from reportlab.lib.utils import ImageReader
UNIT="N"

def scaled_image(path, max_w, max_h):
    """Return an Image scaled to fit within max_w x max_h, preserving aspect ratio."""
    try:
        iw, ih = ImageReader(path).getSize()
        ratio = min(max_w/iw, max_h/ih)
        return Image(path, width=iw*ratio, height=ih*ratio)
    except Exception:
        return Image(path, width=max_w, height=max_h)

def money(v):
    if v in (None,"","-"): return "-"
    try: v=float(v)
    except (TypeError,ValueError): return str(v)
    if abs(v)<0.5: return "-"
    if v<0: return "("+format(abs(round(v)),",.0f")+")"
    return format(round(v),",.0f")

def Sx(name,**kw):
    base=dict(fontName="DV",fontSize=9.5,leading=14,textColor=colors.black); base.update(kw)
    return ParagraphStyle(name,**base)
HY="en_GB"
body=Sx("body",alignment=TA_JUSTIFY,leading=15.5,spaceAfter=7,hyphenationLang=HY,
        embeddedHyphenation=1,uriWasteReduce=0.3); bodyc=Sx("bodyc",leading=15.5)
h1=Sx("h1",fontName="DVB",fontSize=15,textColor=NAVY,leading=19,spaceAfter=2)
h2=Sx("h2",fontName="DVB",fontSize=10.5,textColor=NAVY2,spaceBefore=12,spaceAfter=4,keepWithNext=1)
sub=Sx("sub",fontName="DVI",fontSize=9.5,textColor=GREY,spaceAfter=8,keepWithNext=1)
bullet=Sx("bullet",alignment=TA_JUSTIFY,leftIndent=14,bulletIndent=3,leading=15.5,
          spaceAfter=4,hyphenationLang=HY,embeddedHyphenation=1)
sig=Sx("sig",fontName="DVB",fontSize=9.5,textColor=NAVY2)
sigsub=Sx("sigsub",fontName="DV",fontSize=9,textColor=colors.black)
sigdate=Sx("sigdate",fontName="DVI",fontSize=9,textColor=GREY)
small_i=Sx("small_i",fontName="DVI",fontSize=8.5,textColor=GREY,leading=13,spaceAfter=4)

def draw_watermark(c):
    c.saveState(); c.translate(PAGE_W/2,PAGE_H/2); c.rotate(45)
    c.setFont("DVB",150); c.setFillColor(colors.Color(0.86,0.86,0.86,alpha=0.55))
    c.drawCentredString(0,-45,"DRAFT"); c.restoreState()

def header_footer(c,doc):
    st=doc.afs; c.saveState()
    if st["mode"]=="draft": draw_watermark(c)
    c.setFont("DVB",8.5); c.setFillColor(NAVY)
    c.drawString(LM,PAGE_H-15*mm,st["entity_name"].upper())
    c.setFont("DVI",8.5); c.setFillColor(GREY)
    c.drawRightString(PAGE_W-RM,PAGE_H-15*mm,f"Audited Financial Statements {st['fy']}")
    c.setStrokeColor(GOLD); c.setLineWidth(1.1)
    c.line(LM,PAGE_H-17*mm,PAGE_W-RM,PAGE_H-17*mm)
    c.setStrokeColor(LGREY); c.setLineWidth(0.4); c.line(LM,BM+4*mm,PAGE_W-RM,BM+4*mm)
    c.setFont("DVI",7); c.setFillColor(GREY)
    c.drawString(LM,BM, st.get('auditor') or '')   # footer always shows the firm name (never free-text/address)
    c.drawCentredString(PAGE_W/2,BM,f"Page {doc.page} of {st['total_pages']}")
    c.drawRightString(PAGE_W-RM,BM,"Strictly Confidential"); c.restoreState()

def cover_page(c,doc):
    st=doc.afs
    if st["mode"]=="draft": draw_watermark(c)
    if st.get("logo_image"):
        try:
            ir=ImageReader(st["logo_image"]); iw,ih=ir.getSize(); h=18*mm; w=iw*(h/ih)
            c.drawImage(ir, LM, PAGE_H-32*mm, width=w, height=h, mask='auto', preserveAspectRatio=True)
        except Exception: pass
    c.setStrokeColor(GOLD); c.setLineWidth(2.4); c.line(LM,PAGE_H-40*mm,PAGE_W-RM,PAGE_H-40*mm)
    # Full entity name as the cover title, word-wrapped to fit (never just the first word).
    _name=(st.get("entity_name") or (st.get("short_name","")+" "+st.get("name_line2",""))).upper().strip()
    _maxw=PAGE_W-LM-RM
    def _wrap(txt,sz):
        out=[]; cur=""
        for w in txt.split():
            t=(cur+" "+w).strip()
            if c.stringWidth(t,"DVB",sz)<=_maxw: cur=t
            else:
                if cur: out.append(cur)
                cur=w
        if cur: out.append(cur)
        return out
    _sz=32; _lines=_wrap(_name,_sz)
    while len(_lines)>2 and _sz>18:
        _sz-=2; _lines=_wrap(_name,_sz)
    c.setFillColor(NAVY); c.setFont("DVB",_sz)
    y=PAGE_H-56*mm
    for _ln in _lines:
        c.drawString(LM,y,_ln); y-=(_sz*0.42)*mm
    y-=3*mm; c.setFont("DV",9.5); c.setFillColor(GREY); c.drawString(LM,y,st["rc"])
    y-=30*mm; c.setFillColor(NAVY); c.setFont("DVB",22); c.drawString(LM,y,"AUDITED")
    y-=11*mm; c.drawString(LM,y,"FINANCIAL STATEMENTS")
    y-=9*mm; c.setFont("DV",11); c.setFillColor(GREY); c.drawString(LM,y,f"For the year ended {st['period_end']}")
    yb=52*mm; c.setStrokeColor(GOLD); c.setLineWidth(1.6); c.line(LM,yb,PAGE_W-RM,yb)
    c.setFont("DVI",9); c.setFillColor(GREY); c.drawString(LM,yb-7*mm,"Audited by")
    c.setFont("DVB",11); c.setFillColor(NAVY); c.drawString(LM,yb-13*mm,st["auditor_name"])
    c.setFont("DVI",9); c.setFillColor(GREY); c.drawString(LM,yb-18*mm,"Chartered Accountants")

COLW=[PAGE_W-LM-RM-32*mm-32*mm-14*mm,14*mm,32*mm,32*mm]
def stmt_table(rows,first_year=False,cy_year="2025",py_year="2024"):
    show_py=not first_year; data=[]; styles=[]
    data.append(["","Note",str(cy_year),(str(py_year) if show_py else "")]); r0=0
    styles+=[("FONTNAME",(0,r0),(-1,r0),"DVB"),("FONTSIZE",(0,r0),(-1,r0),8.5),("TEXTCOLOR",(0,r0),(-1,r0),NAVY2),
             ("ALIGN",(1,r0),(-1,r0),"RIGHT"),("LINEBELOW",(0,r0),(-1,r0),0.7,NAVY2),("BOTTOMPADDING",(0,r0),(-1,r0),3)]
    data.append(["","",UNIT,UNIT if show_py else ""]); ru=1
    styles+=[("FONTNAME",(0,ru),(-1,ru),"DVB"),("FONTSIZE",(0,ru),(-1,ru),9),("ALIGN",(2,ru),(-1,ru),"RIGHT"),
             ("TOPPADDING",(0,ru),(-1,ru),0),("BOTTOMPADDING",(0,ru),(-1,ru),2)]
    for row in rows:
        k=row.get("kind","normal")
        if k=="blank":
            data.append(["","","",""]); ri=len(data)-1
            styles+=[("TOPPADDING",(0,ri),(-1,ri),2),("BOTTOMPADDING",(0,ri),(-1,ri),2)]; continue
        if k=="section":
            data.append([row["label"],"","",""]); ri=len(data)-1
            styles+=[("FONTNAME",(0,ri),(-1,ri),"DVB"),("TEXTCOLOR",(0,ri),(0,ri),NAVY2),("FONTSIZE",(0,ri),(-1,ri),9.5)]; continue
        cy=money(row.get("cy")); py=money(row.get("py")) if show_py else ""
        data.append([row["label"],str(row.get("note","")),cy,py]); ri=len(data)-1
        styles+=[("ALIGN",(1,ri),(-1,ri),"RIGHT"),("TOPPADDING",(0,ri),(-1,ri),2.2),("BOTTOMPADDING",(0,ri),(-1,ri),2.2)]
        if row.get("indent"): styles.append(("LEFTPADDING",(0,ri),(0,ri),16))
        if k in ("total","grandtotal","subtotal"):
            styles+=[("FONTNAME",(0,ri),(-1,ri),"DVB"),("LINEABOVE",(0,ri),(-1,ri),0.6,colors.black)]
        if k=="total": styles.append(("LINEBELOW",(2,ri),(-1,ri),0.6,colors.black))
        if k=="grandtotal":
            styles+=[("LINEBELOW",(2,ri),(-1,ri),1.1,NAVY),("TEXTCOLOR",(0,ri),(-1,ri),NAVY)]
    t=Table(data,colWidths=[COLW[0],COLW[1],COLW[2],COLW[3] if show_py else 0])
    t.setStyle(TableStyle([("FONTNAME",(0,0),(-1,-1),"DV"),("FONTSIZE",(0,0),(-1,-1),9.3),
                           ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0)]+styles))
    return t

def note_table(pairs,first_year=False,cy_year="",py_year=""):
    show_py=not first_year
    data=[["",str(cy_year),(str(py_year) if show_py else "")],["",UNIT,UNIT if show_py else ""]]
    styles=[("FONTNAME",(0,0),(-1,1),"DVB"),("ALIGN",(1,0),(-1,1),"RIGHT"),("FONTSIZE",(0,0),(-1,0),8.5),
            ("TEXTCOLOR",(0,0),(-1,0),NAVY2),("FONTSIZE",(0,1),(-1,1),9),
            ("LINEBELOW",(0,1),(-1,1),0.5,LGREY),("BOTTOMPADDING",(0,0),(-1,0),1),("TOPPADDING",(0,1),(-1,1),0)]
    for lbl,cy,py,*rest in pairs:
        tot=rest and rest[0]=="total"
        data.append([lbl,money(cy),money(py) if show_py else ""]); ri=len(data)-1
        styles+=[("ALIGN",(1,ri),(-1,ri),"RIGHT"),("TOPPADDING",(0,ri),(-1,ri),2),("BOTTOMPADDING",(0,ri),(-1,ri),2)]
        if tot: styles+=[("FONTNAME",(0,ri),(-1,ri),"DVB"),("LINEABOVE",(1,ri),(-1,ri),0.6,colors.black),
                         ("LINEBELOW",(1,ri),(-1,ri),1.0,NAVY),("TEXTCOLOR",(0,ri),(-1,ri),NAVY)]
    w0=PAGE_W-LM-RM-34*mm-34*mm
    t=Table(data,colWidths=[w0,34*mm,34*mm if show_py else 0])
    t.setStyle(TableStyle([("FONTNAME",(0,0),(-1,-1),"DV"),("FONTSIZE",(0,0),(-1,-1),9.3),
                           ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0)]+styles))
    return t

def cit_table(rows, cy_year="", py_year=""):
    """Tax-computation style schedule: label | CY | PY, with sections, subtotals and a final total."""
    data=[["",str(cy_year),str(py_year)],["",UNIT,UNIT]]
    st=[("FONTNAME",(0,0),(-1,1),"DVB"),("ALIGN",(1,0),(-1,1),"RIGHT"),("FONTSIZE",(0,0),(-1,0),8.5),
        ("TEXTCOLOR",(0,0),(-1,0),NAVY2),("FONTSIZE",(0,1),(-1,1),9),("LINEBELOW",(0,1),(-1,1),0.5,LGREY),
        ("BOTTOMPADDING",(0,0),(-1,0),1),("TOPPADDING",(0,1),(-1,1),0)]
    for r in rows:
        k=r.get("kind","normal")
        if k=="section":
            data.append([r["label"],"",""]); ri=len(data)-1
            st+=[("FONTNAME",(0,ri),(-1,ri),"DVB"),("TEXTCOLOR",(0,ri),(0,ri),NAVY2),
                 ("TOPPADDING",(0,ri),(-1,ri),4),("BOTTOMPADDING",(0,ri),(-1,ri),1)]; continue
        data.append([r["label"],money(r["cy"]),money(r["py"])]); ri=len(data)-1
        st+=[("ALIGN",(1,ri),(-1,ri),"RIGHT"),("TOPPADDING",(0,ri),(-1,ri),2),("BOTTOMPADDING",(0,ri),(-1,ri),2)]
        if r.get("indent"): st.append(("LEFTPADDING",(0,ri),(0,ri),16))
        if k in ("subtotal","total"):
            st+=[("FONTNAME",(0,ri),(-1,ri),"DVB"),("LINEABOVE",(1,ri),(-1,ri),0.6,colors.black)]
        if k=="total":
            st+=[("LINEBELOW",(1,ri),(-1,ri),1.1,NAVY),("TEXTCOLOR",(0,ri),(-1,ri),NAVY)]
    w0=PAGE_W-LM-RM-34*mm-34*mm
    t=Table(data,colWidths=[w0,34*mm,34*mm])
    t.setStyle(TableStyle([("FONTNAME",(0,0),(-1,-1),"DV"),("FONTSIZE",(0,0),(-1,-1),9.3),
                           ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0)]+st))
    return t

def grid_table(g):
    """Generic multi-column schedule (IFRS notes). g = {headers, rows, money_from, bold_last}.
    Numeric cells in columns >= money_from are right-aligned and money-formatted; a row whose
    last element is the string 'total' is bolded with a rule."""
    headers=g["headers"]; rows=g["rows"]; mf=g.get("money_from",1); bl=g.get("bold_last",False)
    ncol=len(headers)
    data=[list(headers)]
    st=[("FONTNAME",(0,0),(-1,0),"DVB"),("FONTSIZE",(0,0),(-1,0),7.4),("TEXTCOLOR",(0,0),(-1,0),NAVY2),
        ("ALIGN",(mf,0),(-1,0),"RIGHT"),("LINEBELOW",(0,0),(-1,0),0.6,NAVY2),("BOTTOMPADDING",(0,0),(-1,0),3)]
    n=len(rows)
    for idx,row in enumerate(rows):
        cells=list(row)
        kind=cells[-1] if (len(cells)>ncol and isinstance(cells[-1],str)) else None
        if kind is not None: cells=cells[:ncol]
        out=[]
        for i,c in enumerate(cells):
            if i>=mf and isinstance(c,(int,float)): out.append(money(c))
            else: out.append(str(c) if c is not None else "")
        while len(out)<ncol: out.append("")
        data.append(out); ri=len(data)-1
        st+=[("ALIGN",(mf,ri),(-1,ri),"RIGHT"),("TOPPADDING",(0,ri),(-1,ri),1.8),("BOTTOMPADDING",(0,ri),(-1,ri),1.8)]
        if kind=="total" or (bl and idx==n-1):
            st+=[("FONTNAME",(0,ri),(-1,ri),"DVB"),("TEXTCOLOR",(0,ri),(-1,ri),NAVY),
                 ("LINEABOVE",(mf,ri),(-1,ri),0.5,colors.black),("LINEBELOW",(mf,ri),(-1,ri),0.8,NAVY)]
    fixed=(22*mm if ncol>3 else 32*mm)
    w0=PAGE_W-LM-RM-fixed*(ncol-1)
    if w0<38*mm: w0=38*mm
    t=Table(data,colWidths=[w0]+[fixed]*(ncol-1))
    t.setStyle(TableStyle([("FONTNAME",(0,0),(-1,-1),"DV"),("FONTSIZE",(0,0),(-1,-1),7.7),
                           ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0)]+st))
    return t

def soce_table(rows,first_year=False,headers=None):
    data=[headers or ["","Share capital","Retained earnings","Total equity"]]
    styles=[("FONTNAME",(0,0),(-1,0),"DVB"),("ALIGN",(1,0),(-1,0),"RIGHT"),("TEXTCOLOR",(0,0),(-1,0),NAVY2),
            ("FONTSIZE",(0,0),(-1,0),8.6),("LINEBELOW",(0,0),(-1,0),0.7,NAVY2),("BOTTOMPADDING",(0,0),(-1,0),3)]
    data.append(["",UNIT,UNIT,UNIT])
    styles+=[("FONTNAME",(0,1),(-1,1),"DVB"),("ALIGN",(1,1),(-1,1),"RIGHT"),("BOTTOMPADDING",(0,1),(-1,1),2)]
    for r in rows:
        bold=r.get("kind")=="total"
        data.append([r["label"],money(r["sc"]),money(r["re"]),money(r["tot"])]); ri=len(data)-1
        styles+=[("ALIGN",(1,ri),(-1,ri),"RIGHT"),("TOPPADDING",(0,ri),(-1,ri),2.2),("BOTTOMPADDING",(0,ri),(-1,ri),2.2)]
        if bold: styles+=[("FONTNAME",(0,ri),(-1,ri),"DVB"),("LINEABOVE",(0,ri),(-1,ri),0.6,colors.black),
                          ("TEXTCOLOR",(0,ri),(-1,ri),NAVY)]
    w=(PAGE_W-LM-RM-3*30*mm)
    t=Table(data,colWidths=[w,30*mm,30*mm,30*mm])
    t.setStyle(TableStyle([("FONTNAME",(0,0),(-1,-1),"DV"),("FONTSIZE",(0,0),(-1,-1),9),
                           ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0)]+styles))
    return t


def ppe_table(ppe, first_year=False):
    """Fixed-asset schedule by asset class: Class | Cost | Rate | Charge | Acc. deprec'n | NBV."""
    head=["Asset class","Cost","Rate","Charge","Acc. dep'n","NBV"]
    data=[head]
    st=[("FONTNAME",(0,0),(-1,0),"DVB"),("FONTSIZE",(0,0),(-1,0),7.6),("TEXTCOLOR",(0,0),(-1,0),NAVY2),
        ("ALIGN",(1,0),(-1,0),"RIGHT"),("LINEBELOW",(0,0),(-1,0),0.7,NAVY2),("BOTTOMPADDING",(0,0),(-1,0),3)]
    def rate(cost,charge): return (charge/cost) if cost else 0
    for name,cost,dep,charge,nbv in ppe.get("classes",[]):
        data.append([name, money(cost), (f"{rate(cost,charge)*100:.0f}%" if cost else "-"),
                     money(charge), money(dep), money(nbv)])
        ri=len(data)-1
        st+=[("ALIGN",(1,ri),(-1,ri),"RIGHT"),("TOPPADDING",(0,ri),(-1,ri),2),("BOTTOMPADDING",(0,ri),(-1,ri),2)]
    t=ppe.get("total")
    if t:
        cost,dep,charge,nbv=t
        data.append(["TOTAL", money(cost), "", money(charge), money(dep), money(nbv)])
        ri=len(data)-1
        st+=[("FONTNAME",(0,ri),(-1,ri),"DVB"),("ALIGN",(1,ri),(-1,ri),"RIGHT"),("TEXTCOLOR",(0,ri),(-1,ri),NAVY),
             ("LINEABOVE",(0,ri),(-1,ri),0.6,colors.black),("LINEBELOW",(1,ri),(-1,ri),1.0,NAVY)]
    w0=PAGE_W-LM-RM-28*mm-13*mm-26*mm-26*mm-26*mm
    tab=Table(data,colWidths=[w0,28*mm,13*mm,26*mm,26*mm,26*mm])
    tab.setStyle(TableStyle([("FONTNAME",(0,0),(-1,-1),"DV"),("FONTSIZE",(0,0),(-1,-1),8),
                             ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0)]+st))
    return tab

def heading(txt,subtitle=None):
    el=[Paragraph(txt,h1),
        Table([[""]],colWidths=[PAGE_W-LM-RM],
              style=TableStyle([("LINEBELOW",(0,0),(-1,0),1.3,RULE),("TOPPADDING",(0,0),(-1,0),1),("BOTTOMPADDING",(0,0),(-1,0),0)])),
        Spacer(1,6)]
    if subtitle: el.append(Paragraph(subtitle,sub))
    return el

def signatures(names,date_str,role="Director"):
    w=(PAGE_W-LM-RM)/len(names)
    t=Table([[Paragraph("&nbsp;",sig) for _ in names],           # blank signing space above the line
             [Paragraph(nm,sig) for nm in names],
             [Paragraph(role,sigsub) for _ in names],
             [Paragraph(f"Dated: {date_str}",sigdate) for _ in names]],
            colWidths=[w]*len(names), rowHeights=[16*mm,None,None,None])
    t.setStyle(TableStyle([("LINEABOVE",(0,1),(-1,1),0.8,colors.black),("TOPPADDING",(0,1),(-1,1),4),
                           ("VALIGN",(0,0),(-1,0),"BOTTOM"),
                           ("LEFTPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),3)]))
    return t


def _hex(v, default):
    try: return colors.HexColor(v) if v else colors.HexColor(default)
    except Exception: return colors.HexColor(default)

def _apply_theme(meta):
    """Set the document colours from the firm's branding (or Kayode defaults). Called per build."""
    global NAVY, NAVY2, GOLD, RULE
    NAVY = _hex(meta.get("primary_color"), "#13315C"); NAVY2 = NAVY; RULE = NAVY
    GOLD = _hex(meta.get("accent_color"), "#C49A2C")
    for stl in (h1, h2, sig):
        stl.textColor = NAVY

def build(data,out_path):
    global UNIT; UNIT=data.get("meta",{}).get("currency_unit","N")
    _apply_theme(data.get("meta", {}))
    doc=BaseDocTemplate(out_path,pagesize=A4,leftMargin=LM,rightMargin=RM,topMargin=TM,bottomMargin=BM+8*mm)
    frame=Frame(LM,BM+8*mm,PAGE_W-LM-RM,PAGE_H-TM-(BM+8*mm),id="main",leftPadding=0,rightPadding=0,topPadding=0,bottomPadding=0)
    doc.addPageTemplates([PageTemplate(id="cover",frames=[frame],onPage=cover_page),
                          PageTemplate(id="body",frames=[frame],onPage=header_footer)])
    doc.afs=data["meta"]; fy_first=data["meta"]["first_year"]; E=data["entity"]; M=data["meta"]
    el=[]; A=el.append
    L=data.get("labels",{})
    ENT=L.get("entity_word","Company"); GOVS=L.get("gov_plural","Directors"); GOV=L.get("gov_singular","Director")
    T_SOCI=L.get("soci_title","Statement of Profit or Loss and Other Comprehensive Income")
    T_SOCE=L.get("soce_title","Statement of Changes in Equity")
    ADDR=L.get("addressee","Members"); PERF=L.get("perf_phrase","financial performance")
    REPORT_TITLE=L.get("report_title","Directors' Report")
    RESP_TITLE=L.get("resp_title","Statement of Directors' Responsibilities")
    ADVISERS_TITLE=L.get("advisers_title","Directors, Professional Advisers and Registered Office")
    SOCI_FOOT=L.get("soci_footnote","Profit for the year is wholly derived from continuing operations. There were no items of other comprehensive income during the year.")
    A(NextPageTemplate("body")); A(PageBreak())
    # 1 advisers
    A(Spacer(1,2))
    for x in heading(ADVISERS_TITLE): A(x)
    def kv(label,value):
        A(Paragraph(label,h2))
        if isinstance(value,(list,tuple)):
            for v in value: A(Paragraph(v,body))
        else: A(Paragraph(value,body))
    kv("RC Number",E["rc"]); kv(GOVS,E["directors"]); kv("Principal Activity",E["activity"])
    kv("Registered Office",E["office"]); kv("Bankers",E["bankers"]); kv("External Auditor",E["auditor"])
    _firm_addr = M.get("firm_address") or M.get("firm_city")
    if _firm_addr: A(Paragraph(_firm_addr, body))
    A(Spacer(1,10))
    A(Paragraph(f"In accordance with section 357(2) of the Companies and Allied Matters Act, 2020, Messrs {E['auditor']} will continue in office as Auditors without a resolution being passed.",small_i))
    A(Spacer(1,4))
    A(Paragraph(f"Per the firm's house-style signature convention, {M['sig_words']} of the {ENT}'s {GOVS} sign these financial statements on behalf of the {L.get('board_word','Board')}.",small_i))
    A(PageBreak())
    # 2 contents
    for x in heading("Contents"): A(x)
    toc=[(ADVISERS_TITLE,"2"),(REPORT_TITLE,"4"),
         (RESP_TITLE,"6"),("Independent Auditor's Report","7"),
         (T_SOCI,"9"),("Statement of Financial Position","10"),
         ("Statement of Cash Flows","11"),(T_SOCE,"12"),("Notes to the Financial Statements","13")]
    tt=Table([[Paragraph(t,body),Paragraph(p,bodyc)] for t,p in toc],colWidths=[PAGE_W-LM-RM-14*mm,14*mm])
    tt.setStyle(TableStyle([("ALIGN",(1,0),(1,-1),"RIGHT"),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),("LEFTPADDING",(0,0),(-1,-1),0)]))
    A(tt); A(PageBreak())
    # 3 directors report
    for x in heading(REPORT_TITLE): A(x)
    A(Paragraph(f"The {GOVS} present their report together with the audited financial statements of {E['name'].upper()} (“the {ENT}”) for the year ended {M['period_end']}.",body))
    A(Paragraph("Principal activities",h2)); A(Paragraph(E["activity_para"],body))
    A(Paragraph(GOVS,h2))
    A(Paragraph(f"The names of the {GOVS} of the {ENT} are set out on page 2 of these financial statements. In accordance with section 303 of the Companies and Allied Matters Act, 2020, no {GOV} has notified the {ENT} of any declarable interest in contracts or transactions in which the {ENT} was involved during the year under review.",body))
    A(Paragraph("Results for the year",h2)); A(Paragraph(M["results_para"],body))
    A(Paragraph("Property, plant and equipment",h2)); A(Paragraph(M["ppe_para"],body))
    A(Paragraph("Employment of disabled persons",h2))
    A(Paragraph(f"The {ENT} maintains an open employment policy for disabled persons as part of its corporate social responsibility, provided the candidate meets the requirements of the role.",body))
    A(Paragraph("Health, safety and welfare of employees",h2))
    A(Paragraph(f"Arrangements are made for the adequate security and protection of staff at the {ENT}'s premises, and applicable health and safety regulations are observed at all times.",body))
    A(Paragraph("Employee involvement and training",h2))
    A(Paragraph(f"The {ENT} provides facilities for regular on-the-job training of staff. Regular consultative meetings are held by management to keep employees abreast of developments within the {ENT}, including its strategic plans and achievements.",body))
    A(PageBreak())
    A(Paragraph("Post-reporting-date events",h2))
    A(Paragraph(f"There were no significant events after the reporting date which could have materially affected the financial position of the {ENT} as at {M['period_end']} or the results presented in these financial statements that have not been adequately disclosed or provided for.",body))
    A(Paragraph("Auditors",h2))
    A(Paragraph(f"In accordance with the applicable provisions of the Companies and Allied Matters Act, 2020, Messrs {E['auditor']} have indicated their willingness to continue in office as Auditors of the {ENT}.",body))
    A(Spacer(1,14)); A(KeepTogether([Paragraph("BY ORDER OF THE "+L.get('board_word','Board').upper(),sig), Spacer(1,18), signatures(M["signatories"],M["sign_date"],GOV)])); A(PageBreak())
    # 4 responsibilities
    for x in heading(RESP_TITLE,f"In relation to the financial statements for the year ended {M['period_end']}."): A(x)
    A(Paragraph(f"The Companies and Allied Matters Act, 2020 requires the {GOVS} to prepare financial statements for each financial year that give a true and fair view of the financial position of the {ENT} as at the end of the financial year and of its {PERF} and cash flows for the year then ended.",body))
    A(Paragraph(f"In preparing these financial statements, the {GOVS} have ensured that:",body))
    for b in ["proper accounting records are maintained;","appropriate accounting policies are selected and consistently applied;",
              "judgements and estimates made are reasonable and prudent;",f"the applicable {M['framework_short']} has been followed; and",
              f"the financial statements are prepared on a going-concern basis unless it is inappropriate to presume that the {ENT} will continue in business."]:
        A(Paragraph(b,bullet,bulletText="•"))
    A(Spacer(1,4))
    A(Paragraph(f"The {GOVS} are responsible for maintaining adequate accounting records, safeguarding the assets of the {ENT}, and taking reasonable steps for the prevention and detection of fraud and other irregularities. The fulfilment of this responsibility is discharged through the establishment and maintenance of sound management and accounting systems, the maintenance of an organisational structure that provides for the delegation of authority, and the constant review of operational performance against approved plans and budgets.",body))
    A(Paragraph(f"The {GOVS} further confirm that these financial statements have been prepared in accordance with the {M['framework']} and in compliance with the provisions of the Companies and Allied Matters Act, 2020.",body))
    A(Paragraph(f"The {GOVS} have assessed the ability of the {ENT} to continue as a going concern and have no reason to believe that the {ENT} will not remain a going concern in the foreseeable future.",body))
    A(KeepTogether([Spacer(1,18), signatures(M["signatories"],M["sign_date"],GOV)])); A(PageBreak())
    # 5 auditor report
    for x in heading("Independent Auditor's Report",f"To the {ADDR} of {E['name'].upper()}"): A(x)
    A(Paragraph("Report on the Audit of the Financial Statements",h2))
    A(Paragraph("Opinion",sub))
    A(Paragraph(f"We have audited the financial statements of {E['name'].upper()} (\u201cthe {ENT}\u201d), which comprise the Statement of Financial Position as at {M['period_end']}, and the {T_SOCI}, the {T_SOCE} and the Statement of Cash Flows for the year then ended, and the notes to the financial statements, including a summary of significant accounting policies.",body))
    A(Paragraph(f"In our opinion, the accompanying financial statements give a true and fair view of the financial position of the {ENT} as at {M['period_end']}, and of its {PERF} and its cash flows for the year then ended in accordance with the {M['framework_short']} and in the manner required by the Companies and Allied Matters Act, 2020 and the Financial Reporting Council of Nigeria Act.",body))
    A(Paragraph("Basis for Opinion",sub))
    A(Paragraph(f"We conducted our audit in accordance with International Standards on Auditing (ISAs). Our responsibilities under those standards are further described in the Auditor's Responsibilities for the Audit of the Financial Statements section of our report. We are independent of the {ENT} in accordance with the International Ethics Standards Board for Accountants' International Code of Ethics for Professional Accountants (including International Independence Standards) and the ethical requirements that are relevant to our audit of the financial statements in Nigeria, and we have fulfilled our other ethical responsibilities in accordance with these requirements. We believe that the audit evidence we have obtained is sufficient and appropriate to provide a basis for our opinion.",body))
    A(Paragraph("Other Information",sub))
    A(Paragraph(f"The {GOVS} are responsible for the other information. The other information comprises the {REPORT_TITLE} and the {RESP_TITLE}, but does not include the financial statements and our auditor's report thereon. Our opinion on the financial statements does not cover the other information and we do not express any form of assurance conclusion thereon. In connection with our audit of the financial statements, our responsibility is to read the other information and, in doing so, consider whether the other information is materially inconsistent with the financial statements or our knowledge obtained in the audit, or otherwise appears to be materially misstated. If, based on the work we have performed, we conclude that there is a material misstatement of this other information, we are required to report that fact. We have nothing to report in this regard.",body))
    A(Paragraph(f"Responsibilities of the {GOVS} for the Financial Statements",sub))
    A(Paragraph(f"The {GOVS} are responsible for the preparation of financial statements that give a true and fair view in accordance with the {M['framework_short']} and the requirements of the Companies and Allied Matters Act, 2020 and the Financial Reporting Council of Nigeria Act, and for such internal control as the {GOVS} determine is necessary to enable the preparation of financial statements that are free from material misstatement, whether due to fraud or error.",body))
    A(Paragraph(f"In preparing the financial statements, the {GOVS} are responsible for assessing the {ENT}'s ability to continue as a going concern, disclosing, as applicable, matters related to going concern, and using the going-concern basis of accounting unless the {GOVS} either intend to liquidate the {ENT} or to cease operations, or have no realistic alternative but to do so.",body))
    A(Paragraph("Auditor's Responsibilities for the Audit of the Financial Statements",sub))
    A(Paragraph("Our objectives are to obtain reasonable assurance about whether the financial statements as a whole are free from material misstatement, whether due to fraud or error, and to issue an auditor's report that includes our opinion. Reasonable assurance is a high level of assurance, but is not a guarantee that an audit conducted in accordance with ISAs will always detect a material misstatement when it exists. Misstatements can arise from fraud or error and are considered material if, individually or in the aggregate, they could reasonably be expected to influence the economic decisions of users taken on the basis of these financial statements.",body))
    A(Paragraph(f"As part of an audit in accordance with ISAs, we exercise professional judgement and maintain professional scepticism throughout the audit. We also identify and assess the risks of material misstatement, whether due to fraud or error, design and perform audit procedures responsive to those risks, and obtain audit evidence that is sufficient and appropriate to provide a basis for our opinion; obtain an understanding of internal control relevant to the audit in order to design audit procedures that are appropriate in the circumstances; evaluate the appropriateness of accounting policies used and the reasonableness of accounting estimates and related disclosures made by the {GOVS}; conclude on the appropriateness of the {GOVS}' use of the going-concern basis of accounting; and evaluate the overall presentation, structure and content of the financial statements.",body))
    A(Paragraph(f"We communicate with the {GOVS} regarding, among other matters, the planned scope and timing of the audit and significant audit findings, including any significant deficiencies in internal control that we identify during our audit.",body))
    A(Spacer(1,4))
    _sigblk=[Paragraph("Report on Other Legal and Regulatory Requirements",h2),
             Paragraph(f"In accordance with the requirements of Schedule 6 of the Companies and Allied Matters Act, 2020, we confirm that: (a) we have obtained all the information and explanations which to the best of our knowledge and belief were necessary for the purposes of our audit; (b) in our opinion, proper books of account have been kept by the {ENT}, so far as appears from our examination of those books; and (c) the {ENT}'s Statement of Financial Position and {T_SOCI} are in agreement with the books of account.",body),
             Spacer(1,16)]                                  # keep report-on-legal + signature together
    if M["mode"]=="final" and M.get("signature_image"):
        _sig=scaled_image(M["signature_image"], 48*mm, 18*mm); _sig.hAlign="LEFT"; _sigblk += [_sig, Spacer(1,1)]
    _sigblk += [Paragraph(E["auditor_name"],sig), Paragraph("Chartered Accountants",sigdate)]
    if M["mode"]=="final":
        _sigblk += [Spacer(1,4), Paragraph(f"FRC No: {M['frc_no']}",sigsub), Paragraph(f"ICAN Stamp No: {M['ican_stamp_no']}",sigsub)]
        if M.get("stamp_image"):
            _sigblk += [Spacer(1,4), scaled_image(M["stamp_image"], 42*mm, 42*mm)]
    _sigblk.append(Spacer(1,6))
    _aud_loc = M.get("firm_address") or M.get("firm_city")   # auditor's own address, never the client's
    if _aud_loc: _sigblk.append(Paragraph(_aud_loc, body))
    _sigblk.append(Paragraph(f"Dated: {M['sign_date']}",sigdate))
    A(KeepTogether(_sigblk)); A(PageBreak())
    # 6 SOCI
    for x in heading(T_SOCI,f"For the year ended {M['period_end']}"): A(x)
    _cyl=M["fy"] or "CY"; _pyl=(str(int(M["fy"])-1) if (M["fy"] and str(M["fy"]).isdigit()) else "PY")
    A(stmt_table(data["soci"],fy_first,_cyl,_pyl)); A(Spacer(1,8))
    A(Paragraph(SOCI_FOOT,small_i))
    A(Paragraph("The accompanying notes form an integral part of these financial statements.",small_i)); A(PageBreak())
    # 7 SOFP
    for x in heading("Statement of Financial Position",f"As at {M['period_end']}"): A(x)
    A(stmt_table(data["sofp"],fy_first,_cyl,_pyl)); A(Spacer(1,8))
    A(KeepTogether([Paragraph(f"These financial statements were approved by the {L.get('board_word','Board of Directors')} on {M['sign_date']} and signed on its behalf by:",body), Spacer(1,16), signatures(M["signatories"],M["sign_date"],GOV)])); A(Spacer(1,6))
    A(Paragraph("The accompanying notes form an integral part of these financial statements.",small_i)); A(PageBreak())
    # 8 SCF
    for x in heading("Statement of Cash Flows",f"For the year ended {M['period_end']} (indirect method)"): A(x)
    A(stmt_table(data["scf"],fy_first,_cyl,_pyl)); A(Spacer(1,8))
    A(Paragraph("The accompanying notes form an integral part of these financial statements.",small_i)); A(PageBreak())
    # 9 SOCE
    for x in heading(T_SOCE,f"For the year ended {M['period_end']}"): A(x)
    A(soce_table(data["soce"],fy_first,data.get("soce_headers"))); A(Spacer(1,8))
    A(Paragraph("The accompanying notes form an integral part of these financial statements.",small_i)); A(PageBreak())
    # 10 notes
    for x in heading("Notes to the Financial Statements",f"For the year ended {M['period_end']}"): A(x)
    _py=(str(int(M["fy"])-1) if (M["fy"] and str(M["fy"]).isdigit()) else "")
    for note in data["notes"]:
        A(Paragraph(note["title"],h2))                     # h2 keepWithNext prevents an orphaned heading
        for p in note.get("paras",[]): A(Paragraph(p,body))
        if note.get("table"): A(note_table(note["table"],fy_first,(M["fy"] or ""),_py))
        if note.get("ppe"): A(ppe_table(note["ppe"],fy_first))
        for g in (note.get("grids") or []):
            if g.get("subhead"):
                _sh=Paragraph("<b>"+g["subhead"]+"</b>",body); _sh.keepWithNext=1
                A(_sh)
            A(grid_table(g))
        A(Spacer(1,10))
    if data.get("fin_summary"):
        A(PageBreak())
        for x in heading("Five-Year Financial Summary"): A(x)
        A(grid_table(data["fin_summary"]))
        A(Spacer(1,6)); A(Paragraph("The five-year financial summary does not form part of the audited financial statements.",small_i))
    _ts=data.get("tax_schedules")
    if _ts:
        _py3=(str(int(M["fy"])-1) if (M["fy"] and str(M["fy"]).isdigit()) else "")
        A(PageBreak())
        A(Paragraph("Supplementary Schedules",h1))
        A(Paragraph("The following schedules are prepared for taxation purposes and do not form part of the audited financial statements.",small_i))
        A(Spacer(1,10))
        if _ts.get("cit"):
            for x in heading(_ts["cit"]["title"], _ts["cit"].get("subtitle")): A(x)
            A(cit_table(_ts["cit"]["rows"], (M["fy"] or ""), _py3)); A(Spacer(1,14))
        if _ts.get("capallow"):
            for x in heading(_ts["capallow"]["title"]): A(x)
            A(grid_table({"headers":_ts["capallow"]["headers"],"rows":_ts["capallow"]["rows"],"money_from":2}))
    doc.build(el)

if __name__=="__main__":
    import afs_data
    mode=sys.argv[1] if len(sys.argv)>1 else "draft"
    d=afs_data.get_data(mode)
    out=sys.argv[2] if len(sys.argv)>2 else f"LIVINGHEART_AFS_2025_{mode}.pdf"
    build(d,out); print("written",out)
