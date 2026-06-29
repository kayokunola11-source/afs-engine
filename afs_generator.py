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
h2=Sx("h2",fontName="DVB",fontSize=10.5,textColor=NAVY2,spaceBefore=12,spaceAfter=4)
sub=Sx("sub",fontName="DVI",fontSize=9.5,textColor=GREY,spaceAfter=8)
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
    y=PAGE_H-58*mm; c.setFillColor(NAVY); c.setFont("DVB",34); c.drawString(LM,y,st["short_name"].upper())
    y-=13*mm; c.setFont("DVB",15); c.setFillColor(NAVY2); c.drawString(LM,y,st["name_line2"].upper())
    y-=7*mm; c.setFont("DV",9.5); c.setFillColor(GREY); c.drawString(LM,y,st["rc"])
    y-=34*mm; c.setFillColor(NAVY); c.setFont("DVB",22); c.drawString(LM,y,"AUDITED")
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
    data.append(["","","₦","₦" if show_py else ""]); ru=1
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
    data=[["",str(cy_year),(str(py_year) if show_py else "")],["","₦","₦" if show_py else ""]]
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
    data=[["",str(cy_year),str(py_year)],["","\u20a6","\u20a6"]]
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

def soce_table(rows,first_year=False):
    data=[["","Share capital","Retained earnings","Total equity"]]
    styles=[("FONTNAME",(0,0),(-1,0),"DVB"),("ALIGN",(1,0),(-1,0),"RIGHT"),("TEXTCOLOR",(0,0),(-1,0),NAVY2),
            ("FONTSIZE",(0,0),(-1,0),8.6),("LINEBELOW",(0,0),(-1,0),0.7,NAVY2),("BOTTOMPADDING",(0,0),(-1,0),3)]
    data.append(["","₦","₦","₦"])
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
    t=Table([[Paragraph(nm,sig) for nm in names],
             [Paragraph(role,sigsub) for _ in names],
             [Paragraph(f"Dated: {date_str}",sigdate) for _ in names]],colWidths=[w]*len(names))
    t.setStyle(TableStyle([("LINEABOVE",(0,0),(-1,0),0.8,colors.black),("TOPPADDING",(0,0),(-1,0),4),
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
    _apply_theme(data.get("meta", {}))
    doc=BaseDocTemplate(out_path,pagesize=A4,leftMargin=LM,rightMargin=RM,topMargin=TM,bottomMargin=BM+8*mm)
    frame=Frame(LM,BM+8*mm,PAGE_W-LM-RM,PAGE_H-TM-(BM+8*mm),id="main",leftPadding=0,rightPadding=0,topPadding=0,bottomPadding=0)
    doc.addPageTemplates([PageTemplate(id="cover",frames=[frame],onPage=cover_page),
                          PageTemplate(id="body",frames=[frame],onPage=header_footer)])
    doc.afs=data["meta"]; fy_first=data["meta"]["first_year"]; E=data["entity"]; M=data["meta"]
    el=[]; A=el.append
    A(NextPageTemplate("body")); A(PageBreak())
    # 1 advisers
    A(Spacer(1,2))
    for x in heading("Directors, Professional Advisers and Registered Office"): A(x)
    def kv(label,value):
        A(Paragraph(label,h2))
        if isinstance(value,(list,tuple)):
            for v in value: A(Paragraph(v,body))
        else: A(Paragraph(value,body))
    kv("RC Number",E["rc"]); kv("Directors",E["directors"]); kv("Principal Activity",E["activity"])
    kv("Registered Office",E["office"]); kv("Bankers",E["bankers"]); kv("External Auditor",E["auditor"])
    _firm_addr = M.get("firm_address") or M.get("firm_city")
    if _firm_addr: A(Paragraph(_firm_addr, body))
    A(Spacer(1,10))
    A(Paragraph(f"In accordance with section 357(2) of the Companies and Allied Matters Act, 2020, Messrs {E['auditor']} will continue in office as Auditors without a resolution being passed.",small_i))
    A(Spacer(1,4))
    A(Paragraph(f"Per the firm's house-style signature convention, {M['sig_words']} of the Company's Directors sign these financial statements on behalf of the Board.",small_i))
    A(PageBreak())
    # 2 contents
    for x in heading("Contents"): A(x)
    toc=[("Directors, Professional Advisers and Registered Office","2"),("Directors' Report","4"),
         ("Statement of Directors' Responsibilities","6"),("Independent Auditor's Report","7"),
         ("Statement of Profit or Loss and Other Comprehensive Income","9"),("Statement of Financial Position","10"),
         ("Statement of Cash Flows","11"),("Statement of Changes in Equity","12"),("Notes to the Financial Statements","13")]
    tt=Table([[Paragraph(t,body),Paragraph(p,bodyc)] for t,p in toc],colWidths=[PAGE_W-LM-RM-14*mm,14*mm])
    tt.setStyle(TableStyle([("ALIGN",(1,0),(1,-1),"RIGHT"),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),("LEFTPADDING",(0,0),(-1,-1),0)]))
    A(tt); A(PageBreak())
    # 3 directors report
    for x in heading("Directors' Report"): A(x)
    A(Paragraph(f"The Directors present their report together with the audited financial statements of {E['name'].upper()} (“the Company”) for the year ended {M['period_end']}.",body))
    A(Paragraph("Principal activities",h2)); A(Paragraph(E["activity_para"],body))
    A(Paragraph("Directors",h2))
    A(Paragraph("The names of the Directors of the Company are set out on page 2 of these financial statements. In accordance with section 303 of the Companies and Allied Matters Act, 2020, no Director has notified the Company of any declarable interest in contracts or transactions in which the Company was involved during the year under review.",body))
    A(Paragraph("Results for the year",h2)); A(Paragraph(M["results_para"],body))
    A(Paragraph("Property, plant and equipment",h2)); A(Paragraph(M["ppe_para"],body))
    A(Paragraph("Employment of disabled persons",h2))
    A(Paragraph("The Company maintains an open employment policy for disabled persons as part of its corporate social responsibility, provided the candidate meets the requirements of the role.",body))
    A(Paragraph("Health, safety and welfare of employees",h2))
    A(Paragraph("Arrangements are made for the adequate security and protection of staff at the Company's premises, and applicable health and safety regulations are observed at all times.",body))
    A(Paragraph("Employee involvement and training",h2))
    A(Paragraph("The Company provides facilities for regular on-the-job training of staff. Regular consultative meetings are held by management to keep employees abreast of developments within the Company, including its strategic plans and achievements.",body))
    A(PageBreak())
    A(Paragraph("Post-reporting-date events",h2))
    A(Paragraph(f"There were no significant events after the reporting date which could have materially affected the financial position of the Company as at {M['period_end']} or the results presented in these financial statements that have not been adequately disclosed or provided for.",body))
    A(Paragraph("Auditors",h2))
    A(Paragraph(f"In accordance with the applicable provisions of the Companies and Allied Matters Act, 2020, Messrs {E['auditor']} have indicated their willingness to continue in office as Auditors of the Company.",body))
    A(Spacer(1,14)); A(Paragraph("BY ORDER OF THE BOARD",sig)); A(Spacer(1,18)); A(signatures(M["signatories"],M["sign_date"])); A(PageBreak())
    # 4 responsibilities
    for x in heading("Statement of Directors' Responsibilities",f"In relation to the financial statements for the year ended {M['period_end']}."): A(x)
    A(Paragraph("The Companies and Allied Matters Act, 2020 requires the Directors to prepare financial statements for each financial year that give a true and fair view of the financial position of the Company as at the end of the financial year and of its financial performance and cash flows for the year then ended.",body))
    A(Paragraph("In preparing these financial statements, the Directors have ensured that:",body))
    for b in ["proper accounting records are maintained;","appropriate accounting policies are selected and consistently applied;",
              "judgements and estimates made are reasonable and prudent;",f"the applicable {M['framework_short']} has been followed; and",
              "the financial statements are prepared on a going-concern basis unless it is inappropriate to presume that the Company will continue in business."]:
        A(Paragraph(b,bullet,bulletText="•"))
    A(Spacer(1,4))
    A(Paragraph("The Directors are responsible for maintaining adequate accounting records, safeguarding the assets of the Company, and taking reasonable steps for the prevention and detection of fraud and other irregularities. The fulfilment of this responsibility is discharged through the establishment and maintenance of sound management and accounting systems, the maintenance of an organisational structure that provides for the delegation of authority, and the constant review of operational performance against approved plans and budgets.",body))
    A(Paragraph(f"The Directors further confirm that these financial statements have been prepared in accordance with the {M['framework']} and in compliance with the provisions of the Companies and Allied Matters Act, 2020.",body))
    A(Paragraph("The Directors have assessed the ability of the Company to continue as a going concern and have no reason to believe that the Company will not remain a going concern in the foreseeable future.",body))
    A(Spacer(1,18)); A(signatures(M["signatories"],M["sign_date"])); A(PageBreak())
    # 5 auditor report
    for x in heading("Independent Auditor's Report",f"To the Members of {E['name'].upper()}"): A(x)
    A(Paragraph("Report on the Financial Statements",h2))
    A(Paragraph(f"We have audited the accompanying financial statements of {E['name'].upper()} (“the Company”), which comprise the Statement of Financial Position as at {M['period_end']}, the Statement of Profit or Loss and Other Comprehensive Income, the Statement of Cash Flows, the Statement of Changes in Equity for the year then ended, and the notes to the financial statements.",body))
    A(Paragraph("Directors' Responsibility for the Financial Statements",h2))
    A(Paragraph(f"The Directors are responsible for the preparation of financial statements that give a true and fair view in accordance with the {M['framework_short']} and the Statements of Accounting Standards applicable in Nigeria, in the manner required by the Companies and Allied Matters Act of Nigeria and the Financial Reporting Council of Nigeria Act, 2020, and for such internal control as the Directors determine is necessary to enable the preparation of financial statements that are free from material misstatement, whether due to fraud or error.",body))
    A(Paragraph("Auditor's Responsibility",h2))
    A(Paragraph("Our responsibility is to express an opinion on these financial statements based on our audit. We conducted our audit in accordance with International Standards on Auditing. Those standards require that we comply with ethical requirements and plan and perform the audit to obtain reasonable assurance about whether the financial statements are free from material misstatement.",body))
    A(Paragraph("An audit involves performing procedures to obtain audit evidence about the amounts and disclosures in the financial statements. The procedures selected depend on the auditor's judgement, including the assessment of the risks of material misstatement of the financial statements, whether due to fraud or error.",body))
    A(Paragraph("An audit also includes evaluating the appropriateness of accounting policies used and the reasonableness of accounting estimates made by the Directors, as well as evaluating the overall presentation of the financial statements. We believe that the audit evidence we have obtained is sufficient and appropriate to provide a basis for our audit opinion.",body))
    A(Paragraph("Opinion",h2))
    A(Paragraph(f"In our opinion, the accompanying financial statements give a true and fair view of the financial position of {E['name'].upper()} (“the Company”) as at {M['period_end']}, and of its financial performance and cash flows for the year then ended in accordance with the {M['framework_short']} and the Statements of Accounting Standards applicable in Nigeria, and in the manner required by the Companies and Allied Matters Act of Nigeria, 2020.",body))
    A(PageBreak())
    A(Paragraph("Report on Other Legal and Regulatory Requirements",h2))
    A(Paragraph("Compliance with the Requirements of Schedule 6 of the Companies and Allied Matters Act of Nigeria",sub))
    A(Paragraph("In our opinion, proper books of account have been kept by the Company, so far as appears from our examination of those books, and the Company's Statement of Financial Position and Statement of Profit or Loss are in agreement with the books of account.",body))
    A(Spacer(1,16))
    if M["mode"]=="final" and M.get("signature_image"):
        _sig=scaled_image(M["signature_image"], 48*mm, 18*mm); _sig.hAlign="LEFT"; A(_sig); A(Spacer(1,1))
    A(Paragraph(E["auditor_name"],sig)); A(Paragraph("Chartered Accountants",sigdate))
    if M["mode"]=="final":
        A(Spacer(1,4)); A(Paragraph(f"FRC No: {M['frc_no']}",sigsub)); A(Paragraph(f"ICAN Stamp No: {M['ican_stamp_no']}",sigsub))
        if M.get("stamp_image"):
            A(Spacer(1,4)); A(scaled_image(M["stamp_image"], 42*mm, 42*mm))
    A(Spacer(1,6))
    _aud_loc = M.get("firm_address") or M.get("firm_city")   # auditor's own address, never the client's
    if _aud_loc: A(Paragraph(_aud_loc, body))
    A(Paragraph(f"Dated: {M['sign_date']}",sigdate)); A(PageBreak())
    # 6 SOCI
    for x in heading("Statement of Profit or Loss and Other Comprehensive Income",f"For the year ended {M['period_end']}"): A(x)
    _cyl=M["fy"] or "CY"; _pyl=(str(int(M["fy"])-1) if (M["fy"] and str(M["fy"]).isdigit()) else "PY")
    A(stmt_table(data["soci"],fy_first,_cyl,_pyl)); A(Spacer(1,8))
    A(Paragraph("Profit for the year is wholly derived from continuing operations. There were no items of other comprehensive income during the year.",small_i))
    A(Paragraph("The accompanying notes form an integral part of these financial statements.",small_i)); A(PageBreak())
    # 7 SOFP
    for x in heading("Statement of Financial Position",f"As at {M['period_end']}"): A(x)
    A(stmt_table(data["sofp"],fy_first,_cyl,_pyl)); A(Spacer(1,8))
    A(Paragraph(f"These financial statements were approved by the Board of Directors on {M['sign_date']} and signed on its behalf by:",body))
    A(Spacer(1,16)); A(signatures(M["signatories"],M["sign_date"])); A(Spacer(1,6))
    A(Paragraph("The accompanying notes form an integral part of these financial statements.",small_i)); A(PageBreak())
    # 8 SCF
    for x in heading("Statement of Cash Flows",f"For the year ended {M['period_end']} (indirect method)"): A(x)
    A(stmt_table(data["scf"],fy_first,_cyl,_pyl)); A(Spacer(1,8))
    A(Paragraph("The accompanying notes form an integral part of these financial statements.",small_i)); A(PageBreak())
    # 9 SOCE
    for x in heading("Statement of Changes in Equity",f"For the year ended {M['period_end']}"): A(x)
    A(soce_table(data["soce"],fy_first)); A(Spacer(1,8))
    A(Paragraph("The accompanying notes form an integral part of these financial statements.",small_i)); A(PageBreak())
    # 10 notes
    for x in heading("Notes to the Financial Statements",f"For the year ended {M['period_end']}"): A(x)
    _py=(str(int(M["fy"])-1) if (M["fy"] and str(M["fy"]).isdigit()) else "")
    for note in data["notes"]:
        block=[Paragraph(note["title"],h2)]
        for p in note.get("paras",[]): block.append(Paragraph(p,body))
        if note.get("table"): block.append(Spacer(1,2)); block.append(note_table(note["table"],fy_first,(M["fy"] or ""),_py))
        if note.get("ppe"): block.append(Spacer(1,2)); block.append(ppe_table(note["ppe"],fy_first))
        grids=note.get("grids") or []
        start=0
        if grids and not note.get("table") and not note.get("ppe"):
            g=grids[0]                                  # keep the heading with its first schedule
            if g.get("subhead"): block.append(Spacer(1,2)); block.append(Paragraph("<b>"+g["subhead"]+"</b>",body))
            block.append(grid_table(g)); start=1
        A(KeepTogether(block))
        for g in grids[start:]:
            sg=[]
            if g.get("subhead"): sg.append(Paragraph("<b>"+g["subhead"]+"</b>",body))
            sg.append(grid_table(g))
            A(Spacer(1,3)); A(KeepTogether(sg))
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
