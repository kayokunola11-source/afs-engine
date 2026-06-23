# -*- coding: utf-8 -*-
"""AFS engine HTTP service.
POST /generate  (x-api-key)  multipart: workbook=<xlsx>, stamp=<png|optional>
                              form: mode, template, first_year, n_signatories, ican_stamp_no
  -> application/pdf  (headers: X-Tie-Outs, X-Extracted-Summary as JSON)
GET  /health -> {"ok": true}
"""
import os, tempfile, subprocess, json, shutil, base64
from fastapi import FastAPI, UploadFile, File, Form, Header, HTTPException
from fastapi.responses import JSONResponse
import afs_extract, afs_generator

API_KEY = os.environ.get("ENGINE_API_KEY", "")
app = FastAPI(title="AFS Engine")

def recalc(xlsx_in, work):
    """Recalculate the formula-linked workbook with LibreOffice (Excel caches no values).
    Output goes to a SEPARATE dir so LibreOffice never tries to overwrite the input."""
    outdir = os.path.join(work, "recalc_out"); os.makedirs(outdir, exist_ok=True)
    prof = "file://" + os.path.join(work, "lo_profile")     # isolated profile -> concurrency-safe
    subprocess.run(["soffice","--headless","--nologo","--norestore",
                    f"-env:UserInstallation={prof}",
                    "--convert-to","xlsx:Calc MS Excel 2007 XML","--outdir",outdir,xlsx_in],
                   check=True, timeout=180,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return os.path.join(outdir, os.path.splitext(os.path.basename(xlsx_in))[0] + ".xlsx")


def clean_to_transparent(path, crop=True):
    """Make a scanned signature/stamp usable on the page: knock out the white paper
    background to transparency and crop to the ink. Leaves already-transparent PNGs alone."""
    from PIL import Image
    img = Image.open(path).convert("RGBA")
    r, g, b, a = img.split()
    if a.getextrema()[0] < 250:
        new_a = a                                   # already has transparency
    else:
        lum = Image.merge("RGB", (r, g, b)).convert("L")
        new_a = lum.point(lambda v: 0 if v > 226 else int(max(0, min(255, (232 - v) * 255 / 70))))
    out = Image.merge("RGBA", (r, g, b, new_a))
    if crop:
        bbox = new_a.getbbox()
        if bbox:
            w, h = out.size; pad = max(4, int(0.03 * max(w, h)))
            out = out.crop((max(0, bbox[0]-pad), max(0, bbox[1]-pad),
                            min(w, bbox[2]+pad), min(h, bbox[3]+pad)))
    out.save(path)

@app.get("/health")
def health():
    return {"ok": True}

@app.post("/generate")
async def generate(
    workbook: UploadFile = File(...),
    stamp: UploadFile | None = File(None),
    signature: UploadFile | None = File(None),
    mode: str = Form("draft"),
    template: str = Form("SME"),
    first_year: str = Form("auto"),          # "true" | "false" | "auto"
    n_signatories: int = Form(2),
    ican_stamp_no: str = Form(""),
    x_api_key: str = Header(default=""),
):
    if API_KEY and x_api_key != API_KEY:
        raise HTTPException(401, "Invalid API key")
    work = tempfile.mkdtemp(prefix="afs_")
    try:
        src = os.path.join(work, "in.xlsx")
        with open(src, "wb") as f: f.write(await workbook.read())
        recalced = recalc(src, work)

        stamp_path = None
        if stamp is not None:
            stamp_path = os.path.join(work, "stamp.png")
            with open(stamp_path, "wb") as f: f.write(await stamp.read())
            try: clean_to_transparent(stamp_path)
            except Exception: pass
        sig_path = None
        if signature is not None:
            sig_path = os.path.join(work, "signature.png")
            with open(sig_path, "wb") as f: f.write(await signature.read())
            try: clean_to_transparent(sig_path)
            except Exception: pass

        fy = None if first_year == "auto" else (first_year.lower() == "true")
        data = afs_extract.get_data(recalced, mode=mode, first_year=fy,
                                    n_sig=int(n_signatories), template=template,
                                    ican_stamp_no=ican_stamp_no, stamp_image=stamp_path, signature_image=sig_path)
        out = os.path.join(work, "afs.pdf")
        afs_generator.build(data, out)                       # pass 1
        try:
            from pypdf import PdfReader
            data["meta"]["total_pages"] = len(PdfReader(out).pages)
            afs_generator.build(data, out)                   # pass 2 (page count)
        except Exception:
            pass
        pdf = open(out, "rb").read()

        def gv(rows, lbl, k="cy"):
            for r in rows:
                if r.get("label","").upper() == lbl.upper(): return r.get(k)
            return None

        extracted = {
            "entity_name": data["meta"]["entity_name"],
            "period_end":  data["meta"]["period_end"],
            "revenue":     gv(data["soci"], "Revenue"),
            "profit":      gv(data["soci"], "PROFIT/(LOSS) FOR THE YEAR"),
            "total_assets":gv(data["sofp"], "TOTAL ASSETS"),
            "first_year":  data["meta"]["first_year"],
        }
        body = {
            "pdf_base64": base64.b64encode(pdf).decode("ascii"),
            "extracted":  extracted,
            "tie_outs":   data["tie_outs"],          # [{"name","pass"}] x5, Lovable's names
            "flags":      data.get("flags", []),
            "mode":       mode,
            "page_count": data["meta"].get("total_pages"),
        }
        return JSONResponse(body)
    finally:
        shutil.rmtree(work, ignore_errors=True)
