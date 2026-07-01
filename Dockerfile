FROM python:3.11-slim

# LibreOffice (workbook recalculation) + DejaVu fonts (the ₦ glyph)
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-calc fonts-dejavu fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY afs_generator.py afs_extract.py afs_jukes.py afs_am.py afs_ifrs.py afs_notes.py calc_core.py afs_pycore.py disclosure_check.py ifrs_sme_checklist.json main.py ./

ENV PORT=8000
EXPOSE 8000
CMD ["sh","-c","uvicorn main:app --host 0.0.0.0 --port ${PORT}"]
