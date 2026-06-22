# AFS Engine — deploy

Files: afs_generator.py, afs_extract.py, main.py, requirements.txt, Dockerfile

## Deploy on Render (easiest)
1. Put these files in a GitHub repo.
2. Render → New → Web Service → connect the repo → "Docker" runtime.
3. Add environment variable: ENGINE_API_KEY = <your long random secret>.
4. Deploy. Render gives you a URL like https://afs-engine.onrender.com
5. In Lovable secrets:
   - ENGINE_URL     = https://afs-engine.onrender.com
   - ENGINE_API_KEY = <the same secret>

## Endpoints
- GET  /health   -> {"ok": true}
- POST /generate -> multipart workbook (+ optional stamp) + form fields; returns the PDF.
  Header x-api-key must equal ENGINE_API_KEY.

## Test after deploy
curl -s https://YOUR-URL/health
curl -s -X POST https://YOUR-URL/generate -H "x-api-key: YOURKEY" \
  -F "workbook=@client.xlsx" -F "mode=draft" -F "template=SME" -F "n_signatories=2" \
  -o draft.pdf
