# Deploy the AFS Engine to Render — step by step

You only do this once. ~15 minutes. At the end you get the real ENGINE_URL.

Your API key (use the SAME value in Render and in Lovable):
    afs_GGxZCwI2d7mBQoZ9I_-_QcIIb9FqtHYPX3XO2-ezUrM

---

## A. Put the engine folder on GitHub
1. Go to github.com → New repository → name it `afs-engine` → Create.
2. Upload these 6 files from your `engine/` folder (drag-drop into GitHub's "uploading an
   existing file" page, or use GitHub Desktop):
   - afs_generator.py
   - afs_extract.py
   - main.py
   - requirements.txt
   - Dockerfile
   - README_DEPLOY.md
3. Commit.

## B. Create the Render service
4. Go to render.com → sign in → New + → Web Service.
5. Connect your GitHub and pick the `afs-engine` repo.
6. Settings:
   - Language / Runtime: Docker  (Render auto-detects the Dockerfile)
   - Instance type: Starter is fine to begin (Free tier sleeps after inactivity, so the first
     request after idle is slow — Starter avoids that).
   - Region: closest to you.
7. Add an Environment Variable:
   - Key:  ENGINE_API_KEY
   - Value: afs_GGxZCwI2d7mBQoZ9I_-_QcIIb9FqtHYPX3XO2-ezUrM
8. Create Web Service. Wait for the build to finish (first Docker build installs LibreOffice,
   so allow a few minutes). Status goes "Live".
9. Copy the URL Render shows at the top, e.g. https://afs-engine.onrender.com

## C. Test it (optional but recommended)
10. In a browser, open:  https://YOUR-URL/health   → should show {"ok": true}

## D. Point Lovable at it
11. In Lovable → project settings / secrets:
    - ENGINE_URL     = https://afs-engine.onrender.com   (your real Render URL, with https://)
    - ENGINE_API_KEY = afs_GGxZCwI2d7mBQoZ9I_-_QcIIb9FqtHYPX3XO2-ezUrM
12. Re-run Prompt 3's smoke test in Lovable: upload a workbook → Generate Draft → a PDF should
    come back and the tie-out checks should render. Then Prompts 4 and 5 will work too.

---

### Troubleshooting
- /health works but Generate fails with 401 → the key in Lovable doesn't match the key in Render.
- Build fails on LibreOffice → make sure you used the Dockerfile (Runtime = Docker), not a Python
  runtime.
- First request very slow → Free tier cold-start; upgrade to Starter or just retry.
- Timeout on large workbooks → Render → service → Settings → raise the request timeout.

### Railway alternative
Railway works the same way: New Project → Deploy from GitHub → it reads the Dockerfile → add the
ENGINE_API_KEY variable → use the generated domain as ENGINE_URL.
