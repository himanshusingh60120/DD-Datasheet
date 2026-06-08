# Publications Pipeline — Vercel

Frontend (`public/index.html`) + Python serverless backend (`api/`) that runs your
GA4 → Google Sheets pipeline. Upload the monthly Excel in the browser, authorize
Google, hit run. Plus one-click buttons to open each publication's sheet.

## Structure
```
ga4-pipeline/
├── vercel.json          # routes api/process.py to Python runtime (300s, 1GB)
├── requirements.txt     # python deps
├── api/
│   ├── auth.py          # exchanges OAuth code -> tokens
│   └── process.py       # the full pipeline (your script, adapted)
└── public/
    └── index.html       # the frontend UI
```

## One-time Google Cloud setup
1. In Google Cloud Console → **APIs & Services → Credentials**, open your existing
   OAuth client (the one from `client_secret...json`).
2. Under **Authorized redirect URIs**, add your Vercel URL exactly, e.g.
   `https://your-app.vercel.app/` (and `http://localhost:3000/` for local testing).
   The URI must match `REDIRECT_URI` (origin + path) used by the frontend.
3. Keep the Client ID and Client Secret handy.

## Vercel setup
1. Push this folder to a Git repo and import it in Vercel (or `vercel` CLI from this dir).
2. In **Project → Settings → Environment Variables**, add:
   - `GOOGLE_CLIENT_ID`  = your OAuth client id
   - `GOOGLE_CLIENT_SECRET` = your OAuth client secret
3. In `public/index.html`, set the `GOOGLE_CLIENT_ID` constant (top of `<script>`)
   to the same client id. (It's public by design — only the secret stays server-side.)
4. Deploy.

## Notes
- The manual copy-paste localhost flow from Colab is replaced by a proper redirect
  flow that returns to your deployed page and exchanges the code server-side.
- `maxDuration` is 300s. GA4 across 4 properties + writes usually fits; if a month is
  huge and times out, that limit needs a Pro plan (Hobby caps lower on some accounts).
- All config (property IDs, sheet IDs, domains, tab/sheet names) lives in `api/process.py`
  and is identical to your script. Sheet-button IDs also live in `index.html`.
- "Dry run" sends `dry_run=true` so nothing is written — good for verifying GA4 numbers.
```
